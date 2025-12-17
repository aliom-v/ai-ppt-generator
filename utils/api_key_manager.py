"""增强的API密钥安全保护模块"""
import os
import base64
import hashlib
import secrets
import time
from typing import Optional, Dict, Any
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import redis
from utils.logger import get_logger

logger = get_logger("security")


class APIKeyManager:
    """API密钥安全管理器"""

    def __init__(self, redis_url: Optional[str] = None):
        self.redis_client = None
        if redis_url:
            try:
                self.redis_client = redis.from_url(redis_url)
                self.redis_client.ping()
            except Exception as e:
                logger.warning(f"Redis连接失败，使用内存缓存: {e}")
        self.memory_cache: Dict[str, Any] = {}
        self._init_encryption()

    def _init_encryption(self):
        """初始化加密密钥"""
        # 从环境变量获取主密钥，如果没有则生成
        master_key = os.getenv('ENCRYPTION_MASTER_KEY')
        if not master_key:
            logger.warning("未设置ENCRYPTION_MASTER_KEY，使用临时密钥（重启后失效）")
            master_key = Fernet.generate_key().decode()
            # 可选：保存到环境变量或配置文件
            os.environ['ENCRYPTION_MASTER_KEY'] = master_key

        # 使用PBKDF2派生加密密钥
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=b'ai-ppt-generator-salt',
            iterations=100000,
        )
        key = base64.urlsafe_b64encode(kdf.derive(master_key.encode()))
        self.cipher = Fernet(key)

    def encrypt_api_key(self, api_key: str) -> str:
        """加密API密钥"""
        try:
            encrypted = self.cipher.encrypt(api_key.encode())
            # 返回Base64编码的加密字符串
            return base64.urlsafe_b64encode(encrypted).decode()
        except Exception as e:
            logger.error(f"加密API密钥失败: {e}")
            raise ValueError("API密钥加密失败")

    def decrypt_api_key(self, encrypted_key: str) -> str:
        """解密API密钥"""
        try:
            encrypted = base64.urlsafe_b64decode(encrypted_key.encode())
            decrypted = self.cipher.decrypt(encrypted)
            return decrypted.decode()
        except Exception as e:
            logger.error(f"解密API密钥失败: {e}")
            raise ValueError("API密钥解密失败")

    def hash_api_key(self, api_key: str) -> str:
        """生成API密钥的哈希值（用于识别但不暴露密钥）"""
        salt = secrets.token_bytes(32)
        key_hash = hashlib.pbkdf2_hmac('sha256', api_key.encode(), salt, 100000)
        return base64.b64encode(salt + key_hash).decode()

    def verify_api_key_hash(self, api_key: str, stored_hash: str) -> bool:
        """验证API密钥哈希"""
        try:
            decoded = base64.b64decode(stored_hash.encode())
            salt = decoded[:32]
            stored_key_hash = decoded[32:]
            new_key_hash = hashlib.pbkdf2_hmac('sha256', api_key.encode(), salt, 100000)
            return secrets.compare_digest(stored_key_hash, new_key_hash)
        except Exception:
            return False

    def store_api_key(self, user_id: str, api_key: str, provider: str = "openai") -> str:
        """存储加密的API密钥"""
        try:
            # 加密API密钥
            encrypted_key = self.encrypt_api_key(api_key)

            # 生成密钥哈希
            key_hash = self.hash_api_key(api_key)

            # 存储元数据
            metadata = {
                "encrypted_key": encrypted_key,
                "provider": provider,
                "created_at": int(time.time()),
                "last_used": None,
                "usage_count": 0
            }

            # 存储到Redis或内存
            storage_key = f"api_key:{user_id}:{provider}"
            if self.redis_client:
                self.redis_client.hset(storage_key, mapping=metadata)
                self.redis_client.set(f"key_hash:{user_id}:{provider}", key_hash)
            else:
                self.memory_cache[storage_key] = metadata
                self.memory_cache[f"key_hash:{user_id}:{provider}"] = key_hash

            logger.info(f"成功存储API密钥: user={user_id}, provider={provider}")
            return key_hash

        except Exception as e:
            logger.error(f"存储API密钥失败: {e}")
            raise

    def get_api_key(self, user_id: str, provider: str = "openai") -> Optional[str]:
        """获取解密的API密钥"""
        try:
            storage_key = f"api_key:{user_id}:{provider}"

            # 获取加密密钥
            if self.redis_client:
                metadata = self.redis_client.hgetall(storage_key)
                if not metadata:
                    return None
                encrypted_key = metadata.get(b"encrypted_key", b"").decode()
            else:
                metadata = self.memory_cache.get(storage_key)
                if not metadata:
                    return None
                encrypted_key = metadata["encrypted_key"]

            # 解密并返回
            api_key = self.decrypt_api_key(encrypted_key)

            # 更新使用记录
            self._update_usage(user_id, provider)

            return api_key

        except Exception as e:
            logger.error(f"获取API密钥失败: {e}")
            return None

    def _update_usage(self, user_id: str, provider: str):
        """更新使用记录"""
        storage_key = f"api_key:{user_id}:{provider}"
        if self.redis_client:
            self.redis_client.hincrby(storage_key, "usage_count", 1)
            self.redis_client.hset(storage_key, "last_used", int(time.time()))
        else:
            if storage_key in self.memory_cache:
                self.memory_cache[storage_key]["usage_count"] += 1
                self.memory_cache[storage_key]["last_used"] = int(time.time())

    def rotate_api_key(self, user_id: str, new_api_key: str, provider: str = "openai") -> bool:
        """轮换API密钥"""
        try:
            # 验证新密钥格式
            if not self._validate_api_key_format(new_api_key, provider):
                raise ValueError("无效的API密钥格式")

            # 存储新密钥
            self.store_api_key(user_id, new_api_key, provider)
            logger.info(f"成功轮换API密钥: user={user_id}, provider={provider}")
            return True

        except Exception as e:
            logger.error(f"轮换API密钥失败: {e}")
            return False

    def _validate_api_key_format(self, api_key: str, provider: str) -> bool:
        """验证API密钥格式"""
        if provider == "openai":
            # OpenAI API密钥格式：sk-开头，48-51字符
            return api_key.startswith("sk-") and 48 <= len(api_key) <= 51
        elif provider == "claude":
            # Claude API密钥格式：sk-ant-开头
            return api_key.startswith("sk-ant-") and len(api_key) > 30
        else:
            # 通用验证：至少20字符
            return len(api_key) >= 20

    def delete_api_key(self, user_id: str, provider: str = "openai") -> bool:
        """删除API密钥"""
        try:
            storage_key = f"api_key:{user_id}:{provider}"
            hash_key = f"key_hash:{user_id}:{provider}"

            if self.redis_client:
                self.redis_client.delete(storage_key, hash_key)
            else:
                self.memory_cache.pop(storage_key, None)
                self.memory_cache.pop(hash_key, None)

            logger.info(f"成功删除API密钥: user={user_id}, provider={provider}")
            return True

        except Exception as e:
            logger.error(f"删除API密钥失败: {e}")
            return False

    def mask_api_key(self, api_key: str, mask_char: str = "*", visible_chars: int = 8) -> str:
        """遮蔽API密钥用于日志显示"""
        if len(api_key) <= visible_chars:
            return mask_char * len(api_key)

        visible_start = api_key[:visible_chars//2]
        visible_end = api_key[-(visible_chars//2):]
        masked_length = len(api_key) - visible_chars

        return f"{visible_start}{mask_char * masked_length}{visible_end}"

    def get_key_info(self, user_id: str, provider: str = "openai") -> Optional[Dict[str, Any]]:
        """获取API密钥信息（不包含密钥本身）"""
        try:
            storage_key = f"api_key:{user_id}:{provider}"

            if self.redis_client:
                metadata = self.redis_client.hgetall(storage_key)
                if not metadata:
                    return None

                return {
                    "provider": metadata.get(b"provider", b"").decode(),
                    "created_at": int(metadata.get(b"created_at", 0)),
                    "last_used": int(metadata.get(b"last_used", 0)) if metadata.get(b"last_used") else None,
                    "usage_count": int(metadata.get(b"usage_count", 0))
                }
            else:
                metadata = self.memory_cache.get(storage_key)
                if not metadata:
                    return None

                info = metadata.copy()
                del info["encrypted_key"]  # 不返回加密密钥
                return info

        except Exception as e:
            logger.error(f"获取API密钥信息失败: {e}")
            return None


# 全局实例
_key_manager: Optional[APIKeyManager] = None


def get_key_manager() -> APIKeyManager:
    """获取全局API密钥管理器实例"""
    global _key_manager
    if _key_manager is None:
        redis_url = os.getenv('REDIS_URL')
        _key_manager = APIKeyManager(redis_url)
    return _key_manager


def secure_log_api_key(api_key: str, visible_chars: int = 8) -> str:
    """安全地记录API密钥（已遮蔽）"""
    manager = get_key_manager()
    return manager.mask_api_key(api_key, visible_chars=visible_chars)


# 向后兼容的辅助函数
def load_api_key_from_env() -> str:
    """从环境变量加载API密钥（保持向后兼容）"""
    api_key = os.getenv('AI_API_KEY')
    if not api_key:
        raise ValueError("未设置 AI_API_KEY 环境变量")
    return api_key


def validate_api_key(api_key: str, provider: str = "openai") -> bool:
    """验证API密钥格式"""
    manager = get_key_manager()
    return manager._validate_api_key_format(api_key, provider)