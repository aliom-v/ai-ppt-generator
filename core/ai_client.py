"""大模型调用封装模块 - 优化版"""
import json
import time
import threading
from typing import Dict, Any, Optional, Callable
from openai import OpenAI, APIError, APIConnectionError, RateLimitError, AuthenticationError

from config.settings import AIConfig, settings
from core.prompt_builder import get_system_prompt, build_user_prompt
from core.ai_common import (
    clean_json_response,
    calculate_backoff,
    is_retryable_error,
    calculate_batches,
    build_batch_prompt_first,
    build_batch_prompt_continue,
    build_json_error_message,
)
from utils.logger import get_logger
from utils.cache import get_cache

logger = get_logger("ai_client")


# ==================== 客户端连接池 ====================

class AIClientPool:
    """AI 客户端连接池 - 复用连接以提高性能"""

    _instance = None
    _lock = threading.Lock()

    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    cls._instance._clients: Dict[str, OpenAI] = {}
                    cls._instance._client_lock = threading.Lock()
        return cls._instance

    def get_client(self, config: AIConfig) -> OpenAI:
        """获取或创建客户端（基于配置的唯一键）"""
        # 使用 api_base_url 和 api_key 前8位作为键
        key = f"{config.api_base_url}:{config.api_key[:8] if config.api_key else ''}"

        with self._client_lock:
            if key not in self._clients:
                self._clients[key] = OpenAI(
                    api_key=config.api_key,
                    base_url=config.api_base_url,
                    timeout=config.timeout
                )
                logger.debug(f"创建新的 AI 客户端: {config.api_base_url}")
            return self._clients[key]

    def close_all(self):
        """关闭所有客户端连接"""
        with self._client_lock:
            for client in self._clients.values():
                try:
                    client.close()
                except Exception:
                    pass
            self._clients.clear()


def get_ai_client_pool() -> AIClientPool:
    """获取全局客户端池实例"""
    return AIClientPool()


class AIClientError(Exception):
    """AI 客户端错误基类"""
    pass


class APIKeyError(AIClientError):
    """API Key 错误"""
    pass


class RateLimitExceeded(AIClientError):
    """API 限流错误"""
    pass


class JSONParseError(AIClientError):
    """JSON 解析错误"""
    pass


class NetworkError(AIClientError):
    """网络错误"""
    pass


def _call_api_with_retry(
    client: OpenAI,
    model_name: str,
    system_prompt: str,
    user_prompt: str,
    max_retries: int = 3,
    temperature: float = 0.7
) -> str:
    """带智能重试机制的 API 调用

    特性：
    - 指数退避 + 随机抖动
    - 区分可重试/不可重试错误
    - 自动解析 Retry-After 响应头
    - 兼容 OpenAI 和 Claude（通过兼容层如 OpenRouter）
    """
    model_lower = model_name.lower()
    is_claude = "claude" in model_lower
    is_o1 = model_lower.startswith("o1")  # OpenAI o1 系列不支持 system role
    last_error = None

    for attempt in range(max_retries):
        try:
            # 构建消息列表
            if is_o1:
                # o1 系列：不支持 system role，合并到 user 消息
                messages = [
                    {"role": "user", "content": f"{system_prompt}\n\n{user_prompt}"}
                ]
            else:
                # OpenAI GPT 系列和 Claude（通过兼容层）：使用标准格式
                messages = [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ]

            # 构建请求参数
            request_params = {
                "model": model_name,
                "messages": messages,
                "max_tokens": 8192,
            }

            # 温度参数（o1 系列不支持）
            if not is_o1:
                request_params["temperature"] = temperature

            # JSON 模式（Claude 和 o1 不支持 response_format）
            if not is_claude and not is_o1:
                request_params["response_format"] = {"type": "json_object"}

            response = client.chat.completions.create(**request_params)

            # 提取响应内容
            content = None
            if isinstance(response, str):
                content = response
            elif hasattr(response, 'choices') and response.choices:
                message = response.choices[0].message
                content = message.content if message else None

            # 检查内容是否为空
            if not content:
                raise AIClientError("AI 返回了空内容，请重试或更换模型")

            return content

        except AuthenticationError as e:
            # 认证错误不重试，直接抛出
            raise APIKeyError(f"API Key 无效或认证失败: {e}")

        except RateLimitError as e:
            last_error = e
            if attempt < max_retries - 1:
                # 尝试从响应中获取 Retry-After
                retry_after = getattr(e, 'retry_after', None)
                if retry_after:
                    wait_time = float(retry_after)
                else:
                    wait_time = calculate_backoff(attempt, base_delay=2.0, max_delay=60.0)
                logger.warning(f"API 限流（第 {attempt + 1} 次），{wait_time:.1f} 秒后重试...")
                time.sleep(wait_time)
            else:
                raise RateLimitExceeded(f"API 限流，已重试 {max_retries} 次: {e}")

        except APIConnectionError as e:
            last_error = e
            if attempt < max_retries - 1:
                wait_time = calculate_backoff(attempt, base_delay=1.0, max_delay=30.0)
                logger.warning(f"网络错误（第 {attempt + 1} 次），{wait_time:.1f} 秒后重试...")
                time.sleep(wait_time)
            else:
                raise NetworkError(f"网络连接失败，已重试 {max_retries} 次: {e}")

        except APIError as e:
            # 检查是否可重试
            if not is_retryable_error(e):
                if "invalid_api_key" in str(e).lower() or "401" in str(e):
                    raise APIKeyError(f"API Key 无效: {e}")
                raise AIClientError(f"API 错误（不可重试）: {e}")

            last_error = e
            if attempt < max_retries - 1:
                wait_time = calculate_backoff(attempt, base_delay=1.0, max_delay=15.0)
                logger.warning(f"API 错误（第 {attempt + 1} 次），{wait_time:.1f} 秒后重试...")
                time.sleep(wait_time)
            else:
                raise AIClientError(f"API 调用失败: {e}")

    raise AIClientError(f"API 调用失败: {last_error}")


def generate_ppt_plan(
    topic: str,
    audience: str,
    page_count: int = 0,
    description: str = "",
    auto_page_count: bool = False,
    config: Optional[AIConfig] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    use_cache: bool = True
) -> Dict[str, Any]:
    """调用大模型生成 PPT 结构（支持分批生成大型 PPT）

    Args:
        topic: PPT 主题
        audience: 目标受众
        page_count: 内容页数量（不含封面）
        description: 详细描述/要点/参考资料
        auto_page_count: 是否让 AI 自动判断页数
        config: AI 配置（可选，默认使用环境变量）
        progress_callback: 进度回调函数，接收 (current_batch, total_batches, message)
        use_cache: 是否使用缓存（默认开启）

    Returns:
        包含 PPT 结构的字典

    Raises:
        AIClientError: 当 API 调用失败时
        JSONParseError: 当返回格式错误时
    """
    # 使用传入的配置或默认配置
    if config is None:
        config = settings.to_ai_config()

    config.validate()

    # 尝试从缓存获取
    if use_cache and not auto_page_count:
        cache = get_cache()
        cached = cache.get(topic, audience, page_count, description, config.model_name)
        if cached:
            return cached
    
    # 计算是否需要分批
    batches = calculate_batches(page_count) if not auto_page_count and page_count > 35 else [page_count]
    total_batches = len(batches)
    
    if total_batches > 1:
        logger.info(f"页数较多（{page_count}页），将分 {total_batches} 批生成...")
        result = _generate_ppt_plan_batched(
            topic, audience, batches, description, config, progress_callback
        )
    else:
        # 单批生成
        result = _generate_ppt_plan_single(
            topic, audience, page_count, description, auto_page_count, config
        )
    
    # 保存到缓存
    if use_cache and not auto_page_count:
        cache = get_cache()
        cache.set(topic, audience, page_count, result, description, config.model_name)

    return result


def _generate_ppt_plan_batched(
    topic: str,
    audience: str,
    batches: list,
    description: str,
    config: AIConfig,
    progress_callback: Optional[Callable[[int, int, str], None]] = None
) -> Dict[str, Any]:
    """分批生成 PPT 结构"""
    # 使用连接池获取客户端
    client = get_ai_client_pool().get_client(config)

    total_batches = len(batches)
    all_slides = []
    title = ""
    subtitle = ""

    # 记录已生成的内容摘要，用于续写
    generated_summary = []

    for batch_idx, batch_pages in enumerate(batches):
        current_batch = batch_idx + 1

        if progress_callback:
            progress_callback(current_batch, total_batches, f"正在生成第 {current_batch}/{total_batches} 批...")

        logger.info(f"生成第 {current_batch}/{total_batches} 批（{batch_pages} 页）...")

        # 构建分批提示词
        if batch_idx == 0:
            # 第一批：生成开头部分
            batch_prompt = build_batch_prompt_first(
                topic, audience, batch_pages, total_batches, description
            )
        else:
            # 后续批次：续写
            batch_prompt = build_batch_prompt_continue(
                topic, audience, batch_pages, current_batch, total_batches,
                generated_summary, is_last=(current_batch == total_batches)
            )

        system_prompt = get_system_prompt()

        try:
            content = _call_api_with_retry(
                client=client,
                model_name=config.model_name,
                system_prompt=system_prompt,
                user_prompt=batch_prompt,
                max_retries=config.max_retries,
                temperature=config.temperature
            )

            cleaned_content = clean_json_response(content)
            batch_result = json.loads(cleaned_content)

            # 提取标题（只从第一批获取）
            if batch_idx == 0:
                title = batch_result.get("title", topic)
                subtitle = batch_result.get("subtitle", "")

            # 收集 slides
            batch_slides = batch_result.get("slides", [])

            # 过滤掉 ending 页（除了最后一批）
            if current_batch < total_batches:
                batch_slides = [s for s in batch_slides if s.get("type") != "ending"]

            all_slides.extend(batch_slides)

            # 记录摘要用于续写
            for slide in batch_slides:
                slide_title = slide.get("title", "")
                if slide_title:
                    generated_summary.append(slide_title)

            logger.info(f"第 {current_batch} 批完成，获得 {len(batch_slides)} 页")

        except Exception as e:
            logger.error(f"第 {current_batch} 批生成失败: {e}")
            raise

    # 合并结果
    result = {
        "title": title,
        "subtitle": subtitle,
        "slides": all_slides
    }

    logger.info(f"分批生成完成，共 {len(all_slides)} 页")

    # 注意：使用连接池时不再手动关闭客户端

    return result


def _generate_ppt_plan_single(
    topic: str,
    audience: str,
    page_count: int,
    description: str,
    auto_page_count: bool,
    config: AIConfig
) -> Dict[str, Any]:
    """单批生成 PPT 结构（原有逻辑）"""
    # 使用连接池获取客户端
    client = get_ai_client_pool().get_client(config)

    system_prompt = get_system_prompt()
    user_prompt = build_user_prompt(topic, audience, page_count, description, auto_page_count)

    logger.info(f"生成 PPT: {topic} | 受众: {audience} | 模型: {config.model_name}")

    # 初始化 content 变量，避免在异常处理中使用 locals()
    content = ""

    try:
        content = _call_api_with_retry(
            client=client,
            model_name=config.model_name,
            system_prompt=system_prompt,
            user_prompt=user_prompt,
            max_retries=config.max_retries,
            temperature=config.temperature
        )

        logger.debug(f"AI 返回内容长度: {len(content)} 字符")

        content_lower = content.strip().lower()
        if content_lower.startswith('<!doctype') or content_lower.startswith('<html') or '<html' in content_lower[:500]:
            raise AIClientError(
                f"API 返回了 HTML 页面而不是 AI 响应。请检查：\n"
                f"1. API Base URL 是否正确（当前: {config.api_base_url}）\n"
                f"2. 确保 URL 以 /v1 结尾\n"
                f"3. API Key 是否有效"
            )

        cleaned_content = clean_json_response(content)

        if not cleaned_content:
            raise JSONParseError(f"AI 返回了无效内容: {content[:300]}")

        plan_dict = json.loads(cleaned_content)

        return plan_dict

    except json.JSONDecodeError as e:
        error_msg = build_json_error_message(e, content)
        raise JSONParseError(error_msg)
    except AIClientError:
        raise
    except Exception as e:
        raise AIClientError(f"生成失败: {e}")
    # 注意：使用连接池时不再手动关闭客户端


def test_api_connection(config: AIConfig) -> Dict[str, Any]:
    """测试 API 连通性
    
    Args:
        config: AI 配置
        
    Returns:
        测试结果字典，包含 success, message, model_info 等
    """
    result = {
        "success": False,
        "message": "",
        "model": config.model_name,
        "api_base": config.api_base_url,
        "response_time": 0,
    }
    
    try:
        config.validate()
    except ValueError as e:
        result["message"] = str(e)
        return result

    start_time = time.time()

    try:
        client = OpenAI(
            api_key=config.api_key,
            base_url=config.api_base_url,
            timeout=15  # 测试用较短超时
        )
        
        # 发送简单测试请求
        response = client.chat.completions.create(
            model=config.model_name,
            messages=[{"role": "user", "content": "Hi, just testing. Reply with: OK"}],
            max_tokens=10,
            temperature=0
        )
        
        elapsed = time.time() - start_time
        result["response_time"] = round(elapsed * 1000)  # 毫秒
        
        # 检查响应
        content = None
        if hasattr(response, 'choices') and response.choices:
            message = response.choices[0].message
            content = message.content if message else None
        
        # 检查是否有错误状态（某些 API 返回特殊格式）
        if hasattr(response, 'status') and response.status:
            status = str(response.status)
            if status != '200' and status != 'success':
                msg = getattr(response, 'msg', '') or f"状态码: {status}"
                result["message"] = f"API 返回错误: {msg}"
                return result
        
        if not content:
            result["message"] = "API 返回了空响应，请检查模型名称是否正确"
            return result
        
        # 检查是否返回 HTML
        content_lower = content.strip().lower()
        if content_lower.startswith('<!doctype') or content_lower.startswith('<html') or '<html' in content_lower[:500]:
            result["message"] = f"API 返回了 HTML 页面。请检查 API Base URL 是否正确，确保以 /v1 结尾（当前: {config.api_base_url}）"
            return result
        
        result["success"] = True
        result["message"] = f"连接成功！响应时间: {result['response_time']}ms"
        result["response"] = content[:100]
        
    except RateLimitError:
        result["message"] = "API 限流，但连接正常。请稍后再试"
        result["success"] = True  # 限流说明 API 是通的
    except APIConnectionError as e:
        result["message"] = f"网络连接失败: {e}"
    except APIError as e:
        error_str = str(e).lower()
        if "401" in error_str or "invalid_api_key" in error_str:
            result["message"] = "API Key 无效，请检查"
        elif "404" in error_str:
            result["message"] = "模型不存在或 API 路径错误"
        else:
            result["message"] = f"API 错误: {e}"
    except Exception as e:
        result["message"] = f"测试失败: {e}"
    
    return result
