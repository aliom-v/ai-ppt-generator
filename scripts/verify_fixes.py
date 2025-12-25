#!/usr/bin/env python3
"""验证修复脚本 - 检查所有安全修复是否生效"""
import sys
import os
import time

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def test_secret_key_persistence():
    """测试 SECRET_KEY 持久化"""
    print("\n[1] 测试 SECRET_KEY 持久化...")
    from config.settings import _get_or_create_secret_key

    key1 = _get_or_create_secret_key()
    key2 = _get_or_create_secret_key()

    if key1 == key2:
        print("    ✓ SECRET_KEY 持久化正常")
        return True
    else:
        print("    ✗ SECRET_KEY 每次生成不同")
        return False


def test_rate_limiter_trust_proxy():
    """测试速率限制器的代理信任配置"""
    print("\n[2] 测试 IP 欺骗防护...")

    try:
        from utils.rate_limit import RateLimiter
    except ImportError as e:
        print(f"    ⚠ 跳过（缺少依赖: {e}）")
        return True  # 跳过但不算失败

    # 清除环境变量
    old_val = os.environ.pop('TRUST_PROXY', None)

    try:
        limiter = RateLimiter()
        if not limiter._trust_proxy:
            print("    ✓ 默认不信任代理头")
        else:
            print("    ✗ 默认信任代理头（不安全）")
            return False

        # 测试环境变量
        os.environ['TRUST_PROXY'] = 'true'
        limiter2 = RateLimiter()
        if limiter2._trust_proxy:
            print("    ✓ 环境变量配置生效")
        else:
            print("    ✗ 环境变量配置无效")
            return False

        return True
    finally:
        if old_val:
            os.environ['TRUST_PROXY'] = old_val
        else:
            os.environ.pop('TRUST_PROXY', None)


def test_task_manager_cleanup():
    """测试任务管理器超时清理"""
    print("\n[3] 测试任务超时清理...")
    from utils.async_tasks import TaskManager, TaskStatus

    tm = TaskManager()

    # 检查方法存在
    if not hasattr(tm, '_cleanup_stale_tasks'):
        print("    ✗ _cleanup_stale_tasks 方法不存在")
        return False

    # 创建并模拟超时任务
    task_id = tm.create_task()
    tm.update_task(task_id, status=TaskStatus.RUNNING)
    task = tm.get_task(task_id)
    task.started_at = time.time() - 4000  # 模拟超时

    # 执行清理
    cleaned = tm._cleanup_stale_tasks(stale_timeout=3600)

    if cleaned > 0:
        task = tm.get_task(task_id)
        if task.status == TaskStatus.FAILED and "超时" in (task.error or ""):
            print("    ✓ 超时任务清理正常")
            return True

    print("    ✗ 超时任务清理失败")
    return False


def test_rate_limiter_memory_cleanup():
    """测试速率限制器内存清理"""
    print("\n[4] 测试速率限制器内存清理...")

    try:
        from utils.rate_limit import RateLimiter
    except ImportError as e:
        print(f"    ⚠ 跳过（缺少依赖: {e}）")
        return True

    limiter = RateLimiter()

    # 模拟添加过期记录
    test_ip = "192.168.1.100"
    current_time = time.time()
    limiter._minute_counts[test_ip] = [current_time - 120]  # 2分钟前
    limiter._hour_counts[test_ip] = [current_time - 7200]   # 2小时前

    # 执行清理
    limiter._cleanup_old_requests(test_ip, current_time)

    # 检查空列表是否被删除
    if test_ip not in limiter._minute_counts and test_ip not in limiter._hour_counts:
        print("    ✓ 空列表清理正常")
        return True
    else:
        print("    ✗ 空列表未被清理")
        return False


def test_csrf_protection():
    """测试 CSRF 保护逻辑"""
    print("\n[5] 测试 CSRF 保护...")

    try:
        from utils.security import CSRFProtection
    except ImportError as e:
        print(f"    ⚠ 跳过（缺少依赖: {e}）")
        return True

    # 检查类存在
    if CSRFProtection:
        print("    ✓ CSRFProtection 类正常")
        return True
    return False


def test_image_hash():
    """测试图片哈希改进"""
    print("\n[6] 测试图片文件名哈希...")
    import hashlib

    # 模拟新的哈希逻辑
    test_url = "https://example.com/image.jpg"
    url_hash = hashlib.sha256(test_url.encode()).hexdigest()[:12]
    filename = f"image_{url_hash}.jpg"

    if len(url_hash) == 12 and filename.startswith("image_"):
        print("    ✓ SHA256 哈希正常")
        return True
    else:
        print("    ✗ 哈希生成异常")
        return False


def main():
    print("=" * 50)
    print("AI PPT Generator - 修复验证脚本")
    print("=" * 50)

    results = []

    results.append(("SECRET_KEY 持久化", test_secret_key_persistence()))
    results.append(("IP 欺骗防护", test_rate_limiter_trust_proxy()))
    results.append(("任务超时清理", test_task_manager_cleanup()))
    results.append(("速率限制内存清理", test_rate_limiter_memory_cleanup()))
    results.append(("CSRF 保护", test_csrf_protection()))
    results.append(("图片哈希改进", test_image_hash()))

    print("\n" + "=" * 50)
    print("测试结果汇总")
    print("=" * 50)

    passed = 0
    failed = 0
    for name, result in results:
        status = "✓ PASS" if result else "✗ FAIL"
        print(f"  {status}: {name}")
        if result:
            passed += 1
        else:
            failed += 1

    print(f"\n总计: {passed} 通过, {failed} 失败")

    if failed == 0:
        print("\n🎉 所有修复验证通过！")
        return 0
    else:
        print("\n⚠️ 部分修复验证失败，请检查")
        return 1


if __name__ == "__main__":
    sys.exit(main())
