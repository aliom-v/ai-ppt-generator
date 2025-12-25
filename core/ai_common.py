"""AI 客户端公共模块

抽取 ai_client.py 和 ai_client_async.py 的公共函数，避免代码重复。
"""
import json
import random
from typing import List

from openai import AuthenticationError, APIError


def clean_json_response(content: str) -> str:
    """清理 AI 返回的 JSON 内容

    处理常见的格式问题：
    - 移除 markdown 代码块标记
    - 提取 JSON 对象
    - 替换中文引号
    - 移除 BOM 标记
    """
    content = content.strip()

    # 移除 markdown 代码块
    if content.startswith("```json"):
        content = content[7:]
    elif content.startswith("```"):
        content = content[3:]
    if content.endswith("```"):
        content = content[:-3]
    content = content.strip()

    # 提取 JSON 对象
    first_brace = content.find('{')
    last_brace = content.rfind('}')
    if first_brace != -1 and last_brace != -1 and first_brace < last_brace:
        content = content[first_brace:last_brace + 1]

    # 替换中文引号
    content = content.replace('"', '"').replace('"', '"')
    content = content.replace(''', "'").replace(''', "'")

    # 移除 BOM 标记
    if content.startswith('\ufeff'):
        content = content[1:]

    return content


def calculate_backoff(attempt: int, base_delay: float = 1.0, max_delay: float = 30.0) -> float:
    """计算指数退避延迟（带抖动）

    Args:
        attempt: 当前重试次数（从 0 开始）
        base_delay: 基础延迟（秒）
        max_delay: 最大延迟（秒）

    Returns:
        延迟时间（秒）
    """
    delay = min(base_delay * (2 ** attempt), max_delay)
    # 添加 50%-100% 的随机抖动
    jitter = delay * (0.5 + random.random())
    return min(jitter, max_delay)


def is_retryable_error(error: Exception) -> bool:
    """判断错误是否可重试

    不可重试的错误：
    - 认证错误（API Key 无效）
    - 请求格式错误
    - 权限不足
    """
    if isinstance(error, AuthenticationError):
        return False
    if isinstance(error, APIError):
        error_str = str(error).lower()
        non_retryable = ["invalid_api_key", "401", "403", "invalid_request"]
        if any(x in error_str for x in non_retryable):
            return False
    return True


def calculate_batches(page_count: int) -> List[int]:
    """计算分批策略

    根据页数决定如何分批生成，避免单次请求过大。

    Args:
        page_count: 总页数

    Returns:
        每批的页数列表
    """
    if page_count <= 35:
        return [page_count]
    elif page_count <= 70:
        half = page_count // 2
        return [half, page_count - half]
    elif page_count <= 100:
        third = page_count // 3
        return [third, third, page_count - 2 * third]
    elif page_count <= 150:
        return [50, 50, page_count - 100]
    else:
        # 最大支持 200 页
        page_count = min(page_count, 200)
        return [50, 50, 50, page_count - 150]


def build_batch_prompt_first(
    topic: str,
    audience: str,
    pages: int,
    total_batches: int,
    description: str
) -> str:
    """构建第一批的提示词"""
    prompt = f"""请为以下主题创作 PPT 的【开头部分】：

主题：{topic}
目标受众：{audience}
本批页数：{pages} 页（这是第 1/{total_batches} 批，后续还会继续生成）

⚠️ 重要说明：
- 这是分批生成的第一部分，请生成 PPT 的开头内容
- 包含：封面信息（title, subtitle）+ 前 {pages} 页内容
- 不要生成 ending 结束页（后续批次会生成）
- 内容要完整，为后续批次留好衔接"""

    if description:
        prompt += f"\n\n【参考资料】\n{description}"

    prompt += "\n\n请生成 JSON 格式，包含 title、subtitle 和 slides 数组。"
    return prompt


def build_batch_prompt_continue(
    topic: str,
    audience: str,
    pages: int,
    current_batch: int,
    total_batches: int,
    generated_summary: List[str],
    is_last: bool
) -> str:
    """构建续写批次的提示词"""
    # 只取最近 10 个标题作为上下文
    summary_text = "\n".join([f"- {t}" for t in generated_summary[-10:]])

    prompt = f"""请继续生成 PPT 的【第 {current_batch} 部分】：

主题：{topic}
目标受众：{audience}
本批页数：{pages} 页（这是第 {current_batch}/{total_batches} 批）

【已生成的内容摘要】（请续写，不要重复）：
{summary_text}

⚠️ 重要说明：
- 这是续写部分，请接着上面的内容继续
- 不要重复已生成的内容
- 本批生成 {pages} 页新内容"""

    if is_last:
        prompt += "\n- 这是最后一批，请在最后添加 ending 结束页"
    else:
        prompt += "\n- 不要生成 ending 结束页（后续批次会生成）"

    prompt += "\n\n请生成 JSON 格式，只需要 slides 数组（不需要 title 和 subtitle）。"
    return prompt


def build_json_error_message(error: json.JSONDecodeError, content: str) -> str:
    """构建 JSON 解析错误消息

    Args:
        error: JSON 解析错误
        content: 原始内容

    Returns:
        格式化的错误消息
    """
    # 显示错误位置附近的内容
    start = max(0, error.pos - 50)
    end = min(len(content), error.pos + 50)
    context = content[start:end]

    return (
        f"AI 返回的内容不是有效的 JSON 格式。\n"
        f"错误位置: 第 {error.lineno} 行，第 {error.colno} 列\n"
        f"错误内容: {error.msg}\n"
        f"上下文: ...{context}..."
    )
