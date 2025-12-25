"""异步AI客户端 - 支持并发优化"""
import asyncio
import json
import time
import random
from typing import Dict, Any, Optional, Callable, List
from openai import AsyncOpenAI, APIError, APIConnectionError, RateLimitError, AuthenticationError

from config.settings import AIConfig, settings
from core.prompt_builder import get_system_prompt, build_user_prompt
from utils.logger import get_logger
from utils.cache import get_cache

logger = get_logger("ai_client_async")


class AsyncAIClient:
    """异步AI客户端，支持并发调用"""

    def __init__(self, config: AIConfig):
        self.config = config
        self.client = AsyncOpenAI(
            api_key=config.api_key,
            base_url=config.api_base_url,
            timeout=config.timeout
        )
        self.semaphore = asyncio.Semaphore(5)  # 限制并发数

    async def close(self):
        """关闭客户端"""
        await self.client.close()


def _clean_json_response(content: str) -> str:
    """清理 AI 返回的 JSON 内容（与原版保持一致）"""
    content = content.strip()

    if content.startswith("```json"):
        content = content[7:]
    elif content.startswith("```"):
        content = content[3:]
    if content.endswith("```"):
        content = content[:-3]
    content = content.strip()

    first_brace = content.find('{')
    last_brace = content.rfind('}')
    if first_brace != -1 and last_brace != -1 and first_brace < last_brace:
        content = content[first_brace:last_brace + 1]

    content = content.replace('"', '"').replace('"', '"')
    content = content.replace(''', "'").replace(''', "'")

    if content.startswith('\ufeff'):
        content = content[1:]

    return content


def _calculate_backoff(attempt: int, base_delay: float = 1.0, max_delay: float = 30.0) -> float:
    """计算指数退避延迟"""
    delay = min(base_delay * (2 ** attempt), max_delay)
    jitter = delay * (0.5 + random.random())
    return min(jitter, max_delay)


async def _call_api_with_retry_async(
    client: AsyncOpenAI,
    model_name: str,
    system_prompt: str,
    user_prompt: str,
    max_retries: int = 3,
    temperature: float = 0.7,
    semaphore: Optional[asyncio.Semaphore] = None
) -> str:
    """异步带重试机制的 API 调用"""
    if semaphore:
        async with semaphore:
            return await _call_api_impl(
                client, model_name, system_prompt, user_prompt,
                max_retries, temperature
            )
    else:
        return await _call_api_impl(
            client, model_name, system_prompt, user_prompt,
            max_retries, temperature
        )


async def _call_api_impl(
    client: AsyncOpenAI,
    model_name: str,
    system_prompt: str,
    user_prompt: str,
    max_retries: int,
    temperature: float
) -> str:
    """实际的API调用实现"""
    is_claude = "claude" in model_name.lower()
    last_error = None

    for attempt in range(max_retries):
        try:
            if is_claude:
                combined_prompt = f"{system_prompt}\n\n{user_prompt}"
                response = await client.chat.completions.create(
                    model=model_name,
                    messages=[{"role": "user", "content": combined_prompt}],
                    temperature=temperature,
                    max_tokens=8192
                )
            else:
                response = await client.chat.completions.create(
                    model=model_name,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    temperature=temperature,
                    max_tokens=8192,
                    response_format={"type": "json_object"}
                )

            content = None
            if isinstance(response, str):
                content = response
            elif hasattr(response, 'choices') and response.choices:
                message = response.choices[0].message
                content = message.content if message else None

            if not content:
                raise ValueError("AI 返回了空内容")

            return content

        except APIConnectionError as e:
            last_error = e
            if attempt < max_retries - 1:
                wait_time = _calculate_backoff(attempt, base_delay=1.0, max_delay=30.0)
                logger.warning(f"网络错误（第 {attempt + 1} 次），{wait_time:.1f} 秒后重试...")
                await asyncio.sleep(wait_time)
            else:
                raise Exception(f"网络连接失败: {e}")

        except RateLimitError as e:
            last_error = e
            if attempt < max_retries - 1:
                retry_after = getattr(e, 'retry_after', None)
                if retry_after:
                    wait_time = float(retry_after)
                else:
                    wait_time = _calculate_backoff(attempt, base_delay=2.0, max_delay=60.0)
                logger.warning(f"API 限流（第 {attempt + 1} 次），{wait_time:.1f} 秒后重试...")
                await asyncio.sleep(wait_time)
            else:
                raise Exception(f"API 限流: {e}")

        except Exception as e:
            if not _is_retryable_error(e) or attempt == max_retries - 1:
                raise
            last_error = e
            wait_time = _calculate_backoff(attempt, base_delay=1.0, max_delay=15.0)
            logger.warning(f"API 错误（第 {attempt + 1} 次），{wait_time:.1f} 秒后重试...")
            await asyncio.sleep(wait_time)

    raise Exception(f"API 调用失败: {last_error}")


def _is_retryable_error(error: Exception) -> bool:
    """判断错误是否可重试"""
    if isinstance(error, AuthenticationError):
        return False
    if isinstance(error, APIError):
        error_str = str(error).lower()
        if any(x in error_str for x in ["invalid_api_key", "401", "403", "invalid_request"]):
            return False
    return True


async def generate_ppt_plan_async(
    topic: str,
    audience: str,
    page_count: int = 0,
    description: str = "",
    auto_page_count: bool = False,
    config: Optional[AIConfig] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    use_cache: bool = True
) -> Dict[str, Any]:
    """异步生成 PPT 结构（支持并发优化）"""
    if config is None:
        config = settings.to_ai_config()

    config.validate()

    # 缓存检查
    if use_cache and not auto_page_count:
        cache = get_cache()
        cached = cache.get(topic, audience, page_count, description, config.model_name)
        if cached:
            return cached

    # 计算分批策略
    batches = _calculate_batches(page_count) if not auto_page_count and page_count > 35 else [page_count]
    total_batches = len(batches)

    if total_batches > 1:
        # 并发生成多个批次
        result = await _generate_ppt_plan_concurrent(
            topic, audience, batches, description, config, progress_callback
        )
    else:
        # 单批异步生成
        result = await _generate_ppt_plan_single_async(
            topic, audience, page_count, description, auto_page_count, config
        )

    # 保存缓存
    if use_cache and not auto_page_count:
        cache = get_cache()
        cache.set(topic, audience, page_count, result, description, config.model_name)

    return result


async def _generate_ppt_plan_concurrent(
    topic: str,
    audience: str,
    batches: List[int],
    description: str,
    config: AIConfig,
    progress_callback: Optional[Callable[[int, int, str], None]] = None
) -> Dict[str, Any]:
    """并发生成多个批次"""

    total_batches = len(batches)
    client = AsyncOpenAI(
        api_key=config.api_key,
        base_url=config.api_base_url,
        timeout=config.timeout
    )
    semaphore = asyncio.Semaphore(3)  # 限制并发批次数
    summary_lock = asyncio.Lock()  # 保护共享状态

    try:
        # 创建所有批次的任务
        tasks = []
        generated_summary = []

        for batch_idx, batch_pages in enumerate(batches):
            task = _generate_batch_async(
                client, topic, audience, batch_pages, batch_idx,
                total_batches, description, generated_summary,
                config, semaphore, progress_callback, summary_lock
            )
            tasks.append(task)

        # 并发执行所有批次（但按顺序处理结果）
        batch_results = []
        for i, coro in enumerate(asyncio.as_completed(tasks)):
            batch_idx, batch_result = await coro
            batch_results.append((batch_idx, batch_result))

            # 更新生成的摘要列表（加锁保护）
            async with summary_lock:
                slides = batch_result.get("slides", [])
                for slide in slides:
                    slide_title = slide.get("title", "")
                    if slide_title:
                        generated_summary.append(slide_title)

        # 按批次顺序合并结果
        batch_results.sort(key=lambda x: x[0])

        all_slides = []
        title = ""
        subtitle = ""

        for batch_idx, batch_result in batch_results:
            if batch_idx == 0:
                title = batch_result.get("title", topic)
                subtitle = batch_result.get("subtitle", "")

            # 收集slides
            batch_slides = batch_result.get("slides", [])

            # 过滤掉ending页（除了最后一批）
            if batch_idx < len(batches) - 1:
                batch_slides = [s for s in batch_slides if s.get("type") != "ending"]

            all_slides.extend(batch_slides)

        # 合并结果
        result = {
            "title": title,
            "subtitle": subtitle,
            "slides": all_slides
        }

        logger.info(f"并发生成完成，共 {len(all_slides)} 页")
        return result

    finally:
        await client.close()


async def _generate_batch_async(
    client: AsyncOpenAI,
    topic: str,
    audience: str,
    pages: int,
    batch_idx: int,
    total_batches: int,
    description: str,
    generated_summary: List[str],
    config: AIConfig,
    semaphore: asyncio.Semaphore,
    progress_callback: Optional[Callable[[int, int, str], None]],
    summary_lock: asyncio.Lock = None
) -> tuple:
    """异步生成单个批次"""

    async with semaphore:
        current_batch = batch_idx + 1

        if progress_callback:
            progress_callback(current_batch, total_batches, f"正在生成第 {current_batch}/{total_batches} 批...")

        logger.info(f"开始生成第 {current_batch}/{total_batches} 批（{pages} 页）...")

        # 构建提示词
        if batch_idx == 0:
            batch_prompt = _build_batch_prompt_first(
                topic, audience, pages, total_batches, description
            )
        else:
            batch_prompt = _build_batch_prompt_continue(
                topic, audience, pages, current_batch, total_batches,
                generated_summary, is_last=(current_batch == total_batches)
            )

        system_prompt = get_system_prompt()

        content = await _call_api_with_retry_async(
            client=client,
            model_name=config.model_name,
            system_prompt=system_prompt,
            user_prompt=batch_prompt,
            max_retries=config.max_retries,
            temperature=config.temperature,
            semaphore=None  # 内部已经用semaphore限制
        )

        cleaned_content = _clean_json_response(content)
        batch_result = json.loads(cleaned_content)

        logger.info(f"第 {current_batch} 批完成")
        return batch_idx, batch_result


async def _generate_ppt_plan_single_async(
    topic: str,
    audience: str,
    page_count: int,
    description: str,
    auto_page_count: bool,
    config: AIConfig
) -> Dict[str, Any]:
    """单批异步生成"""
    client = AsyncOpenAI(
        api_key=config.api_key,
        base_url=config.api_base_url,
        timeout=config.timeout
    )

    try:
        system_prompt = get_system_prompt()
        user_prompt = build_user_prompt(topic, audience, page_count, description, auto_page_count)

        logger.info(f"异步生成 PPT: {topic} | 受众: {audience} | 模型: {config.model_name}")

        content = await _call_api_with_retry_async(
            client=client,
            model_name=config.model_name,
            system_prompt=system_prompt,
            user_prompt=user_prompt,
            max_retries=config.max_retries,
            temperature=config.temperature
        )

        cleaned_content = _clean_json_response(content)
        plan_dict = json.loads(cleaned_content)

        return plan_dict

    finally:
        await client.close()


def _calculate_batches(page_count: int) -> List[int]:
    """计算分批策略（与原版保持一致）"""
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
        page_count = min(page_count, 200)
        return [50, 50, 50, page_count - 150]


def _build_batch_prompt_first(topic: str, audience: str, pages: int, total_batches: int, description: str) -> str:
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


def _build_batch_prompt_continue(topic: str, audience: str, pages: int,
                                current_batch: int, total_batches: int,
                                generated_summary: List[str], is_last: bool) -> str:
    """构建续写批次的提示词"""
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