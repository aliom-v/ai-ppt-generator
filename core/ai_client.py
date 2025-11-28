"""大模型调用封装模块 - 优化版"""
import json
import time
from typing import Dict, Any, Optional
from openai import OpenAI, APIError, APIConnectionError, RateLimitError

from config.settings import AIConfig, settings
from core.prompt_builder import get_system_prompt, build_user_prompt


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


def _clean_json_response(content: str) -> str:
    """清理 AI 返回的 JSON 内容"""
    content = content.strip()
    
    # 移除 markdown 代码块
    if content.startswith("```json"):
        content = content[7:]
    elif content.startswith("```"):
        content = content[3:]
    if content.endswith("```"):
        content = content[:-3]
    content = content.strip()
    
    # 提取 JSON 部分
    first_brace = content.find('{')
    last_brace = content.rfind('}')
    if first_brace != -1 and last_brace != -1 and first_brace < last_brace:
        content = content[first_brace:last_brace + 1]
    
    # 替换中文引号
    content = content.replace('"', '"').replace('"', '"')
    content = content.replace(''', "'").replace(''', "'")
    
    # 移除 BOM
    if content.startswith('\ufeff'):
        content = content[1:]
    
    return content


def _call_api_with_retry(
    client: OpenAI,
    model_name: str,
    system_prompt: str,
    user_prompt: str,
    max_retries: int = 3,
    temperature: float = 0.7
) -> str:
    """带重试机制的 API 调用"""
    is_claude = "claude" in model_name.lower()
    last_error = None
    
    for attempt in range(max_retries):
        try:
            if is_claude:
                # Claude 模型：合并 system 和 user prompt
                combined_prompt = f"{system_prompt}\n\n{user_prompt}"
                response = client.chat.completions.create(
                    model=model_name,
                    messages=[{"role": "user", "content": combined_prompt}],
                    temperature=temperature,
                    max_tokens=8192
                )
            else:
                # OpenAI 模型：使用标准格式
                response = client.chat.completions.create(
                    model=model_name,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    temperature=temperature,
                    max_tokens=8192,
                    response_format={"type": "json_object"}
                )
            
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
            
        except RateLimitError as e:
            last_error = e
            if attempt < max_retries - 1:
                wait_time = (2 ** attempt) * 2  # 指数退避: 2, 4, 8 秒
                print(f"⚠️ API 限流，{wait_time} 秒后重试...")
                time.sleep(wait_time)
            else:
                raise RateLimitExceeded(f"API 限流，已重试 {max_retries} 次: {e}")
                
        except APIConnectionError as e:
            last_error = e
            if attempt < max_retries - 1:
                wait_time = (2 ** attempt)
                print(f"⚠️ 网络错误，{wait_time} 秒后重试...")
                time.sleep(wait_time)
            else:
                raise NetworkError(f"网络连接失败，已重试 {max_retries} 次: {e}")
                
        except APIError as e:
            if "invalid_api_key" in str(e).lower() or "401" in str(e):
                raise APIKeyError(f"API Key 无效: {e}")
            last_error = e
            if attempt < max_retries - 1:
                time.sleep(1)
            else:
                raise AIClientError(f"API 调用失败: {e}")
    
    raise AIClientError(f"API 调用失败: {last_error}")


def _calculate_batches(page_count: int) -> list:
    """计算分批策略
    
    规则：
    - 35页及以下：1批
    - 36-70页：2批
    - 71-100页：3批
    - 101-150页：3批（每批约50页）
    - 151-200页：4批（每批约50页）
    """
    if page_count <= 35:
        return [page_count]
    elif page_count <= 70:
        # 2批，尽量均分
        half = page_count // 2
        return [half, page_count - half]
    elif page_count <= 100:
        # 3批
        third = page_count // 3
        return [third, third, page_count - 2 * third]
    elif page_count <= 150:
        # 3批，每批约50页
        return [50, 50, page_count - 100]
    else:
        # 4批，每批约50页，最多200页
        page_count = min(page_count, 200)
        return [50, 50, 50, page_count - 150]


def generate_ppt_plan(
    topic: str,
    audience: str,
    page_count: int = 0,
    description: str = "",
    auto_page_count: bool = False,
    config: Optional[AIConfig] = None,
    progress_callback: callable = None
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
    
    # 计算是否需要分批
    batches = _calculate_batches(page_count) if not auto_page_count and page_count > 35 else [page_count]
    total_batches = len(batches)
    
    if total_batches > 1:
        print(f"\n📦 页数较多（{page_count}页），将分 {total_batches} 批生成...")
        return _generate_ppt_plan_batched(
            topic, audience, batches, description, config, progress_callback
        )
    
    # 单批生成
    return _generate_ppt_plan_single(
        topic, audience, page_count, description, auto_page_count, config
    )


def _generate_ppt_plan_batched(
    topic: str,
    audience: str,
    batches: list,
    description: str,
    config: AIConfig,
    progress_callback: callable = None
) -> Dict[str, Any]:
    """分批生成 PPT 结构"""
    from core.prompt_builder import get_system_prompt
    
    client = OpenAI(
        api_key=config.api_key,
        base_url=config.api_base_url,
        timeout=config.timeout
    )
    
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
        
        print(f"\n🔄 生成第 {current_batch}/{total_batches} 批（{batch_pages} 页）...")
        
        # 构建分批提示词
        if batch_idx == 0:
            # 第一批：生成开头部分
            batch_prompt = _build_batch_prompt_first(
                topic, audience, batch_pages, total_batches, description
            )
        else:
            # 后续批次：续写
            batch_prompt = _build_batch_prompt_continue(
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
            
            cleaned_content = _clean_json_response(content)
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
            
            print(f"✓ 第 {current_batch} 批完成，获得 {len(batch_slides)} 页")
            
        except Exception as e:
            print(f"⚠️ 第 {current_batch} 批生成失败: {e}")
            raise
    
    # 合并结果
    result = {
        "title": title,
        "subtitle": subtitle,
        "slides": all_slides
    }
    
    print(f"\n✅ 分批生成完成，共 {len(all_slides)} 页")
    return result


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
    
    prompt += """

请生成 JSON 格式，包含 title、subtitle 和 slides 数组。"""
    return prompt


def _build_batch_prompt_continue(topic: str, audience: str, pages: int, 
                                  current_batch: int, total_batches: int,
                                  generated_summary: list, is_last: bool) -> str:
    """构建续写批次的提示词"""
    summary_text = "\n".join([f"- {t}" for t in generated_summary[-10:]])  # 最近10页摘要
    
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
    
    prompt += """

请生成 JSON 格式，只需要 slides 数组（不需要 title 和 subtitle）。"""
    return prompt


def _generate_ppt_plan_single(
    topic: str,
    audience: str,
    page_count: int,
    description: str,
    auto_page_count: bool,
    config: AIConfig
) -> Dict[str, Any]:
    """单批生成 PPT 结构（原有逻辑）"""
    client = OpenAI(
        api_key=config.api_key,
        base_url=config.api_base_url,
        timeout=config.timeout
    )
    
    system_prompt = get_system_prompt()
    user_prompt = build_user_prompt(topic, audience, page_count, description, auto_page_count)
    
    print(f"\n{'=' * 60}")
    print(f"📝 生成 PPT: {topic}")
    print(f"🎯 目标受众: {audience}")
    print(f"🤖 使用模型: {config.model_name}")
    print(f"{'=' * 60}\n")
    
    try:
        content = _call_api_with_retry(
            client=client,
            model_name=config.model_name,
            system_prompt=system_prompt,
            user_prompt=user_prompt,
            max_retries=config.max_retries,
            temperature=config.temperature
        )
        
        print(f"📄 AI 返回内容长度: {len(content)} 字符")
        
        content_lower = content.strip().lower()
        if content_lower.startswith('<!doctype') or content_lower.startswith('<html') or '<html' in content_lower[:500]:
            raise AIClientError(
                f"API 返回了 HTML 页面而不是 AI 响应。请检查：\n"
                f"1. API Base URL 是否正确（当前: {config.api_base_url}）\n"
                f"2. 确保 URL 以 /v1 结尾\n"
                f"3. API Key 是否有效"
            )
        
        cleaned_content = _clean_json_response(content)
        
        if not cleaned_content:
            raise JSONParseError(f"AI 返回了无效内容: {content[:300]}")
        
        plan_dict = json.loads(cleaned_content)
        
        return plan_dict
        
    except json.JSONDecodeError as e:
        error_msg = _build_json_error_message(e, locals().get('content', ''))
        raise JSONParseError(error_msg)
    except AIClientError:
        raise
    except Exception as e:
        raise AIClientError(f"生成失败: {e}")


def _build_json_error_message(error: json.JSONDecodeError, content: str) -> str:
    """构建 JSON 解析错误消息"""
    msg = "AI 返回的内容不是有效的 JSON 格式。"
    msg += f"  错误详情: {error}"
    
    if content:
        preview = content[:200].replace('\n', ' ')
        msg += f"  返回内容预览: {preview}"
    
    return msg


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
    
    import time
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
