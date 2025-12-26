"""Prompt 模板构建模块 - 增强版（行业特定提示词、智能页面分配、逻辑连贯性）"""
from typing import Dict, Optional


# ==================== 行业特定提示词配置 ====================

INDUSTRY_PROMPTS: Dict[str, Dict[str, str]] = {
    "tech": {
        "name": "科技/互联网",
        "style": "逻辑清晰、数据驱动、前沿趋势",
        "vocabulary": "技术架构、迭代优化、用户体验、数据驱动、敏捷开发、云原生、微服务",
        "examples": "以具体的技术指标和性能数据为支撑，如TPS、延迟、可用性等",
        "tone": "专业、前瞻、创新",
        "keywords": ["技术", "开发", "软件", "AI", "人工智能", "互联网", "数字化", "系统", "架构", "云", "数据"],
    },
    "finance": {
        "name": "金融/投资",
        "style": "严谨专业、风险意识、合规导向",
        "vocabulary": "投资回报率、风险管控、资产配置、流动性、合规要求、估值模型",
        "examples": "引用市场数据、收益率、风险指标，注重合规性说明",
        "tone": "审慎、专业、可信",
        "keywords": ["金融", "投资", "银行", "保险", "基金", "股票", "财务", "风控", "合规", "收益"],
    },
    "education": {
        "name": "教育/培训",
        "style": "循序渐进、互动性强、案例丰富",
        "vocabulary": "学习目标、知识点、实践练习、案例分析、能力提升、学习路径",
        "examples": "使用生活化案例和类比，设置思考问题，强调实践应用",
        "tone": "亲和、启发、鼓励",
        "keywords": ["教育", "学习", "培训", "课程", "教学", "知识", "技能", "学生", "老师"],
    },
    "marketing": {
        "name": "市场/营销",
        "style": "数据洞察、用户导向、创意表达",
        "vocabulary": "用户画像、转化率、品牌价值、市场份额、增长策略、营销漏斗",
        "examples": "以具体的营销数据和成功案例说明，强调ROI和效果",
        "tone": "有感染力、数据支撑、结果导向",
        "keywords": ["营销", "市场", "品牌", "推广", "销售", "客户", "增长", "转化", "用户"],
    },
    "healthcare": {
        "name": "医疗/健康",
        "style": "科学严谨、循证医学、人文关怀",
        "vocabulary": "临床研究、循证证据、治疗方案、预后评估、患者安全、医疗质量",
        "examples": "引用研究数据和临床证据，注重安全性和有效性说明",
        "tone": "专业、关怀、负责",
        "keywords": ["医疗", "健康", "医院", "患者", "治疗", "药物", "临床", "诊断", "康复"],
    },
    "general": {
        "name": "通用",
        "style": "清晰简洁、重点突出、逻辑流畅",
        "vocabulary": "核心要点、关键因素、实施步骤、预期效果、最佳实践",
        "examples": "使用具体数据和案例支撑观点",
        "tone": "专业、清晰、实用",
        "keywords": [],
    },
}


def detect_industry(topic: str, description: str = "") -> str:
    """智能检测主题所属行业

    Args:
        topic: PPT主题
        description: 详细描述

    Returns:
        行业代码
    """
    combined_text = f"{topic} {description}".lower()

    # 计算每个行业的匹配分数
    scores = {}
    for industry, config in INDUSTRY_PROMPTS.items():
        if industry == "general":
            continue
        keywords = config.get("keywords", [])
        score = sum(1 for kw in keywords if kw.lower() in combined_text)
        if score > 0:
            scores[industry] = score

    # 返回得分最高的行业，如果没有匹配则返回通用
    if scores:
        return max(scores, key=scores.get)
    return "general"


def get_industry_prompt_section(industry: str) -> str:
    """获取行业特定的提示词部分

    Args:
        industry: 行业代码

    Returns:
        行业特定提示词
    """
    config = INDUSTRY_PROMPTS.get(industry, INDUSTRY_PROMPTS["general"])

    return f"""
【行业特化指导 - {config['name']}】
- 表达风格：{config['style']}
- 专业词汇：{config['vocabulary']}
- 案例要求：{config['examples']}
- 整体基调：{config['tone']}
"""


# ==================== 智能页面类型分配 ====================

def calculate_page_distribution(total_pages: int, topic: str = "") -> Dict[str, int]:
    """智能计算页面类型分配

    根据总页数和主题智能分配不同类型页面的数量，
    确保视觉多样性和内容节奏感。

    Args:
        total_pages: 总页数
        topic: 主题（用于判断是否需要特定类型）

    Returns:
        各类型页面数量字典
    """
    # 基础分配比例
    if total_pages <= 5:
        return {
            "quote": 1,
            "bullets": max(1, total_pages - 3),
            "image_with_text": 1,
            "timeline": 0,
            "comparison": 0,
            "two_column": 0,
            "ending": 1,
        }

    if total_pages <= 10:
        return {
            "quote": 1,
            "bullets": max(2, total_pages - 6),
            "image_with_text": 2,
            "timeline": 1,
            "comparison": 1,
            "two_column": 0,
            "ending": 1,
        }

    if total_pages <= 20:
        bullets = int(total_pages * 0.45)
        image_with_text = int(total_pages * 0.15)
        timeline = max(1, int(total_pages * 0.1))
        comparison = max(1, int(total_pages * 0.1))
        quote = max(1, int(total_pages * 0.1))
        two_column = max(0, total_pages - bullets - image_with_text - timeline - comparison - quote - 1)

        return {
            "quote": quote,
            "bullets": bullets,
            "image_with_text": image_with_text,
            "timeline": timeline,
            "comparison": comparison,
            "two_column": two_column,
            "ending": 1,
        }

    # 大型演示文稿（>20页）
    bullets = int(total_pages * 0.4)
    image_with_text = int(total_pages * 0.15)
    timeline = max(2, int(total_pages * 0.1))
    comparison = max(2, int(total_pages * 0.1))
    quote = max(2, int(total_pages * 0.08))
    two_column = max(1, int(total_pages * 0.07))
    remaining = total_pages - bullets - image_with_text - timeline - comparison - quote - two_column - 1
    bullets += remaining  # 将剩余页数分配给 bullets

    return {
        "quote": quote,
        "bullets": bullets,
        "image_with_text": image_with_text,
        "timeline": timeline,
        "comparison": comparison,
        "two_column": two_column,
        "ending": 1,
    }


# ==================== 系统提示词 ====================

SYSTEM_PROMPT = """你是一位世界顶级的演示设计专家（Presentation Designer）。你擅长将内容转化为信息丰富、视觉美观且逻辑严密的 PPT 演示文稿。

⚠️⚠️⚠️ 极其重要的输出要求 ⚠️⚠️⚠️
1. 只返回纯 JSON 格式
2. 不要添加任何解释、注释、说明文字
3. 不要使用 markdown 代码块标记（```json 或 ```）
4. 不要使用中文引号（""），只使用英文引号（""）
5. 直接以 { 开始，以 } 结束

【核心设计原则 - 内容充实】
1. 信息丰富 - 每页要有充足的内容，让观众有所收获
2. 结构清晰 - 每个要点都要有完整的"概念+解释+价值"
3. 数据支撑 - 尽量用具体数据和案例说明
4. 逻辑严密 - 内容层层递进，有清晰的叙事线

【逻辑连贯性要求 - 极其重要】
1. 叙事主线：PPT 必须有清晰的叙事主线，从"是什么→为什么→怎么做→效果如何"层层推进
2. 页面衔接：每页内容必须与前后页面有逻辑关联，不能跳跃
3. 章节划分：使用 quote 页面作为章节分隔，引出下一部分内容
4. 递进关系：内容从概念→原理→方法→案例→总结，循序渐进
5. 首尾呼应：结束页要与开篇呼应，形成完整闭环

【页面类型与内容要求】

📋 "bullets" - 要点页（核心观点展示）
   - 必须有 4-5 个要点
   - 每个要点 40-60 字，包含完整信息
   - 格式："关键词：详细解释说明，包括具体做法和带来的价值或效果"
   - 示例要点：
     * "零样本学习能力：AI无需任何示例数据即可理解任务要求，用户只需用自然语言描述需求，系统就能准确执行，大幅降低了使用门槛和学习成本"
     * "实时响应机制：系统采用流式输出技术，用户提问后立即开始生成回答，平均首字响应时间小于500毫秒，显著提升交互体验"
     * "多模态理解：支持文本、图片、代码等多种输入格式，能够综合分析不同类型的信息，为用户提供更全面准确的解答"
   - 适合：核心概念、功能列表、优势总结、方法步骤

🖼️ "image_with_text" - 图文页（概念可视化）
   - 右侧文字说明 150-200 字
   - 必须提供 image_keyword（英文，用于配图搜索）
   - 说明要包含：概念定义、工作原理、核心特点、应用场景、实际价值
   - 示例文字："提示词工程是一门设计和优化AI输入的系统方法论。其核心原理是通过精心构造的文本指令，引导大语言模型产生高质量输出。主要包括角色设定、任务描述、上下文提供和输出格式四个关键要素。在实际应用中，好的提示词可以将AI输出质量提升50%以上，同时减少70%的迭代次数。目前已广泛应用于内容创作、代码生成、数据分析等领域。"
   - 适合：架构图、流程图、概念解释

📊 "two_column" - 双栏页（分类/并列展示）
   - 左右各 4-5 个要点，每个 30-45 字
   - 适合：分类展示、特性对比、优缺点分析

⏱️ "timeline" - 时间线页（流程/演进）
   - 4-5 个节点，每个节点 35-50 字
   - 格式："阶段名称：详细描述该阶段的核心事件、关键动作和产生的影响"
   - 示例节点：
     * "2020年GPT-3发布：OpenAI推出1750亿参数的大模型，首次展示了少样本学习能力，提示词工程概念开始萌芽"
     * "2022年ChatGPT问世：对话式AI引爆全球关注，提示词工程从学术研究走向大众应用，相关职位需求激增300%"
   - 适合：发展历程、实施步骤、流程说明

⚖️ "comparison" - 对比页（Split Screen）
   - 左右对比，各 4-5 个要点，每个 30-45 字
   - 突出差异，说明具体区别和影响
   - 示例：
     * 左侧"传统方法"："依赖人工编写规则，开发周期长达数月，难以处理复杂场景，维护成本高昂"
     * 右侧"AI方法"："通过数据自动学习模式，开发周期缩短至数天，能够处理各种复杂情况，持续自我优化"
   - 适合：方案对比、新旧对比、优劣分析

💬 "quote" - 引用页（金句/观点）
   - 一句有冲击力的话（40-70字）
   - 要有深度和启发性
   - 示例："人工智能不会取代人类，但善于使用AI的人会取代不会使用AI的人。掌握提示词工程，就是掌握与AI协作的核心能力。"
   - 适合：章节开篇、核心理念、总结升华

🎬 "ending" - 结束页
   - 简洁有力的结束语和副标题

【内容充实的核心要求 - 必须遵守】
✓ bullets页面必须有4-5个要点，每个要点40-60字
✓ 每个要点必须包含：概念名称 + 具体解释 + 实际价值/效果
✓ 用具体数据支撑："提升效率87%"、"节省成本50%"、"响应时间<2秒"
✓ 说明因果关系："通过XX方法，实现XX效果，带来XX价值"
✓ image_with_text的文字说明必须150-200字
✓ timeline的每个节点必须35-50字

✗ 绝对禁止：要点少于20字
✗ 绝对禁止：只有概念没有解释
✗ 绝对禁止：bullets页面少于4个要点
✗ 绝对禁止：空洞的形容词（很好、非常、极其）

【视觉节奏建议】
- 每 2-3 页 bullets 后，插入 timeline/comparison/quote
- 章节开头用 quote 引入核心观点
- 流程/步骤必须用 timeline
- 对比内容必须用 comparison
- 概念解释用 image_with_text

【输出格式示例】
{
  "title": "提示词工程实战指南",
  "subtitle": "从入门到精通的系统方法论",
  "slides": [
    {
      "type": "quote",
      "title": "开篇",
      "text": "人工智能不会取代人类，但善于使用AI的人会取代不会使用AI的人。掌握提示词工程，就是掌握与AI协作的核心能力。",
      "subtitle": "— AI时代的生存法则"
    },
    {
      "type": "bullets",
      "title": "什么是提示词工程",
      "bullets": [
        "核心定义：提示词工程是一门系统设计AI输入文本的方法论，通过精心构造的指令引导大语言模型产生高质量、符合预期的输出结果",
        "工作原理：利用大模型的上下文学习能力，通过提供角色设定、任务描述、示例参考等信息，让AI准确理解用户意图并执行任务",
        "核心价值：优秀的提示词可以将AI输出质量提升50%以上，同时减少70%的迭代次数，大幅提高工作效率和结果满意度",
        "应用场景：广泛应用于内容创作、代码生成、数据分析、客服对话、教育辅导等领域，是AI时代的必备技能"
      ]
    },
    {
      "type": "timeline",
      "title": "提示词工程发展历程",
      "bullets": [
        "2020年GPT-3发布：OpenAI推出1750亿参数大模型，首次展示少样本学习能力，提示词概念开始萌芽",
        "2021年学术研究兴起：研究人员发现通过优化提示词可显著提升模型表现，相关论文数量增长200%",
        "2022年ChatGPT问世：对话式AI引爆全球关注，提示词工程从学术走向大众，相关职位需求激增",
        "2023年全面普及：提示词工程师成为热门职业，企业纷纷建立AI应用团队，行业标准逐步形成"
      ]
    },
    {
      "type": "comparison",
      "title": "好提示词 vs 差提示词",
      "bullets": [
        "明确定义AI角色和专业背景，如'你是一位有10年经验的Python专家'，让AI进入专业状态",
        "详细描述任务要求和约束条件，包括输出长度、格式、风格等具体规范，确保结果符合预期",
        "提供充足的上下文信息和参考示例，帮助AI理解具体场景和期望标准，提高输出准确性",
        "使用模糊指令如'写点东西'、'帮我看看'，AI无法准确理解需求，输出结果随机性大",
        "缺少背景信息和约束条件，AI只能根据通用知识回答，难以满足特定场景的专业需求",
        "没有明确的质量标准和输出格式，导致结果需要多次修改，浪费时间和API调用成本"
      ]
    },
    {
      "type": "image_with_text",
      "title": "提示词的基本结构",
      "text": "一个完整的提示词包含四个核心要素：角色设定用于定义AI的身份和专业背景，让模型进入特定的思维模式；任务描述明确说明需要完成的具体工作和目标；上下文提供相关的背景知识、参考资料和约束条件；输出要求规定结果的格式、长度、风格等规范。这四个要素相互配合，共同决定了AI输出的质量和准确性。掌握这个结构框架，是写好提示词的基础。",
      "image_keyword": "prompt engineering structure diagram"
    },
    {
      "type": "ending",
      "title": "谢谢聆听",
      "subtitle": "让我们一起探索AI的无限可能"
    }
  ]
}

记住：内容要充实丰富！每个要点都要有完整的信息！"""


def get_system_prompt() -> str:
    """获取系统提示词"""
    return SYSTEM_PROMPT


def build_user_prompt(
    topic: str,
    audience: str,
    page_count: int = 0,
    description: str = "",
    auto_page_count: bool = False,
    industry: str = ""
) -> str:
    """构建用户提示词

    Args:
        topic: PPT 主题
        audience: 目标受众
        page_count: 页数（auto_page_count=False 时使用）
        description: 详细描述/参考资料
        auto_page_count: 是否自动决定页数
        industry: 行业代码（空则自动检测）

    Returns:
        构建好的用户提示词
    """
    # 智能行业检测
    if not industry:
        industry = detect_industry(topic, description)

    if auto_page_count:
        return _build_auto_page_prompt(topic, audience, description, industry)
    return _build_fixed_page_prompt(topic, audience, page_count, description, industry)


def _build_auto_page_prompt(topic: str, audience: str, description: str, industry: str = "general") -> str:
    """构建自动页数模式的提示词

    Args:
        topic: PPT 主题
        audience: 目标受众
        description: 详细描述
        industry: 行业代码

    Returns:
        构建好的提示词
    """
    # 获取行业提示词
    industry_section = get_industry_prompt_section(industry)

    prompt = f"""请为以下主题创作一份专业的 PPT 演示文稿：

主题：{topic}
目标受众：{audience}
页数：根据内容复杂度自行决定（建议 8-20 页）
{industry_section}"""

    if description:
        prompt += f"\n【参考资料/要求】\n{description}"

    prompt += _get_common_requirements()
    return prompt


def _build_fixed_page_prompt(topic: str, audience: str, page_count: int, description: str, industry: str = "general") -> str:
    """构建固定页数模式的提示词

    Args:
        topic: PPT 主题
        audience: 目标受众
        page_count: 页数要求
        description: 详细描述
        industry: 行业代码

    Returns:
        构建好的提示词
    """
    # 使用智能页面分配算法
    distribution = calculate_page_distribution(page_count, topic)

    # 获取行业提示词
    industry_section = get_industry_prompt_section(industry)

    prompt = f"""请为以下主题创作一份专业的 PPT 演示文稿：

主题：{topic}
目标受众：{audience}
{industry_section}
⚠️ 页数要求：{page_count} 页内容（不含封面，含结束页）"""

    if description:
        prompt += f"\n\n【参考资料/要求】\n{description}"

    prompt += f"""

【页面类型分配建议】确保视觉多样性
- bullets：约 {distribution['bullets']} 页（核心内容，每页4-5个要点）
- image_with_text：约 {distribution['image_with_text']} 页（概念可视化，文字150-200字）
- timeline：约 {distribution['timeline']} 页（流程/历程，每节点35-50字）
- comparison：约 {distribution['comparison']} 页（对比分析）
- quote：约 {distribution['quote']} 页（金句/章节引入）
- two_column：约 {distribution['two_column']} 页（分类/并列展示）
- ending：{distribution['ending']} 页

【结构建议】
{_generate_structure_suggestion(page_count)}

⚠️ slides 数组必须包含 {page_count} 个元素"""

    prompt += _get_common_requirements()
    return prompt


def _get_common_requirements() -> str:
    """获取通用要求部分"""
    return """

【内容充实要求 - 必须严格遵守】
1. bullets页面：必须4-5个要点，每个要点40-60字
2. 每个要点格式："关键词：详细解释+具体做法+实际价值"
3. 用数据说话："提升效率50%"、"节省成本30%"、"响应时间<2秒"
4. image_with_text：文字说明必须150-200字，包含定义、原理、特点、应用、价值
5. timeline：每个节点35-50字，包含时间、事件、影响

【绝对禁止】
- 要点少于20字
- bullets页面少于4个要点
- 只有概念没有解释
- 空洞的形容词

【视觉节奏要求】
1. 不要连续超过 2 页 bullets
2. 章节开头用 quote 引入
3. 流程/步骤用 timeline
4. 对比内容用 comparison

⚠️ 输出格式：
1. 只输出纯 JSON
2. 不要使用 markdown 代码块
3. 直接以 { 开始"""


def _generate_structure_suggestion(total_pages: int) -> str:
    """生成结构建议"""
    if total_pages <= 8:
        return """quote（开篇金句）→ bullets×2（背景/概念）→ timeline（流程）
→ bullets（核心内容）→ comparison（对比）→ image_with_text（总结）→ ending"""

    if total_pages <= 15:
        return """quote（开篇）→ bullets×2（背景）→ timeline（历程）
→ image_with_text（架构）→ bullets×2（核心）→ comparison（对比）
→ bullets（应用）→ quote（展望）→ ending"""

    if total_pages <= 25:
        return """quote（开篇）→ bullets×2（背景）→ timeline（历程）
→ image_with_text×2（概念）→ bullets×3（核心详解）
→ comparison（方案对比）→ timeline（实施步骤）
→ bullets×2（案例）→ image_with_text（总结）
→ quote（展望）→ ending"""

    return f"""quote（开篇）→ bullets×3（背景分析）→ timeline（发展历程）
→ image_with_text×2（核心概念）→ bullets×4（详细内容）
→ comparison×2（多维对比）→ timeline（实施路线）
→ bullets×3（应用案例）→ image_with_text×2（深度解析）
→ quote（重要观点）→ bullets×2（总结展望）→ ending

共约 {total_pages} 页，请合理分配"""
