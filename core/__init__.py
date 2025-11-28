"""核心模块"""
from core.ai_client import (
    generate_ppt_plan,
    AIClientError,
    APIKeyError,
    RateLimitExceeded,
    JSONParseError,
    NetworkError,
)
from core.ppt_plan import PptPlan, Slide, ppt_plan_from_dict, ppt_plan_to_dict
from core.prompt_builder import get_system_prompt, build_user_prompt

__all__ = [
    "generate_ppt_plan",
    "AIClientError",
    "APIKeyError", 
    "RateLimitExceeded",
    "JSONParseError",
    "NetworkError",
    "PptPlan",
    "Slide",
    "ppt_plan_from_dict",
    "ppt_plan_to_dict",
    "get_system_prompt",
    "build_user_prompt",
]
