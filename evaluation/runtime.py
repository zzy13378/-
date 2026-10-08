"""运行配置与模型客户端。统一从环境变量读取，避免多处硬编码。"""

from __future__ import annotations

import os
from dataclasses import dataclass


@dataclass
class RunConfig:
    model: str
    temperature: float = 0.2
    top_k: int = 3
    name: str = ""

    def label(self) -> str:
        return self.name or self.model


def default_config(top_k: int = 3) -> RunConfig:
    model = os.environ.get("LLM_MODEL", "deepseek-chat")
    return RunConfig(model=model, temperature=0.2, top_k=top_k, name=f"真实模型（{model}）")


def base_url() -> str:
    return os.environ.get("OPENAI_BASE_URL", "http://127.0.0.1:15721/v1")


def api_key() -> str:
    return os.environ.get("OPENAI_API_KEY", "any")


def client(timeout: float = 120.0):
    from openai import OpenAI

    return OpenAI(api_key=api_key(), base_url=base_url(), timeout=timeout, max_retries=0)


def model_ready() -> tuple[bool, str]:
    """检查模型配置是否完整，返回（是否可用，说明）。"""

    if not os.environ.get("OPENAI_BASE_URL"):
        return False, "缺少 OPENAI_BASE_URL"
    if not os.environ.get("OPENAI_API_KEY"):
        return False, "缺少 OPENAI_API_KEY"
    return True, f"{os.environ.get('LLM_MODEL', '(默认)')} @ {base_url()}"
