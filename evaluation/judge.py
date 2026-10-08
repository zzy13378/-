"""真实模式的判定：答案正确性判定与句子级引用核实。

两个判定都以真实证据为准，不使用人工编写的期望关键词。
"""

from __future__ import annotations

import json
import time
from typing import Optional, Tuple

from . import runtime

GROUNDED_PROMPT = """你是评测复核员。下面给出一个问题、这段问题的标准证据、被测系统实际拿到的证据，以及它的回答。
请判断这个回答是否正确回答了问题。判断标准：
1. 以标准证据为准，回答是否给出了与之一致的正确内容；
2. 如果标准证据里明确包含答案，而被测系统回答“证据不足”或拒答，属于没有答对，应判失败；
3. 如果被测系统给出的内容与标准证据冲突，或引入了标准证据中不存在的事实，判失败；
4. 表达方式与标准证据不同不算错，只要事实一致即可。
只输出 JSON：{"pass":true,"reason":"一句话理由"}"""

SUPPORT_PROMPT = """你是证据核对员。下面给出一个论断句、它引用的证据，以及原论文的相关证据。
请判断该论断句是否被引用的证据支持。注意：中文句子引用英文证据时，用词不同但事实一致也算支持。
只输出 JSON：{"supported":true,"reason":"一句话理由"}"""


class Judge:
    """统一的模型判定入口，记录 token 消耗便于统计成本。"""

    def __init__(self, model: str = "") -> None:
        self.client = runtime.client()
        self.model = model or runtime.default_config().model
        self.tokens = 0

    def _call(self, system: str, user: str, max_tokens: int = 300) -> Optional[dict]:
        for _ in range(3):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
                    temperature=0.0,
                    max_tokens=max_tokens,
                )
                content = (response.choices[0].message.content or "").strip()
                self.tokens += int(getattr(response.usage, "total_tokens", 0) or 0)
                start, end = content.find("{"), content.rfind("}")
                if start >= 0 and end > start:
                    return json.loads(content[start:end + 1])
            except Exception:  # noqa: BLE001
                time.sleep(2)
        return None

    def answer_correct(self, question: str, gold_text: str, retrieved_text: str,
                       answer: str) -> Tuple[Optional[bool], str]:
        """对照标准证据判断回答是否正确。"""

        payload = self._call(GROUNDED_PROMPT, (
            f"问题：{question}\n\n"
            f"标准证据：\n{gold_text or '（无）'}\n\n"
            f"被测系统实际拿到的证据：\n{retrieved_text or '（无检索结果）'}\n\n"
            f"回答：{answer or '（空）'}"
        ))
        if payload is None:
            return None, "判定失败"
        return bool(payload.get("pass")), str(payload.get("reason", ""))[:200]

    def sentence_supported(self, sentence: str, cited_text: str,
                           gold_text: str = "") -> Tuple[Optional[bool], str]:
        """判断某个论断句是否被它引用的证据支持。"""

        payload = self._call(SUPPORT_PROMPT, (
            f"论断句：{sentence}\n\n"
            f"它引用的证据：\n{cited_text or '（无）'}\n\n"
            f"原论文相关证据：\n{gold_text or '（未提供）'}"
        ), max_tokens=250)
        if payload is None:
            return None, "核实失败"
        return bool(payload.get("supported")), str(payload.get("reason", ""))[:160]


# 兼容旧调用名
SentenceVerifier = Judge
