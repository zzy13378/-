"""从真实语料生成评测问题。

随机抽取语料块，让模型只依据该块写一个可回答的事实性问题，金标准就是这个块。
这样问题的正确答案来自语料本身，不需要人工编写期望关键词。
"""

from __future__ import annotations

import hashlib
import json
import random
import time
from typing import List

PROMPT = """你是评测集设计者。下面是一段真实论文证据，请基于它写一个具体的事实性问题。
要求：
1. 问题必须能且只能由这段证据回答；
2. 问题要具体，包含研究对象、方法或结论中的关键信息；
3. 不要出现“这段材料”“根据上文”这类指代；
4. 只输出 JSON：{"question":"问题内容","focus":"考察点"}
"""


class QuestionGenerator:
    def __init__(self, model: str = "", seed: int = 7) -> None:
        from openai import OpenAI
        import os

        base_url = os.environ.get("OPENAI_BASE_URL", "http://127.0.0.1:15721/v1")
        api_key = os.environ.get("OPENAI_API_KEY", "any")
        self.model = model or os.environ.get("LLM_MODEL", "deepseek-chat")
        self.client = OpenAI(api_key=api_key, base_url=base_url, timeout=120, max_retries=0)
        self.rng = random.Random(seed)
        self.tokens = 0

    def generate(self, chunks, count: int = 12) -> List[dict]:
        pool = [chunk for chunk in chunks if len(chunk.text) >= 150]
        self.rng.shuffle(pool)
        picked = pool[:count]
        questions = []
        for index, chunk in enumerate(picked, 1):
            user = f"论文标题：{chunk.title}\n证据片段：\n{chunk.text}"
            payload = self._call(user)
            if not payload or not payload.get("question"):
                continue
            questions.append({
                "id": f"RQ{index:02d}",
                "question": str(payload["question"]).strip(),
                "focus": str(payload.get("focus", ""))[:60],
                "gold_chunk": chunk.id,
                "gold_doc": chunk.doc_id,
                "gold_title": chunk.title,
                "gold_text": chunk.text,
            })
            time.sleep(0.2)
        return questions

    def _call(self, user: str):
        for _ in range(3):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[{"role": "system", "content": PROMPT}, {"role": "user", "content": user}],
                    temperature=0.4,
                    max_tokens=400,
                )
                content = (response.choices[0].message.content or "").strip()
                self.tokens += int(getattr(response.usage, "total_tokens", 0) or 0)
                start, end = content.find("{"), content.rfind("}")
                if start >= 0 and end > start:
                    return json.loads(content[start:end + 1])
            except Exception:  # noqa: BLE001
                time.sleep(2)
        return None
