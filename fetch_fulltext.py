#!/usr/bin/env python3
"""为开放获取论文抓取全文，提升证据深度。

只处理标记为开放获取（is_oa）且带 PDF 链接的论文，闭源论文保持摘要级证据，
不去访问需要订阅的内容。抓到的全文按论文存到 data/fulltext/，
并把路径与字符数写回语料记录，抽取阶段会自动优先使用全文。

用法：
    python fetch_fulltext.py --limit 50        # 先抓 50 篇
    python fetch_fulltext.py --limit 300       # 继续抓，已抓过的会跳过
    python fetch_fulltext.py --status          # 查看全文覆盖率
"""

from __future__ import annotations

import argparse
import io
import json
import re
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from urllib.parse import urlparse

from local_corpus import CORPUS_PATH, PROJECT_DIR, load_corpus, save_corpus

FULLTEXT_DIR = PROJECT_DIR / "data" / "fulltext"
USER_AGENT = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
MAX_BYTES = 25 * 1024 * 1024
MAX_CHARS = 120000

# 这些出版商会拒绝程序化下载（实测 ScienceDirect 返回 403），直接跳过，
# 既省时间也避免无谓的请求。
BLOCKED_HOSTS = (
    "sciencedirect.com", "link.springer.com", "onlinelibrary.wiley.com",
    "tandfonline.com", "sagepub.com", "cambridge.org", "jstor.org",
    "ieeexplore.ieee.org", "dl.acm.org", "pubs.acs.org", "pubs.rsc.org",
    "academic.oup.com", "journals.plos.org",
)


def safe_name(paper_id: str) -> str:
    return re.sub(r"[^A-Za-z0-9_\-]", "_", paper_id)[:80] or "paper"


def download_pdf(url: str, timeout: int = 30) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": "application/pdf,*/*"})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        data = response.read(MAX_BYTES + 1)
    if len(data) > MAX_BYTES:
        raise ValueError("文件超过大小上限")
    if not data[:5].startswith(b"%PDF"):
        raise ValueError("返回内容不是 PDF")
    return data


def unpaywall_url(doi: str, email: str = "p4literature@outlook.com") -> str:
    """按 DOI 查询开放获取链接；Unpaywall 的额度独立于 OpenAlex，限流压力更小。"""

    import urllib.parse as _parse

    endpoint = f"https://api.unpaywall.org/v2/{_parse.quote(doi)}?email={email}"
    request = urllib.request.Request(endpoint, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=20) as response:
        data = json.loads(response.read().decode())
    if not data.get("is_oa"):
        return ""
    location = data.get("best_oa_location") or {}
    return str(location.get("url_for_pdf") or location.get("url") or "")


def host_blocked(url: str) -> bool:
    host = urlparse(url).netloc.lower()
    return any(blocked in host for blocked in BLOCKED_HOSTS)


def extract_text(data: bytes) -> str:
    from pypdf import PdfReader

    reader = PdfReader(io.BytesIO(data))
    parts = []
    for page in reader.pages[:60]:
        try:
            parts.append(page.extract_text() or "")
        except Exception:  # noqa: BLE001
            continue
        if sum(len(part) for part in parts) > MAX_CHARS:
            break
    text = "\n".join(parts)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()[:MAX_CHARS]


def status() -> None:
    records = load_corpus()
    total = len(records)
    with_ft = [r for r in records if r.fulltext_chars]
    oa = [r for r in records if r.is_oa]
    print(f"语料库 {total} 篇")
    print(f"  标记开放获取：{len(oa)} 篇（{len(oa)/total:.0%}）")
    print(f"  已抓全文：{len(with_ft)} 篇（{len(with_ft)/total:.0%}）")
    if with_ft:
        avg = sum(r.fulltext_chars for r in with_ft) / len(with_ft)
        print(f"  全文平均字符数：{avg:.0f}")
    pending = [r for r in records if r.is_oa and r.oa_url and not r.fulltext_chars]
    print(f"  待抓取：{len(pending)} 篇")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=50)
    parser.add_argument("--sleep", type=float, default=2.0)
    parser.add_argument("--status", action="store_true")
    parser.add_argument("--source", choices=["corpus", "unpaywall"], default="corpus",
                        help="corpus：用语料里已存的 OA 链接；unpaywall：按 DOI 现查")
    args = parser.parse_args()

    if args.status:
        status()
        return

    records = load_corpus()
    FULLTEXT_DIR.mkdir(parents=True, exist_ok=True)
    if args.source == "corpus":
        pending = [r for r in records if r.is_oa and r.oa_url and not r.fulltext_chars and not host_blocked(r.oa_url)]
    else:
        pending = [r for r in records if r.doi and r.abstract and not r.fulltext_chars]
    print(f"待抓取 {len(pending)} 篇，本次最多处理 {args.limit} 篇\n")

    done = failed = skipped = 0
    for index, record in enumerate(pending[: args.limit], 1):
        target = FULLTEXT_DIR / f"{safe_name(record.paper_id)}.txt"
        try:
            url = record.oa_url
            if args.source == "unpaywall":
                url = unpaywall_url(record.doi)
                if not url:
                    skipped += 1
                    print(f"[{index}] 跳过 {record.paper_id} 非开放获取")
                    time.sleep(0.6)
                    continue
                if host_blocked(url):
                    skipped += 1
                    print(f"[{index}] 跳过 {record.paper_id} 出版商禁止下载（{urlparse(url).netloc}）")
                    time.sleep(0.6)
                    continue
                record.oa_url = url
                record.is_oa = True
            data = download_pdf(url)
            text = extract_text(data)
            if len(text) < 1200:
                raise ValueError(f"提取文本过短({len(text)}字符)")
            target.write_text(text, encoding="utf-8")
            record.fulltext_path = str(target.relative_to(PROJECT_DIR))
            record.fulltext_chars = len(text)
            done += 1
            print(f"[{index}/{min(len(pending), args.limit)}] 成功 {record.paper_id} {len(text)} 字符 | {record.title[:40]}")
        except Exception as exc:  # noqa: BLE001
            failed += 1
            print(f"[{index}/{min(len(pending), args.limit)}] 失败 {record.paper_id} {type(exc).__name__}: {str(exc)[:60]}")
        time.sleep(args.sleep)
        if index % 20 == 0:
            save_corpus(records)

    save_corpus(records)
    print(f"\n本次成功 {done} 篇，失败 {failed} 篇，跳过 {skipped} 篇")
    status()


if __name__ == "__main__":
    main()
