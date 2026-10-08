#!/usr/bin/env python3
"""把本地 PDF 批量并入内置文献库，为论文补上全文。

自动抓全文只能覆盖开放获取论文，而且多数出版商禁止程序化下载。学校图书馆能
合法获取的付费文献，可以自己下载后放进 pdf_inbox/ 目录，用这个脚本一次性入库：
脚本会从 PDF 里抽取文本、尝试识别 DOI 和标题，命中已有记录就补全文，
没命中就当新论文加入语料库。

用法：
    mkdir -p pdf_inbox && cp ~/Downloads/*.pdf pdf_inbox/
    python ingest_pdfs.py                 # 处理 pdf_inbox 下的全部 PDF
    python ingest_pdfs.py --dir 我的文献    # 指定目录
"""

from __future__ import annotations

import argparse
import re
from datetime import datetime
from pathlib import Path

from local_corpus import PROJECT_DIR, CorpusRecord, dedup_key, load_corpus, save_corpus

FULLTEXT_DIR = PROJECT_DIR / "data" / "fulltext"
DEFAULT_DIR = PROJECT_DIR / "pdf_inbox"
MAX_CHARS = 120000
DOI_PATTERN = re.compile(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+")


def extract_pdf_text(path: Path) -> str:
    from pypdf import PdfReader
    import io

    reader = PdfReader(io.BytesIO(path.read_bytes()))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception:  # noqa: BLE001
            raise ValueError("PDF 已加密")
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


def guess_title(text: str, fallback: str) -> str:
    for line in (text or "").splitlines()[:40]:
        line = line.strip()
        if 20 <= len(line) <= 200 and not re.search(r"^(abstract|摘要|keywords|doi|http)", line, re.I):
            if sum(ch.isalpha() or "\u4e00" <= ch <= "\u9fff" for ch in line) > len(line) * 0.6:
                return line
    return fallback


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dir", default=str(DEFAULT_DIR))
    args = parser.parse_args()

    inbox = Path(args.dir)
    if not inbox.exists():
        raise SystemExit(f"目录不存在：{inbox}（先建好并放入 PDF）")
    pdfs = sorted(inbox.glob("*.pdf")) + sorted(inbox.glob("*.PDF"))
    if not pdfs:
        raise SystemExit(f"{inbox} 里没有 PDF")

    FULLTEXT_DIR.mkdir(parents=True, exist_ok=True)
    records = load_corpus()
    by_doi = {r.doi.lower(): r for r in records if r.doi}
    by_title = {re.sub(r"\W+", "", r.title).casefold(): r for r in records if r.title}

    matched = added = failed = 0
    for index, path in enumerate(pdfs, 1):
        try:
            text = extract_pdf_text(path)
            if len(text) < 1200:
                raise ValueError(f"文本过短({len(text)}字符)，可能是扫描件")
            doi_match = DOI_PATTERN.search(text[:8000])
            doi = doi_match.group(0).rstrip(".,;)") if doi_match else ""
            record = by_doi.get(doi.lower()) if doi else None
            if record is None:
                title = guess_title(text, path.stem)
                record = by_title.get(re.sub(r"\W+", "", title).casefold())
            if record is not None:
                matched += 1
                action = f"命中已有记录 {record.paper_id}"
            else:
                title = guess_title(text, path.stem)
                record = CorpusRecord(
                    paper_id=f"local-{re.sub(r'[^A-Za-z0-9]+', '-', path.stem)[:40]}",
                    title=title,
                    year=None,
                    abstract="",
                    doi=doi,
                    language="zh" if re.search(r"[\u4e00-\u9fff]", title) else "en",
                    source="本地导入全文",
                    fetched_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                )
                records.append(record)
                which = dedup_key(record)
                if doi:
                    by_doi[doi.lower()] = record
                by_title[re.sub(r"\W+", "", title).casefold()] = record
                added += 1
                action = f"新增论文 {record.paper_id}"

            target = FULLTEXT_DIR / f"{re.sub(r'[^A-Za-z0-9_\-]', '_', record.paper_id)[:80]}.txt"
            target.write_text(text, encoding="utf-8")
            record.fulltext_path = str(target.relative_to(PROJECT_DIR))
            record.fulltext_chars = len(text)
            if doi and not record.doi:
                record.doi = doi
            print(f"[{index}/{len(pdfs)}] {action} | {len(text)} 字符 | {path.name[:40]}")
        except Exception as exc:  # noqa: BLE001
            failed += 1
            print(f"[{index}/{len(pdfs)}] 失败 {path.name[:40]} {type(exc).__name__}: {str(exc)[:60]}")

    save_corpus(records)
    print(f"\n完成：补全 {matched} 篇，新增 {added} 篇，失败 {failed} 篇")
    total_full = sum(1 for r in records if r.fulltext_chars)
    print(f"语料库共 {len(records)} 篇，其中带全文 {total_full} 篇")
    print("提示：入库后可删除 pdf_inbox 里的文件，全文已存到 data/fulltext/")


if __name__ == "__main__":
    main()
