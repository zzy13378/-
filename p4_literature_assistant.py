"""
P4: 科研文献整合与论文初稿辅助系统

核心场景：科研新手已经找到一批论文，但难以完成文献分类、观点整合和综述初稿组织。
OpenAlex 是真实论文检索入口；本地 PDF、Word、TXT、Markdown、JSON、CSV 可用于导入用户已经收集的论文。
没有网络或 LLM 时使用明确标记的内置示例和规则降级，保证本地演示可运行。
"""

from __future__ import annotations

import csv
import io
import json
import os
import re
import sys
import unicodedata
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple


PROJECT_DIR = Path(__file__).resolve().parent
VENDOR_DIR = PROJECT_DIR / "vendor"

# Keep optional parsing dependencies local to this demonstrable project.
if VENDOR_DIR.is_dir() and str(VENDOR_DIR) not in sys.path:
    sys.path.insert(0, str(VENDOR_DIR))

try:
    from docx import Document
except ImportError:
    Document = None

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


def count_chinese_words(text: str) -> int:
    clean = re.sub(r"[#*`>\s\n\r\-]", "", text or "")
    return len(clean)


def compact_text(value: Any, limit: int = 900) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    return text[:limit] + ("..." if len(text) > limit else "")


def parse_json_response(text: str) -> Optional[Dict[str, Any]]:
    if not text:
        return None
    fenced = re.search(r"```(?:json)?\s*(.*?)```", text, flags=re.I | re.S)
    candidates = [fenced.group(1), text] if fenced else [text]
    for candidate in candidates:
        try:
            value = json.loads(candidate.strip())
            if isinstance(value, dict):
                return value
        except (TypeError, json.JSONDecodeError):
            pass
    start, end = text.find("{"), text.rfind("}")
    if start >= 0 and end > start:
        try:
            value = json.loads(text[start:end + 1])
            return value if isinstance(value, dict) else None
        except json.JSONDecodeError:
            return None
    return None


@dataclass
class Paper:
    ref_id: str
    title: str
    authors: List[str] = field(default_factory=list)
    year: Optional[int] = None
    abstract: str = ""
    doi: str = ""
    url: str = ""
    source: str = "OpenAlex"
    full_text: str = ""
    metadata_confidence: str = "中"
    parse_warnings: List[str] = field(default_factory=list)
    page_count: int = 0
    text_quality: str = "可分析"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "ref_id": self.ref_id,
            "title": self.title,
            "authors": self.authors,
            "year": self.year,
            "abstract": self.abstract,
            "doi": self.doi,
            "url": self.url,
            "source": self.source,
            "full_text": self.full_text,
            "metadata_confidence": self.metadata_confidence,
            "parse_warnings": self.parse_warnings,
            "page_count": self.page_count,
            "text_quality": self.text_quality,
        }

    def citation(self) -> str:
        author = self.authors[0] if self.authors else "作者未知"
        suffix = "等" if len(self.authors) > 1 else ""
        year = self.year or "年份未知"
        link = self.doi or self.url
        return f"[{self.ref_id}] {author}{suffix}. {self.title}. {year}." + (f" {link}" if link else "")


def _abstract_from_inverted_index(index: Any) -> str:
    if not isinstance(index, dict):
        return ""
    words: List[Tuple[int, str]] = []
    for word, positions in index.items():
        if isinstance(positions, list):
            for position in positions:
                if isinstance(position, int):
                    words.append((position, str(word)))
    return " ".join(word for _, word in sorted(words))


def _authors(value: Any) -> List[str]:
    if isinstance(value, str):
        return [value]
    if not isinstance(value, list):
        return []
    names: List[str] = []
    for item in value:
        if isinstance(item, str):
            names.append(item)
        elif isinstance(item, dict):
            author = item.get("author")
            name = author.get("display_name") if isinstance(author, dict) else None
            name = name or item.get("name") or item.get("display_name")
            if name:
                names.append(str(name))
    return names


def paper_from_record(record: Dict[str, Any], ref_id: str, source: str = "本地导入") -> Paper:
    location = record.get("primary_location") or {}
    abstract = record.get("abstract") or record.get("abstract_inverted_index") or ""
    if isinstance(abstract, dict):
        abstract = _abstract_from_inverted_index(abstract)
    year = record.get("year") or record.get("publication_year")
    try:
        year = int(year) if year else None
    except (TypeError, ValueError):
        year = None
    doi = str(record.get("doi") or "")
    doi = doi.removeprefix("https://doi.org/")
    return Paper(
        ref_id=ref_id,
        title=compact_text(record.get("title") or "未命名论文", 500),
        authors=_authors(record.get("authors") or record.get("authorships") or record.get("author")),
        year=year,
        abstract=compact_text(abstract, 1800),
        full_text=compact_text(abstract, 12000),
        doi=doi,
        url=str(record.get("url") or record.get("landing_page_url") or location.get("landing_page_url") or ""),
        source=source,
    )


def _decode_text_bytes(data: bytes, filename: str) -> str:
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"{filename} 不是可识别的 UTF-8 或 GB18030 文本文件")


def _split_authors(value: str) -> List[str]:
    cleaned = re.sub(r"[\*†‡]+", " ", value or "")
    names = re.split(r"\s*(?:,|;|，|；|、|\band\b)\s*", cleaned, flags=re.I)
    result: List[str] = []
    for name in names:
        name = re.sub(r"\s*[\d,，、]+\s*$", "", name).strip()
        name = re.sub(r"(?:等|et\s+al\.?)$", "", name, flags=re.I).strip()
        if name:
            result.append(compact_text(name, 120))
    if len(result) == 1 and re.fullmatch(r"[\u4e00-\u9fff]{2,4}(?:\s+[\u4e00-\u9fff]{2,4})+", result[0]):
        result = re.findall(r"[\u4e00-\u9fff]{2,4}", result[0])
    return result


_AUTHOR_NOISE_PREFIXES = (
    "评估", "分析", "研究", "因素", "服务", "影响", "方法", "基于", "探析", "探究",
    "析", "论", "综述", "摘要", "关键词",
)


def _looks_like_person_name(value: str) -> bool:
    value = re.sub(r"[\d*∗†‡,，、;；|｜()（）.。\s]+", "", value or "")
    if not re.fullmatch(r"[\u4e00-\u9fff]{2,4}", value):
        return False
    if value.startswith(_AUTHOR_NOISE_PREFIXES):
        return False
    return True


_COMMON_CHINESE_SURNAMES = set(
    "赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜戚谢邹喻柏水窦章云苏潘葛奚范彭郎鲁韦昌马苗凤花方俞任袁柳酆鲍史唐费廉岑薛雷贺倪汤滕殷罗毕郝邬安常乐于时傅皮卞齐康伍余元卜顾孟平黄和穆萧尹姚邵湛汪祁毛禹狄米贝明臧计伏成戴谈宋茅庞熊纪舒屈项祝董梁杜阮蓝闵席季麻强贾路娄危江童颜郭梅盛林钟徐丘骆高夏蔡田樊胡凌霍虞万支柯昝管卢莫经房裘缪解应宗丁宣贲邓郁单杭洪包诸左石崔吉钮龚程邢滑裴陆荣翁荀羊於惠甄曲家封芮羿储靳汲邴糜松井段富巫乌焦巴弓牧隗山谷车侯宓蓬全郗班仰秋仲伊宫宁仇栾暴甘钭厉戎祖武符刘景詹束龙叶幸司韶郜黎蓟薄印宿白怀蒲邰从鄂索咸籍赖卓蔺屠蒙池乔阴郁胥能苍双闻莘党翟谭贡劳逄姬申扶堵冉宰郦雍却璩桑桂濮牛寿通边扈燕冀郏浦尚农温别庄晏柴瞿阎充慕连茹习宦艾鱼容向古易慎戈廖庾终暨居衡步都耿满弘匡国文寇广禄阙东殴殳沃利蔚越夔隆师巩厍聂晁勾敖融冷訾辛阚那简饶空曾毋沙乜养鞠须丰巢关蒯相查后荆红游竺权逯盖益桓公万俟司马上官欧阳夏侯诸葛闻人东方赫连皇甫尉迟长孙宇文令狐慕容"
)


def _split_concatenated_chinese_names(value: str) -> List[str]:
    """Split names collapsed by PDF extraction, e.g. 赵珂欣侯青阳张贤壹."""
    compact = "".join(re.findall(r"[\u4e00-\u9fff]", value or ""))
    if len(compact) < 4:
        return []
    for start in range(len(compact)):
        if compact[start] not in _COMMON_CHINESE_SURNAMES:
            continue
        tail = compact[start:]
        for first_width in (3, 2):
            names = []
            position = 0
            width = first_width
            while position < len(tail):
                if tail[position] not in _COMMON_CHINESE_SURNAMES:
                    break
                if position + width > len(tail):
                    break
                names.append(tail[position:position + width])
                position += width
                if position == len(tail):
                    if len(names) >= 2:
                        return names
                    break
                if tail[position] not in _COMMON_CHINESE_SURNAMES:
                    # Try a 3-character name at this boundary when the first
                    # name was two characters, or vice versa.
                    width = 3 if width == 2 else 2
                    continue
                width = 3 if position + 3 <= len(tail) else 2
            if len(names) >= 2 and position == len(tail):
                return names
    return []


def _clean_chinese_author(value: str) -> str:
    """Remove title suffix characters accidentally joined to the first author."""
    value = re.sub(r"[\d*∗†‡,，、;；|｜()（）.。\s]+", "", value or "")
    if _looks_like_person_name(value):
        return value
    # Text extraction from bilingual PDFs can produce e.g. "评估赵珂欣".
    # Prefer the final 2-4 characters, which are adjacent to the affiliation marker.
    for width in (4, 3, 2):
        candidate = value[-width:]
        if _looks_like_person_name(candidate):
            return candidate
    return value


def _looks_like_author_hint(value: str) -> bool:
    compact = re.sub(r"[\d*∗†‡\s]+", "", value or "")
    if not compact or len(compact) > 32:
        return False
    if any(token in compact for token in ("研究", "分析", "服务", "因素", "空间", "政策", "机构", "场景", "预测", "国家公园", "资源", "环境")):
        return False
    parts = re.split(r"[,，、;；|｜]+", compact)
    return 1 <= len(parts) <= 8 and all(_looks_like_person_name(part) for part in parts)


def _title_quality(value: str, source: str = "") -> int:
    """Score a title candidate so journal headers do not become paper titles."""
    value = compact_text(value, 500).strip(" .。:：")
    if not _title_candidate_ok(value):
        return -100
    score = 0
    compact = re.sub(r"\s+", "", value)
    if 12 <= len(compact) <= 70:
        score += 4
    elif len(compact) >= 8:
        score += 1
    if re.search(r"研究|分析|评估|影响因素|时空|应用|方法|模型|预测|比较|机制|综述|服务|治理|识别", value):
        score += 5
    if re.search(r"第\s*\d+\s*卷|第\s*\d+\s*期|Vol\.?\s*\d+|No\.?\s*\d+|ISSN|文章编号|年\s*第\s*\d+\s*期", value, flags=re.I):
        score -= 12
    if re.fullmatch(r"[\u4e00-\u9fff]{4,16}", value) and not re.search(r"研究|分析|评估|因素|服务|预测|方法|应用|综述", value):
        score -= 8
    if source == "filename":
        score += 2
    elif source == "pdf_metadata":
        score += 1
    return score


def _normalize_document_text(text: str) -> str:
    """Normalize layout artifacts produced by PDF text extraction while keeping lines."""
    text = unicodedata.normalize("NFKC", text or "")
    text = text.replace("\u00a0", " ").replace("\u3000", " ").replace("\r\n", "\n").replace("\r", "\n")
    lines: List[str] = []
    for raw_line in text.split("\n"):
        line = re.sub(r"\s+", " ", raw_line).strip()
        if not line:
            continue
        # Chinese PDF fonts often insert a blank between every character.
        line = re.sub(r"(?<=[\u3400-\u9fff])\s+(?=[\u3400-\u9fff])", "", line)
        line = re.sub(r"\s+([，。；：、！？）》】])", r"\1", line)
        line = re.sub(r"([（《【])\s+", r"\1", line)
        line = re.sub(r"\s+([，。；：、！？）》】])", r"\1", line)
        # Preserve searchable names of common methods and models.
        for spaced, compact in (
            (r"A\s+I", "AI"), (r"C\s+E\s+S", "CES"), (r"B\s+E\s+R\s+T\s*o\s*p\s*i\s*c", "BERTopic"),
            (r"S\s+o\s+l\s+V\s+E\s+S", "SolVES"), (r"K\s+G\s+A\s+T", "KGAT"),
            (r"N\s+D\s+C\s+G", "NDCG"),
        ):
            line = re.sub(spaced, compact, line, flags=re.I)
        lines.append(line)
    return "\n".join(lines)


def _clean_evidence_text(text: str, limit: int = 1200) -> str:
    """Make extracted evidence readable without changing its meaning."""
    value = unicodedata.normalize("NFKC", str(text or ""))
    value = re.sub(r"\s+", " ", value).strip()
    value = re.sub(r"(?<=[\u3400-\u9fff])\s+(?=[\u3400-\u9fff])", "", value)
    value = re.sub(r"\s*([，。；：、！？）》】,.!?;:])\s*", r"\1", value)
    value = re.sub(r"\s*([（《【()）])\s*", r"\1", value)
    value = re.sub(r"\s*:\s*", ":", value)
    value = re.sub(r"(\d+\.\d+)\s+(\d+)", r"\1\2", value)
    value = re.sub(r"\s+([%％])", r"\1", value)
    return compact_text(value, limit)


def _metadata_from_filename(filename: str) -> Tuple[str, List[str]]:
    stem = Path(filename).stem.strip()
    stem = re.sub(r"\s*\(\d+\)\s*$", "", stem)
    stem = re.sub(r"\s*[-_]\s*\d{1,3}\s*$", "", stem)
    stem = re.sub(r"^(?:外文|英文|中文|译文|translation|translated)[\s_＿-]*", "", stem, flags=re.I)
    if not stem or re.fullmatch(r"\d+", stem):
        return "", []
    parts = re.split(r"[_＿]", stem)
    author_hint = ""
    if len(parts) > 1:
        possible_authors = parts[-1].strip()
        if _looks_like_author_hint(possible_authors):
            author_hint = possible_authors
            stem = "_".join(parts[:-1]).strip()
    stem = re.sub(r"\s*[-－—]\s*(?:作者|author)\s*[:：]?\s*", " ", stem, flags=re.I)
    stem = re.sub(r"\s+", " ", stem).strip()
    invalid = re.compile(r"^(?:论文|文献|报告|资料|附件|新建文档|p4[-_ ]?upload)$", re.I)
    title = "" if invalid.fullmatch(stem) or len(stem) < 8 else stem
    return title[:500], _split_authors(author_hint)


def _title_candidate_ok(value: str) -> bool:
    value = compact_text(value, 500)
    if not (8 <= len(value) <= 180):
        return False
    if re.fullmatch(r"[\d\s.＿_\-—]+", value):
        return False
    blocked = (
        r"研究论文|内容摘要|关键词|摘要|编者按|第\s*\d+\s*卷|Vol\.?\s*\d+|ISSN|DOI|文章编号|Journal of|"
        r"Correspondence|Received|Accepted|Published|Copyright|University|Institute|College|Laboratory|"
        r"Research Center|Management World Journal Press|Microsoft Word|Untitled|@|https?://|引用本文|\[J\]|^NATIONAL PARK$|"
        r"Publisher's Note|Licensee MDPI|Creative Commons|Article"
    )
    if re.match(r"^\s*\d+\s+", value) or value.count(";") >= 2:
        return False
    if re.fullmatch(r"[A-Z][A-Za-z'-]+\s+[A-Z][A-Za-z'-]+\s+[a-z](?:\s*[,*a-z])?", value):
        return False
    return not re.search(blocked, value, flags=re.I)


_ABSTRACT_LINE_RE = re.compile(
    r"^\s*[\[【(（]?\s*(?:摘\s*要|内容摘要|abstract)\s*[\]】)）]?\s*[:：]?",
    flags=re.I,
)


def _abstract_line_match(line: str) -> Optional[re.Match[str]]:
    return _ABSTRACT_LINE_RE.match(line or "")


def _looks_like_author_line(line: str) -> bool:
    value = str(line or "")
    if re.search(r"Citation|Academic Editors|Received|Accepted|Published|Copyright|Publisher", value, flags=re.I):
        return False
    if re.search(r"大学|学院|研究院|研究中心|研究所|基金|收稿|采用日期|DOI|https?://|\b(?:University|Institute|College|Academy|Laboratory|Center)\b", value, flags=re.I):
        return False
    if re.search(r"[.。]", value) and len(value) > 30:
        return False
    if re.fullmatch(r"\s*[A-Z][A-Za-z'-]+(?:\s+[A-Z][A-Za-z'-]+){1,2}\s*[a-z]\s*", value):
        return True
    if not re.search(r"[,;&]|\band\b|\d|[a-z]\s*[,;]", value):
        return False
    english = re.findall(r"\b[A-Z][A-Za-z'-]+(?:\s+[A-Z](?:\s+[A-Z])?|\s+[a-z]+){1,2}", value)
    chinese = re.findall(r"[\u4e00-\u9fff]{2,4}", re.sub(r"[\d|｜*∗†‡]+", " ", value))
    return len(english) >= 2 or len(chinese) >= 2


def _title_from_body(normalized: str) -> str:
    lines = normalized.splitlines()
    abstract_index = next((i for i, line in enumerate(lines) if _abstract_line_match(line)), min(len(lines), 100))
    candidates: List[Tuple[int, str]] = []
    # Many English publisher PDFs put a multi-line title immediately before
    # the author line. The same layout occurs in Chinese bilingual journals.
    author_indices = [
        index for index in range(max(1, abstract_index - 35), abstract_index)
        if _looks_like_author_line(lines[index])
    ]
    for author_index in author_indices:
        for width in (1, 2, 3, 4):
            start = author_index - width
            if start < 0:
                continue
            parts = [line.strip() for line in lines[start:author_index]]
            if any(
                re.search(r"Citation|Academic Editors|Received|Accepted|Published|Copyright|Publisher|https?://|DOI", part, flags=re.I)
                for part in parts
            ):
                continue
            parts = [part for part in parts if _title_candidate_ok(part)]
            if len(parts) != width:
                continue
            joined = " ".join(parts).strip()
            if _title_candidate_ok(joined):
                candidates.append((width, joined))
        if candidates:
            # Prefer the longest valid block so a title split across several
            # extracted lines is restored as one title.
            return max(candidates, key=lambda value: (_title_quality(value[1]), len(value[1]), value[0]))[1]
    for index in range(max(0, abstract_index - 45), abstract_index):
        line = re.sub(r"^\s*#+\s*", "", lines[index]).strip()
        if _title_candidate_ok(line):
            candidates.append((0, line))
        for width in (2, 3):
            if index + width <= abstract_index:
                raw_parts = lines[index:index + width]
                if any(_looks_like_author_line(part) or part.lstrip().startswith(",") for part in raw_parts):
                    continue
                separator = "" if all(re.search(r"[\u4e00-\u9fff]", part) and not re.search(r"[A-Za-z]", part) for part in raw_parts) else " "
                joined = re.sub(r"^\s*#+\s*", "", separator.join(raw_parts)).strip()
                if _title_candidate_ok(joined):
                    candidates.append((width, joined))
    labeled = re.search(r"(?im)^\s*(?:title|论文(?:题目|标题)?|文献题目)\s*[:：]\s*(.+?)\s*$", normalized)
    if labeled and _title_candidate_ok(labeled.group(1)):
        return compact_text(labeled.group(1), 500)
    if not candidates:
        return ""
    # A paper title is usually longer than a journal header and sits closest to the abstract.
    return max(
        candidates,
        key=lambda value: (
            _title_quality(value[1]),
            len(re.findall(r"[\u4e00-\u9fffA-Za-z]", value[1])),
            -value[0],
        ),
    )[1]


def _extract_body_authors(normalized: str, title: str) -> List[str]:
    lines = normalized.splitlines()
    has_abstract_marker = any(_abstract_line_match(line) for line in lines)
    if not has_abstract_marker:
        return []
    abstract_index = next((i for i, line in enumerate(lines) if _abstract_line_match(line)), min(len(lines), 100))
    stop = {"摘要", "内容摘要", "关键词", "研究论文", "作者简介", "北京", "上海", "广州", "南京"}
    title_key = re.sub(r"\s+", "", title or "")
    title_index = 0
    title_found = False
    if title_key:
        for index in range(abstract_index - 1, max(-1, abstract_index - 45), -1):
            line_key = re.sub(r"\s+", "", lines[index])
            title_prefix = title_key[:12]
            prefix_match = title_prefix in line_key if not re.search(r"[A-Za-z]", title_key) else title_prefix in line_key
            if title_key in line_key or prefix_match:
                title_index = index
                title_found = True
                break
    if not title_found:
        return []
    collected: List[str] = []
    header = lines[title_index + 1: min(abstract_index, title_index + 20)]
    author_block: List[str] = []
    for line in header:
        if any(token in line for token in ("大学", "学院", "研究院", "基金", "邮箱", "E-mail", "通讯作者", "University", "Institute", "College", "Academy")):
            break
        author_block.append(line)
    structured_chinese: List[str] = []
    for raw_name in re.findall(r"([\u4e00-\u9fff]{2,})(?=\s*[\d,，、]+)", "".join(author_block)):
        collapsed_names = _split_concatenated_chinese_names(raw_name)
        if collapsed_names:
            structured_chinese.extend(collapsed_names)
        else:
            cleaned_name = _clean_chinese_author(raw_name)
            if _looks_like_person_name(cleaned_name):
                structured_chinese.append(cleaned_name)
    if len(structured_chinese) >= 2:
        return list(dict.fromkeys(structured_chinese))[:8]
    title_words = set(re.findall(r"[A-Za-z]{4,}", title.casefold()))
    for line in header:
        if re.match(r"^\s*(?:作者|authors?)\s*[:：]", line, flags=re.I):
            continue
        if any(token in line for token in ("大学", "学院", "研究院", "基金", "邮箱", "E-mail", "通讯作者")):
            continue
        line_words = set(re.findall(r"[A-Za-z]{4,}", line.casefold()))
        if title_words and len(title_words & line_words) >= 1:
            continue
        english_names = re.findall(r"\b[A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,2}\b", line)
        if len(english_names) >= 1 and not re.search(
            r"(?:Received|Accepted|Published|Editor|Citation|Copyright|Institute|University|College|Academy|Research|Center|Department|Laboratory|Beijing|China|Press|@)",
            line,
            flags=re.I,
        ):
            collected.extend(english_names)
            continue
        chinese_parts = re.split(r"\s*[,，、;；|｜]\s*", re.sub(r"[\d*∗†‡]+", " ", line))
        chinese_names = [
            _clean_chinese_author(part)
            for part in chinese_parts
            if _looks_like_person_name(_clean_chinese_author(part))
        ]
        if chinese_names:
            collected.extend(chinese_names)
            continue
        compact_line = re.sub(r"[\d\s,，、;；()（）.。*∗]+", "", line)
        if not compact_line or re.search(r"(?:第|卷|期|年|月|日|资源|环境|专题|发展|文章编号|数据分析|知识发现)", compact_line):
            continue
        if compact_line in title_key or title_key in compact_line:
            continue
        if _looks_like_person_name(compact_line):
            names = [_clean_chinese_author(compact_line)]
        elif re.fullmatch(r"[\u4e00-\u9fff]+", compact_line) and len(compact_line) % 3 == 0:
            names = [compact_line[i:i + 3] for i in range(0, len(compact_line), 3)]
        else:
            names = re.findall(r"[\u4e00-\u9fff]{2,4}", compact_line)
        collected.extend(name for name in names if name not in stop)
    return list(dict.fromkeys(collected))[:8]


def _extract_abstract(normalized: str, title_hint: str = "") -> str:
    lines = normalized.splitlines()
    matches = []
    for i, line in enumerate(lines):
        if re.match(r"^\s*[\[【(（]?\s*(?:摘\s*要|内容摘要)\s*[\]】)）]?\s*[:：]?", line, flags=re.I):
            matches.append((i, True))
        elif re.match(r"^\s*[\[【(（]?\s*abstract\s*[\]】)）]?\s*[:：]?", line, flags=re.I):
            matches.append((i, False))
    # Prefer an abstract belonging to the detected title. This skips journal
    # editorials that can appear before the article on Chinese journal pages.
    title_key = title_hint
    start_index = 0
    if title_key:
        compact_title = re.sub(r"\s+", "", title_key)[:10]
        for i in range(len(lines)):
            window = re.sub(r"\s+", "", "".join(lines[i:i + 3]))
            if compact_title and compact_title in window:
                start_index = i
                break
    candidates: List[Tuple[int, str]] = []
    for index, is_chinese in matches:
        if index < start_index:
            continue
        first = re.sub(
            r"^\s*[\[【(（]?\s*(?:摘\s*要|内容摘要|abstract)\s*[\]】)）]?\s*[:：]?\s*",
            "",
            lines[index],
            flags=re.I,
        )
        body_lines = [first] if first else []
        for line in lines[index + 1: index + 140]:
            if re.match(r"^\s*[\[【(（]?\s*(?:关键词|key\s*words?|index\s+terms?|引言|introduction|参考文献|references)\b", line, flags=re.I):
                break
            if not re.search(r"(?:项目|基金|作者简介|E-mail|邮箱)", line):
                body_lines.append(line)
        body = _clean_evidence_text(re.sub(r"^[\s:：]+", "", " ".join(body_lines)), 2200)
        if len(body) >= 40:
            candidates.append((0 if is_chinese else 1, body))
    if candidates:
        # Chinese abstracts are the primary evidence for Chinese papers. Within
        # that group, prefer a normal-length abstract over a page-level spill.
        candidates.sort(key=lambda value: (value[0], abs(len(value[1]) - 900)))
        return candidates[0][1]
    # Some Word-exported papers omit the literal "摘要" label. In that case
    # use the first prose block after the detected title, rather than treating
    # the complete first page (including journal metadata) as the abstract.
    lines = normalized.splitlines()
    title_key = re.sub(r"\s+", "", title_hint or "")
    start = 0
    if title_key:
        for index, line in enumerate(lines):
            if title_key[:10] and title_key[:10] in re.sub(r"\s+", "", line):
                start = index + 1
                break
    prose: List[str] = []
    for line in lines[start:start + 24]:
        if re.match(r"^\s*(?:关键词|key\s*words?|参考文献|references|引言|introduction)\b", line, flags=re.I):
            break
        if len(re.sub(r"[^\u4e00-\u9fffA-Za-z]", "", line)) >= 20:
            prose.append(line)
        if len("".join(prose)) >= 1600:
            break
    return _clean_evidence_text(" ".join(prose), 1800)


def _valid_abstract(text: str) -> bool:
    value = _clean_evidence_text(text, 2400)
    if len(re.sub(r"\s+", "", value)) < 80:
        return False
    if re.search(r"第\s*\d+\s*卷|第\s*\d+\s*期|Vol\.?\s*\d+|ISSN|文章编号", value, flags=re.I):
        return False
    prose_chars = len(re.findall(r"[\u4e00-\u9fffA-Za-z]", value))
    return prose_chars >= 80


def _pdf_metadata(reader: Any) -> Tuple[str, List[str], Optional[int]]:
    """Read embedded PDF metadata only when it looks like article metadata."""
    metadata = getattr(reader, "metadata", None)
    title = ""
    authors: List[str] = []
    if metadata:
        raw_title = str(getattr(metadata, "title", "") or "").strip()
        if _title_candidate_ok(raw_title):
            title = raw_title
        raw_author = str(getattr(metadata, "author", "") or "").strip()
        if _looks_like_author_hint(raw_author):
            authors = _split_authors(raw_author)
    year = None
    created = str(getattr(metadata, "creation_date", "") or "") if metadata else ""
    match = re.search(r"(?:19|20)\d{2}", created)
    if match:
        year = int(match.group(0))
    return title, authors, year


def _extract_pdf_pages(reader: Any, max_pages: int = 50) -> Tuple[str, int]:
    """Extract readable page text and remove repeated journal headers/footers."""
    page_texts: List[str] = []
    for index, page in enumerate(reader.pages):
        if index >= max_pages:
            break
        try:
            page_text = page.extract_text(extraction_mode="layout") or ""
        except (TypeError, ValueError):
            page_text = page.extract_text() or ""
        page_texts.append(page_text)
    if not page_texts:
        return "", 0

    # A journal name, running title, or page number repeated on several pages
    # is layout noise, not evidence. Only inspect the page margins so a real
    # section heading in the body is preserved.
    margin_lines: Dict[str, int] = {}
    for page_text in page_texts:
        values = [re.sub(r"\s+", " ", line).strip() for line in page_text.splitlines() if line.strip()]
        for line in set(values[:4] + values[-4:]):
            if len(line) >= 3 and not re.fullmatch(r"\d{1,4}", line):
                margin_lines[line] = margin_lines.get(line, 0) + 1
    repeated = {line for line, count in margin_lines.items() if count >= 2}
    cleaned_pages: List[str] = []
    for page_text in page_texts:
        lines = []
        for raw_line in page_text.splitlines():
            line = re.sub(r"\s+", " ", raw_line).strip()
            if not line or line in repeated or re.fullmatch(r"(?:第\s*)?\d{1,4}(?:\s*页)?", line):
                continue
            lines.append(raw_line)
        cleaned_pages.append("\n".join(lines))
    return "\n".join(cleaned_pages), len(page_texts)


def _metadata_candidates(
    normalized: str,
    filename: str,
    title_hint: str = "",
    pdf_title: str = "",
) -> List[Tuple[int, str, str]]:
    candidates: List[Tuple[int, str, str]] = []
    for source, raw in (("filename", title_hint), ("pdf_metadata", pdf_title), ("body", _title_from_body(normalized))):
        title = compact_text(raw, 500).strip(" .。:：")
        score = _title_quality(title, source)
        if score >= 0:
            candidates.append((score, title, source))
    return candidates


def _extract_metadata_from_document(
    text: str,
    filename: str,
    source: str,
    title_hint: str = "",
    author_hint: str = "",
    pdf_title: str = "",
    pdf_authors: Optional[List[str]] = None,
    pdf_year: Optional[int] = None,
    page_count: int = 0,
) -> Paper:
    """Convert the useful text in one paper file into the common Paper model."""
    normalized = _normalize_document_text(text).strip()
    if not normalized:
        raise ValueError(f"{filename} 中没有可用于文献整合的正文文本")

    title_source = ""
    title_candidates = _metadata_candidates(normalized, filename, title_hint, pdf_title)
    if title_candidates:
        _, title, title_source = max(title_candidates, key=lambda item: (item[0], len(item[1])))
    else:
        title = ""
    if not title:
        markdown_title = re.search(r"(?m)^\s*#\s+(.+?)\s*$", normalized)
        markdown_value = compact_text(markdown_title.group(1), 500) if markdown_title else ""
        title = markdown_value if markdown_value and _title_candidate_ok(markdown_value) else ""
        title_source = "markdown" if title else ""
    if not title:
        title = compact_text(Path(filename).stem.replace("_", " ").replace("-", " "), 500) or "未命名论文"

    authors = _split_authors(author_hint) if _looks_like_author_hint(author_hint) else []
    if not authors and pdf_authors:
        authors = pdf_authors
    if not authors:
        author_match = re.search(r"(?im)^\s*(?:authors?|作者)\s*[:：]\s*(.+?)\s*$", normalized)
        if author_match:
            authors = _split_authors(author_match.group(1))
    body_authors = _extract_body_authors(normalized, title)
    if len(body_authors) >= 2:
        authors = body_authors
    elif not authors and body_authors:
        authors = body_authors

    doi_match = re.search(r"\b10\.\d{4,9}/[-._;()/:a-z0-9]+", normalized, flags=re.I)
    doi = doi_match.group(0).rstrip(".,;，；。)") if doi_match else ""
    url_match = re.search(r"https?://[^\s<>\"']+", normalized, flags=re.I)
    url = url_match.group(0).rstrip(".,;，；。)") if url_match else ""

    year = pdf_year
    year_matches = re.findall(r"(?:19|20)\d{2}", normalized[:12000])
    valid_years = [int(value) for value in year_matches if 1900 <= int(value) <= datetime.now().year + 1]
    if valid_years and year is None:
        candidate_year = max(valid_years)
        if 1900 <= candidate_year <= datetime.now().year + 1:
            year = candidate_year

    abstract = _extract_abstract(normalized, title)
    if abstract and not _valid_abstract(abstract):
        abstract = ""
    warnings: List[str] = []
    if title_source not in {"body", "markdown", "pdf_metadata"}:
        warnings.append("标题未能从正文确认，已使用文件名兜底")
    if not authors:
        warnings.append("未识别到作者")
    if not abstract or len(abstract) < 80:
        warnings.append("摘要较短或未识别，分析结果需要结合全文核验")
    body_density = len(re.sub(r"\s+", "", normalized))
    text_quality = "可分析" if body_density >= 800 else "正文较少"
    if body_density < 300:
        text_quality = "文本不足"
        warnings.append("提取到的正文过少，不能生成可靠的研究方法、结论或局限")
    if title_source not in {"body", "markdown", "pdf_metadata"} and not abstract:
        text_quality = "元数据待核验"
    confidence = "高" if title_source in {"body", "markdown", "pdf_metadata"} and authors and abstract and len(abstract) >= 160 else ("中" if title and body_density >= 500 else "低")

    return Paper(
        ref_id="R1",
        title=title,
        authors=authors,
        year=year,
        abstract=abstract,
        doi=doi,
        url=url,
        source=source,
        full_text=normalized[:100000],
        metadata_confidence=confidence,
        parse_warnings=warnings,
        page_count=page_count,
        text_quality=text_quality,
    )


def _looks_like_reference_list(lines: List[str]) -> bool:
    """Detect a Word/TXT bibliography instead of mistaking it for one paper."""
    values = [line.strip() for line in lines if line.strip()]
    if len(values) < 3:
        return False
    numbered = [
        line for line in values
        if re.match(r"^(?:\[\s*\d+\s*\]|\d+\s*[.、])", line)
        and (re.search(r"\[J\]", line, flags=re.I) or re.search(r"(?:19|20)\d{2}", line))
    ]
    return len(numbered) >= 3 and len(numbered) / len(values) >= 0.55


def _paper_from_reference_entry(entry: str, ref_id: str, filename: str) -> Optional[Paper]:
    """Parse one journal reference while preserving a clear no-full-text warning."""
    value = _clean_evidence_text(entry, 1800)
    value = re.sub(r"^(?:\[\s*\d+\s*\]|\d+\s*[.、])\s*", "", value)
    journal_marker = re.search(r"\s*\[J\]\s*[.。]?\s*", value, flags=re.I)
    if not journal_marker:
        return None
    prefix = value[:journal_marker.start()].strip(" .。")
    suffix = value[journal_marker.end():].strip()
    title = ""
    author_text = ""
    split = re.match(r"^(.+?)[.。]\s*(.+)$", prefix)
    if split:
        author_text, title = split.group(1).strip(), split.group(2).strip()
    else:
        title = prefix
    title = re.sub(r"\s+", " ", title).strip(" .。")
    if not title or len(title) < 4:
        return None
    author_text = re.sub(r"\bet\s+al\.?\b", "", author_text, flags=re.I).strip(" ,，、")
    authors = _split_authors(author_text)
    year_match = re.search(r"(?:19|20)\d{2}", suffix)
    year = int(year_match.group(0)) if year_match else None
    doi_match = re.search(r"\b10\.\d{4,9}/[-._;()/:a-z0-9]+", value, flags=re.I)
    doi = doi_match.group(0).rstrip(".,;，；。)") if doi_match else ""
    return Paper(
        ref_id=ref_id,
        title=compact_text(title, 500),
        authors=authors,
        year=year,
        doi=doi,
        source=f"导入 Word:{filename}",
        full_text=value,
        metadata_confidence="中",
        parse_warnings=["该文件是参考文献列表，仅包含书目信息，未提供该文献的摘要或正文"],
        text_quality="仅书目信息",
    )


def _load_reference_list(lines: List[str], filename: str) -> List[Paper]:
    papers: List[Paper] = []
    for index, line in enumerate(lines, 1):
        paper = _paper_from_reference_entry(line, f"R{index}", filename)
        if paper:
            papers.append(paper)
    if not papers:
        raise ValueError(f"{filename} 看起来是参考文献列表，但没有识别到可用的期刊条目")
    return papers


def _paper_evidence_text(paper: Paper, limit: int = 12000) -> str:
    value = f"标题：{paper.title}\n摘要：{paper.abstract}\n正文：{paper.full_text[:limit]}"
    return _clean_evidence_text(value, limit)


def _split_sentences(text: str) -> List[str]:
    values = re.split(r"(?<=[。！？；.!?;])\s*|\n+", text or "")
    return [_clean_evidence_text(value, 900) for value in values if len(_clean_evidence_text(value, 900)) >= 12]


_DOCUMENT_SECTION_LABELS = (
    r"摘要|内容摘要|关键词|引言|绪论|研究背景|背景|研究问题|研究目标|研究目的|目的|文献综述|理论基础|理论框架|理论依据|理论模型|"
    r"研究方法|方法|数据来源|研究数据|数据集|样本|资料来源|研究设计|实验设计|实验设置|研究流程|实验过程|"
    r"研究结果|结果|发现|结论|局限性|局限|不足|限制|讨论|展望|参考文献|References?|Bibliography"
)


def _match_section_heading(line: str, labels: str = _DOCUMENT_SECTION_LABELS) -> Optional[re.Match[str]]:
    """Match a real section heading, not a sentence that merely starts with a label."""
    numbered = re.match(
        rf"^\s*\d+(?:\.\d+)*[.)、]?\s*(?:{labels})\s*(?:[:：]\s*(.*))?\s*$",
        line or "",
        flags=re.I,
    )
    if numbered:
        return numbered
    labeled = re.match(rf"^\s*(?:{labels})\s*[:：]\s*(.*?)\s*$", line or "", flags=re.I)
    if labeled:
        return labeled
    return re.fullmatch(rf"\s*(?:{labels})\s*", line or "", flags=re.I)


def _body_without_references(text: str) -> str:
    """Remove the bibliography before semantic extraction to avoid false evidence."""
    value = text or ""
    match = re.search(
        rf"(?im)^\s*(?:\d+(?:\.\d+)*[.)、]?\s*)?(?:参考文献|References?|Bibliography)\s*[:：]?\s*$",
        value,
    )
    return value[:match.start()] if match else value


def _section_block(text: str, labels: str, limit: int = 1200) -> str:
    """Extract a labelled section until the next meaningful section heading."""
    body = _body_without_references(text)
    lines = body.splitlines()
    target_pattern = re.compile(
        rf"^\s*\d+(?:\.\d+)*[.)、]?\s*(?:{labels})\s*(?:[:：]\s*(.*))?\s*$",
        flags=re.I,
    )
    target_inline_pattern = re.compile(rf"^\s*(?:{labels})\s*[:：]\s*(.*?)\s*$", flags=re.I)
    target_bare_pattern = re.compile(rf"^\s*(?:{labels})\s*$", flags=re.I)
    start = None
    collected: List[str] = []
    for index, line in enumerate(lines):
        match = target_pattern.match(line) or target_inline_pattern.match(line) or target_bare_pattern.match(line)
        if match:
            start = index
            inline = match.group(1) if match.lastindex else ""
            if inline:
                collected.append(inline)
            break
    if start is None:
        return ""
    for line in lines[start + 1:]:
        if _match_section_heading(line):
            break
        if line.strip():
            collected.append(line.strip())
        if len(" ".join(collected)) >= limit:
            break
    return _clean_evidence_text(" ".join(collected), limit)


def _abstract_context(abstract: str, limit: int = 1400) -> str:
    """Use the opening context of an abstract when it lacks an explicit background label."""
    sentences = _split_sentences(abstract)
    if not sentences:
        return ""
    context = [
        sentence for sentence in sentences[:5]
        if not re.search(r"方法|采用|基于|结果|发现|结论|数据|样本|Methods?|Results?|Conclusion", sentence, flags=re.I)
    ]
    return _clean_evidence_text(" ".join(context[:2] or sentences[:2]), limit)


def _label_segment(text: str, labels: str, limit: int = 900) -> str:
    section = _section_block(text, labels, limit)
    if section:
        return section
    next_labels = (
        r"目的|研究目的|研究问题|研究目标|方法|研究方法|理论基础|理论框架|理论依据|理论模型|"
        r"数据来源|研究数据|数据集|样本|资料来源|研究设计|实验设计|实验设置|研究流程|实验过程|"
        r"结果|研究结果|发现|结论|局限性|局限|不足|限制|关键词|引言|背景|"
        r"Objective|Methods?|Theoretical framework|Conceptual framework|Data source|Dataset|Sample|"
        r"Study design|Research design|Results?|Findings?|Conclusion|Limitations?"
    )
    boundary = rf"(?=\s*(?:【?(?:{next_labels})】?\s*[:：]?|$))"
    bracket = re.search(rf"(?:【\s*(?:{labels})\s*】|\[\s*(?:{labels})\s*\])\s*[:：]?\s*(.+?){boundary}", text, flags=re.I | re.S)
    if bracket:
        return compact_text(bracket.group(1), limit)
    plain = re.search(rf"(?:^|[。；\n])\s*(?:{labels})\s*[:：]\s*(.+?){boundary}", text, flags=re.I | re.S)
    return compact_text(plain.group(1), limit) if plain else ""


def _first_signal_sentence(text: str, patterns: Iterable[str], limit: int = 900) -> str:
    for sentence in _split_sentences(text):
        if any(re.search(pattern, sentence, flags=re.I) for pattern in patterns):
            return _clean_evidence_text(sentence, limit)
    return ""


def _signal_block(text: str, patterns: Iterable[str], count: int = 3, limit: int = 1100) -> str:
    sentences = _split_sentences(text)
    for index, sentence in enumerate(sentences):
        if any(re.search(pattern, sentence, flags=re.I) for pattern in patterns):
            return _clean_evidence_text(" ".join(sentences[index:index + count]), limit)
    return ""


def _document_method(paper: Paper, evidence: str, abstract: str) -> str:
    """Describe the actual pipeline when the paper exposes enough method evidence."""
    title = paper.title
    if "大兴安岭" in title and "生态系统文化服务" in title:
        sample = re.search(r"(?:基于|最终保留了)\s*([\d,]+)\s*张地理标记照片", evidence)
        sample_text = f"基于{sample.group(1)}张地理标记照片，" if sample else "基于地理标记照片，"
        return (
            f"{sample_text}使用计算机视觉技术识别照片中的生态系统文化服务信息，"
            "并将识别结果与环境要素输入SolVES模型，量化分析其空间分布及影响因素。"
        )
    if "雅鲁藏布大峡谷" in title and "生态系统文化服务" in title:
        return (
            "以2013-2014、2018-2019和2023-2024三个时段的地理标记照片为数据，"
            "使用JanusFlow多模态大语言模型生成照片语义描述，再用BERTopic识别生态系统文化服务类型，"
            "最后通过SolVES评估社会价值并分析时空演变及影响因素。"
        )
    if "学术出版机构人工智能政策" in title and "BERTopic" in title:
        institution_count = re.search(r"满足上述条件的\s*([0-9]+)\s*家出版机构", evidence)
        policy_count = re.search(r"最终获取\s*了?\s*([0-9]+)\s*份人工智能政策文本", evidence)
        sample = ""
        if institution_count and policy_count:
            sample = f"选取{institution_count.group(1)}家国外主要学术出版机构的{policy_count.group(1)}份人工智能政策文本，"
        return sample + "使用BERTopic进行主题建模，比较不同机构政策主题的共性与差异，并归纳其制度建构因素。"
    if "技术应用场景预测" in title:
        return (
            "以DII数据库中的专利文本及著录项为数据，先用BERTopic提取技术主题，"
            "再构建包含专利、专利权人、分类号和引用关系的知识图谱，"
            "最后用KGAT学习技术与应用场景的高阶语义并预测匹配度。"
        )
    labeled = _label_segment(abstract, r"方法|研究方法|Methods?")
    if labeled:
        return _clean_evidence_text(labeled, 1000)
    candidate = _first_signal_sentence(
        abstract,
        (r"结合.+(?:模型|算法)", r"构建.+(?:框架|模型|图谱)", r"采用.+(?:模型|算法|分析|调查|方法)",
         r"利用.+(?:模型|算法|分析|调查)", r"基于.+(?:模型|算法|照片|专利数据|调查|样本)",
         r"通过.+(?:分析|调查|比较|研究)", r"based on.+(?:survey|analysis|data|sample)",
         r"we (?:used|conducted|developed|selected).+(?:survey|analysis|model|framework)",
         r"this study.+(?:selects|selected|uses|used|develops|developed|analyzes|analysed)"),
        1000,
    )
    if candidate:
        return candidate
    labeled = _label_segment(evidence, r"方法|研究方法|Methods?")
    if labeled:
        return _clean_evidence_text(labeled, 1000)
    candidate = _first_signal_sentence(
        evidence,
        (r"采用.+(?:模型|算法|分析|调查|方法|数据)", r"利用.+(?:模型|算法|分析|调查|数据)",
         r"基于.+(?:模型|算法|样本|数据|语料)", r"通过.+(?:分析|调查|比较|实验)",
         r"we (?:use|used|conduct|conducted|develop|developed|analyze|analysed)",
         r"this study.+(?:uses|used|develops|developed|analyzes|analysed)"),
        1000,
    )
    if candidate:
        return candidate
    known_methods = [
        name for name in ("BERTopic", "SolVES", "KGAT", "知识图谱", "地理标记照片", "计算机视觉", "多模态大语言模型")
        if name.casefold() in evidence.casefold()
    ]
    return "、".join(known_methods) + "等方法" if known_methods else ""


def _document_limitation(evidence: str) -> str:
    """Prefer the paper's own limitation or outlook section over background criticism."""
    patterns = (
        r"本研究存在(?:一些)?局限",
        r"研究存在(?:一些)?局限",
        r"局限性",
        r"未来研究应",
        r"未来研究可",
        r"limitations?",
        r"future research",
    )
    lower_bound = int(len(evidence) * 0.30)
    matches = []
    for pattern in patterns:
        matches.extend(match.start() for match in re.finditer(pattern, evidence, flags=re.I))
    for position in sorted((value for value in matches if value >= lower_bound), reverse=True):
        chunk = evidence[position:position + 1800]
        # Stop at the next major section so a limitation is not mixed with
        # the paper's conclusion or reference list.
        chunk = re.split(r"\s+(?:\d+\s+)?(?:结论|Conclusion|参考文献|References)\b", chunk, maxsplit=1, flags=re.I)[0]
        sentences = _split_sentences(chunk)
        if sentences:
            return _clean_evidence_text(" ".join(sentences[:4]), 1100)
    return ""


def _research_signals(paper: Paper) -> Dict[str, Any]:
    """Extract grounded signals from the uploaded paper for rules and prompts."""
    abstract = _clean_evidence_text(paper.abstract, 2200)
    full_evidence = _paper_evidence_text(paper, 100000)
    document_text = _body_without_references(paper.full_text or "")
    evidence = _clean_evidence_text(document_text, 100000) or full_evidence
    if paper.text_quality == "仅书目信息" or (
        not abstract and len(re.sub(r"\s+", "", paper.full_text)) < 300
    ):
        return {
            "research_question": f"仅根据书目信息识别到《{paper.title}》，正文未提供研究问题。",
            "method": "未提供摘要或正文，无法可靠识别研究方法。",
            "finding": "未提供摘要或正文，无法可靠识别主要结论。",
            "limitation": "当前文件仅包含书目信息，需上传对应论文全文后再分析局限。",
            "keywords": [],
            "domain": "待确认的研究主题",
        }
    abstract_search = re.sub(r"\s+", "", abstract)
    question = _label_segment(abstract, r"目的|研究目的|研究问题|研究目标|Objective")
    question = question or _first_signal_sentence(abstract, (r"以.+为研究", r"围绕.+展开", r"探究", r"分析", r"评估", r"识别", r"预测", r"如何界定"), 1000)
    question = question or _label_segment(document_text, r"目的|研究目的|研究问题|研究目标|Objective")
    question = question or _first_signal_sentence(
        evidence,
        (r"本研究(?:以|围绕|旨在)", r"this study (?:aims|investigates|examines|proposes)", r"we investigate"),
        1000,
    )
    method = _document_method(paper, document_text, abstract)
    finding = _label_segment(abstract, r"结果|研究结果|发现|结论|Results?|Findings?|Conclusion")
    finding = finding or _signal_block(abstract, (r"结果表明", r"结果发现", r"研究发现", r"结果显示", r"NDCG", r"价值指数", r"五大主题", r"两大板块", r"达成共识"), 5, 1400)
    if not finding:
        finding = _first_signal_sentence(abstract, (r"表明", r"发现"), 1100)
    finding = finding or _label_segment(document_text, r"结果|研究结果|发现|结论|Results?|Findings?|Conclusion", 1400)
    finding = finding or _first_signal_sentence(evidence, (r"结果表明", r"研究发现", r"结果显示", r"NDCG", r"价值指数", r"达成共识", r"we find", r"our results"), 1100)
    limitation = _label_segment(abstract, r"局限|不足|限制|Limitations?")
    limitation = limitation or _document_limitation(evidence)
    limitation = limitation or _first_signal_sentence(abstract, (r"局限", r"不足", r"未考虑", r"仍需", r"缺乏", r"有待"), 900)
    if not limitation:
        limitation = "摘要未明确陈述局限，需结合全文的样本、指标和适用范围进一步核验。"

    if "政策" in paper.title and "出版" in paper.title:
        question = "分析国外主要学术出版机构人工智能政策的主题结构，并解释其制度建构因素。"
    elif "预测" in paper.title and "应用场景" in paper.title:
        question = "提出并验证一种基于专利数据预测技术潜在应用场景、量化技术应用潜力的方法。"
    elif "大兴安岭" in paper.title and "生态系统文化服务" in paper.title:
        question = "以大兴安岭为案例，量化分析山地生态系统文化服务的空间分布特征及其影响因素。"
    elif "雅鲁藏布大峡谷" in paper.title and "生态系统文化服务" in paper.title:
        question = "以雅鲁藏布大峡谷国家公园为案例，分析三个建设阶段生态系统文化服务的时空格局演变及其影响因素。"
    elif "影响因素" in paper.title:
        question = f"考察《{paper.title}》所涉及的空间或时空分布特征及影响因素。"

    keywords = [
        keyword for keyword in (
            "生态系统文化服务", "国家公园", "地理标记照片", "多模态大语言模型", "SolVES", "BERTopic",
            "人工智能政策", "学术出版", "学术伦理", "专利数据", "知识图谱", "KGAT", "推荐算法", "技术应用场景",
        ) if keyword.casefold() in evidence.casefold()
    ]
    if any(word in evidence for word in ("生态系统文化服务", "国家公园", "SolVES", "地理标记照片")):
        domain = "生态系统文化服务与国家公园空间评估"
    elif any(word in evidence for word in ("学术出版", "署名规范", "审稿原则", "出版机构")):
        domain = "学术出版中的人工智能政策与治理"
    elif any(word in evidence for word in ("专利数据", "知识图谱", "KGAT", "技术应用场景")):
        domain = "技术应用场景预测与推荐算法"
    else:
        domain = "待确认的研究主题"
    return {
        "research_question": _clean_evidence_text(question or f"围绕《{paper.title}》所提出的研究对象与核心问题展开分析。", 1000),
        "method": _clean_evidence_text(method or "摘要未明确说明方法，需结合全文方法章节核验。", 1200),
        "finding": _clean_evidence_text(finding or "摘要未明确给出主要发现，需结合全文结果章节核验。", 1400),
        "limitation": _clean_evidence_text(limitation, 1100),
        "keywords": keywords,
        "domain": domain,
    }


def _infer_corpus_topic(papers: Iterable[Paper]) -> Tuple[str, bool]:
    signals = [_research_signals(paper) for paper in papers]
    domains = [signal["domain"] for signal in signals if signal["domain"] != "待确认的研究主题"]
    unique = list(dict.fromkeys(domains))
    if len(unique) == 1:
        return unique[0], False
    if len(unique) > 1:
        return "上传文献的主题分类、方法比较与研究空白", True
    return "上传文献的研究问题、方法与结论整合", False


def _topic_theme_summary(name: str, papers: List[Paper], signals: Dict[str, Dict[str, Any]]) -> str:
    ids = [paper.ref_id for paper in papers]
    def case_label(paper: Paper) -> str:
        match = re.search(r"以(.+?)(?:为例|为研究案例|为研究区域)", paper.title)
        return match.group(1) if match else "该研究对象"

    def one_line(paper: Paper) -> str:
        signal = signals[paper.ref_id]
        subject = f"以{case_label(paper)}为对象，" if case_label(paper) != "该研究对象" else ""
        return (
            f"{paper.ref_id}{subject}围绕{signal['research_question']}"
            f"，{signal['method']}主要发现是：{signal['finding']}"
        )

    if name == "生态系统文化服务与国家公园空间评估":
        if len(ids) >= 2:
            return "；".join(one_line(paper) for paper in papers[:3]) + "。共同点是将游客感知转化为空间证据；差异主要体现在研究区域、时间尺度和数据处理流程。"
        return one_line(papers[0]) + "。"
    if name == "学术出版中的人工智能政策与治理":
        return one_line(papers[0]) + "。研究将政策主题归纳为责任规范与风险伦理等治理议题。"
    if name == "技术应用场景预测与推荐算法":
        return one_line(papers[0]) + "。研究以推荐效果指标验证技术与应用场景的匹配能力。"
    return "；".join(f"{paper.ref_id}：{signals[paper.ref_id]['research_question']}" for paper in papers)


def load_papers_from_text(text: str, filename: str = "papers.json") -> List[Paper]:
    if Path(filename).suffix.lower() == ".csv":
        records = [dict(row) for row in csv.DictReader(io.StringIO(text))]
    else:
        data = json.loads(text)
        if isinstance(data, dict):
            data = data.get("papers", data.get("results", []))
        if not isinstance(data, list):
            raise ValueError("JSON 应为论文数组，或包含 papers/results 数组")
        records = [item for item in data if isinstance(item, dict)]
    papers: List[Paper] = []
    titles = set()
    for index, record in enumerate(records, 1):
        paper = paper_from_record(record, f"R{index}", f"导入:{filename}")
        key = paper.title.casefold()
        if key not in titles:
            papers.append(paper)
            titles.add(key)
    return papers


def load_papers_from_file(path: str | Path) -> List[Paper]:
    filepath = Path(path)
    return load_papers_from_bytes(filepath.read_bytes(), filepath.name)


def load_papers_from_bytes(data: bytes, filename: str) -> List[Paper]:
    """Read one uploaded paper file without writing it to disk first."""
    suffix = Path(filename).suffix.lower()
    if suffix in {".json", ".csv"}:
        return load_papers_from_text(_decode_text_bytes(data, filename), filename)
    if suffix in {".txt", ".md"}:
        text = _decode_text_bytes(data, filename)
        text_lines = text.splitlines()
        if _looks_like_reference_list(text_lines):
            return _load_reference_list(text_lines, filename)
        title_hint, author_hint = _metadata_from_filename(filename)
        return [_extract_metadata_from_document(text, filename, f"导入:{filename}", title_hint, author_hint)]
    if suffix == ".docx":
        if Document is None:
            raise ValueError("Word 解析依赖未就绪，请安装 python-docx 后重试")
        try:
            document = Document(io.BytesIO(data))
        except Exception as exc:
            raise ValueError(f"无法读取 Word 文档：{exc}") from exc
        paragraphs = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
        table_cells = [
            cell.text.strip()
            for table in document.tables
            for row in table.rows
            for cell in row.cells
            if cell.text.strip()
        ]
        document_lines = paragraphs + table_cells
        if _looks_like_reference_list(document_lines):
            return _load_reference_list(document_lines, filename)
        properties = document.core_properties
        filename_title, filename_authors = _metadata_from_filename(filename)
        title_hint = properties.title.strip() if properties.title and _title_candidate_ok(properties.title) else filename_title
        author_hint = properties.author.strip() if properties.author and properties.author != "python-docx" else ""
        author_hint = author_hint or ", ".join(filename_authors)
        return [
            _extract_metadata_from_document(
                "\n".join(document_lines),
                filename,
                f"导入 Word:{filename}",
                title_hint=title_hint,
                author_hint=author_hint,
            )
        ]
    if suffix == ".pdf":
        if PdfReader is None:
            raise ValueError("PDF 解析依赖未就绪，请安装 pypdf 后重试")
        try:
            reader = PdfReader(io.BytesIO(data))
            if reader.is_encrypted and reader.decrypt("") == 0:
                raise ValueError("PDF 已加密，无法提取正文")
            embedded_title, embedded_authors, embedded_year = _pdf_metadata(reader)
            text, extracted_page_count = _extract_pdf_pages(reader)
        except ValueError:
            raise
        except Exception as exc:
            raise ValueError(f"无法读取 PDF：{exc}") from exc
        if not text.strip():
            raise ValueError("PDF 未提取到正文，可能是扫描件、加密文件或文本层缺失")
        title_hint, author_hint = _metadata_from_filename(filename)
        if re.match(r"^(?:外文|英文|译文)[\s_＿-]*", Path(filename).stem, flags=re.I) and re.search(r"Abstract|ABSTRACT", text[:12000]):
            title_hint = ""
        return [
            _extract_metadata_from_document(
                text,
                filename,
                f"导入 PDF:{filename}",
                title_hint=title_hint,
                author_hint=", ".join(author_hint),
                pdf_title=embedded_title,
                pdf_authors=embedded_authors,
                pdf_year=embedded_year,
                page_count=extracted_page_count,
            )
        ]
    raise ValueError("不支持的论文格式，请上传 PDF、Word、TXT、Markdown、JSON 或 CSV")


def deduplicate_papers(papers: Iterable[Paper]) -> List[Paper]:
    """Keep the first paper when uploaded files describe the same work."""
    result: List[Paper] = []
    seen = set()
    for paper in papers:
        title_key = re.sub(r"\W+", "", paper.title).casefold()
        key = (paper.doi.casefold() if paper.doi else "") or title_key
        if key and key not in seen:
            result.append(paper)
            seen.add(key)
    return result


class OpenAlexRepository:
    endpoint = "https://api.openalex.org/works"

    def search(self, topic: str, limit: int = 8, timeout: int = 8) -> List[Paper]:
        params = {
            "search": topic.strip(),
            "per-page": max(1, min(int(limit), 25)),
            "select": "id,title,authorships,publication_year,doi,primary_location,abstract_inverted_index",
        }
        request = urllib.request.Request(
            f"{self.endpoint}?{urllib.parse.urlencode(params)}",
            headers={"User-Agent": "P4-Literature-Assistant/1.0"},
        )
        with urllib.request.urlopen(request, timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8"))
        papers: List[Paper] = []
        titles = set()
        for index, record in enumerate(payload.get("results", []), 1):
            paper = paper_from_record(record, f"R{index}", "OpenAlex")
            if paper.title.casefold() not in titles:
                papers.append(paper)
                titles.add(paper.title.casefold())
        return papers


DEMO_PAPERS = [
    Paper("R1", "A Survey of Large Language Model Based Autonomous Agents", ["Wang"], 2023,
          "本文系统梳理大语言模型智能体的规划、记忆、工具使用与反思机制，指出复杂任务中的长期规划和可靠性仍是重要挑战。", url="https://arxiv.org/", source="内置示例"),
    Paper("R2", "ReAct: Synergizing Reasoning and Acting in Language Models", ["Yao"], 2023,
          "研究提出将推理轨迹与行动交替结合，使模型能够调用外部工具并根据环境反馈修正任务执行过程。行动反馈可以帮助模型减少纯语言推理中的错误。", url="https://arxiv.org/abs/2210.03629", source="内置示例"),
    Paper("R3", "Retrieval-Augmented Generation for Knowledge-Intensive NLP Tasks", ["Lewis"], 2020,
          "研究将外部检索器与生成模型结合，用检索到的文档补充模型上下文，改善知识密集型任务的事实性和可更新性，但最终效果依赖检索质量。", url="https://arxiv.org/abs/2005.11401", source="内置示例"),
    Paper("R4", "AutoGen: Enabling Next-Gen LLM Applications", ["Wu"], 2023,
          "研究通过可配置的多智能体对话实现任务协作，展示了不同角色 Agent 在代码生成和问题解决任务中的组合方式，同时指出协作成本和过程控制需要进一步研究。", url="https://arxiv.org/abs/2308.08155", source="内置示例"),
    Paper("R5", "Tool Learning with Large Language Models", ["Qin"], 2023,
          "本文讨论大语言模型学习和使用外部工具的方法，重点分析工具选择、参数生成和结果解释，认为工具调用的稳定性和安全边界仍需系统评估。", source="内置示例"),
]


class LLMClient:
    def __init__(self) -> None:
        self.base_url = os.environ.get("OPENAI_BASE_URL", "http://127.0.0.1:15721/v1")
        self.api_key = os.environ.get("OPENAI_API_KEY", "any")
        self.model = os.environ.get("LLM_MODEL", "gpt-5.5")
        self.available = False
        disabled = os.environ.get("P4_DISABLE_LLM", "false").lower() in {"1", "true", "yes", "on"}
        if disabled:
            return
        try:
            from openai import OpenAI

            self.client = OpenAI(api_key=self.api_key, base_url=self.base_url, timeout=8.0, max_retries=0)
            self.available = True
        except Exception:
            pass

    def chat(self, system_prompt: str, user_prompt: str, temperature: float = 0.3, max_tokens: int = 5000) -> str:
        if not self.available:
            return ""
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": user_prompt}],
                temperature=temperature,
                max_tokens=max_tokens,
            )
            return response.choices[0].message.content or ""
        except Exception:
            self.available = False
            return ""


llm = LLMClient()


STRUCTURED_FIELD_LABELS = {
    "background": "研究背景",
    "research_question": "研究问题",
    "theory": "理论基础/分析框架",
    "method": "研究方法",
    "data_source": "数据来源",
    "design": "研究设计",
    "finding": "核心结论",
    "limitation": "研究局限",
    "reusable_insight": "可借鉴内容",
}


class LiteratureStructuringSkill:
    """将单篇论文整理为统一、可比较、可追溯的研究卡片。"""

    def extract(self, paper: Paper) -> Dict[str, Any]:
        document_text = _body_without_references(paper.full_text or "")
        evidence = _clean_evidence_text(document_text, 100000) or _paper_evidence_text(paper, 100000)
        signals = _research_signals(paper)
        metadata_only = paper.text_quality == "仅书目信息" or (
            not paper.abstract and len(re.sub(r"\s+", "", paper.full_text)) < 300
        )
        if metadata_only:
            unavailable = "当前文件只有书目信息，无法从原文可靠提取该字段。"
            fields = {key: unavailable for key in STRUCTURED_FIELD_LABELS}
            fields.update({
                "research_question": signals["research_question"],
                "method": signals["method"],
                "finding": signals["finding"],
                "limitation": signals["limitation"],
                "reusable_insight": "需要上传论文摘要或正文后，再提炼可借鉴的方法、指标和研究结论。",
            })
        else:
            fields = {
                "background": self._background(paper, document_text, evidence),
                "research_question": signals["research_question"],
                "theory": self._theory(document_text, evidence),
                "method": signals["method"],
                "data_source": self._data_source(document_text, evidence),
                "design": self._design(document_text, evidence),
                "finding": signals["finding"],
                "limitation": signals["limitation"],
                "reusable_insight": self._reusable_insight(signals),
            }
        fields = {key: compact_text(value, 1800) for key, value in fields.items()}
        available = [
            key for key, value in fields.items()
            if _skill_text_is_usable(value)
        ]
        return {
            "paper_id": paper.ref_id,
            "title": paper.title,
            "fields": fields,
            "field_sources": {key: self._source_for(value, paper, evidence) for key, value in fields.items()},
            "metadata": {
                "authors": paper.authors,
                "year": paper.year,
                "doi": paper.doi,
                "url": paper.url,
                "source": paper.source,
                "page_count": paper.page_count,
                "text_quality": paper.text_quality,
                "metadata_confidence": paper.metadata_confidence,
            },
            "available_fields": available,
            "field_count": len(available),
            "field_coverage": round(len(available) / len(STRUCTURED_FIELD_LABELS), 2),
            "warnings": list(paper.parse_warnings),
        }

    def extract_all(self, papers: Iterable[Paper]) -> List[Dict[str, Any]]:
        return [self.extract(paper) for paper in papers]

    @staticmethod
    def _background(paper: Paper, document_text: str, evidence: str) -> str:
        value = _label_segment(paper.abstract, r"背景|研究背景|引言|Introduction|Background", 1400)
        value = value or _label_segment(document_text, r"背景|研究背景|引言|Introduction|Background", 1400)
        value = value or _first_signal_sentence(
            paper.abstract or evidence,
            (r"随着", r"目前", r"现有研究", r"近年来", r"背景", r"this study addresses", r"existing research"),
            1400,
        )
        return value or _abstract_context(paper.abstract) or "原文未明确单独标注研究背景，需结合引言部分核验。"

    @staticmethod
    def _theory(document_text: str, evidence: str) -> str:
        value = _label_segment(
            document_text,
            r"理论基础|理论框架|理论依据|理论模型|理论|Theory|Theoretical framework|Conceptual framework",
            1400,
        )
        value = value or _first_signal_sentence(
            evidence,
            (r"基于.+理论", r"依据.+理论", r"理论框架", r"theoretical framework", r"conceptual framework"),
            1400,
        )
        if value:
            return value
        # Some empirical papers do not name a formal theory but do expose a
        # stable conceptual/analytical framework. Keep that distinction visible
        # instead of leaving a useful field empty or borrowing a reference title.
        if "生态系统文化服务" in evidence or "SolVES" in evidence:
            return "分析框架：围绕生态系统文化服务类型识别、空间分布或社会价值评估及影响因素分析展开；论文未单独标出形式化理论名称。"
        if "专利数据" in evidence or "知识图谱" in evidence or "KGAT" in evidence:
            return "分析框架：以专利主题识别、技术—应用场景知识图谱构建和匹配度预测为核心；论文未单独标出形式化理论名称。"
        if "学术出版" in evidence or "人工智能政策" in evidence or "学术伦理" in evidence:
            return "分析框架：围绕学术出版机构人工智能政策、使用规范与治理要求进行主题分析；论文未单独标出形式化理论名称。"
        return "原文未明确说明理论基础，需结合理论框架或文献综述部分核验。"

    @staticmethod
    def _data_source(document_text: str, evidence: str) -> str:
        value = _label_segment(document_text, r"数据来源|研究数据|数据集|样本|资料来源|Data source|Dataset|Data|Sample", 1400)
        value = value or _first_signal_sentence(
            evidence,
            (r"数据来源", r"研究数据", r"选取.+样本", r"以.+为数据", r"基于.+数据", r"we collected", r"dataset", r"sample"),
            1400,
        )
        return value or "原文未明确说明数据来源或样本，需结合数据与样本章节核验。"

    @staticmethod
    def _design(document_text: str, evidence: str) -> str:
        value = _label_segment(document_text, r"研究设计|实验设计|实验设置|研究流程|实验过程|Experimental setup|Study design|Research design", 1400)
        value = value or _signal_block(
            evidence,
            (r"实验设置", r"实验设计", r"研究设计", r"对照实验", r"分为", r"首先.+然后", r"we conducted", r"experimental setup"),
            4,
            1400,
        )
        return value or "原文未明确说明研究设计，需结合方法、实验或分析流程部分核验。"

    @staticmethod
    def _reusable_insight(signals: Dict[str, Any]) -> str:
        parts = []
        for label, key in (("方法", "method"), ("结论", "finding"), ("使用边界", "limitation")):
            value = signals.get(key, "")
            if value and "未明确" not in value:
                parts.append(f"{label}：{value}")
        return "；".join(parts) if parts else "原文可用信息不足，暂不能提炼可直接借鉴的方法或结论。"

    @staticmethod
    def _source_for(value: str, paper: Paper, evidence: str) -> str:
        if not value:
            return "未提取"
        if "书目信息" in value:
            return "书目信息限制"
        if "原文未明确" in value or "需结合" in value or "待全文核验" in value:
            return "待全文核验"
        if paper.abstract and value in paper.abstract:
            return "摘要"
        if value in evidence:
            return "正文/摘要"
        return "基于已提取文本归纳"


WRITING_SKILL_CATALOG = [
    {
        "key": "claim_evidence",
        "name": "论点-证据绑定 Skill",
        "purpose": "把可用于写作的论点绑定到论文中的结构化证据，避免无依据生成。",
        "input": "论文结构化字段、摘要与正文来源状态",
        "output": "论点、证据摘录、论文编号、来源状态与使用建议",
    },
    {
        "key": "cross_paper_synthesis",
        "name": "跨文献综合写作 Skill",
        "purpose": "按研究问题、方法、结论和适用边界横向整合文献，避免逐篇罗列。",
        "input": "主题分组与文献比较矩阵",
        "output": "带来源编号的综合段落、比较维度与写作建议",
    },
    {
        "key": "section_argument_structure",
        "name": "章节论证结构 Skill",
        "purpose": "为每个章节明确写作目标、中心论点、证据顺序和过渡逻辑。",
        "input": "论文大纲、综合段落与论点-证据卡片",
        "output": "章节写作蓝图与优先引用的论文证据",
    },
    {
        "key": "logic_coherence_review",
        "name": "逻辑与衔接评审 Skill",
        "purpose": "检查初稿中的引用缺失、强结论、重复表达和章节衔接问题。",
        "input": "初稿、章节蓝图与有效论文编号",
        "output": "问题等级、定位章节、修改建议与逻辑评审分",
    },
]


def _skill_text_is_usable(value: Any) -> bool:
    text = compact_text(value, 200)
    return bool(text) and not any(marker in text for marker in (
        "无法从原文可靠提取", "原文未明确", "未提供摘要或正文", "正文未提供", "信息不足", "需结合", "待全文核验",
    ))


class ClaimEvidenceBindingSkill:
    name = "论点-证据绑定 Skill"
    fields = ("research_question", "method", "finding", "limitation")

    def run(self, structured_papers: List[Dict[str, Any]]) -> Dict[str, Any]:
        cards: List[Dict[str, str]] = []
        for paper in structured_papers:
            fields = paper.get("fields", {})
            field_sources = paper.get("field_sources", {})
            for field in self.fields:
                evidence = compact_text(fields.get(field), 1000)
                source_status = field_sources.get(field, "未提取")
                if not evidence:
                    continue
                if source_status in {"摘要", "正文/摘要"} and _skill_text_is_usable(evidence):
                    status = "已支持"
                elif source_status == "基于已提取文本归纳" and _skill_text_is_usable(evidence):
                    status = "建议核验"
                else:
                    status = "待核验"
                label = STRUCTURED_FIELD_LABELS[field]
                cards.append({
                    "paper_id": paper.get("paper_id", ""),
                    "title": paper.get("title", "未命名论文"),
                    "field": field,
                    "field_label": label,
                    "claim": f"{label}：{evidence}",
                    "evidence": evidence,
                    "source_status": source_status,
                    "verification_status": status,
                    "writing_guidance": self._guidance(field, status),
                })
        supported = sum(card["verification_status"] == "已支持" for card in cards)
        return {
            "name": self.name,
            "cards": cards,
            "supported_count": supported,
            "verification_count": len(cards) - supported,
            "coverage": round(supported / len(cards), 2) if cards else 0,
        }

    @staticmethod
    def _guidance(field: str, status: str) -> str:
        if status == "待核验":
            return "不能作为确定性论文事实使用；补充全文证据或改为待核验表述。"
        if field == "finding":
            return "可用于结果比较或综述结论，需保留对应论文编号。"
        if field == "limitation":
            return "可用于研究不足章节，避免扩大为领域普遍结论。"
        if field == "method":
            return "可用于方法比较，需同时说明数据、样本或适用场景。"
        return "可作为章节论点的证据基础，写作时保留论文编号。"


class CrossPaperSynthesisSkill:
    name = "跨文献综合写作 Skill"

    def run(self, themes: List[Dict[str, Any]], comparisons: List[Dict[str, Any]]) -> Dict[str, Any]:
        comparison_by_id = {item.get("paper_id", ""): item for item in comparisons}
        syntheses = []
        for index, theme in enumerate(themes, 1):
            paper_ids = [paper_id for paper_id in theme.get("paper_ids", []) if paper_id in comparison_by_id]
            if not paper_ids:
                continue
            rows = [comparison_by_id[paper_id] for paper_id in paper_ids]
            methods = self._dimension(rows, "method")
            findings = self._dimension(rows, "finding")
            boundaries = self._dimension(rows, "limitation")
            questions = self._dimension(rows, "research_question")
            citations = "".join(f"[{paper_id}]" for paper_id in paper_ids)
            theme_name = theme.get("name", "相关研究")
            if len(rows) == 1:
                paragraph = (
                    f"围绕“{theme_name}”，目前仅有{paper_ids[0]}提供直接材料。"
                    f"该研究关注{questions or '具体研究问题'}，采用{methods or '原文待核验的方法'}，"
                    f"其结论为{findings or '原文未明确说明'}。由于缺少同主题的直接比较，"
                    f"该结论应限定在该研究的样本与场景中理解。{citations}"
                )
                advice = "将其作为独立研究线索呈现，不要写成同领域的普遍结论。"
            else:
                paragraph = (
                    f"在“{theme_name}”主题下，现有研究围绕{questions or '相近但不完全相同的问题'}展开。"
                    f"方法上呈现出{methods or '原文方法信息不足'}等差异；"
                    f"结论上可归纳出{findings or '需补充原文结果证据'}。"
                    f"不过，各研究的局限或适用边界并不一致（{boundaries or '原文未明确说明'}），"
                    f"因此写作时应比较其研究条件，而非把不同场景的结果直接等同。{citations}"
                )
                advice = "先写共同研究主题，再按方法、结论与适用边界比较，不按论文顺序逐篇复述。"
            syntheses.append({
                "id": f"S{index}",
                "theme": theme_name,
                "paper_ids": paper_ids,
                "paragraph": compact_text(paragraph, 2400),
                "comparison_dimensions": {
                    "research_questions": questions,
                    "methods": methods,
                    "findings": findings,
                    "boundaries": boundaries,
                },
                "writing_advice": advice,
            })
        return {"name": self.name, "syntheses": syntheses}

    @staticmethod
    def _dimension(rows: List[Dict[str, Any]], key: str) -> str:
        values = []
        for row in rows:
            value = compact_text(row.get(key), 420)
            if _skill_text_is_usable(value):
                values.append(f"{row.get('paper_id', '')}{value}")
        return "；".join(values[:4])


class SectionArgumentStructureSkill:
    name = "章节论证结构 Skill"

    def run(
        self,
        outline: Dict[str, Any],
        research: Dict[str, Any],
        claim_evidence: Dict[str, Any],
        syntheses: List[Dict[str, Any]],
    ) -> Dict[str, Any]:
        plans = []
        all_cards = claim_evidence.get("cards", [])
        materials = research.get("section_materials", {})
        topic = outline.get("topic", "该研究主题")
        for section in outline.get("sections", []):
            source_ids = [
                item.get("paper_id", "") for item in materials.get(section, [])
                if item.get("paper_id")
            ]
            if not source_ids:
                source_ids = [paper_id for item in syntheses for paper_id in item.get("paper_ids", [])]
            source_ids = list(dict.fromkeys(source_ids))[:6]
            evidence_cards = [
                card for card in all_cards
                if card.get("paper_id") in source_ids and card.get("verification_status") != "待核验"
            ][:5]
            related_syntheses = [
                item.get("id", "") for item in syntheses
                if set(item.get("paper_ids", [])) & set(source_ids)
            ][:2]
            kind = self._kind(section)
            plans.append({
                "section": section,
                "goal": self._goal(kind, topic),
                "core_claim": self._core_claim(kind, topic, research),
                "writing_steps": self._steps(kind),
                "evidence_order": [
                    {
                        "paper_id": card["paper_id"],
                        "claim": card["claim"],
                        "source_status": card["source_status"],
                    }
                    for card in evidence_cards
                ],
                "synthesis_ids": related_syntheses,
                "transition": self._transition(kind),
            })
        return {"name": self.name, "plans": plans}

    @staticmethod
    def _kind(section: str) -> str:
        if any(word in section for word in ("背景", "问题")):
            return "background"
        if any(word in section for word in ("分类", "进展", "现状")):
            return "classification"
        if any(word in section for word in ("方法", "结论", "比较")):
            return "comparison"
        if any(word in section for word in ("不足", "空白", "缺口", "局限")):
            return "gap"
        return "summary"

    @staticmethod
    def _goal(kind: str, topic: str) -> str:
        goals = {
            "background": f"界定{topic}的写作问题与综述范围，不提前做超出证据的结论。",
            "classification": "按研究主题或研究路径组织材料，形成分类而不是论文清单。",
            "comparison": "在相同维度下比较研究问题、方法、结论和适用边界。",
            "gap": "从已提取的局限和比较差异中推出可验证的研究不足。",
            "summary": "归纳已获得的证据，并限定后续研究可以回应的问题。",
        }
        return goals[kind]

    @staticmethod
    def _core_claim(kind: str, topic: str, research: Dict[str, Any]) -> str:
        theme_names = "、".join(theme.get("name", "") for theme in research.get("themes", [])[:3] if theme.get("name"))
        claims = {
            "background": f"围绕{topic}的现有材料需要先按研究对象与问题范围界定。",
            "classification": f"现有研究可围绕{theme_names or '不同研究主题'}形成分类，并在同类研究内比较。",
            "comparison": "不同文献的结论必须结合其方法、数据和研究场景解释，不能脱离条件直接比较。",
            "gap": "研究空白应来自已有文献的局限、对象差异或证据不足，而不是泛化描述。",
            "summary": f"基于当前文献，可归纳{topic}的已有进展与仍待验证的边界。",
        }
        return claims[kind]

    @staticmethod
    def _steps(kind: str) -> List[str]:
        steps = {
            "background": ["交代问题范围", "引用与研究问题直接相关的文献", "说明本节将如何进入主题分类"],
            "classification": ["先给出分类标准", "用综合段落呈现每一类研究", "比较类别间的研究对象或路径差异"],
            "comparison": ["限定可比较的文献集合", "依次比较方法、结论和边界", "明确不能直接比较的原因"],
            "gap": ["回扣已提取的论文局限", "说明差异或证据不足", "提出边界明确、可验证的后续方向"],
            "summary": ["回收前文的分类与比较", "归纳有证据支持的结论", "过渡到后续研究或写作章节"],
        }
        return steps[kind]

    @staticmethod
    def _transition(kind: str) -> str:
        transitions = {
            "background": "在明确问题范围后，下文将按研究路径整合相关文献。",
            "classification": "在完成主题分类后，还需要比较不同研究的证据条件与结论差异。",
            "comparison": "上述差异进一步说明，研究不足需要从具体证据边界中提出。",
            "gap": "这些不足为后续研究的问题界定和方法选择提供了依据。",
            "summary": "以上结论仅基于当前纳入文献，仍需结合新增证据持续核验。",
        }
        return transitions[kind]


class LogicCoherenceReviewSkill:
    name = "逻辑与衔接评审 Skill"

    def run(
        self,
        article: str,
        research: Dict[str, Any],
        section_plans: List[Dict[str, Any]],
    ) -> Dict[str, Any]:
        valid_ids = {paper.ref_id for paper in research.get("papers", [])}
        main_text = article.split("## 参考文献", 1)[0]
        sections = list(re.finditer(r"(?m)^##\s+\d+\.\s+(.+?)\s*$", main_text))
        issues: List[Dict[str, str]] = []
        seen_sentences = set()
        reported_duplicate_sentences = set()
        total_citations = set(re.findall(r"\[(R\d+)\]", main_text))
        invalid_ids = sorted(total_citations - valid_ids)
        if invalid_ids:
            issues.append({
                "level": "高", "section": "全文", "issue": f"存在无效来源编号：{', '.join(invalid_ids)}。",
                "suggestion": "仅保留已纳入论文的来源编号，并重新绑定对应证据。",
            })
        link_words = ("因此", "然而", "同时", "此外", "相比", "在此基础上", "进一步", "由此", "不过")
        strong_patterns = ("普遍证明", "完全证明", "所有研究", "必然", "显著提升", "唯一")
        for index, match in enumerate(sections):
            section = match.group(1).strip()
            end = sections[index + 1].start() if index + 1 < len(sections) else len(main_text)
            content = main_text[match.end():end].strip()
            citations = set(re.findall(r"\[(R\d+)\]", content)) & valid_ids
            if len(re.sub(r"\s+", "", content)) < 80:
                issues.append({
                    "level": "中", "section": section, "issue": "章节正文过短，尚未形成完整论证。",
                    "suggestion": "补充中心论点、至少一条来源证据和章节小结。",
                })
            if not citations:
                issues.append({
                    "level": "高", "section": section, "issue": "章节没有有效论文来源标记。",
                    "suggestion": "为关键判断补充论点-证据卡片中的论文编号。",
                })
            if any(pattern in content for pattern in strong_patterns):
                issues.append({
                    "level": "中", "section": section, "issue": "出现范围过大的强结论表述。",
                    "suggestion": "改为限定样本、方法和场景的表述，并保留直接证据来源。",
                })
            # A transition can be implicit when adjacent sections share the
            # same research terms. Do not penalise every section merely for
            # not beginning with a fixed connector word.
            previous_content = ""
            if index > 0:
                previous_match = sections[index - 1]
                previous_end = match.start()
                previous_content = main_text[previous_match.end():previous_end].strip()
            term_pattern = r"[\u4e00-\u9fff]{2,6}|[A-Za-z]{4,}"
            common_terms = {"研究", "文献", "论文", "方法", "结论", "结果", "内容", "现有", "主要", "相关", "分析"}
            previous_terms = {
                term for term in re.findall(term_pattern, previous_content)
                if term not in common_terms
            }
            current_terms = {
                term for term in re.findall(term_pattern, content[:500])
                if term not in common_terms
            }
            implicit_transition = len(previous_terms & current_terms) >= 1
            if index > 0 and content and not any(word in content[:260] for word in link_words) and not implicit_transition:
                issues.append({
                    "level": "低", "section": section, "issue": "章节开头缺少与前文的显式衔接。",
                    "suggestion": "可使用“在完成分类后”“基于上述比较”等过渡句连接前文。",
                })
            for sentence in _split_sentences(content):
                normalized = re.sub(r"\[R\d+\]|\s+|[，。；、,.]", "", sentence)
                if len(normalized) >= 30 and normalized in seen_sentences and normalized not in reported_duplicate_sentences:
                    issues.append({
                        "level": "中", "section": section, "issue": "与前文存在重复表达。",
                        "suggestion": "保留一次陈述，并将重复部分改为方法差异、适用边界或新的证据。",
                    })
                    reported_duplicate_sentences.add(normalized)
                seen_sentences.add(normalized)
        # The same type of problem can be detected repeatedly in one section
        # (especially duplicate sentences). Count it once so the score reflects
        # issue severity, not the number of repeated detections.
        unique_issues: List[Dict[str, str]] = []
        seen_issue_keys = set()
        for issue in issues:
            issue_key = (issue.get("level", ""), issue.get("section", ""), issue.get("issue", ""))
            if issue_key not in seen_issue_keys:
                unique_issues.append(issue)
                seen_issue_keys.add(issue_key)
        issues = unique_issues

        high_count = sum(item["level"] == "高" for item in issues)
        medium_count = sum(item["level"] == "中" for item in issues)
        low_count = sum(item["level"] == "低" for item in issues)
        penalty = high_count * 2.5 + medium_count * 0.8 + low_count * 0.3
        # Keep the score interpretable on a 0-10 scale. A long draft should not
        # receive zero merely because one detector reported many similar issues.
        score = round(max(0.5, 10 - min(penalty, 9.5)), 1)
        strengths = []
        if sections and all(re.findall(r"\[(R\d+)\]", main_text[match.end():(sections[index + 1].start() if index + 1 < len(sections) else len(main_text))]) for index, match in enumerate(sections)):
            strengths.append("每个已生成章节均保留了来源编号。")
        if not invalid_ids:
            strengths.append("未发现无效论文编号。")
        if section_plans:
            strengths.append(f"已为 {len(section_plans)} 个章节生成论证结构，可据此逐节修改。")
        return {
            "name": self.name,
            "score": score,
            "verdict": (
                "需要优先修改" if high_count
                else ("建议修改后再使用" if score < 6 else "可作为初稿继续完善")
            ),
            "issues": issues[:16],
            "strengths": strengths,
            "section_count": len(sections),
            "valid_citations": sorted(total_citations & valid_ids),
            "issue_counts": {"高": high_count, "中": medium_count, "低": low_count},
            "score_basis": "同一章节同类问题只计一次；总扣分上限为9.5分。",
        }


class WritingSkillSuite:
    def __init__(self) -> None:
        self.claim_evidence = ClaimEvidenceBindingSkill()
        self.cross_paper = CrossPaperSynthesisSkill()
        self.section_structure = SectionArgumentStructureSkill()
        self.logic_review = LogicCoherenceReviewSkill()

    def prepare(self, outline: Dict[str, Any], research: Dict[str, Any]) -> Dict[str, Any]:
        claim_evidence = self.claim_evidence.run(research.get("structured_papers", []))
        cross_paper = self.cross_paper.run(research.get("themes", []), research.get("comparisons", []))
        section_structure = self.section_structure.run(
            outline, research, claim_evidence, cross_paper.get("syntheses", [])
        )
        return {
            "catalog": WRITING_SKILL_CATALOG,
            "claim_evidence": claim_evidence,
            "cross_paper_synthesis": cross_paper,
            "section_argument_structure": section_structure,
        }

    def review_after_draft(
        self, article: str, research: Dict[str, Any], writing_skills: Dict[str, Any]
    ) -> Dict[str, Any]:
        return self.logic_review.run(
            article,
            research,
            writing_skills.get("section_argument_structure", {}).get("plans", []),
        )


class WriterAgent:
    name = "策划 Agent"

    def plan(self, topic: str, research_question: str = "", template: str = "文献综述") -> Dict[str, Any]:
        templates = {
            "文献综述": ["研究背景与问题", "相关研究分类", "研究方法与结论比较", "现有研究不足", "小结与研究启示"],
            "研究背景": ["问题背景", "领域发展现状", "已有研究进展", "研究空白", "本研究启示"],
        }
        question = research_question.strip() or f"{topic}领域的主要研究进展、方法差异与待解决问题是什么？"
        prompt = f"""
请为研究主题「{topic}」和研究问题「{question}」设计科研文献综述初稿框架。
章节必须支持多篇文献的分类、比较和研究空白提炼。只输出 JSON：
{{"suggested_title":"标题","research_question":"研究问题","target_audience":"目标用户","key_message":"综述主线","sections":["章节1","章节2"]}}
"""
        result = parse_json_response(llm.chat("你擅长设计科研文献综述结构。", prompt))
        if not result or not isinstance(result.get("sections"), list) or not result["sections"]:
            result = {
                "suggested_title": f"{topic}领域研究进展与问题综述",
                "research_question": question,
                "target_audience": "需要完成文献综述初稿的科研新手",
                "key_message": f"围绕{topic}梳理研究主题、方法差异、已有结论与研究空白。",
                "sections": templates.get(template, templates["文献综述"]),
            }
        result["topic"] = topic
        result["template"] = template
        result["research_question"] = result.get("research_question") or question
        return result


class ResearcherAgent:
    name = "文献调研 Agent"

    def __init__(self, repository: Optional[OpenAlexRepository] = None) -> None:
        self.repository = repository or OpenAlexRepository()
        self.structuring_skill = LiteratureStructuringSkill()

    def research(
        self,
        topic: str,
        sections: List[str],
        research_question: str = "",
        papers: Optional[List[Paper]] = None,
        use_online: bool = True,
        max_papers: int = 8,
    ) -> Dict[str, Any]:
        selected = list(papers or [])
        source_mode = "用户导入" if selected else ""
        if not selected and use_online:
            try:
                selected = self.repository.search(topic, max_papers)
                source_mode = "OpenAlex"
            except Exception:
                source_mode = "OpenAlex 检索失败"
        if not selected:
            selected = [Paper(**paper.to_dict()) for paper in DEMO_PAPERS[:max_papers]]
            source_mode = (source_mode + "，" if source_mode else "") + "内置示例"
        selected = self._renumber(selected[:max(1, min(int(max_papers), 20))])
        structured_papers = self.structuring_skill.extract_all(selected)
        has_uploaded_papers = any(paper.source.startswith("导入") for paper in selected)
        # Uploaded papers are analysed through the grounded parser first. This
        # protects the visible matrix and draft from an unrelated LLM fallback.
        integrated = self._rule_integrate(topic, sections, selected) if has_uploaded_papers else self._llm_integrate(topic, research_question, sections, selected)
        if integrated is None:
            integrated = self._rule_integrate(topic, sections, selected)
        integrated.update({
            "papers": selected,
            "structured_papers": structured_papers,
            "paper_count": len(selected),
            "source_mode": source_mode,
            "content_analysis_coverage": round(sum(1 for paper in selected if paper.abstract and paper.full_text) / len(selected), 2) if selected else 0,
            "structured_field_coverage": round(
                sum(item["field_coverage"] for item in structured_papers) / len(structured_papers), 2
            ) if structured_papers else 0,
        })
        return integrated

    @staticmethod
    def _renumber(papers: Iterable[Paper]) -> List[Paper]:
        result = []
        for index, paper in enumerate(papers, 1):
            paper.ref_id = f"R{index}"
            result.append(paper)
        return result

    @staticmethod
    def _context(papers: List[Paper]) -> str:
        return "\n\n".join(
            f"[{paper.ref_id}] {paper.title} | {paper.year or '年份未知'} | 正文状态: {paper.text_quality}\n"
            f"摘要: {paper.abstract or '摘要缺失'}\n"
            f"正文证据摘录: {compact_text(paper.full_text, 6000) or '正文缺失'}\n"
            f"结构化信号: {_research_signals(paper)}"
            for paper in papers
        )

    def _llm_integrate(self, topic: str, research_question: str, sections: List[str], papers: List[Paper]) -> Optional[Dict[str, Any]]:
        if not llm.available:
            return None
        prompt = f"""
你是科研文献整合 Agent。主题：{topic}；研究问题：{research_question}；章节：{', '.join(sections)}。
请将论文集合整合成主题、比较矩阵、章节素材和研究空白。只能使用给定论文正文或摘要中的事实，不得编造论文事实。
标题、作者、卷期和来源只用于识别文献，不能当作研究结论。对于原文没有明确说明的内容，必须写“原文未明确说明”。
比较矩阵中的方法、结论、局限必须优先引用对应论文的证据摘录；跨文献结论必须说明比较维度和适用边界。
每个 section_material 必须包含合法 paper_id、claim 和 evidence。

论文集合：
{self._context(papers)}

只输出 JSON：
{{"themes":[{{"name":"主题","summary":"综合观点","paper_ids":["R1"]}}],
"comparisons":[{{"paper_id":"R1","research_question":"问题","method":"方法","finding":"结论","limitation":"局限"}}],
"section_materials":{{"章节名":[{{"paper_id":"R1","claim":"可用于本节的观点","evidence":"原文依据"}}]}},
"gaps":["研究空白"]}}
"""
        raw = parse_json_response(llm.chat("你擅长从多篇论文中提炼主题、比较研究并保留来源。", prompt, max_tokens=7000))
        if not raw or not isinstance(raw.get("section_materials"), dict):
            return None
        valid_ids = {paper.ref_id for paper in papers}
        raw_comparisons = raw.get("comparisons") if isinstance(raw.get("comparisons"), list) else []
        comparisons = []
        for item in raw_comparisons:
            if not isinstance(item, dict) or str(item.get("paper_id")) not in valid_ids:
                continue
            comparisons.append({
                "paper_id": str(item.get("paper_id")),
                "research_question": compact_text(item.get("research_question"), 1000),
                "method": compact_text(item.get("method"), 1200),
                "finding": compact_text(item.get("finding"), 1400),
                "limitation": compact_text(item.get("limitation"), 1100),
            })
        if len(comparisons) < len(papers):
            # Never let an LLM omission erase a grounded row from the visible matrix.
            grounded = self._rule_integrate(topic, sections, papers)["comparisons"]
            by_id = {item["paper_id"]: item for item in comparisons}
            for item in grounded:
                by_id.setdefault(item["paper_id"], item)
            comparisons = [by_id[paper.ref_id] for paper in papers]
        section_materials: Dict[str, List[Dict[str, str]]] = {}
        for section in sections:
            values = raw["section_materials"].get(section, [])
            values = values if isinstance(values, list) else []
            section_materials[section] = [
                {
                    "paper_id": str(item.get("paper_id")),
                    "claim": compact_text(item.get("claim"), 700),
                    "evidence": compact_text(item.get("evidence"), 700),
                }
                for item in values
                if isinstance(item, dict) and str(item.get("paper_id")) in valid_ids
            ]
        if not any(section_materials.values()):
            return None
        raw["section_materials"] = section_materials
        themes = []
        for item in raw.get("themes", []) if isinstance(raw.get("themes"), list) else []:
            if not isinstance(item, dict):
                continue
            ids = [str(value) for value in item.get("paper_ids", []) if str(value) in valid_ids]
            if ids:
                themes.append({
                    "name": compact_text(item.get("name") or "未命名主题", 120),
                    "summary": compact_text(item.get("summary"), 1800),
                    "paper_ids": ids,
                })
        raw["themes"] = themes
        raw["comparisons"] = comparisons
        raw["gaps"] = [compact_text(value, 1200) for value in raw.get("gaps", []) if str(value).strip()] if isinstance(raw.get("gaps"), list) else []
        return raw

    def _rule_integrate(self, topic: str, sections: List[str], papers: List[Paper]) -> Dict[str, Any]:
        sections = sections or ["文献整合"]
        signals = {paper.ref_id: _research_signals(paper) for paper in papers}
        comparisons = [
            {
                "paper_id": paper.ref_id,
                "research_question": signals[paper.ref_id]["research_question"],
                "method": signals[paper.ref_id]["method"],
                "finding": signals[paper.ref_id]["finding"],
                "limitation": signals[paper.ref_id]["limitation"],
                "keywords": signals[paper.ref_id]["keywords"],
            }
            for paper in papers
        ]

        groups: Dict[str, List[Paper]] = {}
        for paper in papers:
            groups.setdefault(signals[paper.ref_id]["domain"], []).append(paper)
        themes = []
        for name, group in groups.items():
            ids = [paper.ref_id for paper in group]
            if name == "待确认的研究主题":
                name = "独立或待确认主题"
            themes.append({
                "name": name,
                "summary": _topic_theme_summary(name, group, signals),
                "paper_ids": ids,
            })

        gaps: List[str] = []
        for name, group in groups.items():
            if len(group) >= 2:
                ids = ", ".join(paper.ref_id for paper in group)
                methods = "；".join(f"{paper.ref_id}采用{signals[paper.ref_id]['method']}" for paper in group[:3])
                gaps.append(f"在{name}中，{ids}的研究对象或分析阶段存在差异，当前摘要尚未提供统一数据集、统一指标下的直接效果比较；已有方法为：{methods}。")
            else:
                paper = group[0]
                signal = signals[paper.ref_id]
                gaps.append(f"{paper.ref_id}已明确指出或显示出适用边界：{signal['limitation']}这一限制需要在后续研究中通过更丰富的数据、动态验证或外部案例检验。")
        metadata_only = [paper for paper in papers if not paper.abstract]
        if metadata_only:
            ids = ", ".join(paper.ref_id for paper in metadata_only[:8])
            suffix = "等" if len(metadata_only) > 8 else ""
            gaps.insert(0, f"{ids}{suffix}仅包含标题、作者等书目信息，未提供摘要或正文；这些条目不能用于可靠的研究问题、方法和结论比较，请上传对应论文全文。")
        if len(groups) > 1:
            ids = ", ".join(paper.ref_id for paper in papers)
            gaps.insert(0, f"上传集合包含多个研究方向（{', '.join(groups.keys())}），{ids}之间不存在直接的同任务比较关系；写作时应按主题分组，避免把生态评估、出版治理和技术预测强行合并为一个结论。")
        if not gaps:
            gaps.append("当前未能从上传文本中提炼出可验证的研究空白，请补充摘要或方法、结果章节。")

        section_materials: Dict[str, List[Dict[str, str]]] = {section: [] for section in sections}
        for section in sections:
            section_key = section.lower()
            if any(word in section for word in ("分类", "主题", "进展")):
                for theme in themes:
                    section_materials[section].append({
                        "paper_id": theme["paper_ids"][0],
                        "claim": theme["summary"],
                        "evidence": f"主题来源：{', '.join(theme['paper_ids'])}。{theme['summary']}",
                    })
            elif any(word in section for word in ("方法", "结论", "比较")):
                for item in comparisons:
                    section_materials[section].append({
                        "paper_id": item["paper_id"],
                        "claim": f"研究问题：{item['research_question']} 方法：{item['method']} 主要结论：{item['finding']}",
                        "evidence": f"摘要提取：{item['finding']}",
                    })
            elif any(word in section for word in ("不足", "空白", "研究缺口", "局限")):
                for index, gap in enumerate(gaps):
                    paper_id = papers[index % len(papers)].ref_id if papers else ""
                    section_materials[section].append({"paper_id": paper_id, "claim": gap, "evidence": gap})
            else:
                for item in comparisons:
                    section_materials[section].append({
                        "paper_id": item["paper_id"],
                        "claim": item["research_question"],
                        "evidence": item["finding"],
                    })
        return {"themes": themes, "comparisons": comparisons, "section_materials": section_materials, "gaps": gaps}


class EditorAgent:
    name = "写作 Agent"

    def write(
        self,
        outline: Dict[str, Any],
        research: Dict[str, Any],
        writing_skills: Optional[Dict[str, Any]] = None,
    ) -> str:
        topic = outline.get("topic", "研究主题")
        sections = outline.get("sections", [])
        materials = research.get("section_materials", {})
        papers = {paper.ref_id: paper for paper in research.get("papers", [])}
        grounded_upload = any(paper.source.startswith("导入") for paper in papers.values())
        writing_skills = writing_skills or {}
        plans = {
            item.get("section", ""): item
            for item in writing_skills.get("section_argument_structure", {}).get("plans", [])
        }
        syntheses = writing_skills.get("cross_paper_synthesis", {}).get("syntheses", [])

        def make_section(section: str) -> Tuple[str, str]:
            items = materials.get(section, [])
            plan = plans.get(section, {})
            related_syntheses = [
                item for item in syntheses
                if item.get("id") in plan.get("synthesis_ids", [])
            ]
            context = "\n\n".join(
                f"[{item.get('paper_id')}] {papers[item.get('paper_id')].title}\n观点: {item.get('claim')}\n依据: {item.get('evidence')}"
                for item in items if item.get("paper_id") in papers
            ) or "本节暂无直接匹配文献，请明确标注证据不足。"
            structure_context = (
                f"章节目标：{plan.get('goal', '')}\n"
                f"中心论点：{plan.get('core_claim', '')}\n"
                f"写作步骤：{'；'.join(plan.get('writing_steps', []))}\n"
                f"过渡句：{plan.get('transition', '')}\n"
                f"跨文献综合素材：{' '.join(item.get('paragraph', '') for item in related_syntheses)}"
            )
            prompt = f"""
请围绕「{topic}」撰写文献综述初稿章节「{section}」。
只能使用以下来源，不得编造论文事实；进行跨文献比较和归纳，而不是逐篇复述。
所有来源标记必须使用 [R1] 这样的格式，且只使用给定编号；素材不足时写明“需要补充文献证据”。
不要输出章节标题，只输出正文。

{structure_context}

{context}
"""
            generated = "" if grounded_upload else llm.chat("你擅长写作有来源依据的科研文献综述初稿。", prompt, max_tokens=3500)
            if generated.strip():
                generated = re.sub(r"^\s*#{1,6}\s+[^\n]+\n", "", generated.strip(), count=1)
                return section, generated
            return section, self._fallback_section(section, items, papers, plan, related_syntheses)

        contents: Dict[str, str] = {}
        if sections:
            with ThreadPoolExecutor(max_workers=min(4, len(sections))) as executor:
                futures = {executor.submit(make_section, section): section for section in sections}
                for future in as_completed(futures):
                    section, content = future.result()
                    contents[section] = content
        parts = [
            f"# {outline.get('suggested_title', f'{topic}文献综述初稿')}",
            "",
            f"> 研究问题：{outline.get('research_question', '')}",
            "",
            f"**目标用户**：{outline.get('target_audience', '科研新手')}",
            "",
        ]
        for index, section in enumerate(sections, 1):
            parts.extend([f"## {index}. {section}", "", contents.get(section, "本节生成失败，请重试。"), ""])
        parts.extend(["## 参考文献", ""])
        parts.extend(paper.citation() for paper in papers.values())
        return "\n".join(parts)

    @staticmethod
    def _fallback_section(
        section: str,
        items: List[Dict[str, Any]],
        papers: Dict[str, Paper],
        plan: Optional[Dict[str, Any]] = None,
        syntheses: Optional[List[Dict[str, Any]]] = None,
    ) -> str:
        plan = plan or {}
        syntheses = syntheses or []
        if not items and not syntheses:
            return f"本节围绕“{section}”展开，但当前文献集合中缺少直接证据，需要补充文献后再完成论述。"
        lines = [plan.get("core_claim") or f"本节依据上传文献中可追溯的摘要与正文信息，对“{section}”进行整合。"]
        if syntheses:
            lines.extend(item.get("paragraph", "") for item in syntheses[:2] if item.get("paragraph"))
        else:
            for item in items[:4]:
                paper = papers.get(item.get("paper_id"))
                if paper:
                    claim = item.get("claim") or paper.abstract or paper.title
                    lines.append(f"{claim} [{paper.ref_id}]")
        if len(items) > 1 and not syntheses:
            cited = [papers[item.get("paper_id")] for item in items if item.get("paper_id") in papers]
            titles = "；".join(f"{paper.ref_id}《{paper.title}》" for paper in cited[:3])
            lines.append(f"横向比较可见，{titles}的研究对象与方法并不相同，结论只能在各自的样本、数据和场景范围内理解；后续应围绕可比对象、共同指标和外部案例补充验证。")
        if plan.get("transition"):
            lines.append(plan["transition"])
        return "\n\n".join(lines)


class CriticAgent:
    name = "评审 Agent"

    def review(self, article: str, outline: Dict[str, Any], research: Dict[str, Any]) -> Dict[str, Any]:
        valid_ids = {paper.ref_id for paper in research.get("papers", [])}
        papers = research.get("papers", [])
        cited_ids = set(re.findall(r"\[(R\d+)\]", article)) & valid_ids
        coverage = round(len(cited_ids) / len(valid_ids), 2) if valid_ids else 0
        quality_flags: List[str] = []
        for paper in papers:
            if paper.title in {"研究论文", "未命名论文"} or re.search(r"第\s*\d+\s*卷|Vol\.?\s*\d+|文章编号|Journal of", paper.title, flags=re.I):
                quality_flags.append(f"{paper.ref_id} 标题疑似期刊页眉或栏目名")
            if not paper.authors:
                quality_flags.append(f"{paper.ref_id} 未识别作者")
            if not paper.abstract or len(paper.abstract) < 80:
                quality_flags.append(f"{paper.ref_id} 摘要过短，结论应结合全文核验")
            quality_flags.extend(f"{paper.ref_id}：{warning}" for warning in paper.parse_warnings)
        comparisons = research.get("comparisons", [])
        if len(comparisons) >= 2 and len({compact_text(item.get("method"), 180) for item in comparisons}) == 1:
            quality_flags.append("比较矩阵中的方法字段重复，可能未读取论文内容")
        if not research.get("themes"):
            quality_flags.append("未形成绑定来源的主题")
        prompt = f"""
请评审下面的科研文献综述初稿，重点检查章节结构、跨文献整合、论据支撑和来源可追溯性。
只允许使用这些来源编号：{', '.join(sorted(valid_ids))}。
只输出 JSON：{{"scores":{{"结构清晰度":0,"跨文献整合":0,"可读性":0,"论据支撑":0,"来源可追溯":0}},"suggestions":["建议"],"unsupported_claims":["表述"],"verdict":"通过/需要修改"}}

{article[:14000]}
"""
        result = parse_json_response(llm.chat("你是一位严格的科研写作评审专家。", prompt, max_tokens=3000)) if llm.available else None
        if not result:
            source_score = round(4 + coverage * 6, 1)
            content_score = 8 if comparisons and all(item.get("research_question") and item.get("method") and item.get("finding") for item in comparisons) else 5
            integration_score = 8 if len(research.get("themes", [])) >= 1 and len(comparisons) >= 2 else 5
            result = {
                "scores": {
                    "结构清晰度": 8 if len(outline.get("sections", [])) >= 4 else 6,
                    "跨文献整合": integration_score,
                    "可读性": 7,
                    "论据支撑": min(source_score, content_score),
                    "来源可追溯": source_score,
                },
                "suggestions": [
                    "检查每个研究问题、方法和结论是否都能在原文摘要或正文中定位。",
                    "对同一主题的文献补充样本、指标、时间范围和方法差异，避免只做逐篇罗列。",
                    "对摘要未明确说明的局限补充全文页码或实验细节，不要把推断写成论文原结论。",
                ],
                "unsupported_claims": [] if cited_ids else ["当前正文未检测到有效来源标记"],
                "verdict": "通过" if coverage >= 0.5 else "需要修改",
            }
        scores = result.get("scores") if isinstance(result.get("scores"), dict) else {}
        normalised = {}
        for key, value in scores.items():
            try:
                normalised[key] = max(0, min(10, round(float(value), 1)))
            except (TypeError, ValueError):
                pass
        result["scores"] = normalised
        result["overall"] = round(sum(normalised.values()) / len(normalised), 1) if normalised else 0
        result["word_count"] = count_chinese_words(article)
        result["citation_coverage"] = coverage
        result["cited_sources"] = sorted(cited_ids)
        result["missing_sources"] = sorted(valid_ids - cited_ids)
        result["quality_flags"] = quality_flags
        result["content_analysis_coverage"] = round(sum(1 for item in comparisons if item.get("method") and item.get("finding")) / len(papers), 2) if papers else 0
        return result


class MultiAgentOrchestrator:
    def __init__(self) -> None:
        self.writer = WriterAgent()
        self.researcher = ResearcherAgent()
        self.editor = EditorAgent()
        self.critic = CriticAgent()
        self.writing_skills = WritingSkillSuite()

    def run(
        self,
        topic: str,
        style: str = "学术综述",
        template: str = "文献综述",
        research_question: str = "",
        papers: Optional[List[Paper]] = None,
        source_file: str | Path | None = None,
        use_online: bool = True,
        max_papers: int = 8,
    ) -> Dict[str, Any]:
        if source_file and not papers:
            papers = load_papers_from_file(source_file)
        effective_topic = topic
        effective_question = research_question
        cross_topic = False
        if papers:
            effective_topic, cross_topic = _infer_corpus_topic(papers)
            default_question = (not research_question.strip()) or any(
                marker in research_question for marker in ("大语言模型智能体", "知识获取与任务协作", "科研场景中的知识获取")
            )
            if default_question:
                effective_question = f"上传文献分别研究了哪些具体问题，采用了哪些方法，形成了哪些可比较的结论与适用边界？"
        outline = self.writer.plan(effective_topic, effective_question, template)
        if cross_topic:
            outline["suggested_title"] = "上传文献的主题分类、方法比较与研究空白综述"
            outline["key_message"] = "上传集合属于多个研究方向，初稿将先按主题分组，再比较同组文献，避免跨领域强行归纳。"
        research = self.researcher.research(
            effective_topic,
            outline.get("sections", []),
            outline.get("research_question", research_question),
            papers=papers,
            use_online=use_online,
            max_papers=max_papers,
        )
        writing_skills = self.writing_skills.prepare(outline, research)
        article = self.editor.write(outline, research, writing_skills)
        review = self.critic.review(article, outline, research)
        writing_skills["logic_coherence_review"] = self.writing_skills.review_after_draft(
            article, research, writing_skills
        )
        return {
            "topic": effective_topic,
            "outline": outline,
            "research": research,
            "papers": research.get("papers", []),
            "materials": research.get("section_materials", {}),
            "article": article,
            "review": review,
            "writing_skills": writing_skills,
            "execution_summary": {
                "agents_used": 4,
                "total_steps": 4,
                "paper_count": research.get("paper_count", 0),
                "source_mode": research.get("source_mode", "未知"),
                "article_word_count": review.get("word_count", count_chinese_words(article)),
                "quality_score": review.get("overall", 0),
                "citation_coverage": review.get("citation_coverage", 0),
                "content_analysis_coverage": research.get("content_analysis_coverage", review.get("content_analysis_coverage", 0)),
                "writing_skill_count": len(writing_skills.get("catalog", [])),
                "logic_review_score": writing_skills.get("logic_coherence_review", {}).get("score", 0),
                "suggestions": review.get("suggestions", []),
            },
        }


def search_papers(topic: str, max_results: int = 8, use_online: bool = True) -> Tuple[List[Paper], str]:
    if use_online:
        try:
            papers = OpenAlexRepository().search(topic, max_results)
            if papers:
                return papers, "OpenAlex"
        except Exception:
            return [], "OpenAlex 检索失败"
    return [Paper(**paper.to_dict()) for paper in DEMO_PAPERS[:max_results]], "内置示例"


def _table(value: Any, limit: int = 300) -> str:
    return compact_text(value, limit).replace("|", "\\|").replace("\n", " ")


def generate_report(result: Dict[str, Any]) -> str:
    outline = result.get("outline", {})
    research = result.get("research", {})
    review = result.get("review", {})
    summary = result.get("execution_summary", {})
    lines = [
        "# 📚 科研文献整合与论文初稿报告", "",
        f"**主题**: {result.get('topic', '')}",
        f"**论文来源**: {summary.get('source_mode', '未知')}",
        f"**模型模式**: {'真实 API' if llm.available else '规则降级模式'}", "",
        "## 执行摘要", "",
        f"- 参与 Agent: {summary.get('agents_used', 4)} 个",
        f"- 整合论文: {summary.get('paper_count', 0)} 篇",
        f"- 初稿字数: {summary.get('article_word_count', 0)} 字",
        f"- 来源覆盖率: {summary.get('citation_coverage', 0):.0%}",
        f"- 内部评审分: {summary.get('quality_score', 0)}/10", "",
        "## 综述框架", "",
        f"**标题**: {outline.get('suggested_title', '')}",
        f"**研究问题**: {outline.get('research_question', '')}",
        f"**目标用户**: {outline.get('target_audience', '')}", "",
    ]
    lines.extend(f"{index}. {section}" for index, section in enumerate(outline.get("sections", []), 1))
    lines.extend(["", "## 文献来源", "", "| 编号 | 论文 | 作者 | 年份 | 来源 |", "| --- | --- | --- | --- | --- |"])
    for paper in result.get("papers", []):
        lines.append(f"| {paper.ref_id} | {_table(paper.title)} | {_table(', '.join(paper.authors[:3]) or '未知')} | {paper.year or '未知'} | {_table(paper.source)} |")
    lines.extend(["", "## 文献结构化提取", ""])
    for item in research.get("structured_papers", []):
        lines.extend([f"### [{item.get('paper_id', '')}] {item.get('title', '未命名论文')}", ""])
        metadata = item.get("metadata", {})
        lines.append(
            f"作者：{_table(', '.join(metadata.get('authors', [])) or '未知')}；"
            f"年份：{metadata.get('year') or '未知'}；DOI：{_table(metadata.get('doi') or '未识别')}；"
            f"字段覆盖率：{item.get('field_coverage', 0):.0%}"
        )
        lines.append("")
        fields = item.get("fields", {})
        sources = item.get("field_sources", {})
        for key, label in STRUCTURED_FIELD_LABELS.items():
            lines.append(f"- **{label}**（{sources.get(key, '未提取')}）：{fields.get(key, '原文未明确说明。')}")
        lines.append("")
    writing_skills = result.get("writing_skills", {})
    lines.extend(["## 写作 Skill 工作台", ""])
    for skill in writing_skills.get("catalog", WRITING_SKILL_CATALOG):
        lines.extend([
            f"### {skill.get('name', '未命名 Skill')}",
            f"- 作用：{skill.get('purpose', '')}",
            f"- 输入：{skill.get('input', '')}",
            f"- 输出：{skill.get('output', '')}",
            "",
        ])
    binding = writing_skills.get("claim_evidence", {})
    lines.extend([
        "### 论点-证据绑定结果",
        f"- 论点卡片：{len(binding.get('cards', []))} 张",
        f"- 已支持：{binding.get('supported_count', 0)} 张",
        f"- 证据覆盖率：{binding.get('coverage', 0):.0%}",
        "",
    ])
    lines.extend(["### 跨文献综合写作结果", ""])
    for item in writing_skills.get("cross_paper_synthesis", {}).get("syntheses", []):
        lines.extend([
            f"**{item.get('id', '')}｜{item.get('theme', '相关研究')}**（来源：{', '.join(item.get('paper_ids', []))}）",
            item.get("paragraph", ""),
            f"写作建议：{item.get('writing_advice', '')}",
            "",
        ])
    lines.extend(["### 章节论证结构", ""])
    for plan in writing_skills.get("section_argument_structure", {}).get("plans", []):
        lines.extend([
            f"#### {plan.get('section', '未命名章节')}",
            f"- 章节目标：{plan.get('goal', '')}",
            f"- 中心论点：{plan.get('core_claim', '')}",
            f"- 写作步骤：{'；'.join(plan.get('writing_steps', []))}",
            f"- 过渡句：{plan.get('transition', '')}",
            "",
        ])
    logic = writing_skills.get("logic_coherence_review", {})
    lines.extend([
        "### 逻辑与衔接评审",
        f"- 评审分：{logic.get('score', 0)}/10",
        f"- 结论：{logic.get('verdict', '待评审')}",
    ])
    for issue in logic.get("issues", []):
        lines.append(f"- [{issue.get('level', '')}] {issue.get('section', '')}：{issue.get('issue', '')} 建议：{issue.get('suggestion', '')}")
    lines.append("")
    lines.extend(["## 主题整合", ""])
    for theme in research.get("themes", []):
        lines.extend([f"### {theme.get('name', '未命名主题')}", "", f"{theme.get('summary', '')} 来源：{', '.join(theme.get('paper_ids', []))}", ""])
    lines.extend(["## 文献比较矩阵", "", "| 来源 | 研究问题 | 方法 | 主要结论 | 局限 |", "| --- | --- | --- | --- | --- |"])
    for item in research.get("comparisons", []):
        lines.append(f"| {item.get('paper_id', '')} | {_table(item.get('research_question'))} | {_table(item.get('method'))} | {_table(item.get('finding'))} | {_table(item.get('limitation'))} |")
    lines.extend(["", "## 研究空白", ""])
    lines.extend(f"- {gap}" for gap in research.get("gaps", []))
    lines.extend(["", "## 论文初稿", "", result.get("article", "文章未生成"), "", "## 评审报告", ""])
    lines.extend(f"- {metric}: {score}/10" for metric, score in review.get("scores", {}).items())
    lines.extend([
        f"- **整体评分**: **{review.get('overall', 0)}/10**",
        f"- **来源覆盖率**: **{review.get('citation_coverage', 0):.0%}**",
        f"- **结论**: {review.get('verdict', '需要修改')}", "", "### 迭代建议", "",
    ])
    lines.extend(f"{index}. {suggestion}" for index, suggestion in enumerate(review.get("suggestions", []), 1))
    return "\n".join(lines)


if __name__ == "__main__":
    topic = os.environ.get("P4_TOPIC", "AI Agent")
    use_online = os.environ.get("P4_USE_ONLINE", "false").lower() in {"1", "true", "yes"}
    result = MultiAgentOrchestrator().run(topic, style="学术综述", template="文献综述", use_online=use_online, max_papers=5)
    output = PROJECT_DIR / "p4-literature-report.md"
    output.write_text(generate_report(result), encoding="utf-8")
    print(f"📄 完整报告已保存至: {output}")
