"""Extract knowledge cards from the General Zoology docx."""
from __future__ import annotations

import json
import re
from pathlib import Path
from zipfile import ZipFile

BASE_DIR = Path(__file__).resolve().parent.parent
DOCX_GLOB = "11*.docx"
OUT_JSON = BASE_DIR / "static" / "zoology_cards_extracted.json"

CHAPTER_TITLES = {
    1: "绪论",
    2: "动物体的基本结构与机能",
    3: "原生动物门",
    4: "多细胞动物的起源",
    5: "海绵动物门",
    6: "腔肠动物门",
    7: "扁形动物门",
    8: "假体腔动物",
    9: "环节动物门",
    10: "软体动物门",
    11: "节肢动物门",
    12: "触手冠动物",
    13: "棘皮动物门",
    14: "半索动物门",
    15: "脊索动物门",
    16: "圆口纲",
    17: "鱼纲",
    18: "两栖纲",
    19: "爬行纲",
    20: "鸟纲",
    21: "哺乳纲",
    22: "动物进化基本原理",
    23: "动物地理",
    24: "动物生态",
}

NOISE = re.compile(
    r"大学专业|1对1|QQ\s*:?\s*1359097968|微信|13220409168|"
    r"兰州课真题|高等教育出版社|landraco|Zoology|版权所有|ISBN|CIP|"
    r"策划编辑|责任编辑|封面设计|版式设计|责任校对|"
    r"www\.|http://|\.com|\.cn|\.edu|购书热线|防伪查询"
)
SKIP_TITLE = re.compile(r"[⋯…\.]{3,}|^\d+$")
SEC_HEAD = re.compile(
    r"(?<![\d.])(\d{1,2})\.(\d{1,2})\s*"
    r"([\u4e00-\u9fffA-Za-z（）()·、，：；\-\s]{4,55}?)"
    r"(?=\d{1,2}\.\d{1,2}\s*[\u4e00-\u9fff]|第\d{1,2}章|$)"
)


def clean(text: str) -> str:
    text = NOISE.sub("", text)
    text = re.sub(r"([\u4e00-\u9fff])\1+", r"\1", text)
    text = re.sub(r"\.{2,}|…+|⋯+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def load_docx_text() -> str:
    docx = next(
        p for p in BASE_DIR.glob(DOCX_GLOB) if not p.name.startswith("~$")
    )
    with ZipFile(docx) as zf:
        xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
    tags = re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml)
    return clean("".join(tags))


def body_slice(raw: str) -> str:
    """Skip front matter / TOC; start near first in-body chapter paragraph."""
    markers = [
        m
        for m in re.finditer(r"第1章.{0,12}绪论", raw)
        if "⋯" not in m.group() and "…" not in m.group()
    ]
    if markers:
        return raw[markers[0].start() :]
    for m in re.finditer(r"1\.1\s*生物的分界", raw):
        return raw[max(0, m.start() - 80) :]
    return raw


def extract_points(content: str) -> list[str]:
    patterns = [
        r"[^。]{6,45}(?:是指|定义为|即指|称为|叫做)[^。]{6,90}。",
        r"[^。]{4,35}(?:主要特征|共同特征|特点|基本特征)[^。]{8,110}。",
        r"[^。]{6,30}(?:包括|具有|由[^。]{4,55}构成|属于)[^。]{8,90}。",
        r"[^。]{6,40}(?:适应|功能|作用|意义)[^。]{8,90}。",
    ]
    points: list[str] = []
    for pat in patterns:
        points.extend(re.findall(pat, content))
    points = [clean(p) for p in points if 12 < len(clean(p)) < 200]
    deduped = list(dict.fromkeys(points))
    if deduped:
        return deduped[:5]
    sents = re.split(r"[。！？]", content)
    return [
        s.strip() + "。"
        for s in sents
        if 18 < len(s.strip()) < 100 and not re.search(r"\d{3,}", s)
    ][:4]


def build_cards(raw: str) -> list[dict]:
    text = body_slice(raw)
    matches = list(SEC_HEAD.finditer(text))
    by_chapter: dict[int, list[dict]] = {i: [] for i in range(1, 25)}
    seen: set[str] = set()

    for i, m in enumerate(matches):
        ch_n = int(m.group(1))
        sec_n = int(m.group(2))
        if ch_n < 1 or ch_n > 24 or sec_n > 30:
            continue
        sec_id = f"{ch_n}.{sec_n}"
        title = clean(m.group(3))
        if SKIP_TITLE.search(title) or len(title) < 4:
            continue
        if sec_id in seen:
            continue

        end = matches[i + 1].start() if i + 1 < len(matches) else m.end() + 2200
        content = clean(text[m.end() : end])
        if len(content) < 40:
            continue
        points = extract_points(content)
        if len(points) < 2:
            continue
        seen.add(sec_id)
        by_chapter[ch_n].append(
            {"id": sec_id, "title": title[:56], "points": points[:4]}
        )

    cards: list[dict] = []
    for ch in range(1, 25):
        sections = by_chapter[ch][:10]
        if not sections:
            continue
        name = CHAPTER_TITLES[ch]
        cards.append(
            {
                "chapter": ch,
                "title": f"第{ch}章 {name}",
                "summary": sections[0]["points"][0] if sections[0]["points"] else "",
                "sections": sections,
            }
        )
    return cards


def main() -> None:
    raw = load_docx_text()
    cards = build_cards(raw)
    OUT_JSON.parent.mkdir(exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(cards, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    total_secs = sum(len(c["sections"]) for c in cards)
    print(f"Wrote {len(cards)} chapters, {total_secs} sections -> {OUT_JSON}")


if __name__ == "__main__":
    main()
