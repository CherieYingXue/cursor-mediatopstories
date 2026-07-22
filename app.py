"""Russia Media Reports — Flask app.

只做一件事：抓俄语原文站的 RSS，按 5 大板块（政治/经济/社会/普京/梅金斯基）归类，
把标题翻译成中文，并在 /russia 页面上呈现，支持打开自动更新 + 手动更新按钮。

原始 media top stories 项目的域名清单、日程调度器、Excel 导入、Zoology 卡片、
IBO 小测验等一切无关逻辑都不在本项目中；见 README。
"""

from __future__ import annotations

import concurrent.futures
import datetime as dt
import os
import time
from pathlib import Path
from typing import Any

import feedparser
import requests
from flask import Flask, jsonify, redirect, request, send_from_directory

BASE_DIR = Path(__file__).parent

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev-secret-change-in-production")

# ---------------------------------------------------------------------------
# 抓取俄媒 RSS 时使用的 UA 与请求头（模仿 Android + Chrome Mobile，减少被拒率）
# ---------------------------------------------------------------------------
MOBILE_HTTP_USER_AGENT = (
    "Mozilla/5.0 (Linux; Android 15; SM-S928B) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/133.0.0.0 Mobile Safari/537.36"
)
MOBILE_REQUEST_HEADERS = {
    "User-Agent": MOBILE_HTTP_USER_AGENT,
    "Accept-Language": "ru,en;q=0.9,zh;q=0.8",
    "Sec-CH-UA": '"Google Chrome";v="133", "Chromium";v="133", "Not_A Brand";v="24"',
    "Sec-CH-UA-Mobile": "?1",
    "Sec-CH-UA-Platform": '"Android"',
}


# ---------------------------------------------------------------------------
# 路由：页面
# ---------------------------------------------------------------------------
@app.route("/", methods=["GET"])
def root():
    """打开根路径直接跳到 /russia（可用 LANDING_PATH 覆盖）。"""
    landing = os.getenv("LANDING_PATH", "/russia").strip() or "/russia"
    return redirect(landing)


@app.route("/russia")
def russia_page():
    """俄罗斯主流媒体重点新闻聚合页（俄语标题 + 中文翻译）。"""
    return send_from_directory(BASE_DIR / "static", "russia-top-stories-cn.html")


@app.route("/healthz")
def healthz():
    """给 Render / uptime 用的健康检查端点。"""
    return jsonify({"ok": True, "service": "russia-media-reports"})


# ---------------------------------------------------------------------------
# /api/russia-news : 并行抓 RSS + 关键词归类 + 谷歌无鉴权翻译
# ---------------------------------------------------------------------------
RUSSIA_FEEDS: list[tuple[str, str]] = [
    # 通讯社 & 主流报刊
    ("ТАСС", "https://tass.ru/rss/v2.xml"),
    ("РИА Новости", "https://ria.ru/export/rss2/archive/index.xml"),
    ("РТ на русском", "https://russian.rt.com/rss"),
    ("Lenta.ru", "https://lenta.ru/rss/news"),
    ("Российская газета", "https://rg.ru/xml/index.xml"),
    ("Коммерсантъ", "https://www.kommersant.ru/RSS/news.xml"),
    ("Известия", "https://iz.ru/xml/rss/all.xml"),
    ("Взгляд", "https://vz.ru/rss.xml"),
    ("Kremlin.ru", "http://kremlin.ru/events/all/feed"),
    # 智库 & 政策分析
    ("Валдайский клуб", "https://ru.valdaiclub.com/export/rss/feed.xml"),
    ("Международная жизнь", "https://interaffairs.ru/rss/"),
    ("РСМД", "https://russiancouncil.ru/rss/all/"),
    (
        "Carnegie Russia-Eurasia",
        "https://news.google.com/rss/search?"
        "q=site:carnegieendowment.org+(russia+OR+eurasia+OR+putin+OR+kremlin)"
        "&hl=en-US&gl=US&ceid=US:en",
    ),
]

# 重点专家：标题或作者中出现任何一个即进入"重点专家"板块（最高优先级）
EXPERT_KEYWORDS: tuple[str, ...] = (
    "лузянин", "luzyanin",
    "маслов", "maslov",
    "кортунов", "kortunov",
    "кашин", "kashin",
)

# 源优先级：智库最高，官方媒体次之，通讯社最后。用于每个板块内部排序，
# 让分析类文章在通讯社快讯挤满 8 个位置之前先被选中。
SOURCE_PRIORITY: dict[str, int] = {
    "Валдайский клуб": 0,
    "Международная жизнь": 0,
    "РСМД": 0,
    "Carnegie Russia-Eurasia": 0,
    "Kremlin.ru": 1,
    "Российская газета": 1,
    "Известия": 1,
    "Взгляд": 1,
    "Коммерсантъ": 2,
    "ТАСС": 3,
    "РИА Новости": 3,
    "РТ на русском": 3,
    "Lenta.ru": 3,
}

RUSSIA_CACHE_TTL = int(os.getenv("RUSSIA_CACHE_TTL", "600"))  # 秒；默认 10 分钟
_russia_cache: dict[str, Any] = {"ts": 0.0, "data": None}

_POLITICS_KW = (
    "мид ", "госдум", "цик", "санкц", "нато", "всу", "сво", "переговор",
    "дипломат", "захаров", "песков", "лавров", "кремл", "парламент",
    "выбор", "оборон", "мигрант", "гражданств", "экстремист", "фсб",
    "теракт", "война", "министр", "закон", "депутат", "украин", "зеленск",
    "трамп", "байден", "белорусс", "лукашенк", "балт",
)
_ECONOMY_KW = (
    "рубл", "доллар", "евро", "юан", "цб ", "центробанк", "инфляц", "ввп",
    "экономик", "цена", "бизнес", "налог", "банк", "нефт", "газ", "экспорт",
    "импорт", "бирж", "рынок", "топлив", "бензин", "дизель", "тариф",
    "инвестиц", "торгов", "прибыл", "убыт", "промышленн", "компани",
    "wildberries", "яндекс", "росатом", "газпром", "лукойл", "сбер",
)
_SOCIAL_KW = (
    "обществ", "паводок", "потоп", "затопил", "погиб", "ранен", "пожар",
    "авари", "жерт", "школ", "медицин", "здоров", "спорт", "культур",
    "происшеств", "уголов", "склад", "чс", "спас", "траге",
    "наводнен", "дтп",
)


def categorize_ru(title: str, author: str = "", summary: str = "") -> str:
    """归类：重点专家 > 梅金斯基 > 普京 > 政治 > 经济 > 社会 > other。

    重点专家匹配范围包括标题、作者、以及正文简介前 500 字，
    因为像 РСМД、Международная жизнь 的 RSS 通常不带 author 字段，
    专家的姓氏往往只出现在文章简介或者副标题里。"""
    expert_haystack = " ".join([title, author or "", (summary or "")[:500]]).lower()
    for kw in EXPERT_KEYWORDS:
        if kw in expert_haystack:
            return "experts"
    t = title.lower()
    if "медински" in t:
        return "medinsky"
    if "путин" in t or "владимир владимирович" in t:
        return "putin"
    for kw in _POLITICS_KW:
        if kw in t:
            return "politics"
    for kw in _ECONOMY_KW:
        if kw in t:
            return "economy"
    for kw in _SOCIAL_KW:
        if kw in t:
            return "social"
    return "other"


def _fetch_feed(name_url: tuple[str, str]) -> list[dict[str, Any]]:
    """抓一路 RSS。Google News RSS 的标题末尾会被清洗；对个别拒绝 CH-UA 头的源退化到桌面 UA。"""
    name, url = name_url
    parsed = None
    for headers in (MOBILE_REQUEST_HEADERS, {"User-Agent": "Mozilla/5.0", "Accept": "*/*"}):
        try:
            resp = requests.get(url, headers=headers, timeout=12)
            if resp.status_code != 200:
                continue
            parsed = feedparser.parse(resp.content)
            if getattr(parsed, "entries", None):
                break
        except Exception:
            continue
    if parsed is None or not getattr(parsed, "entries", None):
        return []

    is_google_news = "news.google.com" in url
    out: list[dict[str, Any]] = []
    for entry in parsed.entries[:25]:
        title = (entry.get("title") or "").strip()
        link = entry.get("link") or ""
        author = entry.get("author") or entry.get("dc_creator") or ""
        published = entry.get("published") or entry.get("updated") or ""
        summary = entry.get("summary") or entry.get("description") or ""

        # 对 Google News 包装做清理：标题末尾 " - Carnegie Endowment for..." 去掉
        if is_google_news and " - " in title:
            title = title.rsplit(" - ", 1)[0].strip()

        if title and link:
            out.append(
                {
                    "title": title,
                    "url": link,
                    "source": name,
                    "published": published,
                    "author": author,
                    "summary": summary,
                }
            )
    return out


def _translate_to_zh(text: str) -> str:
    """谷歌无鉴权翻译端点。用 sl=auto 兼容俄语原标题与 Carnegie 的英文标题。"""
    if not text:
        return ""
    try:
        resp = requests.get(
            "https://translate.googleapis.com/translate_a/single",
            params={
                "client": "gtx",
                "sl": "auto",
                "tl": "zh-CN",
                "dt": "t",
                "q": text,
            },
            headers={"User-Agent": MOBILE_HTTP_USER_AGENT},
            timeout=6,
        )
        resp.raise_for_status()
        data = resp.json()
        return "".join(seg[0] for seg in data[0] if seg and seg[0]).strip()
    except Exception:
        return ""


@app.route("/api/russia-news")
def api_russia_news():
    """并行抓 8 家俄语原文 RSS + 归类 + 翻译；结果 10 分钟内存缓存。"""
    refresh = request.args.get("refresh") == "1"
    now = time.time()
    if (
        not refresh
        and _russia_cache["data"]
        and (now - _russia_cache["ts"]) < RUSSIA_CACHE_TTL
    ):
        payload = dict(_russia_cache["data"])
        payload["cached"] = True
        payload["age_seconds"] = int(now - _russia_cache["ts"])
        return jsonify(payload)

    with concurrent.futures.ThreadPoolExecutor(max_workers=max(8, len(RUSSIA_FEEDS))) as ex:
        fetched = list(ex.map(_fetch_feed, RUSSIA_FEEDS))
    all_items = [it for lst in fetched for it in lst]

    seen: set[str] = set()
    unique: list[dict[str, Any]] = []
    for it in all_items:
        if it["url"] and it["url"] not in seen:
            seen.add(it["url"])
            unique.append(it)

    buckets: dict[str, list[dict[str, Any]]] = {
        "politics": [],
        "economy": [],
        "social": [],
        "putin": [],
        "medinsky": [],
        "experts": [],
    }
    # 分类并记录每条新闻在**同一 source 内**的序号，用于轮询排序。
    per_source_seen: dict[str, int] = {}
    for idx, it in enumerate(unique):
        cat = categorize_ru(it["title"], it.get("author", ""), it.get("summary", ""))
        pos_in_source = per_source_seen.get(it["source"], 0)
        per_source_seen[it["source"]] = pos_in_source + 1
        # 排序键：(优先级 tier, 该源在这个板块里的第 N 条, 全局 idx)
        # → 效果：先取每个高优先级源的第 1 条，再取第 2 条…
        #        这样智库分析 + 主流媒体快讯自然交错，不会某一家把整个板块占满。
        it["_rank"] = (SOURCE_PRIORITY.get(it["source"], 9), pos_in_source, idx)
        if cat in buckets:
            buckets[cat].append(it)

    LIMITS = {
        "politics": 8, "economy": 6, "social": 6,
        "putin": 6, "medinsky": 4, "experts": 6,
    }
    for cat, limit in LIMITS.items():
        buckets[cat].sort(key=lambda it: it["_rank"])
        buckets[cat] = buckets[cat][:limit]
    for lst in buckets.values():
        for it in lst:
            it.pop("_rank", None)
            it.pop("summary", None)  # 简介只用于分类，不必回给前端

    to_translate = [it for lst in buckets.values() for it in lst]

    def _apply(it: dict[str, Any]) -> None:
        it["title_zh"] = _translate_to_zh(it["title"])

    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as ex:
        list(ex.map(_apply, to_translate))

    result = {
        "generated_at": dt.datetime.now(dt.timezone.utc)
        .isoformat(timespec="seconds")
        .replace("+00:00", "Z"),
        "total": sum(len(v) for v in buckets.values()),
        "sources": sorted({it["source"] for lst in buckets.values() for it in lst}),
        "buckets": buckets,
        "cached": False,
        "age_seconds": 0,
    }
    _russia_cache["ts"] = now
    _russia_cache["data"] = result
    return jsonify(result)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=False)
