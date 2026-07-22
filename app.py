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
    ("ТАСС", "https://tass.ru/rss/v2.xml"),
    ("РИА Новости", "https://ria.ru/export/rss2/archive/index.xml"),
    ("РТ на русском", "https://russian.rt.com/rss"),
    ("Lenta.ru", "https://lenta.ru/rss/news"),
    ("Российская газета", "https://rg.ru/xml/index.xml"),
    ("Коммерсантъ", "https://www.kommersant.ru/RSS/news.xml"),
    ("Известия", "https://iz.ru/xml/rss/all.xml"),
    ("Kremlin.ru", "http://kremlin.ru/events/all/feed"),
]

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


def categorize_ru(title: str) -> str:
    """人物类别优先，其次按关键词归类。"""
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
    name, url = name_url
    try:
        resp = requests.get(url, headers=MOBILE_REQUEST_HEADERS, timeout=8)
        resp.raise_for_status()
        parsed = feedparser.parse(resp.content)
    except Exception:
        return []

    out: list[dict[str, Any]] = []
    for entry in parsed.entries[:25]:
        title = (entry.get("title") or "").strip()
        link = entry.get("link") or ""
        published = entry.get("published") or entry.get("updated") or ""
        if title and link:
            out.append(
                {
                    "title": title,
                    "url": link,
                    "source": name,
                    "published": published,
                }
            )
    return out


def _translate_ru_to_zh(text: str) -> str:
    """谷歌无鉴权翻译端点。失败返回空串；前端会退化为只显示俄语原标题。"""
    if not text:
        return ""
    try:
        resp = requests.get(
            "https://translate.googleapis.com/translate_a/single",
            params={
                "client": "gtx",
                "sl": "ru",
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

    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as ex:
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
    }
    for it in unique:
        cat = categorize_ru(it["title"])
        if cat in buckets:
            buckets[cat].append(it)

    LIMITS = {"politics": 8, "economy": 6, "social": 6, "putin": 6, "medinsky": 4}
    for cat, limit in LIMITS.items():
        buckets[cat] = buckets[cat][:limit]

    to_translate = [it for lst in buckets.values() for it in lst]

    def _apply(it: dict[str, Any]) -> None:
        it["title_zh"] = _translate_ru_to_zh(it["title"])

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
