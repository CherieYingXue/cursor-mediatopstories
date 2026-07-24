"""Russia Media Reports — Flask app.

只做一件事：抓俄语原文站的 RSS，按 5 大板块（政治/经济/社会/普京/梅金斯基）归类，
把标题翻译成中文，并在 /russia 页面上呈现，支持打开自动更新 + 手动更新按钮。

原始 media top stories 项目的域名清单、日程调度器、Excel 导入、Zoology 卡片、
IBO 小测验等一切无关逻辑都不在本项目中；见 README。
"""

from __future__ import annotations

import calendar
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
    # 俄乌前线 / 乌方视角（用于填充"俄乌关系"板块）
    (
        "Dzen · СВО",
        "https://news.google.com/rss/search?"
        "q=site:dzen.ru+(%D1%81%D0%B2%D0%BE+OR+%D1%81%D0%BF%D0%B5%D1%86%D0%BE%D0%BF%D0%B5%D1%80%D0%B0%D1%86%D0%B8%D1%8F+OR+%D1%83%D0%BA%D1%80%D0%B0%D0%B8%D0%BD)"
        "&hl=ru&gl=RU&ceid=RU:ru",
    ),
    ("Укринформ", "https://www.ukrinform.net/rss/block-lastnews"),
    ("Интерфакс-Украина", "https://ru.interfax.com.ua/news/last.rss"),
]

# 重点专家：标题或作者中出现任何一个即进入"重点专家"板块（最高优先级）
EXPERT_KEYWORDS: tuple[str, ...] = (
    "лузянин", "luzyanin",
    "маслов", "maslov",
    "кортунов", "kortunov",
    "кашин", "kashin",
)

# =====================================================================
# 是否与俄罗斯直接相关：白名单 → 黑名单 → 默认丢弃
# =====================================================================
# 白名单：出现任何一个关键词（小写、substring 匹配）即视为俄罗斯相关。
# 关键词根覆盖变格（Россия/России/Российский/Российская 都能命中"росси"）。
RUSSIA_RELATED_KEYWORDS: tuple[str, ...] = (
    # -------- 国名 / 政府 / 首都 --------
    "росси", "рф ", " рф", "рф.", "рф,", "рф:",
    "кремл", "москв", "санкт-петербург", "спб",
    # -------- 领导人、高官、专家、梅金斯基 --------
    "путин", "медведев", "лавров", "песков", "захаров", "шойгу",
    "мишустин", "володин", "матвиенко", "кириенко", "дюмин", "белоусов",
    "мантуров", "силуанов", "набиуллин", "попова", "мурашко", "куренков",
    "патрушев", "козак", "аксёнов", "кадыров", "минниханов", "собянин",
    "беглов", "воробьев", "слуцкий", "миронов", "прилепин", "хинштейн",
    "медински", "лузянин", "маслов", "кортунов", "кашин",
    # -------- 联邦机构（含各种缩写） --------
    "минобороны", "мид рф", "мидроссии", "мид россии", "госдум", "совфед",
    "фсб", "свр", "гру", "цб рф", "центробанк", "минфин", "роскомнадзор",
    "минцифр", "минпромторг", "минпросвещени", "минобрнауки", "минздрав",
    "минюст", "минтранс", "минэнерго", "минтруд", "минстрой", "минсельхоз",
    "минкультуры", "минцифры", "минспорт",
    "роскосмос", "росатом", "росстат", "росгвард", "росприрод",
    "роспотребнадзор", "рособрнадзор", "росмолодежь", "роскомсвязь",
    "вс рф", "вкс рф", "вмф рф", "мчс россии", "мчс рф", "рвсн", "гувд",
    "мвд рф", "мвд россии", "фнс", "фас", "цик рф", "цик россии",
    "генштаб", "росавиац", "росаккредитация", "рособоронэкспорт",
    "фомс", "пфр", "сфр", "сбр",
    # -------- 特别军事行动 & 俄乌前线 --------
    "сво", "всу", "спецопераци", "днр", "лнр", "донбасс", "донец",
    "луган", "херсон", "запорож", "мариупол", "энергодар", "зая",
    "новоросси", "приднестров",
    # -------- 联邦主体（85 个 + 主要城市） --------
    # Республики
    "адыге", "алтай", "башк", "бурят", "дагестан", "ингуш", "кабардин",
    "калмык", "карачаев", "карел", "коми ", " коми", "марий эл", "мордови",
    "осетин", "татарстан", "тув", "удмурт", "хакас", "чечн", "чуваш",
    "якут", "саха ",
    # Края
    "алтайск", "забайкал", "камчат", "краснодар", "красноярск", "перм",
    "приморь", "ставропол", "хабаровск",
    # Области
    "амурск", "архангельск", "астрахан", "белгород", "брянск", "владимир",
    "волгоград", "вологод", "воронеж", "иванов", "иркутск", "калининград",
    "калуж", "кемеров", "киров", "кострома", "курган", "курск",
    "ленинград", "липецк", "магадан", "мурманск", "нижегород",
    "новгород", "новосибир", "омск", "оренбург", "орловск", "пенз",
    "псков", "ростов", "рязан", "самар", "саратов", "сахалин", "свердлов",
    "смоленск", "тамбов", "твер", "томск", "тул", "тюмен", "ульяновск",
    "челябинск", "ярославск",
    # Автономные округа + Севастополь + Крым
    "чукот", "ханты-мансийск", "ямало", "югр", "ненецк", "евре а",
    "крым", "севастопол",
    # 常见城市和非首府大城市
    "казан", "владивосток", "хабаров", "уф", "самара ", "тольятт",
    "омск ", "пермь", "воронеж", "волгоград", "красноярск", "барнаул",
    "иркутск", "тюмень", "тула ", "рязань", "владивосток", "махачкал",
    "симферопол", "ялт", "феодоси", "джанкой", "керч", "евпатор",
    "мурманск", "архангельск", "калининград", "сочи", "анапа", "геленджик",
    "новороссийск", "невинномысск", "армавир", "ирбит", "электростал",
    "котовск", "энгельс", "стерлитамак", "чебоксары", "ижевск",
    "магнитогорск", "нижний тагил", "северодвинск", "тобольск", "сургут",
    "нижневартовск", "ноябрьск", "надым", "уренгой", "южно-сахалинск",
    "магадан ", "мурманск", "апатиты", "воркут", "нарьян-мар",
    # 首都主要地点/机场
    "пулково", "домодедово", "шереметьево", "внуково", "жуковский",
    "кольцово", "толмачёво", "толмачево", "курумоч", "гумрак", "стригино",
    "храброво", "витязево", "пашковский",
    "мгу", "спбгу", "мгимо", "ргу", "рггу", "ргб", "ргэу", "миид",
    "эрмитаж", "третьяков", "большой театр", "мариинск",
    # 民航 / 铁路 / 交通
    "аэрофлот", "s7", "победа", "уральские авиалин", "россия аэ",
    "ржд", "ласточка", "сапсан", "москва-казань",
    # 大企业 / 品牌
    "газпром", "роснефть", "лукойл", "новатэк", "татнефт", "башнефт",
    "сбер", "втб", "россельхозбанк", "альфа-банк", "тинькофф", "т-банк",
    "ростех", "ростелеком", "т-платформы", "яндекс", "вконтакте",
    "вкусвилл", "магнит ", "х5 ", "х5 групп", "лента ", "ozon",
    "wildberries", "wb ", " wb", "мтс ", "мегафон", "билайн", "теле2",
    "норильск", "норникел", "русал", "полюс", "евраз", "нлмк", "ммк",
    "аэрофлот", "иркут", "объединенн авиастроительн", "оак", "уралвагон",
    # 军事装备
    "бпла", "пво", "искандер", "с-400", "с-500", "калибр", "герань",
    "ланцет", "орешник", "т-90", "т-14", "армата", "су-35", "су-57",
    "миг-31", "миг-35", "калашников", "ак-47", "ак-74",
    # 货币 / 经济术语
    "рубл", "цб россии", "ставка цб", "ключевая ставка", "мосбирж",
    "минэкономразвит", "инфляция в россии", "автокредит", "ипотека",
    # 政党 / 政治机构
    "единая россия", "кпрф", "лдпр", "справедливая россия", "новые люди",
    "яблоко", "русская партия", "родина", "госуслуги", "мфц",
    # 白俄（俄白联盟国相关）
    "белорус", "лукашенк", "белтелеком", "оршанск",
    # 联合国 / 独联体 / 集安条约（俄罗斯主导的一系列组织）
    "снг ", "одкб", "евразэс", "еаэс",
    # ————————— 英文（Carnegie 英文标题等） —————————
    "russia", "russian", "kremlin", "moscow", "putin", "medvedev",
    "lavrov", "peskov", "zakharova", "medinsky", "shoigu", "mishustin",
    "sobyanin", "kadyrov", "patrushev", "nabiullina",
    "luzyanin", "maslov", "kortunov", "kashin",
    "donbass", "crimea", "sevastopol", "urals", "siberia", "kaliningrad",
    "sberbank", "gazprom", "rosneft", "lukoil", "rosatom", "novatek",
    "wildberries", "yandex", "aeroflot", "vtb",
    "ruble", "belarus", "lukashenko",
    "ukraine war", "russian army", "russian forces", "russian navy",
    "russo-", "sino-russian", "russian-chinese",
)

# 黑名单：明确的外国话题触发词。命中 → 一律剔除（除非同时命中白名单，白名单已在前一步命中优先）
FOREIGN_ONLY_KEYWORDS: tuple[str, ...] = (
    # 伊朗-以色列-中东（不涉俄）
    "тегеран", "иранск", "аятолл", "ксир", "хаменеи", "раиси",
    "нетаньяху", "тель-авив", "иерусалим", "цахал",
    "газ ", "газы", "хамас", "хизбалл", "ливан", "бейрут", "сирийск",
    "асад ", "саудовск", "эр-риад", "оаэ ", "катар ", "йемен", "хут",
    "ормуз",
    # 美国内政（无俄罗斯关联）
    "трамп подпис", "трамп одобрил", "конгресс сша", "белый дом",
    "верховный суд сша", "фбр", "цру",
    # 欧洲内政（非俄涉）
    "макрон", "меркель", "шольц", "олаф шольц",
    "стармер", "сунак", "риши сунак", "джонсон",
    "эрдоган", "италии", "мелони", "испани", "португалии",
    "олландск", "нидерланд", "бельгийск", "греции",
    # 拉美/亚洲外国内政
    "лула ", "болсонар", "мадуро", "боривар",
    "нарендра моди", "си цзиньпин", "ли цян",
    "трюдо", "джастин трюдо",
    "виктор орб", "фицо",
    # 好莱坞 / 娱乐八卦
    "пит", "анджелина", "джоли", "кардашь", "джей-ло", "тейлор свифт",
    "голливуд", "оскар ", "эмми ", "грэмми", "каннск",
    "нба", "нхл", "нфл", "уефа", "фифа",
)

# 5 家源是俄罗斯焦点媒体/智库，其内容一律视为与俄罗斯相关，跳过关键词过滤。
RUSSIA_FOCUSED_SOURCES: frozenset[str] = frozenset({
    "Kremlin.ru",
    "Валдайский клуб",
    "Международная жизнь",
    "РСМД",
    "Carnegie Russia-Eurasia",
})


def is_russia_related(source: str, title: str, summary: str = "") -> bool:
    """判断一条 RSS 条目是否与俄罗斯直接关联。

    - 5 家俄罗斯焦点媒体/智库 → 直接放行
    - 命中白名单 → 保留
    - 命中黑名单（外国话题词汇）→ 丢弃
    - 其余 → 丢弃（保守策略，宁少不多）
    """
    if source in RUSSIA_FOCUSED_SOURCES:
        return True
    haystack = (title + " " + (summary or "")[:600]).lower()
    if any(kw in haystack for kw in RUSSIA_RELATED_KEYWORDS):
        return True
    return False


def is_foreign_only(title: str, summary: str = "") -> bool:
    """辅助函数：标题/简介明显只讲外国事务（不涉俄）时返回 True。"""
    haystack = (title + " " + (summary or "")[:600]).lower()
    return any(kw in haystack for kw in FOREIGN_ONLY_KEYWORDS)


# 源优先级：智库最高，官方媒体次之，通讯社最后。用于每个板块内部排序，
# 让分析类文章在通讯社快讯挤满 8 个位置之前先被选中。
SOURCE_PRIORITY: dict[str, int] = {
    # 智库：最高优先级
    "Валдайский клуб": 0,
    "Международная жизнь": 0,
    "РСМД": 0,
    "Carnegie Russia-Eurasia": 0,
    # 官方媒体
    "Kremlin.ru": 1,
    "Российская газета": 1,
    "Известия": 1,
    "Взгляд": 1,
    # 商业主流报刊
    "Коммерсантъ": 2,
    # 乌方视角（俄乌关系板块的重要补充）
    "Укринформ": 2,
    "Интерфакс-Украина": 2,
    # 通讯社快讯 & 聚合
    "ТАСС": 3,
    "РИА Новости": 3,
    "РТ на русском": 3,
    "Lenta.ru": 3,
    "Dzen · СВО": 3,
}

RUSSIA_CACHE_TTL = int(os.getenv("RUSSIA_CACHE_TTL", "600"))  # 秒；默认 10 分钟
RUSSIA_MAX_AGE_HOURS = float(os.getenv("RUSSIA_MAX_AGE_HOURS", "24"))
_russia_cache: dict[str, Any] = {"ts": 0.0, "data": None}


def _parsed_to_ts(struct_time_or_none) -> float | None:
    """把 feedparser 的 published_parsed（UTC struct_time）转成 Unix 时间戳。"""
    if not struct_time_or_none:
        return None
    try:
        return float(calendar.timegm(struct_time_or_none))
    except Exception:
        return None

_POLITICS_KW = (
    "мид ", "госдум", "цик", "санкц", "нато", "переговор",
    "дипломат", "захаров", "песков", "лавров", "кремл", "парламент",
    "выбор", "оборон", "мигрант", "гражданств", "экстремист", "фсб",
    "теракт", "министр", "закон", "депутат",
    "трамп", "байден", "белорусс", "лукашенк", "балт",
)

# 俄乌关系板块（"所有与乌克兰相关内容"）
_UKRAINE_KW = (
    # 俄方视角常用词
    "украин", "зеленск", "всу", "сво ", " сво", "спецопераци",
    "днр", "лнр", "донбасс", "донец", "луган", "херсон", "запорож",
    "мариупол", "энергодар", "азовск", "азовст",
    # 乌方地区
    "харьк", "одесс", "николаев", "днепропетровск", "днепр", "черниг",
    "житомир", "львов", "полтав", "суми", "черкасс", "винниц",
    "хмельниц", "ровенск", "ужгород", "тернопол", "чернов", "кременчуг",
    "мелитопол", "бердянск", "кривий рог", "кривой рог",
    "киев", "киеве", "киева", "киево", "kyiv", "kiev",
    # 乌方军政人物
    "ермак", "буданов", "залужн", "сырский", "клич",
    "стефанчук", "шмигаль", "уманский", "умеров", "камыш", "камышин",
    # 西方援乌武器
    "himars", "atacms", "taurus", "storm shadow", "patriot", "abrams",
    "leopard", "флеминго", "нептун",
    # 英文
    "ukraine", "ukrainian", "kharkiv", "kherson", "donetsk", "luhansk",
    "mariupol", "azov", "zelensky", "zelenskyy", "zaluzhny", "syrsky",
    "himars", "atacms", "russo-ukrainian", "kyiv independent",
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


# 阿列克谢·久明专属关键词（含变格 + 拉丁转写）
# Алексей Геннадьевич Дюмин — Государственный совет РФ 秘书兼总统助理，
# 曾任图拉州州长，普京核心圈成员。
DYUMIN_KEYWORDS: tuple[str, ...] = (
    "дюмин", "дюмина", "дюмину", "дюминым", "дюмине", "дюминовск",
    "алексей дюмин", "алексея дюмина",
    "dyumin", "diumin", "alexei dyumin", "aleksey dyumin",
)


def categorize_ru(title: str, author: str = "", summary: str = "") -> str:
    """归类：重点专家 > 梅金斯基 > 久明 > 普京 > 俄乌关系 > 政治 > 经济 > 社会 > other。

    - 「重点专家」匹配范围包括标题、作者、正文简介前 500 字，因为像 РСМД、
      Международная жизнь 的 RSS 通常不带 author 字段，专家的姓氏往往只
      出现在简介或副标题里。
    - 梅金斯基、久明这两位克宫核心圈人物排在普京前面，让他们的专属报道
      优先入自己的板块；普京自己的表态/会见落到「普京新闻」；其它一切
      与乌克兰相关的报道（ВСУ、泽连斯基、СВО 战报、乌军城市等）落到
      「俄乌关系」。
    """
    expert_haystack = " ".join([title, author or "", (summary or "")[:500]]).lower()
    for kw in EXPERT_KEYWORDS:
        if kw in expert_haystack:
            return "experts"
    t = title.lower()
    if "медински" in t:
        return "medinsky"
    for kw in DYUMIN_KEYWORDS:
        if kw in t:
            return "dyumin"
    if "путин" in t or "владимир владимирович" in t:
        return "putin"
    for kw in _UKRAINE_KW:
        if kw in t:
            return "ukraine"
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
        published_ts = _parsed_to_ts(
            entry.get("published_parsed") or entry.get("updated_parsed")
        )

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
                    "published_ts": published_ts,
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
        "dyumin": [],
        "experts": [],
        "ukraine": [],
    }
    # (1) 时间过滤：只保留过去 RUSSIA_MAX_AGE_HOURS 小时（默认 24 小时）内发布的条目
    #     没有 published_ts 的条目保守保留（避免把没日期字段的智库分析全砍掉）。
    cutoff_ts = now - RUSSIA_MAX_AGE_HOURS * 3600
    dropped_stale = 0
    dropped_no_date = 0
    fresh: list[dict[str, Any]] = []
    for it in unique:
        ts = it.get("published_ts")
        if ts is None:
            dropped_no_date += 1
            fresh.append(it)  # 保守：没日期字段的智库分析类文章仍保留
        elif ts >= cutoff_ts:
            fresh.append(it)
        else:
            dropped_stale += 1

    # (2) 内容过滤：只保留与俄罗斯直接相关的条目
    dropped_non_russia = 0
    russia_related: list[dict[str, Any]] = []
    for it in fresh:
        if is_russia_related(it["source"], it["title"], it.get("summary", "")):
            russia_related.append(it)
        else:
            dropped_non_russia += 1

    # 分类并记录每条新闻在**同一 source 内**的序号，用于轮询排序。
    per_source_seen: dict[str, int] = {}
    for idx, it in enumerate(russia_related):
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
        "politics": 8, "economy": 6, "social": 5,
        "putin": 6, "medinsky": 4, "dyumin": 4,
        "experts": 6, "ukraine": 8,
    }
    for cat, limit in LIMITS.items():
        buckets[cat].sort(key=lambda it: it["_rank"])
        buckets[cat] = buckets[cat][:limit]
    for lst in buckets.values():
        for it in lst:
            it.pop("_rank", None)
            it.pop("summary", None)  # 简介只用于分类，不必回给前端
            it.pop("published_ts", None)  # 只用于时间过滤，不必回给前端

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
        "stats": {
            "fetched_unique": len(unique),
            "within_last_hours": len(fresh),
            "russia_related": len(russia_related),
            "dropped_stale": dropped_stale,
            "kept_no_date": dropped_no_date,
            "dropped_non_russia": dropped_non_russia,
            "max_age_hours": RUSSIA_MAX_AGE_HOURS,
        },
    }
    _russia_cache["ts"] = now
    _russia_cache["data"] = result
    return jsonify(result)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=False)
