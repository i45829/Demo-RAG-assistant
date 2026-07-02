"""
app.py — Демо-система «Литературный обзор по теме» для презентации руководству.

ЧТО ДЕЛАЕТ:
  1. Пользователь вводит тему в браузере (никакой установки — просто веб-страница).
  2. Приложение ищет источники через бесплатные научные API (PubMed, arXiv,
     Europe PMC, DOAJ, Semantic Scholar, Crossref + Unpaywall для платных статей),
     а также опционально через открытые патентные базы (USPTO PatentsView, EPO OPS).
     OpenAlex исключён из активных источников из-за постоянных ошибок доступа с IP
     облачных хостингов (частые 403/429) — вместо него используется Semantic Scholar.
  3. Собранные данные передаются в LLM (Claude или Gemini — на выбор), которая
     синтезирует обзор литературы с нумерованными ссылками [n].
  4. Результат показывается в браузере и скачивается как .docx.

ЗАПУСК ЛОКАЛЬНО (для проверки перед деплоем, не обязателен для демо):
  pip install streamlit httpx python-docx
  streamlit run app.py

ДЕПЛОЙ (бесплатно, без установки на ПК начальства) — см. файл
deploy_streamlit_cloud.md в комплекте.
"""

import io
import os
import time
import base64
import json
from datetime import date

import httpx
import streamlit as st
from docx import Document
from docx.shared import Pt

# ============================== НАСТРОЙКИ СТРАНИЦЫ ==============================

st.set_page_config(page_title="Литературный обзор — демо", page_icon="📚", layout="wide")

# Значение по умолчанию; перезаписывается реальным email из сайдбара при запуске —
# так запросы к API попадают в "вежливый пул" (polite pool) и реже получают отказы.
CONTACT_EMAIL = "demo@example.com"


class _RetryableHTTPError(Exception):
    """Понятная ошибка с кодом статуса — чтобы предупреждения в интерфейсе были
    диагностируемыми, а не просто 'не удалось получить данные'."""
    def __init__(self, status_code, url):
        self.status_code = status_code
        super().__init__(f"HTTP {status_code} от {url}")


def _request_with_retry(method, url, retries=3, backoff=1.5, **kwargs):
    """Обёртка над httpx с повторными попытками для временных сбоев:
    429 (превышен лимит запросов — САМАЯ частая причина отказов у бесплатных API
    при вызовах с общего IP облачного хостинга, как у Streamlit Cloud), 502/503/504,
    таймауты. Для 429 уважает заголовок Retry-After, если он есть.
    follow_redirects=True по умолчанию — некоторые API (например arXiv) отдают 301
    с http на https, а httpx по умолчанию редиректы НЕ проходит и падает с ошибкой."""
    kwargs.setdefault("follow_redirects", True)
    headers = kwargs.get("headers") or {}
    headers.setdefault("User-Agent", f"lit-review-demo/1.0 (mailto:{CONTACT_EMAIL})")
    kwargs["headers"] = headers
    last_exc = None
    for attempt in range(retries + 1):
        try:
            r = getattr(httpx, method)(url, **kwargs)
            if r.status_code == 429 and attempt < retries:
                wait = float(r.headers.get("Retry-After", backoff * (attempt + 1) * 2))
                time.sleep(min(wait, 15))
                continue
            if r.status_code in (502, 503, 504) and attempt < retries:
                time.sleep(backoff * (attempt + 1))
                continue
            if r.status_code >= 400:
                last_exc = _RetryableHTTPError(r.status_code, url)
                raise last_exc
            return r
        except (httpx.TimeoutException, httpx.TransportError) as exc:
            last_exc = exc
            if attempt < retries:
                time.sleep(backoff * (attempt + 1))
                continue
            raise
        except _RetryableHTTPError:
            raise
    raise last_exc


def _get(url, **kwargs):
    return _request_with_retry("get", url, **kwargs)


def _post(url, **kwargs):
    return _request_with_retry("post", url, **kwargs)


# ============================== ПОИСК ПО ОТКРЫТЫМ API ==============================


def _trim(text, limit=1500):
    text = (text or "").strip().replace("\n", " ")
    return text if len(text) <= limit else text[:limit] + "…"




def search_pubmed(query, max_results=5):
    """PubMed E-utilities (NCBI) — 36M+ публикаций в медицине, биологии, химии, фармакологии.
    Без ключа (ключ увеличивает лимит до 10 запросов/сек, с ним до 100K/день).
    Отличие от Europe PMC: другая поисковая система, часто комплементарные результаты."""
    try:
        # Шаг 1: поиск PMID
        r = _get(
            "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi",
            params={"db": "pubmed", "term": query, "retmax": max_results,
                    "retmode": "json", "sort": "relevance"},
            timeout=20,
        )
        r.raise_for_status()
        id_list = r.json().get("esearchresult", {}).get("idlist", [])
        if not id_list:
            return []
        # Шаг 2: детали по PMID
        r2 = _get(
            "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi",
            params={"db": "pubmed", "id": ",".join(id_list), "retmode": "json"},
            timeout=20,
        )
        r2.raise_for_status()
        result_data = r2.json().get("result", {})
        # Шаг 3: аннотации
        r3 = _get(
            "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi",
            params={"db": "pubmed", "id": ",".join(id_list), "rettype": "abstract", "retmode": "xml"},
            timeout=25,
        )
        r3.raise_for_status()
        # Простое извлечение аннотаций из XML без lxml
        abstracts = {}
        import re as _re
        for m in _re.finditer(r'<PMID[^>]*>(\d+)</PMID>.*?<AbstractText[^>]*>(.*?)</AbstractText>',
                              r3.text, _re.DOTALL):
            pmid, abst = m.group(1), _re.sub(r'<[^>]+>', '', m.group(2))
            if pmid not in abstracts:
                abstracts[pmid] = abst.strip()
        items = []
        for pmid in id_list:
            meta = result_data.get(pmid, {})
            title = meta.get("title", "(без названия)")
            year = meta.get("pubdate", "н/д")[:4]
            authors = "; ".join(
                f"{a.get('name','')}" for a in (meta.get("authors") or [])[:4]
            ) or "н/д"
            items.append({
                "title": title,
                "authors": authors,
                "year": year,
                "source": "PubMed (NCBI)",
                "url": f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/",
                "doi": "",
                "abstract": _trim(abstracts.get(pmid, "")),
            })
        return items
    except Exception as exc:
        st.warning(f"PubMed: не удалось получить данные ({exc})")
        return []


def search_arxiv(query, max_results=5):
    try:
        import feedparser
        r = _get(
            "https://export.arxiv.org/api/query",
            params={"search_query": f"all:{query}", "start": 0,
                    "max_results": max_results, "sortBy": "relevance"},
            timeout=20,
        )
        r.raise_for_status()
        feed = feedparser.parse(r.text)
        items = []
        for e in feed.entries:
            items.append({
                "title": _trim(getattr(e, "title", "(без названия)"), 300),
                "authors": ", ".join(a.name for a in getattr(e, "authors", [])[:5]) or "н/д",
                "year": (getattr(e, "published", "н/д") or "н/д")[:4],
                "source": "arXiv",
                "url": getattr(e, "link", ""),
                "doi": "",
                "abstract": _trim(getattr(e, "summary", "")),
            })
        return items
    except Exception as exc:
        st.warning(f"arXiv: не удалось получить данные ({exc})")
        return []


def search_europepmc(query, max_results=5):
    try:
        r = _get(
            "https://www.ebi.ac.uk/europepmc/webservices/rest/search",
            params={"query": query, "format": "json", "resultType": "core",
                    "pageSize": max_results, "email": CONTACT_EMAIL},
            timeout=20,
        )
        r.raise_for_status()
        items = []
        for res in (r.json().get("resultList") or {}).get("result", []):
            doi = res.get("doi") or ""
            items.append({
                "title": res.get("title") or "(без названия)",
                "authors": res.get("authorString") or "н/д",
                "year": res.get("pubYear") or "н/д",
                "source": "Europe PMC",
                "url": f"https://doi.org/{doi}" if doi else "",
                "doi": doi,
                "abstract": _trim(res.get("abstractText") or ""),
            })
        return items
    except Exception as exc:
        st.warning(f"Europe PMC: не удалось получить данные ({exc})")
        return []


def search_doaj(query, max_results=5):
    try:
        from urllib.parse import quote
        r = _get(
            f"https://doaj.org/api/search/articles/{quote(query)}",
            params={"pageSize": max_results}, timeout=20,
        )
        r.raise_for_status()
        items = []
        for res in r.json().get("results", []):
            bib = res.get("bibjson", {})
            doi = next((i.get("id", "") for i in bib.get("identifier", [])
                        if (i.get("type") or "").lower() == "doi"), "")
            link = next((l.get("url", "") for l in bib.get("link", [])
                         if l.get("type") == "fulltext"), "") or (
                f"https://doi.org/{doi}" if doi else "")
            items.append({
                "title": bib.get("title") or "(без названия)",
                "authors": ", ".join(a.get("name", "") for a in (bib.get("author") or [])[:5]) or "н/д",
                "year": bib.get("year") or "н/д",
                "source": "DOAJ",
                "url": link,
                "doi": doi,
                "abstract": _trim(bib.get("abstract") or ""),
            })
        return items
    except Exception as exc:
        st.warning(f"DOAJ: не удалось получить данные ({exc})")
        return []


def search_semanticscholar(query, max_results=5, api_key=""):
    """Semantic Scholar — 200M+ работ. Используется вместо OpenAlex: у OpenAlex стабильно
    возникали ошибки доступа (403/429) при запросах с IP облачных хостингов.

    ВАЖНО про 429: без ключа запрос идёт в ОБЩИЙ анонимный пул лимитов Semantic
    Scholar, который делят между собой ВСЕ бесключевые обращения к API в мире —
    включая другие приложения, размещённые на том же облачном хостинге (например,
    Streamlit Community Cloud раздаёт IP из общего набора на тысячи чужих приложений).
    429 может прилетать из-за чужого трафика, а не из-за количества ваших запросов.
    Бесплатный личный ключ (https://www.semanticscholar.org/product/api#api-key)
    даёт отдельный, куда более щедрый лимит — рекомендуется, если 429 повторяются."""
    headers = {"x-api-key": api_key.strip()} if api_key.strip() else {}
    try:
        r = _get(
            "https://api.semanticscholar.org/graph/v1/paper/search",
            params={
                "query": query, "limit": max_results,
                "fields": "title,year,authors,abstract,externalIds,venue,url",
            },
            headers=headers,
            timeout=20,
        )
        r.raise_for_status()
        items = []
        for p in (r.json().get("data") or []):
            doi = (p.get("externalIds") or {}).get("DOI", "")
            items.append({
                "title": p.get("title") or "(без названия)",
                "authors": ", ".join(a.get("name", "") for a in (p.get("authors") or [])[:5]) or "н/д",
                "year": p.get("year") or "н/д",
                "source": "Semantic Scholar",
                "url": f"https://doi.org/{doi}" if doi else (p.get("url") or ""),
                "doi": doi,
                "abstract": _trim(p.get("abstract") or ""),
            })
        return items
    except Exception as exc:
        st.warning(f"Semantic Scholar: не удалось получить данные ({exc})")
        return []


def search_crossref(query, max_results=5):
    """Метаданные почти любого DOI, включая закрытые журналы (ScienceDirect и т.п.)."""
    try:
        r = _get(
            "https://api.crossref.org/works",
            params={"query": query, "rows": max_results, "mailto": CONTACT_EMAIL},
            timeout=20,
        )
        r.raise_for_status()
        items = []
        for it in (r.json().get("message") or {}).get("items", []):
            doi = it.get("DOI", "")
            year = None
            for k in ("published-print", "published-online", "published"):
                dp = (it.get(k) or {}).get("date-parts")
                if dp:
                    year = dp[0][0]
                    break
            items.append({
                "title": (it.get("title") or ["(без названия)"])[0],
                "authors": ", ".join(
                    f"{a.get('given','')} {a.get('family','')}".strip()
                    for a in (it.get("author") or [])[:5]
                ) or "н/д",
                "year": year or "н/д",
                "source": f"Crossref ({it.get('publisher', 'н/д')})",
                "url": f"https://doi.org/{doi}" if doi else "",
                "doi": doi,
                "abstract": "",  # Crossref обычно не даёт аннотацию
            })
        return items
    except Exception as exc:
        st.warning(f"Crossref: не удалось получить данные ({exc})")
        return []


def check_unpaywall(doi):
    """По DOI проверяет легальную бесплатную копию платной статьи."""
    if not doi:
        return None
    try:
        r = _get(f"https://api.unpaywall.org/v2/{doi}",
                       params={"email": CONTACT_EMAIL}, timeout=15)
        r.raise_for_status()
        data = r.json()
        if not data.get("is_oa"):
            return None
        best = data.get("best_oa_location") or {}
        return best.get("url_for_pdf") or best.get("url")
    except Exception:
        return None


# ============================== ПАТЕНТНЫЕ БАЗЫ (нужны свои ключи) ==============================
# Каждая функция возвращает (items, error) — error=None при успехе, иначе строка с
# понятным описанием проблемы (модель показывает её пользователю, а не падает).


# ============================== ПАТЕНТНЫЕ БАЗЫ ==============================
# Проверено на июль 2026:
#  - legacy api.patentsview.org ОТКЛЮЧЁН с 1 мая 2025 (301 на страницу миграции) —
#    старая функция всегда падала с ошибкой, поэтому обзор "терял" патентные источники.
#  - У WIPO PATENTSCOPE НЕТ публичного REST/JSON API вообще (подтверждено официальным
#    каталогом API WIPO, apicatalog.wipo.int) — прежняя функция обращалась к
#    несуществующему эндпоинту и не могла работать в принципе. Убрана.
#  - Рабочие открытые варианты: новый PatentsView PatentSearch API (нужен бесплатный
#    ключ, search.patentsview.org) и EPO OPS/Espacenet (нужна бесплатная регистрация).


def search_patentsview(query, max_results=5, api_key=""):
    """USPTO PatentsView PatentSearch API (новая версия, 2025+).
    Старый api.patentsview.org отключён — актуальный эндпоинт search.patentsview.org
    требует бесплатный ключ (X-Api-Key). Получить: https://patentsview.org/apis/purpose
    (ссылка "Request an API Key"). Покрывает патенты США с 1976 года + аннотации.
    Примечание: по состоянию на середину 2026 сервис сообщался как нестабильный
    (эпизодические 500-е ошибки) — в этом случае используйте EPO OPS как основной
    источник патентов."""
    if not api_key.strip():
        return [], ("USPTO PatentsView: нужен бесплатный API-ключ (старый доступ без "
                     "ключа отключён с мая 2025). Получить: patentsview.org/apis/purpose")
    try:
        r = _get(
            "https://search.patentsview.org/api/v1/patent/",
            params={
                "q": json.dumps({"_text_any": {"patent_abstract": query}}),
                "f": json.dumps(["patent_id", "patent_title", "patent_date",
                                 "patent_abstract"]),
                "o": json.dumps({"size": max_results}),
            },
            headers={"X-Api-Key": api_key.strip()},
            timeout=25,
        )
        r.raise_for_status()
        items = []
        for p in (r.json().get("patents") or [])[:max_results]:
            pn = p.get("patent_id") or ""
            items.append({
                "title": p.get("patent_title") or f"US Patent {pn}",
                "authors": "н/д",
                "year": (p.get("patent_date") or "н/д")[:4],
                "source": "USPTO PatentsView",
                "url": f"https://worldwide.espacenet.com/patent/search?q=pn%3DUS{pn}",
                "doi": "",
                "abstract": _trim(p.get("patent_abstract") or ""),
            })
        return items, None
    except Exception as exc:
        return [], f"USPTO PatentsView: ошибка запроса ({exc})"


def search_core_oa(query, max_results=5):
    """CORE — крупнейший агрегатор открытых полных текстов (300M+ документов).
    Включает часть патентов EPO и отчёты. Бесплатный ключ CORE_API_KEY (core.ac.uk)."""
    api_key = os.environ.get("CORE_API_KEY", "")
    if not api_key:
        st.info("CORE: для поиска задайте CORE_API_KEY (бесплатно на core.ac.uk/services/api)")
        return []
    try:
        r = _get(
            "https://api.core.ac.uk/v3/search/outputs",
            params={"q": query, "limit": max_results},
            headers={"Authorization": f"Bearer {api_key}"},
            timeout=25,
        )
        r.raise_for_status()
        items = []
        for d in (r.json().get("results") or [])[:max_results]:
            doi = d.get("doi") or ""
            items.append({
                "title": d.get("title") or "(без названия)",
                "authors": ", ".join(a.get("name","") for a in (d.get("authors") or [])[:4]) or "н/д",
                "year": str(d.get("yearPublished") or "н/д"),
                "source": "CORE",
                "url": d.get("downloadUrl") or (f"https://doi.org/{doi}" if doi else ""),
                "doi": doi,
                "abstract": _trim(d.get("abstract") or ""),
            })
        return items
    except Exception as exc:
        st.warning(f"CORE: не удалось получить данные ({exc})")
        return []


def search_epo_ops_simple(query, max_results=5, consumer_key="", consumer_secret=""):
    """EPO OPS (Espacenet) — патенты всех крупных ведомств через европейское патентное ведомство.
    Бесплатная регистрация на developers.epo.org → My Apps → Consumer Key/Secret."""
    if not consumer_key or not consumer_secret:
        return [], "EPO OPS (Espacenet): укажите Consumer Key и Secret (бесплатная регистрация на developers.epo.org)."
    try:
        tok_r = _post(
            "https://ops.epo.org/3.2/auth/accesstoken",
            headers={"Authorization": "Basic " + base64.b64encode(
                f"{consumer_key}:{consumer_secret}".encode()).decode(),
                "Content-Type": "application/x-www-form-urlencoded"},
            data={"grant_type": "client_credentials"}, timeout=20,
        )
        tok_r.raise_for_status()
        token = tok_r.json()["access_token"]
    except Exception as exc:
        return [], f"EPO OPS: ошибка авторизации ({exc})"
    cql = query if "=" in query else f'ti="{query}" or ab="{query}"'
    try:
        r = _get(
            "https://ops.epo.org/3.2/rest-services/published-data/search/biblio",
            params={"q": cql, "Range": f"1-{max_results}"},
            headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
            timeout=25,
        )
        r.raise_for_status()
        data = r.json()
    except Exception as exc:
        return [], f"EPO OPS: ошибка поиска ({exc})"
    items = []
    try:
        biblio = (data.get("ops:world-patent-data") or {}).get("ops:biblio-search") or {}
        docs = (biblio.get("ops:search-result") or {}).get("ops:publication-reference", [])
        if isinstance(docs, dict):
            docs = [docs]
        for d in docs[:max_results]:
            did = d.get("document-id", {})
            country = (did.get("country") or {}).get("$", "")
            num = (did.get("doc-number") or {}).get("$", "")
            kind = (did.get("kind") or {}).get("$", "")
            pn = f"{country}{num}{kind}"
            items.append({
                "title": f"Патент {pn}",
                "authors": "н/д", "year": "н/д",
                "source": "Espacenet (EPO OPS)",
                "url": f"https://worldwide.espacenet.com/patent/search?q=pn%3D{country}{num}",
                "doi": "", "abstract": "",
            })
    except Exception as exc:
        return [], f"EPO OPS: не удалось разобрать ответ ({exc})"
    return items, None


PATENT_SOURCES = {
    # Каждая функция здесь возвращает (items, error) — error=None при успехе.
    # Поля учётных данных: (подпись, ключ хранения в session_state, имя kwarg функции).
    "USPTO PatentsView (нужен бесплатный API-ключ)": (
        search_patentsview, [("API-ключ PatentsView", "patentsview_api_key", "api_key")],
    ),
    "Espacenet / EPO OPS (бесплатная регистрация, Consumer Key+Secret)": (
        search_epo_ops_simple,
        [("Consumer Key (EPO)", "epo_consumer_key", "consumer_key"),
         ("Consumer Secret (EPO)", "epo_consumer_secret", "consumer_secret")],
    ),
}

SOURCES = {
    "PubMed / NCBI (медицина, химия, биология)": search_pubmed,
    "Europe PMC (PMC + Европейская коллекция)": search_europepmc,
    "Semantic Scholar (200M+ работ)": search_semanticscholar,
    "Crossref (метаданные DOI + Unpaywall)": search_crossref,
    "DOAJ (открытые журналы)": search_doaj,
    "arXiv (препринты, тема на англ.)": search_arxiv,
}

# ============================== СИНТЕЗ ОБЗОРА (LLM) ==============================

REVIEW_SYSTEM_PROMPT = """Ты — научный аналитик, готовящий раздел «Обзор литературы» для диссертации по естественнонаучной/технической теме.

ПРАВИЛО СОДЕРЖАНИЯ — ГЛАВНОЕ:
Обзор состоит из фактов предметной области. Из каждой аннотации извлекай:
- точные названия веществ, соединений, материалов, реагентов, препаратов, штаммов, методов, приборов, алгоритмов — как написано в источнике;
- химические формулы и обозначения (CH3COONa, HPMC K15M, EC50 = 0.3 мкМ) — переноси буквально;
- количественные данные: концентрации, дозы, температуры, pH, вязкость, размеры частиц, эффективность, проценты, p-значение;
- механизмы и принципы действия — как описаны в источнике;
- конкретные числовые результаты экспериментов.

СТРОГО ЗАПРЕЩЕНО в теле обзора:
- «было найдено X источников», «источники были проанализированы», «проведён поиск» — это не научная информация;
- обобщения без цифр («многие исследования показывают» без [n] и конкретных данных).

ОБЯЗАТЕЛЬНОЕ ПРАВИЛО — ТЫ ВСЕГДА ПИШЕШЬ ОБЗОР:
- Используй ВСЕ предоставленные источники. Никогда не отказывайся от написания обзора целиком.
- Если источник слабо связан с темой — используй из него то, что относится, и отметь в одном предложении его ограниченную релевантность.
- Если аннотации не содержат технических деталей — напиши об этом в разделе «Выводы» как об ограничении, но раздел «Обзор литературы» всё равно заполни тем, что есть в источниках.
- Не выдумывай факты, которых нет в аннотациях.

Каждое фактическое утверждение сопровождай ссылкой [n] на номер источника.

Структура ответа (строго Markdown, на русском языке):
## Введение
2-4 предложения: тема, актуальность, границы обзора. Без клише вроде «в данной работе рассматривается».

## Обзор литературы
Тематический синтез с конкретными данными: вещества, формулы, параметры, механизмы, результаты. Ссылка [n] на каждый факт. Подзаголовки ### по подтемам, если нужно.

## Выводы и направления дальнейших исследований
- Обоснованные выводы из данных [n]
- Пробелы в изученности (какие параметры/механизмы не охвачены)
- Рекомендации (явно как предположение)

## Список литературы
[n] Авторы. Название. Год. Источник. URL.

ЯЗЫК ОБЗОРА И ИСТОЧНИКОВ:
- Тело обзора (Введение, Обзор литературы, Выводы) пишется ТОЛЬКО на русском языке, независимо от языка источников.
- Аннотации источников могут быть на английском или другом языке — извлекай факты и передавай их по-русски своими словами. Не оставляй фрагменты на английском в тексте обзора.
- Специальные термины, названия веществ/методов/приборов и химические формулы сохраняй как в источнике (например, «гидроксипропилметилцеллюлоза (HPMC K15M)», «CH3COONa», «EC50 = 0.3 мкМ»).
- В «Списке литературы» название статьи оставляй на языке оригинала — не переводи. Авторы, год, источник, URL — как есть."""


def build_user_message(topic, items):
    n_with_abstract = sum(1 for it in items if it.get("abstract", "").strip())
    lines = [
        f"ТЕМА ОБЗОРА: {topic}", "",
        f"ИСТОЧНИКОВ ПЕРЕДАНО: {len(items)}, из них с аннотацией: {n_with_abstract}.",
        "Если аннотация пуста ('н/д') — используй название/авторов/год,",
        "пометив '(аннотация недоступна)'. НЕ оставляй обзор пустым.", "",
    ]
    for i, it in enumerate(items, 1):
        lines.append(
            f"[{i}] {it['title']} ({it['year']}). Авторы: {it['authors']}. "
            f"Источник: {it['source']}. URL: {it['url'] or 'н/д'}\n"
            f"Аннотация: {it['abstract'] or 'н/д'}\n"
        )
    return "\n".join(lines)


def call_gemini(system_prompt, user_message, api_key, model="gemini-2.0-flash", temperature=0.3):
    """Известная особенность моделей линейки Gemini 2.5/3.x: их "thinking"-токены
    расходуются из ОБЩЕГО бюджета maxOutputTokens. Если не задать этот параметр
    явно (лимит по умолчанию невелик) и не ограничить бюджет на размышления,
    модель может потратить весь лимит на внутренний thinking и вернуть
    finishReason=MAX_TOKENS с пустым text. Раньше это приводило к необработанному
    KeyError при разборе ответа. Теперь:
      1) явно задаём maxOutputTokens с запасом;
      2) для thinking-моделей ограничиваем thinkingBudget, оставляя место под
         сам текст обзора (для остальных моделей делаем retry без этого поля на 400);
      3) вместо KeyError кидаем понятную ошибку с диагностикой (finishReason,
         сколько токенов ушло на thinking), которую увидит пользователь."""
    generation_config = {"temperature": temperature, "maxOutputTokens": 8192}
    if any(tag in model for tag in ("2.5", "3.0", "3-", "thinking")):
        generation_config["thinkingConfig"] = {"thinkingBudget": 1024}

    def _do_request(cfg):
        return httpx.post(
            f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
            params={"key": api_key},
            json={
                "system_instruction": {"parts": [{"text": system_prompt}]},
                "contents": [{"role": "user", "parts": [{"text": user_message}]}],
                "generationConfig": cfg,
            },
            timeout=120,
        )

    r = _do_request(generation_config)
    if r.status_code == 400 and "thinkingConfig" in generation_config:
        generation_config = dict(generation_config)
        generation_config.pop("thinkingConfig")
        r = _do_request(generation_config)
    r.raise_for_status()
    data = r.json()

    candidates = data.get("candidates") or []
    if not candidates:
        block_reason = (data.get("promptFeedback") or {}).get("blockReason")
        raise RuntimeError(
            "Gemini не вернул ни одного варианта ответа"
            + (f" (blockReason={block_reason})" if block_reason else "")
        )

    cand = candidates[0]
    parts = (cand.get("content") or {}).get("parts") or []
    text = "".join(p.get("text", "") for p in parts).strip()
    if not text:
        finish_reason = cand.get("finishReason", "н/д")
        thoughts = (data.get("usageMetadata") or {}).get("thoughtsTokenCount", 0)
        raise RuntimeError(
            f"Gemini вернул пустой текст (finishReason={finish_reason}, "
            f"токенов на 'размышление': {thoughts}). Модель, вероятно, "
            f"израсходовала весь лимит токенов на внутренний thinking. "
            f"Попробуйте модель без thinking (например gemini-2.0-flash) "
            f"или повторите запрос."
        )
    return text


def call_polza(system_prompt, user_message, api_key, model, temperature=0.3):
    """Polza.ai — агрегатор моделей с OpenAI-совместимым API.
    Формат model: 'provider/model', например 'openai/gpt-4o', 'anthropic/claude-3.7-sonnet'."""
    r = httpx.post(
        "https://polza.ai/api/v1/chat/completions",
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json={
            "model": model,
            "temperature": temperature,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_message},
            ],
        },
        timeout=120,
    )
    r.raise_for_status()
    data = r.json()
    return data["choices"][0]["message"]["content"]


def call_llm(provider, model_id, system_prompt, user_message, api_key, temperature=0.3):
    if provider == "Google Gemini":
        return call_gemini(system_prompt, user_message, api_key, model=model_id, temperature=temperature)
    return call_polza(system_prompt, user_message, api_key, model=model_id, temperature=temperature)


# ============================== ДВУЯЗЫЧНЫЙ ПОИСК ==============================

TRANSLATE_SYSTEM_PROMPT = """Ты переводишь научные темы поиска с любого языка на английский.
Задача: получить 1–2 англоязычных поисковых запроса, которые будут релевантны научным
базам (PubMed, Crossref, Semantic Scholar). Правила:
- Используй устоявшуюся англоязычную научную терминологию, а не дословный перевод.
- НЕ добавляй кавычки, операторы AND/OR, комментарии, объяснения.
- НЕ пиши преамбулы вроде "Here are the queries".
- Верни строго JSON-массив строк. Пример: ["oil dispersions plant protection", "adjuvant oil emulsion pesticide formulation"]
- Если тема уже на английском — верни массив с 1–2 уточнёнными формулировками той же темы."""


def translate_topic_to_english(topic, provider, model_id, api_key):
    """Переводит тему на английский с помощью той же LLM. Возвращает список из 1–2 англ. вариантов
    (или пустой список, если что-то пошло не так — тогда двуязычный поиск просто не выполнится,
    но обычный русский поиск остаётся). Использует temperature=0 для стабильности перевода."""
    try:
        raw = call_llm(provider, model_id, TRANSLATE_SYSTEM_PROMPT,
                       f"Тема: {topic}", api_key, temperature=0.0)
        raw = raw.strip()
        # Модели иногда оборачивают JSON в ```json ... ``` — снимаем
        if raw.startswith("```"):
            raw = raw.strip("`")
            if raw.lower().startswith("json"):
                raw = raw[4:].lstrip()
        # Ищем первый массив в тексте (на случай, если модель добавила преамбулу)
        start = raw.find("[")
        end = raw.rfind("]")
        if start == -1 or end == -1 or end <= start:
            return []
        arr = json.loads(raw[start:end + 1])
        if not isinstance(arr, list):
            return []
        # Оставляем только непустые строки, до 2 запросов, отличающиеся от исходной темы
        cleaned = []
        for q in arr:
            if isinstance(q, str):
                q = q.strip()
                if q and q.lower() != topic.strip().lower() and q not in cleaned:
                    cleaned.append(q)
        return cleaned[:2]
    except Exception:
        return []


def _dedup_items(items):
    """Убирает дубликаты статей, найденных по разным языковым запросам.
    Элемент считается уже виденным, если совпал хотя бы по одному идентификатору:
    DOI, URL или нормализованное название. Это важно, потому что один и тот же
    источник по русскому и английскому запросу может прийти с чуть разным набором
    полей (у одной копии есть DOI, у другой — только URL, и т.п.)."""
    seen_doi, seen_url, seen_title = set(), set(), set()
    result = []
    for it in items:
        doi = (it.get("doi") or "").strip().lower()
        url = (it.get("url") or "").strip().lower()
        title_norm = "".join((it.get("title") or "").lower().split())
        if (doi and doi in seen_doi) or (url and url in seen_url) \
                or (title_norm and title_norm in seen_title):
            continue
        if not (doi or url or title_norm):
            continue  # совсем пустая запись — пропускаем
        if doi:
            seen_doi.add(doi)
        if url:
            seen_url.add(url)
        if title_norm:
            seen_title.add(title_norm)
        result.append(it)
    return result


# ------------------------- Загрузка списка доступных моделей -------------------------


@st.cache_data(ttl=600, show_spinner=False)
def fetch_gemini_models(api_key):
    """Список моделей Gemini, поддерживающих generateContent."""
    r = httpx.get(
        "https://generativelanguage.googleapis.com/v1beta/models",
        params={"key": api_key}, timeout=20,
    )
    r.raise_for_status()
    out = []
    for m in r.json().get("models", []):
        if "generateContent" in (m.get("supportedGenerationMethods") or []):
            model_id = m["name"].removeprefix("models/")
            out.append((model_id, m.get("displayName") or model_id))
    return out


@st.cache_data(ttl=600, show_spinner=False)
def fetch_polza_models(api_key):
    """Список чат-моделей, доступных через Polza.ai (агрегатор OpenAI/Anthropic/Google и др.)."""
    r = httpx.get(
        "https://polza.ai/api/v1/models",
        params={"type": "chat"},
        headers={"Authorization": f"Bearer {api_key}"}, timeout=20,
    )
    r.raise_for_status()
    out = []
    for m in r.json().get("data", []):
        out.append((m["id"], m.get("name") or m["id"]))
    return out


def fetch_models(provider, api_key):
    if provider == "Google Gemini":
        return fetch_gemini_models(api_key)
    return fetch_polza_models(api_key)


# ============================== ЭКСПОРТ В WORD ==============================


def build_docx(topic, review_markdown):
    doc = Document()
    doc.add_heading("Литературный обзор", level=0)
    doc.add_paragraph(f"Тема: {topic}")
    doc.add_paragraph(f"Дата формирования: {date.today().strftime('%d.%m.%Y')}")
    doc.add_paragraph("Сгенерировано автоматически — демонстрационная версия.").italic = True

    for raw_line in review_markdown.split("\n"):
        line = raw_line.rstrip()
        if line.startswith("## "):
            doc.add_heading(line[3:], level=1)
        elif line.startswith("# "):
            doc.add_heading(line[2:], level=0)
        elif line.startswith("- "):
            doc.add_paragraph(line[2:], style="List Bullet")
        elif line.strip():
            p = doc.add_paragraph(line)
            p.style.font.size = Pt(11)
        else:
            doc.add_paragraph("")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ==================================== UI ====================================

st.title("📚 Демо: автоматический литературный обзор")
st.caption(
    "Вводите тему → система ищет источники в открытых научных базах → "
    "ИИ формирует обзор со ссылками → скачиваете Word-документ."
)

with st.sidebar:
    st.header("Настройки")

    provider = st.selectbox(
        "Провайдер LLM",
        ["Google Gemini", "Polza.ai (агрегатор моделей)"],
        help="Google Gemini — бесплатный тариф от Google. "
             "Polza.ai — платный (в рублях) агрегатор, даёт доступ к сотням моделей "
             "(OpenAI, Anthropic, Google и др.) через один ключ.",
    )
    key_label = "Google AI Studio" if provider == "Google Gemini" else "Polza.ai"
    api_key = st.text_input(
        f"API-ключ {key_label}",
        type="password",
        help="Ключ не сохраняется — используется только для этого запроса в этой сессии.",
    )

    model_id = None
    if api_key:
        try:
            with st.spinner("Загружаю список доступных моделей…"):
                model_options = fetch_models(provider, api_key)
        except Exception as exc:
            model_options = []
            st.error(f"Не удалось загрузить список моделей: {exc}")

        if model_options:
            labels = [f"{name}  ·  {mid}" if name != mid else mid for mid, name in model_options]
            idx = st.selectbox(
                "Модель", range(len(model_options)), format_func=lambda i: labels[i]
            )
            model_id = model_options[idx][0]
            st.caption(f"Доступно моделей: {len(model_options)}")
        else:
            st.warning("Список моделей пуст — проверьте ключ.")
    else:
        st.caption("Введите API-ключ, чтобы загрузить список моделей.")

    st.divider()
    st.subheader("Источники для поиска")
    contact_email_input = st.text_input(
        "Email для научных API (необязательно)",
        placeholder="you@example.com",
        help="Некоторые API (Crossref, Europe PMC, PubMed) реже отказывают в "
             "обслуживании, если передавать реальный email — это их 'вежливый пул'. "
             "Ключ/пароль здесь не нужен, только адрес."
    )
    if contact_email_input.strip():
        CONTACT_EMAIL = contact_email_input.strip()

    selected_sources = st.multiselect(
        "Выберите базы", list(SOURCES.keys()),
        default=["PubMed / NCBI (медицина, химия, биология)",
                 "Semantic Scholar (200M+ работ)",
                 "Crossref (метаданные DOI + Unpaywall)",
                 "Europe PMC (PMC + Европейская коллекция)"],
    )
    per_source = st.slider("Источников на каждую базу", 2, 10, 4)
    check_oa = st.checkbox(
        "Проверять открытый доступ (Unpaywall) для найденных DOI", value=True,
        help="Ищет легальную бесплатную копию, если статья из платного журнала."
    )
    bilingual_search = st.checkbox(
        "Искать также на английском (авто-перевод темы через ту же ИИ-модель)",
        value=True,
        help="Тема будет переведена на английский (1–2 варианта устоявшейся англ. "
             "терминологии) и передана в каждый источник дополнительно к русскому "
             "запросу. Результаты объединяются с удалением дубликатов по DOI/URL/"
             "названию. Обзор всё равно пишется на русском — модель переводит "
             "факты из англоязычных аннотаций при синтезе. Требует 1 доп. запроса "
             "к ИИ-модели (короткий, дешёвый). Если модель не выдала перевод — "
             "работает только русский поиск."
    )
    semantic_scholar_key = st.text_input(
        "API-ключ Semantic Scholar (необязательно)", type="password",
        help="Без ключа запрос идёт в общий анонимный пул лимитов, который делят "
             "все бесключевые запросы к Semantic Scholar в мире — на нём часто "
             "ловится 429 из-за чужого трафика, особенно с общих IP облачных "
             "хостингов. Бесплатный личный ключ снимает эту проблему. Форма: "
             "semanticscholar.org/product/api#api-key-form (ключ приходит на "
             "почту, иногда через несколько дней).",
    )

    st.divider()
    st.subheader("Патентные базы (опционально)")
    st.caption(
        "У обеих баз нет полностью бесплатного анонимного доступа (проверено — "
        "полностью открытого API патентов сейчас не существует). "
        "**EPO OPS/Espacenet (рекомендуется)**: developers.epo.org/user/register → "
        "доступ 'Non-paying' → подтвердить email → 'My Apps' → 'Add a new App' → "
        "получите Consumer Key/Secret (~5 минут, 3.5 ГБ/неделю бесплатно). "
        "**PatentsView**: форма на ключ по ссылке ниже, но по свежим отчётам их "
        "сервис поддержки и сам API нестабильны — используйте как резервный, не единственный."
    )
    PATENT_CRED_HELP = {
        "epo_consumer_key": "developers.epo.org/user/register → 'Non-paying' → "
                             "подтвердить email → 'My Apps' → 'Add a new App'.",
        "epo_consumer_secret": "Выдаётся в той же карточке приложения, что и Consumer Key "
                                "(вкладка 'Keys' в 'My Apps').",
        "patentsview_api_key": "patentsview-support.atlassian.net/servicedesk/customer/"
                                "portal/1/create/18 — форма 'Request a PatentSearch API Key'. "
                                "Сервис нестабилен, используйте как резервный вариант.",
    }
    selected_patent_sources = st.multiselect(
        "Выберите патентные базы", list(PATENT_SOURCES.keys()), default=[]
    )
    patent_credentials = {}
    for name in selected_patent_sources:
        _fn, cred_fields = PATENT_SOURCES[name]
        for label, storage_key, _kwarg_name in cred_fields:
            patent_credentials[storage_key] = st.text_input(
                f"{label} — {name}", type="password", key=f"patcred_{storage_key}",
                help=PATENT_CRED_HELP.get(storage_key),
            )


tab1, tab2 = st.tabs(["📚 Обзор", "⚙️ Настройка промта"])

with tab1:
    temperature = st.slider(
        "Температура ответа", min_value=0.0, max_value=1.0, value=0.5, step=0.05,
        help="Ниже (0.0–0.3) — более точный, предсказуемый, "
             "\"academic\" ответ. Выше (0.7–1.0) — больше вариативности и "
             "\"творчества\", но выше риск неточностей."
    )

    topic = st.text_area(
        "Тема литературного обзора",
        placeholder="Например: Барьеры перехода на электромобили в Европе (2023–2025): "
                    "инфраструктура, цена, меры господдержки",
        height=100,
    ).strip()

    go = st.button(
        "🔎 Сформировать обзор", type="primary",
        disabled=not (topic and api_key and model_id and selected_sources),
    )

    if not api_key:
        st.info("Введите API-ключ в боковой панели, чтобы начать.")
    elif not model_id:
        st.info("Выберите модель в боковой панели, чтобы начать.")

    if go:
        is_non_latin = any(ord(ch) > 127 for ch in topic)

        # Формируем список поисковых запросов: русский оригинал + англ. варианты (если включено)
        search_queries = [topic]
        if bilingual_search and is_non_latin:
            with st.spinner("Перевожу тему на английский для расширенного поиска…"):
                try:
                    en_queries = translate_topic_to_english(topic, provider, model_id, api_key)
                except Exception as exc:
                    en_queries = []
                    st.warning(f"Не удалось перевести тему на английский ({exc}). "
                                "Поиск пойдёт только по русскому запросу.")
            if en_queries:
                search_queries.extend(en_queries)
                st.info("🌐 Двуязычный поиск. Запросы: " +
                        " · ".join(f"«{q}»" for q in search_queries))
            else:
                st.info("Перевод не получен, поиск только по русскому запросу.")
        elif bilingual_search and not is_non_latin:
            # Тема уже на латинице — возможно, уже английская. Не дёргаем LLM ради перевода.
            st.info("Тема уже на латинице, дополнительный перевод не требуется.")

        if is_non_latin and "arXiv (препринты, тема на англ.)" in selected_sources \
                and len(search_queries) == 1:
            st.info(
                "Тема на нелатинице, а arXiv — англоязычная база препринтов: его "
                "полнотекстовый поиск плохо понимает смысл нелатинских запросов и "
                "может вернуть случайные совпадения по отдельным словам. "
                "Включите «Искать также на английском» — arXiv тогда получит "
                "англоязычный вариант темы."
            )

        all_items = []
        n_queries = len(search_queries)
        total_steps = (len(selected_sources) + len(selected_patent_sources)) * n_queries
        progress = st.progress(0.0, text="Поиск источников…")
        SS_NAME = "Semantic Scholar (200M+ работ)"
        step = 0
        # Уменьшаем per_source при двуязычном поиске, чтобы итоговое количество
        # источников на базу не раздувалось в N раз (после дедупа всё равно
        # получится около нужного количества уникальных).
        eff_per_source = max(2, per_source // n_queries) if n_queries > 1 else per_source

        for q in search_queries:
            for name in selected_sources:
                progress.progress(step / max(total_steps, 1),
                                  text=f"Ищу в {name} ({q[:40]}…)")
                step += 1
                if name in SOURCES:
                    if name == SS_NAME:
                        found = SOURCES[name](q, eff_per_source, api_key=semantic_scholar_key)
                    else:
                        found = SOURCES[name](q, eff_per_source)
                    all_items.extend(found)

            for name in selected_patent_sources:
                progress.progress(step / max(total_steps, 1),
                                  text=f"Ищу в {name} ({q[:40]}…)")
                step += 1
                fn, cred_fields = PATENT_SOURCES[name]
                kwargs = {
                    kwarg_name: patent_credentials.get(storage_key, "")
                    for _label, storage_key, kwarg_name in cred_fields
                }
                found, error = fn(q, eff_per_source, **kwargs)
                if error:
                    st.warning(error)
                all_items.extend(found)
        progress.progress(1.0, text="Поиск завершён")

        # Дедупликация: одна и та же статья могла всплыть и по русскому, и по англ. запросу
        before = len(all_items)
        all_items = _dedup_items(all_items)
        if before != len(all_items):
            st.caption(f"Удалено дубликатов: {before - len(all_items)} "
                       f"(одна статья найдена по нескольким языковым запросам).")

        # Фолбэк: если выбранные источники ничего не дали (например, все словили
        # 429/сбой одновременно) — автоматически пробуем остальные бесплатные базы,
        # прежде чем сдаваться. Так временный сбой одного API не обнуляет весь обзор.
        if not all_items:
            remaining = [n for n in SOURCES if n not in selected_sources]
            if remaining:
                st.info(
                    "Выбранные базы не вернули результатов (возможно, временный сбой "
                    "или лимит запросов). Пробую остальные бесплатные базы автоматически…"
                )
                for q in search_queries:
                    for name in remaining:
                        if name == SS_NAME:
                            found = SOURCES[name](q, eff_per_source, api_key=semantic_scholar_key)
                        else:
                            found = SOURCES[name](q, eff_per_source)
                        all_items.extend(found)
                all_items = _dedup_items(all_items)

        if check_oa:
            with st.spinner("Проверяю открытый доступ по DOI (Unpaywall)…"):
                for it in all_items:
                    if it.get("doi"):
                        oa_url = check_unpaywall(it["doi"])
                        if oa_url:
                            it["url"] = oa_url
                            it["source"] += " [открытая копия найдена]"

        if not all_items:
            st.error("Ничего не найдено. Попробуйте переформулировать тему или выбрать другие источники.")
            st.stop()

        with st.expander(f"📎 Найдено источников: {len(all_items)} (нажмите, чтобы посмотреть)"):
            for i, it in enumerate(all_items, 1):
                st.markdown(f"**[{i}] {it['title']}** ({it['year']}) — *{it['source']}*  \n{it['url']}")

        with st.spinner("ИИ формирует обзор литературы…"):
            user_msg = build_user_message(topic, all_items)
            try:
                active_prompt = st.session_state.get("system_prompt", REVIEW_SYSTEM_PROMPT)
                review = call_llm(provider, model_id, active_prompt, user_msg, api_key, temperature=temperature)
            except Exception as exc:
                st.error(f"Ошибка обращения к модели: {exc}")
                st.stop()

        # Проверяем что обзор реально сформирован (не пустая строка)
        if not review or len(review.strip()) < 300:
            n_with_abstract = sum(1 for it in all_items if it.get("abstract","").strip())
            if n_with_abstract == 0:
                st.error(
                    "❌ Ни один источник не вернул аннотацию — модель не может сформировать "
                    "содержательный обзор из пустых полей. "
                    "Попробуйте: добавить Europe PMC / Semantic Scholar / arXiv (они возвращают "
                    "полные аннотации), или сформулировать тему на английском языке для лучшего "
                    "покрытия в международных базах."
                )
            else:
                st.error(
                    "❌ Модель вернула пустой ответ. Попробуйте: повысить температуру (0.5–0.7), "
                    "выбрать другую модель, или уточнить тему."
                )
            st.stop()

        st.success("Обзор готов!")
        st.markdown(review)

        docx_buf = build_docx(topic, review)
        st.download_button(
            "⬇️ Скачать как Word (.docx)",
            data=docx_buf,
            file_name=f"literature_review_{date.today().isoformat()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    st.divider()
    st.caption(
        "Демо-версия. Научные базы: PubMed, Europe PMC, Semantic Scholar, Crossref, "
        "DOAJ, arXiv. Патентные базы (опционально, нужны бесплатные ключ/регистрация): "
        "USPTO PatentsView, EPO OPS (Espacenet). "
        "Полный список используемых баз и логика проверки цитирования — в промт-шаблоне проекта."
    )

with tab2:
    st.subheader("Системный промт для генерации обзора")
    st.caption(
        "Здесь задаётся инструкция, по которой модель формирует обзор литературы. "
        "Можно скорректировать структуру, стиль, требования к цитированию и т.д."
    )

    if "system_prompt" not in st.session_state:
        st.session_state["system_prompt"] = REVIEW_SYSTEM_PROMPT

    edited_prompt = st.text_area(
        "Промт", value=st.session_state["system_prompt"], height=500,
        key="prompt_editor",
    )

    col1, col2 = st.columns(2)
    with col1:
        if st.button("💾 Сохранить промт"):
            st.session_state["system_prompt"] = edited_prompt
            st.success("Промт сохранён и будет использован при следующей генерации.")
    with col2:
        if st.button("↩️ Сбросить к варианту по умолчанию"):
            st.session_state["system_prompt"] = REVIEW_SYSTEM_PROMPT
            st.rerun()
