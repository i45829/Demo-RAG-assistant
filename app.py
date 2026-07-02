"""
app.py — Демо-система «Литературный обзор по теме» для презентации руководству.

ЧТО ДЕЛАЕТ:
  1. Пользователь вводит тему в браузере (никакой установки — просто веб-страница).
  2. Приложение ищет источники через бесплатные научные API (OpenAlex, arXiv,
     Europe PMC, DOAJ, Crossref + Unpaywall для платных статей).
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


def search_semanticscholar(query, max_results=5):
    """Semantic Scholar — 200M+ работ. Без ключа лимит ~1 req/s с shared IP (Streamlit Cloud
    это бьёт). С бесплатным ключом лимит ~10 req/s. Ключ: semanticscholar.org/product/api.
    Ключ читается из переменной SS_API_KEY или из session_state['ss_api_key']."""
    ss_key = (os.environ.get("SS_API_KEY") or
              st.session_state.get("ss_api_key", "")).strip()
    headers = {}
    if ss_key:
        headers["x-api-key"] = ss_key
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
        st.warning(f"Semantic Scholar: не удалось получить данные ({exc})"
                   + ("" if ss_key else " — попробуйте добавить бесплатный ключ S2 в настройки источников"))
        return []


def search_openaire(query, max_results=5):
    """OpenAIRE — европейский агрегатор открытого доступа (85M+ публикаций),
    без ключа, хорошо покрывает Европу, ВОЗ, ООН, европейские университеты.
    Полноценная замена/дополнение к Semantic Scholar без ограничений по IP."""
    try:
        r = _get(
            "https://api.openaire.eu/search/publications",
            params={"keywords": query, "format": "json",
                    "size": max_results, "sortBy": "relevancescore,descending"},
            timeout=25,
        )
        r.raise_for_status()
        data = r.json()
        results = (data.get("response", {})
                   .get("results", {})
                   .get("result", []) or [])
        items = []
        for res in results[:max_results]:
            meta = res.get("metadata", {}).get("oaf:entity", {}).get("oaf:result", {})
            title_raw = meta.get("title", {})
            title = (title_raw.get("$") if isinstance(title_raw, dict)
                     else (title_raw[0].get("$") if isinstance(title_raw, list) and title_raw else "")) or "(без названия)"
            date = meta.get("dateofacceptance", {})
            year = (date.get("$", "н/д") if isinstance(date, dict) else str(date))[:4]
            creators = meta.get("creator", [])
            if isinstance(creators, dict):
                creators = [creators]
            authors = ", ".join(c.get("$", "") for c in creators[:5]) or "н/д"
            pids = meta.get("pid", [])
            if isinstance(pids, dict):
                pids = [pids]
            doi = next((p.get("$","") for p in pids
                        if isinstance(p, dict) and p.get("@classid","").lower() == "doi"), "")
            best_url = f"https://doi.org/{doi}" if doi else ""
            descr = meta.get("description", {})
            abstract = _trim(descr.get("$","") if isinstance(descr, dict) else "")
            items.append({
                "title": title, "authors": authors, "year": year,
                "source": "OpenAIRE", "url": best_url, "doi": doi, "abstract": abstract,
            })
        return items
    except Exception as exc:
        st.warning(f"OpenAIRE: не удалось получить данные ({exc})")
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


# ============================== ПАТЕНТНЫЕ БАЗЫ (БЕЗ КЛЮЧЕЙ) ==============================

def search_patentsview_legacy(query, max_results=5):
    """USPTO PatentsView — legacy endpoint, НЕ требует ключа (подтверждено 2026).
    Покрывает все выданные патенты США с 1976 года + аннотации."""
    try:
        r = _post(
            "https://api.patentsview.org/patents/query",
            json={
                "q": {"_text_any": {"patent_abstract": query}},
                "f": ["patent_number", "patent_title", "patent_date",
                      "assignee_organization", "patent_abstract", "inventor_last_name"],
                "o": {"per_page": max_results},
                "s": [{"patent_date": "desc"}],
            },
            timeout=25,
        )
        r.raise_for_status()
        items = []
        for p in (r.json().get("patents") or [])[:max_results]:
            pn = p.get("patent_number") or ""
            items.append({
                "title": p.get("patent_title") or f"US Patent {pn}",
                "authors": p.get("inventor_last_name") or "н/д",
                "year": (p.get("patent_date") or "н/д")[:4],
                "source": "USPTO PatentsView",
                "url": f"https://worldwide.espacenet.com/patent/search?q=pn%3DUS{pn}",
                "doi": "",
                "abstract": _trim(p.get("patent_abstract") or ""),
            })
        return items
    except Exception as exc:
        st.warning(f"USPTO PatentsView: не удалось получить данные ({exc})")
        return []


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


def search_wipo_patentscope(query, max_results=5, token=""):
    """WIPO PATENTSCOPE — международные патенты (PCT + 100+ ведомств).
    Бесплатный доступ через REST API с токеном. Токен: patentscope.wipo.int → API Tools.
    При пустом токене пробует публичный эндпоинт (ограниченные данные)."""
    headers = {"Accept": "application/json"}
    if token.strip():
        headers["Authorization"] = f"Bearer {token.strip()}"
    try:
        r = _get(
            "https://patentscope.wipo.int/api/v1/pct/patents",
            params={"q": query, "maxCount": max_results, "offset": 0},
            headers=headers,
            timeout=25,
        )
        r.raise_for_status()
        data = r.json()
        items = []
        for d in (data.get("results") or data.get("patents") or [])[:max_results]:
            pnum = d.get("applicationNumber") or d.get("pctNumber") or ""
            title_obj = d.get("IMPACT_EN") or d.get("invention_title") or {}
            title = title_obj if isinstance(title_obj, str) else (
                title_obj.get("text") if isinstance(title_obj, dict) else str(d.get("title",""))
            ) or "(без названия)"
            items.append({
                "title": title,
                "authors": ", ".join(d.get("inventors", []))[:3] if d.get("inventors") else "н/д",
                "year": str(d.get("filingDate","н/д"))[:4],
                "source": "WIPO PATENTSCOPE",
                "url": f"https://patentscope.wipo.int/search/en/detail.jsf?docId={pnum}" if pnum else "",
                "doi": "",
                "abstract": _trim(str(d.get("abstract",""))),
            })
        return items, None
    except Exception as exc:
        return [], f"WIPO PATENTSCOPE: ошибка запроса ({exc}). Если запрос без токена — доступ может быть ограничен."


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
    "USPTO PatentsView (патенты США, без ключа)": (
        search_patentsview_legacy, []
    ),
    "WIPO PATENTSCOPE (PCT + 100 ведомств, токен опционален)": (
        search_wipo_patentscope,
        [("Access Token WIPO (оставьте пустым для публичного доступа)", "wipo_token")],
    ),
    "Espacenet / EPO OPS (бесплатная регистрация, Consumer Key+Secret)": (
        search_epo_ops_simple,
        [("Consumer Key (EPO)", "epo_key"), ("Consumer Secret (EPO)", "epo_secret")],
    ),
}



# PatentsView возвращает (items, None) / ([], msg) — адаптируем в обёртку
def _patentsview_wrap(query, max_results=5):
    """Обёртка, чтобы PatentsView вписывался в обычный SOURCES (возвращает list)."""
    result = search_patentsview_legacy(query, max_results)
    return result  # search_patentsview_legacy уже возвращает list и сам вызывает st.warning

SOURCES = {
    "PubMed / NCBI (медицина, химия, биология)": search_pubmed,
    "Europe PMC (PMC + Европейская коллекция)": search_europepmc,
    "OpenAIRE (85M+ открытых публикаций, без ключа)": search_openaire,
    "Semantic Scholar (200M+ работ, опц. ключ)": search_semanticscholar,
    "Crossref (метаданные DOI + Unpaywall)": search_crossref,
    "DOAJ (открытые журналы)": search_doaj,
    "arXiv (препринты, тема на англ.)": search_arxiv,
}

# ============================== ИЗВЛЕЧЕНИЕ ПОИСКОВЫХ ЗАПРОСОВ ==============================
# Источники, для которых нужен английский запрос (они плохо ищут по русским словам)
ENGLISH_CENTRIC = {
    "PubMed / NCBI (медицина, химия, биология)",
    "Semantic Scholar (200M+ работ, опц. ключ)",
    "OpenAIRE (85M+ открытых публикаций, без ключа)",
    "arXiv (препринты, тема на англ.)",
    "DOAJ (открытые журналы)",
}


def extract_search_queries(topic: str, api_key: str, provider: str, model_id: str) -> dict:
    """Вызывает LLM чтобы извлечь ключевые научные термины из темы на русском и английском.
    Возвращает {"ru": "...", "en": "..."}.
    Это самый важный шаг: без него keyword-поиск по полной фразе темы
    находит мусор (совпадение по случайным словам вроде «тренды», «анализ», «год»)."""
    system = (
        "Ты — помощник для научного поиска. "
        "Извлеки из темы исследования 2-4 ключевых научных термина. "
        "Верни ТОЛЬКО JSON без Markdown, без пояснений. "
        'Схема: {"ru": "термины на русском", "en": "same in English"}. '
        "Примеры: "
        'тема «Сверхкритическая вода. Научные тренды 2023-2026» -> {"ru": "сверхкритическая вода", "en": "supercritical water"}; '
        'тема «Масляные дисперсии как препаративная форма для СЗР» -> {"ru": "масляные дисперсии агрохимия", "en": "oil dispersion agrochemical formulation"}. '
        "ТОЛЬКО JSON."
    )
    try:
        raw = call_llm(provider, model_id, system, topic, api_key, temperature=0.0)
        import json as _json, re as _re
        # Убираем возможные markdown-бэктики
        clean = _re.sub(r"```[a-z]*", "", raw).strip().strip("`")
        parsed = _json.loads(clean)
        ru = str(parsed.get("ru", topic)).strip() or topic
        en = str(parsed.get("en", topic)).strip() or topic
        return {"ru": ru, "en": en}
    except Exception:
        # Фолбэк: убираем самые частые «шумовые» слова из русского текста
        import re as _re
        stop_words = ["научные", "тренды", "обзор", "анализ", "исследование", "года", "год"]
        stop_patterns = [r"20\d{2}[–\-]20\d{2}", r"20\d{2}"]
        clean = topic
        for sw in stop_words:
            clean = _re.sub(sw, "", clean, flags=_re.IGNORECASE)
        for pat in stop_patterns:
            clean = _re.sub(pat, "", clean)
        clean = " ".join(clean.split())  # убираем лишние пробелы
        return {"ru": clean or topic, "en": clean or topic}


# ============================== СИНТЕЗ ОБЗОРА (LLM) ==============================

REVIEW_SYSTEM_PROMPT = """Ты — научный аналитик, готовящий раздел «Обзор литературы» для диссертации по естественнонаучной/технической теме.

ШАГ 1 — ФИЛЬТРАЦИЯ (молча, до написания текста):
Оцени каждый источник: относится ли он к теме по СУЩЕСТВУ, а не по случайному слову.
Если источник явно не по теме (например, экономический анализ при теме по физической химии) — исключи его полностью и молча. В итоговый текст и список литературы он не попадает. Никогда не пытайся «натянуть» нерелевантный источник на тему.

ШАГ 2 — НАПИСАНИЕ ОБЗОРА (по отфильтрованным источникам):
Пиши только то, что реально есть в аннотациях релевантных источников. Извлекай:
- точные названия веществ, материалов, соединений, реагентов, методов, приборов;
- химические формулы и обозначения — буквально как в источнике;
- числа: температуры, давления, концентрации, pH, эффективность, проценты, p-значения;
- механизмы и принципы действия — как описаны в источнике;
- конкретные результаты экспериментов с числами.

ЗАПРЕЩЕНО в теле обзора:
- «было найдено X источников», «источники были проанализированы», «проведён поиск»;
- обобщения без цифр и ссылок; пересказ того, что ты делал.

ГАРАНТИЯ НЕПУСТОГО ВЫВОДА:
- Если после фильтрации осталось ≥1 релевантного источника — напиши полноценный обзор по нему/ним.
- Если релевантных источников не осталось совсем — напиши в разделе «Обзор литературы» одно честное предложение об этом, затем в разделе «Выводы» перечисли, что именно нужно искать и в каких базах.
- Никогда не возвращай пустой текст.

Каждое фактическое утверждение — со ссылкой [n] на номер источника.

Структура ответа (строго Markdown, на русском языке):
## Введение
2–4 предложения: тема, актуальность, границы обзора. Без канцелярских клише.

## Обзор литературы
Тематический синтез конкретных данных из источников. Ссылка [n] на каждый факт. Подзаголовки ### по подтемам при необходимости.

## Выводы и направления дальнейших исследований
- Обоснованные выводы из данных [n]
- Пробелы в изученности темы (каких данных не хватает)
- Рекомендации (явно помечены как предположения)

## Список литературы
[n] Авторы. Название. Год. Источник. URL."""


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
    r = httpx.post(
        f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
        params={"key": api_key},
        json={
            "system_instruction": {"parts": [{"text": system_prompt}]},
            "contents": [{"role": "user", "parts": [{"text": user_message}]}],
            "generationConfig": {"temperature": temperature},
        },
        timeout=120,
    )
    r.raise_for_status()
    data = r.json()
    return data["candidates"][0]["content"]["parts"][0]["text"]


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
        help="Некоторые API (OpenAlex, Crossref, Europe PMC) реже отказывают в "
             "обслуживании, если передавать реальный email — это их 'вежливый пул'. "
             "Ключ/пароль здесь не нужен, только адрес."
    )
    if contact_email_input.strip():
        CONTACT_EMAIL = contact_email_input.strip()

    selected_sources = st.multiselect(
        "Выберите базы", list(SOURCES.keys()),
        default=["PubMed / NCBI (медицина, химия, биология)",
                 "OpenAIRE (85M+ открытых публикаций, без ключа)",
                 "Crossref (метаданные DOI + Unpaywall)"],
    )
    per_source = st.slider("Источников на каждую базу", 2, 10, 4)
    check_oa = st.checkbox(
        "Проверять открытый доступ (Unpaywall) для найденных DOI", value=True,
        help="Ищет легальную бесплатную копию, если статья из платного журнала."
    )

    if "Semantic Scholar (200M+ работ, опц. ключ)" in selected_sources:
        ss_key_input = st.text_input(
            "API-ключ Semantic Scholar (необязательно)",
            type="password",
            help="Без ключа Semantic Scholar даёт ~1 запрос/сек — "
                 "с общего IP Streamlit Cloud это вызывает 429. "
                 "Бесплатный ключ: semanticscholar.org/product/api → Request API Key. "
                 "Повышает лимит до ~10 req/s.",
        )
        if ss_key_input.strip():
            st.session_state["ss_api_key"] = ss_key_input.strip()

    # Все источники теперь в SOURCES, PATENT_SOURCES оставлен пустым
    patent_credentials = {}


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
        # Шаг 0: извлекаем ключевые научные термины из темы до поиска
        with st.spinner("🔍 Анализирую тему и формирую поисковые запросы…"):
            queries = extract_search_queries(topic, api_key, provider, model_id)

        with st.expander("🔎 Поисковые запросы к базам данных", expanded=False):
            st.markdown(
                f"**Русский запрос:** `{queries['ru']}`  \n"
                f"**Английский запрос:** `{queries['en']}`  \n"
                "_Запросы автоматически извлечены из темы, чтобы избежать случайных "
                "совпадений по вспомогательным словам (тренды, анализ, год и т.п.)_"
            )

        all_items = []
        progress = st.progress(0.0, text="Поиск источников…")
        for i, name in enumerate(selected_sources):
            progress.progress((i) / len(selected_sources), text=f"Ищу в {name}…")
            # Англоязычным базам — английский запрос, остальным — русский
            api_query = queries["en"] if name in ENGLISH_CENTRIC else queries["ru"]
            if name in SOURCES:
                found = SOURCES[name](api_query, per_source)
                all_items.extend(found)
        progress.progress(1.0, text="Поиск завершён")

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
        "Демо-версия. Источники: Europe PMC, Semantic Scholar, Crossref, arXiv, DOAJ, USPTO PatentsView. "
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
