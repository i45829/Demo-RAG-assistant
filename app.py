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


def _reconstruct_abstract(inverted_index):
    """OpenAlex хранит аннотацию инвертированным индексом {слово: [позиции]}."""
    if not inverted_index:
        return ""
    pos = []
    for word, idxs in inverted_index.items():
        for i in idxs:
            pos.append((i, word))
    pos.sort(key=lambda p: p[0])
    return " ".join(w for _, w in pos)


def search_openalex(query, max_results=5):
    try:
        r = _get(
            "https://api.openalex.org/works",
            params={"search": query, "per-page": max_results, "mailto": CONTACT_EMAIL},
            timeout=20,
        )
        r.raise_for_status()
        items = []
        for w in r.json().get("results", []):
            doi = w.get("doi") or ""
            url = (
                (w.get("open_access") or {}).get("oa_url")
                or (w.get("primary_location") or {}).get("landing_page_url")
                or doi
            )
            items.append({
                "title": w.get("title") or "(без названия)",
                "authors": ", ".join(
                    a["author"]["display_name"]
                    for a in (w.get("authorships") or [])[:5] if a.get("author")
                ) or "н/д",
                "year": w.get("publication_year") or "н/д",
                "source": "OpenAlex",
                "url": url or "",
                "doi": doi.replace("https://doi.org/", "") if doi else "",
                "abstract": _trim(_reconstruct_abstract(w.get("abstract_inverted_index"))),
            })
        return items
    except Exception as exc:
        st.warning(f"OpenAlex: не удалось получить данные ({exc})")
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
    """Semantic Scholar — 200M+ работ, без ключа (аноним. доступ ограничен по скорости,
    но для демо этого достаточно). Резервный/дублирующий источник на случай перебоев
    у OpenAlex — не полагаться на единственный источник надёжнее."""
    try:
        r = _get(
            "https://api.semanticscholar.org/graph/v1/paper/search",
            params={
                "query": query, "limit": max_results,
                "fields": "title,year,authors,abstract,externalIds,venue,url",
            },
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


def _epo_ops_token(client_key, client_secret):
    basic = base64.b64encode(f"{client_key}:{client_secret}".encode()).decode()
    r = _post(
        "https://ops.epo.org/3.2/auth/accesstoken",
        headers={"Authorization": f"Basic {basic}",
                 "Content-Type": "application/x-www-form-urlencoded"},
        data={"grant_type": "client_credentials"}, timeout=30,
    )
    r.raise_for_status()
    return r.json()["access_token"]


def search_epo_ops(query, max_results, client_key, client_secret):
    """Espacenet (данные EPO) через официальный Open Patent Services API.
    Бесплатная регистрация: developers.epo.org → My Apps → Consumer Key/Secret."""
    if not client_key or not client_secret:
        return [], "Espacenet (EPO OPS): не указан ключ/секрет — источник пропущен."
    try:
        token = _epo_ops_token(client_key, client_secret)
    except Exception as exc:
        return [], f"Espacenet (EPO OPS): ошибка авторизации ({exc})"
    cql = query if "=" in query else f'ti="{query}" or ab="{query}"'
    try:
        r = _get(
            "https://ops.epo.org/3.2/rest-services/published-data/search/biblio",
            params={"q": cql, "Range": f"1-{max_results}"},
            headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
            timeout=30,
        )
        r.raise_for_status()
        data = r.json()
    except Exception as exc:
        return [], f"Espacenet (EPO OPS): ошибка запроса ({exc})"

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
                "authors": "н/д",
                "year": "н/д",
                "source": "Espacenet (EPO OPS)",
                "url": f"https://worldwide.espacenet.com/patent/search?q=pn%3D{country}{num}",
                "doi": "",
                "abstract": "",
            })
    except Exception as exc:
        return [], f"Espacenet (EPO OPS): не удалось разобрать ответ ({exc})"
    return items, None


def search_patentsview(query, max_results, api_key):
    """USPTO PatentsView — патенты США. Бесплатный ключ: заявка через форму
    PatentsView Support Portal (search.patentsview.org)."""
    if not api_key:
        return [], "USPTO PatentsView: не указан ключ — источник пропущен."
    try:
        q = {"_text_any": {"patent_title": query}}
        f = ["patent_id", "patent_title", "patent_date"]
        o = {"size": max_results}
        r = _get(
            "https://search.patentsview.org/api/v1/patent/",
            params={"q": json.dumps(q), "f": json.dumps(f), "o": json.dumps(o)},
            headers={"X-Api-Key": api_key}, timeout=30,
        )
        r.raise_for_status()
        data = r.json()
    except Exception as exc:
        return [], f"USPTO PatentsView: ошибка запроса ({exc})"

    items = []
    for p in (data.get("patents") or [])[:max_results]:
        pid = p.get("patent_id", "")
        items.append({
            "title": p.get("patent_title") or f"Патент US{pid}",
            "authors": "н/д",
            "year": (p.get("patent_date") or "н/д")[:4],
            "source": "USPTO PatentsView",
            "url": f"https://worldwide.espacenet.com/patent/search?q=pn%3DUS{pid}",
            "doi": "",
            "abstract": "",
        })
    return items, None


def search_lens_patents(query, max_results, token):
    """Lens.org — глобальная патентная база (100+ ведомств: USPTO, EPO, WIPO и др.).
    Бесплатный доступ (научные/некоммерческие цели) — заявка на lens.org/lens/user/subscriptions,
    токен выдаётся не мгновенно, в отличие от остальных источников."""
    if not token:
        return [], "Lens.org: не указан токен — источник пропущен."
    try:
        r = _post(
            "https://api.lens.org/patent/search",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"query": query, "size": max_results, "include": ["biblio", "lens_id"]},
            timeout=30,
        )
        r.raise_for_status()
        data = r.json()
    except Exception as exc:
        return [], f"Lens.org: ошибка запроса ({exc})"

    items = []
    for d in (data.get("data") or [])[:max_results]:
        biblio = d.get("biblio") or {}
        title_field = biblio.get("invention_title")
        if isinstance(title_field, list):
            title = (title_field[0] or {}).get("text") if title_field else None
        else:
            title = title_field
        lens_id = d.get("lens_id", "")
        items.append({
            "title": title or f"Патент {lens_id}",
            "authors": "н/д",
            "year": "н/д",
            "source": "Lens.org",
            "url": f"https://www.lens.org/lens/patent/{lens_id}" if lens_id else "",
            "doi": "",
            "abstract": "",
        })
    return items, None


# name -> (функция, список (метка поля, session_key) для нужных реквизитов)
PATENT_SOURCES = {
    "Espacenet / EPO OPS (патенты, бесплатный ключ)": (
        search_epo_ops,
        [("Consumer Key (EPO)", "epo_key"), ("Consumer Secret (EPO)", "epo_secret")],
    ),
    "USPTO PatentsView (патенты США, бесплатный ключ)": (
        search_patentsview,
        [("API-ключ PatentsView", "patentsview_key")],
    ),
    "Lens.org (патенты, 100+ ведомств, токен по заявке)": (
        search_lens_patents,
        [("Access Token (Lens.org)", "lens_token")],
    ),
}


SOURCES = {
    "OpenAlex (научные работы)": search_openalex,
    "arXiv (препринты)": search_arxiv,
    "Europe PMC (PMC/PubMed)": search_europepmc,
    "DOAJ (открытые журналы)": search_doaj,
    "Semantic Scholar (резервный источник)": search_semanticscholar,
    "Crossref (+ проверка Unpaywall)": search_crossref,
}

# ============================== СИНТЕЗ ОБЗОРА (LLM) ==============================

REVIEW_SYSTEM_PROMPT = """Ты — научный аналитик, готовящий раздел «Обзор литературы» для диссертации
по естественнонаучной/технической теме. Твоя задача — извлечь и изложить КОНКРЕТНОЕ
научно-техническое содержание источников, а не рассказать о процессе поиска.

ГЛАВНОЕ ПРАВИЛО СОДЕРЖАНИЯ:
Обзор должен состоять из фактов предметной области, а не из метакомментариев о самих
источниках. Из каждой аннотации извлекай и включай в текст, если это в ней есть:
- точные названия веществ, материалов, соединений, реагентов, штаммов, методов,
  приборов, алгоритмов — как они названы в источнике, без упрощения и обобщения;
- химические формулы, обозначения соединений, единицы измерения — переноси как есть;
- количественные данные: концентрации, температуры, давления, эффективность, KPI,
  проценты, размеры выборки, статистическую значимость, диапазоны значений;
- механизмы, принципы действия, схемы процессов — как они описаны в источнике;
- конкретные результаты экспериментов или наблюдений, а не общие фразы вида
  "исследование показало положительный эффект".

СТРОГО ЗАПРЕЩЕНО (это не обзор, а имитация обзора):
- фразы уровня "было найдено N источников", "источники были проанализированы",
  "проведён поиск по теме", "рассмотрены различные аспекты" — такие фразы НЕ несут
  научной информации и не должны появляться в теле обзора;
- пересказ процесса своей работы вместо содержания источников;
- обобщения без опоры на конкретные данные источника ("многие исследования
  показывают" без указания, что именно и с какими цифрами показывает каждое [n]).

Используй ТОЛЬКО информацию из предоставленных источников. Никогда не выдумывай
факты, формулы, вещества, цифры или выводы, которых нет в материалах. Если аннотация
источника не содержит технических деталей (только общие слова) — честно отметь это
у данного источника, а не заполняй пробел общими фразами.

ФИЛЬТРАЦИЯ ПО РЕЛЕВАНТНОСТИ (обязательный внутренний шаг перед синтезом):
Источники присланы автоматическим поиском по ключевым словам и могут включать
случайные, не относящиеся к теме совпадения (особенно при многоязычных запросах —
поисковая система могла подобрать источник по формальному пересечению слов, а не по
смыслу). Прежде чем писать обзор, мысленно оцени каждый источник: относится ли он
РЕАЛЬНО к теме обзора по существу, а не только по случайному слову в названии.
- Источники, явно не относящиеся к теме, ПОЛНОСТЬЮ исключи: не цитируй их в тексте,
  не включай в «Список литературы», не упоминай, что они были исключены.
- Не сообщай в выводе, сколько источников отфильтровано — сделай это молча, обзор
  должен выглядеть так, будто нерелевантных источников не было вовсе.
- Если после фильтрации релевантных источников осталось мало или не осталось совсем —
  честно скажи об этом в начале раздела «Обзор литературы» одним предложением, вместо
  того чтобы притягивать нерелевантные источники к теме искусственно.

Каждое фактическое утверждение сопровождай числовой ссылкой [n], соответствующей
номеру источника в списке. Пиши как тематический синтез: где источники согласуются
в конкретных цифрах/веществах/механизмах, где расходятся — с указанием, в чём именно
расхождение (какие значения или подходы отличаются).

Структура ответа (строго Markdown, на русском языке):
## Введение
Тема и границы обзора — 2-4 предложения, по существу, без вводных клише.

## Обзор литературы
Тематический синтез КОНКРЕТНОГО содержания источников (вещества, формулы, параметры,
механизмы, результаты) со ссылками [n] на каждый факт. Если источники относятся к
разным подтемам — раздели подзаголовками уровня ###.

## Выводы и направления дальнейших исследований
- Обоснованные выводы с опорой на конкретные данные [n]
- Пробелы в изученности темы (что именно не покрыто — по конкретным параметрам/веществам)
- Рекомендации и гипотезы (явно помечены как предположение, а не факт источника)

## Список литературы
Нумерованный список: [n] Автор(ы). Название. Год. Источник."""


def build_user_message(topic, items):
    lines = [f"ТЕМА ОБЗОРА: {topic}", "", "ИСТОЧНИКИ (используй только их, ссылайся по номеру):", ""]
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

    all_source_names = list(SOURCES.keys()) + list(PATENT_SOURCES.keys())
    selected_sources = st.multiselect(
        "Выберите базы", all_source_names,
        default=["OpenAlex (научные работы)", "Semantic Scholar (резервный источник)",
                 "Crossref (+ проверка Unpaywall)"],
    )
    per_source = st.slider("Источников на каждую базу", 2, 10, 4)
    check_oa = st.checkbox(
        "Проверять открытый доступ (Unpaywall) для найденных DOI", value=True,
        help="Ищет легальную бесплатную копию, если статья из платного журнала."
    )

    selected_patent_sources = [s for s in selected_sources if s in PATENT_SOURCES]
    patent_credentials = {}  # {название базы: {session_key: значение}}
    if selected_patent_sources:
        st.divider()
        st.subheader("Ключи патентных API")
        st.caption(
            "Для каждой выбранной патентной базы нужен свой бесплатный ключ. "
            "Без ключа источник будет пропущен при формировании обзора."
        )
        for name in selected_patent_sources:
            _, cred_fields = PATENT_SOURCES[name]
            with st.expander(name, expanded=True):
                values = {}
                for field_label, session_key in cred_fields:
                    values[session_key] = st.text_input(
                        field_label, type="password", key=f"cred_{session_key}"
                    )
                patent_credentials[name] = values


tab1, tab2 = st.tabs(["📚 Обзор", "⚙️ Настройка промта"])

with tab1:
    temperature = st.slider(
        "Температура ответа", min_value=0.0, max_value=1.0, value=0.3, step=0.05,
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
        if is_non_latin and "arXiv (препринты)" in selected_sources:
            st.info(
                "Тема на нелатинице, а arXiv — англоязычная база препринтов: его "
                "полнотекстовый поиск плохо понимает смысл нелатинских запросов и "
                "может вернуть случайные совпадения по отдельным словам. "
                "Для тем не на английском точнее работают OpenAlex, Crossref, "
                "Semantic Scholar и Europe PMC — у них многоязычные метаданные."
            )
        all_items = []
        progress = st.progress(0.0, text="Поиск источников…")
        for i, name in enumerate(selected_sources):
            progress.progress((i) / len(selected_sources), text=f"Ищу в {name}…")
            if name in SOURCES:
                found = SOURCES[name](topic, per_source)
                all_items.extend(found)
            elif name in PATENT_SOURCES:
                fn, cred_fields = PATENT_SOURCES[name]
                creds = patent_credentials.get(name, {})
                args = [creds.get(session_key) for _, session_key in cred_fields]
                found, err = fn(topic, per_source, *args)
                if err:
                    st.warning(err)
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
        "Демо-версия. Источники: OpenAlex, arXiv, Europe PMC, DOAJ, Crossref (+ Unpaywall). "
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
