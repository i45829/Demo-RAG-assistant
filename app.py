import io
import os
from datetime import date

import httpx
import streamlit as st
from docx import Document
from docx.shared import Pt

# ============================== НАСТРОЙКИ СТРАНИЦЫ ==============================

st.set_page_config(page_title="Литературный обзор — демо", page_icon="📚", layout="wide")

CONTACT_EMAIL = "demo@example.com"  # свой email — нужен для "вежливого" доступа к API

# ============================== ПОИСК ПО ОТКРЫТЫМ API ==============================


def _trim(text, limit=600):
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
        r = httpx.get(
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
        r = httpx.get(
            "http://export.arxiv.org/api/query",
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
        r = httpx.get(
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
        r = httpx.get(
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


def search_crossref(query, max_results=5):
    """Метаданные почти любого DOI, включая закрытые журналы (ScienceDirect и т.п.)."""
    try:
        r = httpx.get(
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
        r = httpx.get(f"https://api.unpaywall.org/v2/{doi}",
                       params={"email": CONTACT_EMAIL}, timeout=15)
        r.raise_for_status()
        data = r.json()
        if not data.get("is_oa"):
            return None
        best = data.get("best_oa_location") or {}
        return best.get("url_for_pdf") or best.get("url")
    except Exception:
        return None


SOURCES = {
    "OpenAlex (научные работы)": search_openalex,
    "arXiv (препринты)": search_arxiv,
    "Europe PMC (PMC/PubMed)": search_europepmc,
    "DOAJ (открытые журналы)": search_doaj,
    "Crossref (+ проверка Unpaywall)": search_crossref,
}

# ============================== СИНТЕЗ ОБЗОРА (LLM) ==============================

REVIEW_SYSTEM_PROMPT = """Ты — научный аналитик, готовящий раздел «Обзор литературы» для диссертации.
Используй ТОЛЬКО информацию из предоставленных источников. Никогда не выдумывай факты,
авторов, цифры или выводы, которых нет в материалах. Если по какому-то аспекту темы
источников не хватает — прямо скажи об этом, не додумывай.

Каждое фактическое утверждение сопровождай числовой ссылкой [n], соответствующей номеру
источника в списке. Пиши академическим языком, как тематический синтез (что говорят
источники, где они согласуются или расходятся), а не как пересказ по очереди.

Структура ответа (строго Markdown, на русском языке):
## Введение
Тема, границы обзора (2-4 предложения).

## Обзор литературы
Тематический синтез с ссылками [n] на каждое фактическое утверждение.

## Выводы и направления дальнейших исследований
- Обоснованные выводы [n]
- Пробелы в изученности темы
- Рекомендации и гипотезы (с пометкой, что это предположение, а не факт источника)

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


def call_gemini(system_prompt, user_message, api_key, model="gemini-2.0-flash"):
    r = httpx.post(
        f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
        params={"key": api_key},
        json={
            "system_instruction": {"parts": [{"text": system_prompt}]},
            "contents": [{"role": "user", "parts": [{"text": user_message}]}],
        },
        timeout=120,
    )
    r.raise_for_status()
    data = r.json()
    return data["candidates"][0]["content"]["parts"][0]["text"]


def call_polza(system_prompt, user_message, api_key, model):
    """Polza.ai — агрегатор моделей с OpenAI-совместимым API.
    Формат model: 'provider/model', например 'openai/gpt-4o', 'anthropic/claude-3.7-sonnet'."""
    r = httpx.post(
        "https://polza.ai/api/v1/chat/completions",
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json={
            "model": model,
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


def call_llm(provider, model_id, system_prompt, user_message, api_key):
    if provider == "Google Gemini":
        return call_gemini(system_prompt, user_message, api_key, model=model_id)
    return call_polza(system_prompt, user_message, api_key, model=model_id)


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
    selected_sources = st.multiselect(
        "Выберите базы", list(SOURCES.keys()), default=list(SOURCES.keys())[:3]
    )
    per_source = st.slider("Источников на каждую базу", 2, 10, 4)
    check_oa = st.checkbox(
        "Проверять открытый доступ (Unpaywall) для найденных DOI", value=True,
        help="Ищет легальную бесплатную копию, если статья из платного журнала."
    )

topic = st.text_area(
    "Тема литературного обзора",
    placeholder="Например: Барьеры перехода на электромобили в Европе (2023–2025): "
                "инфраструктура, цена, меры господдержки",
    height=100,
)

go = st.button(
    "🔎 Сформировать обзор", type="primary",
    disabled=not (topic and api_key and model_id and selected_sources),
)

if not api_key:
    st.info("Введите API-ключ в боковой панели, чтобы начать.")
elif not model_id:
    st.info("Выберите модель в боковой панели, чтобы начать.")

if go:
    all_items = []
    progress = st.progress(0.0, text="Поиск источников…")
    for i, name in enumerate(selected_sources):
        progress.progress((i) / len(selected_sources), text=f"Ищу в {name}…")
        found = SOURCES[name](topic, per_source)
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
            review = call_llm(provider, model_id, REVIEW_SYSTEM_PROMPT, user_msg, api_key)
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
