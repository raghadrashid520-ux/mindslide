
# -*- coding: utf-8 -*-
"""
JUST StudyMate — مساعد دراسي لطلاب جامعة العلوم والتكنولوجيا الأردنية
التشغيل:  streamlit run app.py
"""
from __future__ import annotations

import html
import json
import os
import random
import re
import secrets
import time
from datetime import date, datetime, timedelta
from io import BytesIO
from typing import Any

import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from pypdf import PdfReader

st.set_page_config(page_title="JUST StudyMate", page_icon="🎓", layout="wide")

# ───────────────────────────── الإعدادات ─────────────────────────────
MAX_CHARS = 100_000
SHARE_TTL_HOURS = 48
# سلسلة النماذج الاحتياطية (يمكن تغييرها من st.secrets["MODEL_CHAIN"] مفصولة بفواصل)
DEFAULT_MODELS = ["gemini-3.5-flash", "gemini-3.1-flash-lite", "gemini-2.5-flash"]
# ⚠️ جدول التقديرات قابل للتعديل من داخل التطبيق — طابقه مع نشرة القبول والتسجيل في JUST
DEFAULT_GRADE_SCALE = {"A": 4.0, "B+": 3.5, "B": 3.0, "C+": 2.5, "C": 2.0, "D": 1.0, "F": 0.0}
DAYS = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"]
TASK_TYPES = ["واجب", "كويز", "امتحان", "مشروع", "أخرى"]
TASK_COLS = ["المقرر", "النوع", "الموعد", "مكتمل"]
SCHED_COLS = ["اليوم", "من", "إلى", "المقرر", "المكان"]
GPA_COLS = ["المقرر", "الساعات", "التقدير"]
ARABIC_RE = re.compile(r"[\u0600-\u06FF]")
FONT_PATHS = [
    "fonts/NotoNaskhArabic-Regular.ttf", "fonts/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    "/Library/Fonts/Arial Unicode.ttf", "C:/Windows/Fonts/arial.ttf",
]

PROFESSOR_SYSTEM = (
    "أنت «أستاذ المادة»: أستاذ جامعي أكاديمي، صبور وداعم ومحفّز. تخاطب الطالب بدفء واحترام، "
    "تشرح السبب والنتيجة والعلاقة بين المفاهيم، وتشجّعه دون مبالغة. "
    "التزم بنص المحاضرة ولا تخترع معلومات غير موجودة فيه؛ وإن كان النص ناقصاً فصرّح بذلك."
)

ANALYSIS_INSTRUCTION = """
حلّل المحاضرة التالية تحليلاً أكاديمياً عميقاً ومفيداً للمذاكرة.
أعد JSON صحيحاً فقط (بلا markdown وبلا أي نص خارجه) بالمفاتيح التالية حرفياً:
{
  "detected_language": "ar أو en أو غيرها",
  "title": "عنوان المحاضرة",
  "overview": "ملخص تنفيذي من 3 إلى 5 جمل",
  "key_takeaways": ["أهم فكرة"],
  "detailed_explanation": [{"heading": "...", "explanation": "...", "examples": ["..."]}],
  "terminology": [{"term": "المصطلح بلغته الأصلية", "translation": "ترجمته للغة الأخرى (عربي/إنجليزي)", "definition": "تعريف دقيق", "example": "مثال"}],
  "mind_map": {"central_topic": "...", "branches": [{"title": "...", "points": ["..."]}]},
  "quiz": [{"question": "...", "difficulty": "سهل أو متوسط أو صعب", "options": ["أ","ب","ج","د"],
            "correct_index": 0, "correct_answer": "نص الإجابة الصحيحة", "explanation": "..."}],
  "essay_questions": [{"question": "...", "difficulty": "سهل أو متوسط أو صعب", "model_answer": "..."}],
  "study_tips": ["..."]
}
القواعد: 6–10 مصطلحات، 8–12 سؤال اختيار من متعدد موزعة على الصعوبات الثلاث،
3–4 أسئلة مقالية، إجابة صحيحة واحدة فقط لكل سؤال، وكل الأسئلة مرتبطة بالمحاضرة حصراً.
"""

LANG_RULES = {
    "auto": "اكتشف لغة المحاضرة واكتب التحليل بها نفسها، واجعل ترجمة كل مصطلح باللغة الأخرى (عربي↔إنجليزي).",
    "ar": "اكتب كل التحليل بالعربية الفصحى السهلة، وأبقِ المصطلح الإنجليزي الأصلي في حقل term.",
    "en": "Write the whole analysis in clear academic English; put the Arabic translation in the translation field.",
}


# ───────────────────────────── أدوات عامة ─────────────────────────────
def esc(x: Any) -> str:
    return html.escape("" if x is None else str(x))


def as_list(x: Any) -> list:
    return x if isinstance(x, list) else []


def get_secret(name: str, default: str = "") -> str:
    try:
        return str(st.secrets.get(name, default))
    except Exception:
        return default


def get_api_key() -> str:
    return (get_secret("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY", "")
            or st.session_state.get("manual_key", "")).strip()


def get_models() -> tuple[str, ...]:
    raw = get_secret("MODEL_CHAIN")
    models = [m.strip() for m in raw.split(",") if m.strip()] if raw else DEFAULT_MODELS
    return tuple(models)


def to_date(v: Any) -> date | None:
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return date.fromisoformat(str(v)[:10])
    except ValueError:
        return None


def _clean(v: Any) -> Any:
    if v is None or (not isinstance(v, (str, list, dict)) and pd.isna(v)):
        return None
    if isinstance(v, (date, datetime)):
        return v.isoformat()[:10]
    return v.item() if hasattr(v, "item") else v


def df_records(df: pd.DataFrame) -> list[dict]:
    return [{k: _clean(v) for k, v in r.items()} for r in df.to_dict("records")]


# ───────────────────── محرك Gemini: إعادة محاولة + نماذج احتياطية ─────────────────────
class AuthError(Exception):
    pass


class AllModelsFailed(Exception):
    pass


def error_code(exc: Exception) -> int | None:
    code = getattr(exc, "code", None) or getattr(exc, "status_code", None)
    if isinstance(code, int):
        return code
    m = re.search(r"\b([45]\d\d)\b", str(exc))
    return int(m.group(1)) if m else None


def call_gemini(api_key: str, contents: str, models: list[str], *, system: str = "",
                json_mode: bool = True, temperature: float = 0.35, validator=None,
                retries: int = 2, rounds: int = 2):
    """يجرّب كل نموذج بالترتيب؛ 503/500 → إعادة محاولة بتأخير تصاعدي، 429/404 → النموذج التالي فوراً."""
    client = genai.Client(api_key=api_key.strip())
    cfg = types.GenerateContentConfig(
        system_instruction=system or None,
        temperature=temperature,
        response_mime_type="application/json" if json_mode else None,
    )
    last: Exception | None = None
    for rnd in range(rounds):
        for model in models:
            for attempt in range(retries):
                try:
                    resp = client.models.generate_content(model=model, contents=contents, config=cfg)
                    text = (resp.text or "").strip()
                    if not text:
                        raise ValueError("empty response")
                    return (validator(text) if validator else text), model
                except Exception as exc:  # noqa: BLE001
                    last = exc
                    code = error_code(exc)
                    if code in (401, 403) or (code == 400 and "api key" in str(exc).lower()):
                        raise AuthError(str(exc)) from exc
                    if code in (404, 429, 400):
                        break  # النموذج غير متاح/مستنفد الحصة → جرّب التالي
                    if attempt < retries - 1:
                        time.sleep(min(8.0, 1.5 * 2 ** attempt) + random.random())
        if rnd < rounds - 1:
            time.sleep(4)  # فترة تهدئة قبل جولة ثانية على كل النماذج
    raise AllModelsFailed(str(last)) from last


def parse_analysis(raw: str) -> dict:
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw.strip())
    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        s, e = cleaned.find("{"), cleaned.rfind("}")
        if s == -1 or e <= s:
            raise ValueError("invalid json")
        data = json.loads(cleaned[s:e + 1])
    if not isinstance(data, dict) or not data.get("overview"):
        raise ValueError("incomplete analysis")
    return data


# الاستثناءات لا تُخزَّن مؤقتاً، لذلك لا تُحفظ النتائج الفاشلة أبداً.
# المعامل _api_key مستثنى من مفتاح التخزين المؤقت (البادئة _).
@st.cache_data(ttl=6 * 3600, max_entries=40, show_spinner=False)
def cached_extract(file_bytes: bytes) -> tuple[str, int]:
    reader = PdfReader(BytesIO(file_bytes))
    pages = []
    for n, page in enumerate(reader.pages, start=1):
        t = (page.extract_text() or "").strip()
        if t:
            pages.append(f"--- صفحة {n} ---\n{t}")
    return "\n\n".join(pages), len(reader.pages)


@st.cache_data(ttl=6 * 3600, max_entries=40, show_spinner=False)
def cached_analysis(_api_key: str, lecture_text: str, lang: str, models: tuple) -> dict:
    prompt = f"{ANALYSIS_INSTRUCTION}\n\nقاعدة اللغة: {LANG_RULES[lang]}\n\nنص المحاضرة:\n{lecture_text[:MAX_CHARS]}"
    data, model = call_gemini(_api_key, prompt, list(models), system=PROFESSOR_SYSTEM, validator=parse_analysis)
    data["_model"] = model
    return data


@st.cache_data(ttl=6 * 3600, max_entries=100, show_spinner=False)
def cached_translate(_api_key: str, text: str, target: str, models: tuple) -> str:
    prompt = (f"ترجم النص التالي إلى {target} بدقة أكاديمية مع الحفاظ على المصطلحات التقنية. "
              f"اكتشف لغة المصدر تلقائياً وأعد الترجمة فقط بلا شرح:\n\n{text}")
    out, _ = call_gemini(_api_key, prompt, list(models), json_mode=False, temperature=0.2)
    return out


@st.cache_data(ttl=6 * 3600, max_entries=100, show_spinner=False)
def cached_grade(_api_key: str, question: str, model_answer: str, student: str, models: tuple) -> str:
    prompt = (f"السؤال: {question}\nالإجابة النموذجية: {model_answer}\nإجابة الطالب: {student}\n\n"
              "قيّم إجابة الطالب من 10 بأسلوب الأستاذ الداعم: اذكر درجته، ما أصاب فيه، وما ينقصه، "
              "ثم نصيحة واحدة قصيرة للتحسين. كن موجزاً.")
    out, _ = call_gemini(_api_key, prompt, list(models), system=PROFESSOR_SYSTEM, json_mode=False)
    return out


def friendly_error(kind: str) -> str:
    return {
        "auth": "مفتاح Gemini غير صالح أو منتهي. تحقق منه في إعدادات النشر (Secrets).",
        "busy": "الخوادم مشغولة أو الحصة المجانية مستنفدة حالياً بعد تجربة كل النماذج. انتظر قليلاً ثم اضغط «إعادة محاولة ذكية».",
        "other": "حدث خطأ غير متوقع. اضغط «إعادة محاولة ذكية».",
    }[kind]


def safe_ai(fn, *args):
    """ينفّذ دالة ذكاء اصطناعي ويعيد (نتيجة، رسالة_خطأ)."""
    try:
        return fn(*args), None
    except AuthError:
        return None, friendly_error("auth")
    except AllModelsFailed:
        return None, friendly_error("busy")
    except Exception:  # noqa: BLE001
        return None, friendly_error("other")


# ───────────────────────────── التصدير (Word / PDF) ─────────────────────────────
def doc_blocks(res: dict, parts: set[str]) -> tuple:
    b: list[tuple[str, str]] = [("title", str(res.get("title") or "ملخص المحاضرة"))]
    if "summary" in parts:
        b += [("h1", "الملخص"), ("p", str(res.get("overview", "")))]
        b += [("li", str(t)) for t in as_list(res.get("key_takeaways"))]
        for s in as_list(res.get("detailed_explanation")):
            if isinstance(s, dict):
                b += [("h2", str(s.get("heading", ""))), ("p", str(s.get("explanation", "")))]
                b += [("li", str(e)) for e in as_list(s.get("examples"))]
    if "terms" in parts:
        b.append(("h1", "المصطلحات"))
        for t in as_list(res.get("terminology")):
            if isinstance(t, dict):
                b.append(("li", f"{t.get('term','')} — {t.get('translation','')}: {t.get('definition','')}"))
    if "mindmap" in parts:
        mm = res.get("mind_map") if isinstance(res.get("mind_map"), dict) else {}
        b += [("h1", "الخريطة الذهنية"), ("h2", str(mm.get("central_topic", "")))]
        for br in as_list(mm.get("branches")):
            if isinstance(br, dict):
                b.append(("h2", str(br.get("title", ""))))
                b += [("li", str(p)) for p in as_list(br.get("points"))]
    if "questions" in parts:
        b.append(("h1", "بنك الأسئلة"))
        letters = "أبجد"
        for i, q in enumerate(as_list(res.get("quiz")), 1):
            if not isinstance(q, dict):
                continue
            b.append(("p", f"{i}. [{q.get('difficulty','')}] {q.get('question','')}"))
            for j, o in enumerate(as_list(q.get("options"))):
                b.append(("li", f"{letters[j % 4]}) {o}"))
            b.append(("p", f"الإجابة: {q.get('correct_answer','')} — {q.get('explanation','')}"))
        for i, q in enumerate(as_list(res.get("essay_questions")), 1):
            if isinstance(q, dict):
                b += [("p", f"مقالي {i}. [{q.get('difficulty','')}] {q.get('question','')}"),
                      ("p", f"الإجابة النموذجية: {q.get('model_answer','')}")]
    return tuple(b)


@st.cache_data(show_spinner=False, max_entries=30)
def build_docx(blocks: tuple) -> bytes:
    from docx import Document
    from docx.oxml import OxmlElement

    doc = Document()
    for kind, text in blocks:
        if kind == "title":
            par = doc.add_heading(text, 0)
        elif kind in ("h1", "h2"):
            par = doc.add_heading(text, 1 if kind == "h1" else 2)
        elif kind == "li":
            par = doc.add_paragraph(text, style="List Bullet")
        else:
            par = doc.add_paragraph(text)
        if ARABIC_RE.search(text):  # اتجاه من اليمين لليسار
            par._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
            for run in par.runs:
                run._r.get_or_add_rPr().append(OxmlElement("w:rtl"))
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


@st.cache_data(show_spinner=False, max_entries=30)
def build_pdf(blocks: tuple) -> bytes:
    from fpdf import FPDF

    font = next((p for p in FONT_PATHS if os.path.exists(p)), None)
    if not font:
        raise RuntimeError("no_font")
    pdf = FPDF()
    pdf.set_auto_page_break(True, margin=15)
    pdf.add_page()
    pdf.add_font("AR", fname=font)
    pdf.set_text_shaping(True)  # يتطلب uharfbuzz لوصل الحروف العربية
    sizes = {"title": 20, "h1": 16, "h2": 13, "p": 11, "li": 11}
    for kind, text in blocks:
        pdf.set_font("AR", size=sizes.get(kind, 11))
        txt = f"• {text}" if kind == "li" else text
        pdf.multi_cell(0, 8, txt, align="R" if ARABIC_RE.search(text) else "L", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(1)
    return bytes(pdf.output())


# ───────────────────────────── المظهر (Dark / Light) ─────────────────────────────
STATIC_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Cairo', sans-serif; }
.stApp { background: var(--bg); }
.block-container, [data-testid="stSidebar"] { direction: rtl; text-align: right; }
.block-container { max-width: 1180px; padding-top: 2rem; }
.block-container, .block-container p, .block-container label, .block-container li,
.block-container h1, .block-container h2, .block-container h3, .block-container h4,
.block-container span, .block-container summary { color: var(--ink); }
[data-testid="stSidebar"] { background: var(--side); }
[data-testid="stSidebar"] * { color: #eef4ff !important; }
.block-container input, .block-container textarea, .block-container [data-baseweb="select"] > div {
  background: var(--card) !important; color: var(--ink) !important; }
div.stButton > button[kind="primary"] { background: linear-gradient(100deg,#376ccf,#4d80df); border: 0; border-radius: 11px; font-weight: 800; }
div.stButton > button[kind="primary"] * { color: #fff !important; }
div.stButton > button, div.stDownloadButton > button { border-radius: 11px; }
.stTabs [data-baseweb="tab-list"] { gap: .3rem; direction: rtl; overflow-x: auto; }
.hero { font-size: clamp(1.7rem, 4vw, 2.8rem); font-weight: 800; line-height: 1.25; margin: 0 0 .3rem; }
.hero span { color: #4d80df; }
.sub { color: var(--muted) !important; margin-bottom: 1rem; }
.card { background: var(--card); border: 1px solid var(--line); border-radius: 16px; padding: 1rem 1.2rem; margin-bottom: .8rem; }
.intro { background: linear-gradient(110deg, rgba(30,154,138,.15), rgba(61,112,214,.15)); border: 1px solid var(--line);
  border-radius: 16px; padding: 1.1rem 1.3rem; margin-bottom: 1rem; }
.term { border-right: 4px solid #1e9a8a; }
.fcard { min-height: 190px; display: flex; flex-direction: column; justify-content: center; align-items: center;
  text-align: center; font-size: 1.25rem; font-weight: 700; border-radius: 20px; padding: 1.5rem;
  border: 2px solid #4d80df; background: var(--card); }
.task { border-right: 6px solid var(--c, #1e9a8a); }
.muted { color: var(--muted) !important; font-size: .85rem; }
@media (max-width: 768px) { .block-container { padding: 1rem .8rem 3rem; } .fcard { min-height: 150px; font-size: 1.05rem; } }
"""


def inject_css(dark: bool) -> None:
    p = ({"bg": "#0e1525", "card": "#172036", "ink": "#e8eefc", "muted": "#9fb0cf", "line": "#26324d", "side": "#0a101d"}
         if dark else
         {"bg": "#f6f8fc", "card": "#ffffff", "ink": "#172238", "muted": "#66738b", "line": "#e5ebf3", "side": "#16243d"})
    root = ":root{" + "".join(f"--{k}:{v};" for k, v in p.items()) + "}"
    st.markdown(f"<style>{root}{STATIC_CSS}</style>", unsafe_allow_html=True)


# ───────────────────────────── عرض النتائج ─────────────────────────────
def render_summary(res: dict) -> None:
    st.markdown(f'<div class="intro"><b style="font-size:1.25rem">{esc(res.get("title","المحاضرة"))}</b>'
                f'<p style="line-height:2">{esc(res.get("overview",""))}</p></div>', unsafe_allow_html=True)
    if res.get("key_takeaways"):
        st.markdown("#### أهم ما يجب أن تتذكره")
        for t in as_list(res.get("key_takeaways")):
            st.markdown(f"- {t}")
    if res.get("study_tips"):
        st.markdown("#### نصائح الأستاذ للمذاكرة")
        for t in as_list(res.get("study_tips")):
            st.markdown(f"- {t}")


def render_explanation(res: dict) -> None:
    for i, s in enumerate(as_list(res.get("detailed_explanation")), 1):
        if not isinstance(s, dict):
            continue
        with st.expander(f"{i:02d}  {s.get('heading', '')}", expanded=i == 1):
            st.write(s.get("explanation", ""))
            for e in as_list(s.get("examples")):
                st.markdown(f"- {e}")


def render_terms(res: dict) -> None:
    for t in as_list(res.get("terminology")):
        if isinstance(t, dict):
            st.markdown(
                f'<div class="card term"><b>{esc(t.get("term"))}</b> '
                f'<span style="color:#1e9a8a">— {esc(t.get("translation"))}</span>'
                f'<div>{esc(t.get("definition"))}</div>'
                f'<div class="muted">مثال: {esc(t.get("example"))}</div></div>', unsafe_allow_html=True)


def mindmap_dot(mm: dict) -> str:
    def q(s: Any) -> str:
        return '"' + str(s).replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ") + '"'

    center = q(mm.get("central_topic", "المحاضرة"))
    lines = ["digraph G {", "rankdir=RL;", 'bgcolor="transparent";',
             'node [shape=box, style="rounded,filled", fillcolor="#eaf1ff", fontname="Arial", color="#4d80df"];',
             f'{center} [fillcolor="#16243d", fontcolor="white"];']
    for br in as_list(mm.get("branches")):
        if isinstance(br, dict):
            t = q(br.get("title", ""))
            lines.append(f"{center} -> {t};")
            lines += [f"{t} -> {q(p)};" for p in as_list(br.get("points"))]
    return "\n".join(lines + ["}"])


def render_mindmap(res: dict) -> None:
    mm = res.get("mind_map") if isinstance(res.get("mind_map"), dict) else {}
    if not as_list(mm.get("branches")):
        st.info("لا توجد فروع للخريطة الذهنية.")
        return
    st.graphviz_chart(mindmap_dot(mm), use_container_width=True)
    with st.expander("عرض كقائمة نصية"):
        for br in as_list(mm.get("branches")):
            if isinstance(br, dict):
                st.markdown(f"**{br.get('title','')}**")
                for p in as_list(br.get("points")):
                    st.markdown(f"- {p}")


# ───────────────────────────── البطاقات التعليمية ─────────────────────────────
def _fc_move(d: int, n: int) -> None:
    ss = st.session_state
    ss.fc_i = (ss.fc_i + d) % n
    ss.fc_flip = False


def _fc_toggle() -> None:
    st.session_state.fc_flip = not st.session_state.fc_flip


def _fc_mark(idx: int, n: int) -> None:
    ss = st.session_state
    ss.fc_known ^= {idx}
    _fc_move(1, n)


def _fc_shuffle(n: int) -> None:
    order = list(range(n))
    random.shuffle(order)
    st.session_state.update(fc_order=order, fc_i=0, fc_flip=False)


def render_flashcards(res: dict) -> None:
    terms = [t for t in as_list(res.get("terminology")) if isinstance(t, dict) and t.get("term")]
    if not terms:
        st.info("حلّل محاضرة أولاً لتظهر البطاقات هنا.")
        return
    n, ss = len(terms), st.session_state
    ss.setdefault("fc_i", 0)
    ss.setdefault("fc_flip", False)
    ss.setdefault("fc_known", set())
    if not ss.get("fc_order") or len(ss.fc_order) != n:
        ss.fc_order = list(range(n))
    idx = ss.fc_order[ss.fc_i % n]
    t = terms[idx]
    st.progress(len(ss.fc_known & set(range(n))) / n, text=f"أتقنت {len(ss.fc_known & set(range(n)))} من {n} • البطاقة {ss.fc_i % n + 1}/{n}")
    if ss.fc_flip:
        body = (f'{esc(t.get("translation"))}<div class="muted" style="font-weight:400;margin-top:.6rem">'
                f'{esc(t.get("definition"))}</div>')
    else:
        body = esc(t.get("term"))
    st.markdown(f'<div class="fcard">{body}</div>', unsafe_allow_html=True)
    c = st.columns(5)
    c[0].button("➡️ التالي", on_click=_fc_move, args=(1, n), key="fc_next", use_container_width=True)
    c[1].button("🔄 اقلب", on_click=_fc_toggle, key="fc_turn", use_container_width=True)
    c[2].button("⬅️ السابق", on_click=_fc_move, args=(-1, n), key="fc_prev", use_container_width=True)
    c[3].button("✅ أتقنتها" if idx not in ss.fc_known else "↩️ إلغاء", on_click=_fc_mark, args=(idx, n),
                key="fc_ok", use_container_width=True)
    c[4].button("🔀 خلط", on_click=_fc_shuffle, args=(n,), key="fc_mix", use_container_width=True)


# ───────────────────────────── بنك الأسئلة ─────────────────────────────
def resolve_correct(q: dict, opts: list[str]) -> int | None:
    ci = q.get("correct_index")
    if isinstance(ci, int) and 0 <= ci < len(opts):
        return ci
    ca = str(q.get("correct_answer", "")).strip()
    for i, o in enumerate(opts):
        if ca and (o.strip() == ca or ca in o):
            return i
    return None


def _reset_quiz() -> None:
    for k in [k for k in st.session_state if str(k).startswith(("mcq_", "essay_", "fb_"))]:
        del st.session_state[k]


def render_quiz(res: dict, api_key: str, models: tuple) -> None:
    quiz = [q for q in as_list(res.get("quiz")) if isinstance(q, dict)]
    essays = [q for q in as_list(res.get("essay_questions")) if isinstance(q, dict)]
    if not quiz and not essays:
        st.info("حلّل محاضرة أولاً لتظهر الأسئلة هنا.")
        return
    level = st.selectbox("مستوى الصعوبة", ["الكل", "سهل", "متوسط", "صعب"], key="quiz_level")
    mcq_tab, essay_tab = st.tabs(["اختيار من متعدد", "مقالي"])
    with mcq_tab:
        score_box = st.empty()
        right = answered = shown = 0
        for i, q in enumerate(quiz):
            if level != "الكل" and q.get("difficulty") != level:
                continue
            shown += 1
            opts = [str(o) for o in as_list(q.get("options"))]
            st.markdown(f"**{shown}. [{q.get('difficulty','')}] {q.get('question','')}**")
            if not opts:
                continue
            ans = st.radio("اختر", opts, index=None, key=f"mcq_{i}", label_visibility="collapsed")
            if ans is not None:
                answered += 1
                ci = resolve_correct(q, opts)
                if ci is not None and opts.index(ans) == ci:
                    right += 1
                    st.success("✅ إجابة صحيحة")
                else:
                    st.error(f"❌ الإجابة الصحيحة: {opts[ci] if ci is not None else q.get('correct_answer','—')}")
                st.caption(q.get("explanation", ""))
            st.divider()
        score_box.metric("نتيجتك", f"{right} / {answered}" if answered else "—")
        st.button("🔁 إعادة الاختبار", on_click=_reset_quiz, key="quiz_reset")
    with essay_tab:
        for i, q in enumerate(essays):
            if level != "الكل" and q.get("difficulty") != level:
                continue
            st.markdown(f"**[{q.get('difficulty','')}] {q.get('question','')}**")
            mine = st.text_area("إجابتك", key=f"essay_{i}", height=120)
            if st.button("🧑‍🏫 قيّم إجابتي", key=f"grade_{i}") and mine.strip():
                if not api_key:
                    st.warning("أضف مفتاح Gemini أولاً.")
                else:
                    with st.spinner("الأستاذ يقرأ إجابتك..."):
                        out, err = safe_ai(cached_grade, api_key, str(q.get("question")),
                                           str(q.get("model_answer")), mine, models)
                    st.session_state[f"fb_{i}"] = out or err
            if st.session_state.get(f"fb_{i}"):
                st.info(st.session_state[f"fb_{i}"])
            with st.expander("الإجابة النموذجية"):
                st.write(q.get("model_answer", ""))
            st.divider()


# ───────────────────────────── حاسبة المعدل ─────────────────────────────
def calc_gpa(df: pd.DataFrame, scale: dict[str, float]) -> tuple[float | None, float, float]:
    hours = points = 0.0
    for _, r in df.iterrows():
        g, h = r.get("التقدير"), r.get("الساعات")
        if g in scale and pd.notna(h) and h > 0:
            hours += float(h)
            points += float(h) * scale[g]
    return (points / hours if hours else None), hours, points


def render_gpa() -> None:
    ss = st.session_state
    with st.expander("⚙️ جدول التقديرات والنقاط (عدّله ليطابق نظام الجامعة)"):
        st.caption("القيم الافتراضية تقريبية؛ راجع نشرة القبول والتسجيل الرسمية في JUST وعدّل الجدول عند الحاجة.")
        sc = st.data_editor(ss.scale_base, num_rows="dynamic", hide_index=True, use_container_width=True, key="scale_editor")
    scale = {str(r["التقدير"]): float(r["النقاط"]) for _, r in sc.iterrows()
             if pd.notna(r.get("التقدير")) and pd.notna(r.get("النقاط"))}
    ss.scale_cur = scale
    st.markdown("#### مقررات الفصل")
    courses = st.data_editor(
        ss.gpa_base, num_rows="dynamic", hide_index=True, use_container_width=True, key="gpa_editor",
        column_config={
            "المقرر": st.column_config.TextColumn(),
            "الساعات": st.column_config.NumberColumn(min_value=0, max_value=6, step=1, default=3),
            "التقدير": st.column_config.SelectboxColumn(options=list(scale.keys())),
        })
    ss.gpa_cur = courses
    sem, hrs, pts = calc_gpa(courses, scale)
    c1, c2 = st.columns(2)
    c1.metric("المعدل الفصلي", f"{sem:.2f}" if sem is not None else "—")
    c2.metric("ساعات الفصل", f"{hrs:g}")
    st.markdown("#### المعدل التراكمي")
    a, b = st.columns(2)
    prev_gpa = a.number_input("معدلك التراكمي السابق", 0.0, max(scale.values(), default=4.0), 0.0, 0.01, key="prev_gpa")
    prev_h = b.number_input("الساعات المنجزة سابقاً", 0, 250, 0, 1, key="prev_h")
    total_h = prev_h + hrs
    cum = (prev_gpa * prev_h + pts) / total_h if total_h else None
    st.metric("المعدل التراكمي الجديد", f"{cum:.2f}" if cum is not None else "—")


# ───────────────────────────── المهام والمواعيد ─────────────────────────────
def task_status(due: date, done: bool) -> tuple[str, str]:
    if done:
        return "#7a8aa5", "مكتمل"
    d = (due - date.today()).days
    if d < 0:
        return "#d64545", f"متأخر {-d} يوم"
    if d <= 2:
        return "#f08a24", "اليوم" if d == 0 else f"بعد {d} يوم"
    if d <= 7:
        return "#e0b100", f"بعد {d} أيام"
    return "#1e9a8a", f"بعد {d} يوم"


def render_tasks() -> None:
    ss = st.session_state
    st.caption("أضف الواجبات والكويزات والامتحانات؛ تُرتَّب وتُلوَّن حسب القرب.")
    df = st.data_editor(
        ss.tasks_base, num_rows="dynamic", hide_index=True, use_container_width=True, key="tasks_editor",
        column_config={
            "المقرر": st.column_config.TextColumn(required=True),
            "النوع": st.column_config.SelectboxColumn(options=TASK_TYPES, default="واجب"),
            "الموعد": st.column_config.DateColumn(format="YYYY-MM-DD"),
            "مكتمل": st.column_config.CheckboxColumn(default=False),
        })
    ss.tasks_cur = df
    order = st.radio("الترتيب", ["الأقرب أولاً", "الأبعد أولاً (تنازلي)"], horizontal=True, key="task_order")
    rows = [(to_date(r["الموعد"]), r) for _, r in df.iterrows() if to_date(r["الموعد"]) and pd.notna(r["المقرر"])]
    rows.sort(key=lambda x: x[0], reverse=order.startswith("الأبعد"))
    for due, r in rows:
        color, label = task_status(due, bool(r["مكتمل"]))
        st.markdown(f'<div class="card task" style="--c:{color}"><b>{esc(r["المقرر"])}</b> — {esc(r["النوع"])}'
                    f'<div class="muted">{due.isoformat()} • <b style="color:{color}">{label}</b></div></div>',
                    unsafe_allow_html=True)


def collect_notifications() -> list[str]:
    ss, out = st.session_state, []
    for _, r in ss.get("tasks_cur", ss.tasks_base).iterrows():
        d = to_date(r.get("الموعد"))
        if d and not bool(r.get("مكتمل")) and pd.notna(r.get("المقرر")):
            color, label = task_status(d, False)
            if (d - date.today()).days <= 2:
                out.append(f"⏰ {r['المقرر']} ({r['النوع']}): {label}")
    for n in ss.notes:
        d = to_date(n.get("reminder"))
        if d and d <= date.today():
            out.append(f"📓 تذكير ملاحظة: {n['lecture']}")
    return out


# ───────────────────────────── الجدول الدراسي ─────────────────────────────
def render_schedule() -> None:
    ss = st.session_state
    st.caption("اكتب الأوقات بصيغة HH:MM مثل 08:00.")
    df = st.data_editor(
        ss.sched_base, num_rows="dynamic", hide_index=True, use_container_width=True, key="sched_editor",
        column_config={"اليوم": st.column_config.SelectboxColumn(options=DAYS, default=DAYS[0])})
    ss.sched_cur = df
    valid = df.dropna(subset=["اليوم", "المقرر"])
    for day in DAYS:
        part = valid[valid["اليوم"] == day].sort_values("من")
        if part.empty:
            continue
        items = "".join(f'<div>🕒 {esc(r["من"])} – {esc(r["إلى"])} • <b>{esc(r["المقرر"])}</b> '
                        f'<span class="muted">{esc(r["المكان"])}</span></div>' for _, r in part.iterrows())
        st.markdown(f'<div class="card"><b>{day}</b>{items}</div>', unsafe_allow_html=True)
    st.markdown("#### ⏳ العدّ التنازلي للامتحان")
    exam = st.date_input("موعد الامتحان", value=None, key="exam_date")
    if exam:
        left = (exam - date.today()).days
        if left < 0:
            st.info("انتهى الموعد، بالتوفيق في الامتحان القادم!")
        else:
            st.metric("الأيام المتبقية", left)
            st.caption("اقتراح: كتل مذاكرة 50 دقيقة تتبعها استراحة 10 دقائق، ومراجعة سريعة للمادة كل 3 أيام.")


# ───────────────────────────── الملاحظات ─────────────────────────────
def _save_note() -> None:
    ss = st.session_state
    if ss.n_lec.strip() and ss.n_txt.strip():
        ss.notes.append({"id": secrets.token_hex(3), "lecture": ss.n_lec.strip(), "text": ss.n_txt.strip(),
                         "reminder": ss.n_rem.isoformat() if ss.get("n_use_rem") and ss.n_rem else None})
        ss.n_lec = ss.n_txt = ""
        ss.n_use_rem = False


def _del_note(note_id: str) -> None:
    st.session_state.notes = [n for n in st.session_state.notes if n["id"] != note_id]


def render_notes() -> None:
    ss = st.session_state
    st.text_input("المحاضرة / المقرر", key="n_lec")
    st.text_area("ملاحظتك", key="n_txt", height=140)
    c1, c2 = st.columns([1, 2])
    c1.checkbox("أضف تذكيراً", key="n_use_rem")
    if ss.get("n_use_rem"):
        c2.date_input("تاريخ التذكير", value=date.today() + timedelta(days=1), key="n_rem")
    st.button("💾 حفظ الملاحظة", on_click=_save_note, type="primary", key="n_save")
    for n in ss.notes:
        with st.expander(f"{n['lecture']}" + (f"  🔔 {n['reminder']}" if n.get("reminder") else "")):
            st.write(n["text"])
            st.button("🗑️ حذف", on_click=_del_note, args=(n["id"],), key=f"del_{n['id']}")
    if ss.notes:
        blocks = tuple([("title", "ملاحظاتي")] + [x for n in ss.notes for x in (("h1", n["lecture"]), ("p", n["text"]))])
        a, b = st.columns(2)
        a.download_button("⬇️ Word", build_docx(blocks), "notes.docx",
                          "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="n_docx")
        txt = "\n\n".join(f"## {n['lecture']}\n{n['text']}" for n in ss.notes)
        b.download_button("⬇️ نص TXT", txt.encode("utf-8"), "notes.txt", "text/plain", key="n_txt_dl")


# ───────────────────────────── المشاركة المؤقتة ─────────────────────────────
@st.cache_resource
def share_store() -> dict:
    return {}


def create_share(res: dict) -> str:
    store = share_store()
    now = time.time()
    for k in [k for k, v in store.items() if v["exp"] < now]:
        del store[k]
    token = secrets.token_urlsafe(6)
    keep = ("title", "overview", "key_takeaways", "detailed_explanation", "terminology", "mind_map", "study_tips")
    store[token] = {"data": {k: res.get(k) for k in keep}, "exp": now + SHARE_TTL_HOURS * 3600}
    return token


def render_shared(token: str) -> None:
    item = share_store().get(token)
    if not item or item["exp"] < time.time():
        st.error("انتهت صلاحية هذا الرابط أو أنه غير صحيح.")
        return
    res = item["data"]
    st.caption("ملخص مشارَك من زميل — يعرض للقراءة فقط.")
    t1, t2, t3, t4 = st.tabs(["الملخص", "الشرح", "المصطلحات", "الخريطة"])
    with t1:
        render_summary(res)
    with t2:
        render_explanation(res)
    with t3:
        render_terms(res)
    with t4:
        render_mindmap(res)


def render_share(res: dict | None) -> None:
    if not res:
        st.info("حلّل محاضرة أولاً ثم شاركها من هنا.")
        return
    st.write(f"أنشئ رابطاً مؤقتاً (يعمل {SHARE_TTL_HOURS} ساعة) لمشاركة الملخص والشرح والمصطلحات والخريطة مع زملائك.")
    if st.button("🔗 إنشاء رابط مشاركة", key="mk_share"):
        st.session_state.share_token = create_share(res)
    tok = st.session_state.get("share_token")
    if tok:
        base = get_secret("APP_URL").rstrip("/")
        st.code(f"{base}/?share={tok}" if base else f"?share={tok}  (أضفه بعد رابط التطبيق)")
        st.caption("⚠️ كل من يملك الرابط يستطيع قراءة الملخص. الروابط تُخزَّن في ذاكرة الخادم وتزول عند إعادة تشغيله.")


# ───────────────────────────── الدعم النفسي والنصائح الإدارية ─────────────────────────────
def render_wellbeing() -> None:
    st.markdown("#### 💚 عامل نفسك بلطف في فترة الامتحانات")
    with st.expander("تهدئة التوتر بسرعة", expanded=True):
        st.markdown("- تنفّس 4-6: شهيق 4 ثوانٍ وزفير 6 ثوانٍ، لخمس دقائق.\n"
                    "- قسّم المذاكرة لكتل 50 دقيقة تتبعها استراحة 10 دقائق.\n"
                    "- نم 7 ساعات على الأقل؛ الذاكرة تتثبّت أثناء النوم.\n"
                    "- حرّك جسمك 15 دقيقة يومياً، واشرب الماء، وتجنّب الإفراط في المنبّهات.")
    with st.expander("التخطيط للامتحانات"):
        st.markdown("- ابدأ بالمواد الأصعب أو الأقرب موعداً، واترك آخر يوم للمراجعة السريعة فقط.\n"
                    "- حوّل المحاضرة إلى أسئلة وبطاقات، فالاسترجاع الفعّال يتفوّق على إعادة القراءة.\n"
                    "- لا تقارن تقدّمك بغيرك؛ قارنه بأمسك أنت.")
    with st.expander("نصائح إدارية"):
        st.markdown("- راجع المرشد الأكاديمي في كليتك قبل قرارات السحب والإضافة وتوزيع الساعات.\n"
                    "- تابع التقويم الأكاديمي الرسمي وقرارات عمادة شؤون الطلبة على موقع الجامعة.\n"
                    "- إن تعذّر حضور امتحان لعذر قاهر، تواصل مع مدرّس المادة وعمادة الكلية مبكراً ومعك الإثباتات.")
    st.info("إن شعرت بإرهاق شديد أو يأس مستمر، فتحدّث إلى شخص تثق به أو إلى مختص نفسي في الجامعة. "
            "وفي أي خطر مباشر تواصل مع جهات الطوارئ المحلية فوراً.")


# ───────────────────────────── النسخ الاحتياطي ─────────────────────────────
def export_state() -> bytes:
    ss = st.session_state
    payload = {"tasks": df_records(ss.get("tasks_cur", ss.tasks_base)),
               "schedule": df_records(ss.get("sched_cur", ss.sched_base)),
               "courses": df_records(ss.get("gpa_cur", ss.gpa_base)),
               "scale": ss.get("scale_cur", DEFAULT_GRADE_SCALE), "notes": ss.notes}
    return json.dumps(payload, ensure_ascii=False, indent=1).encode("utf-8")


def import_state(raw: bytes) -> None:
    ss, data = st.session_state, json.loads(raw.decode("utf-8"))
    ss.tasks_base = pd.DataFrame(data.get("tasks", []), columns=TASK_COLS).astype(object)
    ss.tasks_base["الموعد"] = ss.tasks_base["الموعد"].map(to_date)
    ss.tasks_base["مكتمل"] = ss.tasks_base["مكتمل"].fillna(False).astype(bool)
    ss.sched_base = pd.DataFrame(data.get("schedule", []), columns=SCHED_COLS).astype(object)
    ss.gpa_base = pd.DataFrame(data.get("courses", []), columns=GPA_COLS).astype(object)
    ss.scale_base = pd.DataFrame(list(data.get("scale", DEFAULT_GRADE_SCALE).items()), columns=["التقدير", "النقاط"])
    ss.notes = data.get("notes", [])
    for k in ("tasks_editor", "sched_editor", "gpa_editor", "scale_editor", "tasks_cur", "sched_cur", "gpa_cur"):
        ss.pop(k, None)
    ss.toasted = False


def init_state() -> None:
    ss = st.session_state
    if "tasks_base" in ss:
        return
    ss.tasks_base = pd.DataFrame({"المقرر": pd.Series(dtype="object"), "النوع": pd.Series(dtype="object"),
                                  "الموعد": pd.Series(dtype="object"), "مكتمل": pd.Series(dtype="bool")})
    ss.sched_base = pd.DataFrame({c: pd.Series(dtype="object") for c in SCHED_COLS})
    ss.gpa_base = pd.DataFrame({"المقرر": pd.Series(dtype="object"), "الساعات": pd.Series(dtype="float"),
                                "التقدير": pd.Series(dtype="object")})
    ss.scale_base = pd.DataFrame(list(DEFAULT_GRADE_SCALE.items()), columns=["التقدير", "النقاط"])
    ss.notes, ss.results, ss.job, ss.last_error, ss.toasted = [], None, None, None, False


# ───────────────────────────── تبويب التحليل ─────────────────────────────
def run_job(api_key: str, models: tuple) -> None:
    ss, job = st.session_state, st.session_state.job
    try:
        with st.spinner("الأستاذ يحلّل محاضرتك... (قد يبدّل النموذج تلقائياً عند الضغط)"):
            res = cached_analysis(api_key, job["text"], job["lang"], models)
        ss.results, ss.last_error = res, None
        _reset_quiz()
        ss.update(fc_i=0, fc_flip=False, fc_known=set(), fc_order=None)
    except AuthError as e:
        ss.last_error = ("auth", str(e))
    except AllModelsFailed as e:
        ss.last_error = ("busy", str(e))
    except Exception as e:  # noqa: BLE001
        ss.last_error = ("other", str(e))


def render_analysis_tab(api_key: str, models: tuple) -> None:
    ss = st.session_state
    up = st.file_uploader("ارفع ملف المحاضرة (PDF)", type=["pdf"])
    lang = st.selectbox("لغة المخرجات", ["auto", "ar", "en"], format_func=lambda x: {
        "auto": "تلقائي (نفس لغة الملف + ترجمة المصطلحات)", "ar": "العربية", "en": "English"}[x])
    if st.button("🚀 ابدأ التحليل", type="primary", use_container_width=True, key="go"):
        if not up:
            st.warning("ارفع ملف PDF أولاً.")
        elif not api_key:
            st.warning("لم يُضبط مفتاح Gemini. أضفه في Secrets أو من الشريط الجانبي.")
        else:
            try:
                text, pages = cached_extract(up.getvalue())
            except Exception:  # noqa: BLE001
                text, pages = "", 0
                st.error("تعذّرت قراءة الملف؛ قد يكون تالفاً أو محمياً بكلمة مرور.")
            if pages and not text.strip():
                st.error("لا يوجد نص قابل للاستخراج (ملف ممسوح ضوئياً). جرّب نسخة نصية أو OCR.")
            elif text:
                if len(text) > MAX_CHARS:
                    st.warning(f"النص طويل؛ سيُحلَّل أول {MAX_CHARS:,} حرفاً فقط.")
                ss.update(job={"text": text, "lang": lang}, pages=pages)
                run_job(api_key, models)
    err = ss.get("last_error")
    if err:
        st.error(friendly_error(err[0]))
        with st.expander("تفاصيل تقنية"):
            st.code(err[1][:800])
        if ss.job and err[0] != "auth" and st.button("🔄 إعادة محاولة ذكية", key="retry"):
            run_job(api_key, models)
            st.rerun()
    res = ss.get("results")
    if not res:
        st.caption("ارفع محاضرتك وابدأ؛ النتائج المتكررة تُخزَّن مؤقتاً فلا تستهلك الحصة.")
        return
    st.success(f"اكتمل التحليل • اللغة المكتشفة: {res.get('detected_language','—')} • النموذج: {res.get('_model','—')}")
    s1, s2, s3, s4, s5 = st.tabs(["الملخص", "الشرح الأكاديمي", "المصطلحات", "الخريطة الذهنية", "⬇️ تصدير"])
    with s1:
        render_summary(res)
    with s2:
        render_explanation(res)
    with s3:
        render_terms(res)
    with s4:
        render_mindmap(res)
    with s5:
        labels = {"summary": "الملخص والشرح", "terms": "المصطلحات", "mindmap": "الخريطة الذهنية", "questions": "بنك الأسئلة"}
        parts = st.multiselect("ما الذي تريد تصديره؟", list(labels), default=list(labels), format_func=labels.get)
        blocks = doc_blocks(res, set(parts))
        a, b = st.columns(2)
        a.download_button("⬇️ Word (.docx)", build_docx(blocks), "lecture.docx",
                          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                          use_container_width=True, key="dl_docx")
        try:
            b.download_button("⬇️ PDF", build_pdf(blocks), "lecture.pdf", "application/pdf",
                              use_container_width=True, key="dl_pdf")
        except Exception:  # noqa: BLE001
            b.warning("تصدير PDF العربي يحتاج fpdf2 وuharfbuzz وخطاً عربياً (راجع requirements/packages). استخدم Word مؤقتاً.")


def render_translator(api_key: str, models: tuple) -> None:
    text = st.text_area("النص المراد ترجمته (يُكتشف لغته تلقائياً)", height=150, key="tr_text")
    target = st.radio("الترجمة إلى", ["العربية", "English"], horizontal=True, key="tr_target")
    if st.button("🌐 ترجم", key="tr_go") and text.strip():
        if not api_key:
            st.warning("أضف مفتاح Gemini أولاً.")
            return
        with st.spinner("جارٍ الترجمة..."):
            out, err = safe_ai(cached_translate, api_key, text.strip(), "العربية" if target == "العربية" else "English", models)
        st.info(out) if out else st.error(err)


# ───────────────────────────── الدالة الرئيسية ─────────────────────────────
def main() -> None:
    inject_css(st.session_state.get("dark", False))
    token = st.query_params.get("share")
    if token:
        st.markdown('<div class="hero">🎓 ملخص <span>مشارَك</span></div>', unsafe_allow_html=True)
        render_shared(str(token))
        st.stop()
    init_state()
    ss, models = st.session_state, get_models()

    with st.sidebar:
        st.markdown("## 🎓 JUST StudyMate")
        st.toggle("🌙 الوضع الداكن", key="dark")
        if not (get_secret("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY")):
            st.text_input("مفتاح Gemini API", type="password", key="manual_key",
                          help="الأفضل ضبطه في Secrets على الخادم.")
        notif_slot = st.container()
        with st.expander("💾 نسخة احتياطية"):
            st.download_button("تنزيل بياناتي", export_state(), "studymate-backup.json", "application/json", key="bk_dl")
            up = st.file_uploader("استعادة نسخة", type=["json"], key="bk_up")
            if up and st.button("استعادة", key="bk_go"):
                try:
                    import_state(up.getvalue())
                    st.rerun()
                except Exception:  # noqa: BLE001
                    st.error("ملف النسخة غير صالح.")
            st.caption("بياناتك تُحفظ في جلسة المتصفح فقط؛ نزّل نسخة قبل الإغلاق.")

    api_key = get_api_key()
    st.markdown('<div class="hero">حوّل محاضرتك إلى <span>خطة تفوّق</span></div>', unsafe_allow_html=True)
    st.markdown('<div class="sub">مساعد دراسي لطلاب جامعة العلوم والتكنولوجيا الأردنية</div>', unsafe_allow_html=True)

    tabs = st.tabs(["📄 التحليل", "🗂️ البطاقات", "📝 الأسئلة", "🌐 المترجم", "🎯 المعدل", "⏰ المهام",
                    "📅 الجدول", "📓 الملاحظات", "🔗 المشاركة", "💚 الدعم"])
    with tabs[0]:
        render_analysis_tab(api_key, models)
    with tabs[1]:
        render_flashcards(ss.results or {})
    with tabs[2]:
        render_quiz(ss.results or {}, api_key, models)
    with tabs[3]:
        render_translator(api_key, models)
    with tabs[4]:
        render_gpa()
    with tabs[5]:
        render_tasks()
    with tabs[6]:
        render_schedule()
    with tabs[7]:
        render_notes()
    with tabs[8]:
        render_share(ss.results)
    with tabs[9]:
        render_wellbeing()

    notes = collect_notifications()
    with notif_slot:
        with st.expander(f"🔔 الإشعارات ({len(notes)})", expanded=bool(notes)):
            for n in notes or ["لا توجد تنبيهات حالياً 🎉"]:
                st.write(n)
    if notes and not ss.toasted:
        for n in notes[:3]:
            st.toast(n)
        ss.toasted = True


if __name__ == "__main__":
    main()
