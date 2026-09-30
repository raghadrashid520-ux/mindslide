"""SlideMind AI — interactive Arabic lecture analysis with Gemini and Streamlit."""

from __future__ import annotations

import json
import os
from io import BytesIO
from typing import Any

from google import genai
from google.genai import types
import streamlit as st
from pypdf import PdfReader


st.set_page_config(
    page_title="SlideMind AI",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="expanded",
)


# حقن ملف الـ Manifest لكي يقرأه موقع PWABuilder
st.markdown(
    '<link rel="manifest" href="https://raw.githubusercontent.com/raghadrashid520-ux/mindslide/main/manifest.json">',
    unsafe_allow_html=True
)

MODEL_OPTIONS = {
    "Gemini 2.0 Flash — سريع ومتوازن": "gemini-3.8-flash",
    "Gemini 1.5 Flash — اقتصادي للمحاضرات الطويلة": "gemini-3.8-flash",
}

ANALYSIS_INSTRUCTION = """
أنت أستاذ جامعي خبير ومصمم مناهج ومراجع امتحانات. حلّل محاضرة الطالب المرفقة
تحليلاً أكاديمياً عميقاً ودقيقاً، وكن واضحاً وعملياً ومفيداً للمذاكرة. اكتب
بالعربية الفصحى السهلة، مع إبقاء المصطلح الإنجليزي بين قوسين عند الحاجة.

أعد النتيجة بصيغة JSON صحيحة فقط، من دون markdown أو أي نص خارج JSON، وبالمفاتيح
التالية حرفياً:
{
  "title": "عنوان المحاضرة المستنتج",
  "overview": "ملخص تنفيذي من 3 إلى 5 جمل",
  "key_takeaways": ["أهم فكرة 1", "أهم فكرة 2"],
  "detailed_explanation": [
    {"heading": "عنوان المحور", "explanation": "شرح عميق بروح أستاذ المادة", "examples": ["مثال أو تطبيق"]}
  ],
  "terminology": [
    {"term": "Technical term", "translation": "الترجمة العربية", "definition": "تعريف دقيق", "example": "مثال مختصر"}
  ],
  "mind_map": {
    "central_topic": "الفكرة المركزية",
    "branches": [
      {"title": "محور رئيسي", "points": ["فكرة فرعية", "فكرة فرعية"]}
    ]
  },
  "quiz": [
    {
      "question": "السؤال",
      "type": "اختيار من متعدد",
      "options": ["الخيار أ", "الخيار ب", "الخيار ج", "الخيار د"],
      "correct_answer": "الإجابة الصحيحة كما تظهر في الخيارات",
      "explanation": "سبب صحة الإجابة"
    }
  ],
  "study_tips": ["نصيحة مذاكرة مرتبطة بالمحاضرة"]
}

قواعد الجودة:
- استند إلى النص فقط، ولا تخترع معلومات غير مدعومة به. إذا كان النص ناقصاً، صرّح بذلك.
- اجعل الشرح مترابطاً: اشرح السبب والنتيجة والعلاقة بين المفاهيم، لا تكتفِ بالتعريف.
- أنشئ 6 إلى 10 مصطلحات تقنية و8 إلى 12 سؤالاً متنوعاً، ووازن بين التذكر والفهم والتطبيق.
- اجعل خيارات الاختيار من متعدد معقولة، وإجابة صحيحة واحدة فقط لكل سؤال.
- لا تستخدم أسئلة أو نصوصاً عامة لا علاقة لها بالمحاضرة.
"""


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&display=swap');

        :root {
            --ink: #172238;
            --muted: #66738b;
            --navy: #16243d;
            --blue: #3d70d6;
            --teal: #1e9a8a;
            --mist: #f4f7fb;
            --line: #e5ebf3;
        }

        html, body, [class*="css"] {
            font-family: 'Cairo', sans-serif;
        }

        .stApp {
            background: #f8fafc;
            color: var(--ink);
        }

        [data-testid="stHeader"] { background: transparent; }
        [data-testid="stSidebar"] {
            background: var(--navy);
            border-left: 1px solid rgba(255,255,255,.08);
        }
        [data-testid="stSidebar"] * { color: #eef4ff; }
        [data-testid="stSidebar"] .stCaption { color: #aab9d2; }
        [data-testid="stSidebar"] input {
            background: rgba(255,255,255,.09);
            color: #fff;
            border: 1px solid rgba(255,255,255,.20);
        }
        [data-testid="stSidebar"] .stSelectbox > div > div {
            background: rgba(255,255,255,.09);
            border-color: rgba(255,255,255,.20);
        }

        .block-container {
            max-width: 1180px;
            padding-top: 2.4rem;
            padding-bottom: 4rem;
        }

        .brand-mark {
            display: flex;
            align-items: center;
            gap: .72rem;
            direction: rtl;
            margin: .2rem 0 2.2rem;
        }
        .brand-symbol {
            width: 42px;
            height: 42px;
            display: grid;
            place-items: center;
            border-radius: 13px;
            background: linear-gradient(135deg, #63d6c1, #4779da);
            color: #fff;
            font-size: 1.45rem;
            font-weight: 800;
            box-shadow: 0 10px 25px rgba(56, 124, 205, .26);
        }
        .brand-name { font-size: 1.18rem; font-weight: 800; letter-spacing: -.02em; }
        .brand-subtitle { font-size: .67rem; color: #aab9d2; margin-top: -.16rem; }

        .eyebrow {
            color: var(--blue);
            font-size: .76rem;
            font-weight: 800;
            letter-spacing: .11em;
            text-transform: uppercase;
            margin-bottom: .55rem;
        }
        .hero-title {
            font-size: clamp(2rem, 4vw, 3.35rem);
            line-height: 1.16;
            letter-spacing: -.055em;
            color: var(--navy);
            font-weight: 800;
            margin: 0;
        }
        .hero-title span { color: var(--blue); }
        .hero-copy {
            color: var(--muted);
            font-size: 1rem;
            line-height: 2;
            max-width: 700px;
            margin: .85rem 0 1.6rem;
        }
        .hero-rule {
            height: 4px;
            width: 74px;
            border-radius: 99px;
            background: linear-gradient(90deg, var(--teal), var(--blue));
            margin: 0 0 2rem;
        }

        .panel {
            background: #fff;
            border: 1px solid var(--line);
            border-radius: 18px;
            padding: 1.4rem 1.5rem;
            box-shadow: 0 11px 35px rgba(23, 34, 56, .055);
        }
        .panel-title {
            display: flex;
            align-items: center;
            gap: .55rem;
            color: var(--navy);
            font-size: 1rem;
            font-weight: 800;
            margin-bottom: .3rem;
        }
        .panel-title .number {
            display: grid;
            place-items: center;
            width: 25px;
            height: 25px;
            border-radius: 8px;
            color: #fff;
            background: var(--blue);
            font-size: .76rem;
        }
        .panel-help { color: var(--muted); font-size: .82rem; line-height: 1.8; }

        .metric-card {
            background: #fff;
            border: 1px solid var(--line);
            border-radius: 14px;
            padding: .95rem 1.1rem;
            min-height: 83px;
        }
        .metric-label { color: var(--muted); font-size: .72rem; }
        .metric-value { color: var(--navy); font-size: 1.3rem; font-weight: 800; margin-top: .2rem; }

        .result-intro {
            background: linear-gradient(110deg, #eaf8f5, #edf3ff);
            border: 1px solid #d7e9eb;
            border-radius: 17px;
            padding: 1.25rem 1.35rem;
            margin: .7rem 0 1.2rem;
        }
        .result-title { color: var(--navy); font-size: 1.3rem; font-weight: 800; }
        .result-copy { color: #43536d; line-height: 2; margin-top: .35rem; }

        .term-card {
            border: 1px solid var(--line);
            border-right: 4px solid var(--teal);
            border-radius: 13px;
            background: #fff;
            padding: .85rem 1rem;
            margin-bottom: .75rem;
        }
        .term-name { color: var(--navy); font-weight: 800; }
        .term-translation { color: var(--teal); font-size: .84rem; }
        .term-definition { color: #43536d; font-size: .87rem; line-height: 1.85; margin-top: .3rem; }

        .map-center {
            background: var(--navy);
            color: white;
            border-radius: 14px;
            padding: 1rem 1.25rem;
            text-align: center;
            font-weight: 800;
            margin-bottom: 1rem;
        }
        .map-branch {
            height: 100%;
            border-top: 3px solid var(--blue);
            border-radius: 12px;
            background: #fff;
            border-left: 1px solid var(--line);
            border-right: 1px solid var(--line);
            border-bottom: 1px solid var(--line);
            padding: .9rem 1rem;
        }
        .map-branch-title { color: var(--blue); font-weight: 800; margin-bottom: .4rem; }

        .status-note {
            border-radius: 12px;
            background: rgba(255,255,255,.07);
            color: #c5d1e7;
            padding: .75rem .85rem;
            font-size: .77rem;
            line-height: 1.85;
        }
        .footer-note { text-align: center; color: #8a96aa; font-size: .75rem; margin-top: 3rem; }

        div.stButton > button[kind="primary"] {
            background: linear-gradient(100deg, #376ccf, #4d80df);
            border: 0;
            border-radius: 11px;
            color: white;
            font-weight: 800;
            padding: .65rem 1.3rem;
        }
        div.stButton > button[kind="primary"]:hover { background: #285bbd; }
        .stTabs [data-baseweb="tab-list"] { gap: .35rem; direction: rtl; }
        .stTabs [data-baseweb="tab"] { font-family: 'Cairo', sans-serif; }
        .stExpander { border-color: var(--line); border-radius: 12px; background: #fff; }
        .stAlert { direction: rtl; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def extract_pdf_text(file_bytes: bytes) -> tuple[str, int]:
    reader = PdfReader(BytesIO(file_bytes))
    pages: list[str] = []
    for page_number, page in enumerate(reader.pages, start=1):
        page_text = (page.extract_text() or "").strip()
        if page_text:
            pages.append(f"--- الصفحة {page_number} ---\n{page_text}")
    return "\n\n".join(pages), len(reader.pages)


def parse_gemini_json(raw_text: str) -> dict[str, Any]:
    cleaned = raw_text.strip()
    if cleaned.startswith("```"):
        cleaned = cleaned.split("\n", 1)[1] if "\n" in cleaned else cleaned
        cleaned = cleaned.rsplit("```", 1)[0].strip()
    try:
        parsed = json.loads(cleaned)
        return parsed if isinstance(parsed, dict) else {"detailed_explanation": [{"heading": "الناتج", "explanation": cleaned}]}
    except json.JSONDecodeError:
        start, end = cleaned.find("{"), cleaned.rfind("}")
        if start != -1 and end > start:
            try:
                parsed = json.loads(cleaned[start : end + 1])
                return parsed if isinstance(parsed, dict) else {"raw_response": cleaned}
            except json.JSONDecodeError:
                pass
    return {
        "title": "نتيجة التحليل",
        "overview": "تعذر تنظيم الإجابة تلقائياً بصيغة الحقول. يمكنك قراءة النص الكامل أدناه.",
        "detailed_explanation": [{"heading": "إجابة Gemini", "explanation": cleaned, "examples": []}],
        "terminology": [],
        "mind_map": {"central_topic": "المحاضرة", "branches": []},
        "quiz": [],
        "study_tips": [],
        "raw_response": cleaned,
    }


def run_analysis(api_key: str, model_name: str, lecture_text: str) -> dict[str, Any]:
    client = genai.Client(api_key=api_key.strip())
    prompt = (
        f"{ANALYSIS_INSTRUCTION}\n\n"
        "ابدأ التحليل الآن. هذا هو النص المستخرج من المحاضرة:\n\n"
        f"{lecture_text[:100_000]}"
    )
    response = client.models.generate_content(
        model=model_name,
        contents=prompt,
        config=types.GenerateContentConfig(
            temperature=0.35,
            response_mime_type="application/json",
        ),
    )
    return parse_gemini_json(response.text or "")


def as_list(value: Any) -> list[Any]:
    if isinstance(value, list):
        return value
    return []


def render_results(results: dict[str, Any], page_count: int, char_count: int) -> None:
    title = results.get("title") or "تحليل المحاضرة"
    overview = results.get("overview") or "لم يتم توفير ملخص موجز."
    takeaways = as_list(results.get("key_takeaways"))

    st.markdown(
        f"""
        <div class="result-intro" dir="rtl">
            <div class="result-title">{title}</div>
            <div class="result-copy">{overview}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    metric_cols = st.columns(4)
    metrics = [
        ("صفحات محللة", str(page_count)),
        ("حروف مستخرجة", f"{char_count:,}"),
        ("محاور الشرح", str(len(as_list(results.get("detailed_explanation"))))),
        ("أسئلة امتحانية", str(len(as_list(results.get("quiz"))))),
    ]
    for col, (label, value) in zip(metric_cols, metrics):
        with col:
            st.markdown(
                f'<div class="metric-card" dir="rtl"><div class="metric-label">{label}</div>'
                f'<div class="metric-value">{value}</div></div>',
                unsafe_allow_html=True,
            )

    tab_summary, tab_explanation, tab_terms, tab_map, tab_quiz, tab_text = st.tabs(
        [
            "الملخص",
            "الشرح الأكاديمي",
            "الترجمة والمصطلحات",
            "الخريطة الذهنية",
            "أسئلة الامتحان",
            "النص المستخرج",
        ]
    )

    with tab_summary:
        st.markdown("### أهم ما يجب أن تتذكره")
        if takeaways:
            for item in takeaways:
                st.markdown(f"- {item}")
        else:
            st.info("لم تُستخرج نقاط مختصرة. انتقل إلى تبويب شرح الأستاذ.")
        tips = as_list(results.get("study_tips"))
        if tips:
            st.markdown("### نصائح مذاكرة مخصصة")
            for tip in tips:
                st.markdown(f"- {tip}")

    with tab_explanation:
        sections = as_list(results.get("detailed_explanation"))
        if not sections:
            st.info("لا يوجد شرح منظم في النتيجة.")
        for index, section in enumerate(sections, start=1):
            if not isinstance(section, dict):
                st.markdown(str(section))
                continue
            heading = section.get("heading") or f"المحور {index}"
            with st.expander(f"{index:02d}  {heading}", expanded=index == 1):
                st.write(section.get("explanation") or "")
                examples = as_list(section.get("examples"))
                if examples:
                    st.markdown("**تطبيقات وأمثلة**")
                    for example in examples:
                        st.markdown(f"- {example}")

    with tab_terms:
        terms = as_list(results.get("terminology"))
        if not terms:
            st.info("لم يتم العثور على مصطلحات منظمة.")
        for term in terms:
            if not isinstance(term, dict):
                st.markdown(str(term))
                continue
            st.markdown(
                f"""
                <div class="term-card" dir="rtl">
                    <div class="term-name">{term.get("term", "مصطلح")}</div>
                    <div class="term-translation">{term.get("translation", "")}</div>
                    <div class="term-definition">{term.get("definition", "")}</div>
                    <div class="term-definition"><b>مثال:</b> {term.get("example", "—")}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    with tab_map:
        mind_map = results.get("mind_map") if isinstance(results.get("mind_map"), dict) else {}
        st.markdown(
            f'<div class="map-center" dir="rtl">الفكرة المركزية: {mind_map.get("central_topic", title)}</div>',
            unsafe_allow_html=True,
        )
        branches = as_list(mind_map.get("branches"))
        if branches:
            branch_cols = st.columns(min(3, len(branches)))
            for index, branch in enumerate(branches):
                if not isinstance(branch, dict):
                    continue
                with branch_cols[index % len(branch_cols)]:
                    points = as_list(branch.get("points"))
                    point_html = "".join(f"<li>{point}</li>" for point in points)
                    st.markdown(
                        f'<div class="map-branch" dir="rtl"><div class="map-branch-title">'
                        f'{branch.get("title", "محور")}</div><ul>{point_html}</ul></div>',
                        unsafe_allow_html=True,
                    )
        else:
            st.info("لم يتم بناء فروع للخريطة الذهنية.")

    with tab_quiz:
        questions = as_list(results.get("quiz"))
        if not questions:
            st.info("لم يتم توليد أسئلة.")
        for index, question in enumerate(questions, start=1):
            if not isinstance(question, dict):
                continue
            prompt = question.get("question", f"السؤال {index}")
            options = [str(option) for option in as_list(question.get("options"))]
            st.markdown(f"**{index}. {prompt}**")
            if options:
                st.radio(
                    "اختر إجابتك",
                    options,
                    index=None,
                    key=f"quiz_answer_{index}",
                    label_visibility="collapsed",
                )
            with st.expander("إظهار الإجابة والشرح"):
                st.markdown(f"**الإجابة الصحيحة:** {question.get('correct_answer', '—')}")
                st.write(question.get("explanation") or "—")
            if index != len(questions):
                st.divider()

    with tab_text:
        st.info("النص المستخرج معروض للمراجعة فقط، وهو النص الذي أرسل إلى Gemini.")
        st.text_area("النص", st.session_state.get("lecture_text", ""), height=500, label_visibility="collapsed")

    download_payload = json.dumps(results, ensure_ascii=False, indent=2)
    st.download_button(
        "تنزيل التحليل بصيغة JSON",
        data=download_payload,
        file_name="slidemind-analysis.json",
        mime="application/json",
        use_container_width=False,
    )


def main() -> None:
    inject_styles()

    with st.sidebar:
        st.markdown(
            """
            <div class="brand-mark" dir="rtl">
                <div class="brand-symbol">◈</div>
                <div>
                    <div class="brand-name">SlideMind AI</div>
                    <div class="brand-subtitle">STUDY SMARTER</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("### إعداد التحليل")
        api_key = st.text_input(
            "مفتاح Gemini API",
            value=os.getenv("GEMINI_API_KEY", ""),
            type="password",
            placeholder="AIza...",
            help="يُستخدم داخل جلسة التحليل فقط ولا يتم حفظه في التطبيق.",
        )
        selected_model_label = st.selectbox("النموذج", list(MODEL_OPTIONS.keys()), index=0)
        st.markdown(
            """
            <div class="status-note" dir="rtl">
                <b>خصوصيتك أولاً</b><br>
                يتم استخراج النص في الذاكرة المؤقتة للتطبيق وإرساله إلى Gemini عند طلب التحليل فقط.
                لا يتم حفظ ملفات المحاضرات على القرص.
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("---")
        st.caption("SlideMind AI · مساعد مذاكرة أكاديمي")

    st.markdown('<div class="eyebrow" dir="rtl">AI-POWERED LECTURE ANALYSIS</div>', unsafe_allow_html=True)
    st.markdown(
        '<h1 class="hero-title" dir="rtl">حوّل محاضرتك إلى <span>خطة تفوّق</span></h1>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="hero-copy" dir="rtl">ارفع ملف PDF، ودع SlideMind AI يشرح المفاهيم بروح أستاذ المادة، '
        'يفكك المصطلحات، ويرسم لك خريطة واضحة مع أسئلة امتحانية ذكية.</p>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="hero-rule"></div>', unsafe_allow_html=True)

    upload_col, guide_col = st.columns([1.55, 1], gap="large")
    with upload_col:
        st.markdown(
            '<div class="panel-title" dir="rtl"><span class="number">1</span> ارفع ملف المحاضرة</div>'
            '<div class="panel-help" dir="rtl">ملف PDF واحد، ويفضل أن يكون النص قابلاً للتحديد.</div>',
            unsafe_allow_html=True,
        )
        uploaded_file = st.file_uploader("ملف PDF", type=["pdf"], label_visibility="collapsed")
        if uploaded_file:
            st.success(f"تم تجهيز الملف: {uploaded_file.name}")
    with guide_col:
        st.markdown(
            """
            <div class="panel" dir="rtl">
                <div class="panel-title"><span class="number">2</span> ماذا ستحصل عليه؟</div>
                <div class="panel-help">
                    شرح عميق ومترابط<br>
                    ترجمة المصطلحات التقنية<br>
                    خريطة ذهنية منظمة<br>
                    بنك أسئلة متدرج الصعوبة
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("ابدأ تحليل المحاضرة", type="primary", use_container_width=True):
        if not uploaded_file:
            st.warning("يرجى رفع ملف PDF أولاً.")
        elif not api_key.strip():
            st.warning("أدخل مفتاح Gemini API من الشريط الجانبي للمتابعة.")
        else:
            try:
                with st.spinner("جاري استخراج النص وبناء تحليل أكاديمي متكامل..."):
                    lecture_text, page_count = extract_pdf_text(uploaded_file.getvalue())
                    if not lecture_text.strip():
                        st.error("لم يتم العثور على نص قابل للاستخراج. جرّب ملف PDF نصياً أو استخدم OCR أولاً.")
                        return
                    results = run_analysis(api_key, MODEL_OPTIONS[selected_model_label], lecture_text)
                    st.session_state["results"] = results
                    st.session_state["lecture_text"] = lecture_text
                    st.session_state["page_count"] = page_count
                    st.session_state["char_count"] = len(lecture_text)
                st.success("اكتمل التحليل. تصفّح النتائج من التبويبات أدناه.")
            except Exception as error:
                st.error(
                    "تعذر إكمال التحليل. تحقق من صلاحية مفتاح Gemini واتصال الإنترنت، "
                    "ثم حاول مرة أخرى."
                )
                with st.expander("تفاصيل تقنية للمساعدة"):
                    st.code(str(error))

    if st.session_state.get("results"):
        st.markdown("---")
        render_results(
            st.session_state["results"],
            st.session_state.get("page_count", 0),
            st.session_state.get("char_count", 0),
        )
    else:
        st.markdown(
            '<div class="footer-note" dir="rtl">ابدأ برفع محاضرتك. كلما كان النص أوضح، كان التحليل أدق.</div>',
            unsafe_allow_html=True,
        )


if __name__ == "__main__":
    main()
