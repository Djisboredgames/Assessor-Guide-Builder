"""
CDU Assessor Guide Builder
Run with: streamlit run app.py
"""

import streamlit as st
import os
import json
import tempfile
from datetime import datetime
from pathlib import Path
from tga_fetcher import fetch_unit_data, UnitData, KnowledgeItem, Element, PerformanceCriteria, FoundationSkill
from docx_generator import build_assessor_guide
from ai_generator import (
    check_ollama_available, get_ollama_models, check_claude_api, generate_questions
)

# ── Settings persistence ──────────────────────────────────────────
SETTINGS_FILE = Path(__file__).parent / ".settings.json"

def load_settings() -> dict:
    defaults = {
        "ai_provider": "none",
        "claude_api_key": "",
        "ollama_model": "llama3.1:8b",
        "dark_mode": False,
        "cdu_locations": "CDU Casuarina, Alice Springs or Haileybury Rendall School",
        "default_ai_statement": (
            "Students may use AI tools to assist with research and understanding concepts, "
            "but all assessment responses must be in the student's own words. "
            "AI-generated content submitted as the student's own work will be treated as academic misconduct."
        ),
    }
    try:
        if SETTINGS_FILE.exists():
            with open(SETTINGS_FILE) as f:
                saved = json.load(f)
                defaults.update(saved)
    except Exception:
        pass
    return defaults

def save_settings(settings: dict):
    try:
        with open(SETTINGS_FILE, "w") as f:
            json.dump(settings, f, indent=2)
    except Exception:
        pass


# ── Page config ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Assessor Guide Builder",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

def _dark_overrides() -> str:
    """CSS overrides for every Streamlit native widget in dark mode."""
    return """
/* ═══ DARK MODE COMPONENT OVERRIDES ═══ */

/* ── Markdown & general text ── */
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stMarkdownContainer"] span,
[data-testid="stMarkdownContainer"] strong,
[data-testid="stMarkdownContainer"] em,
[data-testid="stMarkdownContainer"] h1,
[data-testid="stMarkdownContainer"] h2,
[data-testid="stMarkdownContainer"] h3,
[data-testid="stText"] p,
.stMarkdown p { color: #E2E8F0 !important; }

/* ── Captions ── */
[data-testid="stCaptionContainer"] p,
small { color: #8B949E !important; }

/* ── Form labels ── */
label,
[data-testid="stWidgetLabel"] p,
[data-testid="stWidgetLabel"] span { color: #CBD5E0 !important; }

/* ── Text inputs & textareas ── */
.stTextInput input,
.stTextArea textarea,
[data-baseweb="input"] input,
[data-baseweb="textarea"] textarea,
input[type="text"], input[type="password"] {
    background-color: #1E2432 !important;
    color: #E2E8F0 !important;
    border-color: #3D4A5C !important;
}

/* ── Number input ── */
[data-testid="stNumberInput"] input {
    background-color: #1E2432 !important;
    color: #E2E8F0 !important;
    border-color: #3D4A5C !important;
}
[data-testid="stNumberInput"] button {
    background-color: #2D3748 !important;
    color: #E2E8F0 !important;
    border-color: #3D4A5C !important;
}

/* ── Selectbox ── */
[data-baseweb="select"] > div:first-child {
    background-color: #1E2432 !important;
    border-color: #3D4A5C !important;
}
[data-baseweb="select"] span,
[data-baseweb="select"] div { color: #E2E8F0 !important; }
[data-baseweb="select"] svg { fill: #E2E8F0 !important; }

/* ── Dropdown / popover ── */
[data-baseweb="popover"],
[data-baseweb="menu"] { background-color: #1E2432 !important; }
[data-baseweb="popover"] *,
[data-baseweb="menu"] * { color: #E2E8F0 !important; }
[role="option"]:hover,
[aria-selected="true"] { background-color: #2D3748 !important; }

/* ── Multiselect ── */
[data-testid="stMultiSelect"] [data-baseweb="tag"] {
    background-color: #2C5282 !important;
}
[data-testid="stMultiSelect"] [data-baseweb="tag"] span { color: #E2E8F0 !important; }
[data-testid="stMultiSelect"] input {
    background-color: #1E2432 !important;
    color: #E2E8F0 !important;
}

/* ── Radio buttons ── */
[data-testid="stRadio"] label span { color: #E2E8F0 !important; }

/* ── Checkboxes ── */
[data-testid="stCheckbox"] label span { color: #E2E8F0 !important; }

/* ── Toggle ── */
[data-testid="stToggle"] label span { color: #E2E8F0 !important; }
[data-testid="stToggleLabel"] { color: #E2E8F0 !important; }

/* ── Expanders ── */
[data-testid="stExpander"] {
    background-color: #1A1D27 !important;
    border: 1px solid #2D3748 !important;
}
[data-testid="stExpander"] > div { background-color: #1A1D27 !important; }

/* ── Alert / info / warning / error boxes ── */
[data-testid="stAlert"] {
    background-color: rgba(255,255,255,0.06) !important;
    border-color: rgba(255,255,255,0.12) !important;
}
[data-testid="stAlert"] p,
[data-testid="stAlert"] span { color: #E2E8F0 !important; }

/* ── Info/success message containers ── */
[data-testid="stSuccessMessage"] p,
[data-testid="stInfoMessage"] p,
[data-testid="stWarningMessage"] p,
[data-testid="stErrorMessage"] p { color: #E2E8F0 !important; }

/* ── Horizontal rule ── */
hr { border-color: #2D3748 !important; }

/* ── Code blocks ── */
code { background-color: #1E2432 !important; color: #E2E8F0 !important; }
pre  { background-color: #1A1D27 !important; }
pre code { background-color: transparent !important; }

/* ── Spinner text ── */
[data-testid="stSpinner"] p { color: #E2E8F0 !important; }
"""


def get_css(dark: bool) -> str:
    if dark:
        bg          = "#0E1117"
        bg2         = "#1A1D27"
        text        = "#FAFAFA"
        border      = "#2D3748"
        hdr_grad    = "linear-gradient(135deg, #0D1B2A 0%, #162444 100%)"
        sec_color   = "#90CDF4"
        sec_border  = "#2D3748"
        sec_num     = "#2C5282"
        sb_bg       = "#12161F"
        sb_border   = "#2D3748"
        exp_color   = "#90CDF4"
        ps_done     = "background:#1C3A27;border-color:#38A169;color:#68D391"
        ps_now      = "background:#1A2744;border-color:#4A90D9;color:#90CDF4;font-weight:700"
        ps_later    = "background:#1A1D27;border-color:#2D3748;color:#4A5568"
    else:
        bg          = "#FFFFFF"
        bg2         = "#F0F4F8"
        text        = "#1A1A1A"
        border      = "#E2E8F0"
        hdr_grad    = "linear-gradient(135deg, #1A3A5C 0%, #2C5282 100%)"
        sec_color   = "#1A3A5C"
        sec_border  = "#E2ECF5"
        sec_num     = "#1A3A5C"
        sb_bg       = "#F8FAFC"
        sb_border   = "#E2E8F0"
        exp_color   = "#1A3A5C"
        ps_done     = "background:#EBF5EB;border-color:#38A169;color:#276749"
        ps_now      = "background:#EBF0F8;border-color:#1A3A5C;color:#1A3A5C;font-weight:700"
        ps_later    = "background:#F7FAFC;border-color:#CBD5E0;color:#A0AEC0"

    return f"""<style>
/* ── Hide Streamlit chrome ── */
#MainMenu, footer {{ visibility: hidden; }}
header[data-testid="stHeader"] {{ display: none !important; }}
.stDeployButton {{ display: none !important; }}

/* ── Typography ── */
html, body, [class*="css"] {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    color: {text};
}}

/* ── Page background & layout ── */
.stApp, [data-testid="stAppViewContainer"] {{ background-color: {bg} !important; }}
.block-container {{
    padding-top: 1rem !important;
    padding-bottom: 2rem !important;
    max-width: 1080px !important;
}}

/* ── App header banner ── */
.app-header {{
    background: {hdr_grad};
    border-radius: 10px;
    padding: 1.2rem 1.75rem;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
}}
.app-header-left {{ flex: 1; }}
.app-header-title {{ font-size:1.4rem; font-weight:700; color:white; margin:0; letter-spacing:-0.3px; }}
.app-header-sub   {{ font-size:0.82rem; color:rgba(255,255,255,0.68); margin:0.2rem 0 0 0; }}
.app-header-badge {{
    background: rgba(255,255,255,0.15);
    border: 1px solid rgba(255,255,255,0.3);
    color: rgba(255,255,255,0.9);
    border-radius: 6px;
    padding: 0.3rem 0.75rem;
    font-size: 0.78rem;
    font-weight: 600;
    white-space: nowrap;
}}

/* ── Step progress bar ── */
.progress-wrap {{ display:flex; gap:6px; margin-bottom:1.5rem; }}
.progress-step {{
    flex: 1; padding: 0.55rem 0.5rem; border-radius: 8px;
    font-size: 0.8rem; font-weight: 500; text-align: center; border: 1.5px solid;
}}
.ps-done   {{ {ps_done}; }}
.ps-now    {{ {ps_now}; }}
.ps-later  {{ {ps_later}; }}

/* ── Section title ── */
.section-title {{
    display: flex; align-items: center; gap: 9px;
    font-size: 1.05rem; font-weight: 600; color: {sec_color};
    margin-bottom: 1rem; padding-bottom: 0.55rem;
    border-bottom: 2px solid {sec_border};
}}
.section-num {{
    display: inline-flex; align-items: center; justify-content: center;
    width: 24px; height: 24px; background: {sec_num}; color: white;
    border-radius: 50%; font-size: 0.75rem; font-weight: 700; flex-shrink: 0;
}}

/* ── Sidebar ── */
div[data-testid="stSidebar"],
[data-testid="stSidebarContent"] {{ background: {sb_bg} !important; border-right: 1px solid {sb_border} !important; }}
.sidebar-logo {{
    background: #1A3A5C; color: white; padding: 0.75rem 1rem; border-radius: 8px;
    margin-bottom: 1rem; font-weight: 700; font-size: 0.8rem;
    text-align: center; letter-spacing: 1px; text-transform: uppercase;
}}
section[data-testid="stSidebar"] label {{ font-size: 0.82rem !important; }}

/* ── Primary buttons ── */
div[data-testid="stButton"] button[kind="primary"] {{
    background: #1A3A5C !important; border: none !important;
    border-radius: 7px !important; font-weight: 600 !important;
}}
div[data-testid="stButton"] button[kind="primary"]:hover {{ background: #2C5282 !important; }}

/* ── Download button ── */
div[data-testid="stDownloadButton"] button {{
    background: #276749 !important; color: white !important;
    border: none !important; border-radius: 7px !important; font-weight: 600 !important;
}}
div[data-testid="stDownloadButton"] button:hover {{ background: #2F855A !important; }}

/* ── Expander summary ── */
details summary p {{ font-weight: 600 !important; color: {exp_color} !important; }}

{_dark_overrides() if dark else ""}
</style>"""


def step_header(n, title):
    st.markdown(
        f'<div class="section-title">'
        f'<span class="section-num">{n}</span>'
        f'<span>{title}</span>'
        f'</div>',
        unsafe_allow_html=True
    )


def init():
    if "step" not in st.session_state:
        st.session_state.step = 1
    if "unit_data" not in st.session_state:
        st.session_state.unit_data = None
    if "unit_code" not in st.session_state:
        st.session_state.unit_code = ""
    if "fetch_complete" not in st.session_state:
        st.session_state.fetch_complete = False
    if "tasks" not in st.session_state:
        st.session_state.tasks = []
    if "generated_path" not in st.session_state:
        st.session_state.generated_path = None
    if "fetch_id" not in st.session_state:
        st.session_state.fetch_id = 0
    if "ai_questions" not in st.session_state:
        st.session_state.ai_questions = None
    if "settings" not in st.session_state:
        st.session_state.settings = load_settings()


def k(name):
    return f"{name}_{st.session_state.fetch_id}"


# ── Sidebar: Settings ─────────────────────────────────────────────
def render_sidebar():
    s = st.session_state.settings

    with st.sidebar:
        st.markdown('<div class="sidebar-logo">Assessor Guide Builder</div>', unsafe_allow_html=True)

        # Dark mode toggle — auto-saves immediately
        dark = st.toggle("Dark mode", value=s.get("dark_mode", False), key="sb_dark")
        if dark != s.get("dark_mode", False):
            s["dark_mode"] = dark
            save_settings(s)
            st.rerun()

        st.markdown("---")
        st.markdown("**AI Question Generation**")
        provider = st.radio(
            "Provider",
            options=["none", "ollama", "claude"],
            format_func=lambda x: {
                "none": "Off — basic templates",
                "ollama": "Ollama (free, local)",
                "claude": "Claude API (paid, higher quality)"
            }[x],
            index=["none", "ollama", "claude"].index(s.get("ai_provider", "none")),
            key="sb_provider"
        )
        s["ai_provider"] = provider

        if provider == "ollama":
            st.markdown("---")
            if check_ollama_available():
                st.success("Ollama connected")
                models = get_ollama_models()
                if models:
                    model = st.selectbox("Model", models, key="sb_model")
                    s["ollama_model"] = model
                else:
                    st.warning("No models installed.\nRun: `ollama pull llama3.2:3b`")
            else:
                st.error("Ollama not running")
                st.caption("Open the Ollama app or run `ollama serve`")

        elif provider == "claude":
            st.markdown("---")
            api_key = st.text_input(
                "API Key",
                value=s.get("claude_api_key", ""),
                type="password",
                key="sb_apikey",
                placeholder="sk-ant-..."
            )
            s["claude_api_key"] = api_key

            col1, col2 = st.columns(2)
            with col1:
                if api_key and st.button("Test", key="sb_test", use_container_width=True):
                    with st.spinner("..."):
                        if check_claude_api(api_key):
                            st.success("Valid")
                        else:
                            st.error("Invalid")
            with col2:
                if st.button("Save key", key="sb_save_key", use_container_width=True):
                    save_settings(s)
                    st.success("Saved")

            if not api_key:
                st.caption("[Get a key](https://console.anthropic.com) — ~$0.15 per guide")

        st.markdown("---")
        st.markdown("**CDU Defaults**")

        s["cdu_locations"] = st.text_input(
            "Workshop locations",
            value=s.get("cdu_locations", ""),
            key="sb_locations",
            help="Used in Assessment Conditions responses"
        )

        s["default_ai_statement"] = st.text_area(
            "Default AI statement",
            value=s.get("default_ai_statement", ""),
            height=80,
            key="sb_ai_stmt"
        )

        if st.button("Save settings", use_container_width=True, key="sb_save_all"):
            save_settings(s)
            st.success("Settings saved")

        st.markdown("---")
        st.caption("CDU Assessor Guide Builder v1.0")


# ══════════════════════════════════════════════════════════════════
def main():
    init()
    st.markdown(get_css(st.session_state.settings.get("dark_mode", False)), unsafe_allow_html=True)
    render_sidebar()

    st.markdown(
        '<div class="app-header">'
        '<div class="app-header-left">'
        '<p class="app-header-title">Assessor Guide Builder</p>'
        '<p class="app-header-sub">Generate VET Assessor Guides from training.gov.au unit data — Charles Darwin University</p>'
        '</div>'
        '<div class="app-header-badge">v1.0</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    steps = ["Fetch Unit", "Review Data", "Configure Tasks", "Generate"]
    step_html = '<div class="progress-wrap">'
    for i, label in enumerate(steps, 1):
        if i < st.session_state.step:
            cls = "ps-done"
            icon = "✓ "
        elif i == st.session_state.step:
            cls = "ps-now"
            icon = f"{i}. "
        else:
            cls = "ps-later"
            icon = f"{i}. "
        step_html += f'<div class="progress-step {cls}">{icon}{label}</div>'
    step_html += "</div>"
    st.markdown(step_html, unsafe_allow_html=True)
    {1: step1, 2: step2, 3: step3, 4: step4}[st.session_state.step]()


# ── STEP 1 ────────────────────────────────────────────────────────
def step1():
    step_header(1, "Enter Unit Code")

    col1, col2 = st.columns([2, 1])
    with col1:
        code = st.text_input(
            "Unit Code",
            value=st.session_state.unit_code,
            placeholder="e.g. AURHTF102"
        ).strip().upper()
        st.session_state.unit_code = code
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        fetch = st.button("Fetch from training.gov.au", type="primary", use_container_width=True)

    if fetch and code:
        with st.spinner("Fetching unit data..."):
            progress = st.empty()
            ud = fetch_unit_data(code, progress_callback=lambda m: progress.text(m))
            progress.empty()

        st.session_state.unit_data = ud
        st.session_state.fetch_id += 1
        st.session_state.ai_questions = None

        if ud.is_populated():
            st.success(f"{ud.code} — {ud.title}")
            items = []
            if ud.application: items.append("Application")
            if ud.elements: items.append(f"{len(ud.elements)} Elements")
            if ud.knowledge_evidence: items.append(f"{len(ud.knowledge_evidence)} KE items")
            if ud.performance_evidence: items.append(f"{len(ud.performance_evidence)} PE items")
            if ud.foundation_skills: items.append(f"{len(ud.foundation_skills)} Foundation Skills")
            if ud.assessment_conditions: items.append("Assessment Conditions")
            if items:
                st.info(f"Data loaded: {', '.join(items)}")
        else:
            st.warning("Could not auto-fetch all data. You can enter details manually in the next step.")
            st.session_state.unit_data = UnitData(code=code)

        for err in ud.fetch_errors:
            st.warning(err)

        st.session_state.fetch_complete = True

    if st.session_state.fetch_complete:
        if st.button("Next: Review Data", type="primary", use_container_width=True):
            st.session_state.step = 2
            st.rerun()


# ── STEP 2 ────────────────────────────────────────────────────────
def step2():
    step_header(2, "Review and Edit Unit Data")

    ud = st.session_state.unit_data
    if not ud:
        st.error("No data loaded.")
        if st.button("Back"): st.session_state.step = 1; st.rerun()
        return

    if ud.title:
        st.info(f"{ud.code} — {ud.title}")

    with st.expander("Unit Details", expanded=True):
        ud.code = st.text_input("Unit Code", value=ud.code, key=k("code"))
        ud.title = st.text_input("Unit Title", value=ud.title, key=k("title"))
        ud.application = st.text_area(
            "Application / Descriptor",
            value=ud.application or "",
            height=120, key=k("app")
        )
        c1, c2 = st.columns(2)
        with c1:
            ud.has_prerequisites = st.checkbox("Has prerequisites?", value=ud.has_prerequisites, key=k("prereq"))
        with c2:
            ud.has_licensing = st.checkbox("Has licensing requirements?", value=ud.has_licensing, key=k("lic"))

    with st.expander("Assessment Conditions"):
        ud.has_additional_assessor_reqs = st.checkbox(
            "Additional assessor requirements?",
            value=ud.has_additional_assessor_reqs, key=k("areq")
        )
        ud.has_simulated_environment = st.checkbox(
            "Simulated environment included?",
            value=ud.has_simulated_environment, key=k("sim")
        )
        ud.has_required_resources = st.checkbox(
            "CDU has all required resources?",
            value=ud.has_required_resources, key=k("res")
        )
        ud.assessment_conditions = st.text_area(
            "Assessment Conditions",
            value=ud.assessment_conditions or "",
            height=150, key=k("ac")
        )

    with st.expander("Performance Evidence"):
        ud.has_volume_frequency = st.checkbox(
            "Volume and frequency requirement?",
            value=ud.has_volume_frequency, key=k("vf")
        )
        ud.has_work_placement = st.checkbox(
            "Work placement hours required?",
            value=ud.has_work_placement, key=k("wp")
        )
        pe_val = ud.performance_evidence_raw or "\n".join(ud.performance_evidence)
        pe_text = st.text_area("Performance Evidence (one per line)", value=pe_val, height=200, key=k("pe"))
        ud.performance_evidence = [l.strip() for l in pe_text.split("\n") if l.strip()]
        ud.performance_evidence_raw = pe_text
        if ud.has_volume_frequency:
            ud.volume_frequency_detail = st.text_area(
                "Volume and frequency details",
                value=ud.volume_frequency_detail or "",
                height=100, key=k("vfd")
            )

    with st.expander("Knowledge Evidence"):
        ke_val = ud.knowledge_evidence_raw or "\n".join(
            [x.text if hasattr(x, 'text') else str(x) for x in ud.knowledge_evidence]
        )
        ke_text = st.text_area("Knowledge Evidence (one per line)", value=ke_val, height=300, key=k("ke"))
        ud.knowledge_evidence = [KnowledgeItem(text=l.strip()) for l in ke_text.split("\n") if l.strip()]
        ud.knowledge_evidence_raw = ke_text

    with st.expander("Elements and Performance Criteria"):
        if ud.elements:
            st.info(f"{len(ud.elements)} elements loaded")
            for el in ud.elements:
                st.markdown(f"**Element {el.number}: {el.title}**")
                for pc in el.performance_criteria:
                    st.caption(f"    {pc.number} {pc.text}")
        else:
            st.warning("No elements found. Paste below if needed.")
        el_text = st.text_area(
            "Override elements (format: '1. Title' then '1.1 PC text')",
            height=150, key=k("el")
        )
        if el_text.strip():
            ud.elements = _parse_elements_text(el_text)

    with st.expander("Foundation Skills"):
        if ud.foundation_skills:
            st.info(f"{len(ud.foundation_skills)} foundation skills loaded")
            for fs in ud.foundation_skills:
                st.caption(f"{fs.skill}: {fs.description[:100]}...")
        else:
            st.warning("No foundation skills found.")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Back", use_container_width=True):
            st.session_state.step = 1; st.rerun()
    with c2:
        if st.button("Next: Configure Tasks", type="primary", use_container_width=True):
            st.session_state.step = 3; st.rerun()


def _parse_elements_text(text):
    import re
    elements, current = [], None
    for line in text.strip().split("\n"):
        line = line.strip()
        if not line: continue
        el = re.match(r'^(\d+)\.\s+(.+)', line)
        pc = re.match(r'^(\d+\.\d+)\s+(.+)', line)
        if el and not pc:
            current = Element(number=el.group(1), title=el.group(2))
            elements.append(current)
        elif pc and current:
            current.performance_criteria.append(PerformanceCriteria(number=pc.group(1), text=pc.group(2)))
    return elements


# ── STEP 3 ────────────────────────────────────────────────────────
def step3():
    step_header(3, "Configure Assessment Tasks")
    ud = st.session_state.unit_data
    s = st.session_state.settings

    methods = {
        "questioning": "Questioning (written or verbal)",
        "observation": "Direct Observation (reflecting authentic workplace activities)",
        "project": "Structured Activities (project, role plays, activity sheets)",
        "portfolio": "Product based (work products, logbook, portfolio)",
        "third_party": "Third party (reports, interviews and logbook verification)"
    }
    default_attempts = {"questioning": 3, "observation": 2, "project": 2, "portfolio": 2, "third_party": 2}
    type_options = list(methods.keys())
    type_labels = {
        "questioning": "Knowledge Questions",
        "observation": "Direct Observation",
        "third_party": "Third Party Report",
        "project": "Project",
        "portfolio": "Portfolio"
    }

    num_tasks = st.number_input("Number of assessment tasks", 1, 6, 3, key=k("ntasks"))

    tasks = []
    for i in range(int(num_tasks)):
        with st.expander(f"Assessment Task {i+1}", expanded=(i < 3)):
            c1, c2 = st.columns(2)
            with c1:
                tt = st.selectbox(
                    "Type", type_options,
                    format_func=lambda x: type_labels.get(x, x.replace('_',' ').title()),
                    index=min(i, len(type_options)-1),
                    key=k(f"tt{i}")
                )
            with c2:
                att = st.number_input("Attempts", 1, 5, default_attempts.get(tt, 2), key=k(f"att{i}"))
            desc = st.text_area("Description (leave blank for default)", height=60, key=k(f"desc{i}"))
            ctx = st.multiselect(
                "Context of Assessment",
                ["Simulated workplace", "Active workplace", "Training Room", "Own Environment"],
                default=(
                    ["Training Room", "Own Environment"] if tt == "questioning"
                    else ["Active workplace"] if tt == "third_party"
                    else ["Simulated workplace"]
                ),
                key=k(f"ctx{i}")
            )
            tasks.append({
                "number": i+1, "type": tt, "method": methods.get(tt, ""),
                "attempts": att, "title": "", "description": desc,
                "context": "  ".join([f"☒ {c}" for c in ctx]),
                "evidence": "", "questions": []
            })

    st.session_state.tasks = tasks

    st.markdown("---")
    st.text_area(
        "AI Statement",
        value=s.get("default_ai_statement", ""),
        height=80, key="ai_statement"
    )

    # AI generation
    st.markdown("---")
    has_questioning = any(t["type"] == "questioning" for t in tasks)
    provider = s.get("ai_provider", "none")

    if has_questioning and provider != "none":
        st.markdown("**Generate Assessment Questions**")
        ke_count = len(ud.knowledge_evidence) if ud else 0
        provider_name = "Ollama" if provider == "ollama" else "Claude API"
        st.info(f"{ke_count} knowledge evidence items will be used to generate questions via {provider_name}")

        can_generate = True
        if provider == "ollama" and not check_ollama_available():
            st.error("Ollama is not running. Open the Ollama app or run `ollama serve`.")
            can_generate = False
        elif provider == "claude" and not s.get("claude_api_key"):
            st.warning("Enter your Claude API key in Settings (sidebar).")
            can_generate = False

        if can_generate:
            if st.button("Generate Questions", type="primary"):
                with st.spinner("Generating questions — this may take a few minutes..."):
                    progress = st.empty()
                    questions = generate_questions(
                        ud,
                        provider=provider,
                        api_key=s.get("claude_api_key", ""),
                        model=s.get("ollama_model", "llama3.1:8b"),
                        progress_callback=lambda m: progress.text(m)
                    )
                    progress.empty()

                if questions:
                    st.session_state.ai_questions = questions
                    st.success(f"Generated {len(questions)} questions with benchmark answers")
                else:
                    st.error("Generation failed. Check your provider settings in the sidebar.")

        if st.session_state.ai_questions:
            with st.expander(f"Preview: {len(st.session_state.ai_questions)} questions", expanded=False):
                for q in st.session_state.ai_questions:
                    st.markdown(f"**Q{q.get('number', '?')}:** {q.get('question', '')}")
                    bench = q.get('benchmark', '')
                    if bench:
                        st.caption(f"Benchmark: {bench[:200]}{'...' if len(bench) > 200 else ''}")
                    st.markdown("---")

    elif has_questioning and provider == "none":
        st.caption("Enable AI question generation in Settings (sidebar) for better output quality.")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Back", use_container_width=True):
            st.session_state.step = 2; st.rerun()
    with c2:
        if st.button("Next: Generate Document", type="primary", use_container_width=True):
            st.session_state.step = 4; st.rerun()


# ── STEP 4 ────────────────────────────────────────────────────────
def step4():
    step_header(4, "Generate Assessor Guide")
    ud = st.session_state.unit_data
    tasks = st.session_state.tasks

    st.markdown(f"**Unit:** {ud.code} — {ud.title}")
    for t in tasks:
        st.caption(f"AT{t['number']}: {t['method']}")

    ai_qs = st.session_state.ai_questions
    if ai_qs:
        st.info(f"{len(ai_qs)} AI-generated questions with benchmarks will be included")
    else:
        st.caption("No AI questions generated — using basic templates. Go back to Step 3 to generate them.")

    st.markdown("---")

    if st.button("Generate Assessor Guide", type="primary", use_container_width=True):
        with st.spinner("Building document..."):
            try:
                if ai_qs:
                    for task in tasks:
                        if task["type"] == "questioning":
                            task["questions"] = ai_qs
                            break

                config = {
                    "tasks": tasks,
                    "ai_statement": st.session_state.get("ai_statement", ""),
                    "assessor_req_detail": "",
                }
                output_dir = tempfile.mkdtemp()
                filename = f"{ud.code}_Assessor_Guide_{datetime.now().strftime('%Y%m%d')}.docx"
                path = os.path.join(output_dir, filename)
                build_assessor_guide(ud, config, path)
                st.session_state.generated_path = path
                st.session_state.generated_filename = filename
                st.success("Document generated successfully")
            except Exception as e:
                import traceback
                import logging
                logging.error("Document generation failed:\n%s", traceback.format_exc())
                st.error(f"Document generation failed: {e}")
                st.caption("If this error persists, review your unit data in Step 2 and try again.")

    if st.session_state.get("generated_path") and os.path.exists(st.session_state.generated_path):
        with open(st.session_state.generated_path, "rb") as f:
            st.download_button(
                "Download Assessor Guide (.docx)",
                f.read(),
                st.session_state.get("generated_filename", "Assessor_Guide.docx"),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
        st.info(
            "Open in Microsoft Word to review questions, complete the mapping matrix, "
            "and add any CDU-specific details."
        )

    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Back", use_container_width=True):
            st.session_state.step = 3; st.rerun()
    with c2:
        if st.button("Start New Guide", use_container_width=True):
            for key in list(st.session_state.keys()):
                if key != "settings":
                    del st.session_state[key]
            st.rerun()


if __name__ == "__main__":
    main()
