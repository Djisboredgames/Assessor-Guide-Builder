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
    page_icon="AG",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-title { font-size:1.8rem; font-weight:700; color:#1A1A2E; margin-bottom:0.2rem; }
    .main-sub { font-size:0.95rem; color:#555; margin-bottom:1.5rem; }
    .step-indicator { display:inline-block; background:#1A1A2E; color:white; border-radius:50%;
        width:28px; height:28px; text-align:center; line-height:28px; font-weight:600;
        font-size:0.85rem; margin-right:6px; }
    .step-label { font-size:1.15rem; font-weight:600; color:#1A1A2E; }
    .settings-header { font-size:1rem; font-weight:600; color:#1A1A2E; margin-top:1rem; }
    div[data-testid="stSidebar"] { background:#fafafa; }
    section[data-testid="stSidebar"] .stTextInput label,
    section[data-testid="stSidebar"] .stTextArea label,
    section[data-testid="stSidebar"] .stSelectbox label { font-size:0.85rem; }
</style>
""", unsafe_allow_html=True)


def step_header(n, title):
    st.markdown(
        f'<span class="step-indicator">{n}</span>'
        f'<span class="step-label">{title}</span>',
        unsafe_allow_html=True
    )
    st.markdown("---")


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
        st.markdown('<p class="settings-header">Settings</p>', unsafe_allow_html=True)

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
        st.caption("Assessor Guide Builder v1.0")


# ══════════════════════════════════════════════════════════════════
def main():
    init()
    render_sidebar()

    st.markdown('<p class="main-title">Assessor Guide Builder</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="main-sub">Generate VET Assessor Guides from training.gov.au unit data</p>',
        unsafe_allow_html=True
    )

    steps = ["Fetch Unit", "Review Data", "Configure Tasks", "Generate"]
    cols = st.columns(len(steps))
    for i, (col, label) in enumerate(zip(cols, steps), 1):
        with col:
            if i < st.session_state.step:
                st.success(f"Step {i}: {label}")
            elif i == st.session_state.step:
                st.info(f"Step {i}: {label}")
            else:
                st.markdown(f"Step {i}: {label}")

    st.markdown("")
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
                st.error(f"Error: {e}")
                import traceback
                st.code(traceback.format_exc())

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
