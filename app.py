"""
ESG Value Creation Module - Streamlit app.
Run with:  streamlit run app.py
"""
import json
import streamlit as st
import anthropic

import core

st.set_page_config(page_title="ESG Value Creation Module", layout="wide")

# Model choices (Opus by default, with cheaper/faster fallbacks)
MODELS = {
    "Claude Opus (best quality)": "claude-opus-4-8",
    "Claude Sonnet (faster / cheaper)": "claude-sonnet-4-6",
    "Claude Haiku (fastest)": "claude-haiku-4-5-20251001",
}

# ---- session state init ----
ss = st.session_state
ss.setdefault("findings", None)
ss.setdefault("initiatives", None)
ss.setdefault("tracking", {})
ss.setdefault("board", None)


def get_client(api_key):
    if not api_key:
        return None
    try:
        return anthropic.Anthropic(api_key=api_key)
    except Exception:
        return None


# ---------------------------------------------------------------- sidebar
with st.sidebar:
    st.header("Setup")
    # API key: from Streamlit secrets if present, else a text box
    secret_key = ""
    try:
        secret_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:
        secret_key = ""
    api_key = secret_key or st.text_input("Anthropic API key", type="password",
                                           help="Paste your key from console.anthropic.com")
    model_label = st.selectbox("Model", list(MODELS.keys()), index=0)
    model = MODELS[model_label]

    st.divider()
    st.header("Prioritisation weights")
    st.caption("Move the sliders to change what matters most. The plan re-ranks instantly.")
    weights = {}
    for key, label in core.DIMENSIONS:
        weights[key] = st.slider(label, 0, 100, core.DEFAULT_WEIGHTS[key], 5)

client = get_client(api_key)

# ---------------------------------------------------------------- header
st.title("ESG Value Creation Module")
st.caption("Turn an ESG due-diligence report into a prioritised 100-day value-creation plan.")

# ---------------------------------------------------------------- step 1
st.subheader("1) Upload the DD report and generate findings")
uploaded = st.file_uploader("ESG DD report (.pptx)", type=["pptx"])

if st.button("Run ESG DD -> Generate findings", type="primary"):
    if not client:
        st.error("Please paste your Anthropic API key in the sidebar first.")
    elif not uploaded:
        st.error("Please upload a .pptx DD report first.")
    else:
        with st.spinner("Reading the report and extracting findings (this can take ~30-60s on Opus)..."):
            try:
                text = core.extract_pptx_text(uploaded.read())
                raw = core.call_claude(client, model, core.EXTRACTION_PROMPT % text[:60000])
                ss.findings = core.parse_json(raw)
                ss.initiatives = None
                ss.board = None
                st.success("Findings generated.")
            except Exception as e:
                st.error("Something went wrong while generating findings: %s" % e)

if ss.findings:
    f = ss.findings
    comp = f.get("company", {})
    st.markdown("**%s** - %s" % (comp.get("name", "Company"), comp.get("summary", "")))
    st.markdown("Overall verdict: *%s*" % f.get("overall", {}).get("headline", ""))
    with st.expander("See extracted findings by theme"):
        for t in f.get("themes", []):
            st.markdown("**%s - %s** (%s)" % (t.get("code"), t.get("name"), t.get("current_maturity")))
            st.write(t.get("key_findings", ""))

# ---------------------------------------------------------------- step 2
if ss.findings:
    st.subheader("2) Convert findings into a 100-day plan")
    if st.button("Generate 100-day value-creation plan", type="primary"):
        if not client:
            st.error("Please paste your Anthropic API key in the sidebar first.")
        else:
            with st.spinner("Scoring and sequencing initiatives..."):
                try:
                    raw = core.call_claude(
                        client, model,
                        core.PLANNER_PROMPT % json.dumps(ss.findings),
                        max_tokens=8000,
                    )
                    ss.initiatives = core.parse_json(raw)["initiatives"]
                    ss.tracking = {i["id"]: False for i in ss.initiatives}
                    ss.board = None
                    st.success("Plan generated. Adjust the weight sliders on the left to re-prioritise.")
                except Exception as e:
                    st.error("Something went wrong while building the plan: %s" % e)

# ---------------------------------------------------------------- step 3
if ss.initiatives:
    inits = ss.initiatives
    metrics = core.compute_metrics(inits, ss.tracking)

    st.subheader("3) The 100-day plan")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Annual value creation", "EUR {:,.0f}".format(metrics["total_value_creation_eur"]))
    c2.metric("Investment", "EUR {:,.0f}".format(metrics["total_investment_eur"]))
    c3.metric("Return on investment", "%sx" % metrics["roi_multiple"])
    c4.metric("Initiatives", "%d" % metrics["initiatives_total"])

    # rank live by current weights
    for i in inits:
        i["composite_score"] = core.compute_composite(i, weights)
    ranked = sorted(inits, key=lambda i: i["composite_score"], reverse=True)

    cols = st.columns(3)
    for col, (pkey, plabel) in zip(cols, core.PHASES):
        with col:
            st.markdown("#### %s" % plabel)
            phase_items = [i for i in ranked if i.get("phase") == pkey]
            if not phase_items:
                st.caption("No initiatives in this phase.")
            for i in phase_items:
                with st.container(border=True):
                    st.markdown("**%s**  \n`%s` · score %.1f · owner: %s"
                                % (i["title"], i["theme"], i["composite_score"], i.get("owner_role", "-")))
                    if i.get("value_creation_eur"):
                        st.caption("Value EUR {:,.0f}  |  Invest EUR {:,.0f}".format(
                            core._num(i.get("value_creation_eur")), core._num(i.get("investment_eur"))))
                    done = st.checkbox("Done", value=ss.tracking.get(i["id"], False), key="trk_" + i["id"])
                    ss.tracking[i["id"]] = done
                    with st.expander("Why / KPIs"):
                        st.write(i.get("rationale", ""))
                        for k in i.get("kpis", []):
                            st.caption("KPI: %s (%s)" % (k.get("name", ""), k.get("target", "")))

    # downloadable standard JSON
    plan_obj = core.build_plan_object(ss.findings, inits, weights, ss.tracking)
    st.download_button("Download plan as JSON",
                       data=json.dumps(plan_obj, indent=2),
                       file_name="100_day_esg_plan.json", mime="application/json")

# ---------------------------------------------------------------- step 4
if ss.initiatives:
    st.subheader("4) Generate the board update")
    if st.button("Generate board update", type="primary"):
        if not client:
            st.error("Please paste your Anthropic API key in the sidebar first.")
        else:
            with st.spinner("Writing the board update..."):
                try:
                    metrics = core.compute_metrics(ss.initiatives, ss.tracking)
                    payload = {
                        "company": ss.findings.get("company", {}),
                        "metrics": metrics,
                        "initiatives": [
                            {"title": i["title"], "phase": i.get("phase"),
                             "status": "complete" if ss.tracking.get(i["id"]) else "open",
                             "value_creation_eur": i.get("value_creation_eur")}
                            for i in ss.initiatives
                        ],
                    }
                    raw = core.call_claude(client, model, core.BOARD_PROMPT % json.dumps(payload))
                    ss.board = core.parse_json(raw)
                    st.success("Board update ready.")
                except Exception as e:
                    st.error("Something went wrong while writing the board update: %s" % e)

if ss.board:
    metrics = core.compute_metrics(ss.initiatives, ss.tracking)
    name = ss.findings.get("company", {}).get("name", "Company")
    md = core.board_markdown(name, ss.board, metrics)
    st.markdown(md)
    cc1, cc2 = st.columns(2)
    cc1.download_button("Download board update (Markdown)", data=md,
                        file_name="board_update.md", mime="text/markdown")
    cc2.download_button("Download board update (PowerPoint)",
                        data=core.board_pptx(name, ss.board, metrics),
                        file_name="board_update.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
