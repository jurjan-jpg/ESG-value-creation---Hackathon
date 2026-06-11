"""ESG Value Creation Module - Streamlit app.  Run:  streamlit run app.py"""
import json
import streamlit as st
import anthropic
import core

st.set_page_config(page_title="ESG Value Creation Module", layout="wide")

MODELS = {"Claude Opus (best quality)": "claude-opus-4-8",
          "Claude Sonnet (faster / cheaper)": "claude-sonnet-4-6",
          "Claude Haiku (fastest)": "claude-haiku-4-5-20251001"}

ss = st.session_state
ss.setdefault("extract", None)
ss.setdefault("advice", None)
ss.setdefault("tracking", {})
ss.setdefault("board", None)


def get_client(key):
    if not key:
        return None
    try:
        return anthropic.Anthropic(api_key=key)
    except Exception:
        return None


with st.sidebar:
    st.header("Setup")
    try:
        secret_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:
        secret_key = ""
    api_key = secret_key or st.text_input("Anthropic API key", type="password")
    model = MODELS[st.selectbox("Model", list(MODELS.keys()), index=0)]
    st.divider()
    st.header("Prioritisation weights")
    st.caption("Move the sliders to change what matters most. The plan re-ranks instantly.")
    weights = {k: st.slider(lbl, 0, 100, core.DEFAULT_WEIGHTS[k], 5) for k, lbl in core.DIMENSIONS}

client = get_client(api_key)
st.title("ESG Value Creation Module")
st.caption("Two stages: EXTRACT facts from the DD (with sources), then ADVISE a detailed 100-day plan.")

# ---- Step 1: EXTRACT
st.subheader("1) Extract facts from the DD")
uploaded = st.file_uploader("ESG DD report (.pptx)", type=["pptx"])
if st.button("Extract facts", type="primary"):
    if not client:
        st.error("Paste your Anthropic API key in the sidebar first.")
    elif not uploaded:
        st.error("Upload a .pptx DD report first.")
    else:
        with st.spinner("Reading the report and pulling facts (with sources)..."):
            try:
                text = core.extract_pptx_text(uploaded.read())
                ss.extract = core.parse_json(core.call_claude(client, model, core.extract_prompt(text[:60000])))
                ss.advice = None
                ss.board = None
                st.success("Facts extracted.")
            except Exception as e:
                st.error("Extract failed: %s" % e)

if ss.extract:
    e = ss.extract
    st.markdown("**%s** — %s" % (e.get("company_profile", {}).get("name", "Company"),
                                 e.get("overall", {}).get("headline", "")))
    with st.expander("See extracted facts and sources"):
        st.write("**Company profile**", e.get("company_profile", {}))
        st.write("**Key metrics**", e.get("key_metrics", []))
        st.write("**Themes / findings**", [{"theme": t.get("code"), "finding": t.get("finding"),
                                             "source": t.get("source")} for t in e.get("themes", [])])
    gaps = e.get("data_gaps", [])
    if gaps:
        st.warning("Data to confirm (not found in the DD) — %d item(s):" % len(gaps))
        for g in gaps:
            st.markdown("- %s" % g)

# ---- Step 2: ADVISE
if ss.extract:
    st.subheader("2) Advise a detailed 100-day plan")
    if st.button("Build 100-day plan", type="primary"):
        if not client:
            st.error("Paste your Anthropic API key in the sidebar first.")
        else:
            with st.spinner("Designing the detailed roadmap, targets and value story..."):
                try:
                    ss.advice = core.parse_json(core.call_claude(
                        client, model, core.advise_prompt(json.dumps(ss.extract)), max_tokens=8000))
                    ss.tracking = {i["id"]: False for i in ss.advice.get("initiatives", [])}
                    ss.board = None
                    st.success("Plan built. Adjust the weight sliders on the left to re-prioritise.")
                except Exception as e:
                    st.error("Advise failed: %s" % e)

# ---- Step 3: the plan
if ss.advice:
    inits = ss.advice.get("initiatives", [])
    metrics = core.compute_metrics(ss.extract, inits, ss.tracking)
    st.subheader("3) The 100-day plan")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Annual value creation", "EUR {:,.0f}".format(metrics["total_value_creation_eur"]))
    c2.metric("Investment", "EUR {:,.0f}".format(metrics["total_investment_eur"]))
    c3.metric("Return on investment", "%sx" % metrics["roi_multiple"])
    c4.metric("Initiatives", "%d" % metrics["initiatives_total"])

    for i in inits:
        i["composite_score"] = core.compute_composite(i, weights)
    ranked = sorted(inits, key=lambda i: i["composite_score"], reverse=True)
    cols = st.columns(3)
    for col, (pkey, plabel) in zip(cols, core.PHASES):
        with col:
            st.markdown("#### %s" % plabel)
            for i in [x for x in ranked if x.get("phase") == pkey]:
                with st.container(border=True):
                    st.markdown("**%s**  \n`%s` · score %.1f · %s"
                                % (i.get("title", ""), i.get("theme", ""), i["composite_score"], i.get("owner_role", "-")))
                    st.caption(i.get("objective", ""))
                    for a in i.get("activities", []):
                        st.markdown("- %s" % a)
                    kpi = i.get("recommended_kpi", {}) or {}
                    st.caption("KPI: %s | baseline: %s → target: %s"
                               % (kpi.get("name", "-"), kpi.get("baseline", "-"), kpi.get("target", "-")))
                    ss.tracking[i["id"]] = st.checkbox("Done", value=ss.tracking.get(i["id"], False), key="trk_" + i["id"])

    plan_obj = core.build_plan_object(ss.extract, ss.advice, weights, ss.tracking)
    d1, d2, d3 = st.columns(3)
    d1.download_button("Download findings one-pager (PowerPoint)", data=core.findings_onepager_pptx(plan_obj),
                       file_name="esg_findings_onepager.pptx",
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    d2.download_button("Download detailed 100-day plan (PowerPoint)", data=core.detailed_plan_pptx(plan_obj),
                       file_name="100_day_esg_plan_detailed.pptx",
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    d3.download_button("Download plan as data (JSON)", data=json.dumps(plan_obj, indent=2),
                       file_name="100_day_esg_plan.json", mime="application/json")

# ---- Step 4: board update
if ss.advice:
    st.subheader("4) Generate the board update")
    if st.button("Generate board update", type="primary"):
        if not client:
            st.error("Paste your Anthropic API key in the sidebar first.")
        else:
            with st.spinner("Writing the board update..."):
                try:
                    metrics = core.compute_metrics(ss.extract, ss.advice.get("initiatives", []), ss.tracking)
                    payload = {"company": ss.extract.get("company_profile", {}), "metrics": metrics,
                               "initiatives": [{"title": i.get("title"), "phase": i.get("phase"),
                                                "status": "complete" if ss.tracking.get(i["id"]) else "open"}
                                               for i in ss.advice.get("initiatives", [])]}
                    ss.board = core.parse_json(core.call_claude(client, model, core.board_prompt(json.dumps(payload))))
                    st.success("Board update ready.")
                except Exception as e:
                    st.error("Board update failed: %s" % e)

if ss.board:
    metrics = core.compute_metrics(ss.extract, ss.advice.get("initiatives", []), ss.tracking)
    name = ss.extract.get("company_profile", {}).get("name", "Company")
    md = core.board_markdown(name, ss.board, metrics)
    st.markdown(md)
    b1, b2 = st.columns(2)
    b1.download_button("Board update (Markdown)", data=md, file_name="board_update.md", mime="text/markdown")
    b2.download_button("Board update (PowerPoint)", data=core.board_pptx(name, ss.board, metrics),
                       file_name="board_update.pptx",
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
