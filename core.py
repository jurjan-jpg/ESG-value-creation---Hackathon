"""
Core logic for the ESG Value Creation Module.
Kept separate from the Streamlit UI so it can be tested on its own.
"""
import json
import re
from io import BytesIO
from datetime import date

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# The six prioritisation dimensions (key, label)
DIMENSIONS = [
    ("risk_reduction", "Risk reduction"),
    ("ebitda_impact", "EBITDA impact"),
    ("revenue_growth", "Revenue growth"),
    ("financing_benefit", "Financing benefits"),
    ("operational_efficiency", "Operational efficiency"),
    ("implementation_feasibility", "Implementation feasibility"),
]

DEFAULT_WEIGHTS = {
    "risk_reduction": 25,
    "ebitda_impact": 20,
    "revenue_growth": 10,
    "financing_benefit": 10,
    "operational_efficiency": 15,
    "implementation_feasibility": 20,
}

PHASES = [
    ("0-30", "Day 0-30 · Stabilise & assign ownership"),
    ("30-60", "Day 30-60 · Implement quick wins"),
    ("60-100", "Day 60-100 · Embed KPIs & prepare board reporting"),
]


# ---------------------------------------------------------------- extraction
def extract_pptx_text(file_bytes):
    """Pull all readable text and tables out of a .pptx file."""
    prs = Presentation(BytesIO(file_bytes))
    chunks = []
    for i, slide in enumerate(prs.slides):
        parts = []
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                parts.append(shape.text_frame.text.strip())
            if shape.has_table:
                rows = []
                for r in shape.table.rows:
                    rows.append(" | ".join(c.text.strip() for c in r.cells))
                parts.append("TABLE:\n" + "\n".join(rows))
        if parts:
            chunks.append("--- Slide %d ---\n%s" % (i + 1, "\n".join(parts)))
    return "\n\n".join(chunks)


def parse_json(text):
    """Robustly parse JSON that may be wrapped in ``` fences or prose."""
    text = text.strip()
    text = re.sub(r"^```[a-zA-Z]*", "", text).strip()
    text = re.sub(r"```$", "", text).strip()
    start, end = text.find("{"), text.rfind("}")
    if start != -1 and end != -1:
        text = text[start:end + 1]
    return json.loads(text)


def call_claude(client, model, prompt, max_tokens=8000):
    msg = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        messages=[{"role": "user", "content": prompt}],
    )
    return "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")


# ----------------------------------------------------------------- prompts
EXTRACTION_PROMPT = """You are an ESG analyst. Below is the full text extracted from an ESG due diligence (DD) report. Extract a structured summary of the material ESG themes and their findings.

Return ONLY valid JSON (no markdown, no commentary), in EXACTLY this shape:
{
 "company": {"name": "...", "sector": "...", "summary": "one sentence about the company"},
 "overall": {"abstain_from_deal": "Yes or No", "headline": "one sentence overall ESG verdict"},
 "themes": [
   {"code": "E1", "name": "theme name", "pillar": "E or S or G",
    "current_maturity": "the maturity rating word used in the report, or null",
    "key_findings": "2-3 sentences on the key risks/gaps for this theme",
    "recommended_actions": ["short concrete action", "another action"],
    "investment_eur": 100000,
    "value_creation_eur": 400000,
    "suggested_kpis": ["kpi name", "kpi name"]}
 ]
}

Rules:
- Use the report's own euro figures where given. Convert 'k' to thousands and 'm'/'M' to millions (e.g. "€100k" -> 100000, "€1.2M" -> 1200000).
- If a euro number is missing, "To be done", or unknown, use null.
- pillar: E = environmental, S = social, G = governance.
- Include every material theme you can find.

REPORT TEXT:
%s
"""

PLANNER_PROMPT = """You are a private equity value-creation expert. Turn these ESG due-diligence findings into a prioritised 100-day post-acquisition action plan.

For EACH recommended action across the themes, create one initiative. Score each initiative from 1 (low) to 5 (high) on these six dimensions:
- risk_reduction: how much it reduces ESG / regulatory / business risk
- ebitda_impact: effect on EBITDA via cost savings or margin
- revenue_growth: effect on revenue or commercial win-rate
- financing_benefit: effect on financing terms (sustainability-linked loan margins, refinancing, investor appeal)
- operational_efficiency: process / operational improvement
- implementation_feasibility: how easy and fast it is to do (5 = very easy / quick)

Assign each initiative to exactly ONE phase:
- "0-30"  = Day 0-30, Stabilise & assign ownership (foundational items, governance, baselines, highest-risk fixes)
- "30-60" = Day 30-60, Implement quick wins (high feasibility, visible impact)
- "60-100" = Day 60-100, Embed KPIs & prepare board reporting (targets, dashboards, reporting cadence)
Use dependency logic: anything that needs a baseline, an owner, or another initiative done first should go in a later phase.

Carry the investment_eur and value_creation_eur from the theme. If a theme has several initiatives, put the euro figures on the single most relevant initiative and use null for the others, so the totals are NOT inflated.
Give each initiative a generic owner_role such as: "CFO", "Head of Operations", "Head of HR", "Head of IT / Security", "ESG Lead", "General Counsel".

Return ONLY valid JSON (no markdown), in EXACTLY this shape:
{
 "initiatives": [
   {"id": "E1-01", "theme": "E1", "theme_name": "...", "pillar": "E",
    "title": "short imperative title", "description": "1-2 sentences",
    "scores": {"risk_reduction": 4, "ebitda_impact": 3, "revenue_growth": 1,
               "financing_benefit": 2, "operational_efficiency": 4, "implementation_feasibility": 5},
    "phase": "0-30", "owner_role": "ESG Lead",
    "investment_eur": 20000, "value_creation_eur": 80000,
    "kpis": [{"name": "...", "target": "..."}],
    "dependencies": [], "rationale": "one sentence on the scoring and phase choice"}
 ]
}

FINDINGS JSON:
%s
"""

BOARD_PROMPT = """You are preparing a board update on a 100-day ESG value-creation plan for a private equity owner. Below is the plan with current tracking status and pre-computed metrics. Write a concise, commercially framed update.

Return ONLY valid JSON (no markdown), in EXACTLY this shape:
{
 "executive_summary": "3-4 sentences, commercial framing: value at stake, ROI, risk reduced, momentum",
 "phases": [
   {"phase": "0-30", "status": "On track / At risk / Complete", "highlights": ["...", "..."]},
   {"phase": "30-60", "status": "...", "highlights": ["..."]},
   {"phase": "60-100", "status": "...", "highlights": ["..."]}
 ],
 "top_priorities": ["...", "...", "..."],
 "risks_and_asks": ["...", "..."]
}

PLAN + TRACKING + METRICS JSON:
%s
"""


# ------------------------------------------------------------- scoring / math
def compute_composite(initiative, weights):
    s = initiative.get("scores", {})
    total_w = sum(weights.values()) or 1
    val = sum(float(s.get(k, 0)) * weights[k] for k in weights) / total_w
    return round(val, 2)


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def compute_metrics(initiatives, tracking):
    total_value = sum(_num(i.get("value_creation_eur")) for i in initiatives)
    total_invest = sum(_num(i.get("investment_eur")) for i in initiatives)
    roi = (total_value / total_invest) if total_invest else 0.0
    done = sum(1 for i in initiatives if tracking.get(i["id"]))
    n = len(initiatives)
    return {
        "total_value_creation_eur": total_value,
        "total_investment_eur": total_invest,
        "roi_multiple": round(roi, 1),
        "initiatives_total": n,
        "initiatives_complete": done,
        "percent_complete": round(100 * done / n) if n else 0,
    }


def build_plan_object(findings, initiatives, weights, tracking):
    """Assemble the full, reusable plan JSON (the standard output)."""
    ranked = sorted(initiatives, key=lambda i: compute_composite(i, weights), reverse=True)
    for rank, i in enumerate(ranked, 1):
        i["composite_score"] = compute_composite(i, weights)
        i["priority_rank"] = rank
        i["status"] = "complete" if tracking.get(i["id"]) else "open"
    phases = {}
    for key, label in PHASES:
        phases[key] = {
            "label": label,
            "initiative_ids": [i["id"] for i in ranked if i.get("phase") == key],
        }
    return {
        "company": findings.get("company", {}),
        "overall": findings.get("overall", {}),
        "generated_on": str(date.today()),
        "weights": weights,
        "metrics": compute_metrics(initiatives, tracking),
        "phases": phases,
        "initiatives": ranked,
    }


# ----------------------------------------------------------- board rendering
def board_markdown(company_name, narrative, metrics):
    eur = lambda v: "EUR {:,.0f}".format(v)
    lines = []
    lines.append("# 100-Day ESG Value Creation Plan - Board Update")
    lines.append("**Company:** %s  |  **Date:** %s" % (company_name, date.today()))
    lines.append("")
    lines.append("## Executive summary")
    lines.append(narrative.get("executive_summary", ""))
    lines.append("")
    lines.append("## Key metrics")
    lines.append("- Total annual value creation opportunity: **%s**" % eur(metrics["total_value_creation_eur"]))
    lines.append("- Total implementation investment: **%s**" % eur(metrics["total_investment_eur"]))
    lines.append("- Return on investment: **%sx**" % metrics["roi_multiple"])
    lines.append("- Initiatives: **%d total, %d complete (%d%%)**" % (
        metrics["initiatives_total"], metrics["initiatives_complete"], metrics["percent_complete"]))
    lines.append("")
    lines.append("## Progress by phase")
    for ph in narrative.get("phases", []):
        lines.append("**%s - %s**" % (ph.get("phase", ""), ph.get("status", "")))
        for h in ph.get("highlights", []):
            lines.append("- %s" % h)
        lines.append("")
    lines.append("## Top priorities")
    for p in narrative.get("top_priorities", []):
        lines.append("- %s" % p)
    lines.append("")
    lines.append("## Risks & asks")
    for r in narrative.get("risks_and_asks", []):
        lines.append("- %s" % r)
    return "\n".join(lines)


def board_pptx(company_name, narrative, metrics):
    """Build a simple 2-slide board update deck and return it as bytes."""
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    navy = RGBColor(0x1F, 0x2A, 0x6E)
    orange = RGBColor(0xD9, 0x53, 0x1E)
    eur = lambda v: "EUR {:,.0f}".format(v)

    # Slide 1 - title + metrics
    s = prs.slides.add_slide(blank)
    tb = s.shapes.add_textbox(Inches(0.6), Inches(0.5), Inches(12), Inches(1.2)).text_frame
    tb.text = "100-Day ESG Value Creation Plan"
    tb.paragraphs[0].runs[0].font.size = Pt(34)
    tb.paragraphs[0].runs[0].font.bold = True
    tb.paragraphs[0].runs[0].font.color.rgb = orange
    p = tb.add_paragraph()
    p.text = "%s  |  Board update  |  %s" % (company_name, date.today())
    p.runs[0].font.size = Pt(16)
    p.runs[0].font.color.rgb = navy

    cards = [
        ("Value creation / yr", eur(metrics["total_value_creation_eur"])),
        ("Investment", eur(metrics["total_investment_eur"])),
        ("ROI", "%sx" % metrics["roi_multiple"]),
        ("Progress", "%d%% complete" % metrics["percent_complete"]),
    ]
    x = 0.6
    for label, value in cards:
        box = s.shapes.add_textbox(Inches(x), Inches(2.0), Inches(2.9), Inches(1.3)).text_frame
        box.word_wrap = True
        box.text = value
        box.paragraphs[0].runs[0].font.size = Pt(26)
        box.paragraphs[0].runs[0].font.bold = True
        box.paragraphs[0].runs[0].font.color.rgb = navy
        lp = box.add_paragraph()
        lp.text = label
        lp.runs[0].font.size = Pt(13)
        x += 3.05

    sb = s.shapes.add_textbox(Inches(0.6), Inches(3.6), Inches(12.1), Inches(3.4)).text_frame
    sb.word_wrap = True
    sb.text = "Executive summary"
    sb.paragraphs[0].runs[0].font.size = Pt(18)
    sb.paragraphs[0].runs[0].font.bold = True
    ep = sb.add_paragraph()
    ep.text = narrative.get("executive_summary", "")
    ep.runs[0].font.size = Pt(14)

    # Slide 2 - progress by phase
    s2 = prs.slides.add_slide(blank)
    t2 = s2.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12), Inches(0.8)).text_frame
    t2.text = "Progress by phase"
    t2.paragraphs[0].runs[0].font.size = Pt(26)
    t2.paragraphs[0].runs[0].font.bold = True
    t2.paragraphs[0].runs[0].font.color.rgb = orange
    body = s2.shapes.add_textbox(Inches(0.6), Inches(1.3), Inches(12.1), Inches(5.6)).text_frame
    body.word_wrap = True
    first = True
    for ph in narrative.get("phases", []):
        para = body.paragraphs[0] if first else body.add_paragraph()
        first = False
        para.text = "%s  -  %s" % (ph.get("phase", ""), ph.get("status", ""))
        para.runs[0].font.size = Pt(16)
        para.runs[0].font.bold = True
        para.runs[0].font.color.rgb = navy
        for h in ph.get("highlights", []):
            hp = body.add_paragraph()
            hp.text = "  -  %s" % h
            hp.runs[0].font.size = Pt(13)

    out = BytesIO()
    prs.save(out)
    return out.getvalue()
