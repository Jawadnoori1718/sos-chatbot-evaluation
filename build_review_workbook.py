#!/usr/bin/env python3
"""
Builds the criterion-review workbook requested by Dr Aislinn Gomez Bergin,
expert review session, 11 August 2026.

Her specification (transcript 12:00):
    "...have the responses, then all of the criteria... the reasoning that each
     of these judges give under each of these criteria, but also including the
     ones that they don't think are relevant."

Reviewers assess the JUDGES' REASONING, not the scores (transcript 11:27).

Inputs
------
  sos-v03-stakeholder.yml   the blueprint  -> criterion text, weights, prompts
  run.json                  a Weval export -> scores, judge std-dev, agreement

Two Weval files exist. The *analysis export* (small) carries scores but not the
judge reasoning or the model responses. The *comparison* file written by the CLI
carries everything. If the comparison file is supplied, the reasoning columns
fill automatically; otherwise they are left blank and clearly marked.

Usage
-----
    python build_review_workbook.py sos-v03-stakeholder.yml run.json
"""

import json
import sys
from pathlib import Path

import yaml
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "Weval-Criterion-Review.xlsx"

NAVY, ACC = "10233C", "C0431F"
HDRFILL = PatternFill("solid", fgColor=NAVY)
INFILL = PatternFill("solid", fgColor="8A4A2E")
SUBFILL = PatternFill("solid", fgColor="E8ECF2")
INPUT = PatternFill("solid", fgColor="FFF6DA")
FLAG = PatternFill("solid", fgColor="FDEEE9")
GAP = PatternFill("solid", fgColor="F0F2F5")
HDRFONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
BOLD = Font(name="Arial", size=10, bold=True)
BODY = Font(name="Arial", size=10)
SMALL = Font(name="Arial", size=9, color="55606E")
NOTE = Font(name="Arial", size=9, italic=True, color="8A8F98")
TITLE = Font(name="Arial", size=16, bold=True, color=NAVY)
THIN = Side(style="thin", color="D5DCE5")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOP = Alignment(vertical="top", wrap_text=True)
TOPC = Alignment(vertical="top", horizontal="center", wrap_text=True)

PRETTY = {
    "anthropic:claude-haiku-4-5-20251001": "Claude Haiku 4.5",
    "openai:gpt-4.1-mini": "GPT-4.1-mini",
    "openrouter:google/gemini-2.5-flash": "Gemini 2.5 Flash",
    "openrouter:meta-llama/llama-3.1-8b-instruct": "Llama 3.1 8B",
}
ORDER = ["Claude Haiku 4.5", "Gemini 2.5 Flash", "GPT-4.1-mini", "Llama 3.1 8B"]


def pretty(mid):
    base = mid.split("[")[0]
    return PRETTY.get(base, base)


def user_message(ctx):
    if isinstance(ctx, str):
        return ctx
    if isinstance(ctx, list) and ctx:
        return ctx[-1].get("content", "")
    return ""


# ───────────────────────────────── load ─────────────────────────────────
def load(bp_path, run_path):
    docs = list(yaml.safe_load_all(Path(bp_path).read_text()))
    header = docs[0] if isinstance(docs[0], dict) else {}
    prompts = next(d for d in docs if isinstance(d, list))
    crit = {p["id"]: [str(s) for s in (p.get("should") or [])] for p in prompts}
    cite = {p["id"]: p.get("citation", "") for p in prompts}

    run = json.loads(Path(run_path).read_text())
    cov = run["evaluationResults"]["llmCoverageScores"]
    weights = cov.get("__promptWeights", {}) or {}
    ctxs = run.get("promptContexts", {}) or {}
    resps = run.get("allFinalAssistantResponses", {}) or {}

    # scenario order: worst overall last, matching the poster
    order, means = [], {}
    for pid in run["promptIds"]:
        vals = [r["avgCoverageExtent"] for m, r in cov.get(pid, {}).items()
                if isinstance(r, dict) and r.get("avgCoverageExtent") is not None]
        means[pid] = sum(vals) / len(vals) * 100 if vals else 0
    order = sorted(run["promptIds"], key=lambda p: -means[p])

    resp_rows, rev_rows, n = [], [], 0
    have_reasoning = False

    for pid in order:
        w = weights.get(pid) or 1
        wtxt = "\u00D7%g" % w
        msg = user_message(ctxs.get(pid))
        models = cov.get(pid, {}) or {}
        named = sorted(models.items(), key=lambda kv: ORDER.index(pretty(kv[0]))
                       if pretty(kv[0]) in ORDER else 99)
        for mid, res in named:
            if not isinstance(res, dict):
                continue
            mname = pretty(mid)
            ja = res.get("judgeAgreement") or {}
            alpha = ja.get("krippendorffsAlpha") if isinstance(ja, dict) else ja
            rel = (ja.get("interpretation", "").capitalize()
                   if isinstance(ja, dict) and ja.get("interpretation") else "")
            body = (resps.get(pid, {}) or {}).get(mid)
            resp_rows.append([
                pid, wtxt, mname, msg, body or "",
                (res.get("avgCoverageExtent") or 0) * 100, alpha, rel,
            ])

            pts = res.get("pointAssessments") or []
            texts = crit.get(pid, [])
            for i, pt in enumerate(pts):
                n += 1
                ind = pt.get("individualJudgements") or []
                def pick(tag):
                    for j in ind:
                        if tag in (j.get("judgeModelId") or ""):
                            return j.get("coverageExtent"), (j.get("reflection") or "").strip()
                    return None, ""
                a_s, a_r = pick("judge-openai")
                b_s, b_r = pick("judge-anthropic")
                if a_s is None and b_s is None and ind:          # fallback to order
                    a_s = ind[0].get("coverageExtent"); a_r = (ind[0].get("reflection") or "").strip()
                    if len(ind) > 1:
                        b_s = ind[1].get("coverageExtent"); b_r = (ind[1].get("reflection") or "").strip()
                if a_r or b_r:
                    have_reasoning = True
                rev_rows.append([
                    f"R{n:04d}", pid, wtxt, mname,
                    pt.get("keyPointText") or (texts[i] if i < len(texts) else f"[criterion {i+1}]"),
                    "Inverted (should NOT)" if pt.get("isInverted") else "Standard",
                    pt.get("coverageExtent"),
                    (abs(a_s - b_s) if (a_s is not None and b_s is not None) else pt.get("judgeStdDev")),
                    a_s, a_r, b_s, b_r,
                    None,                       # M disagreement formula
                    "", "", "", "", "", "",     # N-S reviewer
                ])
    scen = [(pid, "\u00D7%g" % (weights.get(pid) or 1), cite.get(pid, "")) for pid in order]
    return resp_rows, rev_rows, scen, have_reasoning, header


# ──────────────────────────────── sheets ────────────────────────────────
def readme(wb, have_reasoning, n_resp, n_rev):
    ws = wb.create_sheet("Read me")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 112

    lines = [
        ("t", "Judging the judge \u2014 criterion review"),
        ("s", "Project 17 \u00B7 Student mental-health chatbot evaluation \u00B7 University of Nottingham"),
        ("s", f"Run sos-v03-runA, 16 August 2026 \u00B7 {n_resp} model responses \u00B7 {n_rev} criterion assessments"),
        ("g", ""),
        ("h", "What this is"),
        ("p", "Four candidate language models were tested against 11 scenarios and 74 criteria written by "
              "mental-health practitioners. Every response was scored independently by two LLM judges. "
              "This workbook sets out what the judges decided, criterion by criterion."),
        ("g", ""),
        ("h", "What we are asking you to review"),
        ("p", "Please assess the judges' JUDGEMENT, not their arithmetic. The question is not whether 0.63 is "
              "the right number, but whether the decision the judge reached is one you would have reached as "
              "a practitioner."),
        ("g", ""),
        ("h", "Two things to look for"),
        ("b", "1.  Where the two judges disagreed. Column M flags these. 110 of the 292 scored criteria "
              "(38%) had some disagreement between the judges."),
        ("b", "2.  Criteria the judges did not assess. Where a criterion was skipped, the row is marked. "
              "If you think it SHOULD have applied to that response, that is an important finding "
              "\u2014 record it in column P."),
        ("g", ""),
        ("h", "How to fill it in"),
        ("p", "Work on the 'Criterion review' sheet. The cream columns (N to S) are yours; everything to the "
              "left is evidence and should not be edited. Use the filter on row 1 to work scenario by "
              "scenario, or to show only the rows where the judges disagreed."),
        ("g", ""),
        ("h", "If it is too much"),
        ("p", "The 'Coverage summary' sheet shows how many rows fall under each scenario, so a subset can be "
              "agreed \u2014 for example only the weighted safety-critical scenarios, or splitting scenarios "
              "between reviewers. Please say what is workable rather than attempting all of it."),
        ("g", ""),
        ("h", "Criteria we should have written and did not"),
        ("p", "Use the 'Missing criteria' sheet \u2014 including anything specific to a UK deployment context."),
    ]
    if not have_reasoning:
        lines += [
            ("g", ""),
            ("w", "Note on columns J and L (judge reasoning) and column E on the Responses sheet"),
            ("p", "These are blank in this version. The file exported from the Weval dashboard carries the "
                  "scores but not the judges' written justifications or the full model responses; those live "
                  "in the run's comparison file, which is being retrieved. The workbook will be reissued with "
                  "those columns filled before review begins."),
        ]
    lines += [("g", ""), ("s", "Prepared by Jawad Noori.")]

    r = 2
    for kind, text in lines:
        c = ws.cell(row=r, column=2, value=text)
        if kind == "t":
            c.font = TITLE
        elif kind == "s":
            c.font = SMALL
        elif kind == "h":
            c.font = Font(name="Arial", size=11, bold=True, color=ACC)
        elif kind == "w":
            c.font = Font(name="Arial", size=11, bold=True, color="8A4A2E")
        elif kind == "b":
            c.font = BODY
            c.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
            ws.row_dimensions[r].height = 30
        else:
            c.font = BODY
        if kind == "p":
            c.alignment = TOP
            ws.row_dimensions[r].height = 32
        r += 1
    return ws


RESP = [("Scenario", 32), ("Weight", 8), ("Model", 18), ("Student message", 50),
        ("Model response", 80), ("Coverage %", 11), ("Judge agreement \u03B1", 14),
        ("Reliability", 12)]

def responses(wb, rows, have_reasoning):
    ws = wb.create_sheet("Responses")
    ws.freeze_panes = "A2"
    for i, (h, w) in enumerate(RESP, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font, c.fill, c.alignment, c.border = HDRFONT, HDRFILL, TOPC, BOX
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 30
    for r, row in enumerate(rows, 2):
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            c.font, c.alignment, c.border = BODY, TOP, BOX
        ws.cell(row=r, column=6).number_format = "0.0"
        ws.cell(row=r, column=7).number_format = "0.000"
        if str(row[7]).lower().startswith("unrel"):
            for i in (7, 8):
                ws.cell(row=r, column=i).fill = FLAG
        if not row[4]:
            c = ws.cell(row=r, column=5, value="[to be filled from the run comparison file]")
            c.font, c.fill = NOTE, GAP
        ws.row_dimensions[r].height = 56
    ws.auto_filter.ref = f"A1:H{len(rows)+1}"
    return ws


REV = [("Row ID", 8), ("Scenario", 30), ("Weight", 8), ("Model", 17),
       ("Criterion", 60), ("Criterion type", 14),
       ("Consensus score", 11), ("Gap between judges", 12),
       ("Judge A score", 10), ("Judge A \u2014 reasoning", 58),
       ("Judge B score", 10), ("Judge B \u2014 reasoning", 58),
       ("Judges disagreed?", 13),
       ("Reviewer", 12), ("Was this criterion relevant here?", 16),
       ("If not assessed: should it have applied?", 18),
       ("Judge A judgement sound?", 15), ("Judge B judgement sound?", 15),
       ("Whose judgement is closer to yours?", 16), ("Your comment", 56)]
FIRST_IN = 14

def review(wb, rows, have_reasoning):
    ws = wb.create_sheet("Criterion review")
    ws.freeze_panes = "F2"
    for i, (h, w) in enumerate(REV, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font, c.alignment, c.border = HDRFONT, TOPC, BOX
        c.fill = INFILL if i >= FIRST_IN else HDRFILL
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 46

    for r, row in enumerate(rows, 2):
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            c.font, c.alignment, c.border = BODY, TOP, BOX
        for col in (7, 8, 9, 11):
            cc = ws.cell(row=r, column=col)
            cc.number_format = "0.00"
            cc.alignment = TOPC
        f = ws.cell(row=r, column=13)
        f.value = (f'=IF(H{r}="","not assessed",'
                   f'IF(H{r}=0,"",'
                   f'IF(H{r}>=0.5,"YES  \u2014  wide",'
                   f'IF(H{r}>=0.25,"YES  \u2014  clear","yes  \u2014  slight"))))')
        f.alignment, f.font, f.border = TOPC, BOLD, BOX
        if row[7] is None or (row[7] or 0) > 0:
            f.fill = FLAG
        if not have_reasoning:
            for col in (10, 12):
                c = ws.cell(row=r, column=col, value="[to be filled from the run comparison file]")
                c.font, c.fill = NOTE, GAP
        for i in range(FIRST_IN, len(REV) + 1):
            ws.cell(row=r, column=i).fill = INPUT
        ws.row_dimensions[r].height = 60

    last = len(rows) + 1
    def dv(f, col):
        d = DataValidation(type="list", formula1=f, allow_blank=True, showDropDown=False)
        ws.add_data_validation(d)
        d.add(f"{col}2:{col}{last}")
    dv('"Yes,No,Unsure"', "O")
    dv('"Yes it should,No it was right to skip,Unsure,n/a"', "P")
    dv('"Agree,Partly agree,Disagree,Not assessed"', "Q")
    dv('"Agree,Partly agree,Disagree,Not assessed"', "R")
    dv('"Judge A,Judge B,Both equally,Neither"', "S")
    ws.auto_filter.ref = f"A1:{get_column_letter(len(REV))}{last}"
    return ws


def missing(wb):
    ws = wb.create_sheet("Missing criteria")
    heads = [("Reviewer", 14), ("Scenario", 30), ("Model (if specific)", 18),
             ("Criterion you would add", 60), ("Why it matters", 60),
             ("Core or context-specific?", 20)]
    for i, (h, w) in enumerate(heads, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font, c.fill, c.alignment, c.border = HDRFONT, HDRFILL, TOPC, BOX
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 32
    ex = ["[example]", "disclosure-direct", "all models",
          "Uses UK service terminology (A&E, 111, GP) rather than US equivalents",
          "A response directed a UK student to the 'nearest emergency room'. "
          "No existing criterion captured this.", "Core"]
    for i, v in enumerate(ex, 1):
        c = ws.cell(row=2, column=i, value=v)
        c.font, c.alignment, c.border = NOTE, TOP, BOX
    ws.row_dimensions[2].height = 44
    for r in range(3, 70):
        for i in range(1, 7):
            c = ws.cell(row=r, column=i)
            c.fill, c.border, c.alignment, c.font = INPUT, BOX, TOP, BODY
        ws.row_dimensions[r].height = 28
    d = DataValidation(type="list", formula1='"Core,Context-specific,Unsure"',
                       allow_blank=True, showDropDown=False)
    ws.add_data_validation(d)
    d.add("F3:F69")
    return ws


def summary(wb, scen):
    ws = wb.create_sheet("Coverage summary")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 34
    for col in "BCDE":
        ws.column_dimensions[col].width = 16
    ws.column_dimensions["F"].width = 62

    ws["A1"] = "How much work is this?"
    ws["A1"].font = TITLE
    ws["A2"] = ("Rows per scenario, so a workable subset can be agreed. Shaded rows are the "
                "safety-critical scenarios carrying practitioner weights.")
    ws["A2"].font = SMALL

    heads = ["Scenario", "Weight", "Rows to review", "Judges disagreed",
             "Reviewed so far", "What this scenario tests"]
    for i, h in enumerate(heads, 1):
        c = ws.cell(row=4, column=i, value=h)
        c.font, c.fill, c.alignment, c.border = HDRFONT, HDRFILL, TOPC, BOX
    ws.row_dimensions[4].height = 30

    r = 5
    for name, w, cite in scen:
        ws.cell(row=r, column=1, value=name).font = BODY
        ws.cell(row=r, column=2, value=w).font = BODY
        ws.cell(row=r, column=3, value=f"=COUNTIF('Criterion review'!$B:$B,$A{r})")
        ws.cell(row=r, column=4,
                value=f"=COUNTIFS('Criterion review'!$B:$B,$A{r},'Criterion review'!$M:$M,\"*yes*\")")
        ws.cell(row=r, column=5,
                value=f"=COUNTIFS('Criterion review'!$B:$B,$A{r},'Criterion review'!$N:$N,\"<>\")")
        ws.cell(row=r, column=6, value=cite).font = SMALL
        for i in range(1, 7):
            c = ws.cell(row=r, column=i)
            c.border = BOX
            c.alignment = TOPC if 2 <= i <= 5 else TOP
            if w != "\u00D71":
                c.fill = FLAG
        ws.row_dimensions[r].height = 40
        r += 1

    ws.cell(row=r, column=1, value="TOTAL").font = BOLD
    for col, L in ((3, "C"), (4, "D"), (5, "E")):
        c = ws.cell(row=r, column=col, value=f"=SUM({L}5:{L}{r-1})")
        c.font = BOLD
    for i in range(1, 7):
        c = ws.cell(row=r, column=i)
        c.fill, c.border = SUBFILL, BOX
        c.alignment = TOPC if 2 <= i <= 5 else TOP

    ws.cell(row=r + 2, column=1,
            value="Reviewers assess the judges' judgement, not the scores. "
                  "'Gap between judges' on the review sheet is the difference between the two "
                  "judges' scores for that criterion; 0 means they agreed exactly.").font = SMALL
    return ws


def main():
    bp = sys.argv[1] if len(sys.argv) > 1 else "sos-v03-stakeholder.yml"
    run = sys.argv[2] if len(sys.argv) > 2 else "run.json"
    resp_rows, rev_rows, scen, have_reasoning, _ = load(bp, run)
    print(f"{len(resp_rows)} responses, {len(rev_rows)} criterion rows, "
          f"reasoning present: {have_reasoning}")

    wb = Workbook(); wb.remove(wb.active)
    readme(wb, have_reasoning, len(resp_rows), len(rev_rows))
    responses(wb, resp_rows, have_reasoning)
    review(wb, rev_rows, have_reasoning)
    missing(wb)
    summary(wb, scen)
    wb.save(OUT)
    print("Written", OUT)


if __name__ == "__main__":
    main()
