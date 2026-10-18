"""Generate the fictional SAP HR internship datasets, visuals, and reports.

All records and calculations in this project are synthetic, reproducible
teaching examples. No SAP system is configured or accessed.
"""

from __future__ import annotations

import math
import random
import re
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Iterable

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUT_DATE = "18 October 2026"
AS_OF = date(2026, 9, 30)
SEED = 20261018
random.seed(SEED)

NAVY = "17365D"
BLUE = "2878A5"
TEAL = "2A9D8F"
GOLD = "E9C46A"
RED = "C84C4C"
PALE = "EAF1F7"
GRAY = "5B6573"
FONT = "Aptos"

AUTO_TABLES = {
    "Week_1_Workforce_Data_Analysis.docx": {
        "Data collection, synthetic cohort, and methodology": ["Department snapshot"],
        "Findings and interpretation": ["Demographic snapshot", "Recruitment source sample"],
    },
    "Week_2_Payroll_Configuration.docx": {
        "Prerequisites and conceptual configuration sequence": ["Conceptual payroll controls"],
        "Fictional error scenarios, corrective actions, and test strategy": ["Conceptual payroll test catalogue"],
    },
    "Week_3_Performance_Management.docx": {
        "Weighted scorecard and fictional calculations": ["Illustrative rating bands"],
    },
}

AUTO_FIGURES = {
    "Week_1_Workforce_Data_Analysis.docx": {
        "Findings and interpretation": ["charts/workforce_recruitment_source.png"],
    },
}

SUPPLEMENTARY_SECTIONS = {
    "Week_1_Workforce_Data_Analysis.docx": [{
        "title": "Decision use, data quality, and responsible follow-up",
        "body": """A production-ready workforce analysis would establish a governed data dictionary before building a dashboard. For each field, record the business meaning, source system, system-of-record owner, effective-date convention, allowed values, sensitivity, retention, and downstream transformations. Employee identity resolution needs an explicit rule for rehires, concurrent assignments, contingent workers, and transfers. A worker can hold more than one assignment, so a person count and an assignment count are different. For headcount, define treatment of unpaid leave, notice periods, part-time employees, and workers whose termination is effective on the snapshot date. Version that logic and test it against a set of anonymized known cases approved by HR and finance.\n\nThe current simulation uses a fixed seed and a closed cohort, but a real collection plan would be purpose-specific. A recruitment-funnel question needs requisitions and applicant-stage transitions, including rejected, withdrawn, offer, accepted, and start dates. Without those exposures, the employee recruitment-source chart only reports labels among people who appear in the staff file. A learning-effectiveness question requires learning objectives, pre/post evidence or workplace demonstration, and an appropriate comparison; training-hour totals alone show participation, not capability change. Absence analysis requires policy-approved absence categories, scheduled work exposure, and protections that prevent managers from inferring diagnoses. These requirements should be assessed before any new source is joined.\n\nData-quality checks should run at ingestion and refresh. Examples include valid date ordering, a unique and stable employee key, allowed department and location values, plausible employment intervals, missing required fields by process, and reconciliation to approved HR control totals. Investigate rather than silently coerce malformed dates or blanks to zero. Maintain an exception log with data owner, severity, resolution, and impact on published measures. Refreshes should be reproducible from a versioned script and immutable source extract or controlled query, with a recorded run date and formula version.\n\nDecision use should be proportionate to evidence. Aggregate trends can inform hiring capacity, skills investment, and scenario planning, but no single indicator should trigger an individual adverse action. A decision memo should explain the population, uncertainty, alternative explanations, privacy assessment, and owner for follow-up. Before publishing a small subgroup, assess re-identification risk and apply a documented threshold or broader aggregation. Invite employee-representative or privacy input when required by policy or law. Review whether each recurring report still has a valid purpose, remove fields and recipients no longer needed, and retire stale extracts safely.""",
    }],
    "Week_2_Payroll_Configuration.docx": [{
        "title": "Period controls, statutory change, and operational runbook",
        "body": """A payroll calendar is an operational control artifact. It should identify input cut-offs, time evaluation cut-offs if applicable, off-cycle rules, calculation windows, approval deadlines, pay dates, bank holidays, finance posting periods, and named owners. Payroll areas should group employees only where they share an approved periodic schedule and control process. A late input policy defines which changes can be processed in the current run, which require an off-cycle run, and which are handled as retroactive adjustments. The policy must be consistent with local law, employment terms, the payroll product, and finance close. Calendar exceptions need an authorised decision and documented employee communication.\n\nStatutory maintenance should be treated as controlled change, not manual guesswork. A legal change intake identifies the authority and effective date, jurisdiction and affected population, impacted wage types or tax parameters, reporting or remittance implications, configuration owner, reviewer, and evidence source. Configuration moves should be tested in a representative non-production environment and compared with independently calculated expected outcomes. Keep historic effective-dated parameters available for retroactive periods. Where a change arrives near a cut-off, record the approved treatment and quantify the impact; never obscure the decision by overwriting historical data.\n\nA practical runbook uses explicit stop/go criteria. Stop if employee population reconciliation fails, a critical interface file is incomplete, a payroll-area period is not in the expected status, high-severity exceptions remain unexplained, or the payment and accounting control totals do not balance. A named payroll lead reviews exception volume and material values; a separate approver confirms results and payment output. Any rerun has a unique identifier and documented reason, scope, before/after comparison, and approval. Restrict direct changes to payroll results; repair the originating source or approved configuration and recompute through the normal engine.\n\nReconciliation should bridge gross earnings, employer and employee deductions, net pay, payment file totals, rejected payments, clearing accounts, payroll liabilities, and expense postings. Differences may reflect timing, rounding, or specific legal and finance treatments, but each must have an owner and explanation. The close package should retain period status, input and output totals, error resolution, approval evidence, payment confirmation, posting document references, and unresolved items. Access to evidence is restricted and retention follows the applicable schedule. These control expectations remain conceptual; actual transactions and report names differ by deployment and release.""",
    }],
    "Week_3_Performance_Management.docx": [{
        "title": "Goal quality, development evidence, and process evaluation",
        "body": """A useful SMART objective is not merely a numeric target. Specificity states the outcome and scope; measurability states the evidence and baseline; achievability considers resources and dependencies; relevance links the outcome to a team or enterprise priority; and time-bound means a clear end date or review interval. A KPI supports monitoring but can distort behaviour when used without a quality, risk, or customer guardrail. For example, a target on cases closed should be balanced with quality and reopened-case evidence. Employees should understand which measures they can influence and how exceptional external factors will be treated.\n\nOKRs usually work best as a transparent alignment and learning mechanism with a short cycle. An objective describes a meaningful direction; key results are measurable outcomes that indicate progress. A key result should avoid becoming a list of tasks. OKRs do not automatically translate into individual ratings or compensation. If an organisation chooses to connect them to an appraisal, explain that relationship, prevent double counting, and account for shared results and dependencies. A person may contribute to a key result without being solely accountable for it. Teams should review progress and adjust assumptions rather than retroactively redefining success to match an outcome.\n\nCompetency assessment needs behavioural anchors appropriate to job level and opportunity. Evidence might include a documented customer outcome, a peer collaboration example, a quality review, or observed decision-making in a stretch assignment. Avoid inferring competencies from confidence, visibility, or similarity to the assessor. Employees need an opportunity to provide context and identify structural barriers. Where a role changed materially during the cycle, preserve the original goal history and agree a fair replacement instead of rating an obsolete target as if circumstances were unchanged.\n\nDevelopment planning should be jointly owned. Convert one priority into a practice-based commitment—for example, lead a scoped project review, receive a mentor's feedback on a defined skill, or demonstrate a new tool capability on a suitable work task. Identify budget, workload, accessibility, manager support, and a measurable demonstration date. Review progress at the next check-in and revise the plan when the opportunity is unavailable. Learning content completion can be recorded, but the effectiveness measure should examine whether the employee can apply the skill in the job.\n\nAfter each cycle, evaluate completion and quality together: objective clarity, on-time conversation, employee understanding, manager training coverage, appeal and correction themes, perceived fairness, development follow-through, and differences in opportunity. Treat group comparisons carefully, protect small cells, and investigate rather than assume a causal explanation. Use findings to improve guidance and manager support. A pilot should include representative roles and accessibility needs, a support channel, change ownership, a clear data-retention rule, and a decision gate for whether to continue or redesign.""",
    }],
    "Week_4_Succession_Planning.docx": [{
        "title": "Role-specific roadmaps and continuity scenarios",
        "body": """The matrix becomes actionable when each critical role has a concise, current role profile. Include purpose, key decisions, stakeholder relationships, essential capabilities, mandatory credentials, delegated authority, control conflicts, knowledge artifacts, and continuity impact. Review criticality when strategy, regulation, operating model, or technology changes. Avoid marking every leadership job as critical by default: too broad a population dilutes resources and makes emergency readiness difficult to maintain. Include specialist and operational roles where expertise or control ownership is difficult to replace, not only senior executives.\n\nA fictional roadmap for the payroll operations lead can sequence three evidence-building experiences: reconcile a controlled payroll close with finance, lead a simulated statutory change assessment, and coordinate a response exercise for a rejected payment file. Readiness review would consider accuracy, control discipline, stakeholder communication, and ability to lead under time pressure. For the HRIS product owner, a scoped release should demonstrate requirements prioritisation, access/privacy review, test evidence, cutover planning, and measurable adoption. For a regional operations director, a temporary project or rotation can build budget ownership and cross-site leadership without prematurely exposing the candidate to unsupported risk. Each assignment needs a sponsor, a mentor or manager, an end date, and a feedback record.\n\nContingency planning must work even when no permanent successor is ready. Identify the authorised interim decision-maker, the duration and boundaries of delegated authority, required dual controls, sensitive system roles, essential supplier contacts, and where the approved procedure is stored. A tabletop scenario can simulate a sudden vacancy during payroll cut-off or an incident affecting the data protection lead. Participants test call trees, access requests, approval authority, stakeholder communication, and recovery objectives. Gaps should become time-bound actions with accountable owners. An interim cover plan is not a promise that an employee will be promoted.\n\nReadiness validation should use evidence gathered since the previous review and confirm candidate interest, availability, and development progress through appropriate channels. A label must not be updated solely to improve a KPI. If a candidate leaves, is unavailable, or declines a move, promptly revise the plan and provide a fair career conversation. Track whether development experiences are equitably distributed, while preserving candidate confidentiality and respecting local privacy and employment rules. For small talent cohorts, individual-level reporting should be restricted to the authorised review group and not included in broadly accessible dashboards.\n\nA robust review cycle asks leaders to challenge assumptions: what has changed in the role, what evidence supports readiness, what critical capability remains uncovered, what dependency would block an interim, and what alternative exists if the primary candidate exits? Review both bench depth and role resilience. Where internal readiness cannot be developed in time, consider knowledge transfer, external hiring, service continuity, or role redesign rather than overstating a successor's preparedness.""",
    }],
    "Week_5_HR_Data_Security.docx": [{
        "title": "Privacy-by-design lifecycle and control assurance",
        "body": """Privacy-by-design begins before a system field, report, or integration is introduced. Document the purpose, data subjects, categories, sources, recipients, locations, retention, access, and legal basis or other permitted ground applicable to the processing. Collect only data necessary for that purpose and give clear notices in a language and channel employees can understand. Design access so that HR operations, line management, payroll processing, analytics, and security investigation see different minimum subsets. A role that permits viewing an employee record should not automatically permit bulk export, compensation change, or bank-detail maintenance.\n\nA data-flow inventory should include source and destination systems, file and API interfaces, service accounts, subprocessors, analytics copies, backups, and deletion paths. For each flow, record authentication, encryption, integrity checks, error handling, monitoring, data owner, and retention. Test access changes across the full identity lifecycle: joiner provisioning, role or location change, temporary assignment, leave, termination, and emergency access. Coordinate HR events with identity governance so that delayed feeds or failed deprovisioning are visible and escalated. Validate cloud role target populations and classic organizational authorization fields with positive and negative tests after every material change.\n\nControl assurance requires more than a policy document. Sample privileged access quarterly or at a risk-based interval, reconcile the list to actual accounts, inspect exceptions, test that logs reach a protected central store, and replay a restore exercise. Monitor supplier notifications, security advisories, vulnerabilities, and contract changes. For encryption, verify the actual coverage of application traffic, databases, backups, and exported files as well as key custody and rotation. For deletion, test live records and understand how backup expiry, legal hold, and immutable logs are handled. Document residual risks accepted by a named, authorised owner with an expiry or review date.\n\nAn incident scenario involving payroll diversion should verify that a bank detail change generates a notification, independent verification uses a known channel rather than the submitted contact, payment release remains segregated, and the employee can report a concern. A privacy disclosure scenario should test containment, log preservation, identification of affected records and recipients, legal review, notification decision authority, and recovery. The playbook must avoid emailing exposed data as evidence. Exercise outcomes should result in assigned corrective actions and retest.\n\nLegal review is an operational dependency. Maintain a register of relevant jurisdictions, entity roles, processing purposes, contracts, notices, retention rules, transfer mechanisms, and regulatory dates. For India, check the official Gazette for the DPDP Rules and commencement notifications in force for the relevant processing and obligation. For GDPR, confirm territorial scope and controller/processor responsibilities. HIPAA analysis depends on covered-entity or business-associate context; employer status alone is not enough. Update security and privacy controls when law, product architecture, or data use changes, and seek qualified counsel where application is uncertain.""",
    }],
}


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paragraph.add_run("Page ")
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    paragraph._p.append(field)


def set_cell_text(cell, value: object, bold: bool = False, color: str | None = None) -> None:
    cell.text = str(value)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = FONT
            run.font.size = Pt(8.3)
            run.bold = bold
            if color:
                run.font.color.rgb = RGBColor.from_string(color)


def add_table(doc: Document, headers: list[str], rows: Iterable[Iterable[object]], caption: str) -> None:
    cap = doc.add_paragraph(f"Table {len(doc.tables) + 1}. {caption}", style="Caption")
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Light Shading Accent 1"
    for i, header in enumerate(headers):
        set_cell_text(table.rows[0].cells[i], header, True, "FFFFFF")
        shade_cell(table.rows[0].cells[i], NAVY)
    set_repeat_table_header(table.rows[0])
    for values in rows:
        cells = table.add_row().cells
        for i, value in enumerate(values):
            set_cell_text(cells[i], value)
    doc.add_paragraph()


def add_bullet(doc: Document, text: str) -> None:
    p = doc.add_paragraph(style="List Bullet")
    p.add_run(text)


def add_numbered(doc: Document, text: str) -> None:
    p = doc.add_paragraph(style="List Number")
    p.add_run(text)


def add_figure(doc: Document, image: Path, caption: str, width: float = 6.0) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(image), width=Inches(width))
    cap = doc.add_paragraph(caption, style="Caption")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER


def paragraph_text(doc: Document, text: str) -> None:
    for piece in text.split("\n\n"):
        if piece.strip():
            doc.add_paragraph(piece.strip())


def make_dataset() -> pd.DataFrame:
    departments = ["People", "Finance", "Technology", "Sales", "Operations", "Customer Support"]
    locations = ["Bengaluru", "Hyderabad", "Pune", "Chennai", "Gurugram"]
    jobs = {
        "People": ["HR Analyst", "Recruiter", "HR Business Partner"],
        "Finance": ["Payroll Analyst", "Accountant", "Finance Manager"],
        "Technology": ["Software Engineer", "Data Analyst", "IT Lead"],
        "Sales": ["Sales Associate", "Account Executive", "Sales Manager"],
        "Operations": ["Operations Analyst", "Team Lead", "Process Specialist"],
        "Customer Support": ["Support Associate", "Service Lead", "Quality Analyst"],
    }
    sources = ["Employee referral", "Career site", "Campus", "Agency", "Internal move"]
    records = []
    for n in range(1, 121):
        department = random.choice(departments)
        start = date(2019, 1, 1) + timedelta(days=random.randint(0, (AS_OF - date(2019, 1, 1)).days))
        ended = random.random() < 0.17 and start < date(2026, 1, 1)
        termination = None
        if ended:
            termination = start + timedelta(days=random.randint(100, min(1800, (AS_OF - start).days)))
            if termination > AS_OF:
                termination = None
        gender = random.choices(["Woman", "Man", "Non-binary", "Undisclosed"], [46, 48, 2, 4])[0]
        age = random.randint(22, 60)
        records.append({
            "employee_id": f"EMP{n:04d}",
            "department": department,
            "job_title": random.choice(jobs[department]),
            "location": random.choice(locations),
            "gender": gender,
            "age": age,
            "hire_date": start.isoformat(),
            "termination_date": termination.isoformat() if termination else "",
            "employment_status": "Terminated" if termination else "Active",
            "annual_salary_inr": random.randrange(420000, 2400000, 10000),
            "recruitment_source": random.choice(sources),
            "absence_days_fy2025": random.choices(range(0, 13), [20, 24, 20, 14, 8, 5, 3, 2, 1, 1, 1, 1, 1])[0],
            "training_hours_fy2025": random.randint(0, 56),
            "performance_rating_fy2025": round(random.uniform(2.2, 4.9), 1),
            "as_of_date": AS_OF.isoformat(),
        })
    frame = pd.DataFrame(records)
    frame.to_csv(ROOT / "datasets" / "synthetic_employee_workforce.csv", index=False)

    payroll = pd.DataFrame([
        ["PAY001", "Regular salary", 85000, 0, 0, "Illustrative clean run"],
        ["PAY002", "Regular salary", 72000, 0, 0, "Missing bank master-data record"],
        ["PAY003", "Regular salary", 96000, 1, 0, "Payroll area period not released"],
        ["PAY004", "Regular salary", 68000, 0, 1, "Duplicate recurring allowance"],
        ["PAY005", "Regular salary", 110000, 0, 0, "Retroactive job change requires recalculation"],
        ["PAY006", "Regular salary", 54000, 0, 0, "Incorrect tax declaration / withholding setup"],
    ], columns=["case_id", "pay_component", "illustrative_gross_inr", "blocked_control_record", "duplicate_recurring_item", "scenario"])
    payroll.to_csv(ROOT / "datasets" / "synthetic_payroll_scenarios.csv", index=False)

    performance = frame.loc[frame.employment_status == "Active", ["employee_id", "department", "performance_rating_fy2025", "training_hours_fy2025"]].copy()
    performance["objective_score"] = (performance.performance_rating_fy2025 * 20).round(0).astype(int)
    performance["competency_score"] = (performance.performance_rating_fy2025 * 20 + performance.training_hours_fy2025 / 5).clip(0, 100).round(0).astype(int)
    performance["values_score"] = (performance.performance_rating_fy2025 * 20 - 2).clip(0, 100).round(0).astype(int)
    performance["weighted_score"] = (performance.objective_score * .60 + performance.competency_score * .25 + performance.values_score * .15).round(1)
    performance.to_csv(ROOT / "datasets" / "synthetic_performance_scores.csv", index=False)

    succession = pd.DataFrame([
        ["Chief People Officer", "CPO-01", "Candidate A", "Ready now", "Executive stakeholder influence", "Board simulation; sponsor mentoring"],
        ["Payroll Operations Lead", "PAY-04", "Candidate B", "Ready in 1-2 years", "Multi-country payroll governance", "Cross-functional payroll rotation"],
        ["HRIS Product Owner", "HRIS-02", "Candidate C", "Ready in 1-2 years", "Roadmap prioritisation", "Lead a scoped HRIS release"],
        ["Regional Operations Director", "OPS-07", "Candidate D", "Ready in 3-5 years", "P&L and strategy experience", "Regional assignment; leadership mentor"],
        ["Head of Talent Acquisition", "TA-03", "Candidate E", "Emergency cover only", "Strategic workforce planning", "Interim delegation drill; capability plan"],
        ["Data Protection Officer", "DPO-01", "Candidate F", "No identified successor", "Independent privacy expertise", "External interim cover; develop internal bench"],
    ], columns=["critical_role", "role_code", "fictional_candidate", "readiness", "gap", "development_or_contingency"])
    succession.to_csv(ROOT / "datasets" / "synthetic_succession_matrix.csv", index=False)
    return frame


def active_at(frame: pd.DataFrame, when: date) -> int:
    hires = pd.to_datetime(frame.hire_date).dt.date
    terms = pd.to_datetime(frame.termination_date.replace("", pd.NA), errors="coerce").dt.date
    return int(((hires <= when) & (terms.isna() | (terms >= when))).sum())


def calculate_metrics(frame: pd.DataFrame) -> dict:
    active = frame[frame.employment_status == "Active"].copy()
    leavers_2025 = frame[pd.to_datetime(frame.termination_date, errors="coerce").dt.year == 2025]
    opening_2025 = active_at(frame, date(2025, 1, 1))
    closing_2025 = active_at(frame, date(2025, 12, 31))
    average_2025 = (opening_2025 + closing_2025) / 2
    attrition = 100 * len(leavers_2025) / average_2025 if average_2025 else 0
    return {
        "headcount": len(active),
        "employees": len(frame),
        "departments": active.department.nunique(),
        "mean_age": active.age.mean(),
        "women_share": 100 * (active.gender == "Woman").sum() / len(active),
        "median_salary": active.annual_salary_inr.median(),
        "attrition_2025": attrition,
        "leavers_2025": len(leavers_2025),
        "opening_2025": opening_2025,
        "closing_2025": closing_2025,
        "absence_mean": active.absence_days_fy2025.mean(),
        "training_mean": active.training_hours_fy2025.mean(),
        "rating_mean": active.performance_rating_fy2025.mean(),
        "new_hires_2025": int((pd.to_datetime(frame.hire_date).dt.year == 2025).sum()),
        "active": active,
    }


def chart_workforce(frame: pd.DataFrame, metrics: dict) -> None:
    active = metrics["active"]
    plt.figure(figsize=(8, 4.4))
    active.groupby("department").size().sort_values().plot(kind="barh", color=f"#{BLUE}")
    plt.title("Active headcount by department (synthetic, 30 Sep 2026)")
    plt.xlabel("Employees"); plt.ylabel("")
    plt.tight_layout()
    plt.savefig(ROOT / "charts" / "workforce_headcount_by_department.png", dpi=180)
    plt.close()

    quarters = pd.date_range("2023-12-31", "2026-09-30", freq="QE")
    values = [active_at(frame, d.date()) for d in quarters]
    plt.figure(figsize=(8, 4.4))
    plt.plot(quarters, values, marker="o", color=f"#{TEAL}", linewidth=2.4)
    plt.title("Quarter-end headcount trajectory (synthetic)")
    plt.ylabel("Employees"); plt.xlabel("Quarter-end")
    plt.grid(axis="y", alpha=.25); plt.tight_layout()
    plt.savefig(ROOT / "charts" / "workforce_headcount_trend.png", dpi=180)
    plt.close()

    by_source = frame.groupby("recruitment_source").size().sort_values()
    plt.figure(figsize=(8, 4.2))
    by_source.plot(kind="bar", color=f"#{GOLD}")
    plt.title("Hiring source represented in employee sample")
    plt.ylabel("Records"); plt.xlabel("")
    plt.xticks(rotation=20, ha="right"); plt.tight_layout()
    plt.savefig(ROOT / "charts" / "workforce_recruitment_source.png", dpi=180)
    plt.close()


def draw_flow(path: Path, title: str, steps: list[str]) -> None:
    fig, ax = plt.subplots(figsize=(9, max(2.6, len(steps) * .72)))
    ax.set_xlim(0, 10); ax.set_ylim(0, len(steps) + .2); ax.axis("off")
    yvals = list(reversed(range(len(steps))))
    for idx, (y, text) in enumerate(zip(yvals, steps)):
        ax.text(5, y + .45, text, ha="center", va="center", color=f"#{NAVY}", fontsize=9,
                bbox=dict(boxstyle="round,pad=.55", fc="#EAF1F7", ec=f"#{BLUE}", lw=1.3))
        if idx < len(steps) - 1:
            ax.annotate("", xy=(5, yvals[idx + 1] + .9), xytext=(5, y + .05),
                        arrowprops=dict(arrowstyle="->", color=f"#{TEAL}", lw=1.5))
    ax.set_title(title, color=f"#{NAVY}", fontsize=14, weight="bold", pad=12)
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def draw_succession_visuals() -> None:
    levels = ["Developing", "Effective", "Exceptional"]
    values = [[10, 5, 2], [8, 16, 7], [3, 9, 6]]
    fig, ax = plt.subplots(figsize=(6.5, 5))
    img = ax.imshow(values, cmap="YlGnBu")
    ax.set_xticks(range(3), levels); ax.set_yticks(range(3), levels)
    ax.set_xlabel("Potential →"); ax.set_ylabel("Performance →")
    ax.set_title("Illustrative nine-box calibration (fictional cohort)")
    for i in range(3):
        for j in range(3):
            ax.text(j, i, str(values[i][j]), ha="center", va="center", color="black", weight="bold")
    fig.colorbar(img, ax=ax, label="Employees")
    fig.tight_layout()
    fig.savefig(ROOT / "charts" / "succession_nine_box.png", dpi=180)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(9, 3.8))
    ax.axis("off")
    labels = [("Identify", "Role criticality\nand vacancy impact"),
              ("Review", "Evidence-led\ncalibration"),
              ("Develop", "Experience, coaching,\nlearning"),
              ("Validate", "Readiness evidence\nand retention"),
              ("Activate", "Contingency cover\nand refresh")]
    xs = [0.1, 0.3, 0.5, 0.7, 0.9]
    for i, (head, body) in enumerate(labels):
        ax.text(xs[i], .5, f"{head}\n\n{body}", transform=ax.transAxes, ha="center", va="center",
            fontsize=9, bbox=dict(boxstyle="round,pad=.7", fc="#EAF1F7", ec=f"#{BLUE}"))
        if i < len(labels)-1:
            ax.annotate("", xy=(xs[i+1]-.075, .5), xytext=(xs[i]+.075, .5),
                        xycoords=ax.transAxes, textcoords=ax.transAxes,
                        arrowprops=dict(arrowstyle="->", color=f"#{TEAL}", lw=1.6))
    ax.set_title("Succession governance cycle", color=f"#{NAVY}", fontsize=14, weight="bold")
    fig.tight_layout()
    fig.savefig(ROOT / "diagrams" / "succession_governance_cycle.png", dpi=180, bbox_inches="tight")
    plt.close(fig)


def draw_security_matrix() -> None:
    matrix = [[1, 2, 3, 4, 5], [2, 4, 6, 8, 10], [3, 6, 9, 12, 15],
              [4, 8, 12, 16, 20], [5, 10, 15, 20, 25]]
    fig, ax = plt.subplots(figsize=(6.3, 5))
    ax.imshow(matrix, cmap="YlOrRd", origin="lower", vmin=1, vmax=25)
    ax.set_xticks(range(5), ["1 Rare", "2 Unlikely", "3 Possible", "4 Likely", "5 Almost certain"], rotation=25, ha="right")
    ax.set_yticks(range(5), ["1 Negligible", "2 Minor", "3 Moderate", "4 Major", "5 Severe"])
    ax.set_xlabel("Likelihood"); ax.set_ylabel("Impact")
    ax.set_title("Illustrative inherent likelihood × impact scoring")
    for y in range(5):
        for x in range(5):
            ax.text(x, y, str(matrix[y][x]), ha="center", va="center", fontsize=9)
    fig.tight_layout()
    fig.savefig(ROOT / "charts" / "hr_security_risk_matrix.png", dpi=180)
    plt.close(fig)


def set_document_style(doc: Document) -> None:
    styles = doc.styles
    styles["Normal"].font.name = FONT
    styles["Normal"].font.size = Pt(10)
    styles["Normal"].font.color.rgb = RGBColor.from_string("233142")
    styles["Normal"].paragraph_format.space_after = Pt(6)
    for name, size, color in [("Title", 28, NAVY), ("Heading 1", 17, NAVY),
                              ("Heading 2", 12, BLUE), ("Heading 3", 10.5, TEAL)]:
        style = styles[name]
        style.font.name = FONT; style.font.size = Pt(size)
        style.font.bold = True; style.font.color.rgb = RGBColor.from_string(color)
    styles["Caption"].font.name = FONT
    styles["Caption"].font.size = Pt(8)
    styles["Caption"].font.italic = True
    for section in doc.sections:
        section.top_margin = Inches(.72)
        section.bottom_margin = Inches(.7)
        section.left_margin = Inches(.78)
        section.right_margin = Inches(.78)
        footer = section.footer.paragraphs[0]
        footer.add_run("SAP HR Internship | Fictional academic simulation     ")
        add_page_number(footer)


def cover(doc: Document, week: int, title: str, subtitle: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(85)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"WEEK {week}")
    run.bold = True; run.font.size = Pt(15); run.font.color.rgb = RGBColor.from_string(TEAL)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(16)
    run = p.add_run(title)
    run.bold = True; run.font.name = FONT; run.font.size = Pt(27); run.font.color.rgb = RGBColor.from_string(NAVY)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    run = p.add_run(subtitle)
    run.font.size = Pt(14); run.font.color.rgb = RGBColor.from_string(GRAY)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(60)
    p.add_run("Academic project report | Prepared " + OUT_DATE).font.size = Pt(11)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(24)
    r = p.add_run("All people, organisations, measures, and outcomes in the worked examples are fictional.\n")
    r.bold = True
    p.add_run("Conceptual learning exercise only; no live SAP system was configured.")
    doc.add_page_break()


def create_report(week: int, title: str, subtitle: str, summary: str, sections: list[dict],
                  table_specs: list[tuple[str, list[str], list[list[object]]]],
                  figure_specs: list[tuple[str, str]], references: list[str], out_file: str) -> None:
    sections = [*sections, *SUPPLEMENTARY_SECTIONS.get(out_file, [])]
    doc = Document()
    set_document_style(doc)
    cover(doc, week, title, subtitle)
    doc.add_heading("Executive Summary", level=1)
    paragraph_text(doc, summary)
    doc.add_heading("Contents", level=1)
    for i, section in enumerate(sections, 1):
        doc.add_paragraph(f"{i}. {section['title']}")
    doc.add_paragraph("Tables, figures, and references follow the relevant discussion.")
    doc.add_page_break()
    for i, section in enumerate(sections, 1):
        doc.add_heading(f"{i}. {section['title']}", level=1)
        paragraph_text(doc, section["body"])
        if section.get("bullets"):
            for item in section["bullets"]:
                add_bullet(doc, item)
        if section.get("steps"):
            for item in section["steps"]:
                add_numbered(doc, item)
        table_names = [*section.get("tables", []), *AUTO_TABLES.get(out_file, {}).get(section["title"], [])]
        if section.get("table"):
            table_names = [section["table"], *table_names]
        for name in table_names:
            spec = next(x for x in table_specs if x[0] == name)
            doc.add_heading(name, level=2)
            add_table(doc, spec[1], spec[2], name)
        for image, caption in figure_specs:
            if section.get("figure") == image:
                add_figure(doc, ROOT / image, caption)
        for image in AUTO_FIGURES.get(out_file, {}).get(section["title"], []):
            caption = next(caption for path, caption in figure_specs if path == image)
            add_figure(doc, ROOT / image, caption)
    doc.add_heading(f"{len(sections)+1}. References", level=1)
    doc.add_paragraph("Official source portals were checked on 18 October 2026 where automated access was available; some government legal pages restrict automated retrieval. Product details and applicable legal obligations should be confirmed against the current official documentation and the organisation's jurisdiction before implementation.")
    for i, ref in enumerate(references, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(.2)
        p.paragraph_format.first_line_indent = Inches(-.2)
        p.add_run(f"[{i}] {ref}")
    doc.core_properties.title = f"Week {week}: {title}"
    doc.core_properties.subject = "SAP HR internship academic simulation"
    doc.core_properties.author = "SAP HR Internship Project"
    doc.save(ROOT / "reports" / out_file)


SAP_REFS = [
    "SAP Help Portal, SAP SuccessFactors Employee Central documentation landing area: https://help.sap.com/docs/SAP_SUCCESSFACTORS_EMPLOYEE_CENTRAL",
    "SAP Help Portal, SAP SuccessFactors Platform documentation landing area: https://help.sap.com/docs/SAP_SUCCESSFACTORS_PLATFORM",
    "SAP Help Portal, SAP ERP documentation landing area (including Human Capital Management topics): https://help.sap.com/docs/SAP_ERP",
]
SECURITY_REFS = [
    "European Parliament and Council, Regulation (EU) 2016/679 (General Data Protection Regulation), official text: https://eur-lex.europa.eu/eli/reg/2016/679/oj",
    "Government of India, Ministry of Electronics and Information Technology, Digital Personal Data Protection Act / data protection framework: https://www.meity.gov.in/data-protection-framework",
    "Government of India, e-Gazette of India, authoritative Gazette publication portal (verify applicable DPDP Rules and commencement notifications): https://egazette.gov.in/",
    "NIST, SP 800-53 Rev. 5, Security and Privacy Controls for Information Systems and Organizations: https://csrc.nist.gov/pubs/sp/800/53/r5/upd1/final",
    "NIST, Cybersecurity Framework 2.0: https://www.nist.gov/cyberframework",
    "ISO/IEC 27001:2022, Information security management systems — Requirements (standard overview): https://www.iso.org/standard/27001",
    "U.S. Department of Health and Human Services, HIPAA for Professionals: https://www.hhs.gov/hipaa/for-professionals/index.html",
]


def build_reports(frame: pd.DataFrame, metrics: dict) -> None:
    dept_counts = metrics["active"].department.value_counts().sort_values(ascending=False)
    dept_rows = [[dept, int(n), f"{100*n/metrics['headcount']:.1f}%"] for dept, n in dept_counts.items()]
    demographic_rows = []
    for key, count in metrics["active"].gender.value_counts().items():
        demographic_rows.append([key, int(count), f"{100*count/metrics['headcount']:.1f}%"])
    source_rows = [[str(k), int(v)] for k, v in frame.recruitment_source.value_counts().items()]
    kpi_rows = [
        ["Active headcount", "Count active at 30 Sep 2026", metrics["headcount"], "Snapshot; not a time series by itself"],
        ["2025 attrition", "2025 leavers ÷ average 2025 headcount × 100", f"{metrics['attrition_2025']:.1f}%", "Small, simulated population; annualised rate is volatile"],
        ["Absence days", "Mean FY2025 absence days per active person", f"{metrics['absence_mean']:.1f}", "No exposure-adjusted rate or leave-type split"],
        ["Training hours", "Mean FY2025 recorded hours per active person", f"{metrics['training_mean']:.1f}", "Completion is not evidence of skill transfer"],
        ["Performance", "Mean FY2025 rating on 1–5 scale", f"{metrics['rating_mean']:.2f}", "Ratings need calibration and context"],
        ["Women representation", "Women ÷ active population × 100", f"{metrics['women_share']:.1f}%", "Categories are self-identification fields in mock data"],
    ]
    workforce_tables = [
        ("Department snapshot", ["Department", "Active headcount", "Share"], dept_rows),
        ("Demographic snapshot", ["Recorded category", "Employees", "Share"], demographic_rows),
        ("KPI definitions and sample results", ["Measure", "Formula / definition", "Synthetic value", "Interpretation caveat"], kpi_rows),
        ("Recruitment source sample", ["Recorded source", "Records"], source_rows),
    ]
    create_report(
        1, "Workforce Data Analysis", "People analytics foundations and workforce planning",
        "This report presents a reproducible workforce analytics exercise using 120 fictional employee records. It explains the distinct roles of SAP HCM, Personnel Administration (PA), Organizational Management (OM), SAP SuccessFactors Employee Central, and analytics tooling. A 30 September 2026 snapshot and FY2025 indicators are calculated from synthetic hires, exits, demographics, department, attendance, learning, and ratings. The sample produces a headcount of {hc}, mean recorded absence of {abs:.1f} days, mean training of {tr:.1f} hours, and mean performance rating {rating:.2f}/5. These figures describe only this generated dataset; they are not benchmarks or evidence about a real employer. The report defines measures, describes limitations and governance, and translates the patterns into responsible workforce-planning actions.".format(hc=metrics["headcount"], abs=metrics["absence_mean"], tr=metrics["training_mean"], rating=metrics["rating_mean"]),
        [
            {"title":"Scope, SAP HR landscape, and analytical question","body":"Workforce analytics combines reliable people data with a clearly defined decision question. SAP HCM is the established human-capital-management application family used in SAP ERP environments. Its Personnel Administration (PA) functions maintain employee master data and employment actions; Organizational Management (OM) represents organizational units, jobs, positions, reporting relationships, and validity periods. These are related but distinct structures: a person assignment is not identical to an organisation chart object. Exact component scope depends on the licensed release and implementation.\n\nSAP SuccessFactors Employee Central is SAP's cloud core-HR system of record for employee and organisational data, with effective-dated records and configurable business processes. It is not simply a renamed on-premise PA/OM screen. Customers may use Employee Central with other SuccessFactors products, integrate it with SAP payroll or other enterprise applications, or operate a hybrid landscape. SAP ERP HCM and HCM for SAP S/4HANA deployment options must be evaluated against SAP's current product strategy, release, maintenance, licensing, and customer-specific architecture; moving an ERP database to S/4HANA does not by itself establish which HR application is in use.\n\nAnalytics may be supplied through operational reports, embedded product reporting, SAP Analytics Cloud or a data platform depending on the architecture and entitlements. The analytical layer should not become an uncontrolled shadow master. This exercise asks: what does a small, de-identified, point-in-time synthetic cohort suggest about workforce composition and indicators that merit follow-up?" ,"bullets":["Use a documented business definition and reference date for every KPI.","Keep operational correction workflows separate from aggregate decision reporting.","Use access controls and aggregation thresholds for sensitive or small demographic groups."]},
            {"title":"Data collection, synthetic cohort, and methodology","body":"The reproducible generator creates 120 fictional records using a fixed random seed. Each record includes a pseudonymous employee identifier, department, job family title, location, a coarse demographic category, age, hire and possible termination dates, salary, recruitment source, FY2025 absence days, training hours, and performance rating. All dates are generated between 2019 and 30 September 2026. The file is explicitly synthetic and must not be combined with production employee data. No person-level inference is appropriate.\n\nThe point-in-time headcount counts records hired on or before the reference date and not terminated before that date. FY2025 attrition is defined as FY2025 terminations divided by the arithmetic mean of headcount at 1 January and 31 December, multiplied by 100. This simplified convention is transparent but not universal: organizations may define voluntary, regrettable, internal, or rolling attrition differently. Absence days and training hours are arithmetic means among employees active at the snapshot, not rates per scheduled workday or exposure. Performance ratings are a 1–5 ordinal scale; averaging them is an exploratory summary, not proof of equal distances between rating levels.\n\nThe analysis uses pandas to produce counts, proportions, means, median salary, turnover estimate, quarterly headcount snapshots, and graphics. Missingness is represented only where explicitly generated; the compact simulation has no realistic data-entry errors, duplicate identifiers, leaves, part-time fractions, transfers, or changing job assignments. Real collection would require purpose limitation, lawful basis, notice, minimisation, retention rules, access controls, data-quality stewardship, and reconciliation to source systems."},
            {"title":"Findings and interpretation","body":f"The generated sample contains {metrics['employees']} total employee records and {metrics['headcount']} active records at the reference date across {metrics['departments']} departments. Active cohort mean age is {metrics['mean_age']:.1f} years; median annual salary is INR {metrics['median_salary']:,.0f}; women represent {metrics['women_share']:.1f}% of the active sample. The distributions in the accompanying tables and charts describe this seed and should not be interpreted as population estimates. Department size differences can reflect the sampling method, not business demand or productivity.\n\nThe 2025 attrition estimate is {metrics['attrition_2025']:.1f}% ({metrics['leavers_2025']} generated departures; opening and closing headcounts {metrics['opening_2025']} and {metrics['closing_2025']}). A small numerator and denominator make the result unstable. A chart of headcount over time can show net movement only; it does not distinguish hiring, exits, transfers, or changes in FTE. Recruitment-source counts show the source values attached to the whole employee file and are not a valid channel conversion or hiring-efficiency measure because failed applicants, time-to-fill, costs, and offer acceptance are absent.\n\nCurrent-cohort mean recorded FY2025 absence is {metrics['absence_mean']:.1f} days, training is {metrics['training_mean']:.1f} hours, and performance is {metrics['rating_mean']:.2f}/5. These aggregates cannot establish whether workload, health, manager practices, programme quality, or rating bias explains variation. Averages conceal distributions, tenure, role mix, leave eligibility, seasonality, and uncertainty. Statistical comparisons are deliberately not presented as causal effects."},
            {"title":"KPI catalogue and analytical controls","body":"The KPI table states the formulas and caveats. Headcount is a count at a specified date; FTE is a separate measure requiring contracted hours and work schedules. Representation should be reported by a carefully defined population and category, alongside privacy safeguards and intersectional suppression rules. A useful absence-rate numerator may be absence hours while the denominator is scheduled hours, but approved leave and policy context determine the definition. Training measures should separate assigned, started, completed, and competency-demonstrated learning.\n\nFor robust operational reporting, version metric definitions, effective dates, source lineage, population filters, refresh cadence, owner, quality checks, and permitted use. Reconcile counts between HR, payroll, and analytics after agreed treatment of contingent workers and concurrent employments. Test validity, uniqueness, completeness, timeliness, and referential integrity. Compare aggregates with source control totals; document exceptions rather than silently substituting zero. Preserve only fields needed for the approved question and restrict row-level exports."},
            {"title":"Workforce planning recommendations and limitations","body":"Start with decision-specific follow-up, not individual ranking. Validate critical roles and forecast demand with business leaders; compare required capabilities to current capacity; monitor recruitment funnel stages and realistic attrition cohorts; and scenario-plan capacity, skills, and internal mobility. Establish monthly headcount and vacancy reconciliation, quarterly workforce reviews, and a documented action owner for each material gap. For absence, examine policy-compliant aggregate patterns with occupational-health and employee-relations context; never infer an individual's health condition from an absence count. For training, evaluate transfer through skills demonstrations and manager follow-up instead of attendance alone. Review rating distributions for calibration fairness before using them for workforce decisions.\n\nLimitations include small sample size, random synthetic sampling, no full applicant pipeline, only one annual absence/training/rating period, no FTE or leave eligibility, no cost or productivity measures, no age bands chosen for legal/privacy review, and no causal design. The illustrative dataset omits real-world uncertainty and is unsuitable for compensation decisions, hiring selection, disciplinary decisions, legal compliance assessments, or benchmarking. A live project needs stakeholder-defined questions, privacy impact assessment where appropriate, data minimisation, governed access, and reproducible validated extracts." ,"table":"KPI definitions and sample results","figure":"charts/workforce_headcount_by_department.png","bullets":["Validate an agreed HR and finance headcount definition before publication.","Add applicant, offer, requisition, and vacancy data only under an approved purpose and access model.","Report small groups only when disclosure risk is acceptably controlled.","Use trends and scenarios to trigger discussion, not automated employment decisions."]},
            {"title":"Conclusion","body":"A basic analytics workflow can produce useful descriptive summaries when the question, dates, formula, and limitations are explicit. SAP HCM PA and OM data structures, Employee Central cloud core HR, and analytics products occupy different roles in the HR information landscape. The synthetic demonstration illustrates reproducible calculation and presentation, not a configuration or organizational finding. Its strongest practical lesson is governance: trustworthy workforce planning depends on complete definitions, data quality, privacy-preserving access, and human review." ,"figure":"charts/workforce_headcount_trend.png"},
        ],
        workforce_tables,
        [("charts/workforce_headcount_by_department.png","Figure 1. Active headcount by department; generated synthetic snapshot."),
         ("charts/workforce_headcount_trend.png","Figure 2. Synthetic quarter-end headcount counts; net totals do not explain underlying flows."),
         ("charts/workforce_recruitment_source.png","Figure 3. Recruitment source labels in synthetic records; not channel effectiveness.")],
        SAP_REFS + ["SAP Analytics Cloud documentation portal: https://help.sap.com/docs/SAP_ANALYTICS_CLOUD"],
        "Week_1_Workforce_Data_Analysis.docx"
    )

    payroll_table = [
        ("Conceptual payroll controls", ["Control area", "Illustrative control / configuration concept", "Evidence to retain"],
         [["Organisational assignment","Personnel number, company, personnel area/subarea, payroll area, cost assignment","Effective-dated master-data validation"],
          ["Payroll calendar","Period start/end, pay date, retro window, period status","Approved payroll calendar/version"],
          ["Wage types","Basic pay, recurring/non-recurring earnings, deductions; eligibility and processing behavior","Mapping, test cases, signed change record"],
          ["Calculation","Schema/PCR logic in classic SAP Payroll; country localization and legal parameters","Transport/change record and test results"],
          ["Posting","Symbolic accounts to GL accounts, cost objects, balancing rules","Posting simulation and finance reconciliation"],
          ["Access","Separate master-data maintenance, payroll execution, approval, and audit","Role matrix, logs, quarterly recertification"]]),
        ("Fictional payroll error register", ["Scenario", "Observed symptom", "Likely cause", "Corrective action / prevention"],
         [["PAY002","Net pay cannot reach valid bank destination","Bank master data incomplete or invalid","Validate bank details through controlled workflow; reconcile rejected payments; protect account data"],
          ["PAY003","Payroll run is blocked before calculation","Payroll control record is not released for the period","Check payroll area, period and status; authorized payroll owner releases after pre-run checks"],
          ["PAY004","Unexpected duplicate allowance","Recurring wage-type entry duplicated or overlapping","Inspect effective dates/source, correct one authorized record, rerun simulation, compare retro results"],
          ["PAY005","Prior-period difference after job change","Retroactive master-data change triggers retro-calculation","Validate effective date and retro limits; review delta and accounting impact before approval"],
          ["PAY006","Withholding differs from expected illustration","Tax declarations or localization parameters incomplete/outdated","Verify jurisdiction, declarations and current approved statutory configuration; do not hard-code rates"]]),
        ("Conceptual payroll test catalogue", ["Test class", "Example test", "Pass evidence"],
         [["Positive","Standard monthly salary, recurring benefit, approved one-time payment","Gross-to-net results match independently approved expected values"],
          ["Negative","Missing bank record, invalid wage-type eligibility, blocked control record","Run blocks or routes to a defined exception without silent payment"],
          ["Boundary","Period end, leap day, threshold boundary, part-period hire","Calendar and prorating behavior matches approved localization rules"],
          ["Retro","Backdated salary or organizational change within allowed retro window","Prior and current periods reconcile with traceable differences"],
          ["Integration","Payroll result to finance posting and payment file","Totals balance; rejected items and unmatched accounts are explained"]]),
    ]
    create_report(
        2, "Payroll Configuration", "Conceptual payroll architecture, controls, and testing",
        "This report explains a conceptual payroll lifecycle and the difference between classic SAP HCM Payroll and SAP SuccessFactors Employee Central Payroll (ECP). It covers master data, payroll area and control record, periods, wage types, classic infotypes, schemas and personnel calculation rules, deductions, localization, posting, reconciliation, testing, troubleshooting, and security. Tables map controls and fictional failure scenarios. Example amounts are simulated only: statutory tax, contribution, and deduction values are intentionally not calculated or represented as current law. No live SAP tenant was accessed and no configuration was performed.",
        [
            {"title":"Product roles and deployment context","body":"SAP HCM Payroll is the payroll calculation capability in SAP Human Capital Management deployments, traditionally associated with SAP ERP HCM and customer-managed landscapes. Product options, deployment, maintenance, country versions, and compatibility depend on the exact SAP release and current product lifecycle. SAP HCM for SAP S/4HANA is a distinct deployment context; it is unsafe to assume that all classic HCM functions are automatically part of a base S/4HANA installation.\n\nSAP SuccessFactors Employee Central is cloud core HR and maintains foundational employee/job/pay information. Employee Central Payroll (ECP) is SAP's cloud payroll offering based on SAP Payroll functionality and designed to integrate with Employee Central; it is not the same product as Employee Central itself. Some organizations use ECP as managed cloud payroll, classic payroll in an SAP ERP/HCM or HCM for SAP S/4HANA landscape, or a hybrid with other HR and finance systems. Architecture, country availability, localization, integration, service scope, and responsibilities require current SAP and implementation-partner confirmation.\n\nPayroll is a controlled calculation and disbursement process, not merely a monthly batch. It transforms effective-dated employee, time, compensation, and statutory inputs into gross-to-net results, payment instructions, accounting entries, and legally required outputs. This document is a learning design, not configuration guidance for production."},
            {"title":"Prerequisites and conceptual configuration sequence","body":"A payroll design begins with legal entity, country localization, pay frequency, calendar, population, currency, employment relationships, time inputs, finance requirements, data ownership, cut-off dates, and approval roles. Confirm the target deployment and supported localization before choosing configuration objects. Prepare master data through governed transactions or Employee Central data flows, define payroll areas and calendars, map wage types and eligibility, define calculation logic with supported country schema/PCR conventions, configure results and posting, integrate payment and finance, and establish test evidence and operations controls.\n\nIn classic SAP HCM, commonly relevant infotypes include 0000 Actions, 0001 Organizational Assignment, 0002 Personal Data, 0006 Addresses, 0007 Planned Working Time, 0008 Basic Pay, 0009 Bank Details, 0014 Recurring Payments/Deductions, 0015 Additional Payments, and 0105 Communications; country- and process-specific infotypes vary. Infotype numbers and use should be confirmed in the applicable release and localization. Employee Central uses cloud data models and effective-dated records rather than exposing every classic infotype as the same user-facing object. Integration mappings are not one-to-one by assumption.\n\nPayroll areas group employees for a pay period and run schedule. A control record governs the period state and activities for that area in classic payroll operations; process names and controls differ across products and releases. Payroll periods define dates, cut-offs, retroactivity limits, and payment dates. A schema orchestrates payroll processing steps and personnel calculation rules (PCRs) apply logic to wage types in classic payroll. Country-specific schemas, wage type processing classes, cumulations, evaluation classes, and legal rules must not be copied blindly between countries or releases." ,"steps":["Confirm deployment, country localization, pay population, and process ownership.","Validate employee and time master data and effective dates before period close.","Lock the approved payroll calendar, period status, and authorized release controls.","Simulate and inspect exceptions; correct source records with traceability.","Approve payroll results, payment output, accounting posting, and reconciliation.","Close the period with retained evidence and controlled retroactive change handling."]},
            {"title":"Wage types, deductions, and gross-to-net controls","body":"Wage types classify earnings and deductions and influence payroll processing, cumulation, valuation, tax treatment, posting, and reporting. Illustrative types include base salary, overtime, shift premium, bonus, meal allowance, employee contribution, tax withholding, loan recovery, and net-pay adjustment. A naming label alone does not define its behavior. Each type requires eligibility, frequency, calculation basis, validity dates, country treatment, account mapping, retroactivity behavior, and appropriate controls. One-time items and recurring payments need separate approval and duplicate detection.\n\nPayroll deductions may include statutory withholding, social insurance or pension contributions, benefit premiums, court-ordered deductions, and voluntary deductions, depending on jurisdiction and eligibility. Never use the illustrative INR amounts in the dataset to infer a tax rate, legal limit, or expected statutory deduction. Payroll localization and current legislation determine thresholds, rounding, exemptions, reporting, and remittance timing. Maintain effective-dated tax parameters through controlled updates and test relevant boundaries.\n\nGross-to-net control should compare input totals, gross pay, deductions, net pay, payment file totals, and accounting. Exception thresholds should route unusual amounts for review. Run-to-run comparisons need context for starters, leavers, one-off payments, retro changes, and calendar effects. Variances are investigated; they should not be suppressed to achieve a target total." ,"table":"Conceptual payroll controls"},
            {"title":"Payroll run, posting, security, and reconciliation","body":"A typical monthly cycle includes time and master-data cut-off, input validation, payroll simulation, exception resolution, authorized payroll release, productive calculation, result review, payment file preparation, approval, payment execution, accounting posting, reconciliation, and period closure. A rejected bank instruction is not proof that a payroll result was wrong; it is a payment exception that needs secure handling. Finance posting uses configured symbolic accounts and account determination in classic SAP contexts, with integration patterns varying by architecture. Reconcile payroll liabilities and expense by legal entity, cost centre, and period, and explain timing differences.\n\nSegregation of duties separates personnel-data changes, payroll configuration, payroll execution, result approval, payment release, and audit. Restrict display of bank and compensation details; mask or avoid exports; log privileged operations; recertify roles; and secure interfaces, background users, keys, and files. Test non-production systems with synthetic or masked data. Production access for troubleshooting should be time-bound and approved. Retain payroll evidence under applicable law and policy, with secure deletion after the retention period." ,"figure":"diagrams/payroll_process_flow.png","bullets":["Reconcile employee counts and high-level control totals before calculation.","Use approved transport/change practices and dual review for payroll logic changes.","Keep payment file generation, authorization, and release distinct when feasible.","Document run identifiers, parameters, exceptions, approvals, and reruns."]},
            {"title":"Fictional error scenarios, corrective actions, and test strategy","body":"The error register gives four fictional scenarios: missing bank details, an unreleased payroll-area period, a duplicate recurring allowance, a retroactive job change, and an incorrect tax declaration/localization setup. These are teaching examples, not observed SAP incidents. Troubleshooting starts with the exact employee/period/run context and a protected audit trail. Confirm input master data and time data, process messages, wage-type results, configuration version, control-record status, interfaces, and posting status. Determine whether the failure originates in source data, configuration, authorization, integration, or operational timing. Correct the root cause through an authorized path and rerun only the affected scope where product procedure permits; compare before/after results and obtain approval.\n\nA test catalogue should include positive, negative, boundary, retroactive, volume, security, integration, and regression cases. Build expected results independently from approved business rules and localization, not from the payroll engine alone. Include joiners, leavers, unpaid time, salary changes, garnishments where lawful, multiple assignments if supported, and leap-year/calendar boundaries. Reconcile payroll to payment and general ledger outputs. Capture anonymized evidence, defects, severity, owner, fix version, retest, and sign-off. A parallel run is appropriate where governance and jurisdiction permit, but has controlled scope, access, and communication." ,"table":"Fictional payroll error register","bullets":["PAY002: confirm bank data and payment rejection feedback without placing account data in unsecured logs.","PAY003: verify period and authorized control-record action before release.","PAY004: inspect effective dates and source transactions; avoid editing calculated results directly.","PAY005: confirm retro trigger, period scope, and approved delta before posting.","PAY006: use current approved localization; never replace statutory tables with a guessed percentage."]},
            {"title":"Conclusion","body":"Payroll reliability depends on clean effective-dated data, localization, configuration governance, controlled execution, financial reconciliation, and secure separation of duties. Classic SAP Payroll configuration concepts such as wage types, schemas, PCRs, infotypes, and control records are useful for understanding the processing model, but they must be mapped carefully to the selected cloud, ERP, or S/4HANA-related deployment. The fictional tables provide a structured design and test checklist only, not production-ready settings or legal calculations."},
        ],
        payroll_table,
        [("diagrams/payroll_process_flow.png","Figure 1. Generic payroll control flow; actual steps and product interfaces vary by deployment.")],
        SAP_REFS + ["SAP Help Portal, SAP SuccessFactors Employee Central Payroll documentation entry point: https://help.sap.com/docs/SAP_SUCCESSFACTORS_EMPLOYEE_CENTRAL_PAYROLL"],
        "Week_2_Payroll_Configuration.docx"
    )

    performance_tables = [
        ("Weighted scorecard", ["Dimension", "Weight", "Fictional score", "Weighted contribution"],
         [["Outcome objectives", "60%", "80 / 100", "48.0"],
          ["Competencies", "25%", "84 / 100", "21.0"],
          ["Values / behaviours", "15%", "78 / 100", "11.7"],
          ["Total", "100%", "—", "80.7 / 100"]]),
        ("Illustrative rating bands", ["Composite score", "Interpretation", "Required manager action"],
         [["90–100", "Exceptional contribution", "Validate evidence and sustainability; agree growth plan"],
          ["75–89.9", "Strong / effective", "Recognise impact; target one or two development priorities"],
          ["60–74.9", "Partially meets", "Diagnose barriers; specific support and check-ins"],
          ["Below 60", "Needs focused improvement", "Fair evidence review, clear expectations, support, and follow-up"]]),
    ]
    create_report(
        3, "Performance Management", "Objectives, evidence, feedback, and employee development",
        "This report designs a conceptual, employee-centred performance process using SAP SuccessFactors Performance & Goals as a capability reference. It distinguishes SMART objectives, key performance indicators (KPIs), objectives and key results (OKRs), competencies, and behaviours; proposes a weighted scorecard and fictional calculated example; and outlines the appraisal workflow, continuous feedback, learning plans, bias controls, change management, and implementation. SAP capabilities and tenant entitlements depend on the licensed release and configuration. All scores and examples are synthetic and cannot be used for real employment decisions.",
        [
            {"title":"Performance concepts and SAP SuccessFactors context","body":"Performance management is a recurring conversation linking role expectations, observable outcomes, capabilities, employee support, and organisational priorities. SMART objectives are specific, measurable, achievable, relevant, and time-bound. KPIs are selected measures of progress or outcome, not objectives by themselves. OKRs express an objective and a set of measurable key results; they are a goal-setting method rather than a rating scale. Competencies describe capabilities or behaviours expected in a role. A good system distinguishes output from controllable effort, context, teamwork, learning, and conduct.\n\nSAP SuccessFactors Performance & Goals is part of the cloud talent-management portfolio. Depending on the product generation, licence, configuration, and release, capabilities can support goals, performance forms, review cycles, competencies, feedback, calibration and development processes. SAP SuccessFactors Learning may support learning assignment and completion. These are product capabilities to evaluate, not guarantees that a particular tenant has a feature enabled. Employee Central supplies core employment context and role data; integration and data permissions require design. This conceptual system does not configure a tenant."},
            {"title":"Cycle design and evidence standards","body":"A balanced cycle starts with strategy translation and role-specific goal setting, followed by employee-manager alignment, periodic check-ins, evidence capture, self-reflection, manager assessment, calibration, discussion, development planning, and process evaluation. Check-in frequency should suit job cadence; quarterly is an example, not a universal prescription. Continuous feedback is timely, specific, respectful, and actionable. Year-end review should synthesise the full period rather than over-weight recent events.\n\nGoals must identify scope, baseline, target, deadline, data source, owner, dependencies, and controllability. An example objective: reduce average customer case resolution time by 10% by 31 December without worsening the agreed quality metric, using the approved case system. The goal owner should clarify how cases are counted and handle demand mix. A manager should not change goals retrospectively without documenting the reason and discussing it with the employee. Define permitted evidence sources and avoid collecting unrelated personal or sensitive data." ,"steps":["Set role-aligned objectives with measurable outcomes and relevant quality or risk guardrails.","Agree baseline, targets, support, dependencies, and check-in timing with each employee.","Hold documented check-ins and update goals when business priorities legitimately change.","Gather balanced evidence from the full review period and invite employee self-reflection.","Complete manager assessment, calibration with safeguards, and a feedback conversation.","Agree development commitments, learning resources, and a follow-up date."]},
            {"title":"Weighted scorecard and fictional calculations","body":"The sample scorecard weights outcome objectives at 60%, competencies at 25%, and values/behaviours at 15%. The example contributions are 80 × 0.60 = 48.0, 84 × 0.25 = 21.0, and 78 × 0.15 = 11.7, producing 80.7/100. The component values are fictional and converted to 0–100 solely for demonstration. A valid design needs descriptors, evidence expectations, role relevance, accessibility, and a clear method for goals that become obsolete or are blocked by external dependencies.\n\nWeighted composites create trade-offs and should not hide unacceptable risks. Consider minimum standards for safety, compliance, and conduct rather than allowing high output to offset a serious breach. Ratings are not precision measurements; decimal output can imply greater accuracy than the evidence supports. Use the composite as a conversation aid, not an automatic pay, promotion, or termination rule. If linked to compensation, disclose the process, validate equity impacts, and follow applicable law and collective arrangements." ,"table":"Weighted scorecard","figure":"charts/performance_score_distribution.png"},
            {"title":"Workflow, calibration, and bias prevention","body":"A configured workflow could assign a review form, employee self-review, manager review, optional second-level review, calibration meeting, final discussion, acknowledgement, and development follow-up. The process should state what acknowledgement means; acknowledgement of receipt is not necessarily agreement. Participants need role-appropriate access to forms and supporting evidence. Apply the data-minimisation and retention rules established for performance data.\n\nBias safeguards include structured behaviourally anchored criteria, consistent opportunities to set goals, accessible process design, review of rating distributions by relevant groups where lawful and privacy-protective, calibration notes based on evidence rather than reputation, guidance on recency/halo/similarity bias, and an appeal or correction route. Calibration should improve consistency, not force a predetermined distribution or erase individual context. Managers should disclose conflicts and avoid using protected traits, leave, or accommodation as negative proxies. Small group analysis needs suppression and careful interpretation." ,"bullets":["Train managers to separate observed behaviour from personality labels.","Audit objective quality and review completion, not just rating distributions.","Provide employee access to their review record and a transparent correction mechanism.","Restrict calibration materials and establish retention and export controls."]},
            {"title":"Development, learning, change, and implementation","body":"Development plans convert feedback into action. Each action should name a capability, practice activity, support, evidence of progress, and review date. Examples include shadowing a senior colleague, a stretch assignment with bounded risk, mentoring, peer learning, a course, or a task rotation. Training completion is an input; demonstrated transfer to work is the outcome. Avoid punitive or unrealistic development plans and agree workload and manager support.\n\nImplementation should begin with sponsor alignment, listening sessions, role/competency design, policy and privacy review, workflow and integration design, accessibility testing, pilot cohorts, manager and employee training, communications, controlled launch, and post-cycle evaluation. Define change ownership, configuration approval, form versioning, role permissions, data retention, support channels, and rollback or correction processes. Pilot measures may include goal quality, on-time check-ins, employee understanding, completion, evidence quality, appeal volume, and perceived fairness. Interpret scores alongside qualitative feedback and business context." ,"table":"Illustrative rating bands","bullets":["Avoid equating an employee's score with potential, value, or future readiness.","Provide manager enablement, employee guides, examples, and accessible support.","Evaluate fairness, trust, and learning transfer before expanding a pilot."]},
            {"title":"Conclusion","body":"A credible performance system sets expectations collaboratively, reviews meaningful evidence, gives continuous feedback, and funds realistic development. SuccessFactors Performance & Goals can provide configurable workflow and goal-management capabilities, but process quality, local policy, licensed features, data access, and human judgement determine whether the system is fair and usable. The calculated scorecard demonstrates arithmetic only; it is not a validated rating model."},
        ],
        performance_tables,
        [("charts/performance_score_distribution.png","Figure 1. Distribution of fictional composite performance scores among active records.")],
        SAP_REFS + ["SAP Help Portal, SAP SuccessFactors Performance & Goals documentation entry point: https://help.sap.com/docs/SAP_SUCCESSFACTORS_PERFORMANCE_AND_GOALS"],
        "Week_3_Performance_Management.docx"
    )

    succession_tables = [
        ("Fictional succession matrix", ["Critical role", "Role code", "Fictional candidate", "Readiness", "Key gap", "Next experience / contingency"],
         pd.read_csv(ROOT / "datasets" / "synthetic_succession_matrix.csv").values.tolist()),
        ("Role risk and continuity plan", ["Risk", "Likelihood", "Impact", "Preventive response", "Contingency"],
         [["Single incumbent leaves suddenly", "Possible", "Major", "Identify at least one interim delegate; document knowledge","Cross-trained deputy and authorised interim appointment"],
          ["Candidate exits before readiness", "Possible", "Major", "Career discussion, retention review, fair mobility","Maintain multiple candidates and external market plan"],
          ["Critical knowledge is undocumented", "Likely", "Major", "Role-based process maps, access-controlled handover","Time-bound transition and recovery procedure"],
          ["Readiness label is stale or biased", "Possible", "Moderate", "Evidence-based quarterly refresh and calibration","Reassess candidates; do not treat labels as promises"]]),
        ("Succession KPIs", ["Indicator", "Definition", "Example target for governance discussion"],
         [["Critical-role coverage","Critical roles with at least one validated successor ÷ total critical roles","Baseline first; improve coverage without forcing nominations"],
          ["Ready-now coverage","Roles with a validated ready-now option ÷ roles where immediate cover is required","Role-specific; explicitly document emergency-only cover"],
          ["Development milestone rate","Due roadmap milestones completed with evidence ÷ milestones due","Set cohort-specific target after baseline"],
          ["Internal placement rate","Critical vacancies filled internally ÷ filled critical vacancies","Monitor alongside fairness and retention"],
          ["Successor retention","Named candidates retained over period ÷ candidates at period start","Use trend with privacy and cohort-size safeguards"]]),
    ]
    create_report(
        4, "Succession Planning", "Critical-role continuity, talent development, and readiness",
        "This report defines a governed succession-planning method for a fictional organisation. It covers critical-role identification, talent reviews, competency mapping, readiness categories, succession matrices, nine-box assessment, skills gaps, leadership development, mentoring, job rotation, retention, risk, and continuity. The accompanying matrix contains six fictional critical roles and candidates; the nine-box counts are illustrative. SAP SuccessFactors Succession & Development is discussed as a configurable capability family, subject to current licensing and release. Succession nominations are confidential planning hypotheses—not guarantees, promises, or employment decisions.",
        [
            {"title":"Purpose, critical-role identification, and governance","body":"Succession planning reduces continuity risk for roles whose vacancy would materially impair strategy, safety, regulatory obligations, customer delivery, or scarce expertise. Criticality is role-based, not a judgement that one incumbent is more valuable as a person. Define criteria such as decision authority, replacement lead time, knowledge concentration, external talent scarcity, control ownership, and consequence of interruption. Confirm role boundaries, interim delegation authority, required credentials, and segregation-of-duties constraints before naming a successor.\n\nGovernance should include an accountable executive sponsor, HR talent team, business leaders, and risk/control owners. Establish a consistent review cadence, evidence standards, access-limited records, correction and challenge routes, retention period, and decision log. Candidates must not be told that a nomination guarantees promotion. Nomination should be based on role-relevant evidence and development interest, with an equitable process and documented uncertainty."},
            {"title":"Talent review, competency maps, and readiness levels","body":"A talent review compares demonstrated role outcomes, competencies, experiences, aspirations, mobility constraints, and development needs with future-role requirements. Competency maps should be specific and observable: strategic judgement, financial stewardship, people leadership, operational resilience, privacy/security governance, or technical expertise where relevant. Define proficiency levels and acceptable evidence. Current performance and future potential are distinct questions; exceptional performance in a current role does not automatically establish readiness for a different role.\n\nUse readiness labels with operational meaning, such as ready now, ready in 1–2 years, ready in 3–5 years, emergency cover only, or no identified successor. Time bands are planning estimates that must be validated periodically against experience and role needs. Distinguish an interim deputy from a fully qualified permanent successor. The fictional matrix identifies gaps and a development or contingency action for each role. The names and role codes are fictional labels, not actual people." ,"table":"Fictional succession matrix"},
            {"title":"Nine-box assessment and skills-gap analysis","body":"A nine-box grid crosses performance with potential to structure discussion, but it is not an objective test or a self-validating algorithm. Potential criteria should be transparent, job-related, and evidence-based, such as learning agility demonstrated in new contexts, ability to operate at greater scope, and motivation for the role. Assessors need common anchors and calibration training. A low or high box is not a permanent label, and development access should not be restricted only to high-potential employees.\n\nThe chart uses a fictional cohort distribution across developing, effective, and exceptional performance and potential categories. Counts are generated solely for illustration and do not come from the workforce ratings or a validated assessment. Skills-gap analysis compares the future-role competency profile with the candidate's verified evidence. Mark each gap with confidence and source; separate a genuine gap from missing evidence. Prioritise gaps by role criticality and time to build, and identify whether learning, experience, hiring, knowledge capture, or contingency can close the risk." ,"figure":"charts/succession_nine_box.png","bullets":["Use multiple evidence sources and structured calibration, not intuition alone.","Document uncertainty, candidate interest, mobility and accessibility context, and conflicts of interest.","Review representation and access to development without exposing small-group personal data."]},
            {"title":"Development roadmaps and leadership pipeline","body":"A roadmap links a named role requirement to learning through experience, exposure, and education. For a payroll-operations lead, a bounded rotation through payroll controls, finance reconciliation, and incident management may test cross-functional readiness. For an HRIS product owner, leading a scoped release can develop prioritisation and change governance. Senior leadership candidates may receive sponsor mentoring, executive simulations, budget responsibility, and stakeholder feedback. Each assignment needs an explicit objective, supervisor, risk boundary, success evidence, and review date.\n\nMentoring supports reflection and networks but does not replace accountable line management or role competence. Job rotation should not disrupt critical operations, expose restricted data unnecessarily, or create segregation-of-duties conflicts. Provide equitable access to stretch work and review workload, accommodation, and employee consent. Retention actions should address meaningful work, development, succession transparency at an appropriate level, and fair reward; avoid treating a person as a retention risk based on protected characteristics." ,"figure":"diagrams/succession_governance_cycle.png"},
            {"title":"Risk, contingency, metrics, and SAP capabilities","body":"Continuity planning covers sudden vacancy, candidate departure, readiness delays, undocumented knowledge, security-sensitive access, and stale or biased assessments. Preventive measures include cross-training, documented procedures, approved delegation, access provisioning controls, multiple candidates where appropriate, external-market plans, and tabletop exercises. Contingency steps should identify who can appoint an interim, what controls remain mandatory, who receives notification, and how long the arrangement can continue.\n\nThe KPI table offers definitions for critical-role coverage, ready-now coverage, development milestones, internal placements, and successor retention. Establish a baseline and role-specific targets rather than choosing an arbitrary universal percentage. Review both coverage and quality: an inflated coverage measure can conceal weak evidence or over-nomination. Monitor disparities and candidate experience with adequate group sizes and privacy safeguards.\n\nSAP SuccessFactors Succession & Development can support talent profiles, succession org charts, position-based succession planning, talent pools, readiness, and development planning depending on licensing, configuration, and current release. Product terminology and feature availability change; validate in current SAP Help and the customer's tenant. Establish permission roles and confidentiality because succession data can be highly sensitive. A product feature does not replace governance or continuity documentation." ,"table":"Role risk and continuity plan"},
            {"title":"Conclusion","body":"Effective succession planning starts with genuinely critical roles, clarifies the evidence and readiness meaning, develops capability through supported experiences, and maintains credible contingencies. The matrix, roadmap, and nine-box are fictional teaching artefacts. They demonstrate how a succession conversation can be structured; they do not predict career outcomes or identify actual successors. Review decisions periodically, protect the records, and measure resilience as well as development." ,"table":"Succession KPIs"},
        ],
        succession_tables,
        [("charts/succession_nine_box.png","Figure 1. Synthetic nine-box distribution; cells are fictional counts and are not a validated potential measure."),
         ("diagrams/succession_governance_cycle.png","Figure 2. Governance cycle for role identification, talent review, development, validation, and continuity.")],
        SAP_REFS + ["SAP Help Portal, SAP SuccessFactors Succession documentation entry point: https://help.sap.com/docs/SAP_SUCCESSFACTORS_SUCCESSION"],
        "Week_4_Succession_Planning.docx"
    )

    security_tables = [
        ("HR data asset and control map", ["Asset / activity", "Primary risk", "Illustrative preventive / detective controls"],
         [["Employee identifiers and demographics","Over-broad access, disclosure, inaccurate use","Purpose limitation, minimisation, role-based field access, aggregation/suppression"],
          ["Compensation and bank details","Fraud, identity theft, payment diversion","Restricted roles, maker-checker payment change, masking, secure transmission"],
          ["Performance and succession","Unfair decisions, retaliation, confidential disclosure","Need-to-know access, transparent process, audit trail, retention limit"],
          ["Payroll integrations and files","Interception, tampering, duplicate payment","Authenticated encrypted interfaces, integrity checks, reconciliation, key management"],
          ["Privileged configuration and logs","Unapproved change or evidence alteration","Separate admin roles, monitored emergency access, protected central logs"]]),
        ("Inherent and residual risk register", ["Threat scenario", "Likelihood (1-5)", "Impact (1-5)", "Inherent score", "Controls / mitigation", "Residual risk (illustrative)"],
         [["Payroll bank-detail change fraud","3","5","15","Dual approval, out-of-band verification, change alerts, payment reconciliation","Medium; residual likelihood remains non-zero"],
          ["Manager views unauthorised performance data","3","4","12","Role-based field permissions, periodic access review, access logs","Low-Medium; misassignment and exports remain"],
          ["Phishing compromises HR user","4","4","16","MFA, awareness, mail controls, conditional access, response drill","Medium; credential theft and social engineering persist"],
          ["Unencrypted or misrouted export","3","5","15","Disable unnecessary exports, secure transfer, DLP, encryption, recipient checks","Low-Medium; human error remains"],
          ["Excessive retention of applicant/employee records","3","4","12","Retention schedule, legal holds, automated deletion, evidence of disposal","Low-Medium; legal holds and backups complicate deletion"],
          ["Availability disruption to payroll period close","2","5","10","Backups, tested recovery, documented manual contingency, vendor escalation","Medium; recovery and external dependency risks remain"]]),
        ("Audit and security KPIs", ["KPI", "Formula", "Control interpretation"],
         [["Privileged access recertification","Reviewed privileged accounts ÷ accounts due for review","Target complete, on time; investigate exceptions"],
          ["MFA coverage","In-scope users with enforced MFA ÷ in-scope users","Segment privileged and workforce access"],
          ["Security incident response time","Median time from alert to triage / containment","Measure by severity and clock definition"],
          ["Access removal timeliness","Leavers disabled within policy SLA ÷ leavers sampled","Reconcile HR events to identity systems"],
          ["Sensitive export exceptions","Count and rate of unauthorised / unexplained exports","Trend and investigate; zero is not assumed"],
          ["Restore test success","Successful tested restores ÷ restores planned","Test recoverability and payroll continuity"]]),
    ]
    create_report(
        5, "HR Data Security", "Confidentiality, privacy, access control, and resilience",
        "This report applies confidentiality, integrity, and availability principles to HR records, payroll, performance and succession data. It covers SAP authorization concepts, least privilege, role-based access, segregation of duties, encryption, logs, monitoring, retention, incident response, awareness, GDPR, India's Digital Personal Data Protection Act 2023 and Rules, and recognized security standards. A likelihood-impact matrix and six-row illustrative risk register show inherent and residual risk. Legal applicability and commencement are fact-specific; the report is educational and is not legal advice. The matrix scores and scenarios are fictional.",
        [
            {"title":"Scope, data classification, and security objectives","body":"HR systems process identity, contact, employment, organisational, compensation, bank, attendance, performance, learning, and sometimes health or accommodation information. These data types do not all carry the same sensitivity or legal treatment. Classify records, identify owners and processors, document purposes and flows, minimise collection, and define access and retention according to jurisdiction and business need. Confidentiality prevents unauthorised disclosure; integrity preserves accuracy and completeness; availability ensures authorised users and processes can access data when needed. Accountability, privacy, and resilience complement the CIA objectives.\n\nSAP landscapes may include on-premise ERP HCM, SAP HCM for SAP S/4HANA contexts, SuccessFactors cloud services, identity providers, payroll providers, banks, analytics platforms, and local interfaces. Responsibility is shared across customer, SAP, and other providers according to contracts and deployment model. Verify current SAP security documentation, service descriptions, data-processing terms, configuration, and integration boundaries. Do not infer that a cloud service makes customer role design or lawful processing automatic."},
            {"title":"Access control, least privilege, and SAP authorization concepts","body":"SAP access control commonly combines user identity, roles, authorizations, organizational restrictions, application controls, and auditability. In classic ABAP-based SAP systems, authorization objects and field values are evaluated by authorization checks; role administration and derived/composite role patterns are release and design dependent. Organizational fields can restrict access by company or personnel area when correctly implemented and tested. SuccessFactors uses its own role-based permissions and target-population concepts; do not assume an ABAP authorization design maps one-to-one to cloud roles. Product-specific configuration and identity integration must be reviewed in current documentation.\n\nLeast privilege grants only the minimum required access for the shortest appropriate period. Segregation of duties separates incompatible tasks—for example changing bank details, processing payroll, approving results, and releasing payments. A role matrix should identify job function, system, data object, actions, organisational scope, approver, and recertification cadence. Test negative access: a user must not see or alter data outside the authorised population. Control emergency access with approval, time limits, session logs, and independent review." ,"table":"HR data asset and control map"},
            {"title":"Technical and operational safeguards","body":"Protect data in transit with authenticated, current cryptographic protocols and at rest using supported encryption and key management. Confirm which application, database, backup, and interface layers are covered; encryption does not compensate for excessive access or insecure endpoints. Secure secrets and integration credentials, rotate keys, restrict service identities, validate certificates, and monitor failed authentication. Mask or pseudonymise non-production data; synthetic data is preferable for training and routine development.\n\nAudit logs should record who accessed or changed sensitive records, what changed, when, the relevant identity and context, and the approval or reason where appropriate. Protect logs from unauthorised alteration, limit access to security investigators, synchronise time, and define retention. Monitor anomalous exports, privileged access, role changes, bank-detail changes, mass downloads, failed logins, and integration errors with privacy-proportionate alerts. Retention schedules should cover active records, backups, legal holds, and secure disposal. Incident plans should include triage, containment, preservation, impact assessment, notification decisions, recovery, and lessons learned; statutory notification timelines vary."},
            {"title":"Privacy law and standards: applicability boundaries","body":"The EU General Data Protection Regulation (GDPR) applies when its territorial and material scope tests are met; it is not automatically applicable to every employer. Relevant principles include lawfulness, fairness and transparency, purpose limitation, data minimisation, accuracy, storage limitation, integrity/confidentiality, and accountability. Employment processing also interacts with national law, worker consultation, and restrictions on special-category data. Determine controller/processor roles, lawful basis, notices, rights, international transfers, records, and impact assessment duties with qualified counsel.\n\nIndia's Digital Personal Data Protection Act, 2023 establishes a framework for digital personal data processing within its scope. The Digital Personal Data Protection Rules and commencement notifications should be checked in the official Gazette for the processing date because commencement and obligations may be phased. Confirm applicability, notices, consent or other permitted grounds, data-principal rights, security safeguards, breach obligations, retention, and any exemptions with counsel. This report does not claim every provision is in force for every entity or date. The Information Technology Act / rules and sector-specific employment or financial obligations may also be relevant depending on facts.\n\nISO/IEC 27001 provides requirements for an information security management system; NIST SP 800-53 provides a control catalogue, and NIST CSF 2.0 is a cybersecurity risk-management framework. They are not substitutes for privacy law or a jurisdiction-specific compliance assessment. HIPAA in the United States applies to covered entities and business associates in defined contexts; it does not automatically apply to every employer merely because the employer holds health-related information. Specialist counsel should determine role and scope." ,"figure":"charts/hr_security_risk_matrix.png"},
            {"title":"Risk assessment, mitigation, and residual risk","body":"The risk matrix multiplies ordinal likelihood and impact scores from 1 to 5 for a simple inherent score from 1 to 25. It is a prioritisation aid, not a statistically calibrated probability. Impact includes confidentiality, employee rights, financial loss, payroll disruption, legal exposure, and trust. The register describes six fictional threats, mitigating controls, and qualitative residual risk. Scores must be recalibrated with asset context, threat intelligence, control effectiveness, vulnerability, and risk appetite.\n\nResidual risk remains after controls: phishing can succeed despite MFA, administrators may misconfigure roles, encrypted files can be misdirected to authorised but unintended recipients, retention deletion can be constrained by backups or legal holds, and vendor outages can exceed recovery assumptions. Assign a risk owner, treatment, due date, evidence, and acceptance authority. Reassess after system changes, incidents, new integrations, acquisitions, or legal changes. Test controls through access samples, negative tests, restore exercises, incident tabletop exercises, and log review." ,"table":"Inherent and residual risk register"},
            {"title":"Incident response, awareness, and audit measures","body":"An HR incident playbook identifies reporting channels, on-call security and privacy contacts, payroll operations, legal, communications, and business continuity. Initial actions preserve evidence and stop ongoing exposure without destroying logs. Determine affected data, people, systems, countries, recipients, duration, and likelihood of harm. Restrict access, rotate credentials where needed, validate payroll/payment integrity, and use approved notification decision paths. Notification duties and deadlines are determined by applicable law and contract; avoid promising timelines before assessment. After recovery, communicate lessons, remediate root causes, and verify completion.\n\nAwareness training should be role-specific. HR and payroll users need practice spotting payroll diversion attempts, verifying identity changes, handling sensitive exports, and reporting suspicious access. Managers should understand confidentiality of performance and succession data. Administrators need secure configuration and break-glass procedures. Audit KPIs should have an owner, denominator, target, data source, and period. Review exceptions and effectiveness, not just completion percentages. The KPI table contains sample definitions; targets must be approved through governance." ,"table":"Audit and security KPIs","bullets":["Report suspected disclosure quickly through an accessible channel; do not investigate by forwarding sensitive data.","Test joiner, mover, and leaver access events end to end.","Retest after permission or integration changes and retain evidence securely.","Review vendor, interface, and business-continuity dependencies in tabletop exercises."]},
            {"title":"Conclusion","body":"HR security is a lifecycle responsibility covering lawful purpose, data minimisation, identity and role governance, secure processing, monitoring, retention, incident response, and recovery. SAP authorization mechanisms are deployment-specific controls within a wider operating model. GDPR, India's DPDP framework, and security standards have different scope and effect; verify the law and Gazette notifications applicable at the time. The fictional risk register supports discussion, not certification or legal advice."},
        ],
        security_tables,
        [("charts/hr_security_risk_matrix.png","Figure 1. Ordinal likelihood-impact matrix; scores are illustrative and not measured probabilities.")],
        SAP_REFS + SECURITY_REFS,
        "Week_5_HR_Data_Security.docx"
    )


def make_performance_chart() -> None:
    scores = pd.read_csv(ROOT / "datasets" / "synthetic_performance_scores.csv").weighted_score
    plt.figure(figsize=(8, 4.4))
    plt.hist(scores, bins=[40, 50, 60, 70, 80, 90, 100], color=f"#{TEAL}", edgecolor="white")
    plt.title("Fictional weighted performance score distribution")
    plt.xlabel("Composite score / 100"); plt.ylabel("Employee records")
    plt.tight_layout()
    plt.savefig(ROOT / "charts" / "performance_score_distribution.png", dpi=180)
    plt.close()


def make_descriptions(metrics: dict) -> None:
    descriptions = {
        1: f"""Week 1 examines workforce data analysis as a disciplined decision-support activity rather than a collection of charts. The report explains SAP HCM, Personnel Administration, Organizational Management, SAP SuccessFactors Employee Central, and how operational HR reporting differs from analytics. The working dataset contains 120 reproducibly generated fictional records with a fixed random seed. Each record includes department, location, job title, demographic fields, hire and possible termination dates, salary, recruitment source, FY2025 absence days, learning hours, and a performance rating. The file is explicitly synthetic; it contains no real employee information and is not suitable for employment decisions.

The methodology applies a 30 September 2026 snapshot for active headcount and defines FY2025 attrition as exits during 2025 divided by the average of opening and closing headcount. It reports department and demographic composition, recruitment-source counts, absence, training, and performance summaries. Python and pandas create the data and calculate the measures; matplotlib generates department headcount, a quarterly headcount trend, and a recruitment-source chart. The report includes KPI formulas, summary tables, interpretation, collection and data-quality considerations, limitations, and references to SAP documentation.

The main finding is that a reproducible descriptive pipeline can make workforce assumptions explicit, but sample statistics cannot be treated as facts about a real company. The small synthetic cohort, incomplete applicant funnel, absent FTE and leave eligibility, and lack of causal design limit interpretation. Recommendations include agreeing metric definitions, reconciling HR/payroll populations, using privacy-protective access and suppression, validating data lineage, and using trends to prompt workforce-planning discussion rather than automated individual decisions. The corresponding submission is the Week 1 report together with its workforce description.""",
        2: """Week 2 presents payroll as a controlled gross-to-net and accounting process, not as a live SAP configuration. It distinguishes classic SAP HCM Payroll from SAP SuccessFactors Employee Central Payroll, explains the core-HR role of Employee Central, and identifies that product scope, licensing, country localization, and deployment context must be verified against the current SAP release. The report introduces payroll prerequisites; employee master data; common classic infotypes; payroll areas and control records; payroll periods; wage types; schemas and personnel calculation rules; deductions; statutory localization; posting; reconciliation; testing; troubleshooting; and security.

The design methodology uses a generic lifecycle from calendar and master-data readiness through simulation, exception resolution, authorized calculation, payment approval, finance posting, reconciliation, and close. Tables organize conceptual configuration controls, fictional payroll exception scenarios, and a test catalogue. The scenarios include missing bank details, a blocked payroll period, a duplicate recurring allowance, a retroactive job change, and an incorrect withholding setup. Each is paired with root-cause checks and a corrective control. No real employee data is used; a small synthetic payroll scenario dataset contains illustrative amounts solely to make the examples concrete.

The report deliberately does not publish tax rates, contribution thresholds, or legal deduction calculations. Statutory rules depend on jurisdiction and current effective dates, so the recommendation is to use approved local configuration and independent validation. Core recommendations are to separate maintenance, calculation, approval, payment, and audit responsibilities; prove expected results with positive, negative, boundary, retro, integration, and regression testing; retain evidence; and reconcile payroll to payment and general ledger totals. The payroll flow diagram provides a concise process view. Submit this report with the Week 2 description.""",
        3: """Week 3 designs a conceptual performance-management cycle using SAP SuccessFactors Performance & Goals as a capability reference. It defines SMART objectives, KPIs, OKRs, competencies, appraisals, continuous feedback, and employee development without treating these terms as interchangeable. The methodology combines a fictional scorecard with an evidence-oriented workflow: collaborative goal-setting, periodic check-ins, self-review, manager review, calibration, a feedback conversation, and a development plan.

The weighted example assigns 60 percent to outcomes, 25 percent to competencies, and 15 percent to values or behaviours. Component values are converted to a 100-point scale and combined with transparent arithmetic. A generated set of active synthetic employee records also yields fictional weighted scores for a histogram. These example scores are not validated assessments, and the report warns that decimal outputs can suggest false precision. A rating band table shows how a composite could support a conversation while not replacing minimum standards for safety, conduct, or compliance.

The report recommends role-relevant measurable objectives with clear baselines, deadlines, owners, dependencies, and quality guardrails. It covers common bias risks, structured evidence, accessible process design, calibration without forced distributions, conflict disclosure, privacy controls, employee correction pathways, and review of fairness. Development actions identify a capability, an experience or learning activity, support, evidence, and a follow-up date. Implementation guidance covers sponsor alignment, design, privacy review, pilot, manager enablement, launch, and evaluation. The key conclusion is that software workflow cannot itself guarantee fairness: governance, meaningful evidence, manager support, and employee trust determine whether the process is useful. Submit the Week 3 report with this description.""",
        4: """Week 4 develops a fictional succession-planning approach focused on critical-role continuity and the deliberate development of capability. The report explains how to identify roles by vacancy impact, replacement lead time, regulatory or control importance, scarce expertise, and knowledge concentration. It distinguishes role criticality from an incumbent's personal value and establishes governance for confidentiality, evidence, review cadence, conflicts, and candidate challenge.

The supporting dataset contains a six-role fictional succession matrix with readiness labels, role-relevant gaps, and development or contingency actions. The report also explains competency mapping, readiness bands, succession matrices, skills-gap analysis, talent reviews, leadership development, mentoring, job rotation, and retention. A nine-box diagram presents made-up category counts for illustrative calibration only; it is not a validated potential instrument, and its values are not inferred from real performance data. A governance-cycle diagram links identification, review, development, validation, and continuity planning.

The findings emphasise that readiness labels are time-bounded hypotheses, not promises of promotion. Candidate performance in a current role is not equivalent to future potential in a different role. Development roadmaps should build specific skills through supported assignments, mentoring, simulations, and role rotations with clear learning evidence and risk boundaries. Continuity planning addresses sudden departures, candidate attrition, undocumented knowledge, stale or biased assessments, and interim authority. The report proposes metrics for critical-role coverage, ready-now cover, development milestones, internal placement, and successor retention, while advising against arbitrary targets that encourage over-nomination. SAP SuccessFactors Succession & Development capabilities are described cautiously as configurable and licensing-dependent. Submit the Week 4 report with this matching description.""",
        5: """Week 5 reviews HR data security and privacy as lifecycle responsibilities. It applies confidentiality, integrity, and availability to employee, payroll, performance, and succession records and explains why classification and purpose determine how data should be protected. The report distinguishes SAP authorization concepts in classic ABAP-based landscapes from role-based permissions in SuccessFactors cloud products, and covers least privilege, organizational restrictions, segregation of duties, encryption, key management, audit logs, monitoring, retention, incident response, recovery, and awareness.

The methodology is a qualitative fictional risk assessment. A likelihood-impact grid scores likelihood and impact on five-point ordinal scales and multiplies them to create an illustrative inherent priority score. Six risk scenarios include payroll bank-detail fraud, inappropriate performance-data access, phishing, misrouted exports, excessive retention, and payroll service disruption. Each row pairs mitigation controls with a qualitative residual-risk note. A control map and audit KPI table define example measures for privileged access reviews, multi-factor authentication coverage, incident response time, leaver-access removal, export exceptions, and restore testing. The scores are prompts for discussion, not empirically measured probabilities or certification findings.

Legal discussion differentiates GDPR's scope from universal applicability and notes that India's Digital Personal Data Protection Act 2023, relevant Rules, and commencement notifications must be checked against official Gazette publications and the processing date. It also clarifies that HIPAA does not automatically apply to every employer just because health-related information is held. NIST and ISO standards are presented as control or management frameworks, not substitutes for law. Recommendations include restricted role design, tested negative permissions, strong authentication, maker-checker payment changes, protected logs, secure deletion, rehearsed incident procedures, and jurisdiction-specific counsel. Submit the Week 5 report with this description.""",
    }
    for week, text in descriptions.items():
        path = ROOT / "submission_descriptions" / f"Week_{week}_Description.txt"
        path.write_text(text.strip() + "\n", encoding="utf-8")


def write_references() -> None:
    text = """# References and source-verification notes

Sources were selected from official product, government, and recognized standards organizations. Product landing pages are entry points rather than citations to an assumed feature entitlement; verify the current release-specific topic and tenant before implementation. Legal conclusions in the reports are educational summaries, not legal advice.

## SAP product documentation

- SAP Help Portal, SAP SuccessFactors Employee Central: https://help.sap.com/docs/SAP_SUCCESSFACTORS_EMPLOYEE_CENTRAL
- SAP Help Portal, SAP SuccessFactors Platform: https://help.sap.com/docs/SAP_SUCCESSFACTORS_PLATFORM
- SAP Help Portal, SAP ERP (including HCM topics): https://help.sap.com/docs/SAP_ERP
- SAP Help Portal, SAP SuccessFactors Employee Central Payroll: https://help.sap.com/docs/SAP_SUCCESSFACTORS_EMPLOYEE_CENTRAL_PAYROLL
- SAP Help Portal, SAP SuccessFactors Performance & Goals: https://help.sap.com/docs/SAP_SUCCESSFACTORS_PERFORMANCE_AND_GOALS
- SAP Help Portal, SAP SuccessFactors Succession: https://help.sap.com/docs/SAP_SUCCESSFACTORS_SUCCESSION
- SAP Help Portal, SAP Analytics Cloud: https://help.sap.com/docs/SAP_ANALYTICS_CLOUD

## Privacy, security, and law

- European Union, Regulation (EU) 2016/679 (GDPR), official text: https://eur-lex.europa.eu/eli/reg/2016/679/oj
- Government of India, MeitY, data protection framework: https://www.meity.gov.in/data-protection-framework
- Government of India, e-Gazette portal; verify the Digital Personal Data Protection Rules and commencement notifications applicable on the date of processing: https://egazette.gov.in/
- NIST SP 800-53 Rev. 5, Security and Privacy Controls: https://csrc.nist.gov/pubs/sp/800/53/r5/upd1/final
- NIST Cybersecurity Framework 2.0: https://www.nist.gov/cyberframework
- ISO/IEC 27001:2022 standard overview: https://www.iso.org/standard/27001
- U.S. HHS, HIPAA for Professionals: https://www.hhs.gov/hipaa/for-professionals/index.html

Reference review date: 18 October 2026. Links lead to official documentation or source portals; no unsourced quotations or claimed implementation outcomes are used. Some government legal pages restricted automated retrieval, so the final Act text, Rules, and Gazette notifications were not independently fetched in this environment. The MeitY and e-Gazette pages may change their page structure; confirm the controlling statute, rules, commencement notification, and jurisdictional application with qualified counsel.
"""
    (ROOT / "references" / "README.md").write_text(text, encoding="utf-8")


def validate_project() -> list[str]:
    errors: list[str] = []
    expected_reports = sorted((ROOT / "reports").glob("Week_*.docx"))
    expected_descriptions = sorted((ROOT / "submission_descriptions").glob("Week_*_Description.txt"))
    if len(expected_reports) != 5:
        errors.append(f"Expected 5 DOCX reports, found {len(expected_reports)}")
    if len(expected_descriptions) != 5:
        errors.append(f"Expected 5 descriptions, found {len(expected_descriptions)}")
    for path in expected_descriptions:
        words = re.findall(r"\b[\w’'-]+\b", path.read_text(encoding="utf-8"))
        if len(words) < 250:
            errors.append(f"{path.name} has only {len(words)} words")
    for path in expected_reports:
        try:
            doc = Document(path)
            text = "\n".join(p.text for p in doc.paragraphs)
            all_text = "\n".join(
                [text, *[cell.text for table in doc.tables for row in table.rows for cell in row.cells]]
            )
            if len(all_text.split()) < 1800:
                errors.append(f"{path.name} appears too short ({len(all_text.split())} document words)")
            if "References" not in text or "Executive Summary" not in text or "Contents" not in text:
                errors.append(f"{path.name} is missing required front/back matter")
            if len(doc.tables) < 2:
                errors.append(f"{path.name} has fewer than two tables")
            image_count = len(doc.inline_shapes)
            if image_count < 1:
                errors.append(f"{path.name} has no embedded images")
            if re.search(r"\b(TODO|TBD|PLACEHOLDER|lorem ipsum|insert here)\b", text, re.I):
                errors.append(f"{path.name} contains an unfinished placeholder")
        except Exception as exc:
            errors.append(f"Cannot read {path.name}: {exc}")
    employee_path = ROOT / "datasets" / "synthetic_employee_workforce.csv"
    if employee_path.exists():
        employees = pd.read_csv(employee_path)
        if len(employees) < 100:
            errors.append(f"Synthetic employee dataset has only {len(employees)} rows")
        if employees.employee_id.nunique() != len(employees):
            errors.append("Employee IDs are not unique")
    else:
        errors.append("Missing workforce dataset")
    required_images = [
        "charts/workforce_headcount_by_department.png",
        "charts/workforce_headcount_trend.png",
        "charts/workforce_recruitment_source.png",
        "charts/performance_score_distribution.png",
        "charts/succession_nine_box.png",
        "charts/hr_security_risk_matrix.png",
        "diagrams/payroll_process_flow.png",
        "diagrams/succession_governance_cycle.png",
    ]
    for relative in required_images:
        if not (ROOT / relative).is_file() or (ROOT / relative).stat().st_size < 2000:
            errors.append(f"Missing or empty image: {relative}")
    return errors


def main() -> None:
    for folder in ["reports", "submission_descriptions", "datasets", "charts", "diagrams", "scripts", "references"]:
        (ROOT / folder).mkdir(parents=True, exist_ok=True)
    frame = make_dataset()
    metrics = calculate_metrics(frame)
    chart_workforce(frame, metrics)
    draw_flow(ROOT / "diagrams" / "payroll_process_flow.png", "Conceptual payroll run and control flow",
              ["Set calendar, scope, statutory and finance rules", "Validate master data, time, and payroll-area period",
               "Simulate calculation and inspect exceptions", "Correct root causes; review gross-to-net and retro items",
               "Approve results and payment output", "Post to finance, reconcile, close, retain evidence"])
    draw_succession_visuals()
    draw_security_matrix()
    make_performance_chart()
    build_reports(frame, metrics)
    make_descriptions(metrics)
    write_references()
    errors = validate_project()
    if errors:
        raise SystemExit("\n".join(errors))
    print(f"Generated 5 reports, 5 descriptions, {len(frame)} employee records, 4 datasets, and all visuals in {ROOT}")
    print(f"Validation passed: descriptions >=250 words; reports have contents, tables, figures, references, and no unfinished placeholders.")
    print(f"Snapshot: active headcount={metrics['headcount']}; FY2025 attrition={metrics['attrition_2025']:.1f}% (fictional).")


if __name__ == "__main__":
    main()
