# CAPA Closure Assistant

**Python analytics workflow for CAPA triage, closure readiness, and management visibility**

## Executive Summary

The CAPA Closure Assistant is a portfolio project that demonstrates how Python can turn a corrective and preventive action (CAPA) tracker into a prioritized management view.

The workflow identifies overdue and due-soon records, highlights incomplete closure elements, calculates a transparent triage score, and produces Excel and text outputs that can support recurring CAPA review meetings.

**Portfolio note:** All records in this repository are synthetic. The project is informed by general quality-system concepts and domain experience; it does not reproduce a proprietary company system, dataset, or confidential workflow.

## Business Problem

CAPA trackers can become difficult to manage when teams rely on manual review to determine:

- Which open records need immediate attention
- Which records are approaching their due dates
- Which closure elements are incomplete
- Which records appear ready for closure review
- Where aging or ownership patterns may require follow-up

The goal of this project is to convert a flat tracker into a repeatable, risk-oriented review process.

## What the Analysis Does

The Python workflow:

1. Loads a CAPA tracker from CSV or Excel.
2. Validates required columns and standardizes date fields.
3. Separates open from closed/canceled records.
4. Calculates CAPA age and days until due.
5. Flags overdue and due-soon records.
6. Identifies missing root-cause, action, verification, and effectiveness elements.
7. Calculates a transparent triage score.
8. Identifies records that meet the defined closure-readiness rule.
9. Exports management-ready Excel tabs and a weekly text summary.

## Triage Logic

The demonstration score is intentionally simple and explainable:

| Condition | Points |
|---|---:|
| Overdue | +3 |
| Due within configured window | +2 |
| One or more closure elements missing | +2 |
| Open 30+ days | +1 |

The score is a prioritization aid for this portfolio project, **not a validated regulatory risk model**.

## Technologies Demonstrated

- Python
- pandas
- NumPy
- Excel/CSV data processing
- Data validation
- Datetime calculations
- Rule-based prioritization
- KPI creation
- Data visualization
- Business-process analytics
- Quality/CAPA domain translation

## Repository Structure

```text
CAPA-Closure-Assistant/
├── README.md
├── capa_closure_assistant.py
├── generate_mock_capa_data.py
├── requirements.txt
├── .gitignore
├── mock_capa_data.csv
├── mock_capa_data.xlsx
├── assets/
└── outputs/
```

## How to Run

Install dependencies:

```bash
pip install -r requirements.txt
```

Generate a fresh synthetic dataset:

```bash
python generate_mock_capa_data.py
```

Run the analysis:

```bash
python capa_closure_assistant.py --input mock_capa_data.xlsx --outdir outputs
```

Optional: change the due-soon window:

```bash
python capa_closure_assistant.py --input mock_capa_data.xlsx --outdir outputs --due-soon-days 21
```

## Outputs

The workflow creates:

- `outputs/capa_triage.xlsx`
  - Open
  - Overdue
  - DueSoon
  - MissingInfo
  - ClosureReady
  - Summary
- `outputs/weekly_update.txt`
- `assets/capa_health_overview.png`
- `assets/capa_age_distribution.png`
- `assets/capa_overdue_by_owner.png`

## Example Business Questions

This project can help answer questions such as:

- How many CAPAs are currently open?
- Which CAPAs are overdue or approaching their due date?
- What closure elements are most frequently incomplete?
- Which owners have overdue items requiring follow-up?
- Which records meet the defined closure-readiness criteria?
- How is the open CAPA population distributed by age?

## Limitations

- The dataset is synthetic and does not represent actual company performance.
- The scoring logic is illustrative and would require governance and validation before use in a regulated production environment.
- The project does not replace quality-system procedures or required human review.
- The current workflow is file-based rather than connected to a production QMS or database.

## Interview Talking Point

A concise way to explain this project:

> I built a Python workflow that converts a synthetic CAPA tracker into a prioritized management view. Using pandas, I calculate aging and due-date metrics, flag incomplete closure elements, apply an explainable triage score, and export review-ready Excel outputs. I chose a transparent rule-based approach because CAPA prioritization needs to be understandable to quality stakeholders. In a production environment, I would validate the business rules with Quality leadership, add access controls and auditability, and connect the workflow to an approved source system.

## Why This Project Matters

The project demonstrates the ability to combine **domain knowledge, data transformation, analytical logic, and business communication**—a combination relevant to quality analytics, operations analytics, business analysis, process improvement, and manufacturing analytics roles.
