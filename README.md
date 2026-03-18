# CAPA Closure Assistant
**A Python tool for CAPA triage, closure readiness, and audit visibility**

## Overview
The **CAPA Closure Assistant** is a Python-based analytics tool that helps regulated manufacturing and laboratory teams **prioritize, track, and close CAPAs more effectively**.

It converts CAPA tracker exports into **audit‑ready insights** by identifying overdue items, missing closure elements, and aging trends that commonly lead to repeat audit findings.

> ⚠️ All data used in this repository is **mocked and anonymized** to protect company confidentiality.  
> The structure, logic, and workflows reflect real quality‑system practices.

---

## Problem
In regulated environments, CAPAs frequently remain open or are closed prematurely due to:
- Incomplete root cause analysis
- Missing verification or effectiveness checks
- Manual tracking that does not scale
- Limited visibility into closure readiness

These gaps increase audit risk and drive repeat findings.

---

## Solution
This tool provides a **data‑driven approach** to CAPA management by:
- Calculating CAPA age, due‑soon, and overdue status
- Flagging missing closure requirements
- Applying a risk‑based triage score
- Producing audit‑friendly reports and summaries

---

## Key Outputs
**Inputs:** The tool consumes a mocked and anonymized CAPA export (CSV or Excel) containing standard fields such as CAPA ID, owner, status, dates, root cause, actions, and verification/effectiveness indicators.

These outputs support CAPA review meetings, audit preparation, and leadership updates.

---

## How to Run
```bash
capa --input mock_capa_data.xlsx --outdir outputs
```
