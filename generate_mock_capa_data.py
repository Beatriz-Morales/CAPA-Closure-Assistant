"""
Generate mock_capa_data.xlsx (synthetic/anonymized CAPA dataset)

Run:
  python generate_mock_capa_data.py
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from typing import Dict

import numpy as np
import pandas as pd


# CAPA status lifecycle (PD/RC/CP/VE/CE/CX) aligned to common CAPA workflows:
# PD = Problem Description
# RC = Root Cause
# CP = Corrective/Preventive Action
# VE = Verification of Effectiveness
# CE = Closed – Effective
# CX = Canceled
STATUS_ORDER = ["PD", "RC", "CP", "VE", "CE", "CX"]


@dataclass(frozen=True)
class MockDataConfig:
    seed: int = 42
    n_records: int = 60
    days_back_created: int = 90
    due_days_min: int = 15
    due_days_max: int = 60
    out_xlsx: str = "mock_capa_data.xlsx"
    out_csv: str = "mock_capa_data.csv"


def generate_mock_capa_data(cfg: MockDataConfig) -> pd.DataFrame:
    np.random.seed(cfg.seed)
    today = date.today()

    owners = ["Owner A", "Owner B", "Owner C", "Owner D", "Owner E"]
    sites = ["Site A", "Site B", "Site C"]
    areas = ["Micro Lab", "QA", "Packaging", "Warehouse", "Sanitation", "IT/QA", "Operations"]
    titles = [
        "Temp log gaps", "Label traceability", "Training record gaps", "Sampling SOP update",
        "Chemical ID gaps", "Data backup gaps", "Equipment list mismatch", "Sanitizer concentration",
        "Deviation documentation", "Hold/release timing", "Environmental monitoring gaps", "Calibration tracking gaps"
    ]

    capa_ids = [f"CAPA-{i:04d}" for i in range(1, cfg.n_records + 1)]

    status_probs = np.array([0.20, 0.22, 0.22, 0.20, 0.10, 0.06])
    status_probs = status_probs / status_probs.sum()
    statuses = np.random.choice(STATUS_ORDER, size=cfg.n_records, p=status_probs)

    created_offsets = np.random.randint(1, cfg.days_back_created + 1, size=cfg.n_records)
    created_dates = [today - timedelta(days=int(x)) for x in created_offsets]

    due_offsets = np.random.randint(cfg.due_days_min, cfg.due_days_max + 1, size=cfg.n_records)
    due_dates = [cd + timedelta(days=int(x)) for cd, x in zip(created_dates, due_offsets)]

    closed_dates = []
    for st, cd, dd in zip(statuses, created_dates, due_dates):
        if st == "CE":
            delta = np.random.randint(-10, 11)
            closed_dates.append(dd + timedelta(days=int(delta)))
        elif st == "CX":
            closed_dates.append(cd + timedelta(days=int(np.random.randint(3, 21))))
        else:
            closed_dates.append(None)

    def stage_flags(st: str) -> Dict[str, str]:
        if st == "PD":
            return dict(
                root_cause_complete="N",
                actions_complete="N",
                verification_complete="N",
                effectiveness_complete="N",
            )
        if st == "RC":
            return dict(
                root_cause_complete=np.random.choice(["Y", "N"], p=[0.65, 0.35]),
                actions_complete="N",
                verification_complete="N",
                effectiveness_complete="N",
            )
        if st == "CP":
            return dict(
                root_cause_complete="Y",
                actions_complete=np.random.choice(["Y", "N"], p=[0.70, 0.30]),
                verification_complete="N",
                effectiveness_complete="N",
            )
        if st == "VE":
            return dict(
                root_cause_complete="Y",
                actions_complete="Y",
                verification_complete=np.random.choice(["Y", "N"], p=[0.75, 0.25]),
                effectiveness_complete=np.random.choice(["Y", "N"], p=[0.55, 0.45]),
            )
        if st == "CE":
            return dict(
                root_cause_complete="Y",
                actions_complete="Y",
                verification_complete="Y",
                effectiveness_complete="Y",
            )

        return dict(
            root_cause_complete=np.random.choice(["Y", "N"], p=[0.40, 0.60]),
            actions_complete=np.random.choice(["Y", "N"], p=[0.30, 0.70]),
            verification_complete="N",
            effectiveness_complete="N",
        )

    flags = [stage_flags(st) for st in statuses]

    df = pd.DataFrame({
        "capa_id": capa_ids,
        "title": np.random.choice(titles, size=cfg.n_records),
        "site": np.random.choice(sites, size=cfg.n_records),
        "area": np.random.choice(areas, size=cfg.n_records),
        "owner": np.random.choice(owners, size=cfg.n_records),
        "status": statuses,
        "created_date": pd.to_datetime(created_dates),
        "due_date": pd.to_datetime(due_dates),
        "closed_date": pd.to_datetime(closed_dates),
        "root_cause_complete": [f["root_cause_complete"] for f in flags],
        "actions_complete": [f["actions_complete"] for f in flags],
        "verification_complete": [f["verification_complete"] for f in flags],
        "effectiveness_complete": [f["effectiveness_complete"] for f in flags],
    })

    return df


def main() -> None:
    cfg = MockDataConfig()
    df = generate_mock_capa_data(cfg)

    df.to_excel(cfg.out_xlsx, index=False, engine="openpyxl")
    df.to_csv(cfg.out_csv, index=False)

    print("✅ Mock CAPA dataset created")
    print(f"- {cfg.out_xlsx}")
    print(f"- {cfg.out_csv}")
    print(f"Rows: {len(df)} | Columns: {len(df.columns)}")


if __name__ == "__main__":
    main()
