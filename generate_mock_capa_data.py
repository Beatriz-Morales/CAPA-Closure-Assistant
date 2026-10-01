"""Generate a synthetic CAPA dataset for the portfolio project."""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path
import argparse
import numpy as np
import pandas as pd

STATUS_ORDER = ["PD", "RC", "CP", "VE", "CE", "CX"]

def stage_flags(status: str, rng: np.random.Generator) -> dict[str, str]:
    if status == "PD":
        return {"root_cause_complete":"N","actions_complete":"N","verification_complete":"N","effectiveness_complete":"N"}
    if status == "RC":
        return {"root_cause_complete":rng.choice(["Y","N"], p=[.65,.35]),"actions_complete":"N","verification_complete":"N","effectiveness_complete":"N"}
    if status == "CP":
        return {"root_cause_complete":"Y","actions_complete":rng.choice(["Y","N"], p=[.70,.30]),"verification_complete":"N","effectiveness_complete":"N"}
    if status == "VE":
        return {"root_cause_complete":"Y","actions_complete":"Y",
                "verification_complete":rng.choice(["Y","N"], p=[.75,.25]),
                "effectiveness_complete":rng.choice(["Y","N"], p=[.55,.45])}
    if status == "CE":
        return {"root_cause_complete":"Y","actions_complete":"Y","verification_complete":"Y","effectiveness_complete":"Y"}
    return {"root_cause_complete":rng.choice(["Y","N"], p=[.40,.60]),
            "actions_complete":rng.choice(["Y","N"], p=[.30,.70]),
            "verification_complete":"N","effectiveness_complete":"N"}

def generate(n_records: int = 60, seed: int = 42) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    today = date.today()
    owners = ["Owner A","Owner B","Owner C","Owner D","Owner E"]
    sites = ["Site A","Site B","Site C"]
    areas = ["Micro Lab","QA","Packaging","Warehouse","Sanitation","IT/QA","Operations"]
    titles = ["Temperature log gaps","Label traceability","Training record gaps","Sampling SOP update",
              "Chemical identification gaps","Data backup gaps","Equipment list mismatch",
              "Sanitizer concentration","Deviation documentation","Hold/release timing",
              "Environmental monitoring gaps","Calibration tracking gaps"]

    statuses = rng.choice(STATUS_ORDER, n_records, p=[.20,.22,.22,.20,.10,.06])
    created = [today - timedelta(days=int(x)) for x in rng.integers(1, 91, n_records)]
    due = [c + timedelta(days=int(x)) for c, x in zip(created, rng.integers(15, 61, n_records))]
    closed = []
    for st, c, d in zip(statuses, created, due):
        if st == "CE":
            closed.append(d + timedelta(days=int(rng.integers(-10, 11))))
        elif st == "CX":
            closed.append(c + timedelta(days=int(rng.integers(3, 21))))
        else:
            closed.append(None)

    flags = [stage_flags(st, rng) for st in statuses]
    return pd.DataFrame({
        "capa_id":[f"CAPA-{i:04d}" for i in range(1,n_records+1)],
        "title":rng.choice(titles,n_records),
        "site":rng.choice(sites,n_records),
        "area":rng.choice(areas,n_records),
        "owner":rng.choice(owners,n_records),
        "status":statuses,
        "created_date":pd.to_datetime(created),
        "due_date":pd.to_datetime(due),
        "closed_date":pd.to_datetime(closed),
        "root_cause_complete":[x["root_cause_complete"] for x in flags],
        "actions_complete":[x["actions_complete"] for x in flags],
        "verification_complete":[x["verification_complete"] for x in flags],
        "effectiveness_complete":[x["effectiveness_complete"] for x in flags],
    })

def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--records", type=int, default=60)
    p.add_argument("--seed", type=int, default=42)
    p.add_argument("--prefix", default="mock_capa_data")
    args = p.parse_args()
    df = generate(args.records, args.seed)
    df.to_csv(f"{args.prefix}.csv", index=False)
    df.to_excel(f"{args.prefix}.xlsx", index=False)
    print(f"Created {len(df)} synthetic CAPA records.")

if __name__ == "__main__":
    main()
