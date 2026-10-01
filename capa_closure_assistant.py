"""CAPA Closure Assistant: analyze a synthetic or appropriately authorized CAPA tracker."""

from __future__ import annotations

from pathlib import Path
import argparse
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

REQUIRED = [
    "capa_id","title","site","area","owner","status","created_date","due_date","closed_date",
    "root_cause_complete","actions_complete","verification_complete","effectiveness_complete"
]
CLOSED_STATUSES = {"CE","CX"}

def load_data(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in {".xlsx",".xls"}:
        df = pd.read_excel(path)
    elif path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    else:
        raise ValueError("Input must be CSV or Excel.")
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")
    for c in ["created_date","due_date","closed_date"]:
        df[c] = pd.to_datetime(df[c], errors="coerce")
    if df["created_date"].isna().any() or df["due_date"].isna().any():
        raise ValueError("created_date and due_date must contain valid dates.")
    return df

def analyze(df: pd.DataFrame, due_soon_days: int) -> pd.DataFrame:
    out = df.copy()
    today = pd.Timestamp.today().normalize()
    out["is_closed"] = out["status"].isin(CLOSED_STATUSES) | out["closed_date"].notna()
    out = out.loc[~out["is_closed"]].copy()
    out["age_days"] = (today - out["created_date"]).dt.days
    out["days_to_due"] = (out["due_date"] - today).dt.days
    out["is_overdue"] = out["days_to_due"] < 0
    out["is_due_soon"] = out["days_to_due"].between(0, due_soon_days)

    mapping = {
        "root_cause_complete":"missing_root_cause",
        "actions_complete":"missing_actions",
        "verification_complete":"missing_verification",
        "effectiveness_complete":"missing_effectiveness",
    }
    for source, target in mapping.items():
        out[target] = out[source].astype(str).str.upper().ne("Y")
    miss_cols = list(mapping.values())
    out["missing_info_count"] = out[miss_cols].sum(axis=1)
    out["blockers"] = out.apply(
        lambda r: ", ".join([
            label for flag, label in [
                ("missing_root_cause","Root Cause"),("missing_actions","Actions"),
                ("missing_verification","Verification"),("missing_effectiveness","Effectiveness")
            ] if r[flag]
        ]) or "None", axis=1
    )

    out["triage_bucket"] = np.select(
        [out["is_overdue"], out["is_due_soon"], out["missing_info_count"].gt(0)],
        ["OVERDUE","DUE_SOON","MISSING_INFO"], default="OK"
    )
    out["triage_score"] = (
        out["is_overdue"].astype(int)*3 +
        out["is_due_soon"].astype(int)*2 +
        out["missing_info_count"].gt(0).astype(int)*2 +
        out["age_days"].ge(30).astype(int)
    )
    out["ready_to_close"] = (
        out["status"].eq("VE") &
        out[["root_cause_complete","actions_complete","verification_complete","effectiveness_complete"]]
        .apply(lambda s: s.astype(str).str.upper().eq("Y")).all(axis=1)
    )
    out["age_bucket"] = pd.cut(
        out["age_days"], bins=[-1,14,30,60,float("inf")],
        labels=["0–14","15–30","31–60","60+"]
    )
    return out.sort_values(["triage_score","age_days"], ascending=[False,False])

def export_outputs(open_df: pd.DataFrame, outdir: Path, due_soon_days: int) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    overdue = open_df[open_df["is_overdue"]]
    due_soon = open_df[open_df["is_due_soon"]]
    missing = open_df[open_df["missing_info_count"] > 0]
    ready = open_df[open_df["ready_to_close"]]

    summary = pd.DataFrame({
        "metric":["Open CAPAs","Overdue","Due Soon","Missing Closure Info","Closure Ready"],
        "count":[len(open_df),len(overdue),len(due_soon),len(missing),len(ready)]
    })
    with pd.ExcelWriter(outdir/"capa_triage.xlsx", engine="openpyxl") as w:
        open_df.to_excel(w,index=False,sheet_name="Open")
        overdue.to_excel(w,index=False,sheet_name="Overdue")
        due_soon.to_excel(w,index=False,sheet_name="DueSoon")
        missing.to_excel(w,index=False,sheet_name="MissingInfo")
        ready.to_excel(w,index=False,sheet_name="ClosureReady")
        summary.to_excel(w,index=False,sheet_name="Summary")

    blocker_cols = ["missing_root_cause","missing_actions","missing_verification","missing_effectiveness"]
    blockers = open_df[blocker_cols].sum().sort_values(ascending=False)
    lines = [
        "Weekly CAPA Triage Summary (Synthetic Portfolio Data)",
        "="*52,
        f"Open CAPAs: {len(open_df)}",
        f"Overdue: {len(overdue)}",
        f"Due soon (next {due_soon_days} days): {len(due_soon)}",
        f"Missing closure information: {len(missing)}",
        f"Closure-ready under demonstration rule: {len(ready)}",
        "",
        "Most common closure blockers:"
    ]
    for name, count in blockers.items():
        lines.append(f"- {name.replace('missing_','').replace('_',' ').title()}: {int(count)}")
    (outdir/"weekly_update.txt").write_text("\n".join(lines), encoding="utf-8")

def make_visuals(open_df: pd.DataFrame, assets: Path, due_soon_days: int) -> None:
    assets.mkdir(parents=True, exist_ok=True)
    metrics = {
        "Open CAPAs":len(open_df),
        "Overdue":int(open_df["is_overdue"].sum()),
        f"Due Soon ({due_soon_days}d)":int(open_df["is_due_soon"].sum()),
        "Missing Info":int((open_df["missing_info_count"]>0).sum())
    }
    fig, ax = plt.subplots(figsize=(9,5))
    ax.bar(metrics.keys(), metrics.values())
    ax.set_title("CAPA Health Overview — Synthetic Data")
    ax.set_ylabel("CAPA Count")
    plt.xticks(rotation=15, ha="right")
    plt.tight_layout()
    fig.savefig(assets/"capa_health_overview.png", dpi=180)
    plt.close(fig)

    counts = open_df["age_bucket"].value_counts(sort=False)
    fig, ax = plt.subplots(figsize=(8,5))
    ax.bar(counts.index.astype(str), counts.values)
    ax.set_title("Open CAPA Age Distribution — Synthetic Data")
    ax.set_xlabel("Days Open")
    ax.set_ylabel("CAPA Count")
    plt.tight_layout()
    fig.savefig(assets/"capa_age_distribution.png", dpi=180)
    plt.close(fig)

    overdue = open_df[open_df["is_overdue"]]["owner"].value_counts().sort_values()
    fig, ax = plt.subplots(figsize=(8,5))
    if overdue.empty:
        ax.text(.5,.5,"No overdue CAPAs",ha="center",va="center")
        ax.axis("off")
    else:
        ax.barh(overdue.index, overdue.values)
        ax.set_title("Overdue CAPAs by Owner — Synthetic Data")
        ax.set_xlabel("Overdue CAPA Count")
    plt.tight_layout()
    fig.savefig(assets/"capa_overdue_by_owner.png", dpi=180)
    plt.close(fig)

def main() -> None:
    p = argparse.ArgumentParser(description="CAPA triage and closure-readiness analysis.")
    p.add_argument("--input", required=True, type=Path)
    p.add_argument("--outdir", default=Path("outputs"), type=Path)
    p.add_argument("--assets", default=Path("assets"), type=Path)
    p.add_argument("--due-soon-days", default=14, type=int)
    args = p.parse_args()
    df = load_data(args.input)
    open_df = analyze(df, args.due_soon_days)
    export_outputs(open_df, args.outdir, args.due_soon_days)
    make_visuals(open_df, args.assets, args.due_soon_days)
    print(f"Analyzed {len(df)} records; {len(open_df)} are open.")
    print(f"Outputs: {args.outdir.resolve()}")
    print(f"Visuals: {args.assets.resolve()}")

if __name__ == "__main__":
    main()
