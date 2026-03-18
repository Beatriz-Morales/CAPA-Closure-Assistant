"""
CAPA Closure Assistant — Mock Visual Generator
----------------------------------------------
Generates (synthetic/anonymized):
- mock_capa_data.xlsx + mock_capa_data.csv
- outputs/capa_triage.xlsx (multi-tab)
- outputs/weekly_update.txt
- assets/*.png visuals (dashboard + charts + table snapshots)

Dependencies:
  pip install pandas matplotlib openpyxl python-dateutil

Run:
  python create_mock_capa_visuals.py
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import List, Dict

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# Status lifecycle aligned with a typical CAPA workflow (PD/RC/CP/VE/CE/CX):
# PD=Problem Description, RC=Root Cause, CP=Corrective/Preventive Action,
# VE=Verification of Effectiveness, CE=Closed–Effective, CX=Canceled. [1](https://nestle.sharepoint.com/teams/NPPCSmallBusSys/SBS/_layouts/15/Doc.aspx?sourcedoc=%7BC077EF0E-F4CC-402C-AFC8-05C089454577%7D&file=CAPA%20Support.docx&action=default&mobileredirect=true&DefaultItemOpen=1)
STATUS_ORDER = ["PD", "RC", "CP", "VE", "CE", "CX"]
STATUS_LABELS = {
    "PD": "PD (Problem Description)",
    "RC": "RC (Root Cause)",
    "CP": "CP (Corrective/Preventive)",
    "VE": "VE (Verification of Effectiveness)",
    "CE": "CE (Closed – Effective)",
    "CX": "CX (Canceled)",
}

COLOR = {
    "OVERDUE": "#D9534F",
    "DUE_SOON": "#F0AD4E",
    "MISSING_INFO": "#5BC0DE",
    "OK": "#5CB85C",
    "NEUTRAL": "#4E5D6C",
}


# -----------------------------
# Config
# -----------------------------

@dataclass(frozen=True)
class Config:
    seed: int = 42
    n_records: int = 60
    due_soon_days: int = 14
    output_dir: str = "outputs"
    assets_dir: str = "assets"
    mock_data_xlsx: str = "mock_capa_data.xlsx"
    mock_data_csv: str = "mock_capa_data.csv"


# -----------------------------
# Helpers
# -----------------------------

def ensure_dirs(*dirs: str) -> None:
    for d in dirs:
        Path(d).mkdir(parents=True, exist_ok=True)


def today_ts() -> pd.Timestamp:
    return pd.Timestamp(date.today())


def fmt_date(d: pd.Timestamp | None) -> str:
    if d is None or pd.isna(d):
        return ""
    return pd.Timestamp(d).strftime("%Y-%m-%d")


def bucket_age(age_days: int) -> str:
    if age_days <= 14:
        return "0–14"
    if age_days <= 30:
        return "15–30"
    if age_days <= 60:
        return "31–60"
    return "60+"


def compute_blockers(row: pd.Series) -> str:
    blockers = []
    if row.get("root_cause_complete") == "N":
        blockers.append("Root Cause")
    if row.get("actions_complete") == "N":
        blockers.append("Actions")
    if row.get("verification_complete") == "N":
        blockers.append("Verification")
    if row.get("effectiveness_complete") == "N":
        blockers.append("Effectiveness")
    return ", ".join(blockers) if blockers else "—"


# -----------------------------
# Data generation
# -----------------------------

def generate_mock_data(cfg: Config) -> pd.DataFrame:
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

    # Bias statuses: most open items in PD/RC/CP/VE, fewer CE/CX
    status_probs = np.array([0.20, 0.22, 0.22, 0.20, 0.10, 0.06])
    status_probs = status_probs / status_probs.sum()
    statuses = np.random.choice(STATUS_ORDER, size=cfg.n_records, p=status_probs)

    # Created dates: last 90 days
    created_offsets = np.random.randint(1, 91, size=cfg.n_records)
    created_dates = [today - timedelta(days=int(x)) for x in created_offsets]

    # Due dates: created date + 15–60 days
    due_offsets = np.random.randint(15, 61, size=cfg.n_records)
    due_dates = [cd + timedelta(days=int(x)) for cd, x in zip(created_dates, due_offsets)]

    # Closed dates: only for CE and some CX
    closed_dates = []
    for st, cd, dd in zip(statuses, created_dates, due_dates):
        if st == "CE":
            delta = np.random.randint(-10, 11)  # around due date
            closed_dates.append(dd + timedelta(days=int(delta)))
        elif st == "CX":
            closed_dates.append(cd + timedelta(days=int(np.random.randint(3, 21))))
        else:
            closed_dates.append(None)

    def stage_flags(st: str) -> Dict[str, str]:
        if st == "PD":
            return dict(root_cause_complete="N", actions_complete="N", verification_complete="N", effectiveness_complete="N")
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
            return dict(root_cause_complete="Y", actions_complete="Y", verification_complete="Y", effectiveness_complete="Y")

        # CX (Canceled)
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


# -----------------------------
# Triage computation
# -----------------------------

def compute_triage(df: pd.DataFrame, cfg: Config) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = df.copy()
    today = today_ts()

    # Closed if status is CE/CX or closed_date exists
    df["is_closed"] = df["status"].isin(["CE", "CX"]) | df["closed_date"].notna()
    open_df = df[~df["is_closed"]].copy()

    open_df["age_days"] = (today - open_df["created_date"]).dt.days
    open_df["days_to_due"] = (open_df["due_date"] - today).dt.days

    open_df["is_overdue"] = open_df["days_to_due"] < 0
    open_df["is_due_soon"] = (open_df["days_to_due"] >= 0) & (open_df["days_to_due"] <= cfg.due_soon_days)

    open_df["missing_root_cause"] = (open_df["root_cause_complete"] == "N")
    open_df["missing_actions"] = (open_df["actions_complete"] == "N")
    open_df["missing_verification"] = (open_df["verification_complete"] == "N")
    open_df["missing_effectiveness"] = (open_df["effectiveness_complete"] == "N")

    open_df["missing_info_count"] = open_df[
        ["missing_root_cause", "missing_actions", "missing_verification", "missing_effectiveness"]
    ].sum(axis=1)

    open_df["triage_bucket"] = "OK"
    open_df.loc[open_df["missing_info_count"] > 0, "triage_bucket"] = "MISSING_INFO"
    open_df.loc[open_df["is_due_soon"], "triage_bucket"] = "DUE_SOON"
    open_df.loc[open_df["is_overdue"], "triage_bucket"] = "OVERDUE"

    open_df["triage_score"] = (
        open_df["is_overdue"].astype(int) * 3
        + open_df["is_due_soon"].astype(int) * 2
        + (open_df["missing_info_count"] > 0).astype(int) * 2
        + (open_df["age_days"] >= 30).astype(int) * 1
    )

    open_df["blockers"] = open_df.apply(compute_blockers, axis=1)
    open_df["age_bucket"] = open_df["age_days"].apply(bucket_age)

    open_df = open_df.sort_values(
        by=["triage_score", "is_overdue", "is_due_soon", "age_days"],
        ascending=[False, False, False, False]
    )

    return df, open_df


# -----------------------------
# Outputs (Excel + weekly text)
# -----------------------------

def write_excel_outputs(open_df: pd.DataFrame, cfg: Config) -> Path:
    out_path = Path(cfg.output_dir) / "capa_triage.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    overdue = open_df[open_df["triage_bucket"] == "OVERDUE"]
    due_soon = open_df[open_df["triage_bucket"] == "DUE_SOON"]
    missing = open_df[open_df["triage_bucket"] == "MISSING_INFO"]

    summary = pd.DataFrame({
        "metric": ["open_total", "overdue", "due_soon", "missing_info"],
        "count": [len(open_df), len(overdue), len(due_soon), len(missing)]
    })

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        open_df.to_excel(writer, index=False, sheet_name="Open")
        overdue.to_excel(writer, index=False, sheet_name="Overdue")
        due_soon.to_excel(writer, index=False, sheet_name="DueSoon")
        missing.to_excel(writer, index=False, sheet_name="MissingInfo")
        summary.to_excel(writer, index=False, sheet_name="Summary")

    return out_path


def write_weekly_update(open_df: pd.DataFrame, cfg: Config) -> Path:
    out_path = Path(cfg.output_dir) / "weekly_update.txt"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    total = len(open_df)
    overdue = int((open_df["triage_bucket"] == "OVERDUE").sum())
    due_soon = int((open_df["triage_bucket"] == "DUE_SOON").sum())
    missing = int((open_df["triage_bucket"] == "MISSING_INFO").sum())

    lines: List[str] = []
    lines.append("Weekly CAPA Triage Summary (Mock Data)")
    lines.append("=" * 36)
    lines.append(f"Open CAPAs: {total}")
    lines.append(f"Overdue: {overdue}")
    lines.append(f"Due soon (next {cfg.due_soon_days} days): {due_soon}")
    lines.append(f"Missing closure info: {missing}")
    lines.append("")

    if overdue > 0:
        lines.append("Overdue CAPAs by Owner (Top 5):")
        top = open_df[open_df["triage_bucket"] == "OVERDUE"]["owner"].value_counts().head(5)
        for owner, cnt in top.items():
            lines.append(f"  - {owner}: {cnt}")
        lines.append("")

    blockers = ["missing_root_cause", "missing_actions", "missing_verification", "missing_effectiveness"]
    miss_counts = open_df[blockers].sum().sort_values(ascending=False)
    lines.append("Most common closure blockers:")
    for k, v in miss_counts.items():
        nice = k.replace("missing_", "").replace("_", " ").title()
        lines.append(f"  - {nice}: {int(v)}")

    out_path.write_text("\n".join(lines), encoding="utf-8")
    return out_path


# -----------------------------
# Visuals
# -----------------------------

def save_kpi_tiles(open_df: pd.DataFrame, cfg: Config) -> Path:
    total = len(open_df)
    overdue = int((open_df["triage_bucket"] == "OVERDUE").sum())
    due_soon = int((open_df["triage_bucket"] == "DUE_SOON").sum())
    missing = int((open_df["triage_bucket"] == "MISSING_INFO").sum())

    fig, axes = plt.subplots(2, 2, figsize=(10, 6))
    fig.suptitle("CAPA Health Overview (Mock Data)", fontsize=14, fontweight="bold")

    tiles = [
        ("Open CAPAs", total, COLOR["NEUTRAL"]),
        ("Overdue", overdue, COLOR["OVERDUE"]),
        ("Due Soon (Next 14 Days)", due_soon, COLOR["DUE_SOON"]),
        ("Missing Closure Info", missing, COLOR["MISSING_INFO"]),
    ]

    for ax, (label, value, color) in zip(axes.flatten(), tiles):
        ax.axis("off")
        ax.add_patch(plt.Rectangle((0, 0), 1, 1, transform=ax.transAxes, color=color, alpha=0.18))
        ax.text(0.05, 0.72, label, fontsize=11, fontweight="bold", transform=ax.transAxes)
        ax.text(0.05, 0.22, str(value), fontsize=28, fontweight="bold", transform=ax.transAxes)
        ax.text(0.05, 0.05, "Mocked/anonymized dataset for portfolio use", fontsize=8, alpha=0.8, transform=ax.transAxes)

    plt.tight_layout(rect=[0, 0.02, 1, 0.93])
    out_path = Path(cfg.assets_dir) / "capa_health_overview.png"
    plt.savefig(out_path, dpi=200)
    plt.close(fig)
    return out_path


def save_age_distribution(open_df: pd.DataFrame, cfg: Config) -> Path:
    order = ["0–14", "15–30", "31–60", "60+"]
    counts = open_df["age_bucket"].value_counts().reindex(order).fillna(0).astype(int)

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.bar(counts.index, counts.values, color=COLOR["NEUTRAL"])
    ax.set_title("CAPA Age Distribution (Days Open)", fontweight="bold")
    ax.set_xlabel("Age Bucket (Days)")
    ax.set_ylabel("Number of CAPAs")
    ax.grid(axis="y", alpha=0.25)
    ax.text(0.99, -0.18, "Mock Data", transform=ax.transAxes, ha="right", va="top", fontsize=8, alpha=0.8)

    plt.tight_layout()
    out_path = Path(cfg.assets_dir) / "capa_age_distribution.png"
    plt.savefig(out_path, dpi=200)
    plt.close(fig)
    return out_path


def save_overdue_by_owner(open_df: pd.DataFrame, cfg: Config) -> Path:
    overdue = open_df[open_df["triage_bucket"] == "OVERDUE"]
    counts = overdue["owner"].value_counts().sort_values(ascending=True)

    fig, ax = plt.subplots(figsize=(9, 5))
    if len(counts) == 0:
        ax.text(0.5, 0.5, "No overdue CAPAs (Mock Data)", ha="center", va="center", fontsize=12)
        ax.axis("off")
    else:
        ax.barh(counts.index, counts.values, color=COLOR["OVERDUE"])
        ax.set_title("Overdue CAPAs by Owner", fontweight="bold")
        ax.set_xlabel("Number of Overdue CAPAs")
        ax.set_ylabel("Owner")
        ax.grid(axis="x", alpha=0.25)

    plt.tight_layout()
    out_path = Path(cfg.assets_dir) / "capa_overdue_by_owner.png"
    plt.savefig(out_path, dpi=200)
    plt.close(fig)
    return out_path


def save_status_pipeline(all_df: pd.DataFrame, cfg: Config) -> Path:
    counts = all_df["status"].value_counts().reindex(STATUS_ORDER).fillna(0).astype(int)

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.bar(counts.index, counts.values, color=COLOR["NEUTRAL"])
    ax.set_title("CAPA Status Pipeline (PD → RC → CP → VE → CE)", fontweight="bold")
    ax.set_xlabel("CAPA Status")
    ax.set_ylabel("Number of CAPAs")
    ax.grid(axis="y", alpha=0.25)

    mapping_text = "\n".join([f"{k}: {STATUS_LABELS[k]}" for k in STATUS_ORDER])
    ax.text(1.02, 0.98, mapping_text, transform=ax.transAxes, va="top", fontsize=8)

    plt.tight_layout()
    out_path = Path(cfg.assets_dir) / "capa_status_pipeline.png"
    plt.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return out_path


def save_table_image(
    df: pd.DataFrame,
    out_path: Path,
    title: str,
    col_order: List[str],
    max_rows: int = 12
) -> Path:
    view = df[col_order].head(max_rows).copy()

    for c in ["created_date", "due_date", "closed_date"]:
        if c in view.columns:
            view[c] = view[c].apply(fmt_date)

    fig, ax = plt.subplots(figsize=(12, 0.6 + 0.45 * len(view)))
    ax.axis("off")
    ax.set_title(title, fontweight="bold", pad=12)

    table = ax.table(
        cellText=view.values,
        colLabels=view.columns,
        cellLoc="left",
        loc="center"
    )
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 1.3)

    for (row, col), cell in table.get_celld().items():
        if row == 0:
            cell.set_text_props(fontweight="bold", color="white")
            cell.set_facecolor(COLOR["NEUTRAL"])
        else:
            cell.set_facecolor("#F7F7F7" if row % 2 == 0 else "white")

    plt.tight_layout()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return out_path


def build_closure_readiness(open_df: pd.DataFrame) -> pd.DataFrame:
    df = open_df.copy()
    df["ready_to_close"] = np.where(
        (df["status"] == "VE")
        & (df["root_cause_complete"] == "Y")
        & (df["actions_complete"] == "Y")
        & (df["verification_complete"] == "Y")
        & (df["effectiveness_complete"] == "Y"),
        "Y", "N"
    )

    cols = [
        "capa_id", "status",
        "root_cause_complete", "actions_complete",
