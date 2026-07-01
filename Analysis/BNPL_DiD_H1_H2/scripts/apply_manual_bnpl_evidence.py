#!/usr/bin/env python3
"""Apply manual BNPL adoption evidence to the public-company evidence table."""

from __future__ import annotations

import argparse
import csv
import datetime as dt
from collections import defaultdict
from pathlib import Path


EXCLUDED_CONFIRMATION_STRENGTHS = {
    "planned_not_final_adoption",
    "expansion_not_initial",
    "not_merchant_adoption",
    "strong_non_us",
}

CORE_STATUSES = {
    "confirmed_or_manually_mapped_public",
    "auto_public_match_needs_light_review",
}


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, rows: list[dict[str, str]], fieldnames: list[str]) -> None:
    with path.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def parse_date(value: str) -> dt.date:
    value = (value or "").strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"unsupported date format: {value!r}")


def display_date(value: str) -> str:
    parsed = parse_date(value)
    return f"{parsed.year}/{parsed.month}/{parsed.day}"


def base_research_note(value: str) -> str:
    marker = " Manual evidence:"
    if marker in value:
        return value.split(marker, 1)[0].rstrip()
    return (value or "").rstrip()


def is_confirmable(evidence: dict[str, str]) -> bool:
    return evidence.get("evidence_strength", "") not in EXCLUDED_CONFIRMATION_STRENGTHS


def choose_confirmed(evidence_rows: list[dict[str, str]]) -> dict[str, str] | None:
    candidates = [row for row in evidence_rows if is_confirmable(row)]
    if not candidates:
        return None
    return sorted(
        candidates,
        key=lambda row: (
            parse_date(row["evidence_date"]),
            row.get("evidence_strength", ""),
            row.get("provider", ""),
        ),
    )[0]


def manual_note(evidence_rows: list[dict[str, str]]) -> str:
    if not evidence_rows:
        return ""
    pieces = []
    for row in sorted(evidence_rows, key=lambda item: (parse_date(item["evidence_date"]), item["provider"])):
        pieces.append(
            "{provider} {date} {date_type} [{strength}]: {url}".format(
                provider=row["provider"],
                date=row["evidence_date"],
                date_type=row["date_type"],
                strength=row["evidence_strength"],
                url=row["source_url"],
            )
        )
    return " Manual evidence: " + " | ".join(pieces)


def apply_evidence(
    rows: list[dict[str, str]], evidence_by_key: dict[str, list[dict[str, str]]]
) -> list[dict[str, str]]:
    enriched = []
    for row in rows:
        row = dict(row)
        evidence_rows = evidence_by_key.get(row["merchant_key"], [])
        chosen = choose_confirmed(evidence_rows)
        for col in (
            "confirmed_bnpl_adoption_date",
            "confirmed_bnpl_provider",
            "confirmed_date_type",
            "confirmed_evidence_strength",
            "confirmed_evidence_source_type",
            "confirmed_evidence_url",
            "confirmed_evidence_quote_or_note",
        ):
            row[col] = ""
        if chosen:
            row["confirmed_bnpl_adoption_date"] = display_date(chosen["evidence_date"])
            row["confirmed_bnpl_provider"] = chosen["provider"]
            row["confirmed_date_type"] = chosen["date_type"]
            row["confirmed_evidence_strength"] = chosen["evidence_strength"]
            row["confirmed_evidence_source_type"] = chosen["source_type"]
            row["confirmed_evidence_url"] = chosen["source_url"]
            row["confirmed_evidence_quote_or_note"] = chosen["note"]
        row["research_notes"] = base_research_note(row.get("research_notes", "")) + manual_note(evidence_rows)
        enriched.append(row)
    return enriched


def write_notes(path: Path, full_rows: list[dict[str, str]], core_rows: list[dict[str, str]], full_name: str, core_name: str) -> None:
    full_confirmed = sum(1 for row in full_rows if row["confirmed_bnpl_adoption_date"])
    core_confirmed = sum(1 for row in core_rows if row["confirmed_bnpl_adoption_date"])
    weak_count = sum(1 for row in full_rows if row["weak_provider_directory_first_observed_date"])
    false_matches = sum(1 for row in full_rows if row["public_match_audit_status"] == "exclude_false_public_match")
    path.write_text(
        "\n".join(
            [
                "# BNPL adoption-date evidence notes",
                "",
                "Generated from the prior enriched public-company evidence table plus `work/manual_bnpl_press_evidence.csv`.",
                "",
                "## Output files",
                "",
                f"- `{full_name}`",
                f"  - {len(full_rows)} rows, matching the workbook's public analysis sample.",
                "  - Includes public-match audit status, confirmed/first-observed adoption evidence, weak provider-directory evidence, source URLs, and notes.",
                f"- `{core_name}`",
                f"  - {len(core_rows)} core rows after keeping manually mapped/confirmed public rows and multi-provider automatic matches.",
                "  - Excludes obvious false public-company name collisions and single-provider candidates needing public-match review.",
                "",
                "## Evidence columns",
                "",
                "- `confirmed_bnpl_adoption_date`",
                "  - Filled only when a usable official announcement, credible report, or Wayback official merchant/help/payment page was found.",
                "  - Check `confirmed_date_type` before using it as an event date.",
                "- `confirmed_date_type`",
                "  - `official_*`: stronger adoption/availability announcement.",
                "  - `first_observed_official_page`: Wayback first-observed date on an official merchant/help/payment page; this is not necessarily the true launch date.",
                "  - `reported_*`: credible news report; use as medium-strength evidence.",
                "- `weak_provider_directory_first_observed_date`",
                "  - Provider directory first-observed date from Wayback/Availability API.",
                "  - Do not use this alone as formal adoption date.",
                "",
                "## Current coverage",
                "",
                f"- {len(full_rows)} total public-analysis-sample rows.",
                f"- {len(core_rows)} core rows.",
                f"- {full_confirmed} full-sample rows have confirmed or official-page first-observed BNPL dates.",
                f"- {core_confirmed} core rows have confirmed or official-page first-observed BNPL dates.",
                f"- {weak_count} rows have weak provider-directory first-observed dates.",
                f"- {false_matches} rows are marked obvious false public-company matches.",
                "",
            ]
        ),
        encoding="utf-8",
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--base-full",
        default="outputs/bnpl_public_company_adoption_evidence_enriched_20260701_154326.csv",
    )
    parser.add_argument("--manual", default="work/manual_bnpl_press_evidence.csv")
    parser.add_argument("--output-dir", default="outputs")
    args = parser.parse_args()

    base_full = Path(args.base_full)
    manual_path = Path(args.manual)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    base_rows = read_csv(base_full)
    manual_rows = read_csv(manual_path)
    evidence_by_key: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in manual_rows:
        evidence_by_key[row["merchant_key"]].append(row)

    enriched_rows = apply_evidence(base_rows, evidence_by_key)
    fieldnames = list(enriched_rows[0].keys())
    core_rows = [row for row in enriched_rows if row["public_match_audit_status"] in CORE_STATUSES]

    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    full_path = output_dir / f"bnpl_public_company_adoption_evidence_enriched_{stamp}.csv"
    core_path = output_dir / f"bnpl_public_company_adoption_evidence_core_enriched_{stamp}.csv"
    notes_path = output_dir / f"bnpl_adoption_date_methodology_notes_{stamp}.md"

    write_csv(full_path, enriched_rows, fieldnames)
    write_csv(core_path, core_rows, fieldnames)
    write_notes(notes_path, enriched_rows, core_rows, full_path.name, core_path.name)

    print(f"wrote {full_path}")
    print(f"wrote {core_path}")
    print(f"wrote {notes_path}")


if __name__ == "__main__":
    main()
