#!/usr/bin/env python3
"""
Lease Audit Multi-Agent Orchestrator
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Agent 1 (agent_011CaFZQQXJDwn8okgf5pR3i): Extract data from lease PDFs
Agent 2 (agent_011CaFZ6gN3WQfpwbtET4ZKQ): Compare extracted data to rent roll

Usage:
  python3 orchestrator.py --rent-roll path/to/rent_roll.xlsx
  python3 orchestrator.py --rent-roll path/to/rent_roll.xlsx --limit 5
  python3 orchestrator.py --rent-roll path/to/rent_roll.xlsx --all
"""

import os
import sys
import json
import argparse
import re
import time
from pathlib import Path
from datetime import datetime

import anthropic
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Agent IDs ─────────────────────────────────────────────────────────────────
AGENT_1_ID = "agent_011CaFZQQXJDwn8okgf5pR3i"   # Lease extractor
AGENT_2_ID = "agent_011CaFZ6gN3WQfpwbtET4ZKQ"    # Rent roll comparator

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).parent
RESIDENT_DOCS = BASE_DIR / "Resident_Documents_04-01-2026_14_59_28"
OUTPUT_DIR    = BASE_DIR / "audit_output"
ENV_ID_FILE   = BASE_DIR / ".audit_env_id"

# ── Valid lease subfolders (anything outside these triggers a warning) ─────────
LEASE_SUBFOLDERS = {"Lease Documents", "Signed Lease Documents"}

# ── Rent roll: names/stems that are NOT a valid rent roll ─────────────────────
BLOCKED_RENT_ROLL_NAMES = {"box score", "box_score", "application", "screening"}
VALID_RENT_ROLL_EXTENSIONS = {".csv", ".xlsx", ".xls"}


# ══════════════════════════════════════════════════════════════════════════════
# File validation
# ══════════════════════════════════════════════════════════════════════════════

def validate_rent_roll(path: Path) -> None:
    """Exits with a clear message if the provided file is not a valid rent roll."""
    ext = path.suffix.lower()
    name_lower = path.stem.lower()

    if not path.exists():
        _wrong_file(f"File not found: {path}")

    if ext not in VALID_RENT_ROLL_EXTENSIONS:
        _wrong_file(
            f"Wrong file found: '{path.name}'\n"
            f"  Expected a rent roll (.csv / .xlsx / .xls) but got '{ext}' file.\n"
            f"  Please provide your property rent roll, not this file."
        )

    for blocked in BLOCKED_RENT_ROLL_NAMES:
        if blocked in name_lower:
            _wrong_file(
                f"Wrong file found: '{path.name}'\n"
                f"  This looks like a '{path.stem}' report, not a rent roll.\n"
                f"  Please provide your property rent roll file instead."
            )


def find_lease_pdfs(resident_dir: Path) -> tuple[list[Path], list[str]]:
    """
    Returns (valid_lease_pdfs, warnings).
    Warns about any files found outside the approved lease subfolders.
    """
    lease_pdfs = []
    warnings = []

    for item in sorted(resident_dir.rglob("*")):
        if not item.is_file():
            continue

        # Determine which subfolder this file lives in
        try:
            rel = item.relative_to(resident_dir)
            top_subfolder = rel.parts[0] if len(rel.parts) > 1 else None
        except ValueError:
            top_subfolder = None

        if top_subfolder in LEASE_SUBFOLDERS:
            # Must be a PDF
            if item.suffix.lower() == ".pdf":
                lease_pdfs.append(item)
            else:
                warnings.append(
                    f"  ⚠  Wrong file found in lease folder: '{item.name}' "
                    f"(expected PDF, got '{item.suffix or 'no extension'}')"
                )
        # Files outside lease folders are silently ignored — they're intentional
        # (applications, IDs, income docs, screening results, etc.)

    return lease_pdfs, warnings


def _wrong_file(msg: str) -> None:
    print(f"\n{'!'*60}")
    print(f"  WRONG FILE FOUND")
    print(f"{'!'*60}")
    print(f"  {msg}")
    print(f"{'!'*60}\n")
    sys.exit(1)


# ══════════════════════════════════════════════════════════════════════════════
# Environment (one-time setup, cached)
# ══════════════════════════════════════════════════════════════════════════════

def get_or_create_env(client: anthropic.Anthropic) -> str:
    if ENV_ID_FILE.exists():
        env_id = ENV_ID_FILE.read_text().strip()
        print(f"  Using cached environment: {env_id}")
        return env_id
    print("  Creating new cloud environment...")
    env = client.beta.environments.create(
        name="lease-audit-env",
        config={"type": "cloud", "networking": {"type": "unrestricted"}},
    )
    ENV_ID_FILE.write_text(env.id)
    print(f"  Environment created: {env.id}")
    return env.id


# ══════════════════════════════════════════════════════════════════════════════
# Agent sessions
# ══════════════════════════════════════════════════════════════════════════════

def upload_file(client: anthropic.Anthropic, path: Path) -> str:
    with open(path, "rb") as f:
        result = client.beta.files.upload(file=f)
    return result.id


RATE_LIMIT_RETRY_WAIT = 90   # seconds to wait before retrying a rate-limited session
MAX_RETRIES = 3


def run_agent_session(
    client: anthropic.Anthropic,
    agent_id: str,
    env_id: str,
    prompt: str,
    resources: list[dict],
    label: str,
    log_fn: callable = print,
    stream_to_stdout: bool = True,
) -> str:
    for attempt in range(1, MAX_RETRIES + 1):
        session = client.beta.sessions.create(
            agent=agent_id,
            environment_id=env_id,
            resources=resources,
        )
        output_parts = []
        rate_limited = False

        with client.beta.sessions.events.stream(session_id=session.id) as stream:
            client.beta.sessions.events.send(
                session_id=session.id,
                events=[{
                    "type": "user.message",
                    "content": [{"type": "text", "text": prompt}],
                }],
            )
            for event in stream:
                if event.type == "agent.message":
                    for block in event.content:
                        if block.type == "text":
                            if stream_to_stdout:
                                sys.stdout.write(block.text)
                                sys.stdout.flush()
                            output_parts.append(block.text)
                elif event.type == "session.status_idle":
                    stop_type = getattr(event.stop_reason, "type", None)
                    if stop_type != "requires_action":
                        break
                elif event.type == "session.status_terminated":
                    break
                elif event.type == "session.error":
                    err_type = getattr(event.error, "type", "")
                    if err_type == "model_rate_limited_error":
                        rate_limited = True
                    else:
                        err_msg = getattr(event.error, "message", str(event))
                        log_fn(f"Error in {label}: {err_msg}")
                    break

        try:
            client.beta.sessions.archive(session_id=session.id)
        except Exception:
            pass

        if rate_limited:
            if attempt < MAX_RETRIES:
                log_fn(f"⏳ Rate limited — waiting {RATE_LIMIT_RETRY_WAIT}s then retrying ({attempt}/{MAX_RETRIES - 1})…")
                time.sleep(RATE_LIMIT_RETRY_WAIT)
                continue
            else:
                log_fn(f"✗ Rate limited after {MAX_RETRIES} attempts — giving up on {label}")
                return ""

        return "".join(output_parts)

    return ""


# ══════════════════════════════════════════════════════════════════════════════
# Per-resident pipeline
# ══════════════════════════════════════════════════════════════════════════════

def process_resident(
    client: anthropic.Anthropic,
    env_id: str,
    rent_roll_file_id: str,
    resident_dir: Path,
) -> dict:
    name = resident_dir.name
    print(f"\n{'─'*60}")
    print(f"  Resident: {name}")

    lease_pdfs, file_warnings = find_lease_pdfs(resident_dir)

    for w in file_warnings:
        print(w)

    if not lease_pdfs:
        print("  ⚠  No lease PDFs found in 'Lease Documents' or 'Signed Lease Documents' — skipping")
        return {"resident": name, "skipped": True, "reason": "no lease PDFs"}

    # Upload only the primary lease PDF (100-page API limit per session)
    primary_pdf = lease_pdfs[0]
    if len(lease_pdfs) > 1:
        print(f"  ℹ  {len(lease_pdfs)} lease PDFs found — uploading primary: {primary_pdf.name}")
    resources_a1 = []
    uploaded_ids = []
    file_id = upload_file(client, primary_pdf)
    uploaded_ids.append(file_id)
    resources_a1.append({
        "type": "file",
        "file_id": file_id,
        "mount_path": f"/workspace/{primary_pdf.name}",
    })
    print(f"  ↑ Uploaded lease: {primary_pdf.name}")

    # ── Agent 1: extract lease terms ──────────────────────────────────────────
    print(f"\n  [Agent 1] Extracting lease data...")
    a1_prompt = (
        f"Resident: {name}\n\n"
        "Extract all key lease terms from the uploaded PDF(s). "
        "Return ONLY a JSON object with these fields:\n"
        "  unit_number, resident_name, lease_start, lease_end,\n"
        "  monthly_rent, security_deposit, late_fee, pet_fee,\n"
        "  parking_fee, utility_charges, other_charges (array of {name, amount}),\n"
        "  total_monthly_charges, special_concessions (array of {description, amount})\n\n"
        "Use null for any field not found. Do not include any explanation, only JSON."
    )
    extraction = run_agent_session(
        client, AGENT_1_ID, env_id, a1_prompt, resources_a1, label="Agent1"
    )

    for fid in uploaded_ids:
        try:
            client.beta.files.delete(fid)
        except Exception:
            pass

    # Skip Agent 2 if Agent 1 produced no usable extraction
    if not extraction.strip() or "error" in extraction.lower()[:50]:
        print("\n\n  ⚠  Agent 1 produced no output — skipping Agent 2")
        return {"resident": name, "extraction": extraction, "comparison": ""}

    # Brief cooldown between Agent 1 and Agent 2
    print("  ⏳ Cooling down 60s before Agent 2...")
    time.sleep(60)

    # ── Agent 2: compare to rent roll ─────────────────────────────────────────
    print(f"\n  [Agent 2] Comparing to rent roll...")
    a2_prompt = (
        f"Resident: {name}\n\n"
        f"=== Lease Extraction ===\n{extraction}\n\n"
        "=== Task ===\n"
        "The file at /workspace/rent_roll is the property rent roll. "
        "Find this resident's record and compare every charge and date "
        "from the lease extraction against what the rent roll shows.\n\n"
        "Return ONLY a JSON object with:\n"
        "  resident (string),\n"
        "  unit_number (string),\n"
        "  discrepancies (array — each item: {field, lease_value, rent_roll_value, "
        "severity (high|medium|low), notes}),\n"
        "  summary (one sentence — or 'No discrepancies found')\n\n"
        "Do not include any explanation outside the JSON."
    )
    rent_roll_resource = [{
        "type": "file",
        "file_id": rent_roll_file_id,
        "mount_path": "/workspace/rent_roll",
    }]
    comparison = run_agent_session(
        client, AGENT_2_ID, env_id, a2_prompt, rent_roll_resource, label="Agent2"
    )

    print("\n")
    return {
        "resident": name,
        "extraction": extraction,
        "comparison": comparison,
    }


# ══════════════════════════════════════════════════════════════════════════════
# Excel report
# ══════════════════════════════════════════════════════════════════════════════

# Colour palette
CLR_HEADER    = "1F3864"   # dark navy
CLR_HIGH      = "FF4444"   # red
CLR_MEDIUM    = "FFA500"   # orange
CLR_LOW       = "FFD700"   # yellow
CLR_OK        = "70AD47"   # green
CLR_ROW_ALT   = "EEF2F7"   # light blue-grey alternate row
CLR_TAB_HEAD  = "2E75B6"   # mid-blue subheader

SEVERITY_CLR  = {"high": CLR_HIGH, "medium": CLR_MEDIUM, "low": CLR_LOW}


def _scalar(v):
    """Coerce any value to an Excel-safe type (str/int/float/None)."""
    if v is None:
        return ""
    if isinstance(v, (dict, list)):
        return json.dumps(v)
    return v


def _hdr_font(bold=True, color="FFFFFF", size=11):
    return Font(bold=bold, color=color, size=size)


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def _thin_border() -> Border:
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def _set_col_widths(ws, widths: dict[str, int]) -> None:
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


def _extract_json(text: str) -> dict:
    """Pull the first JSON object out of agent output text."""
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass
    return {}


def _parse_discrepancies(comparison_text: str) -> tuple[dict, list[dict]]:
    """Returns (comparison_obj, discrepancy_list)."""
    obj = _extract_json(comparison_text)
    discs = obj.get("discrepancies", [])
    if not isinstance(discs, list):
        discs = []
    return obj, discs


def _safe_excel(val):
    """Convert a value to an Excel-compatible type (str/int/float/bool or empty string)."""
    if val is None:
        return ""
    if isinstance(val, (str, int, float, bool)):
        return val
    return str(val)


def build_excel_report(results: list[dict], output_path: Path) -> None:
    wb = openpyxl.Workbook()

    # ── Sheet 1: Summary ──────────────────────────────────────────────────────
    ws_sum = wb.active
    ws_sum.title = "Summary"

    summary_headers = [
        "Resident", "Unit", "High", "Medium", "Low", "Total Issues", "Status", "Summary"
    ]
    ws_sum.append(summary_headers)
    for col_idx, _ in enumerate(summary_headers, 1):
        cell = ws_sum.cell(row=1, column=col_idx)
        cell.font = _hdr_font()
        cell.fill = _fill(CLR_HEADER)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _thin_border()
    ws_sum.row_dimensions[1].height = 28

    for row_idx, result in enumerate(results, 2):
        is_alt = (row_idx % 2 == 0)
        alt_fill = _fill(CLR_ROW_ALT) if is_alt else None

        if result.get("skipped") or result.get("error"):
            status = "Skipped" if result.get("skipped") else "Error"
            row_data = [_safe_excel(result["resident"]), "", 0, 0, 0, 0, status,
                        _safe_excel(result.get("reason") or result.get("error", ""))]
        else:
            comp_obj, discs = _parse_discrepancies(result.get("comparison", ""))
            high   = sum(1 for d in discs if str(d.get("severity","")).lower() == "high")
            medium = sum(1 for d in discs if str(d.get("severity","")).lower() == "medium")
            low    = sum(1 for d in discs if str(d.get("severity","")).lower() == "low")
            total  = len(discs)
            status = "Issues Found" if total > 0 else "Clean"
            summary = _safe_excel(comp_obj.get("summary", ""))
            unit    = _safe_excel(comp_obj.get("unit_number", ""))
            row_data = [_safe_excel(result["resident"]), unit, high, medium, low, total, status, summary]

        ws_sum.append(row_data)

        for col_idx, value in enumerate(row_data, 1):
            cell = ws_sum.cell(row=row_idx, column=col_idx)
            cell.border = _thin_border()
            cell.alignment = Alignment(vertical="center", wrap_text=(col_idx == 8))
            if alt_fill:
                cell.fill = alt_fill

            # Colour status cell
            if col_idx == 7 and isinstance(value, str):
                if value == "Clean":
                    cell.font = Font(bold=True, color=CLR_OK)
                elif value == "Issues Found":
                    cell.font = Font(bold=True, color=CLR_HIGH)
                else:
                    cell.font = Font(italic=True, color="888888")

            # Colour severity counts
            if col_idx == 3 and isinstance(value, int) and value > 0:
                cell.font = Font(bold=True, color=CLR_HIGH)
            if col_idx == 4 and isinstance(value, int) and value > 0:
                cell.font = Font(bold=True, color=CLR_MEDIUM)
            if col_idx == 5 and isinstance(value, int) and value > 0:
                cell.font = Font(bold=True, color="B8860B")

    _set_col_widths(ws_sum, {
        "A": 30, "B": 10, "C": 8, "D": 10, "E": 8, "F": 14, "G": 16, "H": 55
    })
    ws_sum.freeze_panes = "A2"

    # ── Sheet 2: Discrepancies ────────────────────────────────────────────────
    ws_disc = wb.create_sheet("Discrepancies")

    disc_headers = ["Resident", "Unit", "Field", "Lease Value",
                    "Rent Roll Value", "Severity", "Notes"]
    ws_disc.append(disc_headers)
    for col_idx, _ in enumerate(disc_headers, 1):
        cell = ws_disc.cell(row=1, column=col_idx)
        cell.font = _hdr_font()
        cell.fill = _fill(CLR_TAB_HEAD)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = _thin_border()
    ws_disc.row_dimensions[1].height = 24

    disc_row = 2
    for result in results:
        if result.get("skipped") or result.get("error"):
            continue
        comp_obj, discs = _parse_discrepancies(result.get("comparison", ""))
        unit = comp_obj.get("unit_number", "")
        for d in discs:
            sev = str(d.get("severity", "")).lower()
            row_data = [
                result["resident"],
                unit,
                d.get("field", ""),
                d.get("lease_value", ""),
                d.get("rent_roll_value", ""),
                sev.capitalize(),
                d.get("notes", ""),
            ]
            ws_disc.append(row_data)
            sev_color = SEVERITY_CLR.get(sev, "FFFFFF")
            for col_idx, _ in enumerate(row_data, 1):
                cell = ws_disc.cell(row=disc_row, column=col_idx)
                cell.border = _thin_border()
                cell.alignment = Alignment(vertical="center", wrap_text=(col_idx == 7))
                if disc_row % 2 == 0:
                    cell.fill = _fill(CLR_ROW_ALT)
            # Highlight severity cell
            sev_cell = ws_disc.cell(row=disc_row, column=6)
            sev_cell.fill = _fill(sev_color)
            sev_cell.font = Font(bold=True, color="FFFFFF" if sev in ("high", "medium") else "333333")
            sev_cell.alignment = Alignment(horizontal="center", vertical="center")
            disc_row += 1

    _set_col_widths(ws_disc, {
        "A": 28, "B": 8, "C": 22, "D": 20, "E": 20, "F": 12, "G": 45
    })
    ws_disc.freeze_panes = "A2"

    # ── Sheet 3: Lease Extractions ────────────────────────────────────────────
    ws_ext = wb.create_sheet("Lease Extractions")

    ext_headers = [
        "Resident", "Unit", "Resident Name",
        "Lease Start", "Lease End",
        "Monthly Rent", "Security Deposit", "Late Fee",
        "Pet Fee", "Parking Fee", "Utility Charges",
        "Total Monthly", "Concessions",
    ]
    ws_ext.append(ext_headers)
    for col_idx, _ in enumerate(ext_headers, 1):
        cell = ws_ext.cell(row=1, column=col_idx)
        cell.font = _hdr_font()
        cell.fill = _fill(CLR_TAB_HEAD)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _thin_border()
    ws_ext.row_dimensions[1].height = 28

    for row_idx, result in enumerate(results, 2):
        if result.get("skipped") or result.get("error"):
            continue
        ext = _extract_json(result.get("extraction", ""))
        concessions = ext.get("special_concessions") or []
        conc_str = "; ".join(
            f"{c.get('description','')} ({c.get('amount','')})" for c in concessions
        ) if isinstance(concessions, list) else str(concessions)

        row_data = [
            result["resident"],
            _scalar(ext.get("unit_number")),
            _scalar(ext.get("resident_name")),
            _scalar(ext.get("lease_start")),
            _scalar(ext.get("lease_end")),
            _scalar(ext.get("monthly_rent")),
            _scalar(ext.get("security_deposit")),
            _scalar(ext.get("late_fee")),
            _scalar(ext.get("pet_fee")),
            _scalar(ext.get("parking_fee")),
            _scalar(ext.get("utility_charges")),
            _scalar(ext.get("total_monthly_charges")),
            conc_str,
        ]
        ws_ext.append(row_data)
        for col_idx, _ in enumerate(row_data, 1):
            cell = ws_ext.cell(row=row_idx, column=col_idx)
            cell.border = _thin_border()
            cell.alignment = Alignment(vertical="center", wrap_text=(col_idx == 13))
            if row_idx % 2 == 0:
                cell.fill = _fill(CLR_ROW_ALT)

    _set_col_widths(ws_ext, {
        "A": 28, "B": 8, "C": 22, "D": 13, "E": 13,
        "F": 14, "G": 16, "H": 12, "I": 10, "J": 12,
        "K": 14, "L": 14, "M": 35,
    })
    ws_ext.freeze_panes = "A2"

    # ── Metadata tab ──────────────────────────────────────────────────────────
    ws_meta = wb.create_sheet("Run Info")
    ws_meta.append(["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")])
    ws_meta.append(["Agent 1 (Extractor)", AGENT_1_ID])
    ws_meta.append(["Agent 2 (Comparator)", AGENT_2_ID])
    ws_meta.append(["Residents processed", sum(1 for r in results if not r.get("skipped") and not r.get("error"))])
    ws_meta.append(["Skipped", sum(1 for r in results if r.get("skipped"))])
    ws_meta.append(["Errors", sum(1 for r in results if r.get("error"))])
    ws_meta.column_dimensions["A"].width = 24
    ws_meta.column_dimensions["B"].width = 40

    wb.save(output_path)


# ══════════════════════════════════════════════════════════════════════════════
# Main
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description="Lease audit multi-agent orchestrator")
    parser.add_argument("--rent-roll", required=True, metavar="FILE",
                        help="Path to the property rent roll file (.csv or .xlsx)")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--limit", type=int, default=3,
                       help="Process first N residents (default 3 for testing)")
    group.add_argument("--all", action="store_true", help="Process all residents")
    args = parser.parse_args()

    rent_roll_path = Path(args.rent_roll)
    validate_rent_roll(rent_roll_path)          # exits with message if wrong file

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        sys.exit("Error: ANTHROPIC_API_KEY environment variable not set.")

    OUTPUT_DIR.mkdir(exist_ok=True)
    client = anthropic.Anthropic(api_key=api_key)

    print("═" * 60)
    print("  Lease Audit Orchestrator")
    print("═" * 60)
    print(f"  Rent roll : {rent_roll_path.name}")

    print("\n[Setup]")
    env_id = get_or_create_env(client)

    print(f"  Uploading rent roll...")
    rent_roll_file_id = upload_file(client, rent_roll_path)
    print(f"  Rent roll uploaded: {rent_roll_file_id}")

    resident_dirs = sorted(d for d in RESIDENT_DOCS.iterdir() if d.is_dir())
    limit = None if args.all else args.limit
    if limit:
        resident_dirs = resident_dirs[:limit]

    print(f"\n[Processing {len(resident_dirs)} resident(s)]")

    results = []
    for i, resident_dir in enumerate(resident_dirs):
        try:
            result = process_resident(client, env_id, rent_roll_file_id, resident_dir)
            results.append(result)
        except Exception as exc:
            print(f"\n  ✗ Error: {exc}")
            results.append({"resident": resident_dir.name, "error": str(exc)})

        if i < len(resident_dirs) - 1:
            print("  ⏳ Pausing 30s to avoid rate limits...")
            time.sleep(30)

    try:
        client.beta.files.delete(rent_roll_file_id)
    except Exception:
        pass

    # Save JSON
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_path  = OUTPUT_DIR / f"audit_results_{ts}.json"
    excel_path = OUTPUT_DIR / f"Lease_Audit_Report_{ts}.xlsx"

    with open(json_path, "w") as f:
        json.dump(results, f, indent=2)

    print("Building Excel report...")
    build_excel_report(results, excel_path)

    processed = sum(1 for r in results if not r.get("skipped") and not r.get("error"))
    skipped   = sum(1 for r in results if r.get("skipped"))
    errors    = sum(1 for r in results if r.get("error"))

    print("\n" + "═" * 60)
    print("  DONE")
    print("═" * 60)
    print(f"  Processed  : {processed}")
    print(f"  Skipped    : {skipped}")
    print(f"  Errors     : {errors}")
    print(f"  JSON       : {json_path}")
    print(f"  Excel      : {excel_path}")
    print("═" * 60)


if __name__ == "__main__":
    main()
