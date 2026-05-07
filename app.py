#!/usr/bin/env python3
"""
Lease Audit Dashboard
Run with:  streamlit run app.py
"""

import json
import os
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import anthropic
import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))
from orchestrator import (
    AGENT_1_ID,
    AGENT_2_ID,
    BLOCKED_RENT_ROLL_NAMES,
    OUTPUT_DIR,
    VALID_RENT_ROLL_EXTENSIONS,
    _extract_json,
    _parse_discrepancies,
    build_excel_report,
    find_lease_pdfs,
    get_or_create_env,
    run_agent_session,
    upload_file,
)

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Lease Audit Dashboard",
    page_icon="🏠",
    layout="wide",
)

# ── Header ─────────────────────────────────────────────────────────────────────
st.title("🏠 Lease Audit Dashboard")
st.caption("AI-powered lease vs. rent roll comparison")

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    api_key = st.text_input(
        "Anthropic API Key",
        type="password",
        value=os.environ.get("ANTHROPIC_API_KEY", ""),
        help="Your Anthropic API key — never stored",
    )

    st.divider()
    st.subheader("Residents to process")
    limit_choice = st.radio("Mode", ["Test — first 3", "Custom", "All"], index=0)
    if limit_choice == "Custom":
        limit: int | None = st.number_input("Count", min_value=1, max_value=500, value=10)
    elif limit_choice == "All":
        limit = None
    else:
        limit = 3

    st.divider()
    st.caption("Agent IDs")
    st.code(f"Extract : {AGENT_1_ID}", language=None)
    st.code(f"Compare : {AGENT_2_ID}", language=None)

# ── File uploads ───────────────────────────────────────────────────────────────
# Cache uploaded file bytes in session_state so they survive the button-click rerun
col_a, col_b = st.columns(2)
with col_a:
    zip_upload = st.file_uploader(
        "📁 Lease Documents (ZIP)",
        type=["zip"],
        help="ZIP containing resident folders with 'Lease Documents' subfolders",
    )
    if zip_upload is not None:
        st.session_state["zip_bytes"] = zip_upload.getvalue()
        st.session_state["zip_name"]  = zip_upload.name

with col_b:
    rr_upload = st.file_uploader(
        "📊 Rent Roll",
        type=["xlsx", "xls", "csv"],
        help="Property rent roll — .xlsx, .xls, or .csv",
    )
    if rr_upload is not None:
        st.session_state["rr_bytes"] = rr_upload.getvalue()
        st.session_state["rr_name"]  = rr_upload.name

can_run = bool(
    st.session_state.get("zip_bytes") and
    st.session_state.get("rr_bytes") and
    api_key
)
run_btn = st.button(
    "▶  Run Audit",
    disabled=not can_run,
    type="primary",
    use_container_width=True,
)
if not api_key:
    st.info("Enter your Anthropic API key in the sidebar to get started.")


# ── Helper: find resident dirs inside extracted zip ────────────────────────────
def find_resident_dirs(extracted_root: Path) -> list[Path]:
    """
    Handles two common zip layouts:
      - zip root = resident folders directly
      - zip root = one dated folder (e.g. Resident_Documents_04-01-2026/) → go one level deeper
    """
    top = [
        d for d in sorted(extracted_root.iterdir())
        if d.is_dir() and d.name not in ("__MACOSX", ".DS_Store")
    ]
    if len(top) == 1:
        inner = [
            d for d in sorted(top[0].iterdir())
            if d.is_dir() and not d.name.startswith(".")
        ]
        if inner:
            return inner
    return top


# ── Main audit ─────────────────────────────────────────────────────────────────
if run_btn:
    # Pull files from session state (survives the button-click rerun)
    zip_bytes = st.session_state["zip_bytes"]
    rr_bytes  = st.session_state["rr_bytes"]
    rr_name   = st.session_state["rr_name"]

    # Validate rent roll filename before doing anything else
    rr_stem = Path(rr_name).stem.lower()
    rr_ext  = Path(rr_name).suffix.lower()
    if any(b in rr_stem for b in BLOCKED_RENT_ROLL_NAMES) or rr_ext not in VALID_RENT_ROLL_EXTENSIONS:
        st.error(
            f"⛔ **WRONG FILE FOUND**: '{rr_name}' doesn't look like a rent roll.  \n"
            f"Please upload the property rent roll (.xlsx / .xls / .csv)."
        )
        st.stop()

    client  = anthropic.Anthropic(api_key=api_key)
    results: list[dict] = []
    review_flags: list[dict] = []

    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp)

        # Write zip to disk then extract
        zip_path = tmp / "leases.zip"
        zip_path.write_bytes(zip_bytes)
        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(tmp / "leases")

        resident_dirs = find_resident_dirs(tmp / "leases")
        if not resident_dirs:
            st.error(
                "No resident folders found in the ZIP.  \n"
                "Expected folders that contain 'Lease Documents' or 'Signed Lease Documents' subfolders."
            )
            st.stop()

        if limit:
            resident_dirs = resident_dirs[:limit]

        # Save rent roll locally
        rr_path = tmp / rr_name
        rr_path.write_bytes(rr_bytes)

        st.success(f"✅ ZIP extracted — **{len(resident_dirs)}** resident(s) queued  |  Rent roll: {rr_name}")

        # Environment + rent roll upload ───────────────────────────────────────
        with st.status("Setting up…", expanded=False) as setup_box:
            st.write("Loading cloud environment…")
            env_id = get_or_create_env(client)
            st.write(f"Environment: `{env_id}`")
            st.write("Uploading rent roll…")
            rr_file_id = upload_file(client, rr_path)
            st.write(f"Rent roll uploaded: `{rr_file_id}`")
            setup_box.update(label="✅ Setup complete", state="complete")

        total_steps = len(resident_dirs) * 2   # Agent 1 + Agent 2 per resident
        progress = st.progress(0, text="Starting Phase 1…")
        step = 0

        # ══════════════════════════════════════════════════════════════════════
        # PHASE 1 — Agent 1: extract lease data for every resident
        # ══════════════════════════════════════════════════════════════════════
        st.subheader("Phase 1 — Lease Extraction")
        extractions: dict[str, dict] = {}   # name → {extraction, ext_obj, res_dir}

        for i, res_dir in enumerate(resident_dirs):
            name = res_dir.name
            progress.progress(step / total_steps, text=f"Phase 1 · {name} ({i + 1}/{len(resident_dirs)})…")

            with st.status(f"📄 {name}", expanded=True) as s1:
                log = st.write

                lease_pdfs, warnings = find_lease_pdfs(res_dir)
                for w in warnings:
                    log(f"⚠️ {w}")
                    review_flags.append({"resident": name, "issue": w, "type": "warning"})

                if not lease_pdfs:
                    log("⚠️ No lease PDFs found — skipping")
                    s1.update(label=f"⏭ {name} — No lease PDFs", state="error")
                    results.append({"resident": name, "skipped": True, "reason": "no lease PDFs"})
                    review_flags.append({"resident": name, "issue": "No lease PDFs in approved subfolders", "type": "skip"})
                    step += 2   # skip both agent steps
                    continue

                primary = lease_pdfs[0]
                if len(lease_pdfs) > 1:
                    log(f"ℹ️ {len(lease_pdfs)} PDFs found — uploading primary: {primary.name}")
                log(f"⬆️ Uploading {primary.name}…")
                fid = upload_file(client, primary)

                log("🤖 **Agent 1**: extracting lease terms…")
                a1_prompt = (
                    f"Resident: {name}\n\n"
                    "Extract all key lease terms from the uploaded PDF. "
                    "Return ONLY a JSON object with these fields: "
                    "unit_number, resident_name, lease_start, lease_end, monthly_rent, "
                    "security_deposit, late_fee, pet_fee, parking_fee, utility_charges, "
                    "other_charges (array of {name, amount}), total_monthly_charges, "
                    "special_concessions (array of {description, amount}). "
                    "Use null for any field not found. No explanation — only JSON."
                )
                extraction = run_agent_session(
                    client, AGENT_1_ID, env_id, a1_prompt,
                    [{"type": "file", "file_id": fid, "mount_path": f"/workspace/{primary.name}"}],
                    "Agent1",
                )

                try:
                    client.beta.files.delete(fid)
                except Exception:
                    pass

                if not extraction.strip():
                    log("❌ Agent 1 returned no output")
                    s1.update(label=f"❌ {name} — Extraction failed", state="error")
                    results.append({"resident": name, "extraction": "", "comparison": ""})
                    review_flags.append({"resident": name, "issue": "Agent 1 extraction failed — manual review needed", "type": "error"})
                    step += 2
                    continue

                ext_obj = _extract_json(extraction)
                log(f"✅ Done — {len(ext_obj)} fields extracted")
                with st.expander("Extracted lease data"):
                    st.json(ext_obj)

                s1.update(label=f"✅ {name} — Extraction complete", state="complete")
                extractions[name] = {"extraction": extraction, "ext_obj": ext_obj, "res_dir": res_dir}

            step += 1
            progress.progress(step / total_steps, text=f"Phase 1 · {name} complete")

        if extractions:
            st.info(f"✅ Phase 1 complete — {len(extractions)} lease(s) extracted. Starting Agent 2…")

        # ══════════════════════════════════════════════════════════════════════
        # PHASE 2 — Agent 2: compare each extraction to the rent roll
        # ══════════════════════════════════════════════════════════════════════
        st.subheader("Phase 2 — Rent Roll Comparison")
        extraction_list = list(extractions.items())

        for i, (name, data) in enumerate(extraction_list):
            extraction = data["extraction"]
            ext_obj    = data["ext_obj"]

            progress.progress(step / total_steps, text=f"Phase 2 · {name} ({i + 1}/{len(extraction_list)})…")

            with st.status(f"🔍 {name}", expanded=True) as s2:
                log = st.write

                log("🤖 **Agent 2**: comparing to rent roll…")
                a2_prompt = (
                    f"Resident: {name}\n\n"
                    f"=== Lease Extraction ===\n{extraction}\n\n"
                    "The file at /workspace/rent_roll is the property rent roll. "
                    "Find this resident's record and compare every charge and date "
                    "from the lease extraction against what the rent roll shows.\n\n"
                    "Return ONLY a JSON object: "
                    "{ resident, unit_number, "
                    "discrepancies: [{field, lease_value, rent_roll_value, "
                    "severity (high|medium|low), notes}], "
                    "summary }. No explanation outside the JSON."
                )
                comparison = run_agent_session(
                    client, AGENT_2_ID, env_id, a2_prompt,
                    [{"type": "file", "file_id": rr_file_id, "mount_path": "/workspace/rent_roll"}],
                    "Agent2",
                )

                comp_obj, discs = _parse_discrepancies(comparison)
                high_ct = sum(1 for d in discs if str(d.get("severity", "")).lower() == "high")
                med_ct  = sum(1 for d in discs if str(d.get("severity", "")).lower() == "medium")
                low_ct  = sum(1 for d in discs if str(d.get("severity", "")).lower() == "low")

                if not comparison.strip():
                    log("⚠️ Agent 2 returned no output — manual review needed")
                    s2.update(label=f"⚠️ {name} — Comparison failed", state="error")
                    review_flags.append({"resident": name, "issue": "Agent 2 comparison failed — manual review needed", "type": "error"})
                elif discs:
                    log(f"🚨 {len(discs)} discrepancy(ies) — High: {high_ct} / Med: {med_ct} / Low: {low_ct}")
                    if high_ct:
                        review_flags.append({"resident": name, "issue": f"{high_ct} high-severity issue(s): {comp_obj.get('summary', '')}", "type": "high"})
                    icon = "🚨" if high_ct else "⚠️"
                    s2.update(label=f"{icon} {name} — {len(discs)} issue(s)", state="error" if high_ct else "complete")
                else:
                    log("✅ No discrepancies found")
                    s2.update(label=f"✅ {name} — Clean", state="complete")

                with st.expander("Comparison result"):
                    st.json(comp_obj if comp_obj else {"note": "No output from Agent 2"})

                results.append({"resident": name, "extraction": extraction, "comparison": comparison})

            step += 1
            progress.progress(step / total_steps, text=f"Phase 2 · {name} complete")


        progress.progress(1.0, text="Audit complete!")

        # Cleanup rent roll file
        try:
            client.beta.files.delete(rr_file_id)
        except Exception:
            pass

        # Build Excel ─────────────────────────────────────────────────────────
        OUTPUT_DIR.mkdir(exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = OUTPUT_DIR / f"Lease_Audit_Report_{ts}.xlsx"
        build_excel_report(results, excel_path)

        # ── Summary section ────────────────────────────────────────────────────
        st.divider()
        st.header("📊 Audit Results")

        processed = sum(1 for r in results if not r.get("skipped") and not r.get("error"))
        skipped   = sum(1 for r in results if r.get("skipped"))

        all_disc_rows: list[dict] = []
        for r in results:
            if r.get("skipped") or r.get("error"):
                continue
            comp_obj, discs = _parse_discrepancies(r.get("comparison", ""))
            unit = comp_obj.get("unit_number", "")
            for d in discs:
                all_disc_rows.append({
                    "Resident":        r["resident"],
                    "Unit":            unit,
                    "Field":           d.get("field", ""),
                    "Lease Value":     str(d.get("lease_value", "")),
                    "Rent Roll Value": str(d.get("rent_roll_value", "")),
                    "Severity":        str(d.get("severity", "")).capitalize(),
                    "Notes":           d.get("notes", ""),
                })

        high_total = sum(1 for d in all_disc_rows if d["Severity"] == "High")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Residents Processed", processed)
        m2.metric("Total Discrepancies", len(all_disc_rows))
        m3.metric("High Severity",       high_total)
        m4.metric("Flagged for Review",  len(review_flags))

        if all_disc_rows:
            st.subheader("All Discrepancies")

            df = pd.DataFrame(all_disc_rows)

            def _color_row(row):
                palette = {
                    "High":   "background-color: #ffd6d6",
                    "Medium": "background-color: #ffecd1",
                    "Low":    "background-color: #fffde7",
                }
                c = palette.get(row["Severity"], "")
                return [c] * len(row)

            st.dataframe(
                df.style.apply(_color_row, axis=1),
                use_container_width=True,
                hide_index=True,
            )

        # ── Human review flags ─────────────────────────────────────────────────
        if review_flags:
            st.divider()
            st.subheader("🔍 Needs Human Review")
            st.caption("These residents require a manual check before finalising the audit.")

            for flag in review_flags:
                icon = {"high": "🚨", "error": "❌", "warning": "⚠️", "skip": "⏭"}.get(
                    flag["type"], "ℹ️"
                )
                msg = f"{icon} **{flag['resident']}** — {flag['issue']}"
                if flag["type"] in ("high", "error"):
                    st.error(msg)
                else:
                    st.warning(msg)

        # ── Download ───────────────────────────────────────────────────────────
        st.divider()
        with open(excel_path, "rb") as fh:
            st.download_button(
                "⬇️  Download Excel Report",
                data=fh.read(),
                file_name=excel_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
