"""Core per-file processing logic, shared by GUI and headless modes."""

import os
from typing import Dict, Optional, Set

from src.core.handler import PPLXHandler
from src.core.logic import (
    analyze_mr_note_for_aux_data,
    extract_scid_from_filename,
    clean_scid_keywords,
)
from src.core.utils import safe_filename_part

POLE_TAG_BLANK = "NO TAG"


def process_single_file(
    file_path: str,
    condition_value: str,
    output_dir: str,
    excel_data: Optional[Dict],
    valid_scids: Set[str],
    auto_fill_aux1: bool,
    auto_fill_aux2: bool,
    keyword_payload: Dict,
) -> Dict:
    """
    Process a single PPLX file.
    Returns {"status": "success"|"skipped"|"failed", "logs": [...], "csv_row": dict|None}
    """
    logs = []
    filename = os.path.basename(file_path)

    try:
        scid = extract_scid_from_filename(filename)
        clean_pole_number = clean_scid_keywords(scid)

        if excel_data and scid not in valid_scids:
            logs.append(f"Skipping {filename}: SCID '{scid}' not found in Excel data")
            return {"status": "skipped", "logs": logs, "csv_row": None}

        logs.append(
            f"Processing: {filename} (SCID: {scid}, Pole Number: {scid} -> {clean_pole_number})"
        )

        handler = PPLXHandler(file_path)

        _set_aux(handler, 3, condition_value, logs, prefix="Auto-set")

        pole_tag = POLE_TAG_BLANK
        mr_note = ""

        if excel_data and scid in excel_data:
            row_data = excel_data[scid]

            if auto_fill_aux1:
                pole_owner = row_data.get("pole_tag_company", "")
                if pole_owner:
                    _set_aux(handler, 1, pole_owner, logs, prefix="Auto-filled")

            if auto_fill_aux2:
                excel_pole_tag = row_data.get("pole_tag_tagtext", "").strip()
                pole_tag = excel_pole_tag if excel_pole_tag else POLE_TAG_BLANK
                _set_aux(handler, 2, pole_tag, logs, prefix="Auto-filled")
            else:
                _set_aux(handler, 2, pole_tag, logs, prefix="Set (manual)")

            mr_note = row_data.get("mr_note", "")
            aux_data_4, aux_data_5 = analyze_mr_note_for_aux_data(mr_note, **keyword_payload)
            _set_aux(handler, 4, aux_data_4, logs, prefix="Auto-filled")
            if mr_note:
                logs.append(
                    f"    Based on mr_note: {mr_note[:50]}{'...' if len(mr_note) > 50 else ''}"
                )
            _set_aux(handler, 5, aux_data_5, logs, prefix="Auto-filled")
        else:
            _set_aux(handler, 2, pole_tag, logs, prefix="Set (fallback)")

        if excel_data and scid in excel_data:
            if auto_fill_aux2:
                pole_tag = excel_data[scid].get("pole_tag_tagtext", pole_tag)
            mr_note = excel_data[scid].get("mr_note", mr_note)

        final_aux = handler.get_aux_data()
        aux_data_4 = final_aux.get("Aux Data 4", "")
        if aux_data_4 == "PCO":
            clean_pole_number = f"{clean_pole_number} PCO"
            logs.append(
                f"  Aux Data 4 is 'PCO', appending to pole number: {clean_pole_number}"
            )

        clean_pole_number_safe = safe_filename_part(
            clean_pole_number, ". " if aux_data_4 == "PCO" else ""
        )
        new_filename = (
            f"{clean_pole_number_safe}_{safe_filename_part(pole_tag, ' ')}"
            f"_{safe_filename_part(condition_value)}.pplx"
        )
        output_file = os.path.join(output_dir, new_filename)

        description_override = os.path.splitext(new_filename)[0]
        handler.set_pole_attribute("Pole Number", clean_pole_number)
        logs.append(f"  Set Pole Number: {clean_pole_number}")
        handler.set_pole_attribute("DescriptionOverride", description_override)
        logs.append(f"  Set DescriptionOverride: {description_override}")
        handler.save_file(output_file)
        logs.append(f"  Saved: {os.path.basename(output_file)}")

        csv_row = {
            "File Name": filename,
            "MR Note": mr_note,
            "Aux Data 1": final_aux.get("Aux Data 1", "Unset"),
            "Aux Data 2": final_aux.get("Aux Data 2", "Unset"),
            "Aux Data 3": final_aux.get("Aux Data 3", "Unset"),
            "Aux Data 4": final_aux.get("Aux Data 4", "Unset"),
            "Aux Data 5": final_aux.get("Aux Data 5", "Unset"),
        }

    except Exception as e:
        logs.append(f"  Error processing {filename}: {str(e)}")
        return {"status": "failed", "logs": logs, "csv_row": None}

    return {"status": "success", "logs": logs, "csv_row": csv_row}


def _set_aux(handler, aux_num: int, value: str, logs: list, prefix: str = "Set") -> bool:
    """Set aux data on handler and append a log entry."""
    success = handler.set_aux_data(aux_num, value)
    action = "ERROR: Failed to set" if not success else prefix
    logs.append(f"  {action} Aux Data {aux_num}: {value}")
    return success
