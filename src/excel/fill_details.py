"""
Create PPLX_Fill_Details.xlsx from PPLX files and Excel data.
"""

import glob
import os
import shutil
from pathlib import Path

from src.core.handler import PPLXHandler
from src.core.logic import (
    analyze_mr_note_for_aux_data,
    extract_scid_from_filename,
    DEFAULT_AUX_VALUES,
)
from src.core.utils import parse_keywords
from src.config.manager import PPLXConfigManager
from src.excel.loader import load_excel_data

try:
    from openpyxl import Workbook
    from openpyxl.worksheet.table import Table, TableStyleInfo
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


def load_keyword_settings(config_filename: str = None) -> tuple:
    """Load keyword overrides from config profile."""
    mgr = PPLXConfigManager(config_name=config_filename)
    return (
        parse_keywords(mgr.get("comm_keywords", "")),
        parse_keywords(mgr.get("power_keywords", "")),
        parse_keywords(mgr.get("pco_keywords", "")),
        parse_keywords(mgr.get("aux5_keywords", "")),
        mgr.get("power_label", "POWER"),
    )


def _load_excel_mappings(excel_file_path: str) -> tuple:
    """Load Excel data using shared loader and build MR note/aux mappings."""
    mr_note_mapping = {}
    full_data_mapping = {}
    data = load_excel_data(excel_file_path)
    if not data:
        return mr_note_mapping, full_data_mapping

    for scid, row in data.items():
        scid_str = str(scid).strip()
        if not scid_str:
            continue
        mr_note = str(row.get("mr_note", "") or "").strip()
        row_data = {
            "pole_tag_company": row.get("pole_tag_company") or "MVEC",
            "pole_tag_tagtext": row.get("pole_tag_tagtext") or "",
            "mr_note": mr_note,
        }
        padded = scid_str.zfill(3) if scid_str.isdigit() else scid_str
        for key in (scid_str, padded):
            mr_note_mapping[key] = mr_note
            full_data_mapping[key] = row_data

    return mr_note_mapping, full_data_mapping


def create_pplx_excel(
    source_pplx_dir: str = "pplx_files",
    modified_pplx_dir: str = "pplx_files/Modified PPLX",
    excel_file_path: str = "pplx_files/MNMT002 v2 Nodes-Sections-Connections XLSX.xlsx",
    output_excel: str = "PPLX_Fill_Details.xlsx",
) -> None:
    """Create PPLX_Fill_Details.xlsx from filtered PPLX files and Excel data."""
    if not OPENPYXL_AVAILABLE:
        print("Error: openpyxl is required for create_pplx_excel")
        return

    if os.path.exists(modified_pplx_dir):
        shutil.rmtree(modified_pplx_dir)
    os.makedirs(modified_pplx_dir, exist_ok=True)

    print("Loading Excel data...")
    mr_note_mapping, full_data_mapping = _load_excel_mappings(excel_file_path)
    if not mr_note_mapping:
        print("No filtered Excel data found.")
        return

    valid_scids = set(mr_note_mapping.keys())
    pplx_files = sorted(glob.glob(os.path.join(source_pplx_dir, "*.pplx")))
    comm_keywords, power_keywords, pco_keywords, aux5_keywords, power_label = load_keyword_settings()

    csv_data = []
    processed_count = 0
    skipped_count = 0
    updated_count = 0

    for pplx_file in pplx_files:
        filename = os.path.basename(pplx_file)
        file_number = extract_scid_from_filename(filename)
        if file_number not in valid_scids:
            skipped_count += 1
            continue

        processed_count += 1
        modified_path = os.path.join(modified_pplx_dir, filename)
        shutil.copy2(pplx_file, modified_path)
        mr_note = mr_note_mapping.get(file_number, "")
        row_data = full_data_mapping.get(file_number, {})

        # Determine aux data values
        aux_updates = {}
        aux_updates["Aux Data 1"] = row_data.get("pole_tag_company", DEFAULT_AUX_VALUES["Aux Data 1"])
        pole_tag = row_data.get("pole_tag_tagtext", "")
        if not pole_tag or pole_tag.strip() == "" or str(pole_tag).lower() == "nan":
            aux_updates["Aux Data 2"] = DEFAULT_AUX_VALUES["Aux Data 2"]
        else:
            aux_updates["Aux Data 2"] = str(pole_tag).strip()
        aux_updates["Aux Data 3"] = DEFAULT_AUX_VALUES["Aux Data 3"]
        aux_data_4, aux_data_5 = analyze_mr_note_for_aux_data(
            mr_note,
            comm_keywords=comm_keywords,
            power_keywords=power_keywords,
            pco_keywords=pco_keywords,
            aux5_keywords=aux5_keywords,
            power_label=power_label,
        )
        aux_updates["Aux Data 4"] = aux_data_4
        aux_updates["Aux Data 5"] = aux_data_5

        # Use PPLXHandler to update the file — no duplicate XML code
        handler = PPLXHandler(modified_path)
        was_updated = False
        for aux_name, aux_val in aux_updates.items():
            num = int(aux_name.split()[-1])
            if handler.set_aux_data(num, aux_val):
                was_updated = True
        if was_updated and handler.save_file():
            updated_count += 1

        aux_data = handler.get_aux_data()
        csv_data.append(
            {
                "File Name": filename,
                "Aux Data 1": aux_data.get("Aux Data 1", "Unset"),
                "Aux Data 2": aux_data.get("Aux Data 2", "Unset"),
                "Aux Data 3": aux_data.get("Aux Data 3", "Unset"),
                "Aux Data 4": aux_data.get("Aux Data 4", "Unset"),
                "Aux Data 5": aux_data.get("Aux Data 5", "Unset"),
                "mr_note": mr_note,
            }
        )

    if csv_data:
        wb = Workbook()
        ws = wb.active
        ws.title = "PPLX Fill Details"
        headers = [
            "File Name",
            "Aux Data 1",
            "Aux Data 2",
            "Aux Data 3",
            "Aux Data 4",
            "Aux Data 5",
            "mr_note",
        ]
        ws.append(headers)
        for row in csv_data:
            ws.append([row.get(h, "") for h in headers])
        table_range = f"A1:{chr(65 + len(headers) - 1)}{len(csv_data) + 1}"
        table = Table(displayName="PPLXData", ref=table_range)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=True,
        )
        ws.add_table(table)
        for column in ws.columns:
            max_length = max(
                (len(str(cell.value)) for cell in column if cell.value is not None),
                default=0,
            )
            ws.column_dimensions[column[0].column_letter].width = min(
                max_length + 2, 50
            )
        wb.save(output_excel)
        print(f"\nExcel file created: {output_excel}")
        print(f"Processed: {processed_count}, Skipped: {skipped_count}, Updated: {updated_count}")
    else:
        print("No data to write.")
