"""
Create PPLX_Fill_Details.xlsx from PPLX files and Excel data.
"""

import glob
import json
import os
import shutil
from pathlib import Path

import xml.etree.ElementTree as ET

from src.core.utils import parse_keywords

try:
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.worksheet.table import Table, TableStyleInfo
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

from src.core.logic import (
    determine_aux_data_values,
    extract_scid_from_filename,
)
from src.config.manager import get_active_config_name


def extract_aux_data_from_pplx(pplx_file_path: str) -> dict:
    """Extract Aux Data 1-5 from a PPLX file."""
    aux_data = {
        "Aux Data 1": "Unset",
        "Aux Data 2": "Unset",
        "Aux Data 3": "Unset",
        "Aux Data 4": "Unset",
        "Aux Data 5": "Unset",
    }
    try:
        tree = ET.parse(pplx_file_path)
        root = tree.getroot()
        for wood_pole in root.findall(".//WoodPole"):
            attributes = wood_pole.find("ATTRIBUTES")
            if attributes is not None:
                for value in attributes.findall("VALUE"):
                    name = value.get("NAME")
                    if name and name.startswith("Aux Data ") and name in aux_data:
                        aux_data[name] = value.text or "Unset"
                break
    except Exception as e:
        print(f"Error processing {pplx_file_path}: {str(e)}")
    return aux_data


def load_keyword_settings(config_filename: str = None) -> tuple:
    """Load keyword overrides from config file in config/ folder."""
    root = Path(__file__).resolve().parents[2]
    config_name = config_filename or get_active_config_name()
    if not config_name.endswith(".json"):
        config_name = f"{config_name}.json"
    candidate_paths = [
        root / "config" / config_name,
        root / "config" / "OPPD.json",
    ]
    for path in candidate_paths:
        try:
            if path.is_file():
                with open(path, "r") as f:
                    config = json.load(f)
                return (
                    parse_keywords(config.get("comm_keywords")) or [],
                    parse_keywords(config.get("power_keywords")) or [],
                    parse_keywords(config.get("pco_keywords")) or [],
                    parse_keywords(config.get("aux5_keywords")) or [],
                    config.get("power_label", "POWER"),
                )
        except Exception:
            continue
    return [], [], [], [], "POWER"


def _load_excel_data_pandas(excel_file_path: str) -> tuple:
    """Load Excel data using pandas (for create_pplx_excel script)."""
    mr_note_mapping = {}
    full_data_mapping = {}
    if not PANDAS_AVAILABLE:
        return mr_note_mapping, full_data_mapping

    try:
        df = pd.read_excel(excel_file_path, sheet_name="nodes")
        required = ["scid", "node_type", "pole_status", "mr_note"]
        if any(c not in df.columns for c in required):
            return mr_note_mapping, full_data_mapping

        filtered = df[
            (df["node_type"].str.lower() == "pole")
            & (df["pole_status"].str.lower() != "underground")
        ]
        for _, row in filtered.iterrows():
            scid = str(row["scid"]).strip()
            mr_note = str(row["mr_note"]).strip() if pd.notna(row["mr_note"]) else ""
            padded = scid.zfill(3) if scid.isdigit() else scid
            row_data = {
                "pole_tag_company": str(row["pole_tag_company"])
                if pd.notna(row["pole_tag_company"])
                else "MVEC",
                "pole_tag_tagtext": str(row["pole_tag_tagtext"])
                if pd.notna(row["pole_tag_tagtext"])
                else "",
                "mr_note": mr_note,
            }
            for s in (scid, padded):
                mr_note_mapping[s] = mr_note
                full_data_mapping[s] = row_data
    except Exception as e:
        print(f"Error loading Excel: {e}")
    return mr_note_mapping, full_data_mapping


def update_pplx_aux_data(pplx_file_path: str, aux_data_updates: dict) -> bool:
    """Update multiple Aux Data fields in a PPLX file."""
    try:
        tree = ET.parse(pplx_file_path)
        root = tree.getroot()
        updated = False
        for wood_pole in root.findall(".//WoodPole"):
            attributes = wood_pole.find("ATTRIBUTES")
            if attributes is not None:
                for value in attributes.findall("VALUE"):
                    name = value.get("NAME")
                    if name in aux_data_updates:
                        value.text = aux_data_updates[name]
                        updated = True
                if updated:
                    tree.write(pplx_file_path, encoding="utf-8", xml_declaration=True)
                    return True
                break
    except Exception as e:
        print(f"Error updating {pplx_file_path}: {e}")
    return False


def create_pplx_excel(
    source_pplx_dir: str = "pplx_files",
    modified_pplx_dir: str = "pplx_files/Modified PPLX",
    excel_file_path: str = "pplx_files/MNMT002 v2 Nodes-Sections-Connections XLSX.xlsx",
    output_excel: str = "PPLX_Fill_Details.xlsx",
) -> None:
    """Create PPLX_Fill_Details.xlsx from filtered PPLX files and Excel data."""
    if not PANDAS_AVAILABLE:
        print("Error: pandas and openpyxl required for create_pplx_excel")
        return

    if os.path.exists(modified_pplx_dir):
        shutil.rmtree(modified_pplx_dir)
    os.makedirs(modified_pplx_dir, exist_ok=True)

    print("Loading Excel data...")
    mr_note_mapping, full_data_mapping = _load_excel_data_pandas(excel_file_path)
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

        aux_data_updates = determine_aux_data_values(
            file_number,
            mr_note,
            full_data_mapping,
            comm_keywords=comm_keywords,
            power_keywords=power_keywords,
            pco_keywords=pco_keywords,
            aux5_keywords=aux5_keywords,
            power_label=power_label,
        )
        if update_pplx_aux_data(modified_path, aux_data_updates):
            updated_count += 1

        aux_data = extract_aux_data_from_pplx(modified_path)
        csv_data.append(
            {
                "File Name": filename,
                "Aux Data 1": aux_data["Aux Data 1"],
                "Aux Data 2": aux_data["Aux Data 2"],
                "Aux Data 3": aux_data["Aux Data 3"],
                "Aux Data 4": aux_data["Aux Data 4"],
                "Aux Data 5": aux_data["Aux Data 5"],
                "mr_note": mr_note,
            }
        )

    if csv_data:
        df = pd.DataFrame(csv_data)
        wb = Workbook()
        ws = wb.active
        ws.title = "PPLX Fill Details"
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        table_range = f"A1:{chr(65 + len(df.columns) - 1)}{len(df) + 1}"
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
                (len(str(cell.value)) for cell in column),
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
