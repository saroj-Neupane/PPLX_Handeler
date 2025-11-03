#!/usr/bin/env python3
"""
Script to create PPLX_Fill_Details.xlsx from updated PPLX files and Excel data.
"""

import os
import xml.etree.ElementTree as ET
import pandas as pd
from pathlib import Path
import glob
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo

# Import shared configuration and logic
from pplx_config import (
    determine_aux_data_values,
    extract_scid_from_filename,
    clean_scid_keywords
)

def extract_aux_data_from_pplx(pplx_file_path):
    """
    Extract Aux Data 1-5 from a PPLX file.
    
    Args:
        pplx_file_path (str): Path to the PPLX file
        
    Returns:
        dict: Dictionary containing Aux Data 1-5 values
    """
    aux_data = {
        'Aux Data 1': 'Unset',
        'Aux Data 2': 'Unset',
        'Aux Data 3': 'Unset',
        'Aux Data 4': 'Unset',
        'Aux Data 5': 'Unset'
    }
    
    try:
        # Parse the XML file
        tree = ET.parse(pplx_file_path)
        root = tree.getroot()
        
        # Find WoodPole elements and extract Aux Data
        for wood_pole in root.findall('.//WoodPole'):
            attributes = wood_pole.find('ATTRIBUTES')
            if attributes is not None:
                for value in attributes.findall('VALUE'):
                    name = value.get('NAME')
                    if name and name.startswith('Aux Data ') and name in aux_data:
                        aux_data[name] = value.text or 'Unset'
                break  # Take the first WoodPole found
                
    except Exception as e:
        print(f"Error processing {pplx_file_path}: {str(e)}")
    
    return aux_data

def load_excel_data(excel_file_path):
    """
    Load Excel data with SCID filtering and create mappings for mr_note and full data.
    Only includes records where node_type='pole' and pole_status!='underground'
    
    Args:
        excel_file_path (str): Path to the Excel file
        
    Returns:
        tuple: (mr_note_mapping, full_data_mapping)
    """
    mr_note_mapping = {}
    full_data_mapping = {}
    
    try:
        # Read Excel file from 'nodes' sheet
        df = pd.read_excel(excel_file_path, sheet_name='nodes')
        
        # Print column names to help understand structure
        print("Excel file columns:", list(df.columns))
        
        # Check required columns for filtering
        required_columns = ['scid', 'node_type', 'pole_status', 'mr_note']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"Missing required columns: {missing_columns}")
            return mr_note_mapping, full_data_mapping
        
        # Apply filters: node_type = 'pole' AND pole_status != 'underground'
        filtered_df = df[
            (df['node_type'].str.lower() == 'pole') & 
            (df['pole_status'].str.lower() != 'underground')
        ]
        
        print(f"Total records in Excel: {len(df)}")
        print(f"Filtered records (pole, not underground): {len(filtered_df)}")
        
        # Create mappings using scid
        for _, row in filtered_df.iterrows():
            scid = str(row['scid']).strip()
            mr_note = str(row['mr_note']).strip() if pd.notna(row['mr_note']) else ''
            
            # Pad SCID with zeros to match filename format (001, 002, etc.)
            if scid.isdigit():
                padded_scid = scid.zfill(3)
                
                # mr_note mapping
                mr_note_mapping[padded_scid] = mr_note
                mr_note_mapping[scid] = mr_note
                
                # Full data mapping for aux data determination
                full_data_mapping[padded_scid] = {
                    'pole_tag_company': str(row['pole_tag_company']) if pd.notna(row['pole_tag_company']) else 'MVEC',
                    'pole_tag_tagtext': str(row['pole_tag_tagtext']) if pd.notna(row['pole_tag_tagtext']) else '',
                    'mr_note': mr_note
                }
                full_data_mapping[scid] = full_data_mapping[padded_scid]
            else:
                mr_note_mapping[scid] = mr_note
                full_data_mapping[scid] = {
                    'pole_tag_company': str(row['pole_tag_company']) if pd.notna(row['pole_tag_company']) else 'MVEC',
                    'pole_tag_tagtext': str(row['pole_tag_tagtext']) if pd.notna(row['pole_tag_tagtext']) else '',
                    'mr_note': mr_note
                }
    
    except Exception as e:
        print(f"Error loading Excel file: {str(e)}")
        print("Continuing without Excel data...")
    
    return mr_note_mapping, full_data_mapping


def update_pplx_aux_data(pplx_file_path, aux_data_updates):
    """
    Update multiple Aux Data fields in a PPLX file.
    
    Args:
        pplx_file_path (str): Path to the PPLX file
        aux_data_updates (dict): Dictionary of aux data field names to values
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Parse the XML file
        tree = ET.parse(pplx_file_path)
        root = tree.getroot()
        
        updated = False
        # Find WoodPole elements and update Aux Data
        for wood_pole in root.findall('.//WoodPole'):
            attributes = wood_pole.find('ATTRIBUTES')
            if attributes is not None:
                for value in attributes.findall('VALUE'):
                    name = value.get('NAME')
                    if name in aux_data_updates:
                        value.text = aux_data_updates[name]
                        updated = True
                
                if updated:
                    # Save the file
                    tree.write(pplx_file_path, encoding='utf-8', xml_declaration=True)
                    return True
                break  # Take the first WoodPole found
                
    except Exception as e:
        print(f"Error updating {pplx_file_path}: {str(e)}")
        return False
    
    return False


def create_pplx_excel():
    """
    Create PPLX_Fill_Details.xlsx from filtered PPLX files and Excel data.
    Only includes files with SCIDs that exist in filtered Excel data.
    Sets Aux Data 4 to 'NO MAKE READY' when mr_note is empty.
    """
    # Paths
    source_pplx_dir = "pplx_files"
    modified_pplx_dir = "pplx_files/Modified PPLX"
    excel_file_path = "pplx_files/MNMT002 v2 Nodes-Sections-Connections XLSX.xlsx"
    output_excel = "PPLX_Fill_Details.xlsx"
    
    # Create Modified PPLX directory if it doesn't exist, clear it if it does
    if os.path.exists(modified_pplx_dir):
        # Clear existing files in the directory
        import shutil
        shutil.rmtree(modified_pplx_dir)
    os.makedirs(modified_pplx_dir, exist_ok=True)
    
    # Load Excel data for mr_note mapping (already filtered)
    print("Loading Excel data...")
    mr_note_mapping, full_data_mapping = load_excel_data(excel_file_path)
    
    if not mr_note_mapping:
        print("No filtered Excel data found. Cannot proceed.")
        return
    
    # Get valid SCIDs from Excel data
    valid_scids = set(mr_note_mapping.keys())
    print(f"Valid SCIDs from Excel: {len(valid_scids)}")
    
    # Get all PPLX files from source directory
    pplx_files = glob.glob(os.path.join(source_pplx_dir, "*.pplx"))
    pplx_files.sort()
    
    print(f"Found {len(pplx_files)} PPLX files to process")
    
    # Prepare CSV data
    csv_data = []
    processed_count = 0
    skipped_count = 0
    updated_count = 0
    
    for pplx_file in pplx_files:
        filename = os.path.basename(pplx_file)
        file_number = extract_scid_from_filename(filename)
        
        # Check if SCID is in filtered Excel data
        if file_number not in valid_scids:
            print(f"Skipping {filename}: SCID '{file_number}' not in filtered Excel data")
            skipped_count += 1
            continue
        
        print(f"Processing {filename}...")
        processed_count += 1
        
        # Copy file to Modified PPLX directory
        import shutil
        modified_file_path = os.path.join(modified_pplx_dir, filename)
        shutil.copy2(pplx_file, modified_file_path)
        
        # Get mr_note from Excel mapping
        mr_note = mr_note_mapping.get(file_number, '')
        
        # Determine all Aux Data values based on Excel data and mr_note
        aux_data_updates = determine_aux_data_values(file_number, mr_note, full_data_mapping)
        
        # Update the PPLX file with new Aux Data values
        if update_pplx_aux_data(modified_file_path, aux_data_updates):
            print(f"  Updated Aux Data fields in {filename}")
            updated_count += 1
        
        # Extract final Aux Data from updated PPLX file for CSV
        aux_data = extract_aux_data_from_pplx(modified_file_path)
        
        # Create row data
        row_data = {
            'File Name': filename,
            'Aux Data 1': aux_data['Aux Data 1'],
            'Aux Data 2': aux_data['Aux Data 2'],
            'Aux Data 3': aux_data['Aux Data 3'],
            'Aux Data 4': aux_data['Aux Data 4'],
            'Aux Data 5': aux_data['Aux Data 5'],
            'mr_note': mr_note
        }
        
        csv_data.append(row_data)
    
    # Write Excel file
    if csv_data:
        # Convert to DataFrame
        df = pd.DataFrame(csv_data)
        
        # Create Excel workbook and worksheet
        wb = Workbook()
        ws = wb.active
        ws.title = "PPLX Fill Details"
        
        # Add data to worksheet
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        # Create table
        table_range = f"A1:{chr(65 + len(df.columns) - 1)}{len(df) + 1}"
        table = Table(displayName="PPLXData", ref=table_range)
        
        # Add table style
        style = TableStyleInfo(
            name="TableStyleMedium9", 
            showFirstColumn=False,
            showLastColumn=False, 
            showRowStripes=True, 
            showColumnStripes=True
        )
        table.tableStyleInfo = style
        ws.add_table(table)
        
        # Auto-fit columns
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # Auto-fit rows (set row height to auto)
        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = None
        
        # Save the workbook
        wb.save(output_excel)
        
        print(f"\nExcel file created successfully: {output_excel}")
        print(f"Processed files: {processed_count}")
        print(f"Skipped files (SCID not in Excel): {skipped_count}")
        print(f"PPLX files updated with proper Aux Data: {updated_count}")
        print(f"Total Excel rows: {len(csv_data)}")
        
        # Display first few rows as preview
        print("\nPreview of first 5 rows:")
        print("-" * 100)
        for i, row in enumerate(csv_data[:5]):
            print(f"Row {i+1}: {row}")
    else:
        print("No data to write to Excel file.")

if __name__ == "__main__":
    create_pplx_excel() 