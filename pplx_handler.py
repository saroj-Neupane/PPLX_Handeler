#!/usr/bin/env python3
"""
PPLX File Handler - XML Editor for Pole Line Engineering Files

This script allows you to read, analyze, and edit pplx files which are XML files
containing pole line engineering data. It provides specific functionality to edit
Aux Data fields and other attributes.

Author: Saroj Neupane
"""

import xml.etree.ElementTree as ET
import os
import glob
from typing import List, Dict, Optional, Tuple, Any
import json
import argparse
from pathlib import Path


class PPLXHandler:
    """Handler class for reading and editing PPLX XML files."""
    
    def __init__(self, file_path: Optional[str] = None):
        """
        Initialize the PPLX handler.
        
        Args:
            file_path (str, optional): Path to a specific pplx file
        """
        self.file_path = file_path
        self.tree: Optional[Any] = None
        self.root: Optional[Any] = None
        if file_path and os.path.exists(file_path):
            self.load_file(file_path)
    
    def load_file(self, file_path: str) -> bool:
        """
        Load a pplx XML file.
        
        Args:
            file_path (str): Path to the pplx file
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            self.tree = ET.parse(file_path)
            self.root = self.tree.getroot()
            self.file_path = file_path
            print(f"Successfully loaded: {file_path}")
            return True
        except ET.ParseError as e:
            print(f"Error parsing XML file {file_path}: {e}")
            return False
        except FileNotFoundError:
            print(f"File not found: {file_path}")
            return False
    
    def get_file_info(self) -> Dict:
        """
        Get basic information about the loaded pplx file.
        
        Returns:
            Dict: Dictionary containing file metadata
        """
        if not self.root:
            return {}
        
        info = {
            'file_path': self.file_path,
            'date': self.root.get('DATE', 'Unknown'),
            'user': self.root.get('USER', 'Unknown'),
            'workstation': self.root.get('WORKSTATION', 'Unknown')
        }
        
        # Get scene attributes
        scene = self.root.find('PPLScene')
        if scene:
            attributes = scene.find('ATTRIBUTES')
            if attributes:
                for value in attributes.findall('VALUE'):
                    name = value.get('NAME')
                    if name in ['Latitude', 'Longitude', 'ElevationMetersAboveMSL', 'PPLVersion']:
                        info[name] = value.text
        
        return info
    
    def find_wood_poles(self) -> List[ET.Element]:
        """
        Find all WoodPole elements in the file.
        
        Returns:
            List[ET.Element]: List of WoodPole elements
        """
        if not self.root:
            return []
        
        return self.root.findall('.//WoodPole')
    
    def get_aux_data(self, pole_element: ET.Element = None) -> Dict[str, str]:
        """
        Get all Aux Data values from a pole or the first pole if none specified.
        
        Args:
            pole_element (ET.Element, optional): Specific pole element
            
        Returns:
            Dict[str, str]: Dictionary of aux data values
        """
        if pole_element is None:
            poles = self.find_wood_poles()
            if not poles:
                return {}
            pole_element = poles[0]
        
        aux_data = {}
        attributes = pole_element.find('ATTRIBUTES')
        if attributes:
            for value in attributes.findall('VALUE'):
                name = value.get('NAME')
                if name and name.startswith('Aux Data'):
                    aux_data[name] = value.text or 'Unset'
        
        return aux_data
    
    def set_aux_data(self, aux_data_number: int, new_value: str, pole_element: ET.Element = None) -> bool:
        """
        Set the value of a specific Aux Data field.
        
        Args:
            aux_data_number (int): Number of the aux data field (1-8)
            new_value (str): New value to set
            pole_element (ET.Element, optional): Specific pole element
            
        Returns:
            bool: True if successful, False otherwise
        """
        if not 1 <= aux_data_number <= 8:
            print(f"Error: Aux Data number must be between 1 and 8, got {aux_data_number}")
            return False
        
        if pole_element is None:
            poles = self.find_wood_poles()
            if not poles:
                print("Error: No WoodPole elements found")
                return False
            pole_element = poles[0]
        
        aux_data_name = f"Aux Data {aux_data_number}"
        attributes = pole_element.find('ATTRIBUTES')
        success = False
        owner_updated = False
        
        if attributes:
            # Update Aux Data field
            for value in attributes.findall('VALUE'):
                if value.get('NAME') == aux_data_name:
                    old_value = value.text or 'Unset'
                    value.text = new_value
                    print(f"Updated {aux_data_name}: '{old_value}' -> '{new_value}'")
                    success = True
                    break

            # If Aux Data field not found, create it
            if not success:
                aux_element = ET.SubElement(attributes, 'VALUE')
                aux_element.set('NAME', aux_data_name)
                aux_element.set('TYPE', 'String')
                aux_element.text = new_value
                print(f"Created {aux_data_name}: '{new_value}'")
                success = True

            # If this is Aux Data 1, also update the Owner field
            if aux_data_number == 1 and success:
                for value in attributes.findall('VALUE'):
                    if value.get('NAME') == 'Owner':
                        old_owner = value.text or 'Unset'
                        value.text = new_value
                        print(f"Updated Owner field to match Aux Data 1: '{old_owner}' -> '{new_value}'")
                        owner_updated = True
                        break
                
                # If Owner field wasn't found, create it
                if not owner_updated:
                    owner_element = ET.SubElement(attributes, 'VALUE')
                    owner_element.set('NAME', 'Owner')
                    owner_element.set('TYPE', 'String')
                    owner_element.text = new_value
                    print(f"Created Owner field with value: '{new_value}'")
                    owner_updated = True
        
        if not success:
            print(f"Error: Could not find {aux_data_name} field")
        elif aux_data_number == 1 and not owner_updated:
            print("Warning: Could not update Owner field")
            
        return success
    
    def get_pole_attributes(self, pole_element: ET.Element = None) -> Dict[str, str]:
        """
        Get all attributes of a pole.
        
        Args:
            pole_element (ET.Element, optional): Specific pole element
            
        Returns:
            Dict[str, str]: Dictionary of all pole attributes
        """
        if pole_element is None:
            poles = self.find_wood_poles()
            if not poles:
                return {}
            pole_element = poles[0]
        
        attributes_dict = {}
        attributes = pole_element.find('ATTRIBUTES')
        if attributes:
            for value in attributes.findall('VALUE'):
                name = value.get('NAME')
                type_attr = value.get('TYPE')
                text_value = value.text or 'Unset'
                attributes_dict[name] = {
                    'value': text_value,
                    'type': type_attr
                }
        
        return attributes_dict
    
    def set_pole_attribute(self, attribute_name: str, new_value: str, pole_element: ET.Element = None) -> bool:
        """
        Set any pole attribute value.
        
        Args:
            attribute_name (str): Name of the attribute to modify
            new_value (str): New value to set
            pole_element (ET.Element, optional): Specific pole element
            
        Returns:
            bool: True if successful, False otherwise
        """
        if pole_element is None:
            poles = self.find_wood_poles()
            if not poles:
                print("Error: No WoodPole elements found")
                return False
            pole_element = poles[0]
        
        attributes = pole_element.find('ATTRIBUTES')
        if attributes:
            for value in attributes.findall('VALUE'):
                if value.get('NAME') == attribute_name:
                    old_value = value.text or 'Unset'
                    value.text = new_value
                    print(f"Updated {attribute_name}: '{old_value}' -> '{new_value}'")
                    return True
        
        print(f"Error: Could not find attribute '{attribute_name}'")
        return False
    
    def save_file(self, output_path: str = None) -> bool:
        """
        Save the modified XML to a file.
        
        Args:
            output_path (str, optional): Output file path. If None, overwrites original.
            
        Returns:
            bool: True if successful, False otherwise
        """
        if not self.tree:
            print("Error: No file loaded")
            return False
        
        if output_path is None:
            output_path = self.file_path
        
        try:
            self.tree.write(output_path, encoding='utf-8', xml_declaration=True)
            print(f"File saved successfully: {output_path}")
            return True
        except Exception as e:
            print(f"Error saving file: {e}")
            return False
    
    def list_all_elements(self) -> Dict[str, int]:
        """
        List all unique element types in the XML and their counts.
        
        Returns:
            Dict[str, int]: Dictionary of element types and their counts
        """
        if not self.root:
            return {}
        
        element_counts = {}
        for elem in self.root.iter():
            tag = elem.tag
            element_counts[tag] = element_counts.get(tag, 0) + 1
        
        return element_counts
    
    def find_elements_by_type(self, element_type: str) -> List[ET.Element]:
        """
        Find all elements of a specific type.
        
        Args:
            element_type (str): Type of element to find (e.g., 'Insulator', 'Span')
            
        Returns:
            List[ET.Element]: List of matching elements
        """
        if not self.root:
            return []
        
        return self.root.findall(f'.//{element_type}')
    
    def export_structure_to_json(self, output_file: str = None) -> Dict:
        """
        Export the XML structure to a JSON file for analysis.
        
        Args:
            output_file (str, optional): Output JSON file path
            
        Returns:
            Dict: Dictionary representation of the structure
        """
        if not self.root:
            return {}
        
        def xml_to_dict(element):
            result = {'tag': element.tag, 'attributes': element.attrib}
            if element.text and element.text.strip():
                result['text'] = element.text.strip()
            
            children = []
            for child in element:
                children.append(xml_to_dict(child))
            
            if children:
                result['children'] = children
            
            return result
        
        structure = xml_to_dict(self.root)
        
        if output_file:
            try:
                with open(output_file, 'w', encoding='utf-8') as f:
                    json.dump(structure, f, indent=2, ensure_ascii=False)
                print(f"Structure exported to: {output_file}")
            except Exception as e:
                print(f"Error exporting structure: {e}")
        
        return structure


class PPLXBatchProcessor:
    """Batch processor for multiple PPLX files."""
    
    def __init__(self, directory_path: str = "pplx_files"):
        """
        Initialize the batch processor.
        
        Args:
            directory_path (str): Path to directory containing pplx files
        """
        self.directory_path = directory_path
        self.pplx_files = self._find_pplx_files()
    
    def _find_pplx_files(self) -> List[str]:
        """Find all pplx files in the directory."""
        pattern = os.path.join(self.directory_path, "*.pplx")
        return sorted(glob.glob(pattern))
    
    def list_files(self) -> List[str]:
        """List all found pplx files."""
        return self.pplx_files
    
    def batch_update_aux_data(self, aux_data_number: int, new_value: str, 
                             file_pattern: str = "*") -> Dict[str, bool]:
        """
        Update Aux Data for multiple files.
        
        Args:
            aux_data_number (int): Aux Data field number (1-8)
            new_value (str): New value to set
            file_pattern (str): Pattern to match files (default: all files)
            
        Returns:
            Dict[str, bool]: Results for each file
        """
        results = {}
        
        for file_path in self.pplx_files:
            filename = os.path.basename(file_path)
            if file_pattern == "*" or file_pattern in filename:
                handler = PPLXHandler(file_path)
                success = handler.set_aux_data(aux_data_number, new_value)
                if success:
                    success = handler.save_file()
                results[filename] = success
        
        return results
    
    def generate_report(self) -> Dict:
        """
        Generate a report of all files and their Aux Data.
        
        Returns:
            Dict: Report data
        """
        report = {
            'total_files': len(self.pplx_files),
            'files': {}
        }
        
        for file_path in self.pplx_files:
            filename = os.path.basename(file_path)
            handler = PPLXHandler(file_path)
            
            file_info = handler.get_file_info()
            aux_data = handler.get_aux_data()
            
            report['files'][filename] = {
                'info': file_info,
                'aux_data': aux_data
            }
        
        return report


def main():
    """Main function for command-line interface."""
    parser = argparse.ArgumentParser(description='PPLX File Handler - XML Editor for Pole Line Engineering Files')
    parser.add_argument('file', nargs='?', help='Path to pplx file')
    parser.add_argument('--list-files', action='store_true', help='List all pplx files in directory')
    parser.add_argument('--info', action='store_true', help='Show file information')
    parser.add_argument('--aux-data', action='store_true', help='Show all Aux Data values')
    parser.add_argument('--set-aux', type=int, metavar='N', help='Set Aux Data N (1-8)')
    parser.add_argument('--value', type=str, help='Value to set for Aux Data')
    parser.add_argument('--attributes', action='store_true', help='Show all pole attributes')
    parser.add_argument('--set-attr', type=str, help='Set attribute by name')
    parser.add_argument('--output', '-o', type=str, help='Output file path')
    parser.add_argument('--export-json', type=str, help='Export structure to JSON file')
    parser.add_argument('--batch-report', action='store_true', help='Generate batch report')
    
    args = parser.parse_args()
    
    if args.list_files:
        processor = PPLXBatchProcessor()
        files = processor.list_files()
        print(f"Found {len(files)} pplx files:")
        for file_path in files:
            print(f"  {os.path.basename(file_path)}")
        return
    
    if args.batch_report:
        processor = PPLXBatchProcessor()
        report = processor.generate_report()
        print(f"Batch Report - {report['total_files']} files processed:")
        for filename, data in report['files'].items():
            print(f"\n{filename}:")
            print(f"  Pole Number: {data['info'].get('Pole Number', 'Unknown')}")
            print(f"  Date: {data['info'].get('date', 'Unknown')}")
            aux_data = data['aux_data']
            for aux_name, aux_value in aux_data.items():
                if aux_value != 'Unset':
                    print(f"  {aux_name}: {aux_value}")
        return
    
    if not args.file:
        print("Please provide a file path or use --list-files or --batch-report")
        return
    
    # Load the specified file
    handler = PPLXHandler(args.file)
    
    if args.info:
        info = handler.get_file_info()
        print("File Information:")
        for key, value in info.items():
            print(f"  {key}: {value}")
    
    if args.aux_data:
        aux_data = handler.get_aux_data()
        print("\nAux Data Values:")
        for name, value in aux_data.items():
            print(f"  {name}: {value}")
    
    if args.attributes:
        attributes = handler.get_pole_attributes()
        print("\nPole Attributes:")
        for name, data in attributes.items():
            print(f"  {name} ({data['type']}): {data['value']}")
    
    if args.set_aux and args.value:
        success = handler.set_aux_data(args.set_aux, args.value)
        if success:
            handler.save_file(args.output)
    
    if args.set_attr and args.value:
        success = handler.set_pole_attribute(args.set_attr, args.value)
        if success:
            handler.save_file(args.output)
    
    if args.export_json:
        handler.export_structure_to_json(args.export_json)


if __name__ == "__main__":
    main()
