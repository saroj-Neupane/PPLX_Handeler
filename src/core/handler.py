"""
PPLX File Handler - XML Editor for Pole Line Engineering Files.
"""

import glob
import json
import os
from pathlib import Path
from typing import Any, Dict, List, Optional

import xml.etree.ElementTree as ET


class PPLXHandler:
    """Handler class for reading and editing PPLX XML files."""

    def __init__(self, file_path: Optional[str] = None):
        self.file_path = file_path
        self.tree: Optional[Any] = None
        self.root: Optional[Any] = None
        if file_path and os.path.exists(file_path):
            self.load_file(file_path)

    def load_file(self, file_path: str) -> bool:
        """Load a pplx XML file."""
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
        """Get basic information about the loaded pplx file."""
        if not self.root:
            return {}
        info = {
            "file_path": self.file_path,
            "date": self.root.get("DATE", "Unknown"),
            "user": self.root.get("USER", "Unknown"),
            "workstation": self.root.get("WORKSTATION", "Unknown"),
        }
        scene = self.root.find("PPLScene")
        if scene:
            attributes = scene.find("ATTRIBUTES")
            if attributes:
                for value in attributes.findall("VALUE"):
                    name = value.get("NAME")
                    if name in [
                        "Latitude",
                        "Longitude",
                        "ElevationMetersAboveMSL",
                        "PPLVersion",
                    ]:
                        info[name] = value.text
        return info

    def find_wood_poles(self) -> List[ET.Element]:
        """Find all WoodPole elements in the file."""
        if not self.root:
            return []
        return self.root.findall(".//WoodPole")

    def get_scene_lat_lon(self) -> tuple:
        """Return (latitude, longitude) from PPLScene ATTRIBUTES, or (None, None) if missing."""
        if not self.root:
            return (None, None)
        scene = self.root.find("PPLScene")
        if scene is None:
            return (None, None)
        attrs = scene.find("ATTRIBUTES")
        if attrs is None:
            return (None, None)
        lat = lon = None
        for value in attrs.findall("VALUE"):
            name = value.get("NAME")
            if name == "Latitude" and value.text:
                try:
                    lat = float(value.text)
                except (TypeError, ValueError):
                    pass
            elif name == "Longitude" and value.text:
                try:
                    lon = float(value.text)
                except (TypeError, ValueError):
                    pass
        return (lat, lon)

    def find_spans_by_type(self, span_type: str) -> List[ET.Element]:
        """Find all Span elements whose SpanType equals span_type (e.g. 'Primary')."""
        if not self.root:
            return []
        out = []
        for span in self.root.findall(".//Span"):
            attrs = span.find("ATTRIBUTES")
            if attrs is None:
                continue
            for value in attrs.findall("VALUE"):
                if value.get("NAME") == "SpanType" and (value.text or "").strip() == span_type:
                    out.append(span)
                    break
        return out

    def get_span_conductor_type(self, span_element: ET.Element) -> Optional[str]:
        """Get the Type (conductor/wire spec) of a Span element."""
        attrs = span_element.find("ATTRIBUTES")
        if attrs is None:
            return None
        for value in attrs.findall("VALUE"):
            if value.get("NAME") == "Type":
                return (value.text or "").strip() or None
        return None

    def get_span_length_inches(self, span_element: ET.Element) -> Optional[float]:
        """Get SpanDistanceInInches of a Span element."""
        attrs = span_element.find("ATTRIBUTES")
        if attrs is None:
            return None
        for value in attrs.findall("VALUE"):
            if value.get("NAME") == "SpanDistanceInInches" and value.text:
                try:
                    return float(value.text)
                except (TypeError, ValueError):
                    return None
        return None

    def get_spans_by_type_and_length(
        self,
    ) -> Dict[str, Dict[float, str]]:
        """
        Returns {SpanType: {length_inches: conductor_type}}.
        Groups spans by SpanType; for each length, stores first conductor Type seen.
        """
        result = {}
        for span in self.root.findall(".//Span") if self.root else []:
            attrs = span.find("ATTRIBUTES")
            if attrs is None:
                continue
            span_type = length = conductor = None
            for value in attrs.findall("VALUE"):
                n = value.get("NAME")
                if n == "SpanType":
                    span_type = (value.text or "").strip()
                elif n == "SpanDistanceInInches" and value.text:
                    try:
                        length = float(value.text)
                    except (TypeError, ValueError):
                        pass
                elif n == "Type":
                    conductor = (value.text or "").strip() or None
            if span_type and length is not None:
                if span_type not in result:
                    result[span_type] = {}
                if length not in result[span_type]:
                    result[span_type][length] = conductor or ""
        return result

    def set_span_conductor_type(self, span_element: ET.Element, new_value: str) -> bool:
        """Set the Type (conductor/wire spec) of a Span element."""
        attrs = span_element.find("ATTRIBUTES")
        if attrs is None:
            return False
        for value in attrs.findall("VALUE"):
            if value.get("NAME") == "Type":
                value.text = new_value
                return True
        elem = ET.SubElement(attrs, "VALUE")
        elem.set("NAME", "Type")
        elem.set("TYPE", "String")
        elem.text = new_value
        return True

    def get_aux_data(self, pole_element: Optional[ET.Element] = None) -> Dict[str, str]:
        """Get all Aux Data values from a pole or the first pole if none specified."""
        if pole_element is None:
            poles = self.find_wood_poles()
            if not poles:
                return {}
            pole_element = poles[0]
        aux_data = {}
        attributes = pole_element.find("ATTRIBUTES")
        if attributes:
            for value in attributes.findall("VALUE"):
                name = value.get("NAME")
                if name and name.startswith("Aux Data"):
                    aux_data[name] = value.text or "Unset"
        return aux_data

    def set_aux_data(
        self,
        aux_data_number: int,
        new_value: str,
        pole_element: Optional[ET.Element] = None,
    ) -> bool:
        """Set the value of a specific Aux Data field (1-8)."""
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
        attributes = pole_element.find("ATTRIBUTES")
        success = False
        owner_updated = False

        if attributes:
            for value in attributes.findall("VALUE"):
                if value.get("NAME") == aux_data_name:
                    old_value = value.text or "Unset"
                    value.text = new_value
                    print(f"Updated {aux_data_name}: '{old_value}' -> '{new_value}'")
                    success = True
                    break
            if not success:
                aux_element = ET.SubElement(attributes, "VALUE")
                aux_element.set("NAME", aux_data_name)
                aux_element.set("TYPE", "String")
                aux_element.text = new_value
                print(f"Created {aux_data_name}: '{new_value}'")
                success = True

            if aux_data_number == 1 and success:
                for value in attributes.findall("VALUE"):
                    if value.get("NAME") == "Owner":
                        old_owner = value.text or "Unset"
                        value.text = new_value
                        print(f"Updated Owner: '{old_owner}' -> '{new_value}'")
                        owner_updated = True
                        break
                if not owner_updated:
                    owner_element = ET.SubElement(attributes, "VALUE")
                    owner_element.set("NAME", "Owner")
                    owner_element.set("TYPE", "String")
                    owner_element.text = new_value
                    owner_updated = True

        if not success:
            print(f"Error: Could not find {aux_data_name} field")
        elif aux_data_number == 1 and not owner_updated:
            print("Warning: Could not update Owner field")
        return success

    def get_pole_attributes(
        self, pole_element: Optional[ET.Element] = None
    ) -> Dict[str, Dict]:
        """Get all attributes of a pole."""
        if pole_element is None:
            poles = self.find_wood_poles()
            if not poles:
                return {}
            pole_element = poles[0]
        attributes_dict = {}
        attributes = pole_element.find("ATTRIBUTES")
        if attributes:
            for value in attributes.findall("VALUE"):
                name = value.get("NAME")
                type_attr = value.get("TYPE")
                text_value = value.text or "Unset"
                attributes_dict[name] = {"value": text_value, "type": type_attr}
        return attributes_dict

    def set_pole_attribute(
        self,
        attribute_name: str,
        new_value: str,
        pole_element: Optional[ET.Element] = None,
    ) -> bool:
        """Set any pole attribute value."""
        if pole_element is None:
            poles = self.find_wood_poles()
            if not poles:
                print("Error: No WoodPole elements found")
                return False
            pole_element = poles[0]
        attributes = pole_element.find("ATTRIBUTES")
        if attributes:
            for value in attributes.findall("VALUE"):
                if value.get("NAME") == attribute_name:
                    value.text = new_value
                    return True
        print(f"Error: Could not find attribute '{attribute_name}'")
        return False

    def save_file(self, output_path: Optional[str] = None) -> bool:
        """Save the modified XML to a file."""
        if not self.tree:
            print("Error: No file loaded")
            return False
        if output_path is None:
            output_path = self.file_path
        try:
            self.tree.write(output_path, encoding="utf-8", xml_declaration=True)
            print(f"File saved successfully: {output_path}")
            return True
        except Exception as e:
            print(f"Error saving file: {e}")
            return False

    def list_all_elements(self) -> Dict[str, int]:
        """List all unique element types in the XML and their counts."""
        if not self.root:
            return {}
        element_counts = {}
        for elem in self.root.iter():
            tag = elem.tag
            element_counts[tag] = element_counts.get(tag, 0) + 1
        return element_counts

    def find_elements_by_type(self, element_type: str) -> List[ET.Element]:
        """Find all elements of a specific type."""
        if not self.root:
            return []
        return self.root.findall(f".//{element_type}")

    def export_structure_to_json(self, output_file: Optional[str] = None) -> Dict:
        """Export the XML structure to a JSON file."""

        def xml_to_dict(element):
            result = {"tag": element.tag, "attributes": element.attrib}
            if element.text and element.text.strip():
                result["text"] = element.text.strip()
            children = [xml_to_dict(child) for child in element]
            if children:
                result["children"] = children
            return result

        if not self.root:
            return {}
        structure = xml_to_dict(self.root)
        if output_file:
            try:
                with open(output_file, "w", encoding="utf-8") as f:
                    json.dump(structure, f, indent=2, ensure_ascii=False)
                print(f"Structure exported to: {output_file}")
            except Exception as e:
                print(f"Error exporting structure: {e}")
        return structure


class PPLXBatchProcessor:
    """Batch processor for multiple PPLX files."""

    def __init__(self, directory_path: str = "pplx_files"):
        self.directory_path = directory_path
        self.pplx_files = self._find_pplx_files()

    def _find_pplx_files(self) -> List[str]:
        """Find all pplx files in the directory."""
        pattern = os.path.join(self.directory_path, "*.pplx")
        return sorted(glob.glob(pattern))

    def list_files(self) -> List[str]:
        """List all found pplx files."""
        return self.pplx_files

    def batch_update_aux_data(
        self, aux_data_number: int, new_value: str, file_pattern: str = "*"
    ) -> Dict[str, bool]:
        """Update Aux Data for multiple files."""
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
        """Generate a report of all files and their Aux Data."""
        report = {"total_files": len(self.pplx_files), "files": {}}
        for file_path in self.pplx_files:
            filename = os.path.basename(file_path)
            handler = PPLXHandler(file_path)
            report["files"][filename] = {
                "info": handler.get_file_info(),
                "aux_data": handler.get_aux_data(),
            }
        return report
