"""
PPLX File Handler - XML Editor for Pole Line Engineering Files.
Uses lxml for fast parsing when available (C-based); falls back to stdlib xml.etree.ElementTree.
"""

import json
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

try:
    from lxml import etree as ET
    _XML_PARSER = "lxml"
except ImportError:
    import xml.etree.ElementTree as ET
    _XML_PARSER = "stdlib"
import math

# Log parser on first import (only once)
if not hasattr(ET, "_parser_logged"):
    print(f"PPLX XML parser: {_XML_PARSER}")
    ET._parser_logged = True


# ---------------------------------------------------------------------------
# Generic XML attribute helpers
# ---------------------------------------------------------------------------

def _get_attr_str(elem: Any, name: str) -> Optional[str]:
    """Return the text of the first VALUE with NAME from element's ATTRIBUTES, stripped.
    Returns None when element/ATTRIBUTES is missing or no VALUE matches."""
    attrs = elem.find("ATTRIBUTES")
    if attrs is None:
        return None
    for value in attrs.findall("VALUE"):
        if value.get("NAME") == name:
            return (value.text or "").strip() or None
    return None


def _get_attr_float(elem: Any, name: str) -> Optional[float]:
    """Return the text of the first VALUE with NAME as float, or None."""
    raw = _get_attr_str(elem, name)
    if raw is None:
        return None
    try:
        return float(raw)
    except (TypeError, ValueError):
        return None


def _set_attr_value(elem: Any, name: str, new_value: str, type_str: str = "String") -> bool:
    """Set or create a VALUE element under ATTRIBUTES.  Returns True on success."""
    attrs = elem.find("ATTRIBUTES")
    if attrs is None:
        return False
    for value in attrs.findall("VALUE"):
        if value.get("NAME") == name:
            value.text = new_value
            return True
    # Create new VALUE element if not found
    new_elem = ET.SubElement(attrs, "VALUE")
    new_elem.set("NAME", name)
    new_elem.set("TYPE", type_str)
    new_elem.text = new_value
    return True


def _parse_span_attrs(span: Any) -> Tuple[Optional[str], Optional[float], Optional[str]]:
    """Extract (SpanType, length_inches, conductor_type) from a Span element in one pass."""
    attrs = span.find("ATTRIBUTES")
    if attrs is None:
        return None, None, None
    span_type = length = conductor = None
    for value in attrs.findall("VALUE"):
        n = value.get("NAME")
        if n == "SpanType":
            span_type = (value.text or "").strip() or None
        elif n == "SpanDistanceInInches" and value.text:
            try:
                length = float(value.text)
            except (TypeError, ValueError):
                pass
        elif n == "Type":
            conductor = (value.text or "").strip() or None
    return span_type, length, conductor


class PPLXHandler:
    """Handler class for reading and editing PPLX XML files."""

    def __init__(self, file_path: Optional[str] = None):
        self.file_path = file_path
        self.tree: Optional[Any] = None
        self.root: Optional[Any] = None
        self._cached_spans: Optional[List[Any]] = None
        if file_path and os.path.exists(file_path):
            self.load_file(file_path)

    def load_file(self, file_path: str) -> bool:
        """Load a pplx XML file."""
        try:
            self.tree = ET.parse(file_path)
            self.root = self.tree.getroot()
            self.file_path = file_path
            self._cached_spans = None
            return True
        except (ET.ParseError, getattr(ET, "XMLSyntaxError", ET.ParseError)) as e:
            print(f"Error parsing XML file {file_path}: {e}")
            return False
        except FileNotFoundError:
            print(f"File not found: {file_path}")
            return False

    # ------------------------------------------------------------------
    # Read helpers
    # ------------------------------------------------------------------

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
            for name in ("Latitude", "Longitude", "ElevationMetersAboveMSL", "PPLVersion"):
                val = _get_attr_str(scene, name)
                if val is not None:
                    info[name] = val
        return info

    def find_wood_poles(self) -> List[ET.Element]:
        """Find all WoodPole elements in the file."""
        if not self.root:
            return []
        return self.root.findall(".//WoodPole")

    def get_scene_lat_lon(self) -> tuple:
        """Return (latitude, longitude) from PPLScene ATTRIBUTES, or (None, None)."""
        if not self.root:
            return (None, None)
        scene = self.root.find("PPLScene")
        if scene is None:
            return (None, None)
        lat = _get_attr_float(scene, "Latitude")
        lon = _get_attr_float(scene, "Longitude")
        return (lat, lon)

    def _get_all_spans(self) -> List[Any]:
        """Get all spans, using cache if available."""
        if self.root is None:
            return []
        if self._cached_spans is None:
            self._cached_spans = self.root.findall(".//Span")
        return self._cached_spans

    # ------------------------------------------------------------------
    # Span queries
    # ------------------------------------------------------------------

    def get_spans_by_type_and_length(self) -> Dict[str, Dict[float, str]]:
        """
        Returns {SpanType: {length_inches: conductor_type}}.
        Groups spans by SpanType; for each length, stores first conductor Type seen.
        """
        result: Dict[str, Dict[float, str]] = {}
        for span in self._get_all_spans():
            span_type, length, conductor = _parse_span_attrs(span)
            if span_type and length is not None:
                if span_type not in result:
                    result[span_type] = {}
                if length not in result[span_type]:
                    result[span_type][length] = conductor or ""
        return result

    def get_span_type_angle_pairs(self) -> List[Tuple[str, float]]:
        """
        Returns [(SpanType, angle_rad), ...] for each span.
        Angle is the span's direction: Span's CoordinateA if set, else parent Insulator's.
        """
        out: List[Tuple[str, float]] = []
        if self.root is None:
            return out
        for pole in self.find_wood_poles():
            for insulator in pole.findall(".//Insulator"):
                ins_angle = _get_attr_float(insulator, "CoordinateA")
                for span in insulator.findall(".//Span"):
                    span_type = _get_attr_str(span, "SpanType")
                    span_angle = _get_attr_float(span, "CoordinateA")
                    angle = span_angle if span_angle is not None else ins_angle
                    if span_type and angle is not None:
                        out.append((span_type, angle % (2.0 * math.pi)))
        return out

    def get_span_type_length_pairs(self) -> List[Tuple[str, float]]:
        """One pass over all spans. Returns [(SpanType, length_inches), ...]."""
        out: List[Tuple[str, float]] = []
        if self.root is None:
            return out
        for span in self._get_all_spans():
            span_type, length, _ = _parse_span_attrs(span)
            if span_type and length is not None:
                out.append((span_type, length))
        return out

    def get_span_type_length_pairs_for_spans_qc(self) -> List[Tuple[str, float]]:
        """
        Span data for Spans QC:
        - For CATV / Fiber / Telco spans inside a SpanBundle, treat each SpanBundle
          as a single span per SpanType (even if multiple child spans exist).
        - All other spans are counted one-to-one.
        Returns [(SpanType, length_inches), ...].
        """
        out: List[Tuple[str, float]] = []
        if self.root is None:
            return out

        comm_types = {"catv", "fiber", "telco"}
        spans_in_comm_bundles: set[int] = set()

        # First pass: handle SpanBundle children for comm types
        for bundle in self.root.findall(".//SpanBundle"):
            seen_in_bundle: set[str] = set()
            for span in bundle.findall(".//Span"):
                span_type, length, _ = _parse_span_attrs(span)
                if not span_type or length is None:
                    continue
                key = span_type.strip().lower()
                if key in comm_types:
                    if key not in seen_in_bundle:
                        out.append((span_type, length))
                        seen_in_bundle.add(key)
                    spans_in_comm_bundles.add(id(span))

        # Second pass: all spans not already counted above
        for span in self._get_all_spans():
            if id(span) in spans_in_comm_bundles:
                continue
            span_type, length, _ = _parse_span_attrs(span)
            if span_type and length is not None:
                out.append((span_type, length))
        return out

    def get_span_type_length_angle_triples_for_spans_qc(self) -> List[Tuple[str, float, Optional[float]]]:
        """
        Span data for Spans QC with directional angles:
        - For CATV / Fiber / Telco spans inside a SpanBundle, treat each SpanBundle
          as a single span per SpanType and USE THE BUNDLE'S ABSOLUTE ANGLE.
          (Bundle's CoordinateA is RELATIVE to parent Insulator's angle.)
        - All other spans are counted one-to-one, using Span's CoordinateA or parent Insulator's.
        Returns [(SpanType, length_inches, angle_rad), ...].
        """
        out: List[Tuple[str, float, Optional[float]]] = []
        if self.root is None:
            return out

        comm_types = {"catv", "fiber", "telco"}
        spans_in_comm_bundles: set[int] = set()

        # Build parent map for finding insulator parent of bundles
        parent_map = {c: p for p in self.root.iter() for c in p}

        # First pass: handle SpanBundle children for comm types
        # Bundle's CoordinateA is RELATIVE to parent Insulator's angle
        # We need to compute ABSOLUTE angle = insulator_angle + bundle_relative_angle
        for bundle in self.root.findall(".//SpanBundle"):
            bundle_relative_angle = _get_attr_float(bundle, "CoordinateA")

            # Find parent Insulator by walking up the parent map
            parent_insulator = None
            current = bundle
            while current in parent_map:
                parent = parent_map[current]
                if parent.tag == "Insulator":
                    parent_insulator = parent
                    break
                current = parent

            # Compute absolute angle: insulator's angle + bundle's relative angle
            absolute_angle = None
            if parent_insulator is not None:
                ins_angle = _get_attr_float(parent_insulator, "CoordinateA")
                if ins_angle is not None and bundle_relative_angle is not None:
                    absolute_angle = (ins_angle + bundle_relative_angle) % (2.0 * math.pi)
                elif ins_angle is not None:
                    absolute_angle = ins_angle
            elif bundle_relative_angle is not None:
                # Fallback: use bundle angle if no parent insulator found
                absolute_angle = bundle_relative_angle % (2.0 * math.pi)

            seen_in_bundle: set[str] = set()
            # Spans in bundle are under PPLChildElements/Span
            for span in bundle.findall(".//Span"):
                span_type, length, _ = _parse_span_attrs(span)
                if not span_type or length is None:
                    continue
                key = span_type.strip().lower()
                if key in comm_types:
                    if key not in seen_in_bundle:
                        out.append((span_type, length, absolute_angle))
                        seen_in_bundle.add(key)
                    # Track ALL comm spans in bundles (even if we only add one to output)
                    spans_in_comm_bundles.add(id(span))

        # Second pass: all spans not already counted above
        # For power spans, use Span's CoordinateA or parent Insulator's
        #
        # Cross-insulator artifact dedup: the same physical span can appear under both a
        # direction-specific insulator (ins_angle=X, span_rel=0) AND a base insulator
        # (ins_angle=0, span_rel=X), producing the same absolute angle.  Skip the copy from
        # the base insulator when a copy from a different-base-angle insulator already exists.
        # Spans sharing the SAME insulator base-angle are kept in full (e.g. 3 secondary
        # conductors each modeled on their own ins_angle=270.7° insulator are all legitimate).
        #
        # Key: (type_lower, len_rounded_inch, abs_angle_bucket_half_deg)
        # Value: the base angle (in degrees, rounded to 0.5°) of the first insulator that emitted it
        seen_dedup: Dict[tuple, int] = {}  # key -> first insulator's base angle bucket
        for pole in self.find_wood_poles():
            for insulator in pole.findall(".//Insulator"):
                ins_angle = _get_attr_float(insulator, "CoordinateA")
                ins_base_bucket = round(math.degrees(ins_angle) * 2) if ins_angle is not None else 0
                for span in insulator.findall(".//Span"):
                    span_id = id(span)
                    # CRITICAL: Skip ANY span that's in a comm bundle, regardless of type
                    if span_id in spans_in_comm_bundles:
                        continue  # Skip this span entirely - it was already added in first pass
                    span_type, length, _ = _parse_span_attrs(span)
                    if span_type and length is not None:
                        span_angle = _get_attr_float(span, "CoordinateA")
                        # Span's CoordinateA is relative to parent Insulator's angle
                        if ins_angle is not None and span_angle is not None:
                            angle = (ins_angle + span_angle) % (2.0 * math.pi)
                        elif ins_angle is not None:
                            angle = ins_angle % (2.0 * math.pi)
                        elif span_angle is not None:
                            angle = span_angle % (2.0 * math.pi)
                        else:
                            angle = None
                        # Dedup: skip if same (type, length, abs-angle) was already emitted
                        # from an insulator with a DIFFERENT base angle (artifact duplicate).
                        # Copies from insulators at the same base angle are all kept.
                        angle_bucket = round(math.degrees(angle) * 2) if angle is not None else None
                        dedup_key = (span_type.lower(), round(length), angle_bucket)
                        prev_base = seen_dedup.get(dedup_key)
                        if prev_base is not None and prev_base != ins_base_bucket:
                            continue  # Artifact: same span re-encoded on a different-angle insulator
                        seen_dedup.setdefault(dedup_key, ins_base_bucket)
                        out.append((span_type, length, angle))
        return out

    def get_span_type_counts_for_length(
        self, length_inches: float, tolerance: float = 0.15
    ) -> Dict[str, int]:
        """
        Return count of spans per SpanType for spans whose SpanDistanceInInches
        matches length_inches within tolerance (fraction, e.g. 0.15 = 15%).
        """
        counts: Dict[str, int] = {}
        if self.root is None or length_inches <= 0:
            return counts
        for span in self._get_all_spans():
            span_type, length, _ = _parse_span_attrs(span)
            if not span_type or length is None:
                continue
            if abs(length - length_inches) / max(length_inches, 1) <= tolerance:
                counts[span_type] = counts.get(span_type, 0) + 1
        return counts

    # ------------------------------------------------------------------
    # Span mutation
    # ------------------------------------------------------------------

    def set_span_conductor_type(self, span_element: ET.Element, new_value: str) -> bool:
        """Set the Type (conductor/wire spec) of a Span element."""
        return _set_attr_value(span_element, "Type", new_value)

    # ------------------------------------------------------------------
    # Aux Data
    # ------------------------------------------------------------------

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
        success = _set_attr_value(pole_element, aux_data_name, new_value)

        if success:
            print(f"Updated {aux_data_name}: '{new_value}'")
        else:
            print(f"Error: Could not find {aux_data_name} field")
            return False

        # When setting Aux Data 1, also sync the Owner field
        if aux_data_number == 1:
            if not _set_attr_value(pole_element, "Owner", new_value):
                print("Warning: Could not update Owner field")

        return success

    # ------------------------------------------------------------------
    # Pole attributes
    # ------------------------------------------------------------------

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

    # ------------------------------------------------------------------
    # File I/O
    # ------------------------------------------------------------------

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

    # ------------------------------------------------------------------
    # Inspection / debugging utilities
    # ------------------------------------------------------------------

    def list_all_elements(self) -> Dict[str, int]:
        """List all unique element types and their counts."""
        if not self.root:
            return {}
        element_counts: Dict[str, int] = {}
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
