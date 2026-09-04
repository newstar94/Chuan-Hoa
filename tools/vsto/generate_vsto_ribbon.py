from __future__ import annotations

import json
import re
import zipfile
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[2]
CONTRACT_PATH = ROOT / "shared" / "contracts" / "ribbon" / "ribbon-contract.v1.json"
SOURCE_DOTM_PATH = ROOT / "shared" / "ChuanHoaTheThuc_Full_Ribbon.dotm"
OUTPUT_DIRECTORY = ROOT / "src" / "ChuanHoa.AddIn.Vsto" / "Ribbon"
OUTPUT_XML_PATH = OUTPUT_DIRECTORY / "ChuanHoaRibbon.xml"
OUTPUT_CALLBACKS_PATH = OUTPUT_DIRECTORY / "ChuanHoaRibbon.Callbacks.g.cs"
RIBBON_NAMESPACE = "http://schemas.microsoft.com/office/2009/07/customui"
CALLBACK_ATTRIBUTES = {
    "getEnabled",
    "getImage",
    "getItemCount",
    "getItemLabel",
    "getPressed",
    "getSelectedItemIndex",
    "onAction",
}


def load_contract() -> dict:
    return json.loads(CONTRACT_PATH.read_text(encoding="utf-8"))


def load_source_xml() -> ET.ElementTree:
    with zipfile.ZipFile(SOURCE_DOTM_PATH) as archive:
        xml_bytes = archive.read("customUI/customUI14.xml")
    parser = ET.XMLParser()
    return ET.ElementTree(ET.fromstring(xml_bytes, parser=parser))


def update_xml(tree: ET.ElementTree, contract: dict) -> None:
    ET.register_namespace("", RIBBON_NAMESPACE)
    root = tree.getroot()
    root.set("onLoad", "RibbonOnLoad")
    elements_by_id = {
        element.attrib["id"]: element
        for element in root.iter()
        if "id" in element.attrib
    }

    expected_ids = {control["id"] for control in contract["controls"]}
    for control in contract["controls"]:
        if control["id"] in elements_by_id:
            continue
        parent_id = control.get("parentContainerId") or control["groupId"]
        parent = elements_by_id.get(parent_id)
        if parent is None:
            raise RuntimeError(
                f"Cannot add {control['id']}: Ribbon parent {parent_id} is missing"
            )
        element = ET.SubElement(
            parent,
            f"{{{RIBBON_NAMESPACE}}}{control['controlType']}",
            {"id": control["id"]},
        )
        elements_by_id[control["id"]] = element

    parent_by_child = {
        child: parent
        for parent in root.iter()
        for child in parent
    }
    expected_group_ids = {group["id"] for group in contract["groups"]}
    for control_id, element in list(elements_by_id.items()):
        local_name = element.tag.rsplit("}", 1)[-1]
        if local_name in {"button", "menu", "dropDown", "checkBox"} and control_id not in expected_ids:
            parent_by_child[element].remove(element)
            del elements_by_id[control_id]
        elif local_name == "group" and control_id not in expected_group_ids:
            parent_by_child[element].remove(element)
            del elements_by_id[control_id]

    for control in contract["controls"]:
        element = elements_by_id[control["id"]]
        element.set("label", control["label"])
        update_optional_attribute(element, "size", control.get("size"))
        update_optional_attribute(element, "imageMso", control.get("imageMso"))
        update_optional_attribute(element, "screentip", control.get("screenTip"))
        update_optional_attribute(element, "supertip", control.get("superTip"))

        for callback_attribute in CALLBACK_ATTRIBUTES:
            element.attrib.pop(callback_attribute, None)
        for callback_attribute, callback_name in control["callbacks"].items():
            element.set(callback_attribute, callback_name)

    tab = elements_by_id[contract["tab"]["id"]]
    tab.set("label", contract["tab"]["label"])


def update_optional_attribute(element: ET.Element, name: str, value: str | None) -> None:
    if value is None or value == "":
        element.attrib.pop(name, None)
    else:
        element.set(name, value)


def write_xml(tree: ET.ElementTree) -> None:
    OUTPUT_DIRECTORY.mkdir(parents=True, exist_ok=True)
    ET.indent(tree, space="  ")
    tree.write(OUTPUT_XML_PATH, encoding="utf-8", xml_declaration=True)


def build_callback_index(contract: dict) -> dict[str, list[tuple[str, str, str]]]:
    callbacks: dict[str, list[tuple[str, str, str]]] = defaultdict(list)
    for control in contract["controls"]:
        for callback_type, callback_name in control["callbacks"].items():
            callbacks[callback_name].append((callback_type, control["controlType"], control["id"]))
    return dict(callbacks)


def validate_callback_shapes(callbacks: dict[str, list[tuple[str, str, str]]]) -> None:
    for callback_name, uses in callbacks.items():
        callback_types = {callback_type for callback_type, _, _ in uses}
        if len(callback_types) != 1:
            raise RuntimeError(f"Callback {callback_name} is used for multiple callback types: {uses}")

        callback_type = next(iter(callback_types))
        if callback_type == "onAction":
            control_types = {control_type for _, control_type, _ in uses}
            if len(control_types) != 1:
                raise RuntimeError(f"onAction callback {callback_name} has incompatible controls: {uses}")


def generate_callbacks(contract: dict) -> None:
    callbacks = build_callback_index(contract)
    validate_callback_shapes(callbacks)
    methods: list[str] = []
    methods.append(
        """        public void RibbonOnLoad(Office.IRibbonUI ribbonUi)
        {
            if (ribbonUi == null)
            {
                return;
            }

            CompleteRibbonLoad(ribbonUi);
        }"""
    )

    for callback_name in sorted(callbacks):
        uses = callbacks[callback_name]
        callback_type = uses[0][0]
        control_type = uses[0][1]
        if callback_type == "getEnabled":
            body = f"""        public bool {callback_name}(Office.IRibbonControl control)
        {{
            return Runtime.IsEnabled(RequireControlId(control));
        }}"""
        elif callback_type == "getPressed":
            body = f"""        public bool {callback_name}(Office.IRibbonControl control)
        {{
            return Runtime.GetPressed(RequireControlId(control));
        }}"""
        elif callback_type == "getSelectedItemIndex":
            body = f"""        public int {callback_name}(Office.IRibbonControl control)
        {{
            return Runtime.GetSelectedItemIndex(RequireControlId(control));
        }}"""
        elif callback_type == "getItemCount":
            body = f"""        public int {callback_name}(Office.IRibbonControl control)
        {{
            return Runtime.GetItemCount(RequireControlId(control));
        }}"""
        elif callback_type == "getItemLabel":
            body = f"""        public string {callback_name}(Office.IRibbonControl control, int index)
        {{
            return Runtime.GetItemLabel(RequireControlId(control), index);
        }}"""
        elif callback_type == "getImage":
            body = f"""        public object {callback_name}(Office.IRibbonControl control)
        {{
            return Runtime.GetImage(RequireControlId(control));
        }}"""
        elif callback_type == "onAction" and control_type == "button":
            body = f"""        public void {callback_name}(Office.IRibbonControl control)
        {{
            Runtime.ExecuteButton(RequireControlId(control));
        }}"""
        elif callback_type == "onAction" and control_type == "dropDown":
            body = f"""        public void {callback_name}(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {{
            Runtime.SelectDropDownItem(RequireControlId(control), selectedId, selectedIndex);
        }}"""
        elif callback_type == "onAction" and control_type == "checkBox":
            body = f"""        public void {callback_name}(Office.IRibbonControl control, bool pressed)
        {{
            Runtime.SetPressed(RequireControlId(control), pressed);
        }}"""
        else:
            raise RuntimeError(f"Unsupported callback signature for {callback_name}: {uses}")
        methods.append(body)

    source = f"""// <auto-generated />
using System;
using Office = Microsoft.Office.Core;

namespace ChuanHoa.AddIn.Vsto.Ribbon
{{
    public sealed partial class ChuanHoaRibbon
    {{
{chr(10).join(methods)}
    }}
}}
"""
    OUTPUT_CALLBACKS_PATH.write_text(source, encoding="utf-8", newline="\n")


def validate_outputs(contract: dict) -> dict:
    tree = ET.parse(OUTPUT_XML_PATH)
    root = tree.getroot()
    counts = defaultdict(int)
    callback_names = set()
    control_ids = set()

    for element in root.iter():
        local_name = element.tag.rsplit("}", 1)[-1]
        if local_name in {"button", "menu", "dropDown", "checkBox"}:
            counts[local_name] += 1
            control_ids.add(element.attrib["id"])
        for callback_attribute in CALLBACK_ATTRIBUTES:
            callback_name = element.attrib.get(callback_attribute)
            if callback_name:
                callback_names.add(callback_name)

    source = OUTPUT_CALLBACKS_PATH.read_text(encoding="utf-8")
    missing_methods = sorted(
        callback_name
        for callback_name in callback_names | {"RibbonOnLoad"}
        if re.search(rf"\b{re.escape(callback_name)}\s*\(", source) is None
    )
    if missing_methods:
        raise RuntimeError(f"Generated callback methods are missing: {missing_methods}")

    expected_counts = contract["counts"]
    actual_counts = {
        "buttons": counts["button"],
        "menus": counts["menu"],
        "dropDowns": counts["dropDown"],
        "checkBoxes": counts["checkBox"],
        "interactiveControls": len(control_ids),
    }
    for key, expected in expected_counts.items():
        if key in actual_counts and actual_counts[key] != expected:
            raise RuntimeError(f"Ribbon count mismatch for {key}: {actual_counts[key]} != {expected}")

    return {
        "status": "PASS",
        "xml": str(OUTPUT_XML_PATH),
        "callbacks": str(OUTPUT_CALLBACKS_PATH),
        "counts": actual_counts,
        "callbackMethods": len(callback_names) + 1,
    }


def main() -> None:
    contract = load_contract()
    tree = load_source_xml()
    update_xml(tree, contract)
    write_xml(tree)
    generate_callbacks(contract)
    print(json.dumps(validate_outputs(contract), ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
