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
    arrange_compact_columns(root, elements_by_id)


def arrange_compact_columns(root: ET.Element, elements: dict) -> None:
    """Keep secondary actions in three-row columns so later groups fit.

    Word owns responsive group collapse; oversized indivisible groups can leave
    unused space when the next group cannot expand. Avoid a one-button tab column
    and six large actions in the startup group. IDs/callbacks stay unchanged.
    """
    columns = [
        ("grpKhoiDong", "boxReviewActions", ["btnKiemTra", "btnKiemTraChinhTa", "btnChuyenDoiUnicode"]),
        ("grpKhoiDong", "boxFixActions", ["btnSuaLoiDangChon", "btnSuaTatCaChinhTa"]),
        ("grpDinhDang", "boxKeepPageNum", ["btnParagraph", "btnKeepWithNext", "btnXoaTabKhongLeader"]),
        ("grpDinhDang", "boxCharacterTools", ["btnScaleGiam", "btnScale100", "btnScaleTang"]),
        ("grpDinhDang", "boxPageTools", ["btnChenTrangNgang", "btnChenTrangDoc", "btnChenSoTrang"]),
        ("grpKhoiDong", "boxFixActions", ["btnXoaTrangThua"]),
        ("grpAbout", "boxSettingsTools", ["btnThietLap", "mnuThongTinTienIch"]),
        ("grpChinhTaSo", "boxLanguageActions", ["mnuBoDau", "btnDoiDauThapPhan", "btnTuDienCaNhan"]),
    ]
    for control_id in ("ddQuyDinh", "ddLoaiVanBan"):
        elements[control_id].set("visible", "false")
    for group_id, box_id, control_ids in columns:
        box = elements.get(box_id)
        if box is None:
            box = ET.SubElement(elements[group_id], f"{{{RIBBON_NAMESPACE}}}box",
                                {"id": box_id, "boxStyle": "vertical"})
            elements[box_id] = box
        for control_id in control_ids:
            control = elements[control_id]
            parent = next(parent for parent in root.iter() if control in list(parent))
            parent.remove(control)
            control.set("size", "normal")
            control.set("label", control.get("label", "").strip())
            box.append(control)
    # Two compact horizontal rows: scale and character spacing.
    scale_row = ET.SubElement(elements["boxCharacterTools"], f"{{{RIBBON_NAMESPACE}}}box",
                              {"id": "boxScaleIcons", "boxStyle": "horizontal"})
    for control_id, icon in (("btnScaleGiam", "CharacterSpacingCondensed"),
                             ("btnScale100", "CharacterSpacingNormal"),
                             ("btnScaleTang", "CharacterSpacingExpanded")):
        control = elements[control_id]
        elements["boxCharacterTools"].remove(control)
        control.set("showLabel", "true")
        control.set("showImage", "false")
        control.set("label", {"btnScaleGiam": "A−", "btnScale100": "100%", "btnScaleTang": "A+"}[control_id])
        control.attrib.pop("imageMso", None)
        control.attrib.pop("getImage", None)
        scale_row.append(control)
    spacing = elements["boxGianChu"]
    parent = next(parent for parent in root.iter() if spacing in list(parent))
    parent.remove(spacing)
    elements["boxCharacterTools"].append(spacing)
    # Pair controls by column instead of two independently measured rows.
    # The wider 100% caption must not shift the increase/decrease positions.
    character_tools = elements["boxCharacterTools"]
    character_tools.set("boxStyle", "horizontal")
    character_tools.remove(scale_row)
    character_tools.remove(spacing)
    for column_id, scale_id, spacing_id in (
        ("boxCharacterDecrease", "btnScaleGiam", "btnCoChu"),
        ("boxCharacterReset", "btnScale100", "btnGianChuNormal"),
        ("boxCharacterIncrease", "btnScaleTang", "btnGianChuRa"),
    ):
        column = ET.SubElement(character_tools, f"{{{RIBBON_NAMESPACE}}}box",
                               {"id": column_id, "boxStyle": "vertical"})
        scale_control = elements[scale_id]
        spacing_control = elements[spacing_id]
        scale_row.remove(scale_control)
        spacing.remove(spacing_control)
        # Small text-only buttons preserve the two-row layout. Large buttons
        # consume the Ribbon height and cannot be stacked into these columns.
        for control in (scale_control, spacing_control):
            control.set("size", "normal")
            control.set("showLabel", "true")
            control.set("showImage", "false")
            control.attrib.pop("imageMso", None)
            control.attrib.pop("getImage", None)
        spacing_control.set("label", spacing_control.get("label", "").strip())
        if scale_id == "btnScale100":
            # Ribbon normal-button icons are square. Keep 100% as native text;
            # never squeeze a wide text bitmap into that icon slot.
            scale_control.set("label", "100%")
            spacing_control.set("label", "⬤")
            spacing_control.set("showLabel", "true")
            spacing_control.set("showImage", "false")
            spacing_control.attrib.pop("getImage", None)
        if scale_id in ("btnScaleGiam", "btnScaleTang"):
            scale_control.set("label", "A−" if scale_id == "btnScaleGiam" else "A+")
        column.append(scale_control)
        column.append(spacing_control)
    elements["grpKhoiDong"].set("label", "Kiểm tra và sửa lỗi")
    elements["grpAbout"].set("label", "Thiết lập")
    tab = next(parent for parent in root.iter() if elements["grpAbout"] in list(parent))
    info_group = ET.SubElement(tab, f"{{{RIBBON_NAMESPACE}}}group",
                               {"id": "grpInformation", "label": "Thông tin"})
    # Independent controls can use the available width without collapsing a
    # single indivisible vertical box containing both settings and information.
    for control_id in ("btnThietLap", "mnuThongTinTienIch"):
        control = elements[control_id]
        elements["boxSettingsTools"].remove(control)
        control.set("size", "large")
        (info_group if control_id == "mnuThongTinTienIch" else elements["grpAbout"]).append(control)
    elements["grpAbout"].remove(elements["boxSettingsTools"])
    for box_id in ("boxDocDuLieu", "boxChenTrang"):
        box = elements[box_id]
        if not any(child.get("visible") != "false" for child in box):
            box.set("visible", "false")
    # Remove historical label padding; native Ribbon handles spacing itself.
    for element in root.iter():
        if "label" in element.attrib:
            element.set("label", element.get("label", "").strip(" \t\r\n"))


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
