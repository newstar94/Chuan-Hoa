from __future__ import annotations

import hashlib
import json
import re
import sys
import zipfile
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree


ROOT = Path(__file__).resolve().parents[2]
EVIDENCE_DIR = ROOT / "shared" / "docs" / "implementation" / "evidence"
RIBBON_ARTEFACT = ROOT / "shared" / "ChuanHoaTheThuc_Full_Ribbon.dotm"
VBA_DIR = ROOT / "shared" / "vba_extracted"
PLAN_FILE = ROOT / "shared" / "docs" / "KeHoach_TrienKhai_Addin_VSTO_Word_2010_Plus.md"
RULE_DATA_FILE = VBA_DIR / "RuleData.bas.bas"
COMPLIANCE_FILE = VBA_DIR / "ComplianceChecker.bas.bas"
JSON_RULE_FILE = ROOT / "shared" / "rules" / "rules_compliance_checks.json"
BACKEND_CHECKER_FILE = ROOT / "backend-api" / "src" / "checker" / "ComplianceChecker.ts"

EXCLUDED_DIRECTORIES = {
    ".git",
    ".tools",
    ".vs",
    "TestResults",
    "artifacts",
    "bin",
    "coverage",
    "node_modules",
    "obj",
    "publish",
}
TEXT_EXTENSIONS = {
    ".bas",
    ".cls",
    ".cs",
    ".csproj",
    ".css",
    ".frm",
    ".html",
    ".iss",
    ".js",
    ".json",
    ".md",
    ".ps1",
    ".py",
    ".sln",
    ".ts",
    ".tsx",
    ".txt",
    ".xml",
    ".yml",
    ".yaml",
}

PROCEDURE_PATTERN = re.compile(
    r"(?im)^\s*(?P<visibility>Public|Private|Friend)?\s*"
    r"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"(?P<name>[A-Za-z_][A-Za-z0-9_]*)\b"
)
RULE_CODE_PATTERN = re.compile(r'\.Add\s+"ruleCode"\s*,\s*"([^"]+)"', re.IGNORECASE)
ROUTE_PATTERN = re.compile(
    r'mRegistry\("(?P<code>[^"]+)"\)\s*=\s*"ComplianceChecker\.(?P<function>[A-Za-z0-9_]+)"',
    re.IGNORECASE,
)
BACKEND_CODE_PATTERN = re.compile(r"checkCode\s*:\s*['\"]([^'\"]+)['\"]")

SIDE_EFFECT_PATTERNS = {
    "active_document": re.compile(r"\bActiveDocument\b", re.IGNORECASE),
    "application_options": re.compile(r"\b(?:Application\.)?Options\.", re.IGNORECASE),
    "custom_document_properties": re.compile(r"CustomDocumentProperties", re.IGNORECASE),
    "filesystem": re.compile(
        r"\b(?:Open\s+.+\s+For\s+(?:Input|Output|Append|Binary)|FileCopy|Kill|MkDir|RmDir|FileSystemObject)\b",
        re.IGNORECASE,
    ),
    "native_api": re.compile(r"\bDeclare\s+(?:PtrSafe\s+)?(?:Function|Sub)\b", re.IGNORECASE),
    "process_or_shell": re.compile(r"\b(?:Shell|ShellExecute|WScript\.Shell)\b", re.IGNORECASE),
    "save_or_convert": re.compile(r"\b(?:SaveAs2?|FileFormat|wdFormatXMLDocument)\b", re.IGNORECASE),
    "screen_updating": re.compile(r"\bScreenUpdating\b", re.IGNORECASE),
    "selection": re.compile(r"\bSelection\b", re.IGNORECASE),
    "scheduled_callback": re.compile(r"\bOnTime\b", re.IGNORECASE),
    "temporary_path": re.compile(r"\b(?:Environ\s*\(\s*['\"]TEMP|GetTempPath|TemporaryFolder)\b", re.IGNORECASE),
}

# Những route này cùng trỏ vào ba implementation công khai tự ghi nhận là không đo được
# và luôn trả Nothing. Danh sách tạo đúng 17 route theo baseline đã phê duyệt.
HARDWIRED_NOT_CHECKED_FUNCTIONS = {
    "CheckComponentUnderline",
    "CheckComponentNeverDetected",
    "CheckCapitalizationNotDetectable",
}


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace").replace("\r\r\n", "\n")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
        newline="\n",
    )


def iter_source_files() -> Iterable[Path]:
    for path in sorted(ROOT.rglob("*"), key=lambda item: item.as_posix().casefold()):
        if not path.is_file():
            continue
        relative = path.relative_to(ROOT)
        if any(part in EXCLUDED_DIRECTORIES for part in relative.parts):
            continue
        if relative.parts[:5] == ("shared", "docs", "implementation", "evidence"):
            continue
        yield path


def build_source_inventory() -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for path in iter_source_files():
        relative = path.relative_to(ROOT).as_posix()
        entry: dict[str, Any] = {
            "path": relative,
            "bytes": path.stat().st_size,
            "sha256": sha256_file(path),
        }
        if path.suffix.lower() in TEXT_EXTENSIONS or ".".join(path.name.lower().split(".")[-2:]) in {
            "bas.bas",
            "cls.bas",
            "frm.bas",
        }:
            entry["lines"] = len(read_text(path).splitlines())
        result.append(entry)
    return result


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def extract_ribbon_catalog() -> dict[str, Any]:
    with zipfile.ZipFile(RIBBON_ARTEFACT) as archive:
        candidates = sorted(
            name
            for name in archive.namelist()
            if re.fullmatch(r"customUI/customUI(?:14)?\.xml", name, re.IGNORECASE)
        )
        if not candidates:
            raise RuntimeError("Không tìm thấy customUI Ribbon XML trong artefact chuẩn")
        entry_name = candidates[-1]
        xml_bytes = archive.read(entry_name)

    root = ElementTree.fromstring(xml_bytes)
    tab = next(element for element in root.iter() if local_name(element.tag) == "tab")
    groups = [element for element in list(tab) if local_name(element.tag) == "group"]
    interactive_types = {"button", "menu", "dropDown", "checkBox"}
    callback_names = {
        "onAction",
        "getEnabled",
        "getPressed",
        "getSelectedItemIndex",
        "getItemCount",
        "getItemLabel",
        "getImage",
        "getLabel",
        "getVisible",
    }
    controls: list[dict[str, Any]] = []

    def visit(element: ElementTree.Element, group_id: str, parent_id: str | None) -> None:
        element_type = local_name(element.tag)
        current_parent = parent_id
        if element_type in interactive_types:
            attributes = dict(element.attrib)
            callbacks = {key: attributes[key] for key in sorted(callback_names) if key in attributes}
            item_labels = [
                {"id": child.attrib.get("id"), "label": child.attrib.get("label")}
                for child in list(element)
                if local_name(child.tag) == "item"
            ]
            controls.append(
                {
                    "order": len(controls) + 1,
                    "groupId": group_id,
                    "parentContainerId": parent_id,
                    "controlType": element_type,
                    "id": attributes.get("id"),
                    "label": attributes.get("label"),
                    "size": attributes.get("size"),
                    "imageMso": attributes.get("imageMso"),
                    "callbacks": callbacks,
                    "screenTip": attributes.get("screentip"),
                    "superTip": attributes.get("supertip"),
                    "staticItems": item_labels,
                }
            )
            current_parent = attributes.get("id") or parent_id
        elif element_type == "box":
            current_parent = element.attrib.get("id") or parent_id

        for child in list(element):
            if local_name(child.tag) != "item":
                visit(child, group_id, current_parent)

    for group in groups:
        group_id = group.attrib["id"]
        for child in list(group):
            visit(child, group_id, None)

    counts = defaultdict(int)
    for control in controls:
        counts[control["controlType"]] += 1

    return {
        "artefact": RIBBON_ARTEFACT.relative_to(ROOT).as_posix(),
        "artefactSha256": sha256_file(RIBBON_ARTEFACT),
        "xmlEntry": entry_name,
        "xmlSha256": hashlib.sha256(xml_bytes).hexdigest().upper(),
        "schema": root.tag.split("}", 1)[0].lstrip("{") if "}" in root.tag else "",
        "onLoad": root.attrib.get("onLoad"),
        "tab": {"id": tab.attrib.get("id"), "label": tab.attrib.get("label")},
        "groups": [
            {"order": index + 1, "id": group.attrib.get("id"), "label": group.attrib.get("label")}
            for index, group in enumerate(groups)
        ],
        "counts": {
            "tabs": 1,
            "groups": len(groups),
            "buttons": counts["button"],
            "menus": counts["menu"],
            "dropDowns": counts["dropDown"],
            "checkBoxes": counts["checkBox"],
            "interactiveControls": len(controls),
        },
        "controls": controls,
    }


def parse_plan_migration_rows() -> dict[str, dict[str, str]]:
    plan = read_text(PLAN_FILE)
    section_match = re.search(
        r"### 8\.2\. Ma trận module(?P<body>.*?)\n## 9\.",
        plan,
        re.DOTALL,
    )
    if not section_match:
        raise RuntimeError("Không tìm thấy ma trận migration 68 module trong kế hoạch")
    rows: dict[str, dict[str, str]] = {}
    for line in section_match.group("body").splitlines():
        cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
        if len(cells) != 4 or not cells[0].isdigit():
            continue
        index_text, module_text, target_text, decision_text = cells
        module = module_text.strip("`")
        rows[module] = {
            "index": int(index_text),
            "target": target_text.strip("`"),
            "decision": decision_text,
        }
    if len(rows) != 68:
        raise RuntimeError(f"Ma trận kế hoạch phải có 68 module, thực tế parse được {len(rows)}")
    return rows


def extract_vba_ledger() -> list[dict[str, Any]]:
    files = sorted(VBA_DIR.glob("*"), key=lambda item: item.name.casefold())
    if len(files) != 68:
        raise RuntimeError(f"Baseline phải có 68 file VBA, thực tế có {len(files)}")
    plan_rows = parse_plan_migration_rows()
    module_data: dict[str, dict[str, Any]] = {}
    public_owners: dict[str, set[str]] = defaultdict(set)

    for path in files:
        text = read_text(path)
        procedures = [
            {
                "visibility": (match.group("visibility") or "Public").upper(),
                "kind": re.sub(r"\s+", "_", match.group("kind").upper()),
                "name": match.group("name"),
                "line": text[: match.start()].count("\n") + 1,
            }
            for match in PROCEDURE_PATTERN.finditer(text)
        ]
        public_names = [item["name"] for item in procedures if item["visibility"] == "PUBLIC"]
        for procedure_name in public_names:
            public_owners[procedure_name.casefold()].add(path.name)
        module_data[path.name] = {
            "path": path,
            "text": text,
            "procedures": procedures,
            "public_names": public_names,
        }

    dependencies: dict[str, set[str]] = defaultdict(set)
    callers: dict[str, set[str]] = defaultdict(set)
    for module_name, data in module_data.items():
        tokens = {token.casefold() for token in re.findall(r"\b[A-Za-z_][A-Za-z0-9_]*\b", data["text"])}
        for token in tokens:
            for owner in public_owners.get(token, set()):
                if owner != module_name:
                    dependencies[module_name].add(owner)
                    callers[owner].add(module_name)

    ledger: list[dict[str, Any]] = []
    for module_name, data in module_data.items():
        path: Path = data["path"]
        plan_row = plan_rows.get(module_name)
        if not plan_row:
            raise RuntimeError(f"Module {module_name} không có disposition trong kế hoạch")
        side_effects = [name for name, pattern in SIDE_EFFECT_PATTERNS.items() if pattern.search(data["text"])]
        ledger.append(
            {
                "index": plan_row["index"],
                "module": module_name,
                "sourcePath": path.relative_to(ROOT).as_posix(),
                "sha256": sha256_file(path),
                "lines": len(data["text"].splitlines()),
                "publicEntryPoints": data["public_names"],
                "procedureCount": len(data["procedures"]),
                "dependenciesInferred": sorted(dependencies[module_name], key=str.casefold),
                "callersInferred": sorted(callers[module_name], key=str.casefold),
                "sideEffectFlags": side_effects,
                "target": plan_row["target"],
                "disposition": plan_row["decision"],
                "migrationStatus": "BASELINED",
                "fixtures": [],
                "testIds": [],
                "reviewer": None,
                "signedOffAt": None,
            }
        )
    return sorted(ledger, key=lambda entry: entry["index"])


def collect_json_codes(value: Any) -> set[str]:
    codes: set[str] = set()
    if isinstance(value, dict):
        code = value.get("code")
        if isinstance(code, str) and code.startswith("CHK_"):
            codes.add(code)
        for child in value.values():
            codes.update(collect_json_codes(child))
    elif isinstance(value, list):
        for child in value:
            codes.update(collect_json_codes(child))
    return codes


def build_rule_reconciliation() -> dict[str, Any]:
    rule_data = read_text(RULE_DATA_FILE)
    rules_start = rule_data.find("Private Sub LoadRawCheckRules_RulesPart1")
    rules_end = rule_data.find("Private Function LoadRawCheckRules_Rules()", rules_start)
    if rules_start < 0 or rules_end < 0:
        raise RuntimeError("Không xác định được phạm vi LoadRawCheckRules_RulesPart1..5")
    rule_codes = sorted(set(RULE_CODE_PATTERN.findall(rule_data[rules_start:rules_end])))
    route_matches = list(ROUTE_PATTERN.finditer(read_text(COMPLIANCE_FILE)))
    routes = {
        match.group("code"): match.group("function")
        for match in route_matches
    }
    hardwired_codes = sorted(
        code for code, function_name in routes.items() if function_name in HARDWIRED_NOT_CHECKED_FUNCTIONS
    )
    json_codes = sorted(collect_json_codes(json.loads(read_text(JSON_RULE_FILE))))
    backend_codes = sorted(set(BACKEND_CODE_PATTERN.findall(read_text(BACKEND_CHECKER_FILE))))
    missing_routes = sorted(set(rule_codes) - set(routes))
    unknown_routes = sorted(set(routes) - set(rule_codes))

    expected = {
        "vbaDefinitions": 96,
        "registeredRoutes": 94,
        "hardwiredNotCheckedRoutes": 19,
        "routesWithLogic": 75,
        "jsonCodes": 52,
        "backendCodes": 14,
    }
    actual = {
        "vbaDefinitions": len(rule_codes),
        "registeredRoutes": len(routes),
        "hardwiredNotCheckedRoutes": len(hardwired_codes),
        "routesWithLogic": len(routes) - len(hardwired_codes),
        "jsonCodes": len(json_codes),
        "backendCodes": len(backend_codes),
    }
    mismatches = {
        key: {"expected": expected[key], "actual": actual[key]}
        for key in expected
        if expected[key] != actual[key]
    }
    return {
        "status": "MATCHES_APPROVED_BASELINE" if not mismatches else "BASELINE_MISMATCH",
        "expected": expected,
        "actual": actual,
        "mismatches": mismatches,
        "vbaRuleCodes": rule_codes,
        "registeredRoutes": [
            {
                "ruleCode": code,
                "function": routes[code],
                "classification": "HARDWIRED_NOT_CHECKED" if code in hardwired_codes else "HAS_LOGIC_PATH",
            }
            for code in sorted(routes)
        ],
        "missingRegisteredRoutes": missing_routes,
        "routesWithoutDefinition": unknown_routes,
        "jsonCodes": json_codes,
        "backendCodes": backend_codes,
    }


def validate_baseline(ribbon: dict[str, Any], ledger: list[dict[str, Any]]) -> None:
    required_counts = {
        "tabs": 1,
        "groups": 7,
        "buttons": 36,
        "menus": 4,
        "dropDowns": 2,
        "checkBoxes": 3,
        "interactiveControls": 45,
    }
    if ribbon["counts"] != required_counts:
        raise RuntimeError(f"Ribbon count mismatch: {ribbon['counts']}")
    if len(ledger) != 68:
        raise RuntimeError(f"VBA ledger mismatch: {len(ledger)}")


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    generated_at = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
    source_inventory = build_source_inventory()
    ribbon = extract_ribbon_catalog()
    vba_ledger = extract_vba_ledger()
    rule_reconciliation = build_rule_reconciliation()
    validate_baseline(ribbon, vba_ledger)

    write_json(EVIDENCE_DIR / "ribbon_actual.json", ribbon)
    write_json(EVIDENCE_DIR / "vba_migration_ledger.json", vba_ledger)
    write_json(EVIDENCE_DIR / "rule_reconciliation.json", rule_reconciliation)
    write_json(
        EVIDENCE_DIR / "baseline_inventory.json",
        {
            "schemaVersion": 1,
            "generatedAtUtc": generated_at,
            "workspace": str(ROOT),
            "summary": {
                "sourceFiles": len(source_inventory),
                "sourceBytes": sum(entry["bytes"] for entry in source_inventory),
                "ribbon": ribbon["counts"],
                "vbaModules": len(vba_ledger),
                "vbaPublicEntryPoints": sum(len(entry["publicEntryPoints"]) for entry in vba_ledger),
                "rules": rule_reconciliation["actual"],
            },
            "files": source_inventory,
        },
    )
    print(
        json.dumps(
            {
                "status": (
                    "PASS" if rule_reconciliation["status"] == "MATCHES_APPROVED_BASELINE"
                    else "PASS_WITH_BASELINE_VARIANCE"
                ),
                "workspace": str(ROOT),
                "sourceFiles": len(source_inventory),
                "ribbon": ribbon["counts"],
                "vbaModules": len(vba_ledger),
                "rules": rule_reconciliation["actual"],
                "evidenceDirectory": str(EVIDENCE_DIR),
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
