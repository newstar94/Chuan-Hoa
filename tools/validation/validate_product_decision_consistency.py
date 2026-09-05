from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[2]
CONTRACT_PATH = ROOT / "shared" / "contracts" / "ribbon" / "ribbon-contract.v1.json"
RIBBON_PATH = ROOT / "src" / "ChuanHoa.AddIn.Vsto" / "Ribbon" / "ChuanHoaRibbon.xml"
CAPABILITY_PATH = ROOT / "src" / "ChuanHoa.AddIn.Vsto" / "Runtime" / "WordDocumentCapabilityProvider.cs"
ANNOTATION_PATH = ROOT / "src" / "ChuanHoa.Client.Core" / "Annotations" / "AnnotationPlanner.cs"
DECISIONS_PATH = ROOT / "shared" / "docs" / "implementation" / "Decision_Register.md"
ADR_PATH = ROOT / "shared" / "docs" / "implementation" / "Architecture_Decisions.md"
RETIRED_CONTROLS = {
    "btnLuuDocx",
    "btnKieuI",
    "btnKieuY",
    "mnuIY",
    "btnDocDuLieu",
    "btnChenQrCode",
}
REQUIRED_CONTROLS = {"btnKieuOaUy", "btnKieuOaUy2", "btnTuDienCaNhan"}
REQUIRED_DECISIONS = {
    "DEC-031": "Tài liệu chưa lưu",
    "DEC-032": "Nội dung comment",
    "DEC-033": "AI giai đoạn hiện tại",
    "DEC-034": "Thành phần Ribbon retired",
}


def read(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def analyze(
    contract: dict,
    ribbon_source: str,
    capability_source: str,
    annotation_source: str,
    decisions: str,
    adr: str,
) -> list[str]:
    failures: list[str] = []
    counts = contract.get("counts", {})
    expected_counts = {
        "tabs": 1,
        "groups": 6,
        "buttons": 34,
        "menus": 3,
        "dropDowns": 2,
        "checkBoxes": 0,
        "interactiveControls": 39,
    }
    if counts != expected_counts:
        failures.append(f"Ribbon contract counts changed: {counts}")
    if contract.get("tab", {}).get("label") != "Chuẩn hóa":
        failures.append("Ribbon contract product label is not Chuẩn hóa")

    removed: set[str] = set()
    change_ids: set[str] = set()
    for change in contract.get("approvedProductChanges", []):
        change_ids.add(str(change.get("changeId", "")))
        removed.update(str(value) for value in change.get("removedControlIds", []))
    if not RETIRED_CONTROLS.issubset(removed):
        failures.append(f"retired controls missing from contract decisions: {sorted(RETIRED_CONTROLS - removed)}")
    for change_id in (
        "REMOVE_SAVE_AS_DOCX_AND_SUPPORT_DOC_DOCX",
        "REMOVE_MANUAL_READ_DATA_USE_COMMAND_SCOPED_ANALYSIS",
        "REMOVE_QR_FEATURE",
        "REMOVE_VIEW_OPTIONS_GROUP",
        "MERGE_TYPOGRAPHY_INTO_QUICK_SPELLING",
    ):
        if change_id not in change_ids:
            failures.append(f"approved product change missing: {change_id}")

    ribbon_root = ET.fromstring(ribbon_source)
    active_ids = {
        element.attrib["id"]
        for element in ribbon_root.iter()
        if "id" in element.attrib
    }
    active_retired = RETIRED_CONTROLS & active_ids
    if active_retired:
        failures.append(f"retired controls are active in Ribbon XML: {sorted(active_retired)}")
    missing_required = REQUIRED_CONTROLS - active_ids
    if missing_required:
        failures.append(f"required controls are missing from Ribbon XML: {sorted(missing_required)}")
    tabs = [element for element in ribbon_root.iter() if element.tag.endswith("tab")]
    if len(tabs) != 1 or tabs[0].attrib.get("label") != "Chuẩn hóa":
        failures.append("Ribbon XML must expose exactly one Chuẩn hóa tab")

    if not all(
        fragment in capability_source
        for fragment in (
            "var isSaved = !string.IsNullOrWhiteSpace(document.Path);",
            "var supported = isSaved",
            ": (saveFormat == 0 || saveFormat == 12 || saveFormat == 16 || saveFormat == 24 || string.IsNullOrEmpty(extension));",
        )
    ):
        failures.append("unsaved-document capability policy is not explicit in VSTO source")

    comment_match = re.search(
        r"private static string BuildComment\(AnnotationFinding finding\)\s*\{(?P<body>.*?)\n\s*\}\s*\n\s*private static bool AreEquivalent",
        annotation_source,
        flags=re.DOTALL,
    )
    if not comment_match:
        failures.append("could not isolate AnnotationPlanner.BuildComment")
    else:
        body = comment_match.group("body")
        for fragment in ('builder.Append("Hiện tại: ")', 'builder.Append("Yêu cầu đúng: ")'):
            if fragment not in body:
                failures.append(f"visible two-line comment contract is missing: {fragment}")
        for forbidden in ("RuleCode", "Severity", "Citation", "FindingId", "SourceFamily"):
            if forbidden in body:
                failures.append(f"visible comment leaks metadata: {forbidden}")

    for decision_id, label in REQUIRED_DECISIONS.items():
        pattern = rf"\|\s*{re.escape(decision_id)}\s*\|\s*LOCKED\s*\|\s*{re.escape(label)}\s*\|"
        if not re.search(pattern, decisions):
            failures.append(f"locked decision is missing or changed: {decision_id}")

    adr_lower = adr.lower()
    for fragment in (
        "tài liệu mới chưa lưu (`Document1`) vẫn dùng được toàn bộ chức năng",
        "Phần comment nhìn thấy chỉ có hai dòng",
        "Rule-based ở giai đoạn hiện tại, AI là nâng cấp mới",
    ):
        if fragment.lower() not in adr_lower:
            failures.append(f"current ADR statement missing: {fragment}")
    for stale in (
        "tài liệu chưa lưu hoặc extension/SaveFormat không khớp đều fail closed",
        "sau khi backend findings",
        "About, ba tùy chọn hiển thị",
    ):
        if stale in adr:
            failures.append(f"stale ADR statement remains: {stale}")
    return failures


def negative_self_tests(inputs: tuple[dict, str, str, str, str, str]) -> dict[str, bool]:
    contract, ribbon, capability, annotation, decisions, adr = inputs
    cases: dict[str, tuple[dict, str, str, str, str, str]] = {}

    bad_ribbon = ribbon.replace("</group>", '<button id="btnChenQrCode" /></group>', 1)
    cases["retired_ribbon_control"] = (contract, bad_ribbon, capability, annotation, decisions, adr)

    bad_capability = capability.replace("var supported = isSaved", "var supported = true", 1)
    cases["unsaved_capability_drift"] = (contract, ribbon, bad_capability, annotation, decisions, adr)

    bad_comment = annotation.replace(
        'builder.Append("Hiện tại: ")',
        'builder.Append(finding.RuleCode).Append(" Hiện tại: ")',
        1,
    )
    cases["comment_metadata_leak"] = (contract, ribbon, capability, bad_comment, decisions, adr)

    bad_decisions = decisions.replace("| DEC-033 | LOCKED |", "| DEC-033 | PENDING |", 1)
    cases["decision_unlock"] = (contract, ribbon, capability, annotation, bad_decisions, adr)

    return {name: bool(analyze(*case)) for name, case in cases.items()}


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
        sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser()
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--write-evidence", type=Path)
    args = parser.parse_args()

    inputs = (
        json.loads(read(CONTRACT_PATH)),
        read(RIBBON_PATH),
        read(CAPABILITY_PATH),
        read(ANNOTATION_PATH),
        read(DECISIONS_PATH),
        read(ADR_PATH),
    )
    failures = analyze(*inputs)
    self_tests = negative_self_tests(inputs) if args.self_test else {}
    failures.extend(f"negative self-test did not fail: {name}" for name, passed in self_tests.items() if not passed)

    evidence = {
        "schemaVersion": 1,
        "testId": "PRODUCT-DECISION-CONSISTENCY-001",
        "status": "PASS" if not failures else "FAIL",
        "verifiedOn": date.today().isoformat(),
        "productVersion": read(ROOT / "Directory.Build.props").split("<ProductVersion>", 1)[1].split("</ProductVersion>", 1)[0],
        "ribbonCounts": inputs[0].get("counts"),
        "requiredLockedDecisions": sorted(REQUIRED_DECISIONS),
        "negativeSelfTests": self_tests,
        "hashes": {
            "ribbonContractSha256": sha256(CONTRACT_PATH),
            "ribbonXmlSha256": sha256(RIBBON_PATH),
            "decisionRegisterSha256": sha256(DECISIONS_PATH),
            "architectureDecisionsSha256": sha256(ADR_PATH),
        },
        "failures": failures,
    }
    if args.write_evidence:
        path = args.write_evidence if args.write_evidence.is_absolute() else ROOT / args.write_evidence
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(evidence, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    if failures:
        print("PRODUCT_DECISION_CONSISTENCY: FAIL")
        for failure in failures:
            print("- " + failure)
        return 1
    print(
        "PRODUCT_DECISION_CONSISTENCY: PASS "
        f"CONTROLS={inputs[0]['counts']['interactiveControls']} "
        f"DECISIONS={len(REQUIRED_DECISIONS)} "
        f"NEGATIVE_TESTS={sum(self_tests.values())}/{len(self_tests)}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
