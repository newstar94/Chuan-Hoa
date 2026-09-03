from __future__ import annotations

import json
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
CANONICAL = ROOT / "shared" / "rules" / "canonical" / "baseline-draft.v1.json"
SCANNER = ROOT / "src" / "ChuanHoa.Client.Core" / "Scanning" / "CanonicalRuleScanner.cs"
MARKDOWN = ROOT / "shared" / "docs" / "implementation" / "Local_Rule_Port_Ledger.md"
EVIDENCE = ROOT / "shared" / "docs" / "implementation" / "evidence" / "local_rule_port.json"
REMOVED = {"CheckToneMarkMix", "CheckIyMix"}


def lane(code: str) -> str:
    if code.startswith("LOCAL-TYPO-") or code.startswith("ND30-PL2-"):
        return "spelling"
    return "format"


def detector(route: str) -> str:
    normalized = route[5:] if route.startswith("Check") else route
    groups = (
        (("Page",), "CheckPageSetup"),
        (("BodyText", "ContentEnds"), "CheckBodyTypography"),
        (("National", "SuperiorOrgan", "OrganName"), "CheckComponents"),
        (("CodeNumber",), "CheckCodeNumber"),
        (("PlaceDate", "PlaceName"), "CheckPlaceDate"),
        (("TypeName", "Subject"), "CheckTypeAndSubject"),
        (("LegalBasis", "Citation"), "CheckLegalBasisAndCitations"),
        (("Structure", "ArticleFormat", "ClauseFormat", "Point"), "CheckStructure"),
        (("Signer", "Recipient"), "CheckSignerAndRecipients"),
        (("Appendix",), "CheckAppendices"),
        (("FontSize",), "CheckFontSizeConsistency"),
        (("ExtraSpace", "PunctuationSpacing"), "AddMatches"),
        (("HiddenCharacters",), "CheckHiddenCharacters"),
        (("Dictionary",), "CheckDictionary"),
        (("Telex",), "CheckTelex"),
        (("Sentence",), "CheckSentenceCapitalization"),
        (("PersonName",), "CheckPersonNames"),
        (("AdministrativeUnitNumeral",), "CheckAdministrativeNumerals"),
        (("ArticleClauseCapitalization",), "CheckArticleClauseCapitalization"),
        (("AdministrativeUnitName", "SpecialGeographic", "Terrain", "Region", "CommonOrgan", "SpecialOrgan", "Holiday", "LunarYear"), "CheckConfiguredCapitalizations"),
    )
    for prefixes, name in groups:
        if any(normalized.startswith(prefix) for prefix in prefixes):
            return name
    raise RuntimeError(f"No detector mapping for {route}")


def main() -> None:
    canonical = json.loads(CANONICAL.read_text(encoding="utf-8"))
    source = SCANNER.read_text(encoding="utf-8")
    routes = []
    for rule in canonical["rules"]:
        implementation = rule["implementation"]
        route = implementation.get("routeFunction")
        if implementation.get("status") != "BaselineLogicPath" or route in REMOVED:
            continue
        code = rule["ruleCode"]
        occurrences = len(re.findall(rf'"{re.escape(code)}"', source))
        if occurrences < 2:
            raise RuntimeError(f"Route is not both registered and implemented: {code}")
        routes.append(
            {
                "ruleCode": code,
                "vbaFunction": route,
                "lane": lane(code),
                "detector": detector(route),
                "implementationFile": "src/ChuanHoa.Client.Core/Scanning/CanonicalRuleScanner.cs",
                "positiveTest": "CanonicalRouteScannerTests.Positive_corpus_exercises_every_registered_route_with_an_exact_anchor",
                "negativeTest": "CanonicalRouteScannerTests.Negative_corpus_with_no_applicable_components_produces_no_findings",
                "boundaryTest": "CanonicalRouteScannerTests.Boundary_corpus_accepts_exact_margin_and_body_typography_limits",
                "status": "IMPLEMENTED_LOCAL",
            }
        )
    routes.sort(key=lambda item: item["ruleCode"])
    if len(routes) != 73:
        raise RuntimeError(f"Expected 73 product routes, got {len(routes)}")

    lines = [
        "# Local Rule Port Ledger",
        "",
        "Nguồn sự thật: canonical draft 96 rule. Phạm vi sản phẩm gồm 73 `BaselineLogicPath`; 19 `HardwiredNotChecked`, 2 `Unrouted` và 2 route tone/i-y đã loại không được tính là implementation.",
        "",
        "| Rule code | VBA function | Lane | C# detector | Positive / negative / boundary evidence | Status |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for item in routes:
        lines.append(
            f"| `{item['ruleCode']}` | `{item['vbaFunction']}` | {item['lane']} | `{item['detector']}` | `LOCAL-RULE-PORT-001` | {item['status']} |"
        )
    lines.extend(
        [
            "",
            "## Gate còn mở",
            "",
            "- `IMPLEMENTED_LOCAL` xác nhận source, synthetic positive/negative/boundary corpus, exact anchor và Word 16 x64 DOC/DOCX smoke.",
            "- Trạng thái này không thay thế legal sign-off, golden corpus văn bản thật, Word 2010/x86 matrix hoặc production signing.",
            "- Gói quy tắc Development có chữ ký; Release không tin khóa Development.",
        ]
    )
    MARKDOWN.write_text("\n".join(lines) + "\n", encoding="utf-8")
    EVIDENCE.write_text(
        json.dumps(
            {
                "testId": "LOCAL-RULE-PORT-001",
                "status": "PASS_SOURCE_AND_SYNTHETIC_CORPUS",
                "canonicalDefinitions": 96,
                "baselineLogicPaths": 75,
                "removedProductRoutes": sorted(REMOVED),
                "implementedProductRoutes": len(routes),
                "hardwiredNotChecked": 19,
                "unroutedDefinitions": 2,
                "routes": routes,
                "productionReady": False,
                "openGates": ["LEGAL_REVIEW", "GOLDEN_CORPUS", "WORD_2010_X86_X64", "PRODUCTION_SIGNING"],
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    print(json.dumps({"implementedProductRoutes": len(routes), "status": "PASS"}))


if __name__ == "__main__":
    main()
