from __future__ import annotations

import hashlib
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
RECONCILIATION_PATH = (
    ROOT
    / "shared"
    / "docs"
    / "implementation"
    / "evidence"
    / "rule_reconciliation.json"
)
RULE_DATA_PATH = ROOT / "shared" / "vba_extracted" / "RuleData.bas.bas"
OUTPUT_PATH = ROOT / "shared" / "rules" / "canonical" / "baseline-draft.v1.json"
EVIDENCE_PATH = (
    ROOT
    / "shared"
    / "docs"
    / "implementation"
    / "evidence"
    / "canonical_rule_baseline.json"
)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def regime_for(rule_code: str) -> str:
    if rule_code.startswith("ND30-"):
        return "ND30"
    if rule_code.startswith("QD1989-"):
        return "VIETNAMESE_ORTHOGRAPHY"
    if rule_code.startswith("LOCAL-"):
        return "LOCAL_TYPOGRAPHY"
    return "UNKNOWN"


def main() -> None:
    reconciliation = json.loads(RECONCILIATION_PATH.read_text(encoding="utf-8"))
    route_index = {
        route["ruleCode"]: route
        for route in reconciliation["registeredRoutes"]
    }
    source_hash = sha256(RULE_DATA_PATH)
    created_at = "2026-09-01T00:00:00Z"
    rules = []
    status_counts = {
        "Unrouted": 0,
        "HardwiredNotChecked": 0,
        "BaselineLogicPath": 0,
    }

    for rule_code in reconciliation["vbaRuleCodes"]:
        route = route_index.get(rule_code)
        if route is None:
            implementation_status = "Unrouted"
            route_function = None
            source_symbol = "LoadRawCheckRules_RulesPart1..5"
        elif route["classification"] == "HARDWIRED_NOT_CHECKED":
            implementation_status = "HardwiredNotChecked"
            route_function = route["function"]
            source_symbol = route_function
        else:
            implementation_status = "BaselineLogicPath"
            route_function = route["function"]
            source_symbol = route_function

        status_counts[implementation_status] += 1
        rules.append(
            {
                "ruleCode": rule_code,
                "regimeCode": regime_for(rule_code),
                "lifecycleStatus": "Draft",
                "title": f"Baseline {rule_code}",
                "description": (
                    "Unreviewed VBA baseline definition. Semantic wording, legal traceability, "
                    "fixtures, detector behavior, and fix policy must be approved before publication."
                ),
                "provenance": {
                    "sourceKind": "VBA_BASELINE",
                    "sourcePath": "shared/vba_extracted/RuleData.bas.bas",
                    "sourceSymbol": source_symbol,
                    "sourceHash": source_hash,
                },
                "implementation": {
                    "status": implementation_status,
                    "routeFunction": route_function,
                    "detectorId": None,
                    "verifiedEngineVersion": None,
                },
                "legalReview": {
                    "status": "Unreviewed",
                    "authority": None,
                    "instrument": None,
                    "provision": None,
                    "reviewer": None,
                    "reviewedAtUtc": None,
                },
                "fixtures": {
                    "positiveFixtureIds": [],
                    "negativeFixtureIds": [],
                    "boundaryFixtureIds": [],
                },
                "fixPolicy": "Blocked",
                "documentTypeCodes": [],
                "effectiveFromUtc": None,
                "effectiveUntilUtc": None,
            }
        )

    release = {
        "schema": "chuanhoa.canonical-rule-release.v1",
        "releaseId": "VBA-BASELINE-2026-09-01",
        "status": "BaselineDraft",
        "createdAtUtc": created_at,
        "effectiveFromUtc": None,
        "effectiveUntilUtc": None,
        "sourceBaselineId": "CORRECTION-RULE-ROUTE-BASELINE-001",
        "rules": rules,
    }
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text(
        json.dumps(release, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    evidence = {
        "testId": "RULE-CANONICAL-BASELINE-001",
        "status": "PASS_DRAFT_ONLY",
        "releaseId": release["releaseId"],
        "ruleCount": len(rules),
        "implementationStatusCounts": status_counts,
        "legalReviewStatus": "Unreviewed",
        "publishable": False,
        "publicationBlockers": [
            "96 rules do not have approved legal traceability",
            "96 rules do not have positive, negative, and boundary fixtures",
            "75 logic paths are baseline observations, not verified engine implementations",
            "19 routes are hardwired to NotChecked",
            "2 rule definitions are not routed",
            "all fix policies remain Blocked",
        ],
        "outputSha256": sha256(OUTPUT_PATH),
    }
    EVIDENCE_PATH.write_text(
        json.dumps(evidence, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(evidence, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
