from __future__ import annotations

import argparse
import json
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
DOTNET = ROOT / ".tools" / "dotnet" / "dotnet.exe"
NUGET_SOURCE = "https://api.nuget.org/v3/index.json"
TARGETS = (
    Path("ChuanHoa.slnx"),
    Path("tests/ChuanHoa.Infrastructure.IntegrationTests/ChuanHoa.Infrastructure.IntegrationTests.csproj"),
)


def normalize_project(path: str) -> str:
    candidate = Path(path)
    try:
        return candidate.resolve().relative_to(ROOT.resolve()).as_posix()
    except (OSError, ValueError):
        return candidate.as_posix()


def vulnerability_records(value: object, project: str = "") -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    if isinstance(value, dict):
        current_project = project
        if isinstance(value.get("path"), str) and str(value["path"]).lower().endswith(".csproj"):
            current_project = normalize_project(str(value["path"]))
        vulnerabilities = value.get("vulnerabilities")
        if isinstance(vulnerabilities, list) and vulnerabilities:
            records.append(
                {
                    "project": current_project,
                    "package": value.get("id") or value.get("name") or "<unknown>",
                    "resolvedVersion": value.get("resolvedVersion") or value.get("resolved"),
                    "vulnerabilities": vulnerabilities,
                }
            )
        for child in value.values():
            records.extend(vulnerability_records(child, current_project))
    elif isinstance(value, list):
        for child in value:
            records.extend(vulnerability_records(child, project))
    return records


def audit_target(target: Path) -> tuple[dict[str, object], list[dict[str, object]]]:
    command = [
        str(DOTNET),
        "package",
        "list",
        "--project",
        str(target),
        "--vulnerable",
        "--include-transitive",
        "--source",
        NUGET_SOURCE,
        "--format",
        "json",
        "--output-version",
        "1",
        "--no-restore",
    ]
    process = subprocess.run(
        command,
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    record: dict[str, object] = {
        "target": target.as_posix(),
        "command": "dotnet package list --project " + target.as_posix()
        + " --vulnerable --include-transitive --source " + NUGET_SOURCE
        + " --format json --output-version 1 --no-restore",
        "exitCode": process.returncode,
    }
    if process.returncode != 0:
        record["error"] = (process.stderr or process.stdout).strip()[:2000]
        return record, []
    try:
        payload = json.loads(process.stdout)
    except json.JSONDecodeError as error:
        record["error"] = f"invalid JSON output: {error}"
        return record, []
    projects = payload.get("projects", [])
    record["projectCount"] = len(projects) if isinstance(projects, list) else 0
    sources = payload.get("sources", [])
    record["sources"] = sources
    if NUGET_SOURCE not in sources:
        record["error"] = "official NuGet source was not present in audit output"
    return record, vulnerability_records(payload)


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
        sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser()
    parser.add_argument("--write-evidence", type=Path)
    args = parser.parse_args()
    failures: list[str] = []
    if not DOTNET.is_file():
        failures.append("local .NET SDK was not found")

    audits: list[dict[str, object]] = []
    vulnerabilities: list[dict[str, object]] = []
    if not failures:
        for target in TARGETS:
            record, found = audit_target(target)
            audits.append(record)
            vulnerabilities.extend(found)
            if record.get("exitCode") != 0 or record.get("error"):
                failures.append(f"NuGet audit failed for {target.as_posix()}: {record.get('error')}")
    audited_projects = sum(int(record.get("projectCount", 0)) for record in audits)
    if not failures and audited_projects != 14:
        failures.append(f"expected 14 audited projects, found {audited_projects}")
    if vulnerabilities:
        failures.append(f"NuGet reported {len(vulnerabilities)} vulnerable package record(s)")

    status = "PASS" if not failures else "FAIL"
    evidence = {
        "schemaVersion": 1,
        "testId": "NUGET-VULNERABILITY-AUDIT-001",
        "status": status,
        "verifiedAtUtc": datetime.now(timezone.utc).isoformat(),
        "advisorySource": NUGET_SOURCE,
        "audits": audits,
        "auditedProjectCount": audited_projects,
        "vulnerablePackageCount": len(vulnerabilities),
        "vulnerabilities": vulnerabilities,
        "limitations": [
            "The result is a point-in-time NuGet advisory query and must be rerun for every release.",
            "It does not scan non-NuGet prerequisites, source code, the installer, containers or operating-system components.",
        ],
        "failures": failures,
    }
    if args.write_evidence:
        path = args.write_evidence if args.write_evidence.is_absolute() else ROOT / args.write_evidence
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(evidence, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    if failures:
        print("NUGET_VULNERABILITY_AUDIT: FAIL")
        for failure in failures:
            print("- " + failure)
        return 1
    print(
        "NUGET_VULNERABILITY_AUDIT: PASS "
        f"PROJECTS={audited_projects} VULNERABLE_PACKAGES={len(vulnerabilities)}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
