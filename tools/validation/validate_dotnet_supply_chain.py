from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[2]
PROPS = ROOT / "Directory.Build.props"
GENERATOR = ROOT / "tools" / "validation" / "generate_dotnet_sbom.py"
CI_WORKFLOW = ROOT / ".github" / "workflows" / "source-quality.yml"
DEPENDABOT = ROOT / ".github" / "dependabot.yml"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def package_projects() -> list[Path]:
    result: list[Path] = []
    for source_root in (ROOT / "src", ROOT / "tests"):
        for path in source_root.rglob("*.csproj"):
            if "<PackageReference" in path.read_text(encoding="utf-8-sig"):
                result.append(path)
    return sorted(result)


def validate_locks() -> tuple[list[str], list[dict[str, str]]]:
    failures: list[str] = []
    records: list[dict[str, str]] = []
    for project in package_projects():
        project_root = ET.parse(project).getroot()
        direct: dict[str, str] = {}
        for reference in project_root.iter("PackageReference"):
            package_id = reference.attrib.get("Include") or reference.attrib.get("Update")
            version = reference.attrib.get("Version") or reference.findtext("Version") or ""
            if not package_id:
                continue
            if not re.fullmatch(r"[0-9]+(?:\.[0-9A-Za-z-]+)+", version):
                failures.append(
                    f"non-exact PackageReference version {package_id}={version!r}: "
                    f"{project.relative_to(ROOT).as_posix()}"
                )
            direct[package_id.lower()] = version

        lock_path = project.parent / "packages.lock.json"
        if not lock_path.is_file():
            failures.append(f"missing lock: {lock_path.relative_to(ROOT).as_posix()}")
            continue
        try:
            lock = json.loads(lock_path.read_text(encoding="utf-8-sig"))
        except (OSError, json.JSONDecodeError) as error:
            failures.append(f"invalid lock {lock_path.relative_to(ROOT).as_posix()}: {error}")
            continue
        if lock.get("version") != 1:
            failures.append(f"unsupported lock version: {lock_path.relative_to(ROOT).as_posix()}")
        dependencies = lock.get("dependencies", {})
        if not isinstance(dependencies, dict) or not dependencies:
            failures.append(f"lock has no frameworks: {lock_path.relative_to(ROOT).as_posix()}")
            continue
        for framework, entries in dependencies.items():
            if not isinstance(entries, dict):
                failures.append(f"invalid framework dependency map {framework}: {lock_path.name}")
                continue
            normalized = {name.lower(): data for name, data in entries.items()}
            for package_id, declared_version in direct.items():
                metadata = normalized.get(package_id)
                if not isinstance(metadata, dict):
                    failures.append(
                        f"direct package missing from lock {package_id}: "
                        f"{lock_path.relative_to(ROOT).as_posix()}"
                    )
                    continue
                if not str(metadata.get("type", "")).lower().startswith("direct"):
                    failures.append(f"package is not Direct in lock: {package_id} ({framework})")
                if metadata.get("resolved") != declared_version:
                    failures.append(
                        f"resolved direct version differs for {package_id}: "
                        f"declared={declared_version} resolved={metadata.get('resolved')}"
                    )
        records.append(
            {
                "project": project.relative_to(ROOT).as_posix(),
                "lockFile": lock_path.relative_to(ROOT).as_posix(),
                "sha256": sha256(lock_path),
            }
        )
    return failures, records


def generate_twice() -> tuple[list[str], dict, str]:
    failures: list[str] = []
    with tempfile.TemporaryDirectory(prefix="ChuanHoa-Sbom-") as temporary:
        first = Path(temporary) / "first.cdx.json"
        second = Path(temporary) / "second.cdx.json"
        for output in (first, second):
            process = subprocess.run(
                [sys.executable, str(GENERATOR), "--output", str(output)],
                cwd=ROOT,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            if process.returncode != 0:
                failures.append(process.stdout.strip() or process.stderr.strip())
        if failures:
            return failures, {}, ""
        first_hash = sha256(first)
        second_hash = sha256(second)
        if first_hash != second_hash:
            failures.append("SBOM output is not deterministic for unchanged lock files")
        sbom = json.loads(first.read_text(encoding="utf-8"))
        if sbom.get("bomFormat") != "CycloneDX" or sbom.get("specVersion") != "1.6":
            failures.append("SBOM is not CycloneDX 1.6 JSON")
        components = sbom.get("components")
        if not isinstance(components, list) or not components:
            failures.append("SBOM has no NuGet components")
        elif any(not item.get("purl", "").startswith("pkg:nuget/") for item in components):
            failures.append("SBOM contains a dependency without a NuGet purl")
        return failures, sbom, first_hash


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
        sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser()
    parser.add_argument("--write-evidence", type=Path)
    args = parser.parse_args()

    props = ET.parse(PROPS).getroot()
    values = {child.tag: (child.text or "").strip() for group in props.findall("PropertyGroup") for child in group}
    failures: list[str] = []
    if values.get("RestorePackagesWithLockFile") != "true":
        failures.append("RestorePackagesWithLockFile must be true")
    if values.get("RestoreLockedMode") != "true":
        failures.append("RestoreLockedMode must be true for CI")
    props_text = PROPS.read_text(encoding="utf-8-sig")
    if "RestoreLockedMode Condition=\"'$(CI)' == 'true'\"" not in props_text:
        failures.append("RestoreLockedMode must be scoped to CI=true")

    if not CI_WORKFLOW.is_file():
        failures.append("managed source-quality GitHub Actions workflow is missing")
    else:
        workflow = CI_WORKFLOW.read_text(encoding="utf-8-sig")
        required_workflow_fragments = (
            "permissions:\n  contents: read",
            "fetch-depth: 0",
            "persist-credentials: false",
            "global-json-file: global.json",
            "run_source_quality_gates.ps1",
            "if: always()",
            "artifacts/sbom/chuanhoa-dotnet.cdx.json",
        )
        for fragment in required_workflow_fragments:
            if fragment not in workflow:
                failures.append(f"source-quality workflow is missing: {fragment!r}")
        action_references = re.findall(r"uses:\s*([^\s#]+)", workflow)
        for reference in action_references:
            if not re.fullmatch(r"[^/@]+/[^/@]+@[0-9a-f]{40}", reference):
                failures.append(f"GitHub Action is not pinned by full commit SHA: {reference}")
    if not DEPENDABOT.is_file() or "package-ecosystem: github-actions" not in DEPENDABOT.read_text(
        encoding="utf-8-sig"
    ):
        failures.append("Dependabot must monitor GitHub Actions dependencies")

    lock_failures, lock_records = validate_locks()
    sbom_failures, sbom, sbom_hash = generate_twice()
    failures.extend(lock_failures)
    failures.extend(sbom_failures)
    sbom_properties = {
        item.get("name"): item.get("value")
        for item in sbom.get("metadata", {}).get("properties", [])
        if isinstance(item, dict)
    }

    evidence = {
        "schemaVersion": 2,
        "testId": "DOTNET-SUPPLY-CHAIN-001",
        "status": "PASS" if not failures else "FAIL",
        "verifiedOn": date.today().isoformat(),
        "lockedMode": "CI=true",
        "managedCiWorkflow": CI_WORKFLOW.relative_to(ROOT).as_posix(),
        "managedCiConfigured": CI_WORKFLOW.is_file() and not any(
            "workflow" in failure or "GitHub Action" in failure for failure in failures
        ),
        "packageProjects": len(package_projects()),
        "lockFiles": lock_records,
        "sbom": {
            "format": sbom.get("bomFormat"),
            "specVersion": sbom.get("specVersion"),
            "componentCount": len(sbom.get("components", [])),
            "lockFileCount": int(sbom_properties.get("chuanhoa:lockFileCount", "0")),
            "sha256": sbom_hash,
            "deterministic": not sbom_failures,
        },
        "limitations": [
            "This gate covers NuGet dependency locking and a verifiable CycloneDX SBOM.",
            "VSTO framework/Office references are not NuGet packages and are documented as external prerequisites.",
            "The repository configures a managed CI workflow, but a local run cannot prove a remote GitHub Actions execution or branch-protection policy.",
            "Production signing, key separation and provider-side secret revocation remain separate release gates.",
        ],
        "failures": failures,
    }
    if args.write_evidence:
        path = args.write_evidence if args.write_evidence.is_absolute() else ROOT / args.write_evidence
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(evidence, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    if failures:
        print("DOTNET_SUPPLY_CHAIN: FAIL")
        for failure in failures:
            print("- " + failure)
        return 1
    print(
        "DOTNET_SUPPLY_CHAIN: PASS "
        f"PACKAGE_PROJECTS={len(package_projects())} "
        f"LOCKS={len(lock_records)} COMPONENTS={len(sbom['components'])} "
        f"SBOM_SHA256={sbom_hash}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
