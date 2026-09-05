from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from uuid import UUID
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[2]
SOLUTION = ROOT / "ChuanHoa.slnx"
PROPS = ROOT / "Directory.Build.props"
DEFAULT_OUTPUT = ROOT / "artifacts" / "sbom" / "chuanhoa-dotnet.cdx.json"
LOCK_EXCLUDED_PROJECTS = {
    Path("src/ChuanHoa.AddIn.Vsto/ChuanHoa.AddIn.Vsto.csproj"),
}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().lower()


def product_version() -> str:
    root = ET.parse(PROPS).getroot()
    value = root.findtext("./PropertyGroup/ProductVersion")
    if not value or not re.fullmatch(r"\d+\.\d+\.\d+\.\d+", value):
        raise RuntimeError("Directory.Build.props has no valid ProductVersion")
    return value


def solution_projects() -> list[Path]:
    result: set[Path] = set()
    root = ET.parse(SOLUTION).getroot()
    for element in root.iter("Project"):
        relative = Path(element.attrib["Path"])
        if relative.suffix.lower() == ".csproj":
            result.add(relative)
    for source_root in (ROOT / "src", ROOT / "tests"):
        for project_path in source_root.rglob("*.csproj"):
            if "<PackageReference" in project_path.read_text(
                encoding="utf-8-sig", errors="strict"
            ):
                result.add(project_path.relative_to(ROOT))
    return sorted(result, key=lambda path: path.as_posix().lower())


def package_references(project: Path) -> dict[str, str]:
    root = ET.parse(ROOT / project).getroot()
    result: dict[str, str] = {}
    for element in root.iter("PackageReference"):
        package_id = element.attrib.get("Include") or element.attrib.get("Update")
        version = element.attrib.get("Version") or element.findtext("Version")
        if package_id and version and not version.startswith("$("):
            result[package_id.lower()] = version
    return result


def framework_dependencies(project: Path, lock: dict) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {}
    dependencies = lock.get("dependencies", {})
    if not isinstance(dependencies, dict) or not dependencies:
        raise RuntimeError(f"lock file has no dependency frameworks: {project.parent / 'packages.lock.json'}")
    for framework, entries in dependencies.items():
        if not isinstance(entries, dict):
            raise RuntimeError(f"invalid dependency map for {project}: {framework}")
        direct_refs = package_references(project)
        for package_id, version in direct_refs.items():
            match = next((item for name, item in entries.items() if name.lower() == package_id), None)
            if not isinstance(match, dict):
                raise RuntimeError(f"direct package {package_id} missing from lock for {project}")
            resolved = str(match.get("resolved", ""))
            if not resolved:
                raise RuntimeError(f"direct package {package_id} has no resolved version for {project}")
            if not str(match.get("type", "")).lower().startswith("direct"):
                raise RuntimeError(f"package {package_id} is not marked Direct for {project}")
            result.setdefault(package_id, {})[framework] = resolved
        for package_id, metadata in entries.items():
            if not isinstance(metadata, dict):
                continue
            resolved = str(metadata.get("resolved", ""))
            if resolved:
                result.setdefault(package_id.lower(), {})[framework] = resolved
    return result


def validate_locks(projects: list[Path]) -> tuple[dict[str, set[str]], list[dict[str, str]]]:
    components: dict[str, set[str]] = {}
    lock_records: list[dict[str, str]] = []
    for project in projects:
        if project in LOCK_EXCLUDED_PROJECTS:
            continue
        lock_path = ROOT / project.parent / "packages.lock.json"
        if not lock_path.is_file():
            raise RuntimeError(f"missing packages.lock.json: {lock_path.relative_to(ROOT)}")
        lock = json.loads(lock_path.read_text(encoding="utf-8-sig"))
        if lock.get("version") != 1:
            raise RuntimeError(f"unsupported lock version: {lock_path.relative_to(ROOT)}")
        for package_id, framework_versions in framework_dependencies(project, lock).items():
            components.setdefault(package_id, set()).update(framework_versions.values())
        lock_records.append(
            {
                "project": project.as_posix(),
                "lockFile": lock_path.relative_to(ROOT).as_posix(),
                "sha256": sha256(lock_path),
            }
        )
    return components, lock_records


def build_sbom(output: Path) -> dict:
    projects = solution_projects()
    components, lock_records = validate_locks(projects)
    lock_identity = "\n".join(
        f"{record['lockFile']}:{record['sha256']}" for record in lock_records
    )
    serial = UUID(bytes=hashlib.sha256(lock_identity.encode("utf-8")).digest()[:16])
    component_list = []
    for package_id, versions in sorted(components.items()):
        for version in sorted(versions):
            component_list.append(
                {
                    "type": "library",
                    "bom-ref": f"pkg:nuget/{package_id}@{version}",
                    "name": package_id,
                    "version": version,
                    "purl": f"pkg:nuget/{package_id}@{version}",
                    "scope": "required",
                }
            )
    return {
        "bomFormat": "CycloneDX",
        "specVersion": "1.6",
        "serialNumber": f"urn:uuid:{serial}",
        "version": 1,
        "metadata": {
            "tools": {
                "components": [
                    {
                        "type": "application",
                        "name": "ChuanHoa deterministic lockfile SBOM generator",
                        "version": "1",
                    }
                ]
            },
            "component": {
                "type": "application",
                "bom-ref": f"pkg:generic/chuanhoa@{product_version()}",
                "name": "ChuanHoa",
                "version": product_version(),
            },
            "properties": [
                {"name": "chuanhoa:solution", "value": "ChuanHoa.slnx"},
                {"name": "chuanhoa:lockFileCount", "value": str(len(lock_records))},
                {
                    "name": "chuanhoa:excludedLegacyProject",
                    "value": "src/ChuanHoa.AddIn.Vsto/ChuanHoa.AddIn.Vsto.csproj uses only framework/Office/project references",
                },
                {
                    "name": "chuanhoa:lockEvidence",
                    "value": json.dumps(lock_records, separators=(",", ":"), sort_keys=True),
                },
            ],
        },
        "components": component_list,
    }


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
        sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    output = args.output if args.output.is_absolute() else ROOT / args.output
    try:
        sbom = build_sbom(output)
    except Exception as error:
        print(f"DOTNET_SBOM: FAIL - {error}")
        return 1
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(sbom, indent=2) + "\n", encoding="utf-8")
    try:
        output_label = output.relative_to(ROOT).as_posix()
    except ValueError:
        output_label = output.as_posix()
    print(
        "DOTNET_SBOM: PASS "
        f"COMPONENTS={len(sbom['components'])} "
        f"OUTPUT={output_label} "
        f"SHA256={sha256(output).upper()}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
