from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
VSTO_ROOT = ROOT / "src" / "ChuanHoa.AddIn.Vsto"
NETWORK_ALLOWLIST = {Path("Runtime/LocalAccessManager.cs")}
ENDPOINT_ALLOWLIST = {"/v1/development/bootstrap"}
REQUEST_FIELD_ALLOWLIST = {"deviceThumbprint", "clientReleaseId"}

NETWORK_MARKERS = (
    "using System.Net",
    "HttpClient",
    "HttpRequestMessage",
    "WebRequest",
    "WebClient",
    "TcpClient",
    "UdpClient",
    "Socket(",
)
UPLOAD_MARKERS = (
    "MultipartFormDataContent",
    "StreamContent",
    "ByteArrayContent",
    "File.ReadAllBytes",
    "File.OpenRead",
    "UploadFile",
    "UploadData",
    "OpenWrite",
)
DOCUMENT_DATA_MARKERS = (
    "ActiveDocument",
    "DocumentContext",
    "DocumentSnapshot",
    "Word.Document",
    "Microsoft.Office.Interop.Word",
    "Range.Text",
    "StoryRanges",
    "document.FullName",
    "document.Path",
    "document.Name",
    "FileName",
    "Filename",
)
REMOTE_TELEMETRY_MARKERS = (
    "ApplicationInsights",
    "TelemetryClient",
    "OpenTelemetry",
    "SentrySdk",
    "TrackEvent(",
    "TrackException(",
)


@dataclass(frozen=True)
class PrivacyResult:
    failures: tuple[str, ...]
    source_files: int
    network_files: tuple[str, ...]
    endpoints: tuple[str, ...]
    request_fields: tuple[str, ...]


def source_map() -> dict[Path, str]:
    result: dict[Path, str] = {}
    for path in VSTO_ROOT.rglob("*"):
        if (
            path.is_file()
            and path.suffix.lower() in {".cs", ".csproj", ".config", ".xml"}
            and not any(part in {"bin", "obj"} for part in path.parts)
        ):
            result[path.relative_to(VSTO_ROOT)] = path.read_text(
                encoding="utf-8-sig", errors="strict"
            )
    return result


def analyze(sources: dict[Path, str]) -> PrivacyResult:
    failures: list[str] = []
    network_files = {
        path
        for path, content in sources.items()
        if any(marker in content for marker in NETWORK_MARKERS)
    }
    unexpected_network = network_files - NETWORK_ALLOWLIST
    missing_network = NETWORK_ALLOWLIST - network_files
    for path in sorted(unexpected_network):
        failures.append(f"network API outside allowlist: {path.as_posix()}")
    for path in sorted(missing_network):
        failures.append(f"expected reviewed network seam is missing: {path.as_posix()}")

    for path, content in sorted(sources.items(), key=lambda item: item[0].as_posix()):
        for marker in REMOTE_TELEMETRY_MARKERS:
            if marker.lower() in content.lower():
                failures.append(
                    f"remote telemetry requires privacy review ({marker}): {path.as_posix()}"
                )

    access_source = sources.get(Path("Runtime/LocalAccessManager.cs"), "")
    endpoints = set(re.findall(r'"(/v1/[A-Za-z0-9_{}./-]+)"', access_source))
    if endpoints != ENDPOINT_ALLOWLIST:
        failures.append(
            "reviewed VSTO endpoint set changed: "
            f"expected={sorted(ENDPOINT_ALLOWLIST)} actual={sorted(endpoints)}"
        )

    refresh_match = re.search(
        r"private void Refresh\(\)\s*\{(?P<body>.*?)\n\s*private RsaSha256ArtifactVerifier",
        access_source,
        flags=re.DOTALL,
    )
    refresh_source = refresh_match.group("body") if refresh_match else ""
    if not refresh_match:
        failures.append("could not isolate LocalAccessManager.Refresh for payload review")
    access_outside_refresh = (
        access_source[: refresh_match.start("body")]
        + access_source[refresh_match.end("body") :]
        if refresh_match
        else access_source
    )
    for marker in (
        "HttpRequestMessage",
        "StringContent(",
        ".SendAsync(",
        ".PostAsync(",
        ".PutAsync(",
        ".GetAsync(",
    ):
        if marker in access_outside_refresh:
            failures.append(f"network transmission outside reviewed Refresh method: {marker}")

    body_match = re.search(
        r"var body\s*=\s*(?P<body>.*?);\s*\n\s*using\s*\(var content\s*=\s*new StringContent",
        refresh_source,
        flags=re.DOTALL,
    )
    body_source = body_match.group("body") if body_match else ""
    request_fields = set(re.findall(r'\\"([A-Za-z][A-Za-z0-9]*)\\"\s*:', body_source))
    if request_fields != REQUEST_FIELD_ALLOWLIST:
        failures.append(
            "reviewed request payload changed: "
            f"expected={sorted(REQUEST_FIELD_ALLOWLIST)} actual={sorted(request_fields)}"
        )
    if refresh_source.count("new StringContent(body, Encoding.UTF8, \"application/json\")") != 1:
        failures.append("bootstrap request must have exactly one reviewed JSON body")
    if refresh_source.count("request.Content = content") != 1:
        failures.append("bootstrap request content assignment changed")

    for marker in UPLOAD_MARKERS + DOCUMENT_DATA_MARKERS:
        if marker.lower() in access_source.lower():
            failures.append(f"document/upload marker in network seam: {marker}")

    all_vsto = "\n".join(sources.values())
    for marker in ("/document-scan", "DocumentScanController", "DocumentScanRequest"):
        if marker.lower() in all_vsto.lower():
            failures.append(f"legacy server scan is referenced by VSTO: {marker}")

    project = sources.get(Path("ChuanHoa.AddIn.Vsto.csproj"), "")
    for forbidden_reference in ("ChuanHoa.Api", "ChuanHoa.Infrastructure"):
        if forbidden_reference.lower() in project.lower():
            failures.append(f"VSTO references server project: {forbidden_reference}")

    return PrivacyResult(
        failures=tuple(dict.fromkeys(failures)),
        source_files=len(sources),
        network_files=tuple(sorted(path.as_posix() for path in network_files)),
        endpoints=tuple(sorted(endpoints)),
        request_fields=tuple(sorted(request_fields)),
    )


def run_negative_self_tests(sources: dict[Path, str]) -> dict[str, bool]:
    access_path = Path("Runtime/LocalAccessManager.cs")
    cases: dict[str, dict[Path, str]] = {}

    extra_network = dict(sources)
    extra_network[Path("ThisAddIn.cs")] += "\nusing System.Net.Http;\n"
    cases["network_outside_allowlist"] = extra_network

    document_payload = dict(sources)
    document_payload[access_path] = document_payload[access_path].replace(
        '\\",\\"clientReleaseId\\":\\"',
        '\\",\\"documentText\\":\\"x\\",\\"clientReleaseId\\":\\"',
        1,
    )
    cases["document_payload_field"] = document_payload

    scan_endpoint = dict(sources)
    scan_endpoint[access_path] = scan_endpoint[access_path].replace(
        "/v1/development/bootstrap", "/v1/document-scan", 1
    )
    cases["legacy_scan_endpoint"] = scan_endpoint

    upload_content = dict(sources)
    upload_content[access_path] = upload_content[access_path].replace(
        "var body =", "var forbidden = new StreamContent(null);\n            var body =", 1
    )
    cases["binary_upload_primitive"] = upload_content

    return {name: bool(analyze(mutated).failures) for name, mutated in cases.items()}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def write_evidence(path: Path, result: PrivacyResult, self_tests: dict[str, bool]) -> None:
    access_path = VSTO_ROOT / "Runtime" / "LocalAccessManager.cs"
    evidence = {
        "schemaVersion": 1,
        "testId": "DOCUMENT-PRIVACY-REGRESSION-001",
        "status": "PASS" if not result.failures and all(self_tests.values()) else "FAIL",
        "verifiedOn": date.today().isoformat(),
        "scope": "src/ChuanHoa.AddIn.Vsto source; no Word document was opened or uploaded",
        "vstoSourceFilesReviewed": result.source_files,
        "networkSourceAllowlist": list(result.network_files),
        "allowedEndpoints": list(result.endpoints),
        "allowedRequestFields": list(result.request_fields),
        "localAccessManagerSha256": sha256(access_path),
        "negativeSelfTests": self_tests,
        "assertions": [
            "VSTO network APIs exist only in Runtime/LocalAccessManager.cs.",
            "The Development bootstrap sends only deviceThumbprint and clientReleaseId.",
            "VSTO does not reference the legacy document-scan endpoint or server projects.",
            "The reviewed network seam contains no document text/path/name or upload primitive.",
            "No remote telemetry SDK/sink is present in VSTO source.",
        ],
        "limitations": [
            "This is a source regression gate, not production network observability or a penetration test.",
            "Any production identity, lease, rule or telemetry endpoint must update the explicit allowlist and receive privacy review.",
        ],
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(evidence, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
        sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser()
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--write-evidence", type=Path)
    args = parser.parse_args()

    sources = source_map()
    result = analyze(sources)
    self_tests = run_negative_self_tests(sources) if args.self_test else {}
    failures = list(result.failures)
    if args.self_test:
        failures.extend(name for name, passed in self_tests.items() if not passed)
    if args.write_evidence:
        evidence_path = args.write_evidence
        if not evidence_path.is_absolute():
            evidence_path = ROOT / evidence_path
        write_evidence(evidence_path, result, self_tests)

    if failures:
        print("DOCUMENT_PRIVACY: FAIL")
        for failure in failures:
            print("- " + failure)
        return 1
    print(
        "DOCUMENT_PRIVACY: PASS "
        f"NETWORK_FILES={len(result.network_files)} "
        f"ENDPOINTS={len(result.endpoints)} "
        f"REQUEST_FIELDS={len(result.request_fields)} "
        f"NEGATIVE_TESTS={sum(self_tests.values())}/{len(self_tests)}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
