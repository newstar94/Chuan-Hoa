from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
FORBIDDEN_SUFFIXES = {".pfx", ".p12", ".snk", ".key", ".pem"}
REQUIRED_IGNORE_ENTRIES = {
    ".dev-secrets/",
    "*.pfx",
    "*.p12",
    "*.snk",
    "*.key",
    "*.pem",
}
MAX_TEXT_BYTES = 5 * 1024 * 1024


@dataclass(frozen=True)
class Finding:
    rule: str
    path: str


def candidate_paths() -> list[Path]:
    process = subprocess.run(
        ["git", "ls-files", "--cached", "--others", "--exclude-standard", "-z"],
        cwd=ROOT,
        capture_output=True,
        check=True,
    )
    return [Path(item.decode("utf-8")) for item in process.stdout.split(b"\0") if item]


def text_files(paths: list[Path]) -> tuple[dict[Path, str], int]:
    result: dict[Path, str] = {}
    skipped = 0
    for relative in paths:
        full_path = ROOT / relative
        if not full_path.is_file() or full_path.stat().st_size > MAX_TEXT_BYTES:
            skipped += 1
            continue
        data = full_path.read_bytes()
        if b"\0" in data[:8192]:
            skipped += 1
            continue
        try:
            result[relative] = data.decode("utf-8-sig")
        except UnicodeDecodeError:
            skipped += 1
    return result, skipped


def history_text_blobs() -> tuple[dict[Path, str], int, list[str], int]:
    inventory = subprocess.run(
        ["git", "rev-list", "--objects", "--all"],
        cwd=ROOT,
        capture_output=True,
        check=True,
        text=True,
        encoding="utf-8",
    ).stdout.splitlines()
    objects: list[tuple[str, str]] = []
    private_paths: list[str] = []
    for line in inventory:
        oid, separator, path = line.partition(" ")
        if not separator or not path:
            continue
        objects.append((oid, path))
        if Path(path).suffix.lower() in FORBIDDEN_SUFFIXES:
            private_paths.append(f"{path}@{oid[:12]}")

    sources: dict[Path, str] = {}
    skipped = 0
    blob_count = 0
    process = subprocess.Popen(
        ["git", "cat-file", "--batch"],
        cwd=ROOT,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
    )
    assert process.stdin is not None and process.stdout is not None
    try:
        for oid, path in objects:
            process.stdin.write((oid + "\n").encode("ascii"))
            process.stdin.flush()
            header = process.stdout.readline().decode("ascii", errors="replace").strip()
            fields = header.split()
            if len(fields) != 3 or fields[1] != "blob":
                if len(fields) == 3 and fields[2].isdigit():
                    process.stdout.read(int(fields[2]) + 1)
                continue
            size = int(fields[2])
            blob_count += 1
            data = process.stdout.read(size)
            process.stdout.read(1)
            if size > MAX_TEXT_BYTES or b"\0" in data[:8192]:
                skipped += 1
                continue
            try:
                content = data.decode("utf-8-sig")
            except UnicodeDecodeError:
                skipped += 1
                continue
            sources[Path("history") / oid[:12] / Path(path)] = content
    finally:
        process.stdin.close()
        process.stdout.close()
        process.wait(timeout=10)
    if process.returncode != 0:
        raise subprocess.CalledProcessError(process.returncode, process.args)
    return sources, skipped, sorted(set(private_paths)), blob_count


def scan(sources: dict[Path, str]) -> list[Finding]:
    findings: list[Finding] = []
    patterns = {
        "PEM_PRIVATE_KEY": re.compile(
            r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----"
            r"[A-Za-z0-9+/=\r\n]{40,}"
            r"-----END (?:RSA |EC |OPENSSH )?PRIVATE KEY-----"
        ),
        "GITHUB_TOKEN": re.compile(r"gh[pousr]_[A-Za-z0-9]{36,}"),
        "AWS_ACCESS_KEY": re.compile(r"AKIA[0-9A-Z]{16}"),
        "SLACK_TOKEN": re.compile(r"xox[baprs]-[A-Za-z0-9-]{20,}"),
        "JWT": re.compile(r"eyJ[A-Za-z0-9_-]{12,}\.[A-Za-z0-9_-]{12,}\.[A-Za-z0-9_-]{12,}"),
        "PASSWORD_ASSIGNMENT": re.compile(
            r"(?i)\b(?:password|passwd|pwd)\s*=\s*"
            r"(?!\s*(?:\$\(|\$\{|%|<|REDACTED|CHANGEME|PLACEHOLDER|EXAMPLE))"
            r"[^;\s\"']{6,}"
        ),
        "JSON_SECRET": re.compile(
            r'(?i)"(?:clientSecret|apiKey|secretKey|accessToken)"\s*:\s*'
            r'"(?!\s*(?:|REDACTED|CHANGEME|PLACEHOLDER|EXAMPLE)\s*")[^"\r\n]{8,}"'
        ),
    }
    for path, content in sources.items():
        if (
            "<RSAKeyValue" in content
            and re.search(r"<D>[A-Za-z0-9+/=]{40,}</D>", content)
            and re.search(r"<P>[A-Za-z0-9+/=]{20,}</P>", content)
        ):
            findings.append(Finding("RSA_XML_PRIVATE_KEY", path.as_posix()))
        for name, pattern in patterns.items():
            if pattern.search(content):
                findings.append(Finding(name, path.as_posix()))
    return findings


def negative_self_tests() -> dict[str, bool]:
    fixtures = {
        "pem": "-----BEGIN PRIVATE KEY-----\n" + "A" * 80 + "\n-----END PRIVATE KEY-----",
        "rsa_xml": "<RSAKeyValue><P>" + "A" * 24 + "</P><D>" + "B" * 48 + "</D></RSAKeyValue>",
        "github": "token = ghp_" + "A" * 40,
        "password": "Password=" + "not-a-real-secret-but-detect-me",
        "json_secret": '{"client' + 'Secret":"' + "not-a-real-secret-but-detect-me" + '"}',
    }
    return {
        name: bool(scan({Path(f"negative/{name}.txt"): content}))
        for name, content in fixtures.items()
    }


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
        sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser()
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--history", action="store_true")
    parser.add_argument("--write-evidence", type=Path)
    args = parser.parse_args()

    failures: list[str] = []
    try:
        paths = candidate_paths()
    except (OSError, subprocess.CalledProcessError) as error:
        print(f"REPOSITORY_SECRETS: FAIL - cannot enumerate repository files: {error}")
        return 1
    tracked_private_files = [path.as_posix() for path in paths if path.suffix.lower() in FORBIDDEN_SUFFIXES]
    failures.extend(f"private-key file is tracked or pending: {path}" for path in tracked_private_files)

    gitignore = (ROOT / ".gitignore").read_text(encoding="utf-8-sig").splitlines()
    ignore_entries = {line.strip() for line in gitignore if line.strip() and not line.startswith("#")}
    failures.extend(
        f"required secret ignore rule is missing: {entry}"
        for entry in sorted(REQUIRED_IGNORE_ENTRIES - ignore_entries)
    )
    indexed_dev_secrets = subprocess.run(
        ["git", "ls-files", "--", ".dev-secrets"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
    ).stdout.strip()
    if indexed_dev_secrets:
        failures.append(".dev-secrets contains indexed files")

    sources, skipped = text_files(paths)
    findings = scan(sources)
    failures.extend(f"{finding.rule}: {finding.path}" for finding in findings)
    history_sources: dict[Path, str] = {}
    history_skipped = 0
    history_private_paths: list[str] = []
    history_blob_count = 0
    if args.history:
        try:
            history_sources, history_skipped, history_private_paths, history_blob_count = (
                history_text_blobs()
            )
        except (OSError, subprocess.CalledProcessError) as error:
            failures.append(f"cannot scan Git history: {error}")
        history_findings = scan(history_sources)
        failures.extend(
            f"HISTORY_{finding.rule}: {finding.path}" for finding in history_findings
        )
        failures.extend(
            f"private-key file exists in Git history: {path}"
            for path in history_private_paths
        )
    else:
        history_findings = []
    self_tests = negative_self_tests() if args.self_test else {}
    failures.extend(f"negative self-test did not detect {name}" for name, passed in self_tests.items() if not passed)

    evidence = {
        "schemaVersion": 2,
        "testId": "REPOSITORY-SECRET-REGRESSION-001",
        "status": "PASS" if not failures else "FAIL",
        "verifiedOn": date.today().isoformat(),
        "candidateFileCount": len(paths),
        "textFileCount": len(sources),
        "binaryOrLargeSkippedCount": skipped,
        "findingCount": len(findings),
        "trackedPrivateKeyFileCount": len(tracked_private_files),
        "historyScanEnabled": args.history,
        "historyBlobCount": history_blob_count,
        "historyTextBlobCount": len(history_sources),
        "historyBinaryOrLargeSkippedCount": history_skipped,
        "historyFindingCount": len(history_findings),
        "historyPrivateKeyPathCount": len(history_private_paths),
        "negativeSelfTests": self_tests,
        "limitations": [
            "Git history is scanned only when --history is supplied.",
            "Pattern scanning does not replace secret rotation or provider-side token revocation.",
            "Ignored local development secrets are checked for index exclusion but their contents are intentionally not emitted or scanned into evidence.",
        ],
        "failures": failures,
    }
    if args.write_evidence:
        path = args.write_evidence if args.write_evidence.is_absolute() else ROOT / args.write_evidence
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(evidence, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    if failures:
        print("REPOSITORY_SECRETS: FAIL")
        for failure in failures:
            print("- " + failure)
        return 1
    print(
        "REPOSITORY_SECRETS: PASS "
        f"TEXT_FILES={len(sources)} SKIPPED={skipped} "
        f"HISTORY_BLOBS={history_blob_count if args.history else 'NOT_RUN'} "
        f"NEGATIVE_TESTS={sum(self_tests.values())}/{len(self_tests)}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
