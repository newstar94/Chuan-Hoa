from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
SOURCE_ROOTS = (ROOT / "src", ROOT / "tests")
TEXT_SUFFIXES = {".cs", ".csproj", ".props", ".targets", ".json", ".xml"}
FORBIDDEN_TEXT = (
    "VietnameseEngine",
    "LOCAL-TYPO-AI",
    "onnxruntime",
    "model_manifest",
    "InferenceSession",
    "CheckWithAiEngine",
    "btnChenQrCode",
    "InsertQrCode",
    "QrCodeInputDialog",
    "QRCoder",
)
FORBIDDEN_SUFFIXES = (".onnx", ".onnx.data")


def main() -> int:
    failures: list[str] = []
    for forbidden_directory in (ROOT / "tools" / "VietnameseEngine", ROOT / "training"):
        if forbidden_directory.exists():
            failures.append(f"forbidden directory exists: {forbidden_directory.relative_to(ROOT)}")

    for source_root in SOURCE_ROOTS:
        for path in source_root.rglob("*"):
            if not path.is_file() or any(part in {"bin", "obj"} for part in path.parts):
                continue
            relative = path.relative_to(ROOT)
            lower_name = path.name.lower()
            if lower_name.endswith(FORBIDDEN_SUFFIXES):
                failures.append(f"forbidden model artifact: {relative}")
                continue
            if path.suffix.lower() not in TEXT_SUFFIXES:
                continue
            content = path.read_text(encoding="utf-8", errors="replace")
            for marker in FORBIDDEN_TEXT:
                if marker.lower() in content.lower():
                    failures.append(f"forbidden runtime marker {marker!r}: {relative}")

    installer = (ROOT / "tools" / "vsto" / "build_development_test_exe.ps1").read_text(
        encoding="utf-8"
    )
    if "Copy-Item -Path" in installer or "artifacts\\models" in installer:
        failures.append("Development installer contains a wildcard/model payload copy")
    required_allowlist = (
        "ChuanHoa.AddIn.Vsto.vsto",
        "ChuanHoa.AddIn.Vsto.dll.manifest",
        "ChuanHoa.AddIn.Vsto.dll",
        "ChuanHoa.Client.Core.dll",
    )
    for payload_name in required_allowlist:
        if payload_name not in installer:
            failures.append(f"Development installer allowlist is missing {payload_name}")
    for marker in ("QRCoder", "btnChenQrCode", "QrCodeInputDialog", "InsertQrCode"):
        if marker.lower() not in installer.lower():
            failures.append(f"Development installer deny-list is missing retired QR marker {marker!r}")

    if failures:
        print("RULE_ONLY_PRODUCT: FAIL")
        for failure in failures:
            print("- " + failure)
        return 1
    print("RULE_ONLY_PRODUCT: PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
