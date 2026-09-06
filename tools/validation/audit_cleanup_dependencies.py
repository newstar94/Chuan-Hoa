"""Read-only inventory for cleanup; never decides that an unreferenced file is safe to delete."""
from pathlib import Path
import json
import subprocess

ROOT = Path(__file__).resolve().parents[2]
CANDIDATES = ("backend-api", "client-vsto-csharp", "client-web-addin", "shared/vba_extracted",
              "shared/ChuanHoaTheThuc_Full_Ribbon.dotm", "tmp")
TEXT_SUFFIXES = {".py", ".ps1", ".yml", ".yaml", ".json", ".csproj", ".slnx", ".md", ".cs", ".ts", ".js"}

def audit():
    tracked = subprocess.check_output(["git", "ls-files", "-z"], cwd=ROOT).decode("utf-8").split("\0")
    result = []
    for candidate in CANDIDATES:
        refs = []
        needles = {candidate.lower(), candidate.replace("/", "\\").lower()}
        for relative in tracked:
            if not relative or relative == candidate or relative.startswith(candidate + "/"):
                continue
            path = ROOT / relative
            if path.suffix.lower() not in TEXT_SUFFIXES or not path.is_file():
                continue
            for number, line in enumerate(path.read_text(encoding="utf-8", errors="replace").splitlines(), 1):
                if any(needle in line.lower() for needle in needles):
                    # Locations only: do not copy document text or credentials to evidence.
                    refs.append({"file": relative, "line": number})
        target = ROOT / candidate
        files = list(target.rglob("*")) if target.is_dir() else [target]
        result.append({"candidate": candidate, "exists": target.exists(),
                       "files": sum(p.is_file() for p in files), "references": refs,
                       "decision": "RETAIN_PENDING_DEPENDENCY_MIGRATION" if refs else "REVIEW_NOT_AUTO_DELETE"})
    return {"mode": "READ_ONLY", "candidates": result}

if __name__ == "__main__":
    print(json.dumps(audit(), ensure_ascii=True, indent=2))
