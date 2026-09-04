from __future__ import annotations

import sys
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[2]
SOLUTION = ROOT / "ChuanHoa.slnx"


def main() -> int:
    tree = ET.parse(SOLUTION)
    missing: list[str] = []
    for project in tree.getroot().iter("Project"):
        relative = project.attrib.get("Path", "")
        if not relative or not (ROOT / relative).is_file():
            missing.append(relative or "<missing Path attribute>")
    if missing:
        print("SOLUTION_PROJECTS: FAIL")
        for path in missing:
            print("- " + path)
        return 1
    print("SOLUTION_PROJECTS: PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
