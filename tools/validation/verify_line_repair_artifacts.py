"""Compare DOCX regression artifacts without printing customer document text."""
import argparse
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

NS = {"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

def inspect(path):
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
        lines = []
        for anchor in root.findall(".//wp:anchor", NS):
            identity = anchor.find("wp:docPr", NS)
            # Legacy connectors can use custom geometry instead of prstGeom=line.
            # Include all anchored names so removal checks cannot miss that form.
            lines.append(identity.get("name", "") if identity is not None else "")
        media = sorted(hashlib.sha256(archive.read(name)).hexdigest()
                       for name in archive.namelist() if name.startswith("word/media/"))
        return lines, media

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("result", type=Path)
    parser.add_argument("--expected-owned", type=int, required=True)
    parser.add_argument("--removed", action="append", default=[])
    args = parser.parse_args()
    original, before_media = inspect(args.source)
    lines, after_media = inspect(args.result)
    owned = [name for name in lines if name.startswith(("CHUANHOA2_", "CHUANHOA_"))]
    assert len(owned) == args.expected_owned, "Generated line count mismatch"
    assert len(owned) == len(set(owned)), "Duplicate generated line name"
    assert before_media == after_media, "Embedded image/media payload changed"
    for name in args.removed:
        assert name in original, "Requested legacy regression line absent from source"
        assert name not in lines, "Legacy separator survived repair"
    print(json.dumps({"status": "PASS", "generatedLines": len(owned),
                      "preservedMedia": len(before_media), "removedLegacySeparators": len(args.removed)}))

if __name__ == "__main__":
    main()
