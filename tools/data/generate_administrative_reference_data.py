#!/usr/bin/env python3
"""Generate deterministic Vietnamese administrative reference data.

The checked-in source snapshot is extracted from the two official DOC attachments
of Decision 19/2025/QD-TTg.  Runtime assets are generated only from that snapshot;
the original office files are not required by the product or CI.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import unicodedata
from pathlib import Path
from typing import Any, Iterable


ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / "tools/data/sources/qd19_2025_administrative_units.json"
ADMIN = ROOT / "shared/dictionaries/administrative_units.json"
HISTORICAL = ROOT / "shared/dictionaries/historical_administrative_aliases.json"
SPECIAL = ROOT / "shared/dictionaries/special_capitalizations.json"
MANIFEST = ROOT / "shared/dictionaries/reference_data_manifest.json"

SOURCE_SYMBOL = "Quyet dinh 19/2025/QD-TTg"
SOURCE_URL = "https://congbao.chinhphu.vn/van-ban/quyet-dinh-so-19-2025-qd-ttg-45430/57438.htm"
EFFECTIVE_FROM = "2025-07-01"

HISTORICAL_PROVINCES = {
    "Tỉnh Hà Giang": ("08", "Tỉnh Tuyên Quang"),
    "Tỉnh Yên Bái": ("15", "Tỉnh Lào Cai"),
    "Tỉnh Bắc Kạn": ("19", "Tỉnh Thái Nguyên"),
    "Tỉnh Vĩnh Phúc": ("25", "Tỉnh Phú Thọ"),
    "Tỉnh Hòa Bình": ("25", "Tỉnh Phú Thọ"),
    "Tỉnh Bắc Giang": ("24", "Tỉnh Bắc Ninh"),
    "Tỉnh Thái Bình": ("33", "Tỉnh Hưng Yên"),
    "Tỉnh Hải Dương": ("31", "Thành phố Hải Phòng"),
    "Tỉnh Hà Nam": ("37", "Tỉnh Ninh Bình"),
    "Tỉnh Nam Định": ("37", "Tỉnh Ninh Bình"),
    "Tỉnh Quảng Bình": ("44", "Tỉnh Quảng Trị"),
    "Tỉnh Quảng Nam": ("48", "Thành phố Đà Nẵng"),
    "Tỉnh Kon Tum": ("51", "Tỉnh Quảng Ngãi"),
    "Tỉnh Bình Định": ("52", "Tỉnh Gia Lai"),
    "Tỉnh Ninh Thuận": ("56", "Tỉnh Khánh Hòa"),
    "Tỉnh Phú Yên": ("66", "Tỉnh Đắk Lắk"),
    "Tỉnh Đắk Nông": ("68", "Tỉnh Lâm Đồng"),
    "Tỉnh Bình Thuận": ("68", "Tỉnh Lâm Đồng"),
    "Tỉnh Bình Phước": ("75", "Tỉnh Đồng Nai"),
    "Tỉnh Bình Dương": ("79", "Thành phố Hồ Chí Minh"),
    "Tỉnh Bà Rịa - Vũng Tàu": ("79", "Thành phố Hồ Chí Minh"),
    "Tỉnh Long An": ("80", "Tỉnh Tây Ninh"),
    "Tỉnh Tiền Giang": ("82", "Tỉnh Đồng Tháp"),
    "Tỉnh Bến Tre": ("86", "Tỉnh Vĩnh Long"),
    "Tỉnh Trà Vinh": ("86", "Tỉnh Vĩnh Long"),
    "Tỉnh Kiên Giang": ("91", "Tỉnh An Giang"),
    "Tỉnh Hậu Giang": ("92", "Thành phố Cần Thơ"),
    "Tỉnh Sóc Trăng": ("92", "Thành phố Cần Thơ"),
    "Tỉnh Bạc Liêu": ("96", "Tỉnh Cà Mau"),
}

CURATED_SPECIAL = [
    "Đảng Cộng sản Việt Nam", "Ban Chấp hành Trung ương", "Bộ Chính trị",
    "Ban Bí thư", "Ủy ban Kiểm tra Trung ương", "Văn phòng Trung ương Đảng",
    "Quốc hội", "Ủy ban Thường vụ Quốc hội", "Hội đồng Dân tộc",
    "Văn phòng Quốc hội", "Chính phủ", "Thủ tướng Chính phủ",
    "Phó Thủ tướng Chính phủ", "Văn phòng Chính phủ", "Bộ Quốc phòng",
    "Bộ Công an", "Bộ Ngoại giao", "Bộ Tư pháp", "Bộ Tài chính",
    "Bộ Công Thương", "Bộ Xây dựng", "Bộ Nội vụ", "Bộ Y tế",
    "Bộ Giáo dục và Đào tạo", "Bộ Khoa học và Công nghệ",
    "Bộ Văn hóa, Thể thao và Du lịch", "Bộ Nông nghiệp và Môi trường",
    "Bộ Dân tộc và Tôn giáo", "Thanh tra Chính phủ",
    "Ngân hàng Nhà nước Việt Nam", "Tòa án nhân dân tối cao",
    "Viện kiểm sát nhân dân tối cao", "Kiểm toán nhà nước",
    "Mặt trận Tổ quốc Việt Nam", "Tổng Liên đoàn Lao động Việt Nam",
    "Đoàn Thanh niên Cộng sản Hồ Chí Minh", "Hội Liên hiệp Phụ nữ Việt Nam",
    "Hội Nông dân Việt Nam", "Hội Cựu chiến binh Việt Nam",
    "Cộng hòa xã hội chủ nghĩa Việt Nam", "Ngày Quốc khánh",
    "Ngày Giải phóng miền Nam, thống nhất đất nước", "Ngày Quốc tế Lao động",
    "Ngày thành lập Đảng Cộng sản Việt Nam", "Chủ tịch Hồ Chí Minh",
]


def nfc(value: str) -> str:
    return unicodedata.normalize("NFC", value).strip()


def stable_write(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def iter_doc_blocks(document: Any) -> Iterable[tuple[str, Any]]:
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    for child in document.element.body.iterchildren():
        if child.tag.endswith("}p"):
            yield "paragraph", Paragraph(child, document)
        elif child.tag.endswith("}tbl"):
            yield "table", Table(child, document)


def extract_source(docx_paths: list[Path]) -> dict[str, Any]:
    from docx import Document

    provinces: list[dict[str, str]] = []
    communes: list[dict[str, str]] = []
    current_province_code: str | None = None
    province_heading = re.compile(r"^(\d{2})\.\s+(?:TỈNH|THÀNH PHỐ)\s+")

    for part_index, docx_path in enumerate(docx_paths):
        document = Document(str(docx_path))
        large_table_index = 0
        for kind, block in iter_doc_blocks(document):
            if kind == "paragraph":
                text = nfc(block.text)
                match = province_heading.match(text)
                if match:
                    current_province_code = match.group(1)
                continue

            if len(block.rows) <= 3:
                continue
            large_table_index += 1
            header = [nfc(cell.text) for cell in block.rows[0].cells]
            if part_index == 0 and large_table_index == 1 and len(header) == 3:
                for row in block.rows[1:]:
                    cells = [nfc(cell.text) for cell in row.cells]
                    if len(cells) >= 3 and re.fullmatch(r"\d{2}", cells[1]):
                        provinces.append({"code": cells[1], "name": cells[2]})
                continue

            if len(header) != 2 or current_province_code is None:
                continue
            for row in block.rows[1:]:
                cells = [nfc(cell.text) for cell in row.cells]
                if len(cells) >= 2 and re.fullmatch(r"\d{5}", cells[0]):
                    communes.append({
                        "code": cells[0], "name": cells[1],
                        "provinceCode": current_province_code,
                    })

    return {
        "metadata": {
            "source": SOURCE_SYMBOL,
            "sourceUrl": SOURCE_URL,
            "effectiveFrom": EFFECTIVE_FROM,
            "snapshotDate": "2026-09-07",
            "sourceFiles": [
                {"name": path.name, "sha256": file_sha256(path)} for path in docx_paths
            ],
        },
        "provinces": provinces,
        "communes": communes,
    }


def unit_type(name: str) -> str:
    if name.startswith("Thành phố "):
        return "CentrallyGovernedCity"
    if name.startswith("Tỉnh "):
        return "Province"
    if name.startswith("Phường "):
        return "Ward"
    if name.startswith("Xã "):
        return "Commune"
    if name.startswith("Đặc khu "):
        return "SpecialZone"
    raise ValueError(f"Unsupported administrative type: {name!r}")


def validate_source(source: dict[str, Any]) -> None:
    provinces = source["provinces"]
    communes = source["communes"]
    if len(provinces) != 34 or len(communes) != 3321:
        raise ValueError(f"Unexpected current counts: provinces={len(provinces)} communes={len(communes)}")
    province_codes = {item["code"] for item in provinces}
    if len(province_codes) != 34:
        raise ValueError("Duplicate province code")
    commune_codes = {item["code"] for item in communes}
    if len(commune_codes) != 3321:
        raise ValueError("Duplicate commune code")
    if any(item["provinceCode"] not in province_codes for item in communes):
        raise ValueError("Commune has an unknown parent province")
    if any(not nfc(item["name"]).isprintable() for item in provinces + communes):
        raise ValueError("Administrative name contains invalid characters")


def generate(source: dict[str, Any]) -> None:
    validate_source(source)
    metadata = source["metadata"]
    units: list[dict[str, Any]] = []
    for province in source["provinces"]:
        units.append({
            "code": province["code"], "name": nfc(province["name"]),
            "canonicalName": nfc(province["name"]), "level": "Province",
            "type": unit_type(province["name"]), "provinceCode": None,
            "effectiveFrom": EFFECTIVE_FROM, "effectiveTo": None,
            "status": "Active", "source": SOURCE_SYMBOL,
        })
    for commune in source["communes"]:
        units.append({
            "code": commune["code"], "name": nfc(commune["name"]),
            "canonicalName": nfc(commune["name"]), "level": "Commune",
            "type": unit_type(commune["name"]), "provinceCode": commune["provinceCode"],
            "effectiveFrom": EFFECTIVE_FROM, "effectiveTo": None,
            "status": "Active", "source": SOURCE_SYMBOL,
        })
    units.sort(key=lambda item: (0 if item["level"] == "Province" else 1, item["code"]))

    historical = [
        {
            "name": name, "normalizedName": nfc(name).casefold(),
            "formerCode": None, "formerLevel": "Province", "formerProvince": name,
            "effectiveFrom": None, "effectiveTo": "2025-06-30",
            "successorCode": successor[0], "successorName": successor[1],
            "status": "Historical", "source": "Nghị quyết 202/2025/QH15",
        }
        for name, successor in HISTORICAL_PROVINCES.items()
    ]
    historical.sort(key=lambda item: item["normalizedName"])

    special_names = set(CURATED_SPECIAL)
    special_names.update(item["canonicalName"] for item in units)
    special_names.update(item["name"] for item in historical)
    special = sorted((nfc(item) for item in special_names), key=lambda item: item.casefold())

    stable_write(ADMIN, units)
    stable_write(HISTORICAL, historical)
    stable_write(SPECIAL, special)
    stable_write(MANIFEST, {
        "schemaVersion": 1,
        "generatedBy": "tools/data/generate_administrative_reference_data.py",
        "source": SOURCE_SYMBOL,
        "sourceUrl": SOURCE_URL,
        "effectiveFrom": EFFECTIVE_FROM,
        "snapshotDate": metadata["snapshotDate"],
        "sourceFiles": metadata["sourceFiles"],
        "expectedCounts": {
            "activeProvinces": 34, "activeCommunes": 3321,
            "activeDistricts": 0, "historicalAliases": len(historical),
        },
        "outputs": {
            "administrativeUnitsSha256": file_sha256(ADMIN),
            "historicalAliasesSha256": file_sha256(HISTORICAL),
            "specialCapitalizationsSha256": file_sha256(SPECIAL),
        },
        "specialCapitalizationSources": [
            SOURCE_SYMBOL,
            "Nghị quyết 202/2025/QH15",
            "Nghị định 30/2020/NĐ-CP, Phụ lục II",
            "Hướng dẫn 05-HD/VPTW",
        ],
    })


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--extract-docx", nargs=2, type=Path, metavar=("PART1", "PART2"))
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()

    if args.extract_docx:
        stable_write(SOURCE, extract_source(args.extract_docx))
    if not SOURCE.exists():
        raise SystemExit(f"Missing checked-in source snapshot: {SOURCE}")
    source = json.loads(SOURCE.read_text(encoding="utf-8"))
    validate_source(source)
    if args.check:
        before = {path: path.read_bytes() if path.exists() else None for path in (ADMIN, HISTORICAL, SPECIAL, MANIFEST)}
        generate(source)
        changed = [str(path.relative_to(ROOT)) for path, data in before.items() if path.read_bytes() != data]
        if changed:
            raise SystemExit("Generated reference data was stale: " + ", ".join(changed))
    else:
        generate(source)


if __name__ == "__main__":
    main()
