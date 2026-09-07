#!/usr/bin/env python3
"""Fail closed when production language/reference assets regress."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import tempfile
import unicodedata
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[2]
DICTIONARIES = ROOT / "shared/dictionaries"
ADMIN = DICTIONARIES / "administrative_units.json"
HISTORICAL = DICTIONARIES / "historical_administrative_aliases.json"
SPECIAL = DICTIONARIES / "special_capitalizations.json"
MANIFEST = DICTIONARIES / "reference_data_manifest.json"
CONFUSION_SOURCE = ROOT / "src/ChuanHoa.Client.Core/Lexicon/VietnameseConfusionSets.cs"


def load_unique_json(path: Path) -> Any:
    def unique(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f"{path}: duplicate JSON key {key!r}")
            result[key] = value
        return result

    return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=unique)


def normalized(value: str) -> str:
    return unicodedata.normalize("NFC", value).strip()


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def validate(root: Path = DICTIONARIES,
             confusion_source: Path = CONFUSION_SOURCE) -> dict[str, int]:
    admin_path = root / ADMIN.name
    history_path = root / HISTORICAL.name
    special_path = root / SPECIAL.name
    manifest_path = root / MANIFEST.name
    units = load_unique_json(admin_path)
    historical = load_unique_json(history_path)
    special = load_unique_json(special_path)
    manifest = load_unique_json(manifest_path)

    require(isinstance(units, list), "Administrative data must be a JSON array")
    require(isinstance(historical, list), "Historical aliases must be a JSON array")
    require(isinstance(special, list), "Special capitalizations must be a JSON array")
    require(manifest.get("source") and manifest.get("sourceUrl"), "Reference-data source metadata is missing")
    require(manifest.get("effectiveFrom") == "2025-07-01", "Current effective date is missing or stale")

    required = {"code", "name", "canonicalName", "level", "type", "provinceCode",
                "effectiveFrom", "effectiveTo", "status", "source"}
    active = [item for item in units if item.get("status") == "Active"]
    provinces = [item for item in active if item.get("level") == "Province"]
    communes = [item for item in active if item.get("level") == "Commune"]
    require(len(provinces) == 34, f"Expected 34 active provinces; got {len(provinces)}")
    require(len(communes) == 3321, f"Expected 3321 active communes; got {len(communes)}")
    require(not any(item.get("level") == "District" for item in active), "Active district layer is forbidden")

    province_codes = {item["code"] for item in provinces}
    require(len(province_codes) == len(provinces), "Duplicate active province code")
    commune_codes = {item["code"] for item in communes}
    require(len(commune_codes) == len(communes), "Duplicate active commune code")
    require(province_codes.isdisjoint(commune_codes), "Province and commune codes overlap")
    representative = {"01": "Thành phố Hà Nội", "46": "Thành phố Huế",
                      "48": "Thành phố Đà Nẵng", "79": "Thành phố Hồ Chí Minh",
                      "92": "Thành phố Cần Thơ"}
    by_province_code = {item["code"]: item["canonicalName"] for item in provinces}
    require(all(by_province_code.get(code) == name for code, name in representative.items()),
            "Representative province codes/names do not match Decision 19/2025/QD-TTg")

    allowed_types = {"Province", "CentrallyGovernedCity", "Commune", "Ward", "SpecialZone"}
    for item in active:
        require(required.issubset(item), f"Administrative record is missing fields: {item!r}")
        code = item["code"]
        expected_length = 2 if item["level"] == "Province" else 5
        require(isinstance(code, str) and re.fullmatch(rf"\d{{{expected_length}}}", code) is not None,
                f"Invalid administrative code: {code!r}")
        name = item["canonicalName"]
        require(isinstance(name, str) and len(normalized(name)) >= 4 and any(c.isalpha() for c in name),
                f"Invalid administrative name: {name!r}")
        require(name == normalized(name) and unicodedata.is_normalized("NFC", name),
                f"Administrative name is not normalized NFC: {name!r}")
        require(item["type"] in allowed_types, f"Invalid administrative type: {item['type']!r}")
        require(item["source"] and item["effectiveFrom"] == "2025-07-01",
                f"Missing provenance/effective date for {code}")
        if item["level"] == "Commune":
            require(item["provinceCode"] in province_codes, f"Unknown parent province for {code}")

    historical_names: set[str] = set()
    for item in historical:
        require(item.get("status") == "Historical", "Historical record incorrectly marked active")
        name = item.get("name", "")
        key = normalized(name).casefold()
        require(len(normalized(name)) >= 4 and key not in historical_names, f"Invalid/duplicate historical alias: {name!r}")
        historical_names.add(key)
        require(item.get("effectiveTo") and item.get("source"), f"Historical provenance missing: {name!r}")
        require(item.get("successorCode") in province_codes, f"Historical successor is invalid: {name!r}")

    special_keys: set[str] = set()
    for entry in special:
        require(isinstance(entry, str), "Special capitalization entry must be text")
        clean = normalized(entry)
        key = clean.casefold()
        require(clean == entry and unicodedata.is_normalized("NFC", entry), f"Special capitalization is not NFC: {entry!r}")
        require(len(clean) >= 4 and any(c.isalpha() for c in clean), f"Special capitalization fragment: {entry!r}")
        require(not re.search(r"[{}<>]|\b(?:function|private|public|version)\b", clean, re.I),
                f"Special capitalization contains source/code debris: {entry!r}")
        require(key not in special_keys, f"Duplicate special capitalization: {entry!r}")
        special_keys.add(key)
    require(all(item["canonicalName"].casefold() in special_keys for item in active),
            "Current administrative names are missing from capitalization data")

    for dictionary_name in ("typo_dictionary.json", "iy_dictionary.json"):
        corrections = load_unique_json(root / dictionary_name)
        require(isinstance(corrections, dict), f"{dictionary_name} must be a JSON object")
        for wrong, replacement in corrections.items():
            require(isinstance(replacement, str) and normalized(replacement), f"Empty correction: {wrong!r}")
            require(normalized(wrong).casefold() != normalized(replacement).casefold(),
                    f"Identity correction in {dictionary_name}: {wrong!r}")
            require(not (wrong in corrections and corrections.get(replacement) == wrong),
                    f"Correction cycle in {dictionary_name}: {wrong!r}")

    confusion_text = confusion_source.read_text(encoding="utf-8")
    marker = "AdministrativeConfusionPairs ="
    require(marker in confusion_text, "Administrative confusion mapping declaration is missing")
    confusion_block = confusion_text.split(marker, 1)[1].split("/// <summary>", 1)[0]
    confusion_pairs = re.findall(r'\{\s*"([^"]+)",\s*"([^"]+)"\s*\}', confusion_block)
    require(confusion_pairs, "Administrative confusion mappings were not found")
    normalized_pairs = {normalized(wrong).casefold(): normalized(replacement).casefold()
                        for wrong, replacement in confusion_pairs}
    require(len(normalized_pairs) == len(confusion_pairs),
            "Duplicate administrative confusion mapping")
    for wrong, replacement in normalized_pairs.items():
        require(wrong != replacement,
                f"Identity correction in VietnameseConfusionSets: {wrong!r}")
        require(normalized_pairs.get(replacement) != wrong,
                f"Correction cycle in VietnameseConfusionSets: {wrong!r}")

    outputs = manifest.get("outputs", {})
    require(outputs.get("administrativeUnitsSha256") == sha256(admin_path), "Administrative checksum mismatch")
    require(outputs.get("historicalAliasesSha256") == sha256(history_path), "Historical checksum mismatch")
    require(outputs.get("specialCapitalizationsSha256") == sha256(special_path), "Capitalization checksum mismatch")
    counts = manifest.get("expectedCounts", {})
    require(counts == {"activeProvinces": 34, "activeCommunes": 3321,
                       "activeDistricts": 0, "historicalAliases": len(historical)},
            "Manifest counts do not match validated data")
    return {"activeProvinces": len(provinces), "activeCommunes": len(communes),
            "activeDistricts": 0, "historicalAliases": len(historical),
            "specialCapitalizations": len(special)}


def self_test() -> None:
    with tempfile.TemporaryDirectory(prefix="ChuanHoa-reference-data-") as directory:
        target = Path(directory)
        for path in (ADMIN, HISTORICAL, SPECIAL, MANIFEST,
                     DICTIONARIES / "typo_dictionary.json", DICTIONARIES / "iy_dictionary.json"):
            (target / path.name).write_bytes(path.read_bytes())
        data = load_unique_json(target / ADMIN.name)
        data[0]["status"] = "Historical"
        (target / ADMIN.name).write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        try:
            validate(target)
        except ValueError:
            pass
        else:
            raise ValueError("Reference-data self-test failed to reject a corrupted active count")

        confusion = target / CONFUSION_SOURCE.name
        confusion.write_text(CONFUSION_SOURCE.read_text(encoding="utf-8").replace(
            "        /// <summary>",
            '                { "same", "same" },\n\n        /// <summary>', 1),
            encoding="utf-8")
        for path in (ADMIN, HISTORICAL, SPECIAL, MANIFEST,
                     DICTIONARIES / "typo_dictionary.json", DICTIONARIES / "iy_dictionary.json"):
            (target / path.name).write_bytes(path.read_bytes())
        try:
            validate(target, confusion)
        except ValueError:
            return
        raise ValueError("Reference-data self-test failed to reject an identity correction")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        self_test()
    counts = validate()
    print("REFERENCE_DATA: PASS " + " ".join(f"{key}={value}" for key, value in counts.items()))


if __name__ == "__main__":
    main()
