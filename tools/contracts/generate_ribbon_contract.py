from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[2]
ACTUAL_PATH = ROOT / "shared" / "docs" / "implementation" / "evidence" / "ribbon_actual.json"
OUTPUT_PATH = ROOT / "shared" / "contracts" / "ribbon" / "ribbon-contract.v1.json"
TARGET_PRODUCT_LABEL = "Chuẩn hóa"
REMOVED_SAVE_CONTROL_IDS = {"btnLuuDocx"}
REMOVED_OPTIONAL_LANGUAGE_CONTROL_IDS = {
    "mnuIY",
    "btnKieuI",
    "btnKieuY",
}
REMOVED_MANUAL_READ_CONTROL_IDS = {"btnDocDuLieu"}
REMOVED_QR_CONTROL_IDS = {"btnChenQrCode"}
REMOVED_VIEW_CONTROL_IDS = {
    "chkRanhGioiVanBan",
    "chkDauGoc",
    "chkKyHieuSoanThao",
}
REMOVED_TARGET_CONTROL_IDS = (
    REMOVED_SAVE_CONTROL_IDS
    | REMOVED_OPTIONAL_LANGUAGE_CONTROL_IDS
    | REMOVED_MANUAL_READ_CONTROL_IDS
    | REMOVED_QR_CONTROL_IDS
    | REMOVED_VIEW_CONTROL_IDS
)


def meta(
    execution_mode: str,
    entitlement: str,
    risk_tier: str,
    mutation_scope: str,
    preconditions: list[str],
    capability: str,
    undo_policy: str,
    backup_policy: str,
) -> dict[str, Any]:
    return {
        "executionMode": execution_mode,
        "entitlement": entitlement,
        "riskTier": risk_tier,
        "mutationScope": mutation_scope,
        "preconditions": preconditions,
        "capability": capability,
        "undoPolicy": undo_policy,
        "backupPolicy": backup_policy,
    }


HAS_DOCUMENT = ["ACTIVE_DOCUMENT", "NO_ACTIVE_MUTATION_JOB"]
WRITABLE_DOCUMENT = ["ACTIVE_DOCUMENT", "DOCUMENT_WRITABLE", "NOT_PROTECTED", "NO_ACTIVE_MUTATION_JOB"]
SELECTION_REQUIRED = WRITABLE_DOCUMENT + ["VALID_SELECTION"]
TABLE_SELECTION_REQUIRED = WRITABLE_DOCUMENT + ["SELECTION_IN_TABLE"]
REMOTE_SCAN = HAS_DOCUMENT + ["AUTHENTICATED", "DEVICE_ACTIVE", "ENTITLEMENT_ACTIVE", "RELEASE_ALLOWED", "NETWORK_AVAILABLE"]
REMOTE_ANNOTATION = WRITABLE_DOCUMENT + [
    "AUTHENTICATED",
    "DEVICE_ACTIVE",
    "ENTITLEMENT_ACTIVE",
    "RELEASE_ALLOWED",
    "NETWORK_AVAILABLE",
    "FINDING_ANCHORS_EXACT",
    "DOCUMENT_FINGERPRINT_MATCH",
]
REMOTE_MUTATION = WRITABLE_DOCUMENT + [
    "AUTHENTICATED",
    "DEVICE_ACTIVE",
    "ENTITLEMENT_ACTIVE",
    "RELEASE_ALLOWED",
    "NETWORK_AVAILABLE",
    "SIGNED_FIX_PLAN_VALID",
    "DOCUMENT_FINGERPRINT_MATCH",
]
LOCAL_GRANT_MUTATION = WRITABLE_DOCUMENT + [
    "AUTHENTICATED",
    "DEVICE_ACTIVE",
    "ENTITLEMENT_ACTIVE",
    "RELEASE_ALLOWED",
    "SIGNED_EXECUTION_GRANT_VALID",
    "DOCUMENT_FINGERPRINT_MATCH",
]


CONTROL_META: dict[str, dict[str, Any]] = {
    "btnAutoFixAll2026": meta("SIGNED_FIX_PLAN", "AUTOFIX", "CONFIRM", "AUTHORIZED_FIX_PLAN_OPERATIONS", REMOTE_MUTATION, "WORD_2010_OBJECT_MODEL", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED_BEFORE_APPLY"),
    "ddQuyDinh": meta("LOCAL_CONTEXT", "DOCUMENT_READ", "SAFE", "DOCUMENT_CONTEXT_ONLY", ["ACTIVE_DOCUMENT"], "ALL_SUPPORTED_WORD", "NOT_APPLICABLE", "NOT_APPLICABLE"),
    "ddLoaiVanBan": meta("LOCAL_CONTEXT_WITH_SERVER_CATALOG", "DOCUMENT_READ", "SAFE", "DOCUMENT_CONTEXT_ONLY", ["ACTIVE_DOCUMENT", "REGIME_SELECTED"], "DYNAMIC_DROPDOWN", "NOT_APPLICABLE", "NOT_APPLICABLE"),
    "btnKiemTra": meta("REMOTE_FINDINGS", "SCAN_FORMAT", "REPORT_ONLY", "OWNED_FINDING_ANNOTATIONS_ONLY", REMOTE_ANNOTATION, "WORD_SNAPSHOT_AND_EXACT_ANCHOR_V1", "ONE_CUSTOM_UNDO_RECORD_WHEN_SUPPORTED", "NOT_APPLICABLE"),
    "btnKiemTraChinhTa": meta("REMOTE_FINDINGS", "SCAN_SPELLING", "REPORT_ONLY", "OWNED_FINDING_ANNOTATIONS_ONLY", REMOTE_ANNOTATION, "UNICODE_SNAPSHOT_AND_EXACT_ANCHOR_V1", "ONE_CUSTOM_UNDO_RECORD_WHEN_SUPPORTED", "NOT_APPLICABLE"),
    "btnChuyenDoiUnicode": meta("SIGNED_FIX_PLAN", "ENCODING_CONVERT", "CONFIRM", "AUTHORIZED_STORY_RANGES", REMOTE_MUTATION, "TCVN3_FULL_VNI_GATED", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED_BEFORE_APPLY"),
    "btnDinhDangTrangGiay": meta("SIGNED_FIX_PLAN", "PAGE_FORMAT", "SAFE", "AUTHORIZED_SECTIONS_PAGE_SETUP", REMOTE_MUTATION, "WORD_SECTION_PAGE_SETUP", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED_FOR_MULTI_SECTION"),
    "btnChenTrangNgang": meta("LOCAL_EXECUTION_GRANT", "SECTION_INSERT", "CONFIRM", "SELECTION_AND_NEW_SECTIONS", LOCAL_GRANT_MUTATION + ["VALID_SELECTION"], "WORD_SECTION_BREAKS", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED"),
    "btnChenTrangDoc": meta("LOCAL_EXECUTION_GRANT", "SECTION_INSERT", "CONFIRM", "SELECTION_AND_NEW_SECTIONS", LOCAL_GRANT_MUTATION + ["VALID_SELECTION"], "WORD_SECTION_BREAKS", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED"),
    "btnXoaTrangThua": meta("SIGNED_FIX_PLAN", "TRAILING_PAGE_REMOVE", "CONFIRM", "VERIFIED_TRAILING_EMPTY_PAGE_ONLY", REMOTE_MUTATION, "WORD_PAGE_COUNT_VERIFY", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED"),
    "mnuDungBoStyle": meta("CONTAINER", "STYLE_BUILD", "SAFE", "NONE", ["ACTIVE_DOCUMENT"], "ALL_SUPPORTED_WORD", "NOT_APPLICABLE", "NOT_APPLICABLE"),
    "btnDungBoStyleCo15": meta("SIGNED_FIX_PLAN", "STYLE_BUILD", "SAFE", "AUTHORIZED_STYLE_DEFINITIONS", REMOTE_MUTATION, "WORD_STYLES", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED_IF_STYLE_COLLISION"),
    "btnDungBoStyleCo14": meta("SIGNED_FIX_PLAN", "STYLE_BUILD", "SAFE", "AUTHORIZED_STYLE_DEFINITIONS", REMOTE_MUTATION, "WORD_STYLES", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED_IF_STYLE_COLLISION"),
    "btnDungBoStyleCo13": meta("SIGNED_FIX_PLAN", "STYLE_BUILD", "SAFE", "AUTHORIZED_STYLE_DEFINITIONS", REMOTE_MUTATION, "WORD_STYLES", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED_IF_STYLE_COLLISION"),
    "btnCoChu15": meta("SIGNED_FIX_PLAN", "FONT_SIZE_SET", "CONFIRM", "AUTHORIZED_COMPONENT_RANGES", REMOTE_MUTATION, "ROLE_AWARE_FORMATTING", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED"),
    "btnCoChu14": meta("SIGNED_FIX_PLAN", "FONT_SIZE_SET", "CONFIRM", "AUTHORIZED_COMPONENT_RANGES", REMOTE_MUTATION, "ROLE_AWARE_FORMATTING", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED"),
    "btnCoChu13": meta("SIGNED_FIX_PLAN", "FONT_SIZE_SET", "CONFIRM", "AUTHORIZED_COMPONENT_RANGES", REMOTE_MUTATION, "ROLE_AWARE_FORMATTING", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED"),
    "btnKeepWithNext": meta("LOCAL_EXECUTION_GRANT", "PARAGRAPH_KEEP", "SAFE", "SELECTED_OR_CURRENT_PARAGRAPHS", LOCAL_GRANT_MUTATION + ["VALID_SELECTION_OR_CARET"], "WORD_PARAGRAPH_FORMAT", "ONE_CUSTOM_UNDO_RECORD", "NOT_REQUIRED"),
    "btnChenSoTrang": meta("SIGNED_FIX_PLAN", "PAGE_NUMBERING", "CONFIRM", "AUTHORIZED_HEADERS_FOOTERS", REMOTE_MUTATION, "WORD_SECTION_HEADER_FOOTER", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED"),
    "btnCoChu": meta("LOCAL_EXECUTION_GRANT", "CHARACTER_SPACING", "SAFE", "SELECTED_TEXT", SELECTION_REQUIRED + ["SIGNED_EXECUTION_GRANT_VALID"], "WORD_FONT_SPACING", "ONE_CUSTOM_UNDO_RECORD", "NOT_REQUIRED"),
    "btnGianChuNormal": meta("LOCAL_EXECUTION_GRANT", "CHARACTER_SPACING", "SAFE", "SELECTED_TEXT", SELECTION_REQUIRED + ["SIGNED_EXECUTION_GRANT_VALID"], "WORD_FONT_SPACING", "ONE_CUSTOM_UNDO_RECORD", "NOT_REQUIRED"),
    "btnGianChuRa": meta("LOCAL_EXECUTION_GRANT", "CHARACTER_SPACING", "SAFE", "SELECTED_TEXT", SELECTION_REQUIRED + ["SIGNED_EXECUTION_GRANT_VALID"], "WORD_FONT_SPACING", "ONE_CUSTOM_UNDO_RECORD", "NOT_REQUIRED"),
    "btnLapDongTieuDe": meta("SIGNED_FIX_PLAN", "TABLE_HEADER_REPEAT", "SAFE", "AUTHORIZED_OUTER_TABLE_ROWS", REMOTE_MUTATION, "WORD_TABLE_HEADER", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED_FOR_MULTI_TABLE"),
    "btnChuanHoaBang": meta("SIGNED_FIX_PLAN", "TABLE_FORMAT", "CONFIRM", "AUTHORIZED_TABLES", REMOTE_MUTATION, "WORD_TABLE_AUTOFIT", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED"),
    "btnChuanHoaAnh": meta("SIGNED_FIX_PLAN", "IMAGE_FORMAT", "CONFIRM", "AUTHORIZED_IMAGES", REMOTE_MUTATION, "WORD_INLINE_AND_FLOATING_SHAPES", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED"),
    "btnCanDinhO": meta("LOCAL_EXECUTION_GRANT", "CELL_ALIGNMENT", "SAFE", "SELECTED_TABLE_CELLS", TABLE_SELECTION_REQUIRED + ["SIGNED_EXECUTION_GRANT_VALID"], "WORD_CELL_VERTICAL_ALIGNMENT", "ONE_CUSTOM_UNDO_RECORD", "NOT_REQUIRED"),
    "btnCanGiuaO": meta("LOCAL_EXECUTION_GRANT", "CELL_ALIGNMENT", "SAFE", "SELECTED_TABLE_CELLS", TABLE_SELECTION_REQUIRED + ["SIGNED_EXECUTION_GRANT_VALID"], "WORD_CELL_ALIGNMENT", "ONE_CUSTOM_UNDO_RECORD", "NOT_REQUIRED"),
    "btnXoaKyTuThuaBangExcel": meta("SIGNED_FIX_PLAN", "TABLE_TEXT_CLEAN", "CONFIRM", "AUTHORIZED_TABLE_CELL_RANGES", REMOTE_MUTATION + ["SELECTION_IN_TABLE_OR_IDENTIFIED_TABLES"], "PROTECTED_SPANS_V1", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED"),
    "mnuBoDau": meta("CONTAINER", "TONE_PLACEMENT", "SAFE", "NONE", ["ACTIVE_DOCUMENT"], "ALL_SUPPORTED_WORD", "NOT_APPLICABLE", "NOT_APPLICABLE"),
    "btnKieuOaUy": meta("LOCAL_EXECUTION_GRANT", "TONE_PLACEMENT", "CONFIRM", "ALL_EDITABLE_STORIES", LOCAL_GRANT_MUTATION, "VIETNAMESE_TONE_MAIN_VOWEL", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED"),
    "btnKieuOaUy2": meta("LOCAL_EXECUTION_GRANT", "TONE_PLACEMENT", "CONFIRM", "ALL_EDITABLE_STORIES", LOCAL_GRANT_MUTATION, "VIETNAMESE_TONE_FIRST_VOWEL", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED"),
    "btnDoiDauThapPhan": meta("SIGNED_FIX_PLAN", "DECIMAL_NORMALIZE", "CONFIRM", "AUTHORIZED_NUMERIC_TEXT_RANGES", REMOTE_MUTATION, "NUMERIC_PROTECTED_SPANS", "ONE_CUSTOM_UNDO_RECORD", "REQUIRED"),
    "mnuThongTinTienIch": meta("CONTAINER", "ABOUT", "SAFE", "NONE", [], "ALL_SUPPORTED_WORD", "NOT_APPLICABLE", "NOT_APPLICABLE"),
    "btnKiemTraPhienBanMoi": meta("REMOTE_RELEASE_STATUS", "ABOUT", "REPORT_ONLY", "NONE_READ_ONLY", ["RELEASE_METADATA_AVAILABLE_OR_REFRESHABLE"], "SYSTEM_BROWSER_OPTIONAL", "NOT_APPLICABLE", "NOT_APPLICABLE"),
    "btnGuiPhanHoi": meta("SYSTEM_BROWSER", "FEEDBACK", "REPORT_ONLY", "NONE_READ_ONLY", ["OFFICIAL_FEEDBACK_URL_CONFIGURED", "USER_CONSENT_FOR_TECHNICAL_METADATA"], "SYSTEM_BROWSER", "NOT_APPLICABLE", "NOT_APPLICABLE"),
    "btnGioiThieu": meta("LOCAL_SIGNED_METADATA", "ABOUT", "REPORT_ONLY", "NONE_READ_ONLY", [], "TRANSIENT_DIALOG", "NOT_APPLICABLE", "NOT_APPLICABLE"),
}

SELECTED_FINDING_CONTROL: dict[str, Any] = {
    "groupId": "grpKhoiDong",
    "parentContainerId": None,
    "controlType": "button",
    "id": "btnSuaLoiDangChon",
    "label": "Sửa lỗi đang chọn",
    "size": "large",
    "imageMso": "ReviewAcceptChange",
    "callbacks": {
        "getEnabled": "GetEnabledAutoFixAll2026",
        "onAction": "OnSuaLoiDangChon",
    },
    "screenTip": "Sửa lỗi đang chọn",
    "superTip": (
        "Bấm vào phần văn bản đang được Chuẩn hóa comment hoặc tô đỏ, sau đó bấm nút này "
        "để sửa một lỗi có phương án an toàn. Comment sẽ biến mất khi kiểm tra lại xác nhận lỗi đã hết."
    ),
    "staticItems": [],
}

SELECTED_FINDING_META = meta(
    "LOCAL_SIGNED_RULE_FIX",
    "AUTOFIX",
    "CONFIRM",
    "SELECTED_OWNED_FINDING_ONLY",
    WRITABLE_DOCUMENT + ["ENTITLEMENT_ACTIVE", "SELECTED_OWNED_FINDING", "DETERMINISTIC_SAFE_FIX_AVAILABLE"],
    "SELECTED_FINDING_FIX_V1",
    "ONE_CUSTOM_UNDO_RECORD_WHEN_SUPPORTED",
    "NOT_REQUIRED",
)

QUICK_SPELLING_CONTROL: dict[str, Any] = {
    "groupId": "grpKhoiDong",
    "parentContainerId": None,
    "controlType": "button",
    "id": "btnSuaTatCaChinhTa",
    "label": "Sửa nhanh chính tả",
    "size": "large",
    "imageMso": "AutoCorrect",
    "callbacks": {
        "getEnabled": "GetEnabledAutoFixAll2026",
        "onAction": "OnSuaTatCaChinhTa",
    },
    "screenTip": "Sửa nhanh chính tả và dọn lỗi gõ",
    "superTip": (
        "Sửa các lỗi chính tả rule-based có phương án an toàn; đồng thời dọn "
        "khoảng trắng, khoảng cách dấu câu và chuẩn hóa dấu ngoặc."
    ),
    "staticItems": [],
}

QUICK_SPELLING_META = meta(
    "LOCAL_APPLY", "AUTOFIX", "CONFIRM", "SPELLING_AND_TYPOGRAPHY_ALL_EDITABLE_STORIES",
    WRITABLE_DOCUMENT, "WORD_2010_OBJECT_MODEL", "ONE_CUSTOM_UNDO_RECORD", "NOT_APPLICABLE"
)

PERSONAL_DICTIONARY_CONTROL: dict[str, Any] = {
    "groupId": "grpChinhTaSo",
    "parentContainerId": None,
    "controlType": "button",
    "id": "btnTuDienCaNhan",
    "label": "Từ điển cá nhân",
    "size": "large",
    "imageMso": None,
    "callbacks": {
        "getEnabled": "GetEnabledAlways",
        "getImage": "GetImageTuDienCaNhan",
        "onAction": "OnTuDienCaNhan",
    },
    "screenTip": "Quản lý từ điển cá nhân",
    "superTip": "Xem, thêm và xóa thuật ngữ riêng mà bộ kiểm tra chính tả cần bỏ qua.",
    "staticItems": [],
}

PERSONAL_DICTIONARY_META = meta(
    "LOCAL_DIALOG", "DOCUMENT_TOOLS", "SAFE", "NONE", [],
    "WINFORMS_MODAL_DIALOG", "NOT_APPLICABLE", "NOT_APPLICABLE"
)

TARGET_CONTROL_OVERRIDES: dict[str, dict[str, str]] = {
    "btnKiemTra": {
        "superTip": "Kiểm tra khổ giấy, lề, kiểu chữ và thành phần thể thức; thêm comment giải thích và tô đỏ đúng chỗ sai.",
    },
    "btnKiemTraChinhTa": {
        "superTip": "Kiểm tra từ điển chính tả và viết hoa; thêm comment giải thích và tô đỏ đúng chỗ sai.",
    },
    "mnuBoDau": {
        "label": "Đặt dấu thanh",
        "screenTip": "Đặt dấu thanh",
        "superTip": "Đồng nhất cách đặt dấu theo kiểu oà, uý hoặc òa, úy; không thay đổi quy tắc i/y.",
    },
}


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n")


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    actual = json.loads(ACTUAL_PATH.read_text(encoding="utf-8"))
    actual_controls = [
        control
        for control in actual["controls"]
        if control["id"] not in REMOVED_TARGET_CONTROL_IDS
    ]
    actual_ids = {control["id"] for control in actual_controls}
    metadata_ids = set(CONTROL_META)
    if actual_ids != metadata_ids:
        missing = sorted(actual_ids - metadata_ids)
        extra = sorted(metadata_ids - actual_ids)
        raise RuntimeError(f"Metadata Ribbon không khớp; thiếu={missing}, dư={extra}")

    controls: list[dict[str, Any]] = []
    for control in actual_controls:
        control_id = control["id"]
        target = dict(control)
        target.update(TARGET_CONTROL_OVERRIDES.get(control_id, {}))
        target["order"] = len(controls) + 1
        target["callbacks"] = dict(target["callbacks"])
        target["commandContract"] = {
            "commandId": control_id,
            **CONTROL_META[control_id],
            "telemetryAllowlist": ["commandId", "resultCode", "durationBucket", "clientRelease", "wordVersion", "bitness"],
            "testIds": [f"RIBBON-{target['order']:03d}", f"CONTROL-{control_id}"],
        }
        if control_id == "btnAutoFixAll2026":
            target["callbacks"]["onAction"] = "OnAutoFixAll2026"
            target["callbacks"]["getEnabled"] = "GetEnabledAutoFixAll2026"
        controls.append(target)

    selected_finding = dict(SELECTED_FINDING_CONTROL)
    selected_finding["order"] = len(controls) + 1
    selected_finding["commandContract"] = {
        "commandId": selected_finding["id"],
        **SELECTED_FINDING_META,
        "telemetryAllowlist": ["commandId", "resultCode", "durationBucket", "clientRelease", "wordVersion", "bitness"],
        "testIds": [
            f"RIBBON-{selected_finding['order']:03d}",
            "CONTROL-btnSuaLoiDangChon",
            "SELECTED-FINDING-FIX-001",
        ],
    }
    controls.append(selected_finding)

    for synthetic_control, synthetic_meta in (
        (QUICK_SPELLING_CONTROL, QUICK_SPELLING_META),
        (PERSONAL_DICTIONARY_CONTROL, PERSONAL_DICTIONARY_META),
    ):
        target = dict(synthetic_control)
        target["callbacks"] = dict(synthetic_control["callbacks"])
        target["order"] = len(controls) + 1
        target["commandContract"] = {
            "commandId": target["id"],
            **synthetic_meta,
            "telemetryAllowlist": ["commandId", "resultCode", "durationBucket", "clientRelease", "wordVersion", "bitness"],
            "testIds": [f"RIBBON-{target['order']:03d}", f"CONTROL-{target['id']}"],
        }
        controls.append(target)

    button_controls = [control for control in controls if control["controlType"] == "button"]
    menu_controls = [control for control in controls if control["controlType"] == "menu"]
    if len(controls) != 39 or len(button_controls) != 34 or len(menu_controls) != 3:
        raise RuntimeError("Ribbon contract không đạt 39 controls, 34 button commands và 3 menu containers")
    if any("onAction" not in control["callbacks"] for control in button_controls):
        raise RuntimeError("Có button thiếu onAction trong contract đích")
    if any(control["commandContract"]["executionMode"] != "CONTAINER" for control in menu_controls):
        raise RuntimeError("Menu phải là container, không phải command handler")

    contract = {
        "schemaVersion": 1,
        "contractId": "CHUAN_HOA_WORD_RIBBON_V1",
        "sourceArtefact": actual["artefact"],
        "sourceArtefactSha256": actual["artefactSha256"],
        "ribbonSchema": actual["schema"],
        "tab": {
            **actual["tab"],
            "label": TARGET_PRODUCT_LABEL,
        },
        "groups": [group for group in actual["groups"] if group["id"] != "grpHienThi"],
        "counts": {
            "tabs": 1,
            "groups": 6,
            "buttons": len(button_controls),
            "menus": len(menu_controls),
            "dropDowns": sum(control["controlType"] == "dropDown" for control in controls),
            "checkBoxes": sum(control["controlType"] == "checkBox" for control in controls),
            "interactiveControls": len(controls),
        },
        "approvedProductChanges": [
            {
                "changeId": "REMOVE_SAVE_AS_DOCX_AND_SUPPORT_DOC_DOCX",
                "removedControlIds": sorted(REMOVED_SAVE_CONTROL_IDS),
                "supportedDocumentExtensions": [".doc", ".docx"],
                "preserveCurrentDocumentFormat": True,
            },
            {
                "changeId": "REMOVE_IY_NORMALIZATION_KEEP_TONE_PLACEMENT",
                "removedControlIds": sorted(REMOVED_OPTIONAL_LANGUAGE_CONTROL_IDS),
            },
            {
                "changeId": "PRODUCT_NAME_CHUAN_HOA",
                "ribbonTabLabel": TARGET_PRODUCT_LABEL,
            },
            {
                "changeId": "REMOVE_MANUAL_READ_DATA_USE_COMMAND_SCOPED_ANALYSIS",
                "removedControlIds": sorted(REMOVED_MANUAL_READ_CONTROL_IDS),
                "documentReadPolicy": "USER_INITIATED_COMMAND_ONLY",
                "analysisPolicy": "MINIMUM_REQUIRED_LANES",
            },
            {
                "changeId": "REMOVE_QR_FEATURE",
                "removedControlIds": sorted(REMOVED_QR_CONTROL_IDS),
                "productDecision": "QR is intentionally retired from the current product",
            },
            {
                "changeId": "REMOVE_VIEW_OPTIONS_GROUP",
                "removedControlCount": len(REMOVED_VIEW_CONTROL_IDS),
                "removedGroupCount": 1,
                "productDecision": "Word view toggles are intentionally left to Word Options",
            },
            {
                "changeId": "MERGE_TYPOGRAPHY_INTO_QUICK_SPELLING",
                "commandId": "btnSuaTatCaChinhTa",
                "mergedCapabilities": [
                    "RULE_BASED_SAFE_SPELLING_FIX",
                    "WHITESPACE_AND_PUNCTUATION_CLEANUP",
                    "STANDARD_BRACKET_AND_QUOTATION_MARKS",
                ],
                "separateRibbonButtons": False,
            },
        ],
        "invariants": [
            "EXACTLY_ONE_CUSTOM_TAB",
            "NO_PERSISTENT_TASK_PANE",
            "SCAN_COMMANDS_ARE_READ_ONLY",
            "DOCUMENT_AND_WINDOW_STATE_ARE_NOT_GLOBAL_STATIC",
            "DISABLED_COMMANDS_EXPOSE_A_REASON",
            "UNKNOWN_OPERATION_FAILS_CLOSED",
        ],
        "controls": controls,
    }
    write_json(OUTPUT_PATH, contract)
    print(
        json.dumps(
            {
                "status": "PASS",
                "output": str(OUTPUT_PATH),
                "controls": len(controls),
                "buttonCommands": len(button_controls),
                "menuContainers": len(menu_controls),
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
