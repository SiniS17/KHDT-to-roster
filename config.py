"""User-editable labels for roster markers and the Excel log."""

# ------------------------------------------------------------
# App version  (must be plain semver so the updater can compare)
# Bump this string before each release and tag the GitHub
# release with the same value prefixed by "v" (e.g. "v1.0.0").
# ------------------------------------------------------------
APP_VERSION = "1.0.0"
APP_LAST_UPDATE = "29 AUG 26"

# ------------------------------------------------------------
# Auto-update settings
# Set GITHUB_REPO to "your-username/your-repo-name" to enable.
# Leave as "" to disable update checks entirely.
# For private repos add a fine-grained token with "Contents: read"
# permission as an env-var called GITHUB_UPDATE_TOKEN.
# ------------------------------------------------------------
GITHUB_REPO = "SiniS17/KHDT-to-roster"          # ← fill this in, e.g. "acme/amos-validator"
UPDATE_CHECK_ENABLED = bool(GITHUB_REPO)   # auto-disabled when repo is blank


# Prefix written in front of every course marker.
MARKER_PREFIX = "H"

# Class-code prefix (the part before the last two "/"-separated segments)
# mapped to the short label used in the roster.
COURSE_ABBREVIATIONS = {
    "ATHK": "ATHK",
    "VHNT-GD2": "VHNT",
    "VHDN-VAE": "VHDN",
    "AMOS-LB": "AMOS",
    "EWIS-ADV-I": "EWIS",
    "A320/321-B12-P": "CL321 PRAC",
}

# Full course names mapped to short labels when the class-code field is empty.
# Matching ignores accents, case, and repeated spaces.
COURSE_NAME_ABBREVIATIONS = {
    "Lớp 3 – INS-COA/2603/S - Thực hành nhóm 1": "INS-COA G1",
    "Lớp 3 – INS-COA/2603/S - Thực hành nhóm 2": "INS-COA G2",
    "Cập nhật chính sách mới về thuế và kế toán,": "CN THUẾ KT",
    "Văn hóa doanh nghiệp VAECO": "VHDN",
    "Đào tạo, huấn luyện định kỳ - Vận hành thiết bị điều hòa không khí": "ĐT ĐỊNH KỲ ĐHKK",
    "EASA MTOE": "EASA MTOE",
    "EWIS (Electrical Wiring Interconnection Systems) - Advanced Initial Training": "EWIS ADV",
    "A320/321 (CFM56/V2500) To A320/321 (PW1100G) Aircraft Maintenance Difference Course - Theoretical Part": "CL321 NEO THEO",
    "A320/321 (CFM56/V2500) To A320/321 (PW1100G) Aircraft Maintenance Difference Course - Practical Part": "CL321 NEO PRAC",
    "Type Training - A320/321 (CFM56/V2500) Cat B1+B2 - Theoretical Part": "CL321 THEO",
    "Supplement Practical Basic for completing section 4 - Knowledge and Skill Training - CAT B1": "SECTION 4",
    "An toàn hàng không": "ATHK",
    "Human Factor Continuation Training for EASA Roster": "HF CONT EASA",
}

# Change these values if the log should use another language or terminology.
# In particular, LOG_COLUMN_HEADERS["event"] controls the "Event" column name.
LOG_COLUMN_HEADERS = {
    "name": "Name",
    "id": "ID",
    "date": "Date",
    "event": "Event",
    "reason": "Reason",
    "details": "Details",
}

LOG_EVENT_LABELS = {
    "note": "Đã xếp lịch",
    "conflict": "Có lịch đi làm",
    "overlap": "Nhiều lớp",
    "overwrite_day_off": "Học trùng N đen",
    "guarantee_day_off_conflict": "Học trùng N đỏ",
    "name_mismatch": "Sai tên",
    "skipped": "Bỏ qua",
    "no_date_overlap": "Ngoài khoảng ngày",
}