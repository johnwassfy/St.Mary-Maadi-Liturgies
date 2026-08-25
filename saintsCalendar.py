# saintsCalendar.py
# Centralized saints and angels feast calendar.
# Same pattern as find_Readings_Date — pure in-memory dicts, zero I/O.
#
# HOW TO ADD A NEW SAINT:
#   1. Author slides in the relevant .pptx template under Data\CopyData\
#   2. Run update_section_names() to extract section IDs into Files Data.xlsx
#   3. Add saint_id → feast days to SAINTS_CALENDAR
#   4. Add saint_id → GUIDs per file to SAINTS_SECTIONS
#   5. For month-variant sections, add a helper like get_mikhael_mrd_ebrksis_id()

# ---------------------------------------------------------------------------
# SAINTS CALENDAR
# saint_id → list of (coptic_month, coptic_day)
# coptic_month = 0  means "every month"
# ---------------------------------------------------------------------------
SAINTS_CALENDAR = {
    7: [(0, 12)],   # رئيس الملائكة ميخائيل — كل شهر يوم ١٢
    8: [(0, 21)],   # رئيس الملائكة غبريال  — كل شهر يوم ٢١
    # أضف هنا بعد إنشاء الشرائح:
    # 212: [(8, 23)],
    # 305: [(4, 28), (8, 5)],
}

# ---------------------------------------------------------------------------
# SAINTS SECTIONS
# saint_id → { file_key: { "show": [GUIDs], "hide": [GUIDs] } }
#
# file_key matches sheet names in Files Data.xlsx exactly:
#   "القداس" | "رفع بخور" | "التسبحة" | "تسبحة كيهك" | "الذكصولوجيات"
#
# GUIDs are copied from Files Data.xlsx after authoring and extracting —
# same workflow used in odasat.py / baker.py / tasbha.py everywhere.
#
# Michael note: his fixed three sections (هيتينية + تكملة + ربع) are here.
# His month-variant مرد ابركسيس is handled by get_mikhael_mrd_ebrksis_id()
# and injected directly in odasSanawy which already has the per-month logic.
# ---------------------------------------------------------------------------
SAINTS_SECTIONS = {

    # -----------------------------------------------------------------------
    # 7 — رئيس الملائكة ميخائيل
    # Sourced from: odasSanawy (lines 410-413), odasEl8ytas (lines 2831-2836),
    #               odas3ydElrosol (lines 6598-6607), odasKiahk (lines 1222-1227)
    # -----------------------------------------------------------------------
    7: {
        "القداس": {
            "show": [
                '{E95B1DDC-4235-4C02-91A4-DCB7A2808C33}',  # هيتينية الملاك ميخائيل
                '{9EF543FB-A75B-4171-B358-2EB549C98411}',  # تكملة للملاك ميخائيل 2
                '{71232865-63AA-40C5-8F02-BABBFE7297D3}',  # ربع للملاك ميخائيل
            ],
            "hide": [
                '{681FF6A7-4230-4171-8F41-83FD64E8C960}',  # مرد الانجيل السنوي
            ],
        },
        "رفع بخور":      {"show": [], "hide": []},  # أضف بعد إنشاء الشرائح
        "التسبحة":       {"show": [], "hide": []},
        "تسبحة كيهك":   {"show": [], "hide": []},
        "الذكصولوجيات": {"show": [], "hide": []},
    },

    # -----------------------------------------------------------------------
    # 8 — رئيس الملائكة غبريال
    # Sourced from: odas3ydElrosol (lines 6610-6612), odasKiahk (line 1191)
    # odasSanawy Virgin+Gabriel block (line 390) stays untouched — it is
    # season-entangled (seasons 30/31 + day 21 + month 9 day 1).
    # -----------------------------------------------------------------------
    8: {
        "القداس": {
            "show": [
                '{AE7AA37A-5543-45C2-9921-1F5B7FF26544}',  # هيتينية الملاك غبريال
                '{D7004C9A-722E-4972-BC92-78D7310ECF12}',  # مرد ابركسيس كيهك 2و4 والملاك غبريال
                '{D291E41B-2C53-4536-8FD1-348E9CDB2155}',  # ربع للملاك غبريال
            ],
            "hide": [],
        },
        "رفع بخور":      {"show": [], "hide": []},
        "التسبحة":       {"show": [], "hide": []},
        "تسبحة كيهك":   {"show": [], "hide": []},
        "الذكصولوجيات": {"show": [], "hide": []},
    },
}


# ---------------------------------------------------------------------------
# MONTH-VARIANT HELPERS
# ---------------------------------------------------------------------------

def get_mikhael_mrd_ebrksis_id(coptic_month: int) -> str:
    """
    Returns the correct mrd ebrksis GUID for Michael based on Coptic month.
    Used only in odasSanawy where the per-month logic already exists.
    Sourced from odasSanawy lines 419-424.
    """
    month_variants = {
        3:  '{14A3F09D-ACCA-461F-AD67-08484F44D518}',  # هاتور
        10: '{02EBDDE5-1CBF-452A-A12A-A3F76FE68DDC}',  # برمودة
    }
    return month_variants.get(coptic_month, '{56E5BC0D-5FFC-4411-AC9C-78085E58A9E3}')


# ---------------------------------------------------------------------------
# LOOKUP FUNCTIONS
# ---------------------------------------------------------------------------

def get_active_saints(coptic_month: int, coptic_day: int) -> list:
    """
    Returns list of saint_ids active on this Coptic date.
    Called once at startup and on date change, stored in MainWindow.active_saints.
    Zero I/O — pure dict lookup.
    """
    active = []
    for saint_id, feast_days in SAINTS_CALENDAR.items():
        for (m, d) in feast_days:
            if d == coptic_day and (m == 0 or m == coptic_month):
                active.append(saint_id)
                break
    return active


def get_saint_sections(saint_id: int, file_key: str) -> dict:
    """
    Returns {"show": [...], "hide": [...]} for a saint in a specific file.
    Returns empty lists if no sections have been authored yet for that file.
    """
    if saint_id not in SAINTS_SECTIONS:
        return {"show": [], "hide": []}
    return SAINTS_SECTIONS[saint_id].get(file_key, {"show": [], "hide": []})


def get_saint_show_hide(active_saints: list, file_key: str) -> tuple:
    """
    Returns (show_guids, hide_guids) — two flat lists ready to extend
    the existing show/hide arrays before show_hide_insertImage_replaceText().

    Usage inside any odas/baker/tasbha function, right before the call:

        saint_show, saint_hide = get_saint_show_hide(active_saints, "القداس")
        my_show_full_sections.extend(saint_show)
        my_hide_full_sections.extend(saint_hide)

    Args:
        active_saints — passed from the handler via self.active_saints,
                        e.g. [7] on Michael's day, [] on a normal day
        file_key      — sheet name: "القداس", "رفع بخور", "التسبحة", etc.
    """
    all_show = []
    all_hide = []
    if not active_saints:
        return all_show, all_hide
    for saint_id in active_saints:
        data = get_saint_sections(saint_id, file_key)
        all_show.extend(data["show"])
        all_hide.extend(data["hide"])
    return all_show, all_hide
