# saintsCalendar.py
# Centralized saints and angels feast calendar.
# Same pattern as find_Readings_Date in commonFunctions.py — pure in-memory dicts, no I/O.
#
# HOW TO ADD A NEW SAINT:
#   1. Author the slides in the relevant .pptx template under Data\CopyData\
#   2. Run update_section_names() to extract the new section IDs into Files Data.xlsx
#   3. Add the saint_id → feast days to SAINTS_CALENDAR below
#   4. Add the saint_id → section GUIDs (per file) to SAINTS_SECTIONS below
#   5. For saints with month-variant sections (like Michael's مرد ابركسيس),
#      add a helper function following the pattern of get_mikhael_mrd_ebrksis_id()

# ---------------------------------------------------------------------------
# SAINTS CALENDAR
# Maps saint_id → list of (coptic_month, coptic_day)
# coptic_month = 0  means "every month" (for Michael, Gabriel, etc.)
# ---------------------------------------------------------------------------
SAINTS_CALENDAR = {
    7: [(0, 12)],   # رئيس الملائكة ميخائيل — كل شهر يوم ١٢
    8: [(0, 21)],   # رئيس الملائكة غبريال  — كل شهر يوم ٢١
    # أضف القديسين هنا بعد إنشاء الشرائح الخاصة بهم
    # مثال:
    # 212: [(8, 23)],          # القديس جاورجيوس
    # 305: [(4, 28), (8, 5)],  # مار مرقس — عيدان
}

# ---------------------------------------------------------------------------
# SAINTS SECTIONS
# Maps saint_id → { file_key: { "show": [GUIDs], "hide": [GUIDs] } }
#
# file_key values match the sheet names in Files Data.xlsx exactly:
#   "القداس" | "رفع بخور" | "التسبحة" | "تسبحة كيهك" | "الذكصولوجيات"
#
# GUIDs are copied from Files Data.xlsx after authoring and extracting sections.
# These are the same GUIDs used throughout odasat.py, baker.py, etc.
#
# NOTE — Michael's مرد ابركسيس varies by month and is handled separately
# via get_mikhael_mrd_ebrksis_id(). The three fixed sections below are the
# ones that appear in every month identically (هيتينية + تكملة + ربع).
# The month-variant مرد ابركسيس is injected by odasSanawy directly using
# get_mikhael_mrd_ebrksis_id() until that function is also migrated here.
# ---------------------------------------------------------------------------
SAINTS_SECTIONS = {

    # -----------------------------------------------------------------------
    # ٧ — رئيس الملائكة ميخائيل
    # Sources: odasSanawy lines 410–413, odasEl8ytas lines 2831–2836,
    #          odas3ydElrosol lines 6598–6607, odasKiahk lines 1222–1227
    # -----------------------------------------------------------------------
    7: {
        "القداس": {
            # هيتينية الملاك ميخائيل + تكملة للملاك ميخائيل 2 + ربع للملاك ميخائيل
            "show": [
                '{E95B1DDC-4235-4C02-91A4-DCB7A2808C33}',
                '{9EF543FB-A75B-4171-B358-2EB549C98411}',
                '{71232865-63AA-40C5-8F02-BABBFE7297D3}',
            ],
            # مرد الانجيل (السنوي) — يُخفى لأن ميخائيل له مرد ابركسيس خاص
            "hide": ['{681FF6A7-4230-4171-8F41-83FD64E8C960}'],
        },
        "رفع بخور":      {"show": [], "hide": []},  # أضف بعد إنشاء الشرائح
        "التسبحة":       {"show": [], "hide": []},
        "تسبحة كيهك":   {"show": [], "hide": []},
        "الذكصولوجيات": {"show": [], "hide": []},
    },

    # -----------------------------------------------------------------------
    # ٨ — رئيس الملائكة غبريال
    # Sources: odas3ydElrosol lines 6610–6612, odasKiahk line 1191
    # Note: in odasSanawy, Gabriel (day 21) is merged with the Virgin Mary
    # block (season 30/31) — that logic stays in odasSanawy for now because
    # it is season-entangled. The sections below cover all OTHER functions.
    # -----------------------------------------------------------------------
    8: {
        "القداس": {
            # هيتينية الملاك غبريال + مرد ابركسيس كيهك 2و4 و الملاك غبريال + ربع للملاك غبريال
            "show": [
                '{AE7AA37A-5543-45C2-9921-1F5B7FF26544}',
                '{D7004C9A-722E-4972-BC92-78D7310ECF12}',
                '{D291E41B-2C53-4536-8FD1-348E9CDB2155}',
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
# Used for saints whose section IDs differ depending on the Coptic month.
# ---------------------------------------------------------------------------

def get_mikhael_mrd_ebrksis_id(coptic_month: int) -> str:
    """
    Returns the correct 'مرد ابركسيس' section GUID for Michael
    based on the current Coptic month.
    Used in odasSanawy (and any other function that needs the month-variant).

    Month mapping comes from odasSanawy lines 419–424:
        هاتور  (month 3)  → {14A3F09D-...}
        برمودة (month 10) → {02EBDDE5-...}
        default           → {56E5BC0D-...}
    """
    month_variants = {
        3:  '{14A3F09D-ACCA-461F-AD67-08484F44D518}',   # هاتور
        10: '{02EBDDE5-1CBF-452A-A12A-A3F76FE68DDC}',   # برمودة
    }
    return month_variants.get(coptic_month, '{56E5BC0D-5FFC-4411-AC9C-78085E58A9E3}')


# ---------------------------------------------------------------------------
# LOOKUP FUNCTIONS
# ---------------------------------------------------------------------------

def get_active_saints(coptic_month: int, coptic_day: int) -> list:
    """
    Returns a list of saint_ids whose feast falls on this Coptic date.
    Called once at startup (and on date change) and stored in
    MainWindow.active_saints. Zero I/O — pure dict lookup.

    Args:
        coptic_month: the Coptic month (1–13)
        coptic_day:   the Coptic day (1–30)

    Returns:
        list of int saint_ids, e.g. [7] on day 12, [8] on day 21,
        [] on a day with no registered saints.
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
    Returns {"show": [...], "hide": [...]} of section GUIDs for a saint
    in a specific file.  Returns empty lists if the saint has no special
    sections for that file (i.e. no slides have been authored yet).

    Args:
        saint_id:  integer saint identifier from SAINTS_CALENDAR
        file_key:  sheet name string, e.g. "القداس", "رفع بخور"
    """
    if saint_id not in SAINTS_SECTIONS:
        return {"show": [], "hide": []}
    return SAINTS_SECTIONS[saint_id].get(file_key, {"show": [], "hide": []})
