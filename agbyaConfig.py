# Single source of truth for the Agbya (الأجبية) prayer hours; add a new hour here only.
AGBYA_PRAYER_ORDER = ["first", "third", "sixth", "ninth", "sunset", "sleep", "midnight"]

AGBYA_PRAYERS = {
    "first": {
        "label": "الأولى",
        "template": r"01- الأولى.pptx",
        "working_file": r"الأولى.pptx",
        "sheet": "صلاة الأولى",
        "sub_services": None,
        "implemented": False,
    },
    "third": {
        "label": "الثالثة",
        "template": r"03- الثالثة.pptx",
        "working_file": r"الثالثة.pptx",
        "sheet": "صلاة الثالثة",
        "sub_services": None,
        "implemented": False,
    },
    "sixth": {
        "label": "السادسة",
        "template": r"04- السادسة.pptx",
        "working_file": r"السادسة.pptx",
        "sheet": "صلاة السادسة",
        "sub_services": None,
        "implemented": False,
    },
    "ninth": {
        "label": "التاسعة",
        "template": r"05- التاسعة.pptx",
        "working_file": r"التاسعة.pptx",
        "sheet": "صلاة التاسعة",
        "sub_services": None,
        "implemented": False,
    },
    "sunset": {
        "label": "الغروب",
        "template": r"06- الغروب.pptx",
        "working_file": r"الغروب.pptx",
        "sheet": "صلاة الغروب",
        "sub_services": None,
        "implemented": False,
    },
    "sleep": {
        "label": "النوم",
        "template": r"02- النوم.pptx",
        "working_file": r"النوم.pptx",
        "sheet": "صلاة النوم",
        "sub_services": None,
        "implemented": False,
    },
    "midnight": {
        "label": "نصف الليل",
        "template": r"07- منتصف الليل.pptx",
        "working_file": r"منتصف الليل.pptx",
        "sheet": "صلاة نصف الليل",
        # (service_key, arabic_label, section-name keyword) — keyword must appear verbatim
        # in the PowerPoint section's name (via PowerPoint's own "Rename Section") for that
        # section to belong to this service; sections with none of these keywords are fixed.
        "sub_services": [
            ("service1", "الخدمة الأولى", "الخدمة الأولى"),
            ("service2", "الخدمة الثانية", "الخدمة الثانية"),
            ("service3", "الخدمة الثالثة", "الخدمة الثالثة"),
        ],
        "implemented": True,
    },
}
