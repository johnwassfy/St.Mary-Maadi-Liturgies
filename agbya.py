import os
from commonFunctions import (relative_path, replacefile, show_hide_insertImage_replaceText,
                              open_presentation_relative_path, classify_sections_by_keyword,
                              list_section_rows)
from agbyaConfig import AGBYA_PRAYERS

VIDEO_EXTENSIONS = {".mp4", ".wmv", ".avi", ".mov", ".m4v", ".mpg", ".mpeg"}

def open_agbya_prayer(prayer_key, selected_sub_services=None, hidden_section_ids=None,
                       bishop=False, guestBishop=0, hymn_insertions=None, media_insertions=None):
    # bishop/guestBishop are accepted for signature consistency with other service functions;
    # no bishop-specific agbya sections exist yet, so they're unused for now.
    config = AGBYA_PRAYERS[prayer_key]
    prs = relative_path(config["working_file"])
    excel = relative_path(r"Files Data.xlsx")
    sheet = config["sheet"]

    template_path = os.path.join(r"Data\CopyData\الأجبية", config["template"])
    replacefile(prs, relative_path(template_path))

    if config["sub_services"]:
        keyword_map = {service_key: keyword for service_key, _, keyword in config["sub_services"]}
        buckets = classify_sections_by_keyword(excel, sheet, keyword_map)
        selected = set(selected_sub_services or [])
        show_sections = list(buckets["fixed"])
        hide_sections = []
        for service_key in keyword_map:
            if service_key in selected:
                show_sections.extend(buckets[service_key])
            else:
                hide_sections.extend(buckets[service_key])
    else:
        show_sections = [guid for _, guid in list_section_rows(excel, sheet)]
        hide_sections = []

    if hidden_section_ids:
        hidden_set = set(hidden_section_ids)
        show_sections = [guid for guid in show_sections if guid not in hidden_set]
        hide_sections.extend(guid for guid in hidden_section_ids if guid not in hide_sections)

    show_hide_insertImage_replaceText(prs, excel, sheet, show_sections, hide_sections)

    presentation = open_presentation_relative_path(prs)
    if hymn_insertions:
        _insert_hymns(presentation, hymn_insertions)
    if media_insertions:
        _insert_media(presentation, media_insertions)

    return presentation

def _get_blank_layout(presentation):
    """Return the "Blank" CustomLayout under the destination's own "2_Default Design", or None."""
    designs = presentation.Designs
    for i in range(1, designs.Count + 1):
        design = designs.Item(i)
        if design.Name == "2_Default Design":
            layouts = design.SlideMaster.CustomLayouts
            for j in range(1, layouts.Count + 1):
                layout = layouts.Item(j)
                if layout.Name.strip().lower() == "blank":
                    return layout
    return None

def _group_by_anchor(insertions):
    """Group a flat list of {"after_guid": ...} dicts by anchor section, preserving first-seen order."""
    grouped = {}
    anchor_order = []
    for insertion in insertions:
        anchor_guid = insertion["after_guid"]
        if anchor_guid not in grouped:
            grouped[anchor_guid] = []
            anchor_order.append(anchor_guid)
        grouped[anchor_guid].append(insertion)
    return grouped, anchor_order

def _section_index_by_guid(presentation, anchor_guid):
    sections = presentation.SectionProperties
    for i in range(1, sections.Count + 1):
        if sections.SectionID(i) == anchor_guid:
            return i
    return None

def _insert_pptx_range(presentation, source_path, cursor, first_slide, last_slide):
    """Insert slides [first_slide, last_slide] from source_path right after `cursor`, preceded by a
    native spacer slide whose layout is explicitly re-pointed at the destination's own "2_Default
    Design" > Blank layout — see _insert_hymns' docstring for why the spacer is needed. Returns
    (start_slide, inserted_count)."""
    spacer = presentation.Slides.Add(cursor + 1, 12)  # 12 = ppLayoutBlank
    spacer.SlideShowTransition.Hidden = True
    blank_layout = _get_blank_layout(presentation)
    if blank_layout is not None:
        spacer.CustomLayout = blank_layout
    cursor += 1

    start_slide = cursor + 1
    inserted_count = presentation.Slides.InsertFromFile(source_path, cursor, first_slide, last_slide)
    return start_slide, inserted_count

def _insert_hymns(presentation, hymn_insertions):
    """Silent version: InsertFromFile reads the hymn book straight off disk (no Copy/Paste, no
    clipboard, no window activation/flicker). A pasted/inserted slide inherits its layout from
    whatever slide already sits at the insertion point — not from the source — which is why the hymn
    kept picking up the destination's own master ("2_Default Design"). Fix: add a native spacer slide
    at the target position, hide it, then explicitly switch its layout to the destination's own
    "2_Default Design" > Blank layout (a separate post-creation step, not the layout picked at Add()
    time), so the slide sitting right before the real insertion is on a known-good destination master
    before InsertFromFile brings in the hymn's own slides right after it. Each hymn also gets forced
    visible (in case it was hidden in the hymn book) and wrapped in its own PowerPoint section."""
    if not hymn_insertions:
        return

    source_path = relative_path(r"Data\CopyData\كتاب المدائح.pptx")
    grouped, anchor_order = _group_by_anchor(hymn_insertions)

    sections = presentation.SectionProperties
    for anchor_guid in anchor_order:
        section_index = _section_index_by_guid(presentation, anchor_guid)
        if section_index is None:
            continue

        # Re-derived live from COM each group, so earlier insertions already shift this correctly.
        cursor = sections.FirstSlide(section_index) + sections.SlidesCount(section_index) - 1
        for insertion in grouped[anchor_guid]:
            hymn_start_slide, inserted_count = _insert_pptx_range(
                presentation, source_path, cursor, insertion["first_slide"], insertion["last_slide"]
            )
            if inserted_count != insertion["num_slides"]:
                raise RuntimeError(
                    f"Expected {insertion['num_slides']} slides for hymn '{insertion['name']}', inserted {inserted_count}"
                )
            cursor = hymn_start_slide + inserted_count - 1

            # Some hymn slides are marked hidden in كتاب المدائح.pptx itself — always show them here.
            for slide_index in range(hymn_start_slide, hymn_start_slide + inserted_count):
                presentation.Slides(slide_index).SlideShowTransition.Hidden = False

            presentation.SectionProperties.AddBeforeSlide(hymn_start_slide, insertion["name"])

    presentation.Save()

def _get_pptx_slide_count(presentation, path):
    """python-pptx only understands the modern .pptx (zip/OOXML) format — a legacy binary .ppt file
    makes it raise PackageNotFoundError, so for that extension ask PowerPoint itself via COM instead
    (reusing the destination's own Application instance rather than spawning a new one)."""
    if os.path.splitext(path)[1].lower() == ".pptx":
        from pptx import Presentation as PptxPresentation
        return len(PptxPresentation(path).slides)

    app = presentation.Application
    # Match the Open() call style used everywhere else in this codebase (just WithWindow) — combining
    # it with ReadOnly=True as a second keyword arg raised a generic COM exception via dynamic dispatch.
    source = app.Presentations.Open(path, WithWindow=False)
    try:
        return source.Slides.Count
    finally:
        source.Close()

def _insert_media(presentation, media_insertions):
    """User-picked, session-only additions to an Agbya gathering: either a full external .pptx/.ppt
    (all of its slides inserted the same way _insert_hymns does, via the spacer-slide fix) or a video
    file (a single new blank slide with the video shape stretched to fill it, click-to-play —
    PowerPoint's own default behavior for a freshly inserted media shape with no extra
    animation/timeline set)."""
    if not media_insertions:
        return

    grouped, anchor_order = _group_by_anchor(media_insertions)
    sections = presentation.SectionProperties
    slide_width = presentation.PageSetup.SlideWidth
    slide_height = presentation.PageSetup.SlideHeight

    for anchor_guid in anchor_order:
        section_index = _section_index_by_guid(presentation, anchor_guid)
        if section_index is None:
            continue

        cursor = sections.FirstSlide(section_index) + sections.SlidesCount(section_index) - 1
        for insertion in grouped[anchor_guid]:
            # Defensive: normalize again here too (the UI already does this), since a forward-slash
            # path made PowerPoint's COM Open/AddMediaObject2 fail with a generic "file not found".
            media_path = os.path.normpath(insertion["path"])
            if insertion["media_kind"] == "pptx":
                total_slides = _get_pptx_slide_count(presentation, media_path)
                start_slide, inserted_count = _insert_pptx_range(
                    presentation, media_path, cursor, 1, total_slides
                )
                for slide_index in range(start_slide, start_slide + inserted_count):
                    presentation.Slides(slide_index).SlideShowTransition.Hidden = False
                presentation.SectionProperties.AddBeforeSlide(start_slide, insertion["name"])
                cursor = start_slide + inserted_count - 1
            else:
                new_slide = presentation.Slides.Add(cursor + 1, 12)  # 12 = ppLayoutBlank
                new_slide.Shapes.AddMediaObject2(
                    media_path, LinkToFile=False, SaveWithDocument=True,
                    Left=0, Top=0, Width=slide_width, Height=slide_height
                )
                presentation.SectionProperties.AddBeforeSlide(cursor + 1, insertion["name"])
                cursor += 1

    presentation.Save()




