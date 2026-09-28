"""PowerPoint text utilities.

Before running this file in VS Code, edit WORKFLOW and the default paths in
the configuration section below. Then use "Run Python File" normally.

Available workflows:
    chapters      Replace Arabic chapter names throughout a presentation.
    superscript   Convert numbers in one named PowerPoint object.
    both          Run both workflows in sequence.
    merge_chapters
                  Combine consecutive slides with the same chapter title.
    all           Replace chapters, merge chapters, then superscript numbers.
    all_files     Run the complete workflow for every presentation in a folder.
    split_rendered
                  Split each rendered text box into groups of four lines.
    split_rendered_files
                  Run the rendered-line split for every presentation in a folder.
    phase3_files
                  Add chapter titles and sections to every Phase 2 presentation.
    phase3_all_files
                  Run Phase 3 for every New and Old Testament presentation.
    format_phase3_layouts
                  Format the Title + Text layout and slides in all Phase 3 presentations.
    merge_all_books
                  Insert all Phase 3 books into a copy of the master Bible presentation.
"""

import re
import os
import shutil
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn


# ---------------------------------------------------------------------------
# Defaults
# ---------------------------------------------------------------------------

DEFAULT_INPUT_FILE = Path(r"الكتاب المقدس\عهد جديد\01 انجيل متي.pptx")
DEFAULT_OUTPUT_FILE = Path(r"الكتاب المقدس\عهد جديد\phase2\01 انجيل متي.pptx")
DEFAULT_INPUT_FOLDER = Path(r"الكتاب المقدس\عهد جديد")
DEFAULT_OUTPUT_FOLDER = DEFAULT_INPUT_FOLDER / "processed"
DEFAULT_PHASE2_OUTPUT_FOLDER = DEFAULT_INPUT_FOLDER / "phase2"
DEFAULT_PHASE3_OUTPUT_FOLDER = DEFAULT_INPUT_FOLDER / "phase3"
DEFAULT_PHASE3_MASTER_BACKGROUND = Path(r"Data\Designs\Picture1.png")
DEFAULT_PHASE3_VERSE_FONT_NAME = "Times New Roman"
DEFAULT_PHASE3_HEADER_FONT_NAME = "Times New Roman"
DEFAULT_PHASE3_HEADER_FONT_SIZE = 34
DEFAULT_PHASE3_VERSE_FONT_SIZE = 44
DEFAULT_MASTER_PRESENTATION = Path(r"Data\CopyData\الكتاب المقدس.pptx")
DEFAULT_FULL_BIBLE_OUTPUT = Path(r"Data\CopyData\FullBible.pptx")
DEFAULT_OBJECT_NAME = "Text Placeholder 2"
TITLE_OBJECT_NAME = "Title 1"
WORKFLOW = "format_phase3_layouts"  # Choose a workflow listed above.
CHAPTER_HEAD_FONT_SIZE = 34
CHAPTER_HEAD_FONT_NAME = "Times New Roman"
CHAPTER_HEAD_COLOR_RGB = 49407  # #FFC000 in Office RGB integer format.
SUPERSCRIPT_OFFSET = "40000"  # 40% baseline offset in PowerPoint OOXML.
TEXTBOX_LEFT_CM = 0
TEXTBOX_TOP_NO_SUBTITLE_CM = 12.18
TEXTBOX_TOP_WITH_SUBTITLE_CM = 10.25
LINES_WITHOUT_SUBTITLE = 4
LINES_WITH_SUBTITLE = 5  # 1 subtitle line + 4 verse lines
SUBTITLE_SPACE_AFTER_POINTS = 6
VERSE_FONT_SIZE_POINTS = 44
LINE_SPACING_MULTIPLE = 0.9

CHAPTERS = {
    "الإصحاح الأول": "الإصحاح الـ1",
    "الإصحاح الثاني": "الإصحاح الـ2",
    "الإصحاح الثالث": "الإصحاح الـ3",
    "الإصحاح الرابع": "الإصحاح الـ4",
    "الإصحاح الخامس": "الإصحاح الـ5",
    "الإصحاح السادس": "الإصحاح الـ6",
    "الإصحاح السابع": "الإصحاح الـ7",
    "الإصحاح الثامن": "الإصحاح الـ8",
    "الإصحاح التاسع": "الإصحاح الـ9",
    "الإصحاح العاشر": "الإصحاح الـ10",
}
RESTORE_CHAPTERS = {new: old for old, new in CHAPTERS.items()}

ARABIC_DIGITS = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")

BOOK_SHORTCUTS = {
    "عهد جديد": {
        1: "مت", 2: "مر", 3: "لو", 4: "يو", 5: "أع", 6: "رو",
        7: "1كو", 8: "2كو", 9: "غل", 10: "أف", 11: "في", 12: "كو",
        13: "1تس", 14: "2تس", 15: "1تي", 16: "2تي", 17: "تي",
        18: "فل", 19: "عب", 20: "يع", 21: "1بط", 22: "2بط",
        23: "1يو", 24: "2يو", 25: "3يو", 26: "يه", 27: "رؤ",
    },
    "عهد قديم": {
        1: "تك", 2: "خر", 3: "لا", 4: "عد", 5: "تث", 6: "يش",
        7: "قض", 8: "را", 9: "1صم", 10: "2صم", 11: "1مل", 12: "2مل",
        13: "1أخ", 14: "2أخ", 15: "عز", 16: "نح", 17: "طو", 18: "يه",
        19: "أس", 20: "أي", 21: "مز", 22: "أم", 23: "جا", 24: "نش",
        25: "حك", 26: "سي", 27: "إش", 28: "إر", 29: "مرا", 30: "با",
        31: "حز", 32: "دا", 33: "هو", 34: "يوئ", 35: "عا", 36: "عو",
        37: "يون", 38: "مي", 39: "نا", 40: "حب", 41: "صف", 42: "حج",
        43: "زك", 44: "ملا", 45: "1مك", 46: "2مك",
    },
}

BOOK_SECTION_NAMES = {
    "عهد جديد": {
        1: "متى", 2: "مرقس", 3: "لوقا", 4: "يوحنا", 5: "أعمال الرسل",
        6: "رومية", 7: "كورنثوس الأولى", 8: "كورنثوس الثانية", 9: "غلاطية",
        10: "أفسس", 11: "فيلبي", 12: "كولوسي", 13: "تسالونيكي الأولى",
        14: "تسالونيكي الثانية", 15: "تيموثاوس الأولى", 16: "تيموثاوس الثانية",
        17: "تيطس", 18: "فليمون", 19: "العبرانيين", 20: "يعقوب",
        21: "بطرس الأولى", 22: "بطرس الثانية", 23: "يوحنا الأولى",
        24: "يوحنا الثانية", 25: "يوحنا الثالثة", 26: "يهوذا", 27: "الرؤيا",
    },
    "عهد قديم": {
        1: "تكوين", 2: "خروج", 3: "لاويين", 4: "عدد", 5: "تثنية",
        6: "يشوع", 7: "قضاة", 8: "راعوث", 9: "صموئيل الأول",
        10: "صموئيل الثاني", 11: "ملوك الأول", 12: "ملوك الثاني",
        13: "أخبار الأيام الأول", 14: "أخبار الأيام الثاني", 15: "عزرا",
        16: "نحميا", 17: "طوبيا", 18: "يهوديت", 19: "أستير", 20: "أيوب",
        21: "المزامير", 22: "الأمثال", 23: "الجامعة", 24: "نشيد الأنشاد",
        25: "الحكمة", 26: "يشوع بن سيراخ", 27: "إشعياء", 28: "إرميا",
        29: "مراثي إرميا", 30: "باروخ", 31: "حزقيال", 32: "دانيال",
        33: "هوشع", 34: "يوئيل", 35: "عاموس", 36: "عوبديا", 37: "يونان",
        38: "ميخا", 39: "ناحوم", 40: "حبقوق", 41: "صفنيا", 42: "حجي",
        43: "زكريا", 44: "ملاخي", 45: "المكابيين الأول", 46: "المكابيين الثاني",
    },
}


# ---------------------------------------------------------------------------
# Shared PowerPoint traversal
# ---------------------------------------------------------------------------

def iter_text_frames(shape):
    """Yield text frames in a shape, including tables and groups."""
    if getattr(shape, "has_text_frame", False):
        yield shape.text_frame

    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                yield cell.text_frame

    if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
        for child in shape.shapes:
            yield from iter_text_frames(child)


def iter_shapes(shape_collection):
    """Yield shapes recursively, including shapes inside groups."""
    for shape in shape_collection:
        yield shape
        if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
            yield from iter_shapes(shape.shapes)


def find_named_shape(slide, object_name):
    """Return the first shape with the requested name on a slide."""
    return next(
        (shape for shape in iter_shapes(slide.shapes) if shape.name == object_name),
        None,
    )


# ---------------------------------------------------------------------------
# Chapter replacement workflow
# ---------------------------------------------------------------------------

def replace_chapter_in_paragraph(paragraph, replacements=CHAPTERS):
    """Replace one chapter phrase, even when it spans multiple runs."""
    runs = list(paragraph.runs)
    full_text = "".join(run.text for run in runs)

    if not full_text:
        return False

    old_text, new_text = next(
        (
            (old, new)
            for old, new in sorted(
                replacements.items(),
                key=lambda item: len(item[0]),
                reverse=True,
            )
            if old in full_text
        ),
        (None, None),
    )
    if old_text is None:
        return False

    start = full_text.index(old_text)
    end = start + len(old_text)
    run_positions = []
    current_position = 0

    for index, run in enumerate(runs):
        run_start = current_position
        run_end = run_start + len(run.text)
        if run_end > start and run_start < end:
            run_positions.append(index)
        current_position = run_end

    if not run_positions:
        return False

    first_index = run_positions[0]
    last_index = run_positions[-1]
    first_run = runs[first_index]
    last_run = runs[last_index]
    first_run_start = sum(len(run.text) for run in runs[:first_index])
    last_run_start = sum(len(run.text) for run in runs[:last_index])
    prefix = first_run.text[:start - first_run_start]
    suffix = last_run.text[end - last_run_start:] if end <= last_run_start + len(last_run.text) else ""

    first_run.text = prefix + new_text
    for index in range(first_index + 1, last_index + 1):
        runs[index].text = ""

    if first_index == last_index:
        first_run.text += suffix
    else:
        last_run.text = suffix

    return True


def replace_chapters_in_frame(text_frame, replacements=CHAPTERS):
    """Replace chapter phrases in every paragraph of a text frame."""
    return sum(
        replace_chapter_in_paragraph(paragraph, replacements)
        for paragraph in text_frame.paragraphs
    )


def replace_chapters(input_file, output_file):
    """Replace chapter phrases throughout every slide and save the result."""
    presentation = Presentation(input_file)
    changed_count = 0

    for slide in presentation.slides:
        for shape in slide.shapes:
            for text_frame in iter_text_frames(shape):
                changed_count += replace_chapters_in_frame(text_frame)

    presentation.save(output_file)
    print(f"Chapter replacements: {changed_count}")
    print(f"Output: {output_file}")


def restore_chapters(input_file, output_file):
    """Restore the original chapter names throughout a presentation."""
    presentation = Presentation(input_file)
    changed_count = 0

    for slide in presentation.slides:
        for shape in slide.shapes:
            for text_frame in iter_text_frames(shape):
                changed_count += replace_chapters_in_frame(
                    text_frame,
                    RESTORE_CHAPTERS,
                )

    presentation.save(output_file)
    print(f"Chapter names restored: {changed_count}")
    print(f"Output: {output_file}")


# ---------------------------------------------------------------------------
# Superscript workflow
# ---------------------------------------------------------------------------

def to_arabic_number(number_text):
    """Convert Western digits to Arabic-Indic digits."""
    return number_text.translate(ARABIC_DIGITS)


def set_superscript(run):
    """Set a PowerPoint run to superscript with the configured offset."""
    run._r.get_or_add_rPr().set("baseline", SUPERSCRIPT_OFFSET)


def copy_run_format(source_run, target_run):
    """Copy run-level formatting from one run to another."""
    source_properties = source_run._r.find(qn("a:rPr"))
    if source_properties is None:
        return

    target_properties = target_run._r.get_or_add_rPr()
    for child in list(target_properties):
        target_properties.remove(child)
    for attribute, value in source_properties.attrib.items():
        target_properties.set(attribute, value)
    for child in source_properties:
        target_properties.append(deepcopy(child))


def superscript_run(run):
    """Split a run so only integers from 1 through 100 are superscripted."""
    matches = [
        match for match in re.finditer(r"\d+", run.text)
        if 1 <= int(match.group()) <= 100
    ]
    if not matches:
        return False

    original_element = run._r
    parent = original_element.getparent()
    position = 0

    for match in matches:
        if match.start() > position:
            new_run = run._parent.add_run()
            new_run.text = run.text[position:match.start()]
            copy_run_format(run, new_run)
            parent.insert(parent.index(original_element), new_run._r)

        number_run = run._parent.add_run()
        number_run.text = to_arabic_number(match.group())
        copy_run_format(run, number_run)
        set_superscript(number_run)
        parent.insert(parent.index(original_element), number_run._r)
        position = match.end()

    if position < len(run.text):
        new_run = run._parent.add_run()
        new_run.text = run.text[position:]
        copy_run_format(run, new_run)
        parent.insert(parent.index(original_element), new_run._r)

    parent.remove(original_element)
    return True


def superscript_frame(text_frame):
    """Superscript eligible numbers in every paragraph of a text frame."""
    changed_count = 0
    for paragraph in text_frame.paragraphs:
        for run in list(paragraph.runs):
            changed_count += superscript_run(run)
    return changed_count


def superscript_object(input_file, output_file, object_name):
    """Process matching named objects and save the result."""
    presentation = Presentation(input_file)
    object_count = 0
    changed_count = 0

    for slide_number, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.name != object_name:
                continue

            print(f"Found '{object_name}' on slide {slide_number}")
            object_count += 1
            for text_frame in iter_text_frames(shape):
                changed_count += superscript_frame(text_frame)

    presentation.save(output_file)
    print(f"Objects processed: {object_count}")
    print(f"Superscripted runs: {changed_count}")
    print(f"Output: {output_file}")


# ---------------------------------------------------------------------------
# Merge chapters workflow
# ---------------------------------------------------------------------------

def get_shape_text(shape):
    """Return the visible text from a named text shape."""
    if shape is None or not getattr(shape, "has_text_frame", False):
        return ""
    return shape.text_frame.text.strip()


def append_text_frame_content(target_frame, source_frame):
    """Append source paragraphs while retaining their PowerPoint formatting."""
    target_body = target_frame._txBody
    source_body = source_frame._txBody

    for paragraph in source_body.iterchildren(qn("a:p")):
        target_body.append(deepcopy(paragraph))


def clean_paragraph(paragraph):
    """Remove empty line breaks while preserving real lines and formatting."""
    text_nodes = list(paragraph.iter(qn("a:t")))
    if not text_nodes:
        return False, 0

    for text_node in text_nodes:
        text_node.text = re.sub(
            r"[ \t\u00a0]{2,}",
            " ",
            text_node.text or "",
        )

    first_text_node = next((node for node in text_nodes if node.text), None)
    last_text_node = next((node for node in reversed(text_nodes) if node.text), None)
    if first_text_node is None or last_text_node is None:
        return False, 0

    first_text_node.text = first_text_node.text.lstrip()
    last_text_node.text = last_text_node.text.rstrip()

    if not "".join(node.text or "" for node in text_nodes).strip():
        return False, 0

    # Shift+Enter creates a direct a:br child instead of a new paragraph.
    # Remove breaks at the paragraph edges and repeated breaks, but preserve
    # one break between two pieces of actual text, such as a subtitle line.
    text_children = [
        child for child in paragraph
        if child.tag == qn("a:r") and "".join(
            node.text or "" for node in child.iter(qn("a:t"))
        ).strip()
    ]
    removed_breaks = 0
    saw_content = False
    previous_was_break = False
    for child in list(paragraph):
        if child.tag == qn("a:pPr") or child.tag == qn("a:endParaRPr"):
            continue

        if child.tag == qn("a:br"):
            remaining_text_children = [
                text_child for text_child in text_children
                if paragraph.index(text_child) > paragraph.index(child)
            ]
            if not saw_content or not remaining_text_children or previous_was_break:
                paragraph.remove(child)
                removed_breaks += 1
            else:
                previous_was_break = True
            continue

        if child.tag == qn("a:r"):
            if child in text_children:
                saw_content = True
                previous_was_break = False

    return True, removed_breaks


def paragraph_is_white(paragraph):
    """Return whether a paragraph uses only white or inherited text color."""
    for run in paragraph.findall(qn("a:r")):
        for color in run.findall(f".//{qn('a:srgbClr')}"):
            if color.get("val", "").upper() not in {"FFFFFF", "FFF"}:
                return False
        for color in run.findall(f".//{qn('a:schemeClr')}"):
            if color.get("val", "").lower() not in {"lt1", "bg1"}:
                return False
    return True


def join_paragraphs_with_space(first_paragraph, second_paragraph):
    """Join two paragraphs while retaining all source run formatting."""
    last_text_node = next(
        (
            node
            for node in reversed(list(first_paragraph.iter(qn("a:t"))))
            if node.text
        ),
        None,
    )
    if last_text_node is None:
        return

    last_text_node.text = f"{last_text_node.text} "
    insert_index = next(
        (
            index for index, child in enumerate(first_paragraph)
            if child.tag == qn("a:endParaRPr")
        ),
        len(first_paragraph),
    )

    for child in second_paragraph:
        if child.tag in {qn("a:pPr"), qn("a:endParaRPr")}:
            continue
        first_paragraph.insert(insert_index, deepcopy(child))
        insert_index += 1


def clean_text_frame(text_frame):
    """Remove empty lines and tidy spaces while preserving non-empty paragraphs."""
    body = text_frame._txBody
    removed_lines = 0
    joined_lines = 0
    previous_paragraph = None

    for paragraph in list(body.iterchildren(qn("a:p"))):
        has_content, removed_breaks = clean_paragraph(paragraph)
        removed_lines += removed_breaks
        if not has_content:
            body.remove(paragraph)
            removed_lines += 1
            continue

        if (
            previous_paragraph is not None
            and paragraph_is_white(previous_paragraph)
            and paragraph_is_white(paragraph)
        ):
            join_paragraphs_with_space(previous_paragraph, paragraph)
            body.remove(paragraph)
            joined_lines += 1
        else:
            previous_paragraph = paragraph

    return removed_lines, joined_lines


def remove_slide(presentation, slide_index):
    """Remove a slide and its relationship from a presentation."""
    slide_id = presentation.slides._sldIdLst[slide_index]
    presentation.part.drop_rel(slide_id.rId)
    presentation.slides._sldIdLst.remove(slide_id)


def merge_chapters(input_file, output_file, title_name, text_name):
    """Merge consecutive same-title slides into the first slide of each group."""
    presentation = Presentation(input_file)
    slides = list(presentation.slides)
    merged_slide_count = 0
    chapter_count = 0
    empty_lines_removed = 0
    white_lines_joined = 0
    slide_index = 0

    while slide_index < len(slides):
        title_shape = find_named_shape(slides[slide_index], title_name)
        title_text = get_shape_text(title_shape)
        text_shape = find_named_shape(slides[slide_index], text_name)
        group_end = slide_index + 1

        if not title_text or text_shape is None:
            slide_index += 1
            continue

        target_text_frame = text_shape.text_frame
        if title_text and text_shape is not None:
            while group_end < len(slides):
                next_title = get_shape_text(
                    find_named_shape(slides[group_end], title_name)
                )
                if next_title != title_text:
                    break

                next_text_shape = find_named_shape(slides[group_end], text_name)
                if text_shape is not None and next_text_shape is not None:
                    append_text_frame_content(
                        text_shape.text_frame,
                        next_text_shape.text_frame,
                    )
                group_end += 1

            if group_end > slide_index + 1:
                chapter_count += 1
                merged_slide_count += group_end - slide_index - 1

                for remove_index in range(group_end - 1, slide_index, -1):
                    remove_slide(presentation, remove_index)

                slides = list(presentation.slides)

            removed_lines, joined_lines = clean_text_frame(target_text_frame)
            empty_lines_removed += removed_lines
            white_lines_joined += joined_lines

        slide_index += 1

    presentation.save(output_file)
    print(f"Chapters merged: {chapter_count}")
    print(f"Slides removed: {merged_slide_count}")
    print(f"Empty lines removed: {empty_lines_removed}")
    print(f"White lines joined with spaces: {white_lines_joined}")
    print(f"Output: {output_file}")


def process_all(input_file, output_file):
    """Run all three stages and clean up temporary files after success."""
    chapters_file = output_file.with_name(
        f"{output_file.stem}_chapters{output_file.suffix}"
    )
    merged_file = output_file.with_name(
        f"{output_file.stem}_merged{output_file.suffix}"
    )

    replace_chapters(input_file, chapters_file)
    merge_chapters(
        chapters_file,
        merged_file,
        TITLE_OBJECT_NAME,
        DEFAULT_OBJECT_NAME,
    )
    superscript_object(
        merged_file,
        output_file,
        DEFAULT_OBJECT_NAME,
    )
    restore_chapters(output_file, output_file)
    chapters_file.unlink(missing_ok=True)
    merged_file.unlink(missing_ok=True)
    print("Intermediate files removed.")


def process_all_files(input_folder, output_folder):
    """Process every PPTX in a folder into a separate output folder."""
    output_folder.mkdir(parents=True, exist_ok=True)
    input_files = sorted(input_folder.glob("*.pptx"))

    if not input_files:
        raise FileNotFoundError(f"No PowerPoint files found in {input_folder}")

    print(f"Processing {len(input_files)} presentations...")
    for index, input_file in enumerate(input_files, start=1):
        output_file = output_folder / input_file.name
        print(f"[{index}/{len(input_files)}] {input_file.name}")
        process_all(input_file, output_file)


def append_soft_break(text_range):
    """Add a PowerPoint Shift+Enter line break after the slide text."""
    if not text_range.Text.endswith("\v"):
        text_range.InsertAfter("\v")


def is_line_subtitle(line_range):
    """Check if a line uses subtitle coloring (not default white)."""
    try:
        return line_range.Font.Color.RGB != 16777215
    except Exception:
        return False


def apply_rendered_slide_format(text_shape, has_subtitle=False):
    """Apply the liturgy paragraph layout and position based on subtitle presence."""
    points_per_cm = 28.3464567
    text_range = text_shape.TextFrame.TextRange
    text_range.Font.Size = VERSE_FONT_SIZE_POINTS
    text_shape.Left = TEXTBOX_LEFT_CM * points_per_cm

    if has_subtitle:
        text_shape.Top = TEXTBOX_TOP_WITH_SUBTITLE_CM * points_per_cm
        paras_count = text_range.Paragraphs().Count
        if paras_count >= 1:
            sub_para = text_range.Paragraphs(1)
            try:
                sub_para.Font.Underline = -1  # msoTrue
            except Exception:
                pass
            sub_format = sub_para.ParagraphFormat
            for prop, val in (
                ("SpaceBefore", 0),
                ("SpaceAfter", SUBTITLE_SPACE_AFTER_POINTS),
                ("SpaceWithin", LINE_SPACING_MULTIPLE),
                ("Alignment", 2),  # ppAlignCenter
            ):
                try:
                    setattr(sub_format, prop, val)
                except Exception:
                    pass

            for p in range(2, paras_count + 1):
                verse_para = text_range.Paragraphs(p)
                try:
                    verse_para.Font.Underline = 0  # msoFalse
                except Exception:
                    pass
                v_format = verse_para.ParagraphFormat
                for prop, val in (
                    ("SpaceBefore", 0),
                    ("SpaceAfter", 0),
                    ("SpaceWithin", LINE_SPACING_MULTIPLE),
                    ("Alignment", 7),  # ppAlignJustifyLow
                ):
                    try:
                        setattr(v_format, prop, val)
                    except Exception:
                        pass
    else:
        text_shape.Top = TEXTBOX_TOP_NO_SUBTITLE_CM * points_per_cm
        paras_count = text_range.Paragraphs().Count
        for p in range(1, paras_count + 1):
            verse_para = text_range.Paragraphs(p)
            try:
                verse_para.Font.Underline = 0  # msoFalse
            except Exception:
                pass
            v_format = verse_para.ParagraphFormat
            for prop, val in (
                ("SpaceBefore", 0),
                ("SpaceAfter", 0),
                ("SpaceWithin", LINE_SPACING_MULTIPLE),
                ("Alignment", 7),  # ppAlignJustifyLow
            ):
                try:
                    setattr(v_format, prop, val)
                except Exception:
                    pass

    try:
        text_shape.TextFrame2.TextRange.ParagraphFormat.TextDirection = 2
    except Exception:
        pass

    append_soft_break(text_range)


def format_title_text_layout(presentation):
    """Format the Title + Text layout and its existing slides consistently."""
    points_per_cm = 72 / 2.54
    subtitle_color = 49407
    white_color = 16777215
    formatted_slides = 0
    found_layout = False
    slide_width = presentation.PageSetup.SlideWidth
    slide_height = presentation.PageSetup.SlideHeight
    title_left = -145
    title_top = 0
    title_width = 145
    title_height = 66
    header_left = 0
    header_top = 0
    header_width = slide_width
    header_height = 66
    body_left = 0
    body_width = slide_width
    body_top = TEXTBOX_TOP_NO_SUBTITLE_CM * points_per_cm
    body_height = max(0, slide_height - body_top)

    for design_index in range(1, presentation.Designs.Count + 1):
        layouts = presentation.Designs(design_index).SlideMaster.CustomLayouts
        for layout_index in range(1, layouts.Count + 1):
            layout = layouts(layout_index)
            if layout.Name != "Title + Text":
                continue
            found_layout = True
            title_placeholder = None
            body_placeholder = None
            header_placeholder = None
            header_placeholder_index = 0
            placeholder_index = 0
            for shape_index in range(1, layout.Shapes.Count + 1):
                shape = layout.Shapes(shape_index)
                if shape.Type == 14:
                    placeholder_index += 1
                if shape.Name == "Title 1":
                    title_placeholder = shape
                elif shape.Name == "Text Placeholder 3":
                    body_placeholder = shape
                elif shape.Name == "Text Placeholder 4":
                    header_placeholder = shape
                    header_placeholder_index = placeholder_index
                try:
                    placeholder_type = shape.PlaceholderFormat.Type
                except Exception:
                    continue
                if title_placeholder is None and placeholder_type == 1:
                    title_placeholder = shape
                elif body_placeholder is None and placeholder_type == 2:
                    body_placeholder = shape

            if title_placeholder is None or body_placeholder is None:
                raise ValueError(
                    f"Layout {layout.Name!r} must contain title and body placeholders"
                )
            if header_placeholder is None:
                header_placeholder = layout.Shapes.AddPlaceholder(
                    2,
                    header_left,
                    header_top,
                    header_width,
                    header_height,
                )
                header_placeholder.Name = "Text Placeholder 4"
                header_placeholder_index = sum(
                    1
                    for shape_index in range(1, layout.Shapes.Count + 1)
                    if layout.Shapes(shape_index).Type == 14
                    and layout.Shapes(shape_index).Name == "Text Placeholder 4"
                )

            title_placeholder.Left = title_left
            title_placeholder.Top = title_top
            title_placeholder.Width = title_width
            title_placeholder.Height = title_height
            header_placeholder.Left = header_left
            header_placeholder.Top = header_top
            header_placeholder.Width = header_width
            header_placeholder.Height = header_height
            body_placeholder.Left = body_left
            body_placeholder.Top = body_top
            body_placeholder.Width = body_width
            body_placeholder.Height = body_height

            title_range = title_placeholder.TextFrame.TextRange
            title_range.Font.Name = DEFAULT_PHASE3_HEADER_FONT_NAME
            title_range.Font.Size = DEFAULT_PHASE3_HEADER_FONT_SIZE
            title_range.Font.Color.RGB = CHAPTER_HEAD_COLOR_RGB
            title_range.ParagraphFormat.Alignment = 2
            title_range.ParagraphFormat.SpaceBefore = 0
            title_range.ParagraphFormat.SpaceAfter = 0
            title_range.ParagraphFormat.SpaceWithin = LINE_SPACING_MULTIPLE
            title_placeholder.TextFrame2.TextRange.ParagraphFormat.TextDirection = 2

            header_range = header_placeholder.TextFrame.TextRange
            header_range.Font.Name = DEFAULT_PHASE3_HEADER_FONT_NAME
            header_range.Font.Size = DEFAULT_PHASE3_HEADER_FONT_SIZE
            header_range.Font.Color.RGB = CHAPTER_HEAD_COLOR_RGB
            header_range.ParagraphFormat.Alignment = 2
            header_range.ParagraphFormat.SpaceBefore = 0
            header_range.ParagraphFormat.SpaceAfter = 0
            header_range.ParagraphFormat.SpaceWithin = LINE_SPACING_MULTIPLE
            header_placeholder.TextFrame2.TextRange.ParagraphFormat.TextDirection = 2
            for paragraph_index in range(1, header_range.Paragraphs().Count + 1):
                header_range.Paragraphs(paragraph_index).ParagraphFormat.Bullet.Visible = 0

            body_range = body_placeholder.TextFrame.TextRange
            body_range.Font.Name = DEFAULT_PHASE3_VERSE_FONT_NAME
            body_range.Font.Size = DEFAULT_PHASE3_VERSE_FONT_SIZE
            body_range.ParagraphFormat.Alignment = 7
            body_range.ParagraphFormat.SpaceBefore = 0
            body_range.ParagraphFormat.SpaceAfter = 0
            body_range.ParagraphFormat.SpaceWithin = LINE_SPACING_MULTIPLE
            body_placeholder.TextFrame2.TextRange.ParagraphFormat.TextDirection = 2

            for slide_index in range(1, presentation.Slides.Count + 1):
                slide = presentation.Slides(slide_index)
                chapter_header = next(
                    (
                        slide.Shapes(shape_index)
                        for shape_index in range(1, slide.Shapes.Count + 1)
                        if slide.Shapes(shape_index).Name in ("Chapter Header", "Text Placeholder 4")
                    ),
                    None,
                )
                chapter_text = (
                    chapter_header.TextFrame.TextRange.Text
                    if chapter_header is not None and chapter_header.HasTextFrame
                    else ""
                )
                if chapter_header is not None:
                    chapter_header.Delete()

                slide.CustomLayout = layout
                if slide.CustomLayout.Index != layout_index:
                    continue

                if header_placeholder_index > slide.Shapes.Placeholders.Count:
                    raise ValueError(
                        f"Slide {slide_index} did not receive the layout header placeholder"
                    )
                chapter_header = slide.Shapes.Placeholders(header_placeholder_index)
                chapter_header.TextFrame.TextRange.Text = chapter_text.strip()
                header_range = chapter_header.TextFrame.TextRange
                header_range.Font.Name = DEFAULT_PHASE3_HEADER_FONT_NAME
                header_range.Font.Size = DEFAULT_PHASE3_HEADER_FONT_SIZE
                header_range.Font.Color.RGB = CHAPTER_HEAD_COLOR_RGB
                header_range.ParagraphFormat.Alignment = 2
                header_range.ParagraphFormat.SpaceBefore = 0
                header_range.ParagraphFormat.SpaceAfter = 0
                header_range.ParagraphFormat.SpaceWithin = LINE_SPACING_MULTIPLE
                chapter_header.TextFrame2.TextRange.ParagraphFormat.TextDirection = 2
                for paragraph_index in range(1, header_range.Paragraphs().Count + 1):
                    header_range.Paragraphs(paragraph_index).ParagraphFormat.Bullet.Visible = 0

                slide_title = next(
                    (
                        slide.Shapes(shape_index)
                        for shape_index in range(1, slide.Shapes.Count + 1)
                        if slide.Shapes(shape_index).Name == "Chapter Slide Title"
                    ),
                    None,
                )
                if slide_title is None:
                    slide_title = next(
                        (
                            slide.Shapes(shape_index)
                            for shape_index in range(1, slide.Shapes.Count + 1)
                            if slide.Shapes(shape_index).Name == "Title 1"
                        ),
                        None,
                    )
                if slide_title is not None:
                    title_text = slide_title.TextFrame.TextRange
                    title_text.Font.Name = DEFAULT_PHASE3_HEADER_FONT_NAME
                    title_text.Font.Size = DEFAULT_PHASE3_HEADER_FONT_SIZE
                    title_text.Font.Color.RGB = CHAPTER_HEAD_COLOR_RGB
                    title_text.ParagraphFormat.Alignment = 2
                    slide_title.TextFrame2.TextRange.ParagraphFormat.TextDirection = 2

                body_shape = next(
                    (
                        slide.Shapes(shape_index)
                        for shape_index in range(1, slide.Shapes.Count + 1)
                        if slide.Shapes(shape_index).Name in ("Text Placeholder 2", "Text Placeholder 3")
                    ),
                    None,
                )
                if body_shape is None:
                    body_shape = next(
                        (
                            slide.Shapes(shape_index)
                            for shape_index in range(1, slide.Shapes.Count + 1)
                            if slide.Shapes(shape_index).PlaceholderFormat.Type == 2
                            and slide.Shapes(shape_index).Name != "Text Placeholder 4"
                        ),
                        None,
                    )
                if body_shape is None or not body_shape.HasTextFrame:
                    continue

                body_text = body_shape.TextFrame.TextRange
                body_text.Font.Name = DEFAULT_PHASE3_VERSE_FONT_NAME
                body_text.Font.Size = DEFAULT_PHASE3_VERSE_FONT_SIZE
                paragraph_count = body_text.Paragraphs().Count
                has_subtitle = False
                if paragraph_count:
                    first_paragraph = body_text.Paragraphs(1)
                    if first_paragraph.Text.strip():
                        try:
                            has_subtitle = first_paragraph.Font.Color.RGB != white_color
                        except Exception:
                            has_subtitle = False

                body_shape.Left = body_left
                body_shape.Top = (
                    TEXTBOX_TOP_WITH_SUBTITLE_CM if has_subtitle
                    else TEXTBOX_TOP_NO_SUBTITLE_CM
                ) * points_per_cm
                body_shape.Width = body_width
                body_shape.Height = max(0, slide_height - body_shape.Top)
                body_shape.TextFrame2.TextRange.ParagraphFormat.TextDirection = 2

                for paragraph_index in range(1, paragraph_count + 1):
                    paragraph = body_text.Paragraphs(paragraph_index)
                    paragraph.Font.Name = DEFAULT_PHASE3_VERSE_FONT_NAME
                    paragraph.Font.Size = DEFAULT_PHASE3_VERSE_FONT_SIZE
                    paragraph_format = paragraph.ParagraphFormat
                    paragraph_format.SpaceBefore = 0
                    paragraph_format.SpaceWithin = LINE_SPACING_MULTIPLE
                    if has_subtitle and paragraph_index == 1:
                        paragraph.Font.Color.RGB = subtitle_color
                        paragraph.Font.Underline = -1
                        paragraph_format.Alignment = 2
                        paragraph_format.SpaceAfter = SUBTITLE_SPACE_AFTER_POINTS
                    else:
                        paragraph.Font.Underline = 0
                        paragraph_format.Alignment = 7
                        paragraph_format.SpaceAfter = 0

                formatted_slides += 1

    if not found_layout:
        raise ValueError("Presentation has no 'Title + Text' custom layout")
    return formatted_slides


def populate_phase3_chapter_titles_and_headers(
    presentation,
    book_name,
    shortcut,
    header_book_name=None,
):
    """Fill chapter titles and repeating headers without changing slide formatting."""
    header_book_name = header_book_name or book_name
    sections = presentation.SectionProperties
    chapter_sections = []
    for section_index in range(1, sections.Count + 1):
        section_name = sections.Name(section_index)
        prefix = f"{book_name} "
        if not section_name.startswith(prefix):
            continue
        chapter_number = section_name[len(prefix):].strip()
        if not chapter_number.isdigit():
            continue
        chapter_sections.append(
            (sections.FirstSlide(section_index), int(chapter_number))
        )

    if not chapter_sections:
        raise ValueError(f"No numbered chapter sections found for {book_name}")

    chapter_sections.sort()
    starts = {start_slide: chapter_number for start_slide, chapter_number in chapter_sections}
    chapter_headers = {}
    short_title_pattern = re.compile(rf"^{re.escape(shortcut)}\s+\d+$")
    for chapter_index, (start_slide, chapter_number) in enumerate(chapter_sections):
        end_slide = (
            chapter_sections[chapter_index + 1][0] - 1
            if chapter_index + 1 < len(chapter_sections)
            else presentation.Slides.Count
        )
        source_heading = ""
        for slide_number in range(start_slide, end_slide + 1):
            slide = presentation.Slides(slide_number)
            for placeholder_index in (3, 1):
                if placeholder_index > slide.Shapes.Placeholders.Count:
                    continue
                candidate = slide.Shapes.Placeholders(
                    placeholder_index
                ).TextFrame.TextRange.Text.strip()
                if not candidate or short_title_pattern.fullmatch(candidate):
                    continue
                source_heading = candidate
                break
            if source_heading:
                break

        if source_heading:
            if "الإصحاح" in source_heading:
                heading_prefix = source_heading.split("الإصحاح", 1)[0]
                heading_prefix = heading_prefix.strip().rstrip(" -–—\u200f").strip()
                chapter_headers[chapter_number] = (
                    f"{heading_prefix} - الإصحاح الـ{chapter_number}"
                )
            else:
                chapter_headers[chapter_number] = source_heading
        else:
            chapter_headers[chapter_number] = (
                f"سفر {header_book_name} - الإصحاح الـ{chapter_number}"
            )

    active_chapter = None
    for slide_number in range(1, presentation.Slides.Count + 1):
        slide = presentation.Slides(slide_number)
        if slide.Shapes.Placeholders.Count < 3:
            raise ValueError(f"Slide {slide_number} is missing title or header placeholders")

        title_placeholder = slide.Shapes.Placeholders(1)
        header_placeholder = slide.Shapes.Placeholders(3)
        if title_placeholder.PlaceholderFormat.Type != 1:
            raise ValueError(f"Slide {slide_number} placeholder 1 is not a title")

        if slide_number in starts:
            active_chapter = starts[slide_number]
        title_placeholder.TextFrame.TextRange.Text = (
            f"{shortcut} {active_chapter}"
            if slide_number in starts
            else ""
        )
        header_placeholder.TextFrame.TextRange.Text = (
            chapter_headers[active_chapter]
            if active_chapter is not None
            else ""
        )

    return len(chapter_sections)


def update_index_slide_navigation(presentation):
    """Remove named index pictures and link each chapter button shape to its start."""
    chapter_starts = {}
    for slide_index in range(2, len(presentation.slides) + 1):
        slide = presentation.slides[slide_index - 1]
        title_placeholder = next(
            (
                shape
                for shape in slide.shapes
                if shape.is_placeholder
                and shape.placeholder_format.idx == 0
                and shape.has_text_frame
            ),
            None,
        )
        if title_placeholder is None:
            continue
        match = re.search(r"(\d+)\s*$", title_placeholder.text.strip())
        if match:
            chapter_starts[int(match.group(1))] = slide_index

    if not chapter_starts:
        raise ValueError("No chapter-start title placeholders found")

    index_slide = presentation.slides[0]
    index_shapes = list(iter_shapes(index_slide.shapes))
    picture_shapes = [
        shape for shape in index_shapes if shape.name.startswith("Picture")
    ]
    for shape in picture_shapes:
        shape._element.getparent().remove(shape._element)

    buttons = {}
    for shape in iter_shapes(index_slide.shapes):
        if not shape.has_text_frame:
            continue
        label = shape.text.strip()
        if label.isdigit():
            chapter_number = int(label)
            if chapter_number in buttons:
                raise ValueError(f"Duplicate chapter button: {chapter_number}")
            buttons[chapter_number] = shape

    unknown_buttons = set(buttons) - set(chapter_starts)
    if unknown_buttons:
        raise ValueError(
            "Index buttons refer to unknown chapters: "
            f"buttons={sorted(unknown_buttons)}, chapters={sorted(chapter_starts)}"
        )

    for chapter_number, button in buttons.items():
        button.click_action.target_slide = presentation.slides[
            chapter_starts[chapter_number] - 1
        ]

    return len(picture_shapes), len(buttons)


def correct_subtitle_spacing_xml(presentation_file, object_name):
    """Store subtitle after-spacing as points rather than a percentage."""
    presentation_file = Path(presentation_file)
    namespaces = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    }
    temp_file = None

    try:
        with tempfile.NamedTemporaryFile(
            dir=presentation_file.parent,
            suffix=".pptx",
            delete=False,
        ) as temp:
            temp_file = Path(temp.name)

        with zipfile.ZipFile(presentation_file, "r") as source:
            with zipfile.ZipFile(temp_file, "w") as target:
                for item in source.infolist():
                    data = source.read(item.filename)
                    if re.fullmatch(r"ppt/slides/slide\d+\.xml", item.filename):
                        root = etree.fromstring(data)
                        shapes = root.xpath(
                            ".//p:sp[p:nvSpPr/p:cNvPr[@name=$name]]",
                            namespaces=namespaces,
                            name=object_name,
                        )
                        for shape in shapes:
                            for paragraph in shape.xpath("./p:txBody/a:p", namespaces=namespaces):
                                ppr = paragraph.find(qn("a:pPr"))
                                if ppr is None:
                                    continue
                                spacing = ppr.find(qn("a:spcAft"))
                                percent = (
                                    spacing.find(qn("a:spcPct"))
                                    if spacing is not None
                                    else None
                                )
                                if percent is None or percent.get("val") != "600000":
                                    continue
                                points_spacing = etree.Element(qn("a:spcAft"))
                                etree.SubElement(points_spacing, qn("a:spcPts"), val="600")
                                ppr.replace(spacing, points_spacing)
                        data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                    target.writestr(item, data)

        os.replace(temp_file, presentation_file)
    finally:
        if temp_file is not None and temp_file.exists():
            temp_file.unlink()


def build_rendered_chunks(lines):
    """Group lines so subtitles are strictly on line 1 of a 5-line slide.

    Non-subtitle slides have up to 4 lines. If a subtitle appears on lines
    2-4, the current chunk is terminated early and the subtitle starts line 1
    of the next chunk.
    """
    chunks = []
    current_idx = 0
    total_lines = len(lines)

    while current_idx < total_lines:
        first_line = lines[current_idx]
        has_subtitle = first_line[2]
        max_lines = LINES_WITH_SUBTITLE if has_subtitle else LINES_WITHOUT_SUBTITLE

        chunk = [first_line]
        next_idx = current_idx + 1

        while next_idx < total_lines and len(chunk) < max_lines:
            next_line = lines[next_idx]
            if next_line[2]:  # Encounted subtitle: break early so it starts next slide
                break
            chunk.append(next_line)
            next_idx += 1

        chunks.append(chunk)
        current_idx = next_idx

    return chunks


def split_rendered_file(input_file, output_file, object_name):
    """Split rendered PowerPoint lines into new slides using PowerPoint COM."""
    import win32com.client as win32

    powerpoint = win32.DispatchEx("PowerPoint.Application")
    presentation = powerpoint.Presentations.Open(
        str(Path(input_file).resolve()),
        WithWindow=False,
    )
    created_slides = 0
    source_slide_count = presentation.Slides.Count

    try:
        # Work backwards so newly inserted duplicates do not affect source indexes.
        for slide_number in range(source_slide_count, 0, -1):
            source_slide = presentation.Slides(slide_number)
            source_shape = next(
                (
                    source_slide.Shapes(shape_number)
                    for shape_number in range(1, source_slide.Shapes.Count + 1)
                    if source_slide.Shapes(shape_number).Name == object_name
                ),
                None,
            )
            if source_shape is None or not source_shape.HasTextFrame:
                continue

            # Ensure font size is set before rendered line calculation
            source_shape.TextFrame.TextRange.Font.Size = VERSE_FONT_SIZE_POINTS

            text_range = source_shape.TextFrame.TextRange
            lines = []
            line_number = 1
            while True:
                line = text_range.Lines(line_number, 1)
                if line.Length == 0:
                    break
                is_sub = is_line_subtitle(line)
                lines.append((line.Start, line.Length, is_sub))
                line_number += 1

            chunks = build_rendered_chunks(lines)

            if len(chunks) <= 1:
                has_subtitle = chunks[0][0][2] if (chunks and chunks[0]) else False
                apply_rendered_slide_format(source_shape, has_subtitle)
                continue

            original_text_length = text_range.Length
            original_slide_index = source_slide.SlideIndex
            duplicate_slides = []

            for chunk in chunks:
                duplicate = source_slide.Duplicate().Item(1)
                duplicate_slides.append(duplicate)
                duplicate_shape = next(
                    duplicate.Shapes(shape_number)
                    for shape_number in range(1, duplicate.Shapes.Count + 1)
                    if duplicate.Shapes(shape_number).Name == object_name
                )
                chunk_start = chunk[0][0]
                chunk_end = chunk[-1][0] + chunk[-1][1] - 1
                duplicate_range = duplicate_shape.TextFrame.TextRange

                if chunk_end < original_text_length:
                    duplicate_range.Characters(
                        chunk_end + 1,
                        original_text_length - chunk_end,
                    ).Text = ""
                if chunk_start > 1:
                    duplicate_range.Characters(1, chunk_start - 1).Text = ""

                has_subtitle = chunk[0][2]
                apply_rendered_slide_format(duplicate_shape, has_subtitle)
                created_slides += 1

            # Duplicate() can insert slides at an implementation-defined
            # location. Move each generated chunk beside its source explicitly.
            for offset, duplicate in enumerate(duplicate_slides, start=1):
                duplicate.MoveTo(original_slide_index + offset)

            source_slide.Delete()

        presentation.SaveAs(str(Path(output_file).resolve()))
    finally:
        try:
            presentation.Close()
        except Exception:
            pass
        try:
            powerpoint.Quit()
        except Exception:
            pass

    correct_subtitle_spacing_xml(output_file, object_name)
    print(f"Rendered source slides: {source_slide_count}")
    print(f"Rendered slides created: {created_slides}")
    print(f"Output: {output_file}")


def split_rendered_files(input_folder, output_folder):
    """Split every presentation in a folder into a separate output folder."""
    output_folder.mkdir(parents=True, exist_ok=True)
    input_files = sorted(input_folder.glob("*.pptx"))

    if not input_files:
        raise FileNotFoundError(f"No PowerPoint files found in {input_folder}")

    print(f"Splitting {len(input_files)} presentations into rendered groups of four lines...")
    for index, input_file in enumerate(input_files, start=1):
        output_file = output_folder / input_file.name
        print(f"[{index}/{len(input_files)}] {input_file.name}")
        split_rendered_file(
            input_file,
            output_file,
            DEFAULT_OBJECT_NAME,
        )


def phase3_book(input_file, output_file, testament_folder):
    """Add distinct chapter titles, formatted headers, and sections to Phase 2."""
    import win32com.client as win32

    input_file = Path(input_file).resolve()
    output_file = Path(output_file).resolve()
    if input_file == output_file:
        raise ValueError("Phase 3 output must not overwrite its Phase 2 input")

    match = re.match(r"\s*(\d+)", input_file.stem)
    if match is None:
        raise ValueError(f"Could not determine book number from filename: {input_file.name}")
    book_number = int(match.group(1))
    shortcut = BOOK_SHORTCUTS.get(testament_folder, {}).get(book_number)
    book_name = BOOK_SECTION_NAMES.get(testament_folder, {}).get(book_number)
    if shortcut is None or book_name is None:
        raise ValueError(
            f"No Phase 3 book mapping for {testament_folder} book {book_number}: {input_file.name}"
        )

    staged_presentation = Presentation(input_file)
    header_count = 0
    for slide in staged_presentation.slides:
        for shape in slide.shapes:
            if shape.name != TITLE_OBJECT_NAME:
                continue
            nv_sp_pr = shape._element.find(qn("p:nvSpPr"))
            nv_pr = nv_sp_pr.find(qn("p:nvPr")) if nv_sp_pr is not None else None
            placeholder = nv_pr.find(qn("p:ph")) if nv_pr is not None else None
            if nv_pr is not None and placeholder is not None:
                nv_pr.remove(placeholder)
            shape._element.nvSpPr.cNvPr.set("name", "Chapter Header")
            header_count += 1
    if not header_count:
        raise ValueError(f"No chapter header shapes found in {input_file.name}")

    output_file.parent.mkdir(parents=True, exist_ok=True)
    staged_presentation.save(output_file)

    powerpoint = win32.DispatchEx("PowerPoint.Application")
    presentation = powerpoint.Presentations.Open(
        str(output_file),
        WithWindow=False,
    )

    try:
        chapter_starts = []
        previous_heading = None
        for slide_number in range(1, presentation.Slides.Count + 1):
            slide = presentation.Slides(slide_number)
            title_shape = next(
                (
                    slide.Shapes(shape_number)
                    for shape_number in range(1, slide.Shapes.Count + 1)
                    if slide.Shapes(shape_number).Name == "Chapter Header"
                    and slide.Shapes(shape_number).HasTextFrame
                ),
                None,
            )
            if title_shape is None:
                continue

            heading = title_shape.TextFrame.TextRange.Text.strip()
            if not heading or heading == previous_heading:
                continue

            chapter_starts.append((slide_number, heading))
            previous_heading = heading

        if not chapter_starts:
            raise ValueError(f"No chapter headings found in {input_file.name}")

        sections = presentation.SectionProperties
        if sections.Count:
            raise ValueError(
                f"Expected Phase 2 input without sections: {input_file.name}"
            )

        sections.AddBeforeSlide(chapter_starts[0][0], f"{book_name} 1")
        for chapter_number, (start_slide, _) in enumerate(chapter_starts[1:], start=2):
            sections.AddBeforeSlide(start_slide, f"{book_name} {chapter_number}")
        sections.Rename(1, book_name)

        for chapter_index, (start_slide, heading) in enumerate(chapter_starts, start=1):
            end_slide = (
                chapter_starts[chapter_index][0] - 1
                if chapter_index < len(chapter_starts)
                else presentation.Slides.Count
            )
            for slide_number in range(start_slide, end_slide + 1):
                slide = presentation.Slides(slide_number)
                header_shape = next(
                    (
                        slide.Shapes(shape_number)
                        for shape_number in range(1, slide.Shapes.Count + 1)
                        if slide.Shapes(shape_number).Name == "Chapter Header"
                        and slide.Shapes(shape_number).HasTextFrame
                    ),
                    None,
                )
                if header_shape is None:
                    continue

                if (
                    testament_folder == "عهد قديم" and book_number == 21
                ) or (
                    testament_folder == "عهد جديد" and book_number in (24, 25)
                ):
                    normalized_heading = heading
                else:
                    heading_prefix, separator, _ = heading.partition("الإصحاح")
                    if not separator:
                        normalized_heading = f"{shortcut} الإصحاح الـ{chapter_index}"
                    else:
                        normalized_heading = f"{heading_prefix}الإصحاح الـ{chapter_index}"
                header_range = header_shape.TextFrame.TextRange
                header_range.Text = normalized_heading
                header_range.Font.Size = CHAPTER_HEAD_FONT_SIZE
                header_range.Font.Name = CHAPTER_HEAD_FONT_NAME
                header_range.Font.Color.RGB = CHAPTER_HEAD_COLOR_RGB
                header_range.ParagraphFormat.Alignment = 2  # ppAlignCenter
                header_shape.TextFrame2.TextRange.ParagraphFormat.TextDirection = 2

                slide_title = slide.Shapes.AddPlaceholder(1, -58, 8, 58, 32)
                slide_title.Name = "Chapter Slide Title"
                slide_title.Left = -58
                slide_title.Top = 8
                slide_title.Width = 58
                slide_title.Height = 32
                title_range = slide_title.TextFrame.TextRange
                chapter_title = f"{shortcut} {chapter_index}" if slide_number == start_slide else ""
                title_range.Text = chapter_title
                if chapter_title:
                    title_range.Font.Size = 18

            formatted_slide_count = format_title_text_layout(presentation)
        presentation.Save()
    finally:
        try:
            presentation.Close()
        except Exception:
            pass
        try:
            powerpoint.Quit()
        except Exception:
            pass

    print(f"Chapters: {len(chapter_starts)}")
    print(f"Title + Text slides formatted: {formatted_slide_count}")
    print(f"Output: {output_file}")


def phase3_files(input_folder, output_folder):
    """Create Phase 3 copies for every Phase 2 book in a testament folder."""
    input_folder = Path(input_folder)
    output_folder = Path(output_folder)
    files = sorted(input_folder.glob("*.pptx"))
    if not files:
        raise FileNotFoundError(f"No Phase 2 presentations found in {input_folder}")

    testament_folder = input_folder.parent.name
    if testament_folder not in BOOK_SHORTCUTS:
        raise ValueError(f"Unknown testament folder: {testament_folder}")

    output_folder.mkdir(parents=True, exist_ok=True)
    print(f"Creating Phase 3 for {len(files)} books in {input_folder}...")
    failures = []
    for index, input_file in enumerate(files, start=1):
        output_file = output_folder / input_file.name
        print(f"[{index}/{len(files)}] {input_file.name}")
        try:
            phase3_book(input_file, output_file, testament_folder)
        except Exception as error:
            failures.append((input_file.name, str(error)))
            print(f"FAILED: {input_file.name}: {error}")
    return failures


def phase3_all_files(bible_folder):
    """Create Phase 3 copies for every book in both testament folders."""
    bible_folder = Path(bible_folder)
    failures = []
    for testament_folder in ("عهد جديد", "عهد قديم"):
        source_folder = bible_folder / testament_folder / "phase2"
        output_folder = bible_folder / testament_folder / "phase3"
        failures.extend(phase3_files(source_folder, output_folder))
    if failures:
        print(f"Phase 3 failures ({len(failures)}):")
        for filename, error in failures:
            print(f"- {filename}: {error}")
    else:
        print("Phase 3 completed for all books without errors.")


def format_phase3_layouts_all_files(bible_folder):
    """Format Title + Text layouts and slides in every Phase 3 book deck."""
    import win32com.client as win32

    bible_folder = Path(bible_folder).resolve()
    presentations = []
    for testament_folder in ("عهد قديم", "عهد جديد"):
        phase3_folder = bible_folder / testament_folder / "phase3"
        if not phase3_folder.is_dir():
            raise FileNotFoundError(f"Phase 3 folder not found: {phase3_folder}")
        presentations.extend(
            path for path in sorted(phase3_folder.glob("*.pptx"))
            if not path.name.startswith("~$")
            and re.match(r"^\d+\s", path.name)
        )
    if not presentations:
        raise FileNotFoundError(f"No Phase 3 presentations found under {bible_folder}")

    locked = [
        path for path in presentations
        if path.with_name(f"~${path.name}").exists()
    ]
    if locked:
        names = ", ".join(path.name for path in locked)
        raise RuntimeError(f"Close these Phase 3 presentations in PowerPoint before formatting: {names}")

    powerpoint = win32.DispatchEx("PowerPoint.Application")
    failures = []
    try:
        for index, presentation_file in enumerate(presentations, start=1):
            temporary_file = None
            presentation = None
            try:
                with tempfile.NamedTemporaryFile(
                    dir=presentation_file.parent,
                    prefix=f".{presentation_file.stem}_format_",
                    suffix=presentation_file.suffix,
                    delete=False,
                ) as temporary:
                    temporary_file = Path(temporary.name)
                shutil.copy2(presentation_file, temporary_file)
                presentation = powerpoint.Presentations.Open(
                    str(temporary_file),
                    WithWindow=False,
                )
                slide_count = format_title_text_layout(presentation)
                presentation.Save()
                presentation.Close()
                presentation = None
                os.replace(temporary_file, presentation_file)
                temporary_file = None
                print(
                    f"[{index}/{len(presentations)}] {presentation_file.name}: "
                    f"formatted {slide_count} slides"
                )
            except Exception as error:
                failures.append((presentation_file.name, str(error)))
                print(f"[{index}/{len(presentations)}] {presentation_file.name}: FAILED: {error}")
            finally:
                if presentation is not None:
                    presentation.Close()
                if temporary_file is not None and temporary_file.exists():
                    temporary_file.unlink()
    finally:
        powerpoint.Quit()

    if failures:
        detail = "; ".join(f"{name}: {error}" for name, error in failures)
        raise RuntimeError(f"Formatting failed for {len(failures)} presentation(s): {detail}")
    print(f"Phase 3 layouts formatted: {len(presentations)} presentations")


def apply_phase3_master_backgrounds(bible_folder, background_file):
    """Add the blue artwork to each Phase 3 presentation's slide master."""
    import win32com.client as win32

    bible_folder = Path(bible_folder).resolve()
    background_file = Path(background_file).resolve()
    if not background_file.is_file():
        raise FileNotFoundError(f"Master background image not found: {background_file}")

    points_per_cm = 72 / 2.54
    picture_left = 0
    picture_top = 12.08 * points_per_cm
    picture_width = 25.4 * points_per_cm
    picture_height = 6.97 * points_per_cm
    picture_name = "Phase 3 Blue Background"
    presentations = []
    for testament_folder in ("عهد قديم", "عهد جديد"):
        phase3_folder = bible_folder / testament_folder / "phase3"
        if not phase3_folder.is_dir():
            raise FileNotFoundError(f"Phase 3 folder not found: {phase3_folder}")
        presentations.extend(
            path for path in sorted(phase3_folder.glob("*.pptx"))
            if not path.name.startswith("~$")
        )
    if not presentations:
        raise FileNotFoundError(f"No Phase 3 presentations found under {bible_folder}")

    powerpoint = win32.DispatchEx("PowerPoint.Application")
    try:
        for index, presentation_file in enumerate(presentations, start=1):
            temporary_file = None
            presentation = None
            try:
                with tempfile.NamedTemporaryFile(
                    dir=presentation_file.parent,
                    prefix=f".{presentation_file.stem}_master_",
                    suffix=presentation_file.suffix,
                    delete=False,
                ) as temporary:
                    temporary_file = Path(temporary.name)
                shutil.copy2(presentation_file, temporary_file)

                presentation = powerpoint.Presentations.Open(
                    str(temporary_file),
                    WithWindow=False,
                )
                added_count = 0
                for design_index in range(1, presentation.Designs.Count + 1):
                    master_shapes = presentation.Designs(design_index).SlideMaster.Shapes
                    existing = next(
                        (
                            master_shapes(shape_index)
                            for shape_index in range(1, master_shapes.Count + 1)
                            if master_shapes(shape_index).Name == picture_name
                        ),
                        None,
                    )
                    if existing is not None:
                        continue
                    picture = master_shapes.AddPicture(
                        str(background_file),
                        0,
                        -1,
                        picture_left,
                        picture_top,
                        picture_width,
                        picture_height,
                    )
                    picture.Name = picture_name
                    picture.ZOrder(1)  # msoSendToBack
                    added_count += 1

                presentation.Save()
                presentation.Close()
                presentation = None
                os.replace(temporary_file, presentation_file)
                temporary_file = None
                print(
                    f"[{index}/{len(presentations)}] {presentation_file.name}: "
                    f"{'background added' if added_count else 'already updated'}"
                )
            finally:
                if presentation is not None:
                    presentation.Close()
                if temporary_file is not None and temporary_file.exists():
                    temporary_file.unlink()
    finally:
        powerpoint.Quit()

    print(f"Phase 3 master backgrounds checked: {len(presentations)}")


def merge_all_books_into_master(master_file, bible_folder, output_file):
    """Replace every master book placeholder with its complete Phase 3 presentation."""
    import win32com.client as win32

    master_file = Path(master_file).resolve()
    bible_folder = Path(bible_folder).resolve()
    output_file = Path(output_file).resolve()
    if output_file == master_file:
        raise ValueError("Merged output must be separate from the master presentation")
    if output_file.exists():
        raise FileExistsError(f"Refusing to overwrite existing merged output: {output_file}")
    if not master_file.is_file():
        raise FileNotFoundError(f"Master presentation not found: {master_file}")

    book_sources = []
    for testament_folder in ("عهد قديم", "عهد جديد"):
        expected_numbers = set(BOOK_SECTION_NAMES[testament_folder])
        phase3_folder = bible_folder / testament_folder / "phase3"
        if not phase3_folder.is_dir():
            raise FileNotFoundError(f"Phase 3 folder not found: {phase3_folder}")

        files_by_number = {}
        for candidate in phase3_folder.glob("*.pptx"):
            if candidate.name.startswith("~$"):
                continue
            match = re.match(r"^(\d+)\s", candidate.stem)
            if not match:
                raise ValueError(f"Cannot determine book number from {candidate.name!r}")
            number = int(match.group(1))
            if number in files_by_number:
                raise ValueError(f"Duplicate Phase 3 file for {testament_folder} book {number}")
            files_by_number[number] = candidate.resolve()

        missing = sorted(expected_numbers - files_by_number.keys())
        unexpected = sorted(files_by_number.keys() - expected_numbers)
        if missing or unexpected:
            raise ValueError(
                f"Unexpected Phase 3 files in {phase3_folder}; missing={missing}, unexpected={unexpected}"
            )

        for book_number in sorted(expected_numbers):
            book_file = files_by_number[book_number]
            if output_file == book_file:
                raise ValueError("Merged output must be separate from every source presentation")
            book_sources.append({
                "testament": testament_folder,
                "number": book_number,
                "name": BOOK_SECTION_NAMES[testament_folder][book_number],
                "path": book_file,
            })

    powerpoint = win32.DispatchEx("PowerPoint.Application")
    try:
        for book in book_sources:
            source = powerpoint.Presentations.Open(str(book["path"]), WithWindow=False)
            try:
                sections = source.SectionProperties
                book["slide_count"] = source.Slides.Count
                if sections.Count == 0 and book["name"] == "المزامير":
                    index_slide = source.Slides(1)
                    index_has_book_name = any(
                        shape.HasTextFrame
                        and shape.TextFrame.HasText
                        and shape.TextFrame.TextRange.Text.strip() == book["name"]
                        for shape in (
                            index_slide.Shapes(i)
                            for i in range(1, index_slide.Shapes.Count + 1)
                        )
                    )
                    if not index_has_book_name:
                        raise ValueError(f"Missing the Psalms index slide in {book['path'].name}")

                    chapter_starts = []
                    previous_heading = None
                    for slide_number in range(2, book["slide_count"] + 1):
                        slide = source.Slides(slide_number)
                        header = next(
                            (
                                slide.Shapes(i)
                                for i in range(1, slide.Shapes.Count + 1)
                                if slide.Shapes(i).Name == "Chapter Header"
                                and slide.Shapes(i).HasTextFrame
                            ),
                            None,
                        )
                        heading = (
                            header.TextFrame.TextRange.Text.strip()
                            if header is not None and header.TextFrame.HasText
                            else ""
                        )
                        if heading and heading != previous_heading:
                            chapter_starts.append(
                                (slide_number, f"{book['name']} {len(chapter_starts) + 1}")
                            )
                            previous_heading = heading
                    if not chapter_starts or chapter_starts[0][0] != 2:
                        raise ValueError(f"Could not derive Psalms chapter starts in {book['path'].name}")
                    book["chapters"] = chapter_starts
                else:
                    if sections.Count < 2 or sections.Name(1) != book["name"]:
                        raise ValueError(
                            f"{book['path'].name} does not start with the {book['name']!r} index section"
                        )
                    if sections.FirstSlide(1) != 1 or sections.SlidesCount(1) != 1:
                        raise ValueError(f"Expected one linked index slide at the start of {book['path'].name}")
                    book["chapters"] = [
                        (sections.FirstSlide(index), sections.Name(index))
                        for index in range(2, sections.Count + 1)
                    ]
                if not book["chapters"] or any(first < 2 for first, _ in book["chapters"]):
                    raise ValueError(f"Invalid chapter sections in {book['path'].name}")
                if any(
                    first >= next_first
                    for (first, _), (next_first, _) in zip(book["chapters"], book["chapters"][1:])
                ):
                    raise ValueError(f"Chapter sections are not ordered in {book['path'].name}")
            finally:
                source.Close()
    except Exception:
        powerpoint.Quit()
        raise

    output_file.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        dir=output_file.parent,
        prefix=f".{output_file.stem}_",
        suffix=".pptx",
        delete=False,
    ) as temporary:
        temporary_file = Path(temporary.name)

    try:
        shutil.copy2(master_file, temporary_file)
        presentation = powerpoint.Presentations.Open(
            str(temporary_file),
            WithWindow=False,
        )
        try:
            sections = presentation.SectionProperties
            old_heading = next(
                (i for i in range(1, sections.Count + 1) if sections.Name(i) == "العهد القديم"),
                None,
            )
            new_heading = next(
                (i for i in range(1, sections.Count + 1) if sections.Name(i) == "العهد الجديد"),
                None,
            )
            if old_heading is None or new_heading is None or old_heading >= new_heading:
                raise ValueError("Master must contain ordered Old and New Testament sections")

            old_book_name = BOOK_SECTION_NAMES["عهد قديم"][1]
            genesis_index_section = next(
                (i for i in range(old_heading + 1, new_heading) if sections.Name(i) == old_book_name),
                None,
            )
            genesis_chapter_section = next(
                (
                    i for i in range(old_heading + 1, new_heading)
                    if sections.Name(i) == f"{old_book_name} 1"
                ),
                None,
            )
            old_book_sections = [
                i for i in range(old_heading + 1, new_heading)
                if sections.FirstSlide(i) > 0 and i != genesis_chapter_section
            ]
            new_book_sections = [
                i for i in range(new_heading + 1, sections.Count + 1)
                if sections.FirstSlide(i) > 0
            ]
            if (
                genesis_index_section is None
                or genesis_chapter_section is None
                or len(old_book_sections) != len(BOOK_SECTION_NAMES["عهد قديم"])
                or len(new_book_sections) != len(BOOK_SECTION_NAMES["عهد جديد"])
            ):
                raise ValueError("Master book placeholders do not match the expected 46 Old and 27 New Testament books")

            target_sections = {
                "عهد قديم": old_book_sections,
                "عهد جديد": new_book_sections,
            }
            slot_slide_ids = {}
            for testament, book_section_indices in target_sections.items():
                for book_number, section_index in enumerate(book_section_indices, start=1):
                    slide_count = sections.SlidesCount(section_index)
                    if slide_count != 1:
                        raise ValueError(
                            f"Expected one master placeholder slide in section {sections.Name(section_index)!r}"
                        )
                    first_slide = sections.FirstSlide(section_index)
                    slot_slide_ids[(testament, book_number)] = [
                        presentation.Slides(first_slide).SlideID
                    ]

            genesis_index_slide = sections.FirstSlide(genesis_index_section)
            genesis_chapter_slide = sections.FirstSlide(genesis_chapter_section)
            if (
                sections.SlidesCount(genesis_index_section) != 1
                or sections.SlidesCount(genesis_chapter_section) != 1
                or genesis_chapter_slide != genesis_index_slide + 1
            ):
                raise ValueError("Expected Genesis index and chapter placeholders to be adjacent single slides")
            slot_slide_ids[("عهد قديم", 1)] = [
                presentation.Slides(genesis_index_slide).SlideID,
                presentation.Slides(genesis_chapter_slide).SlideID,
            ]

            for book in book_sources:
                slot_ids = slot_slide_ids[(book["testament"], book["number"])]
                index_slot_id = slot_ids[0]
                index_slot = next(
                    presentation.Slides(i)
                    for i in range(1, presentation.Slides.Count + 1)
                    if presentation.Slides(i).SlideID == index_slot_id
                )
                index_slot_number = index_slot.SlideIndex
                is_genesis = book["testament"] == "عهد قديم" and book["number"] == 1
                insert_after = index_slot_number if is_genesis else index_slot_number - 1
                print(f'Book "{book["name"]}" is being added')
                inserted = presentation.Slides.InsertFromFile(
                    str(book["path"]),
                    insert_after,
                    1,
                    book["slide_count"],
                )
                if inserted != book["slide_count"]:
                    raise RuntimeError(
                        f"PowerPoint inserted {inserted} slides for {book['name']}; "
                        f"expected {book['slide_count']}"
                    )

                first_imported_id = presentation.Slides(insert_after + 1).SlideID
                for slide_id in reversed(slot_ids):
                    slot = next(
                        (
                            presentation.Slides(i)
                            for i in range(1, presentation.Slides.Count + 1)
                            if presentation.Slides(i).SlideID == slide_id
                        ),
                        None,
                    )
                    if slot is None:
                        raise RuntimeError(f"Could not locate a master placeholder for {book['name']}")
                    slot.Delete()

                sections = presentation.SectionProperties
                if is_genesis:
                    stale_chapter_section = next(
                        (
                            i for i in range(1, sections.Count + 1)
                            if sections.Name(i) == f"{book['name']} 1"
                        ),
                        None,
                    )
                    if stale_chapter_section is None:
                        raise RuntimeError("Could not locate the Genesis chapter placeholder section")
                    sections.Delete(stale_chapter_section, False)

                first_imported = next(
                    presentation.Slides(i)
                    for i in range(1, presentation.Slides.Count + 1)
                    if presentation.Slides(i).SlideID == first_imported_id
                )
                imported_start = first_imported.SlideIndex
                index_section = next(
                    (
                        i for i in range(1, sections.Count + 1)
                        if sections.FirstSlide(i) == imported_start
                    ),
                    None,
                )
                if index_section is None:
                    sections.AddBeforeSlide(imported_start, book["name"])
                    sections = presentation.SectionProperties
                    index_section = next(
                        (
                            i for i in range(1, sections.Count + 1)
                            if sections.FirstSlide(i) == imported_start
                        ),
                        None,
                    )
                if index_section is None:
                    raise RuntimeError(
                        f"Could not create a section for the imported {book['name']!r} index slide"
                    )
                sections.Rename(index_section, book["name"])

                for source_start, section_name in reversed(book["chapters"]):
                    sections.AddBeforeSlide(
                        imported_start + source_start - 1,
                        section_name,
                    )

                sections = presentation.SectionProperties
                for source_start, section_name in book["chapters"]:
                    destination_start = imported_start + source_start - 1
                    if not any(
                        sections.Name(i) == section_name
                        and sections.FirstSlide(i) == destination_start
                        for i in range(1, sections.Count + 1)
                    ):
                        raise RuntimeError(
                            f"Section {section_name!r} was not created at slide {destination_start}"
                        )
                print(f'Book "{book["name"]}" finished')

            presentation.Save()
        finally:
            presentation.Close()

        os.replace(temporary_file, output_file)
    finally:
        powerpoint.Quit()
        if temporary_file.exists():
            temporary_file.unlink()

    print(f"Books merged: {len(book_sources)}")
    print(f"Inserted slides: {sum(book['slide_count'] for book in book_sources)}")
    print(f"Merged output: {output_file}")


# ---------------------------------------------------------------------------
# VS Code run configuration
# ---------------------------------------------------------------------------

def run():
    """Run the workflow selected in the configuration section above."""
    if WORKFLOW == "chapters":
        replace_chapters(DEFAULT_INPUT_FILE, DEFAULT_OUTPUT_FILE)
        return

    if WORKFLOW == "superscript":
        superscript_object(
            DEFAULT_INPUT_FILE,
            DEFAULT_OUTPUT_FILE,
            DEFAULT_OBJECT_NAME,
        )
        return

    if WORKFLOW == "both":
        intermediate_file = DEFAULT_OUTPUT_FILE.with_name(
            f"{DEFAULT_OUTPUT_FILE.stem}_chapters{DEFAULT_OUTPUT_FILE.suffix}"
        )
        replace_chapters(DEFAULT_INPUT_FILE, intermediate_file)
        superscript_object(
            intermediate_file,
            DEFAULT_OUTPUT_FILE,
            DEFAULT_OBJECT_NAME,
        )
        return

    if WORKFLOW == "merge_chapters":
        merge_chapters(
            DEFAULT_INPUT_FILE,
            DEFAULT_OUTPUT_FILE,
            TITLE_OBJECT_NAME,
            DEFAULT_OBJECT_NAME,
        )
        return

    if WORKFLOW == "all":
        process_all(DEFAULT_INPUT_FILE, DEFAULT_OUTPUT_FILE)
        return

    if WORKFLOW == "all_files":
        process_all_files(DEFAULT_INPUT_FOLDER, DEFAULT_OUTPUT_FOLDER)
        return

    if WORKFLOW == "split_rendered":
        split_rendered_file(
            DEFAULT_INPUT_FILE,
            DEFAULT_OUTPUT_FILE,
            DEFAULT_OBJECT_NAME,
        )
        return

    if WORKFLOW == "split_rendered_files":
        split_rendered_files(
            DEFAULT_INPUT_FOLDER,
            DEFAULT_PHASE2_OUTPUT_FOLDER,
        )
        return

    if WORKFLOW == "phase3_all_files":
        phase3_all_files(DEFAULT_INPUT_FOLDER.parent)
        return

    if WORKFLOW == "format_phase3_layouts":
        format_phase3_layouts_all_files(DEFAULT_INPUT_FOLDER.parent)
        return

    if WORKFLOW == "merge_all_books":
        merge_all_books_into_master(
            DEFAULT_MASTER_PRESENTATION,
            DEFAULT_INPUT_FOLDER.parent,
            DEFAULT_FULL_BIBLE_OUTPUT,
        )
        return

    raise ValueError(
        f"Unknown WORKFLOW {WORKFLOW!r}. "
        "Choose a workflow listed in the configuration section."
    )


if __name__ == "__main__":
    run()