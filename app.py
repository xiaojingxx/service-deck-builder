import os
import io
import base64
import tempfile
import subprocess
import shutil
from io import BytesIO
from shutil import which

import fitz  # PyMuPDF
import gspread
import psutil
import re
import streamlit as st
from streamlit_ace import st_ace
from google.oauth2.service_account import Credentials
from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Pt, Emu
from lxml import etree
from PIL import Image
from pptx_tools.creator import PPTXCreator
from docx import Document


# =========================================================
# CONFIG
# =========================================================
SOFFICE_PATH = os.environ.get("SOFFICE_PATH", "soffice")

SHEET_KEY = st.secrets["SHEET_KEY"]
WORKSHEET_NAME = st.secrets["WORKSHEET_NAME"]

FIRST_LAYOUT_NAME = "TEMPLATE_FIRST"
REST_LAYOUT_NAME = "TEMPLATE_REST"
DIVIDER_LAYOUT_NAME = "SECTION_DIVIDER"

# Song-part footers (e.g. "Chorus" bottom-left, "Verse 1 | Verse 2 | Chorus" links bottom-right).
# Add text placeholders with these names (Selection Pane) to TEMPLATE_FIRST / TEMPLATE_REST
# in Slide Master to control their look. Without them, plain text boxes are used.
PART_LABEL_PLACEHOLDER_NAME = "PART_LABEL"
PART_NAV_PLACEHOLDER_NAME = "PART_NAV"
PART_NAV_SEPARATOR = " | "
PART_FOOTER_FONT_SIZE_PT = 14  # only used for the text-box fallback

# Lyrics lines like "[Verse 1]" / "[Chorus]" (or "#Verse 1#" / "#Chorus#") start a new
# song part (and a new slide). The marker must be the whole line.
PART_MARKER_RE = re.compile(r"^\s*(?:#\s*([^#]*?)\s*#|\[\s*([^\[\]]*?)\s*\])\s*$")
# A bare "[]" (or "##") line ends the current part (following lines are unlabelled).
# Parts whose content is metadata, not lyrics, so they never become slides.
METADATA_PART_LABELS = {"title"}


# =========================================================
# STREAMLIT PAGE
# =========================================================
st.set_page_config(page_title="Service Deck Builder", layout="wide")
st.title("Service Deck Builder")


# =========================================================
# GOOGLE SHEETS
# =========================================================
@st.cache_resource
def get_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credentials = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes,
    )
    gc = gspread.authorize(credentials)
    return gc.open_by_key(SHEET_KEY).worksheet(WORKSHEET_NAME)


sheet = get_sheet()


@st.cache_data(show_spinner=False)
def get_all_records_cached():
    return sheet.get_all_records()


# =========================================================
# SESSION STATE
# =========================================================
DEFAULTS = {
    "setlist": [],
    "loaded_song": None,
    "editing_setlist_index": None,
    "pending_setlist_load": None,
    "reset_editor_pending": False,
    "setlist_selected_index": 0,
    "uploaded_templates": {},
    "selected_template_name": None,
    "editor_umh": "",
    "editor_title": "",
    "editor_text": "",
    "editor_override_lyrics_font_size": False,
    "editor_override_line_spacing": False,
    "editor_lyrics_font_size_pt": 32,
    "editor_line_spacing": 1.2,
    "refresh_on_new_line": True,
    "editor_ace_key": 0,
    "last_editor_text": "",
    "last_current_song_signature": None,
    "editor_status_message": "",
    "current_preview_slide": 1,
    "preview_mode": "song",
    "current_song_preview_images": None,
    "current_song_preview_stats": None,
    "service_preview_images": None,
    "service_preview_stats": None,
    "service_preview_error": None,
    "service_song_start_slides": [],
    "ppt_data": None,
    "last_split_settings": None,
    "preserve_template_slides": True,
    "template_sections": [],
    "hidden_section_ids": [],
    "service_output_mode": "full",  # full | songs
    "selected_song_section_id": None,
    "setlist_selectbox_sidebar": 0,
    "pending_setlist_selectbox_index": None,
    "service_order_blocks": [],
    "song_store": {},
    "last_docx_section_mapping": [],
    "selected_song_id": None,
    "song_selectbox_version": 0,
    "smart_split_enabled": False,
    "smart_split_max_chars": 30,
    "smart_split_max_lines_per_slide": 4,
    "last_docx_import_report": None,
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# =========================================================
# BASIC HELPERS
# =========================================================
def soffice_available() -> bool:
    if SOFFICE_PATH == "soffice":
        return which("soffice") is not None
    return os.path.exists(SOFFICE_PATH)


def clear_service_outputs():
    st.session_state["ppt_data"] = None
    st.session_state["service_preview_images"] = None
    st.session_state["service_preview_stats"] = None
    st.session_state["service_preview_error"] = None
    st.session_state["service_song_start_slides"] = []


def format_bytes(num_bytes: int) -> str:
    if num_bytes < 1024:
        return f"{num_bytes} B"
    if num_bytes < 1024**2:
        return f"{num_bytes / 1024:.1f} KB"
    if num_bytes < 1024**3:
        return f"{num_bytes / (1024**2):.2f} MB"
    return f"{num_bytes / (1024**3):.2f} GB"


def preview_stats(images):
    if not images:
        return {
            "count": 0,
            "total_bytes": 0,
            "avg_bytes": 0,
            "max_bytes": 0,
        }

    sizes = [len(img) for img in images]
    total = sum(sizes)

    return {
        "count": len(images),
        "total_bytes": total,
        "avg_bytes": total // len(images),
        "max_bytes": max(sizes),
    }


def get_runtime_resource_stats():
    proc = psutil.Process(os.getpid())
    mem = proc.memory_info()

    vm = psutil.virtual_memory()
    disk = shutil.disk_usage("/")

    return {
        "process_rss": mem.rss,
        "process_vms": mem.vms,
        "system_available_ram": vm.available,
        "system_total_ram": vm.total,
        "system_used_ram": vm.used,
        "disk_free": disk.free,
        "disk_total": disk.total,
        "disk_used": disk.used,
    }


def find_row_by_umh(umh_number: str):
    umh_number = str(umh_number).strip()
    for row in get_all_records_cached():
        if str(row.get("UMH Number", "")).strip() == umh_number:
            return row
    return None


def find_row_by_title(title: str):
    target = re.sub(r"\s+", " ", str(title or "")).strip().lower()
    if not target:
        return None
    for row in get_all_records_cached():
        row_title = re.sub(r"\s+", " ", str(row.get("Title", ""))).strip().lower()
        if row_title == target:
            return row
    return None


def search_titles(keyword: str, limit: int = 20):
    keyword = keyword.lower().strip()
    matches = []
    if not keyword:
        return matches

    for row in get_all_records_cached():
        title = str(row.get("Title", "")).strip()
        if keyword in title.lower():
            matches.append(row)

    return matches[:limit]

def split_slides_by_line_count(text: str, lines_per_slide: int = 4) -> list[list[str]]:
    if lines_per_slide < 1:
        lines_per_slide = 1

    verses = []
    current_verse = []

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if line == "":
            if current_verse:
                verses.append(current_verse)
                current_verse = []
        else:
            current_verse.append(line)

    if current_verse:
        verses.append(current_verse)

    slides = []
    for verse in verses:
        for i in range(0, len(verse), lines_per_slide):
            chunk = verse[i:i + lines_per_slide]
            if chunk:
                slides.append(chunk)

    return slides

def parse_part_marker(line):
    """Return the label for a '[Label]' / '#Label#' line ('' for '[]' / '##'), or None."""
    m = PART_MARKER_RE.match(str(line or ""))
    if not m:
        return None
    label = m.group(1) if m.group(1) is not None else m.group(2)
    return label.strip()


def split_text_into_parts(text: str) -> list[dict]:
    """
    Split lyrics on '#Label#' marker lines.
    Returns [{"label": str | None, "text": str}, ...] in order, skipping
    metadata parts (e.g. [Title]) and parts with no lyrics.
    """
    raw_parts = []
    current = {"label": None, "lines": []}

    for raw_line in str(text or "").splitlines():
        label = parse_part_marker(raw_line)
        if label is None:
            current["lines"].append(raw_line)
            continue
        raw_parts.append(current)
        current = {"label": label or None, "lines": []}
    raw_parts.append(current)

    parts = []
    for part in raw_parts:
        label = part["label"]
        if label and label.lower() in METADATA_PART_LABELS:
            continue
        if all(is_effectively_blank(line) for line in part["lines"]):
            continue
        parts.append({"label": label, "text": "\n".join(part["lines"])})
    return parts


def default_slide_parts(slides):
    return [{"label": None, "new_part": i == 0} for i in range(len(slides))]


def normalize_slide_parts(slides, slide_parts):
    if not slide_parts or len(slide_parts) != len(slides):
        return default_slide_parts(slides)
    return slide_parts


def get_current_slides_and_parts(text: str, smart_split=None):
    """
    Returns (slides, slide_parts):
      slides:      list[list[str]]
      slide_parts: list[{"label": str | None, "new_part": bool}], one per slide.
                   new_part is True on the first slide of each part occurrence.
    Every part starts on a new slide; splitting happens within each part.
    """
    # ALWAYS expand refrain first
    processed_text = expand_refrain_blocks(text)

    if smart_split is None:
        smart_split = st.session_state.get("smart_split_enabled", False)

    slides = []
    slide_parts = []

    for part in split_text_into_parts(processed_text):
        if smart_split:
            part_slides = split_slides_balanced(
                part["text"],
                max_chars=st.session_state.get("smart_split_max_chars", 30),
                max_lines_per_slide=st.session_state.get("smart_split_max_lines_per_slide", 4),
            )
        else:
            part_slides = split_slides_manual(part["text"])

        for j, slide_lines in enumerate(part_slides):
            slides.append(slide_lines)
            slide_parts.append({"label": part["label"], "new_part": j == 0})

    return slides, slide_parts


def get_current_slides(text: str) -> list[list[str]]:
    return get_current_slides_and_parts(text)[0]


def song_slides_to_editor_text(slides, slide_parts=None) -> str:
    """Rebuild editor text from stored slides, re-inserting '#Label#' markers."""
    slide_parts = normalize_slide_parts(slides, slide_parts)
    chunks = []
    prev_label = None
    for slide_lines, part in zip(slides, slide_parts):
        body = "\n".join(slide_lines)
        label = part.get("label")
        if part.get("new_part") and label:
            body = f"[{label}]\n{body}"
        elif part.get("new_part") and prev_label:
            body = f"[]\n{body}"  # "[]" ends the previous labelled part
        prev_label = label
        chunks.append(body)
    return "\n\n".join(chunks)

def open_presentation_from_bytes(template_bytes: bytes):
    return Presentation(BytesIO(template_bytes))


def get_layout_by_name(prs, layout_name):
    for slide_master in prs.slide_masters:
        for slide_layout in slide_master.slide_layouts:
            if slide_layout.name.strip() == layout_name:
                return slide_layout
    return None


def find_layout_placeholder_idx(layout, name):
    for ph in layout.placeholders:
        if getattr(ph, "name", "").strip().upper() == name:
            return ph.placeholder_format.idx
    return None


def get_part_footer_idxs(layout):
    return {
        "label": find_layout_placeholder_idx(layout, PART_LABEL_PLACEHOLDER_NAME),
        "nav": find_layout_placeholder_idx(layout, PART_NAV_PLACEHOLDER_NAME),
    }


def get_body_placeholder(slide):
    footer_idxs = {
        idx for idx in get_part_footer_idxs(slide.slide_layout).values() if idx is not None
    }
    placeholders = [
        ph for ph in slide.placeholders
        if ph.placeholder_format.idx not in footer_idxs
    ]

    for ph in placeholders:
        name = getattr(ph, "name", "").lower()
        if "title" not in name and getattr(ph, "has_text_frame", False):
            return ph

    if len(placeholders) > 1:
        return placeholders[1]
    if placeholders:
        return placeholders[0]
    return None


def set_shape_text(shape, text, font_size_pt=None, line_spacing=None):
    if shape is None or not getattr(shape, "has_text_frame", False):
        return

    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True

    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER

        if line_spacing is not None:
            p.line_spacing = line_spacing

        run = p.add_run()
        run.text = line

        if font_size_pt is not None:
            run.font.size = Pt(font_size_pt)


def delete_all_slides(prs):
    while len(prs.slides) > 0:
        slide_id = prs.slides._sldIdLst[0]
        rId = slide_id.rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]


def delete_slide_by_index(prs, slide_index: int):
    slide_id = prs.slides._sldIdLst[slide_index]
    rId = slide_id.rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[slide_index]

def apply_docx_heading_alias(heading: str) -> str:
    """
    Map common DOCX headings to template section names.
    This is a controlled alias layer (safe hardcoding).
    """
    simple = simplify_heading_text(heading)

    alias_map = {
        "closing Hymn": "response",
        "closing song": "response",
        "hymn of response": "response",
        "response hymn": "response",

        "offertory": "tithe offering",
        "offering": "tithe offering",

        "announcements": "announcements",
        "welcome announcements": "announcements",

        "doxology": "doxology",
        "gloria patri": "gloria patri",
    }

    return alias_map.get(simple, heading)

def get_flat_song_index_by_song_id(setlist, song_id):
    if not song_id:
        return 0

    for i, song in enumerate(setlist):
        if song.get("song_id") == song_id:
            return i
    return 0


def restore_selected_index_from_song_id():
    setlist = st.session_state.get("setlist", [])
    restored_index = get_flat_song_index_by_song_id(
        setlist,
        st.session_state.get("selected_song_id"),
    )

    if setlist:
        restored_index = max(0, min(restored_index, len(setlist) - 1))
    else:
        restored_index = 0

    st.session_state["setlist_selected_index"] = restored_index
    st.session_state["pending_setlist_selectbox_index"] = restored_index

def is_effectively_blank(line: str) -> bool:
    if line is None:
        return True
    line = str(line).replace("\xa0", " ")
    line = line.replace("\u200b", "")
    line = line.strip()
    return line == ""

def split_line_balanced_by_length(line: str, max_chars: int = 30) -> list[str]:
    line = str(line).strip()
    if not line:
        return []

    if len(line) <= max_chars:
        return [line]

    words = line.split()
    if len(words) <= 1:
        return [line]

    mid = len(words) // 2
    best_split = mid
    best_score = None

    for idx in range(max(1, mid - 2), min(len(words), mid + 3)):
        left = " ".join(words[:idx]).strip()
        right = " ".join(words[idx:]).strip()

        if not left or not right:
            continue

        score = abs(len(left) - len(right))

        if best_score is None or score < best_score:
            best_score = score
            best_split = idx

    left = " ".join(words[:best_split]).strip()
    right = " ".join(words[best_split:]).strip()

    return [left, right]


def split_slides_manual(text: str) -> list[list[str]]:
    if not text:
        return []

    lines = text.splitlines()
    slides = []
    current_slide = []

    for raw_line in lines:
        line = str(raw_line).replace("\xa0", " ").replace("\u200b", "")

        if is_effectively_blank(line):
            if current_slide:
                slides.append(current_slide)
                current_slide = []
        else:
            current_slide.append(line.strip())

    if current_slide:
        slides.append(current_slide)

    return slides


def split_slides_balanced(
    text: str,
    max_chars: int = 30,
    max_lines_per_slide: int = 4,
) -> list[list[str]]:
    if not text:
        return []

    slides = []
    current_slide = []

    for raw_line in text.splitlines():
        line = str(raw_line).replace("\xa0", " ").replace("\u200b", "").strip()

        # blank line = new slide
        if is_effectively_blank(line):
            if current_slide:
                slides.append(current_slide)
                current_slide = []
            continue

        split_lines = split_line_balanced_by_length(line, max_chars=max_chars)

        for split_line in split_lines:
            if len(current_slide) >= max_lines_per_slide:
                slides.append(current_slide)
                current_slide = []
            current_slide.append(split_line)

    if current_slide:
        slides.append(current_slide)

    return slides

def expand_refrain_blocks(text: str) -> str:
    import re

    if not text:
        return text

    lines = text.splitlines()

    refrain_start_patterns = [
        re.compile(r"^\s*refrain\s*:?\s*$", re.IGNORECASE),
        re.compile(r"^\s*chorus\s*:?\s*$", re.IGNORECASE),
    ]

    refrain_ref_patterns = [
        re.compile(r"^\s*\(?refrain\)?\s*$", re.IGNORECASE),
        re.compile(r"^\s*\(?chorus\)?\s*$", re.IGNORECASE),
    ]

    def is_blank(line: str) -> bool:
        return str(line).replace("\xa0", " ").strip() == ""

    refrain_lines = None
    refrain_label = "Refrain"
    output = []
    i = 0

    while i < len(lines):
        line = lines[i]

        # Detect "Refrain:" block
        if any(p.match(line) for p in refrain_start_patterns):
            refrain_label = "Chorus" if "chorus" in line.lower() else "Refrain"
            i += 1
            block = []

            while i < len(lines) and not is_blank(lines[i]):
                block.append(lines[i].strip())
                i += 1

            if block:
                refrain_lines = block
                # label it as a song part, so it gets a footer
                output.append(f"[{refrain_label}]")
                output.extend(block)
                output.append("[]")  # end the refrain part; following lines are unlabelled

            # preserve spacing
            if i < len(lines) and is_blank(lines[i]):
                output.append("")
            i += 1
            continue

        # Replace "(Refrain)" later
        if any(p.match(line) for p in refrain_ref_patterns):
            if refrain_lines:
                output.append(f"[{refrain_label}]")
                output.extend(refrain_lines)
                output.append("[]")
            else:
                output.append(line)
            i += 1
            continue

        output.append(line)
        i += 1

    return "\n".join(output)


# =========================================================
# SONG PART FOOTER HELPERS
# =========================================================
def _part_key(label):
    return str(label or "").strip().lower()


def build_part_nav(slide_parts):
    """
    Returns (nav_labels, starts_by_key):
      nav_labels:    display labels in nav order. Parts that occur once (verses) come first
                     in song order; repeated parts (chorus/refrain) go last.
      starts_by_key: key -> list of slide indices where each occurrence of that part starts.
    """
    starts_by_key = {}
    display_by_key = {}
    order = []

    for i, part in enumerate(slide_parts):
        label = part.get("label")
        if not label or not part.get("new_part"):
            continue
        key = _part_key(label)
        if key not in starts_by_key:
            starts_by_key[key] = []
            display_by_key[key] = label
            order.append(key)
        starts_by_key[key].append(i)

    singles = [k for k in order if len(starts_by_key[k]) == 1]
    repeats = [k for k in order if len(starts_by_key[k]) > 1]
    nav_labels = [display_by_key[k] for k in singles + repeats]
    return nav_labels, starts_by_key


def get_current_part_start(slide_parts, slide_index):
    for i in range(slide_index, -1, -1):
        if slide_parts[i].get("new_part"):
            return i
    return 0


def pick_part_link_target(starts, current_part_start):
    """
    Link to the next occurrence of the part at or after the current part
    (so jumping to the Chorus from Verse 2 lands on the chorus after Verse 2),
    else to the last occurrence.
    """
    for s in starts:
        if s >= current_part_start:
            return s
    return starts[-1]


def remove_shape(shape):
    if shape is not None:
        el = shape._element
        el.getparent().remove(el)


def get_slide_placeholder_by_idx(slide, idx):
    if idx is None:
        return None
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            return ph
    return None


def add_fallback_footer_textbox(prs, slide, kind):
    """Text box used when the layout has no PART_LABEL / PART_NAV placeholder."""
    slide_w = prs.slide_width
    slide_h = prs.slide_height
    top = int(slide_h * 0.9066)          # 6217623 EMU on 16:9, same row as the slide number
    height = int(slide_h * 0.0765)
    margin = int(slide_w * 0.0341)        # 415600 EMU on 16:9

    if kind == "label":
        left, width, align = margin, int(slide_w * 0.39), PP_ALIGN.LEFT
    else:
        # stop short of the slide number in the bottom-right corner
        right_edge = int(slide_w * 0.918)
        width = int(slide_w * 0.48)
        left, align = right_edge - width, PP_ALIGN.RIGHT

    box = slide.shapes.add_textbox(Emu(left), Emu(top), Emu(width), Emu(height))
    box.name = PART_LABEL_PLACEHOLDER_NAME if kind == "label" else PART_NAV_PLACEHOLDER_NAME
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.paragraphs[0].alignment = align
    return box, True


def get_footer_shape(prs, slide, kind):
    idx = get_part_footer_idxs(slide.slide_layout)[kind]
    ph = get_slide_placeholder_by_idx(slide, idx)
    if ph is not None:
        return ph, False
    return add_fallback_footer_textbox(prs, slide, kind)


def remove_footer_placeholders(slide):
    for idx in get_part_footer_idxs(slide.slide_layout).values():
        remove_shape(get_slide_placeholder_by_idx(slide, idx))


def link_run_to_slide(run, source_slide, target_slide):
    rId = source_slide.part.relate_to(target_slide.part, RT.SLIDE)
    rPr = run._r.get_or_add_rPr()
    hlink = etree.SubElement(rPr, qn("a:hlinkClick"))
    hlink.set(qn("r:id"), rId)
    hlink.set("action", "ppaction://hlinksldjump")


def apply_song_part_footers(prs, song_slides, slide_parts):
    """
    song_slides: the pptx slide objects of ONE song, in order.
    slide_parts: matching [{"label", "new_part"}] list.
    """
    slide_parts = normalize_slide_parts(song_slides, slide_parts)
    nav_labels, starts_by_key = build_part_nav(slide_parts)

    # Song has no part markers -> no footers at all
    if not nav_labels:
        for slide in song_slides:
            remove_footer_placeholders(slide)
        return

    for i, (slide, part) in enumerate(zip(song_slides, slide_parts)):
        current_key = _part_key(part.get("label"))
        current_part_start = get_current_part_start(slide_parts, i)

        # ---- left: current part label
        label_shape, is_fallback = get_footer_shape(prs, slide, "label")
        if part.get("label"):
            tf = label_shape.text_frame
            tf.clear()
            run = tf.paragraphs[0].add_run()
            run.text = part["label"]
            if is_fallback:
                tf.paragraphs[0].alignment = PP_ALIGN.LEFT
                run.font.size = Pt(PART_FOOTER_FONT_SIZE_PT)
        else:
            remove_shape(label_shape)

        # ---- right: links to the start of each part
        nav_shape, is_fallback = get_footer_shape(prs, slide, "nav")
        tf = nav_shape.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        if is_fallback:
            p.alignment = PP_ALIGN.RIGHT

        for n, label in enumerate(nav_labels):
            if n > 0:
                sep = p.add_run()
                sep.text = PART_NAV_SEPARATOR
                if is_fallback:
                    sep.font.size = Pt(PART_FOOTER_FONT_SIZE_PT)

            key = _part_key(label)
            target_idx = pick_part_link_target(starts_by_key[key], current_part_start)

            run = p.add_run()
            run.text = label
            if is_fallback:
                run.font.size = Pt(PART_FOOTER_FONT_SIZE_PT)
            if key == current_key:
                run.font.bold = True  # highlight the part we're in
            link_run_to_slide(run, slide, song_slides[target_idx])

# =========================================================
# MOVE HELPERS
# =========================================================
def get_live_index_by_slide_id(slides, slide_id: int) -> int:
    """
    Find the current (live) index of a slide using its slide_id.

    slides: ppt.prs.slides
    slide_id: int
    """
    for idx, slide in enumerate(slides):
        if slide.slide_id == slide_id:
            return idx
    raise ValueError(f"Slide id {slide_id} not found")

def get_slide_obj_by_slide_id(slides, slide_id: int):
    for slide in slides:
        if slide.slide_id == slide_id:
            return slide
    raise ValueError(f"Slide id {slide_id} not found")


def build_section_anchor_map(prs):
    """
    Build anchors from the real template-preserved deck BEFORE reordering.

    Returns:
        dict[normalized_section_title] = {
            "divider_slide_id": int,
            "next_divider_slide_id": int | None
        }
    """
    divider_entries = []

    for slide in prs.slides:
        if is_divider_slide(slide):
            # same fallback naming as parse_template_sections(), so untitled
            # dividers ("Section 1", "Section 2", ...) still match
            raw_title = get_divider_title(
                slide, fallback=f"Section {len(divider_entries) + 1}"
            )
            norm_title = canonicalize_section_label(raw_title)
            divider_entries.append(
                {
                    "raw_title": raw_title,
                    "norm_title": norm_title,
                    "slide_id": slide.slide_id,
                }
            )

    anchor_map = {}
    for i, entry in enumerate(divider_entries):
        next_slide_id = None
        if i + 1 < len(divider_entries):
            next_slide_id = divider_entries[i + 1]["slide_id"]

        anchor_map[entry["norm_title"]] = {
            "divider_slide_id": entry["slide_id"],
            "next_divider_slide_id": next_slide_id,
            "raw_title": entry["raw_title"],
        }

    return anchor_map


def find_section_insert_index_from_anchor_map(slides, section_title: str, anchor_map: dict) -> int:
    target_norm = canonicalize_section_label(section_title)
    anchor = anchor_map.get(target_norm)

    if anchor is None:
        available = [v["raw_title"] for v in anchor_map.values()]
        raise ValueError(
            f"Could not find divider for section: {section_title!r}. "
            f"Available dividers: {available}"
        )

    next_divider_slide_id = anchor["next_divider_slide_id"]

    # last section -> insert at end
    if next_divider_slide_id is None:
        return len(slides)

    return get_live_index_by_slide_id(slides, next_divider_slide_id)

# =========================================================
# TEMPLATE / SECTION HELPERS
# =========================================================
def get_slide_layout_name(slide):
    try:
        return slide.slide_layout.name.strip()
    except Exception:
        return ""


def is_divider_slide(slide):
    return get_slide_layout_name(slide) == DIVIDER_LAYOUT_NAME


def get_divider_title(slide, fallback="Untitled Section"):
    if slide.shapes.title and slide.shapes.title.text.strip():
        return slide.shapes.title.text.strip()

    texts = []
    for shape in slide.shapes:
        if hasattr(shape, "text"):
            text = shape.text.strip()
            if text:
                texts.append(text)

    return texts[0] if texts else fallback


def parse_template_sections(template_bytes: bytes):
    prs = open_presentation_from_bytes(template_bytes)

    sections = []
    current_section = None
    section_counter = 0

    for i, slide in enumerate(prs.slides):
        if is_divider_slide(slide):
            current_section = {
                "id": f"sec_{section_counter}",
                "title": get_divider_title(slide, fallback=f"Section {section_counter + 1}"),
                "divider_index": i,
                "content_slide_indices": [],
            }
            sections.append(current_section)
            section_counter += 1
        else:
            if current_section is not None:
                current_section["content_slide_indices"].append(i)

    return sections


def get_section_title_by_id(section_id):
    for sec in st.session_state.get("template_sections", []):
        if sec["id"] == section_id:
            return sec["title"]
    return None

def canonicalize_section_label(s: str) -> str:
    s = str(s or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def validate_template_bytes(template_bytes: bytes):
    prs = open_presentation_from_bytes(template_bytes)

    errors = []
    warnings = []

    first_layout = get_layout_by_name(prs, FIRST_LAYOUT_NAME)
    rest_layout = get_layout_by_name(prs, REST_LAYOUT_NAME)

    if first_layout is None:
        errors.append(f"Missing layout: {FIRST_LAYOUT_NAME}")
    if rest_layout is None:
        errors.append(f"Missing layout: {REST_LAYOUT_NAME}")

    if errors:
        return False, errors, warnings

    first_slide = prs.slides.add_slide(first_layout)
    rest_slide = prs.slides.add_slide(rest_layout)

    title_placeholder = first_slide.shapes.title
    first_body_placeholder = get_body_placeholder(first_slide)
    rest_body_placeholder = get_body_placeholder(rest_slide)

    if title_placeholder is None:
        errors.append(f"{FIRST_LAYOUT_NAME} is missing a title placeholder")
    if first_body_placeholder is None:
        errors.append(f"{FIRST_LAYOUT_NAME} is missing a body/lyrics placeholder")
    if rest_body_placeholder is None:
        errors.append(f"{REST_LAYOUT_NAME} is missing a body/lyrics placeholder")

    for layout in (first_layout, rest_layout):
        missing = [
            name
            for name, idx in (
                (PART_LABEL_PLACEHOLDER_NAME, get_part_footer_idxs(layout)["label"]),
                (PART_NAV_PLACEHOLDER_NAME, get_part_footer_idxs(layout)["nav"]),
            )
            if idx is None
        ]
        if missing:
            warnings.append(
                f"{layout.name.strip()} has no {' / '.join(missing)} placeholder; "
                "song-part footers will use default text boxes."
            )

    return len(errors) == 0, errors, warnings


# =========================================================
# SONG / SECTION HELPERS
# =========================================================
def build_editor_song_item(current_slides, slide_parts=None):
    return {
        "umh_number": st.session_state["editor_umh"].strip(),
        "title": st.session_state["editor_title"].strip(),
        "slides": current_slides,
        "slide_parts": normalize_slide_parts(current_slides, slide_parts),
        "lyrics_font_size_pt": (
            st.session_state["editor_lyrics_font_size_pt"]
            if st.session_state["editor_override_lyrics_font_size"]
            else None
        ),
        "line_spacing": (
            st.session_state["editor_line_spacing"]
            if st.session_state["editor_override_line_spacing"]
            else None
        ),
        "override_lyrics_font_size": st.session_state["editor_override_lyrics_font_size"],
        "override_line_spacing": st.session_state["editor_override_line_spacing"],
        "section_id": st.session_state.get("selected_song_section_id"),
    }


def build_current_song_signature(song_item, selected_template_name):
    return (
        song_item["umh_number"],
        song_item["title"],
        tuple(tuple(slide) for slide in song_item["slides"]),
        tuple(
            (p.get("label"), p.get("new_part"))
            for p in normalize_slide_parts(song_item["slides"], song_item.get("slide_parts"))
        ),
        song_item.get("lyrics_font_size_pt"),
        song_item.get("line_spacing"),
        song_item.get("section_id"),
        selected_template_name
    )


def get_ordered_songs_for_output(setlist):
    sections = st.session_state.get("template_sections", [])
    songs_by_section = {sec["id"]: [] for sec in sections}
    unassigned_songs = []

    for idx, song in enumerate(setlist):
        sec_id = song.get("section_id")
        if sec_id in songs_by_section:
            songs_by_section[sec_id].append((idx, song))
        else:
            unassigned_songs.append((idx, song))

    ordered = []
    for sec in sections:
        ordered.extend(songs_by_section.get(sec["id"], []))
    ordered.extend(unassigned_songs)
    return ordered


def group_songs_by_section_order(setlist):
    sections = st.session_state.get("template_sections", [])
    songs_by_section = {sec["id"]: [] for sec in sections}
    unassigned = []

    for idx, song in enumerate(setlist):
        sec_id = song.get("section_id")
        if sec_id in songs_by_section:
            songs_by_section[sec_id].append((idx, song))
        else:
            unassigned.append((idx, song))

    ordered_groups = []
    for sec in sections:
        section_songs = songs_by_section.get(sec["id"], [])
        if section_songs:
            ordered_groups.append((sec["id"], sec["title"], section_songs))

    if unassigned:
        ordered_groups.append((None, None, unassigned))

    return ordered_groups




def format_song_label(song, idx=None):
    if song.get("umh_number"):
        base = f'UMH {song["umh_number"]} {song["title"]}'.strip()
    else:
        base = song.get("title", "")
    if idx is not None:
        return f"{idx+1}. {base}"
    return base


def create_empty_service_order_blocks_from_template(template_sections):
    return [
        {
            "block_id": f"block_{i}",
            "section_id": sec["id"],
            "section_title": sec["title"],
            "source_heading": sec["title"],
            "match_type": "template",
            "score": 100,
            "items": [],
        }
        for i, sec in enumerate(template_sections)
    ]


def pick_default_service_section(template_sections):
    if not template_sections:
        return None

    skip_canonicals = {
        "welcome to mci",
        "welcome",
        "announcements",
    }

    preferred_canonicals = {
        "call to worship",
        "scripture reading",
        "corporate prayer",
        "sermon",
        "response",
        "benediction",
        "tithe offering",
        "tithes offerings",
        "doxology",
    }

    for sec in template_sections:
        canon = canonicalize_section_label(sec["title"])
        if canon in preferred_canonicals:
            return sec

    for sec in template_sections:
        canon = canonicalize_section_label(sec["title"])
        if canon not in skip_canonicals:
            return sec

    return template_sections[0]


def flatten_blocks_to_setlist(service_order_blocks, song_store):
    flat = []
    for block in service_order_blocks:
        for item in block.get("items", []):
            if item.get("type") != "song":
                continue
            song = song_store.get(item.get("song_id"))
            if not song:
                continue
            flat.append({
                "umh_number": song.get("umh_number", ""),
                "title": song.get("title", ""),
                "slides": song.get("slides", []),
                "slide_parts": song.get("slide_parts"),
                "lyrics_font_size_pt": song.get("lyrics_font_size_pt"),
                "line_spacing": song.get("line_spacing"),
                "override_lyrics_font_size": song.get("override_lyrics_font_size", False),
                "override_line_spacing": song.get("override_line_spacing", False),
                "section_id": block.get("section_id"),
                "song_id": song.get("song_id"),
                "service_block_id": block.get("block_id"),
                "service_block_title": block.get("section_title"),
            })
    return flat


def sync_block_model_from_setlist():
    template_sections = st.session_state.get("template_sections", []) or []
    if not template_sections:
        return

    current_blocks = st.session_state.get("service_order_blocks", []) or []
    song_store = st.session_state.get("song_store", {}) or {}
    setlist = st.session_state.get("setlist", []) or []

    if not current_blocks:
        current_blocks = create_empty_service_order_blocks_from_template(template_sections)

    # preserve non-song items from existing blocks, but rebuild song nesting from setlist
    block_lookup = {}
    new_blocks = []
    for sec in template_sections:
        existing = next((b for b in current_blocks if b.get("section_id") == sec["id"]), None)
        if existing:
            items = [it for it in existing.get("items", []) if it.get("type") != "song"]
            block = {**existing, "section_title": sec["title"], "items": items}
        else:
            block = {
                "block_id": f"block_{len(new_blocks)}",
                "section_id": sec["id"],
                "section_title": sec["title"],
                "source_heading": sec["title"],
                "match_type": "template",
                "score": 100,
                "items": [],
            }
        new_blocks.append(block)
        block_lookup[sec["id"]] = block

    unassigned_items = []
    for idx, song in enumerate(setlist):
        song_id = song.get("song_id") or f"setlist_song_{idx}"
        song_store[song_id] = {
            "song_id": song_id,
            "umh_number": song.get("umh_number", ""),
            "title": song.get("title", ""),
            "slides": song.get("slides", []),
            "slide_parts": song.get("slide_parts"),
            "lyrics_font_size_pt": song.get("lyrics_font_size_pt"),
            "line_spacing": song.get("line_spacing"),
            "override_lyrics_font_size": song.get("override_lyrics_font_size", False),
            "override_line_spacing": song.get("override_line_spacing", False),
        }
        sec_id = song.get("section_id")
        song_item = {"type": "song", "song_id": song_id}
        if sec_id in block_lookup:
            block_lookup[sec_id]["items"].append(song_item)
        else:
            unassigned_items.append(song_item)

    if unassigned_items:
        new_blocks.append({
            "block_id": "block_unassigned",
            "section_id": None,
            "section_title": "Unassigned Songs",
            "source_heading": "Unassigned Songs",
            "match_type": "derived",
            "score": 100,
            "items": unassigned_items,
        })

    st.session_state["service_order_blocks"] = new_blocks
    st.session_state["song_store"] = song_store


def build_template_service_order_view(setlist=None):
    sync_block_model_from_setlist()
    service_order_blocks = st.session_state.get("service_order_blocks", []) or []
    song_store = st.session_state.get("song_store", {}) or {}
    setlist = st.session_state.get("setlist", []) or []

    index_by_song_id = {}
    for i, song in enumerate(setlist):
        song_id = song.get("song_id")
        if song_id:
            index_by_song_id[song_id] = i

    groups = []
    for block in service_order_blocks:
        items = []
        for item in block.get("items", []):
            if item.get("type") != "song":
                continue
            song = song_store.get(item.get("song_id"))
            if not song:
                continue
            items.append({
                "type": "song",
                "index": index_by_song_id.get(item.get("song_id")),
                "song": song,
            })
        groups.append({
            "section_id": block.get("section_id"),
            "section_title": block.get("section_title"),
            "items": items,
        })
    return groups


def render_service_group_items_markdown(group, selected_index=None, editing_index=None):
    st.markdown(f"**{group['section_title']}**")
    for item in group.get("items", []):
        i = item.get("index")
        song = item.get("song")
        if not song:
            continue
        label = format_song_label(song, i) if i is not None else format_song_label(song)
        prefix = ""
        if selected_index is not None and i == selected_index:
            prefix += "🔹 "
        if editing_index is not None and i == editing_index:
            prefix += "✏️ "
        if prefix:
            st.markdown(f"&nbsp;&nbsp;**{prefix}{label}**", unsafe_allow_html=True)
        else:
            st.markdown(f"&nbsp;&nbsp;{label}", unsafe_allow_html=True)


def render_service_group_items_caption(group):
    st.markdown(f"**{group['section_title']}**")
    for item in group.get("items", []):
        song = item.get("song")
        if not song:
            continue
        st.caption(f"{format_song_label(song)} ({len(song['slides'])} slide(s))")


def reset_editor():
    st.session_state["loaded_song"] = None
    st.session_state["editing_setlist_index"] = None
    st.session_state["editor_umh"] = ""
    st.session_state["editor_title"] = ""
    st.session_state["editor_text"] = ""
    st.session_state["editor_override_lyrics_font_size"] = False
    st.session_state["editor_override_line_spacing"] = False
    st.session_state["editor_lyrics_font_size_pt"] = 32
    st.session_state["editor_line_spacing"] = 1.2
    st.session_state["current_song_preview_images"] = None
    st.session_state["current_song_preview_stats"] = None
    st.session_state["last_editor_text"] = ""
    st.session_state["last_current_song_signature"] = None
    st.session_state["editor_status_message"] = ""
    st.session_state["current_preview_slide"] = 1
    st.session_state["editor_ace_key"] += 1
    st.session_state["selected_song_section_id"] = None


def load_song_preview_if_possible():
    if (
        st.session_state.get("selected_template_name")
        and st.session_state["selected_template_name"] in st.session_state["uploaded_templates"]
        and soffice_available()
    ):
        try:
            template_bytes = st.session_state["uploaded_templates"][
                st.session_state["selected_template_name"]
            ]
            template_ok, _, _ = validate_template_bytes(template_bytes)

            if template_ok:
                current_slides, current_slide_parts = get_current_slides_and_parts(
                    st.session_state["editor_text"]
                )
                if current_slides:
                    song_item = build_editor_song_item(current_slides, current_slide_parts)
                    refresh_current_song_preview(song_item, template_bytes)
                    st.session_state["current_preview_slide"] = 1
                    st.session_state["preview_mode"] = "song"
                    st.session_state["editor_status_message"] = "Song preview loaded."
                else:
                    st.session_state["editor_status_message"] = "Song loaded, but no slides detected for preview."
            else:
                st.session_state["editor_status_message"] = "Song loaded, but selected template is invalid."
        except Exception as e:
            st.session_state["editor_status_message"] = preview_error_message(e)
    else:
        if not st.session_state.get("selected_template_name"):
            st.session_state["editor_status_message"] = "Song loaded. Please select a template to generate preview."
        elif not soffice_available():
            st.session_state["editor_status_message"] = "Song loaded. LibreOffice/soffice is not available."


def load_song_into_editor(match):
    lyrics_raw = str(match.get("Lyrics (Raw)", "")).strip()

    st.session_state["loaded_song"] = {
        "umh_number": str(match.get("UMH Number", "")).strip(),
        "title": str(match.get("Title", "")).strip(),
        "lyrics_raw": lyrics_raw,
    }

    st.session_state["editing_setlist_index"] = None
    st.session_state["editor_umh"] = str(match.get("UMH Number", "")).strip()
    st.session_state["editor_title"] = str(match.get("Title", "")).strip()
    st.session_state["editor_text"] = lyrics_raw

    st.session_state["editor_override_lyrics_font_size"] = False
    st.session_state["editor_override_line_spacing"] = False
    st.session_state["editor_lyrics_font_size_pt"] = 32
    st.session_state["editor_line_spacing"] = 1.2

    sections = st.session_state.get("template_sections", [])
    st.session_state["selected_song_section_id"] = sections[0]["id"] if sections else None

    st.session_state["current_song_preview_images"] = None
    st.session_state["current_song_preview_stats"] = None
    st.session_state["last_editor_text"] = lyrics_raw
    st.session_state["last_current_song_signature"] = None
    st.session_state["editor_status_message"] = ""
    st.session_state["current_preview_slide"] = 1
    st.session_state["preview_mode"] = "song"
    st.session_state["editor_ace_key"] += 1

    load_song_preview_if_possible()


def apply_pending_setlist_load():
    idx = st.session_state.get("pending_setlist_load")
    if idx is None:
        return

    if idx >= len(st.session_state["setlist"]):
        st.session_state["pending_setlist_load"] = None
        return

    item = st.session_state["setlist"][idx]
    lyrics_text = song_slides_to_editor_text(item["slides"], item.get("slide_parts"))

    st.session_state["loaded_song"] = {
        "umh_number": item["umh_number"],
        "title": item["title"],
        "lyrics_raw": lyrics_text,
    }

    st.session_state["editing_setlist_index"] = idx
    st.session_state["editor_umh"] = item["umh_number"]
    st.session_state["editor_title"] = item["title"]
    st.session_state["editor_text"] = lyrics_text

    st.session_state["editor_override_lyrics_font_size"] = item.get(
        "override_lyrics_font_size", False
    )
    st.session_state["editor_override_line_spacing"] = item.get(
        "override_line_spacing", False
    )
    st.session_state["editor_lyrics_font_size_pt"] = item.get("lyrics_font_size_pt", 32) or 32
    st.session_state["editor_line_spacing"] = item.get("line_spacing", 1.2) or 1.2
    st.session_state["selected_song_section_id"] = item.get("section_id")

    if (
        st.session_state["selected_song_section_id"] is None
        and st.session_state.get("template_sections")
    ):
        st.session_state["selected_song_section_id"] = st.session_state["template_sections"][0]["id"]

    st.session_state["current_song_preview_images"] = None
    st.session_state["current_song_preview_stats"] = None
    st.session_state["pending_setlist_load"] = None
    st.session_state["last_editor_text"] = lyrics_text
    st.session_state["last_current_song_signature"] = None
    st.session_state["editor_status_message"] = ""
    st.session_state["current_preview_slide"] = 1
    st.session_state["preview_mode"] = "song"
    st.session_state["editor_ace_key"] += 1

    load_song_preview_if_possible()



# =========================================================
# ORDER OF SERVICE DOCX IMPORT
# =========================================================
UMH_IMPORT_RE = re.compile(
    # number is optional: "UMH 140 Title" or "UMH Title"
    r"^(?P<hymnal>UMH)\s*(?P<number>\d+)?\.?\s+(?P<title>.+?)(?:\s*\((?P<stanzas>.+?)\))?\s*$",
    re.IGNORECASE,
)

DOCX_SECTION_ALIASES = {
    "songs of praise": [
        "songs of praise",
        "hymns of praise",
        "opening song",
        "opening songs",
        "opening hymn",
        "opening hymns",
        "praise songs",
        "praise and worship",
        "worship songs",
    ],
    "gloria patri": ["gloria patri", "glory be to the father"],
    "scripture reading": ["scripture reading", "bible reading", "scripture"],
    "sermon": ["sermon", "message", "sermon title", "the message"],
    "call to worship": ["call to worship"],
    "tithes & offerings": [
        "tithes & offerings",
        "tithes and offerings",
        "offertory",
        "offering",
        "offering prayer",
    ],
    "response": [
        "response",
        "closing song",
        "closing songs",
        "closing hymn",
        "closing hymns",
        "hymn of response",
        "response song",
        "response hymn",
        "sending song",
        "closing response",
    ],
    "benediction": ["benediction", "closing blessing"],
}


def normalize_text(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip()).lower()



def simplify_heading_text(s: str) -> str:
    """
    Normalize headings so variations like:
    - 'CALL TO\u000bWORSHIP'
    - 'Call to Worship (based on Ps 130:5-7)'
    all become:
    - 'call to worship'
    """
    s = str(s or "")

    # Remove vertical tabs from PPT titles
    s = s.replace("\u000b", " ")

    # Remove parenthetical notes
    s = re.sub(r"\([^)]*\)", " ", s)

    # Normalize
    s = s.lower().strip()

    # Normalize symbols
    s = s.replace("&", " and ")

    # "Section1" -> "section 1"
    s = re.sub(r"([a-z])(\d)", r"\1 \2", s)
    s = re.sub(r"(\d)([a-z])", r"\1 \2", s)

    # Remove punctuation
    s = re.sub(r"[^a-z0-9 ]+", " ", s)

    # Collapse whitespace
    s = re.sub(r"\s+", " ", s).strip()

    return s


def canonicalize_section_label(label: str) -> str:
    simplified = simplify_heading_text(label)
    if not simplified:
        return ""

    for canonical, aliases in DOCX_SECTION_ALIASES.items():
        alias_pool = [canonical] + aliases
        simplified_aliases = [simplify_heading_text(a) for a in alias_pool]
        if simplified in simplified_aliases:
            return canonical
        for alias in simplified_aliases:
            if alias and (simplified == alias or simplified in alias or alias in simplified):
                return canonical

    return simplified

def read_docx_lines(docx_file) -> list[str]:
    """
    One entry per line. Blank paragraphs are kept as "" (they separate verses),
    and soft line breaks (Shift+Enter) inside a paragraph become separate lines.
    """
    doc = Document(docx_file)
    lines = []
    for para in doc.paragraphs:
        text = para.text.replace("\xa0", " ").replace("\u000b", "\n")
        for line in text.split("\n"):
            lines.append(line.strip())
    return lines



def is_umh_song_line(line: str) -> bool:
    return UMH_IMPORT_RE.match(line.strip()) is not None



def parse_umh_song_line(line: str) -> dict:
    m = UMH_IMPORT_RE.match(line.strip())
    if not m:
        return {
            "umh_number": "",
            "title": line.strip(),
            "stanzas": "",
            "raw": line.strip(),
        }

    return {
        "umh_number": (m.group("number") or "").strip(),
        "title": (m.group("title") or "").strip(),
        "stanzas": (m.group("stanzas") or "").strip(),
        "raw": line.strip(),
    }



def match_template_section_from_heading(heading: str, template_sections: list[dict]) -> dict | None:
    """
    Robust heading-to-template matcher.

    Handles cases like:
    - "Call to Worship (based on Ps 130:5-7)"
    - "CALL TO\u000bWORSHIP"
    """
    if not template_sections or not heading:
        return None

    heading = apply_docx_heading_alias(heading)
    heading_simple = simplify_heading_text(heading)
    
    if not heading_simple:
        return None

    best = None

    for sec in template_sections:
        template_simple = simplify_heading_text(sec["title"])

        score = 0
        match_type = "none"

        if heading_simple == template_simple:
            score = 100
            match_type = "exact"

        elif heading_simple.startswith(template_simple) or template_simple.startswith(heading_simple):
            score = 98
            match_type = "alias"

        elif heading_simple in template_simple or template_simple in heading_simple:
            score = 95
            match_type = "contains"

        else:
            heading_tokens = set(heading_simple.split())
            template_tokens = set(template_simple.split())
            overlap = heading_tokens & template_tokens

            if heading_tokens and template_tokens and overlap:
                ratio = len(overlap) / max(1, min(len(heading_tokens), len(template_tokens)))
                if ratio >= 0.75:
                    score = 90
                    match_type = "alias"
                elif ratio >= 0.5:
                    score = 75
                    match_type = "token-overlap"

        candidate = {
            "section_id": sec["id"],
            "section_title": sec["title"],
            "score": score,
            "match_type": match_type,
            "docx_heading": heading,
        }

        if best is None or candidate["score"] > best["score"]:
            best = candidate

    if best and best["score"] >= 90:
        return best

    return None

def import_service_order_from_docx(docx_file, template_sections):
    """
    Template-driven block import.

    Returns:
      service_order_blocks: list[dict]
      song_store: dict[str, dict]
      missing_songs: list[dict]
      section_mapping_rows: list[dict]
    """
    lines = read_docx_lines(docx_file)

    service_order_blocks = create_empty_service_order_blocks_from_template(template_sections)
    block_by_section_id = {block["section_id"]: block for block in service_order_blocks}
    song_store = {}
    missing_songs = []
    section_mapping_rows = []

    current_block = None
    song_seq = 0

    def pick_starting_block():
        default_sec = pick_default_service_section(template_sections)
        if default_sec:
            return block_by_section_id.get(default_sec["id"])
        return service_order_blocks[0] if service_order_blocks else None

    current_block = pick_starting_block()

    # A song line opens a "pending" song. Numbered songs ("UMH 140 Title") get their
    # lyrics from Google Sheets. Un-numbered songs ("UMH Title") take the lyrics that
    # follow in the DOCX, up to the next song line / section heading / end of file.
    pending_song = None

    def finish_pending_song():
        nonlocal song_seq, pending_song
        if pending_song is None:
            return

        parsed = pending_song["parsed"]
        block = pending_song["block"]
        inline_lyrics = "\n".join(pending_song["lines"]).strip()
        pending_song = None

        if inline_lyrics:
            # lyrics come from the DOCX; borrow the UMH number from Sheets if the title matches
            title_row = find_row_by_title(parsed["title"]) if not parsed["umh_number"] else None
            umh_number = parsed["umh_number"] or (
                str(title_row.get("UMH Number", "")).strip() if title_row else ""
            )
            title = parsed["title"]
            lyrics = inline_lyrics
        else:
            row = (
                find_row_by_umh(parsed["umh_number"])
                if parsed["umh_number"]
                else find_row_by_title(parsed["title"])
            )
            if not row:
                missing_songs.append(parsed)
                return
            umh_number = str(row.get("UMH Number", "")).strip()
            title = str(row.get("Title", "")).strip()
            lyrics = str(row.get("Lyrics (Raw)", "")).strip()

        imported_slides, imported_slide_parts = get_current_slides_and_parts(
            lyrics,
            smart_split=False,
        )
        if not imported_slides:
            missing_songs.append(parsed)
            return

        song_seq += 1
        song_id = f"song_{song_seq}"
        song_store[song_id] = {
            "song_id": song_id,
            "umh_number": umh_number,
            "title": title,
            "slides": imported_slides,
            "slide_parts": imported_slide_parts,
            "lyrics_font_size_pt": None,
            "line_spacing": None,
            "override_lyrics_font_size": False,
            "override_line_spacing": False,
        }
        if block is not None:
            block["items"].append({"type": "song", "song_id": song_id})

    awaiting_title = False  # set after a "[Title]" / "#Title#" line; the next line is the song title

    for raw_line in lines:
        stripped = raw_line.strip()
        collecting = pending_song is not None and pending_song["collect_lyrics"]

        if not stripped:
            if collecting:
                pending_song["lines"].append("")  # blank paragraph = slide break
            continue

        # "[Title]" (or "#Title#") starts a new song (same format as the song editor);
        # the next line is its title and the lines after it are its lyrics
        marker_label = parse_part_marker(stripped)
        if marker_label is not None and marker_label.lower() in METADATA_PART_LABELS:
            finish_pending_song()
            awaiting_title = True
            continue

        if awaiting_title:
            awaiting_title = False
            if is_umh_song_line(stripped):
                parsed = parse_umh_song_line(stripped)
            else:
                parsed = {"umh_number": "", "title": stripped, "stanzas": "", "raw": stripped}
            pending_song = {
                "parsed": parsed,
                "lines": [],
                "block": current_block,
                "collect_lyrics": True,  # lyrics follow, even if the title has a UMH number
            }
            continue

        # "[Verse 1]" / "#Verse 1#" part markers inside lyrics are lyrics, not "#" headings
        if collecting and parse_part_marker(stripped) is not None:
            pending_song["lines"].append(stripped)
            continue

        is_prefixed_heading = stripped.startswith(("+", "#"))
        heading_text = stripped[1:].strip() if is_prefixed_heading else stripped

        # apply alias BEFORE matching
        aliased_heading = apply_docx_heading_alias(heading_text)

        heading_match = match_template_section_from_heading(
            aliased_heading,
            template_sections,
        )

        # While reading lyrics, only an exact heading match ends the song,
        # so lyric lines don't get mistaken for loosely-matching headings.
        strong_match_types = {"exact"} if collecting else {"exact", "alias", "contains"}
        matched_block = None
        if heading_match and heading_match["match_type"] in strong_match_types:
            matched_block = block_by_section_id.get(heading_match["section_id"])

        is_song_line = is_umh_song_line(stripped)

        if collecting and matched_block is None and not is_prefixed_heading and not is_song_line:
            pending_song["lines"].append(stripped)
            continue

        finish_pending_song()

        # allow plain headings OR prefixed headings like "# Closing Hymn"
        # to switch sections if they strongly match
        if matched_block is not None:
            current_block = matched_block
            section_mapping_rows.append({
                "docx_heading": heading_text,
                "mapped_section_id": heading_match["section_id"],
                "mapped_section_title": heading_match["section_title"],
                "match_type": heading_match["match_type"],
                "score": heading_match["score"],
                "source": "prefixed" if is_prefixed_heading else "plain",
            })
            continue

        if is_prefixed_heading:
            if current_block is not None:
                current_block["items"].append({"type": "minor_heading", "text": heading_text})
            continue

        if is_song_line:
            parsed = parse_umh_song_line(stripped)
            pending_song = {
                "parsed": parsed,
                "lines": [],
                "block": current_block,
                "collect_lyrics": not parsed["umh_number"],
            }
            continue

        if current_block is not None:
            current_block["items"].append({"type": "text", "text": stripped})

    finish_pending_song()

    return service_order_blocks, song_store, missing_songs, section_mapping_rows


# =========================================================
# PPT BUILD HELPERS
# =========================================================
def add_song_block_to_prs(prs, song, first_layout, rest_layout):
    start_idx = len(prs.slides)

    umh_number = str(song["umh_number"]).strip()
    title = str(song["title"]).strip()
    slides = song["slides"]

    lyrics_font_size_pt = song.get("lyrics_font_size_pt")
    line_spacing = song.get("line_spacing")

    full_title = f"UMH {umh_number} {title}".strip() if umh_number else title

    song_slide_objs = []

    for i, slide_lines in enumerate(slides):
        lyrics_text = "\n".join(slide_lines)

        if i == 0:
            slide = prs.slides.add_slide(first_layout)
            set_shape_text(slide.shapes.title, full_title)
            set_shape_text(
                get_body_placeholder(slide),
                lyrics_text,
                font_size_pt=lyrics_font_size_pt,
                line_spacing=line_spacing,
            )
        else:
            slide = prs.slides.add_slide(rest_layout)
            set_shape_text(
                get_body_placeholder(slide),
                lyrics_text,
                font_size_pt=lyrics_font_size_pt,
                line_spacing=line_spacing,
            )

        song_slide_objs.append(slide)

    # Footers go on after all of the song's slides exist, so links have targets
    apply_song_part_footers(prs, song_slide_objs, song.get("slide_parts"))

    end_idx = len(prs.slides) - 1
    return start_idx, end_idx


def add_section_song_block_to_prs(prs, section_song_pairs, first_layout, rest_layout):
    start_idx = len(prs.slides)
    for _, song in section_song_pairs:
        add_song_block_to_prs(prs, song, first_layout, rest_layout)
    end_idx = len(prs.slides) - 1
    return start_idx, end_idx

class ExistingDeckTemplate:
    def __init__(
        self,
        pptx_path: str,
        title_layout_name=None,
        default_layout_name=None,
        divider_layout_name=None,
    ):
        self.template = pptx_path
        self.prs = Presentation(pptx_path)

        self.title_layout = self._get_layout(title_layout_name, fallback_index=0)
        self.default_layout = self._get_layout(default_layout_name, fallback_index=1)
        self.divider_layout = self._get_layout(divider_layout_name, fallback_index=0)

    def _get_layout(self, layout_name, fallback_index=0):
        if layout_name:
            for layout in self.prs.slide_layouts:
                try:
                    if layout.name == layout_name:
                        return layout
                except Exception:
                    pass
        return self.prs.slide_layouts[fallback_index]

def create_combined_ppt(setlist, template_bytes: bytes):
    """
    Rewritten to avoid direct _sldIdLst reordering inside create_combined_ppt().
    Uses PPTXCreator.move_slide() for final slide moves.

    Keeps your current behavior:
    - songs mode => songs only
    - preserve_template_slides=False => songs only
    - preserve_template_slides=True => keep template, optionally hide section content,
      then insert songs under each matching section
    """
    
    sync_block_model_from_setlist()
    service_order_blocks = st.session_state.get("service_order_blocks", []) or []
    song_store = st.session_state.get("song_store", {}) or {}

    prs = open_presentation_from_bytes(template_bytes)

    first_layout = get_layout_by_name(prs, FIRST_LAYOUT_NAME)
    rest_layout = get_layout_by_name(prs, REST_LAYOUT_NAME)

    if first_layout is None or rest_layout is None:
        raise ValueError("Template layouts not found.")

    output_mode = st.session_state.get("service_output_mode", "full")
    preserve_template_slides = st.session_state.get("preserve_template_slides", True)

    ordered_song_pairs = []
    for block in service_order_blocks:
        for item in block.get("items", []):
            if item.get("type") == "song":
                song = song_store.get(item.get("song_id"))
                if song:
                    ordered_song_pairs.append((item["song_id"], song))

    # ---------------------------------------------------------
    # CASE 1: songs only
    # ---------------------------------------------------------
    if output_mode == "songs" or not preserve_template_slides:
        delete_all_slides(prs)
        for _, song in ordered_song_pairs:
            add_song_block_to_prs(prs, song, first_layout, rest_layout)

        output = BytesIO()
        prs.save(output)
        output.seek(0)
        return output

    # ---------------------------------------------------------
    # CASE 2: full deck with template preserved
    # ---------------------------------------------------------
    hidden_ids = set(st.session_state.get("hidden_section_ids", []))
    sections = st.session_state.get("template_sections", [])

    # Delete hidden section content slides first
    slides_to_delete = []
    for sec in sections:
        if sec["id"] in hidden_ids:
            slides_to_delete.extend(sec.get("content_slide_indices", []))

    for idx in sorted(set(slides_to_delete), reverse=True):
        delete_slide_by_index(prs, idx)

    # Append all song blocks first, and record where each block starts/ends
    block_records = []
    for block in service_order_blocks:
        section_song_pairs = []
        for item in block.get("items", []):
            if item.get("type") != "song":
                continue
            song = song_store.get(item.get("song_id"))
            if song:
                section_song_pairs.append((item["song_id"], song))

        if not section_song_pairs:
            continue

        start_idx, end_idx = add_section_song_block_to_prs(
            prs,
            section_song_pairs,
            first_layout,
            rest_layout,
        )

        block_records.append(
            {
                "section_id": block.get("section_id"),
                "section_title": block.get("section_title"),
                "start_idx": start_idx,
                "end_idx": end_idx,
                "slide_count": end_idx - start_idx + 1,
            }
        )

    section_anchor_map = build_section_anchor_map(prs)
    
    # Save intermediate PPTX so PPTXCreator can work on it
    with tempfile.TemporaryDirectory() as tmpdir:
        intermediate_path = os.path.join(tmpdir, "combined_before_reorder.pptx")
        final_path = os.path.join(tmpdir, "combined_after_reorder.pptx")

        prs.save(intermediate_path)

        template_obj = ExistingDeckTemplate(
            intermediate_path,
            title_layout_name=FIRST_LAYOUT_NAME,
            default_layout_name=REST_LAYOUT_NAME,
            divider_layout_name=DIVIDER_LAYOUT_NAME,
        )
        
        ppt = PPTXCreator(template_obj)

        # Track each block by slide_id, not start/end/count
        block_slide_ids = {
            i: [prs.slides[j].slide_id for j in range(rec["start_idx"], rec["end_idx"] + 1)]
            for i, rec in enumerate(block_records)
        }

        def move_block_with_interface(block_idx: int, target_idx: int):
            slide_ids = block_slide_ids[block_idx]
            if not slide_ids:
                return

            live_indices = [get_live_index_by_slide_id(ppt.prs.slides, sid) for sid in slide_ids]
            block_start = min(live_indices)
            block_end = max(live_indices)
            count = len(slide_ids)

            if block_start <= target_idx <= block_end + 1:
                return

            adjusted_target = target_idx if target_idx < block_start else target_idx - count

            for offset, sid in enumerate(slide_ids):
                desired_idx = adjusted_target + offset
                current_idx = get_live_index_by_slide_id(ppt.prs.slides, sid)
                slide_obj = get_slide_obj_by_slide_id(ppt.prs.slides, sid)

                if current_idx != desired_idx:
                    ppt.move_slide(slide_obj, desired_idx)

        # move from last block to first block
        for i in range(len(block_records) - 1, -1, -1):
            block = block_records[i]
            section_id = block["section_id"]
            section_title = block["section_title"]

            if section_id is None:
                continue

            target_idx = find_section_insert_index_from_anchor_map(
                ppt.prs.slides,
                section_title,
                section_anchor_map,
            )
            move_block_with_interface(i, target_idx)

        ppt.save(final_path)

        with open(final_path, "rb") as f:
            output = BytesIO(f.read())
            output.seek(0)
            return output


def create_single_song_ppt(song_item, template_bytes: bytes):
    prs = open_presentation_from_bytes(template_bytes)

    first_layout = get_layout_by_name(prs, FIRST_LAYOUT_NAME)
    rest_layout = get_layout_by_name(prs, REST_LAYOUT_NAME)

    if first_layout is None or rest_layout is None:
        raise ValueError("Template layouts not found.")

    delete_all_slides(prs)
    add_song_block_to_prs(prs, song_item, first_layout, rest_layout)

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output


# =========================================================
# PREVIEW HELPERS
# =========================================================
def pptx_to_preview_images(pptx_bytes: BytesIO):
    with tempfile.TemporaryDirectory() as tmpdir:
        pptx_path = os.path.join(tmpdir, "preview.pptx")

        with open(pptx_path, "wb") as f:
            f.write(pptx_bytes.getvalue())

        cmd = [
            SOFFICE_PATH,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", tmpdir,
            pptx_path,
        ]

        result = subprocess.run(cmd, capture_output=True, text=True)

        stderr_text = (result.stderr or "").strip().lower()

        if (
            result.returncode != 0
            or "source file could not be loaded" in stderr_text
            or "error:" in stderr_text
        ):
            raise RuntimeError(
                "Preview conversion failed.\n"
                f"Return code: {result.returncode}\n"
                f"stdout: {result.stdout}\n"
                f"stderr: {result.stderr}"
            )

        pdf_files = [f for f in os.listdir(tmpdir) if f.lower().endswith(".pdf")]
        if not pdf_files:
            raise FileNotFoundError(
                "Preview PDF was not created.\n"
                f"Files in temp dir: {os.listdir(tmpdir)}\n"
                f"stdout: {result.stdout}\n"
                f"stderr: {result.stderr}"
            )

        pdf_path = os.path.join(tmpdir, pdf_files[0])

        try:
            doc = fitz.open(pdf_path)
            repaired_bytes = doc.tobytes(garbage=3, clean=True, deflate=True)
            doc.close()
            doc = fitz.open(stream=repaired_bytes, filetype="pdf")
        except Exception as e:
            raise RuntimeError(f"Unable to open preview PDF in PyMuPDF: {e}")

        images = []
        try:
            for page in doc:
                pix = page.get_pixmap(dpi=60, alpha=False)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                buffer = io.BytesIO()
                img.save(buffer, format="JPEG", quality=45, optimize=True)
                images.append(buffer.getvalue())

                buffer.close()
                del pix
                del img
                del buffer
        finally:
            doc.close()

        return images


def preview_error_message(exc: Exception) -> str:
    msg = str(exc).strip()

    if "source file could not be loaded" in msg:
        return (
            "LibreOffice could not open the generated PowerPoint preview file. "
            "This usually means the combined PPTX structure is invalid for LibreOffice."
        )

    if "Preview conversion failed" in msg:
        return "Preview conversion failed."

    if "Preview PDF was not created" in msg:
        return "Preview PDF was not created."

    if "No common ancestor in structure tree" in msg:
        return (
            "Preview generation failed because the generated PDF structure is malformed."
        )

    return f"Preview generation failed: {msg}"


def render_scrollable_images(images, height=760, active_slide=None):
    if not images:
        st.info("Preview will appear here.")
        return

    container_id = f"preview-scroll-container-{len(images)}-{active_slide}"
    active_slide_js = "null" if active_slide is None else str(active_slide)

    html = f"""
    <div id="{container_id}" style="
        height: {height}px;
        overflow-y: auto;
        border: 1px solid #ddd;
        padding: 12px;
        border-radius: 8px;
        background: #fafafa;
        box-sizing: border-box;
        scroll-behavior: smooth;
    ">
    """

    for i, img_bytes in enumerate(images, start=1):
        b64 = base64.b64encode(img_bytes).decode("utf-8")
        border = "3px solid #2563eb" if active_slide == i else "1px solid #ccc"
        badge = " ← current" if active_slide == i else ""

        html += f"""
        <div id="slide-{i}" style="margin-bottom: 24px;">
            <div style="font-weight: 600; margin-bottom: 8px;">Slide {i}{badge}</div>
            <img
                src="data:image/jpeg;base64,{b64}"
                style="width: 100%; border: {border}; display: block;"
            />
        </div>
        """

    html += "</div>"

    html += f"""
    <script>
    const container = document.getElementById("{container_id}");
    const activeSlide = {active_slide_js};

    function scrollToActiveSlide() {{
        if (!container || activeSlide === null) return;
        const target = document.getElementById("slide-" + activeSlide);
        if (!target) return;
        container.scrollTop = target.offsetTop - 12;
    }}

    setTimeout(() => {{
        scrollToActiveSlide();
    }}, 200);
    </script>
    """

    st.components.v1.html(html, height=height, scrolling=False)


# =========================================================
# AUTO PREVIEW HELPERS
# =========================================================
def refresh_current_song_preview(song_item, template_bytes):
    st.session_state["current_song_preview_images"] = None
    st.session_state["current_song_preview_stats"] = None

    ppt_data = create_single_song_ppt(song_item, template_bytes)
    preview_images = pptx_to_preview_images(ppt_data)

    st.session_state["current_song_preview_images"] = preview_images
    st.session_state["current_song_preview_stats"] = preview_stats(preview_images)
    st.session_state["last_current_song_signature"] = build_current_song_signature(
        song_item,
        st.session_state.get("selected_template_name"),
    )


def blank_separator_added(old_text: str, new_text: str) -> bool:
    def valid_blank_positions(lines):
        positions = []
        for i, line in enumerate(lines):
            if line.strip() == "":
                for j in range(i + 1, len(lines)):
                    if lines[j].strip() != "":
                        positions.append(i)
                        break
        return positions

    old_lines = old_text.splitlines()
    new_lines = new_text.splitlines()
    return len(valid_blank_positions(new_lines)) > len(valid_blank_positions(old_lines))


def detect_new_slide_target_line(old_text: str, new_text: str):
    old_lines = old_text.splitlines()
    new_lines = new_text.splitlines()

    min_len = min(len(old_lines), len(new_lines))
    changed_idx = None

    for i in range(min_len):
        if old_lines[i] != new_lines[i]:
            changed_idx = i
            break

    if changed_idx is None:
        if len(new_lines) > len(old_lines):
            changed_idx = len(old_lines)
        else:
            return None

    for j in range(changed_idx, len(new_lines)):
        if new_lines[j].strip() != "":
            return j

    return None


def get_first_new_blank_separator_index(old_text: str, new_text: str):
    old_blank_positions = [i for i, line in enumerate(old_text.splitlines()) if line.strip() == ""]
    new_blank_positions = [i for i, line in enumerate(new_text.splitlines()) if line.strip() == ""]

    for pos in new_blank_positions:
        if pos not in old_blank_positions:
            return pos
    return None


def get_slide_number_from_line_index(text: str, line_index: int, auto_split: bool, lines_per_slide: int):
    if line_index is None:
        return None

    lines = text.splitlines()

    if auto_split:
        current_verse_indexes = []
        line_to_slide = {}
        slide_num = 1

        for idx, raw_line in enumerate(lines):
            stripped = raw_line.strip()

            if stripped == "":
                if current_verse_indexes:
                    for j in range(0, len(current_verse_indexes), lines_per_slide):
                        chunk = current_verse_indexes[j:j + lines_per_slide]
                        for original_idx in chunk:
                            line_to_slide[original_idx] = slide_num
                        slide_num += 1
                    current_verse_indexes = []
            else:
                current_verse_indexes.append(idx)

        if current_verse_indexes:
            for j in range(0, len(current_verse_indexes), lines_per_slide):
                chunk = current_verse_indexes[j:j + lines_per_slide]
                for original_idx in chunk:
                    line_to_slide[original_idx] = slide_num
                slide_num += 1

        return line_to_slide.get(line_index)

    slide_num = 1
    in_slide = False
    in_metadata_part = False

    for idx, raw_line in enumerate(lines):
        stripped = raw_line.strip()

        # "[Verse 1]" / "#Verse 1#" lines start a new slide and aren't lyrics themselves
        marker_label = parse_part_marker(raw_line)
        if marker_label is not None:
            if in_slide:
                slide_num += 1
                in_slide = False
            in_metadata_part = marker_label.lower() in METADATA_PART_LABELS
            continue
        if in_metadata_part:
            continue

        if stripped == "":
            if in_slide:
                slide_num += 1
                in_slide = False
        else:
            in_slide = True
            if idx == line_index:
                return slide_num

    return None


# =========================================================
# SERVICE PREVIEW POSITION HELPERS
# =========================================================
def get_service_song_start_slides(setlist, template_bytes: bytes):
    starts = [None] * len(setlist)

    if st.session_state.get("service_output_mode") == "songs" or not st.session_state.get("preserve_template_slides", True):
        slide_counter = 1
        ordered = get_ordered_songs_for_output(setlist)
        for original_idx, song in ordered:
            starts[original_idx] = slide_counter
            slide_counter += len(song["slides"])
        return starts

    sections = st.session_state.get("template_sections", [])
    hidden_ids = set(st.session_state.get("hidden_section_ids", []))
    groups = group_songs_by_section_order(setlist)

    songs_by_sec = {sec_id: song_pairs for sec_id, _, song_pairs in groups if sec_id is not None}
    unassigned = next((song_pairs for sec_id, _, song_pairs in groups if sec_id is None), [])

    slide_counter = 1

    for sec in sections:
        slide_counter += 1
        if sec["id"] not in hidden_ids:
            slide_counter += len(sec.get("content_slide_indices", []))

        for original_idx, song in songs_by_sec.get(sec["id"], []):
            starts[original_idx] = slide_counter
            slide_counter += len(song["slides"])

    for original_idx, song in unassigned:
        starts[original_idx] = slide_counter
        slide_counter += len(song["slides"])

    return starts


def refresh_service_preview(setlist, template_bytes):
    sync_block_model_from_setlist()
    st.session_state["service_preview_images"] = None
    st.session_state["service_preview_stats"] = None
    st.session_state["service_preview_error"] = None
    st.session_state["ppt_data"] = None

    ppt_data = create_combined_ppt(setlist, template_bytes)
    st.session_state["ppt_data"] = ppt_data
    st.session_state["service_song_start_slides"] = get_service_song_start_slides(
        setlist, template_bytes
    )

    try:
        preview_images = pptx_to_preview_images(ppt_data)

        if not preview_images:
            raise RuntimeError("No preview images generated from service PPT")

        st.session_state["service_preview_images"] = preview_images
        st.session_state["service_preview_stats"] = preview_stats(preview_images)

    except Exception as e:
        st.session_state["service_preview_images"] = None
        st.session_state["service_preview_stats"] = None
        st.session_state["service_preview_error"] = preview_error_message(e)


def selected_template_info():
    template_name = st.session_state.get("selected_template_name")
    if not template_name:
        return None, False, [], []

    uploaded = st.session_state["uploaded_templates"]
    if template_name not in uploaded:
        return None, False, [], []

    template_bytes = uploaded[template_name]
    ok, errors, warnings = validate_template_bytes(template_bytes)
    return template_bytes, ok, errors, warnings


# =========================================================
# PRE-RUN ACTIONS
# =========================================================
apply_pending_setlist_load()

if st.session_state.get("reset_editor_pending"):
    reset_editor()
    st.session_state["reset_editor_pending"] = False

selected_template_bytes, selected_template_ok, selected_template_errors, selected_template_warnings = selected_template_info()

if selected_template_bytes and selected_template_ok:
    try:
        current_sections = parse_template_sections(selected_template_bytes)
        old_sections = st.session_state.get("template_sections", [])

        old_hidden_by_title = {
            sec["title"]: (sec["id"] in st.session_state.get("hidden_section_ids", []))
            for sec in old_sections
        }

        st.session_state["template_sections"] = current_sections

        new_hidden_ids = []
        for sec in current_sections:
            if old_hidden_by_title.get(sec["title"], False):
                new_hidden_ids.append(sec["id"])
        st.session_state["hidden_section_ids"] = new_hidden_ids

    except Exception:
        st.session_state["template_sections"] = []
        st.session_state["hidden_section_ids"] = []
else:
    st.session_state["template_sections"] = []
    st.session_state["hidden_section_ids"] = []

if st.session_state.get("template_sections"):
    valid_section_ids = {sec["id"] for sec in st.session_state["template_sections"]}
    if st.session_state.get("selected_song_section_id") not in valid_section_ids:
        st.session_state["selected_song_section_id"] = st.session_state["template_sections"][0]["id"]
else:
    st.session_state["selected_song_section_id"] = None

if st.session_state.get("template_sections") and not st.session_state.get("service_order_blocks"):
    st.session_state["service_order_blocks"] = create_empty_service_order_blocks_from_template(
        st.session_state["template_sections"]
    )


# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    sync_block_model_from_setlist()
    st.markdown("### Service Order")

    setlist = st.session_state["setlist"]

    if not setlist:
        st.caption("No songs added yet.")
    else:
        selected_index = st.session_state.get("setlist_selected_index", 0)
        editing_index = st.session_state.get("editing_setlist_index")

        selected_index = max(0, min(selected_index, len(setlist) - 1))
        st.session_state["setlist_selected_index"] = selected_index

        for group in build_template_service_order_view(setlist):
            render_service_group_items_markdown(
                group,
                selected_index=selected_index,
                editing_index=editing_index,
            )

    st.divider()

    st.markdown("### Loaded in Editor")

    loaded_umh = st.session_state.get("editor_umh", "").strip()
    loaded_title = st.session_state.get("editor_title", "").strip()
    loaded_idx = st.session_state.get("editing_setlist_index")
    loaded_section_title = get_section_title_by_id(st.session_state.get("selected_song_section_id"))

    if loaded_title:
        section_line = f"\n\nSection: {loaded_section_title}" if loaded_section_title else ""

        if loaded_idx is not None:
            if loaded_umh:
                st.info(f"Editing setlist item #{loaded_idx + 1}\n\nUMH {loaded_umh} {loaded_title}{section_line}")
            else:
                st.info(f"Editing setlist item #{loaded_idx + 1}\n\n{loaded_title}{section_line}")
        else:
            if loaded_umh:
                st.info(f"New / repository song\n\nUMH {loaded_umh} {loaded_title}{section_line}")
            else:
                st.info(f"New / repository song\n\n{loaded_title}{section_line}")
    else:
        st.caption("No song loaded in editor.")

    st.divider()

    st.header("Controls")

    with st.expander("1. Template", expanded=False):
        uploaded_templates = st.file_uploader(
            "Upload template(s)",
            type=["pptx"],
            accept_multiple_files=True,
            key="template_uploader",
        )

        if uploaded_templates:
            for file in uploaded_templates:
                st.session_state["uploaded_templates"][file.name] = file.getvalue()

        template_names = list(st.session_state["uploaded_templates"].keys())

        if template_names:
            default_index = 0
            if st.session_state["selected_template_name"] in template_names:
                default_index = template_names.index(st.session_state["selected_template_name"])

            chosen_template = st.selectbox(
                "Select template",
                template_names,
                index=default_index,
            )

            if chosen_template != st.session_state.get("selected_template_name"):
                st.session_state["selected_template_name"] = chosen_template
                clear_service_outputs()
                st.session_state["current_song_preview_images"] = None
                st.session_state["current_song_preview_stats"] = None
                st.session_state["last_current_song_signature"] = None
                st.rerun()
            else:
                st.session_state["selected_template_name"] = chosen_template

            selected_template_bytes, selected_template_ok, selected_template_errors, selected_template_warnings = selected_template_info()

            if selected_template_ok:
                st.success("Template valid")
            else:
                st.error("Template invalid")
                for err in selected_template_errors:
                    st.write(f"- {err}")

            if selected_template_warnings:
                for warn in selected_template_warnings:
                    st.warning(warn)

            if st.button("Remove selected template", use_container_width=True):
                del st.session_state["uploaded_templates"][st.session_state["selected_template_name"]]
                st.session_state["selected_template_name"] = None
                st.session_state["template_sections"] = []
                st.session_state["hidden_section_ids"] = []
                st.session_state["selected_song_section_id"] = None
                clear_service_outputs()
                st.rerun()

            st.divider()
            st.checkbox(
                "Keep existing slides in template",
                key="preserve_template_slides",
            )

            st.radio(
                "Service output mode",
                options=["full", "songs"],
                format_func=lambda x: (
                    "Full deck (insert section-by-section)"
                    if x == "full"
                    else "Songs only (for checking)"
                ),
                key="service_output_mode",
            )

            if selected_template_ok and st.session_state["template_sections"]:
                st.divider()
                st.markdown("#### Hide Template Section Contents")

                section_options = [sec["id"] for sec in st.session_state["template_sections"]]
                section_name_map = {
                    sec["id"]: sec["title"] for sec in st.session_state["template_sections"]
                }

                st.multiselect(
                    "Hide contents under these section headers",
                    options=section_options,
                    default=st.session_state.get("hidden_section_ids", []),
                    format_func=lambda sid: section_name_map[sid],
                    key="hidden_section_ids",
                )

                for sec in st.session_state["template_sections"]:
                    hidden = sec["id"] in st.session_state["hidden_section_ids"]
                    status = "content hidden" if hidden else "content shown"
                    st.caption(
                        f'{sec["title"]}: {len(sec["content_slide_indices"])} content slide(s) · divider kept · {status}'
                    )
            elif selected_template_ok:
                st.info(
                    f'No "{DIVIDER_LAYOUT_NAME}" slides detected. '
                    "Section tools will not be available."
                )
        else:
            st.info("Upload at least one template.")

        if not soffice_available():
            st.warning("LibreOffice/soffice is not available.")

    with st.expander("2. Load Song", expanded=False):
        if st.button("Start New Song", use_container_width=True):
            st.session_state["reset_editor_pending"] = True
            st.rerun()

        load_mode = st.radio("Find hymn by", ["UMH Number", "Title"], horizontal=True)

        if load_mode == "UMH Number":
            umh_number_input = st.text_input("UMH Number", placeholder="e.g. 57")
            if st.button("Load by Number", use_container_width=True):
                if umh_number_input.strip():
                    match = find_row_by_umh(umh_number_input)
                    if match:
                        load_song_into_editor(match)
                        st.success("Hymn loaded.")
                        st.rerun()
                    else:
                        st.error("Hymn not found.")
        else:
            keyword = st.text_input("Search title", placeholder="e.g. thousand tongues")
            matches = search_titles(keyword) if keyword.strip() else []

            if matches:
                options = [
                    f'UMH {row.get("UMH Number","")} - {row.get("Title","")}'
                    for row in matches
                ]
                selected = st.selectbox("Select hymn", options)

                if st.button("Load by Title", use_container_width=True):
                    chosen_index = options.index(selected)
                    load_song_into_editor(matches[chosen_index])
                    st.success("Hymn loaded.")
                    st.rerun()
            elif keyword.strip():
                st.info("No matching titles found.")

    with st.expander("3. Service Order", expanded=True):
        setlist = st.session_state["setlist"]

        if not setlist:
            st.info("No songs added yet.")
        else:
            st.markdown("#### Service Order (by DOCX block, using template section headers)")
            for group in build_template_service_order_view(setlist):
                if not group["items"]:
                    continue
                render_service_group_items_caption(group)

            st.divider()
            st.markdown("#### Song Selection")
            labels = []
            for i, song in enumerate(setlist):
                section_title = get_section_title_by_id(song.get("section_id"))
                section_suffix = f" [{section_title}]" if section_title else ""

                if song["umh_number"]:
                    labels.append(f'{i+1}. UMH {song["umh_number"]} {song["title"]}{section_suffix} ({len(song["slides"])})')
                else:
                    labels.append(f'{i+1}. {song["title"]}{section_suffix} ({len(song["slides"])})')

            # stable song selection by song_id
            for i, song in enumerate(setlist):
                if not song.get("song_id"):
                    song["song_id"] = f"song_{i}_{song.get('umh_number','')}_{song.get('title','')}"
            
            ids = [song.get("song_id") for song in setlist if song.get("song_id")]
            label_map = {song["song_id"]: labels[i] for i, song in enumerate(setlist)}
            
            selected_song_id = st.session_state.get("selected_song_id")
            if selected_song_id not in ids:
                selected_song_id = ids[0]
                st.session_state["selected_song_id"] = selected_song_id
            
            previous_selected_index = next(
                (i for i, song in enumerate(setlist) if song.get("song_id") == selected_song_id),
                0,
            )
            
            widget_key = f"song_select_{st.session_state.get('song_selectbox_version', 0)}"
            
            selected_song_id = st.selectbox(
                "Selected song",
                options=ids,
                index=previous_selected_index,
                format_func=lambda sid: label_map[sid],
                key=widget_key,
            )
            
            st.session_state["selected_song_id"] = selected_song_id
            
            selected_index = next(
                (i for i, song in enumerate(setlist) if song.get("song_id") == selected_song_id),
                0,
            )
            st.session_state["setlist_selected_index"] = selected_index

            if (
                selected_index != previous_selected_index
                and st.session_state.get("preview_mode") == "service"
            ):
                starts = st.session_state.get("service_song_start_slides", [])
                st.session_state["current_preview_slide"] = (
                    starts[selected_index] if selected_index < len(starts) and starts[selected_index] is not None else 1
                )
                st.rerun()

            action_cols = st.columns(4)

            with action_cols[0]:
                if st.button("✏️", use_container_width=True, help="Edit selected song"):
                    st.session_state["pending_setlist_load"] = selected_index
                    st.session_state["preview_mode"] = "song"
                    st.session_state["current_song_preview_images"] = None
                    st.session_state["current_song_preview_stats"] = None
                    st.session_state["last_current_song_signature"] = None
                    st.rerun()

            with action_cols[1]:
                if st.button("⬆️", use_container_width=True, help="Move selected song up") and selected_index > 0:
                    moved_song_id = setlist[selected_index].get("song_id")
            
                    setlist[selected_index - 1], setlist[selected_index] = (
                        setlist[selected_index],
                        setlist[selected_index - 1],
                    )
            
                    st.session_state["selected_song_id"] = moved_song_id
                    st.session_state["song_selectbox_version"] += 1
            
                    editing_index = st.session_state.get("editing_setlist_index")
                    if editing_index == selected_index:
                        st.session_state["editing_setlist_index"] = selected_index - 1
                    elif editing_index == selected_index - 1:
                        st.session_state["editing_setlist_index"] = selected_index
            
                        pending = st.session_state.get("pending_setlist_load")
                        if pending == selected_index:
                            st.session_state["pending_setlist_load"] = selected_index - 1
                        elif pending == selected_index - 1:
                            st.session_state["pending_setlist_load"] = selected_index
                
                        clear_service_outputs()
                        st.rerun()
    
            with action_cols[2]:
                if st.button("⬇️", use_container_width=True, help="Move selected song down") and selected_index < len(setlist) - 1:
                    moved_song_id = setlist[selected_index].get("song_id")
            
                    setlist[selected_index + 1], setlist[selected_index] = (
                        setlist[selected_index],
                        setlist[selected_index + 1],
                    )
            
                    st.session_state["selected_song_id"] = moved_song_id
                    st.session_state["song_selectbox_version"] += 1
            
                    editing_index = st.session_state.get("editing_setlist_index")
                    if editing_index == selected_index:
                        st.session_state["editing_setlist_index"] = selected_index + 1
                    elif editing_index == selected_index + 1:
                        st.session_state["editing_setlist_index"] = selected_index
            
                    pending = st.session_state.get("pending_setlist_load")
                    if pending == selected_index:
                        st.session_state["pending_setlist_load"] = selected_index + 1
                    elif pending == selected_index + 1:
                        st.session_state["pending_setlist_load"] = selected_index
            
                    clear_service_outputs()
                    st.rerun()
    
            with action_cols[3]:
                if st.button("🗑️", use_container_width=True, help="Delete selected song"):
                    if len(setlist) > 1:
                        if selected_index < len(setlist) - 1:
                            next_id = setlist[selected_index + 1].get("song_id")
                        else:
                            next_id = setlist[selected_index - 1].get("song_id")
                    else:
                        next_id = None
            
                    setlist.pop(selected_index)
            
                    st.session_state["selected_song_id"] = next_id
                    st.session_state["song_selectbox_version"] += 1
            
                    if st.session_state.get("editing_setlist_index") == selected_index:
                        st.session_state["reset_editor_pending"] = True
                    elif (
                        st.session_state.get("editing_setlist_index") is not None
                        and st.session_state["editing_setlist_index"] > selected_index
                    ):
                        st.session_state["editing_setlist_index"] -= 1
            
                    pending = st.session_state.get("pending_setlist_load")
                    if pending == selected_index:
                        st.session_state["pending_setlist_load"] = None
                    elif pending is not None and pending > selected_index:
                        st.session_state["pending_setlist_load"] = pending - 1
            
                    clear_service_outputs()
                    st.rerun()
    
            if st.button("Clear Setlist", use_container_width=True, type="secondary"):
                st.session_state["setlist"] = []
                st.session_state["service_order_blocks"] = []
                st.session_state["song_store"] = {}
                st.session_state["last_docx_section_mapping"] = []
                st.session_state["editing_setlist_index"] = None
                st.session_state["pending_setlist_load"] = None
                st.session_state["setlist_selected_index"] = 0
                st.session_state["selected_song_id"] = None
                st.session_state["preview_mode"] = "song"
                st.session_state["current_song_preview_images"] = None
                st.session_state["current_song_preview_stats"] = None
                clear_service_outputs()
                st.rerun()



    with st.expander("4. Import Order of Service", expanded=False):
        service_docx_file = st.file_uploader(
            "Upload Order of Service (.docx)",
            type=["docx"],
            key="service_docx_importer",
        )

        replace_existing_setlist = st.checkbox(
            "Replace current setlist",
            value=False,
            key="replace_setlist_from_docx",
        )
        st.caption("Template sections remain the source of truth. DOCX is parsed into ordered blocks, and only major matched headings start a new template section block.")

        if st.button("Import Songs from DOCX", use_container_width=True):
            if service_docx_file is None:
                st.warning("Please upload a .docx file first.")
            else:
                try:
                    service_order_blocks, song_store, missing_songs, section_mapping_rows = import_service_order_from_docx(
                        service_docx_file,
                        st.session_state.get("template_sections", []),
                    )
        
                    st.session_state["service_order_blocks"] = service_order_blocks
                    st.session_state["song_store"] = song_store
                    st.session_state["setlist"] = flatten_blocks_to_setlist(service_order_blocks, song_store)
        
                    ids_after_import = [s.get("song_id") for s in st.session_state["setlist"] if s.get("song_id")]
                    if ids_after_import:
                        if st.session_state.get("selected_song_id") not in ids_after_import:
                            st.session_state["selected_song_id"] = ids_after_import[0]
                    else:
                        st.session_state["selected_song_id"] = None
        
                    st.session_state["song_selectbox_version"] += 1
                    st.session_state["last_docx_section_mapping"] = section_mapping_rows
        
                    clear_service_outputs()
        
                    st.session_state["editing_setlist_index"] = None
                    st.session_state["pending_setlist_load"] = None
                    st.session_state["setlist_selected_index"] = 0
        
                    st.session_state["last_docx_import_report"] = {
                        "song_count": len(st.session_state["setlist"]),
                        "mapping": section_mapping_rows,
                        "missing": missing_songs,
                    }
        
                    st.rerun()
        
                except Exception as e:
                    st.exception(e)

        import_report = st.session_state.get("last_docx_import_report")
        if import_report:
            if import_report["song_count"]:
                st.success(f"Imported {import_report['song_count']} song(s) into the template-driven service order.")
            else:
                st.warning(
                    "No songs were imported. Start each song with a \"[Title]\" line followed by "
                    "the title and lyrics, or a line starting with \"UMH\", e.g. "
                    "\"UMH 140 Great Is Thy Faithfulness\" (lyrics from Google Sheets) or "
                    "\"UMH Great Is Thy Faithfulness\" followed by the lyrics in the DOCX."
                )

            if import_report["mapping"]:
                st.caption("Section mapping:")
                for row in import_report["mapping"]:
                    mapped = row.get("mapped_section_title") or "No template section matched"
                    st.write(
                        f"{row['source']} {row['docx_heading']} → {mapped}"
                        f" ({row['match_type']}, score {row['score']})"
                    )

            if import_report["missing"]:
                st.warning(
                    f"{len(import_report['missing'])} song(s) were found in the DOCX but have no "
                    "lyrics in the DOCX and were not found in Google Sheets."
                )
                for song in import_report["missing"]:
                    st.write(f"- {song['raw']}")


# =========================================================
# MAIN LAYOUT
# =========================================================
main_left, main_right = st.columns([1.15, 1], vertical_alignment="top")

with main_left:
    st.subheader("Song Editor")

    edit_idx = st.session_state.get("editing_setlist_index")
    if edit_idx is not None:
        st.info(f"Editing setlist item #{edit_idx + 1}")

    meta_col1, meta_col2 = st.columns([1, 3])
    with meta_col1:
        st.text_input("UMH", key="editor_umh")
    with meta_col2:
        st.text_input("Title", key="editor_title")

    if st.session_state.get("template_sections"):
        section_options = [sec["id"] for sec in st.session_state["template_sections"]]
        section_name_map = {sec["id"]: sec["title"] for sec in st.session_state["template_sections"]}

        current_section = st.session_state.get("selected_song_section_id")
        if current_section not in section_options:
            current_section = section_options[0]
            st.session_state["selected_song_section_id"] = current_section

        st.selectbox(
            "Place new song under section",
            options=section_options,
            format_func=lambda sid: section_name_map[sid],
            key="selected_song_section_id",
        )
    else:
        st.caption("No template sections detected.")

    st.markdown("#### Slide Splitting")
    split_col1, split_col2 = st.columns([3, 1])
    
    with split_col1:
        st.checkbox("Enable smart split for long lyric lines", key="smart_split_enabled")
    
        if st.session_state["smart_split_enabled"]:
            st.slider(
                "Smart split max characters per line",
                min_value=20,
                max_value=50,
                key="smart_split_max_chars",
            )
            st.slider(
                "Smart split max lines per slide",
                min_value=2,
                max_value=6,
                key="smart_split_max_lines_per_slide",
            )
    
    with split_col2:
        st.write("")
        refresh_song_preview_clicked = st.button("Refresh Song Preview", use_container_width=True)
        old_text = st.session_state.get("last_editor_text", "")
    
    editor_text = st_ace(
        value=st.session_state.get("editor_text", ""),
        language="text",
        theme="textmate",
        keybinding="vscode",
        font_size=16,
        tab_size=2,
        wrap=True,
        show_gutter=False,
        auto_update=True,
        readonly=False,
        height=420,
        key=f"editor_ace_{st.session_state['editor_ace_key']}",
    )

    if editor_text is None:
        editor_text = st.session_state.get("editor_text", "")
    else:
        st.session_state["editor_text"] = editor_text

    if editor_text != old_text:
        clear_service_outputs()

    current_slides, current_slide_parts = get_current_slides_and_parts(editor_text)
    song_item = build_editor_song_item(current_slides, current_slide_parts)

    if st.session_state.get("smart_split_enabled", False):
        st.caption(
            f"{len(current_slides)} slide(s) "
            f"(smart split: max {st.session_state.get('smart_split_max_chars', 30)} chars, "
            f"{st.session_state.get('smart_split_max_lines_per_slide', 4)} lines per slide)"
        )
    else:
        st.caption(f"{len(current_slides)} slide(s) (manual mode: blank lines separate slides)")

    detected_part_labels, _ = build_part_nav(current_slide_parts)
    if detected_part_labels:
        st.caption("Song parts (footer links): " + PART_NAV_SEPARATOR.join(detected_part_labels))
    else:
        st.caption("No song parts detected. Add lines like [Verse 1] or [Chorus] to label slides.")

    if refresh_song_preview_clicked:
        if selected_template_bytes is None:
            st.error("Please upload and select a template first.")
        elif not selected_template_ok:
            st.error("Selected template is invalid.")
        elif not soffice_available():
            st.error("LibreOffice/soffice is not available.")
        elif not current_slides:
            st.error("No slides to preview.")
        else:
            try:
                st.session_state["preview_mode"] = "song"
                st.session_state["current_song_preview_images"] = None
                st.session_state["current_song_preview_stats"] = None
                refresh_current_song_preview(song_item, selected_template_bytes)
                st.session_state["editor_status_message"] = "Song preview refreshed."
                st.rerun()
            except Exception as e:
                st.error(preview_error_message(e))
                
        text_changed = editor_text != old_text
        trigger_refresh = False
        smart_split = st.session_state.get("smart_split_enabled", False)
        
        if smart_split:
            trigger_refresh = get_current_slides(old_text) != get_current_slides(editor_text)
        else:
            trigger_refresh = blank_separator_added(old_text, editor_text)
        
        if text_changed and trigger_refresh:
            if not smart_split:
                blank_idx = get_first_new_blank_separator_index(old_text, editor_text)
                if blank_idx is not None:
                    lines = editor_text.splitlines()
                    target_line_index = None
                    for i in range(blank_idx + 1, len(lines)):
                        if lines[i].strip() != "":
                            target_line_index = i
                            break
        
                    detected_slide = get_slide_number_from_line_index(
                        editor_text,
                        target_line_index,
                        auto_split=False,
                        lines_per_slide=4,
                    )
                    if detected_slide is not None:
                        st.session_state["current_preview_slide"] = detected_slide
            else:
                # Smart split can rebalance multiple lines/slides,
                # so exact slide detection from the old fixed-line logic is unreliable.
                st.session_state["current_preview_slide"] = 1
        
        st.session_state["last_editor_text"] = editor_text
        st.session_state["last_split_settings"] = current_split_settings
        
        if st.session_state["editor_status_message"]:
            st.caption(st.session_state["editor_status_message"])
    
    st.markdown("#### Song Formatting")

    fmt_col1, fmt_col2 = st.columns(2)

    with fmt_col1:
        st.checkbox("Override lyrics font size", key="editor_override_lyrics_font_size")
        if st.session_state["editor_override_lyrics_font_size"]:
            st.slider(
                "Lyrics font size (pt)",
                min_value=12,
                max_value=60,
                key="editor_lyrics_font_size_pt",
                on_change=lambda: st.session_state.update({"last_current_song_signature": None}),
            )
        else:
            st.caption("Using template lyrics size")

    with fmt_col2:
        st.checkbox("Override line spacing", key="editor_override_line_spacing")
        if st.session_state["editor_override_line_spacing"]:
            st.slider(
                "Line spacing",
                min_value=0.8,
                max_value=2.0,
                step=0.1,
                key="editor_line_spacing",
                on_change=lambda: st.session_state.update({"last_current_song_signature": None}),
            )
        else:
            st.caption("Using template line spacing")

    st.markdown("#### Add / Update")
    allow_duplicates = st.checkbox("Allow duplicate songs in setlist", value=False)

    action_col1, action_col2 = st.columns(2)
    with action_col1:
        add_or_update = st.button(
            "Update Song" if edit_idx is not None else "Add to Setlist",
            use_container_width=True,
        )
    with action_col2:
        clear_editor = st.button("Clear Editor", use_container_width=True)

    if add_or_update:
        if not current_slides:
            st.error("No slides to add.")
        else:
            item = build_editor_song_item(current_slides, current_slide_parts)
            edit_idx = st.session_state.get("editing_setlist_index")

            if edit_idx is None:
                duplicate_index = next(
                    (
                        i for i, s in enumerate(st.session_state["setlist"])
                        if s["umh_number"] == item["umh_number"]
                        and s["title"] == item["title"]
                        and s["slides"] == item["slides"]
                        and s.get("section_id") == item.get("section_id")
                    ),
                    None,
                )

                if duplicate_index is not None and not allow_duplicates:
                    st.warning(f"This song is already in the setlist as item #{duplicate_index + 1}.")
                else:
                    st.session_state["setlist"].append(item)
                    clear_service_outputs()
                    st.success(
                        f'Added: {"UMH " + item["umh_number"] + " " if item["umh_number"] else ""}{item["title"]}'
                    )
                    st.session_state["reset_editor_pending"] = True
                    st.rerun()
            else:
                old_item = st.session_state["setlist"][edit_idx]
                for meta_key in ["import_uid", "service_block_id", "service_block_order", "service_block_title"]:
                    if meta_key in old_item:
                        item[meta_key] = old_item[meta_key]
                if not item.get("section_id") and old_item.get("section_id"):
                    item["section_id"] = old_item.get("section_id")
                st.session_state["setlist"][edit_idx] = item
                clear_service_outputs()

                if (
                    selected_template_bytes is not None
                    and selected_template_ok
                    and soffice_available()
                    and current_slides
                ):
                    try:
                        st.session_state["current_song_preview_images"] = None
                        st.session_state["current_song_preview_stats"] = None
                        refresh_current_song_preview(item, selected_template_bytes)
                        st.session_state["current_preview_slide"] = 1
                        st.session_state["preview_mode"] = "song"
                        st.session_state["editor_status_message"] = "Song updated and preview refreshed."
                    except Exception as e:
                        st.session_state["editor_status_message"] = preview_error_message(e)

                st.success(
                    f'Updated: {"UMH " + item["umh_number"] + " " if item["umh_number"] else ""}{item["title"]}'
                )
                st.rerun()

    if clear_editor:
        st.session_state["reset_editor_pending"] = True
        st.rerun()


with main_right:
    st.subheader("Preview")

    resource_stats = get_runtime_resource_stats()
    with st.expander("Runtime resource usage", expanded=False):
        st.write(f'App RAM (RSS): {format_bytes(resource_stats["process_rss"])}')
        st.write(f'App virtual memory: {format_bytes(resource_stats["process_vms"])}')
        st.write(
            f'System RAM: {format_bytes(resource_stats["system_used_ram"])} used / '
            f'{format_bytes(resource_stats["system_total_ram"])} total'
        )
        st.write(
            f'System RAM available now: {format_bytes(resource_stats["system_available_ram"])}'
        )
        st.write(
            f'Disk: {format_bytes(resource_stats["disk_used"])} used / '
            f'{format_bytes(resource_stats["disk_total"])} total'
        )
        st.write(f'Disk free now: {format_bytes(resource_stats["disk_free"])}')

    preview_mode = st.radio(
        "Preview Mode",
        ["Song", "Service"],
        index=0 if st.session_state["preview_mode"] == "song" else 1,
        horizontal=True,
    )
    st.session_state["preview_mode"] = preview_mode.lower()

    if st.session_state["preview_mode"] == "song":
        st.session_state["service_preview_images"] = None
        st.session_state["service_preview_stats"] = None
    else:
        st.session_state["current_song_preview_images"] = None
        st.session_state["current_song_preview_stats"] = None

    preview_images = None

    if st.session_state["preview_mode"] == "song":
        st.caption("Current song preview")

        if selected_template_bytes is None:
            st.warning("Please upload and select a template to see preview.")
        elif not selected_template_ok:
            st.error("Selected template is invalid.")
        elif not soffice_available():
            st.warning("LibreOffice/soffice is required for preview.")

        preview_images = st.session_state.get("current_song_preview_images")

        song_stats = st.session_state.get("current_song_preview_stats")
        if song_stats:
            st.caption(
                f'Preview: {song_stats["count"]} slide(s) · '
                f'total {format_bytes(song_stats["total_bytes"])} · '
                f'avg {format_bytes(song_stats["avg_bytes"])} · '
                f'max {format_bytes(song_stats["max_bytes"])}'
            )

    else:
        if st.session_state.get("service_output_mode") == "songs":
            st.caption("Output mode: songs only (for checking)")
        else:
            st.caption("Output mode: full deck with section-by-section insertion")

        can_generate = (
            selected_template_bytes is not None
            and selected_template_ok
            and soffice_available()
            and len(st.session_state["setlist"]) > 0
        )

        if can_generate:
            refresh_service_now = st.button(
                "Refresh Service Preview",
                use_container_width=True
            )

            if refresh_service_now:
                try:
                    refresh_service_preview(
                        st.session_state["setlist"],
                        selected_template_bytes,
                    )
                except Exception as e:
                    st.error(f"PowerPoint generation failed: {e}")

            starts = st.session_state.get("service_song_start_slides", [])
            selected_index = st.session_state.get("setlist_selected_index", 0)

            if starts and 0 <= selected_index < len(starts) and starts[selected_index] is not None:
                st.session_state["current_preview_slide"] = starts[selected_index]
            else:
                st.session_state["current_preview_slide"] = 1

            if st.session_state.get("service_preview_error"):
                st.warning(
                    "Preview could not be generated, but the PowerPoint file is available for download.\n\n"
                    + st.session_state["service_preview_error"]
                )

            preview_images = st.session_state.get("service_preview_images")

            service_stats = st.session_state.get("service_preview_stats")
            if service_stats:
                st.caption(
                    f'Preview: {service_stats["count"]} slide(s) · '
                    f'total {format_bytes(service_stats["total_bytes"])} · '
                    f'avg {format_bytes(service_stats["avg_bytes"])} · '
                    f'max {format_bytes(service_stats["max_bytes"])}'
                )

            if st.session_state.get("ppt_data") is not None:
                st.download_button(
                    label="Download Service PowerPoint",
                    data=st.session_state["ppt_data"].getvalue(),
                    file_name="service_deck.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )

        else:
            if selected_template_bytes is None:
                st.info("Please upload and select a template first.")
            elif not selected_template_ok:
                st.info("Selected template is invalid.")
            elif not soffice_available():
                st.info("LibreOffice/soffice is not available.")
            elif not st.session_state["setlist"]:
                st.info("Add songs to the setlist to view the service preview.")

    if preview_images:
        render_scrollable_images(
            preview_images,
            height=700,
            active_slide=st.session_state.get("current_preview_slide"),
        )
    else:
        st.info("Preview will appear here.")
