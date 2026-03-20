import os
import io
import base64
import tempfile
import subprocess
from io import BytesIO
from shutil import which
from copy import deepcopy

import fitz  # PyMuPDF
import gspread
import streamlit as st
from streamlit_ace import st_ace
from google.oauth2.service_account import Credentials
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt
from PIL import Image


# =========================================================
# CONFIG
# =========================================================
SOFFICE_PATH = os.environ.get("SOFFICE_PATH", "soffice")

SHEET_KEY = st.secrets["SHEET_KEY"]
WORKSHEET_NAME = st.secrets["WORKSHEET_NAME"]

FIRST_LAYOUT_NAME = "TEMPLATE_FIRST"
REST_LAYOUT_NAME = "TEMPLATE_REST"
DIVIDER_LAYOUT_NAME = "SECTION_DIVIDER"


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
    "loaded_song": None,
    "reset_editor_pending": False,
    "uploaded_templates": {},
    "selected_template_name": None,
    "editor_umh": "",
    "editor_title": "",
    "editor_text": "",
    "editor_override_lyrics_font_size": False,
    "editor_override_line_spacing": False,
    "editor_lyrics_font_size_pt": 32,
    "editor_line_spacing": 1.2,
    "auto_split_by_lines": False,
    "lines_per_slide": 4,
    "refresh_on_new_line": True,
    "editor_ace_key": 0,
    "last_editor_text": "",
    "last_current_song_signature": None,
    "editor_status_message": "",
    "current_preview_slide": 1,
    "preview_mode": "song",
    "current_song_preview_images": None,
    "service_preview_images": None,
    "ppt_data": None,
    "last_split_settings": None,
    "preserve_template_slides": True,
    "service_sections": [],
    "divider_layout_names": [DIVIDER_LAYOUT_NAME],
    "selected_section_id": None,
    "editing_song_location": None,   # {"section_id": "...", "song_index": 0}
    "editor_target_section_id": None,
    "service_song_start_slides": [],
    "template_signature_for_sections": None,
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


def find_row_by_umh(umh_number: str):
    umh_number = str(umh_number).strip()
    for row in get_all_records_cached():
        if str(row.get("UMH Number", "")).strip() == umh_number:
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


def split_slides_manual(text: str) -> list[list[str]]:
    blocks = [block.strip() for block in text.split("\n\n") if block.strip()]
    slides = []

    for block in blocks:
        lines = [line.strip() for line in block.splitlines() if line.strip()]
        if lines:
            slides.append(lines)

    return slides


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


def get_current_slides(text: str) -> list[list[str]]:
    if st.session_state["auto_split_by_lines"]:
        return split_slides_by_line_count(
            text,
            lines_per_slide=st.session_state["lines_per_slide"],
        )
    return split_slides_manual(text)


def open_presentation_from_bytes(template_bytes: bytes):
    return Presentation(BytesIO(template_bytes))


def get_layout_by_name(prs, layout_name):
    for slide_master in prs.slide_masters:
        for slide_layout in slide_master.slide_layouts:
            if slide_layout.name.strip() == layout_name:
                return slide_layout
    return None


def get_body_placeholder(slide):
    placeholders = list(slide.placeholders)

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


def get_slide_layout_name(slide):
    try:
        return slide.slide_layout.name.strip()
    except Exception:
        return ""


def extract_slide_text(slide):
    parts = []
    for shape in slide.shapes:
        if hasattr(shape, "text"):
            text = shape.text.strip()
            if text:
                parts.append(text)
    return "\n".join(parts).strip()


# =========================================================
# TEMPLATE / SECTION HELPERS
# =========================================================
def is_divider_slide(slide):
    layout_name = get_slide_layout_name(slide).strip()
    divider_layout_names = st.session_state.get("divider_layout_names", [])
    return layout_name in divider_layout_names


def get_divider_title(slide, fallback="Untitled Section"):
    if slide.shapes.title and slide.shapes.title.text.strip():
        return slide.shapes.title.text.strip()

    text = extract_slide_text(slide)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return lines[0] if lines else fallback


def template_has_layout(prs, layout_name: str) -> bool:
    for slide_master in prs.slide_masters:
        for slide_layout in slide_master.slide_layouts:
            if slide_layout.name.strip() == layout_name:
                return True
    return False


def count_divider_slides(prs) -> int:
    divider_names = set(st.session_state.get("divider_layout_names", []))
    count = 0
    for slide in prs.slides:
        if get_slide_layout_name(slide).strip() in divider_names:
            count += 1
    return count


def find_divider_slides_missing_titles(prs):
    bad_indices = []
    for i, slide in enumerate(prs.slides, start=1):
        if is_divider_slide(slide):
            title = slide.shapes.title.text.strip() if slide.shapes.title and slide.shapes.title else ""
            if not title:
                bad_indices.append(i)
    return bad_indices


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

    if not template_has_layout(prs, DIVIDER_LAYOUT_NAME):
        errors.append(
            f"Missing divider layout: {DIVIDER_LAYOUT_NAME}. "
            f"Create it in Slide Master and use it for all service divider slides."
        )

    divider_slide_count = count_divider_slides(prs)
    if divider_slide_count == 0:
        errors.append(
            f"No divider slides found using layout {DIVIDER_LAYOUT_NAME}. "
            f"Add at least one divider slide based on that layout."
        )

    if errors:
        return False, errors, warnings

    # validate the song-generation layouts by creating temp slides
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

    missing_title_slides = find_divider_slides_missing_titles(prs)
    if missing_title_slides:
        warnings.append(
            "Divider slides missing title text at slide(s): "
            + ", ".join(map(str, missing_title_slides))
        )

    return len(errors) == 0, errors, warnings


def parse_template_sections(template_bytes: bytes):
    prs = open_presentation_from_bytes(template_bytes)

    sections = []
    current_section = None
    section_counter = 0

    divider_count = count_divider_slides(prs)
    if divider_count == 0:
        raise ValueError(
            f"Template has no {DIVIDER_LAYOUT_NAME} slides. "
            f"Please add divider slides using the {DIVIDER_LAYOUT_NAME} layout."
        )

    for i, slide in enumerate(prs.slides):
        if is_divider_slide(slide):
            current_section = {
                "id": f"sec_{section_counter}",
                "title": get_divider_title(slide, fallback=f"Section {section_counter + 1}"),
                "divider_index": i,
                "template_slide_indices": [],
                "include_divider": True,
                "include_template_content": True,
                "songs": [],
            }
            sections.append(current_section)
            section_counter += 1
        else:
            if current_section is None:
                continue
            current_section["template_slide_indices"].append(i)

    return sections


# =========================================================
# SONG / SECTION STATE HELPERS
# =========================================================
def build_editor_song_item(current_slides):
    return {
        "umh_number": st.session_state["editor_umh"].strip(),
        "title": st.session_state["editor_title"].strip(),
        "slides": current_slides,
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
        "target_section_id": st.session_state.get("editor_target_section_id"),
    }


def build_current_song_signature(song_item, selected_template_name):
    return (
        song_item["umh_number"],
        song_item["title"],
        tuple(tuple(slide) for slide in song_item["slides"]),
        song_item.get("lyrics_font_size_pt"),
        song_item.get("line_spacing"),
        selected_template_name,
        st.session_state["auto_split_by_lines"],
        st.session_state["lines_per_slide"],
    )


def get_section_by_id(section_id):
    for sec in st.session_state.get("service_sections", []):
        if sec["id"] == section_id:
            return sec
    return None


def get_section_index_by_id(section_id):
    sections = st.session_state.get("service_sections", [])
    for i, sec in enumerate(sections):
        if sec["id"] == section_id:
            return i
    return None


def clear_service_outputs():
    st.session_state["ppt_data"] = None
    st.session_state["service_preview_images"] = None
    st.session_state["service_song_start_slides"] = []


def reset_editor():
    st.session_state["loaded_song"] = None
    st.session_state["editor_umh"] = ""
    st.session_state["editor_title"] = ""
    st.session_state["editor_text"] = ""
    st.session_state["editor_override_lyrics_font_size"] = False
    st.session_state["editor_override_line_spacing"] = False
    st.session_state["editor_lyrics_font_size_pt"] = 32
    st.session_state["editor_line_spacing"] = 1.2
    st.session_state["current_song_preview_images"] = None
    st.session_state["last_editor_text"] = ""
    st.session_state["last_current_song_signature"] = None
    st.session_state["editor_status_message"] = ""
    st.session_state["current_preview_slide"] = 1
    st.session_state["editor_ace_key"] += 1
    st.session_state["editing_song_location"] = None
    st.session_state["editor_target_section_id"] = st.session_state.get("selected_section_id")


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


def sync_sections_from_template(template_bytes: bytes):
    template_sig = (
        st.session_state.get("selected_template_name"),
        len(template_bytes),
    )

    existing_sections = st.session_state.get("service_sections", [])
    existing_song_map = {
        sec["title"]: sec.get("songs", [])
        for sec in existing_sections
    }
    existing_include_content_map = {
        sec["title"]: sec.get("include_template_content", True)
        for sec in existing_sections
    }
    existing_include_divider_map = {
        sec["title"]: sec.get("include_divider", True)
        for sec in existing_sections
    }

    parsed_sections = parse_template_sections(template_bytes)

    for sec in parsed_sections:
        sec["songs"] = existing_song_map.get(sec["title"], [])
        sec["include_template_content"] = existing_include_content_map.get(sec["title"], True)
        sec["include_divider"] = existing_include_divider_map.get(sec["title"], True)

    st.session_state["service_sections"] = parsed_sections
    st.session_state["template_signature_for_sections"] = template_sig

    if parsed_sections:
        valid_ids = {sec["id"] for sec in parsed_sections}
        if st.session_state.get("selected_section_id") not in valid_ids:
            st.session_state["selected_section_id"] = parsed_sections[0]["id"]

        if st.session_state.get("editor_target_section_id") not in valid_ids:
            st.session_state["editor_target_section_id"] = st.session_state["selected_section_id"]


# =========================================================
# PPT BUILD HELPERS
# =========================================================
def add_song_to_presentation(prs, song, first_layout, rest_layout):
    umh_number = str(song["umh_number"]).strip()
    title = str(song["title"]).strip()
    slides = song["slides"]

    lyrics_font_size_pt = song.get("lyrics_font_size_pt")
    line_spacing = song.get("line_spacing")

    full_title = f"UMH {umh_number} {title}".strip() if umh_number else title

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


def copy_slide_approx(source_prs, dest_prs, slide_index):
    source_slide = source_prs.slides[slide_index]
    blank_layout = dest_prs.slide_layouts[0]
    new_slide = dest_prs.slides.add_slide(blank_layout)

    for shape in list(new_slide.shapes):
        sp = shape._element
        sp.getparent().remove(sp)

    for shape in source_slide.shapes:
        new_el = deepcopy(shape._element)
        new_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")

    return new_slide


def create_single_song_ppt(song_item, template_bytes: bytes):
    prs = open_presentation_from_bytes(template_bytes)

    first_layout = get_layout_by_name(prs, FIRST_LAYOUT_NAME)
    rest_layout = get_layout_by_name(prs, REST_LAYOUT_NAME)

    if first_layout is None or rest_layout is None:
        raise ValueError("Template layouts not found.")

    delete_all_slides(prs)
    add_song_to_presentation(prs, song_item, first_layout, rest_layout)

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output


def create_combined_ppt_from_sections(service_sections, template_bytes: bytes):
    original_prs = open_presentation_from_bytes(template_bytes)
    first_layout = get_layout_by_name(original_prs, FIRST_LAYOUT_NAME)
    rest_layout = get_layout_by_name(original_prs, REST_LAYOUT_NAME)

    if first_layout is None or rest_layout is None:
        raise ValueError("Template layouts not found.")

    rebuilt = open_presentation_from_bytes(template_bytes)
    delete_all_slides(rebuilt)

    if not st.session_state.get("preserve_template_slides", True):
        for sec in service_sections:
            for song in sec.get("songs", []):
                add_song_to_presentation(rebuilt, song, first_layout, rest_layout)

        output = BytesIO()
        rebuilt.save(output)
        output.seek(0)
        return output

    for sec in service_sections:
        if sec.get("include_divider", True) and sec.get("divider_index") is not None:
            copy_slide_approx(original_prs, rebuilt, sec["divider_index"])

        if sec.get("include_template_content", True):
            for slide_idx in sec.get("template_slide_indices", []):
                copy_slide_approx(original_prs, rebuilt, slide_idx)

        for song in sec.get("songs", []):
            add_song_to_presentation(rebuilt, song, first_layout, rest_layout)

    output = BytesIO()
    rebuilt.save(output)
    output.seek(0)
    return output


def get_service_song_start_slides_from_sections(sections):
    starts = []
    slide_counter = 1

    for sec in sections:
        if sec.get("include_divider", True) and sec.get("divider_index") is not None:
            slide_counter += 1

        if sec.get("include_template_content", True):
            slide_counter += len(sec.get("template_slide_indices", []))

        for song in sec.get("songs", []):
            starts.append({
                "section_id": sec["id"],
                "title": song["title"],
                "start_slide": slide_counter,
            })
            slide_counter += len(song["slides"])

    return starts


def refresh_service_preview(service_sections, template_bytes):
    ppt_data = create_combined_ppt_from_sections(service_sections, template_bytes)
    preview_images = pptx_to_preview_images(ppt_data)

    if not preview_images:
        raise RuntimeError("No preview images generated from service PPT")

    st.session_state["ppt_data"] = ppt_data
    st.session_state["service_preview_images"] = preview_images
    st.session_state["service_song_start_slides"] = get_service_song_start_slides_from_sections(service_sections)


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

        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice conversion failed:\n{result.stderr}")

        pdf_files = [f for f in os.listdir(tmpdir) if f.lower().endswith(".pdf")]
        if not pdf_files:
            raise FileNotFoundError(
                f"No PDF created.\nstdout={result.stdout}\nstderr={result.stderr}"
            )

        pdf_path = os.path.join(tmpdir, pdf_files[0])
        doc = fitz.open(pdf_path)

        images = []
        for page in doc:
            pix = page.get_pixmap(dpi=100)
            mode = "RGB" if pix.alpha == 0 else "RGBA"
            img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)

            if mode == "RGBA":
                img = img.convert("RGB")

            img = img.resize(
                (int(pix.width * 0.7), int(pix.height * 0.7)),
                Image.LANCZOS,
            )

            buffer = io.BytesIO()
            img.save(buffer, format="JPEG", quality=70, optimize=True)
            images.append(buffer.getvalue())

        doc.close()
        return images


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
# EDITOR / AUTO PREVIEW HELPERS
# =========================================================
def refresh_current_song_preview(song_item, template_bytes):
    ppt_data = create_single_song_ppt(song_item, template_bytes)
    preview_images = pptx_to_preview_images(ppt_data)

    st.session_state["current_song_preview_images"] = preview_images
    st.session_state["last_current_song_signature"] = build_current_song_signature(
        song_item,
        st.session_state.get("selected_template_name"),
    )


def load_song_into_editor(match):
    lyrics_raw = str(match.get("Lyrics (Raw)", "")).strip()

    st.session_state["loaded_song"] = {
        "umh_number": str(match.get("UMH Number", "")).strip(),
        "title": str(match.get("Title", "")).strip(),
        "lyrics_raw": lyrics_raw,
    }

    st.session_state["editing_song_location"] = None
    st.session_state["editor_umh"] = str(match.get("UMH Number", "")).strip()
    st.session_state["editor_title"] = str(match.get("Title", "")).strip()
    st.session_state["editor_text"] = lyrics_raw

    st.session_state["editor_override_lyrics_font_size"] = False
    st.session_state["editor_override_line_spacing"] = False
    st.session_state["editor_lyrics_font_size_pt"] = 32
    st.session_state["editor_line_spacing"] = 1.2
    st.session_state["editor_target_section_id"] = st.session_state.get("selected_section_id")

    st.session_state["current_song_preview_images"] = None
    st.session_state["last_editor_text"] = lyrics_raw
    st.session_state["last_current_song_signature"] = None
    st.session_state["editor_status_message"] = ""
    st.session_state["current_preview_slide"] = 1
    st.session_state["preview_mode"] = "song"
    st.session_state["editor_ace_key"] += 1

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
                current_slides = get_current_slides(st.session_state["editor_text"])
                if current_slides:
                    song_item = build_editor_song_item(current_slides)
                    refresh_current_song_preview(song_item, template_bytes)
                    st.session_state["editor_status_message"] = "Song preview loaded."
                else:
                    st.session_state["editor_status_message"] = "Song loaded, but no slides detected for preview."
            else:
                st.session_state["editor_status_message"] = "Song loaded, but selected template is invalid."
        except Exception as e:
            st.session_state["editor_status_message"] = f"Preview load failed: {e}"
    else:
        if not st.session_state.get("selected_template_name"):
            st.session_state["editor_status_message"] = "Song loaded. ⚠️ Please select a template to generate preview."
        elif not soffice_available():
            st.session_state["editor_status_message"] = "Song loaded. LibreOffice/soffice is not available."


def load_section_song_into_editor(section_id, song_index):
    sec = get_section_by_id(section_id)
    if not sec or not (0 <= song_index < len(sec["songs"])):
        return

    item = sec["songs"][song_index]
    lyrics_text = "\n\n".join("\n".join(slide) for slide in item["slides"])

    st.session_state["loaded_song"] = {
        "umh_number": item["umh_number"],
        "title": item["title"],
        "lyrics_raw": lyrics_text,
    }

    st.session_state["editing_song_location"] = {
        "section_id": section_id,
        "song_index": song_index,
    }

    st.session_state["editor_umh"] = item["umh_number"]
    st.session_state["editor_title"] = item["title"]
    st.session_state["editor_text"] = lyrics_text
    st.session_state["editor_override_lyrics_font_size"] = item.get("override_lyrics_font_size", False)
    st.session_state["editor_override_line_spacing"] = item.get("override_line_spacing", False)
    st.session_state["editor_lyrics_font_size_pt"] = item.get("lyrics_font_size_pt", 32) or 32
    st.session_state["editor_line_spacing"] = item.get("line_spacing", 1.2) or 1.2
    st.session_state["editor_target_section_id"] = section_id

    st.session_state["current_song_preview_images"] = None
    st.session_state["last_editor_text"] = lyrics_text
    st.session_state["last_current_song_signature"] = None
    st.session_state["editor_status_message"] = ""
    st.session_state["current_preview_slide"] = 1
    st.session_state["preview_mode"] = "song"
    st.session_state["editor_ace_key"] += 1


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

    for idx, raw_line in enumerate(lines):
        stripped = raw_line.strip()

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
# PRE-RUN ACTIONS
# =========================================================
if st.session_state.get("reset_editor_pending"):
    reset_editor()
    st.session_state["reset_editor_pending"] = False

selected_template_bytes, selected_template_ok, selected_template_errors, selected_template_warnings = selected_template_info()

if selected_template_bytes and selected_template_ok:
    try:
        template_sig = (
            st.session_state.get("selected_template_name"),
            len(selected_template_bytes),
        )
        if st.session_state.get("template_signature_for_sections") != template_sig:
            sync_sections_from_template(selected_template_bytes)
    except Exception:
        st.session_state["service_sections"] = []
else:
    st.session_state["service_sections"] = []
    st.session_state["selected_section_id"] = None


# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.header("Controls")

    with st.expander("1. Template", expanded=True):
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

            selected_name = st.selectbox(
                "Select template",
                template_names,
                index=default_index,
            )
            if selected_name != st.session_state["selected_template_name"]:
                st.session_state["selected_template_name"] = selected_name
                clear_service_outputs()
                st.session_state["current_song_preview_images"] = None
                st.session_state["last_current_song_signature"] = None
                st.rerun()
            else:
                st.session_state["selected_template_name"] = selected_name

            selected_template_bytes, selected_template_ok, selected_template_errors, selected_template_warnings = selected_template_info()

            if selected_template_ok:
                st.success("Template valid")
            else:
                st.error("Template invalid")
                for err in selected_template_errors:
                    st.write(f"- {err}")
                st.info(
                    "Required template structure:\n"
                    f"- layout: {FIRST_LAYOUT_NAME}\n"
                    f"- layout: {REST_LAYOUT_NAME}\n"
                    f"- layout: {DIVIDER_LAYOUT_NAME}\n"
                    f"- at least 1 slide using {DIVIDER_LAYOUT_NAME}"
                )

            if selected_template_warnings:
                for warn in selected_template_warnings:
                    st.warning(warn)

            if st.button("Remove selected template", use_container_width=True):
                del st.session_state["uploaded_templates"][st.session_state["selected_template_name"]]
                st.session_state["selected_template_name"] = None
                st.session_state["service_sections"] = []
                clear_service_outputs()
                st.rerun()

            st.divider()
            st.checkbox("Keep existing slides in template", key="preserve_template_slides")

            if selected_template_ok and st.session_state.get("service_sections"):
                st.markdown("#### Detected Service Sections")
                for sec in st.session_state["service_sections"]:
                    st.caption(
                        f'- {sec["title"]} '
                        f'({len(sec["template_slide_indices"])} content slide(s))'
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

    with st.expander("3. Service Sections", expanded=True):
        sections = st.session_state.get("service_sections", [])

        if not sections:
            st.caption("No sections detected yet.")
        else:
            selected_section_id = st.session_state.get("selected_section_id")
            if selected_section_id is None and sections:
                st.session_state["selected_section_id"] = sections[0]["id"]
                selected_section_id = sections[0]["id"]

            for sec in sections:
                is_selected = sec["id"] == selected_section_id
                prefix = "🔹 " if is_selected else ""

                uploaded_count = len(sec.get("template_slide_indices", []))
                keep_text = "shown" if sec.get("include_template_content", True) else "hidden"

                st.markdown(f"**{prefix}{sec['title']}**")
                st.caption(f"Uploaded content: {uploaded_count} slide(s) · {keep_text}")

                row1, row2, row3 = st.columns(3)

                with row1:
                    if st.button("Select", key=f"select_sec_{sec['id']}", use_container_width=True):
                        st.session_state["selected_section_id"] = sec["id"]
                        st.session_state["editor_target_section_id"] = sec["id"]
                        if st.session_state["preview_mode"] == "service":
                            st.rerun()

                with row2:
                    toggle_label = "Hide content" if sec.get("include_template_content", True) else "Show content"
                    if st.button(toggle_label, key=f"toggle_content_{sec['id']}", use_container_width=True):
                        sec["include_template_content"] = not sec.get("include_template_content", True)
                        clear_service_outputs()
                        st.rerun()

                with row3:
                    toggle_divider_label = "Hide divider" if sec.get("include_divider", True) else "Show divider"
                    if st.button(toggle_divider_label, key=f"toggle_divider_{sec['id']}", use_container_width=True):
                        sec["include_divider"] = not sec.get("include_divider", True)
                        clear_service_outputs()
                        st.rerun()

                if sec.get("songs"):
                    for i, song in enumerate(sec["songs"]):
                        label = (
                            f'• UMH {song["umh_number"]} {song["title"]}'
                            if song["umh_number"] else f'• {song["title"]}'
                        )
                        st.write(label)

                        c1, c2, c3, c4 = st.columns(4)
                        with c1:
                            if st.button("Edit", key=f"edit_{sec['id']}_{i}", use_container_width=True):
                                load_section_song_into_editor(sec["id"], i)
                                st.rerun()

                        with c2:
                            if st.button("Up", key=f"up_{sec['id']}_{i}", use_container_width=True) and i > 0:
                                sec["songs"][i - 1], sec["songs"][i] = sec["songs"][i], sec["songs"][i - 1]
                                clear_service_outputs()
                                st.rerun()

                        with c3:
                            if st.button("Down", key=f"down_{sec['id']}_{i}", use_container_width=True) and i < len(sec["songs"]) - 1:
                                sec["songs"][i + 1], sec["songs"][i] = sec["songs"][i], sec["songs"][i + 1]
                                clear_service_outputs()
                                st.rerun()

                        with c4:
                            if st.button("Del", key=f"del_{sec['id']}_{i}", use_container_width=True):
                                sec["songs"].pop(i)
                                clear_service_outputs()
                                edit_loc = st.session_state.get("editing_song_location")
                                if (
                                    edit_loc is not None
                                    and edit_loc["section_id"] == sec["id"]
                                    and edit_loc["song_index"] == i
                                ):
                                    st.session_state["reset_editor_pending"] = True
                                st.rerun()
                else:
                    st.caption("No songs in this section.")

                st.divider()

            if st.button("Clear All Section Songs", use_container_width=True, type="secondary"):
                for sec in sections:
                    sec["songs"] = []
                clear_service_outputs()
                st.rerun()

    st.divider()
    st.markdown("### Loaded in Editor")

    loaded_umh = st.session_state.get("editor_umh", "").strip()
    loaded_title = st.session_state.get("editor_title", "").strip()
    editing_loc = st.session_state.get("editing_song_location")
    target_sec = get_section_by_id(st.session_state.get("editor_target_section_id"))

    if loaded_title:
        if editing_loc is not None:
            sec = get_section_by_id(editing_loc["section_id"])
            sec_name = sec["title"] if sec else "Unknown section"
            if loaded_umh:
                st.info(f'Editing "{sec_name}"\n\nUMH {loaded_umh} {loaded_title}')
            else:
                st.info(f'Editing "{sec_name}"\n\n{loaded_title}')
        else:
            target_name = target_sec["title"] if target_sec else "No section selected"
            if loaded_umh:
                st.info(f'New / repository song\n\nUMH {loaded_umh} {loaded_title}\n\nTarget: {target_name}')
            else:
                st.info(f'New / repository song\n\n{loaded_title}\n\nTarget: {target_name}')
    else:
        st.caption("No song loaded in editor.")


# =========================================================
# MAIN LAYOUT
# =========================================================
main_left, main_right = st.columns([1.15, 1], vertical_alignment="top")

with main_left:
    st.subheader("Song Editor")

    edit_loc = st.session_state.get("editing_song_location")
    if edit_loc is not None:
        sec = get_section_by_id(edit_loc["section_id"])
        sec_name = sec["title"] if sec else "Unknown section"
        st.info(f"Editing song in section: {sec_name}")

    meta_col1, meta_col2 = st.columns([1, 3])
    with meta_col1:
        st.text_input("UMH", key="editor_umh")
    with meta_col2:
        st.text_input("Title", key="editor_title")

    sections = st.session_state.get("service_sections", [])
    if sections:
        section_ids = [sec["id"] for sec in sections]
        section_name_map = {sec["id"]: sec["title"] for sec in sections}

        current_target = st.session_state.get("editor_target_section_id")
        if current_target not in section_ids:
            current_target = section_ids[0]
            st.session_state["editor_target_section_id"] = current_target

        st.selectbox(
            "Add / move song to section",
            options=section_ids,
            format_func=lambda sid: section_name_map[sid],
            key="editor_target_section_id",
        )
    else:
        st.caption("No service sections detected from template.")

    st.markdown("#### Slide Splitting")
    split_col1, split_col2 = st.columns([3, 1])

    with split_col1:
        st.checkbox("Auto split by lines per slide", key="auto_split_by_lines")
        st.slider("Lines per slide", min_value=1, max_value=8, key="lines_per_slide")
        st.checkbox(
            "Refresh preview only when a new slide break is detected",
            key="refresh_on_new_line",
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

    current_slides = get_current_slides(editor_text)
    song_item = build_editor_song_item(current_slides)

    if st.session_state["auto_split_by_lines"]:
        st.caption(
            f"{len(current_slides)} slide(s) "
            f"({st.session_state['lines_per_slide']} lines per slide, blank lines kept as verse separators)"
        )
    else:
        st.caption(f"{len(current_slides)} slide(s) (manual mode: blank lines separate slides)")

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
                refresh_current_song_preview(song_item, selected_template_bytes)
                st.session_state["editor_status_message"] = "Song preview refreshed."
                st.rerun()
            except Exception as e:
                st.error(f"Preview generation failed: {e}")

    new_signature = build_current_song_signature(
        song_item,
        st.session_state.get("selected_template_name"),
    )

    text_changed = editor_text != old_text
    trigger_refresh = False
    current_split_settings = (
        st.session_state["auto_split_by_lines"],
        st.session_state["lines_per_slide"],
    )

    split_settings_changed = (
        current_split_settings != st.session_state.get("last_split_settings")
    )

    if st.session_state["auto_split_by_lines"]:
        trigger_refresh = get_current_slides(old_text) != get_current_slides(editor_text)
    else:
        trigger_refresh = blank_separator_added(old_text, editor_text)

    if text_changed and trigger_refresh:
        if not st.session_state["auto_split_by_lines"]:
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
                    lines_per_slide=st.session_state["lines_per_slide"],
                )
                if detected_slide is not None:
                    st.session_state["current_preview_slide"] = detected_slide
        else:
            target_line_index = detect_new_slide_target_line(old_text, editor_text)
            detected_slide = get_slide_number_from_line_index(
                editor_text,
                target_line_index,
                auto_split=True,
                lines_per_slide=st.session_state["lines_per_slide"],
            )
            if detected_slide is not None:
                st.session_state["current_preview_slide"] = detected_slide

    should_refresh_preview = (
        (
            text_changed and (
                (st.session_state["refresh_on_new_line"] and trigger_refresh)
                or (not st.session_state["refresh_on_new_line"])
            )
        )
        or split_settings_changed
    ) and (
        selected_template_bytes is not None
        and selected_template_ok
        and soffice_available()
        and bool(current_slides)
        and new_signature != st.session_state.get("last_current_song_signature")
    )

    if should_refresh_preview:
        try:
            refresh_current_song_preview(song_item, selected_template_bytes)
            st.session_state["editor_status_message"] = "Song preview auto-refreshed."
        except Exception as e:
            st.session_state["editor_status_message"] = f"Preview refresh failed: {e}"

    st.session_state["last_editor_text"] = editor_text
    st.session_state["last_split_settings"] = current_split_settings

    if st.session_state["editor_status_message"]:
        if "⚠️" in st.session_state["editor_status_message"]:
            st.warning(st.session_state["editor_status_message"])
        else:
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
    action_col1, action_col2 = st.columns(2)

    edit_location = st.session_state.get("editing_song_location")

    with action_col1:
        add_or_update = st.button(
            "Update Song" if edit_location is not None else "Add to Section",
            use_container_width=True,
        )

    with action_col2:
        clear_editor = st.button("Clear Editor", use_container_width=True)

    if add_or_update:
        if not current_slides:
            st.error("No slides to add.")
        else:
            item = song_item
            target_section_id = item.get("target_section_id")
            target_sec = get_section_by_id(target_section_id)

            if target_sec is None:
                st.error("Please select a valid section.")
            else:
                edit_location = st.session_state.get("editing_song_location")

                if edit_location is None:
                    target_sec["songs"].append(item)
                    clear_service_outputs()
                    st.success(f'Added to "{target_sec["title"]}": {item["title"]}')
                    st.session_state["reset_editor_pending"] = True
                    st.rerun()
                else:
                    old_sec = get_section_by_id(edit_location["section_id"])
                    old_idx = edit_location["song_index"]

                    if old_sec is None or not (0 <= old_idx < len(old_sec["songs"])):
                        st.error("Original song location not found.")
                    else:
                        if old_sec["id"] == target_sec["id"]:
                            old_sec["songs"][old_idx] = item
                        else:
                            old_sec["songs"].pop(old_idx)
                            target_sec["songs"].append(item)

                        st.session_state["editing_song_location"] = None
                        clear_service_outputs()

                        if (
                            selected_template_bytes is not None
                            and selected_template_ok
                            and soffice_available()
                            and current_slides
                        ):
                            try:
                                refresh_current_song_preview(item, selected_template_bytes)
                                st.session_state["preview_mode"] = "song"
                                st.session_state["editor_status_message"] = "Song updated and preview refreshed."
                            except Exception as e:
                                st.session_state["editor_status_message"] = f"Preview refresh failed: {e}"

                        st.success(f'Updated: {item["title"]}')
                        st.rerun()

    if clear_editor:
        st.session_state["reset_editor_pending"] = True
        st.rerun()


with main_right:
    st.subheader("Preview")

    preview_mode = st.radio(
        "Preview Mode",
        ["Song", "Service"],
        index=0 if st.session_state["preview_mode"] == "song" else 1,
        horizontal=True,
    )
    st.session_state["preview_mode"] = preview_mode.lower()

    preview_images = None

    if st.session_state["preview_mode"] == "song":
        st.caption("Current song preview")

        if selected_template_bytes is None:
            st.warning("⚠️ Please upload and select a template to see preview.")
        elif not selected_template_ok:
            st.error("⚠️ Selected template is invalid.")
        elif not soffice_available():
            st.warning("⚠️ LibreOffice/soffice is required for preview.")

        preview_images = st.session_state.get("current_song_preview_images")

    else:
        st.caption("Full service deck preview")

        sections = st.session_state.get("service_sections", [])
        can_generate = (
            selected_template_bytes is not None
            and selected_template_ok
            and soffice_available()
            and len(sections) > 0
        )

        if can_generate:
            refresh_service_now = st.button(
                "Refresh Service Preview",
                use_container_width=True
            )

            need_refresh = (
                refresh_service_now
                or st.session_state.get("service_preview_images") is None
                or st.session_state.get("ppt_data") is None
                or not st.session_state.get("service_song_start_slides")
            )

            if need_refresh:
                try:
                    refresh_service_preview(sections, selected_template_bytes)
                except Exception as e:
                    st.error(f"Service preview generation failed: {e}")

            selected_section_id = st.session_state.get("selected_section_id")
            starts = st.session_state.get("service_song_start_slides", [])

            # highlight first song in selected section; otherwise top of deck
            current_slide = 1
            found = next((x for x in starts if x["section_id"] == selected_section_id), None)
            if found is not None:
                current_slide = found["start_slide"]

            st.session_state["current_preview_slide"] = current_slide
            preview_images = st.session_state.get("service_preview_images")

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
            elif not st.session_state["service_sections"]:
                st.info("No service sections detected from template.")

    if preview_images:
        render_scrollable_images(
            preview_images,
            height=700,
            active_slide=st.session_state.get("current_preview_slide"),
        )
    else:
        st.info("Preview will appear here.")
