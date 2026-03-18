import os
import base64
import tempfile
import subprocess
from io import BytesIO
from shutil import which

import fitz  # PyMuPDF
import gspread
import streamlit as st
from streamlit_ace import st_ace
from google.oauth2.service_account import Credentials
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

# =========================
# PATHS / SETTINGS
# =========================
SOFFICE_PATH = os.environ.get("SOFFICE_PATH", "soffice")

SHEET_KEY = st.secrets["SHEET_KEY"]
WORKSHEET_NAME = st.secrets["WORKSHEET_NAME"]

FIRST_LAYOUT_NAME = "TEMPLATE_FIRST"
REST_LAYOUT_NAME = "TEMPLATE_REST"

# =========================
# GOOGLE SHEETS
# =========================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

credentials = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=SCOPES,
)
gc = gspread.authorize(credentials)
sheet = gc.open_by_key(SHEET_KEY).worksheet(WORKSHEET_NAME)

# =========================
# STREAMLIT PAGE
# =========================
st.set_page_config(page_title="Service Deck Builder", layout="wide")
st.title("Service Deck Builder")

# =========================
# SESSION STATE INIT
# =========================
defaults = {
    "setlist": [],
    "editor_text": "",
    "editor_umh": "",
    "editor_title": "",
    "editor_override_title_font_size": False,
    "editor_override_lyrics_font_size": False,
    "editor_override_line_spacing": False,
    "editor_title_font_size_pt": 28,
    "editor_lyrics_font_size_pt": 32,
    "editor_line_spacing": 1.2,
    "loaded_song": None,
    "ppt_data": None,
    "preview_images": None,
    "current_preview_slide": None,
    "last_detected_edit_line": None,
    "current_song_preview_images": None,
    "editing_setlist_index": None,
    "pending_setlist_load": None,
    "uploaded_templates": {},
    "selected_template_name": None,
    "reset_editor_pending": False,
    "auto_split_by_lines": True,
    "lines_per_slide": 4,
    "refresh_on_new_line": True,
    "last_editor_text": "",
    "last_current_song_signature": None,
    "editor_status_message": "",
    "editor_ace_key": 0,
    "refresh_preview_after_load": False
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# =========================
# HELPERS
# =========================
def soffice_available() -> bool:
    if SOFFICE_PATH == "soffice":
        return which("soffice") is not None
    return os.path.exists(SOFFICE_PATH)


def split_slides(text: str) -> list[list[str]]:
    blocks = [block.strip() for block in text.split("\n\n") if block.strip()]
    slides = []

    for block in blocks:
        lines = [line.strip() for line in block.splitlines() if line.strip()]
        if lines:
            slides.append(lines)

    return slides


def split_slides_by_line_count_with_verse_separators(
    text: str,
    lines_per_slide: int = 4
) -> list[list[str]]:
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
        return split_slides_by_line_count_with_verse_separators(
            text,
            lines_per_slide=st.session_state["lines_per_slide"],
        )
    return split_slides(text)


def find_row_by_umh(ws, umh_number: str):
    records = ws.get_all_records()
    for row in records:
        if str(row.get("UMH Number", "")).strip() == umh_number.strip():
            return row
    return None


def search_titles(ws, keyword: str):
    records = ws.get_all_records()
    keyword = keyword.lower().strip()
    matches = []

    for row in records:
        title = str(row.get("Title", "")).strip()
        if keyword and keyword in title.lower():
            matches.append(row)

    return matches[:20]


def open_presentation_from_bytes(template_bytes: bytes):
    return Presentation(BytesIO(template_bytes))


def get_layout_by_name(prs, layout_name):
    for slide_master in prs.slide_masters:
        for slide_layout in slide_master.slide_layouts:
            if slide_layout.name.strip() == layout_name:
                return slide_layout
    return None


def set_shape_text(shape, text, font_size_pt=None, line_spacing=None):
    if shape is None or not getattr(shape, "has_text_frame", False):
        return

    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True

    lines = text.split("\n")
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        p.alignment = PP_ALIGN.CENTER

        if line_spacing is not None:
            p.line_spacing = line_spacing

        run = p.add_run()
        run.text = line

        if font_size_pt is not None:
            run.font.size = Pt(font_size_pt)


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


def delete_all_slides(prs):
    while len(prs.slides) > 0:
        slide_id = prs.slides._sldIdLst[0]
        rId = slide_id.rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]


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

    return len(errors) == 0, errors, warnings


def create_combined_ppt(setlist, template_bytes: bytes):
    prs = open_presentation_from_bytes(template_bytes)

    first_layout = get_layout_by_name(prs, FIRST_LAYOUT_NAME)
    rest_layout = get_layout_by_name(prs, REST_LAYOUT_NAME)

    if first_layout is None or rest_layout is None:
        raise ValueError("Template layouts not found.")

    delete_all_slides(prs)

    for song in setlist:
        umh_number = str(song["umh_number"]).strip()
        title = str(song["title"]).strip()
        slides = song["slides"]

        title_font_size_pt = song.get("title_font_size_pt")
        lyrics_font_size_pt = song.get("lyrics_font_size_pt")
        line_spacing = song.get("line_spacing")

        full_title = f"UMH {umh_number} {title}".strip() if umh_number else title

        for i, slide_lines in enumerate(slides):
            lyrics_text = "\n".join(slide_lines)

            if i == 0:
                new_slide = prs.slides.add_slide(first_layout)
                title_placeholder = new_slide.shapes.title
                body_placeholder = get_body_placeholder(new_slide)

                set_shape_text(
                    title_placeholder,
                    full_title,
                    font_size_pt=title_font_size_pt,
                    line_spacing=line_spacing,
                )
                set_shape_text(
                    body_placeholder,
                    lyrics_text,
                    font_size_pt=lyrics_font_size_pt,
                    line_spacing=line_spacing,
                )
            else:
                new_slide = prs.slides.add_slide(rest_layout)
                body_placeholder = get_body_placeholder(new_slide)
                set_shape_text(
                    body_placeholder,
                    lyrics_text,
                    font_size_pt=lyrics_font_size_pt,
                    line_spacing=line_spacing,
                )

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output


def create_single_song_ppt(song_item, template_bytes: bytes):
    return create_combined_ppt([song_item], template_bytes)


def pptx_to_preview_images(pptx_bytes):
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
            pix = page.get_pixmap(dpi=160)
            img_path = os.path.join(tmpdir, f"slide_{page.number + 1}.png")
            pix.save(img_path)
            with open(img_path, "rb") as f:
                images.append(f.read())

        doc.close()
        return images


def render_scrollable_images(images, height=760, active_slide=None):
    html = f"""
    <div style="
        height: {height}px;
        overflow-y: auto;
        border: 1px solid #ddd;
        padding: 12px;
        border-radius: 8px;
        background: #fafafa;
        box-sizing: border-box;
    ">
    """

    for i, img_bytes in enumerate(images, start=1):
        b64 = base64.b64encode(img_bytes).decode("utf-8")
        border = "3px solid #2563eb" if active_slide == i else "1px solid #ccc"
        badge = " ← editing here" if active_slide == i else ""

        html += f"""
        <div style="margin-bottom: 24px;">
            <div style="font-weight: 600; margin-bottom: 8px;">Slide {i}{badge}</div>
            <img
                src="data:image/png;base64,{b64}"
                style="width: 100%; border: {border}; display: block;"
            />
        </div>
        """

    html += "</div>"
    st.components.v1.html(html, height=height, scrolling=False)

def build_editor_song_item(current_slides):
    return {
        "umh_number": st.session_state.get("editor_umh", "").strip(),
        "title": st.session_state.get("editor_title", "").strip(),
        "slides": current_slides,
        "title_font_size_pt": (
            st.session_state.get("editor_title_font_size_pt")
            if st.session_state.get("editor_override_title_font_size")
            else None
        ),
        "lyrics_font_size_pt": (
            st.session_state.get("editor_lyrics_font_size_pt")
            if st.session_state.get("editor_override_lyrics_font_size")
            else None
        ),
        "line_spacing": (
            st.session_state.get("editor_line_spacing")
            if st.session_state.get("editor_override_line_spacing")
            else None
        ),
        "override_title_font_size": st.session_state.get("editor_override_title_font_size"),
        "override_lyrics_font_size": st.session_state.get("editor_override_lyrics_font_size"),
        "override_line_spacing": st.session_state.get("editor_override_line_spacing"),
    }


def build_current_song_signature(song_item, selected_template_name):
    return (
        song_item["umh_number"],
        song_item["title"],
        tuple(tuple(slide) for slide in song_item["slides"]),
        song_item.get("title_font_size_pt"),
        song_item.get("lyrics_font_size_pt"),
        song_item.get("line_spacing"),
        selected_template_name,
        st.session_state.get("auto_split_by_lines"),
        st.session_state.get("lines_per_slide"),
    )


def should_refresh_on_new_line(old_text: str, new_text: str) -> bool:
    return new_text.count("\n") > old_text.count("\n")


def refresh_current_song_preview(song_item, template_bytes):
    ppt_data = create_single_song_ppt(song_item, template_bytes)
    preview_images = pptx_to_preview_images(ppt_data)
    st.session_state["current_song_preview_images"] = preview_images
    st.session_state["last_current_song_signature"] = build_current_song_signature(
        song_item,
        st.session_state.get("selected_template_name"),
    )


def load_song_into_editor_from_repository(match):
    song = {
        "umh_number": str(match.get("UMH Number", "")).strip(),
        "title": str(match.get("Title", "")).strip(),
        "lyrics_raw": str(match.get("Lyrics (Raw)", "")).strip(),
    }
    new_text = song["lyrics_raw"]

    st.session_state["loaded_song"] = song
    st.session_state["editor_umh"] = song["umh_number"]
    st.session_state["editor_title"] = song["title"]
    st.session_state["editor_text"] = new_text
    st.session_state["editing_setlist_index"] = None

    st.session_state["editor_override_title_font_size"] = False
    st.session_state["editor_override_lyrics_font_size"] = False
    st.session_state["editor_override_line_spacing"] = False
    st.session_state["editor_title_font_size_pt"] = 28
    st.session_state["editor_lyrics_font_size_pt"] = 32
    st.session_state["editor_line_spacing"] = 1.2

    st.session_state["ppt_data"] = None
    st.session_state["preview_images"] = None
    st.session_state["current_song_preview_images"] = None
    st.session_state["last_editor_text"] = new_text
    st.session_state["last_current_song_signature"] = None
    st.session_state["editor_status_message"] = ""
    st.session_state["current_preview_slide"] = 1
    st.session_state["last_detected_edit_line"] = None
    st.session_state["editor_ace_key"] += 1

    # Generate current-song preview immediately on load
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
                current_slides = (
                    split_slides_by_line_count_with_verse_separators(
                        new_text,
                        lines_per_slide=st.session_state["lines_per_slide"],
                    )
                    if st.session_state["auto_split_by_lines"]
                    else split_slides(new_text)
                )

                if current_slides:
                    song_item = {
                        "umh_number": st.session_state["editor_umh"].strip(),
                        "title": st.session_state["editor_title"].strip(),
                        "slides": current_slides,
                        "title_font_size_pt": None,
                        "lyrics_font_size_pt": None,
                        "line_spacing": None,
                        "override_title_font_size": False,
                        "override_lyrics_font_size": False,
                        "override_line_spacing": False,
                    }
                    refresh_current_song_preview(song_item, template_bytes)
                    st.session_state["editor_status_message"] = "Current-song preview refreshed."
        except Exception as e:
            st.session_state["editor_status_message"] = f"Preview refresh failed: {e}"

def apply_pending_setlist_load():
    pending = st.session_state.get("pending_setlist_load")
    if pending is None:
        return
    if pending >= len(st.session_state["setlist"]):
        st.session_state["pending_setlist_load"] = None
        return

    item = st.session_state["setlist"][pending]
    lyrics_text = "\n\n".join("\n".join(slide) for slide in item["slides"])

    st.session_state["loaded_song"] = {
        "umh_number": item["umh_number"],
        "title": item["title"],
        "lyrics_raw": lyrics_text,
    }
    st.session_state["editor_umh"] = item["umh_number"]
    st.session_state["editor_title"] = item["title"]
    st.session_state["editor_text"] = lyrics_text
    st.session_state["editing_setlist_index"] = pending

    st.session_state["editor_override_title_font_size"] = item.get(
        "override_title_font_size", False
    )
    st.session_state["editor_override_lyrics_font_size"] = item.get(
        "override_lyrics_font_size", False
    )
    st.session_state["editor_override_line_spacing"] = item.get(
        "override_line_spacing", False
    )

    st.session_state["editor_title_font_size_pt"] = (
        item.get("title_font_size_pt", 28) or 28
    )
    st.session_state["editor_lyrics_font_size_pt"] = (
        item.get("lyrics_font_size_pt", 32) or 32
    )
    st.session_state["editor_line_spacing"] = (
        item.get("line_spacing", 1.2) or 1.2
    )

    st.session_state["ppt_data"] = None
    st.session_state["preview_images"] = None
    st.session_state["current_song_preview_images"] = None
    st.session_state["pending_setlist_load"] = None
    st.session_state["last_editor_text"] = lyrics_text
    st.session_state["last_current_song_signature"] = None
    st.session_state["editor_status_message"] = ""
    st.session_state["editor_ace_key"] += 1
    st.session_state["current_preview_slide"] = None


def reset_editor_for_new_song():
    st.session_state["loaded_song"] = None
    st.session_state["editor_umh"] = ""
    st.session_state["editor_title"] = ""
    st.session_state["editor_text"] = ""
    st.session_state["editing_setlist_index"] = None

    st.session_state["editor_override_title_font_size"] = False
    st.session_state["editor_override_lyrics_font_size"] = False
    st.session_state["editor_override_line_spacing"] = False
    st.session_state["editor_title_font_size_pt"] = 28
    st.session_state["editor_lyrics_font_size_pt"] = 32
    st.session_state["editor_line_spacing"] = 1.2

    st.session_state["ppt_data"] = None
    st.session_state["preview_images"] = None
    st.session_state["current_song_preview_images"] = None
    st.session_state["last_editor_text"] = ""
    st.session_state["last_current_song_signature"] = None
    st.session_state["editor_status_message"] = ""
    st.session_state["editor_ace_key"] += 1
    st.session_state["current_preview_slide"] = None

def detect_changed_line_index(old_text: str, new_text: str):
    old_lines = old_text.splitlines()
    new_lines = new_text.splitlines()

    min_len = min(len(old_lines), len(new_lines))

    for i in range(min_len):
        if old_lines[i] != new_lines[i]:
            return i

    if len(new_lines) > len(old_lines):
        return len(old_lines)

    return None


def get_slide_number_from_line_index(
    text: str,
    line_index: int,
    auto_split: bool,
    lines_per_slide: int
):
    if line_index is None:
        return None

    lines = text.splitlines()

    if auto_split:
        current_verse_line_indexes = []
        line_to_slide = {}
        slide_num = 1

        for idx, raw_line in enumerate(lines):
            stripped = raw_line.strip()

            if stripped == "":
                if current_verse_line_indexes:
                    for j in range(0, len(current_verse_line_indexes), lines_per_slide):
                        chunk = current_verse_line_indexes[j:j + lines_per_slide]
                        for original_idx in chunk:
                            line_to_slide[original_idx] = slide_num
                        slide_num += 1
                    current_verse_line_indexes = []
            else:
                current_verse_line_indexes.append(idx)

        if current_verse_line_indexes:
            for j in range(0, len(current_verse_line_indexes), lines_per_slide):
                chunk = current_verse_line_indexes[j:j + lines_per_slide]
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

def count_blank_lines(text: str) -> int:
    return sum(1 for line in text.splitlines() if line.strip() == "")

def blank_line_added(old_text: str, new_text: str) -> bool:
    return count_blank_lines(new_text) > count_blank_lines(old_text)

def slide_count_changed(old_text: str, new_text: str) -> bool:
    old_slides = get_current_slides(old_text)
    new_slides = get_current_slides(new_text)
    return len(new_slides) != len(old_slides)

def get_next_nonblank_line_index(text: str, line_index: int):
    lines = text.splitlines()

    if line_index is None:
        return None

    for i in range(line_index + 1, len(lines)):
        if lines[i].strip() != "":
            return i

    return None

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

    return changed_idx

def get_slide_number_from_line_index(
    text: str,
    line_index: int,
    auto_split: bool,
    lines_per_slide: int
):
    if line_index is None:
        return None

    lines = text.splitlines()

    if auto_split:
        current_verse_line_indexes = []
        line_to_slide = {}
        slide_num = 1

        for idx, raw_line in enumerate(lines):
            stripped = raw_line.strip()

            if stripped == "":
                if current_verse_line_indexes:
                    for j in range(0, len(current_verse_line_indexes), lines_per_slide):
                        chunk = current_verse_line_indexes[j:j + lines_per_slide]
                        for original_idx in chunk:
                            line_to_slide[original_idx] = slide_num
                        slide_num += 1
                    current_verse_line_indexes = []
            else:
                current_verse_line_indexes.append(idx)

        if current_verse_line_indexes:
            for j in range(0, len(current_verse_line_indexes), lines_per_slide):
                chunk = current_verse_line_indexes[j:j + lines_per_slide]
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
    

# Must happen before widgets are created
apply_pending_setlist_load()

if st.session_state.get("reset_editor_pending"):
    reset_editor_for_new_song()
    st.session_state["reset_editor_pending"] = False

selected_template_bytes = None
selected_template_ok = False
selected_template_errors = []
selected_template_warnings = []

# =========================
# ROW 1 — TEMPLATE
# =========================
with st.container():
    st.subheader("Template")

    uploaded_templates = st.file_uploader(
        "Upload one or more Template.pptx files",
        type=["pptx"],
        accept_multiple_files=True,
        key="template_uploader"
    )

    if uploaded_templates:
        for file in uploaded_templates:
            st.session_state["uploaded_templates"][file.name] = file.getvalue()

    template_names = list(st.session_state["uploaded_templates"].keys())

    if template_names:
        default_index = 0
        if st.session_state["selected_template_name"] in template_names:
            default_index = template_names.index(st.session_state["selected_template_name"])

        selected_template_name = st.selectbox(
            "Select template",
            template_names,
            index=default_index
        )
        st.session_state["selected_template_name"] = selected_template_name
        selected_template_bytes = st.session_state["uploaded_templates"][selected_template_name]

        selected_template_ok, selected_template_errors, selected_template_warnings = (
            validate_template_bytes(selected_template_bytes)
        )

        if selected_template_ok:
            st.success(f"Template is usable: {selected_template_name}")
        else:
            st.error(f"Template is invalid: {selected_template_name}")
            for err in selected_template_errors:
                st.write(f"- {err}")

        if selected_template_warnings:
            st.warning("Template warnings:")
            for warn in selected_template_warnings:
                st.write(f"- {warn}")

        if st.button("Remove Selected Template"):
            del st.session_state["uploaded_templates"][selected_template_name]
            st.session_state["selected_template_name"] = None
            st.rerun()
    else:
        st.info("Please upload at least one PowerPoint template to continue.")

    if not soffice_available():
        st.warning(
            "LibreOffice/soffice is not available. On Streamlit Community Cloud, add a packages.txt file with LibreOffice packages."
        )

# =========================
# ROW 2 — LOAD SONG | CURRENT SETLIST
# =========================
with st.container():
    load_col, setlist_col = st.columns([1.2, 1], vertical_alignment="top")

    with load_col:
        st.subheader("Load Song")

        if st.button("Start New Song"):
            st.session_state["reset_editor_pending"] = True
            st.rerun()

        load_mode = st.radio("Find hymn by", ["UMH Number", "Title"], horizontal=True)

        if load_mode == "UMH Number":
            umh_number_input = st.text_input("Enter UMH Number", placeholder="e.g. 57")

            if st.button("Load Hymn by Number"):
                if umh_number_input.strip():
                    match = find_row_by_umh(sheet, umh_number_input)
                    if match:
                        load_song_into_editor_from_repository(match)
                        st.success("Hymn loaded.")
                        st.rerun()
                    else:
                        st.error("Hymn not found.")

        else:
            keyword = st.text_input("Search title", placeholder="e.g. thousand tongues")

            if keyword.strip():
                matches = search_titles(sheet, keyword)
                if matches:
                    options = [
                        f'UMH {row.get("UMH Number","")} - {row.get("Title","")}'
                        for row in matches
                    ]
                    selected = st.selectbox("Select hymn", options)

                    if st.button("Load Hymn by Title"):
                        chosen_index = options.index(selected)
                        match = matches[chosen_index]
                        load_song_into_editor_from_repository(match)
                        st.success("Hymn loaded.")
                        st.rerun()
                else:
                    st.info("No matching titles found.")

    with setlist_col:
        st.subheader("Current Setlist")

        if st.session_state["setlist"]:
            remove_index = None

            for i, song in enumerate(st.session_state["setlist"]):
                if song["umh_number"]:
                    label = f'UMH {song["umh_number"]} {song["title"]}'
                else:
                    label = song["title"]

                total_slides = len(song["slides"])
                col_title, col_edit, col_up, col_down, col_delete = st.columns(
                    [12, 1, 1, 1, 1],
                    gap="small"
                )

                with col_title:
                    st.markdown(f"**{i+1}. {label} ({total_slides})**")

                with col_edit:
                    if st.button("✏️", key=f"edit_{i}"):
                        st.session_state["pending_setlist_load"] = i
                        st.rerun()

                with col_up:
                    if st.button("↑", key=f"up_{i}") and i > 0:
                        st.session_state["setlist"][i - 1], st.session_state["setlist"][i] = (
                            st.session_state["setlist"][i],
                            st.session_state["setlist"][i - 1],
                        )
                        st.session_state["ppt_data"] = None
                        st.session_state["preview_images"] = None
                        st.session_state["current_song_preview_images"] = None

                        current_edit = st.session_state.get("editing_setlist_index")
                        if current_edit == i:
                            st.session_state["editing_setlist_index"] = i - 1
                        elif current_edit == i - 1:
                            st.session_state["editing_setlist_index"] = i

                        st.rerun()

                with col_down:
                    if st.button("↓", key=f"down_{i}") and i < len(st.session_state["setlist"]) - 1:
                        st.session_state["setlist"][i + 1], st.session_state["setlist"][i] = (
                            st.session_state["setlist"][i],
                            st.session_state["setlist"][i + 1],
                        )
                        st.session_state["ppt_data"] = None
                        st.session_state["preview_images"] = None
                        st.session_state["current_song_preview_images"] = None

                        current_edit = st.session_state.get("editing_setlist_index")
                        if current_edit == i:
                            st.session_state["editing_setlist_index"] = i + 1
                        elif current_edit == i + 1:
                            st.session_state["editing_setlist_index"] = i

                        st.rerun()

                with col_delete:
                    if st.button("🗑", key=f"delete_{i}"):
                        remove_index = i

            if remove_index is not None:
                st.session_state["setlist"].pop(remove_index)
                st.session_state["ppt_data"] = None
                st.session_state["preview_images"] = None
                st.session_state["current_song_preview_images"] = None

                current_edit = st.session_state.get("editing_setlist_index")
                if current_edit == remove_index:
                    st.session_state["reset_editor_pending"] = True
                elif current_edit is not None and current_edit > remove_index:
                    st.session_state["editing_setlist_index"] = current_edit - 1

                pending = st.session_state.get("pending_setlist_load")
                if pending == remove_index:
                    st.session_state["pending_setlist_load"] = None
                elif pending is not None and pending > remove_index:
                    st.session_state["pending_setlist_load"] = pending - 1

                st.rerun()

            col1, col2 = st.columns(2)

            if col1.button("Generate Service Preview"):
                if selected_template_bytes is None:
                    st.error("Please upload and select a template first.")
                elif not selected_template_ok:
                    st.error("Cannot generate preview because the selected template is invalid.")
                else:
                    try:
                        ppt_data = create_combined_ppt(
                            st.session_state["setlist"],
                            selected_template_bytes,
                        )
                        st.session_state["ppt_data"] = ppt_data
                        st.success("Service preview generated.")
                    except Exception as e:
                        st.error(f"Preview generation failed: {e}")

            if col2.button("Clear Setlist"):
                st.session_state["setlist"] = []
                st.session_state["ppt_data"] = None
                st.session_state["preview_images"] = None
                st.session_state["current_song_preview_images"] = None
                st.session_state["editing_setlist_index"] = None
                st.session_state["pending_setlist_load"] = None
                st.session_state["reset_editor_pending"] = True
                st.rerun()

            if st.session_state["ppt_data"] is not None:
                download_data = (
                    st.session_state["ppt_data"].getvalue()
                    if hasattr(st.session_state["ppt_data"], "getvalue")
                    else st.session_state["ppt_data"]
                )
                st.download_button(
                    label="Download Service PowerPoint",
                    data=download_data,
                    file_name="service_deck.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
        else:
            st.info("No songs added yet.")

# =========================
# ROW 3 — SONG EDITOR | CURRENT SONG PREVIEW
# =========================
with st.container():
    editor_col, preview_col = st.columns([1.2, 1], vertical_alignment="top")

    with editor_col:
        st.subheader("Song Editor")

        edit_idx = st.session_state.get("editing_setlist_index")
        if edit_idx is not None:
            st.info(f"Editing setlist item #{edit_idx + 1}")

        col1, col2 = st.columns([1, 3])
        with col1:
            st.text_input("UMH", key="editor_umh")
        with col2:
            st.text_input("Title", key="editor_title")

        st.markdown("#### Slide Splitting")

        st.checkbox(
            "Auto split by lines per slide",
            key="auto_split_by_lines"
        )

        st.slider(
            "Lines per slide",
            min_value=1,
            max_value=8,
            key="lines_per_slide"
        )

        st.checkbox(
            "Refresh current-song preview only when slide count changes",
            key="refresh_on_new_line"
        )

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

        old_text = st.session_state.get("last_editor_text", "")
        st.session_state["editor_text"] = editor_text

        current_slides = get_current_slides(editor_text)

        if st.session_state["auto_split_by_lines"]:
            st.caption(
                f"{len(current_slides)} slide(s) for current song "
                f"({st.session_state['lines_per_slide']} lines per slide, blank lines kept as verse separators)"
            )
        else:
            st.caption(
                f"{len(current_slides)} slide(s) for current song "
                f"(manual mode: blank lines separate slides)"
            )

        st.markdown("#### Song Formatting")

        st.checkbox(
            "Override title font size for this song",
            key="editor_override_title_font_size"
        )
        if st.session_state["editor_override_title_font_size"]:
            st.slider(
                "Title font size (pt)",
                min_value=12,
                max_value=60,
                key="editor_title_font_size_pt"
            )
        else:
            st.caption("Title font size: using template default")

        st.checkbox(
            "Override lyrics font size for this song",
            key="editor_override_lyrics_font_size"
        )
        if st.session_state["editor_override_lyrics_font_size"]:
            st.slider(
                "Lyrics font size (pt)",
                min_value=12,
                max_value=60,
                key="editor_lyrics_font_size_pt"
            )
        else:
            st.caption("Lyrics font size: using template default")

        st.checkbox(
            "Override line spacing for this song",
            key="editor_override_line_spacing"
        )
        if st.session_state["editor_override_line_spacing"]:
            st.slider(
                "Line spacing",
                min_value=0.8,
                max_value=2.0,
                step=0.1,
                key="editor_line_spacing"
            )
        else:
            st.caption("Line spacing: using template default")

        song_item = build_editor_song_item(current_slides)
        new_signature = build_current_song_signature(
            song_item,
            st.session_state.get("selected_template_name"),
        )

        text_changed = editor_text != old_text
        old_slide_count = len(get_current_slides(old_text))
        new_slide_count = len(current_slides)
        slide_count_has_changed = old_slide_count != new_slide_count

        if text_changed and slide_count_has_changed:
            target_line_index = detect_new_slide_target_line(old_text, editor_text)

            detected_slide = get_slide_number_from_line_index(
                editor_text,
                target_line_index,
                st.session_state["auto_split_by_lines"],
                st.session_state["lines_per_slide"],
            )

            st.session_state["last_detected_edit_line"] = target_line_index

            if detected_slide is not None:
                st.session_state["current_preview_slide"] = detected_slide
            elif new_slide_count > 0:
                if st.session_state.get("current_preview_slide") is None:
                    st.session_state["current_preview_slide"] = 1
                else:
                    st.session_state["current_preview_slide"] = min(
                        st.session_state["current_preview_slide"],
                        new_slide_count
                    )

        should_refresh_preview = (
            text_changed
            and selected_template_bytes is not None
            and selected_template_ok
            and soffice_available()
            and bool(current_slides)
            and (
                (st.session_state["refresh_on_new_line"] and slide_count_has_changed)
                or (not st.session_state["refresh_on_new_line"])
            )
            and new_signature != st.session_state.get("last_current_song_signature")
        )

        if should_refresh_preview:
            try:
                refresh_current_song_preview(song_item, selected_template_bytes)
                st.session_state["editor_status_message"] = (
                    f"Current-song preview refreshed. "
                    f"Slides: {old_slide_count} → {new_slide_count}. "
                    f"Active slide: {st.session_state.get('current_preview_slide')}"
                )
            except Exception as e:
                st.session_state["editor_status_message"] = f"Preview refresh failed: {e}"

        st.session_state["last_editor_text"] = editor_text
        if st.session_state["editor_status_message"]:
            st.caption(st.session_state["editor_status_message"])

        if st.button("Refresh Current Song Preview"):
            if selected_template_bytes is None:
                st.error("Please upload and select a template first.")
            elif not selected_template_ok:
                st.error("Cannot preview because the selected template is invalid.")
            elif not soffice_available():
                st.error("LibreOffice/soffice is not available.")
            elif not current_slides:
                st.error("No slides to preview.")
            else:
                try:
                    refresh_current_song_preview(song_item, selected_template_bytes)
                    st.session_state["editor_status_message"] = "Current-song preview refreshed."
                    st.rerun()
                except Exception as e:
                    st.error(f"Preview generation failed: {e}")

        allow_duplicates = st.checkbox("Allow duplicate songs in setlist", value=False)

        button_label = (
            "Update Song in Setlist"
            if edit_idx is not None
            else "Add Song to Setlist"
        )

        if st.button(button_label):
            if current_slides:
                item = song_item
                edit_idx = st.session_state.get("editing_setlist_index")

                if edit_idx is None:
                    duplicate_index = next(
                        (
                            i for i, s in enumerate(st.session_state["setlist"])
                            if s["umh_number"] == item["umh_number"]
                            and s["title"] == item["title"]
                            and s["slides"] == item["slides"]
                        ),
                        None
                    )

                    if duplicate_index is not None and not allow_duplicates:
                        st.warning(
                            f"This song is already in the setlist as item #{duplicate_index + 1}."
                        )
                    else:
                        st.session_state["setlist"].append(item)
                        st.session_state["ppt_data"] = None
                        st.success(
                            f'Added: {"UMH " + item["umh_number"] + " " if item["umh_number"] else ""}{item["title"]}'
                        )
                        st.session_state["reset_editor_pending"] = True
                        st.rerun()
                else:
                    st.session_state["setlist"][edit_idx] = item
                    st.session_state["editing_setlist_index"] = None
                    st.session_state["ppt_data"] = None
                    st.success(
                        f'Updated: {"UMH " + item["umh_number"] + " " if item["umh_number"] else ""}{item["title"]}'
                    )
                    st.session_state["reset_editor_pending"] = True
                    st.rerun()
            else:
                st.error("No slides to add.")

        if st.button("Clear Current Editor"):
            st.session_state["reset_editor_pending"] = True
            st.rerun()

    with preview_col:
        st.subheader("Current Song Preview")

        if st.session_state.get("current_song_preview_images"):
            render_scrollable_images(
                st.session_state["current_song_preview_images"],
                height=600,
                active_slide=st.session_state.get("current_preview_slide"),
            )
        else:
            st.info(
                "The current-song preview will appear here after a hymn is loaded or refreshed."
            )

