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


# =========================================================
# CONFIG
# =========================================================
SOFFICE_PATH = os.environ.get("SOFFICE_PATH", "soffice")

SHEET_KEY = st.secrets["SHEET_KEY"]
WORKSHEET_NAME = st.secrets["WORKSHEET_NAME"]

FIRST_LAYOUT_NAME = "TEMPLATE_FIRST"
REST_LAYOUT_NAME = "TEMPLATE_REST"


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
    "editor_override_title_font_size": False,
    "editor_override_lyrics_font_size": False,
    "editor_override_line_spacing": False,
    "editor_title_font_size_pt": 28,
    "editor_lyrics_font_size_pt": 32,
    "editor_line_spacing": 1.2,
    "auto_split_by_lines": False,
    "lines_per_slide": 4,
    "refresh_on_new_line": True,
    "editor_ace_key": 0,
    "last_editor_text": "",
    "last_current_song_signature": None,
    "last_detected_edit_line": None,
    "editor_status_message": "",
    "current_preview_slide": 1,
    "preview_mode": "song",
    "current_song_preview_images": None,
    "service_preview_images": None,
    "service_song_start_slides": [],
    "ppt_data": None,
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# =========================================================
# HELPERS
# =========================================================
def soffice_available() -> bool:
    if SOFFICE_PATH == "soffice":
        return which("soffice") is not None
    return os.path.exists(SOFFICE_PATH)


@st.cache_data(show_spinner=False)
def get_all_records_cached():
    return sheet.get_all_records()


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
        st.session_state["auto_split_by_lines"],
        st.session_state["lines_per_slide"],
    )

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

        lyrics_font_size_pt = song.get("lyrics_font_size_pt")
        line_spacing = song.get("line_spacing")

        full_title = f"UMH {umh_number} {title}".strip() if umh_number else title

        for i, slide_lines in enumerate(slides):
            lyrics_text = "\n".join(slide_lines)

            if i == 0:
                slide = prs.slides.add_slide(first_layout)
                set_shape_text(
                    slide.shapes.title,
                    full_title,
                    font_size_pt=None,  # use template default
                    line_spacing=None,
                )
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

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output


def create_single_song_ppt(song_item, template_bytes: bytes):
    return create_combined_ppt([song_item], template_bytes)


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
            pix = page.get_pixmap(dpi=160)
            img_path = os.path.join(tmpdir, f"slide_{page.number + 1}.png")
            pix.save(img_path)
            with open(img_path, "rb") as f:
                images.append(f.read())

        doc.close()
        return images


def render_scrollable_images(images, height=760, active_slide=None):
    container_id = f"preview-scroll-container-{len(images)}"
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
        badge = " ← editing here" if active_slide == i else ""

        html += f"""
        <div id="slide-{i}" style="margin-bottom: 24px;">
            <div style="font-weight: 600; margin-bottom: 8px;">Slide {i}{badge}</div>
            <img
                src="data:image/png;base64,{b64}"
                style="width: 100%; border: {border}; display: block;"
            />
        </div>
        """

    html += "</div>"

    html += f"""
    <script>
    const container = document.getElementById("{container_id}");
    const activeSlide = {active_slide_js};
    const scrollKey = "preview-scroll-position";

    function saveScroll() {{
        if (container) {{
            sessionStorage.setItem(scrollKey, container.scrollTop);
        }}
    }}

    function restoreScroll() {{
        const saved = sessionStorage.getItem(scrollKey);
        if (container && saved !== null) {{
            container.scrollTop = parseInt(saved, 10);
        }}
    }}

    function scrollToActiveSlide() {{
        if (!container || activeSlide === null) return;
        const target = document.getElementById("slide-" + activeSlide);
        if (!target) return;
        container.scrollTop = target.offsetTop - 12;
        saveScroll();
    }}

    if (container) {{
        container.addEventListener("scroll", saveScroll);
    }}

    setTimeout(() => {{
        if (activeSlide !== null) {{
            scrollToActiveSlide();
        }} else {{
            restoreScroll();
        }}
    }}, 300);
    </script>
    """

    st.components.v1.html(html, height=height, scrolling=False)


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
    st.session_state["last_editor_text"] = ""
    st.session_state["last_current_song_signature"] = None
    st.session_state["editor_status_message"] = ""
    st.session_state["current_preview_slide"] = 1
    st.session_state["editor_ace_key"] += 1


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

    st.session_state["editor_override_title_font_size"] = False
    st.session_state["editor_override_lyrics_font_size"] = False
    st.session_state["editor_override_line_spacing"] = False
    st.session_state["editor_title_font_size_pt"] = 28
    st.session_state["editor_lyrics_font_size_pt"] = 32
    st.session_state["editor_line_spacing"] = 1.2

    st.session_state["current_song_preview_images"] = None
    st.session_state["last_editor_text"] = lyrics_raw
    st.session_state["last_current_song_signature"] = None
    st.session_state["editor_status_message"] = ""
    st.session_state["current_preview_slide"] = 1
    st.session_state["preview_mode"] = "song"
    st.session_state["editor_ace_key"] += 1

    # Generate preview immediately for newly loaded repository song
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
                    st.session_state["current_preview_slide"] = 1
                    st.session_state["editor_status_message"] = "Song preview loaded."
        except Exception as e:
            st.session_state["editor_status_message"] = f"Preview load failed: {e}"

def apply_pending_setlist_load():
    idx = st.session_state.get("pending_setlist_load")
    if idx is None:
        return

    if idx >= len(st.session_state["setlist"]):
        st.session_state["pending_setlist_load"] = None
        return

    item = st.session_state["setlist"][idx]
    lyrics_text = "\n\n".join("\n".join(slide) for slide in item["slides"])

    st.session_state["loaded_song"] = {
        "umh_number": item["umh_number"],
        "title": item["title"],
        "lyrics_raw": lyrics_text,
    }

    st.session_state["editing_setlist_index"] = idx
    st.session_state["editor_umh"] = item["umh_number"]
    st.session_state["editor_title"] = item["title"]
    st.session_state["editor_text"] = lyrics_text

    # Restore per-song formatting
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

    st.session_state["current_song_preview_images"] = None
    st.session_state["pending_setlist_load"] = None
    st.session_state["last_editor_text"] = lyrics_text
    st.session_state["last_current_song_signature"] = None
    st.session_state["editor_status_message"] = ""
    st.session_state["current_preview_slide"] = 1
    st.session_state["preview_mode"] = "song"
    st.session_state["editor_ace_key"] += 1

    # Generate current-song preview immediately after loading
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
                    st.session_state["current_preview_slide"] = 1
                    st.session_state["editor_status_message"] = "Song preview loaded."
        except Exception as e:
            st.session_state["editor_status_message"] = f"Preview load failed: {e}"


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


def refresh_current_song_preview(song_item, template_bytes):
    ppt_data = create_single_song_ppt(song_item, template_bytes)
    preview_images = pptx_to_preview_images(ppt_data)
    st.session_state["current_song_preview_images"] = preview_images
    st.session_state["last_current_song_signature"] = build_current_song_signature(
        song_item,
        st.session_state.get("selected_template_name"),
    )


def get_service_song_start_slides(setlist):
    starts = []
    slide_counter = 1
    for song in setlist:
        starts.append(slide_counter)
        slide_counter += len(song["slides"])
    return starts


def refresh_service_preview(setlist, template_bytes):
    ppt_data = create_combined_ppt(setlist, template_bytes)
    preview_images = pptx_to_preview_images(ppt_data)

    if not preview_images:
        raise RuntimeError("No preview images generated from service PPT")

    st.session_state["ppt_data"] = ppt_data
    st.session_state["service_preview_images"] = preview_images
    st.session_state["service_song_start_slides"] = get_service_song_start_slides(setlist)


def clear_service_outputs():
    st.session_state["ppt_data"] = None
    st.session_state["service_preview_images"] = None
    st.session_state["service_song_start_slides"] = []


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


# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    # =====================================================
    # SETLIST ORDER (always visible)
    # =====================================================
    st.markdown("### Setlist Order")

    setlist = st.session_state["setlist"]

    if not setlist:
        st.caption("No songs added yet.")
    else:
        selected_index = st.session_state.get("setlist_selected_index", 0)
        editing_index = st.session_state.get("editing_setlist_index")

        selected_index = max(0, min(selected_index, len(setlist) - 1))
        st.session_state["setlist_selected_index"] = selected_index

        for i, song in enumerate(setlist):
            is_selected = i == selected_index
            is_editing = i == editing_index

            if song["umh_number"]:
                label = f'{i+1}. UMH {song["umh_number"]} {song["title"]}'
            else:
                label = f'{i+1}. {song["title"]}'

            prefix = ""
            if is_selected:
                prefix += "🔹 "
            if is_editing:
                prefix += "✏️ "

            if prefix:
                st.markdown(f"**{prefix}{label}**")
            else:
                st.markdown(label)

    st.divider()

    # =====================================================
    # LOADED IN EDITOR (always visible)
    # =====================================================
    st.markdown("### Loaded in Editor")

    loaded_umh = st.session_state.get("editor_umh", "").strip()
    loaded_title = st.session_state.get("editor_title", "").strip()
    loaded_idx = st.session_state.get("editing_setlist_index")

    if loaded_title:
        if loaded_idx is not None:
            if loaded_umh:
                st.info(f"Editing setlist item #{loaded_idx + 1}\n\nUMH {loaded_umh} {loaded_title}")
            else:
                st.info(f"Editing setlist item #{loaded_idx + 1}\n\n{loaded_title}")
        else:
            if loaded_umh:
                st.info(f"New / repository song\n\nUMH {loaded_umh} {loaded_title}")
            else:
                st.info(f"New / repository song\n\n{loaded_title}")
    else:
        st.caption("No song loaded in editor.")

    st.divider()

    st.header("Controls")

    # -------------------------
    # TEMPLATE
    # -------------------------
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

            st.session_state["selected_template_name"] = st.selectbox(
                "Select template",
                template_names,
                index=default_index,
            )

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
                st.rerun()
        else:
            st.info("Upload at least one template.")

        if not soffice_available():
            st.warning("LibreOffice/soffice is not available.")

    # -------------------------
    # LOAD SONG
    # -------------------------
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

    # -------------------------
    # SETLIST
    # -------------------------
    with st.expander("3. Setlist", expanded=True):
        setlist = st.session_state["setlist"]
    
        if not setlist:
            st.info("No songs added yet.")
        else:
            labels = []
            for i, song in enumerate(setlist):
                if song["umh_number"]:
                    labels.append(f'{i+1}. UMH {song["umh_number"]} {song["title"]} ({len(song["slides"])})')
                else:
                    labels.append(f'{i+1}. {song["title"]} ({len(song["slides"])})')
    
            # Keep selected index valid
            st.session_state["setlist_selected_index"] = min(
                st.session_state.get("setlist_selected_index", 0),
                len(labels) - 1,
            )
    
            # Sync widget state with selected index
            st.session_state["setlist_selectbox_sidebar"] = st.session_state["setlist_selected_index"]
            previous_selected_index = st.session_state["setlist_selected_index"]
    
            selected_index = st.selectbox(
                "Selected song",
                options=list(range(len(labels))),
                format_func=lambda i: labels[i],
                key="setlist_selectbox_sidebar",
            )
            st.session_state["setlist_selected_index"] = selected_index
    
            # In service mode, changing selected song jumps to that song's first slide
            if (
                selected_index != previous_selected_index
                and st.session_state.get("preview_mode") == "service"
            ):
                starts = st.session_state.get("service_song_start_slides", [])
                st.session_state["current_preview_slide"] = (
                    starts[selected_index] if selected_index < len(starts) else 1
                )
                st.rerun()
    
            action_cols = st.columns(4)
    
            with action_cols[0]:
                if st.button("✏️", use_container_width=True, help="Edit selected song"):
                    st.session_state["pending_setlist_load"] = selected_index
                    st.session_state["preview_mode"] = "song"
                    st.session_state["current_song_preview_images"] = None
                    st.session_state["last_current_song_signature"] = None
                    st.rerun()
    
            with action_cols[1]:
                if (
                    st.button("⬆️", use_container_width=True, help="Move selected song up")
                    and selected_index > 0
                ):
                    setlist[selected_index - 1], setlist[selected_index] = (
                        setlist[selected_index],
                        setlist[selected_index - 1],
                    )
    
                    new_index = selected_index - 1
                    st.session_state["setlist_selected_index"] = new_index
                    st.session_state["setlist_selectbox_sidebar"] = new_index
    
                    editing_index = st.session_state.get("editing_setlist_index")
                    if editing_index == selected_index:
                        st.session_state["editing_setlist_index"] = new_index
                    elif editing_index == new_index:
                        st.session_state["editing_setlist_index"] = selected_index
    
                    pending = st.session_state.get("pending_setlist_load")
                    if pending == selected_index:
                        st.session_state["pending_setlist_load"] = new_index
                    elif pending == new_index:
                        st.session_state["pending_setlist_load"] = selected_index
    
                    clear_service_outputs()
                    st.rerun()
    
            with action_cols[2]:
                if (
                    st.button("⬇️", use_container_width=True, help="Move selected song down")
                    and selected_index < len(setlist) - 1
                ):
                    setlist[selected_index + 1], setlist[selected_index] = (
                        setlist[selected_index],
                        setlist[selected_index + 1],
                    )
    
                    new_index = selected_index + 1
                    st.session_state["setlist_selected_index"] = new_index
                    st.session_state["setlist_selectbox_sidebar"] = new_index
    
                    editing_index = st.session_state.get("editing_setlist_index")
                    if editing_index == selected_index:
                        st.session_state["editing_setlist_index"] = new_index
                    elif editing_index == new_index:
                        st.session_state["editing_setlist_index"] = selected_index
    
                    pending = st.session_state.get("pending_setlist_load")
                    if pending == selected_index:
                        st.session_state["pending_setlist_load"] = new_index
                    elif pending == new_index:
                        st.session_state["pending_setlist_load"] = selected_index
    
                    clear_service_outputs()
                    st.rerun()
    
            with action_cols[3]:
                if st.button("🗑️", use_container_width=True, help="Delete selected song"):
                    setlist.pop(selected_index)
    
                    if setlist:
                        new_index = min(selected_index, len(setlist) - 1)
                    else:
                        new_index = 0
    
                    st.session_state["setlist_selected_index"] = new_index
                    st.session_state["setlist_selectbox_sidebar"] = new_index
    
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
                st.session_state["editing_setlist_index"] = None
                st.session_state["pending_setlist_load"] = None
                st.session_state["setlist_selected_index"] = 0
                st.session_state["setlist_selectbox_sidebar"] = 0
                st.session_state["preview_mode"] = "song"
                st.session_state["current_song_preview_images"] = None
                clear_service_outputs()
                st.rerun()
                
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
        height=520,
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
        text_changed
        and selected_template_bytes is not None
        and selected_template_ok
        and soffice_available()
        and bool(current_slides)
        and (
            (st.session_state["refresh_on_new_line"] and trigger_refresh)
            or (not st.session_state["refresh_on_new_line"])
        )
        and new_signature != st.session_state.get("last_current_song_signature")
    )

    if should_refresh_preview:
        try:
            refresh_current_song_preview(song_item, selected_template_bytes)
            st.session_state["editor_status_message"] = "Song preview auto-refreshed."
        except Exception as e:
            st.session_state["editor_status_message"] = f"Preview refresh failed: {e}"

    st.session_state["last_editor_text"] = editor_text

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
                st.session_state["setlist"][edit_idx] = item
                clear_service_outputs()

                if (
                    selected_template_bytes is not None
                    and selected_template_ok
                    and soffice_available()
                    and current_slides
                ):
                    try:
                        refresh_current_song_preview(item, selected_template_bytes)
                        st.session_state["current_preview_slide"] = 1
                        st.session_state["preview_mode"] = "song"
                        st.session_state["editor_status_message"] = "Song updated and preview refreshed."
                    except Exception as e:
                        st.session_state["editor_status_message"] = f"Preview refresh failed: {e}"

                st.success(
                    f'Updated: {"UMH " + item["umh_number"] + " " if item["umh_number"] else ""}{item["title"]}'
                )
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
        preview_images = st.session_state.get("current_song_preview_images")

    else:
        st.caption("Full service deck preview")

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

            need_refresh = (
                refresh_service_now
                or st.session_state.get("service_preview_images") is None
                or st.session_state.get("ppt_data") is None
                or not st.session_state.get("service_song_start_slides")
            )

            if need_refresh:
                try:
                    refresh_service_preview(
                        st.session_state["setlist"],
                        selected_template_bytes,
                    )
                except Exception as e:
                    st.error(f"Service preview generation failed: {e}")

            starts = st.session_state.get("service_song_start_slides", [])
            selected_index = st.session_state.get("setlist_selected_index", 0)

            if starts and 0 <= selected_index < len(starts):
                st.session_state["current_preview_slide"] = starts[selected_index]
            else:
                st.session_state["current_preview_slide"] = 1

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
            elif not st.session_state["setlist"]:
                st.info("Add songs to the setlist to view the service preview.")

    if preview_images:
        render_scrollable_images(
            preview_images,
            height=860,
            active_slide=st.session_state.get("current_preview_slide"),
        )
    else:
        st.info("Preview will appear here.")
