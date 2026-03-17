import os
import base64
import tempfile
import subprocess
from io import BytesIO

import fitz  # PyMuPDF
import gspread
import streamlit as st
from google.oauth2.service_account import Credentials
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt


# =========================
# APP CONFIG
# =========================
st.set_page_config(page_title="Service Deck Builder", layout="wide")
st.title("Service Deck Builder")

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
# SESSION STATE
# =========================
defaults = {
    "setlist": [],
    "editor_text": "",
    "editor_text_box": "",
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
    "editing_setlist_index": None,
    "pending_setlist_load": None,
    "uploaded_templates": {},
    "selected_template_name": None,
    "reset_editor_pending": False,
    "auto_split_by_lines": True,
    "lines_per_slide": 4,
    "auto_refresh_editor_preview": True,
    "editor_preview_ppt_data": None,
    "editor_preview_images": None,
    "focused_preview_slide": 1,
}
for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# =========================
# DATA HELPERS
# =========================
@st.cache_data(ttl=300, show_spinner=False)
def get_sheet_records():
    return sheet.get_all_records()


def find_row_by_umh(umh_number: str):
    target = umh_number.strip()
    for row in get_sheet_records():
        if str(row.get("UMH Number", "")).strip() == target:
            return row
    return None


def search_titles(keyword: str):
    keyword = keyword.lower().strip()
    matches = []
    for row in get_sheet_records():
        title = str(row.get("Title", "")).strip()
        if keyword and keyword in title.lower():
            matches.append(row)
    return matches[:20]


# =========================
# SLIDE SPLITTING
# =========================
def split_slides_manual(text: str) -> list[list[str]]:
    blocks = [block.strip() for block in text.split("\n\n") if block.strip()]
    slides = []
    for block in blocks:
        lines = [line.strip() for line in block.splitlines() if line.strip()]
        if lines:
            slides.append(lines)
    return slides



def split_slides_auto(text: str, lines_per_slide: int = 4) -> list[list[str]]:
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


# =========================
# POWERPOINT HELPERS
# =========================
def open_presentation_from_bytes(template_bytes: bytes):
    return Presentation(BytesIO(template_bytes))



def get_layout_by_name(prs: Presentation, layout_name: str):
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

    lines = text.split("\n") if text else [""]

    for i, line in enumerate(lines):
        paragraph = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        paragraph.alignment = PP_ALIGN.CENTER

        if line_spacing is not None:
            paragraph.line_spacing = line_spacing

        run = paragraph.add_run()
        run.text = line

        if font_size_pt is not None:
            run.font.size = Pt(font_size_pt)



def delete_all_slides(prs: Presentation):
    while len(prs.slides) > 0:
        slide_id = prs.slides._sldIdLst[0]
        rel_id = slide_id.rId
        prs.part.drop_rel(rel_id)
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

    if first_slide.shapes.title is None:
        errors.append(f"{FIRST_LAYOUT_NAME} is missing a title placeholder")
    if get_body_placeholder(first_slide) is None:
        errors.append(f"{FIRST_LAYOUT_NAME} is missing a body/lyrics placeholder")
    if get_body_placeholder(rest_slide) is None:
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

        for idx, slide_lines in enumerate(slides):
            lyrics_text = "\n".join(slide_lines)

            if idx == 0:
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
        raw_bytes = pptx_bytes.getvalue() if hasattr(pptx_bytes, "getvalue") else pptx_bytes

        with open(pptx_path, "wb") as f:
            f.write(raw_bytes)

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


@st.cache_data(show_spinner=False)
def build_editor_preview_cached(song_item, template_bytes: bytes):
    ppt_data = create_single_song_ppt(song_item, template_bytes)
    preview_images = pptx_to_preview_images(ppt_data)
    return ppt_data.getvalue(), preview_images



def render_scrollable_images(images, height=780, focus_index=None):
    html = f"""
    <div id="preview-container" style="
        height: {height}px;
        overflow-y: auto;
        border: 1px solid #ddd;
        padding: 12px;
        border-radius: 8px;
        background: #fafafa;
        scroll-behavior: smooth;
    ">
    """

    for i, img_bytes in enumerate(images, start=1):
        b64 = base64.b64encode(img_bytes).decode("utf-8")
        border = "2px solid #4a90e2" if focus_index == i else "1px solid #ccc"
        html += f"""
        <div id="slide-{i}" style="margin-bottom: 24px;">
            <div style="font-weight: 600; margin-bottom: 8px;">Slide {i}</div>
            <img src="data:image/png;base64,{b64}" style="width: 100%; border: {border};" />
        </div>
        """

    html += "</div>"

    if focus_index is not None:
        html += f"""
        <script>
            const container = document.getElementById("preview-container");
            const target = document.getElementById("slide-{focus_index}");
            if (container && target) {{
                container.scrollTop = target.offsetTop - 10;
            }}
        </script>
        """

    st.components.v1.html(html, height=height + 20, scrolling=False)


# =========================
# EDITOR STATE HELPERS
# =========================
def get_current_slides_from_raw_text(raw_text: str) -> list[list[str]]:
    if st.session_state["auto_split_by_lines"]:
        return split_slides_auto(raw_text, st.session_state["lines_per_slide"])
    return split_slides_manual(raw_text)



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
    st.session_state["editor_text_box"] = new_text
    st.session_state["editing_setlist_index"] = None

    st.session_state["editor_override_title_font_size"] = False
    st.session_state["editor_override_lyrics_font_size"] = False
    st.session_state["editor_override_line_spacing"] = False
    st.session_state["editor_title_font_size_pt"] = 28
    st.session_state["editor_lyrics_font_size_pt"] = 32
    st.session_state["editor_line_spacing"] = 1.2

    st.session_state["ppt_data"] = None
    st.session_state["preview_images"] = None
    st.session_state["editor_preview_ppt_data"] = None
    st.session_state["editor_preview_images"] = None
    st.session_state["focused_preview_slide"] = 1



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
    st.session_state["editor_text_box"] = lyrics_text
    st.session_state["editing_setlist_index"] = pending

    st.session_state["editor_override_title_font_size"] = item.get("override_title_font_size", False)
    st.session_state["editor_override_lyrics_font_size"] = item.get("override_lyrics_font_size", False)
    st.session_state["editor_override_line_spacing"] = item.get("override_line_spacing", False)

    st.session_state["editor_title_font_size_pt"] = item.get("title_font_size_pt", 28) or 28
    st.session_state["editor_lyrics_font_size_pt"] = item.get("lyrics_font_size_pt", 32) or 32
    st.session_state["editor_line_spacing"] = item.get("line_spacing", 1.2) or 1.2

    st.session_state["ppt_data"] = None
    st.session_state["preview_images"] = None
    st.session_state["editor_preview_ppt_data"] = None
    st.session_state["editor_preview_images"] = None
    st.session_state["pending_setlist_load"] = None
    st.session_state["focused_preview_slide"] = 1



def reset_editor_for_new_song():
    st.session_state["loaded_song"] = None
    st.session_state["editor_umh"] = ""
    st.session_state["editor_title"] = ""
    st.session_state["editor_text"] = ""
    st.session_state["editor_text_box"] = ""
    st.session_state["editing_setlist_index"] = None

    st.session_state["editor_override_title_font_size"] = False
    st.session_state["editor_override_lyrics_font_size"] = False
    st.session_state["editor_override_line_spacing"] = False
    st.session_state["editor_title_font_size_pt"] = 28
    st.session_state["editor_lyrics_font_size_pt"] = 32
    st.session_state["editor_line_spacing"] = 1.2

    st.session_state["ppt_data"] = None
    st.session_state["preview_images"] = None
    st.session_state["editor_preview_ppt_data"] = None
    st.session_state["editor_preview_images"] = None
    st.session_state["focused_preview_slide"] = 1


apply_pending_setlist_load()

if st.session_state.get("reset_editor_pending"):
    reset_editor_for_new_song()
    st.session_state["reset_editor_pending"] = False


# =========================
# TEMPLATE SECTION
# =========================
st.subheader("Template")

uploaded_templates = st.file_uploader(
    "Upload one or more Template.pptx files",
    type=["pptx"],
    accept_multiple_files=True,
    key="template_uploader",
)

if uploaded_templates:
    for file in uploaded_templates:
        st.session_state["uploaded_templates"][file.name] = file.getvalue()

template_names = list(st.session_state["uploaded_templates"].keys())
selected_template_bytes = None
selected_template_ok = False
selected_template_errors = []
selected_template_warnings = []

if template_names:
    default_index = 0
    if st.session_state["selected_template_name"] in template_names:
        default_index = template_names.index(st.session_state["selected_template_name"])

    selected_template_name = st.selectbox(
        "Select template",
        template_names,
        index=default_index,
    )
    st.session_state["selected_template_name"] = selected_template_name
    selected_template_bytes = st.session_state["uploaded_templates"][selected_template_name]

    selected_template_ok, selected_template_errors, selected_template_warnings = validate_template_bytes(
        selected_template_bytes
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


# =========================
# MAIN LAYOUT
# =========================
main_left, main_right = st.columns([1, 1.2])

with main_left:
    st.subheader("Current Setlist")

    if st.session_state["setlist"]:
        for i, song in enumerate(st.session_state["setlist"]):
            label = f'UMH {song["umh_number"]} {song["title"]}' if song["umh_number"] else song["title"]
            st.markdown(f"**{i + 1}. {label} ({len(song['slides'])})**")
    else:
        st.info("No songs added yet.")

    st.markdown("---")
    st.subheader("Load Song")

    if st.button("Start New Song"):
        st.session_state["reset_editor_pending"] = True
        st.rerun()

    load_mode = st.radio("Find hymn by", ["UMH Number", "Title"], horizontal=True)

    if load_mode == "UMH Number":
        umh_number_input = st.text_input("Enter UMH Number", placeholder="e.g. 57")
        if st.button("Load Hymn by Number"):
            if umh_number_input.strip():
                match = find_row_by_umh(umh_number_input)
                if match:
                    load_song_into_editor_from_repository(match)
                    st.success("Hymn loaded.")
                    st.rerun()
                else:
                    st.error("Hymn not found.")
    else:
        keyword = st.text_input("Search title", placeholder="e.g. thousand tongues")
        if keyword.strip():
            matches = search_titles(keyword)
            if matches:
                options = [
                    f'UMH {row.get("UMH Number", "")} - {row.get("Title", "")}'
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

    st.markdown("---")
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
    st.checkbox("Auto split by lines per slide", key="auto_split_by_lines")
    st.slider("Lines per slide", min_value=1, max_value=8, key="lines_per_slide")

    current_slides = get_current_slides_from_raw_text(st.session_state.get("editor_text_box", ""))
    slide_count = len(current_slides)

    if slide_count > 0:
        st.selectbox(
            "Focus preview on slide",
            options=list(range(1, slide_count + 1)),
            key="focused_preview_slide",
        )
    else:
        st.session_state["focused_preview_slide"] = 1

    editor_col, preview_col = st.columns([1, 1.25])

    with editor_col:
        st.markdown("#### Raw Lyrics")
        editor_text = st.text_area(
            "Edit lyrics for this service",
            height=720,
            key="editor_text_box",
            label_visibility="collapsed",
        )
        st.session_state["editor_text"] = editor_text

        current_slides = get_current_slides_from_raw_text(editor_text)
        slide_count = len(current_slides)

        if st.session_state["auto_split_by_lines"]:
            st.caption(
                f"{slide_count} slide(s) for current song "
                f"({st.session_state['lines_per_slide']} lines per slide, blank lines kept as verse separators)"
            )
        else:
            st.caption(
                f"{slide_count} slide(s) for current song "
                f"(manual mode: blank lines separate slides)"
            )

    with preview_col:
        st.markdown("#### Live Preview of Current Song")

        if selected_template_bytes is None:
            st.info("Upload and select a template to preview the current song.")
        elif not selected_template_ok:
            st.info("Selected template is invalid, so live preview is unavailable.")
        elif not current_slides:
            st.info("Enter lyrics to preview the current song.")
        elif st.session_state["editor_preview_images"]:
            preview_images = st.session_state["editor_preview_images"]
            focus_index = min(max(st.session_state.get("focused_preview_slide", 1), 1), len(preview_images))
            st.caption(f"{len(preview_images)} slide(s) — focused on slide {focus_index}")
            render_scrollable_images(preview_images, height=760, focus_index=focus_index)

            if st.session_state["editor_preview_ppt_data"] is not None:
                st.download_button(
                    label="Download Current Song Preview",
                    data=st.session_state["editor_preview_ppt_data"],
                    file_name="current_song_preview.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key="download_current_song_preview",
                )
        else:
            st.info("Click refresh to build the current-song preview.")

    st.markdown("#### Song Formatting")

    st.checkbox("Override title font size for this song", key="editor_override_title_font_size")
    if st.session_state["editor_override_title_font_size"]:
        st.slider("Title font size (pt)", min_value=12, max_value=60, key="editor_title_font_size_pt")
    else:
        st.caption("Title font size: using template default")

    st.checkbox("Override lyrics font size for this song", key="editor_override_lyrics_font_size")
    if st.session_state["editor_override_lyrics_font_size"]:
        st.slider("Lyrics font size (pt)", min_value=12, max_value=60, key="editor_lyrics_font_size_pt")
    else:
        st.caption("Lyrics font size: using template default")

    st.checkbox("Override line spacing for this song", key="editor_override_line_spacing")
    if st.session_state["editor_override_line_spacing"]:
        st.slider("Line spacing", min_value=0.8, max_value=2.0, step=0.1, key="editor_line_spacing")
    else:
        st.caption("Line spacing: using template default")

    current_song_item = {
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

    st.checkbox("Auto-refresh live preview", key="auto_refresh_editor_preview")
    manual_preview = st.button("Refresh Live Preview")

    should_build_live_preview = (
        selected_template_bytes is not None
        and selected_template_ok
        and bool(current_slides)
        and (st.session_state["auto_refresh_editor_preview"] or manual_preview)
    )

    if should_build_live_preview:
        try:
            preview_ppt_bytes, preview_images = build_editor_preview_cached(
                current_song_item,
                selected_template_bytes,
            )
            st.session_state["editor_preview_ppt_data"] = preview_ppt_bytes
            st.session_state["editor_preview_images"] = preview_images
        except Exception as e:
            st.error(f"Live preview failed: {e}")

    st.markdown("---")
    allow_duplicates = st.checkbox("Allow duplicate songs in setlist", value=False)
    button_label = "Update Song in Setlist" if edit_idx is not None else "Add Song to Setlist"

    if st.button(button_label):
        if current_slides:
            item = current_song_item
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
                    st.session_state["ppt_data"] = None
                    st.session_state["preview_images"] = None
                    st.success(
                        f'Added: {"UMH " + item["umh_number"] + " " if item["umh_number"] else ""}{item["title"]}'
                    )
                    st.session_state["reset_editor_pending"] = True
                    st.rerun()
            else:
                st.session_state["setlist"][edit_idx] = item
                st.session_state["editing_setlist_index"] = None
                st.session_state["ppt_data"] = None
                st.session_state["preview_images"] = None
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

with main_right:
    st.subheader("Current Setlist")

    if st.session_state["setlist"]:
        remove_index = None

        for i, song in enumerate(st.session_state["setlist"]):
            label = f'UMH {song["umh_number"]} {song["title"]}' if song["umh_number"] else song["title"]
            total_slides = len(song["slides"])

            col_title, col_edit, col_up, col_down, col_delete = st.columns(
                [12, 1, 1, 1, 1],
                gap="small",
            )

            with col_title:
                st.markdown(f"**{i + 1}. {label} ({total_slides})**")
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
                    preview_images = pptx_to_preview_images(ppt_data)
                    st.session_state["ppt_data"] = ppt_data
                    st.session_state["preview_images"] = preview_images
                    st.success("Service preview generated.")
                except Exception as e:
                    st.error(f"Preview generation failed: {e}")

        if col2.button("Clear Setlist"):
            st.session_state["setlist"] = []
            st.session_state["ppt_data"] = None
            st.session_state["preview_images"] = None
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
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
    else:
        st.info("No songs added yet.")
