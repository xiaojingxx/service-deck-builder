import os
import base64
import tempfile
import subprocess
import textwrap
from io import BytesIO

import fitz  # PyMuPDF
import gspread
import streamlit as st
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
    "editor_text_box": "",
    "editor_umh": "",
    "editor_title": "",
    "loaded_song": None,
    "ppt_data": None,
    "preview_images": None,
    "editing_setlist_index": None,
    "pending_setlist_load": None,
    "uploaded_templates": {},   # name -> bytes
    "selected_template_name": None,
    "reset_editor_pending": False,
    "use_auto_break": True,
    "line_width": 28,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# =========================
# HELPERS
# =========================
def split_slides(text: str) -> list[list[str]]:
    blocks = [block.strip() for block in text.split("\n\n") if block.strip()]
    slides = []
    for block in blocks:
        lines = [line.strip() for line in block.splitlines() if line.strip()]
        if lines:
            slides.append(lines)
    return slides


def auto_line_break_lyrics(text: str, width: int = 28) -> str:
    output = []

    for block in text.split("\n"):
        block = block.strip()

        if not block:
            output.append("")
            continue

        wrapped_lines = textwrap.wrap(
            block,
            width=width,
            break_long_words=False,
            break_on_hyphens=False,
        )

        if wrapped_lines:
            output.extend(wrapped_lines)
        else:
            output.append("")

    return "\n".join(output)


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


def set_shape_text(shape, text, font_size_pt=None):
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


def create_combined_ppt(
    setlist,
    template_bytes: bytes,
    title_font_size_pt=None,
    lyrics_font_size_pt=None
):
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

        full_title = f"UMH {umh_number} {title}".strip() if umh_number else title

        for i, slide_lines in enumerate(slides):
            lyrics_text = "\n".join(slide_lines)

            if i == 0:
                new_slide = prs.slides.add_slide(first_layout)
                title_placeholder = new_slide.shapes.title
                body_placeholder = get_body_placeholder(new_slide)

                set_shape_text(title_placeholder, full_title, font_size_pt=title_font_size_pt)
                set_shape_text(body_placeholder, lyrics_text, font_size_pt=lyrics_font_size_pt)
            else:
                new_slide = prs.slides.add_slide(rest_layout)
                body_placeholder = get_body_placeholder(new_slide)
                set_shape_text(body_placeholder, lyrics_text, font_size_pt=lyrics_font_size_pt)

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output


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

        return images


def render_scrollable_images(images):
    html = """
    <div style="
        height: 760px;
        overflow-y: auto;
        border: 1px solid #ddd;
        padding: 12px;
        border-radius: 8px;
        background: #fafafa;
    ">
    """
    for i, img_bytes in enumerate(images, start=1):
        b64 = base64.b64encode(img_bytes).decode("utf-8")
        html += f"""
        <div style="margin-bottom: 24px;">
            <div style="font-weight: 600; margin-bottom: 8px;">Slide {i}</div>
            <img src="data:image/png;base64,{b64}" style="width: 100%; border: 1px solid #ccc;" />
        </div>
        """
    html += "</div>"
    st.components.v1.html(html, height=780, scrolling=False)


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
    st.session_state["ppt_data"] = None
    st.session_state["preview_images"] = None


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
    st.session_state["ppt_data"] = None
    st.session_state["preview_images"] = None
    st.session_state["pending_setlist_load"] = None


def reset_editor_for_new_song():
    st.session_state["loaded_song"] = None
    st.session_state["editor_umh"] = ""
    st.session_state["editor_title"] = ""
    st.session_state["editor_text"] = ""
    st.session_state["editor_text_box"] = ""
    st.session_state["editing_setlist_index"] = None
    st.session_state["ppt_data"] = None
    st.session_state["preview_images"] = None


# Must happen before widgets are created
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
    key="template_uploader"
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


# =========================
# MAIN LAYOUT
# =========================
left_col, right_col = st.columns([1, 1.2])

title_font_size_pt = None
lyrics_font_size_pt = None

with left_col:
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

    st.markdown("---")
    st.subheader("Font Size")

    use_custom_font_size = st.checkbox(
        "Override slide master font sizes",
        value=False
    )

    if use_custom_font_size:
        st.caption("If unchecked, the deck keeps the font sizes from the slide master/layout.")

        title_font_size_pt = st.slider(
            "Title font size (pt)",
            min_value=12,
            max_value=60,
            value=28,
            step=1
        )

        lyrics_font_size_pt = st.slider(
            "Lyrics font size (pt)",
            min_value=12,
            max_value=60,
            value=32,
            step=1
        )
    else:
        st.caption("Using font sizes from the slide master/layout.")

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

    use_auto_break = st.checkbox("Apply auto line breaks", key="use_auto_break")
    line_width = st.slider(
        "Max characters per line",
        min_value=15,
        max_value=50,
        key="line_width"
    )

    raw_editor_text = st.text_area(
        "Edit lyrics for this service (use blank lines between slides)",
        height=420,
        key="editor_text_box"
    )

    processed_text = (
        auto_line_break_lyrics(raw_editor_text, width=line_width)
        if use_auto_break else raw_editor_text
    )

    if use_auto_break:
        st.caption("Auto-formatted lyrics preview")
        st.text_area(
            "Processed lyrics preview",
            value=processed_text,
            height=220,
            disabled=True
        )

    st.session_state["editor_text"] = processed_text
    current_slides = split_slides(processed_text)
    st.caption(f"{len(current_slides)} slide(s) for current song")

    allow_duplicates = st.checkbox("Allow duplicate songs in setlist", value=False)

    button_label = (
        "Update Song in Setlist"
        if edit_idx is not None
        else "Add Song to Setlist"
    )

    if st.button(button_label):
        if current_slides:
            item = {
                "umh_number": st.session_state.get("editor_umh", "").strip(),
                "title": st.session_state.get("editor_title", "").strip(),
                "slides": current_slides,
            }

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

with right_col:
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
                        title_font_size_pt=title_font_size_pt,
                        lyrics_font_size_pt=lyrics_font_size_pt
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
            st.download_button(
                label="Download Service PowerPoint",
                data=st.session_state["ppt_data"],
                file_name="service_deck.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

        st.subheader("PowerPoint Preview")
        if st.session_state["preview_images"]:
            render_scrollable_images(st.session_state["preview_images"])
        else:
            st.info("Generate the service preview to see the slide images.")
    else:
        st.info("No songs added yet.")
