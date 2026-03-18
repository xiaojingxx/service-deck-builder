import os
import base64
import tempfile
import subprocess
from io import BytesIO
from shutil import which

import fitz  # PyMuPDF
import streamlit as st
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

st.set_page_config(page_title="Live PPT Preview", layout="wide")
st.title("Live PPT Preview with LibreOffice")

SOFFICE_PATH = os.environ.get("SOFFICE_PATH", "soffice")

defaults = {
    "editor_text": "",
    "last_text": "",
    "preview_images": None,
    "ppt_bytes": None,
    "lines_per_slide": 4,
    "auto_split_by_lines": True,
    "refresh_on_new_line": True,
    "last_preview_reason": "Preview not generated yet.",
    "deck_title": "Live Service Deck",
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


def soffice_available() -> bool:
    if SOFFICE_PATH == "soffice":
        return which("soffice") is not None
    return os.path.exists(SOFFICE_PATH)


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


def get_current_slides(text: str) -> list[list[str]]:
    if st.session_state["auto_split_by_lines"]:
        return split_slides_auto(text, st.session_state["lines_per_slide"])
    return split_slides_manual(text)


def should_refresh_preview(old_text: str, new_text: str, refresh_on_new_line: bool) -> bool:
    if new_text == old_text:
        return False
    if not refresh_on_new_line:
        return True
    return new_text.count("\n") > old_text.count("\n")


def build_pptx(slides: list[list[str]], deck_title: str) -> BytesIO:
    prs = Presentation()

    try:
        title_layout = prs.slide_layouts[5]
    except IndexError:
        title_layout = prs.slide_layouts[0]

    for i, slide_lines in enumerate(slides):
        slide = prs.slides.add_slide(title_layout)

        if i == 0 and slide.shapes.title is not None:
            slide.shapes.title.text = deck_title

        left = Pt(40)
        top = Pt(70)
        width = Pt(640)
        height = Pt(360)

        textbox = slide.shapes.add_textbox(left, top, width, height)
        tf = textbox.text_frame
        tf.clear()
        tf.word_wrap = True

        for j, line in enumerate(slide_lines):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = line
            run.font.size = Pt(28)

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output


def pptx_to_preview_images(pptx_bytes: BytesIO) -> list[bytes]:
    if not soffice_available():
        raise RuntimeError("LibreOffice/soffice is not available.")

    with tempfile.TemporaryDirectory() as tmpdir:
        pptx_path = os.path.join(tmpdir, "preview.pptx")
        with open(pptx_path, "wb") as f:
            f.write(pptx_bytes.getvalue())

        cmd = [
            SOFFICE_PATH,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            tmpdir,
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
            pix = page.get_pixmap(dpi=150)
            img_path = os.path.join(tmpdir, f"slide_{page.number + 1}.png")
            pix.save(img_path)
            with open(img_path, "rb") as f:
                images.append(f.read())

        doc.close()
        return images


def render_scrollable_images(images: list[bytes], height: int = 760) -> None:
    html = f"""
    <div style="
        height: {height}px;
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
    st.components.v1.html(html, height=height + 20, scrolling=False)


controls1, controls2, controls3 = st.columns([1, 1, 1])

with controls1:
    st.checkbox("Refresh only when Enter/new line is added", key="refresh_on_new_line")

with controls2:
    st.checkbox("Auto split by lines per slide", key="auto_split_by_lines")

with controls3:
    st.slider("Lines per slide", 1, 8, key="lines_per_slide")

st.text_input("Deck title", key="deck_title")

if not soffice_available():
    st.warning(
        "LibreOffice/soffice is not available yet. "
        "If you are deploying on Streamlit Community Cloud, make sure you added packages.txt and redeployed."
    )

left, right = st.columns([1, 1.2])

with left:
    st.subheader("Editable text")

    new_text = st.text_area(
        "Edit lyrics",
        value=st.session_state["editor_text"],
        height=760,
    )

    old_text = st.session_state["last_text"]
    st.session_state["editor_text"] = new_text

    refresh_now = should_refresh_preview(
        old_text=old_text,
        new_text=new_text,
        refresh_on_new_line=st.session_state["refresh_on_new_line"],
    )

    if refresh_now:
        slides = get_current_slides(new_text)
        try:
            ppt_bytes = build_pptx(slides, st.session_state["deck_title"])
            preview_images = pptx_to_preview_images(ppt_bytes)
            st.session_state["ppt_bytes"] = ppt_bytes
            st.session_state["preview_images"] = preview_images
            if st.session_state["refresh_on_new_line"]:
                st.session_state["last_preview_reason"] = "Preview refreshed because a new line was added."
            else:
                st.session_state["last_preview_reason"] = "Preview refreshed because text changed."
        except Exception as e:
            st.session_state["last_preview_reason"] = f"Preview failed: {e}"

    st.session_state["last_text"] = new_text

    if st.button("Force refresh preview"):
        slides = get_current_slides(new_text)
        try:
            ppt_bytes = build_pptx(slides, st.session_state["deck_title"])
            preview_images = pptx_to_preview_images(ppt_bytes)
            st.session_state["ppt_bytes"] = ppt_bytes
            st.session_state["preview_images"] = preview_images
            st.session_state["last_preview_reason"] = "Preview refreshed manually."
        except Exception as e:
            st.session_state["last_preview_reason"] = f"Preview failed: {e}"

with right:
    st.subheader("LibreOffice-rendered thumbnail preview")
    st.caption(st.session_state["last_preview_reason"])

    if st.session_state["preview_images"]:
        render_scrollable_images(st.session_state["preview_images"], height=760)
    else:
        st.info("No preview yet.")

slides_for_export = get_current_slides(st.session_state["editor_text"])
if slides_for_export:
    ppt_bytes = build_pptx(slides_for_export, st.session_state["deck_title"])
    st.download_button(
        label="Download PowerPoint (.pptx)",
        data=ppt_bytes,
        file_name="live_service_deck.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
