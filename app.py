import html
from io import BytesIO

import streamlit as st
from streamlit_ace import st_ace
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN


st.set_page_config(page_title="Live PPT Editor", layout="wide")
st.title("Live PPT Editor")

# -----------------------------
# Session state
# -----------------------------
defaults = {
    "editor_text": "",
    "last_text": "",
    "refresh_on_enter_only": True,
    "auto_split_by_lines": True,
    "lines_per_slide": 4,
    "preview_slides": [],
    "last_preview_reason": "Preview not generated yet.",
    "deck_title": "Live Service Deck",
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# -----------------------------
# Helpers
# -----------------------------
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


def should_refresh_preview(old_text: str, new_text: str, enter_only: bool) -> bool:
    if new_text == old_text:
        return False
    if not enter_only:
        return True
    return new_text.count("\n") > old_text.count("\n")


def render_slide_preview_html(slides: list[list[str]]) -> str:
    html_parts = [
        """
        <div style="
            height: 760px;
            overflow-y: auto;
            padding: 8px;
            background: #f7f7f7;
            border: 1px solid #ddd;
            border-radius: 10px;
        ">
        """
    ]

    for i, slide in enumerate(slides, start=1):
        lines_html = "".join(
            f'<div style="margin: 0.35rem 0;">{html.escape(line)}</div>'
            for line in slide
        )
        html_parts.append(
            f"""
            <div style="margin-bottom: 24px;">
                <div style="font-weight: 600; margin: 0 0 8px 4px;">Slide {i}</div>
                <div style="
                    width: 100%;
                    min-height: 220px;
                    background: white;
                    border: 1px solid #cfcfcf;
                    border-radius: 10px;
                    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    text-align: center;
                    padding: 24px;
                    box-sizing: border-box;
                    font-size: 26px;
                    line-height: 1.4;
                    font-family: Arial, sans-serif;
                ">
                    <div style="width: 100%;">
                        {lines_html}
                    </div>
                </div>
            </div>
            """
        )

    html_parts.append("</div>")
    return "".join(html_parts)


def build_pptx(slides: list[list[str]], deck_title: str) -> BytesIO:
    prs = Presentation()

    # Use default layout 5 (title only) if available, else layout 0
    try:
        title_layout = prs.slide_layouts[5]
    except IndexError:
        title_layout = prs.slide_layouts[0]

    for i, slide_lines in enumerate(slides):
        slide = prs.slides.add_slide(title_layout)

        # Add title on first slide
        if i == 0 and slide.shapes.title is not None:
            slide.shapes.title.text = deck_title

        # Add textbox for lyrics/content
        left = top = Pt(40)
        width = Pt(640)
        height = Pt(360)

        textbox = slide.shapes.add_textbox(left, top + Pt(40), width, height)
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


# -----------------------------
# Controls
# -----------------------------
c1, c2, c3 = st.columns([1, 1, 1])

with c1:
    st.checkbox("Refresh only when Enter/new line is added", key="refresh_on_enter_only")

with c2:
    st.checkbox("Auto split by lines per slide", key="auto_split_by_lines")

with c3:
    st.slider("Lines per slide", 1, 8, key="lines_per_slide")

st.text_input("Deck title", key="deck_title")

# -----------------------------
# Editor + preview
# -----------------------------
left, right = st.columns([1, 1.2])

with left:
    st.subheader("Editable text")

    new_text = st_ace(
        value=st.session_state["editor_text"],
        language="text",
        theme="textmate",
        keybinding="vscode",
        font_size=16,
        tab_size=2,
        wrap=True,
        show_gutter=False,
        auto_update=True,
        readonly=False,
        height=760,
        key="lyrics_editor",
    )

    if new_text is None:
        new_text = st.session_state["editor_text"]

    old_text = st.session_state["last_text"]
    st.session_state["editor_text"] = new_text

    refresh_now = should_refresh_preview(
        old_text=old_text,
        new_text=new_text,
        enter_only=st.session_state["refresh_on_enter_only"],
    )

    if refresh_now:
        st.session_state["preview_slides"] = get_current_slides(new_text)
        if st.session_state["refresh_on_enter_only"]:
            st.session_state["last_preview_reason"] = "Preview refreshed because a new line was added."
        else:
            st.session_state["last_preview_reason"] = "Preview refreshed because text changed."

    elif not st.session_state["preview_slides"] and new_text.strip():
        st.session_state["preview_slides"] = get_current_slides(new_text)
        st.session_state["last_preview_reason"] = "Initial preview generated."

    st.session_state["last_text"] = new_text

    st.caption(
        "With the setting turned on, the preview refreshes only when the editor gains a new line."
    )

with right:
    st.subheader("Slide preview")
    slides = st.session_state["preview_slides"]

    st.caption(f"{len(slides)} slide(s)")
    st.caption(st.session_state["last_preview_reason"])

    preview_html = render_slide_preview_html(slides)
    st.components.v1.html(preview_html, height=780, scrolling=False)

# -----------------------------
# Download PPTX
# -----------------------------
slides_for_export = get_current_slides(st.session_state["editor_text"])

if slides_for_export:
    pptx_data = build_pptx(slides_for_export, st.session_state["deck_title"])
    st.download_button(
        label="Download PowerPoint (.pptx)",
        data=pptx_data,
        file_name="live_service_deck.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

with st.expander("Current parsed slide data"):
    st.json({"slides": slides_for_export})
