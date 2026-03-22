# app_rewritten_clean.py
import re
import streamlit as st


# =========================
# HELPERS
# =========================
def simplify_heading_text(s: str) -> str:
    s = str(s or "")
    s = s.replace("\u000b", " ")
    s = re.sub(r"\([^)]*\)", " ", s)
    s = s.lower().strip()
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def is_effectively_blank(line: str) -> bool:
    if line is None:
        return True
    line = str(line).replace("\xa0", " ").replace("\u200b", "")
    return line.strip() == ""


def split_slides_manual(text: str):
    if not text:
        return []

    lines = text.splitlines()
    slides = []
    current = []

    for raw in lines:
        line = str(raw).replace("\xa0", " ").replace("\u200b", "")

        if is_effectively_blank(line):
            if current:
                slides.append(current)
                current = []
        else:
            current.append(line.strip())

    if current:
        slides.append(current)

    return slides


# =========================
# SESSION STATE
# =========================
if "setlist" not in st.session_state:
    st.session_state["setlist"] = []

if "selected_song_id" not in st.session_state:
    st.session_state["selected_song_id"] = None

if "song_selectbox_version" not in st.session_state:
    st.session_state["song_selectbox_version"] = 0


# =========================
# APP
# =========================
st.title("Service Builder (Stable Version)")

if st.button("Add Dummy Songs"):
    st.session_state["setlist"] = [
        {"song_id": "s1", "title": "Song A"},
        {"song_id": "s2", "title": "Song B"},
        {"song_id": "s3", "title": "Song C"},
    ]
    st.session_state["selected_song_id"] = "s1"
    st.session_state["song_selectbox_version"] += 1
    st.rerun()

setlist = st.session_state["setlist"]

if setlist:
    ids = []
    title_map = {}

    for i, s in enumerate(setlist):
        song_id = s.get("song_id")
        if not song_id:
            song_id = f"fallback_{s.get('title', '')}_{i}"
            s["song_id"] = song_id

        ids.append(song_id)
        title_map[song_id] = s.get("title", song_id)

    selected_song_id = st.session_state.get("selected_song_id")
    if selected_song_id not in ids:
        selected_song_id = ids[0]
        st.session_state["selected_song_id"] = selected_song_id

    selected_index_for_widget = ids.index(selected_song_id)

    widget_key = f"selected_song_id_selectbox_v{st.session_state['song_selectbox_version']}"
    selected_song_id = st.selectbox(
        "Select song",
        options=ids,
        index=selected_index_for_widget,
        format_func=lambda sid: title_map[sid],
        key=widget_key,
    )

    st.session_state["selected_song_id"] = selected_song_id

    selected_index = next(
        (i for i, s in enumerate(setlist) if s.get("song_id") == selected_song_id),
        0,
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("⬆️") and selected_index > 0:
            moved_song_id = selected_song_id
            setlist[selected_index - 1], setlist[selected_index] = (
                setlist[selected_index],
                setlist[selected_index - 1],
            )
            st.session_state["selected_song_id"] = moved_song_id
            st.session_state["song_selectbox_version"] += 1
            st.rerun()

    with col2:
        if st.button("⬇️") and selected_index < len(setlist) - 1:
            moved_song_id = selected_song_id
            setlist[selected_index + 1], setlist[selected_index] = (
                setlist[selected_index],
                setlist[selected_index + 1],
            )
            st.session_state["selected_song_id"] = moved_song_id
            st.session_state["song_selectbox_version"] += 1
            st.rerun()

    with col3:
        if st.button("🗑️"):
            if len(setlist) > 1:
                if selected_index < len(setlist) - 1:
                    next_id = setlist[selected_index + 1]["song_id"]
                else:
                    next_id = setlist[selected_index - 1]["song_id"]
            else:
                next_id = None

            setlist.pop(selected_index)
            st.session_state["selected_song_id"] = next_id
            st.session_state["song_selectbox_version"] += 1
            st.rerun()

st.write("Current setlist:", setlist)
