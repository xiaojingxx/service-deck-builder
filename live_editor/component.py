import os
import streamlit.components.v1 as components

_RELEASE = False

if _RELEASE:
    parent_dir = os.path.dirname(os.path.abspath(__file__))
    build_dir = os.path.join(parent_dir, "frontend", "dist")
    _component_func = components.declare_component("live_editor", path=build_dir)
else:
    _component_func = components.declare_component(
        "live_editor",
        url="http://localhost:3001",
    )


def live_editor(text: str = "", height: int = 500, key: str | None = None):
    return _component_func(
        text=text,
        height=height,
        key=key,
        default={
            "text": text,
            "cursorLine": 0,
            "cursorCh": 0,
            "lineCount": len(text.splitlines()) if text else 0,
            "pressedEnter": False,
        },
    )