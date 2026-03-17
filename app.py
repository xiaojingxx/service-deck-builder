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
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
    else:
        st.info("No songs added yet.")

    st.markdown("---")
    st.subheader("Live Preview of Current Song")

    if selected_template_bytes is None:
        st.info("Upload and select a template to preview the current song.")
    elif not selected_template_ok:
        st.info("Selected template is invalid, so live preview is unavailable.")
    elif not current_slides:
        st.info("Enter lyrics to preview the current song.")
    elif st.session_state["editor_preview_images"]:
        preview_images = st.session_state["editor_preview_images"]
        st.caption(f"{len(preview_images)} slide(s)")
        render_scrollable_images(preview_images, height=600)

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

    st.markdown("---")
    st.subheader("PowerPoint Preview")

    if st.session_state["preview_images"]:
        render_scrollable_images(st.session_state["preview_images"])
    else:
        st.info("Generate the service preview to see the slide images.")
