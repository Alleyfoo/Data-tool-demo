from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from src.core.processor import recipe_to_template
from src.core.state import SessionState
from src.core.streamlit_io import list_excel_sheets, read_uploaded_dataframe
from src.templates import parse_skiprows, save_template


DEFAULTS = {
    "uploaded_name": None,
    "uploaded_bytes": None,
    "header_row": 0,
    "skiprows": "",
    "delimiter": ",",
    "encoding": "utf-8",
    "sheet_name": None,
    "selected_column": None,
    "mappings": {},
    "metadata_cells": [],
    "meta_row_idx": None,
    "meta_col_idx": None,
    "meta_target": "",
    "meta_type": "metadata",
    "meta_last_pick": None,
    "draft_template_path": None,
    "last_metadata_count": 0,
}


def _render_dataframe_selection(df):
    selection = None
    try:
        selection = st.dataframe(
            df,
            use_container_width=True,
            selection_mode="single-column",
            on_select="rerun",
        )
    except TypeError:
        st.dataframe(df, use_container_width=True)
    return selection


def _extract_selected_column(selection) -> str | None:
    if selection is None:
        return None
    if hasattr(selection, "selection"):
        columns = selection.selection.get("columns", [])
        return columns[0] if columns else None
    if isinstance(selection, dict):
        columns = selection.get("columns") or selection.get("selection", {}).get("columns")
        if columns:
            return columns[0]
    return None


def _apply_metadata_headers(
    state: SessionState,
    metadata_cells: list[dict],
    file_bytes: bytes,
    filename: str,
    sheet_name: str | int | None,
    delimiter: str,
    encoding: str,
) -> None:
    cleaned: list[dict] = []
    for item in metadata_cells:
        try:
            row = int(item.get("row"))
            col = int(item.get("col"))
        except (TypeError, ValueError):
            continue
        target = str(item.get("target", "")).strip()
        if not target:
            continue
        cleaned.append({"row": row, "col": col, "target": target, "meta": item.get("metadata_type")})

    if not cleaned:
        return

    header_row = max(item["row"] for item in cleaned)
    state.header_row = int(header_row)

    raw_df = read_uploaded_dataframe(
        file_bytes,
        filename,
        header_row=None,
        skiprows=[],
        delimiter=delimiter,
        encoding=encoding,
        sheet_name=sheet_name,
        nrows=200,
    )

    cleaned.sort(key=lambda item: (item["row"], item["col"]))
    seen_cols: set[int] = set()
    selected_cols: list[int] = []
    header_names: list[str] = []
    for item in cleaned:
        col = item["col"]
        if col in seen_cols or col >= len(raw_df.columns) or col < 0:
            continue
        seen_cols.add(col)
        selected_cols.append(col)
        header_names.append(item["target"])

    if not selected_cols:
        return

    start_row = header_row + 1
    if start_row >= len(raw_df.index):
        data_df = raw_df.iloc[0:0, selected_cols]
    else:
        data_df = raw_df.iloc[start_row:, selected_cols]

    updated_df = data_df.copy()
    updated_df.columns = header_names

    metadata_signature = tuple(
        sorted(
            (item["row"], item["col"], item["target"], item.get("meta"))
            for item in cleaned
        )
    )
    st.session_state["preview_df"] = updated_df
    st.session_state["preview_settings"] = {
        "uploaded_name": filename,
        "header_row": int(state.header_row),
        "skiprows": parse_skiprows(state.skiprows),
        "delimiter": delimiter,
        "encoding": encoding,
        "sheet_name": sheet_name,
        "metadata_signature": metadata_signature,
    }


def _build_metadata_preview(
    metadata_cells: list[dict],
    file_bytes: bytes,
    filename: str,
    sheet_name: str | int | None,
    delimiter: str,
    encoding: str,
) -> pd.DataFrame:
    cleaned: list[dict] = []
    for item in metadata_cells:
        try:
            row = int(item.get("row"))
            col = int(item.get("col"))
        except (TypeError, ValueError):
            continue
        target = str(item.get("target", "")).strip()
        if not target:
            continue
        cleaned.append({"row": row, "col": col, "target": target})

    if not cleaned:
        return pd.DataFrame()

    header_row = max(item["row"] for item in cleaned)
    raw_df = read_uploaded_dataframe(
        file_bytes,
        filename,
        header_row=None,
        skiprows=[],
        delimiter=delimiter,
        encoding=encoding,
        sheet_name=sheet_name,
        nrows=200,
    )

    cleaned.sort(key=lambda item: (item["row"], item["col"]))
    seen_cols: set[int] = set()
    selected_cols: list[int] = []
    header_names: list[str] = []
    for item in cleaned:
        col = item["col"]
        if col in seen_cols or col >= len(raw_df.columns) or col < 0:
            continue
        seen_cols.add(col)
        selected_cols.append(col)
        header_names.append(item["target"])

    if not selected_cols:
        return pd.DataFrame()

    start_row = header_row + 1
    if start_row >= len(raw_df.index):
        data_df = raw_df.iloc[0:0, selected_cols]
    else:
        data_df = raw_df.iloc[start_row:, selected_cols]

    preview = data_df.copy()
    preview.columns = header_names
    return preview


def render() -> None:
    state = SessionState(DEFAULTS)

    header_left, header_right = st.columns([4, 1])
    with header_left:
        st.title("Upload & Preview")
        st.caption("Step 1 of 3: Upload a file and preview columns.")
    with header_right:
        if st.button("Reset", use_container_width=True):
            state.reset()
            st.rerun()

    uploaded = st.file_uploader(
        "Upload a CSV or Excel file",
        type=["csv", "xlsx", "xls"],
    )

    if uploaded is None and state.uploaded_bytes and state.uploaded_name:
        st.info("Using the last uploaded file from this session. Reset to clear.")
    elif uploaded is None:
        st.info("Upload a file to begin.")
        return

    if uploaded is not None and uploaded.name != state.uploaded_name:
        state.uploaded_name = uploaded.name
        state.uploaded_bytes = uploaded.getvalue()
        state.sheet_name = None
        state.selected_column = None
        state.mappings = {}
        if hasattr(st, "toast"):
            st.toast("File uploaded.")
        else:
            st.info("File uploaded.")

    if not state.uploaded_bytes or not state.uploaded_name:
        st.warning("Uploaded file is empty.")
        return

    active_name = uploaded.name if uploaded is not None else state.uploaded_name
    active_bytes = uploaded.getvalue() if uploaded is not None else state.uploaded_bytes
    is_excel = active_name.lower().endswith((".xlsx", ".xls"))

    if active_name and active_bytes and state.metadata_cells:
        current_meta_count = len(state.metadata_cells)
        if current_meta_count > int(state.last_metadata_count):
            recipe = {
                "headers": [
                    {
                        "name": str(item.get("target", "")).strip(),
                        "row": int(item.get("row", 0)),
                        "column": int(item.get("col", 0)),
                        "is_metadata": True,
                        "metadata_type": item.get("metadata_type", "metadata"),
                    }
                    for item in state.metadata_cells
                    if str(item.get("target", "")).strip()
                ],
                "header_row": int(state.header_row),
                "skiprows": parse_skiprows(state.skiprows),
                "delimiter": state.delimiter,
                "encoding": state.encoding,
                "sheet": state.sheet_name,
                "source_type": "excel" if is_excel else "csv",
                "source_file": active_name,
            }
            template = recipe_to_template(recipe)
            stem = Path(active_name).stem if active_name else "template"
            path = Path("data") / "schemas" / f"{stem}.df-template.json"
            save_template(template, path)
            state.draft_template_path = str(path)
            state.last_metadata_count = current_meta_count
            if hasattr(st, "toast"):
                st.toast("Template auto-saved to Schema Library.")
            else:
                st.success("Template auto-saved to Schema Library.")
    left, right = st.columns([3, 2], gap="large")

    with right:
        st.subheader("Settings")
        header_row = st.number_input(
            "Header row (0-indexed)",
            min_value=0,
            value=int(state.header_row),
        )
        state.header_row = int(header_row)

        skiprows_text = st.text_input(
            "Skip rows (comma-separated)",
            value=state.skiprows,
        )
        state.skiprows = skiprows_text

        if is_excel:
            sheets = list_excel_sheets(active_bytes)
            if not sheets:
                sheets = ["Sheet1"]
            if state.sheet_name not in sheets:
                state.sheet_name = sheets[0]
            state.sheet_name = st.selectbox("Sheet", sheets, index=sheets.index(state.sheet_name))
        else:
            delimiter = st.text_input("Delimiter", value=state.delimiter)
            encoding = st.text_input("Encoding", value=state.encoding)
            state.delimiter = delimiter
            state.encoding = encoding
            state.sheet_name = None

    with left:
        st.subheader("Preview")
        progress = st.progress(0)
        read_ok = False
        try:
            progress.progress(20)
            skiprows = parse_skiprows(state.skiprows)
            df = read_uploaded_dataframe(
                active_bytes,
                active_name,
                int(state.header_row),
                skiprows,
                state.delimiter,
                state.encoding,
                state.sheet_name,
                nrows=200,
            )
            progress.progress(80)
            read_ok = True
        except Exception as exc:
            progress.empty()
            st.error(f"Unable to parse file: {exc}")
            return
        finally:
            if read_ok:
                progress.progress(100)
                progress.empty()
        metadata_signature = tuple(
            sorted(
                (
                    int(item.get("row", -1)),
                    int(item.get("col", -1)),
                    str(item.get("target", "")),
                    str(item.get("metadata_type", "")),
                )
                for item in state.metadata_cells
                if item.get("target")
            )
        )
        cached_signature = st.session_state.get("preview_settings", {}).get(
            "metadata_signature"
        )
        if metadata_signature and cached_signature != metadata_signature:
            try:
                _apply_metadata_headers(
                    state,
                    state.metadata_cells,
                    active_bytes,
                    active_name,
                    state.sheet_name,
                    state.delimiter,
                    state.encoding,
                )
            except Exception:
                pass
        elif not metadata_signature:
            st.session_state["preview_df"] = df
            st.session_state["preview_settings"] = {
                "uploaded_name": active_name,
                "header_row": int(state.header_row),
                "skiprows": skiprows,
                "delimiter": state.delimiter,
                "encoding": state.encoding,
                "sheet_name": state.sheet_name,
                "metadata_signature": None,
            }

        tab_data, tab_meta = st.tabs(
            ["Table Columns (Data)", "Metadata Cells (Titles/Dates)"]
        )

        with tab_data:
            selection = _render_dataframe_selection(df.head(30))
            selected = _extract_selected_column(selection)

            if selected:
                state.selected_column = selected
            elif state.selected_column not in df.columns:
                state.selected_column = None

            if state.selected_column:
                st.success(f"Selected column: {state.selected_column}")
            else:
                fallback = st.selectbox(
                    "Select a column to map",
                    options=["(none)"] + list(df.columns),
                )
                state.selected_column = None if fallback == "(none)" else fallback

        with tab_meta:
            st.subheader("Metadata Cells")
            st.caption("Select a cell to capture titles, dates, or other metadata.")

            selected_row_idx = None
            selected_col_idx = None
            try:
                raw_df = read_uploaded_dataframe(
                    active_bytes,
                    active_name,
                    header_row=None,
                    skiprows=[],
                    delimiter=state.delimiter,
                    encoding=state.encoding,
                    sheet_name=state.sheet_name,
                    nrows=200,
                )
                display_df = raw_df.copy()
                display_df.columns = [f"Col {idx}" for idx in range(len(raw_df.columns))]
                meta_event = st.dataframe(
                    display_df.head(30),
                    use_container_width=True,
                    on_select="rerun",
                    selection_mode=["single-row", "single-column"],
                    height=200,
                )
                if meta_event.selection.rows:
                    selected_row_idx = int(meta_event.selection.rows[0])
                if meta_event.selection.columns:
                    col_name = meta_event.selection.columns[0]
                    if col_name in display_df.columns:
                        selected_col_idx = int(display_df.columns.get_loc(col_name))
            except TypeError:
                st.caption("Selection API not available; using dropdowns.")
                raw_df = read_uploaded_dataframe(
                    active_bytes,
                    active_name,
                    header_row=None,
                    skiprows=[],
                    delimiter=state.delimiter,
                    encoding=state.encoding,
                    sheet_name=state.sheet_name,
                    nrows=200,
                )
                selected_row_idx = st.selectbox(
                    "Row", options=list(range(len(raw_df))), index=0 if len(raw_df) else 0
                )
                selected_col_idx = st.selectbox(
                    "Column",
                    options=list(range(len(raw_df.columns))),
                    format_func=lambda idx: f"Col {idx}",
                    index=0 if len(raw_df.columns) else 0,
                )

            cell_value = None
            if selected_row_idx is not None and selected_col_idx is not None:
                state.meta_row_idx = selected_row_idx
                state.meta_col_idx = selected_col_idx
                col_name = f"Col {selected_col_idx}"
                cell_value = raw_df.iat[selected_row_idx, selected_col_idx]
                pick_key = f"{selected_row_idx}:{selected_col_idx}"
                if state.meta_last_pick != pick_key and cell_value is not None:
                    state.meta_target = str(cell_value)
                    state.meta_last_pick = pick_key
                st.write(
                    f"Selected cell: Row {selected_row_idx}, Column {col_name} -> `{cell_value}`"
                )
            else:
                st.info("Select a row and column to capture metadata.")

            state.meta_target = st.text_input(
                "Metadata target name", value=state.meta_target
            )
            state.meta_type = st.selectbox(
                "Metadata type",
                options=["metadata", "title", "date", "header_def"],
                index=["metadata", "title", "date", "header_def"].index(state.meta_type),
            )

            if cell_value is None and state.meta_row_idx is not None and state.meta_col_idx is not None:
                try:
                    cell_value = raw_df.iat[int(state.meta_row_idx), int(state.meta_col_idx)]
                except Exception:
                    cell_value = None

            add_disabled = (
                state.meta_row_idx is None
                or state.meta_col_idx is None
                or not state.meta_target.strip()
            )
            if st.button("Add Metadata Field", disabled=add_disabled):
                entry = {
                    "row": int(state.meta_row_idx),
                    "col": int(state.meta_col_idx),
                    "value": "" if cell_value is None else str(cell_value),
                    "target": state.meta_target.strip(),
                    "metadata_type": state.meta_type,
                }
                state.metadata_cells = state.metadata_cells + [entry]
                state.meta_target = ""
                try:
                    _apply_metadata_headers(
                        state,
                        state.metadata_cells,
                        active_bytes,
                        active_name,
                        state.sheet_name,
                        state.delimiter,
                        state.encoding,
                    )
                except Exception:
                    pass
                st.toast("Metadata field added.") if hasattr(st, "toast") else st.success(
                    "Metadata field added."
                )

            if state.metadata_cells:
                st.dataframe(state.metadata_cells, use_container_width=True)
                if st.button("Clear Metadata Fields"):
                    state.metadata_cells = []

        if state.metadata_cells:
            st.subheader("Current Recipe")
            recipe_rows = []
            for entry in state.metadata_cells:
                recipe_rows.append(
                    {
                        "target_name": entry.get("target"),
                        "source_type": "metadata",
                        "source_pointer": {"row": entry.get("row"), "col": entry.get("col")},
                        "data_type": "string",
                        "metadata_type": entry.get("metadata_type", "metadata"),
                    }
                )
            st.dataframe(recipe_rows, use_container_width=True)

            if st.button("Apply Metadata as Headers", use_container_width=True):
                try:
                    _apply_metadata_headers(
                        state,
                        state.metadata_cells,
                        active_bytes,
                        active_name,
                        state.sheet_name,
                        state.delimiter,
                        state.encoding,
                    )
                    st.success(f"Applied header row {state.header_row}.")
                except Exception as exc:
                    st.error(f"Failed to rebuild preview: {exc}")

            if st.button("Save Draft Template"):
                recipe = {
                    "headers": [
                        {
                            "name": row.get("target_name"),
                            "row": row.get("source_pointer", {}).get("row", 0),
                            "column": row.get("source_pointer", {}).get("col", 0),
                            "is_metadata": True,
                            "metadata_type": row.get("metadata_type", "metadata"),
                        }
                        for row in recipe_rows
                        if row.get("target_name")
                    ],
                    "header_row": int(state.header_row),
                    "skiprows": parse_skiprows(state.skiprows),
                    "delimiter": state.delimiter,
                    "encoding": state.encoding,
                    "sheet": state.sheet_name,
                    "source_type": "excel" if is_excel else "csv",
                    "source_file": active_name,
                }
                template = recipe_to_template(recipe)
                stem = Path(active_name).stem if active_name else "template"
                path = Path("data") / "schemas" / f"{stem}.df-template.json"
                save_template(template, path)
                state.draft_template_path = str(path)
                st.success(f"Draft template saved: {path}")
            elif state.draft_template_path:
                st.caption(f"Draft template: {state.draft_template_path}")

        st.caption(f"{len(df)} rows x {len(df.columns)} columns (preview)")
        if df.empty:
            st.warning("Preview is empty. Adjust header row or delimiter settings.")

        if active_bytes and active_name:
            with st.expander("Preview Built DataFrame"):
                st.caption("Shows the headers/data that will flow into Mapping and Query Builder.")
                try:
                    if state.metadata_cells:
                        built_df = _build_metadata_preview(
                            state.metadata_cells,
                            active_bytes,
                            active_name,
                            state.sheet_name,
                            state.delimiter,
                            state.encoding,
                        )
                    else:
                        built_df = df
                    st.dataframe(built_df.head(50), use_container_width=True)
                    st.caption(f"Showing {min(len(built_df), 50)} rows.")
                except Exception as exc:
                    st.error(f"Could not build preview: {exc}")
