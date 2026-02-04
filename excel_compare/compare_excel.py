from __future__ import annotations

import argparse
import json
from pathlib import Path

import pandas as pd


def load_excel(path: str, sheet: str | int = 0) -> pd.DataFrame:
    """
    Load an Excel sheet and normalize values for safe comparisons.

    Notes:
    - Column names are stripped.
    - All cells are coerced to Pandas 'string' dtype and stripped.
    - Empty strings are treated as missing values (pd.NA).
    """
    df = pd.read_excel(path, sheet_name=sheet)
    df.columns = df.columns.astype(str).str.strip()

    # Coerce to string dtype (keeps pd.NA) and strip whitespace to prevent false mismatches.
    for col in df.columns:
        df[col] = df[col].astype("string").str.strip()

    df = df.replace({"": pd.NA})
    return df


def _flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Pandas .compare() produces MultiIndex columns. Excel export (index=False) doesn't
    fully support that, so we flatten to single-level column names.
    """
    if not isinstance(df.columns, pd.MultiIndex):
        return df

    df = df.copy()
    df.columns = [
        " | ".join([str(part) for part in tup if part not in (None, "", "nan")]).strip()
        for tup in df.columns.to_list()
    ]
    return df


def _validate_primary_key(df: pd.DataFrame, key: str, label: str) -> None:
    if key not in df.columns:
        raise ValueError(f"Primary key '{key}' not found in {label}. Columns: {list(df.columns)}")

    if df[key].isna().any():
        bad = df[df[key].isna()].head(5)
        raise ValueError(
            f"{label} contains blank/NA values in primary key '{key}'. "
            f"Fix the data or choose a different key. Sample rows:\n{bad}"
        )

    dupes = df[df[key].duplicated(keep=False)][key].head(10).tolist()
    if dupes:
        raise ValueError(
            f"{label} contains duplicate values in primary key '{key}'. "
            f"A primary key must be unique. Example duplicates: {dupes}"
        )


def compare_excels(df1: pd.DataFrame, df2: pd.DataFrame, key: str):
    """
    Compare two DataFrames using a primary key.
    Returns: added_rows_df, removed_rows_df, changed_cells_df, metadata_dict
    """
    cols1 = set(df1.columns)
    cols2 = set(df2.columns)

    only_in_1 = sorted(cols1 - cols2)
    only_in_2 = sorted(cols2 - cols1)
    common_cols = sorted(cols1 & cols2)
    compare_cols = [c for c in common_cols if c != key]

    df1i = df1.set_index(key, drop=False)
    df2i = df2.set_index(key, drop=False)

    added_keys = df2i.index.difference(df1i.index)
    removed_keys = df1i.index.difference(df2i.index)

    added = df2i.loc[added_keys, common_cols].reset_index(drop=True)
    removed = df1i.loc[removed_keys, common_cols].reset_index(drop=True)

    common_keys = df1i.index.intersection(df2i.index)
    left = df1i.loc[common_keys, compare_cols]
    right = df2i.loc[common_keys, compare_cols]

    changed = left.compare(
        right,
        align_axis=1,
        keep_equal=False,
        result_names=("file1", "file2"),
    ).reset_index()
    changed = _flatten_columns(changed)

    meta = {
        "only_in_1": only_in_1,
        "only_in_2": only_in_2,
        "common_cols": common_cols,
        "compare_cols": compare_cols,
        "added_count": int(len(added_keys)),
        "removed_count": int(len(removed_keys)),
        "changed_id_count": int(changed[key].nunique()) if not changed.empty else 0,
    }
    return added, removed, changed, meta


def compare_excels_by_position(df1: pd.DataFrame, df2: pd.DataFrame):
    """
    Compare two DataFrames by row index (row N in file1 vs row N in file2).
    No primary key: "added" = rows at the end of file2 only, "removed" = at the end of file1 only,
    "modified" = same index, different values.
    Returns: added_df, removed_df, modified_list, meta
    """
    cols1 = set(df1.columns)
    cols2 = set(df2.columns)
    only_in_1 = sorted(cols1 - cols2)
    only_in_2 = sorted(cols2 - cols1)
    common_cols = sorted(cols1 & cols2)
    n1, n2 = len(df1), len(df2)

    # Added: rows in file2 at indices n1..n2-1
    if n2 > n1 and common_cols:
        added = df2.iloc[n1:n2][common_cols].copy()
        added.insert(0, "row_index", range(n1, n2))
    else:
        added = pd.DataFrame(columns=(["row_index"] + common_cols) if common_cols else ["row_index"])

    # Removed: rows in file1 at indices n2..n1-1
    if n1 > n2 and common_cols:
        removed = df1.iloc[n2:n1][common_cols].copy()
        removed.insert(0, "row_index", range(n2, n1))
    else:
        removed = pd.DataFrame(columns=(["row_index"] + common_cols) if common_cols else ["row_index"])

    # Modified: same index, different values
    modified_list = []
    for i in range(min(n1, n2)):
        row1 = df1.iloc[i]
        row2 = df2.iloc[i]
        changes = []
        for c in common_cols:
            v1, v2 = row1[c], row2[c]
            same = (pd.isna(v1) and pd.isna(v2)) or (not pd.isna(v1) and not pd.isna(v2) and str(v1).strip() == str(v2).strip())
            if not same:
                old_val = None if pd.isna(v1) else str(v1).strip()
                new_val = None if pd.isna(v2) else str(v2).strip()
                changes.append({"column": c, "old_value": old_val, "new_value": new_val})
        if changes:
            modified_list.append({"row_index": i, "changes": changes})

    meta = {
        "only_in_1": only_in_1,
        "only_in_2": only_in_2,
        "common_cols": common_cols,
        "added_count": len(added),
        "removed_count": len(removed),
        "changed_id_count": len(modified_list),
    }
    return added, removed, modified_list, meta


def _changed_df_to_modified_rows(changed: pd.DataFrame, key: str) -> list[dict]:
    """
    Convert the compare() result DataFrame into a frontend-friendly list of
    modified rows: each item has id (key value) and a list of { column, old_value, new_value }.
    """
    if changed.empty:
        return []

    modified_rows = []
    # Flattened columns are like "ColName | file1" and "ColName | file2"
    key_col = key
    other_cols = [c for c in changed.columns if c != key_col]
    # Group by "base" column name (part before " | ")
    col_pairs: dict[str, tuple[str, str]] = {}
    for c in other_cols:
        if " | file1" in c:
            base = c.replace(" | file1", "").strip()
            file2_col = base + " | file2"
            if file2_col in changed.columns:
                col_pairs[base] = (c, file2_col)
        elif " | file2" in c and c not in [v[1] for v in col_pairs.values()]:
            base = c.replace(" | file2", "").strip()
            file1_col = base + " | file1"
            if file1_col in changed.columns:
                col_pairs[base] = (file1_col, c)

    for _, row in changed.iterrows():
        row_id = row[key_col]
        if pd.isna(row_id):
            row_id = str(row_id)
        else:
            row_id = str(row_id).strip()
        changes = []
        for base, (c1, c2) in col_pairs.items():
            old_val = row.get(c1)
            new_val = row.get(c2)
            if pd.isna(old_val):
                old_val = None
            else:
                old_val = str(old_val).strip()
            if pd.isna(new_val):
                new_val = None
            else:
                new_val = str(new_val).strip()
            if old_val != new_val:
                changes.append({"column": base, "old_value": old_val, "new_value": new_val})
        if changes:
            modified_rows.append({"id": row_id, "changes": changes})
    return modified_rows


def get_comparison_for_frontend(
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    key: str,
) -> dict:
    """
    Run comparison and return a dict suitable for JSON/API and frontend display:
    - summary: counts and config
    - added_rows: list of dicts (column -> value)
    - removed_rows: list of dicts
    - modified_rows: list of { id, changes: [ { column, old_value, new_value } ] }
    """
    added, removed, changed, meta = compare_excels(df1, df2, key)

    def df_to_list_of_dicts(df: pd.DataFrame) -> list[dict]:
        if df.empty:
            return []
        return df.fillna("").astype(str).to_dict(orient="records")

    modified_rows = _changed_df_to_modified_rows(changed, key)

    return {
        "summary": {
            "added_count": meta["added_count"],
            "removed_count": meta["removed_count"],
            "modified_count": meta["changed_id_count"],
            "columns_only_in_file1": meta["only_in_1"],
            "columns_only_in_file2": meta["only_in_2"],
        },
        "added_rows": df_to_list_of_dicts(added),
        "removed_rows": df_to_list_of_dicts(removed),
        "modified_rows": modified_rows,
    }


def get_comparison_for_frontend_by_position(df1: pd.DataFrame, df2: pd.DataFrame) -> dict:
    """
    Compare by row position (no primary key). Returns same shape as get_comparison_for_frontend.
    modified_rows use "id": "Row N" and "row_index": N.
    """
    added, removed, modified_list, meta = compare_excels_by_position(df1, df2)

    def df_to_list_of_dicts(df: pd.DataFrame) -> list[dict]:
        if df.empty:
            return []
        return df.fillna("").astype(str).to_dict(orient="records")

    modified_rows = [
        {"id": f"Row {m['row_index']}", "row_index": m["row_index"], "changes": m["changes"]}
        for m in modified_list
    ]

    return {
        "summary": {
            "added_count": meta["added_count"],
            "removed_count": meta["removed_count"],
            "modified_count": meta["changed_id_count"],
            "columns_only_in_file1": meta["only_in_1"],
            "columns_only_in_file2": meta["only_in_2"],
        },
        "added_rows": df_to_list_of_dicts(added),
        "removed_rows": df_to_list_of_dicts(removed),
        "modified_rows": modified_rows,
        "mode": "position",
    }


def export_to_excel(
    added: pd.DataFrame,
    removed: pd.DataFrame,
    changed: pd.DataFrame,
    meta: dict,
    output_path: str,
    file1_path: str,
    file2_path: str,
    sheet: str | int,
    key: str,
) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    summary = pd.DataFrame(
        [
            {"metric": "file1", "value": file1_path},
            {"metric": "file2", "value": file2_path},
            {"metric": "sheet", "value": str(sheet)},
            {"metric": "primary_key", "value": key},
            {"metric": "added_rows", "value": meta["added_count"]},
            {"metric": "removed_rows", "value": meta["removed_count"]},
            {"metric": "modified_ids", "value": meta["changed_id_count"]},
            {"metric": "columns_only_in_file1", "value": ", ".join(meta["only_in_1"])},
            {"metric": "columns_only_in_file2", "value": ", ".join(meta["only_in_2"])},
        ]
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        added.to_excel(writer, sheet_name="Added_Rows", index=False)
        removed.to_excel(writer, sheet_name="Removed_Rows", index=False)
        changed.to_excel(writer, sheet_name="Changed_Cells", index=False)

    return output_path


def export_to_excel_by_position(
    added: pd.DataFrame,
    removed: pd.DataFrame,
    modified_list: list[dict],
    meta: dict,
    output_path: str,
    file1_path: str,
    file2_path: str,
    sheet: str | int,
) -> Path:
    """Export position-based comparison to Excel (no primary key)."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    summary = pd.DataFrame(
        [
            {"metric": "file1", "value": file1_path},
            {"metric": "file2", "value": file2_path},
            {"metric": "sheet", "value": str(sheet)},
            {"metric": "mode", "value": "position (no primary key)"},
            {"metric": "added_rows", "value": meta["added_count"]},
            {"metric": "removed_rows", "value": meta["removed_count"]},
            {"metric": "modified_rows", "value": meta["changed_id_count"]},
            {"metric": "columns_only_in_file1", "value": ", ".join(meta["only_in_1"])},
            {"metric": "columns_only_in_file2", "value": ", ".join(meta["only_in_2"])},
        ]
    )

    # Flatten modified_list to one row per change
    changed_rows = []
    for m in modified_list:
        for ch in m["changes"]:
            changed_rows.append({
                "row_index": m["row_index"],
                "column": ch["column"],
                "file1_value": ch["old_value"],
                "file2_value": ch["new_value"],
            })
    changed_df = pd.DataFrame(changed_rows) if changed_rows else pd.DataFrame(columns=["row_index", "column", "file1_value", "file2_value"])

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        added.to_excel(writer, sheet_name="Added_Rows", index=False)
        removed.to_excel(writer, sheet_name="Removed_Rows", index=False)
        changed_df.to_excel(writer, sheet_name="Changed_Cells", index=False)

    return output_path


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Compare two Excel files using a primary key.")
    p.add_argument("--file1", default="input/file1.xlsx", help="Path to first Excel file")
    p.add_argument("--file2", default="input/file2.xlsx", help="Path to second Excel file")
    p.add_argument("--output", default="output/comparison_result.xlsx", help="Output Excel path")
    p.add_argument("--key", required=False, default="EMP_ID", help="Primary key column name (omit with --no-key for position-based compare)")
    p.add_argument("--no-key", action="store_true", help="Compare by row position (no primary key)")
    p.add_argument(
        "--sheet",
        default=0,
        help="Sheet name or 0-based index (default: 0)",
    )
    p.add_argument("--json", action="store_true", help="Also write frontend-ready JSON to output dir")
    return p.parse_args()


def main() -> None:
    args = parse_args()

    file1 = str(args.file1)
    file2 = str(args.file2)
    output = str(args.output)
    key = "" if getattr(args, "no_key", False) else str(args.key).strip()
    sheet = args.sheet

    # Allow numeric strings for --sheet (e.g., "--sheet 0")
    if isinstance(sheet, str) and sheet.isdigit():
        sheet = int(sheet)

    print("Loading Excel files...")
    df1 = load_excel(file1, sheet=sheet)
    df2 = load_excel(file2, sheet=sheet)

    if key:
        _validate_primary_key(df1, key, "file1")
        _validate_primary_key(df2, key, "file2")
        print("Comparing (by primary key)...")
        added, removed, changed, meta = compare_excels(df1, df2, key)
        if getattr(args, "json", False):
            frontend_data = get_comparison_for_frontend(df1, df2, key)
            json_path = Path(output).parent / "comparison_result.json"
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(frontend_data, f, indent=2, ensure_ascii=False)
            print(f"JSON for frontend: {json_path}")
        print("Summary")
        print(f"  Added rows   : {meta['added_count']}")
        print(f"  Removed rows : {meta['removed_count']}")
        print(f"  Modified IDs : {meta['changed_id_count']}")
        print("Exporting result...")
        out = export_to_excel(
            added=added,
            removed=removed,
            changed=changed,
            meta=meta,
            output_path=output,
            file1_path=file1,
            file2_path=file2,
            sheet=sheet,
            key=key,
        )
    else:
        print("Comparing (by row position, no primary key)...")
        added, removed, modified_list, meta = compare_excels_by_position(df1, df2)
        if getattr(args, "json", False):
            frontend_data = get_comparison_for_frontend_by_position(df1, df2)
            json_path = Path(output).parent / "comparison_result.json"
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(frontend_data, f, indent=2, ensure_ascii=False)
            print(f"JSON for frontend: {json_path}")
        print("Summary")
        print(f"  Added rows   : {meta['added_count']}")
        print(f"  Removed rows : {meta['removed_count']}")
        print(f"  Modified rows: {meta['changed_id_count']}")
        print("Exporting result...")
        out = export_to_excel_by_position(
            added=added,
            removed=removed,
            modified_list=modified_list,
            meta=meta,
            output_path=output,
            file1_path=file1,
            file2_path=file2,
            sheet=sheet,
        )
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
