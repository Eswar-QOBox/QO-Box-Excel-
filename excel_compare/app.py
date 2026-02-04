"""
Flask app: compare two Excel files and show added / removed / modified rows in a web UI.
"""
import tempfile
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file

from compare_excel import (
    compare_excels,
    compare_excels_by_position,
    export_to_excel,
    export_to_excel_by_position,
    get_comparison_for_frontend,
    get_comparison_for_frontend_by_position,
    load_excel,
    _validate_primary_key,
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32 MB

BASE = Path(__file__).resolve().parent
INPUT_DIR = BASE / "input"
DEFAULT_FILE1 = INPUT_DIR / "file1.xlsx"
DEFAULT_FILE2 = INPUT_DIR / "file2.xlsx"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/file-info", methods=["POST"])
def api_file_info():
    """Return sheet count and names for an uploaded Excel file (for UI display)."""
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "No file"}), 400
    try:
        xl = pd.ExcelFile(f)
        return jsonify({
            "sheets": len(xl.sheet_names),
            "sheet_names": xl.sheet_names,
        })
    except Exception as e:
        return jsonify({"error": str(e), "sheets": None, "sheet_names": []}), 400


PREVIEW_MAX_ROWS = 20


@app.route("/api/compare-sheet-count", methods=["POST"])
def api_compare_sheet_count():
    """Compare only the number of sheets in two uploaded files."""
    f_expected = request.files.get("expected_file") or request.files.get("file1")
    f_actual = request.files.get("actual_file") or request.files.get("file2")
    if not f_expected or not f_actual or not f_expected.filename or not f_actual.filename:
        return jsonify({"error": "Upload both Expected and Actual files"}), 400
    try:
        xl1 = pd.ExcelFile(f_expected)
        xl2 = pd.ExcelFile(f_actual)
        n1, n2 = len(xl1.sheet_names), len(xl2.sheet_names)
        return jsonify({
            "expected_sheets": n1,
            "actual_sheets": n2,
            "match": n1 == n2,
            "expected_filename": f_expected.filename,
            "actual_filename": f_actual.filename,
            "sheet_names_expected": xl1.sheet_names,
            "sheet_names_actual": xl2.sheet_names,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.route("/api/preview", methods=["POST"])
def api_preview():
    """Return first N rows of an uploaded Excel file for preview. Accept form key 'sheet' (index, default 0)."""
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "No file"}), 400
    sheet_raw = request.form.get("sheet", "0")
    sheet = int(sheet_raw) if isinstance(sheet_raw, str) and sheet_raw.isdigit() else 0
    try:
        df = load_excel(f, sheet=sheet)
        columns = list(df.columns)
        # First PREVIEW_MAX_ROWS rows, values as strings for JSON
        head = df.head(PREVIEW_MAX_ROWS)
        rows = []
        for _, r in head.iterrows():
            rows.append({c: (str(v) if pd.notna(v) else "") for c, v in zip(columns, r)})
        return jsonify({"columns": columns, "rows": rows, "total_rows": len(df)})
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.route("/api/compare", methods=["GET", "POST"])
def api_compare():
    """
    Run comparison and return JSON for frontend.
    GET: use default files (input/file1.xlsx, input/file2.xlsx).
    POST: optional form keys file1, file2 (paths), key, sheet; or upload expected_file + actual_file.
    """
    key_raw = (request.args.get("key") or request.form.get("key") or "EMP_ID").strip()
    # Empty or "none" => compare by row position (no primary key)
    key = None if key_raw.lower() in ("", "none") else key_raw
    sheet1_raw = request.args.get("expected_sheet") or request.form.get("expected_sheet") or request.args.get("sheet") or request.form.get("sheet") or "0"
    sheet2_raw = request.args.get("actual_sheet") or request.form.get("actual_sheet") or request.args.get("sheet") or request.form.get("sheet") or "0"
    sheet1 = int(sheet1_raw) if isinstance(sheet1_raw, str) and sheet1_raw.isdigit() else 0
    sheet2 = int(sheet2_raw) if isinstance(sheet2_raw, str) and sheet2_raw.isdigit() else 0

    file1_path = request.args.get("file1") or request.form.get("file1") or str(DEFAULT_FILE1)
    file2_path = request.args.get("file2") or request.form.get("file2") or str(DEFAULT_FILE2)

    f_expected = request.files.get("expected_file") or request.files.get("file1")
    f_actual = request.files.get("actual_file") or request.files.get("file2")

    try:
        if f_expected and f_actual and f_expected.filename and f_actual.filename:
            df1 = load_excel(f_expected, sheet=sheet1)
            df2 = load_excel(f_actual, sheet=sheet2)
            file1_label, file2_label = f_expected.filename, f_actual.filename
        else:
            if not Path(file1_path).is_absolute():
                file1_path = str(BASE / file1_path)
            if not Path(file2_path).is_absolute():
                file2_path = str(BASE / file2_path)
            df1 = load_excel(file1_path, sheet=sheet1)
            df2 = load_excel(file2_path, sheet=sheet2)
            file1_label, file2_label = file1_path, file2_path
    except FileNotFoundError as e:
        return jsonify({
            "error": f"File not found: {e}",
            "code": "file_not_found",
        }), 400
    except Exception as e:
        return jsonify({"error": str(e), "code": "load_error"}), 400

    if key is None:
        # Position-based comparison (no primary key)
        data = get_comparison_for_frontend_by_position(df1, df2)
        data["config"] = {"file1": file1_label, "file2": file2_label, "key": None, "expected_sheet": sheet1, "actual_sheet": sheet2, "mode": "position"}
        return jsonify(data)

    # Primary key mode: key must exist in both files
    cols1 = list(df1.columns)
    cols2 = list(df2.columns)
    if key not in cols1 or key not in cols2:
        in1 = key in cols1
        in2 = key in cols2
        if not in1 and not in2:
            msg = f"Primary key '{key}' not found in either file."
        elif not in1:
            msg = f"Primary key '{key}' not found in file1."
        else:
            msg = f"Primary key '{key}' not found in file2."
        return jsonify({
            "error": msg,
            "code": "primary_key_missing",
            "key": key,
            "columns_file1": cols1,
            "columns_file2": cols2,
        }), 400

    try:
        _validate_primary_key(df1, key, "file1")
        _validate_primary_key(df2, key, "file2")
    except ValueError as e:
        return jsonify({
            "error": str(e),
            "code": "primary_key_invalid",
            "key": key,
        }), 400

    data = get_comparison_for_frontend(df1, df2, key)
    data["config"] = {"file1": file1_label, "file2": file2_label, "key": key, "expected_sheet": sheet1, "actual_sheet": sheet2, "mode": "primary_key"}
    return jsonify(data)


@app.route("/api/export-excel", methods=["POST"])
def api_export_excel():
    """Run same comparison as /api/compare and return Excel file. Same form params as compare."""
    key_raw = (request.form.get("key") or "EMP_ID").strip()
    key = None if key_raw.lower() in ("", "none") else key_raw
    sheet1_raw = request.form.get("expected_sheet") or request.form.get("sheet") or "0"
    sheet2_raw = request.form.get("actual_sheet") or request.form.get("sheet") or "0"
    sheet1 = int(sheet1_raw) if isinstance(sheet1_raw, str) and sheet1_raw.isdigit() else 0
    sheet2 = int(sheet2_raw) if isinstance(sheet2_raw, str) and sheet2_raw.isdigit() else 0
    f_expected = request.files.get("expected_file") or request.files.get("file1")
    f_actual = request.files.get("actual_file") or request.files.get("file2")
    if not f_expected or not f_actual or not f_expected.filename or not f_actual.filename:
        return jsonify({"error": "Upload both Expected and Actual files"}), 400
    try:
        df1 = load_excel(f_expected, sheet=sheet1)
        df2 = load_excel(f_actual, sheet=sheet2)
        file1_label, file2_label = f_expected.filename, f_actual.filename
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    if key is None:
        added, removed, modified_list, meta = compare_excels_by_position(df1, df2)
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            out_path = tmp.name
        try:
            export_to_excel_by_position(
                added, removed, modified_list, meta, out_path,
                file1_label, file2_label, sheet1,
            )
            return send_file(
                out_path,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name="comparison_result.xlsx",
            )
        finally:
            Path(out_path).unlink(missing_ok=True)
    if key not in list(df1.columns) or key not in list(df2.columns):
        return jsonify({"error": f"Primary key '{key}' not in both files"}), 400
    try:
        _validate_primary_key(df1, key, "file1")
        _validate_primary_key(df2, key, "file2")
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    added, removed, changed, meta = compare_excels(df1, df2, key)
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        out_path = tmp.name
    try:
        export_to_excel(
            added, removed, changed, meta, out_path,
            file1_label, file2_label, sheet1, key,
        )
        return send_file(
            out_path,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="comparison_result.xlsx",
        )
    finally:
        Path(out_path).unlink(missing_ok=True)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
