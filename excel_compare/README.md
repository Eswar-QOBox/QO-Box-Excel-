# Excel Compare (Pandas)

Compare two Excel files and export a single Excel report with:

- Added rows
- Removed rows
- Changed cells (before/after)
- Summary (counts + column differences)

## Setup

```bash
pip install -r requirements.txt
```

## Put your files here

- `input/file1.xlsx`
- `input/file2.xlsx`

## Run

From the `excel_compare/` folder:

**With a primary key** (recommended when rows have a unique ID):

```bash
python compare_excel.py --key EMP_ID
```

**By row position only** (no primary key; row 1 vs row 1, row 2 vs row 2, etc.):

```bash
python compare_excel.py --no-key
```

Optional flags:

```bash
python compare_excel.py --file1 input/file1.xlsx --file2 input/file2.xlsx --key EMP_ID --sheet 0 --output output/comparison_result.xlsx
```

## Output

The report is written to `output/comparison_result.xlsx` with sheets:

- `Summary`
- `Added_Rows`
- `Removed_Rows`
- `Changed_Cells`

Use `--json` to also write a frontend-ready JSON file:

```bash
python compare_excel.py --key EMP_ID --json
```

This creates `output/comparison_result.json` with `added_rows`, `removed_rows`, `modified_rows` (id + list of column changes with old/new values), and `summary`.

## Web frontend

To view added / removed / modified rows in a browser:

```bash
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5000 and click **Run comparison**. Check **Compare by row position (no primary key)** to compare by row index instead of a key column. The page shows summary counts and tables for added rows, removed rows, and modified rows (with per-column old → new values).

API: `GET /api/compare` returns the same JSON (uses default `input/file1.xlsx` and `input/file2.xlsx`). Use `?key=none` for position-based comparison. Optional query params: `key`, `sheet`, `file1`, `file2`.

### API documentation (Swagger & Postman)

- **Swagger UI**: With the app running, open **http://127.0.0.1:5000/api-docs** for interactive API docs (OpenAPI 3.0).
- **Postman**: Import the collection from **`postman/Excel_Compare_API.postman_collection.json`**. Set the `baseUrl` variable (e.g. `http://localhost:5000`) and use the requests for file-info, preview, compare, and export.

For production deployment (systemd, Nginx, gunicorn), see **[DEPLOY.md](DEPLOY.md)**.
