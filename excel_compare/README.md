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

---

## Push to GitHub

1. **Create a new repository** on [GitHub](https://github.com/new). Name it e.g. `excel-compare`. Do **not** add a README, .gitignore, or license (this project already has them).

2. **From the `excel_compare` folder** on your PC, run:

```bash
cd excel_compare
git init
git add .
git commit -m "Initial commit: Excel comparison tool with web UI"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/excel-compare.git
git push -u origin main
```

Replace `YOUR_USERNAME/excel-compare` with your GitHub username and repo name. If GitHub asks for login, use a [personal access token](https://github.com/settings/tokens) instead of a password.

3. **Later updates:** `git add .` → `git commit -m "Your message"` → `git push`

The `.gitignore` excludes `venv/`, `__pycache__/`, and generated `output/*.xlsx` so they are not pushed.

---

## Upload this folder to a VPS

Run these from your **PC**, in the folder that **contains** `excel_compare`. Replace `USER` with your VPS username and `VPS_IP` with the server IP.

**Option 1: SCP (simple, works in PowerShell/CMD/Git Bash)**

```bash
scp -r excel_compare USER@VPS_IP:/var/www/
```

**Option 2: rsync (good for re-uploading only changes)**

```bash
rsync -avz --progress excel_compare/ USER@VPS_IP:/var/www/excel_compare/
```

**Option 3: Zip then SCP (if folder is large)**

On PC (PowerShell): `Compress-Archive -Path excel_compare -DestinationPath excel_compare.zip`  
Then: `scp excel_compare.zip USER@VPS_IP:/var/www/`  
On VPS: `cd /var/www && unzip excel_compare.zip`

**Option 4: Git (after pushing to GitHub)**

On the VPS: `git clone https://github.com/YOUR_USERNAME/excel-compare.git` then `cd excel-compare`.

**After upload, on the VPS:**

```bash
cd /var/www/excel_compare
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
gunicorn --bind 127.0.0.1:5001 app:app
```

For full deployment (systemd, Nginx, firewall, running alongside another Flask app), see **[DEPLOY.md](DEPLOY.md)**.
