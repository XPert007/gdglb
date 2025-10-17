from flask import Flask, jsonify, send_file
import pandas as pd
import os, re
from datetime import datetime

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILENAME = (
    "data.xlsx"
)
EXCEL_PATH = os.path.join(BASE_DIR, EXCEL_FILENAME)
DISPLAY_HTML = os.path.join(BASE_DIR, "display.html")


def _cleanup_name(s):
    if s is None:
        return ""
    s = str(s).strip()
    s = re.sub(r"[\x00-\x1f\x7f]", "", s)
    return s


def _extract_two_numbers(cell):
    if cell is None:
        return 0, 0
    s = str(cell)
    nums = re.findall(r"\d+", s)
    if len(nums) >= 2:
        return int(nums[0]), int(nums[1])
    if len(nums) == 1:
        return int(nums[0]), 0
    parts = re.split(r"[/&;,\\|\\-]+", s)
    extracted = []
    for p in parts:
        pnums = re.findall(r"\d+", p)
        if pnums:
            extracted.extend([int(n) for n in pnums])
    if len(extracted) >= 2:
        return extracted[0], extracted[1]
    if len(extracted) == 1:
        return extracted[0], 0
    return 0, 0


def load_data():
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Excel file not found at: {EXCEL_PATH}")

    if EXCEL_PATH.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(EXCEL_PATH, sheet_name=0)
    else:
        df = pd.read_csv(EXCEL_PATH)

    orig_cols = list(df.columns)
    cols_lc = [str(c).strip().lower() for c in orig_cols]
    col_map = dict(zip(cols_lc, orig_cols))

    wanted_badges = "# of skill badges completed"
    wanted_arcade = "# of arcade games completed"
    badges_col = None
    arcade_col = None
    name_col = None

    for k, orig in zip(cols_lc, orig_cols):
        if "name" == k or "user name" == k or "username" == k or "user" == k:
            name_col = orig
            break

    for k, orig in zip(cols_lc, orig_cols):
        if k == wanted_badges:
            badges_col = orig
        if k == wanted_arcade:
            arcade_col = orig

    if name_col is None:
        for k, orig in zip(cols_lc, orig_cols):
            if "name" in k or "player" in k or "user" in k:
                name_col = orig
                break

    if badges_col is None:
        for k, orig in zip(cols_lc, orig_cols):
            if "badge" in k or "skill" in k or "medal" in k:
                badges_col = orig
                break
    if arcade_col is None:
        for k, orig in zip(cols_lc, orig_cols):
            if "arcade" in k or "game" in k or "played" in k:
                arcade_col = orig
                break

    combined_col = None
    for k, orig in zip(cols_lc, orig_cols):
        if ("badge" in k or "skill" in k or "badg" in k) and (
            "game" in k or "arcade" in k or "play" in k
        ):
            combined_col = orig
            break
    if combined_col and (badges_col is None or arcade_col is None):
        badges_col = arcade_col = combined_col

    if (name_col is None or badges_col is None or arcade_col is None) and len(
        orig_cols
    ) == 3:
        name_col = name_col or orig_cols[0]
        remaining = [c for c in orig_cols if c != name_col]
        numeric = df.select_dtypes(include="number").columns.tolist()
        if len(numeric) >= 2:
            possible = [c for c in numeric if c != name_col]
            if len(possible) >= 2:
                badges_col = badges_col or possible[0]
                arcade_col = arcade_col or possible[1]
            else:
                badges_col = badges_col or remaining[0]
                arcade_col = arcade_col or remaining[1]
        else:
            badges_col = badges_col or remaining[0]
            arcade_col = arcade_col or remaining[1]

    if badges_col is None or arcade_col is None:
        numeric_cols = list(df.select_dtypes(include="number").columns)
        if name_col in numeric_cols:
            numeric_cols.remove(name_col)
        if len(numeric_cols) >= 2:
            badges_col = badges_col or numeric_cols[0]
            arcade_col = arcade_col or numeric_cols[1]

    if name_col is None or badges_col is None or arcade_col is None:
        raise ValueError(
            f"Could not detect required columns. Found columns: {orig_cols}\n"
            f"Attempted to match '# of Skill Badges Completed' and '# of Arcade Games Completed' first."
        )

    names = df[name_col].fillna("").astype(str).apply(_cleanup_name)

    badges = []
    arcades = []
    if badges_col == arcade_col:
        for cell in df[badges_col]:
            b, a = _extract_two_numbers(cell)
            badges.append(b)
            arcades.append(a)
    else:
        badges = (
            pd.to_numeric(df[badges_col], errors="coerce")
            .fillna(0)
            .astype(int)
            .tolist()
        )
        arcades = (
            pd.to_numeric(df[arcade_col], errors="coerce")
            .fillna(0)
            .astype(int)
            .tolist()
        )

    recs = []
    for i, nm in enumerate(names):
        recs.append(
            {
                "Name": nm,
                "Badges": int(badges[i]) if i < len(badges) else 0,
                "Arcade": int(arcades[i]) if i < len(arcades) else 0,
            }
        )

    recs = sorted(recs, key=lambda r: (-r["Badges"], -r["Arcade"], r["Name"]))
    mtime = datetime.fromtimestamp(os.path.getmtime(EXCEL_PATH)).isoformat()
    return {"generated_at": mtime, "records": recs}


@app.route("/data")
def data():
    try:
        payload = load_data()
        return jsonify(payload)
    except FileNotFoundError as fnf:
        return jsonify({"error": str(fnf)}), 400
    except Exception as e:
        return jsonify({"error": "Failed to parse spreadsheet: " + str(e)}), 500


@app.route("/display.html")
def display():
    if os.path.exists(DISPLAY_HTML):
        return send_file(DISPLAY_HTML)
    return (
        "display.html not found on server. Place the provided display.html in the same folder.",
        500,
    )


@app.route("/")
def root():
    return display()


if __name__ == "__main__":

    port = int(os.environ.get("PORT", 5000))
    host = os.environ.get("HOST", "0.0.0.0") 
    print(f"Server starting. Excel expected at: {EXCEL_PATH}. Binding to {host}:{port}")

    app.run(host=host, port=port, debug=False)
