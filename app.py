import os
import re
import zipfile
import base64
from io import BytesIO
from datetime import datetime

import pandas as pd
import requests
from flask import (
    Flask, render_template, request,
    send_file, flash, redirect, url_for
)
from dotenv import load_dotenv
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from docx.image.exceptions import UnrecognizedImageError

# -------------------- ENV + FLASK SETUP --------------------

load_dotenv()

SECRET = os.getenv("SECRET_KEY")
NIGHT_CHECK_SHEET_URL = os.getenv("NIGHT_CHECK_SHEET_URL", "")

if not SECRET:
    raise RuntimeError("SECRET_KEY is missing in .env")

app = Flask(__name__)
app.secret_key = SECRET


# -------------------- GOOGLE SHEET HELPERS --------------------

def extract_sheet_id(sheet_input: str) -> str:
    pattern = r"/spreadsheets/d/([a-zA-Z0-9-_]+)"
    m = re.search(pattern, sheet_input)
    return m.group(1) if m else sheet_input.strip()


def load_sheet_via_csv(sheet_input: str, gid: str | None = None) -> pd.DataFrame:
    sheet_id = extract_sheet_id(sheet_input)
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    if gid:
        url += f"&gid={gid}"
    df = pd.read_csv(url)
    return df


# -------------------- SITE NAME PARSER --------------------

def parse_site_name(raw: str):
    if not isinstance(raw, str):
        return "", "", ""
    parts = raw.split("-", 2)
    if len(parts) < 3:
        return "", "", raw.strip()
    zone = parts[0].strip()
    unit_code = parts[1].strip()
    sitename = parts[2].strip()
    return zone, unit_code, sitename


# -------------------- GOOGLE DRIVE IMAGE DOWNLOADER --------------------

def extract_drive_file_id(url: str) -> str:
    patterns = [
        r"id=([A-Za-z0-9_-]+)",
        r"/d/([A-Za-z0-9_-]+)/",
    ]
    for p in patterns:
        m = re.search(p, url)
        if m:
            return m.group(1)
    return ""


def download_drive_image(url: str):
    file_id = extract_drive_file_id(url)
    if not file_id:
        return None

    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"

    try:
        r = requests.get(download_url, timeout=15)
        if r.status_code == 200:
            content_type = r.headers.get("Content-Type", "")
            if content_type.startswith("image/"):
                return BytesIO(r.content)
    except Exception:
        return None

    return None


def image_bytes_to_data_uri(img_bytes: BytesIO, mime_type: str) -> str:
    img_bytes.seek(0)
    b64 = base64.b64encode(img_bytes.read()).decode("ascii")
    return f"data:{mime_type};base64,{b64}"


# -------------------- CONTEXT HELPERS --------------------

def build_context_from_row(row: pd.Series):
    raw_site = row.get("Site Name", "")
    zone, unit_code, sitename = parse_site_name(raw_site)

    # --- NEW: format date as DD/MM/YYYY using Date_parsed if available ---
    formatted_date = ""
    parsed_date = row.get("Date_parsed", None)
    if parsed_date is not None and not pd.isna(parsed_date):
        try:
            formatted_date = pd.to_datetime(parsed_date).strftime("%d/%m/%Y")
        except Exception:
            formatted_date = str(row.get("Date", "") or "")
    else:
        formatted_date = str(row.get("Date", "") or "")

    ctx = {
        "zone": zone,
        "unit_code": unit_code,
        "site_name": sitename or raw_site,
        "date": formatted_date,  # ← use formatted date here
        "time": row.get("Time", ""),
        "attendance_register": row.get("Documentation Check [Attendance Register]", ""),
        "handling_register": row.get("Documentation Check [Handling / Taking Over Register]", ""),
        "material_register": row.get("Documentation Check [Visitor Log Register]", ""),
        "grooming": row.get("Performance Check [Grooming]", ""),
        "alertness": row.get("Performance Check [Alertness]", ""),
        "post_discipline": row.get("Performance Check [Post Discipline]", ""),
        "overall_rating": row.get("Performance Check [Overall Rating]", ""),
        "observation": row.get("Observation", ""),
        "inspected_by": row.get("Inspected By", ""),
    }

    for k, v in list(ctx.items()):
        if pd.isna(v):
            ctx[k] = ""
        elif v is None:
            ctx[k] = ""

    return ctx

def get_image_data_uris_for_row(row: pd.Series):
    images_raw = str(row.get("Images", "") or "")
    if not images_raw:
        return []

    uris = []
    urls = [u.strip() for u in images_raw.split(",") if u.strip()]
    for url in urls:
        img_bytes = download_drive_image(url)
        if img_bytes:
            ext = url.split("?")[0].split(".")[-1].lower()
            if ext in ("jpg", "jpeg"):
                mime = "image/jpeg"
            elif ext == "png":
                mime = "image/png"
            else:
                mime = "image/jpeg"
            try:
                uris.append(image_bytes_to_data_uri(img_bytes, mime))
            except Exception:
                continue
    return uris


# -------------------- DOCX TEMPLATE RENDER --------------------

def render_docx_row(row: pd.Series, template_bytes: bytes | None, template_path: str | None) -> BytesIO:
    if template_bytes:
        tpl_stream = BytesIO(template_bytes)
        tpl = DocxTemplate(tpl_stream)
    else:
        if not template_path or not os.path.exists(template_path):
            raise FileNotFoundError("No valid template.docx found.")
        tpl = DocxTemplate(template_path)

    ctx = build_context_from_row(row)

    images_raw = str(row.get("Images", "") or "")
    inline_images = []

    if images_raw:
        urls = [u.strip() for u in images_raw.split(",") if u.strip()]
        for url in urls:
            img_bytes = download_drive_image(url)
            if img_bytes:
                try:
                    inline_images.append(InlineImage(tpl, img_bytes, width=Inches(2.5)))
                except UnrecognizedImageError:
                    continue

    ctx["images"] = inline_images

    tpl.render(ctx)
    buffer = BytesIO()
    tpl.save(buffer)
    buffer.seek(0)
    return buffer


# -------------------- ROUTE: HOME (FORM + ZIP DOWNLOAD) --------------------

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        sheet_input = request.form.get("sheet_input", "").strip()
        gid = request.form.get("gid", "").strip() or None
        date_str = request.form.get("date", "").strip()
        action = request.form.get("action", "download_zip")

        if not sheet_input or not date_str:
            flash("Please enter sheet URL/ID and date.", "error")
            return redirect(url_for("index"))

        try:
            target_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            flash("Invalid date format.", "error")
            return redirect(url_for("index"))

        try:
            df = load_sheet_via_csv(sheet_input, gid)
        except Exception as e:
            flash(f"Unable to load Google Sheet: {e}", "error")
            return redirect(url_for("index"))

        if "Date" not in df.columns:
            flash("Column 'Date' missing in Google Sheet.", "error")
            return redirect(url_for("index"))

        df["Date_parsed"] = pd.to_datetime(df["Date"], errors="coerce")
        df_date = df[df["Date_parsed"].dt.date == target_date]

        if df_date.empty:
            flash(f"No records found for {target_date}.", "warning")
            return redirect(url_for("index"))

        # template upload or default
        uploaded_template = request.files.get("template_file")
        template_bytes = None
        template_path = None

        if uploaded_template and uploaded_template.filename:
            if not uploaded_template.filename.lower().endswith(".docx"):
                flash("Template must be a .docx file.", "error")
                return redirect(url_for("index"))
            template_bytes = uploaded_template.read()
        else:
            template_path = "template.docx"

        if action == "download_zip":
            if not template_bytes and not os.path.exists(template_path):
                flash("No template uploaded and template.docx not found in project folder.", "error")
                return redirect(url_for("index"))

            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for idx, row in df_date.iterrows():
                    docx_buf = render_docx_row(row, template_bytes, template_path)

                    _, _, sitename = parse_site_name(row.get("Site Name", "Site"))
                    site_slug = (sitename or "Site").replace(" ", "_")
                    filename = f"{target_date}_{site_slug}.docx"
                    zipf.writestr(filename, docx_buf.getvalue())

            zip_buffer.seek(0)
            return send_file(
                zip_buffer,
                as_attachment=True,
                download_name=f"night_checks_{target_date}.zip",
                mimetype="application/zip",
            )

        elif action == "preview":
            # redirect to preview route, starting at index 0
            return redirect(url_for(
                "preview",
                sheet_input=sheet_input,
                gid=gid or "",
                date=date_str,
                idx=0
            ))

        else:
            flash("Unknown action.", "error")
            return redirect(url_for("index"))

    return render_template(
    "index.html",
    night_check_url=NIGHT_CHECK_SHEET_URL
)



# -------------------- ROUTE: PREVIEW WITH NEXT / PREVIOUS --------------------

@app.route("/preview")
def preview():
    sheet_input = request.args.get("sheet_input", "").strip()
    gid = request.args.get("gid", "").strip() or None
    date_str = request.args.get("date", "").strip()
    idx_str = request.args.get("idx", "0").strip()

    if not sheet_input or not date_str:
        flash("Missing sheet or date for preview.", "error")
        return redirect(url_for("index"))

    try:
        target_date = datetime.strptime(date_str, "%Y-%m-%d").date()
    except ValueError:
        flash("Invalid date format.", "error")
        return redirect(url_for("index"))

    try:
        df = load_sheet_via_csv(sheet_input, gid)
    except Exception as e:
        flash(f"Unable to load Google Sheet: {e}", "error")
        return redirect(url_for("index"))

    if "Date" not in df.columns:
        flash("Column 'Date' missing in Google Sheet.", "error")
        return redirect(url_for("index"))

    df["Date_parsed"] = pd.to_datetime(df["Date"], errors="coerce")
    df_date = df[df["Date_parsed"].dt.date == target_date]

    if df_date.empty:
        flash(f"No records found for {target_date}.", "warning")
        return redirect(url_for("index"))

    try:
        idx = int(idx_str)
    except ValueError:
        idx = 0

    total = len(df_date)
    if idx < 0:
        idx = 0
    if idx > total - 1:
        idx = total - 1

    row = df_date.iloc[idx]
    ctx = build_context_from_row(row)
    ctx["images_html"] = get_image_data_uris_for_row(row)

    # Navigation info
    ctx["sheet_input"] = sheet_input
    ctx["gid"] = gid or ""
    ctx["date_str"] = date_str
    ctx["current_index"] = idx
    ctx["total"] = total
    ctx["has_prev"] = idx > 0
    ctx["has_next"] = idx < total - 1

    return render_template("preview.html", **ctx)


if __name__ == "__main__":
    app.run(debug=True)
