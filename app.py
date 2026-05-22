from __future__ import annotations

import logging
import os
import shutil
import time
import uuid
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request, send_from_directory, url_for

from converter import (
    ConversionError,
    convert_pdf_to_editable_docx,
    convert_pdf_to_office2007_docx,
)


BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

RUNTIME_DIR = BASE_DIR / "runtime"
UPLOAD_DIR = RUNTIME_DIR / "uploads"
OUTPUT_DIR = RUNTIME_DIR / "outputs"
MAX_UPLOAD_MB = 80
JOB_TTL_SECONDS = 24 * 60 * 60

for directory in (UPLOAD_DIR, OUTPUT_DIR):
    directory.mkdir(parents=True, exist_ok=True)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/api/convert", methods=["POST"])
def convert():
    cleanup_old_jobs()

    files = request.files.getlist("files")
    if not files or all(not file.filename for file in files):
        return jsonify({"error": "请选择至少一个 PDF 文件。"}), 400

    mode = _parse_mode(request.form.get("mode", "editable"))
    dpi = _parse_dpi(request.form.get("dpi", "220"))
    job_id = uuid.uuid4().hex
    upload_dir = UPLOAD_DIR / job_id
    output_dir = OUTPUT_DIR / job_id
    upload_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    results = []
    errors = []

    try:
        for uploaded_file in files:
            original_name = _clean_filename(uploaded_file.filename)
            if not original_name:
                continue
            if Path(original_name).suffix.lower() != ".pdf":
                errors.append({"file": original_name, "error": "只支持 PDF 文件。"})
                continue

            stored_pdf = upload_dir / f"{uuid.uuid4().hex}.pdf"
            uploaded_file.save(stored_pdf)

            suffix = "editable" if mode == "editable" else "office2007"
            output_name = _unique_filename(output_dir, f"{Path(original_name).stem}_{suffix}.docx")
            output_path = output_dir / output_name

            try:
                if mode == "editable":
                    converted = convert_pdf_to_editable_docx(stored_pdf, output_path)
                else:
                    converted = convert_pdf_to_office2007_docx(
                        stored_pdf,
                        output_path,
                        dpi=dpi,
                    )
            except ConversionError as exc:
                logger.warning("Convert failed for %s: %s", original_name, exc)
                errors.append({"file": original_name, "error": str(exc)})
                continue

            results.append(
                {
                    "source": original_name,
                    "filename": output_name,
                    "pages": converted.page_count,
                    "mode": mode,
                    "dpi": converted.dpi or None,
                    "size": output_path.stat().st_size,
                    "download_url": url_for(
                        "download_file",
                        job_id=job_id,
                        filename=output_name,
                    ),
                }
            )

        if not results:
            shutil.rmtree(output_dir, ignore_errors=True)
            return jsonify({"error": "没有文件转换成功。", "details": errors}), 400

        response = {
            "success": True,
            "job_id": job_id,
            "results": results,
            "errors": errors,
        }

        if len(results) > 1:
            bundle_name = "office2007_docx.zip"
            bundle_path = output_dir / bundle_name
            with ZipFile(bundle_path, "w", ZIP_DEFLATED) as archive:
                for result in results:
                    archive.write(output_dir / result["filename"], result["filename"])

            response["bundle"] = {
                "filename": bundle_name,
                "size": bundle_path.stat().st_size,
                "download_url": url_for(
                    "download_file",
                    job_id=job_id,
                    filename=bundle_name,
                ),
            }

        return jsonify(response)
    except Exception as exc:
        logger.exception("Unexpected conversion error")
        shutil.rmtree(output_dir, ignore_errors=True)
        return jsonify({"error": f"转换失败：{exc}"}), 500
    finally:
        shutil.rmtree(upload_dir, ignore_errors=True)


@app.route("/api/download/<job_id>/<path:filename>")
def download_file(job_id: str, filename: str):
    if not _valid_job_id(job_id):
        return jsonify({"error": "无效的下载地址。"}), 404

    job_dir = OUTPUT_DIR / job_id
    if not job_dir.exists():
        return jsonify({"error": "文件已过期或不存在。"}), 404

    return send_from_directory(job_dir, filename, as_attachment=True)


@app.errorhandler(413)
def file_too_large(_error):
    return jsonify({"error": f"文件太大，单次上传不能超过 {MAX_UPLOAD_MB}MB。"}), 413


def cleanup_old_jobs() -> None:
    cutoff = time.time() - JOB_TTL_SECONDS
    for root in (UPLOAD_DIR, OUTPUT_DIR):
        if not root.exists():
            continue
        for child in root.iterdir():
            try:
                if child.stat().st_mtime < cutoff:
                    if child.is_dir():
                        shutil.rmtree(child, ignore_errors=True)
                    else:
                        child.unlink(missing_ok=True)
            except OSError:
                logger.debug("Unable to remove stale file: %s", child, exc_info=True)


def _parse_dpi(value: str) -> int:
    try:
        dpi = int(value)
    except (TypeError, ValueError):
        return 220
    return min(300, max(100, dpi))


def _parse_mode(value: str) -> str:
    if value == "image":
        return "image"
    return "editable"


def _clean_filename(filename: str) -> str:
    filename = Path(filename or "").name
    invalid_chars = set('<>:"/\\|?*')
    cleaned = "".join(
        char for char in filename if char not in invalid_chars and ord(char) >= 32
    ).strip()
    return cleaned


def _unique_filename(directory: Path, filename: str) -> str:
    path = directory / filename
    if not path.exists():
        return filename

    stem = Path(filename).stem
    suffix = Path(filename).suffix
    index = 2
    while True:
        candidate = f"{stem}_{index}{suffix}"
        if not (directory / candidate).exists():
            return candidate
        index += 1


def _valid_job_id(job_id: str) -> bool:
    return len(job_id) == 32 and all(char in "0123456789abcdef" for char in job_id)


if __name__ == "__main__":
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "7633"))
    debug = os.environ.get("DEBUG", "0").lower() in {"1", "true", "yes"}
    app.run(debug=debug, host=host, port=port, use_reloader=False)
