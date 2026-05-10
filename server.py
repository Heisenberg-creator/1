import os
import tempfile
from pathlib import Path

from flask import Flask, render_template_string, request, send_file, jsonify
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB max

UPLOAD_DIR = Path(tempfile.gettempdir()) / "pdf_word_converter"
UPLOAD_DIR.mkdir(exist_ok=True)

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PDF & Word 互转工具</title>
<style>
  :root {
    --bg: #f5f5f5;
    --card-bg: #ffffff;
    --border: #e0e0e0;
    --text: #333333;
    --text-secondary: #888888;
    --accent: #4f46e5;
    --accent-hover: #4338ca;
    --success: #16a34a;
    --danger: #dc2626;
    --radius: 12px;
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    background: var(--bg);
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
  }
  .container {
    width: 100%;
    max-width: 520px;
  }
  h1 {
    text-align: center;
    font-size: 1.5rem;
    font-weight: 600;
    color: var(--text);
    margin-bottom: 8px;
  }
  .subtitle {
    text-align: center;
    font-size: 0.875rem;
    color: var(--text-secondary);
    margin-bottom: 28px;
  }
  .drop-zone {
    background: var(--card-bg);
    border: 2px dashed var(--border);
    border-radius: var(--radius);
    padding: 48px 24px;
    text-align: center;
    cursor: pointer;
    transition: border-color 0.2s, background 0.2s;
    position: relative;
  }
  .drop-zone.drag-over {
    border-color: var(--accent);
    background: #eef2ff;
  }
  .drop-zone.has-file {
    border-style: solid;
    border-color: var(--accent);
  }
  .drop-icon {
    width: 48px;
    height: 48px;
    margin: 0 auto 16px;
    background: #eef2ff;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .drop-icon svg { width: 24px; height: 24px; stroke: var(--accent); }
  .drop-text { font-size: 1rem; color: var(--text); font-weight: 500; }
  .drop-hint { font-size: 0.8rem; color: var(--text-secondary); margin-top: 6px; }
  .file-info {
    display: none;
    margin-top: 16px;
    padding: 12px 16px;
    background: #f0fdf4;
    border: 1px solid #bbf7d0;
    border-radius: 8px;
    align-items: center;
    gap: 10px;
  }
  .file-info.show { display: flex; }
  .file-info .name { flex: 1; font-size: 0.875rem; color: var(--text); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .file-info .type-badge {
    font-size: 0.7rem;
    padding: 2px 8px;
    border-radius: 100px;
    background: var(--accent);
    color: #fff;
    font-weight: 600;
    text-transform: uppercase;
    flex-shrink: 0;
  }
  .file-info .remove-btn {
    background: none;
    border: none;
    cursor: pointer;
    color: var(--text-secondary);
    padding: 4px;
    line-height: 0;
    flex-shrink: 0;
  }
  .file-info .remove-btn:hover { color: var(--danger); }
  .btn {
    display: none;
    width: 100%;
    margin-top: 16px;
    padding: 12px;
    background: var(--accent);
    color: #fff;
    border: none;
    border-radius: 8px;
    font-size: 0.95rem;
    font-weight: 500;
    cursor: pointer;
    transition: background 0.2s;
  }
  .btn:hover { background: var(--accent-hover); }
  .btn:disabled { opacity: 0.6; cursor: not-allowed; }
  .btn.show { display: block; }
  .result {
    display: none;
    margin-top: 16px;
    padding: 16px;
    background: #f0fdf4;
    border: 1px solid #bbf7d0;
    border-radius: 8px;
    text-align: center;
  }
  .result.show { display: block; }
  .result .success-text { color: var(--success); font-weight: 500; font-size: 0.9rem; }
  .result .download-btn {
    display: inline-block;
    margin-top: 10px;
    padding: 8px 24px;
    background: var(--success);
    color: #fff;
    border-radius: 6px;
    text-decoration: none;
    font-size: 0.875rem;
    font-weight: 500;
    transition: background 0.2s;
  }
  .result .download-btn:hover { background: #15803d; }
  .error-msg {
    display: none;
    margin-top: 16px;
    padding: 12px 16px;
    background: #fef2f2;
    border: 1px solid #fecaca;
    border-radius: 8px;
    color: var(--danger);
    font-size: 0.875rem;
  }
  .error-msg.show { display: block; }
  .spinner {
    display: none;
    width: 20px;
    height: 20px;
    border: 2px solid #e0e0e0;
    border-top-color: var(--accent);
    border-radius: 50%;
    animation: spin 0.6s linear infinite;
    margin: 16px auto 0;
  }
  .spinner.show { display: block; }
  @keyframes spin { to { transform: rotate(360deg); } }
  input[type="file"] { display: none; }
</style>
</head>
<body>
<div class="container">
  <h1>PDF &harr; Word 互转</h1>
  <p class="subtitle">拖入 PDF 或 Word 文件，自动识别并转换</p>

  <div class="drop-zone" id="dropZone">
    <div class="drop-icon">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
        <polyline points="17 8 12 3 7 8"/>
        <line x1="12" y1="3" x2="12" y2="15"/>
      </svg>
    </div>
    <div class="drop-text">拖拽文件到此处</div>
    <div class="drop-hint">支持 .pdf / .docx 格式，最大 50MB</div>
  </div>

  <input type="file" id="fileInput" accept=".pdf,.docx">

  <div class="file-info" id="fileInfo">
    <span class="type-badge" id="typeBadge"></span>
    <span class="name" id="fileName"></span>
    <button class="remove-btn" id="removeBtn" title="移除">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
    </button>
  </div>

  <button class="btn" id="convertBtn">开始转换</button>
  <div class="spinner" id="spinner"></div>
  <div class="result" id="result"></div>
  <div class="error-msg" id="errorMsg"></div>
</div>

<script>
  const dropZone = document.getElementById("dropZone");
  const fileInput = document.getElementById("fileInput");
  const fileInfo = document.getElementById("fileInfo");
  const fileName = document.getElementById("fileName");
  const typeBadge = document.getElementById("typeBadge");
  const removeBtn = document.getElementById("removeBtn");
  const convertBtn = document.getElementById("convertBtn");
  const spinner = document.getElementById("spinner");
  const result = document.getElementById("result");
  const errorMsg = document.getElementById("errorMsg");

  let currentFile = null;

  dropZone.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", (e) => {
    if (e.target.files.length > 0) handleFile(e.target.files[0]);
  });

  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("drag-over");
  });
  dropZone.addEventListener("dragleave", () => dropZone.classList.remove("drag-over"));
  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("drag-over");
    if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
  });

  removeBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    resetUI();
  });

  convertBtn.addEventListener("click", () => {
    if (!currentFile) return;
    convertFile();
  });

  function handleFile(file) {
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext !== "pdf" && ext !== "docx") {
      showError("仅支持 .pdf 和 .docx 格式的文件");
      return;
    }
    currentFile = file;
    fileName.textContent = file.name;
    typeBadge.textContent = ext;
    fileInfo.classList.add("show");
    convertBtn.classList.add("show");
    dropZone.classList.add("has-file");
    result.classList.remove("show");
    errorMsg.classList.remove("show");
    result.innerHTML = "";
  }

  function resetUI() {
    currentFile = null;
    fileInput.value = "";
    fileInfo.classList.remove("show");
    convertBtn.classList.remove("show");
    dropZone.classList.remove("has-file");
    result.classList.remove("show");
    errorMsg.classList.remove("show");
    spinner.classList.remove("show");
    result.innerHTML = "";
  }

  function showError(msg) {
    errorMsg.textContent = msg;
    errorMsg.classList.add("show");
    result.classList.remove("show");
    result.innerHTML = "";
  }

  async function convertFile() {
    convertBtn.style.display = "none";
    spinner.classList.add("show");
    result.classList.remove("show");
    errorMsg.classList.remove("show");

    const formData = new FormData();
    formData.append("file", currentFile);

    try {
      const resp = await fetch("/convert", { method: "POST", body: formData });
      if (!resp.ok) {
        const data = await resp.json();
        throw new Error(data.error || "转换失败");
      }
      const blob = await resp.blob();
      const url = URL.createObjectURL(blob);
      const ext = currentFile.name.split(".").pop().toLowerCase();
      const targetExt = ext === "pdf" ? "docx" : "pdf";
      const downloadName = currentFile.name.replace(new RegExp("\\." + ext + "$"), "." + targetExt);

      result.innerHTML = `
        <div class="success-text">转换完成</div>
        <a class="download-btn" href="${url}" download="${downloadName}">下载 ${targetExt.toUpperCase()} 文件</a>
      `;
      result.classList.add("show");
    } catch (err) {
      showError(err.message);
    } finally {
      spinner.classList.remove("show");
      convertBtn.style.display = "block";
    }
  }
</script>
</body>
</html>"""


@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route("/convert", methods=["POST"])
def convert():
    file = request.files.get("file")
    if not file or file.filename == "":
        return jsonify({"error": "请选择一个文件"}), 400

    filename = secure_filename(file.filename)
    ext = Path(filename).suffix.lower()

    if ext not in (".pdf", ".docx"):
        return jsonify({"error": "仅支持 .pdf 和 .docx 格式"}), 400

    input_path = UPLOAD_DIR / filename
    file.save(str(input_path))

    try:
        if ext == ".pdf":
            # PDF → Word
            output_path = input_path.with_suffix(".docx")
            _pdf_to_word(str(input_path), str(output_path))
        else:
            # Word → PDF
            output_path = input_path.with_suffix(".pdf")
            _word_to_pdf(str(input_path), str(output_path))

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_path.name,
            mimetype="application/octet-stream",
        )
    except Exception as e:
        return jsonify({"error": f"转换失败: {str(e)}"}), 500
    finally:
        # Clean up input file
        try:
            input_path.unlink(missing_ok=True)
        except OSError:
            pass


def _pdf_to_word(input_path: str, output_path: str):
    from pdf2docx import Converter
    cv = Converter(input_path)
    cv.convert(output_path)
    cv.close()


def _word_to_pdf(input_path: str, output_path: str):
    from docx2pdf import convert
    convert(input_path, output_path)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
