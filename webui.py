"""FastAPI Web UI — 上传 JSON → 运行流水线 → 下载 Excel"""
import io
import sys
import uuid
import shutil
import tempfile
from pathlib import Path
from threading import Thread
from typing import Dict, List

from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse

import config  # noqa: F401 — 确保 .env 已加载
from pipeline import run

app = FastAPI(title="产品规划流水线")

tasks: Dict[str, dict] = {}

HTML = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>产品规划流水线</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #f5f5f5; color: #1a1a1a; min-height: 100vh; display: flex; justify-content: center; align-items: flex-start; padding: 40px 20px; }
  .container { width: 100%; max-width: 600px; }
  h1 { font-size: 24px; font-weight: 600; margin-bottom: 8px; }
  .subtitle { color: #666; font-size: 14px; margin-bottom: 24px; }
  .upload-zone { border: 2px dashed #ccc; border-radius: 12px; padding: 48px 24px; text-align: center; background: #fff; cursor: pointer; transition: border-color 0.2s, background 0.2s; margin-bottom: 16px; }
  .upload-zone:hover, .upload-zone.drag-over { border-color: #2563eb; background: #eff6ff; }
  .upload-zone.has-files { border-color: #22c55e; background: #f0fdf4; }
  .upload-icon { font-size: 40px; margin-bottom: 12px; opacity: 0.6; }
  .upload-text { font-size: 15px; color: #333; }
  .upload-hint { font-size: 12px; color: #999; margin-top: 8px; }
  .file-list { margin-bottom: 16px; }
  .file-item { background: #fff; border: 1px solid #e5e5e5; border-radius: 8px; padding: 10px 14px; margin-bottom: 6px; font-size: 13px; display: flex; align-items: center; gap: 8px; }
  .file-item .name { flex: 1; word-break: break-all; }
  .file-item .remove { color: #ef4444; cursor: pointer; font-size: 18px; line-height: 1; padding: 0 4px; }
  .btn { display: block; width: 100%; padding: 12px 24px; border: none; border-radius: 8px; font-size: 16px; font-weight: 500; cursor: pointer; transition: opacity 0.2s; }
  .btn:disabled { opacity: 0.5; cursor: not-allowed; }
  .btn-run { background: #2563eb; color: #fff; }
  .btn-run:hover:not(:disabled) { background: #1d4ed8; }
  .btn-download { background: #16a34a; color: #fff; margin-top: 12px; }
  .btn-download:hover:not(:disabled) { background: #15803d; }
  .log-box { background: #1e1e1e; color: #d4d4d4; border-radius: 8px; padding: 16px; margin-top: 16px; font-family: "Cascadia Code", "Fira Code", monospace; font-size: 12px; line-height: 1.6; max-height: 300px; overflow-y: auto; white-space: pre-wrap; display: none; }
  .log-box.visible { display: block; }
  .status-bar { display: flex; align-items: center; gap: 8px; margin-top: 12px; font-size: 13px; color: #666; }
  .spinner { width: 16px; height: 16px; border: 2px solid #e5e5e5; border-top-color: #2563eb; border-radius: 50%; animation: spin 0.8s linear infinite; display: none; }
  .spinner.active { display: inline-block; }
  @keyframes spin { to { transform: rotate(360deg); } }
  input[type="file"] { display: none; }
</style>
</head>
<body>
<div class="container">
  <h1>产品规划流水线</h1>
  <p class="subtitle">上传大润发商品 JSON 文件，自动解析、提取图片、AI 分析，输出 Excel 报表</p>

  <div class="upload-zone" id="dropZone">
    <div class="upload-icon">JSON</div>
    <div class="upload-text">拖拽 .txt / .json 文件到此处</div>
    <div class="upload-hint">或点击选择文件（支持多选）</div>
  </div>
  <input type="file" id="fileInput" multiple accept=".txt,.json">

  <div class="file-list" id="fileList"></div>

  <button class="btn btn-run" id="btnRun" disabled>开始分析</button>

  <div class="log-box" id="logBox"></div>
  <div class="status-bar">
    <div class="spinner" id="spinner"></div>
    <span id="statusText"></span>
  </div>
  <button class="btn btn-download" id="btnDownload" style="display:none">下载 Excel</button>
</div>

<script>
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileList = document.getElementById('fileList');
const btnRun = document.getElementById('btnRun');
const logBox = document.getElementById('logBox');
const spinner = document.getElementById('spinner');
const statusText = document.getElementById('statusText');
const btnDownload = document.getElementById('btnDownload');

let files = [];

function updateUI() {
  fileList.innerHTML = files.map((f, i) =>
    `<div class="file-item"><span class="name">${f.name}</span><span class="remove" onclick="removeFile(${i})">&times;</span></div>`
  ).join('');
  if (files.length > 0) {
    dropZone.classList.add('has-files');
  } else {
    dropZone.classList.remove('has-files');
  }
  btnRun.disabled = files.length === 0;
}

function addFiles(newFiles) {
  for (const f of newFiles) {
    if (!files.some(ex => ex.name === f.name)) {
      files.push(f);
    }
  }
  updateUI();
}

function removeFile(i) {
  files.splice(i, 1);
  updateUI();
}

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  if (e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
});
fileInput.addEventListener('change', () => {
  if (fileInput.files.length) addFiles(fileInput.files);
  fileInput.value = '';
});

btnRun.addEventListener('click', async () => {
  btnRun.disabled = true;
  logBox.classList.add('visible');
  logBox.textContent = '';
  spinner.classList.add('active');
  statusText.textContent = '运行中...';
  btnDownload.style.display = 'none';

  const form = new FormData();
  files.forEach(f => form.append('files', f));

  try {
    const res = await fetch('/api/run', { method: 'POST', body: form });
    const data = await res.json();
    if (data.status !== 'running') {
      throw new Error(data.msg || '启动失败');
    }
    const taskId = data.task_id;
    pollStatus(taskId);
  } catch (err) {
    spinner.classList.remove('active');
    statusText.textContent = '错误: ' + err.message;
    btnRun.disabled = files.length === 0;
  }
});

async function pollStatus(taskId) {
  try {
    const res = await fetch('/api/status/' + taskId);
    const data = await res.json();

    if (data.log) {
      logBox.textContent = data.log.join('\\n');
      logBox.scrollTop = logBox.scrollHeight;
    }

    if (data.status === 'done') {
      spinner.classList.remove('active');
      statusText.textContent = '完成 — ' + data.records + ' 条商品记录';
      btnDownload.style.display = 'block';
      btnDownload.onclick = () => { window.location.href = '/api/download/' + taskId; };
      btnRun.disabled = files.length === 0;
      return;
    }

    if (data.status === 'error') {
      spinner.classList.remove('active');
      statusText.textContent = '错误: ' + (data.msg || '未知错误');
      btnRun.disabled = files.length === 0;
      return;
    }

    setTimeout(() => pollStatus(taskId), 1000);
  } catch (err) {
    spinner.classList.remove('active');
    statusText.textContent = '连接错误';
    btnRun.disabled = files.length === 0;
  }
}
</script>
</body>
</html>"""


@app.get("/", response_class=HTMLResponse)
async def index():
    return HTML


@app.post("/api/run")
async def api_run(files: List[UploadFile] = File(...)):
    task_id = uuid.uuid4().hex[:8]
    temp_dir = Path(tempfile.mkdtemp())
    for f in files:
        content = await f.read()
        (temp_dir / f.filename).write_bytes(content)

    temp_out = Path(tempfile.mkdtemp())
    tasks[task_id] = {"status": "running", "log": []}

    def _run():
        captured_log = []
        old_stdout = sys.stdout
        sys.stdout = io.StringIO()
        try:
            records = run(input_dir=str(temp_dir), output_dir=str(temp_out))
            output = sys.stdout.getvalue()
            captured_log = [l for l in output.split("\n") if l.strip()]
            xlsx_files = list(temp_out.rglob("*.xlsx"))
            if xlsx_files:
                tasks[task_id] = {
                    "status": "done",
                    "file": str(xlsx_files[0]),
                    "filename": xlsx_files[0].name,
                    "records": len(records),
                    "log": captured_log,
                }
            else:
                output_lines = [l for l in output.split("\n") if l.strip()]
                tasks[task_id] = {"status": "error", "msg": "未生成Excel文件", "log": output_lines}
        except SystemExit:
            output = sys.stdout.getvalue()
            tasks[task_id] = {"status": "error", "msg": "流水线执行失败，请检查上传文件格式", "log": [l for l in output.split("\n") if l.strip()]}
        except Exception:
            import traceback
            tasks[task_id] = {"status": "error", "msg": traceback.format_exc(), "log": captured_log}
        finally:
            sys.stdout = old_stdout
            shutil.rmtree(temp_dir, ignore_errors=True)

    Thread(target=_run, daemon=True).start()
    return {"task_id": task_id, "status": "running"}


@app.get("/api/status/{task_id}")
async def api_status(task_id: str):
    task = tasks.get(task_id)
    if not task:
        return {"status": "not_found"}
    return task


@app.get("/api/download/{task_id}")
async def api_download(task_id: str):
    task = tasks.get(task_id)
    if not task or task["status"] != "done":
        return {"error": "文件不存在"}
    return FileResponse(
        task["file"],
        filename=task["filename"],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("webui:app", host="0.0.0.0", port=7860, reload=True)
