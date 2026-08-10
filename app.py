#!/usr/bin/env python3
"""
Local web frontend for generate_spi_summary.py.

Drop in your weekly Worksheet .xlsx (and optionally a replacement template),
and it runs the generator automatically and hands back the finished
Summary_DD_MM_YYYY.docx -- no commands to type.

SETUP (once):
    pip install flask

RUN (each week, or just leave it running):
    python3 app.py

Then open http://127.0.0.1:5050 in your browser.

Put this file in the same folder as generate_spi_summary.py, spi_template.docx,
and spi_history.csv -- it uses them exactly as they already work.
"""
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile

from flask import Flask, request, send_file, jsonify

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
GENERATOR = os.path.join(SCRIPT_DIR, "generate_spi_summary.py")
DEFAULT_TEMPLATE = os.path.join(SCRIPT_DIR, "spi_template.docx")
HISTORY_PATH = os.path.join(SCRIPT_DIR, "spi_history.csv")

app = Flask(__name__)

PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SPI Weekly Summary</title>
<style>
  :root{
    --paper:#F6F3EC; --ink:#1B2A4A; --ink-soft:#5A6478;
    --gold:#B8862E; --rule:#D8D2C2; --ok:#2E6E6A; --err:#A14B3A; --panel:#FFFFFF;
  }
  *{box-sizing:border-box;}
  body{margin:0;background:var(--paper);color:var(--ink);
    font-family:Georgia,'Iowan Old Style','Palatino Linotype',serif;
    padding:48px 20px 80px;}
  .wrap{max-width:620px;margin:0 auto;}
  .masthead{border-top:3px solid var(--ink);border-bottom:1px solid var(--ink);
    padding:16px 0 12px;margin-bottom:32px;text-align:center;}
  .eyebrow{font-family:'Courier New',monospace;letter-spacing:.18em;font-size:.68rem;
    color:var(--gold);text-transform:uppercase;margin-bottom:6px;}
  h1{margin:0;font-size:1.5rem;font-weight:600;}
  .sub{margin:6px 0 0;font-size:.85rem;color:var(--ink-soft);font-style:italic;}

  .drop{background:var(--panel);border:1.5px dashed var(--rule);border-radius:3px;
    padding:30px 20px;text-align:center;cursor:pointer;margin-bottom:16px;
    transition:border-color .15s, background .15s;}
  .drop:hover,.drop.drag{border-color:var(--gold);background:#FAF6EC;}
  .drop .label{font-weight:600;font-size:.95rem;}
  .drop .hint{font-size:.78rem;color:var(--ink-soft);margin-top:4px;}
  .drop .filename{margin-top:10px;font-size:.85rem;color:var(--ok);font-weight:600;}
  input[type=file]{display:none;}

  .optional-toggle{font-size:.78rem;color:var(--ink-soft);text-align:center;
    margin:-6px 0 18px;cursor:pointer;text-decoration:underline;}

  #status{min-height:60px;background:var(--panel);border:1px solid var(--rule);
    border-radius:3px;padding:16px 18px;font-size:.85rem;line-height:1.6;}
  #status .line{margin:0;}
  #status .err{color:var(--err);}
  #status .ok{color:var(--ok);font-weight:600;}
  #status a{color:var(--ink);}

  .spinner{display:inline-block;width:12px;height:12px;border:2px solid var(--rule);
    border-top-color:var(--gold);border-radius:50%;animation:spin .7s linear infinite;
    vertical-align:-2px;margin-right:6px;}
  @keyframes spin{to{transform:rotate(360deg);}}
</style>
</head>
<body>
<div class="wrap">
  <div class="masthead">
    <div class="eyebrow">Weekly Inflation Brief</div>
    <h1>SPI Summary Generator</h1>
    <p class="sub">Drop this week's worksheet. The report builds itself.</p>
  </div>

  <div class="drop" id="dropExcel">
    <div class="label">Worksheet (.xlsx)</div>
    <div class="hint">Click or drag your weekly worksheet here</div>
    <div class="filename" id="excelName"></div>
    <input type="file" id="excelInput" accept=".xlsx">
  </div>

  <div class="optional-toggle" id="toggleTemplate">+ use a different template for this run</div>
  <div class="drop" id="dropTemplate" style="display:none;">
    <div class="label">Template (.docx)</div>
    <div class="hint">Optional &mdash; defaults to spi_template.docx</div>
    <div class="filename" id="templateName"></div>
    <input type="file" id="templateInput" accept=".docx">
  </div>

  <div id="status"><p class="line">Waiting for a worksheet&hellip;</p></div>
</div>

<script>
let excelFile = null, templateFile = null;

function wireDrop(dropId, inputId, onFile){
  const drop = document.getElementById(dropId), input = document.getElementById(inputId);
  drop.addEventListener('click', () => input.click());
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('drag'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('drag'));
  drop.addEventListener('drop', e => {
    e.preventDefault(); drop.classList.remove('drag');
    if (e.dataTransfer.files.length) onFile(e.dataTransfer.files[0]);
  });
  input.addEventListener('change', () => { if (input.files.length) onFile(input.files[0]); });
}

wireDrop('dropExcel', 'excelInput', f => {
  excelFile = f;
  document.getElementById('excelName').textContent = f.name;
  runGeneration();
});
wireDrop('dropTemplate', 'templateInput', f => {
  templateFile = f;
  document.getElementById('templateName').textContent = f.name;
  if (excelFile) runGeneration();
});

document.getElementById('toggleTemplate').addEventListener('click', () => {
  const el = document.getElementById('dropTemplate');
  el.style.display = el.style.display === 'none' ? 'block' : 'none';
});

function setStatus(html){ document.getElementById('status').innerHTML = html; }

async function runGeneration(){
  if (!excelFile) return;
  setStatus('<p class="line"><span class="spinner"></span>Reading ' + excelFile.name + ' and building the summary&hellip;</p>');

  const form = new FormData();
  form.append('excel', excelFile);
  if (templateFile) form.append('template', templateFile);

  try {
    const res = await fetch('/generate', { method: 'POST', body: form });
    if (!res.ok) {
      const err = await res.json().catch(() => ({error: 'Unknown error'}));
      setStatus('<p class="line err">Could not generate the report: ' + err.error + '</p>');
      return;
    }
    const disposition = res.headers.get('Content-Disposition') || '';
    const match = disposition.match(/filename="?([^"]+)"?/);
    const filename = match ? match[1] : 'Summary.docx';
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
    setStatus('<p class="line ok">&#10003; ' + filename + ' generated and downloaded.</p><p class="line">Drop a new worksheet any time to regenerate.</p>');
  } catch (e) {
    setStatus('<p class="line err">Request failed: ' + e.message + '</p>');
  }
}
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return PAGE


@app.route("/generate", methods=["POST"])
def generate():
    if "excel" not in request.files or request.files["excel"].filename == "":
        return jsonify(error="No worksheet file received"), 400

    # Not using "with tempfile.TemporaryDirectory()" here: on Windows, send_file
    # can still hold the output file open when the context manager tries to
    # delete the folder, causing a PermissionError. Instead we read the finished
    # file fully into memory, THEN clean up the folder, THEN serve the bytes.
    tmp = tempfile.mkdtemp()
    try:
        excel_path = os.path.join(tmp, request.files["excel"].filename)
        request.files["excel"].save(excel_path)

        template_path = DEFAULT_TEMPLATE
        if "template" in request.files and request.files["template"].filename:
            template_path = os.path.join(tmp, request.files["template"].filename)
            request.files["template"].save(template_path)

        # figure out the report date so the output is named the same way
        # generate_spi_summary.py would name it
        out_name = _output_name_from_workbook(excel_path)
        out_path = os.path.join(tmp, out_name)

        result = subprocess.run(
            [sys.executable, GENERATOR,
             "--excel", excel_path, "--output", out_path,
             "--template", template_path, "--history", HISTORY_PATH],
            capture_output=True, text=True,
        )
        if result.returncode != 0 or not os.path.exists(out_path):
            msg = (result.stderr or "Generation failed").strip().splitlines()[-1]
            return jsonify(error=msg), 500

        with open(out_path, "rb") as f:
            file_bytes = f.read()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    buffer = io.BytesIO(file_bytes)
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


def _output_name_from_workbook(excel_path):
    """Reads the report date from the workbook so the download is named
    Summary_DD_MM_YYYY.docx regardless of what the uploaded file was called."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
        ws = wb["Impact Combined"]
        text = ws.cell(row=2, column=1).value
        wb.close()
        m = re.search(r"Current Date\s*=\s*([\d/]+)", text)
        d, mo, y = m.group(1).strip().split("/")
        return f"Summary_{d.zfill(2)}_{mo.zfill(2)}_{y}.docx"
    except Exception:
        stem = os.path.splitext(os.path.basename(excel_path))[0]
        return f"Summary_{stem}.docx"


if __name__ == "__main__":
    if not os.path.exists(GENERATOR):
        sys.exit(f"Can't find generate_spi_summary.py next to app.py at {GENERATOR}")
    if not os.path.exists(DEFAULT_TEMPLATE):
        sys.exit(f"Can't find spi_template.docx next to app.py at {DEFAULT_TEMPLATE}")
    print("SPI Summary Generator running at http://127.0.0.1:5050")
    app.run(host="127.0.0.1", port=5050, debug=False)