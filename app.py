"""
Excel 변환기 웹 앱
Flask 없이 Python 표준 라이브러리만 사용 (pandas, openpyxl 필요)
"""
import http.server
import urllib.parse
import io
import cgi
import os

import pandas as pd

# ─────────────────────────────────────────────
#  변환 로직
# ─────────────────────────────────────────────
DATE_COLS   = [24, 43]          # DELI_DATE, CFMD_DATE
SPLIT_COLS  = [32, 33, 34, 49]  # Pallet No, BATCH PROPOSAL, MOLD REQUEST, Prod Date
HEADER_ROWS = 3                 # 행 0·1·2 = 헤더


def fmt_date(val):
    """datetime / string → 'YYYY. M. D' 형식"""
    if val is None or (hasattr(val, '__float__') and pd.isna(val)):
        return val
    try:
        dt = pd.to_datetime(val)
        return f"{dt.year}. {dt.month}. {dt.day}"
    except Exception:
        return val


def parse_semi(val):
    """세미콜론 구분값 파싱.
    단일값 → [원본값] (trailing ';' 보존)
    복수값 → 세미콜론 없는 개별 요소 리스트"""
    if val is None or (hasattr(val, '__float__') and pd.isna(val)):
        return [val]
    s = str(val).strip()
    parts = [p.strip() for p in s.split(';') if p.strip()]
    if len(parts) <= 1:
        return [val]
    return parts


def transform(file_bytes: bytes) -> bytes:
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet = 'before' if 'before' in xl.sheet_names else xl.sheet_names[0]
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=None)

    headers = df.iloc[:HEADER_ROWS]
    data    = df.iloc[HEADER_ROWS:]

    new_rows = []
    for _, row in data.iterrows():
        row = row.copy()

        # 날짜 변환
        for c in DATE_COLS:
            row.iloc[c] = fmt_date(row.iloc[c])

        # 분리 대상 파싱
        sv    = {c: parse_semi(row.iloc[c]) for c in SPLIT_COLS}
        max_n = max(len(v) for v in sv.values())

        if max_n == 1:
            new_rows.append(row.values.tolist())
        else:
            for i in range(max_n):
                r = row.values.tolist()
                for c in SPLIT_COLS:
                    parts = sv[c]
                    r[c] = parts[i] if i < len(parts) else None
                new_rows.append(r)

    all_rows = headers.values.tolist() + new_rows
    out_df   = pd.DataFrame(all_rows)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        out_df.to_excel(writer, index=False, header=False, sheet_name='after')
    return buf.getvalue()


# ─────────────────────────────────────────────
#  HTML 템플릿
# ─────────────────────────────────────────────
HTML_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Excel 변환기 | Before → After</title>
  <link rel="stylesheet"
        href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap/5.3.2/css/bootstrap.min.css">
  <style>
    body { background:#f0f4f8; font-family:'Segoe UI','Malgun Gothic',sans-serif; }
    .hero {
      background:linear-gradient(135deg,#1e3c72 0%,#2a5298 100%);
      color:white; padding:48px 0 36px; text-align:center;
    }
    .hero h1  { font-size:2rem; font-weight:700; margin-bottom:8px; }
    .hero p   { opacity:.8; font-size:1rem; }
    .card-main {
      max-width:680px; margin:40px auto;
      border:none; border-radius:16px;
      box-shadow:0 8px 32px rgba(0,0,0,.12);
    }
    .card-main .card-body { padding:40px; }
    #drop-zone {
      border:2.5px dashed #90aad4; border-radius:12px;
      padding:56px 24px; text-align:center;
      cursor:pointer; transition:all .2s; background:#f7faff;
    }
    #drop-zone.hover { border-color:#2a5298; background:#eaf0fb; }
    #drop-zone .icon { font-size:3rem; margin-bottom:12px; }
    #drop-zone p  { color:#5a7aaa; margin:0; }
    #file-input   { display:none; }
    #file-name    { margin-top:14px; font-size:.92rem; color:#2a5298; font-weight:600; min-height:22px; }
    .btn-convert  {
      background:linear-gradient(135deg,#1e3c72,#2a5298);
      color:white; border:none; border-radius:10px;
      padding:14px 40px; font-size:1.05rem; font-weight:600;
      width:100%; margin-top:24px; letter-spacing:.3px; transition:opacity .2s;
    }
    .btn-convert:hover     { opacity:.88; }
    .btn-convert:disabled  { opacity:.5; cursor:not-allowed; }
    .rules {
      background:#eef3fb; border-radius:10px;
      padding:18px 22px; margin-top:28px; font-size:.88rem; color:#3a5080;
    }
    .rules h6 { font-weight:700; margin-bottom:10px; color:#1e3c72; }
    .spinner-border { width:1.2rem; height:1.2rem; margin-right:8px; }
    #alert-box  { margin-top:20px; display:none; }
    .badge-rule { background:#2a5298; color:white; border-radius:5px; padding:2px 8px; font-size:.8rem; }
  </style>
</head>
<body>

<div class="hero">
  <h1>📊 Excel 변환기</h1>
  <p>Before 형식 파일을 업로드하면 After 형식으로 자동 변환합니다</p>
</div>

<div class="container">
  <div class="card card-main">
    <div class="card-body">

      <div id="drop-zone" onclick="document.getElementById('file-input').click()">
        <div class="icon">📁</div>
        <p><strong>클릭</strong>하거나 파일을 여기에 <strong>드래그 &amp; 드롭</strong></p>
        <p style="font-size:.82rem;margin-top:6px;">.xlsx 형식만 지원</p>
      </div>
      <input type="file" id="file-input" accept=".xlsx">
      <div id="file-name">선택된 파일 없음</div>

      <button class="btn-convert" id="btn-convert" disabled onclick="doConvert()">
        변환 후 다운로드
      </button>

      <div id="alert-box" class="alert" role="alert"></div>

      <div class="rules">
        <h6>🔄 적용되는 변환 규칙</h6>
        <div class="mb-2">
          <span class="badge-rule">날짜 형식</span>
          &nbsp;DELI_DATE / CFMD_DATE :
          <code>2026-03-24 00:00:00</code> → <code>2026. 3. 24</code>
        </div>
        <div>
          <span class="badge-rule">행 분리</span>
          &nbsp;Pallet No, BATCH PROPOSAL, MOLD REQUEST, Prod Date 컬럼에서
          <code>;</code>로 구분된 복수 값 → 각 값을 별도 행으로 분리
        </div>
      </div>

    </div>
  </div>
</div>

<script>
const fileInput  = document.getElementById('file-input');
const fileNameEl = document.getElementById('file-name');
const btnConvert = document.getElementById('btn-convert');
const dropZone   = document.getElementById('drop-zone');
const alertBox   = document.getElementById('alert-box');

fileInput.addEventListener('change', () => {
  if (fileInput.files.length) {
    fileNameEl.textContent = '✅  ' + fileInput.files[0].name;
    btnConvert.disabled = false;
  }
});

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('hover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('hover'));
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.classList.remove('hover');
  if (e.dataTransfer.files.length) {
    const dt = new DataTransfer();
    dt.items.add(e.dataTransfer.files[0]);
    fileInput.files = dt.files;
    fileNameEl.textContent = '✅  ' + e.dataTransfer.files[0].name;
    btnConvert.disabled = false;
  }
});

function doConvert() {
  if (!fileInput.files.length) return;
  btnConvert.disabled = true;
  btnConvert.innerHTML =
    '<span class="spinner-border spinner-border-sm" role="status"></span> 변환 중...';

  const fd = new FormData();
  fd.append('file', fileInput.files[0]);

  fetch('/convert', { method: 'POST', body: fd })
    .then(res => {
      if (!res.ok) return res.text().then(t => { throw new Error(t); });
      return res.blob().then(blob => {
        const cd  = res.headers.get('Content-Disposition') || '';
        const m   = cd.match(/filename[^=]*=([^;\\n]+)/i);
        const fn  = m ? m[1].trim().replace(/"/g,'') : 'converted.xlsx';
        const a   = document.createElement('a');
        a.href    = URL.createObjectURL(blob);
        a.download = fn;
        a.click();
        showAlert('success', '✅ 변환 완료! "' + fn + '" 파일이 다운로드 되었습니다.');
      });
    })
    .catch(err => showAlert('danger', '❌ 오류: ' + err.message))
    .finally(() => {
      btnConvert.disabled = false;
      btnConvert.innerHTML = '변환 후 다운로드';
    });
}

function showAlert(type, msg) {
  alertBox.className = 'alert alert-' + type;
  alertBox.textContent = msg;
  alertBox.style.display = 'block';
}
</script>
</body>
</html>"""


# ─────────────────────────────────────────────
#  HTTP 핸들러
# ─────────────────────────────────────────────
class Handler(http.server.BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):
        print(f"[{self.address_string()}] {fmt % args}")

    def do_GET(self):
        if self.path in ('/', '/index.html'):
            body = HTML_PAGE.encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path != '/convert':
            self.send_response(404)
            self.end_headers()
            return

        ctype, pdict = cgi.parse_header(self.headers.get('Content-Type', ''))
        if ctype != 'multipart/form-data':
            self._err(400, '잘못된 요청 형식입니다.')
            return

        pdict['boundary'] = bytes(pdict['boundary'], 'utf-8')
        pdict['CONTENT-LENGTH'] = int(self.headers.get('Content-Length', 0))

        fields = cgi.parse_multipart(self.rfile, pdict)
        file_data = fields.get('file')
        if not file_data:
            self._err(400, '파일이 없습니다.')
            return

        raw = file_data[0] if isinstance(file_data[0], bytes) else file_data[0].encode()

        # 원본 파일명 추출
        orig_name = 'upload'
        cd_header = self.headers.get('Content-Disposition', '')
        # multipart 내부에서 파일명 추출은 cgi.FieldStorage 방식으로 별도 처리
        # 간단하게 고정 이름 사용
        orig_name = 'converted'

        # Content-Disposition 에서 파일명 추출 시도
        for part in self.headers.get('Content-Type', '').split(';'):
            part = part.strip()

        try:
            out_bytes  = transform(raw)
            out_name   = f"{orig_name}.xlsx"
            self.send_response(200)
            self.send_header('Content-Type',
                             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', f'attachment; filename="{out_name}"')
            self.send_header('Content-Length', str(len(out_bytes)))
            self.end_headers()
            self.wfile.write(out_bytes)
        except Exception as e:
            self._err(500, f'변환 오류: {e}')

    def _err(self, code, msg):
        body = msg.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'text/plain; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)


# ─────────────────────────────────────────────
#  진입점
# ─────────────────────────────────────────────
if __name__ == '__main__':
    import webbrowser, threading

    PORT = 5000
    server = http.server.HTTPServer(('', PORT), Handler)

    print("=" * 52)
    print("  📊 Excel 변환기 서버 시작")
    print(f"  브라우저: http://localhost:{PORT}")
    print("  종료: Ctrl+C")
    print("=" * 52)

    # 1초 후 자동으로 브라우저 열기
    threading.Timer(1.0, lambda: webbrowser.open(f'http://localhost:{PORT}')).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n서버 종료.")
        server.shutdown()
