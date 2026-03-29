// ── File pickers (File System Access API) ───────────────────────────────────

let wordFileHandle    = null;   // FileSystemFileHandle for the Word file
let pickedImageFiles  = [];     // File[] from the image folder

async function pickWordFile() {
  let handles;
  try {
    handles = await window.showOpenFilePicker({
      multiple: false,
      types: [{ description: 'Word 文件', accept: { 'application/octet-stream': ['.doc', '.docx'] } }]
    });
  } catch (e) {
    if (e.name !== 'AbortError') showError('無法開啟檔案：' + e.message);
    return;
  }
  wordFileHandle = handles[0];
  const file = await wordFileHandle.getFile();
  document.getElementById('wordName').textContent = file.name;
  resetResult();
  inspectDoc(file);
}

async function pickImageFolder() {
  let dirHandle;
  try {
    dirHandle = await window.showDirectoryPicker({
      mode: 'read',
      startIn: wordFileHandle ?? 'pictures'   // ← 預設與 Word 檔案同一資料夾
    });
  } catch (e) {
    if (e.name !== 'AbortError') showError('無法開啟資料夾：' + e.message);
    return;
  }

  const imgExt = /\.(png|jpe?g|gif|bmp|tiff?|webp)$/i;
  const files = [];
  for await (const [, handle] of dirHandle) {
    if (handle.kind === 'file') {
      if (imgExt.test(handle.name)) files.push(await handle.getFile());
    }
  }
  files.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));

  pickedImageFiles = files;
  document.getElementById('imgName').textContent = files.length > 0
    ? `${dirHandle.name}（${files.length} 張圖片）`
    : `${dirHandle.name}（無圖片）`;
  resetResult();
}

// ── Date → ROC preview ──────────────────────────────────────────────────────

document.getElementById('dateInput').addEventListener('change', function () {
  document.getElementById('rocPreview').textContent = toRoc(this.value);
  resetResult();
});

document.getElementById('padZeroCheck').addEventListener('change', function () {
  const dateVal = document.getElementById('dateInput').value;
  document.getElementById('rocPreview').textContent = toRoc(dateVal);
});

function toRoc(isoDate) {
  if (!isoDate) return '—';
  const [y, m, d] = isoDate.split('-');
  const roc = parseInt(y, 10) - 1911;
  const pad = document.getElementById('padZeroCheck').checked;
  const mm = pad ? m : String(parseInt(m, 10));
  const dd = pad ? d : String(parseInt(d, 10));
  return `民國 ${roc}.${mm}.${dd}`;
}

// ── Set today's date as default ─────────────────────────────────────────────

(function () {
  const today = new Date();
  const iso = today.toISOString().slice(0, 10);
  const input = document.getElementById('dateInput');
  input.value = iso;
  document.getElementById('rocPreview').textContent = toRoc(iso);
})();

// ── Reset result area ───────────────────────────────────────────────────────

function resetResult() {
  setVisible('successMsg', false);
  setVisible('errorMsg', false);
  setVisible('status', false);
}

function setVisible(id, visible) {
  const el = document.getElementById(id);
  if (visible) el.classList.add('visible');
  else el.classList.remove('visible');
}

function setProgress(pct, label) {
  document.getElementById('progressBar').style.width = pct + '%';
  document.getElementById('progressLabel').textContent = label;
}

// ── Process ─────────────────────────────────────────────────────────────────

function processDoc() {
  const dateInput = document.getElementById('dateInput');

  if (!wordFileHandle)  { showError('請選擇 Word 檔案'); return; }
  if (!dateInput.value) { showError('請選擇日期'); return; }

  const padZero = document.getElementById('padZeroCheck').checked;
  // getFile() re-reads the current bytes (handles file-changed-on-disk case)
  wordFileHandle.getFile().then(wordFile => _doProcess(wordFile, dateInput.value, padZero));
}

function _doProcess(wordFile, dateValue, padZero) {
  const formData = new FormData();
  formData.append('wordFile', wordFile);
  formData.append('date', dateValue);
  formData.append('padZero', padZero ? 'true' : 'false');

  pickedImageFiles.forEach(f => formData.append('images', f));

  // Show progress area
  resetResult();
  document.getElementById('processBtn').disabled = true;
  document.getElementById('spinner').style.display = 'none';
  setProgress(0, '準備上傳…');
  setVisible('status', true);

  const xhr = new XMLHttpRequest();

  // ── Upload progress (0 → 100%)
  xhr.upload.addEventListener('progress', e => {
    if (!e.lengthComputable) return;
    const pct = Math.round(e.loaded / e.total * 100);
    setProgress(pct, `上傳中… ${pct}%`);
  });

  // ── Upload done → switch to processing spinner
  xhr.upload.addEventListener('load', () => {
    setProgress(100, '處理中…');
    document.getElementById('spinner').style.display = 'block';
  });

  // ── Response received
  xhr.addEventListener('load', () => {
    document.getElementById('spinner').style.display = 'none';
    setVisible('status', false);
    document.getElementById('processBtn').disabled = false;

    if (xhr.status !== 200) {
      // responseType is 'arraybuffer', so decode the error body manually
      let msg = `處理失敗（HTTP ${xhr.status}）`;
      try {
        const text = new TextDecoder().decode(new Uint8Array(xhr.response));
        const json = JSON.parse(text);
        msg = json.error || json.title || text || msg;
      } catch {}
      showError(msg);
      return;
    }

    // Build download from blob
    const blob = new Blob([xhr.response], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);

    const disposition = xhr.getResponseHeader('Content-Disposition') || '';
    const nameMatch = disposition.match(/filename\*?=(?:UTF-8'')?["']?([^"';\n]+)/i);
    const filename = nameMatch ? decodeURIComponent(nameMatch[1]) : 'output.doc';

    const link = document.getElementById('downloadLink');
    link.href = url;
    link.download = filename;

    document.getElementById('successSub').textContent = `檔案名稱：${filename}`;

    // Show process log
    const logEl = document.getElementById('processLog');
    logEl.innerHTML = '';
    try {
      const raw = xhr.getResponseHeader('X-Process-Log');
      if (raw) {
        const entries = JSON.parse(decodeURIComponent(raw));
        entries.forEach(msg => {
          const li = document.createElement('li');
          li.textContent = msg;
          li.className = msg.startsWith('✅') ? 'log-ok' : 'log-warn';
          logEl.appendChild(li);
        });
      }
    } catch {}

    setVisible('successMsg', true);
  });

  xhr.addEventListener('error', () => {
    document.getElementById('spinner').style.display = 'none';
    setVisible('status', false);
    document.getElementById('processBtn').disabled = false;
    showError('網路錯誤，請確認程式是否正常執行');
  });

  xhr.open('POST', '/api/process');
  xhr.responseType = 'arraybuffer';
  xhr.send(formData);
}

function showError(msg) {
  const el = document.getElementById('errorMsg');
  el.textContent = `⚠ ${msg}`;
  setVisible('errorMsg', true);
}

// ── Document Inspector ───────────────────────────────────────────────────────

function hideInspect() {
  document.getElementById('inspectPanel').style.display = 'none';
  document.getElementById('inspectBody').innerHTML = '';
}

async function inspectDoc(file) {
  const panel   = document.getElementById('inspectPanel');
  const body    = document.getElementById('inspectBody');
  const spinner = document.getElementById('inspectSpinner');

  body.innerHTML = '';
  panel.style.display = 'block';
  spinner.style.display = 'inline-block';

  const fd = new FormData();
  fd.append('wordFile', file);

  let data;
  try {
    const res = await fetch('/api/inspect', { method: 'POST', body: fd });
    data = await res.json();
  } catch {
    body.innerHTML = '<p class="inspect-err">掃描失敗</p>';
    spinner.style.display = 'none';
    return;
  }
  spinner.style.display = 'none';

  if (data.isLegacy) {
    body.innerHTML = '<p class="inspect-note">.doc 格式需要轉換後才能掃描（按「開始處理」後可在處理紀錄查看）</p>';
    return;
  }

  const sections = [];

  if (data.floatingShapes?.length) {
    sections.push(renderSection('浮動圖形（Selection Pane 中的物件）', data.floatingShapes.map(s => {
      const badge = `<span class="ibadge">${s.type}</span>`;
      const name  = `<span class="iname">${s.name}</span>`;
      const text  = s.text ? ` <span class="itext">「${s.text}」</span>` : '';
      return badge + name + text;
    })));
  }

  if (data.inlineShapes?.length) {
    sections.push(renderSection('內嵌圖片（嵌在文字流中）', data.inlineShapes.map(s =>
      `<span class="ibadge">${s.type}</span><span class="iname">${s.name}</span>`
    )));
  }

  if (data.textBoxes?.length) {
    sections.push(renderSection('VML 文字方塊', data.textBoxes.map(s =>
      `<span class="ibadge">${s.type}</span><span class="iname">${s.name}</span>` +
      (s.text ? ` <span class="itext">「${s.text}」</span>` : '')
    )));
  }

  if (data.bodyParagraphs?.length) {
    sections.push(renderSection('段落文字（前 20 段）',
      data.bodyParagraphs.slice(0, 20).map(t =>
        `<span class="ipara">${escHtml(t)}</span>`
      )
    ));
  }

  if (!sections.length) {
    body.innerHTML = '<p class="inspect-note">未偵測到任何內容</p>';
  } else {
    body.innerHTML = sections.join('');
  }
}

function renderSection(title, items) {
  return `<div class="isection">
    <div class="isection-title">${title}</div>
    <ul class="ilist">${items.map(i => `<li>${i}</li>`).join('')}</ul>
  </div>`;
}

function escHtml(s) {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
