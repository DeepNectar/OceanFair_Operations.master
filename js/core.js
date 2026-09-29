/* ============================================================
   core.js — shared engine for ALL tabs (common CSS/JS/HTML)
   Replaces: Application objects, MsgBox/InputBox, Worksheets,
   cell formatting, persistence via localStorage.
   ============================================================ */
var OF = {};

/* unified boot: runs once the DOM is ready; delegates to the
   Workbook_Open port (app.js) and the access gate (access-control.js) */
document.addEventListener('DOMContentLoaded', function () {
  if (window.OF_APP_BOOT) OF_APP_BOOT();
});

/* ---------- storage helpers ---------- */
OF.load = function (key, fallback) {
  try { var v = localStorage.getItem('of_' + key); return v ? JSON.parse(v) : fallback; }
  catch (e) { return fallback; }
};
OF.save = function (key, val) { localStorage.setItem('of_' + key, JSON.stringify(val)); };

/* ---------- MsgBox replacement ---------- */
OF.msg = function (text, kind, title) {
  var stack = document.getElementById('toast-stack');
  var t = document.createElement('div');
  t.className = 'toast ' + (kind || 'info');
  t.textContent = (title ? '[' + title + '] ' : '') + text;
  stack.appendChild(t);
  setTimeout(function () { t.remove(); }, 4200);
};
OF.info    = function (m, title) { OF.msg(m, 'info',  title); };
OF.warn    = function (m, title) { OF.msg(m, 'warn',  title); };
OF.error   = function (m, title) { OF.msg(m, 'error', title); };

/* ---------- InputBox replacement (Promise) ---------- */
OF.inputBox = function (prompt, title, def) {
  return new Promise(function (resolve) {
    var bd = document.getElementById('modal-backdrop');
    document.getElementById('modal-title').textContent = title || 'Input';
    document.getElementById('modal-body').textContent = prompt || '';
    var wrap = document.getElementById('modal-input-wrap');
    var inp = document.getElementById('modal-input');
    wrap.classList.remove('hidden');
    inp.value = def || '';
    var acts = document.getElementById('modal-actions');
    acts.innerHTML = '';
    function done(v) { bd.classList.add('hidden'); resolve(v); }
    var ok = document.createElement('button'); ok.className = 'btn btn-primary'; ok.textContent = 'OK';
    var cancel = document.createElement('button'); cancel.className = 'btn'; cancel.textContent = 'Cancel';
    ok.onclick = function () { done(inp.value); };
    cancel.onclick = function () { done(null); }; // VBA: Cancel returns ""
    acts.append(ok, cancel);
    bd.classList.remove('hidden');
    inp.focus();
    inp.onkeydown = function (e) { if (e.key === 'Enter') done(inp.value); };
  });
};

/* ---------- Yes/No/Ok confirmation (MsgBox vbYesNo…) ---------- */
OF.confirmBox = function (question, title, buttons) {
  // buttons: 'ok' | 'yesno' | 'yesnocancel'  -> resolves 'OK'/'Yes'/'No'/'Cancel'
  return new Promise(function (resolve) {
    var bd = document.getElementById('modal-backdrop');
    document.getElementById('modal-title').textContent = title || 'Confirm';
    document.getElementById('modal-body').textContent = question || '';
    document.getElementById('modal-input-wrap').classList.add('hidden');
    var acts = document.getElementById('modal-actions');
    acts.innerHTML = '';
    var opts = buttons === 'yesnocancel' ? ['Yes', 'No', 'Cancel']
             : buttons === 'yesno' ? ['Yes', 'No'] : ['OK'];
    opts.forEach(function (label) {
      var b = document.createElement('button');
      b.className = 'btn ' + (label === 'Yes' || label === 'OK' ? 'btn-primary' : '');
      b.textContent = label;
      b.onclick = function () { bd.classList.add('hidden'); resolve(label); };
      acts.appendChild(b);
    });
    bd.classList.remove('hidden');
  });
};

/* ---------- Sheet grid builder (editable spreadsheet cells) ---------- */
/*  spec = { id, columns:[{name, align}], rows: [[..]], maxRows, onChange(r,c), isLocked(r,c) } */
OF.buildGrid = function (spec) {
  var table = document.getElementById(spec.id);
  // keep runtime edits (e.g. port dropdowns added by addPortDropdown) in sync
  // with spec.rows BEFORE rebuilding, otherwise the rebuild wipes unsaved DOM state.
  // rowsIsFresh: caller already synced spec.rows with the current DOM.
  if (table && table._spec === spec && !spec.rowsIsFresh) {
    var dom = OF.readGrid(spec.id);
    while (dom.length < spec.rows.length) dom.push(spec.rows[dom.length].map(function () { return ''; }));
    spec.rows = dom;
  }
  var html = '<thead><tr><th class="rownum">#</th>';
  spec.columns.forEach(function (col) { html += '<th>' + col.name + '</th>'; });
  html += '</tr></thead><tbody>';
  var data = spec.rows;
  for (var r = 0; r < Math.max(spec.maxRows || 30, data.length); r++) {
    html += '<tr><td class="rownum">' + (r + 1) + '</td>';
    for (var c = 0; c < spec.columns.length; c++) {
      var val = (data[r] && data[r][c] != null) ? String(data[r][c]) : '';
      var cls = (spec.columns[c].align === 'left' ? 'left' : 'center') + ' tnr';
      var locked = spec.isLocked && spec.isLocked(r, c);
      html += '<td class="' + cls + '" data-r="' + r + '" data-c="' + c + '"' +
              (locked ? '' : ' contenteditable="true"') + '>' + OF.esc(val) + '</td>';
    }
    html += '</tr>';
  }
  html += '</tbody>';
  table.innerHTML = html;
  table._spec = spec;

  // wire change events (Worksheet_Change equivalent)
  table.querySelectorAll('td[contenteditable]').forEach(function (td) {
    td.addEventListener('blur', function () {
      var d = OF.readGrid(spec.id);
      OF.save(spec.id, d);
      if (spec.onChange) spec.onChange(d, +td.dataset.r, +td.dataset.c);
    });
    td.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); td.blur(); }
    });
  });
  return table;
};

OF.readGrid = function (id) {
  var table = document.getElementById(id);
  var rows = [];
  table.querySelectorAll('tbody tr').forEach(function (tr) {
    var cells = [];
    tr.querySelectorAll('td:not(.rownum)').forEach(function (td) {
      cells.push(td.textContent.trim());
    });
    rows.push(cells);
  });
  return rows;
};

OF.writeCell = function (id, r, c, val, cls) {
  var table = document.getElementById(id);
  var trs = table.querySelectorAll('tbody tr');
  if (!trs[r]) return;
  var tds = trs[r].querySelectorAll('td:not(.rownum)');
  if (!tds[c]) return;
  tds[c].textContent = val == null ? '' : String(val);
  if (cls !== undefined) { tds[c].className = cls; }
};

OF.gridDataOr = function (id, fallback) {
  var saved = OF.load(id, null);
  return saved && saved.length ? saved : fallback;
};

OF.esc = function (s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
};

/* ---------- mailto: composer (replaces Outlook CreateItem) ---------- */
OF.openMail = function (opts) {
  // opts: to, cc, subject, htmlBody  — body sent as plain-text summary + HTML preview tab
  var plain = OF.htmlToText(opts.htmlBody);
  var href = 'mailto:' + encodeURIComponent(opts.to || '') +
    '?subject=' + encodeURIComponent(opts.subject || '') +
    (opts.cc ? '&cc=' + encodeURIComponent(opts.cc) : '') +
    '&body=' + encodeURIComponent(plain.substring(0, 1800));
  // Also open a full HTML preview window (Outlook .Display replacement)
  var w = window.open('', '_blank');
  if (w) { w.document.write(opts.htmlBody); w.document.close(); }
  location.href = href;
  OF.info('Email draft opened in your mail app + HTML preview tab.', 'Outlook');
};

OF.htmlToText = function (html) {
  var d = document.createElement('div');
  d.innerHTML = html;
  return (d.innerText || d.textContent || '').replace(/\n{3,}/g, '\n\n');
};

/* ---------- date helpers (VBA Format / ordinals) ---------- */
OF.monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
OF.weekdays = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
OF.formatDateLong = function (dt) { // Format(Date,"dddd, mmmm d, yyyy")
  return OF.weekdays[dt.getDay()] + ', ' + OF.monthNames[dt.getMonth()] + ' ' + dt.getDate() + ', ' + dt.getFullYear();
};
OF.formatDateDDMMMYYYY = function (dt) {
  var m = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  return String(dt.getDate()).padStart(2,'0') + '-' + m[dt.getMonth()] + '-' + dt.getFullYear();
};
OF.FormatDateWithOrdinal = function (dt) { // REORDER.bas function port
  var day = dt.getDate(), suffix;
  if ([1,21,31].indexOf(day) >= 0) suffix = 'st';
  else if ([2,22].indexOf(day) >= 0) suffix = 'nd';
  else if ([3,23].indexOf(day) >= 0) suffix = 'rd';
  else suffix = 'th';
  return day + suffix + ' ' + OF.monthNames[dt.getMonth()] + ', ' + dt.getFullYear();
};
OF.isValidEmail = function (email) { // IsValidEmail port
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(email).trim());
};
OF.cleanFileName = function (name) { // CleanFileName port
  return String(name).replace(/[\\/:*?"<>|]/g, '_');
};

/* ---------- Excel import / export helpers (Import & Export buttons) ---------- */
OF.parseDelimited = function (text) {
  text = String(text).replace(/^\uFEFF/, '');
  var delim = text.indexOf('\t') > -1 ? '\t' : ',';
  var lines = text.split(/\r?\n/);
  if (lines.length && lines[lines.length - 1].trim() === '') lines.pop();
  return lines.map(function (l) { return l.split(delim); });
};

OF.csvEscape = function (v) {
  v = String(v == null ? '' : v);
  return /[",\r\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
};

OF.rowsToCsv = function (rows) {
  return rows.map(function (r) { return r.map(OF.csvEscape).join(','); }).join('\r\n');
};

OF.downloadCsv = function (fileName, csvText) {
  var blob = new Blob(['\uFEFF' + csvText], { type: 'text/csv;charset=utf-8;' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 0);
};

// Minimal XLSX reader (no external libs): unzips via DecompressionStream and
// pulls cell values out of the first worksheet's XML (shared strings inlined).
OF.readXlsxAsync = function (file) {
  if (typeof DecompressionStream === 'undefined') {
    return Promise.reject(new Error('This browser cannot read .xlsx directly — please save the sheet as .csv and import that.'));
  }
  return file.arrayBuffer().then(function (buf) { return OF.unzipEntries(buf); })
    .then(function (entries) {
      var names = Object.keys(entries).filter(function (n) { return /^xl\/worksheets\/sheet\d+\.xml$/.test(n); })
        .sort(function (a, b) { return a.localeCompare(b, undefined, { numeric: true }); });
      if (!names.length) throw new Error('No worksheet found inside the .xlsx file.');
      var shared = [];
      if (entries['xl/sharedStrings.xml']) {
        var sDoc = new DOMParser().parseFromString(entries['xl/sharedStrings.xml'], 'application/xml');
        Array.prototype.forEach.call(sDoc.querySelectorAll('si'), function (si) {
          shared.push(si.textContent || '');
        });
      }
      var wDoc = new DOMParser().parseFromString(entries[names[0]], 'application/xml');
      var rowsOut = [];
      Array.prototype.forEach.call(wDoc.querySelectorAll('row'), function (rowEl) {
        var arr = [];
        Array.prototype.forEach.call(rowEl.querySelectorAll('c'), function (c) {
          var ref = c.getAttribute('r') || '';
          var colIdx = 0, ch;
          for (var i = 0; i < ref.length; i++) {
            ch = ref.charCodeAt(i);
            if (ch >= 65 && ch <= 90) colIdx = colIdx * 26 + (ch - 64);
            else break;
          }
          var col = Math.max(colIdx - 1, 0);
          var t = c.getAttribute('t');
          var vEl = c.getElementsByTagName('v')[0];
          var isEl = c.getElementsByTagName('is')[0];
          var val = '';
          if (t === 's' && vEl) val = shared[parseInt(vEl.textContent, 10)] || '';
          else if (t === 'inlineStr' && isEl) val = isEl.textContent || '';
          else if (vEl) val = vEl.textContent || '';
          while (arr.length < col) arr.push('');
          arr[col] = String(val).trim();
        });
        rowsOut.push(arr);
      });
      return rowsOut;
    });
};

OF.unzipEntries = function (arrayBuffer) {
  var view = new DataView(arrayBuffer);
  // locate End Of Central Directory record
  var eocd = -1;
  for (var p = view.byteLength - 22; p >= 0 && p >= view.byteLength - 66000; p--) {
    if (view.getUint32(p, true) === 0x06054b50) { eocd = p; break; }
  }
  if (eocd < 0) throw new Error('Not a valid .xlsx (zip) file.');
  var cdOffset = view.getUint32(eocd + 16, true);
  var cdCount = view.getUint16(eocd + 10, true);
  var tasks = [];
  var ptr = cdOffset;
  for (var n = 0; n < cdCount; n++) {
    if (view.getUint32(ptr, true) !== 0x02014b50) break;
    var method = view.getUint16(ptr + 10, true);
    var compSize = view.getUint32(ptr + 20, true);
    var nameLen = view.getUint16(ptr + 28, true);
    var extraLen = view.getUint16(ptr + 30, true);
    var commentLen = view.getUint16(ptr + 32, true);
    var localOff = view.getUint32(ptr + 42, true);
    var name = new TextDecoder().decode(new Uint8Array(arrayBuffer, ptr + 46, nameLen));
    // parse local header to find data start
    var lhNameLen = view.getUint16(localOff + 26, true);
    var lhExtraLen = view.getUint16(localOff + 28, true);
    var dataStart = localOff + 30 + lhNameLen + lhExtraLen;
    var data = new Uint8Array(arrayBuffer, dataStart, compSize);
    if (method === 0) {
      tasks.push(Promise.resolve({ name: name, bytes: data }));
    } else {
      tasks.push(OF.inflateRaw(data).then(function (bytes) { return { name: name, bytes: bytes }; }));
    }
    ptr += 46 + nameLen + extraLen + commentLen;
  }
  return Promise.all(tasks).then(function (parts) {
    var entries = {};
    parts.forEach(function (pt) {
      if (pt && pt.name && typeof pt.bytes !== 'undefined') {
        entries[pt.name] = new TextDecoder().decode(pt.bytes);
      }
    });
    return entries;
  });
};

OF.inflateRaw = function (data) {
  var ds = null;
  try { ds = new DecompressionStream('deflate-raw'); } catch (e) { /* older browsers */ }
  if (!ds) return Promise.reject(new Error('Browser does not support deflate-raw decompression.'));
  return new Response(new Blob([data]).stream().pipeThrough(ds)).arrayBuffer();
};

/* Generic per-sheet Import / Export wiring.
   spec: { gridId, sheetName, maxRows, onChange(gridId), onAfterLoad(gridId) } */
OF.bindSheetIO = function (spec) {
  var imp = document.getElementById('import-' + spec.gridId);
  var inp = document.getElementById('import-file-' + spec.gridId);
  var exp = document.getElementById('export-' + spec.gridId);
  if (exp) exp.addEventListener('click', function () {
    var rows = OF.readGrid(spec.gridId);
    if (!rows.length) { OF.warn('Nothing to export.', spec.sheetName); return; }
    OF.downloadCsv(OF.cleanFileName(spec.sheetName) + '.csv', OF.rowsToCsv(rows));
    OF.info('Exported "' + OF.cleanFileName(spec.sheetName) + '.csv" — open it in Excel.', 'Export');
  });
  if (imp && inp) {
    imp.addEventListener('click', function () { inp.value = ''; inp.click(); });
    inp.addEventListener('change', function () {
      if (!inp.files.length) return;
      var file = inp.files[0];
      var done = function (rows) {
        if (!rows || !rows.length) { OF.warn('No data found in the selected file.', 'Import'); return; }
        var width = (document.getElementById(spec.gridId)._spec || { columns: [] }).columns.length || 5;
        var norm = rows.slice(0, spec.maxRows || 1000).map(function (r) {
          var out = [];
          for (var c = 0; c < width; c++) out.push(String(r[c] == null ? '' : r[c]));
          return out;
        });
        OF.save(spec.gridId, norm);
        OF.rebuildFromSave(spec.gridId);
        if (spec.onChange) spec.onChange(spec.gridId);
        if (spec.onAfterLoad) spec.onAfterLoad(spec.gridId);
        OF.info('Imported ' + norm.length + ' row(s) from ' + file.name + ' into "' + spec.sheetName + '".', 'Import');
      };
      if (/\.xlsx$/i.test(file.name)) {
        OF.readXlsxAsync(file).then(done)['catch'](function (e) {
          OF.error('Could not read the .xlsx file: ' + e.message, 'Import');
        });
      } else {
        var reader = new FileReader();
        reader.onload = function () { done(OF.parseDelimited(reader.result)); };
        reader.onerror = function () { OF.error('Could not read the file.', 'Import'); };
        reader.readAsText(file);
      }
    });
  }
};
