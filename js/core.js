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
