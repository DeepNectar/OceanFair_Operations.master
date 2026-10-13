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

/* ---------- email composer (replaces Outlook CreateItem / .Display) ----------
   mailto: can NOT carry an HTML body — clients strip tags and drop <style>/<table>
   formatting. So every generated HTML body is normalised here into ONE consistent,
   email-client-safe structure that is used by all tabs/buttons:
     • tableless plain-text version  -> mailto: draft body
     • HTML version                  -> preview window + "Copy HTML" for the mail body
   Every OF.openMail() call passes its htmlBody through OF.buildEmailHtml(), so the
   tables and text of all emails look the same and render correctly.                */

/* ---- CSS subset we can resolve ourselves (mail clients ignore <style> blocks,
        therefore every rule is flattened onto inline style attributes) ---- */
OF.parseCssRules = function (cssText) {
  var rules = [];
  cssText.replace(/\/\*[\s\S]*?\*\//g, '').split('}').forEach(function (chunk) {
    var parts = chunk.split('{');
    if (parts.length < 2) return;
    var selector = parts[0].replace(/\s+/g, ' ').trim();
    if (!selector) return;
    var decls = {};
    parts.slice(1).join('{').split(';').forEach(function (d) {
      var i = d.indexOf(':');
      if (i > 0) decls[d.slice(0, i).trim().toLowerCase()] = d.slice(i + 1).trim();
    });
    selector.split(',').forEach(function (sel) {
      sel = sel.trim();
      if (sel) rules.push({ sel: sel, simple: /^[a-zA-Z]+|\.[\w-]+$/.test(sel), decls: decls });
    });
  });
  return rules;
};

OF.styleFromDecls = function (decls) {
  var out = '';
  Object.keys(decls).forEach(function (prop) {
    out += prop + ':' + decls[prop] + ';';
  });
  return out;
};

OF.matchesSimpleSelector = function (el, sel) {
  if (sel.charAt(0) === '.') {
    return (' ' + (el.getAttribute('class') || '') + ' ')
      .indexOf(' ' + sel.slice(1) + ' ') >= 0;
  }
  return el.tagName.toLowerCase() === sel.toLowerCase();
};

/* Inline <style>/<class> rules are ignored by most mail clients, therefore all
   styles are flattened onto each element (inline style attributes). */
OF.inlineStyles = function (html) {
  var doc = new DOMParser().parseFromString('<!DOCTYPE html><html><head></head><body>' + html + '</body></html>', 'text/html');

  /* collect every stylesheet rule found in the original markup */
  var rules = [];
  var styleEls = doc.getElementsByTagName('style');
  for (var s = 0; s < styleEls.length; s++) rules = rules.concat(OF.parseCssRules(styleEls[s].textContent));

  /* copy the original font declarations down to every text element, because mail
     clients do not inherit from a normalised <body>. */
  var fontCss = OF.extractFontRules(html);

  /* keep whatever the ported VBA markup asked for, but make it mail-safe */
  var defaults = {
    'h1': fontCss + 'font-size:22px;line-height:1.3;margin:0 0 14px;color:#8B0000;font-weight:bold;',
    'h2': fontCss + 'font-size:18px;line-height:1.3;margin:0 0 12px;color:#2c3e50;font-weight:bold;',
    'h3': fontCss + 'font-size:16px;line-height:1.3;margin:0 0 12px;color:#2c3e50;font-weight:bold;',
    'p':  fontCss + 'margin:0 0 12px;line-height:1.5;',
    'li': fontCss + 'margin:0 0 6px;line-height:1.5;',
    'span': fontCss,
    'strong': fontCss,
    'em': fontCss,
    'b': fontCss,
    'i': fontCss,
    'small': fontCss,
    'div': fontCss + 'margin:0 0 12px;',
    'a':  fontCss + 'color:#0563C1;text-decoration:underline;',
    'img': 'max-width:100%;height:auto;border:0;display:block;',
    'blockquote': fontCss + 'margin:0 0 12px;padding:8px 14px;border-left:4px solid #DDDDDD;color:#444444;'
  };

  /* pass 1 — author CSS rules (later rules win), then tag defaults as a fallback */
  var all = doc.getElementsByTagName('*');
  for (var n = 0; n < all.length; n++) {
    var el = all[n], tag = el.tagName.toLowerCase();
    if (tag === 'br' || tag === 'hr' || tag === 'style' || tag === 'script') continue;
    var css = '';
    rules.forEach(function (r) {
      if (r.simple && OF.matchesSimpleSelector(el, r.sel)) css += OF.styleFromDecls(r.decls);
    });
    if (defaults[tag]) css += defaults[tag];
    if (css) OF.mergeStyle(el, css);
  }

  /* pass 2 — tables: fixed layout, one hairline border per cell, no stray gaps */
  var tables = doc.getElementsByTagName('table');
  for (var t = 0; t < tables.length; t++) {
    var tb = tables[t];
    tb.setAttribute('cellspacing', '0');
    tb.setAttribute('cellpadding', '0');
    tb.removeAttribute('border');
    tb.removeAttribute('width');
    OF.mergeStyle(tb, 'border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt;' +
      "width:100%;max-width:900px;font-size:13px;color:#202020;background-color:#ffffff;");

    var rows = tb.getElementsByTagName('tr');
    for (var r2 = 0; r2 < rows.length; r2++) {
      var cells = rows[r2].children;
      var rowStyle = (rows[r2].getAttribute('style') || '').toLowerCase();
      var isHeaderRow = /background(-color)?\s*:\s*#4472c4/.test(rowStyle);
      for (var c = 0; c < cells.length; c++) {
        var cell = cells[c];
        var old = (cell.getAttribute('style') || '').toLowerCase();
        var isHead = cell.tagName === 'TH' || isHeaderRow || /background(-color)?\s*:\s*#4472c4/.test(old);
        var align = /text-align\s*:\s*left/.test(old) ? 'left'
                  : /text-align\s*:\s*right/.test(old) ? 'right' : 'center';
        var css2 = fontCss + 'border:1px solid #BFC9D4;padding:8px 10px;text-align:' + align + ';vertical-align:top;';
        if (isHead) css2 += 'background-color:#4472C4;color:#ffffff;font-weight:bold;';
        else if (/background-color\s*:\s*(#e9e9e9|#f2f2f2)/.test(old)) css2 += 'background-color:#F2F4F7;';
        OF.mergeStyle(cell, css2);
      }
    }
  }

  /* drop <style>/<script>/link — they never survive a real mail client */
  ['style', 'script', 'link', 'meta', 'title'].forEach(function (tag) {
    var els = doc.getElementsByTagName(tag);
    while (els.length) els[0].parentNode.removeChild(els[0]);
  });

  /* class/id attributes are dead weight once everything is inlined */
  var cleaned = doc.getElementsByTagName('*');
  for (var k = 0; k < cleaned.length; k++) {
    cleaned[k].removeAttribute('class');
    cleaned[k].removeAttribute('id');
  }

  return doc.body.innerHTML;
};

/* pull the font declarations out of the body tag / <style> blocks of a raw htmlBody
   so the original look (Times New Roman, Calibri, …) survives inline-only rendering. */
OF.extractFontRules = function (html) {
  var src = String(html || '');
  var rules = {};
  var order = [];
  function add(prop, value) {
    prop = prop.toLowerCase();
    if (!/^(font-family|font-size|color|line-height)$/.test(prop)) return;
    if (rules[prop] === undefined) order.push(prop);
    rules[prop] = value.trim();
  }
  /* 1. inline style on <body ...> */
  var m = src.match(/<body[^>]*style\s*=\s*["']([^"']*)["']/i);
  if (m) m[1].split(';').forEach(function (d) {
    var i = d.indexOf(':'); if (i > 0) add(d.slice(0, i), d.slice(i + 1));
  });
  /* 2. simple `body{...}` / `body,p{...}` rules inside <style> blocks */
  var styles = src.match(/<style[^>]*>([\s\S]*?)<\/style>/gi) || [];
  styles.forEach(function (block) {
    var cssText = block.replace(/<[^>]*>/g, '');
    cssText.split('}').forEach(function (chunk) {
      var parts = chunk.split('{');
      if (parts.length < 2) return;
      var selectors = parts[0].replace(/\s+/g, ' ').trim();
      if (!/(^|[,\s])body(\s*,|$)|(^|[,\s])(p|td|th)(\s*,|$)/i.test(selectors)) return;
      parts[1].split(';').forEach(function (d) {
        var i = d.indexOf(':'); if (i > 0) add(d.slice(0, i), d.slice(i + 1));
      });
    });
  });
  /* sensible fallback when the source declared nothing */
  if (rules['font-family'] === undefined) add('font-family', "Arial, Helvetica, sans-serif");
  if (rules['font-size'] === undefined) add('font-size', '14px');
  if (rules.color === undefined) add('color', '#202020');
  return order.map(function (prop) { return prop + ':' + rules[prop] + ';'; }).join('');
};

/* append default rules without overriding the markup's own declarations */
OF.mergeStyle = function (el, css) {
  var existing = (el.getAttribute('style') || '').trim();
  if (!existing) { el.setAttribute('style', css); return; }
  var keep = {};
  existing.split(';').forEach(function (decl) {
    var idx = decl.indexOf(':');
    if (idx > 0) keep[decl.slice(0, idx).trim().toLowerCase()] = true;
  });
  var extra = css.split(';').filter(function (decl) {
    var idx = decl.indexOf(':');
    return idx > 0 && !keep[decl.slice(0, idx).trim().toLowerCase()];
  }).join(';');
  if (extra) el.setAttribute('style', existing.replace(/;?\s*$/, ';') + extra + ';');
};

/* single wrapper shared by every email body in this app */
OF.buildEmailHtml = function (innerHtml) {
  var inner = OF.inlineStyles(innerHtml || '');
  return '<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="utf-8" />\n' +
    '<meta name="viewport" content="width=device-width, initial-scale=1" />\n' +
    '<!--[if mso]><xml><o:OfficeDocumentSettings><o:PixelsPerInch>96</o:PixelsPerInch>' +
    '</o:OfficeDocumentSettings></xml><![endif]-->\n</head>\n<body style="margin:0;padding:0;background-color:#F2F4F7;">\n' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" ' +
    'style="border-collapse:collapse;background-color:#F2F4F7;">\n<tr><td align="center" style="padding:18px 10px;">\n' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" ' +
    'style="border-collapse:collapse;width:100%;max-width:900px;background-color:#ffffff;' +
    'border:1px solid #DDE3EA;">\n<tr><td style="padding:18px 20px;font-family:Arial,Helvetica,sans-serif;' +
    'font-size:14px;line-height:1.5;color:#202020;">\n' +
    inner + '\n</td></tr>\n</table>\n</td></tr>\n</table>\n</body>\n</html>';
};

/* tableless plain-text rendering: keeps column alignment readable everywhere */
OF.htmlToText = function (html) {
  var doc = new DOMParser().parseFromString('<!DOCTYPE html><html><body>' + (html || '') + '</body></html>', 'text/html');
  var out = [];
  var push = function (line) { out.push(line); };

  var walkBlocks = function (node) {
    var kids = node.children;
    if (!kids || !kids.length) {
      var txt = (node.textContent || '').replace(/\s+/g, ' ').trim();
      if (txt) push(txt);
      return;
    }
    for (var i = 0; i < kids.length; i++) {
      var el = kids[i], tag = el.tagName.toLowerCase();
      if (tag === 'table') { push(OF.tableToText(el)); continue; }
      if (tag === 'ul' || tag === 'ol') {
        Array.prototype.forEach.call(el.children, function (li, n) {
          push((tag === 'ol' ? (n + 1) + '. ' : '- ') + (li.textContent || '').replace(/\s+/g, ' ').trim());
        });
        continue;
      }
      if (tag === 'br' || tag === 'hr') { push(''); continue; }
      var hasBlockChild = el.querySelector('table, ul, ol, p, div, h1, h2, h3, tr');
      if (hasBlockChild) { walkBlocks(el); }
      else {
        var t = (el.textContent || '').replace(/\s+/g, ' ').trim();
        if (t) push(t);
      }
    }
  };
  walkBlocks(doc.body);
  return out.join('\n\n').replace(/\n{3,}/g, '\n\n').trim();
};

OF.tableToText = function (table) {
  var rows = [];
  Array.prototype.forEach.call(table.rows, function (tr) {
    var cells = [];
    Array.prototype.forEach.call(tr.cells, function (td) {
      cells.push((td.textContent || '').replace(/\s+/g, ' ').trim());
    });
    if (cells.join('').length) rows.push(cells);
  });
  if (!rows.length) return '';
  var ncols = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
  var widths = [];
  for (var c = 0; c < ncols; c++) {
    widths[c] = rows.reduce(function (m, r) { return Math.max(m, (r[c] || '').length); }, 0);
  }
  var pad = function (s, w) { s = s || ''; while (s.length < w) s += ' '; return s; };
  var lines = rows.map(function (r) {
    var parts = [];
    for (var c = 0; c < r.length; c++) parts.push(pad(r[c], widths[c]));
    return parts.join(' | ').replace(/\s+$/, '');
  });
  var sep = widths.map(function (w) { return new Array(w + 1).join('-'); }).join('-+-');
  return [lines[0], sep].concat(lines.slice(1)).join('\n');
};

OF.copyToClipboard = function (text) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    return navigator.clipboard.writeText(text)
      .then(function () { return true; })
      .catch(function () { return OF.copyToClipboardFallback(text); });
  }
  return Promise.resolve(OF.copyToClipboardFallback(text));
};
OF.copyToClipboardFallback = function (text) {
  try {
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed'; ta.style.top = '-9999px';
    document.body.appendChild(ta); ta.select();
    var ok = document.execCommand('copy');
    document.body.removeChild(ta);
    return Promise.resolve(ok);
  } catch (e) { return Promise.resolve(false); }
};

OF.downloadFile = function (name, content, mime) {
  var a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([content], { type: mime || 'text/html;charset=utf-8' }));
  a.download = name;
  document.body.appendChild(a); a.click();
  setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 0);
};

OF.openMail = function (opts) {
  // opts: to, cc, subject, htmlBody — HTML body is normalised (tables + text) for every client
  var html = OF.buildEmailHtml(opts.htmlBody);
  var plain = OF.htmlToText(html);
  var href = 'mailto:' + encodeURIComponent(opts.to || '') +
    '?subject=' + encodeURIComponent(opts.subject || '') +
    (opts.cc ? '&cc=' + encodeURIComponent(opts.cc) : '') +
    '&body=' + encodeURIComponent(plain.substring(0, 1800));

  // Preview window = what the email really looks like (Outlook .Display replacement),
  // with buttons to copy the HTML straight into the mail body or save it as a file.
  var w = window.open('', '_blank');
  if (w) {
    var bar = '<div id="of-mail-toolbar" style="position:sticky;top:0;z-index:9;background:#1F3B5C;color:#fff;' +
      "padding:10px 14px;font-family:Arial,Helvetica,sans-serif;font-size:13px;" +
      'display:flex;align-items:center;gap:10px;flex-wrap:wrap;">' +
      '<strong>Email preview — ' + OF.esc(opts.subject || '') + '</strong>' +
      '<span style="opacity:.8;">To: ' + OF.esc(opts.to || '') + '</span>' +
      (opts.cc ? '<span style="opacity:.8;">Cc: ' + OF.esc(opts.cc) + '</span>' : '') +
      '<button id="of-copy-html" style="margin-left:auto;cursor:pointer;background:#4472C4;border:0;' +
      'color:#fff;padding:7px 12px;border-radius:4px;font-size:13px;">Copy HTML for email body</button>' +
      '<button id="of-dl-html" style="cursor:pointer;background:#2E7D32;border:0;color:#fff;' +
      'padding:7px 12px;border-radius:4px;font-size:13px;">Save .html</button>' +
      '<button id="of-copy-text" style="cursor:pointer;background:#607D8B;border:0;color:#fff;' +
      'padding:7px 12px;border-radius:4px;font-size:13px;">Copy plain text</button>' +
      '</div>';
    w.document.write(html.replace(/<\/head>/i, '</head>') );
    // inject toolbar at the top of the rendered document
    w.document.close();
    try {
      var host = w.document.body;
      host.insertAdjacentHTML('afterbegin', bar);
      var fileName = (OF.cleanFileName(opts.subject || 'email') || 'email') + '.html';
      w.document.getElementById('of-copy-html').onclick = function () {
        OF.copyToClipboard(html).then(function (ok) {
          w.alert(ok ? 'Email HTML copied — paste it into the body of your email (rich-text / HTML mode).'
                     : 'Copy blocked by the browser — please use "Save .html" instead.');
        });
      };
      w.document.getElementById('of-copy-text').onclick = function () {
        OF.copyToClipboard(plain).then(function (ok) { if (!ok) w.alert('Copy blocked by the browser.'); });
      };
      w.document.getElementById('of-dl-html').onclick = function () {
        OF.downloadFile(fileName, html);
      };
    } catch (e) { /* preview still usable */ }
  }
  location.href = href;
  OF.info('Email draft opened in your mail app + HTML preview tab (use "Copy HTML for email body" there — mailto cannot carry HTML).', 'Outlook');
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
