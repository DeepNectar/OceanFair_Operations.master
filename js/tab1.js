/* ============================================================
   tab1.js — "Tab 1" folder (ETA INFO sheet)
   Ports: Sheet1.cls Worksheet_Change, ETA1.bas SendStylishEmail…,
          ETA2.bas AddListClass_NoPassword, VesselExtractETAData.bas
   ============================================================ */

OF.PORT_LIST = 'FUJAIRAH, DIBBA, KHOR FAKKAN, KFK, MINA SAQR, RAS AL KHAIMAH, DUBAI, JEBEL ALI, JEBAL ALI, DUBAI MARITIME CITY, DMC, SHARJAH, HAMRIYAH SHARJAH, ABU DHABI, ABU DHABI PORT, KHALIFA PORT, KHALID PORT';

// Excluded group-header values that must NOT be numbered (Sheet1.cls)
OF.EXCLUDED_LOCATIONS = [
  'Fujairah / East Coast',
  'Dubai / Jebel Ali Area / Sharjah / Hamriyah',
  'Other / Unknown',
  'Abu Dhabi Area'
];

// Port color coding from ETA1.bas
OF.PORT_COLORS = {
  'FUJAIRAH': '#B71C1C', 'DIBBA': '#E65100', 'KHOR FAKKAN': '#4A148C', 'KFK': '#4A148C',
  'MINA SAQR': '#880E4F', 'RAS AL KHAIMAH': '#880E4F', 'DUBAI': '#0D47A1',
  'JEBEL ALI': '#1B5E20', 'JEBAL ALI': '#1B5E20', 'DUBAI MARITIME CITY': '#004D40',
  'DMC': '#004D40', 'SHARJAH': '#F57F17', 'HAMRIYAH SHARJAH': '#F57F17',
  'ABU DHABI': '#4E342E', 'ABU DHABI PORT': '#4E342E', 'KHALIFA PORT': '#4E342E',
  'KHALID PORT': '#F57F17', 'DRY DOCK': '#00695C'
};

/* ---------- Sheet1.cls Worksheet_Change port ----------
   Auto SR numbering in column A for rows 2..101 (index 1..100),
   skipping excluded location headers; formats A:E.            */
OF.tab1WorksheetChange = function () {
  var data = OF.readGrid('grid-tab1');
  var counter = 1;
  for (var i = 1; i <= 100 && i < data.length; i++) { // rows 2..101
    var b = (data[i][1] || '').trim();
    if (b !== '' && OF.EXCLUDED_LOCATIONS.indexOf(b) === -1) {
      OF.writeCell('grid-tab1', i, 0, counter, 'center tnr');
      counter++;
    } else {
      OF.writeCell('grid-tab1', i, 0, '', 'center tnr');
    }
  }
  OF.save('grid-tab1', OF.readGrid('grid-tab1'));
};

OF.initTab1 = function () {
  var demo = OF.gridDataOr('grid-tab1', [
    ['SR#', 'SHIP TO/VESSEL', 'PORT', 'ETA-ETB-ETD', 'REMARKS (Confirm for FFV, Bread & Dairy arrangement)'],
    ['', 'MV OCEAN STAR', 'JEBEL ALI', 'ETA 05-10 0800 / ETB 1400', 'FFV required'],
    ['', 'MV ARABIAN GULF', 'FUJAIRAH', 'ETA 06-10 0600', 'Bread + Dairy'],
    ['Fujairah / East Coast', '', '', '', ''],
    ['', 'MV DESERT ROSE', 'DIBBA', 'ETA 07-10 0900', '']
  ]);
  OF.buildGrid({
    id: 'grid-tab1',
    sheetName: 'ETA INFO',
    columns: [
      { name: 'SR#', align: 'center' },
      { name: 'SHIP TO/VESSEL', align: 'left' },
      { name: 'PORT', align: 'center' },
      { name: 'ETA-ETB-ETD', align: 'center' },
      { name: 'REMARKS (FFV, Bread & Dairy)', align: 'center' }
    ],
    rows: demo, maxRows: Math.max(101, demo.length),
    isLocked: function (r) { return r === 0; }, // header row protected like Excel
    onChange: function () { OF.tab1WorksheetChange(); }
  });
  OF.tab1WorksheetChange();
};

/* Re-apply port dropdowns after a grid rebuild (only when the flag is set). */
OF.maybeRestorePortDropdown = function () {
  if (OF.load('grid-tab1-portdropdown', false)) OF.addPortDropdown(true);
};

/* ---------- ETA2.bas AddListClass_NoPassword -> dropdown on C2:C101 ---------- */
OF.addPortDropdown = function (silent) {
  var trs = document.querySelectorAll('#grid-tab1 tbody tr');
  var ports = OF.PORT_LIST.split(',').map(function (p) { return p.trim(); });
  for (var i = 1; i < trs.length && i <= 100; i++) {
    var td = trs[i].querySelectorAll('td:not(.rownum)')[2]; // column C
    if (!td) continue;
    td.innerHTML = '';
    var sel = document.createElement('select');
    sel.style.width = '100%';
    sel.appendChild(new Option('', ''));
    ports.forEach(function (p) { sel.appendChild(new Option(p, p)); });
    sel.value = td.textContent.trim().toUpperCase();
    sel.addEventListener('change', function () {
      this.closest('td').textContent = this.value;
      OF.save('grid-tab1', OF.readGrid('grid-tab1'));
      OF.tab1WorksheetChange();
    });
    td.removeAttribute('contenteditable');
    td.appendChild(sel);
  }
  OF.save('grid-tab1-portdropdown', true); // remember so a later grid rebuild re-applies it
  if (!silent) OF.info('Dropdown lists for ports added successfully in Column C!\nSheet is protected (without password) but dropdowns work.', 'ETA2');
};

/* ---------- ETA1.bas SendStylishEmailWithTableofETA ---------- */
OF.sendEtaEmail = function () {
  var data = OF.readGrid('grid-tab1');
  var htmlTable = "<table style='border-collapse: collapse; font-family: Arial, sans-serif;'>";
  htmlTable += "<thead><tr style='background-color:#40E0D0;color:#FF00FF;'>";
  ['SR#','VESSEL NAME','PORT','ETA-ETB-ETD','REMARKS<br>(Confirm for FFV, Bread &amp; Dairy arrangement)']
    .forEach(function (h) {
      htmlTable += "<th style='padding:12px 8px;border:6px solid #E0E0E0;text-align:center;'>" + h + "</th>";
    });
  htmlTable += "</tr></thead><tbody>";
  var any = false;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] !== '') { // only rows with SR# (rng.Cells(i,1) not empty)
      any = true;
      htmlTable += '<tr>';
      htmlTable += "<td style='padding:6px 12px;border:6px solid #E0E0E0;text-align:center;font-weight:bold;'>" + OF.esc(row[0]) + "</td>";
      htmlTable += "<td style='padding:6px 12px;border:6px solid #E0E0E0;font-weight:bold;'>" + OF.esc(row[1]) + "</td>";
      var port = (row[2] || '').toUpperCase();
      var color = OF.PORT_COLORS[port] || '#263238';
      htmlTable += "<td style='padding:6px 12px;border:6px solid #E0E0E0;text-align:center;font-weight:bold;color:" + color + ";'>" + OF.esc(row[2]) + "</td>";
      htmlTable += "<td style='padding:6px 12px;border:6px solid #E0E0E0;text-align:center;font-weight:bold;color:BLACK;'>" + OF.esc(row[3]) + "</td>";
      htmlTable += "<td style='padding:6px 12px;border:6px solid #E0E0E0;text-align:center;font-weight:bold;'>" + OF.esc(row[4]) + "</td>";
      htmlTable += '</tr>';
    }
  }
  htmlTable += '</tbody></table>';
  if (!any) { OF.warn('No numbered rows found — nothing to send.', 'ETA1'); return; }

  var htmlBody = "<html><head><style>" +
    "body{margin:0;padding:0;background-color:#f5f7fa;}" +
    ".email-container{margin:0 auto;}.email-content{background-color:lightblue;padding:0px;}" +
    "h1{color:#2c3e50;font-size:25pt;font-family:'Kristen ITC','Times New Roman';}" +
    "p{color:#FF00FF;line-height:1;font-family:'Times New Roman';font-size:15pt;margin:5px 0;}" +
    "</style></head><body><div class='email-container'><div class='email-content'>" +
    "<h1 style='color:#8B0000;'>Vessel Schedule Update</h1>" +
    "<p>Dear Team,</p>" +
    "<p>Please advise the exact <b>ETA</b> for the below vessels with <b>REMARKS</b> to enable us to arrange the fresh items accordingly.</p>" +
    htmlTable + "</div></div></body></html>";

  OF.openMail({
    to: 'jebelali@oceanfair.com; fujairah@oceanfair.com; john@oceanfair.com',
    cc: 'bhavesh@oceanfair.com;deep@oceanfair.com;desai@oceanfair.com;dispatch@oceanfair.com;ajay@oceanfair.com;dry@oceanfair.com;frozen@oceanfair.com;sunil@oceanfair.com;bilal@oceanfair.com;technical@oceanfair.com;akash@oceanfair.com;akshay@oceanfair.com',
    subject: 'ETA/ETB - INFO -- ' + OF.formatDateLong(new Date()),
    htmlBody: htmlBody
  });
  OF.info('Your stylish email for E.T.A / E.T.B sent!', 'Success');
};

/* ---------- VesselExtractETAData.bas ExtractVesselData ----------
   Browser can't open external .xlsx silently, so the user uploads
   the file. Real .xlsx files are now parsed properly (zip + sheet
   XML reader in core.js); CSV/TSV/TXT use the delimited reader.  */
OF.extractVesselData = function () {
  return new Promise(function (resolve) {
    var inp = document.createElement('input');
    inp.type = 'file';
    inp.accept = '.xlsx,.csv,.txt,.tsv';
    inp.onchange = function () {
      if (!inp.files.length) { OF.warn('No file selected. Operation cancelled.', 'Extract'); return resolve(); }
      var file = inp.files[0];
      var finish = function (rows) {
        if (!rows || rows.length < 2) { OF.warn('No data found in the source sheet.', 'Extract'); return resolve(); }
        OF.processVesselRows(rows, file.name);
        resolve();
      };
      var fail = function (e) { OF.error('An error occurred: ' + (e && e.message ? e.message : e), 'Extract'); resolve(); };
      if (/\.xlsx$/i.test(file.name)) {
        OF.readXlsxAsync(file).then(finish)['catch'](fail);
      } else {
        var reader = new FileReader();
        reader.onload = function () {
          try { finish(OF.parseDelimited(String(reader.result))); }
          catch (e) { fail(e); }
        };
        reader.onerror = function () { fail(new Error('Could not read the file.')); };
        reader.readAsText(file);
      }
    };
    inp.click();
  });
};

OF.getPortGroup = function (portName) { // GetPortGroup port
  var port = String(portName).trim().toUpperCase();
  if (['FUJAIRAH','DIBBA','KHOR FAKKAN','KFK','MINA SAQR','RAS AL KHAIMAH'].indexOf(port) >= 0)
    return 'Fujairah / East Coast';
  if (['DUBAI','JEBEL ALI','JEBAL ALI','DUBAI MARITIME CITY','DMC','SHARJAH','HAMRIYAH SHARJAH'].indexOf(port) >= 0)
    return 'Dubai / Jebel Ali Area / Sharjah / Hamriyah';
  if (['ABU DHABI','ABU DHABI PORT','KHALIFA PORT','KHALID PORT'].indexOf(port) >= 0)
    return 'Abu Dhabi Area';
  return 'Other / Unknown';
};

OF.processVesselRows = function (rows, fileName) {
  var validCodes = ['YA', 'AF', 'DA', 'BH']; // column G check
  rows = rows.filter(function (r) { return r && r.some(function (c) { return String(c).trim() !== ''; }); }); // drop blank lines
  if (rows.length < 2) { OF.warn('No data found in the source sheet.', 'Extract'); return; }

  var header = (rows[0] || []).map(function (h) { return String(h).trim().toUpperCase(); });
  var findCol = function (re) {
    for (var c = 0; c < header.length; c++) if (re.test(header[c])) return c;
    return -1;
  };
  // Locate columns by header name first (robust against extra/shifted columns);
  // fall back to the original VBA fixed positions: D(3)=vessel, G(6)=code, K(10)=port.
  var vesselCol = findCol(/VESSEL|SHIP\s*TO|SHIPMENT|CLIENT|CUSTOMER/);
  var codeCol   = findCol(/AGENT|CODE|^G$/);
  var portCol   = findCol(/PORT|DISCHARGE|LOADING\s*PORT/);
  if (vesselCol === -1) vesselCol = 3;
  if (codeCol === -1)   codeCol = 6;
  if (portCol === -1)   portCol = 10;

  var cellAt = function (r, idx) { return String(r[idx] === undefined ? '' : r[idx]).trim(); };
  var dict = {}; // vessel -> [port, group]
  for (var i = 1; i < rows.length; i++) { // first pass from row 2
    var r = rows[i];
    var code = cellAt(r, codeCol).toUpperCase(); // col G
    if (validCodes.indexOf(code) === -1) continue;
    var vessel = cellAt(r, vesselCol); // col D
    if (!vessel) continue;
    if (!(vessel in dict)) {
      var port = cellAt(r, portCol).toUpperCase(); // col K
      dict[vessel] = [port, OF.getPortGroup(port)];
    }
  }
  var keys = Object.keys(dict);
  if (!keys.length) {
    OF.warn('No data found matching the criteria (YA, AF, DA, BH) in column "' +
            (header[codeCol] || ('col ' + (codeCol + 1))) + '".\n' +
            'Tip: the file must have a header row and the agent-code column containing YA/AF/DA/BH.',
            'Extract');
    return;
  }

  var groups = ['Fujairah / East Coast', 'Dubai / Jebel Ali Area / Sharjah / Hamriyah', 'Abu Dhabi Area', 'Other / Unknown'];
  var out = [['', 'SHIP TO/VESSEL', 'DELIVERY PORT', 'ETA-ETB-ETD', 'REMARKS']];
  groups.forEach(function (g) {
    var members = keys.filter(function (k) { return dict[k][1] === g; });
    if (!members.length) return;
    out.push(['', g, '', '', '']); // group header (excluded from numbering automatically)
    members.forEach(function (k) { out.push(['', k, dict[k][0], '', '']); });
    out.push(['', '', '', '', '']); // blank spacer row
  });
  OF.save('grid-tab1', out);
  OF.initTab1();
  OF.maybeRestorePortDropdown(); // keep the C-column dropdowns alive after the rebuild
  OF.info('Data extraction completed!\nSource File: ' + fileName + '\nTotal unique vessels extracted: ' + keys.length, 'Extract');
};
