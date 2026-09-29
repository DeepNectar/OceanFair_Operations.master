/* ============================================================
   tab2.js - "tab 2" folder (RE-ORDER sheet)
   Port: REORDER.bas SendDynamicHTMLTableEmailofREORDER +
         frmDatePicker.frm (date picker modal in app.js)
   ============================================================ */

OF.initTab2 = function () {
  var demo = OF.gridDataOr('grid-tab2', [
    ['SR NO#','ITEM CODE#','ITEM#','QTY#','UNIT#','REMARKS#','SUPPLIER SR#'],
    ['1','0135054','WATER, MINERAL, 1.5LTR X12','50','CS','','12'],
    ['2','0110221','BREAD, WHOLE WHEAT, 500G','30','PCS','','7'],
    ['3','0122089','MILK, FRESH, FULL CREAM 1L','40','CS','','16'],
    ['4','0130077','CHICKEN, BREAST FROZEN 1KG','25','CS','','12']
  ]);
  OF.buildGrid({
    id: 'grid-tab2',
    columns: [
      { name: 'SR NO#', align: 'center' },
      { name: 'ITEM CODE#', align: 'center' },
      { name: 'ITEM#', align: 'left' },
      { name: 'QTY#', align: 'center' },
      { name: 'UNIT#', align: 'center' },
      { name: 'REMARKS#', align: 'center' },
      { name: 'SUPPLIER SR#', align: 'center' }
    ],
    rows: demo, maxRows: 40,
    isLocked: function (r) { return r === 0; }
  });
};

/* ---------- CreateHTMLTable port (exact styles from REORDER.bas) ---------- */
OF.createReorderHtmlTable = function (data, srNumber) {
  var html = "<table border='3.3' cellpadding='5' cellspacing='2' style='border-collapse:collapse;font-family:Times New Roman;font-size:14px;width:auto;border:3.3px solid #4472C4;'>";
  html += "<tr style='background-color:#4472C4;color:white;font-weight:bold;'>";
  ['SR NO#','ITEM CODE#','ITEM#','QTY#','UNIT#','REMARKS#'].forEach(function (h) {
    html += "<td style='border:3.3px solid #C47444;text-align:center;padding:5px;'>" + h + "</td>";
  });
  html += '</tr>';

  var rowCounter = 0;
  for (var i = 1; i < data.length; i++) { // from row 2
    var a = String(data[i][0] || '').trim();
    if (a === '') continue;
    var includeRow = false;
    var g = String(data[i][6] || '').trim();
    if (g !== '') {
      g.split(',').forEach(function (s) {
        if (s.trim() === String(srNumber)) includeRow = true;
      });
    } else {
      includeRow = true; // no SR filter -> include all (VBA behavior)
    }
    if (!includeRow) continue;
    rowCounter++;
    html += "<tr style='background-color:" + (rowCounter % 2 === 0 ? '#FFFFFF' : '#E9E9E9') + ";font-weight:bold;'>";
    for (var col = 0; col < 6; col++) { // A..F
      var cellValue = data[i][col] || '';
      if (cellValue === '') cellValue = '&nbsp;';
      html += "<td style='border:3.3px solid #C4D1E0;" +
              (col === 2 ? 'text-align:left;' : 'text-align:center;') +
              "padding:5px;" + (col === 5 ? 'color:#FF0000;' : '') + "'>" +
              OF.esc(cellValue) + "</td>";
    }
    html += '</tr>';
  }
  html += '</table>';
  return { html: html, rowCount: rowCounter };
};

/* ---------- main macro button (async flow: InputBox x2 + date picker) ---------- */
OF.sendReorderEmail = function () {
  return (async function () {
    try {
      var salesOrderNumber = await OF.inputBox('Please enter the Sales Order Number:', 'Sales Order Entry');
      if (salesOrderNumber === null || salesOrderNumber === '') {
        OF.warn('Operation cancelled - no Sales Order Number entered', 'RE-ORDER'); return;
      }
      var vesselName = await OF.inputBox('Please enter the Vessel Name:', 'Vessel Name Entry');
      if (vesselName === null || vesselName === '') {
        OF.warn('Operation cancelled - no Vessel Name entered', 'RE-ORDER'); return;
      }
      var picked = await OF.showDatePicker();          // frmDatePicker.Show vbModal
      if (picked.cancelled) {
        OF.warn('Operation cancelled - no Loading Date selected', 'RE-ORDER'); return;
      }
      var loadingDate = OF.FormatDateWithOrdinal(picked.date);

      var built = OF.createReorderHtmlTable(OF.readGrid('grid-tab2'), salesOrderNumber);
      if (built.rowCount === 0) {
        OF.warn('No matching items found for Sales Order ' + salesOrderNumber, 'RE-ORDER'); return;
      }

      var htmlBody = "<html><head><style>body{font-family:'Times New Roman',serif;font-size:12pt;margin:0;padding:10px;}</style></head><body>" +
        "<p>Dear Purchase team,</p>" +
        "<p>Please arrange below items for sales order# <strong>" + OF.esc(salesOrderNumber) +
        "</strong> <span style='color:red;font-weight:bold;'>Loading</span> <strong style='color:red;'>" +
        OF.esc(loadingDate) + "</strong></p>" + built.html + "</body></html>";

      OF.openMail({
        to: 'bhavesh@oceanfair.com; deep@oceanfair.com; akash@oceanfair.com; akshay@oceanfair.com',
        cc: '',
        subject: 'RE-ORDER STOCK FOR ORDER# ' + vesselName,
        htmlBody: htmlBody
      });
      OF.info('Email is ready for review with your signature', 'RE-ORDER');
    } catch (e) {
      OF.error('Error: ' + e.message + '\nPlease contact the person who created this file if this error persists.', 'RE-ORDER');
    }
  })();
};
