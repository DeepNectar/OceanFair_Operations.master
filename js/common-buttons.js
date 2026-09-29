/* ============================================================
   common-buttons.js - folder "all button only in all the sheet"
   These buttons live in the header so they are available on EVERY
   tab, exactly like the buttons placed on every sheet in the VBA
   workbook. Ports: GoToETAINFO/GoToREORDER/GoToInquiryData/
   GoToEmailLog/GoToSuppliers .bas, QCREPORT.bas, WATER.bas +
   WaterOrderForm.frm.
   ============================================================ */

OF.bindCommonButtons = function () {
  /* ---------- navigation ports (Worksheets("x").Activate) ---------- */
  document.querySelectorAll('.global-nav [data-action]').forEach(function (btn) {
    btn.addEventListener('click', function () {
      var action = btn.dataset.action;
      if (action === 'goto-eta')       { OF.switchTab('tab1');     OF.info("Activated 'ETA INFO' sheet.", 'GoToETAINFO'); }
      if (action === 'goto-reorder')   { OF.switchTab('tab2');     OF.info("Activated 'RE-ORDER' sheet.", 'GoToREORDER'); }
      if (action === 'goto-inquiry')   { OF.switchTab('tab3');     OF.info("Activated 'Inquiry data' sheet.", 'GoToInquiryData'); }
      if (action === 'goto-emaillog')  { OF.renderLog(); OF.switchTab('emaillog'); OF.info("Activated 'EmailLog' sheet.", 'GoToEmailLog'); }
      if (action === 'goto-suppliers') OF.goToSuppliers();
      if (action === 'qc-report')      OF.sendQcReport();
      if (action === 'water-order')    OF.openWaterOrderForm();
    });
  });

  // generic row add/del/clear mini-toolbar buttons on each sheet
  document.querySelectorAll('[data-op]').forEach(function (btn) {
    btn.addEventListener('click', function () {
      var target = btn.dataset.target, op = btn.dataset.op;
      if (op === 'clear') {
        OF.confirmBox('Clear ALL data on this sheet? (Header row kept)', 'Clear Sheet', 'yesno').then(function (y) {
          if (y !== 'Yes') return;
          var d = OF.readGrid(target);
          var cleared = d.map(function (row, idx) { return idx === 0 ? row : row.map(function () { return ''; }); });
          OF.save(target, cleared);
          OF.rebuildFromSave(target);
        });
      }
      if (op === 'row-add') {
        var d = OF.readGrid(target);
        d.push(d[0].map(function () { return ''; }));
        OF.save(target, d); OF.rebuildFromSave(target);
      }
      if (op === 'row-del') {
        var d2 = OF.readGrid(target);
        if (d2.length > 1) { d2.pop(); OF.save(target, d2); OF.rebuildFromSave(target); }
      }
    });
  });
};

// rebuild a grid from saved data using its registered spec
OF.rebuildFromSave = function (gridId) {
  var table = document.getElementById(gridId);
  if (table && table._spec) {
    table._spec.rows = OF.readGrid(gridId);
    OF.buildGrid(table._spec);
    if (gridId === 'grid-tab1') OF.tab1WorksheetChange();
    if (gridId === 'grid-tab4') OF.tab4WorksheetChange(1, 1);
    if (gridId === 'grid-tab3' || gridId === 'grid-tab4') OF.refreshSupplierScroll();
  }
};

/* ---------- GoToSuppliers.bas (device restriction) ---------- */
OF.goToSuppliers = function () {
  var name = AccessControl.getComputerName();
  var restricted = AccessControl.SUPPLIERS_RESTRICTED_COMPUTERS
    .some(function (c) { return c.toUpperCase() === name; });
  if (restricted) {
    OF.error("This device is restricted from using this Button 'Go To Suppliers'.", 'Device Restricted');
    return;
  }
  OF.switchTab('tab4'); // very-hidden sheet: shows password screen unless unlocked
  if (!OF.TAB4_UNLOCKED) {
    OF.info("Suppliers sheet is HIDDEN (xlSheetVeryHidden). Enter the password to open it.", 'GoToSuppliers');
  } else {
    OF.info("Activated 'Suppliers' sheet.", 'GoToSuppliers');
  }
};

/* ---------- QCREPORT.bas SendQCReportMail (device restriction) ---------- */
OF.sendQcReport = function () {
  var name = AccessControl.getComputerName();
  var restricted = AccessControl.QC_RESTRICTED_COMPUTERS.some(function (c) { return c.toUpperCase() === name; });
  if (restricted) {
    OF.error('This device is restricted from using this Button Send QC Report.', 'Device Restricted');
    return;
  }
  OF.inputBox('Please enter the vessel name:', 'Vessel Name').then(function (vesselName) {
    if (vesselName === null || vesselName.trim() === '') {
      OF.info('Operation cancelled - no vessel name entered', 'QC REPORT'); return;
    }
    var quoteNo = (document.getElementById('inq-ref-no').value || '').trim(); // Sheet1.Range("B1") equivalent

    // exact HTML body port from QCREPORT.bas
    var htmlBody = "<html><body style='font-family:Times New Roman; font-size:11pt;'>" +
      "<p><strong><span style='color:Brown;font-size:15pt;'>Dear Team,</span></strong></p>" +
      "<p><strong><span style='color:black;font-size:15pt;'>Please make the </span>" +
      "<span style='color:red; background-color:yellow;font-size:20pt;'>Quality Report</span>" +
      "<span style='color:black;font-size:15pt;'> for all the items in below order.</span></strong></p>" +
      "<p><strong><span style='color:blue; font-size:25pt;'>" + OF.esc(vesselName) + "</span></strong></p>" +
      "<p><strong><span style='color:red; background-color:yellow;font-size:20pt;'>Please issue more than 6 month expiry.</span></strong></p>" +
      "</body></html>";

    OF.openMail({
      to: 'sunil@oceanfair.com; dry@oceanfair.com; frozen@oceanfair.com; technical@oceanfair.com; bilal@oceanfair.com',
      cc: 'bhavesh@oceanfair.com; ciaran@oceanfair.com; jean@oceanfair.com; akshay@oceanfair.com; akash@oceanfair.com',
      subject: vesselName + ' ' + quoteNo + ' {QC REPORT}',
      htmlBody: htmlBody
    });
    OF.info('Email Created!', 'QC REPORT');
  });
};

/* ---------- WATER.bas SendWaterOrder + WaterOrderForm.frm ---------- */
OF.openWaterOrderForm = function () {
  var m = document.getElementById('water-modal');
  ['w-vessel', 'w-order', 'w-qty'].forEach(function (id) { document.getElementById(id).value = ''; });
  m.classList.remove('hidden');
  document.getElementById('w-vessel').focus();
};

OF.bindWaterForm = function () {
  // btnCancel clears inputs; closing via X/backdrop is treated as Cancel (QueryClose port)
  function cancelForm() {
    ['w-vessel', 'w-order', 'w-qty'].forEach(function (id) { document.getElementById(id).value = ''; });
    document.getElementById('water-modal').classList.add('hidden');
  }
  document.getElementById('w-cancel').addEventListener('click', cancelForm);
  document.getElementById('water-modal').addEventListener('click', function (e) {
    if (e.target === this) cancelForm();
  });
  document.getElementById('w-submit').addEventListener('click', function () {
    var qty = document.getElementById('w-qty').value.trim();
    var vessel = document.getElementById('w-vessel').value.trim();
    var order = document.getElementById('w-order').value.trim();
    // validation exactly like btnSubmit_Click
    if (!qty)    { OF.warn('Please enter quantity!', 'Water Order'); return; }
    if (!vessel) { OF.warn('Please enter vessel name!', 'Water Order'); return; }
    if (!order)  { OF.warn('Please enter order number!', 'Water Order'); return; }
    document.getElementById('water-modal').classList.add('hidden');
    OF.sendWaterOrder(vessel, order, qty);
  });
};

OF.sendWaterOrder = function (vessel, order, qty) {
  var today = new Date();
  // exact HTML email port from WATER.bas
  var htmlBody = "<html><head><style>" +
    "body,p{font-family:'Times New Roman',serif;font-size:15.5pt;}" +
    "table{font-size:inherit;border-collapse:collapse;}th{background:#f2f2f2;text-align:center;}td{padding:5px;}" +
    ".center{text-align:center;}.highlight{font-weight:bold;color:black;background-color:#ADD8E6;}" +
    "</style></head><body>" +
    "<p class='larger-text'>Dear Faisal,</p>" +
    "<p class='larger-text'>Please see below order for water requirement and please stage.</p>" +
    "<p><strong>Vessel:</strong> <span class='highlight'>" + OF.esc(vessel) + "</span></p>" +
    "<p><strong>Sales Order Number:</strong> <span class='highlight'>" + OF.esc(order) + "</span></p>" +
    "<table border='1' cellspacing='0'><tr>" +
    "<th width='15%'>Item Code</th><th width='60%'>Description</th><th width='10%'>Quantity</th><th width='15%'>Unit</th></tr>" +
    "<tr><td class='center'><b>0135054</b></td>" +
    "<td>WATER, MINERAL, 1.5 LTR X 12 BTL PER CASE (FUJAIRAH)</td>" +
    "<td class='center'>" + OF.esc(qty) + "</td><td class='center'>CS</td></tr></table></body></html>";

  OF.openMail({
    to: 'fujairah@oceanfair.com',
    cc: 'bhavesh@oceanfair.com; deep@oceanfair.com; akash@oceanfair.com; akshay@oceanfair.com',
    subject: 'Water Order Request - ' + vessel + ' - ' + OF.formatDateDDMMMYYYY(today),
    htmlBody: htmlBody
  });
  OF.info('Water order email has been sent successfully with your signature.', 'Email Sent');
};
