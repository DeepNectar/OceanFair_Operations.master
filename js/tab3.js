/* ============================================================
   tab3.js - "tab 3" folder (Inquiry data sheet)
   Ports: AddMailListToInquirydatasheet.bas (AddListClass),
          SENDINQINEXCEL.bas (SendEmailsToSuppliersInExcelFormat,
          EmailLog, device restriction for supplier 16),
          Sheet3.cls (SpinButton1_Change supplier scroll J7:K29)
   ============================================================ */

OF.SHEET_PASSWORD = 'OC*123458';                       // LOG_SHEET_PASSWORD / WORKBOOK_PASSWORD
OF.MAIL_LIST = ['bhavesh@oceanfair.com','deep@oceanfair.com','akash@oceanfair.com','akshay@oceanfair.com'];
OF.LOG_HEADERS = ['Quote Number','Supplier Code','Supplier Name','Email Addresses','Batches','Deadline','Total Items','Status','Logged At','Device'];
OF.MAX_ITEMS_PER_BATCH = 25;
OF.RESTRICTED_SUPPLIER_CODE = '16';
OF.spinValue = 1;

OF.initTab3 = function () {
  var demo = OF.gridDataOr('grid-tab3', [
    ['ITEM CODE','DESCRIPTION','SPEC','PACK','UOM','PRIORITY','REQ QTY','SUPPLIER SR#'],
    ['0135054','WATER, MINERAL, 1.5 LTR X 12 BTL (FUJAIRAH)','12x1.5L','CS','EA','NORMAL','50','12'],
    ['0110221','BREAD, WHOLE WHEAT, 500G','500G','PCS','EA','HIGH','30','7,12'],
    ['0122089','MILK, FRESH, FULL CREAM 1L','1L','CS','EA','NORMAL','40','16']
  ]);
  OF.buildGrid({
    id: 'grid-tab3',
    sheetName: 'Inquiry data',
    columns: [
      { name: 'A . ITEM CODE', align: 'center' },
      { name: 'B . DESCRIPTION', align: 'left' },
      { name: 'C . SPEC', align: 'center' },
      { name: 'D . PACK', align: 'center' },
      { name: 'E . UOM', align: 'center' },
      { name: 'F . PRIORITY', align: 'center' },
      { name: 'G . REQ QTY', align: 'center' },
      { name: 'H . SUPPLIER SR#', align: 'center' }
    ],
    rows: demo, maxRows: 40,
    isLocked: function (r) { return r === 0; }
  });

  // Mail list dropdowns J3:J6 (AddListClass port)
  [1,2,3,4].forEach(function (n) {
    var sel = document.getElementById('mail-list-' + n);
    sel.innerHTML = '<option value="">(select email)</option>';
    OF.MAIL_LIST.forEach(function (e) { sel.appendChild(new Option(e, e)); });
  });

  OF.refreshSupplierScroll();
};

/* ---------- AddMailListToInquirydatasheet.bas button ---------- */
OF.addMailList = function () {
  // password check like ws.Unprotect "OC*123458"
  OF.inputBox('Sheet is protected. Enter password to configure dropdowns:', 'Workbook Protected')
    .then(function (pw) {
      if (pw !== OF.SHEET_PASSWORD) { OF.warn('Incorrect password or protection error', 'AddListClass'); return; }
      [1,2,3,4].forEach(function (n) {
        var sel = document.getElementById('mail-list-' + n);
        sel.innerHTML = '<option value="">(select email)</option>';
        OF.MAIL_LIST.forEach(function (e) { sel.appendChild(new Option(e, e)); });
      });
      OF.info('Dropdown lists configured successfully!\nCells are protected but dropdowns remain usable.', 'AddListClass');
    });
};

/* ---------- device-restriction ports (SENDINQINEXCEL.bas) ---------- */
OF.shouldDisplaySupplier = function (supplierCode) {
  supplierCode = String(supplierCode);
  if (supplierCode !== OF.RESTRICTED_SUPPLIER_CODE) return true;
  return AccessControl.getComputerName() === AccessControl.ALLOWED_DEVICE_FOR_SUPPLIER16;
};
OF.shouldSendEmailForSupplier = function (supplierCode) {
  return String(supplierCode) !== OF.RESTRICTED_SUPPLIER_CODE;
};

/* ---------- EmailLog helpers (InitializeLogSheet / LogEmailSent ports) ---------- */
OF.getLog = function () { return OF.load('emaillog', []); };
OF.saveLog = function (rows) { OF.save('emaillog', rows); OF.renderLog(); };
OF.isQuoteAlreadySent = function (q) { return OF.getLog().some(function (r) { return r[0] === q; }); };
OF.isSupplierAlreadySent = function (q, code) {
  return OF.getLog().some(function (r) { return r[0] === q && String(r[1]) === String(code); });
};
OF.logEmailSent = function (quote, code, name, emails, batches, deadline, items, status) {
  var log = OF.getLog();
  log.push([quote, code, name, emails, batches, deadline.toLocaleString(), items, status, new Date().toLocaleString(), AccessControl.getComputerName()]);
  OF.saveLog(log);
};
OF.renderLog = function () {
  var rows = [OF.LOG_HEADERS].concat(OF.getLog());
  OF.buildGrid({
    id: 'grid-emaillog',
    sheetName: 'EmailLog',
    columns: OF.LOG_HEADERS.map(function (h) { return { name: h, align: 'center' }; }),
    rows: rows, maxRows: Math.max(rows.length, 12),
    isLocked: function () { return true; } // log sheet is read-only/protected
  });
};

/* ---------- calculateDefaultDeadline port (2 hours, office hours) ---------- */
OF.calculateDefaultDeadline = function () {
  var d = new Date(Date.now() + 2 * 3600 * 1000);
  // AdjustDeadlineToOfficeHours: office Sat-Thu, skip weekends only
  while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + 1);
  return d;
};

/* ---------- MAIN BUTTON: SendEmailsToSuppliersInExcelFormat ---------- */
OF.sendInquiryEmails = function () {
  return (async function () {
    try {
      var referenceNumber = (document.getElementById('inq-ref-no').value || '').trim();
      if (!referenceNumber) { OF.warn('Please enter a Quote Number before sending emails.', 'SEND INQ'); return; }

      var suppliers = OF.load('grid-tab4', null) || OF.defaultSuppliers(); // hidden Suppliers sheet
      var data = OF.readGrid('grid-tab3');

      // duplicate-quote check (Yes/No/Cancel exactly like VBA)
      var checkResponse = 'No';
      if (OF.isQuoteAlreadySent(referenceNumber)) {
        checkResponse = await OF.confirmBox(
          'Emails have already been sent for Quote Number: ' + referenceNumber +
          '\nDo you want to send emails again?\n\nYes = send again | No = cancel | Cancel = clear log for this quote and send',
          'Duplicate Quote Number', 'yesnocancel');
        if (checkResponse === 'No') return;
        if (checkResponse === 'Cancel') {
          OF.saveLog(OF.getLog().filter(function (r) { return r[0] !== referenceNumber; }));
        }
      }

      // response deadline
      var defaultDeadline = OF.calculateDefaultDeadline();
      var useDef = await OF.confirmBox('Use default 2-hour response deadline?\n\nDeadline: ' +
        defaultDeadline.toLocaleString(), 'Response Deadline', 'yesno');
      var responseDeadline;
      if (useDef === 'Yes') {
        responseDeadline = defaultDeadline;
      } else {
        var custom = await OF.inputBox('Enter deadline as YYYY-MM-DD HH:MM:', 'Modified Deadline');
        if (!custom) return;
        responseDeadline = new Date(custom.replace(' ', 'T'));
        if (isNaN(responseDeadline.getTime())) { OF.warn('Invalid date entered.', 'Deadline'); return; }
      }

      // build recipient dictionary from Suppliers sheet (code -> name + emails)
      var recipients = {};
      var restrictedNotice = [];
      for (var i = 1; i < suppliers.length; i++) {
        var srow = suppliers[i];
        var code = String(srow[0] || '').trim();
        if (!code) continue;
        if (OF.shouldDisplaySupplier(code)) {
          if (!recipients[code]) recipients[code] = { name: srow[1] || ('Supplier ' + code), emails: [] };
          for (var j = 2; j < srow.length; j++) if (String(srow[j]).trim()) recipients[code].emails.push(String(srow[j]).trim());
        } else if (code !== OF.RESTRICTED_SUPPLIER_CODE) {
          restrictedNotice.push('Supplier ' + code + ' - ' + (srow[1] || ''));
        }
      }
      if (restrictedNotice.length) {
        OF.warn('The following restricted supplier(s) were COMPLETELY HIDDEN (only available on device ' +
          AccessControl.ALLOWED_DEVICE_FOR_SUPPLIER16 + '):\n' + restrictedNotice.join('\n') +
          '\nCurrent device: ' + AccessControl.getComputerName(), 'Device Restriction Notice');
      }

      // unique supplier codes referenced in column H of Inquiry data
      var uniqueSRs = [];
      for (i = 1; i < data.length; i++) {
        var h = String(data[i][7] || '').trim();
        if (!h) continue;
        h.split(',').forEach(function (part) {
          var sr = part.trim();
          if (sr !== '' && !isNaN(Number(sr))) {
            var norm = String(Number(sr));
            if (OF.shouldDisplaySupplier(norm) && uniqueSRs.indexOf(norm) === -1) uniqueSRs.push(norm);
          }
        });
      }
      if (!uniqueSRs.length) { OF.warn('No supplier numbers found in column H.', 'SEND INQ'); return; }

      var belongsTo = function (cell, code) {
        return String(cell).split(',').some(function (p) {
          return p.trim() !== '' && !isNaN(Number(p)) && String(Number(p)) === String(code);
        });
      };
      var countItems = function (code) {
        var n = 0;
        for (var r = 1; r < data.length; r++) if (belongsTo(data[r][7], code)) n++;
        return n;
      };

      var ccList = ['inq-cc-1','inq-cc-2','inq-cc-3'].map(function (id) {
        return (document.getElementById(id).value || '').trim();
      }).filter(Boolean).join('; ');

      var emailCounter = 0;
      for (var u = 0; u < uniqueSRs.length; u++) {
        var srCode = uniqueSRs[u];
        if (!recipients[srCode]) continue;
        var supplierName = recipients[srCode].name;
        if (checkResponse !== 'Yes' && OF.isSupplierAlreadySent(referenceNumber, srCode)) {
          var again = await OF.confirmBox('Quote: ' + referenceNumber + '\nSupplier: ' + supplierName +
            '\nAn email has already been sent for this quote.\nSend another?', 'Duplicate Supplier', 'yesno');
          if (again === 'No') continue;
        }
        var totalItems = countItems(srCode);
        if (totalItems === 0) { OF.warn('No items found for supplier ' + srCode + ' - ' + supplierName, 'SEND INQ'); continue; }
        var totalBatches = Math.ceil(totalItems / OF.MAX_ITEMS_PER_BATCH);

        var emailsStr = recipients[srCode].emails.join(', ');
        var isPreview = !OF.shouldSendEmailForSupplier(srCode);

        for (var batch = 1; batch <= totalBatches; batch++) {
          var subject = 'Price & Availability Request - ' + referenceNumber + ' - ' + supplierName +
                        (totalBatches > 1 ? ' (Part ' + batch + ' of ' + totalBatches + ')' : '');
          if (isPreview) subject = '[PREVIEW] ' + subject;

          // collect this batch's item rows
          var batchRows = [], seen = 0;
          for (var r2 = 1; r2 < data.length; r2++) {
            if (belongsTo(data[r2][7], srCode)) {
              seen++;
              if (seen > (batch - 1) * OF.MAX_ITEMS_PER_BATCH && seen <= batch * OF.MAX_ITEMS_PER_BATCH) batchRows.push(data[r2]);
            }
          }
          var tbl = "<table border='1' cellpadding='4' style='border-collapse:collapse;font-family:Calibri;font-size:11pt;'>" +
                    "<tr style='background:#4472C4;color:#fff;'><th>Item Code</th><th>Description</th><th>Spec</th><th>Pack</th><th>UOM</th><th>Qty</th><th>Your Price</th><th>Availability</th></tr>";
          batchRows.forEach(function (br) {
            tbl += '<tr><td>' + OF.esc(br[0]) + '</td><td style="text-align:left">' + OF.esc(br[1]) + '</td><td>' + OF.esc(br[2]) +
                   '</td><td>' + OF.esc(br[3]) + '</td><td>' + OF.esc(br[4]) + '</td><td>' + OF.esc(br[6]) + '</td><td>&nbsp;</td><td>&nbsp;</td></tr>';
          });
          tbl += '</table>';

          var htmlBody = "<html><body style='font-family:Calibri;font-size:11pt;'>" +
            "<p>Dear " + OF.esc(supplierName) + ",</p>" +
            "<p>Please quote your best price and availability against Reference <b>" + OF.esc(referenceNumber) + "</b>" +
            (totalBatches > 1 ? " (Part " + batch + " of " + totalBatches + ", Total items: " + totalItems + ")" : "") + ":</p>" +
            tbl +
            "<p style='color:#C00000;font-weight:bold;'>Kindly respond by " + responseDeadline.toLocaleString() + ".</p>" +
            "<p>Note: Attach your signed quotation Excel file when replying.<br/>Best Regards,<br/>Oceanfair Ship Stores</p>" +
            (isPreview ? "<p style='background:#FFF2CC;'>PREVIEW MODE - this email will NOT be sent automatically.</p>" : "") +
            "</body></html>";

          if (isPreview) {
            // supplier 16 rule: display only, never send
            var w = window.open('', '_blank');
            if (w) { w.document.write(htmlBody); w.document.close(); }
          } else {
            OF.openMail({ to: emailsStr, cc: ccList, subject: subject, htmlBody: htmlBody });
            emailCounter++;
          }
          if (batch === 1) {
            OF.logEmailSent(referenceNumber, srCode, supplierName, emailsStr, totalBatches, responseDeadline, totalItems, isPreview ? 'Preview Only' : 'Sent');
          }
        }
      }

      if (emailCounter > 0) {
        OF.info(emailCounter + " email(s) created and sent successfully!\nLog updated in the 'EmailLog' sheet.\nDevice: " +
          AccessControl.getComputerName(), 'SEND INQ');
        // AUTO-CLEAR columns A-H after successful send (ClearInquiryDataColumnsAtoH port)
        var cleared = OF.readGrid('grid-tab3').map(function (row, idx) {
          return idx === 0 ? row : ['', '', '', '', '', '', '', ''];
        });
        OF.save('grid-tab3', cleared);
        OF.initTab3();
        OF.info('Inquiry data columns A:H auto-cleared after sending.', 'Auto-Clear');
      }
    } catch (e) {
      OF.error('Error: ' + e.message, 'SEND INQ');
    }
  })();
};

/* ---------- ViewEmailLog button ---------- */
OF.viewEmailLog = function () {
  OF.renderLog();
  OF.switchTab('emaillog');
  OF.info('Showing EmailLog sheet.', 'View Email Log');
};

/* ---------- ClearLogForQuote / ClearAllLogData buttons ---------- */
OF.clearLogForQuote = function () {
  return (async function () {
    var q = await OF.inputBox('Enter Quote Number whose log entries must be cleared:', 'Clear Log For Quote');
    if (!q) return;
    var before = OF.getLog().length;
    OF.saveLog(OF.getLog().filter(function (r) { return r[0] !== q; }));
    OF.info('Removed ' + (before - OF.getLog().length) + " log entry(ies) for quote '" + q + "'.", 'Clear Log');
  })();
};
OF.clearAllLogData = function () {
  return (async function () {
    var yes = await OF.confirmBox('Delete ALL EmailLog data? This cannot be undone.', 'Clear All Log Data', 'yesno');
    if (yes === 'Yes') { OF.saveLog([]); OF.info('All log data cleared.', 'Clear All'); }
  })();
};

/* ---------- Sheet3.cls SpinButton1_Change port (J7:K29 window) ---------- */
OF.refreshSupplierScroll = function () {
  var el = document.getElementById('grid-supplier-scroll');
  if (!el) return;
  var suppliers = OF.load('grid-tab4', null) || OF.defaultSuppliers();
  var lastRow = 0;
  for (var i = 1; i < suppliers.length; i++) if (String(suppliers[i][0]).trim() !== '') lastRow = i;
  var visibleRows = 23;
  if (OF.spinValue < 1) OF.spinValue = 1;
  if (OF.spinValue > Math.max(lastRow - visibleRows + 1, 1)) OF.spinValue = Math.max(lastRow - visibleRows + 1, 1);

  var out = [['J . SUPPLIER', 'K . INFO']];
  for (var k = 0; k < visibleRows; k++) {
    var srcIdx = OF.spinValue + k;
    if (srcIdx <= lastRow && suppliers[srcIdx]) out.push([suppliers[srcIdx][0], suppliers[srcIdx][1]]);
    else out.push(['', '']);
  }
  OF.buildGrid({
    id: 'grid-supplier-scroll',
    columns: [{ name: 'J . SUPPLIER', align: 'center' }, { name: 'K . INFO', align: 'left' }],
    rows: out, maxRows: 24,
    isLocked: function () { return true; } // display-only scroll area
  });
  document.getElementById('spin-value').textContent = OF.spinValue;
};
OF.spinUp = function () { OF.spinValue++; OF.refreshSupplierScroll(); };
OF.spinDown = function () { if (OF.spinValue > 1) OF.spinValue--; OF.refreshSupplierScroll(); };
