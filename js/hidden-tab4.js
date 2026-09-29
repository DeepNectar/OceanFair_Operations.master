/* ============================================================
   hidden-tab4.js - "tab4.cls" (Suppliers sheet) - VERY HIDDEN
   Ports: tab4.cls Worksheet_Change + ExecuteAllFunctions,
          UnprotectWorkbook / ProtectWorkbook buttons.
   Sheet is hidden like xlSheetVeryHidden; unlocked only with the
   workbook password (OC*123458). Black bg / white text / red borders.
   ============================================================ */

OF.TAB4_UNLOCKED = false;

OF.defaultSuppliers = function () {
  return [
    ['SR#', 'SUPPLIER NAME', 'EMAIL 1', 'EMAIL 2'],
    ['7',  'Al Marwan Marine Supplies',  'sales@almarwan.example.com',  'orders@almarwan.example.com'],
    ['12', 'Gulf Foods International',   'info@gulffoods.example.com',  'sales@gulffoods.example.com'],
    ['16', 'Fujairah Provisions LLC',    'quote@fujprovisions.example.com', ''],
    ['21', 'Emirates Fresh Co.',         'enq@emiratesfresh.example.com', '']
  ];
};

/* ---------- tab4.cls Worksheet_Change port ----------
   Editing A/B -> unprotect, re-number all filled rows, format,
   lock filled cells, protect again.                              */
OF.tab4WorksheetChange = function (changedRow, changedCol) {
  if (!OF.TAB4_UNLOCKED) return;
  if (changedCol !== 0 && changedCol !== 1) return; // only columns A or B
  var data = OF.readGrid('grid-tab4');
  var counter = 1;
  for (var i = 1; i < data.length; i++) { // from row 2
    var a = String(data[i][0] || '').trim();
    var b = String(data[i][1] || '').trim();
    if (a !== '' || b !== '') {
      OF.writeCell('grid-tab4', i, 0, counter, 'center tnr numbered');
      OF.writeCell('grid-tab4', i, 1, b, 'left tnr numbered');
      counter++;
    } else {
      OF.writeCell('grid-tab4', i, 0, '', 'center tnr');
      OF.writeCell('grid-tab4', i, 1, '', 'left tnr');
    }
  }
  OF.save('grid-tab4', OF.readGrid('grid-tab4'));
  OF.refreshSupplierScroll();
};

/* ---------- build grid (only after unlock) ---------- */
OF.initTab4 = function () {
  var demo = OF.load('grid-tab4', null) || OF.defaultSuppliers();
  OF.buildGrid({
    id: 'grid-tab4',
    sheetName: 'Suppliers',
    columns: [
      { name: 'A . SR#', align: 'center' },
      { name: 'B . SUPPLIER NAME', align: 'left' },
      { name: 'C . EMAIL 1', align: 'left' },
      { name: 'D . EMAIL 2', align: 'left' }
    ],
    rows: demo, maxRows: 40,
    isLocked: function (r) { return r === 0; },
    onChange: function (d, r, c) { OF.tab4WorksheetChange(r, c); }
  });
};

/* ---------- Unlock button (password = WORKBOOK_PASSWORD) ---------- */
OF.unlockTab4 = function () {
  var pw = document.getElementById('tab4-password').value;
  if (pw === OF.SHEET_PASSWORD) {
    OF.TAB4_UNLOCKED = true;
    document.querySelector('#tab4 .lock-screen').classList.add('hidden');
    document.querySelector('#tab4 .unlocked-content').classList.remove('hidden');
    OF.initTab4();
    OF.tab4WorksheetChange(1, 1); // apply numbering/format right away
    OF.info('Suppliers sheet unprotected (tab4.cls UnprotectWorkbook port).', 'Unlock OK');
  } else {
    OF.error('Incorrect password! The Suppliers sheet stays hidden.', 'Unlock Failed');
  }
};

/* ---------- ExecuteAllFunctions button port ---------- */
OF.executeAllFunctions = function () {
  if (!OF.TAB4_UNLOCKED) { OF.warn('Unlock the sheet first.', 'Protected'); return; }
  OF.confirmBox(
    'This will unprotect, re-number & format rows, lock filled cells and protect the sheet again. Continue?',
    'Execute All Functions', 'yesno'
  ).then(function (yes) {
    if (yes !== 'Yes') return;

    var data = OF.readGrid('grid-tab4');
    var lastRow = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== '' || String(data[i][1]).trim() !== '') lastRow = i;
    }
    var counter = 1;
    for (i = 1; i <= lastRow; i++) {
      OF.writeCell('grid-tab4', i, 0, counter, 'center tnr numbered locked-cell');
      OF.writeCell('grid-tab4', i, 1, data[i][1], 'left tnr numbered locked-cell');
      // remove editability from filled A/B cells (ProtectionStatus=True)
      var trs = document.querySelectorAll('#grid-tab4 tbody tr');
      var tds = trs[i].querySelectorAll('td:not(.rownum)');
      tds[0].removeAttribute('contenteditable');
      tds[1].removeAttribute('contenteditable');
      counter++;
    }
    OF.save('grid-tab4', OF.readGrid('grid-tab4'));
    OF.refreshSupplierScroll();
    OF.info('Last Row: ' + (lastRow + 1) + '\nTotal Rows Processed: ' + lastRow +
            '\nSheet protected again (UserInterfaceOnly allowed).', 'Execute All Functions Completed');
  });
};
