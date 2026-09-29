/* ============================================================
   app.js - startup / wiring (common for all tabs)
   ============================================================ */

/* Workbook_Open extras that run once access is granted */
OF.onWorkbookOpen = function () {
  if (OF._started) return;   // guard: may be called twice (session already verified)
  OF._started = true;
  OF.renderLog();          // InitializeLogSheet port (log sheet always exists)
  OF.initTab1(); OF.initTab2(); OF.initTab3();
  if (OF.TAB4_UNLOCKED) OF.initTab4(); // suppliers data feeds the Tab 3 scroll area
  OF.bindCommonButtons(); OF.bindWaterForm(); OF.bindMacroButtons();
  OF.switchTab('tab1');    // ThisWorkbook.Sheets(1).Activate
  OF.info('Welcome! Verified computer: ' + AccessControl.getComputerName(), 'Workbook Open');
};

/* central macro registry -> every button with data-macro runs code */
OF.macroMap = {
  extractVesselData:   function () { return OF.extractVesselData(); },   // Tab 1 / VesselExtractETAData.bas
  addPortDropdown:     function () { return OF.addPortDropdown(); },     // Tab 1 / ETA2.bas
  sendEtaEmail:        function () { return OF.sendEtaEmail(); },        // Tab 1 / ETA1.bas
  sendReorderEmail:    function () { return OF.sendReorderEmail(); },    // Tab 2 / REORDER.bas
  addMailList:         function () { return OF.addMailList(); },         // Tab 3 / AddMailListToInquirydatasheet.bas
  sendInquiryEmails:   function () { return OF.sendInquiryEmails(); },   // Tab 3 / SENDINQINEXCEL.bas
  viewEmailLog:        function () { return OF.viewEmailLog(); },        // Tab 3 / ViewEmailLog
  clearLogForQuote:    function () { return OF.clearLogForQuote(); },    // Tab 3 / ClearLogForQuote
  clearAllLogData:     function () { return OF.clearAllLogData(); },     // Tab 3 / ManualClearAllData
  spinUp:              function () { return OF.spinUp(); },              // Tab 3 / Sheet3.cls SpinButton1
  spinDown:            function () { return OF.spinDown(); },            // Tab 3 / Sheet3.cls SpinButton1
  unlockTab4:          function () { return OF.unlockTab4(); },          // hidden tab4.cls
  executeAllFunctions: function () { return OF.executeAllFunctions(); }, // hidden tab4.cls
  clearSheetTab1:      function () { return OF.clearSheet('grid-tab1', 'ETA INFO'); },
  clearSheetTab2:      function () { return OF.clearSheet('grid-tab2', 'RE-ORDER'); },
  clearSheetTab3:      function () { return OF.clearSheet('grid-tab3', 'Inquiry data'); },
  clearSheetTab4:      function () { return OF.clearSheet('grid-tab4', 'Suppliers'); },
  clearSheetEmaillog:  function () { return OF.clearAllLogData(); }      // EmailLog sheet clear = wipe log
};

OF.bindMacroButtons = function () {
  document.querySelectorAll('[data-macro]').forEach(function (btn) {
    btn.addEventListener('click', function () {
      var fn = OF.macroMap[btn.dataset.macro];
      if (fn) fn(); else OF.warn('Macro "' + btn.dataset.macro + '" not found.', 'Registry');
    });
  });
};

/* tab switching (sheet activation) */
OF.switchTab = function (tabId) {
  document.querySelectorAll('.tab-btn').forEach(function (b) {
    b.classList.toggle('active', b.dataset.tab === tabId);
  });
  document.querySelectorAll('.sheet-panel').forEach(function (p) {
    p.classList.toggle('active', p.id === tabId);
  });
};
document.querySelectorAll('.tab-btn').forEach(function (b) {
  b.addEventListener('click', function () { OF.switchTab(b.dataset.tab); });
});

/* ---------- frmDatePicker.frm port (modal) ---------- */
OF.showDatePicker = function () {
  return new Promise(function (resolve) {
    var m = document.getElementById('date-modal');
    var daySel = document.getElementById('dp-day'),
        monSel = document.getElementById('dp-month'),
        yrSel  = document.getElementById('dp-year');
    daySel.innerHTML = ''; monSel.innerHTML = ''; yrSel.innerHTML = '';
    for (var d = 1; d <= 31; d++) daySel.appendChild(new Option(d, d));       // Fill days
    OF.monthNames.forEach(function (mn, i) { monSel.appendChild(new Option(mn, i + 1)); }); // Fill months
    var startYear = Math.max(new Date().getFullYear(), 2025);                  // years >= 2025
    for (var y = startYear; y <= startYear + 20; y++) yrSel.appendChild(new Option(y, y));
    var now = new Date();
    daySel.value = now.getDate(); monSel.value = now.getMonth() + 1; yrSel.value = now.getFullYear();
    m.classList.remove('hidden');

    function finish(cancelled) {
      m.classList.add('hidden');
      if (cancelled) { resolve({ cancelled: true }); return; }                 // Cancelled = True
      var date = new Date(+yrSel.value, +monSel.value - 1, +daySel.value);
      if (date.getMonth() !== +monSel.value - 1 || date.getDate() !== +daySel.value) {
        OF.warn('Invalid date! Please select a valid date.', 'Date Picker');   // IsDateValid port
        OF.showDatePicker().then(resolve);
        return;
      }
      resolve({ cancelled: false, date: date });
    }
    document.getElementById('dp-ok').onclick = function () { finish(false); };
    document.getElementById('dp-cancel').onclick = function () { finish(true); };
  });
};

/* Boot entry called by core.js on DOMContentLoaded:
   verify computer first (ThisWorkbook.cls), then start workbook. */
function OF_APP_BOOT() {
  var gate = document.getElementById('access-gate');
  var app = document.getElementById('app');
  if (AccessControl.isComputerAllowed()) {          // session already verified
    gate.classList.add('hidden');
    app.classList.remove('hidden');
    OF.onWorkbookOpen();
  } else {                                          // show the access gate
    gate.classList.remove('hidden');
    app.classList.add('hidden');
    AccessControl.wireGate(OF.onWorkbookOpen);      // verify button -> callback
  }
}
