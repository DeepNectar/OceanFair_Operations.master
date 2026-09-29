/* ============================================================
   access-control.js  —  ported from ThisWorkbook.cls
   Workbook_Open(): only allowed computers may open the workbook.
   On the web the browser cannot read Environ("ComputerName"), so the
   user must identify the machine once; the value is remembered
   (localStorage) just like the workbook "knew" its computer.
   ============================================================ */
(function () {
  'use strict';

  // Same list as VBA: Array("DC1-AKSHAY-LAP", "DC1-AKASH-PC", "DEEP-LAP", "BHAVESH-PC")
  var ALLOWED_COMPUTERS = ['DC1-AKSHAY-LAP', 'DC1-AKASH-PC', 'DEEP-LAP', 'BHAVESH-PC'];

  // Restricted-device rules used by other modules
  var SUPPLIERS_RESTRICTED_COMPUTERS = ['DC1-AKSHAY-LAP', 'DC1-AKASH-PC', '--']; // GoToSuppliers.bas
  var QC_RESTRICTED_COMPUTERS       = ['DC1-AKSHAY-LAP', 'AKASH', '--'];          // QCREPORT.bas
  var ALLOWED_DEVICE_FOR_SUPPLIER16 = 'DEEP-LAP';                                  // SENDINQINEXCEL.bas

  function getComputerName() {
    return (localStorage.getItem('of_computerName') || '').toUpperCase();
  }

  // UCase(Environ("ComputerName")) comparison from VBA
  function isComputerAllowed() {
    var name = getComputerName();
    return ALLOWED_COMPUTERS.some(function (c) { return c.toUpperCase() === name; });
  }

  function checkAccess() {
    var input = document.getElementById('computer-name-input');
    var err = document.getElementById('gate-error');
    var name = (input.value || '').trim().toUpperCase();
    if (!name) { showGateError(err, 'Please enter the computer name.'); return; }
    localStorage.setItem('of_computerName', name);
    if (isComputerAllowed()) {
      // Allowed -> hide gate, show app, then run the "Workbook_Open" extras
      document.getElementById('access-gate').classList.add('hidden');
      document.getElementById('app').classList.remove('hidden');
      if (window.OF && OF.onWorkbookOpen) OF.onWorkbookOpen();
    } else {
      // VBA: MsgBox "This workbook cannot be opened on this computer." + Close
      showGateError(err, 'This workbook cannot be opened on this computer.');
    }
  }

  function showGateError(el, msg) {
    el.textContent = msg;
    el.classList.remove('hidden');
  }

  window.AccessControl = {
    ALLOWED_COMPUTERS: ALLOWED_COMPUTERS,
    SUPPLIERS_RESTRICTED_COMPUTERS: SUPPLIERS_RESTRICTED_COMPUTERS,
    QC_RESTRICTED_COMPUTERS: QC_RESTRICTED_COMPUTERS,
    ALLOWED_DEVICE_FOR_SUPPLIER16: ALLOWED_DEVICE_FOR_SUPPLIER16,
    getComputerName: getComputerName,
    isComputerAllowed: isComputerAllowed
  };

  function wireGate(onSuccess) {
    var input = document.getElementById('computer-name-input');
    var btn = document.getElementById('btn-check-access');
    function run() {
      var err = document.getElementById('gate-error');
      var name = (input.value || '').trim().toUpperCase();
      if (!name) { showGateError(err, 'Please enter the computer name.'); return; }
      localStorage.setItem('of_computerName', name);
      if (isComputerAllowed()) {
        document.getElementById('access-gate').classList.add('hidden');
        document.getElementById('app').classList.remove('hidden');
        if (onSuccess) onSuccess();
      } else {
        // VBA: MsgBox "This workbook cannot be opened on this computer." + Close
        showGateError(err, 'This workbook cannot be opened on this computer.');
      }
    }
    btn.addEventListener('click', run);
    input.addEventListener('keydown', function (e) { if (e.key === 'Enter') run(); });
  }

  window.AccessControl = Object.assign(window.AccessControl || {}, { wireGate: wireGate });
})();
