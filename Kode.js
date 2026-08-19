// ==============================================================================
// NAMA FILE: Kode.gs
// DESKRIPSI: Backend Apps Script Teroptimasi untuk Form Order Crane PPA Site BIB
// ==============================================================================

const SHEET_ORDERS = "Data_Orders";
const SHEET_MASTER = "Master_Data";

function doGet() {
  try {
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Sistem Monitoring Crane Truck PPA')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Sistem Monitoring Crane Truck PPA')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function saveOrder(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_ORDERS);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_ORDERS);
      sheet.appendRow([
        "ID_Order", "Timestamp", "Nama_Pemohon", "No_WA", "Perusahaan",
        "Departemen", "Section", "Tgl_Pelaksanaan", "Shift", "Waktu_Request",
        "Durasi_Request", "Lokasi", "Tujuan", "Deskripsi", "Foto_Base64",
        "Status", "Unit_CT", "GL", "Operator", "Rigger",
        "Waktu_Start_Aktual", "Waktu_End_Aktual", "Durasi_Aktual", "Alasan_Delay"
      ]);
    }

    const dateStr = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd");
    const timestamp = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm");

    const lastRow = sheet.getLastRow();
    let nextNum = 1;

    // OPTIMASI KECEPATAN: Hanya cek ID di baris paling bawah, tidak perlu me-loop seluruh data
    if (lastRow > 1) {
      const lastIdStr = String(sheet.getRange(lastRow, 1).getValue());
      if (lastIdStr.startsWith(`LIFT-${dateStr}-`)) {
        let numPart = lastIdStr.split('-')[2];
        if (numPart) {
          let lastNum = parseInt(numPart.split('_')[0]);
          if (!isNaN(lastNum)) nextNum = lastNum + 1;
        }
      }
    }

    const orderId = `LIFT-${dateStr}-${nextNum}`;

    // Tanda petik tunggal (') mencegah Google Sheets mengubah format jam menjadi desimal otomatis
    const rowData = [
      orderId, timestamp, data.nama || "", data.wa || "", data.perusahaan || "",
      data.departemen || "", data.section || "", data.tanggal || "", data.shift || "",
      "'" + (data.waktu || "00:00"), data.durasi || 0, data.lokasi || "",
      data.tujuan || "", data.deskripsi || "", data.foto || "",
      "Menunggu Validasi", "", "", "", "", "", "", "", ""
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush();

    return { success: true, id: orderId, message: "Order berhasil disimpan." };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetMaster = ss.getSheetByName(SHEET_MASTER);
  let masterData = [];
  if (sheetMaster) {
    const mData = sheetMaster.getDataRange().getValues();
    mData.shift();
    masterData = mData.filter(r => r[0] && r[1]);
  }

  const sheetOrders = ss.getSheetByName(SHEET_ORDERS);
  let ordersData = [];
  if (sheetOrders && sheetOrders.getLastRow() > 1) {
    const oData = sheetOrders.getDataRange().getValues();
    const headers = oData.shift();

    ordersData = oData.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        let val = row[index];
        if (val instanceof Date) {
          if (header === 'Tgl_Pelaksanaan') obj[header] = Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd");
          else obj[header] = Utilities.formatDate(val, "GMT+8", "HH:mm");
        } else if (typeof val === 'number' && header.includes('Waktu')) {
          let totalMinutes = Math.round(val * 24 * 60);
          let h = Math.floor(totalMinutes / 60).toString().padStart(2, '0');
          let m = (totalMinutes % 60).toString().padStart(2, '0');
          obj[header] = `${h}:${m}`;
        } else {
          obj[header] = val !== undefined && val !== null ? String(val) : "";
        }
      });
      return obj;
    });
  }

  return JSON.stringify({ master: masterData, orders: ordersData });
}

function updateJobRecord(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ORDERS);
    if (!sheet) throw new Error("Sheet order tidak ditemukan");

    const range = sheet.getDataRange();
    const values = range.getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(data.id)) {

        if (data.isResuming) {
          sheet.getRange(i + 1, 16).setValue('Completed');

          let baseId = String(data.id).split('_')[0];
          let maxSplit = 0;
          for (let j = 1; j < values.length; j++) {
            let cid = String(values[j][0]);
            if (cid.startsWith(baseId + "_")) {
              let spl = parseInt(cid.split('_')[1]);
              if (!isNaN(spl) && spl > maxSplit) maxSplit = spl;
            }
          }
          let newId = baseId + "_" + (maxSplit + 1);

          let newRow = [...values[i]];
          newRow[0] = newId;
          newRow[7] = data.tglReq || newRow[7];
          newRow[8] = data.shift || newRow[8];
          newRow[9] = data.waktuReq ? "'" + data.waktuReq : newRow[9];
          newRow[10] = data.durasiReq || newRow[10];
          newRow[13] = data.deskripsiReq || newRow[13];

          newRow[15] = data.status || "On Progress";
          newRow[16] = data.unit || "";
          newRow[17] = data.gl || "";
          newRow[18] = data.operator || "";
          newRow[19] = data.rigger || "";

          newRow[20] = ""; newRow[21] = ""; newRow[22] = ""; newRow[23] = "";
          sheet.appendRow(newRow);

        } else {
          // BATCH WRITE OPTIMIZATION
          let row = values[i];

          if (data.tglReq) row[7] = data.tglReq;
          if (data.shift) row[8] = data.shift;
          if (data.waktuReq) row[9] = "'" + data.waktuReq;
          if (data.durasiReq) row[10] = data.durasiReq;
          if (data.deskripsiReq) row[13] = data.deskripsiReq;

          if (data.status) row[15] = data.status;
          if (data.unit) row[16] = data.unit;
          if (data.gl) row[17] = data.gl;
          if (data.operator) row[18] = data.operator;
          if (data.rigger) row[19] = data.rigger;

          if (data.status === 'Completed' || data.status === 'Canceled' || data.status === 'Pending') {
            if (data.startAktual) row[20] = "'" + data.startAktual;
            if (data.endAktual) row[21] = "'" + data.endAktual;
            if (data.durasiAktual) row[22] = data.durasiAktual;
            if (data.alasanDelay) row[23] = data.alasanDelay;
          }

          sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
        }
        break;
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: "Rekam jejak job berhasil diperbarui." };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteOrder(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ORDERS);
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function addMasterItem(type, value) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_MASTER);
    sheet.appendRow([type, value]);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteMasterItem(type, value) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_MASTER);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === type && String(data[i][1]) === String(value)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    SpreadsheetApp.flush();
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updatePasswords(adminPass, ctPass) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_MASTER);
    const data = sheet.getDataRange().getValues();

    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === 'Password') sheet.deleteRow(i + 1);
    }

    sheet.appendRow(['Password', 'Admin|' + adminPass]);
    sheet.appendRow(['Password', 'CT|' + ctPass]);

    SpreadsheetApp.flush();
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function resetMasterDataOtomatis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_MASTER);
  if (!sheet) sheet = ss.insertSheet(SHEET_MASTER);
  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([["Tipe_Data", "Nilai"]]);
  sheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#c9daf8");
  SpreadsheetApp.flush();
  return "Dikosongkan.";
}

function doPost(e) {
  try {
    let request = {};
    if (e.postData && e.postData.contents) request = JSON.parse(e.postData.contents);

    let action = request.action;
    if (action === 'getAppData') return ContentService.createTextOutput(getAppData()).setMimeType(ContentService.MimeType.JSON);

    let result;
    if (action === 'saveOrder') result = saveOrder(request.payload);
    else if (action === 'updateJobRecord') result = updateJobRecord(request.payload);
    else if (action === 'deleteOrder') result = deleteOrder(request.payload);
    else if (action === 'addMasterItem') result = addMasterItem(request.cat, request.val);
    else if (action === 'deleteMasterItem') result = deleteMasterItem(request.cat, request.val);
    else if (action === 'updatePasswords') result = updatePasswords(request.admin, request.ct);
    else result = { success: false, message: 'Action not found' };

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}