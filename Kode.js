// ======= ID SPREADSHEET PRIBADI ANDA =======
const SPREADSHEET_ID = "1lkRZz2V6ZZK0UueUJ9xcj_wAFOX8wpsKxhrlkgQhQ_k";

function getSheetSmart(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'getAppData') return outputJSON(getAppData(ss));
    if (action === 'saveOrder') return outputJSON(saveOrder(ss, data.payload));
    if (action === 'updateJobRecord') return outputJSON(updateJobRecord(ss, data.payload));
    if (action === 'deleteOrder') return outputJSON(deleteOrder(ss, data.payload));
    if (action === 'addMasterItem') return outputJSON(addMasterItem(ss, data.cat, data.val));
    if (action === 'deleteMasterItem') return outputJSON(deleteMasterItem(ss, data.cat, data.val));
    if (action === 'updatePasswords') return outputJSON(updatePasswords(ss, data.admin, data.ct));

    return outputJSON({ status: 'error', message: 'Unknown action' });
  } catch (err) {
    return outputJSON({ status: 'error', message: err.toString() });
  }
}

function outputJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getAppData(ss) {
  const sheetMaster = getSheetSmart(ss, "Master_Data");
  const sheetOrders = getSheetSmart(ss, "Data_Orders");

  let masterData = sheetMaster.getDataRange().getValues();
  let ordersData = sheetOrders.getDataRange().getValues();

  // Smart Header Detection untuk Master
  if (masterData.length > 0 && (masterData[0][0] === "Tipe_Data" || masterData[0][0].toString().toLowerCase() === "perusahaan" || masterData[0][0].toString().toLowerCase() === "kategori" || masterData[0][0].toString().toLowerCase() === "nama data")) {
    if (masterData[0][1] && masterData[0][1].toString().toLowerCase() !== "pt. ppa") {
      masterData.shift();
    }
  }

  let parsedMaster = [];
  masterData.forEach(row => {
    if (row[0] && row[1]) {
      parsedMaster.push([row[0].toString().trim(), row[1].toString().trim()]);
    }
  });

  // Parse Orders
  let parsedOrders = [];
  if (ordersData.length > 1) {
    let headers = ordersData[0];
    for (let i = 1; i < ordersData.length; i++) {
      let row = ordersData[i];
      if (row[0]) {
        let obj = {};
        headers.forEach((header, index) => {
          obj[header] = row[index] !== undefined ? row[index] : "";
        });
        parsedOrders.push(obj);
      }
    }
  }

  return { master: parsedMaster, orders: parsedOrders };
}

function saveOrder(ss, p) {
  const sheet = getSheetSmart(ss, "Data_Orders");

  if (sheet.getLastRow() === 0) {
    const headers = ["ID_Order", "Timestamp", "Nama_Pemohon", "No_WA", "Perusahaan", "Departemen", "Section", "Tgl_Pelaksanaan", "Shift", "Waktu_Request", "Durasi_Request", "Lokasi", "Tujuan", "Deskripsi", "Foto_Base64", "Status", "Unit_CT", "GL", "Operator", "Rigger", "Waktu_Start_Aktual", "Waktu_End_Aktual", "Durasi_Aktual", "Alasan_Delay"];
    sheet.appendRow(headers);
  }

  // Algoritma ID Super Cepat (hanya membaca baris terakhir)
  const lr = sheet.getLastRow();
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');

  let newIdStr = `LIFT-${yyyy}${mm}${dd}-001`;
  if (lr > 1) {
    const lastId = sheet.getRange(lr, 1).getValue().toString();
    if (lastId.startsWith(`LIFT-${yyyy}${mm}${dd}`)) {
      const lastNum = parseInt(lastId.split('-')[2]);
      if (!isNaN(lastNum)) {
        newIdStr = `LIFT-${yyyy}${mm}${dd}-${String(lastNum + 1).padStart(3, '0')}`;
      }
    }
  }

  const ts = `${dd}/${mm}/${yyyy} ${String(today.getHours()).padStart(2, '0')}:${String(today.getMinutes()).padStart(2, '0')}`;

  const newRow = [
    newIdStr,
    ts,
    p.nama || "",
    p.wa || "",
    p.perusahaan || "",
    p.departemen || "",
    p.section || "",
    p.tanggal || "",
    p.shift || "",
    p.waktu || "",
    p.durasi || "",
    p.lokasi || "",
    p.tujuan || "",
    p.deskripsi || "",
    p.foto || "",
    "Menunggu Validasi",
    "", "", "", "", "", "", "", ""
  ];

  sheet.appendRow(newRow);
  return { status: "success", id: newIdStr };
}

function updateJobRecord(ss, p) {
  const sheet = getSheetSmart(ss, "Data_Orders");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idCol = headers.indexOf("ID_Order");
  if (idCol === -1) return { status: "error", message: "Kolom ID_Order tidak ditemukan" };

  const searchId = p.id.split('_')[0];

  // Logika jika Job Di-Resume / Dipending (Membuat Baris Baru dengan riwayat)
  if (p.isResuming) {
    let oldRowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol].toString().split('_')[0] === searchId) {
        oldRowIdx = i;
        break;
      }
    }
    if (oldRowIdx !== -1) {
      let newRowData = [...data[oldRowIdx]];

      newRowData[idCol] = searchId + "_RESUME_" + Math.floor(Math.random() * 1000);

      const mapUpdate = {
        "Tgl_Pelaksanaan": p.tglReq, "Waktu_Request": p.waktuReq, "Durasi_Request": p.durasiReq,
        "Deskripsi": p.deskripsiReq, "Status": p.status, "Unit_CT": p.unit, "GL": p.gl,
        "Operator": p.operator, "Rigger": p.rigger, "Shift": p.shift,
        "Waktu_Start_Aktual": "", "Waktu_End_Aktual": "", "Durasi_Aktual": "", "Alasan_Delay": ""
      };

      Object.keys(mapUpdate).forEach(key => {
        let colIdx = headers.indexOf(key);
        if (colIdx !== -1 && mapUpdate[key] !== undefined) newRowData[colIdx] = mapUpdate[key];
      });

      let statCol = headers.indexOf("Status");
      if (statCol !== -1) sheet.getRange(oldRowIdx + 1, statCol + 1).setValue("Completed");

      sheet.appendRow(newRowData);
      return { status: "success", resumed: true };
    }
  }

  // Update Normal
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol].toString().split('_')[0] === searchId) {

      const mapFields = {
        "Tgl_Pelaksanaan": p.tglReq, "Waktu_Request": p.waktuReq, "Durasi_Request": p.durasiReq,
        "Deskripsi": p.deskripsiReq, "Status": p.status, "Unit_CT": p.unit, "GL": p.gl,
        "Operator": p.operator, "Rigger": p.rigger, "Shift": p.shift,
        "Waktu_Start_Aktual": p.startAktual, "Waktu_End_Aktual": p.endAktual,
        "Durasi_Aktual": p.durasiAktual, "Alasan_Delay": p.alasanDelay
      };

      Object.keys(mapFields).forEach(key => {
        let colIdx = headers.indexOf(key);
        if (colIdx !== -1 && mapFields[key] !== undefined) {
          sheet.getRange(i + 1, colIdx + 1).setValue(mapFields[key]);
        }
      });

      return { status: "success" };
    }
  }
  return { status: "error", message: "Data Order tidak ditemukan" };
}

function deleteOrder(ss, id) {
  const sheet = getSheetSmart(ss, "Data_Orders");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { status: "success" };
    }
  }
  return { status: "error", message: "Order tidak ditemukan" };
}

function addMasterItem(ss, cat, val) {
  const sheet = getSheetSmart(ss, "Master_Data");
  sheet.appendRow([cat, val]);
  return { status: "success" };
}

function deleteMasterItem(ss, cat, val) {
  const sheet = getSheetSmart(ss, "Master_Data");
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == cat && data[i][1] == val) {
      sheet.deleteRow(i + 1);
      return { status: "success" };
    }
  }
  return { status: "error", message: "Item tidak ditemukan" };
}

function updatePasswords(ss, adminPass, ctPass) {
  const sheet = getSheetSmart(ss, "Master_Data");
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] == 'Password') {
      sheet.deleteRow(i + 1);
    }
  }

  sheet.appendRow(['Password', 'Admin|' + adminPass]);
  sheet.appendRow(['Password', 'CT|' + ctPass]);
  return { status: "success" };
}