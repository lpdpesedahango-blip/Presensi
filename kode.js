// ============================================================
// SIPADAN LPD PESEDAHAN - CODE.GS
// Versi perbaikan login + database otomatis + serialisasi tanggal
// ============================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SIPADAN LPD Pesedahan')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ------------------------------------------------------------
// 1. DATABASE
// ------------------------------------------------------------
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // USERS
  let shUsers = ss.getSheetByName("Users");
  if (!shUsers) {
    shUsers = ss.insertSheet("Users");
    shUsers.appendRow(["ID", "NIK", "Nama", "PIN", "Role", "Jabatan", "Barcode", "Foto_URL"]);
    shUsers.appendRow(["U1", "12345", "Karyawan Demo", "1234", "karyawan", "Kolektor", "12345", ""]);
    shUsers.appendRow(["U2", "99999", "Admin LPD", "admin123", "admin", "", "99999", ""]);
  } else if (shUsers.getLastRow() === 0) {
    shUsers.appendRow(["ID", "NIK", "Nama", "PIN", "Role", "Jabatan", "Barcode", "Foto_URL"]);
  } else {
    let headers = shUsers.getRange(1, 1, 1, Math.max(shUsers.getLastColumn(), 5)).getValues()[0].map(String);
    if (!headers.includes("Jabatan")) {
      const barcodeHeaderIndex = headers.indexOf("Barcode");
      if (barcodeHeaderIndex >= 0) {
        shUsers.insertColumnBefore(barcodeHeaderIndex + 1);
        shUsers.getRange(1, barcodeHeaderIndex + 1).setValue("Jabatan");
      } else {
        shUsers.getRange(1, shUsers.getLastColumn() + 1).setValue("Jabatan");
      }
      headers = shUsers.getRange(1, 1, 1, shUsers.getLastColumn()).getValues()[0].map(String);
    }
    if (!headers.includes("Barcode")) {
      shUsers.getRange(1, shUsers.getLastColumn() + 1).setValue("Barcode");
      headers = shUsers.getRange(1, 1, 1, shUsers.getLastColumn()).getValues()[0].map(String);
    }
    if (!headers.includes("Foto_URL")) {
      shUsers.getRange(1, shUsers.getLastColumn() + 1).setValue("Foto_URL");
    }
  }

  // CONFIG
  let shConfig = ss.getSheetByName("Config");
  if (!shConfig) {
    shConfig = ss.insertSheet("Config");
    shConfig.appendRow(["Key", "Value", "Description"]);
    shConfig.appendRow(["jam_masuk", "07:30", "Jam batas masuk (HH:mm)"]);
    shConfig.appendRow(["toleransi", "15", "Toleransi keterlambatan (menit)"]);
    shConfig.appendRow(["lat_kantor", "-8.650000", "Latitude Kantor LPD"]);
    shConfig.appendRow(["lng_kantor", "115.216667", "Longitude Kantor LPD"]);
    shConfig.appendRow(["radius", "50", "Radius diizinkan (meter)"]);
    shConfig.appendRow(["qr_code", "LPD-PESEDAHAN-VALID", "Teks pada QR Code Kantor"]);
    shConfig.appendRow(["folder_id", "", "ID Folder Google Drive untuk menyimpan Foto"]);
    shConfig.appendRow(["logo_url", "https://drive.google.com/thumbnail?id=18Xafn6aFfRfyMePZvqxPdawAA-cANgxt&sz=w500", "URL Logo Aplikasi E-ABSEN KARYAWAN"]);
  } else if (shConfig.getLastRow() === 0) {
    shConfig.appendRow(["Key", "Value", "Description"]);
  }

  // LOGO APLIKASI - migrasi aman untuk database lama
  const cfgValues = shConfig.getDataRange().getValues();
  const hasLogoUrl = cfgValues.slice(1).some(r => String(r[0] || '').trim() === 'logo_url');
  if (!hasLogoUrl) {
    shConfig.appendRow(["logo_url", "https://drive.google.com/thumbnail?id=18Xafn6aFfRfyMePZvqxPdawAA-cANgxt&sz=w500", "URL Logo Aplikasi E-ABSEN KARYAWAN"]);
  }

  // LOGS PRESENSI
  let shLogs = ss.getSheetByName("Logs_Presensi");
  if (!shLogs) {
    shLogs = ss.insertSheet("Logs_Presensi");
    shLogs.appendRow([
      "ID", "Waktu_Timestamp", "Tanggal", "Jam", "NIK", "Nama",
      "Jarak_Meter", "Status_Kehadiran", "Foto_URL"
    ]);
  }

  // PENGAJUAN IZIN
  let shIzin = ss.getSheetByName("Pengajuan_Izin");
  if (!shIzin) {
    shIzin = ss.insertSheet("Pengajuan_Izin");
    shIzin.appendRow([
      "ID", "Waktu_Pengajuan", "NIK", "Nama", "Jenis_Izin",
      "Tgl_Mulai", "Tgl_Selesai", "Keterangan", "Lampiran_URL",
      "Status_Approval"
    ]);
  }

  return "Setup Database Berhasil!";
}

// Pastikan semua sheet tersedia sebelum dipakai.
function ensureDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Aman dipanggil berulang: membuat sheet yang belum ada dan
  // melakukan migrasi kolom Jabatan/Barcode/Foto_URL pada database lama.
  setupDatabase();
  return ss;
}

// ------------------------------------------------------------
// 2. HELPER
// ------------------------------------------------------------

// Penting: google.script.run tidak boleh menerima Date object.
// Semua Date diubah menjadi string.
function serializeValue(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    if (isNaN(value.getTime())) return "";
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm:ss"
    );
  }

  if (value === null || value === undefined) return "";

  return value;
}

function cleanData(sheet) {
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (!data || data.length === 0) return [];

  const headers = data.shift();

  return data.map(row => {
    const obj = {};

    headers.forEach((header, index) => {
      obj[header] = serializeValue(row[index]);
    });

    return obj;
  });
}

function getConfigObject(sheet) {
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) return {};

  data.shift();

  const config = {};

  data.forEach(row => {
    if (row[0] !== "" && row[0] !== null) {
      config[String(row[0]).trim()] =
        row[1] === null || row[1] === undefined ? "" : String(row[1]);
    }
  });

  return config;
}

// ------------------------------------------------------------
// 3. LOGIN - SERVER SIDE
// ------------------------------------------------------------
function loginUser(nik, pin) {
  try {
    const ss = ensureDatabase();
    const shUsers = ss.getSheetByName("Users");

    if (!shUsers) {
      return {
        status: false,
        message: "Sheet Users tidak ditemukan."
      };
    }

    nik = String(nik || "").trim();
    pin = String(pin || "").trim();

    if (!nik || !pin) {
      return {
        status: false,
        message: "NIK dan PIN wajib diisi."
      };
    }

    const data = shUsers.getDataRange().getValues();

    if (data.length < 2) {
      return {
        status: false,
        message: "Belum ada data pengguna pada Sheet Users."
      };
    }

    const headers = data[0].map(h => String(h).trim());

    const idxNIK = headers.indexOf("NIK");
    const idxNama = headers.indexOf("Nama");
    const idxPIN = headers.indexOf("PIN");
    const idxRole = headers.indexOf("Role");
    const idxID = headers.indexOf("ID");

    if (idxNIK === -1 || idxPIN === -1) {
      return {
        status: false,
        message: "Kolom NIK atau PIN pada Sheet Users tidak ditemukan."
      };
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      const rowNIK = String(row[idxNIK] ?? "").trim();
      const rowPIN = String(row[idxPIN] ?? "").trim();

      if (rowNIK === nik && rowPIN === pin) {
        return {
          status: true,
          user: {
            ID: idxID >= 0 ? String(row[idxID] ?? "") : "",
            NIK: rowNIK,
            Nama: idxNama >= 0 ? String(row[idxNama] ?? "") : "",
            Role: idxRole >= 0
              ? String(row[idxRole] ?? "karyawan").trim().toLowerCase()
              : "karyawan",
            Jabatan: headers.indexOf("Jabatan") >= 0
              ? String(row[headers.indexOf("Jabatan")] ?? "").trim()
              : "",
            Foto_URL: headers.indexOf("Foto_URL") >= 0
              ? String(row[headers.indexOf("Foto_URL")] ?? "").trim()
              : "",
            Barcode: headers.indexOf("Barcode") >= 0
              ? String(row[headers.indexOf("Barcode")] ?? "").trim()
              : ""
          },
          message: "Login berhasil."
        };
      }
    }

    return {
      status: false,
      message: "NIK atau PIN salah."
    };

  } catch (err) {
    console.error("loginUser: " + err.stack);
    return {
      status: false,
      message: "Terjadi kesalahan server: " + err.message
    };
  }
}

// ------------------------------------------------------------
// 4. READ DATA
// ------------------------------------------------------------

function saveUser(user) {
  try {
    ensureDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Users");
    const config = getConfigObject(ss.getSheetByName("Config"));
    if (!sh) throw new Error("Sheet Users tidak ditemukan.");

    let headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 5)).getValues()[0].map(h => String(h).trim());

    if (!headers.includes("Jabatan")) {
      const barcodeIndex = headers.indexOf("Barcode");
      if (barcodeIndex >= 0) {
        sh.insertColumnBefore(barcodeIndex + 1);
        sh.getRange(1, barcodeIndex + 1).setValue("Jabatan");
      } else {
        sh.getRange(1, sh.getLastColumn() + 1).setValue("Jabatan");
      }
      headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
    }
    if (!headers.includes("Barcode")) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue("Barcode");
      headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
    }
    if (!headers.includes("Foto_URL")) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue("Foto_URL");
      headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
    }

    const idxID = headers.indexOf("ID");
    const idxNIK = headers.indexOf("NIK");
    const idxNama = headers.indexOf("Nama");
    const idxPIN = headers.indexOf("PIN");
    const idxRole = headers.indexOf("Role");
    const idxJabatan = headers.indexOf("Jabatan");
    const idxBarcode = headers.indexOf("Barcode");
    const idxFoto = headers.indexOf("Foto_URL");

    const nik = String(user.NIK || "").trim();
    const nama = String(user.Nama || "").trim();
    const pin = String(user.PIN || "").trim();
    const role = String(user.Role || "karyawan").trim().toLowerCase();
    const jabatan = String(user.Jabatan || "").trim();
    const barcode = String(user.Barcode || nik).trim();

    const allowedJabatan = [
      "Pemucuk",
      "Petengen/Kasir",
      "Penyarikan/Tata Usaha",
      "Kabag Perkreditan",
      "Kabag Dana",
      "Kolektor"
    ];

    if (!nik || !nama || !pin) return {status:false, message:"NIK, Nama, dan PIN wajib diisi."};
    if (role !== "admin" && allowedJabatan.indexOf(jabatan) === -1) {
      return {status:false, message:"Jabatan karyawan wajib dipilih dari daftar yang tersedia."};
    }

    const values = sh.getDataRange().getValues();
    let editRow = 0;
    let oldPhotoUrl = "";

    if (user.ID && idxID >= 0) {
      for (let r = 1; r < values.length; r++) {
        if (String(values[r][idxID]) === String(user.ID)) {
          editRow = r + 1;
          oldPhotoUrl = idxFoto >= 0 ? String(values[r][idxFoto] || "") : "";
          break;
        }
      }
    }

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const rowId = idxID >= 0 ? String(row[idxID] || "") : "";
      if (editRow && rowId === String(user.ID)) continue;
      if (idxNIK >= 0 && String(row[idxNIK] || "").trim() === nik) {
        return {status:false, message:"NIK/ID karyawan sudah digunakan."};
      }
      if (idxBarcode >= 0 && String(row[idxBarcode] || "").trim() === barcode) {
        return {status:false, message:"QR ID Card sudah digunakan oleh karyawan lain."};
      }
    }

    const id = editRow ? String(user.ID) : "U" + new Date().getTime();
    let fotoUrl = oldPhotoUrl;

    // Foto baru dikirim dari Index.html sebagai Base64 dan disimpan ke Drive.
    if (user.PhotoBase64) {
      const safeNik = nik.replace(/[^A-Za-z0-9_-]/g, "_");
      const fileName = "Foto_Karyawan_" + safeNik + "_" + Date.now() + ".jpg";
      fotoUrl = saveFileToDrive(user.PhotoBase64, fileName, config.folder_id);
      if (!fotoUrl) return {status:false, message:"Foto gagal disimpan ke Google Drive."};
    } else if (user.PhotoURL !== undefined && String(user.PhotoURL).trim() !== "") {
      fotoUrl = String(user.PhotoURL).trim();
    }

    const width = Math.max(sh.getLastColumn(), headers.length);
    const rowData = new Array(width).fill("");

    // Saat edit, pertahankan kolom lain yang mungkin sudah ada.
    if (editRow) {
      const existing = sh.getRange(editRow, 1, 1, width).getValues()[0];
      for (let i = 0; i < width; i++) rowData[i] = existing[i];
    }

    if (idxID >= 0) rowData[idxID] = id;
    if (idxNIK >= 0) rowData[idxNIK] = nik;
    if (idxNama >= 0) rowData[idxNama] = nama;
    if (idxPIN >= 0) rowData[idxPIN] = pin;
    if (idxRole >= 0) rowData[idxRole] = role;
    if (idxJabatan >= 0) rowData[idxJabatan] = jabatan;
    if (idxBarcode >= 0) rowData[idxBarcode] = barcode;
    if (idxFoto >= 0) rowData[idxFoto] = fotoUrl;

    if (editRow) sh.getRange(editRow, 1, 1, width).setValues([rowData]);
    else sh.getRange(sh.getLastRow() + 1, 1, 1, width).setValues([rowData]);

    return {status:true, message:editRow ? "Data karyawan berhasil diperbarui." : "Karyawan berhasil ditambahkan.", id:id, fotoUrl:fotoUrl};
  } catch (err) {
    return {status:false, message:"Gagal menyimpan karyawan: " + err.message};
  }
}
function deleteUser(id) {
  try {
    ensureDatabase();
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if (!sh) return {status:false, message:"Sheet Users tidak ditemukan."};
    const values = sh.getDataRange().getValues();
    for (let r = 1; r < values.length; r++) {
      if (String(values[r][0]) === String(id)) {
        sh.deleteRow(r + 1);
        return {status:true, message:"Karyawan berhasil dihapus."};
      }
    }
    return {status:false, message:"Data karyawan tidak ditemukan."};
  } catch (err) {
    return {status:false, message:"Gagal menghapus karyawan: " + err.message};
  }
}

function getAllData() {
  try {
    const ss = ensureDatabase();

    return {
      users: cleanData(ss.getSheetByName("Users")),
      config: getConfigObject(ss.getSheetByName("Config")),
      logs: cleanData(ss.getSheetByName("Logs_Presensi")),
      izin: cleanData(ss.getSheetByName("Pengajuan_Izin"))
    };

  } catch (err) {
    console.error("getAllData: " + err.stack);
    throw new Error("Gagal membaca database: " + err.message);
  }
}

// ------------------------------------------------------------
// 5. SIMPAN FILE KE DRIVE
// ------------------------------------------------------------
function saveFileToDrive(base64Data, fileName, folderId) {
  try {
    let folder;

    if (folderId && folderId !== "") {
      folder = DriveApp.getFolderById(folderId);
    } else {
      folder = DriveApp.getRootFolder();
    }

    const splitBase = String(base64Data).split(",");

    if (splitBase.length < 2) {
      throw new Error("Format file Base64 tidak valid.");
    }

    const type = splitBase[0]
      .split(";")[0]
      .replace("data:", "");

    const byteCharacters = Utilities.base64Decode(splitBase[1]);
    const blob = Utilities.newBlob(byteCharacters, type, fileName);

    const file = folder.createFile(blob);

    try {
      file.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW
      );
    } catch (shareErr) {
      console.warn("Sharing file tidak dapat diubah: " + shareErr);
    }

    return file.getUrl();

  } catch (err) {
    console.error("Gagal simpan file: " + err.stack);
    return "";
  }
}

// ------------------------------------------------------------
// 6. SIMPAN PRESENSI
// ------------------------------------------------------------
function saveCheckIn(data) {
  const ss = ensureDatabase();
  const sh = ss.getSheetByName("Logs_Presensi");
  const config = getConfigObject(ss.getSheetByName("Config"));

  // Validasi QR ID Card di server.
  // QR lama pada ID Card berisi nilai dari kolom B (umumnya NIK).
  // Sistem menerima QR yang sama dengan NIK atau kolom Barcode/QR ID.
  const qrValue = String(data.qrValue || "").trim();
  const shUsers = ss.getSheetByName("Users");
  const userRows = shUsers ? shUsers.getDataRange().getValues() : [];
  let matchedUser = null;

  if (userRows.length > 1) {
    const headers = userRows[0].map(h => String(h).trim());
    const idxNIK = headers.indexOf("NIK");
    const idxNama = headers.indexOf("Nama");
    const idxRole = headers.indexOf("Role");
    const idxBarcode = headers.indexOf("Barcode");

    for (let i = 1; i < userRows.length; i++) {
      const row = userRows[i];
      const rowNIK = idxNIK >= 0 ? String(row[idxNIK] ?? "").trim() : "";
      const rowBarcode = idxBarcode >= 0 ? String(row[idxBarcode] ?? "").trim() : "";
      if (rowNIK === String(data.nik).trim()) {
        matchedUser = {
          NIK: rowNIK,
          Nama: idxNama >= 0 ? String(row[idxNama] ?? "").trim() : "",
          Role: idxRole >= 0 ? String(row[idxRole] ?? "karyawan").trim().toLowerCase() : "karyawan",
          Jabatan: headers.indexOf("Jabatan") >= 0 ? String(row[headers.indexOf("Jabatan")] ?? "").trim() : "",
          Barcode: rowBarcode
        };
        break;
      }
    }
  }

  if (!matchedUser) {
    return {status:false, message:"Data karyawan tidak ditemukan."};
  }

  const allowedQr = [matchedUser.NIK, matchedUser.Barcode]
    .map(v => String(v || "").trim())
    .filter(Boolean);

  if (!qrValue || allowedQr.indexOf(qrValue) === -1) {
    return {
      status:false,
      message:"QR ID Card tidak sesuai dengan akun karyawan yang sedang login."
    };
  }

  const logs = cleanData(sh);

  const isAlreadyCheckedIn = logs.some(log =>
    String(log.NIK) === String(data.nik) &&
    String(log.Tanggal) === String(data.tanggal)
  );

  if (isAlreadyCheckedIn) {
    return {
      status: false,
      message: "Anda sudah melakukan absen masuk hari ini!"
    };
  }

  let photoUrl = "";

  if (data.photoBase64) {
    photoUrl = saveFileToDrive(
      data.photoBase64,
      "Absen_" + data.nik + "_" + data.tanggal + ".jpg",
      config.folder_id
    );
  }

  const newId = "LOG-" + new Date().getTime();

  sh.appendRow([
    newId,
    new Date(),
    data.tanggal,
    data.jam,
    matchedUser.NIK,
    matchedUser.Nama || data.nama,
    data.jarak,
    data.statusKehadiran,
    photoUrl
  ]);

  return {
    status: true,
    message: "Absen Masuk Berhasil!"
  };
}

// ------------------------------------------------------------
// 7. SIMPAN IZIN
// ------------------------------------------------------------
function saveLeave(data) {
  const ss = ensureDatabase();
  const sh = ss.getSheetByName("Pengajuan_Izin");
  const config = getConfigObject(ss.getSheetByName("Config"));

  let lampiranUrl = "";

  if (data.fileBase64) {
    lampiranUrl = saveFileToDrive(
      data.fileBase64,
      "Izin_" + data.nik + "_" + new Date().getTime() + ".png",
      config.folder_id
    );
  }

  const newId = "IZIN-" + new Date().getTime();

  sh.appendRow([
    newId,
    new Date(),
    data.nik,
    data.nama,
    data.jenis,
    data.tglMulai,
    data.tglSelesai,
    data.keterangan,
    lampiranUrl,
    "Menunggu"
  ]);

  return {
    status: true,
    message: "Pengajuan izin berhasil dikirim!"
  };
}

// ------------------------------------------------------------
// 8. UPDATE STATUS IZIN
// ------------------------------------------------------------
function updateLeaveStatus(id, newStatus) {
  const ss = ensureDatabase();
  const sh = ss.getSheetByName("Pengajuan_Izin");
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sh.getRange(i + 1, 10).setValue(newStatus);

      return {
        status: true,
        message: "Status Izin Diperbarui"
      };
    }
  }

  return {
    status: false,
    message: "Data tidak ditemukan"
  };
}

// ------------------------------------------------------------
// 9. SIMPAN LOGO APLIKASI
// ------------------------------------------------------------
function saveLogo(base64Data, originalFileName) {
  try {
    const ss = ensureDatabase();
    const shConfig = ss.getSheetByName("Config");
    const config = getConfigObject(shConfig);
    if (!base64Data) return {status:false, message:"File logo belum dipilih."};

    const ext = String(originalFileName || 'logo.png').split('.').pop().toLowerCase().replace(/[^a-z0-9]/g, '');
    const safeExt = ['png','jpg','jpeg','webp'].indexOf(ext) >= 0 ? ext : 'png';
    const fileName = 'Logo_EABSEN_KARYAWAN_' + new Date().getTime() + '.' + safeExt;
    const logoUrl = saveFileToDrive(base64Data, fileName, config.folder_id);
    if (!logoUrl) return {status:false, message:"Logo gagal disimpan ke Google Drive."};

    const data = shConfig.getDataRange().getValues();
    let rowFound = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === 'logo_url') { rowFound = i + 1; break; }
    }
    if (rowFound) shConfig.getRange(rowFound, 2).setValue(logoUrl);
    else shConfig.appendRow(['logo_url', logoUrl, 'URL Logo Aplikasi E-ABSEN KARYAWAN']);

    return {status:true, message:"Logo aplikasi berhasil disimpan.", logoUrl:logoUrl};
  } catch (err) {
    return {status:false, message:"Gagal menyimpan logo: " + err.message};
  }
}

// ------------------------------------------------------------
// 10. SIMPAN CONFIG
// ------------------------------------------------------------
function saveConfig(settings) {
  const ss = ensureDatabase();
  const sh = ss.getSheetByName("Config");
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();

    if (Object.prototype.hasOwnProperty.call(settings, key)) {
      sh.getRange(i + 1, 2).setValue(settings[key]);
    }
  }

  return {
    status: true,
    message: "Pengaturan berhasil disimpan"
  };
}
function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    let response = {};

    if (request.action === "login") {
      response = loginUser(request.nik, request.pin);
    } 
    else if (request.action === "saveCheckIn") {
      response = saveCheckIn(request.data);
    } 
    else {
      response = { status: false, message: "Aksi tidak dikenali" };
    }

    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: false, message: "Error: " + err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}