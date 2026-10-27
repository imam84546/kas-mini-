/**
 * =================================================================
 * KAS MINI - PT IMAM INDONESIA
 * BACKEND - GOOGLE APPS SCRIPT (Kode.gs)
 * =================================================================
 */

// 1. Fungsi Utama (Menampilkan Halaman HTML)
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('KAS MINI - PT IMAM INDONESIA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * =================================================================
 * 2. INISIALISASI DATABASE (Google Sheets)
 * =================================================================
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheets = [
    {
      name: "TRANSAKSI",
      headers: ["ID", "Tanggal", "Jenis", "Kategori", "Keterangan", "Nominal", "Metode", "PenerimaSumber", "Status", "DibuatOleh", "DibuatPada", "DiperbaruiPada"]
    },
    {
      name: "KATEGORI",
      headers: ["NamaKategori", "Jenis", "Status"]
    },
    {
      name: "USERS",
      headers: ["UserID", "NamaLengkap", "Username", "Email", "Telepon", "Peran", "Status", "DibuatPada"]
    },
    {
      name: "AUDIT_LOG",
      headers: ["Waktu", "User", "Aktivitas", "Status"]
    },
    {
      name: "SETTINGS",
      headers: ["Key", "Value"]
    }
  ];

  sheets.forEach(function(s) {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.headers);
      sheet.getRange(1, 1, 1, s.headers.length).setFontWeight("bold").setBackground("#e2e8f0");
    } else {
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(s.headers);
        sheet.getRange(1, 1, 1, s.headers.length).setFontWeight("bold").setBackground("#e2e8f0");
      }
    }
  });

  // Isi data default Settings jika kosong
  let settingsSheet = ss.getSheetByName("SETTINGS");
  if (settingsSheet.getLastRow() <= 1) {
    settingsSheet.appendRow(["NamaPerusahaan", "PT IMAM INDONESIA"]);
    settingsSheet.appendRow(["AlamatKantor", "Jl. Raya Industri No. 123, Malang"]);
    settingsSheet.appendRow(["BatasPersetujuan", "10000000"]);
  }

  // Isi data default Kategori jika kosong
  let katSheet = ss.getSheetByName("KATEGORI");
  if (katSheet.getLastRow() <= 1) {
    katSheet.appendRow(["Penjualan", "MASUK", "Aktif"]);
    katSheet.appendRow(["Pendapatan Jasa", "MASUK", "Aktif"]);
    katSheet.appendRow(["Lain-lain", "MASUK", "Aktif"]);
    katSheet.appendRow(["Operasional", "KELUAR", "Aktif"]);
    katSheet.appendRow(["Gaji", "KELUAR", "Aktif"]);
    katSheet.appendRow(["Transportasi", "KELUAR", "Aktif"]);
  }
  
  return "Database berhasil disiapkan!";
}


/**
 * =================================================================
 * 3. FUNGSI TRANSAKSI (KAS MASUK & KELUAR)
 * =================================================================
 */

function simpanKasMasuk(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("TRANSAKSI");
    
    const tgl = new Date();
    const dateStr = Utilities.formatDate(tgl, "Asia/Jakarta", "yyyyMMdd");
    const rand = Math.floor(100 + Math.random() * 900);
    const id = "KM-" + dateStr + "-" + rand;
    
    const timestamp = Utilities.formatDate(tgl, "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
    
    sheet.appendRow([
      id, data.tanggal, "MASUK", data.kategori, data.keterangan, Number(data.nominal), 
      data.metode || "Kas Utama", data.sumber || "-", "Berhasil", data.user || "Admin", timestamp, timestamp
    ]);

    addAuditLog(data.user || "Admin", "Input Kas Masuk " + id, "Berhasil");
    
    return { success: true, message: "Kas Masuk berhasil dicatat", id: id };
  } catch(e) {
    addAuditLog(data.user || "Sistem", "Gagal input Kas Masuk: " + e.message, "Gagal");
    return { success: false, message: e.message };
  }
}

function simpanKasKeluar(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("TRANSAKSI");
    const settings = getSettingsObject();
    
    const limit = Number(settings["BatasPersetujuan"]) || 10000000;
    const nominal = Number(data.nominal);
    let status = nominal > limit ? "Menunggu Persetujuan" : "Berhasil";
    
    const tgl = new Date();
    const dateStr = Utilities.formatDate(tgl, "Asia/Jakarta", "yyyyMMdd");
    const rand = Math.floor(100 + Math.random() * 900);
    const id = "KK-" + dateStr + "-" + rand;
    const timestamp = Utilities.formatDate(tgl, "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
    
    sheet.appendRow([
      id, data.tanggal, "KELUAR", data.kategori, data.keterangan, nominal, 
      data.rekening || "Kas Utama", data.penerima || "-", status, data.user || "Admin", timestamp, timestamp
    ]);

    addAuditLog(data.user || "Admin", "Input Kas Keluar " + id, "Berhasil");
    
    return { success: true, message: `Kas Keluar dicatat. Status: ${status}`, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}


/**
 * =================================================================
 * 4. FUNGSI PENGAMBILAN DATA (DASHBOARD & TRANSAKSI)
 * =================================================================
 */

function getAllData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Ambil Transaksi
    const sheetTrx = ss.getSheetByName("TRANSAKSI");
    const trxData = sheetToJSON(sheetTrx).reverse();
    
    // Ambil Kategori
    const sheetKat = ss.getSheetByName("KATEGORI");
    const katData = sheetToJSON(sheetKat);

    // Ambil Audit Log
    const sheetAudit = ss.getSheetByName("AUDIT_LOG");
    const auditData = sheetToJSON(sheetAudit).reverse();

    // Hitung Dashboard Summary
    let totalPemasukan = 0;
    let totalPengeluaran = 0;
    let trxHariIni = 0;
    const today = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");

    trxData.forEach(function(row) {
      if(row.Status === "Berhasil" || row.Status === "Disetujui") {
        if(row.Jenis === "MASUK") {
          totalPemasukan += Number(row.Nominal);
        } else if(row.Jenis === "KELUAR") {
          totalPengeluaran += Number(row.Nominal);
        }
      }
      
      let tglTrx = row.Tanggal;
      if (tglTrx instanceof Date) {
        tglTrx = Utilities.formatDate(tglTrx, "Asia/Jakarta", "yyyy-MM-dd");
      }
      if(tglTrx === today) {
        trxHariIni++;
      }
    });

    return {
      success: true,
      transaksi: trxData,
      kategori: katData,
      auditLog: auditData,
      summary: {
        saldoKas: totalPemasukan - totalPengeluaran,
        totalPemasukan: totalPemasukan,
        totalPengeluaran: totalPengeluaran,
        transaksiHariIni: trxHariIni
      }
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}


/**
 * =================================================================
 * 5. MANAJEMEN AKSI TRANSAKSI & KATEGORI
 * =================================================================
 */

function updateStatusTransaksi(id, newStatus, user) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TRANSAKSI");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf("ID");
    const statusIndex = headers.indexOf("Status");
    const updateIndex = headers.indexOf("DiperbaruiPada");
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === id) {
        sheet.getRange(i + 1, statusIndex + 1).setValue(newStatus);
        sheet.getRange(i + 1, updateIndex + 1).setValue(Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss"));
        addAuditLog(user || "Admin", `Ubah status transaksi ${id} menjadi ${newStatus}`, "Berhasil");
        return { success: true, message: `Status ${id} diperbarui menjadi ${newStatus}.` };
      }
    }
    return { success: false, message: "Transaksi tidak ditemukan." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function hapusTransaksi(id, user) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TRANSAKSI");
    const data = sheet.getDataRange().getValues();
    const idIndex = data[0].indexOf("ID");
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === id) {
        sheet.deleteRow(i + 1);
        addAuditLog(user || "Admin", `Menghapus transaksi ${id}`, "Berhasil");
        return { success: true, message: `Transaksi ${id} berhasil dihapus.` };
      }
    }
    return { success: false, message: "Transaksi tidak ditemukan." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function tambahKategori(nama, jenis) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KATEGORI");
    sheet.appendRow([nama, jenis, "Aktif"]);
    addAuditLog("Admin", `Menambah kategori baru: ${nama} (${jenis})`, "Berhasil");
    return { success: true, message: "Kategori berhasil ditambahkan" };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function simpanPengaturanPerusahaan(nama, alamat) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SETTINGS");
    const data = sheet.getDataRange().getValues();
    
    for(let i = 1; i < data.length; i++) {
      if(data[i][0] === "NamaPerusahaan") sheet.getRange(i+1, 2).setValue(nama);
      if(data[i][0] === "AlamatKantor") sheet.getRange(i+1, 2).setValue(alamat);
    }
    addAuditLog("Admin", "Memperbarui informasi perusahaan", "Berhasil");
    return { success: true, message: "Pengaturan berhasil disimpan" };
  } catch(e) {
    return { success: false, message: e.message };
  }
}


/**
 * =================================================================
 * 6. UTILITIES & AUDIT LOG
 * =================================================================
 */

function addAuditLog(user, activity, status) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AUDIT_LOG");
    if(!sheet) return;
    const timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([timestamp, user, activity, status]);
  } catch(e) {
    Logger.log("Log error: " + e.message);
  }
}

function getSettingsObject() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SETTINGS");
  if(!sheet) return {};
  const data = sheet.getDataRange().getValues();
  let settings = {};
  for(let i = 1; i < data.length; i++) {
    settings[data[i][0]] = data[i][1];
  }
  return settings;
}

function sheetToJSON(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    let obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = data[i][j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, "Asia/Jakarta", "yyyy-MM-dd");
      }
      obj[headers[j]] = val;
    }
    result.push(obj);
  }
  return result;
}