const APP_CONFIG = {
  TIMEZONE: "GMT+7",
  SHEETS: {
    PEMASUKAN: "Pemasukan",
    ARUS_KAS: "Arus Kas",
    GAJI: "Gaji Karyawan",
    STOK: "Master_Stok",
    MUTASI_STOK: "Mutasi_Stok",
    AUDIT: "Audit_Log"
  }
};

function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  template.dataBukuJson = ambilDataTerbaruJson();
  return template.evaluate()
      .setTitle('Pembukuan Bisnis')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// Fungsi untuk menyimpan database karyawan secara permanen di server Google
function simpanPengaturanGaji(config) {
  try {
    var properties = PropertiesService.getScriptProperties();
    properties.setProperty("CONFIG_GAJI_KARYAWAN", JSON.stringify(config));
    return { success: true };
  } catch (err) {
    throw new Error("Gagal menyimpan profil karyawan di server: " + err.message);
  }
}

// Fungsi untuk membaca database karyawan yang tersimpan di server
function ambilPengaturanGajiTersimpan() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty("CONFIG_GAJI_KARYAWAN");
    if (raw) {
      return JSON.parse(raw);
    }
  } catch (e) {}
  return null;
}

function pastikanSheetAda(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9ead3");
  }
  return sheet;
}

function getBEPDefault() {
  return {
    margin: 25,
    jumlahHari: 30,
    baselinePct: 10,
    baseline: 0,
    biayaTetapTotal: 11768000,
    gaji: [
      { nama: "Gaji 2 Karyawan", totalHari: 3500000, jumlahHari: 1, totalNominal: 3500000 },
      { nama: "Gaji Tangan Kanan", totalHari: 4768000, jumlahHari: 1, totalNominal: 4768000 }
    ],
    depresiasi: [
      { nama: "Printer L 1300", masaPakai: 36, hargaBarang: 4500000, depresiasiBulanan: 125000 },
      { nama: "Komputer", masaPakai: 60, hargaBarang: 10406500, depresiasiBulanan: 173442 },
      { nama: "Laptop", masaPakai: 36, hargaBarang: 4000000, depresiasiBulanan: 111111 },
      { nama: "Mesin Cutting K 1200", masaPakai: 60, hargaBarang: 16493000, depresiasiBulanan: 274883 },
      { nama: "Mesin Cutting Jinka", masaPakai: 36, hargaBarang: 3359580, depresiasiBulanan: 93322 }
    ],
    operasional: [
      { nama: "Sewa Toko", nominalHari: 83500, jumlahHari: 30, totalNominal: 2505000 },
      { nama: "Listrik & Air", nominalHari: 12000, jumlahHari: 30, totalNominal: 360000 },
      { nama: "Iuran RT", nominalHari: 1200, jumlahHari: 30, totalNominal: 36000 },
      { nama: "Bensin", nominalHari: 6700, jumlahHari: 30, totalNominal: 201000 }
    ]
  };
}

function ambilDataBEP() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Setting BEP");
  if (!sheet) {
    var def = getBEPDefault();
    simpanDataBEP(def);
    return def;
  }
  try {
    var val = sheet.getRange(1, 1).getValue();
    if (val) return JSON.parse(val);
  } catch(e) {}
  return getBEPDefault();
}

function simpanDataBEP(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Setting BEP");
  if (!sheet) sheet = ss.insertSheet("Setting BEP");
  sheet.getRange(1, 1).setValue(JSON.stringify(config));
  SpreadsheetApp.flush();
  return ambilDataTerbaruJson();
}

function getGajiSettingDefault() {
  return {
    defaultJamStandar: 10,
    defaultGajiStandar: 110000,
    defaultTarifLembur: 20000,
    daftarKaryawan: [
      { nama: "Hani", jamStandar: 10, gajiStandar: 99000, tarifLembur: 20000 },
      { nama: "Sunlie", jamStandar: 10, gajiStandar: 99000, tarifLembur: 20000 },
      { nama: "Anis", jamStandar: 9, gajiStandar: 117000, tarifLembur: 25000 },
      { nama: "Noor", jamStandar: 10, gajiStandar: 100000, tarifLembur: 20000 },
      { nama: "Agung", jamStandar: 10, gajiStandar: 50000, tarifLembur: 10000 }
    ]
  };
}

function ambilSettingGaji() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Setting Gaji");
  if (!sheet) {
    var def = getGajiSettingDefault();
    simpanSettingGaji(def);
    return def;
  }
  try {
    var val = sheet.getRange(1, 1).getValue();
    if (val) return JSON.parse(val);
  } catch(e) {}
  return getGajiSettingDefault();
}

function simpanSettingGaji(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Setting Gaji");
  if (!sheet) sheet = ss.insertSheet("Setting Gaji");
  sheet.getRange(1, 1).setValue(JSON.stringify(config));
  SpreadsheetApp.flush();
  return ambilDataTerbaruJson();
}

function ambilSemuaData(modeFull) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;

  var tahunIni = new Date().getFullYear();
  var cutoffTgl = String(tahunIni) + "-01-01";

  // OPTIMASI: Cache semua sheet ke dalam memori agar tidak mencari berulang-ulang
  var allSheets = ss.getSheets();
  var sheetMap = {};
  for (var s = 0; s < allSheets.length; s++) {
    sheetMap[allSheets[s].getName()] = allSheets[s];
  }

  function getSheetDataOptimized(name, tglColIdx) {
    var sheet = sheetMap[name];
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return [];

    var values = sheet.getDataRange().getValues();
    var result = [];

    if (values.length > 0) {
      var header = values[0].slice();
      header.push(1);
      result.push(header);
    }

    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var mappedRow = new Array(row.length + 1);

      for (var c = 0; c < row.length; c++) {
        var cell = row[c];
        if (cell instanceof Date) {
          if (cell.getFullYear() <= 1900) {
            mappedRow[c] = Utilities.formatDate(cell, "GMT+7", "HH:mm");
          } else {
            mappedRow[c] = Utilities.formatDate(cell, "GMT+7", "yyyy-MM-dd");
          }
        } else {
          mappedRow[c] = cell;
        }
      }

      mappedRow[row.length] = i + 1; // Indeks baris asli untuk fungsi edit/hapus

      if (!modeFull && tglColIdx !== undefined && tglColIdx !== null) {
        var valTgl = String(mappedRow[tglColIdx] || "");
        if (valTgl && valTgl < cutoffTgl) continue;
      }
      result.push(mappedRow);
    }
    return result;
  }

  // Prioritaskan mengambil database karyawan dari script properties (agar data baru tidak hilang saat refresh)
  var configKaryawan = ambilPengaturanGajiTersimpan();
  if (!configKaryawan) {
    configKaryawan = ambilSettingGaji();
  }

  return {
  isFullHistory: !!modeFull,
  pemasukan: getSheetDataOptimized("Pemasukan", 1),
  shopee: getSheetDataOptimized("Pembelian Shopee", 1),
  agen: getSheetDataOptimized("Pembelian Agen", 1),
  pengeluaranLain: getSheetDataOptimized("Pengeluaran Lain", 1),
  iklan: getSheetDataOptimized("Iklan", 1),
  laporan: getSheetDataOptimized("Total & Laporan", null),
  arusKas: getSheetDataOptimized("Arus Kas", 0),
  gajiKaryawan: getSheetDataOptimized("Gaji Karyawan", 0),
  penyesuaianGaji: getSheetDataOptimized("Penyesuaian Gaji", 0),
  invoice: getSheetDataOptimized("Riwayat_Invoice", 0),
  detailPenjualan: getSheetDataOptimized("Detail_Penjualan", 0),
  // --- TAMBAHKAN DARI SINI ---
  mutasiStok: getSheetDataOptimized("Mutasi_Stok", 0),
  masterStok: getSheetDataOptimized("Master_Stok", null),
  // --- BATAS TAMBAHAN ---
  bepConfig: ambilDataBEP(),
  gajiConfig: configKaryawan,
  skuAliasMap: buatIndexAliasSku_()
  };
}

function test_cekDataMutasiLabaRugi() {
  try {
    var data = ambilSemuaData(false);
    if (data.mutasiStok && data.masterStok) {
      Logger.log("BERHASIL: Data Mutasi Stok (" + data.mutasiStok.length + " baris) dan Master Stok (" + data.masterStok.length + " baris) siap dikirim ke frontend.");
    } else {
      Logger.log("GAGAL: Data Mutasi/Master Stok tidak ditemukan.");
    }
  } catch (e) {
    Logger.log("Error: " + e.message);
  }
}

// HELPER: Ekstrak File ID dari URL Google Drive dan pindahkan ke Trash
function hapusFilePdfLamaDariUrl(urlDrive) {
  if (!urlDrive || typeof urlDrive !== "string") return;
  try {
    var match = urlDrive.match(/[-\w]{25,}/);
    if (match && match[0]) {
      DriveApp.getFileById(match[0]).setTrashed(true);
    }
  } catch (e) {
    Logger.log("File lama sudah tidak ada atau gagal dihapus: " + e.message);
  }
}

function simpanInvoiceKeDrive(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userAktif = data.user || "Owner";
  
  // 1. Cari atau Buat Folder "Arsip Invoice OFS" di Google Drive
  var namaFolder = "Arsip Invoice OFS";
  var folderIter = DriveApp.getFoldersByName(namaFolder);
  var folderTarget = folderIter.hasNext() ? folderIter.next() : DriveApp.createFolder(namaFolder);
  
  // 2. Dekode Text Base64 kembali menjadi File PDF Asli
  var noInvBersih = String(data.noInvoice || "").replace(/\//g, "-");
  var namaFile = "INV_" + noInvBersih + "_" + String(data.namaPelanggan || "").replace(/\s+/g, '_') + ".pdf";
  var blob = Utilities.newBlob(Utilities.base64Decode(data.fileBase64), 'application/pdf', namaFile);
  
  // 3. Simpan File ke Drive dan dapatkan Link-nya
  var filePdf = folderTarget.createFile(blob);
  var linkFileTersimpan = filePdf.getUrl();
  
  // 4. Pastikan Sheet Riwayat_Invoice memiliki 7 kolom lengkap
  var sheetInvoice = pastikanSheetAda(ss, "Riwayat_Invoice", [
    "Tanggal", "No Invoice", "Pelanggan", "Total Tagihan", "Link PDF", "Status", "Data_JSON"
  ]);
  if (sheetInvoice.getLastColumn() < 7) {
    sheetInvoice.getRange(1, 6, 1, 2).setValues([["Status", "Data_JSON"]]).setFontWeight("bold").setBackground("#d9ead3");
  }

  var editRow = Number(data.editRowIndex) || 0;
  var statusPilihan = data.statusBayar || "Belum Lunas";

  // JIKA MODE EDIT: Hapus file PDF lama di Drive dan timpa barisnya
  if (editRow > 1 && editRow <= sheetInvoice.getLastRow()) {
    var linkLama = String(sheetInvoice.getRange(editRow, 5).getValue() || "");
    hapusFilePdfLamaDariUrl(linkLama);

    sheetInvoice.getRange(editRow, 1, 1, 7).setValues([[
      data.tanggal,
      data.noInvoice,
      data.namaPelanggan,
      Number(data.totalTagihan) || 0,
      linkFileTersimpan,
      statusPilihan,
      data.rawJson || ""
    ]]);
    catatAuditLog("EDIT INVOICE", "Riwayat Invoice", data.noInvoice, "Revisi PDF & Data (" + data.namaPelanggan + ")", userAktif);
  } else {
    // JIKA MODE BARU: Tambahkan baris baru
    sheetInvoice.appendRow([
      data.tanggal, 
      data.noInvoice, 
      data.namaPelanggan, 
      Number(data.totalTagihan) || 0, 
      linkFileTersimpan,
      statusPilihan,
      data.rawJson || ""
    ]);
    catatAuditLog("BUAT INVOICE", "Riwayat Invoice", "-", data.noInvoice + " - " + data.namaPelanggan, userAktif);
  }
  
  SpreadsheetApp.flush();
  return ambilDataTerbaruJson();
}

function hapusInvoiceDrive(rowIndex, user) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Riwayat_Invoice");
  if (!sheet) throw new Error("Sheet Riwayat_Invoice tidak ditemukan.");

  var r = Number(rowIndex);
  if (r <= 1 || r > sheet.getLastRow()) throw new Error("Baris invoice tidak valid.");

  var rowVals = sheet.getRange(r, 1, 1, Math.min(sheet.getLastColumn(), 7)).getValues()[0];
  var noInv = String(rowVals[1] || "");
  var pelanggan = String(rowVals[2] || "");
  var linkPdf = String(rowVals[4] || "");

  // Pindahkan PDF di Google Drive ke Sampah (Trash)
  hapusFilePdfLamaDariUrl(linkPdf);

  // Hapus baris dari Sheet
  sheet.deleteRow(r);
  catatAuditLog("HAPUS INVOICE", "Riwayat Invoice", noInv + " (" + pelanggan + ")", "Dihapus beserta PDF di Drive", user || "Owner");

  SpreadsheetApp.flush();
  return ambilDataTerbaruJson();
}

function ubahStatusInvoiceServer(rowIndex, statusBaru, user) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Riwayat_Invoice");
  if (!sheet) throw new Error("Sheet Riwayat_Invoice tidak ditemukan.");

  var r = Number(rowIndex);
  if (r <= 1 || r > sheet.getLastRow()) throw new Error("Baris invoice tidak valid.");

  if (sheet.getLastColumn() < 6) {
    sheet.getRange(1, 6).setValue("Status").setFontWeight("bold").setBackground("#d9ead3");
  }
  sheet.getRange(r, 6).setValue(statusBaru);

  var noInv = String(sheet.getRange(r, 2).getValue() || "");
  catatAuditLog("STATUS INVOICE", "Riwayat Invoice", noInv, "Status diubah ke: " + statusBaru, user || "Owner");

  SpreadsheetApp.flush();
  return ambilDataTerbaruJson();
}

function ambilDataTerbaruJson(modeFull) {
  SpreadsheetApp.flush();
  return JSON.stringify(ambilSemuaData(modeFull));
}

function ambilDataHistorisPenuhJson() {
  return ambilDataTerbaruJson(true);
}

function arsipkanDataTahun(tahunTarget) {
  var tahunStr = String(tahunTarget).trim();
  if (!tahunStr || tahunStr.length !== 4) throw new Error("Format tahun harus 4 digit (misal: 2025)");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var daftarSheet = [
    { name: "Pemasukan", tglIdx: 1 },
    { name: "Pembelian Shopee", tglIdx: 1 },
    { name: "Pembelian Agen", tglIdx: 1 },
    { name: "Pengeluaran Lain", tglIdx: 1 },
    { name: "Iklan", tglIdx: 1 },
    { name: "Arus Kas", tglIdx: 0 }
  ];

  var totalBarisDipindah = 0;

  daftarSheet.forEach(function(item) {
    var sheet = ss.getSheetByName(item.name);
    if (!sheet || sheet.getLastRow() <= 1) return;

    var data = sheet.getDataRange().getValues();
    var header = data[0];
    var barisTetap = [header];
    var barisArsip = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rawDate = row[item.tglIdx];
      var tglStr = "";
      if (rawDate instanceof Date) {
        tglStr = Utilities.formatDate(rawDate, "GMT+7", "yyyy-MM-dd");
      } else {
        tglStr = String(rawDate || "");
      }

      if (tglStr.startsWith(tahunStr)) {
        barisArsip.push(row);
      } else {
        barisTetap.push(row);
      }
    }

    if (barisArsip.length > 0) {
      var namaSheetArsip = "Arsip_" + item.name.replace(/\s+/g, '_') + "_" + tahunStr;
      var sheetArsip = pastikanSheetAda(ss, namaSheetArsip, header);
      sheetArsip.getRange(sheetArsip.getLastRow() + 1, 1, barisArsip.length, header.length).setValues(barisArsip);

      sheet.clearContents();
      sheet.getRange(1, 1, barisTetap.length, header.length).setValues(barisTetap);
      totalBarisDipindah += barisArsip.length;
    }
  });

  SpreadsheetApp.flush();
  catatAuditLog("ARSIP", "Database", "Arsip Tahun " + tahunStr, totalBarisDipindah + " baris transaksi dipindahkan ke sheet arsip.", "Owner");

  return {
    status: "success",
    tahun: tahunStr,
    totalDipindah: totalBarisDipindah,
    freshDataJson: ambilDataTerbaruJson(false)
  };
}

function catatAuditLog(aksi, kategori, dataLama, dataBaru, userPenginput) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = pastikanSheetAda(ss, "Audit_Log", ["Waktu", "User", "Aksi", "Kategori", "Data Lama", "Data Baru"]);
  var user = userPenginput || "Sistem";
  var waktu = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
  sheet.appendRow([waktu, user, aksi, kategori, String(dataLama), String(dataBaru)]);
}

function simpanDataWeb(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nominal = Number(data.nominal) || 0;
  var userAktif = data.user || "Staf";
  
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  
  try {
    var waktuSekarang = Utilities.formatDate(new Date(), "GMT+7", "HH:mm");
    if (data.jam && String(data.jam).includes(':')) waktuSekarang = data.jam;
    var jamTeks = "'" + waktuSekarang;
    var tokoVal = (data.toko && String(data.toko).trim() !== '') ? String(data.toko).trim() : "-";
    
    var sheet = ss.getSheetByName(data.kategori);
    if (!sheet) {
      sheet = pastikanSheetAda(ss, data.kategori, ["Nama", "Tanggal", "Nominal", "Toko", "Jam"]);
    }
    
    var tokoField = (data.kategori === "Pemasukan" || data.kategori === "Pembelian Shopee" || data.kategori === "Iklan") ? tokoVal : "-";
    sheet.appendRow([data.nama, data.tanggal, nominal, tokoField, jamTeks]);

    var sheetAK = pastikanSheetAda(ss, "Arus Kas", ["Tanggal", "Keterangan", "Tipe", "Nominal", "Jam"]);
    var tipeKas = (data.kategori === "Pemasukan") ? "Kas Masuk" : "Kas Keluar";
    var ketKas = "[" + data.kategori + "] " + data.nama;
    sheetAK.appendRow([data.tanggal, ketKas, tipeKas, nominal, jamTeks]);

    catatAuditLog("INPUT", data.kategori, "-", data.nama + " (Rp " + nominal.toLocaleString('id-ID') + " di " + tokoField + ")", userAktif);

    SpreadsheetApp.flush();
    return ambilDataTerbaruJson();
  } finally {
    lock.releaseLock();
  }
}

function editTransaksi(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nominal = Number(data.nominal) || 0;
  var jamTeks = "'" + data.jam;
  var row = Number(data.rowIndex);
  var tokoVal = (data.toko && String(data.toko).trim() !== '') ? String(data.toko).trim() : "-";
  var userAktif = data.user || "Staf";
  
  // Deteksi jika pengguna memindahkan kategori transaksi saat edit
  var kategoriLama = data.kategoriLama || data.kategori;

  var sheetLama = ss.getSheetByName(kategoriLama);
  if (!sheetLama) throw new Error("Sheet lama tidak ditemukan: " + kategoriLama);
  
  var colCount = Math.min(sheetLama.getLastColumn(), 6);
  var rowValuesLama = sheetLama.getRange(row, 1, 1, colCount).getValues()[0];
  var dataLamaTeks = rowValuesLama.filter(function(v) { return v !== "" && v !== null; }).join(" | ");
  
  // Ekstrak data lama untuk pencarian di Arus Kas
  var tglLama = Utilities.formatDate(new Date(rowValuesLama[1]), "GMT+7", "yyyy-MM-dd");
  var nominalLama = Number(rowValuesLama[2]) || 0;
  var namaLama = String(rowValuesLama[0] || "").trim();

  // 1. UPDATE DI SHEET TRANSAKSI (Sesuai Kategori)
  if (kategoriLama === data.kategori) {
    // Jika tidak pindah sheet kategori
    if (data.kategori === "Pemasukan" || data.kategori === "Pembelian Shopee" || data.kategori === "Iklan") {
      sheetLama.getRange(row, 1, 1, 5).setValues([[data.nama, data.tanggal, nominal, tokoVal, jamTeks]]);
    } else {
      sheetLama.getRange(row, 1, 1, 5).setValues([[data.nama, data.tanggal, nominal, "-", jamTeks]]);
    }
  } else {
    // Jika pindah kategori: Hapus dari sheet lama, tambah ke sheet baru
    sheetLama.deleteRow(row);
    var sheetBaru = pastikanSheetAda(ss, data.kategori, ["Nama", "Tanggal", "Nominal", "Toko", "Jam"]);
    var tokoField = (data.kategori === "Pemasukan" || data.kategori === "Pembelian Shopee" || data.kategori === "Iklan") ? tokoVal : "-";
    sheetBaru.appendRow([data.nama, data.tanggal, nominal, tokoField, jamTeks]);
  }

  // 2. UPDATE DI SHEET ARUS KAS
  var sheetAK = ss.getSheetByName("Arus Kas");
  if (sheetAK && sheetAK.getLastRow() > 1) {
    var akVals = sheetAK.getDataRange().getValues();
    for (var k = akVals.length - 1; k >= 1; k--) {
      var tglK = Utilities.formatDate(new Date(akVals[k][0]), "GMT+7", "yyyy-MM-dd");
      var nomK = Number(akVals[k][3]) || 0;
      var ketK = String(akVals[k][1] || "").trim();
      
      // Cari baris kas yang pasangannya cocok dengan data lama
      if (tglK === tglLama && nomK === nominalLama && ketK.includes(namaLama)) {
        var tipeKasBaru = (data.kategori === "Pemasukan") ? "Kas Masuk" : "Kas Keluar";
        var ketKasBaru = "[" + data.kategori + "] " + data.nama;
        sheetAK.getRange(k + 1, 1, 1, 5).setValues([[data.tanggal, ketKasBaru, tipeKasBaru, nominal, jamTeks]]);
        break; // Cukup perbarui satu transaksi yang cocok
      }
    }
  }

  var dataBaruTeks = [data.nama, data.tanggal, "Rp " + nominal.toLocaleString('id-ID'), tokoVal, data.jam].join(" | ");
  catatAuditLog("EDIT", data.kategori, dataLamaTeks, dataBaruTeks, userAktif);

  SpreadsheetApp.flush();
  return ambilDataTerbaruJson();
}

function hapusTransaksi(sheetName, rowIndex, user) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var userAktif = user || "Staf";
  if (!sheet) throw new Error("Sheet tidak ditemukan: " + sheetName);

  var r = Number(rowIndex);
  if (r <= 1 || r > sheet.getLastRow()) throw new Error("Indeks baris transaksi tidak valid.");

  var rowValues = sheet.getRange(r, 1, 1, Math.min(sheet.getLastColumn(), 8)).getValues()[0];
  var dataLamaTeks = rowValues.filter(function(v) { return v !== "" && v !== null; }).join(" | ");

  // Simpan payload baris asli untuk keperluan pemulihan (Audit Log Restore)
  var payloadBackup = JSON.stringify({
    sheet: sheetName,
    values: rowValues
  });

  // Variabel bantu untuk sinkronisasi ke Gudang & POS
  var tglSync = "";
  var jamSync = "";
  var ketSync = "";

  // =========================================================================
  // KASUS A: PENGHAPUSAN BERASAL DARI TABEL ARUS KAS
  // =========================================================================
  if (sheetName === "Arus Kas") {
    var tglKas = (rowValues[0] instanceof Date) ? Utilities.formatDate(rowValues[0], "GMT+7", "yyyy-MM-dd") : String(rowValues[0]).split('T')[0];
    var ketKas = String(rowValues[1] || "").trim();
    var nominalKas = Number(rowValues[3]) || 0;
    var jamKas = String(rowValues[4] || "").replace(/'/g, "").trim();

    // Bersihkan prefix kategori
    var targetSheetName = "";
    var searchNama = ketKas;

    if (ketKas.indexOf("[") === 0 && ketKas.indexOf("]") > 1) {
      targetSheetName = ketKas.substring(1, ketKas.indexOf("]"));
      searchNama = ketKas.substring(ketKas.indexOf("]") + 1).trim();
    }

    tglSync = tglKas;
    jamSync = jamKas;
    ketSync = searchNama;

    // Cari dan hapus di sheet sumbernya (Laporan)
    var sheetsDicari = (targetSheetName && ss.getSheetByName(targetSheetName)) ? [targetSheetName] : ["Pemasukan", "Pembelian Shopee", "Pembelian Agen", "Pengeluaran Lain", "Iklan"];
    sheetsDicari.forEach(function(sName) {
      var sTarget = ss.getSheetByName(sName);
      if (sTarget && sTarget.getLastRow() > 1) {
        var dVals = sTarget.getDataRange().getValues();
        for (var i = dVals.length - 1; i >= 1; i--) {
          var tglD = (dVals[i][1] instanceof Date) ? Utilities.formatDate(dVals[i][1], "GMT+7", "yyyy-MM-dd") : String(dVals[i][1]).split('T')[0];
          var nomD = Number(dVals[i][2]) || 0;
          var namaD = String(dVals[i][0] || "").trim();

          if (tglD === tglKas && (nomD === nominalKas || Math.abs(nomD) === nominalKas) && (namaD.includes(searchNama) || searchNama.includes(namaD))) {
            sTarget.deleteRow(i + 1);
            break;
          }
        }
      }
    });

    // PENCARIAN BALIK: Bersihkan Master Gaji Karyawan jika yang dihapus adalah Beban Gaji
    if (ketKas.includes("Beban Gaji")) {
      var sGajiKas = ss.getSheetByName("Gaji Karyawan");
      if (sGajiKas && sGajiKas.getLastRow() > 1) {
        var gkVals = sGajiKas.getDataRange().getValues();
        for (var gk = gkVals.length - 1; gk >= 1; gk--) {
          var tglGK = Utilities.formatDate(new Date(gkVals[gk][0]), "GMT+7", "yyyy-MM-dd");
          var nomGK = Number(gkVals[gk][6]) || 0;
          if (tglGK === tglKas && nomGK === nominalKas && ketKas.includes(String(gkVals[gk][2] || "").trim())) {
            sGajiKas.deleteRow(gk + 1);
            break;
          }
        }
      }
    }

  // =========================================================================
  // KASUS B: PENGHAPUSAN BERASAL DARI LAPORAN TRANSAKSI (PEMASUKAN / PENGELUARAN)
  // =========================================================================
  } else if (sheetName !== "Gaji Karyawan") {
    var tglSumber = (rowValues[1] instanceof Date) ? Utilities.formatDate(rowValues[1], "GMT+7", "yyyy-MM-dd") : String(rowValues[1]).split('T')[0];
    var namaSumber = String(rowValues[0] || "").trim();
    var nominalSumber = Number(rowValues[2]) || 0;
    var jamSumber = String(rowValues[4] || "").replace(/'/g, "").trim();

    tglSync = tglSumber;
    jamSync = jamSumber;
    ketSync = namaSumber.split(" [Net]")[0].trim();

    var sKas = ss.getSheetByName("Arus Kas");
    if (sKas && sKas.getLastRow() > 1) {
      var kVals = sKas.getDataRange().getValues();
      for (var k = kVals.length - 1; k >= 1; k--) {
        var tglK = (kVals[k][0] instanceof Date) ? Utilities.formatDate(kVals[k][0], "GMT+7", "yyyy-MM-dd") : String(kVals[k][0]).split('T')[0];
        var nomK = Number(kVals[k][3]) || 0;
        var ketK = String(kVals[k][1] || "").trim();

        if (tglK === tglSumber && (nomK === nominalSumber || nomK === Math.abs(nominalSumber)) && (ketK.includes(namaSumber) || ketK.includes(ketSync))) {
          sKas.deleteRow(k + 1);
          break;
        }
      }
    }

    // PENCARIAN BALIK: Bersihkan Master Gaji Karyawan jika Pengeluaran yang dihapus adalah Beban Gaji
    if (sheetName === "Pengeluaran Lain" && namaSumber.includes("Beban Gaji")) {
      var sGaji = ss.getSheetByName("Gaji Karyawan");
      if (sGaji && sGaji.getLastRow() > 1) {
        var gVals = sGaji.getDataRange().getValues();
        for (var g = gVals.length - 1; g >= 1; g--) {
          var tglG = Utilities.formatDate(new Date(gVals[g][0]), "GMT+7", "yyyy-MM-dd");
          var nomG = Number(gVals[g][6]) || 0;
          if (tglG === tglSumber && nomG === nominalSumber && namaSumber.includes(String(gVals[g][2] || "").trim())) {
            sGaji.deleteRow(g + 1);
            break;
          }
        }
      }
    }

  // =========================================================================
  // KASUS C: PENGHAPUSAN BERASAL DARI RIWAYAT GAJI KARYAWAN
  // =========================================================================
  } else if (sheetName === "Gaji Karyawan") {
    var tglGaji = Utilities.formatDate(new Date(rowValues[0]), "GMT+7", "yyyy-MM-dd");
    var namaKaryawan = String(rowValues[2] || "").trim();
    var totalGaji = Number(rowValues[6]) || 0;

    // Bersihkan dari Pengeluaran Lain
    var sPeng = ss.getSheetByName("Pengeluaran Lain");
    if (sPeng && sPeng.getLastRow() > 1) {
      var pVals = sPeng.getDataRange().getValues();
      for (var p = pVals.length - 1; p >= 1; p--) {
        var tglP = Utilities.formatDate(new Date(pVals[p][1]), "GMT+7", "yyyy-MM-dd");
        var nomP = Number(pVals[p][2]) || 0;
        var namaP = String(pVals[p][0] || "");
        if (tglP === tglGaji && nomP === totalGaji && namaP.includes(namaKaryawan)) {
          sPeng.deleteRow(p + 1);
          break;
        }
      }
    }

    // Bersihkan dari Arus Kas
    var sKasGaji = ss.getSheetByName("Arus Kas");
    if (sKasGaji && sKasGaji.getLastRow() > 1) {
      var kgVals = sKasGaji.getDataRange().getValues();
      for (var kg = kgVals.length - 1; kg >= 1; kg--) {
        var tglKG = Utilities.formatDate(new Date(kgVals[kg][0]), "GMT+7", "yyyy-MM-dd");
        var nomKG = Number(kgVals[kg][3]) || 0;
        var ketKG = String(kgVals[kg][1] || "");
        if (tglKG === tglGaji && nomKG === totalGaji && ketKG.includes(namaKaryawan)) {
          sKasGaji.deleteRow(kg + 1);
          break;
        }
      }
    }
  }

  // =========================================================================
  // SINKRONISASI OTOMATIS KE GUDANG (MASTER STOK, MUTASI, POS, & RETUR)
  // =========================================================================
  if (sheetName !== "Gaji Karyawan" && !ketSync.includes("Beban Gaji") && tglSync && jamSync) {
    var pembeliKunci = ketSync.includes(" - ") ? ketSync.split(" - ")[0].trim() : "";

    // 1. Bersihkan Detail_Penjualan (Nota POS) jika terkait
    var sheetDetail = ss.getSheetByName("Detail_Penjualan");
    if (sheetDetail && sheetDetail.getLastRow() > 1) {
      var dRows = sheetDetail.getDataRange().getValues();
      for (var d = dRows.length - 1; d >= 1; d--) {
        var dTgl = (dRows[d][0] instanceof Date) ? Utilities.formatDate(dRows[d][0], "GMT+7", "yyyy-MM-dd") : String(dRows[d][0]).split('T')[0];
        var dJam = String(dRows[d][1] || "").replace(/'/g, "").trim();
        var dPembeli = String(dRows[d][2] || "").trim();
        if (dTgl === tglSync && dJam === jamSync && (ketSync.includes(dPembeli) || (pembeliKunci && dPembeli.includes(pembeliKunci)))) {
          sheetDetail.deleteRow(d + 1);
          break;
        }
      }
    }

    // 2. Bersihkan Retur_Penjualan jika transaksi yang dihapus adalah Retur Refund
    var sheetRetur = ss.getSheetByName("Retur_Penjualan");
    if (sheetRetur && sheetRetur.getLastRow() > 1) {
      var rRows = sheetRetur.getDataRange().getValues();
      for (var rt = rRows.length - 1; rt >= 1; rt--) {
        var rTgl = (rRows[rt][0] instanceof Date) ? Utilities.formatDate(rRows[rt][0], "GMT+7", "yyyy-MM-dd") : String(rRows[rt][0]).split('T')[0];
        var rJam = String(rRows[rt][1] || "").replace(/'/g, "").trim();
        var rNama = String(rRows[rt][2] || "").trim();
        if (rTgl === tglSync && rJam === jamSync && ketSync.includes(rNama)) {
          sheetRetur.deleteRow(rt + 1);
          break;
        }
      }
    }

    // 3. Kembalikan Stok di Master_Stok & Hapus Riwayat di Mutasi_Stok
    var sheetMutasi = ss.getSheetByName("Mutasi_Stok") || ss.getSheetByName("Mutasi Stok");
    if (sheetMutasi && sheetMutasi.getLastRow() > 1) {
      var mRows = sheetMutasi.getDataRange().getValues();
      var deltaStokMap = {};

      for (var m = mRows.length - 1; m >= 1; m--) {
        var mTgl = (mRows[m][0] instanceof Date) ? Utilities.formatDate(mRows[m][0], "GMT+7", "yyyy-MM-dd") : String(mRows[m][0]).split('T')[0];
        var mJam = String(mRows[m][1] || "").replace(/'/g, "").trim();
        var mTipe = String(mRows[m][2] || "");
        var mTipeLow = mTipe.toLowerCase();
        var mNama = String(mRows[m][3] || "").trim();
        var mQty = Number(mRows[m][4]) || 0;
        var mKet = String(mRows[m][5] || "");

        var isCocokMutasi = (mTgl === tglSync && mJam === jamSync) && (
          ketSync.toLowerCase().includes(mNama.toLowerCase()) ||
          (pembeliKunci && mKet.toLowerCase().includes(pembeliKunci.toLowerCase()))
        );

        if (isCocokMutasi) {
          // Hitung pembalikan stok untuk komponen fisik (bukan baris header paket bundling)
          if (!mTipeLow.includes("penjualan bundling") && !mTipeLow.includes("retur bundling")) {
            if (mTipeLow.includes("keluar") || mTipeLow.includes("jual") || mTipeLow.includes("terpakai") || mTipeLow.includes("kurang") || mTipeLow.includes("rusak")) {
              deltaStokMap[mNama] = (deltaStokMap[mNama] || 0) + mQty;
            } else if (mTipeLow.includes("masuk") || mTipeLow.includes("restock") || mTipeLow.includes("awal") || mTipeLow.includes("lebih") || mTipeLow.includes("kembali")) {
              deltaStokMap[mNama] = (deltaStokMap[mNama] || 0) - mQty;
            }
          }
          sheetMutasi.deleteRow(m + 1);
        }
      }

      // Terapkan perubahan stok kembali ke Master_Stok secara batch
      if (Object.keys(deltaStokMap).length > 0) {
        var sheetStok = ss.getSheetByName("Master_Stok");
        if (sheetStok && sheetStok.getLastRow() > 1) {
          var dataStok = sheetStok.getDataRange().getValues();
          var adaStokBerubah = false;
          for (var s = 1; s < dataStok.length; s++) {
            var namaGudang = String(dataStok[s][1] || "").trim();
            if (deltaStokMap[namaGudang] !== undefined) {
              dataStok[s][3] = Math.max(0, (Number(dataStok[s][3]) || 0) + deltaStokMap[namaGudang]);
              adaStokBerubah = true;
            }
          }
          if (adaStokBerubah) {
            sheetStok.getRange(1, 1, dataStok.length, dataStok[0].length).setValues(dataStok);
          }
        }
      }
    }
  }

  // Hapus baris utama di sheet asalnya
  sheet.deleteRow(r);

  // Catat ke Audit Log dengan backup JSON agar bisa dipulihkan
  catatAuditLog("HAPUS", sheetName, payloadBackup, "Data Dihapus", userAktif);

  SpreadsheetApp.flush();
  return ambilDataTerbaruJson();
}

function pulihkanTransaksiDariAudit(logItem, user, logRowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userAktif = user || "Owner";

  var backupObj = null;
  try {
    backupObj = JSON.parse(logItem.dataLama);
  } catch(e) {
    throw new Error("Data lama tidak memiliki format cadangan yang valid untuk dipulihkan.");
  }

  var targetSheetName = backupObj.sheet;
  var rowValues = backupObj.values;

  // =========================================================
  // PERBAIKAN: MENCEGAH TANGGAL MUNDUR 1 HARI SAAT DIPULIHKAN
  // =========================================================
  // Jika sistem mendeteksi adanya data format waktu global (berisi huruf T dan Z),
  // kembalikan ke zona waktu Indonesia (GMT+7) dan format menjadi YYYY-MM-DD
  for (var i = 0; i < rowValues.length; i++) {
    if (typeof rowValues[i] === "string" && rowValues[i].match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/)) {
      rowValues[i] = Utilities.formatDate(new Date(rowValues[i]), "GMT+7", "yyyy-MM-dd");
    }
  }

  // =========================================================
  // LOGIKA PEMULIHAN PINTAR (SINKRONISASI KE SEMUA SHEET)
  // =========================================================
  
  if (targetSheetName === "Arus Kas") {
    // Jika dihapus dari Arus Kas
    var tglKas = rowValues[0];
    var ketKas = rowValues[1];
    var nomKas = rowValues[3];
    var jamKas = rowValues[4];

    var sheetAK = pastikanSheetAda(ss, "Arus Kas", ["Tanggal", "Keterangan", "Tipe", "Nominal", "Jam"]);
    sheetAK.appendRow(rowValues);

    // Deteksi dari keterangan untuk memulihkan ke sheet asalnya
    var srcSheetName = "";
    var srcNama = ketKas;
    if (ketKas.indexOf("[") === 0 && ketKas.indexOf("]") > 1) {
      srcSheetName = ketKas.substring(1, ketKas.indexOf("]"));
      srcNama = ketKas.substring(ketKas.indexOf("]") + 1).trim();
    }

    if (srcSheetName) {
       var sheetSrc = pastikanSheetAda(ss, srcSheetName, ["Nama", "Tanggal", "Nominal", "Toko", "Jam"]);
       sheetSrc.appendRow([srcNama, tglKas, nomKas, "-", jamKas]);
    } else if (ketKas.includes("Beban Gaji")) {
       var sheetPengeluaran = pastikanSheetAda(ss, "Pengeluaran Lain", ["Nama", "Tanggal", "Nominal", "Toko", "Jam"]);
       sheetPengeluaran.appendRow([ketKas, tglKas, nomKas, "Gaji Karyawan", jamKas]);
       
       var sheetGaji = pastikanSheetAda(ss, "Gaji Karyawan", ["Tanggal", "Hari / Periode", "Nama", "Hadir", "Status", "Rincian", "Total Gaji", "Jam"]);
       var namaPegawai = ketKas.replace("Beban Gaji:", "").split("(")[0].trim();
       var periode = ketKas.includes("(") ? ketKas.split("(")[1].replace(")", "") : "-";
       sheetGaji.appendRow([tglKas, periode, namaPegawai, "-", "-", "Dipulihkan dari Arus Kas", nomKas, jamKas]);
    }

  } else if (targetSheetName === "Gaji Karyawan") {
    // Jika dihapus dari Gaji Karyawan
    var tglGk = rowValues[0];
    var periodeGk = rowValues[1];
    var namaGk = rowValues[2];
    var nomGk = rowValues[6];
    var jamGk = rowValues[7];
    
    var sheetGaji = pastikanSheetAda(ss, "Gaji Karyawan", ["Tanggal", "Hari / Periode", "Nama", "Hadir", "Status", "Rincian", "Total Gaji", "Jam"]);
    sheetGaji.appendRow(rowValues);

    var ketGaji = "Beban Gaji: " + namaGk + " (" + periodeGk + ")";
    
    var sheetPeng = pastikanSheetAda(ss, "Pengeluaran Lain", ["Nama", "Tanggal", "Nominal", "Toko", "Jam"]);
    sheetPeng.appendRow([ketGaji, tglGk, nomGk, "Gaji Karyawan", jamGk]);

    var sheetAK = pastikanSheetAda(ss, "Arus Kas", ["Tanggal", "Keterangan", "Tipe", "Nominal", "Jam"]);
    sheetAK.appendRow([tglGk, ketGaji, "Kas Keluar", nomGk, jamGk]);

  } else {
    // Jika dihapus dari sheet sumber (Pemasukan, Shopee, dll)
    var namaSrc = rowValues[0];
    var tglSrc = rowValues[1];
    var nomSrc = rowValues[2];
    var jamSrc = rowValues[4];

    var sheetSrc = pastikanSheetAda(ss, targetSheetName, ["Nama", "Tanggal", "Nominal", "Toko", "Jam"]);
    sheetSrc.appendRow(rowValues);

    var sheetAK = pastikanSheetAda(ss, "Arus Kas", ["Tanggal", "Keterangan", "Tipe", "Nominal", "Jam"]);
    if (targetSheetName === "Pemasukan") {
       sheetAK.appendRow([tglSrc, "[Pemasukan] " + namaSrc, "Kas Masuk", nomSrc, jamSrc]);
    } else {
       sheetAK.appendRow([tglSrc, "[" + targetSheetName + "] " + namaSrc, "Kas Keluar", nomSrc, jamSrc]);
    }

    if (targetSheetName === "Pengeluaran Lain" && namaSrc.includes("Beban Gaji")) {
       var sheetGaji = pastikanSheetAda(ss, "Gaji Karyawan", ["Tanggal", "Hari / Periode", "Nama", "Hadir", "Status", "Rincian", "Total Gaji", "Jam"]);
       var namaPegawai = namaSrc.replace("Beban Gaji:", "").split("(")[0].trim();
       var periode = namaSrc.includes("(") ? namaSrc.split("(")[1].replace(")", "") : "-";
       sheetGaji.appendRow([tglSrc, periode, namaPegawai, "-", "-", "Dipulihkan", nomSrc, jamSrc]);
    }
  }

  // =========================================================
  // TANDAI LOG AGAR TIDAK BISA DIPULIHKAN BERKALI-KALI
  // =========================================================
  if (logRowIndex) {
     var sheetAudit = ss.getSheetByName("Audit_Log");
     if (sheetAudit) {
        sheetAudit.getRange(logRowIndex, 3).setValue("DIPULIHKAN");
        sheetAudit.getRange(logRowIndex, 6).setValue(String(sheetAudit.getRange(logRowIndex, 6).getValue()) + " (Telah Dipulihkan)");
     }
  }

  SpreadsheetApp.flush();
  return ambilDataTerbaruJson();
}

function simpanGajiKaryawan(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    // PERBAIKAN: Gunakan GMT+7 dan tambahkan tanda petik tunggal (') di depan jam
    // agar Google Sheets membacanya secara paksa sebagai "Teks Biasa".
    var waktuMentah = payload.jam || Utilities.formatDate(new Date(), "GMT+7", "HH:mm");
    var jamSekarang = "'" + waktuMentah; 
    
    var tgl = payload.tanggal || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
    var totalNominal = Number(payload.totalGaji) || 0;
    var userAktif = payload.user || "Owner";

    // 1. Simpan ke sheet Gaji Karyawan
    var sheetGaji = pastikanSheetAda(ss, "Gaji Karyawan", ["Tanggal", "Hari / Periode", "Nama", "Hadir", "Status", "Rincian", "Total Gaji", "Jam"]);
    sheetGaji.appendRow([
      tgl,
      payload.hari,
      payload.nama,
      payload.jamMasuk,
      payload.jamPulang,
      payload.durasiTeks,
      totalNominal,
      jamSekarang
    ]);

    // 2. Jika dicentang, otomatis masukkan ke Pengeluaran Lain & Arus Kas
    if (payload.catatKePembukuan && totalNominal > 0) {
      var keterangan = "Beban Gaji: " + payload.nama + " (" + payload.hari + ")";

      // A. Masuk ke Sheet Pengeluaran Lain
      var sheetPengeluaran = pastikanSheetAda(ss, "Pengeluaran Lain", ["Nama", "Tanggal", "Nominal", "Toko", "Jam"]);
      sheetPengeluaran.appendRow([
        keterangan,
        tgl,
        totalNominal,
        "Gaji Karyawan",
        jamSekarang
      ]);

      // B. Masuk ke Sheet Arus Kas (Mengurangi Kas Keluar)
      var sheetKas = pastikanSheetAda(ss, "Arus Kas", ["Tanggal", "Keterangan", "Tipe", "Nominal", "Jam"]);
      sheetKas.appendRow([
        tgl,
        keterangan,
        "Kas Keluar",
        totalNominal,
        jamSekarang
      ]);
    }

    // 3. CATAT KE AUDIT LOG 
    catatAuditLog(
      "INPUT GAJI", 
      "Gaji Karyawan", 
      "-", 
      payload.nama + " (" + payload.hari + " - Rp " + totalNominal.toLocaleString('id-ID') + ")", 
      userAktif
    );

    SpreadsheetApp.flush();
    return ambilDataTerbaruJson();
  } catch (err) {
    throw new Error("Gagal simpan gaji: " + err.message);
  } finally {
    lock.releaseLock();
  }
}

function simpanPenyesuaianGaji(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = pastikanSheetAda(ss, "Penyesuaian Gaji", ["Tanggal", "Nama", "Tipe", "Nominal", "Keterangan", "Jam"]);
  var waktuSekarang = Utilities.formatDate(new Date(), "GMT+7", "HH:mm");
  var jamTeks = "'" + waktuSekarang;
  var nominal = Number(data.nominal) || 0;
  sheet.appendRow([data.tanggal, data.nama, data.tipe, nominal, data.keterangan || "-", jamTeks]);
  SpreadsheetApp.flush();
  return ambilDataTerbaruJson();
}

function postingGajiKePembukuan(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nominal = Number(data.nominal) || 0;
  var waktuSekarang = Utilities.formatDate(new Date(), "GMT+7", "HH:mm");
  var jamTeks = "'" + waktuSekarang;
  var tgl = data.tanggal || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var ket = data.keterangan || ("Beban Gaji Karyawan (" + data.periode + ")");

  var sheetOps = ss.getSheetByName("Pengeluaran Lain");
  if (!sheetOps) sheetOps = pastikanSheetAda(ss, "Pengeluaran Lain", ["Nama Pengeluaran", "Tanggal", "Nominal", "Toko", "Jam"]);
  sheetOps.appendRow([ket, tgl, nominal, "-", jamTeks]);

  var sheetAK = pastikanSheetAda(ss, "Arus Kas", ["Tanggal", "Keterangan", "Tipe", "Nominal", "Jam"]);
  sheetAK.appendRow([tgl, ket, "Kas Keluar", nominal, jamTeks]);

  SpreadsheetApp.flush();
  return ambilDataTerbaruJson();
}

function cadangkanSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd_HH-mm");
  var backupName = "Backup_" + ss.getName() + "_" + timestamp;
  var copy = file.makeCopy(backupName);
  return { url: copy.getUrl(), name: backupName };
}

function eksekusiAutoBackup() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var fileAsli = DriveApp.getFileById(ss.getId());
    var namaFolderBackup = "Backup Database Pembukuan";
    var folderIter = DriveApp.getFoldersByName(namaFolderBackup);
    var targetFolder = folderIter.hasNext() ? folderIter.next() : DriveApp.createFolder(namaFolderBackup);
    
    var timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd_HH-mm");
    var backupName = "AutoBackup_" + ss.getName() + "_" + timestamp;
    var fileSalinan = fileAsli.makeCopy(backupName, targetFolder);
    
    var batasHari = 30;
    var waktuSekarang = new Date().getTime();
    var filesInFolder = targetFolder.getFiles();
    while (filesInFolder.hasNext()) {
      var f = filesInFolder.next();
      var umurHari = (waktuSekarang - f.getDateCreated().getTime()) / (1000 * 60 * 60 * 24);
      if (umurHari > batasHari) f.setTrashed(true);
    }
    return { status: "success", name: backupName, url: fileSalinan.getUrl() };
  } catch (err) {
    return { status: "error", message: err.message };
  }
}

function kirimRingkasanMalam() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tglHariIni = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
    var tglFormatTeks = Utilities.formatDate(new Date(), "GMT+7", "dd MMMM yyyy");
    
    var semuaData = ambilSemuaData();
    var bepCfg = semuaData.bepConfig || getBEPDefault();
    
    var omzetHariIni = 0;
    var platformMap = {};
    var dataPemasukan = semuaData.pemasukan || [];
    for (var i = 1; i < dataPemasukan.length; i++) {
      var tglRow = String(dataPemasukan[i][1]).split('T')[0];
      if (tglRow === tglHariIni) {
        var nom = Number(dataPemasukan[i][2]) || 0;
        var plat = String(dataPemasukan[i][3] || "Lainnya").trim();
        omzetHariIni += nom;
        platformMap[plat] = (platformMap[plat] || 0) + nom;
      }
    }
    
    function hitungBeban(rows) {
      var tot = 0;
      for (var j = 1; j < rows.length; j++) {
        if (String(rows[j][1]).split('T')[0] === tglHariIni) tot += Number(rows[j][2]) || 0;
      }
      return tot;
    }
    
    var totalBiayaShopee = hitungBeban(semuaData.shopee || []);
    var totalBiayaAgen = hitungBeban(semuaData.agen || []);
    var totalBiayaLain = hitungBeban(semuaData.pengeluaranLain || []);
    var totalBiayaIklan = hitungBeban(semuaData.iklan || []);
    var totalPengeluaran = totalBiayaShopee + totalBiayaAgen + totalBiayaLain + totalBiayaIklan;
    var labaBersihHariIni = omzetHariIni - totalPengeluaran;
    
    var targetBEPHarian = 0;
    if (bepCfg.biayaTetapTotal && bepCfg.margin && bepCfg.jumlahHari) {
      var bepBulanan = Math.round(bepCfg.biayaTetapTotal / (bepCfg.margin / 100));
      targetBEPHarian = Math.round(bepBulanan / bepCfg.jumlahHari);
    }
    
    var bepStatusText = (omzetHariIni >= targetBEPHarian && targetBEPHarian > 0)
      ? "🎉 LOLOS BEP HARIAN! (+Rp " + (omzetHariIni - targetBEPHarian).toLocaleString('id-ID') + ")"
      : "⚠️ Kurang Rp " + (targetBEPHarian - omzetHariIni).toLocaleString('id-ID') + " untuk BEP Harian";
    var bepColor = (omzetHariIni >= targetBEPHarian && targetBEPHarian > 0) ? "#059669" : "#e11d48";

    var htmlPlatform = "";
    var platformKeys = Object.keys(platformMap);
    if (platformKeys.length === 0) {
      htmlPlatform = "<tr><td colspan='2' style='padding:8px; text-align:center; color:#94a3b8; font-size:12px;'>Belum ada transaksi hari ini</td></tr>";
    } else {
      platformKeys.forEach(function(p) {
        htmlPlatform += "<tr><td style='padding:6px 8px; font-size:13px; color:#334155; border-bottom:1px solid #f1f5f9;'>" + p + "</td><td style='padding:6px 8px; font-size:13px; font-weight:bold; text-align:right; color:#4f46e5; border-bottom:1px solid #f1f5f9;'>Rp " + platformMap[p].toLocaleString('id-ID') + "</td></tr>";
      });
    }

    var htmlBody = `
    <div style="font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width:540px; margin:auto; background:#f8fafc; padding:20px; border-radius:16px;">
      <div style="background:#4f46e5; padding:24px; border-radius:12px; text-align:center; color:#ffffff;">
        <span style="font-size:32px;">📊</span>
        <h2 style="margin:8px 0 4px; font-size:20px; font-weight:800;">Laporan Harian Bisnis</h2>
        <p style="margin:0; font-size:13px; opacity:0.85;">${tglFormatTeks}</p>
      </div>
      <div style="background:#ffffff; padding:20px; margin-top:16px; border-radius:12px; border:1px solid #e2e8f0;">
        <table style="width:100%; border-collapse:collapse;">
          <tr><td style="padding:10px 0; color:#64748b; font-size:13px; font-weight:600;">Total Omzet:</td><td style="padding:10px 0; font-size:16px; font-weight:800; color:#0f172a; text-align:right;">Rp ${omzetHariIni.toLocaleString('id-ID')}</td></tr>
          <tr><td style="padding:10px 0; color:#64748b; font-size:13px; font-weight:600;">Total Pengeluaran:</td><td style="padding:10px 0; font-size:16px; font-weight:800; color:#e11d48; text-align:right;">Rp ${totalPengeluaran.toLocaleString('id-ID')}</td></tr>
          <tr style="border-top:2px dashed #cbd5e1;"><td style="padding:12px 0; color:#0f172a; font-size:14px; font-weight:bold;">Laba Bersih Harian:</td><td style="padding:12px 0; font-size:18px; font-weight:900; color:${labaBersihHariIni >= 0 ? '#059669' : '#e11d48'}; text-align:right;">Rp ${labaBersihHariIni.toLocaleString('id-ID')}</td></tr>
        </table>
      </div>
      <div style="background:#ffffff; padding:16px; margin-top:12px; border-radius:12px; border-left:4px solid ${bepColor}; border:1px solid #e2e8f0;">
        <div style="font-size:11px; font-weight:bold; color:#64748b; text-transform:uppercase;">Evaluasi Target BEP Harian</div>
        <div style="font-size:14px; font-weight:800; margin-top:4px; color:${bepColor};">${bepStatusText}</div>
        <div style="font-size:11px; color:#94a3b8; margin-top:2px;">Target Harian: Rp ${targetBEPHarian.toLocaleString('id-ID')}</div>
      </div>
      <div style="background:#ffffff; padding:16px; margin-top:12px; border-radius:12px; border:1px solid #e2e8f0;">
        <div style="font-size:12px; font-weight:bold; color:#334155; margin-bottom:8px;">Rincian Omzet per Platform:</div>
        <table style="width:100%; border-collapse:collapse;">${htmlPlatform}</table>
      </div>
      <div style="background:#ffffff; padding:16px; margin-top:12px; border-radius:12px; border:1px solid #e2e8f0;">
        <div style="font-size:12px; font-weight:bold; color:#334155; margin-bottom:8px;">Rincian Pengeluaran:</div>
        <table style="width:100%; border-collapse:collapse; font-size:12px;">
          <tr><td style="padding:4px 0; color:#64748b;">Belanja Shopee:</td><td style="text-align:right; font-weight:bold; color:#334155;">Rp ${totalBiayaShopee.toLocaleString('id-ID')}</td></tr>
          <tr><td style="padding:4px 0; color:#64748b;">Belanja Agen:</td><td style="text-align:right; font-weight:bold; color:#334155;">Rp ${totalBiayaAgen.toLocaleString('id-ID')}</td></tr>
          <tr><td style="padding:4px 0; color:#64748b;">Biaya Iklan (Ads):</td><td style="text-align:right; font-weight:bold; color:#9333ea;">Rp ${totalBiayaIklan.toLocaleString('id-ID')}</td></tr>
          <tr><td style="padding:4px 0; color:#64748b;">Pengeluaran Lain:</td><td style="text-align:right; font-weight:bold; color:#334155;">Rp ${totalBiayaLain.toLocaleString('id-ID')}</td></tr>
        </table>
      </div>
    </div>`;

    var emailTujuan = Session.getEffectiveUser().getEmail();
    MailApp.sendEmail({
      to: emailTujuan,
      subject: "📊 Ringkasan Keuangan Harian [" + tglHariIni + "] - " + (labaBersihHariIni >= 0 ? "Untung Rp " + labaBersihHariIni.toLocaleString('id-ID') : "Rugi Rp " + Math.abs(labaBersihHariIni).toLocaleString('id-ID')),
      htmlBody: htmlBody
    });
    return { status: "success", email: emailTujuan };
  } catch (err) {
    return { status: "error", message: err.message };
  }
}

function pasangTriggerOtomatis() {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    var handler = allTriggers[i].getHandlerFunction();
    if (handler === "kirimRingkasanMalam" || handler === "eksekusiAutoBackup") {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }
  ScriptApp.newTrigger("kirimRingkasanMalam").timeBased().everyDays(1).atHour(21).create();
  ScriptApp.newTrigger("eksekusiAutoBackup").timeBased().everyDays(1).atHour(23).create();
  return "Trigger aktif: Ringkasan Malam (21.00) & Auto-Backup (23.00).";
}

function ambilDataAuditLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Audit_Log");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  var limit = Math.min(lastRow - 1, 50);
  var startRow = lastRow - limit + 1; 
  var values = sheet.getRange(startRow, 1, limit, 6).getValues();
  var logs = values.map(function(row, index) {
    var waktuStr = (row[0] instanceof Date) ? Utilities.formatDate(row[0], "GMT+7", "yyyy-MM-dd HH:mm") : String(row[0]);
    return {
      waktu: waktuStr,
      user: String(row[1] || "-"),
      aksi: String(row[2] || "-"),
      kategori: String(row[3] || "-"),
      dataLama: String(row[4] || "-"),
      dataBaru: String(row[5] || "-"),
      rowIdx: startRow + index // <-- Mengambil nomor baris asli di Google Sheets
    };
  });
  return logs.reverse(); 
}

function siapkanSheetStok() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = pastikanSheetAda(ss, "Master_Stok", ["Kode / SKU", "Nama Barang", "Kategori", "Stok (Pcs)", "Satuan", "Harga Modal (Rp)", "Batas Min Stok", "Pack_Per_Dus", "Pcs_Per_Pack"]);
  if (sheet.getLastColumn() < 9) {
    sheet.getRange(1, 8).setValue("Pack_Per_Dus").setFontWeight("bold").setBackground("#d9ead3");
    sheet.getRange(1, 9).setValue("Pcs_Per_Pack").setFontWeight("bold").setBackground("#d9ead3");
  }
  pastikanSheetAda(ss, "Master_Bundling", ["ID Bundling", "Nama Paket", "Komposisi Resep (JSON)", "Harga Jual (Rp)", "Keterangan"]);
  pastikanSheetAda(ss, "Mutasi_Stok", ["Tanggal", "Jam", "Tipe", "Nama Barang", "Jumlah", "Keterangan"]);
}

// ================================================================
// INDEX ALIAS SKU: MENJAGA RIWAYAT NAMA PRODUK TETAP TERHUBUNG KE SKU
// ================================================================
function normalisasiNamaSkuServer_(nilai) {
  return String(nilai || "").toLowerCase().normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

function simpanAliasNamaSku_(sku, namaAlias, namaSaatIniBaru, sumber) {
  sku = String(sku || "").trim();
  namaAlias = String(namaAlias || "").trim();
  if (!sku || !namaAlias) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Alias_SKU");
  if (!sh) {
    sh = ss.insertSheet("Alias_SKU");
    sh.getRange(1, 1, 1, 5).setValues([[
      "Kode / SKU", "Nama Alias", "Nama Saat Ini", "Sumber", "Aktif"
    ]]).setFontWeight("bold").setBackground("#d9ead3");
  }

  var master = ss.getSheetByName("Master_Stok");
  var namaSekarang = "";
  if (master && master.getLastRow() > 1) {
    var mRows = master.getRange(2, 1, master.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < mRows.length; i++) {
      if (String(mRows[i][0] || "").trim() === sku) {
        namaSekarang = String(mRows[i][1] || "").trim();
        break;
      }
    }
  }

  var key = sku + "|" + normalisasiNamaSkuServer_(namaAlias);
  var last = sh.getLastRow();
  if (last > 1) {
    var vals = sh.getRange(2, 1, last - 1, 5).getValues();
    for (var r = 0; r < vals.length; r++) {
      var existingKey = String(vals[r][0] || "").trim() + "|" +
        normalisasiNamaSkuServer_(vals[r][1]);
      if (existingKey === key) return;
    }
  }
  sh.appendRow([sku, namaAlias, namaSaatIniBaru || namaSekarang, sumber || "Perubahan Nama Master Stok", "YA"]);
}

function buatIndexAliasSku_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Master_Stok");
  if (!masterSheet || masterSheet.getLastRow() <= 1) return {};

  var masterRows = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1,
    Math.max(6, masterSheet.getLastColumn())).getValues();
  var masterBySku = {};
  var petaNamaKeSku = {};
  var calonBaris = [];
  var sudahAda = {};

  function tambahAlias(sku, alias, sumber) {
    sku = String(sku || "").trim();
    alias = String(alias || "").trim();
    if (!sku || !alias || !masterBySku[sku]) return;
    var keyNama = normalisasiNamaSkuServer_(alias);
    if (!keyNama) return;
    var keyBaris = sku + "|" + keyNama;
    if (!sudahAda[keyBaris]) {
      sudahAda[keyBaris] = true;
      calonBaris.push([sku, alias, masterBySku[sku].nama, sumber || "Alias Historis", "YA"]);
    }
    // Nama Master yang aktif dimasukkan lebih dahulu; jika ada tabrakan,
    // pemetaan nama aktif tidak ditimpa alias lama.
    if (!petaNamaKeSku[keyNama]) petaNamaKeSku[keyNama] = sku;
  }

  // 1. Daftar SKU dan nama barang yang berlaku saat ini.
  masterRows.forEach(function(row) {
    var sku = String(row[0] || "").trim();
    var nama = String(row[1] || "").trim();
    if (sku) masterBySku[sku] = { nama: nama, hpp: Number(row[5]) || 0 };
  });
  masterRows.forEach(function(row) {
    tambahAlias(row[0], row[1], "Nama aktif Master_Stok");
  });

  // 2. Muat alias yang pernah disimpan manual/oleh sistem.
  var aliasSheet = ss.getSheetByName("Alias_SKU");
  if (!aliasSheet) {
    aliasSheet = ss.insertSheet("Alias_SKU");
    aliasSheet.getRange(1, 1, 1, 5).setValues([[
      "Kode / SKU", "Nama Alias", "Nama Saat Ini", "Sumber", "Aktif"
    ]]).setFontWeight("bold").setBackground("#d9ead3");
  } else if (aliasSheet.getLastRow() === 0) {
    aliasSheet.getRange(1, 1, 1, 5).setValues([[
      "Kode / SKU", "Nama Alias", "Nama Saat Ini", "Sumber", "Aktif"
    ]]).setFontWeight("bold").setBackground("#d9ead3");
  }
  if (aliasSheet.getLastRow() > 1) {
    var existingAliases = aliasSheet.getRange(2, 1, aliasSheet.getLastRow() - 1, 5).getValues();
    existingAliases.forEach(function(row) {
      var aktif = String(row[4] == null ? "YA" : row[4]).toUpperCase();
      if (aktif === "TIDAK" || aktif === "FALSE" || aktif === "0") return;
      tambahAlias(row[0], row[1], row[3] || "Alias_SKU (manual)");
    });
  }

  // 3. Rekonstruksi nama sebelum perubahan dari Audit_Log.
  var auditSheet = ss.getSheetByName("Audit_Log");
  if (auditSheet && auditSheet.getLastRow() > 1) {
    var auditRows = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1,
      Math.max(6, auditSheet.getLastColumn())).getValues();
    auditRows.forEach(function(row) {
      if (String(row[2] || "").trim().toUpperCase() !== "EDIT BARANG") return;
      var sebelum = String(row[4] || "").trim();
      var pemisah = sebelum.indexOf(" - ");
      if (pemisah < 1) return;
      var sku = sebelum.substring(0, pemisah).trim();
      var namaLama = sebelum.substring(pemisah + 3).trim();
      tambahAlias(sku, namaLama, "Audit_Log: nama sebelum edit");
    });
  }

  // 4. Resep Master_Bundling menyimpan SKU + nama bahan saat resep dibuat.
  var bundleSheet = ss.getSheetByName("Master_Bundling");
  if (bundleSheet && bundleSheet.getLastRow() > 1) {
    var bundleRows = bundleSheet.getRange(2, 1, bundleSheet.getLastRow() - 1,
      Math.max(3, bundleSheet.getLastColumn())).getValues();
    bundleRows.forEach(function(row) {
      var jsonResep = String(row[2] || "");
      if (!jsonResep) return;
      try {
        var resep = JSON.parse(jsonResep);
        if (!Array.isArray(resep)) return;
        resep.forEach(function(item) {
          tambahAlias(item.kode || item.sku, item.nama, "Master_Bundling: nama bahan di resep");
        });
      } catch (e) {
        Logger.log("Alias SKU: resep bundling gagal dibaca: " + e.message);
      }
    });
  }

  // 5. Simpan alias historis baru ke sheet secara batch (tanpa mengubah mutasi lama).
  var existingKeys = {};
  if (aliasSheet.getLastRow() > 1) {
    var currentRows = aliasSheet.getRange(2, 1, aliasSheet.getLastRow() - 1, 5).getValues();
    currentRows.forEach(function(row) {
      existingKeys[String(row[0] || "").trim() + "|" + normalisasiNamaSkuServer_(row[1])] = true;
    });
  }
  var addRows = calonBaris.filter(function(row) {
    var k = String(row[0]) + "|" + normalisasiNamaSkuServer_(row[1]);
    if (existingKeys[k]) return false;
    existingKeys[k] = true;
    return true;
  });
  if (addRows.length) {
    aliasSheet.getRange(aliasSheet.getLastRow() + 1, 1, addRows.length, 5).setValues(addRows);
  }

  // Segarkan kolom Nama Saat Ini agar setiap alias selalu menunjuk nama master terbaru.
  if (aliasSheet.getLastRow() > 1) {
    var semuaAliasTersimpan = aliasSheet.getRange(2, 1, aliasSheet.getLastRow() - 1, 5).getValues();
    var aliasBerubah = false;
    semuaAliasTersimpan.forEach(function(row) {
      var skuRow = String(row[0] || "").trim();
      if (masterBySku[skuRow] && String(row[2] || "") !== masterBySku[skuRow].nama) {
        row[2] = masterBySku[skuRow].nama;
        aliasBerubah = true;
      }
    });
    if (aliasBerubah) {
      aliasSheet.getRange(2, 1, semuaAliasTersimpan.length, 5).setValues(semuaAliasTersimpan);
    }
  }

  // Kembalikan peta nama ternormalisasi -> SKU untuk dipakai laporan.
  return petaNamaKeSku;
}

function ambilDataStokLengkap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  function getRows(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      siapkanSheetStok();
      sheet = ss.getSheetByName(name);
    }
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var vals = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    return vals.map(function(row) {
      return row.map(function(cell) {
        if (cell instanceof Date) {
          if (cell.getFullYear() <= 1900) return Utilities.formatDate(cell, "GMT+7", "HH:mm");
          return Utilities.formatDate(cell, "GMT+7", "yyyy-MM-dd");
        }
        return cell;
      });
    });
  }
  return JSON.stringify({
    stok: getRows("Master_Stok"),
    bundling: getRows("Master_Bundling"),
    mutasi: getRows("Mutasi_Stok"),
    skuAliasMap: buatIndexAliasSku_()
  });
}

function simpanItemStokBaru(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  siapkanSheetStok();
  var sheet = ss.getSheetByName("Master_Stok");
  var userAktif = data.user || "Staf";

  var inputNama = String(data.nama || "").trim();
  var inputKategori = data.kategori || "Bahan Baku";
  var inputStok = Number(data.stok) || 0;
  var inputHargaModal = Number(data.hargaModal) || 0;
  var inputMinStok = Number(data.minStok) || 5;
  var inputSatuan = data.satuan || "pcs";
  var packPerDus = Number(data.packPerDus) || 1;
  var pcsPerPack = Number(data.pcsPerPack) || 1;

  var vals = sheet.getDataRange().getValues();
  var rowIndexKetemu = -1;

  for (var i = 1; i < vals.length; i++) {
    var namaDiSheet = String(vals[i][1] || "").trim();
    if (namaDiSheet.toLowerCase() === inputNama.toLowerCase()) {
      rowIndexKetemu = i + 1;
      var stokLama = Number(vals[i][3]) || 0;
      var hargaLama = Number(vals[i][5]) || 0;
      var satuanLama = vals[i][4] || inputSatuan;

      var totalStokBaru = stokLama + inputStok;
      var totalNilaiLama = (stokLama > 0 ? stokLama : 0) * hargaLama;
      var totalNilaiBaru = inputStok * inputHargaModal;
      var hargaRataRata = totalStokBaru > 0 ? Math.round((totalNilaiLama + totalNilaiBaru) / totalStokBaru) : inputHargaModal;

      // OPTIMASI: Menulis 6 kolom (kolom 4 sampai 9) sekaligus dalam 1 panggilan API batch
      sheet.getRange(rowIndexKetemu, 4, 1, 6).setValues([[
        totalStokBaru,
        satuanLama,
        hargaRataRata,
        inputMinStok,
        packPerDus,
        pcsPerPack
      ]]);

      catatMutasiStok("Restock / Rata-rata", namaDiSheet, inputStok, "Masuk " + inputStok + " pcs [Oleh: " + userAktif + "]");
      catatAuditLog("RESTOCK", "Master Stok", "Stok: " + stokLama, "Stok: " + totalStokBaru + " (" + inputNama + ")", userAktif);
      break;
    }
  }

  if (rowIndexKetemu === -1) {
    var kode = data.kode || ("BRG-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss"));
    sheet.appendRow([kode, inputNama, inputKategori, inputStok, inputSatuan, inputHargaModal, inputMinStok, packPerDus, pcsPerPack]);
    catatMutasiStok("Stok Awal", inputNama, inputStok, "Barang Baru [Oleh: " + userAktif + "]");
    catatAuditLog("TAMBAH BARANG", "Master Stok", "-", inputNama + " (Stok Awal: " + inputStok + ")", userAktif);
  }

  if (data.catatKePembukuan && Number(data.totalBiayaBelanja) > 0) {
    var totalBiaya = Number(data.totalBiayaBelanja);
    var katBuku = data.kategoriPengeluaran || "Pembelian Agen";
    var tokoBuku = (data.toko && String(data.toko).trim() !== "") ? String(data.toko).trim() : "-";
    var tglHariIni = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
    var jamTeks = "'" + Utilities.formatDate(new Date(), "GMT+7", "HH:mm");
    var ketPengeluaran = "Restock: " + inputNama + " (" + (data.ringkasanQty || (inputStok + " pcs")) + ")";

    var sheetBeban = ss.getSheetByName(katBuku) || pastikanSheetAda(ss, katBuku, ["Nama", "Tanggal", "Nominal", "Toko", "Jam"]);
    sheetBeban.appendRow([ketPengeluaran, tglHariIni, totalBiaya, tokoBuku, jamTeks]);

    var sheetAK = ss.getSheetByName("Arus Kas") || pastikanSheetAda(ss, "Arus Kas", ["Tanggal", "Keterangan", "Tipe", "Nominal", "Jam"]);
    sheetAK.appendRow([tglHariIni, "[" + katBuku + "] Restock " + inputNama, "Kas Keluar", totalBiaya, jamTeks]);
  }

  SpreadsheetApp.flush();
  return ambilDataStokLengkap();
}

function sesuaikanJumlahStok(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  siapkanSheetStok();
  var sheet = ss.getSheetByName("Master_Stok");
  if (!sheet) return ambilDataStokLengkap();
  
  var vals = sheet.getDataRange().getValues();
  var jumlah = Number(data.jumlah) || 0;

  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(data.kode).trim()) {
      var stokLama = Number(vals[i][3]) || 0;
      var stokBaru = (data.tipe === "Masuk") ? (stokLama + jumlah) : Math.max(0, stokLama - jumlah);
      sheet.getRange(i + 1, 4).setValue(stokBaru);
      catatMutasiStok(data.tipe === "Masuk" ? "Barang Masuk" : "Barang Keluar", vals[i][1], jumlah, data.keterangan || "-");
      break;
    }
  }
  SpreadsheetApp.flush();
  return ambilDataStokLengkap();
}

function eksekusiPenguranganStok(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  siapkanSheetStok();
  var sheet = ss.getSheetByName("Master_Stok");
  if (!sheet) return ambilDataStokLengkap();

  var userAktif = data.user || "Staf";
  var vals = sheet.getDataRange().getValues();
  var jumlahPcs = Number(data.jumlah) || 0;
  var namaBarang = "";
  var stokLama = 0; // Deklarasi di luar blok for agar bisa dipakai di Audit Log

  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(data.kode).trim()) {
      namaBarang = vals[i][1];
      stokLama = Number(vals[i][3]) || 0;
      
      // Validasi ketat di sisi server untuk mencegah stok minus
      if (stokLama < jumlahPcs) {
        throw new Error("Transaksi ditolak! Sisa stok " + namaBarang + " tidak mencukupi (Sisa: " + stokLama + " pcs).");
      }
      
      var stokBaru = Math.max(0, stokLama - jumlahPcs);
      sheet.getRange(i + 1, 4).setValue(stokBaru);
      break;
    }
  }

  // Hentikan proses jika barang fiktif (tidak ditemukan di database)
  if (!namaBarang) {
    throw new Error("Barang tidak ditemukan di database stok gudang.");
  }

  // PERBAIKAN BACA TANGGAL DAN PEMBELI
  var tglPilihan = (data.tanggal && String(data.tanggal).trim() !== "")
    ? String(data.tanggal).trim() 
    : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var jamTeks = "'" + Utilities.formatDate(new Date(), "GMT+7", "HH:mm");
  
  var pembeli = (data.pembeli && String(data.pembeli).trim() !== "") ? String(data.pembeli).trim() : "";
  var platform = data.platform || "Toko Offline";

  var tipeMutasi = (data.alasan === "Penjualan") ? "Penjualan Satuan" : "Barang Keluar";
  var ketMutasi = (data.keterangan || "-") + " [Oleh: " + userAktif + "]";
  
  if (data.alasan === "Penjualan") {
    ketMutasi = "Terjual " + jumlahPcs + " pcs (" + platform + ")";
    if (pembeli) ketMutasi += " - Pembeli: " + pembeli;
    ketMutasi += " [Net: " + data.totalOmzet + "] [Oleh: " + userAktif + "]";
  }
  
  catatMutasiStok(tipeMutasi, namaBarang, jumlahPcs, ketMutasi, tglPilihan);
  catatAuditLog("KELUAR STOK", data.alasan, "Sisa Lama: " + stokLama, namaBarang + " (-" + jumlahPcs + " pcs)", userAktif);

  // Otomatis catat ke Laporan Pemasukan & Arus Kas jika penjualan
  if (data.alasan === "Penjualan" && data.catatKePembukuan) {
    var totalOmzet = Number(data.totalOmzet);
    // Format nama buku disamakan dengan bundling: "Nama Pembeli - Nama Barang (Qty)"
    var namaBuku = pembeli 
       ? (pembeli + " - " + namaBarang + " (" + data.ringkasanQty + ")") 
       : ("Penjualan " + namaBarang + " (" + data.ringkasanQty + ")");

    var sheetPemasukan = ss.getSheetByName("Pemasukan");
    if (!sheetPemasukan) sheetPemasukan = pastikanSheetAda(ss, "Pemasukan", ["Nama", "Tanggal", "Nominal", "Toko", "Jam"]);
    sheetPemasukan.appendRow([namaBuku, tglPilihan, totalOmzet, platform, jamTeks]);

    var sheetAK = pastikanSheetAda(ss, "Arus Kas", ["Tanggal", "Keterangan", "Tipe", "Nominal", "Jam"]);
    sheetAK.appendRow([tglPilihan, "[Pemasukan] " + namaBuku, "Kas Masuk", totalOmzet, jamTeks]);
  }

  SpreadsheetApp.flush();
  return JSON.stringify({
    success: true,
    stokData: JSON.parse(ambilDataStokLengkap()),
    bukuData: ambilSemuaData(false)
  });
}

/**
 * Menyimpan atau memperbarui resep paket bundling lengkap dengan konfigurasi biaya marketplace
 * @param {Object} data - Payload data bundling, resep bahan, dan daftar biaya potongan marketplace
 * @return {string} - Data JSON stok terbaru
 */
function simpanResepBundling(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Master_Bundling");
  if (!sheet) {
    sheet = pastikanSheetAda(ss, "Master_Bundling", [
      "ID Bundling", 
      "Nama Paket", 
      "Komposisi Resep (JSON)", 
      "Harga Jual (Rp)", 
      "Keterangan", 
      "Biaya Marketplace (JSON)"
    ]);
  }

  // Pastikan kolom ke-6 tersedia untuk menyimpan biaya marketplace tanpa merusak kolom lama
  if (sheet.getLastColumn() < 6) {
    sheet.getRange(1, 6).setValue("Biaya Marketplace (JSON)").setFontWeight("bold").setBackground("#d9ead3");
  }

  var vals = sheet.getDataRange().getValues();
  var idTarget = String(data.idBundling || "").trim();
  var rowIndexKetemu = -1;

  if (idTarget) {
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === idTarget) {
        rowIndexKetemu = i + 1;
        break;
      }
    }
  }

  var resepJson = JSON.stringify(data.resep || []);
  var biayaJson = JSON.stringify(data.biayaMarketplace || []);
  var hargaJual = Number(data.hargaJual) || 0;
  var keterangan = data.keterangan || "-";

  if (rowIndexKetemu !== -1) {
    // Mode Update / Edit
    sheet.getRange(rowIndexKetemu, 2).setValue(data.namaPaket);
    sheet.getRange(rowIndexKetemu, 3).setValue(resepJson);
    sheet.getRange(rowIndexKetemu, 4).setValue(hargaJual);
    sheet.getRange(rowIndexKetemu, 5).setValue(keterangan);
    sheet.getRange(rowIndexKetemu, 6).setValue(biayaJson);
  } else {
    // Mode Tambah Baru
    var idBundling = "BDL-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss");
    sheet.appendRow([
      idBundling,
      data.namaPaket,
      resepJson,
      hargaJual,
      keterangan,
      biayaJson
    ]);
  }

  SpreadsheetApp.flush();
  return ambilDataStokLengkap();
}

function hapusResepBundling(idBundling) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Master_Bundling");
  if (!sheet) return ambilDataStokLengkap();

  var vals = sheet.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(idBundling)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }

  SpreadsheetApp.flush();
  return ambilDataStokLengkap();
}

function catatMutasiStok(tipe, namaBarang, qty, ket, customTgl) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Mutasi_Stok") || ss.getSheetByName("Mutasi Stok");
  if (!sheet) {
    sheet = pastikanSheetAda(ss, "Mutasi_Stok", ["Tanggal", "Jam", "Tipe", "Nama Barang", "Jumlah", "Keterangan"]);
  }
  var tgl = customTgl || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var jam = Utilities.formatDate(new Date(), "GMT+7", "HH:mm");
  sheet.appendRow([tgl, "'" + jam, tipe, namaBarang, Number(qty) || 0, ket || "-"]);
}

function eksekusiPenjualanBundling(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetStok = ss.getSheetByName("Master_Stok");
  if (!sheetStok) return ambilDataStokLengkap();

  var lastRow = sheetStok.getLastRow();
  if (lastRow <= 1) return ambilDataStokLengkap();

  var rangeStok = sheetStok.getRange(2, 1, lastRow - 1, sheetStok.getLastColumn());
  var dataStok = rangeStok.getValues();

  var qtyJual = Number(data.qty) || 1;
  var resep = data.resep || [];
  var namaPaket = data.namaPaket || "Paket Bundling";
  var hargaSatuan = Number(data.hargaSatuan) || 0;
  var totalOmzet = hargaSatuan * qtyJual;

  var tglPilihan = (data.tanggal && String(data.tanggal).trim() !== "") 
    ? String(data.tanggal).trim() 
    : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");

  var platform = data.platform || "Toko Offline";
  var tokoField = (data.toko && String(data.toko).trim() !== "") 
    ? String(data.toko).trim() 
    : platform;

  var pembeli = (data.namaPembeli && String(data.namaPembeli).trim() !== "") 
    ? String(data.namaPembeli).trim() 
    : "";

  var jamSekarang = Utilities.formatDate(new Date(), "GMT+7", "HH:mm");
  var mutasiBatch = [];

  // OPTIMASI: Hitung semua pemotongan stok di memori (in-memory update) DENGAN VALIDASI
  for (var r = 0; r < resep.length; r++) {
    var butuhTotal = (Number(resep[r].butuh) || 1) * qtyJual;
    for (var s = 0; s < dataStok.length; s++) {
      if (String(dataStok[s][0]) === String(resep[r].kode)) {
        var stokSekarang = Number(dataStok[s][3]) || 0;
        
        // VALIDASI SERVER: Cegah stok minus
        if (stokSekarang < butuhTotal) {
          throw new Error("Transaksi ditolak! Sisa bahan " + dataStok[s][1] + " tidak cukup (Sisa: " + stokSekarang + " pcs, Dibutuhkan: " + butuhTotal + " pcs).");
        }
        
        var sisa = stokSekarang - butuhTotal;
        dataStok[s][3] = sisa; // Potong di array memori

        var ketBahan = "Komponen " + qtyJual + "x " + namaPaket + " (" + tokoField + ")";
        if (pembeli) ketBahan += " - Pembeli: " + pembeli;

        mutasiBatch.push([tglPilihan, "'" + jamSekarang, "Bahan Bundling Terpakai", dataStok[s][1], butuhTotal, ketBahan]);
        break;
      }
    }
  }

  // Tulis kembali stok yang telah dipotong ke Spreadsheet sekaligus dalam 1 panggilan
  rangeStok.setValues(dataStok);

  // Catat penjualan paket bundling
  var ketPaket = "Penjualan via " + tokoField;
  if (pembeli) ketPaket += " - Pembeli: " + pembeli;
  mutasiBatch.push([tglPilihan, "'" + jamSekarang, "Penjualan Bundling", namaPaket, qtyJual, ketPaket]);

  // OPTIMASI: Simpan seluruh mutasi bahan resep sekaligus dalam 1 batch
  var sheetMutasi = ss.getSheetByName("Mutasi_Stok") || ss.getSheetByName("Mutasi Stok");
  if (sheetMutasi && mutasiBatch.length > 0) {
    sheetMutasi.getRange(sheetMutasi.getLastRow() + 1, 1, mutasiBatch.length, 6).setValues(mutasiBatch);
  }

  if (data.catatKePembukuan && totalOmzet > 0) {
    var jamTeks = "'" + jamSekarang;
    var namaBuku = pembeli 
      ? (pembeli + " - " + namaPaket + " (" + qtyJual + "x)") 
      : ("Penjualan " + namaPaket + " (" + qtyJual + " paket)");

    var sheetPemasukan = ss.getSheetByName("Pemasukan") || pastikanSheetAda(ss, "Pemasukan", ["Nama", "Tanggal", "Nominal", "Toko", "Jam"]);
    sheetPemasukan.appendRow([namaBuku, tglPilihan, totalOmzet, tokoField, jamTeks]);

    var sheetAK = ss.getSheetByName("Arus Kas") || pastikanSheetAda(ss, "Arus Kas", ["Tanggal", "Keterangan", "Tipe", "Nominal", "Jam"]);
    sheetAK.appendRow([tglPilihan, "[Pemasukan] " + namaBuku, "Kas Masuk", totalOmzet, jamTeks]);
  }

  SpreadsheetApp.flush();
  return ambilDataStokLengkap();
}

function simpanStokOpname(penyesuaianList, user) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Master_Stok");
  if (!sheet) return ambilDataStokLengkap();

  var userAktif = user || "Staf";
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return ambilDataStokLengkap();

  // Buat lookup Map agar pencarian kode barang instan O(1)
  var opnameMap = {};
  for (var p = 0; p < penyesuaianList.length; p++) {
    if (penyesuaianList[p].selisih !== 0) {
      opnameMap[String(penyesuaianList[p].kode).trim()] = penyesuaianList[p];
    }
  }

  var rangeStok = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  var vals = rangeStok.getValues();
  var mutasiBatch = [];
  var auditBatch = [];
  var tglHariIni = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var jamTeks = Utilities.formatDate(new Date(), "GMT+7", "HH:mm");
  var adaPerubahan = false;

  for (var i = 0; i < vals.length; i++) {
    var kodeBarang = String(vals[i][0]).trim();
    if (opnameMap.hasOwnProperty(kodeBarang)) {
      var itemP = opnameMap[kodeBarang];
      vals[i][3] = itemP.stokFisik; // Perbarui nilai stok di memori
      adaPerubahan = true;

      var tipeMutasi = itemP.selisih < 0 ? "Opname (Kurang/Rusak)" : "Opname (Lebih)";
      mutasiBatch.push([
        tglHariIni,
        "'" + jamTeks,
        tipeMutasi,
        itemP.nama,
        Math.abs(itemP.selisih),
        (itemP.alasan || "Penyesuaian Opname") + " [Oleh: " + userAktif + "]"
      ]);

      auditBatch.push([
        tglHariIni + " " + jamTeks,
        userAktif,
        "OPNAME",
        "Koreksi Fisik",
        "Sistem: " + itemP.stokSistem,
        itemP.nama + " ➔ Fisik: " + itemP.stokFisik + " (" + itemP.alasan + ")"
      ]);
    }
  }

  if (adaPerubahan) {
    // OPTIMASI: Tulis seluruh baris stok yang diperbarui dalam 1 kali eksekusi
    rangeStok.setValues(vals);

    // Simpan semua log mutasi secara batch
    if (mutasiBatch.length > 0) {
      var sheetMutasi = ss.getSheetByName("Mutasi_Stok") || ss.getSheetByName("Mutasi Stok");
      if (sheetMutasi) {
        sheetMutasi.getRange(sheetMutasi.getLastRow() + 1, 1, mutasiBatch.length, 6).setValues(mutasiBatch);
      }
    }

    // Simpan semua audit log secara batch
    if (auditBatch.length > 0) {
      var sheetAudit = ss.getSheetByName("Audit_Log");
      if (sheetAudit) {
        sheetAudit.getRange(sheetAudit.getLastRow() + 1, 1, auditBatch.length, 6).setValues(auditBatch);
      }
    }
  }

  SpreadsheetApp.flush();
  return ambilDataStokLengkap();
}

// ====================================================================
// 1. HELPER KEAMANAN: HASHING PASSWORD SHA-256
// ====================================================================

function hashPassword(str) {
  if (!str) return "";
  var digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(str),
    Utilities.Charset.UTF_8
  );
  var hashHex = "";
  for (var i = 0; i < digest.length; i++) {
    var byteVal = digest[i];
    if (byteVal < 0) byteVal += 256;
    var byteHex = byteVal.toString(16);
    if (byteHex.length === 1) byteHex = "0" + byteHex;
    hashHex += byteHex;
  }
  return hashHex;
}

function isAlreadyHashed(str) {
  return typeof str === "string" && str.length === 64 && /^[0-9a-fA-F]{64}$/.test(str);
}

// ====================================================================
// 2. INISIALISASI & MIGRASI DATABASE USER SECARA AMAN
// ====================================================================

function siapkanSheetUser() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = pastikanSheetAda(ss, "Master_User", [
    "Username", 
    "Password_Hash", 
    "Nama", 
    "Role", 
    "Akses", 
    "Terakhir_Login"
  ]);

  var lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    var defaultAdminHash = hashPassword("admin123");
    sheet.appendRow(["admin", defaultAdminHash, "Owner Toko", "owner", "Semua", "-"]);
    return sheet;
  }

  // Blok auto-hashing telah dihapus agar sandi bisa dilihat sebagai teks biasa
  return sheet;
}

// ====================================================================
// 3. VERIFIKASI LOGIN DENGAN PENCOCOKAN HASH
// ====================================================================

function verifikasiLogin(username, password, portalTujuan) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = siapkanSheetUser();
  var data = sheet.getDataRange().getValues();
  
  var uInput = String(username || "").trim().toLowerCase();
  var pInput = String(password || "").trim();
  var pInputHashed = hashPassword(pInput);

  for (var i = 1; i < data.length; i++) {
    var uDb = String(data[i][0]).trim().toLowerCase();
    var pDb = String(data[i][1]).trim();
    var nama = String(data[i][2]);
    var role = String(data[i][3]).toLowerCase().trim();
    var akses = String(data[i][4]);

    var isPasswordCocok = (pDb === pInputHashed) || (pDb === pInput);

    if (uDb === uInput && isPasswordCocok) {
      // Perintah auto-hash saat login telah dihapus

      var isPetinggi = (role === "owner" || role === "super admin" || role === "superadmin");
      if (!isPetinggi && akses !== "Semua") {
        if (portalTujuan === "pembukuan" && !akses.toLowerCase().includes("pembukuan")) {
          return { success: false, message: "Akun ini tidak memiliki izin untuk modul Pembukuan." };
        }
        if (portalTujuan === "stok" && !akses.toLowerCase().includes("stok")) {
          return { success: false, message: "Akun ini tidak memiliki izin untuk modul Stok & Bundling." };
        }
      }

      var tglLogin = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm");
      sheet.getRange(i + 1, 6).setValue("'" + tglLogin);
      SpreadsheetApp.flush();

      return {
        success: true,
        user: {
          username: data[i][0],
          nama: nama,
          role: role,
          akses: akses
        }
      };
    }
  }

  return { success: false, message: "Username atau password salah!" };
}

// ====================================================================
// 4. PENYIMPANAN PENGGUNA BARU / EDIT PENGGUNA
// ====================================================================

function ambilDaftarUser() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = siapkanSheetUser();
  var data = sheet.getDataRange().getValues();
  var list = [];

  for (var i = 1; i < data.length; i++) {
    list.push({
      username: String(data[i][0]),
      password: String(data[i][1] || ""), // Mengembalikan password asli agar tombol mata berfungsi
      nama: String(data[i][2] || "-"),
      role: String(data[i][3] || "staff"),
      akses: String(data[i][4] || "Semua"),
      terakhirLogin: String(data[i][5] || "-")
    });
  }
  return JSON.stringify(list);
}

function simpanUserBaru(dataUser) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = siapkanSheetUser();
  var data = sheet.getDataRange().getValues();
  var uInput = String(dataUser.username || "").trim().toLowerCase();
  var rawPassword = String(dataUser.password || "").trim();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === uInput) {
      if (rawPassword) {
        sheet.getRange(i + 1, 2).setValue(rawPassword);
      }
      sheet.getRange(i + 1, 3).setValue(String(dataUser.nama));
      sheet.getRange(i + 1, 4).setValue(String(dataUser.role));
      sheet.getRange(i + 1, 5).setValue(String(dataUser.akses));
      SpreadsheetApp.flush();
      return ambilDaftarUser();
    }
  }

  sheet.appendRow([
    dataUser.username.trim(),
    rawPassword,
    dataUser.nama.trim(),
    dataUser.role,
    dataUser.akses,
    "-"
  ]);
  SpreadsheetApp.flush();
  return ambilDaftarUser();
}


function hapusUser(username) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = siapkanSheetUser();
  var data = sheet.getDataRange().getValues();
  var uTarget = String(username || "").trim().toLowerCase();

  if (uTarget === "admin") throw new Error("Akun admin utama tidak boleh dihapus!");

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === uTarget) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  SpreadsheetApp.flush();
  return ambilDaftarUser();
}

// ====================================================================
// FITUR: MANAJEMEN RETUR PENJUALAN & PENGEMBALIAN BARANG
// ====================================================================

/**
 * Memastikan sheet 'Retur_Penjualan' tersedia dengan struktur kolom yang rapi
 * @param {Spreadsheet} ss - Objek Spreadsheet aktif
 * @return {Sheet}
 */
function siapkanSheetRetur(ss) {
  return pastikanSheetAda(ss, "Retur_Penjualan", [
    "Tanggal",
    "Jam",
    "Nama Barang",
    "Jumlah Retur",
    "Kondisi Barang",
    "Masuk Kembali ke Stok?",
    "Platform",
    "Nominal Refund (Rp)",
    "Alasan / Keterangan",
    "User"
  ]);
}

/**
 * Memproses transaksi retur barang dari form antarmuka web
 * @param {Object} data - Objek data payload dari browser
 * @return {string} - Data JSON stok terbaru setelah disinkronkan
 */
/**
 * Memproses transaksi retur barang (mendukung Barang Satuan & Paket Bundling)
 * @param {Object} data - Objek data payload dari browser
 * @return {string} - Data JSON stok terbaru setelah disinkronkan
 */
function eksekusiReturBarang(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetStok = ss.getSheetByName("Master_Stok");
  if (!sheetStok) {
    siapkanSheetStok();
    sheetStok = ss.getSheetByName("Master_Stok");
  }

  var userAktif = data.user || "Staf";
  var tipeRetur = data.tipeRetur || "satuan"; // "satuan" atau "bundling"
  var qtyRetur = Number(data.jumlah) || 0;
  var kondisi = data.kondisi || "Layak Jual"; // "Layak Jual" atau "Rusak / Cacat"
  var masukKeStok = (kondisi === "Layak Jual" || data.kembalikanKeStok === true);
  var nominalRefund = Number(data.nominalRefund) || 0;
  var platform = data.platform || "Toko Offline";
  var alasan = data.alasan || "Retur Pelanggan";

  var tglHariIni = data.tanggal || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var jamSekarang = Utilities.formatDate(new Date(), "GMT+7", "HH:mm");
  var jamTeks = "'" + jamSekarang;

  var namaBarangDisplay = "";

  // ==========================================================
  // KASUS A: RETUR DALAM BENTUK PAKET BUNDLING / HAMPERS
  // ==========================================================
  if (tipeRetur === "bundling") {
    var namaPaket = String(data.namaPaket || data.namaBarang || "").trim();
    namaBarangDisplay = "[Bundling] " + namaPaket;
    var resep = data.resep || [];

    // Jika data resep belum dikirim lengkap dari antarmuka, cari di Master_Bundling
    if (!resep || resep.length === 0) {
      var sheetBundling = ss.getSheetByName("Master_Bundling");
      if (sheetBundling) {
        var bdlRows = sheetBundling.getDataRange().getValues();
        for (var b = 1; b < bdlRows.length; b++) {
          if (String(bdlRows[b][1]).trim().toLowerCase() === namaPaket.toLowerCase() || String(bdlRows[b][0]).trim() === String(data.idBundling).trim()) {
            try { resep = JSON.parse(bdlRows[b][2]); } catch(e) {}
            break;
          }
        }
      }
    }

    var lastRowStok = sheetStok.getLastRow();
    if (lastRowStok > 1) {
      var rangeStok = sheetStok.getRange(2, 1, lastRowStok - 1, sheetStok.getLastColumn());
      var dataStok = rangeStok.getValues();
      var mutasiBatch = [];

      // Telusuri resep komponen pembentuk paket bundling
      for (var r = 0; r < resep.length; r++) {
        var butuhTotal = (Number(resep[r].butuh) || 1) * qtyRetur;
        for (var s = 0; s < dataStok.length; s++) {
          if (String(dataStok[s][0]) === String(resep[r].kode)) {
            if (masukKeStok) {
              // Pulihkan stok komponen ke Master_Stok
              dataStok[s][3] = (Number(dataStok[s][3]) || 0) + butuhTotal;
              var ketBahan = "Komponen " + qtyRetur + "x " + namaPaket + " Kembali (" + platform + ") - " + alasan + " [Oleh: " + userAktif + "]";
              mutasiBatch.push([tglHariIni, jamTeks, "Bahan Bundling Kembali (Retur)", dataStok[s][1], butuhTotal, ketBahan]);
            } else {
              var ketBahanRusak = "Komponen " + qtyRetur + "x " + namaPaket + " Rusak/Afkir (" + platform + ") - " + alasan + " [Oleh: " + userAktif + "]";
              mutasiBatch.push([tglHariIni, jamTeks, "Komponen Bundling Rusak", dataStok[s][1], butuhTotal, ketBahanRusak]);
            }
            break;
          }
        }
      }

      // Tulis kembali stok komponen yang dipulihkan ke Master_Stok secara batch
      if (masukKeStok) {
        rangeStok.setValues(dataStok);
      }

      // Catat ringkasan retur paket bundling ke log mutasi
      mutasiBatch.push([
        tglHariIni,
        jamTeks,
        masukKeStok ? "Retur Bundling (Bagus)" : "Retur Bundling (Rusak)",
        namaPaket,
        qtyRetur,
        "Retur " + qtyRetur + " paket via " + platform + " - " + alasan + " [Oleh: " + userAktif + "]"
      ]);

      var sheetMutasi = ss.getSheetByName("Mutasi_Stok") || ss.getSheetByName("Mutasi Stok");
      if (sheetMutasi && mutasiBatch.length > 0) {
        sheetMutasi.getRange(sheetMutasi.getLastRow() + 1, 1, mutasiBatch.length, 6).setValues(mutasiBatch);
      }
    }

  // ==========================================================
  // KASUS B: RETUR DALAM BENTUK BARANG SATUAN
  // ==========================================================
  } else {
    var kodeBarang = String(data.kodeBarang || "").trim();
    var namaBarang = String(data.namaBarang || "").trim();
    namaBarangDisplay = namaBarang;

    if (masukKeStok && qtyRetur > 0) {
      var dataStok = sheetStok.getDataRange().getValues();
      for (var i = 1; i < dataStok.length; i++) {
        var kodeDiSheet = String(dataStok[i][0]).trim();
        var namaDiSheet = String(dataStok[i][1]).trim().toLowerCase();

        if (kodeDiSheet === kodeBarang || namaDiSheet === namaBarang.toLowerCase()) {
          var stokLama = Number(dataStok[i][3]) || 0;
          var stokBaru = stokLama + qtyRetur;
          sheetStok.getRange(i + 1, 4).setValue(stokBaru);

          catatMutasiStok(
            "Retur Masuk (Restock)",
            dataStok[i][1],
            qtyRetur,
            "Kondisi Layak Jual (" + platform + ") - " + alasan + " [Oleh: " + userAktif + "]"
          );
          break;
        }
      }
    } else if (!masukKeStok && qtyRetur > 0) {
      catatMutasiStok(
        "Retur Barang Rusak",
        namaBarang,
        qtyRetur,
        "Barang Rusak/Afkir (" + platform + ") - " + alasan + " [Oleh: " + userAktif + "]"
      );
    }
  }

  // 1. Catat transaksi ke sheet 'Retur_Penjualan'
  var sheetRetur = siapkanSheetRetur(ss);
  sheetRetur.appendRow([
    tglHariIni,
    jamTeks,
    namaBarangDisplay,
    qtyRetur,
    kondisi,
    masukKeStok ? "Ya (Restock)" : "Tidak (Barang Rusak)",
    platform,
    nominalRefund,
    alasan,
    userAktif
  ]);

  // 2. Catat Kas Keluar jika ada dana refund yang ditransfer balik ke pelanggan
  if (nominalRefund > 0) {
    var sheetAK = ss.getSheetByName("Arus Kas") || pastikanSheetAda(ss, "Arus Kas", ["Tanggal", "Keterangan", "Tipe", "Nominal", "Jam"]);
    var ketKas = "[Retur Refund] " + namaBarangDisplay + " (" + qtyRetur + " " + (tipeRetur === "bundling" ? "paket" : "pcs") + " via " + platform + ")";
    sheetAK.appendRow([tglHariIni, ketKas, "Kas Keluar", nominalRefund, jamTeks]);
  }

  // 3. Catat ke Audit Log
  catatAuditLog(
    "RETUR",
    tipeRetur === "bundling" ? "Bundling & Gudang" : "Penjualan & Gudang",
    "-",
    namaBarangDisplay + " (" + qtyRetur + " item, Refund: Rp " + nominalRefund.toLocaleString('id-ID') + ")",
    userAktif
  );

  SpreadsheetApp.flush();
  return ambilDataStokLengkap();
}

// ====================================================================
// FITUR: PENJUALAN MULTI-BUNDLING / MULTI-ITEM CHECKOUT MARKETPLACE
// ====================================================================

/**
 * Memproses transaksi multi-bundling:
 * 1. Memotong stok bahan secara batch di memori.
 * 2. Mencatat ke Pemasukan dan Arus Kas dengan format standar sistem.
 * 3. Mengembalikan data stok & pembukuan terbaru sekaligus.
 */
function eksekusiPenjualanMultiBundling(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetStok = ss.getSheetByName("Master_Stok");
  if (!sheetStok) return ambilDataStokLengkap();

  var lastRowStok = sheetStok.getLastRow();
  if (lastRowStok <= 1) return ambilDataStokLengkap();

  var rangeStok = sheetStok.getRange(2, 1, lastRowStok - 1, sheetStok.getLastColumn());
  var dataStok = rangeStok.getValues();

  var daftarItem = data.daftarItem || [];
  var tglPilihan = (data.tanggal && String(data.tanggal).trim() !== "") 
    ? String(data.tanggal).trim() 
    : Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");

  var platform = data.platform || "Shopee";
  var tokoField = (data.toko && String(data.toko).trim() !== "") ? String(data.toko).trim() : platform;
  var pembeli = (data.namaPembeli && String(data.namaPembeli).trim() !== "") ? String(data.namaPembeli).trim() : "";
  var userAktif = data.user || "Staf";

  // Perhitungan Subtotal dan Dana Bersih dengan Pengaman (Fallback Calculation)
  var subtotalBruto = Number(data.subtotalBruto) || 0;
  var totalPotonganAdmin = Number(data.totalPotonganAdmin) || 0;
  var totalBersihCair = Number(data.totalBersihCair);

  // PERBAIKAN: Hapus syarat `<= 0` dan hapus `Math.max` agar nilai minus bisa masuk ke Arus Kas
  if (isNaN(totalBersihCair)) {
    var hitungBrutoOtomatis = 0;
    for (var k = 0; k < daftarItem.length; k++) {
      hitungBrutoOtomatis += (Number(daftarItem[k].qty) || 1) * (Number(daftarItem[k].hargaSatuan) || 0);
    }
    subtotalBruto = hitungBrutoOtomatis;
    totalBersihCair = subtotalBruto - totalPotonganAdmin;
  }

  var jamSekarang = Utilities.formatDate(new Date(), "GMT+7", "HH:mm");
  var jamTeks = "'" + jamSekarang;
  var mutasiBatch = [];
  var ringkasanNamaPaket = [];

 // 1. PROSES PEMOTONGAN SELURUH BAHAN RESEP DARI SEMUA PAKET DENGAN VALIDASI
  for (var i = 0; i < daftarItem.length; i++) {
    var itemPesanan = daftarItem[i];
    var namaPaket = itemPesanan.namaPaket;
    var qtyJual = Number(itemPesanan.qty) || 1;
    var resep = itemPesanan.resep || [];

    ringkasanNamaPaket.push(qtyJual + "x " + namaPaket);

    for (var r = 0; r < resep.length; r++) {
      var butuhTotal = (Number(resep[r].butuh) || 1) * qtyJual;
      for (var s = 0; s < dataStok.length; s++) {
        if (String(dataStok[s][0]) === String(resep[r].kode)) {
          var stokSekarang = Number(dataStok[s][3]) || 0;
          
          // VALIDASI SERVER: Cegah stok minus pada Multi-Bundling
          if (stokSekarang < butuhTotal) {
            throw new Error("Transaksi ditolak! Stok bahan " + dataStok[s][1] + " untuk pesanan paket " + namaPaket + " tidak mencukupi (Sisa: " + stokSekarang + " pcs).");
          }
          
          dataStok[s][3] = stokSekarang - butuhTotal;

          var ketBahan = "Komponen " + qtyJual + "x " + namaPaket + " (" + tokoField + ")";
          if (pembeli) ketBahan += " - Pesanan: " + pembeli;

          mutasiBatch.push([tglPilihan, jamTeks, "Bahan Bundling Terpakai", dataStok[s][1], butuhTotal, ketBahan]);
          break;
        }
      }
    }

    // Hitung proporsi nilai bersih (net) khusus untuk paket yang ditarik ini
    var subtotalItem = qtyJual * (Number(itemPesanan.hargaSatuan) || 0);
    var proporsi = (subtotalBruto > 0) ? (subtotalItem / subtotalBruto) : 0;
    var netItem = Math.round(totalBersihCair * proporsi);

    var ketPaket = "Penjualan via " + tokoField;
    if (pembeli) ketPaket += " - Pembeli: " + pembeli;
    
    // PERBAIKAN: Menyisipkan "[Net: 123]"
    ketPaket += " [Net: " + netItem + "]"; 
    mutasiBatch.push([tglPilihan, jamTeks, "Penjualan Bundling", namaPaket, qtyJual, ketPaket]);
  }

  rangeStok.setValues(dataStok);

  var sheetMutasi = ss.getSheetByName("Mutasi_Stok") || ss.getSheetByName("Mutasi Stok");
  if (sheetMutasi && mutasiBatch.length > 0) {
    sheetMutasi.getRange(sheetMutasi.getLastRow() + 1, 1, mutasiBatch.length, 6).setValues(mutasiBatch);
  }

  // 2. CATAT TRANSAKSI KEUANGAN DENGAN FORMAT STANDAR SISTEM
  var isCatatBuku = (data.catatKePembukuan !== false && data.catatKePembukuan !== "false");
      // PERBAIKAN: Syarat > 0 dihapus agar jika total bersihnya minus, tetap memotong saldo Pemasukan/Arus Kas
      if (isCatatBuku) {
    var ringkasPaketStr = ringkasanNamaPaket.join(", ");
    var namaTransaksi = pembeli ? (pembeli + " - " + ringkasPaketStr) : ringkasPaketStr;

    var ketPemasukan = namaTransaksi;
    if (totalPotonganAdmin > 0) {
      ketPemasukan += " [Net] (Bruto: Rp " + subtotalBruto.toLocaleString('id-ID') + " - Potongan: Rp " + totalPotonganAdmin.toLocaleString('id-ID') + ")";
    }

    // A. Masuk ke sheet Pemasukan
    var sheetPemasukan = ss.getSheetByName("Pemasukan") || pastikanSheetAda(ss, "Pemasukan", ["Nama", "Tanggal", "Nominal", "Toko", "Jam"]);
    sheetPemasukan.appendRow([ketPemasukan, tglPilihan, totalBersihCair, tokoField, jamTeks]);

    // B. Masuk ke Arus Kas dengan awalan baku "[Pemasukan] " agar terdeteksi oleh sistem laporan
    var sheetAK = ss.getSheetByName("Arus Kas") || pastikanSheetAda(ss, "Arus Kas", ["Tanggal", "Keterangan", "Tipe", "Nominal", "Jam"]);
    sheetAK.appendRow([tglPilihan, "[Pemasukan] " + namaTransaksi, "Kas Masuk", totalBersihCair, jamTeks]);
  }

  // 3. Catat ke Audit Log
  catatAuditLog(
    "JUAL MULTI-BUNDLING",
    "Penjualan & Gudang",
    "-",
    ringkasanNamaPaket.join(" | ") + " (Net: Rp " + totalBersihCair.toLocaleString('id-ID') + ")",
    userAktif
  );

  SpreadsheetApp.flush();

  // KEMBALIKAN DATA LENGKAP: STOK GUDANG + BUKU PEMBUKUAN TERBARU
  return JSON.stringify({
    success: true,
    stokData: JSON.parse(ambilDataStokLengkap()),
    bukuData: ambilSemuaData(false)
  });
}

// ====================================================================
// --- FUNGSI TESTING MANDIRI SISTEM (SERVER) ---
// ====================================================================
function test_verifikasiStrukturData() {
  try {
    var dataBuku = ambilSemuaData(false);
    Logger.log("Pemasukan baris: " + (dataBuku.pemasukan ? dataBuku.pemasukan.length : 0));
    Logger.log("Arus Kas baris: " + (dataBuku.arusKas ? dataBuku.arusKas.length : 0));
    
    var dataStok = JSON.parse(ambilDataStokLengkap());
    Logger.log("Master Stok SKU: " + (dataStok.stok ? dataStok.stok.length : 0));
    Logger.log("Mutasi Stok log: " + (dataStok.mutasi ? dataStok.mutasi.length : 0));
    Logger.log("✅ Semua struktur data terbaca sempurna tanpa error.");
  } catch(e) {
    Logger.log("❌ Error saat verifikasi: " + e.message);
  }
}



// ====================================================================
// FITUR BARU: BATALKAN MUTASI & PULIHKAN STOK (AUTO-SYNC KEUANGAN & BUNDLING)
// ====================================================================
function batalkanMutasiStok(rowIndex, user) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetMutasi = ss.getSheetByName("Mutasi_Stok") || ss.getSheetByName("Mutasi Stok");
  if (!sheetMutasi) throw new Error("Sheet Mutasi_Stok tidak ditemukan");

  // 1. Baca baris mutasi yang akan dihapus
  var rowValues = sheetMutasi.getRange(rowIndex, 1, 1, 6).getValues()[0];
  var tglRaw = rowValues[0];
  var tgl = (tglRaw instanceof Date) ? Utilities.formatDate(tglRaw, "GMT+7", "yyyy-MM-dd") : String(tglRaw).split('T')[0];
  var jam = String(rowValues[1] || "").replace(/'/g, "").trim();
  var tipe = String(rowValues[2] || "");
  var namaBarang = String(rowValues[3] || "");
  var qty = Number(rowValues[4]) || 0;
  var ketMutasiAsli = String(rowValues[5] || "");
  var userAktif = user || "Owner";

  // Ekstrak nama pembeli dari keterangan mutasi jika ada (misal: "Pembeli: Kak Nisa")
  var pembeliMutasi = "";
  if (ketMutasiAsli.includes("Pembeli:")) {
    pembeliMutasi = ketMutasiAsli.split("Pembeli:")[1].split("[")[0].trim();
  } else if (ketMutasiAsli.includes("Pesanan:")) {
    pembeliMutasi = ketMutasiAsli.split("Pesanan:")[1].split("[")[0].trim();
  }

  var sheetStok = ss.getSheetByName("Master_Stok");
  var dataStok = sheetStok ? sheetStok.getDataRange().getValues() : [];

  // ==========================================================
  // 2. KEMBALIKAN STOK FISIK DI GUDANG (LOGIKA BUNDLING VS SATUAN)
  // ==========================================================
  var isBundling = tipe.toLowerCase().includes("bundling");

  if (isBundling && (tipe === "Penjualan Bundling" || tipe.indexOf("Retur Bundling") === 0)) {
    // A. BONGKAR RESEP BUNDLING
    var sheetBdl = ss.getSheetByName("Master_Bundling");
    var resep = [];
    if (sheetBdl) {
      var dataBdl = sheetBdl.getDataRange().getValues();
      for (var b = 1; b < dataBdl.length; b++) {
        if (String(dataBdl[b][1]).trim().toLowerCase() === namaBarang.toLowerCase()) {
          try { resep = JSON.parse(dataBdl[b][2]); } catch(e){}
          break;
        }
      }
    }

    // B. KEMBALIKAN ATAU TARIK BAHAN BAKU DI MASTER STOK
    for (var r = 0; r < resep.length; r++) {
      var butuhTotal = (Number(resep[r].butuh) || 1) * qty;
      for (var s = 1; s < dataStok.length; s++) {
        if (String(dataStok[s][0]) === String(resep[r].kode)) {
          var stokSaatIni = Number(dataStok[s][3]) || 0;
          if (tipe === "Penjualan Bundling") {
            sheetStok.getRange(s + 1, 4).setValue(stokSaatIni + butuhTotal);
          } else if (tipe === "Retur Bundling (Bagus)") {
            sheetStok.getRange(s + 1, 4).setValue(Math.max(0, stokSaatIni - butuhTotal));
          }
          break;
        }
      }
    }

    // C. HAPUS BARIS KOMPONEN BUNDLING TERKAIT (Yang jam, tanggal, & nama paketnya cocok)
    var mutasiVals = sheetMutasi.getDataRange().getValues();
    for (var m = mutasiVals.length - 1; m >= 1; m--) {
      var mTglRaw = mutasiVals[m][0];
      var mTgl = (mTglRaw instanceof Date) ? Utilities.formatDate(mTglRaw, "GMT+7", "yyyy-MM-dd") : String(mTglRaw).split('T')[0];
      var mJam = String(mutasiVals[m][1] || "").replace(/'/g, "").trim();
      var mTipe = String(mutasiVals[m][2] || "");
      var mKet = String(mutasiVals[m][5] || "");
      
      var isKomponenTerkait = (mTipe === "Bahan Bundling Terpakai" || mTipe === "Bahan Bundling Kembali (Retur)" || mTipe === "Komponen Bundling Rusak");
      if (mTgl === tgl && mJam === jam && isKomponenTerkait && mKet.includes(namaBarang)) {
        sheetMutasi.deleteRow(m + 1);
      }
    }

  } else {
    // LOGIKA NORMAL: PENGEMBALIAN BARANG SATUAN
    for (var i = 1; i < dataStok.length; i++) {
      if (String(dataStok[i][1]).trim().toLowerCase() === namaBarang.toLowerCase()) {
        var stokSaatIni = Number(dataStok[i][3]) || 0;
        var stokBaru = stokSaatIni;
        var tLow = tipe.toLowerCase();
        
        if (tLow.includes("keluar") || tLow.includes("jual") || tLow.includes("terpakai") || tLow.includes("kurang") || tLow.includes("rusak")) {
          stokBaru = stokSaatIni + qty;
        } 
        else if (tLow.includes("masuk") || tLow.includes("restock") || tLow.includes("awal") || tLow.includes("lebih") || tLow.includes("kembali") || tLow.includes("retur")) {
          stokBaru = Math.max(0, stokSaatIni - qty);
        }
        
        sheetStok.getRange(i + 1, 4).setValue(stokBaru);
        break;
      }
    }
  }

  // ==========================================================
  // 3. SAPU BERSIH TRANSAKSI FINANSIAL, NOTA POS, & RETUR TERKAIT
  // ==========================================================
  var sheetsKeuangan = ["Pemasukan", "Pembelian Shopee", "Pembelian Agen", "Pengeluaran Lain", "Arus Kas"];
  sheetsKeuangan.forEach(function(sName) {
    var s = ss.getSheetByName(sName);
    if (s && s.getLastRow() > 1) {
      var vals = s.getDataRange().getValues();
      for (var r = vals.length - 1; r >= 1; r--) {
         var tglRRaw = vals[r][sName === "Arus Kas" ? 0 : 1];
         var tglR = (tglRRaw instanceof Date) ? Utilities.formatDate(tglRRaw, "GMT+7", "yyyy-MM-dd") : String(tglRRaw).split('T')[0];
         var jamR = String(vals[r][4] || "").replace(/'/g, "").trim();
         var ketR = String(vals[r][sName === "Arus Kas" ? 1 : 0] || "").toLowerCase();
         
         if (tglR === tgl && jamR === jam && (ketR.includes(namaBarang.toLowerCase()) || (pembeliMutasi && ketR.includes(pembeliMutasi.toLowerCase())))) {
            s.deleteRow(r + 1);
         }
      }
    }
  });

  // Hapus dari Detail_Penjualan (Nota Kasir POS) jika ada
  var sheetDetail = ss.getSheetByName("Detail_Penjualan");
  if (sheetDetail && sheetDetail.getLastRow() > 1) {
    var dVals = sheetDetail.getDataRange().getValues();
    for (var d = dVals.length - 1; d >= 1; d--) {
      var dTglRaw = dVals[d][0];
      var dTgl = (dTglRaw instanceof Date) ? Utilities.formatDate(dTglRaw, "GMT+7", "yyyy-MM-dd") : String(dTglRaw).split('T')[0];
      var dJam = String(dVals[d][1] || "").replace(/'/g, "").trim();
      var dPembeli = String(dVals[d][2] || "").toLowerCase();
      var dProduk = String(dVals[d][3] || "").toLowerCase();

      if (dTgl === tgl && dJam === jam && (dProduk.includes(namaBarang.toLowerCase()) || (pembeliMutasi && dPembeli.includes(pembeliMutasi.toLowerCase())))) {
        sheetDetail.deleteRow(d + 1);
        break;
      }
    }
  }

  // Hapus dari Retur_Penjualan jika mutasi yang dibatalkan berasal dari Retur
  var sheetRetur = ss.getSheetByName("Retur_Penjualan");
  if (sheetRetur && sheetRetur.getLastRow() > 1) {
    var rVals = sheetRetur.getDataRange().getValues();
    for (var rt = rVals.length - 1; rt >= 1; rt--) {
      var rTglRaw = rVals[rt][0];
      var rTgl = (rTglRaw instanceof Date) ? Utilities.formatDate(rTglRaw, "GMT+7", "yyyy-MM-dd") : String(rTglRaw).split('T')[0];
      var rJam = String(rVals[rt][1] || "").replace(/'/g, "").trim();
      var rNama = String(rVals[rt][2] || "").toLowerCase();

      if (rTgl === tgl && rJam === jam && rNama.includes(namaBarang.toLowerCase())) {
        sheetRetur.deleteRow(rt + 1);
        break;
      }
    }
  }

  // ==========================================================
  // 4. HAPUS BARIS UTAMA DI MUTASI STOK
  // ==========================================================
  var mNewVals = sheetMutasi.getDataRange().getValues();
  for (var mn = mNewVals.length - 1; mn >= 1; mn--) {
      var mnTglRaw = mNewVals[mn][0];
      var mnTgl = (mnTglRaw instanceof Date) ? Utilities.formatDate(mnTglRaw, "GMT+7", "yyyy-MM-dd") : String(mnTglRaw).split('T')[0];
      var mnJam = String(mNewVals[mn][1] || "").replace(/'/g, "").trim();
      var mnTipe = String(mNewVals[mn][2] || "");
      var mnNama = String(mNewVals[mn][3] || "");
      
      if (mnTgl === tgl && mnJam === jam && mnTipe === tipe && mnNama === namaBarang) {
          sheetMutasi.deleteRow(mn + 1);
          break;
      }
  }

  // 5. CATAT TINDAKAN INI KE AUDIT LOG
  catatAuditLog("BATAL MUTASI", "Stok & Keuangan", tipe + " " + qty + " pcs/paket (" + namaBarang + ")", "Dibatalkan & Stok Dipulihkan", userAktif);

  SpreadsheetApp.flush();
  
  return JSON.stringify({
     success: true,
     stokData: JSON.parse(ambilDataStokLengkap()),
     bukuData: ambilSemuaData(false)
  });
}

// ====================================================================
// FITUR EDIT & HAPUS MASTER STOK
// ====================================================================
function editDataMasterStok(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Master_Stok");
  if (!sheet) throw new Error("Sheet Master_Stok tidak ditemukan");
  
  var data = sheet.getDataRange().getValues();
  var rowToEdit = -1;
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(payload.kode)) {
      rowToEdit = i + 1;
      break;
    }
  }
  
  if (rowToEdit !== -1) {
    // Simpan nama lama sebagai alias SKU sebelum nama Master ditimpa.
    var namaSebelumEdit = String(data[rowToEdit - 1][1] || "").trim();
    var namaSesudahEdit = String(payload.nama || "").trim();
    if (namaSebelumEdit && normalisasiNamaSkuServer_(namaSebelumEdit) !== normalisasiNamaSkuServer_(namaSesudahEdit)) {
      simpanAliasNamaSku_(payload.kode, namaSebelumEdit, namaSesudahEdit, "Perubahan Nama Master Stok");
    }

    // Menimpa Kolom B (Nama), C (Kategori), dan G (Min Stok)
    sheet.getRange(rowToEdit, 2).setValue(payload.nama);
    sheet.getRange(rowToEdit, 3).setValue(payload.kategori);
    sheet.getRange(rowToEdit, 7).setValue(payload.minStok);
    
    catatAuditLog("EDIT BARANG", "Master Stok", payload.kode + " - " + payload.nama, "Nama/Kategori/Min Stok diperbarui", payload.user);
  } else {
    throw new Error("Kode barang tidak ditemukan di database.");
  }
  
  SpreadsheetApp.flush();
  return ambilDataStokLengkap(); // Kembalikan data segar ke UI
}

function hapusDataMasterStok(kode, user) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Master_Stok");
  if (!sheet) throw new Error("Sheet Master_Stok tidak ditemukan");
  
  var data = sheet.getDataRange().getValues();
  var rowToDelete = -1;
  var namaBarang = "";
  var kodeBersih = String(kode || "").trim();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === kodeBersih) {
      rowToDelete = i + 1;
      namaBarang = String(data[i][1]).trim();
      break;
    }
  }
  
  if (rowToDelete !== -1) {
    // 1. Hapus dari Master_Stok
    sheet.deleteRow(rowToDelete);

    // 2. Bersihkan kode barang ini jika terdaftar di dalam Resep Master_Bundling
    var sheetBdl = ss.getSheetByName("Master_Bundling");
    if (sheetBdl && sheetBdl.getLastRow() > 1) {
      var dataBdl = sheetBdl.getDataRange().getValues();
      for (var b = 1; b < dataBdl.length; b++) {
        try {
          var resepArr = JSON.parse(dataBdl[b][2] || "[]");
          var resepBersih = resepArr.filter(function(item) {
            return String(item.kode || "").trim() !== kodeBersih;
          });
          if (resepBersih.length !== resepArr.length) {
            sheetBdl.getRange(b + 1, 3).setValue(JSON.stringify(resepBersih));
          }
        } catch (e) {}
      }
    }

    // 3. Bersihkan riwayat Stok Awal / Restock barang ini di Mutasi_Stok
    var sheetMutasi = ss.getSheetByName("Mutasi_Stok") || ss.getSheetByName("Mutasi Stok");
    if (sheetMutasi && sheetMutasi.getLastRow() > 1 && namaBarang) {
      var mVals = sheetMutasi.getDataRange().getValues();
      for (var m = mVals.length - 1; m >= 1; m--) {
        if (String(mVals[m][3] || "").trim().toLowerCase() === namaBarang.toLowerCase()) {
          sheetMutasi.deleteRow(m + 1);
        }
      }
    }

    // 4. Bersihkan transaksi belanja Restock barang ini di Pembukuan & Arus Kas
    if (namaBarang) {
      var namaLower = namaBarang.toLowerCase();
      var sheetsBeban = ["Pembelian Agen", "Pembelian Shopee", "Pengeluaran Lain", "Arus Kas"];
      sheetsBeban.forEach(function(sName) {
        var sBeban = ss.getSheetByName(sName);
        if (sBeban && sBeban.getLastRow() > 1) {
          var bVals = sBeban.getDataRange().getValues();
          for (var r = bVals.length - 1; r >= 1; r--) {
            var ketCol = (sName === "Arus Kas") ? 1 : 0;
            var ketTeks = String(bVals[r][ketCol] || "").toLowerCase();
            if (ketTeks.includes("restock") && ketTeks.includes(namaLower)) {
              sBeban.deleteRow(r + 1);
            }
          }
        }
      });
    }

    catatAuditLog("HAPUS BARANG", "Master Stok", kodeBersih + " - " + namaBarang, "Dihapus permanen beserta relasi mutasi & bundling", user);
  } else {
    throw new Error("Kode barang tidak ditemukan untuk dihapus.");
  }
  
  SpreadsheetApp.flush();
  return ambilDataStokLengkap(); // Kembalikan data segar ke UI
}
/**
 * Memproses transaksi checkout dari Kasir POS secara batch:
 * - Memvalidasi stok fisik sebelum diproses (LAPIS 3)
 * - Memotong stok & mencatat mutasi
 * - Mencatat Arus Kas, Pemasukan, dan Detail_Penjualan
 */
function prosesCheckoutPOS(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    if (!payload || !payload.keranjang || payload.keranjang.length === 0) {
      throw new Error("Keranjang pesanan kosong atau tidak valid.");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = Session.getScriptTimeZone() || "GMT+7";

    var sheetStok = ss.getSheetByName("Master_Stok");
    if (!sheetStok) {
      siapkanSheetStok();
      sheetStok = ss.getSheetByName("Master_Stok");
    }

    var dataStok = sheetStok.getDataRange().getValues();
    var sheetBdl = ss.getSheetByName("Master_Bundling");
    var dataBdl = sheetBdl ? sheetBdl.getDataRange().getValues() : [];

    // ====================================================================
    // --- LAPIS 3: VALIDASI STOK SERVER (Mencegah Stok Minus) ---
    // ====================================================================
    var kebutuhanStok = {}; // Menyimpan total kebutuhan tiap kode barang

    // 1. Kumpulkan semua kebutuhan riil dari keranjang
    for (var v = 0; v < payload.keranjang.length; v++) {
      var itemVal = payload.keranjang[v];
      var qtyVal = Math.max(1, Number(itemVal.qty) || 1);

      if (itemVal.tipe === "Bundling") {
        var resepVal = [];
        for (var b = 1; b < dataBdl.length; b++) {
          if (String(dataBdl[b][0]).trim() === String(itemVal.id).trim() ||
              String(dataBdl[b][1]).trim().toLowerCase() === String(itemVal.nama).trim().toLowerCase()) {
            try { resepVal = JSON.parse(dataBdl[b][2]); } catch (e) {}
            break;
          }
        }
        // Urai bahan baku resepnya
        for (var r = 0; r < resepVal.length; r++) {
          var kodeBahan = String(resepVal[r].kode).trim();
          var butuhTotal = (Number(resepVal[r].butuh) || 1) * qtyVal;
          kebutuhanStok[kodeBahan] = (kebutuhanStok[kodeBahan] || 0) + butuhTotal;
        }
      } else {
        var idSatuan = String(itemVal.id).trim();
        kebutuhanStok[idSatuan] = (kebutuhanStok[idSatuan] || 0) + qtyVal;
      }
    }

    // 2. Cocokkan dengan stok aktual di Master_Stok saat ini
    for (var s = 1; s < dataStok.length; s++) {
      var idGudang = String(dataStok[s][0]).trim();
      var namaGudang = String(dataStok[s][1]).trim();
      var stokGudang = Number(dataStok[s][3]) || 0;

      if (kebutuhanStok[idGudang] !== undefined) {
        if (kebutuhanStok[idGudang] > stokGudang) {
          // CEGAT DAN BATALKAN JIKA STOK KURANG!
          throw new Error("TRANSAKSI DITOLAK: Stok [" + namaGudang + "] tidak mencukupi! Dibutuhkan " + kebutuhanStok[idGudang] + " pcs, namun sisa di gudang saat ini hanya " + stokGudang + " pcs.");
        }
      }
    }
    // ====================================================================

    // Jika lolos validasi, lanjutkan proses seperti biasa
    var tglPilihan = (payload.tanggal && String(payload.tanggal).trim() !== "")
      ? String(payload.tanggal).trim()
      : Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    var jamSekarang = Utilities.formatDate(new Date(), tz, "HH:mm");
    var jamTeks = "'" + jamSekarang;

    var userAktif = payload.user || "Kasir";
    var pembeli = (payload.pembeli && String(payload.pembeli).trim() !== "")
      ? String(payload.pembeli).trim()
      : "Pembeli (POS)";
    var platform = (payload.platform && String(payload.platform).trim() !== "")
      ? String(payload.platform).trim()
      : "Toko Offline";

    var keranjang = payload.keranjang;
    var potonganAdmin = Number(payload.potonganAdmin) || 0;

    var hitungBrutoKeranjang = 0;
    for (var i = 0; i < keranjang.length; i++) {
      var qItem = Math.max(1, Number(keranjang[i].qty) || 1);
      var hItem = Number(keranjang[i].harga) || 0;
      hitungBrutoKeranjang += (qItem * hItem);
    }

    var subtotalBruto = !isNaN(Number(payload.subtotalBruto))
      ? Number(payload.subtotalBruto)
      : hitungBrutoKeranjang;

    var totalBersih = !isNaN(Number(payload.totalBersihCair))
      ? Number(payload.totalBersihCair)
      : (subtotalBruto - potonganAdmin);

    var mutasiBatch = [];
    var ringkasanBarang = [];
    var totalHppTransaksi = 0; 

    // Proses Pemotongan Stok Fisik (Pasti Aman karena sudah divalidasi di atas)
    for (var k = 0; k < keranjang.length; k++) {
      var item = keranjang[k];
      var qtyJual = Math.max(1, Number(item.qty) || 1);
      var hargaJualItem = Number(item.harga) || 0;
      var subItem = qtyJual * hargaJualItem;

      var proporsi = (subtotalBruto > 0) ? (subItem / subtotalBruto) : (1 / keranjang.length);
      var netItem = Math.round(totalBersih * proporsi);

      ringkasanBarang.push(qtyJual + "x " + item.nama);

      if (item.tipe === "Bundling") {
        var resep = [];
        for (var bl = 1; bl < dataBdl.length; bl++) {
          if (String(dataBdl[bl][0]).trim() === String(item.id).trim() ||
              String(dataBdl[bl][1]).trim().toLowerCase() === String(item.nama).trim().toLowerCase()) {
            try { resep = JSON.parse(dataBdl[bl][2]); } catch (errJson) {}
            break;
          }
        }

        var hppBundlingIni = 0;
        for (var r2 = 0; r2 < resep.length; r2++) {
          var butuhTotal2 = (Number(resep[r2].butuh) || 1) * qtyJual;
          for (var s2 = 1; s2 < dataStok.length; s2++) {
            if (String(dataStok[s2][0]).trim() === String(resep[r2].kode).trim()) {
              var stokAsalBdl = Number(dataStok[s2][3]) || 0;
              dataStok[s2][3] = Math.max(0, stokAsalBdl - butuhTotal2); // Aman dipotong
              hppBundlingIni += (Number(dataStok[s2][5]) || 0) * butuhTotal2;

              var ketBahan = "Komponen " + qtyJual + "x " + item.nama + " (" + platform + ") - Pembeli: " + pembeli;
              mutasiBatch.push([tglPilihan, jamTeks, "Bahan Bundling Terpakai", dataStok[s2][1], butuhTotal2, ketBahan]);
              break;
            }
          }
        }
        totalHppTransaksi += hppBundlingIni;

        var ketPaket = "Penjualan via " + platform + " - Pembeli: " + pembeli + " [Net: " + netItem + "]";
        mutasiBatch.push([tglPilihan, jamTeks, "Penjualan Bundling", item.nama, qtyJual, ketPaket]);

      } else {
        var hppSatuanIni = 0;
        for (var st = 1; st < dataStok.length; st++) {
          if (String(dataStok[st][0]).trim() === String(item.id).trim() ||
              String(dataStok[st][1]).trim().toLowerCase() === String(item.nama).trim().toLowerCase()) {
            var stokAsalSat = Number(dataStok[st][3]) || 0;
            dataStok[st][3] = Math.max(0, stokAsalSat - qtyJual); // Aman dipotong
            hppSatuanIni = (Number(dataStok[st][5]) || 0) * qtyJual;
            break;
          }
        }
        totalHppTransaksi += hppSatuanIni;

        var ketSatuan = "Terjual " + qtyJual + " pcs (" + platform + ") - Pembeli: " + pembeli + " [Net: " + netItem + "] [Oleh: " + userAktif + "]";
        mutasiBatch.push([tglPilihan, jamTeks, "Penjualan Satuan", item.nama, qtyJual, ketSatuan]);
      }
    }

    // Eksekusi Simpan Data ke Berbagai Sheet
    if (dataStok.length > 1) {
      sheetStok.getRange(1, 1, dataStok.length, dataStok[0].length).setValues(dataStok);
    }
    if (mutasiBatch.length > 0) {
      var sheetMutasi = ss.getSheetByName("Mutasi_Stok") || pastikanSheetAda(ss, "Mutasi_Stok", ["Tanggal", "Jam", "Tipe", "Nama Barang", "Jumlah", "Keterangan"]);
      sheetMutasi.getRange(sheetMutasi.getLastRow() + 1, 1, mutasiBatch.length, 6).setValues(mutasiBatch);
    }

    var ringkasStr = ringkasanBarang.join(", ");
    var namaTransaksi = pembeli + " - " + ringkasStr;
    var ketPemasukan = namaTransaksi;

    if (potonganAdmin > 0) {
      ketPemasukan += " [Net] (Bruto: Rp " + subtotalBruto.toLocaleString('id-ID') + " - Potongan: Rp " + potonganAdmin.toLocaleString('id-ID') + ")";
    }

    var sheetPemasukan = ss.getSheetByName("Pemasukan") || pastikanSheetAda(ss, "Pemasukan", ["Nama", "Tanggal", "Nominal", "Toko", "Jam"]);
    sheetPemasukan.appendRow([ketPemasukan, tglPilihan, totalBersih, platform, jamTeks]);

    var sheetAK = ss.getSheetByName("Arus Kas") || pastikanSheetAda(ss, "Arus Kas", ["Tanggal", "Keterangan", "Tipe", "Nominal", "Jam"]);
    if (totalBersih >= 0) {
        sheetAK.appendRow([tglPilihan, "[Pemasukan] " + namaTransaksi, "Kas Masuk", totalBersih, jamTeks]);
    } else {
        sheetAK.appendRow([tglPilihan, "[Potongan Admin POS] " + namaTransaksi, "Kas Keluar", Math.abs(totalBersih), jamTeks]);
    }

    var sheetDetail = ss.getSheetByName("Detail_Penjualan") || pastikanSheetAda(ss, "Detail_Penjualan", [
      "Tanggal", "Jam", "Nama Pembeli", "Produk", "Pendapatan Kotor", "Admin", "Pendapatan Bersih", "Total HPP", "Margin (%)"
    ]);

    var marginPersen = 0;
    if (totalBersih > 0) {
      marginPersen = ((totalBersih - totalHppTransaksi) / totalBersih) * 100;
    } else if (totalBersih < 0) {
      marginPersen = ((totalBersih - totalHppTransaksi) / Math.abs(totalBersih)) * 100;
    }
    
    sheetDetail.appendRow([
      tglPilihan, jamTeks, pembeli, ringkasanBarang.join("\n"), subtotalBruto, potonganAdmin, totalBersih, totalHppTransaksi, Math.round(marginPersen)
    ]);

    catatAuditLog("KASIR POS", "Penjualan & Gudang", "-", namaTransaksi + " (Net: Rp " + totalBersih.toLocaleString('id-ID') + " pada " + tglPilihan + ")", userAktif);

    SpreadsheetApp.flush();

    return JSON.stringify({
      success: true,
      stokData: JSON.parse(ambilDataStokLengkap()),
      bukuData: ambilSemuaData(false)
    });

  } catch (e) {
    Logger.log("Error prosesCheckoutPOS: " + e.message);
    // Mengembalikan pesan error langsung ke UI
    return JSON.stringify({ success: false, message: e.message });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Fungsi testing mandiri untuk menguji prosesCheckoutPOS langsung dari Editor Apps Script
 * (Termasuk simulasi perubahan harga, tanggal pesanan, dan potongan admin yang menyebabkan minus)
 */
function test_prosesCheckoutPOS() {
  try {
    var payloadUji = {
      keranjang: [
        { id: "SKU-UJI-01", nama: "Produk Simulasi POS", harga: 12000, tipe: "Satuan", qty: 2 }
      ],
      pembeli: "Pelanggan Uji POS",
      platform: "Shopee",
      potonganAdmin: 26000, // Simulasi potongan lebih besar sehingga hasil Net minus (-2.000)
      subtotalBruto: 24000,
      totalBersihCair: -2000,
      tanggal: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
      user: "Owner"
    };

    var hasilJson = prosesCheckoutPOS(payloadUji);
    var parsed = JSON.parse(hasilJson);
    Logger.log("Status Eksekusi test_prosesCheckoutPOS: " + parsed.success);
  } catch (e) {
    Logger.log(e.message);
  }
}

function hapusTransaksiPOSBerantai(rowIndex, tgl, jam, pembeli) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var adaYangDihapus = false;

    // ====================================================================
    // 1. KEMBALIKAN STOK FISIK & HAPUS RIWAYAT MUTASI
    // ====================================================================
    var sheetMutasi = ss.getSheetByName("Mutasi_Stok");
    var itemDikembalikan = {}; 

    if (sheetMutasi && sheetMutasi.getLastRow() > 1) {
      var dataMutasi = sheetMutasi.getDataRange().getValues();
      
      for (var m = dataMutasi.length - 1; m >= 1; m--) {
        var mTanggalRaw = dataMutasi[m][0];
        var mTanggal = (mTanggalRaw instanceof Date) ? Utilities.formatDate(mTanggalRaw, "GMT+7", "yyyy-MM-dd") : String(mTanggalRaw).split('T')[0];
        var mJam = String(dataMutasi[m][1]).replace(/'/g, ''); 
        var mTipe = String(dataMutasi[m][2]);
        var mNamaBarang = String(dataMutasi[m][3]);
        var mQty = Number(dataMutasi[m][4]) || 0;
        var mKet = String(dataMutasi[m][5]);

        if (mTanggal === tgl && mJam === jam && mKet.indexOf(pembeli) !== -1) {
          if (mTipe === "Penjualan Satuan" || mTipe === "Bahan Bundling Terpakai") {
             if (!itemDikembalikan[mNamaBarang]) itemDikembalikan[mNamaBarang] = 0;
             itemDikembalikan[mNamaBarang] += mQty;
          }
          sheetMutasi.deleteRow(m + 1);
          adaYangDihapus = true;
        }
      }
    }

    if (Object.keys(itemDikembalikan).length > 0) {
      var sheetStok = ss.getSheetByName("Master_Stok");
      if (sheetStok && sheetStok.getLastRow() > 1) {
        var dataStok = sheetStok.getDataRange().getValues();
        var stokBerubah = false;

        for (var s = 1; s < dataStok.length; s++) {
          var namaDiStok = String(dataStok[s][1]).trim();
          if (itemDikembalikan[namaDiStok]) {
             dataStok[s][3] = (Number(dataStok[s][3]) || 0) + itemDikembalikan[namaDiStok];
             stokBerubah = true;
          }
        }
        if (stokBerubah) {
           sheetStok.getRange(1, 1, dataStok.length, dataStok[0].length).setValues(dataStok);
        }
      }
    }

    // ====================================================================
    // 2. HAPUS DARI DETAIL PENJUALAN (Pencarian Presisi, Bukan Urutan Index)
    // ====================================================================
    var sheetDetail = ss.getSheetByName("Detail_Penjualan");
    if (sheetDetail && sheetDetail.getLastRow() > 1) {
      var dataDetail = sheetDetail.getDataRange().getValues();
      for (var d = dataDetail.length - 1; d >= 1; d--) {
        var dTanggalRaw = dataDetail[d][0];
        var dTanggal = (dTanggalRaw instanceof Date) ? Utilities.formatDate(dTanggalRaw, "GMT+7", "yyyy-MM-dd") : String(dTanggalRaw).split('T')[0];
        var dJam = String(dataDetail[d][1]).replace(/'/g, ''); 
        var dPembeli = String(dataDetail[d][2]);
        
        if (dTanggal === tgl && dJam === jam && dPembeli.indexOf(pembeli) !== -1) {
          sheetDetail.deleteRow(d + 1);
          adaYangDihapus = true;
          break; // Hentikan setelah nota tepat ditemukan
        }
      }
    }

    // ====================================================================
    // 3. HAPUS DARI PEMASUKAN
    // ====================================================================
    var sheetPem = ss.getSheetByName("Pemasukan");
    if (sheetPem && sheetPem.getLastRow() > 1) {
      var dataPem = sheetPem.getDataRange().getValues();
      for (var i = dataPem.length - 1; i >= 1; i--) {
        var rowTglRaw = dataPem[i][1];
        var rowTgl = (rowTglRaw instanceof Date) ? Utilities.formatDate(rowTglRaw, "GMT+7", "yyyy-MM-dd") : String(rowTglRaw).split('T')[0];
        var rowJam = String(dataPem[i][4]).replace(/'/g, ''); 
        var rowKet = String(dataPem[i][0]);
        
        if (rowTgl === tgl && rowJam === jam && rowKet.indexOf(pembeli) !== -1) {
          sheetPem.deleteRow(i + 1);
          adaYangDihapus = true;
          break; 
        }
      }
    }

    // ====================================================================
    // 4. HAPUS DARI ARUS KAS
    // ====================================================================
    var sheetAK = ss.getSheetByName("Arus Kas");
    if (sheetAK && sheetAK.getLastRow() > 1) {
      var dataAK = sheetAK.getDataRange().getValues();
      for (var j = dataAK.length - 1; j >= 1; j--) {
        var akTglRaw = dataAK[j][0];
        var akTgl = (akTglRaw instanceof Date) ? Utilities.formatDate(akTglRaw, "GMT+7", "yyyy-MM-dd") : String(akTglRaw).split('T')[0];
        var akJam = String(dataAK[j][4]).replace(/'/g, ''); 
        var akKet = String(dataAK[j][1]);
        
        if (akTgl === tgl && akJam === jam && akKet.indexOf(pembeli) !== -1) {
          sheetAK.deleteRow(j + 1);
          adaYangDihapus = true;
          break;
        }
      }
    }

    if (!adaYangDihapus) {
      throw new Error("Data transaksi tidak ditemukan di database.");
    }

    catatAuditLog("KASIR POS", "Hapus & Restore Stok", "-", "Membatalkan & memulihkan stok transaksi a.n " + pembeli + " (" + tgl + " " + jam + ")", "Admin");

    SpreadsheetApp.flush();

    // Pastikan mengembalikan stokData yang sudah di-restore juga
    return JSON.stringify({
      success: true,
      stokData: JSON.parse(ambilDataStokLengkap()),
      bukuData: ambilSemuaData(false)
    });

  } catch (e) {
    Logger.log("Error hapusTransaksiPOSBerantai: " + e.message);
    return JSON.stringify({ success: false, message: e.message });
  } finally {
    lock.releaseLock();
  }
}

function test_hapusTransaksiPOSBerantai() {
  try {
    var hasil = hapusTransaksiPOSBerantai(0, "2026-10-19", "12:00", "Testing Pembeli");
    Logger.log(hasil);
  } catch(e) {
    Logger.log(e.message);
  }
}


