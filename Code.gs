/**
 * ==============================================================================
 * JURNAL HARIAN SIKAWAN - SDN BABELAN KOTA 01
 * Google Apps Script Backend (Code.gs)
 * ==============================================================================
 * File ini dipasang pada Apps Script (Code.gs) di Google Spreadsheet.
 * Didesain khusus agar tanggal dan hari stabil (imun terhadap pergeseran zona waktu)
 * serta semua interaksi dan tombol di index.html berjalan aman tanpa terkunci.
 */

/**
 * Melayani halaman web HTML utama
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('JURNAL HARIAN SIKAWAN - SDN BABELAN KOTA 01')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Helper: Mengonversi nilai tanggal apa pun (Date object atau string) 
 * secara stabil menjadi format 'YYYY-MM-DD' berbasis zona waktu Asia/Jakarta
 */
function formatTanggalYMD(val) {
  if (!val) return "";
  if (val instanceof Date) {
    var tz = Session.getScriptTimeZone() || "Asia/Jakarta";
    return Utilities.formatDate(val, tz, "yyyy-MM-dd");
  }
  var str = val.toString().trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    return str;
  }
  if (str.indexOf("T") !== -1) {
    var part = str.split("T")[0];
    if (/^\d{4}-\d{2}-\d{2}$/.test(part)) return part;
  }
  var d = new Date(str);
  if (!isNaN(d.getTime())) {
    var tz = Session.getScriptTimeZone() || "Asia/Jakarta";
    return Utilities.formatDate(d, tz, "yyyy-MM-dd");
  }
  return str;
}

/**
 * Helper: Memformat 'YYYY-MM-DD' menjadi Hari, Tanggal Bulan Tahun Indonesia
 */
function formatHariTanggalIndoGAS(ymdStr) {
  try {
    if (!ymdStr) return "";
    var parts = ymdStr.split('-');
    if (parts.length === 3) {
      var y = parseInt(parts[0], 10);
      var m = parseInt(parts[1], 10) - 1;
      var d = parseInt(parts[2], 10);
      var dt = new Date(y, m, d, 12, 0, 0); // Jam 12 siang lokal, imun pergeseran hari
      var days = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
      var months = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
      return days[dt.getDay()] + ", " + dt.getDate() + " " + months[dt.getMonth()] + " " + dt.getFullYear();
    }
  } catch(e) {}
  return ymdStr;
}

/**
 * Mengambil data pegawai dari sheet "Data Pegawai"
 */
function getDataPegawaiFromSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName("Data Pegawai");
    
    if (!sheet) {
      sheet = ss.insertSheet("Data Pegawai");
      sheet.appendRow(["No", "NIP / NI PPPK", "Nama Pegawai", "Jabatan", "Pangkat / Gol", "Status", "Foto Pegawai", "TTD Pegawai"]);
      sheet.getRange(1, 2, 100, 1).setNumberFormat("@"); // Format kolom NIP sebagai teks murni
      sheet.appendRow(["1", "198911102025211006", "ABD. QODIR, S.Pd.I", "Guru PAI", "IX", "PPPK", "", ""]);
      sheet.appendRow(["2", "198806122025211014", "AHMAD SHOHEH, S.Pd", "Guru Kelas", "IX", "PPPK", "", ""]);
      sheet.appendRow(["3", "198911062025211012", "DIDI MULYADI, S.Pd", "Guru Kelas", "IX", "PPPK", "", ""]);
      sheet.appendRow(["4", "197612132008011003", "DODY ARDIANSYAH, S.Pd", "Guru Kelas", "Penata Muda, III/a", "ASN", "", ""]);
      sheet.appendRow(["5", "199411162020122006", "DOLA SARLITA, S.Pd", "Guru Kelas", "Penata Muda, III/a", "ASN", "", ""]);
      sheet.appendRow(["6", "198004032025211031", "EDI SAPUTRA, S.Sos.I", "Guru PJOK", "IX", "PPPK", "", ""]);
      sheet.appendRow(["7", "197605262025212004", "EMI SUHAEMI, S.S.", "Guru Kelas", "IX", "PPPK", "", ""]);
      sheet.appendRow(["8", "199301222025212005", "ERYASANI RAHMAWATI, S.Pd", "Guru Kelas", "IX", "PPPK", "", ""]);
      sheet.appendRow(["9", "198607262009022001", "ESIN RIAWATI, S.Pd", "Guru Kelas", "Penata Muda Tk. I, III/b", "ASN", "", ""]);
      sheet.appendRow(["10", "197305232008012002", "FATKHURIYAH, S.Pd", "Guru Kelas", "Penata Muda Tk. I, III/b", "ASN", "", ""]);
      sheet.appendRow(["11", "197808182008011004", "H. RAHMAT HIDAYATTULLOH, S.Pd", "Guru Kelas", "Penata Tk. I, III/d", "ASN", "", ""]);
      sheet.appendRow(["12", "197606122008012003", "Hj. AISYAH, S.Ag.,M.M", "Guru Kelas", "Penata Tk. I, III/d", "ASN", "", ""]);
      sheet.appendRow(["13", "198802222025212005", "KIKI FUJI LESTARI, S.Pd", "Guru Kelas", "IX", "PPPK", "", ""]);
      sheet.appendRow(["14", "197808202008012005", "LAILATUL FAJRIAH, S.Pd.SD", "Kepala Sekolah", "Penata Tk. I, III/d", "ASN", "", ""]);
      sheet.appendRow(["15", "199503092025212087", "LATIPAH, S.Pd", "Guru Kelas", "IX", "PPPK PW", "", ""]);
      sheet.appendRow(["16", "198506242025211006", "MAHFUDZ, S.Pd.I", "Guru PAI", "IX", "PPPK", "", ""]);
      sheet.appendRow(["17", "199106182025212013", "MARLINAH, S.Pd", "Guru Kelas", "IX", "PPPK", "", ""]);
      sheet.appendRow(["18", "198801282025211015", "MUHAMMAD RIDUAN, S.Pd.I", "Guru PAI", "IX", "PPPK", "", ""]);
      sheet.appendRow(["19", "197212172014082001", "NGATIMAH, S.Pd", "Guru Kelas VI A", "Penata Muda Tk. I, III/b", "ASN", "", ""]);
      sheet.appendRow(["20", "198808022024212007", "NISA UTAMININGRUM, S.Pd", "Guru Kelas", "IX", "PPPK", "", ""]);
      sheet.appendRow(["21", "198509062025211010", "NURDIANSYAH, S.Pd", "Guru Kelas", "IX", "PPPK", "", ""]);
      sheet.appendRow(["22", "197911042025212022", "NURYANIH, S.Pd", "Guru Kelas / Guru PAI", "IX", "PPPK PW", "", ""]);
      sheet.appendRow(["23", "199201072025212010", "PUTRI ALFIANI, S.Pd", "Guru Kelas", "IX", "PPPK", "", ""]);
      sheet.appendRow(["24", "199509072024211007", "RAFLY BAHRI, S.Pd", "Guru PJOK", "IX", "PPPK", "", ""]);
      sheet.appendRow(["25", "198104092008012003", "ROCHIZA EFFENDI, S.Pd", "Guru Kelas I C", "Penata, III/c", "ASN", "", ""]);
      sheet.appendRow(["26", "198411012008012006", "SOPIATI, S.Pd", "Guru Kelas", "Penata, III/c", "ASN", "", ""]);
      sheet.appendRow(["27", "198201272025212007", "SRI SUPRAPTI, S.Pd", "Guru Kelas", "IX", "PPPK", "", ""]);
      sheet.appendRow(["28", "198107252025212007", "SUMARTINI, S.Ag", "Guru PAB", "IX", "PPPK", "", ""]);
      sheet.appendRow(["29", "198111152008012004", "SURYATUN, S.Pd", "Guru Kelas", "Penata, III/c", "ASN", "", ""]);
      sheet.appendRow(["30", "198511092025211006", "SYAFARUDIN AHMED, S.Pd", "Guru Kelas", "IX", "PPPK", "", ""]);
      sheet.appendRow(["31", "196707172008012005", "TITIN, S.Pd", "Guru Kelas", "Penata Muda Tk. I, III/b", "ASN", "", ""]);
      sheet.appendRow(["32", "197801312008012006", "TRI ASTUTY, S.Pd.,M.M", "Guru Kelas I A", "Penata Tk. I, III/d", "ASN", "", ""]);
      sheet.appendRow(["33", "199404172025212015", "WAHYUNING HARDIYANTI, S.Pd", "Guru Kelas", "IX", "PPPK", "", ""]);
      sheet.appendRow(["34", "197604122007012012", "YAYAH SUTINAH, S.Pd", "Guru Kelas I B", "Penata, III/c", "ASN", "", ""]);
      sheet.appendRow(["35", "197507292008012008", "YULI HERLIYANTI, S.Pd", "Guru Kelas", "Penata, III/c", "ASN", "", ""]);
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    // Hapus SAMSUDIN jika sebelumnya pernah masuk ke sheet Data Pegawai
    for (var r = data.length - 1; r >= 1; r--) {
      var rNip = data[r][1] ? data[r][1].toString().trim() : "";
      var rNama = data[r][2] ? data[r][2].toString().trim().toUpperCase() : "";
      if (rNip === "198105102025211008" || rNama.indexOf("SAMSUDIN") !== -1) {
        sheet.deleteRow(r + 1);
        data = sheet.getDataRange().getValues();
      }
    }

    // Pastikan seluruh pegawai baku SDN Babelan Kota 01 selalu tersedia di sheet secara alfabetis
    var defaultPegawaiList = [
      ["198911102025211006", "ABD. QODIR, S.Pd.I", "Guru PAI", "IX", "PPPK"],
      ["198806122025211014", "AHMAD SHOHEH, S.Pd", "Guru Kelas", "IX", "PPPK"],
      ["198911062025211012", "DIDI MULYADI, S.Pd", "Guru Kelas", "IX", "PPPK"],
      ["197612132008011003", "DODY ARDIANSYAH, S.Pd", "Guru Kelas", "Penata Muda, III/a", "ASN"],
      ["199411162020122006", "DOLA SARLITA, S.Pd", "Guru Kelas", "Penata Muda, III/a", "ASN"],
      ["198004032025211031", "EDI SAPUTRA, S.Sos.I", "Guru PJOK", "IX", "PPPK"],
      ["197605262025212004", "EMI SUHAEMI, S.S.", "Guru Kelas", "IX", "PPPK"],
      ["199301222025212005", "ERYASANI RAHMAWATI, S.Pd", "Guru Kelas", "IX", "PPPK"],
      ["198607262009022001", "ESIN RIAWATI, S.Pd", "Guru Kelas", "Penata Muda Tk. I, III/b", "ASN"],
      ["197305232008012002", "FATKHURIYAH, S.Pd", "Guru Kelas", "Penata Muda Tk. I, III/b", "ASN"],
      ["197808182008011004", "H. RAHMAT HIDAYATTULLOH, S.Pd", "Guru Kelas", "Penata Tk. I, III/d", "ASN"],
      ["197606122008012003", "Hj. AISYAH, S.Ag.,M.M", "Guru Kelas", "Penata Tk. I, III/d", "ASN"],
      ["198802222025212005", "KIKI FUJI LESTARI, S.Pd", "Guru Kelas", "IX", "PPPK"],
      ["197808202008012005", "LAILATUL FAJRIAH, S.Pd.SD", "Kepala Sekolah", "Penata Tk. I, III/d", "ASN"],
      ["199503092025212087", "LATIPAH, S.Pd", "Guru Kelas", "IX", "PPPK PW"],
      ["198506242025211006", "MAHFUDZ, S.Pd.I", "Guru PAI", "IX", "PPPK"],
      ["199106182025212013", "MARLINAH, S.Pd", "Guru Kelas", "IX", "PPPK"],
      ["198801282025211015", "MUHAMMAD RIDUAN, S.Pd.I", "Guru PAI", "IX", "PPPK"],
      ["197212172014082001", "NGATIMAH, S.Pd", "Guru Kelas VI A", "Penata Muda Tk. I, III/b", "ASN"],
      ["198808022024212007", "NISA UTAMININGRUM, S.Pd", "Guru Kelas", "IX", "PPPK"],
      ["198509062025211010", "NURDIANSYAH, S.Pd", "Guru Kelas", "IX", "PPPK"],
      ["197911042025212022", "NURYANIH, S.Pd", "Guru Kelas / Guru PAI", "IX", "PPPK PW"],
      ["199201072025212010", "PUTRI ALFIANI, S.Pd", "Guru Kelas", "IX", "PPPK"],
      ["199509072024211007", "RAFLY BAHRI, S.Pd", "Guru PJOK", "IX", "PPPK"],
      ["198104092008012003", "ROCHIZA EFFENDI, S.Pd", "Guru Kelas I C", "Penata, III/c", "ASN"],
      ["198411012008012006", "SOPIATI, S.Pd", "Guru Kelas", "Penata, III/c", "ASN"],
      ["198201272025212007", "SRI SUPRAPTI, S.Pd", "Guru Kelas", "IX", "PPPK"],
      ["198107252025212007", "SUMARTINI, S.Ag", "Guru PAB", "IX", "PPPK"],
      ["198111152008012004", "SURYATUN, S.Pd", "Guru Kelas", "Penata, III/c", "ASN"],
      ["198511092025211006", "SYAFARUDIN AHMED, S.Pd", "Guru Kelas", "IX", "PPPK"],
      ["196707172008012005", "TITIN, S.Pd", "Guru Kelas", "Penata Muda Tk. I, III/b", "ASN"],
      ["197801312008012006", "TRI ASTUTY, S.Pd.,M.M", "Guru Kelas I A", "Penata Tk. I, III/d", "ASN"],
      ["199404172025212015", "WAHYUNING HARDIYANTI, S.Pd", "Guru Kelas", "IX", "PPPK"],
      ["197604122007012012", "YAYAH SUTINAH, S.Pd", "Guru Kelas I B", "Penata, III/c", "ASN"],
      ["197507292008012008", "YULI HERLIYANTI, S.Pd", "Guru Kelas", "Penata, III/c", "ASN"]
    ];

    var needReload = false;
    for (var dIdx = 0; dIdx < defaultPegawaiList.length; dIdx++) {
      var def = defaultPegawaiList[dIdx];
      var exists = false;
      for (var k = 1; k < data.length; k++) {
        var n = data[k][1] ? data[k][1].toString().trim() : "";
        var nm = data[k][2] ? data[k][2].toString().trim() : "";
        if (n === def[0] || (nm !== "" && nm.toUpperCase().indexOf(def[1].toUpperCase()) !== -1)) {
          exists = true;
          break;
        }
      }
      if (!exists) {
        var nextNo = sheet.getLastRow();
        sheet.appendRow([
          nextNo.toString(),
          def[0],
          def[1],
          def[2],
          def[3],
          def[4],
          "",
          ""
        ]);
        needReload = true;
      }
    }
    if (needReload) {
      data = sheet.getDataRange().getValues();
    }

    var pegawaiList = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var nip = row[1] ? row[1].toString().trim() : "";
      var nama = row[2] ? row[2].toString().trim() : "";
      
      // Lewatkan jika SAMSUDIN
      if (nip === "198105102025211008" || nama.toUpperCase().indexOf("SAMSUDIN") !== -1) {
        continue;
      }

      if (nama !== "" || nip !== "") {
        pegawaiList.push({
          no: (pegawaiList.length + 1).toString(),
          nip: nip,
          name: nama,
          jabatan: row[3] ? row[3].toString().trim() : "-",
          pangkat: row[4] ? row[4].toString().trim() : "-",
          status: row[5] ? row[5].toString().trim() : "-",
          poto: row[6] ? row[6].toString().trim() : "",
          ttd: row[7] ? row[7].toString().trim() : ""
        });
      }
    }

    // Urutkan pegawai secara alfabetis A - Z berdasarkan Nama Pegawai
    pegawaiList.sort(function(a, b) {
      var nameA = (a.name || "").toString().toUpperCase();
      var nameB = (b.name || "").toString().toUpperCase();
      return nameA.localeCompare(nameB);
    });

    for (var p = 0; p < pegawaiList.length; p++) {
      pegawaiList[p].no = (p + 1).toString();
    }

    return pegawaiList;
  } catch (error) {
    Logger.log("Error getDataPegawaiFromSheet: " + error.toString());
    return [];
  }
}

/**
 * Menyimpan data pegawai baru atau update jika NIP sudah ada
 */
function simpanDataPegawaiBaru(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { status: "error", message: "Spreadsheet tidak aktif" };
    var sheet = ss.getSheetByName("Data Pegawai");
    
    if (!sheet) {
      sheet = ss.insertSheet("Data Pegawai");
      sheet.appendRow(["No", "NIP / NI PPPK", "Nama Pegawai", "Jabatan", "Pangkat / Gol", "Status", "Foto Pegawai", "TTD Pegawai"]);
    }
    sheet.getRange(1, 2, Math.max(10, sheet.getLastRow() + 5), 1).setNumberFormat("@");
    
    var data = sheet.getDataRange().getValues();
    var existingRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (payload.nip && data[i][1] && data[i][1].toString().trim() === payload.nip.toString().trim()) {
        existingRow = i + 1;
        break;
      }
    }

    if (existingRow > 0) {
      sheet.getRange(existingRow, 2, 1, 7).setValues([[
        payload.nip || "",
        payload.name || "",
        payload.jabatan || "",
        payload.pangkat || "",
        payload.status || "",
        payload.poto || "",
        payload.ttd || ""
      ]]);
      return { status: "success", message: "Data Pegawai berhasil diperbarui di Spreadsheet!" };
    } else {
      var nextRow = sheet.getLastRow() + 1;
      var nomorUrut = nextRow - 1;

      sheet.appendRow([
        nomorUrut,
        payload.nip || "",
        payload.name || "",
        payload.jabatan || "",
        payload.pangkat || "",
        payload.status || "",
        payload.poto || "",
        payload.ttd || ""
      ]);
      return { status: "success", message: "Data Pegawai berhasil disimpan ke Spreadsheet!" };
    }
  } catch (error) {
    return { status: "error", message: "Gagal menyimpan data pegawai: " + error.toString() };
  }
}

/**
 * Menghapus data pegawai dari sheet "Data Pegawai" berdasarkan NIP
 */
function hapusPegawaiFromSheet(nip) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { status: "success" };
    var sheet = ss.getSheetByName("Data Pegawai");
    if (!sheet) return { status: "success" };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === nip.toString().trim()) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return { status: "success", message: "Data Pegawai berhasil dihapus dari Spreadsheet!" };
  } catch (error) {
    return { status: "error", message: "Gagal menghapus data pegawai: " + error.toString() };
  }
}

/**
 * Menyimpan Data Sekolah secara permanen ke Google Sheets
 */
function simpanDataSekolah(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { status: "error", message: "Spreadsheet tidak aktif" };
    var sheet = ss.getSheetByName("Data Sekolah");
    if (!sheet) {
      sheet = ss.insertSheet("Data Sekolah");
    }
    
    sheet.clear();
    sheet.appendRow(["Kunci", "Nilai"]);
    
    sheet.appendRow(["namaKepala", payload.namaKepala || ""]);
    sheet.appendRow(["nipKepala", payload.nipKepala || ""]);
    sheet.appendRow(["ttdKepala", payload.ttdKepala || ""]);
    sheet.appendRow(["stempelSekolah", payload.stempelSekolah || ""]);
    sheet.appendRow(["kopSekolah", payload.kopSekolah || ""]);
    
    return { status: "success", message: "Pengaturan Data Sekolah berhasil disimpan secara permanen!" };
  } catch (error) {
    return { status: "error", message: "Gagal menyimpan Data Sekolah: " + error.toString() };
  }
}

/**
 * Mengambil Data Sekolah dari Google Sheets
 */
function getDataSekolahFromSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var defaultData = {
      namaKepala: "LAILATUL FAJRIAH, S.Pd.SD.",
      nipKepala: "197808202008012005",
      ttdKepala: "",
      stempelSekolah: "",
      kopSekolah: ""
    };
    if (!ss) return defaultData;

    var sheet = ss.getSheetByName("Data Sekolah");
    if (!sheet) {
      sheet = ss.insertSheet("Data Sekolah");
      sheet.appendRow(["Kunci", "Nilai"]);
      sheet.appendRow(["namaKepala", defaultData.namaKepala]);
      sheet.appendRow(["nipKepala", defaultData.nipKepala]);
      sheet.appendRow(["ttdKepala", ""]);
      sheet.appendRow(["stempelSekolah", ""]);
      sheet.appendRow(["kopSekolah", ""]);
      return defaultData;
    }
    
    var data = sheet.getDataRange().getValues();
    var result = {
      namaKepala: defaultData.namaKepala,
      nipKepala: defaultData.nipKepala,
      ttdKepala: "",
      stempelSekolah: "",
      kopSekolah: ""
    };
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var k = data[i][0].toString().trim();
        var v = data[i][1] ? data[i][1].toString().trim() : "";
        if (v !== "") {
          result[k] = v;
        }
      }
    }
    return result;
  } catch (error) {
    Logger.log("Error getDataSekolahFromSheet: " + error.toString());
    return {
      namaKepala: "LAILATUL FAJRIAH, S.Pd.SD.",
      nipKepala: "197808202008012005",
      ttdKepala: "",
      stempelSekolah: "",
      kopSekolah: ""
    };
  }
}

/**
 * Mengambil riwayat pengisian jurnal harian dari sheet "Riwayat Harian"
 * Didesain aman: format tanggal dikonversi ke YYYY-MM-DD teks agar tidak bergeser hari
 */
function getDataRiwayatFromSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName("Riwayat Harian");
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var riwayatList = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[4]) continue;

      var nip = row[3] ? row[3].toString().trim() : "";
      var nama = row[4] ? row[4].toString().trim() : "";
      if (nip === "198105102025211008" || nama.toUpperCase().indexOf("SAMSUDIN") !== -1) {
        continue;
      }

      var rawTgl = row[1];
      var tanggalStr = formatTanggalYMD(rawTgl);

      var rawHariTgl = row[2];
      var hariTanggalStr = rawHariTgl ? rawHariTgl.toString().trim() : "";
      if (!hariTanggalStr && tanggalStr) {
        hariTanggalStr = formatHariTanggalIndoGAS(tanggalStr);
      }

      var kegiatanList = [];
      try {
        kegiatanList = row[10] ? JSON.parse(row[10].toString()) : [];
      } catch (e) {
        kegiatanList = [];
      }

      riwayatList.push({
        id: row[0] ? row[0].toString() : "rw_" + i,
        tanggal: tanggalStr,
        hariTanggal: hariTanggalStr,
        nip: nip,
        nama: nama,
        shift: row[5] ? row[5].toString() : "",
        rentangWaktu: row[6] ? row[6].toString() : "",
        status: row[7] ? row[7].toString() : "",
        jabatan: row[8] ? row[8].toString() : "",
        pangkat: row[9] ? row[9].toString() : "",
        kegiatanList: kegiatanList
      });
    }

    // Urutkan riwayat berdasarkan tanggal terbaru (descending)
    riwayatList.sort(function(a, b) {
      var tglA = (a.tanggal || "").toString();
      var tglB = (b.tanggal || "").toString();
      return tglB.localeCompare(tglA);
    });

    return riwayatList;
  } catch (error) {
    Logger.log("Error getDataRiwayatFromSheet: " + error.toString());
    return [];
  }
}

/**
 * Mengosongkan data riwayat jurnal dari sheet "Riwayat Harian"
 */
function kosongkanRiwayatFromSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { status: "success" };
    var sheet = ss.getSheetByName("Riwayat Harian");
    if (!sheet) return { status: "success" };
    
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    return { status: "success", message: "Riwayat berhasil dikosongkan dari Spreadsheet!" };
  } catch (error) {
    return { status: "error", message: "Gagal mengosongkan riwayat: " + error.toString() };
  }
}

/**
 * Menyimpan atau memperbarui data riwayat jurnal harian ke sheet "Riwayat Harian"
 * Format tanggal dikunci sebagai teks murni string 'YYYY-MM-DD'
 */
function simpanDataRiwayatToSheet(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { status: "error", message: "Spreadsheet tidak aktif" };
    var sheet = ss.getSheetByName("Riwayat Harian");
    
    if (!sheet) {
      sheet = ss.insertSheet("Riwayat Harian");
      sheet.appendRow([
        "ID",
        "Tanggal",
        "Hari Tanggal",
        "NIP",
        "Nama Pegawai",
        "Shift",
        "Rentang Waktu",
        "Status",
        "Jabatan",
        "Pangkat",
        "Detail Kegiatan (JSON)"
      ]);
    }

    // Set kolom Tanggal dan NIP sebagai teks murni (@)
    sheet.getRange(1, 2, Math.max(10, sheet.getLastRow() + 5), 1).setNumberFormat("@");
    sheet.getRange(1, 4, Math.max(10, sheet.getLastRow() + 5), 1).setNumberFormat("@");

    var data = sheet.getDataRange().getValues();
    var existingRow = -1;

    var cleanTanggal = formatTanggalYMD(payload.tanggal);
    var cleanHariTanggal = payload.hariTanggal ? payload.hariTanggal.toString().trim() : "";
    if (cleanTanggal && (!cleanHariTanggal || cleanHariTanggal.indexOf("Invalid") !== -1)) {
      cleanHariTanggal = formatHariTanggalIndoGAS(cleanTanggal);
    }

    // Cek apakah data dengan ID atau (NIP + Tanggal) sudah ada
    for (var i = 1; i < data.length; i++) {
      var rowTgl = formatTanggalYMD(data[i][1]);
      if ((payload.id && data[i][0] && data[i][0].toString() === payload.id.toString()) ||
          (data[i][3] && data[i][3].toString() === payload.nip && rowTgl === cleanTanggal)) {
        existingRow = i + 1;
        break;
      }
    }

    var kegiatanJson = JSON.stringify(payload.kegiatanList || []);
    // Cegah limit ukuran sel di Google Spreadsheet (50.000 karakter)
    if (kegiatanJson.length > 45000) {
      var compactList = (payload.kegiatanList || []).map(function(k) {
        return {
          jamMulai: k.jamMulai || "",
          menitMulai: k.menitMulai || "",
          jamSelesai: k.jamSelesai || "",
          menitSelesai: k.menitSelesai || "",
          uraian: k.uraian || "",
          foto: k.foto ? (k.foto.length > 500 ? k.foto.substring(0, 500) : k.foto) : ""
        };
      });
      kegiatanJson = JSON.stringify(compactList);
    }
    var rowData = [
      payload.id || ("rw_" + new Date().getTime()),
      cleanTanggal,
      cleanHariTanggal,
      payload.nip || "",
      payload.nama || "",
      payload.shift || "",
      payload.rentangWaktu || "",
      payload.status || "",
      payload.jabatan || "",
      payload.pangkat || "",
      kegiatanJson
    ];

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { status: "success", message: "Riwayat jurnal berhasil disimpan ke Spreadsheet!" };
  } catch (error) {
    return { status: "error", message: "Gagal menyimpan riwayat jurnal: " + error.toString() };
  }
}

/**
 * Menghapus data riwayat jurnal dari sheet "Riwayat Harian" berdasarkan ID
 */
function hapusDataRiwayatFromSheet(id) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { status: "success" };
    var sheet = ss.getSheetByName("Riwayat Harian");
    if (!sheet) return { status: "success" };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === id.toString()) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return { status: "success", message: "Data riwayat berhasil dihapus dari Spreadsheet!" };
  } catch (error) {
    return { status: "error", message: "Gagal menghapus riwayat: " + error.toString() };
  }
}
