/**
 * ==============================================================================
 * GOOGLE APPS SCRIPT: SISTEM PENYIMPANAN ASESMEN STRUKTUR DATA
 * UKIN INFORMATIKA FASE E / KELAS X — SMA NEGERI 4 YOGYAKARTA
 * Penyusun: Agasta Pratama Nugraha, S.Kom.
 * ==============================================================================
 * Deskripsi:
 * Script ini mengelola penyimpanan data Pre-Test & Post-Test secara otomatis:
 * 1. Sheet "Rekap_Nilai": Menampilkan skor Pre-Test & Post-Test, selisih kenaikan
 *    (Gain Score), status kenaikan (Meningkat / Tetap / Menurun), keterangan KKTP
 *    (TUNTAS >= 75 / REMEDIAL < 75), dan rekomendasi tindak lanjut (Pengayaan / Remedial).
 * 2. Sheet "Detail_PreTest": Log rincian jawaban murid untuk soal 1 s.d. 6.
 * 3. Sheet "Detail_PostTest": Log rincian jawaban murid untuk soal 1 s.d. 6.
 * ==============================================================================
 */

// KUNCI JAWABAN RESMI RPP (6 BUTIR SOAL)
const KUNCI_JAWABAN = {
  1: "C", // Soal 1 (C4): Aplikasi pemutar musik -> Queue (FIFO)
  2: "C", // Soal 2 (C4): Customer service -> Queue (FIFO)
  3: "B", // Soal 3 (C4): Fitur Back peramban browser -> Stack (LIFO)
  4: "B", // Soal 4 (C5): Antrean loket A, B, C -> Enqueue A, Enqueue B, Dequeue A, Dequeue B
  5: "B", // Soal 5 (C5): Tumpukan Stack [1,2,3] Push(4), Pop(), Push(5) -> [1,2,3,5]
  6: "C"  // Soal 6 (C5): Fitur Batalkan Langkah (Undo) -> Stack
};

// Standar Kriteria Ketercapaian Tujuan Pembelajaran (RPP Hal 12 & 56)
const KKTP = 75;
const TOTAL_SOAL = 6;

/**
 * Handle GET Request (Untuk Uji Coba Endpoint di Browser)
 */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "status";
  
  if (action === "setup") {
    setupSheets();
    return createJsonResponse({
      status: "success",
      message: "Tabel Google Spreadsheet berhasil diinisialisasi!"
    });
  }
  
  return createJsonResponse({
    status: "online",
    message: "Endpoint Web App Google Apps Script Asesmen Struktur Data aktif.",
    timestamp: new Date().toISOString()
  });
}

/**
 * Handle POST Request (Menerima Pengiriman Data dari Website)
 */
function doPost(e) {
  try {
    var payload;
    
    // Parsing payload baik JSON maupun URL Encoded Form
    if (e.postData && e.postData.contents) {
      if (e.postData.type.indexOf("application/json") !== -1) {
        payload = JSON.parse(e.postData.contents);
      } else {
        payload = e.parameter;
        if (typeof payload.jawaban === "string") {
          try {
            payload.jawaban = JSON.parse(payload.jawaban);
          } catch (err) {
            payload.jawaban = {};
          }
        }
      }
    } else if (e.parameter) {
      payload = e.parameter;
    } else {
      throw new Error("Tidak ada data yang diterima.");
    }

    // Ambil data murid
    var nama = (payload.nama || "Siswa Anonim").trim();
    var presensi = (payload.presensi !== undefined && payload.presensi !== null) ? String(payload.presensi).trim() : "-";
    var kelas = (payload.kelas || "E / X5").trim();
    var tipe = (payload.tipe || "pre").toLowerCase(); // "pre" atau "pasca"
    var jawaban = payload.jawaban || {};

    // Hitung skor berdasarkan kunci jawaban resmi
    var benarCount = 0;
    for (var i = 1; i <= TOTAL_SOAL; i++) {
      var ansUser = jawaban[i] || jawaban[String(i)] || "";
      if (String(ansUser).trim().toUpperCase() === KUNCI_JAWABAN[i]) {
        benarCount++;
      }
    }
    var calculatedScore = Math.round((benarCount / TOTAL_SOAL) * 100);

    // Buka spreadsheet aktif
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    setupSheets(ss); // Pastikan tabel sudah terstruktur

    var now = new Date();
    var timestampStr = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");

    // 1. Simpan ke Detail Log
    simpanDetailLog(ss, tipe, {
      timestamp: timestampStr,
      nama: nama,
      presensi: presensi,
      kelas: kelas,
      jawaban: jawaban,
      benar: benarCount,
      skor: calculatedScore
    });

    // 2. Simpan & Perbarui ke Rekap Nilai (Komparasi & Kenaikan)
    var resultRekap = perbaruiRekapNilai(ss, {
      timestamp: timestampStr,
      nama: nama,
      presensi: presensi,
      kelas: kelas,
      tipe: tipe,
      skor: calculatedScore
    });

    return createJsonResponse({
      status: "success",
      message: "Data " + (tipe === "pasca" ? "Post-Test" : "Pre-Test") + " berhasil disimpan!",
      nama: nama,
      presensi: presensi,
      tipe: tipe,
      benar: benarCount,
      skor: calculatedScore,
      rekap: resultRekap
    });

  } catch (error) {
    return createJsonResponse({
      status: "error",
      message: error.toString()
    });
  }
}

/**
 * Simpan rincian jawaban ke sheet log detail
 */
function simpanDetailLog(ss, tipe, data) {
  var sheetName = (tipe === "pasca") ? "Detail_PostTest" : "Detail_PreTest";
  var sheet = ss.getSheetByName(sheetName);
  
  var row = [
    data.timestamp,
    data.nama,
    data.presensi,
    data.kelas,
    data.jawaban[1] || data.jawaban["1"] || "-",
    data.jawaban[2] || data.jawaban["2"] || "-",
    data.jawaban[3] || data.jawaban["3"] || "-",
    data.jawaban[4] || data.jawaban["4"] || "-",
    data.jawaban[5] || data.jawaban["5"] || "-",
    data.jawaban[6] || data.jawaban["6"] || "-",
    data.benar,
    data.skor
  ];
  
  sheet.appendRow(row);
  
  // Format center alignment untuk presensi, kelas, jawaban, skor
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 1).setHorizontalAlignment("center");
  sheet.getRange(lastRow, 3, 1, 10).setHorizontalAlignment("center");
}

/**
 * Perbarui atau Tambahkan Siswa ke Sheet Rekap Nilai
 */
function perbaruiRekapNilai(ss, data) {
  var sheet = ss.getSheetByName("Rekap_Nilai");
  var values = sheet.getDataRange().getValues();
  
  var targetRow = -1;
  // Cari berdasarkan Nomor Presensi (kolom D / index 3) dan kecocokan Nama (kolom C / index 2)
  for (var i = 1; i < values.length; i++) {
    var rowPresensi = String(values[i][3]).trim();
    var rowNama = String(values[i][2]).trim().toLowerCase();
    
    if (data.presensi !== "-" && rowPresensi === data.presensi) {
      targetRow = i + 1;
      break;
    } else if (rowNama === data.nama.toLowerCase()) {
      targetRow = i + 1;
      break;
    }
  }

  var preScore = null;
  var postScore = null;

  if (targetRow !== -1) {
    // Siswa sudah terdaftar di rekap, perbarui nilai
    var existingPre = sheet.getRange(targetRow, 6).getValue();
    var existingPost = sheet.getRange(targetRow, 7).getValue();

    if (data.tipe === "pre") {
      preScore = data.skor;
      postScore = (existingPost !== "" && existingPost !== "-") ? Number(existingPost) : null;
      sheet.getRange(targetRow, 6).setValue(preScore);
    } else {
      postScore = data.skor;
      preScore = (existingPre !== "" && existingPre !== "-") ? Number(existingPre) : null;
      sheet.getRange(targetRow, 7).setValue(postScore);
    }

    sheet.getRange(targetRow, 2).setValue(data.timestamp); // Update timestamp
    sheet.getRange(targetRow, 3).setValue(data.nama);      // Update nama
    sheet.getRange(targetRow, 4).setValue(data.presensi);  // Update presensi
    sheet.getRange(targetRow, 5).setValue(data.kelas);     // Update kelas

  } else {
    // Siswa baru, tambahkan baris baru
    targetRow = sheet.getLastRow() + 1;
    var noUrut = targetRow - 1;

    if (data.tipe === "pre") {
      preScore = data.skor;
    } else {
      postScore = data.skor;
    }

    var newRow = [
      noUrut,
      data.timestamp,
      data.nama,
      data.presensi,
      data.kelas,
      (preScore !== null) ? preScore : "-",
      (postScore !== null) ? postScore : "-",
      "-", // Gain
      "-", // Status Kenaikan
      "-", // Keterangan KKTP
      "-"  // Tindak Lanjut
    ];

    sheet.appendRow(newRow);
  }

  // Hitung Kenaikan, Status Kenaikan, Keterangan KKTP, dan Tindak Lanjut
  var gain = "-";
  var statusKenaikan = "Menunggu Tes";
  var keteranganKktp = "-";
  var tindakLanjut = "-";

  if (preScore !== null && postScore !== null) {
    var selisih = postScore - preScore;
    gain = (selisih >= 0 ? "+" : "") + selisih;
    
    if (selisih > 0) {
      statusKenaikan = "Meningkat (+" + selisih + ")";
    } else if (selisih === 0) {
      statusKenaikan = "Tetap (0)";
    } else {
      statusKenaikan = "Menurun (" + selisih + ")";
    }
  } else if (postScore !== null) {
    statusKenaikan = "Post-Test Saja";
  } else if (preScore !== null) {
    statusKenaikan = "Pre-Test Saja";
  }

  // Evaluasi KKTP (berdasarkan Post-Test jika sudah dikerjakan)
  if (postScore !== null) {
    if (postScore >= KKTP) {
      keteranganKktp = "TUNTAS";
      tindakLanjut = "Pengayaan (Sistem E-Commerce / Riwayat Browser)";
    } else {
      keteranganKktp = "REMEDIAL";
      tindakLanjut = "Remedial (Konsep LIFO/FIFO & Latihan Terbimbing)";
    }
  }

  // Update kolom H, I, J, K pada baris target
  sheet.getRange(targetRow, 8).setValue(gain);
  sheet.getRange(targetRow, 9).setValue(statusKenaikan);
  sheet.getRange(targetRow, 10).setValue(keteranganKktp);
  sheet.getRange(targetRow, 11).setValue(tindakLanjut);

  // Styling baris
  sheet.getRange(targetRow, 1).setHorizontalAlignment("center");
  sheet.getRange(targetRow, 4, 1, 7).setHorizontalAlignment("center");
  sheet.getRange(targetRow, 11).setHorizontalAlignment("left");

  // Highlight warna kolom keterangan KKTP
  var cellKktp = sheet.getRange(targetRow, 10);
  if (keteranganKktp === "TUNTAS") {
    cellKktp.setBackground("#dcfce7"); // Hijau lembut
    cellKktp.setFontColor("#15803d");
    cellKktp.setFontWeight("bold");
  } else if (keteranganKktp === "REMEDIAL") {
    cellKktp.setBackground("#fee2e2"); // Merah lembut
    cellKktp.setFontColor("#b91c1c");
    cellKktp.setFontWeight("bold");
  }

  // Highlight warna status kenaikan
  var cellGain = sheet.getRange(targetRow, 9);
  if (statusKenaikan.indexOf("Meningkat") !== -1) {
    cellGain.setBackground("#e0f2fe"); // Biru lembut
    cellGain.setFontColor("#0369a1");
    cellGain.setFontWeight("bold");
  } else if (statusKenaikan.indexOf("Menurun") !== -1) {
    cellGain.setBackground("#fef3c7"); // Kuning oranye
    cellGain.setFontColor("#b45309");
  }

  return {
    preScore: preScore,
    postScore: postScore,
    gain: gain,
    statusKenaikan: statusKenaikan,
    keteranganKktp: keteranganKktp,
    tindakLanjut: tindakLanjut
  };
}

/**
 * Setup struktur tabel otomatis jika belum tersedia
 */
function setupSheets(targetSs) {
  var ss = targetSs || SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Buat atau format sheet Rekap_Nilai
  var sRekap = ss.getSheetByName("Rekap_Nilai");
  if (!sRekap) {
    sRekap = ss.insertSheet("Rekap_Nilai", 0);
  }
  
  if (sRekap.getLastRow() === 0) {
    var headersRekap = [
      "No",
      "Timestamp Terakhir",
      "Nama Lengkap Siswa",
      "No Presensi",
      "Kelas",
      "Nilai Pre-Test",
      "Nilai Post-Test",
      "Gain Score",
      "Status Kenaikan",
      "Keterangan KKTP (≥75)",
      "Rekomendasi Tindak Lanjut"
    ];
    sRekap.appendRow(headersRekap);
    formatHeader(sRekap, "#0f172a", "#f8fafc");
    sRekap.setColumnWidth(1, 45);
    sRekap.setColumnWidth(2, 150);
    sRekap.setColumnWidth(3, 220);
    sRekap.setColumnWidth(4, 90);
    sRekap.setColumnWidth(5, 80);
    sRekap.setColumnWidth(6, 110);
    sRekap.setColumnWidth(7, 115);
    sRekap.setColumnWidth(8, 95);
    sRekap.setColumnWidth(9, 140);
    sRekap.setColumnWidth(10, 150);
    sRekap.setColumnWidth(11, 280);
    sRekap.setFrozenRows(1);
  }

  // 2. Buat atau format sheet Detail_PreTest
  var sPre = ss.getSheetByName("Detail_PreTest");
  if (!sPre) {
    sPre = ss.insertSheet("Detail_PreTest", 1);
  }
  if (sPre.getLastRow() === 0) {
    var headersPre = [
      "Timestamp",
      "Nama Siswa",
      "No Presensi",
      "Kelas",
      "Soal 1 (Queue)",
      "Soal 2 (Queue)",
      "Soal 3 (Stack)",
      "Soal 4 (Queue)",
      "Soal 5 (Stack)",
      "Soal 6 (Stack)",
      "Jumlah Benar",
      "Skor Pre-Test"
    ];
    sPre.appendRow(headersPre);
    formatHeader(sPre, "#854d0e", "#fef9c3"); // Nuansa Emas/Coklat
    sPre.setColumnWidth(1, 150);
    sPre.setColumnWidth(2, 220);
    sPre.setColumnWidth(3, 90);
    sPre.setColumnWidth(4, 80);
    for (var c = 5; c <= 10; c++) sPre.setColumnWidth(c, 110);
    sPre.setColumnWidth(11, 100);
    sPre.setColumnWidth(12, 110);
    sPre.setFrozenRows(1);
  }

  // 3. Buat atau format sheet Detail_PostTest
  var sPost = ss.getSheetByName("Detail_PostTest");
  if (!sPost) {
    sPost = ss.insertSheet("Detail_PostTest", 2);
  }
  if (sPost.getLastRow() === 0) {
    var headersPost = [
      "Timestamp",
      "Nama Siswa",
      "No Presensi",
      "Kelas",
      "Soal 1 (Queue)",
      "Soal 2 (Queue)",
      "Soal 3 (Stack)",
      "Soal 4 (Queue)",
      "Soal 5 (Stack)",
      "Soal 6 (Stack)",
      "Jumlah Benar",
      "Skor Post-Test"
    ];
    sPost.appendRow(headersPost);
    formatHeader(sPost, "#1e3a8a", "#dbeafe"); // Nuansa Biru
    sPost.setColumnWidth(1, 150);
    sPost.setColumnWidth(2, 220);
    sPost.setColumnWidth(3, 90);
    sPost.setColumnWidth(4, 80);
    for (var cp = 5; cp <= 10; cp++) sPost.setColumnWidth(cp, 110);
    sPost.setColumnWidth(11, 100);
    sPost.setColumnWidth(12, 110);
    sPost.setFrozenRows(1);
  }
}

/**
 * Utilitas format baris header
 */
function formatHeader(sheet, bgColor, textColor) {
  var range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  range.setBackground(bgColor);
  range.setFontColor(textColor);
  range.setFontWeight("bold");
  range.setHorizontalAlignment("center");
  range.setVerticalAlignment("middle");
  sheet.setRowHeight(1, 35);
}

/**
 * Helper response JSON dengan CORS Header
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
