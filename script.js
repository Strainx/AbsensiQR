function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // =========================================================
    // FITUR: AMBIL REKAP KESELURUHAN (Master Siswa vs Status Kehadiran)
    // =========================================================
    if (data.action === "getRekapKeseluruhan") {
      var kegiatan = data.kegiatan;
      var hariIni = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
      
      // 1. Ambil Master Siswa
      var sheetSiswa = ss.getSheetByName("MasterSiswa");
      var dataSiswa = sheetSiswa.getDataRange().getValues();
      var listSiswa = [];
      for (var i = 1; i < dataSiswa.length; i++) {
        if (dataSiswa[i][0]) {
          listSiswa.push({
            id: String(dataSiswa[i][0]).trim(),
            nama: dataSiswa[i][1],
            kelas: dataSiswa[i][2]
          });
        }
      }
      
      // 2. Ambil Log Absen hari ini untuk kegiatan tersebut
      var sheetLog = ss.getSheetByName("LogAbsen");
      var dataLog = sheetLog.getLastRow() > 1 ? sheetLog.getDataRange().getValues() : [];
      var petaAbsen = {}; 
      
      // Bersihkan teks kegiatan untuk pencocokan yang aman dari spasi/huruf besar-kecil
      var targetKegiatan = String(kegiatan || "").trim().toLowerCase();
      
      for (var j = 1; j < dataLog.length; j++) {
        var tglLog = String(dataLog[j][1] || "").trim();
        var kegiatanLog = String(dataLog[j][6] || "").trim().toLowerCase();
        
        // Cocokan tanggal dan kegiatan
        if (tglLog === hariIni && kegiatanLog === targetKegiatan) {
          var idSs = String(dataLog[j][3] || "").trim();
          
          // Konversi & bersihkan waktu menggunakan fungsi helper
          var jamTampil = formatJamBersih(dataLog[j][2]);

          petaAbsen[idSs] = {
            status: "Hadir",
            waktu: jamTampil,
            petugas: String(dataLog[j][7] || "Petugas")
          };
        }
      }
      
      // 3. Gabungkan Master dengan Status Kehadiran
      var rekapGabungan = listSiswa.map(function(siswa, index) {
        var absenInfo = petaAbsen[siswa.id];
        return {
          no: index + 1,
          id: siswa.id,
          nama: siswa.nama,
          kelas: siswa.kelas,
          status: absenInfo ? "Hadir" : "Belum Hadir",
          waktu: absenInfo ? absenInfo.waktu : "-",
          petugas: absenInfo ? absenInfo.petugas : "-"
        };
      });
      
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        rekap: rekapGabungan 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // =========================================================
    // PROSES UTAMA: CEK MASTER & SCAN ABSEN
    // =========================================================
    var sheetSiswa = ss.getSheetByName("MasterSiswa");
    var dataSiswa = sheetSiswa.getDataRange().getValues();
    var siswaDitemukan = null;
    var idScan = String(data.idSiswa).trim();
    
    for (var k = 1; k < dataSiswa.length; k++) {
      if (String(dataSiswa[k][0]).trim() === idScan) {
        siswaDitemukan = {
          id: dataSiswa[k][0],
          nama: dataSiswa[k][1],
          kelas: dataSiswa[k][2]
        };
        break;
      }
    }
    
    if (!siswaDitemukan) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        pesan: "Kartu tidak terdaftar / ID: " + idScan 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Cek duplikat di LogAbsen hari ini (Anti-Double Scan Multi-Device)
    var sheetLog = ss.getSheetByName("LogAbsen");
    var dataLog = sheetLog.getLastRow() > 1 ? sheetLog.getDataRange().getValues() : [];
    var hariIniScan = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var scanKegiatanTarget = String(data.kegiatan || "").trim().toLowerCase();
    
    for (var m = 1; m < dataLog.length; m++) {
      var tglLogScan = String(dataLog[m][1] || "").trim();
      var idLogScan = String(dataLog[m][3] || "").trim();
      var kegiatanLogScan = String(dataLog[m][6] || "").trim().toLowerCase();

      if (tglLogScan === hariIniScan && idLogScan === idScan && kegiatanLogScan === scanKegiatanTarget) {
        var petugasTerdahulu = dataLog[m][7] || "Petugas Lain";
        return ContentService.createTextOutput(JSON.stringify({ 
          status: "warning", 
          pesan: siswaDitemukan.nama + " sudah absen! (Oleh: " + petugasTerdahulu + ")" 
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    var photoUrl = "";
    
    // 1. Simpan foto ke Google Drive jika ada
    if (data.photoBase64 && data.photoName) {
      var folderName = "Dokumentasi_Geotag_WaliAsuh";
      var folders = DriveApp.getFoldersByName(folderName);
      var folder;
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder(folderName);
      }
      
      var splitBase = data.photoBase64.split(",");
      var contentType = splitBase[0].split(":")[1].split(";")[0];
      var bytes = Utilities.base64Decode(splitBase[1]);
      var blob = Utilities.newBlob(bytes, contentType, data.photoName);
      
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoUrl = file.getUrl();
    }
    
    // 2. Generate Refleksi Otomatis menggunakan Gemini AI
    var hasilRefleksi = ANALISIS_AI(data.catatan, "refleksi");

    // 3. Simpan data ke LogAbsen (Kolom A s.d. H)
    var timestamp = new Date();
    var tanggal = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var waktu = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "HH:mm:ss");
    
    sheetLog.appendRow([timestamp, tanggal, waktu, siswaDitemukan.id, siswaDitemukan.nama, siswaDitemukan.kelas, data.kegiatan, data.petugas, hasilRefleksi, photoUrl]);
    
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      siswa: siswaDitemukan 
    })).setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      pesan: err.message 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Membuat analisis refleksi mendalam menggunakan Gemini AI.
 */
function ANALISIS_AI(teksUraian, mode) {
  if (!teksUraian || teksUraian.toString().trim() === "") {
    return "Catatan kegiatan kosong.";
  }
  
  // 🔑 MASUKKAN API KEY GEMINI ASLI ANDA DI SINI (Contoh: "AIzaSy...")
  var apiKey = "PASTE_API_KEY_GEMINI_ASLI_DI_SINI"; 
  
  // Pengecekan: Jika API Key belum diganti, gunakan analisis lokal
  if (apiKey === "PASTE_API_KEY_GEMINI_ASLI_DI_SINI" || apiKey === "") {
    return analisisLokal(teksUraian, mode);
  }

  var prompt = "";
  var modeLower = mode ? mode.toString().toLowerCase() : "refleksi";

  if (modeLower === "refleksi") {
    prompt = "Bertindaklah sebagai pendidik dan wali asuh profesional. " +
             "Analisis uraian kegiatan berikut, lalu buat 1-2 kalimat REFLEKSI DIRI yang mendalam dan bermakna. " +
             "Fokus pada nilai pembelajaran, perkembangan santri/peserta didik, atau evaluasi tindak lanjut ke depan. " +
             "Jangan sekadar mengulang kalimat input atau menambahkan kata pujian di awal.\n\n" +
             "Uraian Kegiatan: \"" + teksUraian + "\"";
  } else {
    prompt = "Tentukan sentimen dari uraian kegiatan berikut (Pilih salah satu: Positif, Netral, atau Evaluatif): \"" + teksUraian + "\"";
  }

  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + apiKey;
  var payload = {
    "contents": [{
      "parts": [{"text": prompt}]
    }],
    "generationConfig": {
      "temperature": 0.7,
      "maxOutputTokens": 150
    }
  };
  
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    
    if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts[0].text) {
      return json.candidates[0].content.parts[0].text.trim();
    } else if (json.error) {
      return "AI Error: " + json.error.message;
    } else {
      return "Gagal memproses AI.";
    }
  } catch (err) {
    return "Error Fetch: " + err.message;
  }
}

// Fungsi cadangan lokal jika API Key belum dikonfigurasi
function analisisLokal(teks, mode) {
  var t = teks.toLowerCase();
  if (mode.toLowerCase() === "sentimen") {
    if (t.includes("lancar") || t.includes("sukses") || t.includes("baik") || t.includes("berhasil")) return "Positif";
    if (t.includes("kendala") || t.includes("masalah") || t.includes("gagal") || t.includes("terlambat")) return "Evaluatif";
    return "Netral";
  } else {
    return "Pendampingan terlaksana dengan baik, perlu pemantauan konsisten terhadap perkembangan peserta didik.";
  }
}

// 🛠️ FUNGSI HELPER: Memastikan nilai waktu HANYA mengembalikan format "HH:mm:ss"
function formatJamBersih(val) {
  if (!val) return "-";
  
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm:ss");
  }
  
  var str = String(val).trim();
  var match = str.match(/\d{2}:\d{2}(:\d{2})?/);
  if (match) {
    return match[0];
  }
  
  try {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), "HH:mm:ss");
    }
  } catch(e){}
  
  return str;
}
