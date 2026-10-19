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
      
      for (var j = 1; j < dataLog.length; j++) {
        // Kolom B=Tanggal(1), C=Waktu(2), D=ID Siswa(3), G=Kegiatan(6), H=Petugas(7)
        if (String(dataLog[j][1]) === hariIni && String(dataLog[j][6]) === kegiatan) {
          var idSs = String(dataLog[j][3]).trim();
          
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
    
    for (var m = 1; m < dataLog.length; m++) {
      if (String(dataLog[m][1]) === hariIniScan && String(dataLog[m][3]) === idScan && String(dataLog[m][6]) === data.kegiatan) {
        var petugasTerdahulu = dataLog[m][7] || "Petugas Lain";
        return ContentService.createTextOutput(JSON.stringify({ 
          status: "warning", 
          pesan: siswaDitemukan.nama + " sudah absen! (Oleh: " + petugasTerdahulu + ")" 
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // Simpan data ke LogAbsen
    var timestamp = new Date();
    var tanggal = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var waktu = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "HH:mm:ss");
    
    sheetLog.appendRow([timestamp, tanggal, waktu, siswaDitemukan.id, siswaDitemukan.nama, siswaDitemukan.kelas, data.kegiatan, data.petugas]);
    
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