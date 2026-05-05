/**
 * BACKEND GOOGLE APPS SCRIPT (VERSI REAL)
 * File: Code.gs
 */

const SPREADSHEET_PEGAWAI_URL = "https://docs.google.com/spreadsheets/d/1vQBjRbWOH3IpODEuX6lnJum0M0Jcnh8Eo-kKEHz8k4o/edit";
const SPREADSHEET_ABSENSI_URL = "https://docs.google.com/spreadsheets/d/1s1zVQStEiZ9va2HGH-LTOZ6blW-OvajopNjgfEv3EsE/edit";

// 1. Fungsi melayani UI Web
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('SPPD Tanjung Puri')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 2. Fungsi inisialisasi / Setup Database otomatis untuk SPPD
function initSheets() {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_PEGAWAI_URL);
  
  let sheetSppd = ss.getSheetByName("Data_SPPD");
  // Jika sheet Data_SPPD belum ada, buat otomatis
  if (!sheetSppd) {
    sheetSppd = ss.insertSheet("Data_SPPD");
    sheetSppd.appendRow(["ID_SPPD", "Waktu_Dibuat", "ID_QR", "Nama_Pegawai", "Tujuan", "Tgl_Berangkat", "Tgl_Pulang", "Status"]);
  }

  let sheetLaporan = ss.getSheetByName("Data_Laporan");
  if (!sheetLaporan) {
    sheetLaporan = ss.insertSheet("Data_Laporan");
    sheetLaporan.appendRow(["Waktu_Dibuat", "ID_SPPD", "Hasil_Kegiatan"]);
  }

  let sheetKwitansi = ss.getSheetByName("Data_Kwitansi");
  if (!sheetKwitansi) {
    sheetKwitansi = ss.insertSheet("Data_Kwitansi");
    sheetKwitansi.appendRow(["Waktu_Dibuat", "ID_SPPD", "Transport", "Penginapan", "Uang_Harian", "Total"]);
  }
}

// 3. Fungsi Ambil Data Awal (Pegawai & Riwayat SPPD) saat Login
function getInitData() {
  try {
    initSheets();
    const ss = SpreadsheetApp.openByUrl(SPREADSHEET_PEGAWAI_URL);
    
    // --- Ambil Data Pegawai ---
    const sheetPegawai = ss.getSheets()[0];
    const dataPegawai = sheetPegawai.getDataRange().getValues();
    let pegawaiList = [];
    
    for (let i = 1; i < dataPegawai.length; i++) {
      if(!dataPegawai[i][0]) continue; // Lewati jika ID_QR kosong
      pegawaiList.push({
        id_qr: dataPegawai[i][0],
        nama: dataPegawai[i][1],
        nip: dataPegawai[i][2],
        jabatan: dataPegawai[i][3],
        status: dataPegawai[i][4],
        lokasi: dataPegawai[i][5],
        pangkat: dataPegawai[i][6]
      });
    }

    // --- Ambil Data SPPD ---
    const sheetSppd = ss.getSheetByName("Data_SPPD");
    const dataSppd = sheetSppd.getDataRange().getValues();
    let sppdList = [];
    
    for (let j = 1; j < dataSppd.length; j++) {
      if(!dataSppd[j][0]) continue;
      
      // Ambil tgl dalam string agar aman di passing ke front-end
      let berangkat = dataSppd[j][5] instanceof Date ? dataSppd[j][5].toISOString().split('T')[0] : dataSppd[j][5];
      let pulang = dataSppd[j][6] instanceof Date ? dataSppd[j][6].toISOString().split('T')[0] : dataSppd[j][6];

      sppdList.push({
        id: dataSppd[j][0],
        id_qr: dataSppd[j][2],
        nama: dataSppd[j][3],
        tujuan: dataSppd[j][4],
        berangkat: berangkat,
        pulang: pulang,
        status: dataSppd[j][7]
      });
    }
    
    return JSON.stringify({status: 'success', pegawai: pegawaiList, sppd: sppdList});
  } catch (e) {
    return JSON.stringify({status: 'error', message: e.toString()});
  }
}

// 4. Fungsi Mengecek Absensi Realtime dari Sheet Absensi
function cekAbsensiReal(id_qr, tanggalStr) {
  try {
    const ss = SpreadsheetApp.openByUrl(SPREADSHEET_ABSENSI_URL);
    const sheet = ss.getSheetByName("Absensi");
    
    if(!sheet) return JSON.stringify({status: 'error', message: 'Sheet "Absensi" tidak ditemukan.'});

    const data = sheet.getDataRange().getValues();
    let statusAbsen = "HADIR"; // Default jika tidak ada data absen di hari tsb
    
    // Looping dari bawah ke atas agar mendapat input absen yang paling baru (jika ada revisi di sheet)
    for(let i = data.length - 1; i >= 1; i--) {
      let tglRow = data[i][1]; // Kolom B
      let tglFormatted = "";
      
      if (tglRow instanceof Date) {
          // Koreksi zona waktu untuk Indonesia agar tgl tidak mundur 1 hari (karena offset UTC)
          let tzDate = new Date(tglRow.getTime() - (tglRow.getTimezoneOffset() * 60000));
          tglFormatted = tzDate.toISOString().split('T')[0];
      } else {
          tglFormatted = tglRow;
      }
      
      let idQrRow = data[i][4]; // Kolom E
      let ketRow = data[i][10]; // Kolom K
      
      // Jika cocok ID QR dan Tanggal
      if(idQrRow == id_qr && tglFormatted == tanggalStr) {
        statusAbsen = ketRow;
        break; 
      }
    }
    
    return JSON.stringify({status: 'success', absen: statusAbsen});
  } catch (e) {
    return JSON.stringify({status: 'error', message: e.toString()});
  }
}

// 5. Fungsi Menyimpan SPT / SPPD secara nyata ke Spreadsheet
function simpanSPPDReal(payload) {
  try {
    const ss = SpreadsheetApp.openByUrl(SPREADSHEET_PEGAWAI_URL);
    const sheet = ss.getSheetByName("Data_SPPD");
    
    let now = new Date();
    // Buat nomor surat / ID unik berdasarkan timestamp
    let id_sppd = "SPT-" + now.getTime().toString().slice(-5);
    
    // Tulis ke row paling bawah di Sheet "Data_SPPD"
    sheet.appendRow([
      id_sppd, 
      now, 
      payload.id_qr, 
      payload.nama, 
      payload.tujuan, 
      payload.berangkat, 
      payload.pulang, 
      "Direncanakan" // Status awal selalu direncanakan
    ]);
    
    // Kembalikan data ke UI agar tampil di tabel tanpa refresh browser
    return JSON.stringify({
      status: 'success', 
      data: {
        id: id_sppd,
        id_qr: payload.id_qr,
        nama: payload.nama,
        tujuan: payload.tujuan,
        berangkat: payload.berangkat,
        pulang: payload.pulang,
        status: "Direncanakan"
      }
    });
  } catch(e) {
    return JSON.stringify({status: 'error', message: e.toString()});
  }
}

// 6. Fungsi Hitung Absensi Hari Ini untuk Dashboard
function getAbsensiHariIni() {
  try {
    const ss = SpreadsheetApp.openByUrl(SPREADSHEET_ABSENSI_URL);
    const sheet = ss.getSheetByName("Absensi");
    if(!sheet) return JSON.stringify({status: 'error', message: 'Sheet Absensi tidak ditemukan'});
    
    const data = sheet.getDataRange().getValues();
    let countAbsen = 0;
    
    // Format tanggal hari ini (YYYY-MM-DD) di Indonesia
    let now = new Date();
    let tzDate = new Date(now.getTime() - (now.getTimezoneOffset() * 60000));
    let todayStr = tzDate.toISOString().split('T')[0];
    
    // Termasuk SAKIT, TL, ALPA, IZIN, CUTI
    let targetStatus = ['CUTI', 'IZIN', 'ALPA', 'TL', 'SAKIT'];
    
    for(let i = 1; i < data.length; i++) {
      let tglRow = data[i][1]; 
      let tglFormatted = "";
      if (tglRow instanceof Date) {
          let tzRow = new Date(tglRow.getTime() - (tglRow.getTimezoneOffset() * 60000));
          tglFormatted = tzRow.toISOString().split('T')[0];
      } else {
          tglFormatted = tglRow;
      }
      
      let ketRow = data[i][10]; // Kolom K (Keterangan)
      
      if(tglFormatted === todayStr && targetStatus.includes(ketRow)) {
        countAbsen++;
      }
    }
    return JSON.stringify({status: 'success', count: countAbsen});
  } catch (e) {
    return JSON.stringify({status: 'error', message: e.toString()});
  }
}

// 7. Fungsi Menyimpan Laporan
function simpanLaporanReal(payload) {
  try {
    const ss = SpreadsheetApp.openByUrl(SPREADSHEET_PEGAWAI_URL);
    const sheet = ss.getSheetByName("Data_Laporan");
    sheet.appendRow([new Date(), payload.id_sppd, payload.hasil]);
    
    // Update status SPPD jadi Selesai
    const sheetSppd = ss.getSheetByName("Data_SPPD");
    const dataSppd = sheetSppd.getDataRange().getValues();
    for(let i = 1; i < dataSppd.length; i++) {
       if(dataSppd[i][0] == payload.id_sppd) {
          sheetSppd.getRange(i+1, 8).setValue("Selesai"); // Kolom 8 = Status
          break;
       }
    }
    return JSON.stringify({status: 'success', id_sppd: payload.id_sppd});
  } catch(e) {
    return JSON.stringify({status: 'error', message: e.toString()});
  }
}

// 8. Fungsi Menyimpan Kwitansi
function simpanKwitansiReal(payload) {
  try {
    const ss = SpreadsheetApp.openByUrl(SPREADSHEET_PEGAWAI_URL);
    const sheet = ss.getSheetByName("Data_Kwitansi");
    sheet.appendRow([
      new Date(), 
      payload.id_sppd, 
      payload.transport, 
      payload.penginapan, 
      payload.uang_harian, 
      payload.total
    ]);
    return JSON.stringify({status: 'success'});
  } catch(e) {
    return JSON.stringify({status: 'error', message: e.toString()});
  }
}
