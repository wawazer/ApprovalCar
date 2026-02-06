/************************************************************
 * CONFIGURATION & MAPPING 
 ************************************************************/
const SPREADSHEET_ID = '1hg3CynBrqFci3kEg7c611ht4jtIDO9Ocyh2RPc1ixE8';
const SHEET_NAME     = 'Request Mobil';
const FORM_ID        = '1k-evV4VwEn29VSHheqlvgML6d0SQfGE7SZiWh0t-9s4';
const WEBAPP_URL     = 'https://script.google.com/macros/s/AKfycbxICLEfVFaYQdZydstk4kwmZHipDNvTSxB2xj1DwATX9wHAmCCW8FZQf9SiwxiEgtlOnQ/exec'; 

const LV1_APPROVER_EMAILS = 'muhammad.wawazer@pelindo.co.id';
const CC_EMAIL = 'wawazer@gmail.com';

const COL_TIMESTAMP      = 1;  
const COL_TGL_BERANGKAT  = 3;  
const COL_TGL_KEPULANGAN = 4;  
const COL_PILIH_KENDARAAN= 5;  
const COL_EMAIL_PEMOHON  = 6;  
const COL_UNIT_KERJA     = 7;  
const COL_DAFTAR_TAMU    = 8;  
const COL_TUJUAN         = 9;  
const COL_HOTEL          = 10; 
const COL_NAMA_PEMOHON   = 13; 
const COL_WA_PEMOHON     = 15; 
const COL_STATUS_LV1     = 17; 
const COL_STATUS_FINAL   = 19; 
const COL_REASON_LV1     = 20; 

/************************************************************
 * 1. TRIGGER: ON FORM SUBMIT 
 ************************************************************/
function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const row = e.range.getRow();
  
  const vals = sheet.getRange(row, 1, 1, 21).getValues()[0]; 
  
  const d = {
    row: row,
    tglMulaiRaw: vals[COL_TGL_BERANGKAT-1],
    tglSelesaiRaw: vals[COL_TGL_KEPULANGAN-1],
    mobilRaw: vals[COL_PILIH_KENDARAAN-1].toString(),
    emailPengaju: vals[COL_EMAIL_PEMOHON-1],
    noWaPengaju: vals[COL_WA_PEMOHON-1], 
    unitKerja: vals[COL_UNIT_KERJA-1],
    daftarTamu: vals[COL_DAFTAR_TAMU-1],
    tujuan: vals[COL_TUJUAN-1],
    hotel: vals[COL_HOTEL-1],
    namaPIC: vals[COL_NAMA_PEMOHON-1]
  };

  let namaMobilBersih = d.mobilRaw.replace("⚠️ ", "").replace("✅ ", "").split(" (")[0].trim();
  d.nomorKendaraan = namaMobilBersih;
  sheet.getRange(row, COL_PILIH_KENDARAAN).setValue(namaMobilBersih);

  const cek = checkJadwalBentrok(namaMobilBersih, d.tglMulaiRaw, d.tglSelesaiRaw, row);
  
  if (cek.bentrok) {
    sheet.getRange(row, COL_STATUS_LV1).setValue("Rejected (System)");
    sheet.getRange(row, COL_STATUS_FINAL).setValue("Rejected");
    sheet.getRange(row, COL_REASON_LV1).setValue(`AUTO-REJECT: Bentrok dengan ${cek.pic} (${cek.tgl})`);
    
    MailApp.sendEmail({
      to: d.emailPengaju,
      subject: 'Mobil Tidak Tersedia - Penolakan Otomatis',
      htmlBody: `Yth <b>${d.namaPIC}</b>,<br><br>Unit <b>${namaMobilBersih}</b> ditolak otomatis karena bentrok jadwal.`
    });


    if (d.noWaPengaju) {
      const pesanBentrok = `❌ *PENGAJUAN DITOLAK OTOMATIS*\n\nHalo *${d.namaPIC}*,\nMohon maaf, unit *${namaMobilBersih}* ditolak oleh sistem karena bentrok dengan jadwal *${cek.pic}* (${cek.tgl}).\n\nSilakan pilih waktu atau unit lain.`;
      kirimWAWatzap(d.noWaPengaju, pesanBentrok);
    }

    return; 
  }

  d.tglBerangkat = d.tglMulaiRaw instanceof Date ? Utilities.formatDate(d.tglMulaiRaw, "GMT+7", "dd/MM/yyyy HH:mm") : d.tglMulaiRaw;
  d.tglKepulangan = d.tglSelesaiRaw instanceof Date ? Utilities.formatDate(d.tglSelesaiRaw, "GMT+7", "dd/MM/yyyy HH:mm") : d.tglSelesaiRaw;

  sendApprovalToLv1_(d);

  const noWaLV1 = "6281803216767"; 
  const pesanLV1 = `🔔 *PENGAJUAN MOBIL BARU*\n\nYth. Bapak/Ibu,\nAda permohonan kendaraan baru.\n\n*PIC:* ${d.namaPIC}\n*Unit:* ${d.nomorKendaraan}\n*Tujuan:* ${d.tujuan}\n*Waktu:* ${d.tglBerangkat}\n\nMohon cek email Anda untuk melakukan persetujuan.`;
  
  Logger.log("Mengirim WA ke LV1: " + noWaLV1);
  kirimWAWatzap(noWaLV1, pesanLV1);
}

/************************************************************
 * 2. WEBAPP HANDLER 
 ************************************************************/
function doGet(e) {
  // Tambahkan Header X-Frame-Options agar bisa dibuka di semua browser
  const output = processRequest(e);
  return output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processRequest(e) {
  try {
    const action = e.parameter.action;
    const row = parseInt(e.parameter.row);
    const reason = e.parameter.reason || "";

    if (action === "reject" && !e.parameter.reason) return renderReasonForm(row);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const status = action === "approve" ? "Approved" : "Rejected";
    
    sheet.getRange(row, COL_STATUS_LV1).setValue(status);
    sheet.getRange(row, COL_STATUS_FINAL).setValue(status);
    if (action === "reject") sheet.getRange(row, COL_REASON_LV1).setValue(reason);

    const vals = sheet.getRange(row, 1, 1, 21).getValues()[0];
    const noWaPemohon = vals[COL_WA_PEMOHON - 1]; 
    const namaPIC = vals[COL_NAMA_PEMOHON - 1];
    const unitMobil = vals[COL_PILIH_KENDARAAN - 1].toString();
    const tujuan = vals[COL_TUJUAN - 1]; 

    // =======================================================
    // TAMBAHAN: UPDATE STATUS & KIRIM WA DRIVER
    // =======================================================
    if (action === "approve") {
      const sheetMaster = ss.getSheetByName("Master_Armada");
      if (sheetMaster) {
        const dataMaster = sheetMaster.getDataRange().getValues();
        for (let i = 1; i < dataMaster.length; i++) {
          if (unitMobil.indexOf(dataMaster[i][0]) !== -1) { 
            // Update status mobil jadi In Use
            sheetMaster.getRange(i + 1, 5).setValue("In Use"); 

            // --- Bagian Kirim WA ke Driver ---
            const namaDriver = dataMaster[i][2]; 
            const noWaDriver = dataMaster[i][3]; 
            
            if (noWaDriver) {
              const linkFormKembali = "https://docs.google.com/forms/d/e/1FAIpQLSdtvfgFGcgfVM-cggtpkJt5Uw-zqoIx0zHzkK5mlVduT_UWog/viewform";
              const pesanDriver = `🚛 *TUGAS BARU: ${unitMobil}*\n\nHalo *${namaDriver}*,\nAda tugas pengantaran:\n📍 *Tujuan:* ${tujuan}\n👤 *PIC:* ${namaPIC}\n\nJika sudah kembali ke kantor, mohon klik link ini untuk lapor KM Akhir:\n👉 ${linkFormKembali}`;
              
              kirimWAWatzap(noWaDriver, pesanDriver);
            }
            break;
          }
        }
      }
    }

    if (noWaPemohon) {
      let pesanWA = "";
      if (action === "approve") {
        pesanWA = `✅ *PERMOHONAN DISETUJUI*\n\nHalo *${namaPIC}*,\nPermohonan mobil *${unitMobil}* Anda telah *DISETUJUI*.\n\nSelamat bertugas!`;
      } else {
        pesanWA = `❌ *PERMOHONAN DITOLAK*\n\nHalo *${namaPIC}*,\nMohon maaf, permohonan mobil *${unitMobil}* Anda *DITOLAK*.\n*Alasan:* ${reason}`;
      }
      kirimWAWatzap(noWaPemohon, pesanWA);
    }

    kirimEmailFinal_(row, status, reason);

    return HtmlService.createHtmlOutput(`
      <div style="font-family:sans-serif;text-align:center;padding-top:50px;">
        <div style="display:inline-block;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,0.1);">
          <h2 style="color:#28a745;">✅ BERHASIL!</h2>
          <p>Status baris <b>${row}</b> sudah di-update menjadi <b>${status}</b>.</p>
          <p>Status mobil & Notifikasi Driver telah diproses.</p>
          <p>Anda bisa menutup halaman ini.</p>
        </div>
      </div>`);
  } catch (err) {
    return HtmlService.createHtmlOutput("<b>ERROR:</b> " + err.message);
  }
}

// function processRequest(e) {
//   try {
//     const action = e.parameter.action;
//     const row = parseInt(e.parameter.row);
//     const reason = e.parameter.reason || "";

//     if (action === "reject" && !e.parameter.reason) return renderReasonForm(row);

//     const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//     const sheet = ss.getSheetByName(SHEET_NAME);
//     const status = action === "approve" ? "Approved" : "Rejected";
    
//     sheet.getRange(row, COL_STATUS_LV1).setValue(status);
//     sheet.getRange(row, COL_STATUS_FINAL).setValue(status);
//     if (action === "reject") sheet.getRange(row, COL_REASON_LV1).setValue(reason);

//     const vals = sheet.getRange(row, 1, 1, 21).getValues()[0];
//     const noWaPemohon = vals[COL_WA_PEMOHON - 1]; 
//     const namaPIC = vals[COL_NAMA_PEMOHON - 1];
//     const unitMobil = vals[COL_PILIH_KENDARAAN - 1].toString();

//     // =======================================================
//     // TAMBAHAN: UPDATE STATUS DI MASTER_ARMADA JADI "In Use"
//     // =======================================================
//     if (action === "approve") {
//       const sheetMaster = ss.getSheetByName("Master_Armada");
//       if (sheetMaster) {
//         const dataMaster = sheetMaster.getDataRange().getValues();
//         for (let i = 1; i < dataMaster.length; i++) {
//           // Cek apakah plat nomor di Master_Armada ada di dalam teks unitMobil
//           if (unitMobil.indexOf(dataMaster[i][0]) !== -1) { 
//             sheetMaster.getRange(i + 1, 5).setValue("In Use"); //
//             break;
//           }
//         }
//       }
//     }
//     // =======================================================

//     if (noWaPemohon) {
//       let pesanWA = "";
//       if (action === "approve") {
//         pesanWA = `✅ *PERMOHONAN DISETUJUI*\n\nHalo *${namaPIC}*,\nPermohonan mobil *${unitMobil}* Anda telah *DISETUJUI*.\n\nSelamat bertugas!`;
//       } else {
//         pesanWA = `❌ *PERMOHONAN DITOLAK*\n\nHalo *${namaPIC}*,\nMohon maaf, permohonan mobil *${unitMobil}* Anda *DITOLAK*.\n*Alasan:* ${reason}`;
//       }
//       kirimWAWatzap(noWaPemohon, pesanWA);
//     }

//     kirimEmailFinal_(row, status, reason);

//     return HtmlService.createHtmlOutput(`
//       <div style="font-family:sans-serif;text-align:center;padding-top:50px;">
//         <div style="display:inline-block;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,0.1);">
//           <h2 style="color:#28a745;">✅ BERHASIL!</h2>
//           <p>Status baris <b>${row}</b> sudah di-update menjadi <b>${status}</b>.</p>
//           <p>Status mobil di Master Armada juga telah diperbarui.</p>
//           <p>Anda bisa menutup halaman ini.</p>
//         </div>
//       </div>`);
//   } catch (err) {
//     return HtmlService.createHtmlOutput("<b>ERROR:</b> " + err.message);
//   }
// }

/************************************************************
 * 3. RENDER FORM REJECT
 ************************************************************/
function renderReasonForm(row) {
  const url = ScriptApp.getService().getUrl();
  return HtmlService.createHtmlOutput(`
    <div style="font-family:sans-serif;padding:20px;max-width:400px;margin:auto;border:1px solid #ddd;border-radius:10px;">
      <form action="${url}" method="get">
        <h3 style="color:#d93025;">Alasan Penolakan</h3>
        <p>Anda akan menolak permintaan pada baris <b>${row}</b>.</p>
        <input type="hidden" name="action" value="reject"><input type="hidden" name="row" value="${row}">
        <textarea name="reason" rows="4" style="width:100%; padding:10px; border-radius:5px; border:1px solid #ccc;" placeholder="Tulis alasan di sini..." required></textarea><br><br>
        <button type="submit" style="background:#d93025; color:white; border:none; padding:10px 20px; cursor:pointer; width:100%; font-weight:bold; border-radius:5px;">Kirim Penolakan</button>
      </form>
    </div>`).setTitle("Form Penolakan");
}

/************************************************************
 * 4. (CHECK BENTROK & EMAIL)
 ************************************************************/
function checkJadwalBentrok(mobilReq, tglMulaiReq, tglSelesaiReq, currentRow) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const startReq = new Date(tglMulaiReq).getTime();
  const endReq = new Date(tglSelesaiReq).getTime();
  if (isNaN(startReq)) return { bentrok: false };

  for (let i = 1; i < data.length; i++) {
    if ((i + 1) === currentRow) continue; 
    if (data[i][COL_STATUS_FINAL - 1] === "Approved") {
      let mobilDiSheet = data[i][COL_PILIH_KENDARAAN - 1].toString().split(" (")[0].replace("⚠️ ", "").replace("✅ ", "").trim();
      const startSheet = new Date(data[i][COL_TGL_BERANGKAT - 1]).getTime();
      const endSheet = new Date(data[i][COL_TGL_KEPULANGAN - 1]).getTime();
      if (mobilDiSheet === mobilReq && startReq < endSheet && endReq > startSheet) {
        return { bentrok: true, pic: data[i][COL_NAMA_PEMOHON - 1], tgl: Utilities.formatDate(new Date(startSheet), "GMT+7", "dd/MM HH:mm") };
      }
    }
  }
  return { bentrok: false };
}

function sendApprovalToLv1_(d) {
  const approveUrl = `${WEBAPP_URL}?action=approve&row=${d.row}`;
  const rejectUrl  = `${WEBAPP_URL}?action=reject&row=${d.row}`;
  
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; border: 1px solid #eee; padding: 20px; border-radius: 8px;">
      <h3 style="color: #004a99; border-bottom: 2px solid #004a99; padding-bottom: 10px;">Permohonan Kendaraan</h3>
      <table style="width: 100%; border-collapse: collapse;">
        <tr><td style="padding:8px 0; font-weight:bold; width:120px;">PIC</td><td>: ${d.namaPIC}</td></tr>
        <tr><td style="padding:8px 0; font-weight:bold;">Unit Kerja</td><td>: ${d.unitKerja}</td></tr>
        <tr style="background:#f9f9f9;"><td style="padding:8px 0; font-weight:bold;">Mobil</td><td style="color:#d93025; font-weight:bold;">: ${d.nomorKendaraan}</td></tr>
        <tr><td style="padding:8px 0; font-weight:bold;">Waktu</td><td>: ${d.tglBerangkat} s.d. ${d.tglKepulangan}</td></tr>
        <tr style="background:#f9f9f9;"><td style="padding:8px 0; font-weight:bold;">Tujuan</td><td>: ${d.tujuan}</td></tr>
        <tr><td style="padding:8px 0; font-weight:bold;">Tamu</td><td>: ${d.daftarTamu}</td></tr>
        <tr><td style="padding:8px 0; font-weight:bold;">Hotel</td><td>: ${d.hotel || "-"}</td></tr>
      </table>
      <div style="margin-top:25px; text-align:center;">
        <a href="${approveUrl}" style="background:#28a745; color:white; padding:12px 25px; text-decoration:none; border-radius:5px; font-weight:bold; margin-right:10px;">APPROVE</a>
        <a href="${rejectUrl}" style="background:#dc3545; color:white; padding:12px 25px; text-decoration:none; border-radius:5px; font-weight:bold;">REJECT</a>
      </div>
    </div>`;

  MailApp.sendEmail({ to: LV1_APPROVER_EMAILS, cc: CC_EMAIL, subject: `[REQUEST] Mobil: ${d.namaPIC} - ${d.unitKerja}`, htmlBody: htmlBody });
}

function kirimEmailFinal_(row, status, reason) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const vals = sheet.getRange(row, 1, 1, 20).getValues()[0];
  const email = vals[COL_EMAIL_PEMOHON-1];
  const pic = vals[COL_NAMA_PEMOHON-1];
  
  // Menentukan teks status untuk subjek agar pasti sesuai dengan input
  var statusTeks = status || "Update"; 

  const htmlBody = `
    <div style="font-family:Arial; padding:20px; border:1px solid #ddd;">
      <h3>Status Permohonan Mobil: <span style="color:${status==='Approved'?'#28a745':'#dc3545'};">${status}</span></h3>
      <p>Halo <b>${pic}</b>,</p>
      <p>Permohonan kendaraan Anda telah diperbarui menjadi <b>${status}</b>.</p>
      ${reason ? `<p style="background:#f2f2f2; padding:10px; border-left:4px solid #dc3545;"><b>Alasan:</b> ${reason}</p>` : ""}
      <p>Terima kasih.</p>
    </div>`;
  
  MailApp.sendEmail({ 
    to: email, 
    subject: "[" + statusTeks + "] Status Peminjaman Mobil", 
    htmlBody: htmlBody 
  });
}

/************************************************************
 * 3. SCHEDULER: UPDATE DROPDOWN FORM
 ************************************************************/
// function filterMobilTersedia() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const masterSheet = ss.getSheetByName('Master_Mobil');
//   const requestSheet = ss.getSheetByName('Request Mobil');
//   const logSheet = ss.getSheetByName('Log_Scheduler');
//   const form = FormApp.openById(FORM_ID); 
  
//   const sekarang = new Date();
//   const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 1).getValues();
//   const listMobilMaster = masterData.map(row => row[0].toString().trim()).filter(String);
//   const dataRequest = requestSheet.getDataRange().getValues();
  
//   let statusMobilMap = {}; 
//   let mobilTerpakai = [];
  
//   // cek masa depan: 48 Jam agar user tahu jadwal besok
//   const rentangMasaDepan = new Date(sekarang.getTime() + (48 * 60 * 60 * 1000)); 

//   for (let i = 1; i < dataRequest.length; i++) {
//     const statusFinal = dataRequest[i][COL_STATUS_FINAL - 1];
//     const tglMulaiRaw = dataRequest[i][COL_TGL_BERANGKAT - 1];
//     const tglSelesaiRaw = dataRequest[i][COL_TGL_KEPULANGAN - 1];
    
//     const tglMulai = new Date(tglMulaiRaw);
//     const tglSelesai = new Date(tglSelesaiRaw);
    
//     const noMobil = dataRequest[i][COL_PILIH_KENDARAAN - 1].toString()
//                     .split(" (")[0]
//                     .replace("⚠️ ", "")
//                     .replace("✅ ", "")
//                     .trim();

//     if (statusFinal === "Approved" && !isNaN(tglMulai.getTime())) {
      
//       // KONDISI 1: MOBIL SEDANG DIGUNAKAN SAAT INI
//       if (sekarang >= tglMulai && sekarang <= tglSelesai) {
//         statusMobilMap[noMobil] = "SEDANG JALAN s.d. " + Utilities.formatDate(tglSelesai, "GMT+7", "dd/MM HH:mm");
//         if (!mobilTerpakai.includes(noMobil)) mobilTerpakai.push(noMobil);
//       } 
      
//       // KONDISI 2: MOBIL SUDAH DI-BOOKED UNTUK JADWAL MENDATANG (Dalam 48 Jam)
//       else if (tglMulai > sekarang && tglMulai <= rentangMasaDepan) {
//         if (!statusMobilMap[noMobil]) {
//           statusMobilMap[noMobil] = "BOOKED " + 
//                                     Utilities.formatDate(tglMulai, "GMT+7", "dd/MM HH:mm") + 
//                                     " s.d. " + 
//                                     Utilities.formatDate(tglSelesai, "GMT+7", "dd/MM HH:mm");
//           if (!mobilTerpakai.includes(noMobil)) mobilTerpakai.push(noMobil);
//         }
//       }
//     }
//   }

//   // Membuat daftar tampilan baru untuk dropdown
//   const listTampilanBaru = listMobilMaster.map(mobil => {
//     if (statusMobilMap[mobil]) {
//       return `⚠️ ${mobil} (${statusMobilMap[mobil]})`;
//     }
//     return `✅ ${mobil} (Tersedia)`;
//   });

//   const item = form.getItems(FormApp.ItemType.LIST).find(i => i.getTitle().trim() === "Pilih Kendaraan");
//   let statusUpdate = "Failed: Item Not Found";
  
//   if (item) {
//     item.asListItem().setChoiceValues(listTampilanBaru);
//     statusUpdate = "Success";
//   }

//   if (logSheet) {
//     logSheet.appendRow([
//       sekarang, 
//       statusUpdate, 
//       mobilTerpakai.join(", ") || "Semua Tersedia", 
//       "Update otomatis dropdown"
//     ]);
//   }
// }

function filterMobilTersedia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('Master_Mobil');
  const requestSheet = ss.getSheetByName('Request Mobil');
  const armadaSheet = ss.getSheetByName('Master_Armada'); 
  const logSheet = ss.getSheetByName('Log_Scheduler');
  const form = FormApp.openById(FORM_ID); 
  
  const sekarang = new Date();
  const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 1).getValues();
  const listMobilMaster = masterData.map(row => row[0].toString().trim()).filter(String);
  const dataRequest = requestSheet.getDataRange().getValues();

  // --- 1. AMBIL STATUS REAL-TIME ---
  const dataArmada = armadaSheet.getDataRange().getValues();
  let statusRealTimeMap = {};
  for (let j = 1; j < dataArmada.length; j++) {
    const plat = dataArmada[j][0].toString().trim();
    const jenis = dataArmada[j][1].toString().trim();
    const gabungan = jenis + " - " + plat;
    const status = dataArmada[j][4] ? dataArmada[j][4].toString().trim() : "";
    statusRealTimeMap[gabungan] = status;
  }
  
  let statusMobilMap = {}; 
  let mobilTerpakai = [];
  // Tingkatkan rentang masa depan ke 7 hari (168 jam) sesuai request Abang sebelumnya
  const rentangMasaDepan = new Date(sekarang.getTime() + (168 * 60 * 60 * 1000)); 

  for (let i = 1; i < dataRequest.length; i++) {
    const statusFinal = dataRequest[i][COL_STATUS_FINAL - 1];
    const tglMulaiRaw = dataRequest[i][COL_TGL_BERANGKAT - 1];
    const tglSelesaiRaw = dataRequest[i][COL_TGL_KEPULANGAN - 1];
    
    const tglMulai = new Date(tglMulaiRaw);
    const tglSelesai = new Date(tglSelesaiRaw);
    
    const noMobil = dataRequest[i][COL_PILIH_KENDARAAN - 1].toString()
                    .split(" (")[0]
                    .replace("⚠️ ", "")
                    .replace("✅ ", "")
                    .trim();

    if (statusFinal === "Approved" && !isNaN(tglMulai.getTime())) {
      
      // --- 2. CEK APAKAH MOBIL SEHARUSNYA JALAN (DIPERTJAM) ---
      // Syarat: Waktu sekarang sudah masuk jadwal DAN status fisik bukan Available
      if (sekarang >= tglMulai && sekarang <= tglSelesai) {
        if (statusRealTimeMap[noMobil] !== "Available") {
          statusMobilMap[noMobil] = "SEDANG JALAN s.d. " + Utilities.formatDate(tglSelesai, "GMT+7", "dd/MM HH:mm");
          if (!mobilTerpakai.includes(noMobil)) mobilTerpakai.push(noMobil);
        }
      } 
      
      // --- 3. CEK APAKAH MOBIL DI-BOOKED (DIPERBAIKI) ---
      // Jika Approved tapi waktu sekarang BELUM masuk jam jalan, masuk kategori BOOKED
      else if (tglMulai > sekarang && tglMulai <= rentangMasaDepan) {
        if (!statusMobilMap[noMobil]) {
          statusMobilMap[noMobil] = "BOOKED " + 
                                    Utilities.formatDate(tglMulai, "GMT+7", "dd/MM HH:mm") + 
                                    " s.d. " + 
                                    Utilities.formatDate(tglSelesai, "GMT+7", "dd/MM HH:mm");
          if (!mobilTerpakai.includes(noMobil)) mobilTerpakai.push(noMobil);
        }
      }
    }
  }

  // --- 4. BANGUN LIST TAMPILAN BARU ---
  const listTampilanBaru = listMobilMaster.map(mobil => {
    const statusFisik = statusRealTimeMap[mobil];
    const statusJadwal = statusMobilMap[mobil];

    if (statusFisik === "Available") {
      if (statusJadwal && statusJadwal.includes("BOOKED")) {
        return `⚠️ ${mobil} (${statusJadwal})`;
      }
      return `✅ ${mobil} (Tersedia)`;
    }

    if (statusJadwal) {
      return `⚠️ ${mobil} (${statusJadwal})`;
    }

    return `✅ ${mobil} (Tersedia)`;
  });

  const item = form.getItems(FormApp.ItemType.LIST).find(i => i.getTitle().trim() === "Pilih Kendaraan");
  let statusUpdate = "Failed: Item Not Found";
  
  if (item) {
    item.asListItem().setChoiceValues(listTampilanBaru);
    statusUpdate = "Success";
  }

  if (logSheet) {
    logSheet.appendRow([
      sekarang, 
      statusUpdate, 
      mobilTerpakai.join(", ") || "Semua Tersedia", 
      "Update otomatis dropdown (Refresh Ulang Form)"
    ]);
  }
}
/***********
 * WA NYA 
 * 
 * 
 * ***********/
function kirimWAWatzap(noHP, pesan) {
  if (!noHP) return;
  
  // Membersihkan karakter non-angka
  let formattedNo = noHP.toString().replace(/[^0-9]/g, "");
  
  // Otomatis ubah 08xxx menjadi 628xxx
  if (formattedNo.startsWith("0")) {
    formattedNo = "62" + formattedNo.slice(1);
  }

  const url = "https://api.watzap.id/v1/send_message";
  const payload = {
    "api_key": "V3ELWOCBWBWHDEMX",
    "number_key": "VcgcGA4Tq9FkpwMJ",
    "phone_no": formattedNo,
    "message": pesan
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    console.log("Respon Watzap: " + response.getContentText());
  } catch (e) {
    console.error("Gagal kirim WA: " + e.message);
  }
}


/***************************************/
function buatDashboardOtomatis() {  
  const ss = SpreadsheetApp.getActiveSpreadsheet();  
  let dash = ss.getSheetByName('DASHBOARD_MONITORING');  
    
  if (!dash) {  
    dash = ss.insertSheet('DASHBOARD_MONITORING');  
  } else {  
    dash.clear(); 
  }  

  // --- BAGIAN 1: HEADER & STATISTIK RINGKAS ---
  dash.getRange("A1:L1").merge()  
    .setValue("FLEET MONITORING & PERFORMANCE DASHBOARD")  
    .setFontSize(18).setFontWeight("bold")  
    .setBackground("#004a99").setFontColor("white")  
    .setHorizontalAlignment("center").setVerticalAlignment("middle");  
    
  dash.getRange("A2").setValue("Terakhir Update:").setFontWeight("bold");  
  dash.getRange("B2").setValue(new Date()).setNumberFormat("dd/MM/yyyy HH:mm");  

  // Kotak Statistik Atas
  dash.getRange("A4").setValue("Total Armada");
  dash.getRange("B4").setFormula("=COUNTA(Master_Mobil!A2:A)");
  dash.getRange("A5").setValue("Unit In-Use");
  dash.getRange("B5").setFormula("=COUNTIF(Master_Armada!E:E; \"In Use\")");
  dash.getRange("A6").setValue("Unit Available");
  dash.getRange("B6").setFormula("=COUNTIF(Master_Armada!E:E; \"Available\")");

  // --- BAGIAN 2: TABEL JADWAL (SISI KIRI) ---
  dash.getRange("A8").setValue("JADWAL BOOKING ARMADA").setFontWeight("bold").setBackground("#444444").setFontColor("white").setHorizontalAlignment("center"); 
  dash.getRange("B8").setFormula("=TODAY()").setNumberFormat("dd/MM (ddd)");  
  dash.getRange("C8:H8").setFormula("=B8+1");  
  dash.getRange("A8:H8").setFontWeight("bold").setBackground("#eeeeee").setHorizontalAlignment("center");  

  // Isi Nama Mobil & Rumus Jadwal
  dash.getRange("A9").setFormula("=QUERY(Master_Mobil!A2:A; \"SELECT A WHERE A IS NOT NULL\")");  
  const rumusJadwal = "=ARRAYFORMULA(IF(A9:A=\"\"; \"\"; IFERROR(MAP(A9:A; LAMBDA(m; MAP(B8:H8; LAMBDA(t; IF(COUNTIFS('Request Mobil'!$E:$E; m; 'Request Mobil'!$S:$S; \"Approved\"; INT('Request Mobil'!$C:$C); t)>0; \"🔴 BOOKED\"; \"✅ AVAILABLE\"))))); \"✅ AVAILABLE\")))";
  dash.getRange("B9").setFormula(rumusJadwal);  

  // --- BAGIAN 3: ANALISIS BEBAN KERJA (SISI KANAN - KOLOM J) ---
  const colStat = "J";
  const bulanIni = new Date().getMonth() + 1;

  dash.getRange(colStat + "8:" + "L8").merge()
    .setValue("UTILITAS ARMADA (BULAN INI)")
    .setFontWeight("bold").setBackground("#444444").setFontColor("white").setHorizontalAlignment("center");

  dash.getRange(colStat + "9").setValue("NAMA UNIT").setFontWeight("bold").setBackground("#eeeeee");
  dash.getRange("K9").setValue("TOTAL TRIP").setFontWeight("bold").setBackground("#eeeeee");
  dash.getRange("L9").setValue("BEBAN %").setFontWeight("bold").setBackground("#eeeeee");

  // Rumus Utilitas (Query mengambil data dari Request Mobil)
  dash.getRange(colStat + "10").setFormula("=QUERY('Request Mobil'!A2:S; \"SELECT E, COUNT(E) WHERE S = 'Approved' AND MONTH(C)+1 = " + bulanIni + " GROUP BY E LABEL COUNT(E) ''\"; 0)");
  
  // Rumus Persentase Beban (Asumsi 22 hari kerja)
  // Menghitung otomatis ke bawah sebanyak jumlah mobil yang muncul di statistik
  dash.getRange("L10").setFormula("=ARRAYFORMULA(IF(K10:K=\"\"; \"\"; K10:K/22))");
  dash.getRange("L10:L35").setNumberFormat("0%");

  // --- BAGIAN 4: STYLING & FINISHING ---
  // Warna Jadwal
  dash.clearConditionalFormatRules();
  const rangeJadwal = dash.getRange("B9:H35");
  const ruleHijau = SpreadsheetApp.newConditionalFormatRule().whenTextContains("✅ AVAILABLE").setBackground("#d4edda").setFontColor("#155724").setRanges([rangeJadwal]).build();
  const ruleMerah = SpreadsheetApp.newConditionalFormatRule().whenTextContains("🔴 BOOKED").setBackground("#f8d7da").setFontColor("#721c24").setRanges([rangeJadwal]).build();
  
  // Warna Utilitas (Heatmap: Makin tinggi % makin merah)
  const rangePersen = dash.getRange("L10:L35");
  const ruleBeban = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxColor("#FF5555")
    .setGradientMinColor("#FFFFFF")
    .setRanges([rangePersen])
    .build();

  dash.setConditionalFormatRules([ruleHijau, ruleMerah, ruleBeban]);

  // Pengaturan Lebar Kolom
  dash.setColumnWidth(1, 180); // Kolom A
  dash.setColumnWidths(2, 7, 110); // Kolom B-H
  dash.setColumnWidth(9, 30); // Kolom I (Pemisah)
  dash.setColumnWidth(10, 180); // Kolom J
  dash.setColumnWidths(11, 2, 90); // Kolom K-L

  // Border
  dash.getRange("A8:H35").setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange("J8:L35").setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  
  dash.getRange("A9:L35").setVerticalAlignment("middle").setHorizontalAlignment("center");
  
  try { dash.setHideGridlines(true); } catch(e) { }

  Browser.msgBox("🚀 DASHBOARD MANAGER GABUNGAN SELESAI!");
}

/************************************************************
 * MENU KUSTOM: 
 ************************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚗 Update Sistem') 
      .addItem('🔄 Update Dropdown Mobil', 'filterMobilTersedia')
.addItem('🚀 Buat Dashboard Otomatis', 'buatDashboardOtomatis')
      .addSeparator()
      .addItem('⚙️ Cek Otorisasi WA', 'cekOtorisasiManual')
      .addToUi();
}

/** * Fungsi pembantu untuk memicu permintaan izin jika 
 * tombol menu di Sheets tidak jalan 
 */
function cekOtorisasiManual() {
  Browser.msgBox("Otorisasi Berhasil", "Sistem sudah memiliki izin untuk menjalankan script.", Browser.Buttons.OK);
}

// ==========================================
// DRIVER ROLE
// ==========================================
function updateFormDropdown() {
  // 1. ID Google Form abang (ambil dari URL saat edit Form)
  const formId = "10_a-XaPZ_U4Ql-Z-R03Z239Yi6J3heP0a_W4MckC1OQ"; 
  const form = FormApp.openById(formId);
  
  // 2. Nama Sheet dan Range data mobil
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMaster = ss.getSheetByName("Master_Armada"); 
  const data = sheetMaster.getRange("A2:A" + sheetMaster.getLastRow()).getValues();
  
  // 3. Ubah data jadi list satu baris
  const listMobil = data.map(row => row[0]).filter(item => item !== "");
  
  // 4. Cari pertanyaan Dropdown di Form (berdasarkan judulnya)
  const items = form.getItems();
  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === "Pilih Unit Mobil") { // Sesuaikan dengan judul di Form
      items[i].asListItem().setChoiceValues(listMobil);
      break;
    }
  }
}

/*************** */
function onFormSubmitDriver(e) {
  if (!e || !e.values) {
    Logger.log("Gagal: Tidak ada data form yang masuk");
    return;
  }

  const responses = e.values;
  const platNomor = responses[1].toString().trim(); 
  const kmAkhir = responses[2]; 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const armadaSheet = ss.getSheetByName('Master_Armada');
  const data = armadaSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === platNomor) {
      armadaSheet.getRange(i + 1, 5).setValue("Available");
      armadaSheet.getRange(i + 1, 6).setValue(kmAkhir);
      
      SpreadsheetApp.flush(); // PAKSA Google Sheets menulis data sekarang juga
      Logger.log("Berhasil Update Status untuk: " + platNomor);
      
      filterMobilTersedia(); 
      return;
    }
  }
}


/*****************
* reset master armada mobil yang dipakai
*****************/
function autoResetStatusMobil() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestSheet = ss.getSheetByName('Request Mobil');
  const armadaSheet = ss.getSheetByName('Master_Armada');
  const data = requestSheet.getDataRange().getValues();
  const sekarang = new Date();

  for (let i = 1; i < data.length; i++) {
    const tglSelesai = new Date(data[i][COL_TGL_KEPULANGAN - 1]);
    const statusFinal = data[i][COL_STATUS_FINAL - 1];
    const unitMobil = data[i][COL_PILIH_KENDARAAN - 1].toString();

    if (statusFinal === "Approved" && (sekarang.getTime() - tglSelesai.getTime()) > (12 * 60 * 60 * 1000)) {
       updateStatusFisik(unitMobil, "Available");
    }
  }
}

function updateStatusFisik(namaUnit, statusBaru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const armadaSheet = ss.getSheetByName('Master_Armada');
  const data = armadaSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (namaUnit.indexOf(data[i][0]) !== -1) {
      armadaSheet.getRange(i + 1, 5).setValue(statusBaru);
      break;
    }
  }
}


/*****************
 * 
 * 
 */
function setupSemuaTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  
  // 1. Trigger saat ada pengajuan baru (Pemohon)
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create();
    
  // 2. Trigger Update Dropdown Mobil tiap 15 menit
  ScriptApp.newTrigger('filterMobilTersedia')
    .timeBased()
    .everyMinutes(15)
    .create();

  // 3. Trigger Pembersih Otomatis tiap 4 jam
  ScriptApp.newTrigger('autoResetStatusMobil')
    .timeBased()
    .everyHours(4)
    .create();
    
  Browser.msgBox("✅ Berhasil! Semua trigger (termasuk Pembersih Otomatis) telah dipasang.");
}
