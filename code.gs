const SPREADSHEET_ID = '1hg3CynBrqFci3kEg7c611ht4jtIDO9Ocyh2RPc1ixE8';
const SHEET_NAME     = 'Request Mobil';

const PIC_EMAIL = 'wawazer@gmail.com';                  
const CC_EMAIL  = 'muhammad.wawazer@pelindo.co.id';     

const WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbxemjnzYVho13q1kY27ptfAD8F9QSbiD_z4XqSjWkVU0dPMT3BX-V7LXKoM1KCcSbrz3g/exec';

const COL_NO             = 1;   
const COL_DASAR_SURAT    = 2;   
const COL_TGL_BERANGKAT  = 3;   
const COL_TGL_KEPULANGAN = 4;   
const COL_NO_KENDARAAN   = 5;   
const COL_EMAIL_PENGAJU  = 6;   
const COL_PIC_PENDAMPING = 7;   
const COL_DRIVER         = 8;   
const COL_DAFTAR_TAMU    = 9;   
const COL_TUJUAN         = 10;  
const COL_HOTEL          = 11;  
const COL_JUMLAH_HARI    = 12;  
const COL_BIAYA          = 13;  
const COL_KETERANGAN     = 14;  
const COL_DIKETAHUI      = 15;  
const COL_STATUS         = 16;  



function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Request Mobil')
    .addItem('Kirim email ke PIC (Pastikan baris sudah benar)', 'sendEmailForSelectedRow')
    .addToUi();
}


function sendEmailForRow(row) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error('Sheet dengan nama "' + SHEET_NAME + '" tidak ditemukan.');
  }

  const lastCol = sheet.getLastColumn();
  const data    = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const no              = data[COL_NO - 1];
  const dasarSurat      = data[COL_DASAR_SURAT - 1];
  const tglBerangkat    = data[COL_TGL_BERANGKAT - 1];
  const tglKepulangan   = data[COL_TGL_KEPULANGAN - 1];
  const nomorKendaraan  = data[COL_NO_KENDARAAN - 1];
  const emailPengaju    = data[COL_EMAIL_PENGAJU - 1];
  const picPendamping   = data[COL_PIC_PENDAMPING - 1];
  const driver          = data[COL_DRIVER - 1];
  const daftarTamu      = data[COL_DAFTAR_TAMU - 1];
  const tujuanPerjalanan= data[COL_TUJUAN - 1];
  const hotelMenginap   = data[COL_HOTEL - 1];
  const jumlahHari      = data[COL_JUMLAH_HARI - 1];
  const biayaPelayanan  = data[COL_BIAYA - 1];
  const keterangan      = data[COL_KETERANGAN - 1];
  const diketahuiOleh   = data[COL_DIKETAHUI - 1];
  const status          = data[COL_STATUS - 1];

  // kalau sudah ada status, jangan kirim ulang ke PIC
  if (status && status !== '') {
    console.log('Row ' + row + ' sudah punya status: ' + status + ', email ke PIC tidak dikirim ulang.');
    return;
  }

  if (!WEBAPP_URL || WEBAPP_URL === '') {
    throw new Error('WEBAPP_URL belum diisi.');
  }

  const approveUrl = `${WEBAPP_URL}?action=approve&row=${row}`;
  const rejectUrl  = `${WEBAPP_URL}?action=reject&row=${row}`;

  const subject = `Request Peminjaman Mobil dari ${picPendamping|| ''}`;

  const htmlBody = `
    <p>Ada request Peminjaman Mobil :</p>
    <table border="0" cellpadding="4" cellspacing="0">
      <tr><td><b>No</b></td><td>: ${no || ''}</td></tr>
      <tr><td><b>Dasar Surat</b></td><td>: ${dasarSurat || ''}</td></tr>
      <tr><td><b>Tanggal Berangkat</b></td><td>: ${tglBerangkat || ''}</td></tr>
      <tr><td><b>Tanggal Kepulangan</b></td><td>: ${tglKepulangan || ''}</td></tr>
      <tr><td><b>Nomor Kendaraan</b></td><td>: ${nomorKendaraan || ''}</td></tr>
      <tr><td><b>Email Pengaju</b></td><td>: ${emailPengaju || ''}</td></tr>
      <tr><td><b>PIC Pendamping</b></td><td>: ${picPendamping || ''}</td></tr>
      <tr><td><b>Driver</b></td><td>: ${driver || ''}</td></tr>
      <tr><td><b>Daftar Tamu yang Dilayani</b></td><td>: ${daftarTamu || ''}</td></tr>
      <tr><td><b>Tujuan Perjalanan</b></td><td>: ${tujuanPerjalanan || ''}</td></tr>
      <tr><td><b>Hotel Menginap</b></td><td>: ${hotelMenginap || ''}</td></tr>
      <tr><td><b>Jumlah Hari</b></td><td>: ${jumlahHari || ''}</td></tr>
      <tr><td><b>Biaya Pelayanan</b></td><td>: ${biayaPelayanan || ''}</td></tr>
      <tr><td><b>Keterangan</b></td><td>: ${keterangan || ''}</td></tr>
      <tr><td><b>Diketahui oleh (Koordinator Kendaraan)</b></td><td>: ${diketahuiOleh || ''}</td></tr>
    </table>
    <p>
      <a href="${approveUrl}">✅ <b>APPROVE</b></a>&nbsp;&nbsp;&nbsp;
      <a href="${rejectUrl}">❌ <b>REJECT</b></a>
    </p>
  `;

  const mailOptions = {
    to: PIC_EMAIL,
    subject: subject,
    htmlBody: htmlBody
  };

  if (CC_EMAIL && CC_EMAIL !== '') {
    mailOptions.cc = CC_EMAIL;
  }

  MailApp.sendEmail(mailOptions);

  console.log('Email request dikirim ke PIC untuk row ' + row);
}

function sendEmailForSelectedRow() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet "' + SHEET_NAME + '" tidak ditemukan.');
    return;
  }

  const range = sheet.getActiveRange();
  const row   = range.getRow();

  if (row === 1) {
    SpreadsheetApp.getUi().alert('Pilih baris data, bukan header.');
    return;
  }

  try {
    sendEmailForRow(row);
    SpreadsheetApp.getUi().alert('Email request telah dikirim ke PIC untuk baris ke-' + row + '.');
  } catch (err) {
    SpreadsheetApp.getUi().alert('Gagal kirim email: ' + err.message);
  }
}

function kirimEmailKePemohon(row, newStatus) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastCol = sheet.getLastColumn();
  const data    = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const no              = data[COL_NO - 1];
  const dasarSurat      = data[COL_DASAR_SURAT - 1];
  const tglBerangkat    = data[COL_TGL_BERANGKAT - 1];
  const tglKepulangan   = data[COL_TGL_KEPULANGAN - 1];
  const nomorKendaraan  = data[COL_NO_KENDARAAN - 1];
  const emailPengaju    = data[COL_EMAIL_PENGAJU - 1];
  const picPendamping   = data[COL_PIC_PENDAMPING - 1];
  const driver          = data[COL_DRIVER - 1];
  const tujuanPerjalanan= data[COL_TUJUAN - 1];

  if (!emailPengaju || emailPengaju === '') {
    console.log('Row ' + row + ' tidak punya Email Pengaju, email ke pemohon tidak dikirim.');
    return;
  }

  const subjectPemohon = `Status Permohonan Mobil No SPPD ${dasarSurat|| ''}: ${newStatus}`;
  const htmlPemohon = `
    <p>Yth. Pemohon,</p>
    <p>Permohonan penggunaan mobil Anda telah diproses dengan status: <b>${newStatus}</b>.</p>
    <p>Detail permohonan:</p>
    <table border="0" cellpadding="4" cellspacing="0">
      <tr><td><b>Dasar Surat</b></td><td>: ${dasarSurat || ''}</td></tr>
      <tr><td><b>Tanggal Berangkat</b></td><td>: ${tglBerangkat || ''}</td></tr>
      <tr><td><b>Tanggal Kepulangan</b></td><td>: ${tglKepulangan || ''}</td></tr>
      <tr><td><b>Nomor Kendaraan</b></td><td>: ${nomorKendaraan || '-'}</td></tr>
      <tr><td><b>PIC Pendamping</b></td><td>: ${picPendamping || '-'}</td></tr>
      <tr><td><b>Driver</b></td><td>: ${driver || '-'}</td></tr>
      <tr><td><b>Tujuan Perjalanan</b></td><td>: ${tujuanPerjalanan || ''}</td></tr>
    </table>
    <p>Terima kasih.</p>
  `;

  MailApp.sendEmail({
    to: emailPengaju,
    subject: subjectPemohon,
    htmlBody: htmlPemohon
  });

  console.log('Email status dikirim ke pemohon untuk row ' + row);
}

function doGet(e) {
  const action = e.parameter.action;
  const rowStr = e.parameter.row;

  if (!action || !rowStr) {
    return ContentService.createTextOutput('Parameter tidak lengkap.');
  }

  const row = parseInt(rowStr, 10);
  if (isNaN(row)) {
    return ContentService.createTextOutput('Row tidak valid.');
  }

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    return ContentService.createTextOutput('Sheet "' + SHEET_NAME + '" tidak ditemukan.');
  }

  const newStatus =
    action === 'approve' ? 'Approved' :
    action === 'reject'  ? 'Rejected' : 'Unknown';

  sheet.getRange(row, COL_STATUS).setValue(newStatus);

  kirimEmailKePemohon(row, newStatus);

  const html = HtmlService.createHtmlOutput(
    '<html><body style="font-family:Arial;padding:20px;">' +
    '<h3>Request mobil telah di-' + newStatus + '.</h3>' +
    '<p>Pemohon akan menerima email notifikasi (jika alamat email tercatat di kolom Email Pengaju).</p>' +
    '</body></html>'
  );

  return html;
}

function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row === 1 || col !== COL_STATUS) return;

  const newValue = (e.value || '').toString().trim();

  if (newValue === 'Approved' || newValue === 'Rejected') {
    kirimEmailKePemohon(row, newValue);
  }
}

function onFormSubmit(e) {
  const formSheet = e.range.getSheet();
  if (formSheet.getName() !== 'Request Mobil') return; // sesuaikan nama sheet respons

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const reqSheet = ss.getSheetByName(SHEET_NAME);

  const values = e.values; 
  // values[0] = Timestamp
  const dasarSurat      = values[1];
  const tglBerangkat    = values[2];
  const tglKepulangan   = values[3];
  const nomorKendaraan  = values[4];
  const emailPengaju    = values[5];
  const picPendamping   = values[6];
  const driver          = values[7];
  const daftarTamu      = values[8];
  const tujuanPerjalanan= values[9];
  const hotelMenginap   = values[10];
  const jumlahHari      = values[11];
  const biayaPelayanan  = values[12];
  const keterangan      = values[13];
  const diketahuiOleh   = values[14];

  const lastRowReq = reqSheet.getLastRow();
  let nextNo = 1;
  if (lastRowReq > 1) { 
    const lastNo = reqSheet.getRange(lastRowReq, COL_NO).getValue();
    nextNo = (Number(lastNo) || 0) + 1;
  }

  const newRow = [
    nextNo,            
    dasarSurat,        
    tglBerangkat,      
    tglKepulangan,     
    nomorKendaraan,    
    emailPengaju,      
    picPendamping,     
    driver,            
    daftarTamu,        
    tujuanPerjalanan,  
    hotelMenginap,     
    jumlahHari,        
    biayaPelayanan,    
    keterangan,        
    diketahuiOleh,     
    ''                 
  ];

  reqSheet.appendRow(newRow);

  const newRowIndex = reqSheet.getLastRow();
  sendEmailForRow(newRowIndex);
}

