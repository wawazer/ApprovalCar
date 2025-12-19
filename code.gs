const SPREADSHEET_ID = '1hg3CynBrqFci3kEg7c611ht4jtIDO9Ocyh2RPc1ixE8';
const SHEET_NAME     = 'Request Mobil';

// const PIC_EMAIL = 'WAWAZER@GMAIL.COM';
// const CC_EMAIL  = ' listio.margianto@pelindo.co.id, Reza.abimoko@pelindo.co.id';
const CC_EMAIL  = 'listio.margianto@pelindo.co.id, Reza.abimoko@pelindo.co.id, wahyu.ekoyulianto@pelindo.co.id';
// const LV1_APPROVER_EMAILS = 'iwan.sulistiono@pelindo.co.id'; 
const LV1_APPROVER_EMAILS = 'iwan.sulistiono@pelindo.co.id'; 
// const LV2_APPROVER_EMAILS = 'delumintu@gmail.com';
const KOOR_WA_NUMBER = '6285331946877'; 
// const WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbxICLEfVFaYQdZydstk4kwmZHipDNvTSxB2xj1DwATX9wHAmCCW8FZQf9SiwxiEgtlOnQ/exec'; langsung ke approval emai
const WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbygGiOvv66mnFf7F36Bp-_BXy-N0HuNWD0yXCIcdhYPBoBLgK3tzI2EjuJ9wNZ9fVVY/exec';
                    


const COL_TIMESTAMP      = 1;  
const COL_DASAR_SURAT    = 2;  
const COL_TGL_BERANGKAT  = 3;  
const COL_TGL_KEPULANGAN = 4;  
const COL_NO_KENDARAAN   = 5;  
const COL_EMAIL_PENGAJU  = 6;  
const COL_UNIT_KERJA     = 7;  
const COL_DRIVER         = 8;  
const COL_DAFTAR_TAMU    = 9;  
const COL_TUJUAN         = 10; 
const COL_HOTEL          = 11; 
const COL_JUMLAH_HARI    = 12; 
const COL_BIAYA          = 13; 
const COL_NAMA_PIC       = 14; 
const COL_STATUS_LV1   = 15; 
const COL_STATUS_LV2   = 16; 
const COL_STATUS_FINAL = 17; 


function generateWaLink(phoneNumber, message) {
  const encodedMessage = encodeURIComponent(message);
  return `https://wa.me/${phoneNumber}?text=${encodedMessage}`;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Request Mobil')
    .addItem('Kirim email ke PIC (Pastikan baris sudah benar)', 'sendEmailForSelectedRow')
    .addToUi();
}

/************************************************************/
function sendEmailForRow(row) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error('Sheet dengan nama "' + SHEET_NAME + '" tidak ditemukan.');
  }

  const lastCol = sheet.getLastColumn();
  const data    = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const timestamp        = data[COL_TIMESTAMP      - 1];
  const dasarSurat       = data[COL_DASAR_SURAT    - 1];
  const tglBerangkat     = data[COL_TGL_BERANGKAT  - 1];
  const tglKepulangan    = data[COL_TGL_KEPULANGAN - 1];
  const nomorKendaraan   = data[COL_NO_KENDARAAN   - 1];
  const emailPengaju     = data[COL_EMAIL_PENGAJU  - 1];
  const unitKerja        = data[COL_UNIT_KERJA     - 1];
  const driver           = data[COL_DRIVER         - 1];
  const daftarTamu       = data[COL_DAFTAR_TAMU    - 1];
  const tujuanPerjalanan = data[COL_TUJUAN         - 1];
  const hotelMenginap    = data[COL_HOTEL          - 1];
  const jumlahHari       = data[COL_JUMLAH_HARI    - 1];
  const biayaPelayanan   = data[COL_BIAYA          - 1];
  const namaPIC          = data[COL_NAMA_PIC       - 1];
  const status           = data[COL_STATUS_FINAL   - 1];


  // kalau sudah ada status, jangan kirim ulang ke PIC
  if (status && status !== '') {
    console.log('Row ' + row + ' sudah punya status: ' + status + ', email ke PIC tidak dikirim ulang.');
    return;
  }

  if (!WEBAPP_URL || WEBAPP_URL === '') {
    throw new Error('WEBAPP_URL belum diisi.');
  }
const approveUrl = `${WEBAPP_URL}?action=approve&row=${row}&level=1`;
const rejectUrl  = `${WEBAPP_URL}?action=reject&row=${row}&level=1`;

  const subject = `Request Peminjaman Mobil dari ${namaPIC || ''}`;

  const htmlBody = `
    <p>Ada request Peminjaman Mobil :</p>
    <table border="0" cellpadding="4" cellspacing="0">
      <tr><td><b>Timestamp</b></td><td>: ${timestamp || ''}</td></tr>
      <tr><td><b>Nama PIC</b></td><td>: ${namaPIC || ''}</td></tr>
      <tr><td><b>Unit Kerja</b></td><td>: ${unitKerja || ''}</td></tr>
      <tr><td><b>Dasar Surat</b></td><td>: ${dasarSurat || ''}</td></tr>
      <tr><td><b>Tanggal Berangkat</b></td><td>: ${tglBerangkat || ''}</td></tr>
      <tr><td><b>Tanggal Kepulangan</b></td><td>: ${tglKepulangan || ''}</td></tr>
      <tr><td><b>Nomor Kendaraan</b></td><td>: ${nomorKendaraan || ''}</td></tr>
      <tr><td><b>Email Pengaju</b></td><td>: ${emailPengaju || ''}</td></tr>
      <tr><td><b>Driver</b></td><td>: ${driver || ''}</td></tr>
      <tr><td><b>Daftar Tamu yang Dilayani</b></td><td>: ${daftarTamu || ''}</td></tr>
      <tr><td><b>Tujuan Perjalanan</b></td><td>: ${tujuanPerjalanan || ''}</td></tr>
      <tr><td><b>Hotel Menginap</b></td><td>: ${hotelMenginap || ''}</td></tr>
      <tr><td><b>Jumlah Hari</b></td><td>: ${jumlahHari || ''}</td></tr>
      <tr><td><b>Biaya Pelayanan</b></td><td>: ${biayaPelayanan || ''}</td></tr>
    </table>
    <p>
      <a href="${approveUrl}">✅ <b>APPROVE</b></a>&nbsp;&nbsp;&nbsp;
      <a href="${rejectUrl}">❌ <b>REJECT</b></a>
    </p>
  `;

const mailOptions = {
  to: LV1_APPROVER_EMAILS,  
  subject: subject,
  htmlBody: htmlBody
};

if (CC_EMAIL && CC_EMAIL !== '') {
  mailOptions.cc = CC_EMAIL; 
}

  MailApp.sendEmail(mailOptions);

  console.log('Email request dikirim ke PIC untuk row ' + row);
}

/************************************************************
 *  KIRIM EMAIL 
 ************************************************************/

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

/************************************************************
 * EMAIL STATUS KE PEMOHON
 ************************************************************/

function kirimEmailKePemohon(row, newStatus) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastCol = sheet.getLastColumn();
  const data    = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const timestamp        = data[COL_TIMESTAMP      - 1];
  const dasarSurat       = data[COL_DASAR_SURAT    - 1];
  const tglBerangkat     = data[COL_TGL_BERANGKAT  - 1];
  const tglKepulangan    = data[COL_TGL_KEPULANGAN - 1];
  const nomorKendaraan   = data[COL_NO_KENDARAAN   - 1];
  const emailPengaju     = data[COL_EMAIL_PENGAJU  - 1];
  const unitKerja        = data[COL_UNIT_KERJA     - 1];
  const driver           = data[COL_DRIVER         - 1];
  const tujuanPerjalanan = data[COL_TUJUAN         - 1];
  const namaPIC          = data[COL_NAMA_PIC       - 1];

  if (!emailPengaju || emailPengaju === '') {
    console.log('Row ' + row + ' tidak punya Email Pengaju, email ke pemohon tidak dikirim.');
    return;
  }

  const subjectPemohon = `Status Permohonan Mobil (${newStatus})`;
  const htmlPemohon = `
    <p>Yth. ${namaPIC || 'Pemohon'},</p>
    <p>Permohonan penggunaan mobil Anda telah diproses dengan status: <b>${newStatus}</b>.</p>
    <p>Detail permohonan:</p>
    <table border="0" cellpadding="4" cellspacing="0">
      <tr><td><b>Timestamp</b></td><td>: ${timestamp || ''}</td></tr>
      <tr><td><b>Dasar Surat</b></td><td>: ${dasarSurat || ''}</td></tr>
      <tr><td><b>Tanggal Berangkat</b></td><td>: ${tglBerangkat || ''}</td></tr>
      <tr><td><b>Tanggal Kepulangan</b></td><td>: ${tglKepulangan || ''}</td></tr>
      <tr><td><b>Nomor Kendaraan</b></td><td>: ${nomorKendaraan || '-'}</td></tr>
      <tr><td><b>Unit Kerja</b></td><td>: ${unitKerja || '-'}</td></tr>
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

/************************************************************
 *HANDLE APPROVE / REJECT
 ************************************************************/

function doGet(e) {
  const action = e.parameter.action;
  const rowStr = e.parameter.row;
  const level  = parseInt(e.parameter.level || '1', 10); // default 1 kalau tidak ada

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

  // --- Baca data baris (dipakai untuk email notifikasi) ---
  const lastCol = sheet.getLastColumn();
  const data    = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const emailPengaju     = data[COL_EMAIL_PENGAJU  - 1];
  const namaPIC          = data[COL_NAMA_PIC       - 1];
  const dasarSurat       = data[COL_DASAR_SURAT    - 1];
  const tglBerangkat     = data[COL_TGL_BERANGKAT  - 1];
  const tglKepulangan    = data[COL_TGL_KEPULANGAN - 1];
  const unitKerja        = data[COL_UNIT_KERJA     - 1];

  // =======================
  // LEVEL 1 
  // =======================
  if (level === 1) {
    sheet.getRange(row, COL_STATUS_LV1).setValue(newStatus);

    if (newStatus === 'Rejected') {
      sheet.getRange(row, COL_STATUS_FINAL).setValue('Rejected');
      if (emailPengaju) {
        kirimEmailKePemohon(row, 'Rejected');
      }
    }
   if (newStatus === 'Approved') {

  const approveUrlLv2 = `${WEBAPP_URL}?action=approve&row=${row}&level=2`;
  const rejectUrlLv2  = `${WEBAPP_URL}?action=reject&row=${row}&level=2`;

  const subjectLv2 = `Approval LV2 Permohonan Mobil dari ${namaPIC || ''}`;
  const htmlLv2 = `
    <p>Yth. Approver Level 2,</p>
    <p>Mohon approval permohonan mobil berikut (telah di-approve Koordinator):</p>
    <ul>
      <li>Nama PIC: ${namaPIC || '-'}</li>
      <li>Unit Kerja: ${unitKerja || '-'}</li>
      <li>Dasar Surat: ${dasarSurat || '-'}</li>
      <li>Tanggal: ${tglBerangkat || '-'} s/d ${tglKepulangan || '-'}</li>
    </ul>
    <p>
      <a href="${approveUrlLv2}">✅ APPROVE </a>&nbsp;&nbsp;&nbsp;
      <a href="${rejectUrlLv2}">❌ REJECT </a>
    </p>
  `;

  MailApp.sendEmail({
    to: LV2_APPROVER_EMAILS,   
    subject: subjectLv2,
    htmlBody: htmlLv2
  });
}


  // =======================
  // LEVEL 2 
  // =======================
  } else if (level === 2) {
    sheet.getRange(row, COL_STATUS_LV2).setValue(newStatus);

    sheet.getRange(row, COL_STATUS_FINAL).setValue(newStatus);

    if (emailPengaju) {
      kirimEmailKePemohon(row, newStatus);
    }
  }
  const html = HtmlService.createHtmlOutput(
    '<html><body style="font-family:Arial;padding:20px;">' +
    '<h3>Approval level ' + level + ' telah di-' + newStatus + '.</h3>' +
    '<p>Anda dapat menutup halaman ini.</p>' +
    '</body></html>'
  );

  return html;
}


// Ambil semua mobil dari sheet "Daftar Mobil"
function getAllCars() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Daftar Mobil');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(r => r[0])
    .filter(v => v && v !== '');
  return data;
}

// Ambil mobil yang AVAILABLE 
function getAvailableCars(startIso, endIso) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const reqSheet = ss.getSheetByName(SHEET_NAME);

  const allCars = getAllCars();
  if (!startIso || !endIso) {
    return allCars;
  }

  const start = new Date(startIso);
  const end   = new Date(endIso);

  const used = new Set();

  const data = reqSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const car         = row[COL_NO_KENDARAAN   - 1];
    const statusFinal = row[COL_STATUS_FINAL   - 1];
    const existStart  = row[COL_TGL_BERANGKAT  - 1];
    const existEnd    = row[COL_TGL_KEPULANGAN - 1];

    if (!car) continue;
    if (statusFinal !== 'Approved') continue;
    if (!(existStart instanceof Date) || !(existEnd instanceof Date)) continue;

    // cek overlap: existingStart < end && existingEnd > start
    if (existStart < end && existEnd > start) {
      used.add(String(car));
    }
  }

  const available = allCars.filter(c => !used.has(String(c)));
  return available;
}

// Submit data dari Web App ke sheet & kirim email approval LV1
function submitRequest(form) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  const now = new Date();

  const rowValues = [
    now,                            
    form.dasarSurat || '',          
    new Date(form.tglBerangkat),    
    new Date(form.tglKepulangan),   
    form.noKendaraan || '',         
    form.emailPengaju || '',        
    form.unitKerja || '',           
    form.driver || '',              
    form.daftarTamu || '',          
    form.tujuan || '',              
    form.hotel || '',               
    form.jumlahHari || '',          
    form.biaya || '',               
    form.namaPIC || '',             
    '',                             
    '',                             
    ''                              
  ];

  sheet.appendRow(rowValues);
  const newRowIndex = sheet.getLastRow();

  // kirim email ke approver level 1 
  sendEmailForRow(newRowIndex);

  // kirim email WA link ke pemohon seperti di onFormSubmit 
  const dasarSurat       = form.dasarSurat || '';
  const tglBerangkat     = form.tglBerangkat || '';
  const tglKepulangan    = form.tglKepulangan || '';
  const nomorKendaraan   = form.noKendaraan || '';
  const emailPengaju     = form.emailPengaju || '';
  const unitKerja        = form.unitKerja || '';
  const namaPIC          = form.namaPIC || '';

  const waMessage =
    'Yth. Koordinator,\n\n' +
    'Saya ' + (namaPIC || '-') + ' dari ' + (unitKerja || '-') + ' mengajukan permohonan penggunaan mobil.\n\n' +
    'Dasar Surat : ' + (dasarSurat || '-') + '\n' +
    'Tanggal     : ' + (tglBerangkat || '-') + ' s/d ' + (tglKepulangan || '-') + '\n' +
    'Nomor Kendaraan (jika ada) : ' + (nomorKendaraan || '-') + '\n\n' +
    'Mohon konfirmasi persetujuan. Terima kasih.';

  const waLink = generateWaLink(KOOR_WA_NUMBER, waMessage);

  if (emailPengaju) {
    const subjectPemohon = 'Link WhatsApp ke Koordinator - Permohonan Mobil';
    const htmlPemohon = `
      <p>Yth. ${namaPIC || 'Pemohon'},</p>
      <p>Terima kasih sudah mengisi form permohonan penggunaan mobil.</p>
      <p>Jika ingin <b>mengirim reminder ke Koordinator via WhatsApp</b>, silakan klik link berikut:</p>
      <p><a href="${waLink}" target="_blank">📲 Kirim WhatsApp ke Koordinator</a></p>
    `;

    MailApp.sendEmail({
      to: emailPengaju,
      subject: subjectPemohon,
      htmlBody: htmlPemohon
    });
  }

  return 'Permohonan berhasil disimpan. Email approval sudah dikirim ke Koordinator.';
}


function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  // hanya trigger kalau yang diubah adalah kolom Status Final
  if (row === 1 || col !== COL_STATUS_FINAL) return;

  const newValue = (e.value || '').toString().trim();

  if (newValue === 'Approved' || newValue === 'Rejected') {
    kirimEmailKePemohon(row, newValue);
  }
}


/************************************************************
 * TRIGGER: FORM SUBMIT → EMAIL PIC + WA 
 ************************************************************/

function onFormSubmit(e) {
  const sheet = e.range.getSheet();

  if (sheet.getName() !== SHEET_NAME) return;

  const row = e.range.getRow();
  const values = e.values;

  const dasarSurat       = values[COL_DASAR_SURAT    - 1];
  const tglBerangkat     = values[COL_TGL_BERANGKAT  - 1];
  const tglKepulangan    = values[COL_TGL_KEPULANGAN - 1];
  const nomorKendaraan   = values[COL_NO_KENDARAAN   - 1];
  const emailPengaju     = values[COL_EMAIL_PENGAJU  - 1];
  const unitKerja        = values[COL_UNIT_KERJA     - 1];
  const driver           = values[COL_DRIVER         - 1];
  const daftarTamu       = values[COL_DAFTAR_TAMU    - 1];
  const tujuanPerjalanan = values[COL_TUJUAN         - 1];
  const hotelMenginap    = values[COL_HOTEL          - 1];
  const jumlahHari       = values[COL_JUMLAH_HARI    - 1];
  const biayaPelayanan   = values[COL_BIAYA          - 1];
  const namaPIC          = values[COL_NAMA_PIC       - 1];

  sendEmailForRow(row);

  const waMessage =
    'Yth. Koordinator,\n\n' +
    'Saya ' + (namaPIC || '-') + ' dari ' + (unitKerja || '-') + ' mengajukan permohonan penggunaan mobil.\n\n' +
    'Dasar Surat : ' + (dasarSurat || '-') + '\n' +
    'Tanggal     : ' + (tglBerangkat || '-') + ' s/d ' + (tglKepulangan || '-') + '\n' +
    'Tujuan      : ' + (tujuanPerjalanan || '-') + '\n' +
    'Nomor Kendaraan (jika ada) : ' + (nomorKendaraan || '-') + '\n\n' +
    'Mohon konfirmasi persetujuan. Terima kasih.';

  const waLink = generateWaLink(KOOR_WA_NUMBER, waMessage);

  if (emailPengaju && emailPengaju !== '') {
    const subjectPemohon = 'Link WhatsApp ke Koordinator - Permohonan Mobil';
    const htmlPemohon = `
      <p>Yth. ${namaPIC || 'Pemohon'},</p>
      <p>Terima kasih sudah mengisi form permohonan penggunaan mobil.</p>
      <p>Jika ingin <b>mengirim reminder ke Koordinator via WhatsApp</b>, silakan klik link berikut:</p>
      <p><a href="${waLink}" target="_blank">📲 Kirim WhatsApp ke Koordinator</a></p>
      <p>Pesan WhatsApp akan otomatis terisi, Anda hanya perlu menekan tombol <b>Send</b> di aplikasi WhatsApp.</p>
    `;

    MailApp.sendEmail({
      to: emailPengaju,
      subject: subjectPemohon,
      htmlBody: htmlPemohon
    });
  }
}
