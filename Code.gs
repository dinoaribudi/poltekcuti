// =====================================================================
// APLIKASI CUTI KARYAWAN — Google Apps Script Backend
// Spreadsheet Utama : 1cvHSkD0b9HJ7e-wao1TJyY6xkHm_XJP89kuN6Hlartc
// Spreadsheet PHL   : 1RagbWc-jel6eCEexNWZrniCHHyaymNrdeid8XraXMns
// =====================================================================

const SPREADSHEET_ID     = '1cvHSkD0b9HJ7e-wao1TJyY6xkHm_XJP89kuN6Hlartc';
const SPREADSHEET_PHL_ID = '1RagbWc-jel6eCEexNWZrniCHHyaymNrdeid8XraXMns';
const SHEET_USERS        = 'Users';
const SHEET_QUOTA        = 'QuotaCuti';
const SHEET_TRANSAKSI    = 'Transaksi';
const SHEET_PHL          = 'PHL';

// =====================================================================
// KONFIGURASI EMAIL
// =====================================================================
const EMAIL_SENDER      = 'hrd.poltek@polteksimasberau.ac.id';
const EMAIL_APP_PASSWORD = 'tufg fkmn okvb vpae';
const APP_NAME          = 'Away — Sistem Cuti';

let _ss    = null;
let _ssPHL = null;

function getSS() {
  if (!_ss) _ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ss;
}

function getSsPHL() {
  if (!_ssPHL) _ssPHL = SpreadsheetApp.openById(SPREADSHEET_PHL_ID);
  return _ssPHL;
}

function getSheet(name) {
  const ss = getSS();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheet(sheet, name);
  }
  return sheet;
}

function getPHLSheet() {
  return getSsPHL().getSheetByName(SHEET_PHL);
}

function initSheet(sheet, name) {
  if (name === SHEET_QUOTA) {
    sheet.getRange(1, 1, 1, 6).setValues([['id','user_id','jenis_cuti','quota','terpakai','sisa']]);
  } else if (name === SHEET_TRANSAKSI) {
    sheet.getRange(1, 1, 1, 12).setValues([['id','user_id','nama_karyawan','jenis_cuti','tgl_mulai','tgl_selesai','jumlah_hari','alasan','status','tgl_pengajuan','tgl_approval','ref_phl_id']]);
  }
}

function fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) return val.toISOString();
  return String(val);
}

// Format tanggal ke YYYY-MM-DD menggunakan timezone WIB (UTC+7)
// Digunakan untuk kolom start_date & end_date agar konsisten di semua device
function fmtDateWIB(val) {
  if (!val) return '';
  const d = val instanceof Date ? val : new Date(val);
  if (isNaN(d)) return String(val);
  // Offset WIB = UTC+7 = 7*60 menit
  const wib = new Date(d.getTime() + 7 * 60 * 60 * 1000);
  const y  = wib.getUTCFullYear();
  const m  = String(wib.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(wib.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${dd}`;
}

// =====================================================================
// ENTRY POINT
// =====================================================================
function doGet(e) {
  _ss = null; _ssPHL = null;
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const p      = e.parameter || {};
    const action = p.action || '';
    const data   = p.data ? JSON.parse(decodeURIComponent(p.data)) : {};
    data.action  = action;
    let result;
    switch (action) {
      case 'login':               result = login(data); break;
      case 'getUsers':            result = getUsers(); break;
      case 'getQuota':            result = getQuota(data); break;
      case 'getAllQuota':         result = getAllQuota(); break;
      case 'getTransaksi':        result = getTransaksi(data); break;
      case 'getAllTransaksi':     result = getAllTransaksi(); break;
      case 'getPendingForAtasan': result = getPendingForAtasan(data); break;
      case 'ajukanCuti':         result = ajukanCuti(data); break;
      case 'approveCuti':        result = approveCuti(data); break;
      case 'hapusTransaksi':     result = hapusTransaksi(data); break;
      case 'withdrawTransaksi':  result = withdrawTransaksi(data); break;
      case 'updateQuota':        result = updateQuota(data); break;
      case 'addQuotaJenis':      result = addQuotaJenis(data); break;
      case 'getPHL':             result = getPHL(data); break;
      case 'ajukanPHL':          result = ajukanPHL(data); break;
      case 'hapusPHL':           result = hapusPHL(data); break;
      case 'resetPHLOnly':       result = resetPHLOnly(data); break;
      case 'batalPHL':           result = batalPHL(data); break;
      default:
        result = { success: false, message: 'Action tidak dikenal: ' + action };
    }
    output.setContent(JSON.stringify(result));
  } catch (err) {
    output.setContent(JSON.stringify({ success: false, message: err.toString() }));
  }
  return output;
}

// =====================================================================
// AUTH
// =====================================================================
function login(data) {
  const sheet = getSheet(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();
  const h     = rows[0];
  const iId   = h.indexOf('id');
  const iUser = h.indexOf('username');
  const iPass = h.indexOf('password');
  const iName = h.indexOf('name');
  const iRole = h.indexOf('role');
  const iDept = h.indexOf('dept');
  const iAtas = h.indexOf('atasan_id');
  const iAtas2 = h.indexOf('atasan_id2');
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[iId]) continue;
    if (String(r[iUser]) === String(data.username) && String(r[iPass]) === String(data.password)) {
      return {
        success: true,
        user: {
          id:         String(r[iId]),
          username:   String(r[iUser]),
          name:       String(r[iName]),
          role:       String(r[iRole]),
          dept:       String(r[iDept]),
          atasan_id:  iAtas  >= 0 ? String(r[iAtas]  || '') : '',
          atasan_id2: iAtas2 >= 0 ? String(r[iAtas2] || '') : ''
        }
      };
    }
  }
  return { success: false, message: 'Username atau password salah' };
}

// =====================================================================
// USERS
// =====================================================================
function getUsers() {
  const sheet = getSheet(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();
  const h     = rows[0];
  const users = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    const obj = {};
    h.forEach((key, idx) => { obj[key] = String(rows[i][idx]); });
    delete obj.password;
    users.push(obj);
  }
  return { success: true, users };
}

// =====================================================================
// QUOTA
// =====================================================================
function getQuota(data) {
  const sheet  = getSheet(SHEET_QUOTA);
  const rows   = sheet.getDataRange().getValues();
  const h      = rows[0];
  const iUid   = h.indexOf('user_id');
  const iStart = h.indexOf('start_date');
  const iEnd   = h.indexOf('end_date');
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (String(rows[i][iUid]) === String(data.user_id)) {
      const obj = {};
      h.forEach((key, idx) => {
        if ((key === 'start_date' || key === 'end_date') && rows[i][idx]) {
          obj[key] = fmtDateWIB(rows[i][idx]);
        } else {
          obj[key] = rows[i][idx];
        }
      });
      result.push(obj);
    }
  }
  return { success: true, quota: result };
}

function getAllQuota() {
  const sheet  = getSheet(SHEET_QUOTA);
  const rows   = sheet.getDataRange().getValues();
  const h      = rows[0];
  const iStart = h.indexOf('start_date');
  const iEnd   = h.indexOf('end_date');
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    const obj = {};
    h.forEach((key, idx) => {
      if ((key === 'start_date' || key === 'end_date') && rows[i][idx]) {
        obj[key] = fmtDateWIB(rows[i][idx]);
      } else {
        obj[key] = rows[i][idx];
      }
    });
    result.push(obj);
  }
  return { success: true, quota: result };
}

function updateQuota(data) {
  const sheet  = getSheet(SHEET_QUOTA);
  const rows   = sheet.getDataRange().getValues();
  const h      = rows[0];
  const iUid   = h.indexOf('user_id');
  const iJenis = h.indexOf('jenis_cuti');
  const iQta   = h.indexOf('quota');
  const iTrp   = h.indexOf('terpakai');
  const iSisa  = h.indexOf('sisa');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][iUid]) === String(data.user_id) && String(rows[i][iJenis]) === String(data.jenis_cuti)) {
      if (data.quota    !== undefined) sheet.getRange(i+1, iQta+1).setValue(Number(data.quota));
      if (data.terpakai !== undefined) sheet.getRange(i+1, iTrp+1).setValue(Number(data.terpakai));
      if (data.sisa     !== undefined) sheet.getRange(i+1, iSisa+1).setValue(Number(data.sisa));
      return { success: true };
    }
  }
  return { success: false, message: 'Data quota tidak ditemukan' };
}

function addQuotaJenis(data) {
  const sheet  = getSheet(SHEET_QUOTA);
  const rows   = sheet.getDataRange().getValues();
  const h      = rows[0];
  const iUid   = h.indexOf('user_id');
  const iJenis = h.indexOf('jenis_cuti');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][iUid]) === String(data.user_id) && String(rows[i][iJenis]) === String(data.jenis_cuti)) {
      return { success: false, message: 'Jenis cuti sudah ada untuk karyawan ini' };
    }
  }
  const newId = 'q' + Date.now();
  const quota = Number(data.quota) || 0;
  sheet.appendRow([newId, String(data.user_id), String(data.jenis_cuti), quota, 0, quota]);
  return { success: true };
}

// =====================================================================
// TRANSAKSI
// =====================================================================
function rowToTx(h, row) {
  const obj = {};
  h.forEach((key, idx) => {
    if (idx === 4 || idx === 5 || idx === 9 || idx === 10) {
      obj[key] = fmtDate(row[idx]);
    } else {
      obj[key] = row[idx];
    }
  });
  return obj;
}

function getTransaksi(data) {
  const sheet  = getSheet(SHEET_TRANSAKSI);
  const rows   = sheet.getDataRange().getValues();
  const h      = rows[0];
  const iUid   = h.indexOf('user_id');
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (String(rows[i][iUid]) === String(data.user_id)) result.push(rowToTx(h, rows[i]));
  }
  return { success: true, transaksi: result };
}

function getAllTransaksi() {
  const sheet  = getSheet(SHEET_TRANSAKSI);
  const rows   = sheet.getDataRange().getValues();
  const h      = rows[0];
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push(rowToTx(h, rows[i]));
  }
  return { success: true, transaksi: result };
}

function getPendingForAtasan(data) {
  const atasanId = String(data.atasan_id);
  const uSheet   = getSheet(SHEET_USERS);
  const uRows    = uSheet.getDataRange().getValues();
  const uH       = uRows[0];
  const iId      = uH.indexOf('id');
  const iAtas    = uH.indexOf('atasan_id');
  const bawahanIds = [];
  for (let i = 1; i < uRows.length; i++) {
    if (!uRows[i][0]) continue;
    if (String(uRows[i][iAtas]) === atasanId) bawahanIds.push(String(uRows[i][iId]));
  }
  const sheet  = getSheet(SHEET_TRANSAKSI);
  const rows   = sheet.getDataRange().getValues();
  const h      = rows[0];
  const iUid   = h.indexOf('user_id');
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (bawahanIds.includes(String(rows[i][iUid]))) result.push(rowToTx(h, rows[i]));
  }
  return { success: true, transaksi: result };
}

function ajukanCuti(data) {
  const txSheet    = getSheet(SHEET_TRANSAKSI);
  const quotaSheet = getSheet(SHEET_QUOTA);
  const jumlahHari = Number(data.jumlah_hari);
  const qRows  = quotaSheet.getDataRange().getValues();
  const qH     = qRows[0];
  const qiUid   = qH.indexOf('user_id');
  const qiJns   = qH.indexOf('jenis_cuti');
  const qiQta   = qH.indexOf('quota');
  const qiTrp   = qH.indexOf('terpakai');
  const qiSisa  = qH.indexOf('sisa');
  const qiStart = qH.indexOf('start_date');
  const qiEnd   = qH.indexOf('end_date');

  // Ambil tanggal hari ini dalam WIB
  const nowWIB  = new Date(new Date().getTime() + 7 * 60 * 60 * 1000);
  const today   = nowWIB.toISOString().substring(0, 10);

  // Cari baris quota yang aktif (sesuai start_date & end_date)
  let foundRow = -1;
  for (let i = 1; i < qRows.length; i++) {
    if (String(qRows[i][qiUid]) !== String(data.user_id)) continue;
    if (String(qRows[i][qiJns]) !== String(data.jenis_cuti)) continue;

    // Cek range tanggal jika kolom ada
    if (qiStart >= 0 && qiEnd >= 0) {
      const startDate = fmtDateWIB(qRows[i][qiStart]);
      const endDate   = fmtDateWIB(qRows[i][qiEnd]);
      if (startDate && endDate) {
        if (today >= startDate && today <= endDate) {
          foundRow = i; break; // aktif
        } else {
          continue; // skip yang tidak aktif
        }
      }
    }
    // Jika tidak ada start/end date, langsung pakai
    foundRow = i; break;
  }

  if (foundRow === -1) return { success: false, message: 'Quota aktif untuk jenis cuti ini tidak ditemukan' };
  const sisaSkrg = Number(qRows[foundRow][qiSisa]);
  if (sisaSkrg < jumlahHari) {
    return { success: false, message: 'Sisa cuti tidak mencukupi. Sisa: ' + sisaSkrg + ' hari' };
  }
  const newTerpakai = Number(qRows[foundRow][qiTrp]) + jumlahHari;
  const newSisa     = Number(qRows[foundRow][qiQta]) - newTerpakai;
  quotaSheet.getRange(foundRow+1, qiTrp+1).setValue(newTerpakai);
  quotaSheet.getRange(foundRow+1, qiSisa+1).setValue(newSisa);

  const newId = 'tx' + Date.now();
  txSheet.appendRow([
    newId, String(data.user_id), String(data.nama_karyawan), String(data.jenis_cuti),
    String(data.tgl_mulai), String(data.tgl_selesai), jumlahHari,
    String(data.alasan || ''), 'pending', new Date().toISOString(), '', ''
  ]);

  // Kirim notifikasi email ke atasan
  const atasanId = getAtasanId(data.user_id);
  console.log('Email: atasanId=' + atasanId + ', user=' + data.user_id);
  if (atasanId) {
    const atasanEmail = getUserEmail(atasanId);
    console.log('Email: atasanEmail=' + atasanEmail);
    emailNotifPengajuan(data.nama_karyawan, atasanId, data.jenis_cuti, data.tgl_mulai, data.tgl_selesai, jumlahHari, data.alasan);
  }

  return { success: true, message: 'Pengajuan cuti berhasil diajukan' };
}

function approveCuti(data) {
  const sheet  = getSheet(SHEET_TRANSAKSI);
  const rows   = sheet.getDataRange().getValues();
  const h      = rows[0];
  const iId    = h.indexOf('id');
  const iSts   = h.indexOf('status');
  const iTglAp = h.indexOf('tgl_approval');
  const iRefPHL = h.indexOf('ref_phl_id');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][iId]) !== String(data.transaksi_id)) continue;
    sheet.getRange(i+1, iSts+1).setValue(data.status);
    const now = new Date().toISOString();
    sheet.getRange(i+1, iTglAp+1).setValue(now);

    const refPhlId = iRefPHL >= 0 ? String(rows[i][iRefPHL] || '') : '';

    if (data.status === 'approved' && refPhlId) {
      updatePHLApproval(refPhlId, data.approver_id || '', now);
    }

    if (data.status === 'rejected') {
      if (refPhlId) {
        resetPHL(refPhlId);
      } else {
        const iUid = h.indexOf('user_id');
        const iJns = h.indexOf('jenis_cuti');
        const iJml = h.indexOf('jumlah_hari');
        kembalikanQuota(String(rows[i][iUid]), String(rows[i][iJns]), Number(rows[i][iJml]));
      }
    }

    // Kirim notifikasi email ke karyawan
    try {
      const iUid2 = h.indexOf('user_id');
      const iNama = h.indexOf('nama_karyawan');
      const iJns2 = h.indexOf('jenis_cuti');
      const iTglM = h.indexOf('tgl_mulai');
      const iTglS = h.indexOf('tgl_selesai');
      const approverName = getUserNameById(data.approver_id || '');
      emailNotifApproval(
        String(rows[i][iUid2]),
        String(rows[i][iNama]),
        String(rows[i][iJns2]),
        String(rows[i][iTglM]),
        String(rows[i][iTglS]),
        data.status,
        approverName
      );
    } catch(e) { console.log('Email error: ' + e); }

    return { success: true };
  }
  return { success: false, message: 'Transaksi tidak ditemukan' };
}

function hapusTransaksi(data) {
  const sheet = getSheet(SHEET_TRANSAKSI);
  const rows  = sheet.getDataRange().getValues();
  const h     = rows[0];
  const iId   = h.indexOf('id');
  const iUid  = h.indexOf('user_id');
  const iJns  = h.indexOf('jenis_cuti');
  const iJml  = h.indexOf('jumlah_hari');
  const iSts  = h.indexOf('status');
  const iRefPHL = h.indexOf('ref_phl_id');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][iId]) !== String(data.transaksi_id)) continue;
    if (String(rows[i][iUid]) !== String(data.user_id)) {
      return { success: false, message: 'Tidak berhak menghapus transaksi ini' };
    }
    const status   = String(rows[i][iSts]);
    const refPhlId = iRefPHL >= 0 ? String(rows[i][iRefPHL] || '') : '';

    if (status === 'pending' || status === 'approved') {
      if (refPhlId) {
        resetPHL(refPhlId);
      } else {
        kembalikanQuota(String(rows[i][iUid]), String(rows[i][iJns]), Number(rows[i][iJml]));
      }
    }
    sheet.deleteRow(i + 1);
    return { success: true, message: 'Pengajuan berhasil dihapus' };
  }
  return { success: false, message: 'Transaksi tidak ditemukan' };
}

function kembalikanQuota(userId, jenisCuti, jumlah) {
  const sheet  = getSheet(SHEET_QUOTA);
  const rows   = sheet.getDataRange().getValues();
  const h      = rows[0];
  const iUid   = h.indexOf('user_id');
  const iJns   = h.indexOf('jenis_cuti');
  const iQta   = h.indexOf('quota');
  const iTrp   = h.indexOf('terpakai');
  const iSisa  = h.indexOf('sisa');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][iUid]) === userId && String(rows[i][iJns]) === jenisCuti) {
      const newTerpakai = Math.max(0, Number(rows[i][iTrp]) - jumlah);
      const newSisa     = Number(rows[i][iQta]) - newTerpakai;
      sheet.getRange(i+1, iTrp+1).setValue(newTerpakai);
      sheet.getRange(i+1, iSisa+1).setValue(newSisa);
      break;
    }
  }
}

// Ubah status transaksi menjadi 'withdrawn' (tidak menghapus row)
// Jika cuti biasa → kembalikan quota. Jika PHL → hanya ubah status.
function withdrawTransaksi(data) {
  const sheet = getSheet(SHEET_TRANSAKSI);
  const rows  = sheet.getDataRange().getValues();
  const h     = rows[0];
  const iId   = h.indexOf('id');
  const iUid  = h.indexOf('user_id');
  const iJns  = h.indexOf('jenis_cuti');
  const iJml  = h.indexOf('jumlah_hari');
  const iSts  = h.indexOf('status');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][iId]) !== String(data.transaksi_id)) continue;
    if (String(rows[i][iUid]) !== String(data.user_id)) {
      return { success: false, message: 'Tidak berhak membatalkan.' };
    }
    const status = String(rows[i][iSts]);
    if (status !== 'pending') {
      return { success: false, message: 'Hanya pengajuan berstatus pending yang bisa dibatalkan.' };
    }

    // Ubah status ke withdrawn
    sheet.getRange(i+1, iSts+1).setValue('withdrawn');

    // Kembalikan quota hanya untuk cuti biasa (bukan PHL)
    const jenisCuti = String(rows[i][iJns]);
    if (jenisCuti !== 'PHL') {
      kembalikanQuota(String(rows[i][iUid]), jenisCuti, Number(rows[i][iJml]));
    }

    return { success: true };
  }
  return { success: false, message: 'Transaksi tidak ditemukan.' };
}

// =====================================================================
// PHL
// =====================================================================
function getPHLHeaders() {
  const sheet = getPHLSheet();
  if (!sheet) return null;
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function getPHL(data) {
  try {
    const sheet = getPHLSheet();
    if (!sheet) return { success: false, message: 'Sheet PHL tidak ditemukan. Cek nama sheet di SHEET_PHL.' };

    const rows = sheet.getDataRange().getValues();
    const h    = rows[0];
    const iId          = h.indexOf('id');
    const iUid         = h.indexOf('user_id');
    const iTglMasuk    = h.indexOf('tgl_masuk');
    const iTglStart    = h.indexOf('tgl_start');
    const iTglEnd      = h.indexOf('tgl_end');
    const iStatus      = h.indexOf('status');
    const iDigunakanTgl= h.indexOf('digunakan_tgl');
    const iKet         = h.indexOf('keterangan');
    const iDisetujuiOleh = h.indexOf('penggunaan_disetujui_oleh');
    const iDisetujuiPada = h.indexOf('penggunaan_disetujui_pada');

    const result = [];
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][iId]) continue;
      if (String(rows[i][iUid]) !== String(data.user_id)) continue;

      result.push({
        id:                      String(rows[i][iId]),
        user_id:                 String(rows[i][iUid]),
        tgl_masuk:               fmtDate(rows[i][iTglMasuk]),
        tgl_start:               fmtDate(rows[i][iTglStart]),
        tgl_end:                 fmtDate(rows[i][iTglEnd]),
        status:                  String(rows[i][iStatus] || ''),
        digunakan_tgl:           fmtDate(rows[i][iDigunakanTgl]),
        keterangan:              String(rows[i][iKet] || ''),
        penggunaan_disetujui_oleh: iDisetujuiOleh >= 0 ? String(rows[i][iDisetujuiOleh] || '') : '',
        penggunaan_disetujui_pada: iDisetujuiPada >= 0 ? fmtDate(rows[i][iDisetujuiPada]) : '',
        row: i + 1  // simpan nomor baris untuk update nanti
      });
    }
    return { success: true, phl: result };
  } catch (err) {
    return { success: false, message: 'Error membaca PHL: ' + err.toString() };
  }
}

function ajukanPHL(data) {
  try {
    const sheet = getPHLSheet();
    if (!sheet) return { success: false, message: 'Sheet PHL tidak ditemukan.' };

    const rows = sheet.getDataRange().getValues();
    const h    = rows[0];
    const iId           = h.indexOf('id');
    const iUid          = h.indexOf('user_id');
    const iDigunakanTgl = h.indexOf('digunakan_tgl');
    const iStatus       = h.indexOf('status');

    // Cari baris PHL yang dimaksud
    let foundRow = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][iId]) === String(data.phl_id) && String(rows[i][iUid]) === String(data.user_id)) {
        foundRow = i;
        break;
      }
    }
    if (foundRow === -1) return { success: false, message: 'Data PHL tidak ditemukan.' };

    // Pastikan belum digunakan
    const digunakanTglSkrg = fmtDate(rows[foundRow][iDigunakanTgl]);
    if (digunakanTglSkrg && digunakanTglSkrg !== '') {
      return { success: false, message: 'PHL ini sudah digunakan.' };
    }

    // Update digunakan_tgl di sheet PHL
    sheet.getRange(foundRow + 1, iDigunakanTgl + 1).setValue(String(data.digunakan_tgl));

    // Simpan transaksi di spreadsheet utama
    const txSheet = getSheet(SHEET_TRANSAKSI);
    const newId   = 'tx' + Date.now();
    txSheet.appendRow([
      newId,
      String(data.user_id),
      String(data.nama_karyawan),
      'PHL',
      String(data.digunakan_tgl),  // tgl_mulai = tgl tidak masuk
      String(data.digunakan_tgl),  // tgl_selesai = sama (1 hari)
      1,
      String(data.alasan || ''),
      'pending',
      new Date().toISOString(),
      '',
      String(data.phl_id)  // ref_phl_id
    ]);

    // Kirim notifikasi email ke atasan
    const atasanId = getAtasanId(data.user_id);
    if (atasanId) {
      emailNotifPengajuan(data.nama_karyawan, atasanId, 'PHL', data.digunakan_tgl, data.digunakan_tgl, 1, data.alasan);
    }

    return { success: true, message: 'Pengajuan PHL berhasil diajukan.' };
  } catch (err) {
    return { success: false, message: 'Error mengajukan PHL: ' + err.toString() };
  }
}

function hapusPHL(data) {
  try {
    // Hapus transaksi
    const txSheet = getSheet(SHEET_TRANSAKSI);
    const txRows  = txSheet.getDataRange().getValues();
    const txH     = txRows[0];
    const iId     = txH.indexOf('id');
    const iUid    = txH.indexOf('user_id');
    const iSts    = txH.indexOf('status');

    let txFound = false;
    for (let i = 1; i < txRows.length; i++) {
      if (String(txRows[i][iId]) !== String(data.transaksi_id)) continue;
      if (String(txRows[i][iUid]) !== String(data.user_id)) {
        return { success: false, message: 'Tidak berhak menghapus.' };
      }
      txSheet.deleteRow(i + 1);
      txFound = true;
      break;
    }
    if (!txFound) return { success: false, message: 'Transaksi tidak ditemukan.' };

    // Reset digunakan_tgl & kolom approval di sheet PHL
    resetPHL(data.phl_id);

    return { success: true, message: 'Pengajuan PHL berhasil dihapus.' };
  } catch (err) {
    return { success: false, message: 'Error menghapus PHL: ' + err.toString() };
  }
}

function resetPHL(phlId) {
  try {
    const sheet = getPHLSheet();
    if (!sheet) return;
    const rows  = sheet.getDataRange().getValues();
    const h     = rows[0];
    const iId           = h.indexOf('id');
    const iDigunakanTgl = h.indexOf('digunakan_tgl');
    const iDisetujuiOleh = h.indexOf('penggunaan_disetujui_oleh');
    const iDisetujuiPada = h.indexOf('penggunaan_disetujui_pada');

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][iId]) === String(phlId)) {
        sheet.getRange(i+1, iDigunakanTgl+1).setValue('');
        if (iDisetujuiOleh >= 0) sheet.getRange(i+1, iDisetujuiOleh+1).setValue('');
        if (iDisetujuiPada >= 0) sheet.getRange(i+1, iDisetujuiPada+1).setValue('');
        break;
      }
    }
  } catch(e) {}
}

function updatePHLApproval(phlId, approverId, timestamp) {
  try {
    const sheet = getPHLSheet();
    if (!sheet) return;
    const rows  = sheet.getDataRange().getValues();
    const h     = rows[0];
    const iId            = h.indexOf('id');
    const iDisetujuiOleh = h.indexOf('penggunaan_disetujui_oleh');
    const iDisetujuiPada = h.indexOf('penggunaan_disetujui_pada');

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][iId]) === String(phlId)) {
        if (iDisetujuiOleh >= 0) sheet.getRange(i+1, iDisetujuiOleh+1).setValue(String(approverId));
        if (iDisetujuiPada >= 0) sheet.getRange(i+1, iDisetujuiPada+1).setValue(timestamp);
        break;
      }
    }
  } catch(e) {}
}

// Reset digunakan_tgl saja — transaksi tetap ada sebagai history
function resetPHLOnly(data) {
  try {
    // Verifikasi user punya akses ke PHL ini
    const sheet = getPHLSheet();
    if (!sheet) return { success: false, message: 'Sheet PHL tidak ditemukan.' };
    const rows  = sheet.getDataRange().getValues();
    const h     = rows[0];
    const iId   = h.indexOf('id');
    const iUid  = h.indexOf('user_id');
    const iDigunakanTgl  = h.indexOf('digunakan_tgl');
    const iDisetujuiOleh = h.indexOf('penggunaan_disetujui_oleh');
    const iDisetujuiPada = h.indexOf('penggunaan_disetujui_pada');

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][iId]) === String(data.phl_id)) {
        if (String(rows[i][iUid]) !== String(data.user_id)) {
          return { success: false, message: 'Tidak berhak mengubah PHL ini.' };
        }
        sheet.getRange(i+1, iDigunakanTgl+1).setValue('');
        if (iDisetujuiOleh >= 0) sheet.getRange(i+1, iDisetujuiOleh+1).setValue('');
        if (iDisetujuiPada >= 0) sheet.getRange(i+1, iDisetujuiPada+1).setValue('');
        return { success: true };
      }
    }
    return { success: false, message: 'Data PHL tidak ditemukan.' };
  } catch(err) {
    return { success: false, message: 'Error: ' + err.toString() };
  }
}

// batalPHL: 
// - Jika PENDING → ubah status menjadi 'withdrawn' + reset digunakan_tgl
// - Jika REJECTED → hanya reset digunakan_tgl (transaksi tetap sebagai history)
// - Jika APPROVED → tidak bisa dibatalkan
function batalPHL(data) {
  try {
    const txSheet = getSheet(SHEET_TRANSAKSI);
    const txRows  = txSheet.getDataRange().getValues();
    const txH     = txRows[0];
    const iId     = txH.indexOf('id');
    const iUid    = txH.indexOf('user_id');
    const iSts    = txH.indexOf('status');

    for (let i = 1; i < txRows.length; i++) {
      if (String(txRows[i][iId]) !== String(data.transaksi_id)) continue;
      if (String(txRows[i][iUid]) !== String(data.user_id)) {
        return { success: false, message: 'Tidak berhak membatalkan.' };
      }

      const status = String(txRows[i][iSts]);

      if (status === 'approved') {
        return { success: false, message: 'PHL yang sudah disetujui tidak dapat dibatalkan.' };
      }

      if (status === 'pending') {
        // Ubah status menjadi withdrawn, bukan dihapus
        txSheet.getRange(i + 1, iSts + 1).setValue('withdrawn');
      }
      // Jika rejected/withdrawn: transaksi tetap, tidak diubah lagi

      // Reset digunakan_tgl di sheet PHL agar bisa diajukan ulang
      resetPHL(data.phl_id);
      return { success: true };
    }
    return { success: false, message: 'Transaksi tidak ditemukan.' };
  } catch(err) {
    return { success: false, message: 'Error: ' + err.toString() };
  }
}

// Helper ambil nama user by ID
function getUserNameById(userId) {
  const sheet = getSheet(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();
  const h     = rows[0];
  const iId   = h.indexOf('id');
  const iName = h.indexOf('name');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][iId]) === String(userId)) return String(rows[i][iName] || '');
  }
  return userId;
}

// =====================================================================
// EMAIL NOTIFICATION
// =====================================================================

// Kirim email via SMTP menggunakan App Password
function sendEmail(to, subject, htmlBody) {
  try {
    if (!to || to === '') {
      console.log('Email skipped: empty recipient');
      return;
    }
    console.log('Sending email to: ' + to + ' | subject: ' + subject);
    MailApp.sendEmail({
      to: to,
      subject: subject,
      htmlBody: htmlBody,
      name: APP_NAME,
      replyTo: EMAIL_SENDER
    });
    console.log('Email sent successfully to: ' + to);
  } catch(e) {
    console.log('Email error: ' + e.toString());
  }
}

// Ambil email user dari sheet Users berdasarkan user_id
function getUserEmail(userId) {
  const sheet = getSheet(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();
  const h     = rows[0];
  const iId   = h.indexOf('id');
  const iEmail = h.indexOf('email');
  if (iEmail < 0) return ''; // kolom email tidak ada
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][iId]) === String(userId)) {
      return String(rows[i][iEmail] || '');
    }
  }
  return '';
}

// Ambil atasan_id dari user
function getAtasanId(userId) {
  const sheet = getSheet(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();
  const h     = rows[0];
  const iId   = h.indexOf('id');
  const iAtas = h.indexOf('atasan_id');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][iId]) === String(userId)) {
      return iAtas >= 0 ? String(rows[i][iAtas] || '') : '';
    }
  }
  return '';
}

function formatTanggal(dateStr) {
  if (!dateStr) return '-';
  const d = new Date(dateStr);
  if (isNaN(d)) return dateStr;
  const months = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
  return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
}

function emailTemplate(title, content) {
  return `
  <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 560px; margin: 0 auto; background: #f3f4f6; padding: 24px;">
    <div style="background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.08);">
      <div style="background: #1a56db; padding: 20px 28px;">
        <div style="font-size: 20px; font-weight: 800; color: white; letter-spacing: -0.5px;">✈️ Away</div>
        <div style="font-size: 10px; color: rgba(255,255,255,0.65); margin-top: 2px;">Application for Work-break And daY off</div>
      </div>
      <div style="padding: 28px;">
        <h2 style="font-size: 17px; font-weight: 700; color: #111928; margin: 0 0 16px;">${title}</h2>
        ${content}
      </div>
      <div style="background: #f3f4f6; padding: 14px 28px; font-size: 11px; color: #9ca3af; text-align: center;">
        Email ini dikirim otomatis oleh sistem Away — Poltek SIMAS Berau.<br>Jangan balas email ini.
      </div>
    </div>
  </div>`;
}

function infoRow(label, value) {
  return `<tr>
    <td style="padding: 6px 0; font-size: 13px; color: #6b7280; width: 140px;">${label}</td>
    <td style="padding: 6px 0; font-size: 13px; color: #111928; font-weight: 600;">${value}</td>
  </tr>`;
}

// Email ke atasan saat karyawan ajukan cuti
function emailNotifPengajuan(namaKaryawan, atasanId, jenisCuti, tglMulai, tglSelesai, jumlahHari, alasan) {
  const atasanEmail = getUserEmail(atasanId);
  if (!atasanEmail) return;
  const content = `
    <p style="font-size: 13px; color: #374151; margin: 0 0 16px;">
      Karyawan berikut telah mengajukan cuti dan menunggu persetujuan Anda:
    </p>
    <table style="width:100%; border-collapse:collapse; margin-bottom:20px;">
      ${infoRow('Karyawan', namaKaryawan)}
      ${infoRow('Jenis Cuti', jenisCuti)}
      ${infoRow('Tanggal Mulai', formatTanggal(tglMulai))}
      ${infoRow('Tanggal Selesai', formatTanggal(tglSelesai))}
      ${infoRow('Jumlah Hari', jumlahHari + ' hari')}
      ${infoRow('Alasan', alasan || '-')}
    </table>
    <p style="font-size: 12px; color: #6b7280;">Silakan login ke aplikasi Away untuk memberikan persetujuan.</p>`;
  sendEmail(atasanEmail, `[Away] Pengajuan Cuti Baru — ${namaKaryawan}`, emailTemplate('Pengajuan Cuti Baru', content));
}

// Email ke karyawan saat atasan approve/reject
function emailNotifApproval(userId, namaKaryawan, jenisCuti, tglMulai, tglSelesai, status, approverName) {
  const userEmail = getUserEmail(userId);
  if (!userEmail) return;
  const isApproved = status === 'approved';
  const statusLabel = isApproved ? '✅ Disetujui' : '❌ Ditolak';
  const statusColor = isApproved ? '#057a55' : '#e02424';
  const content = `
    <p style="font-size: 13px; color: #374151; margin: 0 0 16px;">
      Pengajuan cuti Anda telah diproses:
    </p>
    <table style="width:100%; border-collapse:collapse; margin-bottom:20px;">
      ${infoRow('Jenis Cuti', jenisCuti)}
      ${infoRow('Tanggal Mulai', formatTanggal(tglMulai))}
      ${infoRow('Tanggal Selesai', formatTanggal(tglSelesai))}
      ${infoRow('Status', `<span style="color:${statusColor}; font-weight:700;">${statusLabel}</span>`)}
      ${infoRow('Diproses oleh', approverName || '-')}
    </table>
    <p style="font-size: 12px; color: #6b7280;">Silakan login ke aplikasi Away untuk melihat detail.</p>`;
  sendEmail(userEmail, `[Away] Pengajuan Cuti ${isApproved ? 'Disetujui' : 'Ditolak'} — ${jenisCuti}`, emailTemplate(`Pengajuan Cuti ${isApproved ? 'Disetujui' : 'Ditolak'}`, content));
}
