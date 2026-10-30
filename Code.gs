// ============================================================
// APLIKASI TERPADU LPG 3KG (1 SPREADSHEET)
// ============================================================
const SPREADSHEET_ID = '1RlibIUukGGRee9y3Hgk5VUBRt-DVZam8h8vtjlxal2M'; 
const TZ = Session.getScriptTimeZone() || 'Asia/Jakarta';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sistem Terpadu LPG 3KG')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSpreadsheet() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

// Pastikan semua sheet ada
function setupDatabase() {
  const ss = getSpreadsheet();
  const sheets = {
    'Data_User': ['Username', 'Password', 'Nama', 'Level', 'Status'],
    'Master_SPBE': ['Kode SPBE', 'Nama SPBE', 'Alamat', 'Telepon', 'Harga Beli', 'Ship To', 'Status'],
    'Master_Bank': ['Nama Bank', 'Nomor Rekening', 'Atas Nama', 'Cabang', 'Status'],
    'Data_Pangkalan': ['Kode', 'Nama Pangkalan', 'Alamat', 'Kuota', 'Harga', 'Status'],
    'Data_Supir': ['ID Supir', 'Nama Supir', 'Nomor Polisi', 'Nama Kernet'],
    'Data_Perusahaan': ['Field', 'Value'],
    'Data_Kitir': ['Kode Kitir', 'Nomor SA', 'Jenis SA', 'Sold To', 'Ship To', 'SPBE', 'Total Jumlah', 'Tanggal Buat', 'Status'],
    'Detail_Kitir': ['Kode Kitir', 'Tanggal', 'Hari', 'Jumlah', 'Keterangan'],
    'Data_Penebusan': ['Kode Tebus', 'Tanggal', 'Bulan', 'Nominal Bank', 'Rekening Bank', 'No Bukti Bank', 'SPBE', 'Jumlah Unit', 'Harga Beli', 'Jumlah Harga', 'Pembulatan', 'Jumlah Setor', 'Kode Kitir', 'Nomor SA', 'Dari DO', 'Sampai DO'],
    'Data_Surat_Jalan': ['Tanggal', 'No SJ', 'No SA', 'SPBE', 'Supir', 'Pangkalan', 'Alamat', 'Jumlah', 'Harga', 'Total', 'Status', 'No Antri', 'Kode Tebus', 'Total SJ'],
    'Data_SPP': ['Tanggal', 'No SJ', 'SPBE', 'No SA', 'Supir', 'Total Refill', 'Status', 'No Antri', 'Kode Tebus'],
    'Data_Pengiriman': ['Tanggal', 'No SJ', 'Kode Pangkalan', 'Nama Pangkalan', 'Jumlah Tabung']
  };

  for (const [name, headers] of Object.entries(sheets)) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.appendRow(headers);
      sh.setFrozenRows(1);
      if (name === 'Data_User') sh.appendRow(['admin', 'admin123', 'Administrator', 'ADMIN', 'AKTIF']);
    }
  }
}

// ==========================================
// 1. AUTHENTICATION
// ==========================================
function checkLogin(username, password) {
  try {
    const sh = getSpreadsheet().getSheetByName('Data_User');
    if (!sh) return { success: false, message: 'Sheet Data_User tidak ditemukan.' };
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(username).trim() && String(data[i][1]).trim() === String(password).trim()) {
        if (String(data[i][4] || 'AKTIF').toUpperCase() !== 'AKTIF') return { success: false, message: 'Akun tidak aktif.' };
        const token = Utilities.getUuid();
        const user = { username: data[i][0], nama: data[i][2] || data[i][0], level: data[i][3] || 'USER' };
        CacheService.getScriptCache().put('LOGIN_' + token, JSON.stringify(user), 21600);
        return { success: true, token: token, user: user };
      }
    }
    return { success: false, message: 'Username atau Password salah.' };
  } catch(e) { return { success: false, message: e.message }; }
}

function getSession(token) {
  if (!token) return { success: false };
  const data = CacheService.getScriptCache().get('LOGIN_' + token);
  return data ? { success: true, user: JSON.parse(data) } : { success: false };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('LOGIN_' + token);
  return { success: true };
}

function requireLogin_(token) {
  if (!getSession(token).success) throw new Error('Session berakhir. Silakan login kembali.');
}

// ==========================================
// 2. GET INITIAL DATA
// ==========================================
function getInitialData() {
  return {
    company: getPerusahaan_(),
    spbe: getSPBE_(),
    bank: getBank_(),
    pangkalan: getPangkalan_(),
    supir: getSupir_(),
    kitir: getKitirList_(),
    penebusan: getPenebusanList_(),
    counter: getKitirCounter_()
  };
}

// ==========================================
// 3. MASTER DATA (BACA & TULIS)
// ==========================================
function getPerusahaan_() {
  try {
    const sh = getSpreadsheet().getSheetByName('Data_Perusahaan');
    if (!sh) return {};
    const data = sh.getDataRange().getValues();
    const obj = {};
    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0] || '').toLowerCase().replace(/\s+/g, '');
      if (key) obj[key] = String(data[i][1] || '');
    }
    return { nama: obj.nama || 'PT. Nama Perusahaan', alamat: obj.alamat || '', kota: obj.kota || '', telp: obj.telepon || '', email: obj.email || '', npwp: obj.npwp || '', soldTo: obj.soldto || '' };
  } catch(e) { return {}; }
}

function savePerusahaan(token, data) {
  requireLogin_(token);
  const sh = getSpreadsheet().getSheetByName('Data_Perusahaan');
  const fields = ['nama', 'alamat', 'kota', 'telepon', 'email', 'npwp', 'sold to'];
  const values = [data.nama, data.alamat, data.kota, data.telp, data.email, data.npwp, data.soldTo];
  const cur = sh.getDataRange().getValues();
  const existing = {};
  for (let i = 1; i < cur.length; i++) existing[String(cur[i][0]).toLowerCase()] = i + 1;
  for (let i = 0; i < fields.length; i++) {
    if (existing[fields[i]]) sh.getRange(existing[fields[i]], 2).setValue(values[i]);
    else sh.appendRow([fields[i].charAt(0).toUpperCase() + fields[i].slice(1), values[i]]);
  }
  return { success: true, message: 'Data Perusahaan disimpan.' };
}

function getSPBE_() {
  try {
    const sh = getSpreadsheet().getSheetByName('Master_SPBE');
    if (!sh) return [];
    return sh.getDataRange().getValues().slice(1).filter(r => r[0]).map(r => ({ kode: r[0], nama: r[1], alamat: r[2], telp: r[3], hargaBeli: Number(r[4]), shipTo: r[5] }));
  } catch(e) { return []; }
}

function saveSPBE(token, data) {
  requireLogin_(token);
  const sh = getSpreadsheet().getSheetByName('Master_SPBE');
  const d = sh.getDataRange().getValues();
  let idx = -1;
  for(let i=1; i<d.length; i++) if(String(d[i][0]) === data.kode) idx = i+1;
  const row = [data.kode, data.nama, data.alamat, data.telp, Number(data.hargaBeli), data.shipTo, 'AKTIF'];
  if(idx > 0) sh.getRange(idx, 1, 1, 7).setValues([row]); else sh.appendRow(row);
  return { success: true, message: 'SPBE berhasil disimpan.' };
}

function deleteSPBE(token, kode) {
  requireLogin_(token);
  const sh = getSpreadsheet().getSheetByName('Master_SPBE');
  const d = sh.getDataRange().getValues();
  for(let i=d.length-1; i>=1; i--) if(String(d[i][0]) === kode) { sh.deleteRow(i+1); break; }
  return { success: true };
}

function getBank_() {
  try {
    const sh = getSpreadsheet().getSheetByName('Master_Bank');
    if (!sh) return [];
    return sh.getDataRange().getValues().slice(1).filter(r => r[0]).map(r => ({ nama: r[0], rekening: r[1], atasNama: r[2], cabang: r[3] }));
  } catch(e) { return []; }
}

function saveBank(token, data) {
  requireLogin_(token);
  const sh = getSpreadsheet().getSheetByName('Master_Bank');
  const d = sh.getDataRange().getValues();
  let idx = -1;
  for(let i=1; i<d.length; i++) if(String(d[i][1]) === data.rekening) idx = i+1;
  const row = [data.nama, data.rekening, data.atasNama, data.cabang, 'AKTIF'];
  if(idx > 0) sh.getRange(idx, 1, 1, 5).setValues([row]); else sh.appendRow(row);
  return { success: true, message: 'Bank berhasil disimpan.' };
}

function deleteBank(token, rek) {
  requireLogin_(token);
  const sh = getSpreadsheet().getSheetByName('Master_Bank');
  const d = sh.getDataRange().getValues();
  for(let i=d.length-1; i>=1; i--) if(String(d[i][1]) === rek) { sh.deleteRow(i+1); break; }
  return { success: true };
}

function getPangkalan_() {
  try {
    const sh = getSpreadsheet().getSheetByName('Data_Pangkalan');
    if (!sh) return [];
    return sh.getDataRange().getValues().slice(1).filter(r => r[1]).map(r => ({ kode: r[0], nama: r[1], alamat: r[2], kuota: Number(r[3]), harga: Number(r[4]), status: r[5] || 'AKTIF' }));
  } catch(e) { return []; }
}

function savePangkalan(token, data) {
  requireLogin_(token);
  const sh = getSpreadsheet().getSheetByName('Data_Pangkalan');
  const d = sh.getDataRange().getValues();
  let idx = -1;
  for(let i=1; i<d.length; i++) if(String(d[i][0]) === data.kode) idx = i+1;
  const row = [data.kode, data.nama, data.alamat, Number(data.kuota), Number(data.harga), data.status];
  if(idx > 0) sh.getRange(idx, 1, 1, 6).setValues([row]); else sh.appendRow(row);
  return { success: true, message: 'Pangkalan berhasil disimpan.' };
}

function deletePangkalan(token, kode) {
  requireLogin_(token);
  const sh = getSpreadsheet().getSheetByName('Data_Pangkalan');
  const d = sh.getDataRange().getValues();
  for(let i=d.length-1; i>=1; i--) if(String(d[i][0]) === kode) { sh.deleteRow(i+1); break; }
  return { success: true };
}

function getSupir_() {
  try {
    const sh = getSpreadsheet().getSheetByName('Data_Supir');
    if (!sh) return [];
    return sh.getDataRange().getValues().slice(1).filter(r => r[1]).map(r => ({ id: r[0], nama: r[1], armada: r[2], kernet: r[3] }));
  } catch(e) { return []; }
}

function saveSupir(token, data) {
  requireLogin_(token);
  const sh = getSpreadsheet().getSheetByName('Data_Supir');
  const d = sh.getDataRange().getValues();
  let idx = -1;
  for(let i=1; i<d.length; i++) if(String(d[i][0]) === data.id) idx = i+1;
  const row = [data.id, data.nama, data.armada, data.kernet];
  if(idx > 0) sh.getRange(idx, 1, 1, 4).setValues([row]); else sh.appendRow(row);
  return { success: true, message: 'Pengemudi berhasil disimpan.' };
}

function deleteSupir(token, id) {
  requireLogin_(token);
  const sh = getSpreadsheet().getSheetByName('Data_Supir');
  const d = sh.getDataRange().getValues();
  for(let i=d.length-1; i>=1; i--) if(String(d[i][0]) === id) { sh.deleteRow(i+1); break; }
  return { success: true };
}

// ==========================================
// 4. KITIR & PENEBUSAN (BACA & TULIS)
// ==========================================
function getKitirList_() {
  try {
    const sh = getSpreadsheet().getSheetByName('Data_Kitir');
    if (!sh) return [];
    const header = sh.getDataRange().getValues().slice(1).filter(r => r[0]).map(r => ({ kodeKitir: r[0], nomorSA: r[1], jenisSA: r[2], soldTo: r[3], shipTo: r[4], spbe: r[5], total: Number(r[6]) }));
    const shDet = getSpreadsheet().getSheetByName('Detail_Kitir');
    const detData = shDet ? shDet.getDataRange().getValues() : [];
    const details = {};
    for(let i=1; i<detData.length; i++) {
      const k = String(detData[i][0]);
      if(!details[k]) details[k] = [];
      details[k].push({ tanggal: formatDateOnly_(detData[i][1]), hari: detData[i][2], jumlah: Number(detData[i][3]) });
    }
    header.forEach(h => h.jadwal = details[h.kodeKitir] || []);
    return header;
  } catch(e) { return []; }
}

function getKitirCounter_() {
  try {
    const sh = getSpreadsheet().getSheetByName('Data_Kitir');
    if (!sh) return {};
    const counter = {};
    sh.getDataRange().getValues().forEach((r, i) => {
      if(i === 0) return;
      const match = String(r[0]).match(/^KT(\d{4})\.(\d+)$/);
      if (match) { const k = match[1], n = parseInt(match[2], 10); if (!counter[k] || n > counter[k]) counter[k] = n; }
    });
    return counter;
  } catch(e) { return {}; }
}

function saveKitir(token, data) {
  requireLogin_(token);
  const ss = getSpreadsheet(), shH = ss.getSheetByName('Data_Kitir'), shD = ss.getSheetByName('Detail_Kitir');
  const total = data.jadwal.reduce((s, j) => s + Number(j.jumlah), 0);
  const row = [data.kodeKitir, data.nomorSA, data.jenisSA, data.soldTo, data.shipTo, data.spbe, total, new Date(), 'AKTIF'];
  
  const hD = shH.getDataRange().getValues();
  let idx = -1;
  for(let i=1; i<hD.length; i++) if(String(hD[i][0]) === data.kodeKitir) idx = i+1;
  if(idx > 0) shH.getRange(idx, 1, 1, row.length).setValues([row]); else shH.appendRow(row);
  
  const dD = shD.getDataRange().getValues();
  for(let i=dD.length-1; i>=1; i--) if(String(dD[i][0]) === data.kodeKitir) shD.deleteRow(i+1);
  data.jadwal.forEach(j => shD.appendRow([data.kodeKitir, new Date(j.tanggal + 'T00:00:00'), j.hari, Number(j.jumlah), '']));
  
  return { success: true, message: 'Kitir berhasil disimpan.' };
}

function deleteKitir(token, kode) {
  requireLogin_(token);
  const ss = getSpreadsheet(), shH = ss.getSheetByName('Data_Kitir'), shD = ss.getSheetByName('Detail_Kitir');
  const hD = shH.getDataRange().getValues();
  for(let i=hD.length-1; i>=1; i--) if(String(hD[i][0]) === kode) shH.deleteRow(i+1);
  const dD = shD.getDataRange().getValues();
  for(let i=dD.length-1; i>=1; i--) if(String(dD[i][0]) === kode) shD.deleteRow(i+1);
  return { success: true };
}

function getPenebusanList_() {
  try {
    const sh = getSpreadsheet().getSheetByName('Data_Penebusan');
    if (!sh) return [];
    return sh.getDataRange().getValues().slice(1).filter(r => r[0]).map(r => ({
      kodeTebus: r[0], tanggal: formatDateOnly_(r[1]), spbe: r[6], jumlahUnit: Number(r[7]), kodeKitir: r[12], nomorSA: r[13], doFrom: formatDateOnly_(r[14]), doTo: formatDateOnly_(r[15]), jumlahSetor: Number(r[11]), hargaBeli: Number(r[8])
    }));
  } catch(e) { return []; }
}

function savePenebusan(token, data) {
  requireLogin_(token);
  const sh = getSpreadsheet().getSheetByName('Data_Penebusan');
  const isEdit = data.kodeTebus && data.kodeTebus.length > 0;
  let kodeTebus = data.kodeTebus;
  
  if(!isEdit) {
    const prefix = 'TBD' + Utilities.formatDate(new Date(), TZ, 'yyMM');
    let maxNo = 0;
    sh.getDataRange().getValues().forEach(r => { const k = String(r[0]); if(k.startsWith(prefix)) { const n = parseInt(k.replace(prefix, ''), 10); if(!isNaN(n) && n > maxNo) maxNo = n; } });
    kodeTebus = prefix + String(maxNo + 1).padStart(3, '0');
  }
  
  const jmlHarga = Number(data.jumlahUnit) * Number(data.hargaBeli);
  const pembulatan = Number(data.jumlahSetor) - jmlHarga;
  const row = [kodeTebus, new Date(data.tanggal + 'T00:00:00'), Utilities.formatDate(new Date(data.tanggal + 'T00:00:00'), TZ, 'MMMM yyyy'), Number(data.jumlahSetor), data.rekeningBank, data.noBuktiBank, data.spbe, Number(data.jumlahUnit), Number(data.hargaBeli), jmlHarga, pembulatan, Number(data.jumlahSetor), data.kodeKitir, data.nomorSA, new Date(data.doFrom + 'T00:00:00'), new Date(data.doTo + 'T00:00:00')];
  
  if(isEdit) {
    const d = sh.getDataRange().getValues();
    for(let i=1; i<d.length; i++) if(String(d[i][0]) === kodeTebus) { sh.getRange(i+1, 1, 1, row.length).setValues([row]); break; }
  } else {
    sh.appendRow(row);
  }
  return { success: true, message: 'Penebusan berhasil disimpan.', kodeTebus };
}

function deletePenebusan(token, kode) {
  requireLogin_(token);
  const sh = getSpreadsheet().getSheetByName('Data_Penebusan');
  const d = sh.getDataRange().getValues();
  for(let i=d.length-1; i>=1; i--) if(String(d[i][0]) === kode) { sh.deleteRow(i+1); break; }
  return { success: true };
}

// ==========================================
// 5. SURAT JALAN & PENGIRIMAN
// ==========================================
function generateNoSJ(tanggal) {
  if (!tanggal) return '';
  const prefix = 'SJ' + tanggal.replace(/-/g, '').slice(2) + '.';
  const sh = getSpreadsheet().getSheetByName('Data_Surat_Jalan');
  let maxNo = 0;
  if(sh) {
    sh.getDataRange().getValues().forEach(r => {
      const v = String(r[1] || '');
      if(v.startsWith(prefix)) { const n = parseInt(v.replace(prefix, ''), 10); if(!isNaN(n) && n > maxNo) maxNo = n; }
    });
  }
  return prefix + String(maxNo + 1).padStart(2, '0');
}

function prosesSuratJalan(data) {
  if (data.tujuan.reduce((s, i) => s + Number(i.jumlah), 0) !== 560) return { status: 'error', message: 'Total Harus 560 Tabung!' };
  const ss = getSpreadsheet(), shSJ = ss.getSheetByName('Data_Surat_Jalan'), shSPP = ss.getSheetByName('Data_SPP'), shPengiriman = ss.getSheetByName('Data_Pengiriman');
  
  const rows = data.tujuan.map(i => [data.tanggal, data.noSJ, data.noSA, data.spbe, data.supir, i.pangkalan, i.alamat, i.jumlah, i.harga, i.jumlah * i.harga, 'AKTIF', '-', data.kodeTebus, 560]);
  shSJ.getRange(shSJ.getLastRow() + 1, 1, rows.length, 14).setValues(rows);
  if(shSPP) shSPP.appendRow([data.tanggal, data.noSJ, data.spbe, data.noSA, data.supir, 560, 'AKTIF', '-', data.kodeTebus]);
  if(shPengiriman) shPengiriman.getRange(shPengiriman.getLastRow() + 1, 1, data.tujuan.length, 5).setValues(data.tujuan.map(i => [data.tanggal, data.noSJ, i.kode, i.pangkalan, i.jumlah]));
  
  return { status: 'success', message: 'Surat Jalan berhasil disimpan!' };
}

function getDaftarSJ(tglMulai, tglSelesai) {
  const sh = getSpreadsheet().getSheetByName('Data_SPP');
  if (!sh) return [];
  let out = [];
  sh.getDataRange().getValues().forEach((r, i) => {
    if(i===0 || !r[0]) return;
    const tgl = formatDateOnly_(r[0]);
    if(tgl >= tglMulai && tgl <= tglSelesai) out.push({ tanggal: tgl, noSJ: r[1], spbe: r[2], noSA: r[3], supir: r[4], totalRefill: r[5], status: r[6], kodeTebus: r[8] });
  });
  return out;
}

function batalkanSJ(token, noSJ) {
  requireLogin_(token);
  const ss = getSpreadsheet();
  ['Data_Surat_Jalan', 'Data_SPP'].forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if(sh) {
      const d = sh.getDataRange().getValues();
      for(let i=1; i<d.length; i++) {
        if(String(d[i][1]) === noSJ) {
          if(sheetName === 'Data_Surat_Jalan') sh.getRange(i+1, 11).setValue('BATAL');
          else sh.getRange(i+1, 7).setValue('BATAL');
        }
      }
    }
  });
  const shPeng = ss.getSheetByName('Data_Pengiriman');
  if(shPeng) {
    const d = shPeng.getDataRange().getValues();
    for(let i=d.length-1; i>=1; i--) if(String(d[i][1]) === noSJ) shPeng.deleteRow(i+1);
  }
  return { success: true, message: 'SJ Dibatalkan.' };
}

function getDistribusiPenebusan(tglMulai, tglSelesai) {
  const sjData = getSpreadsheet().getSheetByName('Data_Surat_Jalan').getDataRange().getValues();
  const tebusData = getSpreadsheet().getSheetByName('Data_Penebusan').getDataRange().getValues();
  
  const mapTebus = {};
  for(let i=1; i<tebusData.length; i++) {
    mapTebus[String(tebusData[i][0])] = { tglTebus: formatDateOnly_(tebusData[i][1]), noSA: tebusData[i][13], spbe: tebusData[i][6], jmlTebus: tebusData[i][7] };
  }

  const result = [];
  for(let i=1; i<sjData.length; i++) {
    const tglSJ = formatDateOnly_(sjData[i][0]);
    const kodeTebus = String(sjData[i][12] || '');
    if (kodeTebus && kodeTebus !== '-' && tglSJ >= tglMulai && tglSJ <= tglSelesai) {
      const exist = result.find(r => r.noSJ === sjData[i][1]);
      if (!exist) {
        const info = mapTebus[kodeTebus] || {};
        result.push({ noSJ: sjData[i][1], tanggal: tglSJ, kodeTebus, pengemudi: sjData[i][4], noPolisi: sjData[i][4], jmlPkl: 0, jmlKirim: sjData[i][13], jmlNominal: 0, ...info });
      }
      const target = result.find(r => r.noSJ === sjData[i][1]);
      target.jmlPkl += 1; target.jmlNominal += Number(sjData[i][9] || 0);
    }
  }
  return result;
}

// ==========================================
// 6. BACKUP & HAPUS DATA
// ==========================================
function exportBackupExcel(tglMulai, tglSelesai) {
  const ss = getSpreadsheet();
  const newSS = SpreadsheetApp.create('Backup_Excel_LPG_' + tglMulai + '_sd_' + tglSelesai);
  
  const sh1 = newSS.getSheets()[0];
  sh1.setName('Ringkasan');
  sh1.getRange(1,1,1,4).merge().setValue('BACKUP DATA APLIKASI LPG 3KG').setFontWeight('bold').setFontSize(14);
  sh1.getRange(2,1).setValue('Periode:').setFontWeight('bold'); sh1.getRange(2,2).setValue(tglMulai + ' s/d ' + tglSelesai);
  sh1.getRange(3,1).setValue('Waktu Backup:').setFontWeight('bold'); sh1.getRange(3,2).setValue(new Date());
  
  const copyFiltered = (sheetName, dateColIdx, newSheetName) => {
    const shAsal = ss.getSheetByName(sheetName);
    if (!shAsal) return;
    const data = shAsal.getDataRange().getValues();
    if (data.length < 2) return;

    const newSh = newSS.insertSheet(newSheetName);
    const headers = data[0];
    const filtered = [headers];

    for (let i = 1; i < data.length; i++) {
      let tglVal = data[i][dateColIdx];
      let tglStr = '';
      if (tglVal instanceof Date) tglStr = Utilities.formatDate(tglVal, TZ, 'yyyy-MM-dd');
      else tglStr = String(tglVal || '').substring(0, 10);
      
      if (tglStr >= tglMulai && tglStr <= tglSelesai) filtered.push(data[i]);
    }

    if (filtered.length > 1) {
      newSh.getRange(1, 1, filtered.length, headers.length).setValues(filtered);
      newSh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
      newSh.setFrozenRows(1);
    }
  };

  copyFiltered('Data_Surat_Jalan', 0, 'Data Surat Jalan');
  copyFiltered('Data_Kitir', 7, 'Data Kitir');
  copyFiltered('Data_Penebusan', 1, 'Data Penebusan');
  
  const url = 'https://docs.google.com/spreadsheets/d/' + newSS.getId() + '/export?format=xlsx';
  const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob().setName(newSS.getName() + '.xlsx');
  const file = DriveApp.createFile(blob); DriveApp.getFileById(newSS.getId()).setTrashed(true);
  return { downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId() };
}

function exportBackupPDF(tglMulai, tglSelesai) {
  const ss = getSpreadsheet();
  const tempSS = SpreadsheetApp.create('Temp_Backup_PDF_' + new Date().getTime());
  const sh = tempSS.getSheets()[0]; sh.setName('Laporan Backup');
  
  sh.getRange(1,1,1,5).merge().setValue('LAPORAN BACKUP DATA').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sh.getRange(2,1,1,5).merge().setValue('Periode: ' + tglMulai + ' s/d ' + tglSelesai).setHorizontalAlignment('center');
  
  let r = 4;
  sh.getRange(r, 1, 1, 5).setValues([['Tanggal', 'No SJ', 'SPBE', 'Supir', 'Total Tabung']]).setFontWeight('bold').setBackground('#E2E8F0'); r++;

  const shSJ = ss.getSheetByName('Data_Surat_Jalan');
  if (shSJ) {
    const dataSJ = shSJ.getDataRange().getValues();
    const mapSJ = {};
    for(let i=1; i<dataSJ.length; i++) {
      const tgl = formatDateOnly_(dataSJ[i][0]);
      if (tgl >= tglMulai && tgl <= tglSelesai && String(dataSJ[i][10]).toUpperCase() !== 'BATAL') {
        const noSJ = dataSJ[i][1];
        if(!mapSJ[noSJ]) mapSJ[noSJ] = { tgl: tgl, noSJ: noSJ, spbe: dataSJ[i][3], supir: dataSJ[i][4], qty: 0 };
        mapSJ[noSJ].qty += Number(dataSJ[i][7]) || 0;
      }
    }
    const rowsPDF = [];
    Object.values(mapSJ).sort((a,b) => a.tgl.localeCompare(b.tgl)).forEach(x => rowsPDF.push([x.tgl, x.noSJ, x.spbe, x.supir, x.qty]));
    if(rowsPDF.length > 0) { sh.getRange(r, 1, rowsPDF.length, 5).setValues(rowsPDF); }
  }

  SpreadsheetApp.flush();
  const url = 'https://docs.google.com/spreadsheets/d/' + tempSS.getId() + '/export?format=pdf&portrait=true&fitw=true&gridlines=false';
  const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob().setName('Backup_PDF_LPG.pdf');
  const file = DriveApp.createFile(blob); DriveApp.getFileById(tempSS.getId()).setTrashed(true);
  return { downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId() };
}

function hapusDataTransaksiLama(batasTanggal) {
  const ss = getSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let hapusSJ = 0, hapusKitir = 0;

  try {
    const prosesHapus = (sheetName, colTglIdx) => {
      const sh = ss.getSheetByName(sheetName);
      if (!sh) return;
      const data = sh.getDataRange().getValues();
      const newData = [data[0]];
      for (let i = 1; i < data.length; i++) {
        const tgl = formatDateOnly_(data[i][colTglIdx]);
        if (tgl && tgl < batasTanggal) {
          if(sheetName === 'Data_Surat_Jalan') hapusSJ++;
          if(sheetName === 'Data_Kitir') hapusKitir++;
        } else {
          newData.push(data[i]);
        }
      }
      sh.clearContents();
      if (newData.length > 0) sh.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
    };

    prosesHapus('Data_Surat_Jalan', 0); prosesHapus('Data_SPP', 0); prosesHapus('Data_Pengiriman', 0);
    prosesHapus('Data_Kitir', 7); prosesHapus('Detail_Kitir', 1); prosesHapus('Data_Penebusan', 1);

    return { success: true, message: `Pembersihan berhasil! ${hapusSJ} Data SJ & ${hapusKitir} Data Kitir lama dihapus.` };
  } catch (e) {
    throw new Error('Gagal: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function formatDateOnly_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  return '';
}