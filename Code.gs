const APP_NAME = 'SIM SPPD UPTD Puskesmas Tanjung Puri';
const APP_VERSION = '2.0';
const SHEETS = {
  USERS: 'USERS',
  SETTINGS: 'SETTINGS',
  EMPLOYEES: 'EMPLOYEES',
  COSTS: 'STANDARD_BIAYA',
  ACCOUNTS: 'REKENING_KEGIATAN',
  SPD: 'SPD',
  LOGS: 'AUDIT_LOG'
};
const ROLES = {
  GRAND_ADMIN: 'grand_admin',
  ADMIN: 'admin'
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEETS.USERS, ['ID','USERNAME','PASSWORD_HASH','NAMA','ROLE','ACTIVE','CREATED_AT','UPDATED_AT']);
  ensureSheet_(ss, SHEETS.SETTINGS, ['KEY','VALUE']);
  ensureSheet_(ss, SHEETS.EMPLOYEES, ['ID','NIP','NAMA','JABATAN','PANGKAT_GOL','UNIT_KERJA','ACTIVE','MAX_PER_DAY','TOTAL_PERJALANAN','CREATED_AT','UPDATED_AT']);
  ensureSheet_(ss, SHEETS.COSTS, ['ID','KATEGORI','SUBKATEGORI','TUJUAN','SATUAN','NILAI','KETERANGAN','ACTIVE','UPDATED_AT']);
  ensureSheet_(ss, SHEETS.ACCOUNTS, ['ID','KODE_KEGIATAN','NAMA_KEGIATAN','KODE_SUB_KEGIATAN','NAMA_SUB_KEGIATAN','KODE_REKENING','NAMA_REKENING','ACTIVE','UPDATED_AT']);
  ensureSheet_(ss, SHEETS.SPD, ['ID','NO_SPD','TGL_BERANGKAT','TGL_PULANG','PEGAWAI_ID','NAMA_PEGAWAI','TUJUAN','KEPERLUAN','KEGIATAN','SUB_KEGIATAN','STATUS','CREATED_BY','CREATED_AT','UPDATED_AT']);
  ensureSheet_(ss, SHEETS.LOGS, ['ID','TIMESTAMP','USERNAME','ACTION','DETAIL','STATUS']);

  const settings = getSettingsMap_();
  if (!settings.APP_LOGO_URL) setSetting_('APP_LOGO_URL', '');
  if (!settings.PRIMARY_COLOR) setSetting_('PRIMARY_COLOR', '#0d6efd');
  if (!settings.SECONDARY_COLOR) setSetting_('SECONDARY_COLOR', '#ffc107');

  const usersSheet = getSheet_(SHEETS.USERS);
  if (usersSheet.getLastRow() === 1) {
    usersSheet.appendRow([
      makeId_('USR'),
      'grandadmin',
      hashPassword_('123456'),
      'Grand Admin',
      ROLES.GRAND_ADMIN,
      'Y',
      new Date(),
      new Date()
    ]);
  }

  seedDemoData_();
  return { ok: true, message: 'Setup awal selesai.' };
}

function seedDemoData_() {
  const emp = getSheet_(SHEETS.EMPLOYEES);
  if (emp.getLastRow() === 1) {
    emp.appendRow([makeId_('EMP'),'19870001','Ari Saputra','Kepala TU','III/c','UPTD Puskesmas Tanjung Puri','Y',1,0,new Date(),new Date()]);
    emp.appendRow([makeId_('EMP'),'19870002','Dewi Lestari','Bendahara','III/b','UPTD Puskesmas Tanjung Puri','Y',1,0,new Date(),new Date()]);
  }
  const costs = getSheet_(SHEETS.COSTS);
  if (costs.getLastRow() === 1) {
    costs.appendRow([makeId_('CST'),'Uang Harian','Dalam Daerah','Sintang','Orang/Hari',150000,'Contoh konfigurasi, sesuaikan aturan daerah','Y',new Date()]);
    costs.appendRow([makeId_('CST'),'Transport','Darat','Sintang-Pontianak','Orang/PP',350000,'Contoh konfigurasi, sesuaikan aturan daerah','Y',new Date()]);
  }
  const acc = getSheet_(SHEETS.ACCOUNTS);
  if (acc.getLastRow() === 1) {
    acc.appendRow([makeId_('ACC'),'1.02.01','Pelayanan Kesehatan Dasar','1.02.01.2.01','Perjalanan Dinas Dalam Daerah','5.1.02.04','Belanja Perjalanan Dinas','Y',new Date()]);
  }
}

function login(payload) {
  payload = payload || {};
  const username = String(payload.username || '').trim();
  const password = String(payload.password || '');
  if (!username || !password) return fail_('Username dan password wajib diisi.');

  const user = findUserByUsername_(username);
  if (!user || user.ACTIVE !== 'Y') {
    logAction_('guest', 'LOGIN', username, 'FAILED');
    return fail_('Akun tidak ditemukan atau nonaktif.');
  }
  if (user.PASSWORD_HASH !== hashPassword_(password)) {
    logAction_(username, 'LOGIN', 'Password salah', 'FAILED');
    return fail_('Password salah.');
  }

  const token = Utilities.getUuid();
  const session = {
    token: token,
    username: user.USERNAME,
    name: user.NAMA,
    role: user.ROLE,
    loginAt: new Date().toISOString()
  };
  PropertiesService.getScriptProperties().setProperty('SESSION_' + token, JSON.stringify(session));
  logAction_(username, 'LOGIN', 'Berhasil login', 'SUCCESS');
  return { ok: true, message: 'Login berhasil.', session: session };
}

function logout(token) {
  const session = getSession_(token, false);
  if (session) {
    logAction_(session.username, 'LOGOUT', 'Keluar dari aplikasi', 'SUCCESS');
    PropertiesService.getScriptProperties().deleteProperty('SESSION_' + token);
  }
  return { ok: true, message: 'Logout berhasil.' };
}

function checkSession(token) {
  const session = getSession_(token, false);
  if (!session) return fail_('Sesi tidak valid. Silakan login kembali.');
  return { ok: true, session: session, app: getAppConfig_() };
}

function getBootstrapData(token) {
  const session = getSession_(token, false);
  if (!session) return fail_('Sesi tidak valid.');
  return {
    ok: true,
    session: session,
    app: getAppConfig_(),
    stats: getStats_(),
    admins: session.role === ROLES.GRAND_ADMIN ? listUsers_(session).data : [],
    employees: listEmployees(token).data,
    costs: listCosts(token).data,
    accounts: listAccounts(token).data,
    spd: listSpd(token).data
  };
}

function createAdmin(token, payload) {
  const session = requireRole_(token, [ROLES.GRAND_ADMIN]);
  payload = payload || {};
  const username = String(payload.username || '').trim().toLowerCase();
  const password = String(payload.password || '');
  const name = String(payload.name || '').trim();
  if (!username || !password || !name) return fail_('Nama, username, dan password wajib diisi.');
  if (findUserByUsername_(username)) return fail_('Username sudah digunakan.');

  getSheet_(SHEETS.USERS).appendRow([
    makeId_('USR'), username, hashPassword_(password), name, ROLES.ADMIN, 'Y', new Date(), new Date()
  ]);
  logAction_(session.username, 'CREATE_ADMIN', username, 'SUCCESS');
  return { ok: true, message: 'Admin baru berhasil ditambahkan.' };
}

function listUsers(token) {
  const session = requireRole_(token, [ROLES.GRAND_ADMIN]);
  return listUsers_(session);
}

function listUsers_(session) {
  const rows = getObjects_(SHEETS.USERS).map(function(r){
    return { ID:r.ID, USERNAME:r.USERNAME, NAMA:r.NAMA, ROLE:r.ROLE, ACTIVE:r.ACTIVE };
  });
  return { ok: true, data: rows };
}

function saveEmployee(token, payload) {
  const session = requireRole_(token, [ROLES.GRAND_ADMIN, ROLES.ADMIN]);
  payload = payload || {};
  const sheet = getSheet_(SHEETS.EMPLOYEES);
  const obj = {
    ID: payload.id || makeId_('EMP'),
    NIP: String(payload.nip || '').trim(),
    NAMA: String(payload.nama || '').trim(),
    JABATAN: String(payload.jabatan || '').trim(),
    PANGKAT_GOL: String(payload.pangkat || '').trim(),
    UNIT_KERJA: String(payload.unitKerja || '').trim(),
    ACTIVE: payload.active === false ? 'N' : 'Y',
    MAX_PER_DAY: Number(payload.maxPerDay || 1),
    TOTAL_PERJALANAN: Number(payload.totalPerjalanan || 0),
    CREATED_AT: new Date(),
    UPDATED_AT: new Date()
  };
  if (!obj.NIP || !obj.NAMA) return fail_('NIP dan nama pegawai wajib diisi.');

  upsertById_(sheet, obj, 'ID');
  logAction_(session.username, 'SAVE_EMPLOYEE', obj.NAMA, 'SUCCESS');
  return { ok: true, message: 'Data pegawai tersimpan.' };
}

function listEmployees(token) {
  requireRole_(token, [ROLES.GRAND_ADMIN, ROLES.ADMIN]);
  return { ok: true, data: getObjects_(SHEETS.EMPLOYEES) };
}

function saveCost(token, payload) {
  const session = requireRole_(token, [ROLES.GRAND_ADMIN, ROLES.ADMIN]);
  payload = payload || {};
  const obj = {
    ID: payload.id || makeId_('CST'),
    KATEGORI: String(payload.kategori || '').trim(),
    SUBKATEGORI: String(payload.subkategori || '').trim(),
    TUJUAN: String(payload.tujuan || '').trim(),
    SATUAN: String(payload.satuan || '').trim(),
    NILAI: Number(payload.nilai || 0),
    KETERANGAN: String(payload.keterangan || '').trim(),
    ACTIVE: payload.active === false ? 'N' : 'Y',
    UPDATED_AT: new Date()
  };
  if (!obj.KATEGORI) return fail_('Kategori biaya wajib diisi.');
  upsertById_(getSheet_(SHEETS.COSTS), obj, 'ID');
  logAction_(session.username, 'SAVE_COST', obj.KATEGORI + ' ' + obj.TUJUAN, 'SUCCESS');
  return { ok: true, message: 'Standar biaya tersimpan.' };
}

function listCosts(token) {
  requireRole_(token, [ROLES.GRAND_ADMIN, ROLES.ADMIN]);
  return { ok: true, data: getObjects_(SHEETS.COSTS) };
}

function saveAccount(token, payload) {
  const session = requireRole_(token, [ROLES.GRAND_ADMIN, ROLES.ADMIN]);
  payload = payload || {};
  const obj = {
    ID: payload.id || makeId_('ACC'),
    KODE_KEGIATAN: String(payload.kodeKegiatan || '').trim(),
    NAMA_KEGIATAN: String(payload.namaKegiatan || '').trim(),
    KODE_SUB_KEGIATAN: String(payload.kodeSubKegiatan || '').trim(),
    NAMA_SUB_KEGIATAN: String(payload.namaSubKegiatan || '').trim(),
    KODE_REKENING: String(payload.kodeRekening || '').trim(),
    NAMA_REKENING: String(payload.namaRekening || '').trim(),
    ACTIVE: payload.active === false ? 'N' : 'Y',
    UPDATED_AT: new Date()
  };
  if (!obj.KODE_REKENING || !obj.NAMA_REKENING) return fail_('Kode dan nama rekening wajib diisi.');
  upsertById_(getSheet_(SHEETS.ACCOUNTS), obj, 'ID');
  logAction_(session.username, 'SAVE_ACCOUNT', obj.KODE_REKENING, 'SUCCESS');
  return { ok: true, message: 'Rekening kegiatan tersimpan.' };
}

function listAccounts(token) {
  requireRole_(token, [ROLES.GRAND_ADMIN, ROLES.ADMIN]);
  return { ok: true, data: getObjects_(SHEETS.ACCOUNTS) };
}

function saveSpd(token, payload) {
  const session = requireRole_(token, [ROLES.GRAND_ADMIN, ROLES.ADMIN]);
  payload = payload || {};
  const employeeId = String(payload.pegawaiId || '').trim();
  const tglBerangkat = String(payload.tglBerangkat || '').trim();
  const tglPulang = String(payload.tglPulang || '').trim();
  if (!employeeId || !tglBerangkat || !tglPulang) return fail_('Pegawai, tanggal berangkat, dan tanggal pulang wajib diisi.');

  const employees = getObjects_(SHEETS.EMPLOYEES);
  const emp = employees.find(function(e){ return e.ID === employeeId && e.ACTIVE === 'Y'; });
  if (!emp) return fail_('Pegawai tidak ditemukan atau nonaktif.');

  const dayCount = countSpdForEmployeeOnDate_(employeeId, tglBerangkat, payload.id);
  const maxPerDay = Number(emp.MAX_PER_DAY || 1);
  if (dayCount >= maxPerDay) {
    return fail_('Pegawai sudah mencapai batas perjalanan dinas pada tanggal berangkat tersebut.');
  }

  const obj = {
    ID: payload.id || makeId_('SPD'),
    NO_SPD: payload.noSpd || generateNoSpd_(),
    TGL_BERANGKAT: tglBerangkat,
    TGL_PULANG: tglPulang,
    PEGAWAI_ID: emp.ID,
    NAMA_PEGAWAI: emp.NAMA,
    TUJUAN: String(payload.tujuan || '').trim(),
    KEPERLUAN: String(payload.keperluan || '').trim(),
    KEGIATAN: String(payload.kegiatan || '').trim(),
    SUB_KEGIATAN: String(payload.subKegiatan || '').trim(),
    STATUS: String(payload.status || 'Draft'),
    CREATED_BY: session.username,
    CREATED_AT: new Date(),
    UPDATED_AT: new Date()
  };
  if (!obj.TUJUAN || !obj.KEPERLUAN) return fail_('Tujuan dan keperluan wajib diisi.');

  upsertById_(getSheet_(SHEETS.SPD), obj, 'ID');
  refreshEmployeeTripTotal_(emp.ID);
  logAction_(session.username, 'SAVE_SPD', obj.NO_SPD, 'SUCCESS');
  return { ok: true, message: 'Data SPD tersimpan.', noSpd: obj.NO_SPD };
}

function listSpd(token) {
  requireRole_(token, [ROLES.GRAND_ADMIN, ROLES.ADMIN]);
  return { ok: true, data: getObjects_(SHEETS.SPD) };
}

function updateAppSettings(token, payload) {
  const session = requireRole_(token, [ROLES.GRAND_ADMIN, ROLES.ADMIN]);
  payload = payload || {};
  if (payload.logoUrl !== undefined) setSetting_('APP_LOGO_URL', String(payload.logoUrl || ''));
  if (payload.primaryColor !== undefined) setSetting_('PRIMARY_COLOR', String(payload.primaryColor || '#0d6efd'));
  if (payload.secondaryColor !== undefined) setSetting_('SECONDARY_COLOR', String(payload.secondaryColor || '#ffc107'));
  logAction_(session.username, 'UPDATE_SETTINGS', 'Pengaturan aplikasi', 'SUCCESS');
  return { ok: true, message: 'Pengaturan aplikasi diperbarui.', app: getAppConfig_() };
}

function getAuditLogs(token) {
  requireRole_(token, [ROLES.GRAND_ADMIN, ROLES.ADMIN]);
  return { ok: true, data: getObjects_(SHEETS.LOGS).slice(-200).reverse() };
}

function getStats_() {
  return {
    totalPegawai: Math.max(getSheet_(SHEETS.EMPLOYEES).getLastRow() - 1, 0),
    totalBiaya: Math.max(getSheet_(SHEETS.COSTS).getLastRow() - 1, 0),
    totalRekening: Math.max(getSheet_(SHEETS.ACCOUNTS).getLastRow() - 1, 0),
    totalSpd: Math.max(getSheet_(SHEETS.SPD).getLastRow() - 1, 0)
  };
}

function getAppConfig_() {
  const s = getSettingsMap_();
  return {
    appName: APP_NAME,
    version: APP_VERSION,
    logoUrl: s.APP_LOGO_URL || '',
    primaryColor: s.PRIMARY_COLOR || '#0d6efd',
    secondaryColor: s.SECONDARY_COLOR || '#ffc107'
  };
}

function requireRole_(token, allowedRoles) {
  const session = getSession_(token, true);
  if (allowedRoles.indexOf(session.role) === -1) throw new Error('Anda tidak memiliki hak akses.');
  return session;
}

function getSession_(token, throwOnError) {
  const raw = token ? PropertiesService.getScriptProperties().getProperty('SESSION_' + token) : null;
  if (!raw) {
    if (throwOnError) throw new Error('Sesi tidak valid. Silakan login kembali.');
    return null;
  }
  return JSON.parse(raw);
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  const headerRange = sh.getRange(1,1,1,headers.length);
  headerRange.setValues([headers]).setFontWeight('bold').setBackground('#d9e9ff');
  sh.setFrozenRows(1);
  return sh;
}

function getSheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getObjects_(sheetName) {
  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(function(r){ return r.join('') !== ''; }).map(function(row){
    const obj = {};
    headers.forEach(function(h, i){ obj[h] = row[i]; });
    return obj;
  });
}

function upsertById_(sheet, obj, idField) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf(idField);
  const values = headers.map(function(h){ return obj[h] !== undefined ? obj[h] : ''; });
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === obj[idField]) { foundRow = i + 1; break; }
  }
  if (foundRow > -1) {
    const oldCreated = sheet.getRange(foundRow, headers.indexOf('CREATED_AT') + 1).getValue();
    if (headers.indexOf('CREATED_AT') > -1 && !obj.CREATED_AT) values[headers.indexOf('CREATED_AT')] = oldCreated;
    sheet.getRange(foundRow, 1, 1, headers.length).setValues([values]);
  } else {
    sheet.appendRow(values);
  }
}

function findUserByUsername_(username) {
  const rows = getObjects_(SHEETS.USERS);
  return rows.find(function(r){ return String(r.USERNAME).toLowerCase() === String(username).toLowerCase(); }) || null;
}

function hashPassword_(password) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return raw.map(function(b){
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function makeId_(prefix) {
  return prefix + '-' + Utilities.getUuid().split('-')[0].toUpperCase();
}

function setSetting_(key, value) {
  const sh = getSheet_(SHEETS.SETTINGS);
  const values = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === key) { row = i + 1; break; }
  }
  if (row === -1) sh.appendRow([key, value]);
  else sh.getRange(row, 2).setValue(value);
}

function getSettingsMap_() {
  const rows = getObjects_(SHEETS.SETTINGS);
  const map = {};
  rows.forEach(function(r){ map[r.KEY] = r.VALUE; });
  return map;
}

function logAction_(username, action, detail, status) {
  getSheet_(SHEETS.LOGS).appendRow([makeId_('LOG'), new Date(), username, action, detail, status]);
}

function fail_(message) {
  return { ok: false, message: message };
}

function generateNoSpd_() {
  const sh = getSheet_(SHEETS.SPD);
  const count = Math.max(sh.getLastRow() - 1, 0) + 1;
  const now = new Date();
  return 'SPD/' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy') + '/' + ('000' + count).slice(-4);
}

function countSpdForEmployeeOnDate_(pegawaiId, tglBerangkat, excludeId) {
  const rows = getObjects_(SHEETS.SPD);
  return rows.filter(function(r){
    return r.PEGAWAI_ID === pegawaiId && String(r.TGL_BERANGKAT) === String(tglBerangkat) && r.ID !== excludeId;
  }).length;
}

function refreshEmployeeTripTotal_(pegawaiId) {
  const total = getObjects_(SHEETS.SPD).filter(function(r){ return r.PEGAWAI_ID === pegawaiId; }).length;
  const sh = getSheet_(SHEETS.EMPLOYEES);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf('ID');
  const totalIndex = headers.indexOf('TOTAL_PERJALANAN');
  const updatedIndex = headers.indexOf('UPDATED_AT');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === pegawaiId) {
      sh.getRange(i + 1, totalIndex + 1).setValue(total);
      sh.getRange(i + 1, updatedIndex + 1).setValue(new Date());
      break;
    }
  }
}
