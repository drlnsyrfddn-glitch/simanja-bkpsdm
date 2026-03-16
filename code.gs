/************************************************************
 * BKPSDM Booking System — Code.gs (FULL)
 * Sheets required:
 * - appointments: Timestamp | BookingCode | Nama | NIP | OPD | HP | Email | Layanan | Tanggal | Slot | Keperluan | Status
 * - services: Layanan | StartTime | EndTime | Aktif
 * - pic_users: Layanan | NamaPIC | Username | Password | Role | Aktif
 * - settings: Key | Value  (adminKey stored here)
 ************************************************************/

const SPREADSHEET_ID = "1isp6yn_0N2yeXCNGjmLYVhahKavw_zGB2hyoPMhEmlk"; // kosong jika script bound ke spreadsheet yang sama
const TZ = "Asia/Jakarta";

/* =========================
   Core WebApp Handlers
========================= */
function doGet() {
  return ContentService.createTextOutput("BKPSDM WebApp is running. Use POST JSON.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    const data = JSON.parse(raw || "{}");
    const action = (data.action || "").trim();

    switch (action) {
      case "create_booking":        return output_(create_booking_(data));
      case "cek_booking":           return output_(cek_booking_(data));

      case "get_services":          return output_(get_services_public_());
      case "get_slots":             return output_(get_slots_(data));

      // PIC
      case "pic_login":             return output_(pic_login_(data));
      case "pic_update_schedule":   return output_(pic_update_schedule_(data));
      case "pic_list":              return output_(pic_list_(data));
      case "pic_update_status":     return output_(update_booking_status_(data, "PIC"));

      // ADMIN
      case "admin_services_all":        return output_(admin_services_all_(data));
      case "admin_pic_create":          return output_(admin_pic_create_(data));
      case "admin_pic_list":            return output_(admin_pic_list_(data));
      case "admin_pic_update":          return output_(admin_pic_update_(data));
      case "admin_pic_reset_password":  return output_(admin_pic_reset_password_(data));
      case "admin_pic_delete":          return output_(admin_pic_delete_(data));

      case "admin_booking_list":        return output_(admin_booking_list_(data));

      case "admin_selesai_report": return output_(admin_selesai_report_(data));

      case "admin_booking_update_status": return output_(update_booking_status_(data, "ADMIN"));

      default:
        return output_({ ok: false, message: "Unknown action: " + action });
    }
  } catch (err) {
    return output_({ ok: false, message: "Server error: " + err.message });
  }
}

/* =========================
   Utilities
========================= */
function ss_() {
  return SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error(`Sheet "${name}" tidak ditemukan.`);
  return sh;
}

function output_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function now_() {
  return new Date();
}

function fmtDate_(d, pattern) {
  return Utilities.formatDate(d, TZ, pattern);
}

function normalizeISODate_(v) {
  if (v === null || v === undefined || v === "") return "";
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) {
    return fmtDate_(v, "yyyy-MM-dd");
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  // dd/mm/yyyy
  const m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m1) return `${m1[3]}-${m1[2]}-${m1[1]}`;

  // dd-mm-yyyy
  const m2 = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (m2) return `${m2[3]}-${m2[2]}-${m2[1]}`;

  const d = new Date(s);
  if (!isNaN(d)) return fmtDate_(d, "yyyy-MM-dd");
  return s;
}

function normalizeTimeHHMM_(v) {
  if (v === null || v === undefined || v === "") return "";
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) {
    // NOTE: time-only Date from Sheets can be weird historically; prefer display values when possible
    return Utilities.formatDate(v, "GMT+7", "HH:mm"); // fixed offset to avoid historical tz shifts
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) return `${("0"+m[1]).slice(-2)}:${m[2]}`;
  return s;
}

function parseISODateObj_(iso) {
  if (!iso) return null;
  const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  d.setHours(0,0,0,0);
  return d;
}

/* =========================
   Settings / Admin auth
========================= */
function getAdminKey_() {
  const sh = sheet_("settings");
  const values = sh.getDataRange().getValues();
  for (let i=1; i<values.length; i++) {
    const key = String(values[i][0] || "").trim();
    const val = String(values[i][1] || "").trim();
    if (key === "adminKey") return val;
  }
  return "";
}

function isAdmin_(adminKey) {
  const real = getAdminKey_();
  return real && String(adminKey || "").trim() === real;
}

/* =========================
   SERVICES
========================= */
function readServices_() {
  const sh = sheet_("services");
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const head = values[0].map(String);
  const idx = (name) => head.indexOf(name);

  const iLayanan = idx("Layanan");
  const iStart = idx("StartTime");
  const iEnd = idx("EndTime");
  const iAktif = idx("Aktif");

  const out = [];
  for (let r=1; r<values.length; r++) {
    const row = values[r];
    const layanan = String(row[iLayanan] || "").trim();
    if (!layanan) continue;
    out.push({
      layanan,
      start_time: normalizeTimeHHMM_(row[iStart]),
      end_time: normalizeTimeHHMM_(row[iEnd]),
      aktif: String(row[iAktif] || "YES").trim().toUpperCase() || "YES"
    });
  }
  return out;
}

function get_services_public_() {
  const services = readServices_().filter(s => s.aktif === "YES");
  return { ok:true, services };
}

function admin_services_all_(data) {
  if (!isAdmin_(data.adminKey)) return { ok:false, message:"Unauthorized" };
  const services = readServices_();
  return { ok:true, services };
}

function updateServiceSchedule_(layanan, start_time, end_time, aktif) {
  const sh = sheet_("services");
  const values = sh.getDataRange().getValues();
  const head = values[0].map(String);
  const idx = (name) => head.indexOf(name);

  const iLayanan = idx("Layanan");
  const iStart = idx("StartTime");
  const iEnd = idx("EndTime");
  const iAktif = idx("Aktif");

  for (let r=1; r<values.length; r++) {
    const row = values[r];
    if (String(row[iLayanan]||"").trim() === layanan) {
      sh.getRange(r+1, iStart+1).setValue(start_time);
      sh.getRange(r+1, iEnd+1).setValue(end_time);
      sh.getRange(r+1, iAktif+1).setValue(aktif);
      return true;
    }
  }
  return false;
}

/* =========================
   SLOTS generator (30 min)
========================= */
function makeSlots_(startHHMM, endHHMM) {
  const [sh, sm] = startHHMM.split(":").map(Number);
  const [eh, em] = endHHMM.split(":").map(Number);

  const start = sh * 60 + sm;
  const end = eh * 60 + em;

  const slots = [];
  for (let t = start; t < end; t += 30) {
    const hh = Math.floor(t/60);
    const mm = t%60;
    slots.push(`${("0"+hh).slice(-2)}:${("0"+mm).slice(-2)}`);
  }
  return slots;
}

function get_slots_(data) {
  const layanan = String(data.layanan || "").trim();
  const tanggal = normalizeISODate_(data.tanggal || "");
  if (!layanan) return { ok:false, message:"layanan wajib" };
  if (!tanggal) return { ok:false, message:"tanggal wajib" };

  const service = readServices_().find(s => s.layanan === layanan);
  if (!service) return { ok:false, message:"Layanan tidak ditemukan" };
  if (service.aktif !== "YES") return { ok:false, message:"Layanan sedang nonaktif" };

  const slots = makeSlots_(service.start_time, service.end_time);
  return { ok:true, slots };
}

/* =========================
   BOOKING — appointments
========================= */
function ensureAppointmentsHeader_() {
  const sh = sheet_("appointments");
  const values = sh.getDataRange().getValues();
  if (values.length === 0) return;

  const required = ["Timestamp","BookingCode","Nama","NIP","OPD","HP","Email","Layanan","Tanggal","Slot","Keperluan","Status"];
  const head = values[0].map(String);
  const missing = required.filter(h => head.indexOf(h) < 0);
  if (missing.length) {
    throw new Error("Header appointments belum sesuai. Missing: " + missing.join(", "));
  }
}

function newBookingCode_(tanggalISO) {
  const rnd = Math.floor(1000 + Math.random()*9000);
  const datePart = tanggalISO.replaceAll("-","");
  return `BKPSDM-${datePart}-${rnd}`;
}

function create_booking_(data) {
  ensureAppointmentsHeader_();

  const nama = String(data.nama||"").trim();
  const nip = String(data.nip||"").trim();
  const opd = String(data.opd||"").trim();
  const hp = String(data.hp||"").trim();
  const email = String(data.email||"").trim();
  const layanan = String(data.layanan||"").trim();
  const tanggal = normalizeISODate_(data.tanggal||"");
  const slot = normalizeTimeHHMM_(data.slot||"");
  const keperluan = String(data.keperluan||"").trim();

  if (!nama) return { ok:false, message:'Field "nama" wajib diisi.' };
  if (!nip) return { ok:false, message:'Field "nip" wajib diisi.' };
  if (!opd) return { ok:false, message:'Field "opd" wajib diisi.' };
  if (!hp) return { ok:false, message:'Field "hp" wajib diisi.' };
  if (!layanan) return { ok:false, message:'Field "layanan" wajib diisi.' };
  if (!tanggal) return { ok:false, message:'Field "tanggal" wajib diisi.' };
  if (!slot) return { ok:false, message:'Field "slot" wajib diisi.' };

  const service = readServices_().find(s => s.layanan === layanan);
  if (!service) return { ok:false, message:"Layanan tidak ditemukan." };
  if (service.aktif !== "YES") return { ok:false, message:"Layanan sedang nonaktif." };

  const allowedSlots = makeSlots_(service.start_time, service.end_time);
  if (allowedSlots.indexOf(slot) < 0) return { ok:false, message:"Slot tidak valid untuk layanan ini." };

  const todayISO = fmtDate_(now_(), "yyyy-MM-dd");
  if (tanggal < todayISO) return { ok:false, message:"Tidak bisa booking tanggal yang sudah lewat." };

  const d = parseISODateObj_(tanggal);
  const day = d.getDay();
  if (day === 0 || day === 6) return { ok:false, message:"Tidak melayani weekend/libur." };

  const sh = sheet_("appointments");
  const values = sh.getDataRange().getValues();
  const head = values[0].map(String);
  const idx = (name) => head.indexOf(name);

  const iTanggal = idx("Tanggal");
  const iSlot = idx("Slot");
  const iLayanan = idx("Layanan");
  const iStatus = idx("Status");

  for (let r=1; r<values.length; r++) {
    const row = values[r];
    const tISO = normalizeISODate_(row[iTanggal]);
    const sHH = normalizeTimeHHMM_(row[iSlot]);
    const lay = String(row[iLayanan]||"").trim();
    const st = String(row[iStatus]||"Booked").trim();
    if (tISO === tanggal && sHH === slot && lay === layanan && st !== "Batal") {
      return { ok:false, message:"Slot sudah terisi. Pilih slot lain." };
    }
  }

  const bookingCode = newBookingCode_(tanggal);
  const ts = fmtDate_(now_(), "M/d/yyyy HH:mm:ss");

  const rowOut = [];
  rowOut[idx("Timestamp")] = ts;
  rowOut[idx("BookingCode")] = bookingCode;
  rowOut[idx("Nama")] = nama;
  rowOut[idx("NIP")] = nip;
  rowOut[idx("OPD")] = opd;
  rowOut[idx("HP")] = hp;
  rowOut[idx("Email")] = email;
  rowOut[idx("Layanan")] = layanan;
  rowOut[idx("Tanggal")] = tanggal; // store ISO
  rowOut[idx("Slot")] = slot;       // store HH:mm
  rowOut[idx("Keperluan")] = keperluan;
  rowOut[idx("Status")] = "Booked";

  const totalCols = head.length;
  const finalRow = new Array(totalCols).fill("");
  for (let i=0; i<totalCols; i++) finalRow[i] = rowOut[i] || "";
  sh.appendRow(finalRow);

  return { ok:true, bookingCode, message:"Booking berhasil." };
}

/* ✅ FIX: cek_booking baca tanggal/slot dari DISPLAY VALUES */
function cek_booking_(data) {
  ensureAppointmentsHeader_();
  const nip = String(data.nip||"").trim();
  if (!nip) return { ok:false, message:"NIP wajib diisi." };

  const sh = sheet_("appointments");
  const values = sh.getDataRange().getValues();
  const disp = sh.getDataRange().getDisplayValues();
  const head = values[0].map(String);
  const idx = (name) => head.indexOf(name);

  const iTanggal = idx("Tanggal");
  const iSlot = idx("Slot");
  const iNip = idx("NIP");
  const iStatus = idx("Status");

  const todayISO = fmtDate_(now_(), "yyyy-MM-dd");

  const out = [];
  for (let r=1; r<values.length; r++) {
    const rowNip = String(values[r][iNip]||"").trim();
    if (rowNip !== nip) continue;

    const tanggalISO = normalizeISODate_(disp[r][iTanggal] || values[r][iTanggal]);
    const slotHHMM = normalizeTimeHHMM_(disp[r][iSlot] || values[r][iSlot]);
    const status = String(values[r][iStatus]||"Booked").trim();

    if (tanggalISO < todayISO) continue;
    if (status === "Batal") continue;

    out.push({
      bookingCode: String(values[r][idx("BookingCode")]||""),
      nama: String(values[r][idx("Nama")]||""),
      nip: rowNip,
      layanan: String(values[r][idx("Layanan")]||""),
      tanggal: tanggalISO,
      slot: slotHHMM,
      status
    });
  }

  return { ok:true, nip, rows: out, count: out.length };
}

/* =========================
   PIC USERS (multi-PIC)
========================= */
function readPicUsers_() {
  const sh = sheet_("pic_users");
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const head = values[0].map(String);
  const idx = (name) => head.indexOf(name);

  const out = [];
  for (let r=1; r<values.length; r++) {
    const row = values[r];
    const layanan = String(row[idx("Layanan")]||"").trim();
    const username = String(row[idx("Username")]||"").trim();
    if (!layanan || !username) continue;

    out.push({
      rowIndex: r+1,
      layanan,
      nama: String(row[idx("NamaPIC")]||"").trim(),
      username,
      password: String(row[idx("Password")]||"").trim(),
      role: String(row[idx("Role")]||"PIC").trim(),
      aktif: String(row[idx("Aktif")]||"YES").trim().toUpperCase()
    });
  }
  return out;
}

function pic_login_(data) {
  const layanan = String(data.layanan||"").trim();
  const username = String(data.username||"").trim();
  const password = String(data.password||"").trim();

  if (!layanan) return { ok:false, message:"Layanan wajib dipilih." };
  if (!username) return { ok:false, message:"Username wajib diisi." };
  if (!password) return { ok:false, message:"Password wajib diisi." };

  const users = readPicUsers_();
  const u = users.find(x => x.layanan === layanan && x.username === username);
  if (!u) return { ok:false, message:"User tidak ditemukan untuk layanan ini." };
  if (u.aktif !== "YES") return { ok:false, message:"User nonaktif." };
  if (u.password !== password) return { ok:false, message:"Password salah." };

  const svc = readServices_().find(s => s.layanan === layanan);
  if (!svc) return { ok:false, message:"Layanan tidak ditemukan di sheet services." };

  return {
    ok:true,
    layanan,
    nama_pic: u.nama || u.username,
    role: u.role,
    start_time: svc.start_time,
    end_time: svc.end_time,
    aktif: svc.aktif
  };
}

function pic_update_schedule_(data) {
  const layanan = String(data.layanan||"").trim();
  const start_time = normalizeTimeHHMM_(data.start_time||"");
  const end_time = normalizeTimeHHMM_(data.end_time||"");
  const aktif = String(data.aktif||"YES").trim().toUpperCase();

  if (!layanan) return { ok:false, message:"Layanan wajib." };
  if (!start_time || !end_time) return { ok:false, message:"Jam mulai & selesai wajib." };

  const ok = updateServiceSchedule_(layanan, start_time, end_time, aktif);
  return ok ? { ok:true } : { ok:false, message:"Layanan tidak ditemukan." };
}

/* ✅ FIX: pic_list baca tanggal/slot dari DISPLAY VALUES */
function pic_list_(data) {
  ensureAppointmentsHeader_();
  const layanan = String(data.layanan||"").trim();
  const tanggal = normalizeISODate_(data.tanggal||"");
  if (!layanan) return { ok:false, message:"Layanan wajib." };
  if (!tanggal) return { ok:false, message:"Tanggal wajib." };

  const sh = sheet_("appointments");
  const values = sh.getDataRange().getValues();
  const disp = sh.getDataRange().getDisplayValues();
  const head = values[0].map(String);
  const idx = (name) => head.indexOf(name);

  const out = [];
  for (let r=1; r<values.length; r++) {
    const lay = String(values[r][idx("Layanan")]||"").trim();
    if (lay !== layanan) continue;

    const tISO = normalizeISODate_(disp[r][idx("Tanggal")] || values[r][idx("Tanggal")]);
    if (tISO !== tanggal) continue;

    out.push({
      bookingCode: String(values[r][idx("BookingCode")]||""),
      nama: String(values[r][idx("Nama")]||""),
      nip: String(values[r][idx("NIP")]||""),
      slot: normalizeTimeHHMM_(disp[r][idx("Slot")] || values[r][idx("Slot")]),
      status: String(values[r][idx("Status")]||"Booked").trim()
    });
  }

  out.sort((a,b)=> String(a.slot).localeCompare(String(b.slot)));
  return { ok:true, rows: out };
}

function update_booking_status_(data, actor) {
  ensureAppointmentsHeader_();
  const bookingCode = String(data.bookingCode||"").trim();
  const status = String(data.status||"").trim();
  if (!bookingCode) return { ok:false, message:"bookingCode wajib." };
  if (!status) return { ok:false, message:"status wajib." };

  const sh = sheet_("appointments");
  const values = sh.getDataRange().getValues();
  const head = values[0].map(String);
  const idx = (name) => head.indexOf(name);

  const iBooking = idx("BookingCode");
  const iStatus = idx("Status");

  for (let r=1; r<values.length; r++) {
    if (String(values[r][iBooking]||"").trim() === bookingCode) {
      sh.getRange(r+1, iStatus+1).setValue(status);
      return { ok:true, actor };
    }
  }
  return { ok:false, message:"BookingCode tidak ditemukan." };
}

/* =========================
   ADMIN — PIC CRUD
========================= */
function admin_pic_create_(data) {
  if (!isAdmin_(data.adminKey)) return { ok:false, message:"Unauthorized" };

  const layanan = String(data.layanan||"").trim();
  const nama = String(data.nama||"").trim();
  const username = String(data.username||"").trim();
  const password = String(data.password||"").trim();
  const role = String(data.role||"PIC").trim();
  const aktif = String(data.aktif||"YES").trim().toUpperCase();

  if (!layanan) return { ok:false, message:"Layanan wajib." };
  if (!username) return { ok:false, message:"Username wajib." };
  if (!password) return { ok:false, message:"Password wajib." };

  const sh = sheet_("pic_users");
  const existing = readPicUsers_();
  if (existing.find(u => u.layanan===layanan && u.username===username)) {
    return { ok:false, message:"User sudah ada untuk layanan ini." };
  }

  sh.appendRow([layanan, nama, username, password, role, aktif]);
  return { ok:true };
}

function admin_pic_list_(data) {
  if (!isAdmin_(data.adminKey)) return { ok:false, message:"Unauthorized" };
  const users = readPicUsers_().map(u => ({
    layanan: u.layanan,
    nama: u.nama,
    username: u.username,
    role: u.role,
    aktif: u.aktif
  }));
  return { ok:true, users };
}

function admin_pic_update_(data) {
  if (!isAdmin_(data.adminKey)) return { ok:false, message:"Unauthorized" };

  const layanan = String(data.layanan||"").trim();
  const username = String(data.username||"").trim();
  const role = String(data.role||"PIC").trim();
  const aktif = String(data.aktif||"YES").trim().toUpperCase();

  if (!layanan || !username) return { ok:false, message:"layanan & username wajib." };

  const sh = sheet_("pic_users");
  const values = sh.getDataRange().getValues();
  const head = values[0].map(String);
  const idx = (name) => head.indexOf(name);

  const iL = idx("Layanan");
  const iU = idx("Username");
  const iR = idx("Role");
  const iA = idx("Aktif");

  for (let r=1; r<values.length; r++) {
    if (String(values[r][iL]||"").trim()===layanan && String(values[r][iU]||"").trim()===username) {
      sh.getRange(r+1, iR+1).setValue(role);
      sh.getRange(r+1, iA+1).setValue(aktif);
      return { ok:true };
    }
  }
  return { ok:false, message:"PIC tidak ditemukan." };
}

function admin_pic_reset_password_(data) {
  if (!isAdmin_(data.adminKey)) return { ok:false, message:"Unauthorized" };

  const layanan = String(data.layanan||"").trim();
  const username = String(data.username||"").trim();
  const newPassword = String(data.newPassword||"").trim();

  if (!layanan || !username || !newPassword) return { ok:false, message:"Data kurang." };

  const sh = sheet_("pic_users");
  const values = sh.getDataRange().getValues();
  const head = values[0].map(String);
  const idx = (name) => head.indexOf(name);

  const iL = idx("Layanan");
  const iU = idx("Username");
  const iP = idx("Password");

  for (let r=1; r<values.length; r++) {
    if (String(values[r][iL]||"").trim()===layanan && String(values[r][iU]||"").trim()===username) {
      sh.getRange(r+1, iP+1).setValue(newPassword);
      return { ok:true };
    }
  }
  return { ok:false, message:"PIC tidak ditemukan." };
}

function admin_pic_delete_(data) {
  if (!isAdmin_(data.adminKey)) return { ok:false, message:"Unauthorized" };

  const layanan = String(data.layanan||"").trim();
  const username = String(data.username||"").trim();
  if (!layanan || !username) return { ok:false, message:"Data kurang." };

  const sh = sheet_("pic_users");
  const values = sh.getDataRange().getValues();
  const head = values[0].map(String);
  const idx = (name) => head.indexOf(name);

  const iL = idx("Layanan");
  const iU = idx("Username");

  for (let r=1; r<values.length; r++) {
    if (String(values[r][iL]||"").trim()===layanan && String(values[r][iU]||"").trim()===username) {
      sh.deleteRow(r+1);
      return { ok:true };
    }
  }
  return { ok:false, message:"PIC tidak ditemukan." };
}

/* =========================
   ADMIN — Booking monitoring (range)
   ✅ FIX: gunakan DISPLAY VALUES untuk Tanggal & Slot
========================= */
function admin_booking_list_(data) {
  if (!isAdmin_(data.adminKey)) return { ok:false, message:"Unauthorized" };
  ensureAppointmentsHeader_();

  const date_from = normalizeISODate_(data.date_from || "");
  const date_to = normalizeISODate_(data.date_to || "");
  const layananFilter = String(data.layanan || "ALL").trim();
  const statusFilter = String(data.status || "ALL").trim();
  const nipFilter = String(data.nip || "").trim();

  const fromObj = date_from ? parseISODateObj_(date_from) : null;
  const toObj = date_to ? parseISODateObj_(date_to) : null;
  if (toObj) toObj.setHours(23,59,59,999);

  const sh = sheet_("appointments");
  const range = sh.getDataRange();
  const values = range.getValues();
  const disp = range.getDisplayValues();

  const head = values[0].map(String);
  const idx = (name) => head.indexOf(name);

  const out = [];
  for (let r=1; r<values.length; r++) {
    const bookingCode = String(values[r][idx("BookingCode")]||"").trim();
    if (!bookingCode) continue;

    const tanggalISO = normalizeISODate_(disp[r][idx("Tanggal")] || values[r][idx("Tanggal")]);
    const slotHHMM = normalizeTimeHHMM_(disp[r][idx("Slot")] || values[r][idx("Slot")]);

    const tObj = parseISODateObj_(tanggalISO);

    if (fromObj && (!tObj || tObj < fromObj)) continue;
    if (toObj && (!tObj || tObj > toObj)) continue;

    const layanan = String(values[r][idx("Layanan")]||"").trim();
    if (layananFilter !== "ALL" && layanan !== layananFilter) continue;

    const status = String(values[r][idx("Status")]||"Booked").trim();
    if (statusFilter !== "ALL" && status !== statusFilter) continue;

    const nip = String(values[r][idx("NIP")]||"").trim();
    if (nipFilter && !nip.includes(nipFilter)) continue;

    out.push({
      bookingCode,
      tanggal: tanggalISO,
      slot: slotHHMM,
      layanan,
      nama: String(values[r][idx("Nama")]||""),
      nip,
      status
    });
  }

  out.sort((a,b)=>{
    if (a.tanggal !== b.tanggal) return a.tanggal.localeCompare(b.tanggal);
    return String(a.slot).localeCompare(String(b.slot));
  });

  return { ok:true, rows: out };
}

function admin_selesai_report_(data) {
  if (!isAdmin_(data.adminKey)) return { ok: false, message: "Unauthorized" };
  
  const date_from = normalizeISODate_(data.date_from || "");
  const date_to = normalizeISODate_(data.date_to || "");
  const layananFilter = String(data.layanan || "ALL").trim();

  const sh = sheet_("appointments");
  const values = sh.getDataRange().getValues();
  const disp = sh.getDataRange().getDisplayValues();
  const head = values[0].map(String);
  const idx = (name) => head.indexOf(name);

  const fromObj = date_from ? parseISODateObj_(date_from) : null;
  const toObj = date_to ? parseISODateObj_(date_to) : null;
  if (toObj) toObj.setHours(23, 59, 59, 999);

  let totalSelesai = 0;
  let rows = [];
  let svcMap = {};

  for (let r = 1; r < values.length; r++) {
    const status = String(values[r][idx("Status")] || "").trim();
    if (status !== "Selesai") continue; // Hanya ambil yang selesai

    const tISO = normalizeISODate_(disp[r][idx("Tanggal")] || values[r][idx("Tanggal")]);
    const tObj = parseISODateObj_(tISO);

    if (fromObj && (!tObj || tObj < fromObj)) continue;
    if (toObj && (!tObj || tObj > toObj)) continue;

    const layanan = String(values[r][idx("Layanan")] || "").trim();
    if (layananFilter !== "ALL" && layanan !== layananFilter) continue;

    totalSelesai++;
    svcMap[layanan] = (svcMap[layanan] || 0) + 1;

    rows.push({
      bookingCode: String(values[r][idx("BookingCode")] || ""),
      tanggal: tISO,
      layanan: layanan,
      nama: String(values[r][idx("Nama")] || ""),
      nip: String(values[r][idx("NIP")] || ""),
      keperluan: String(values[r][idx("Keperluan")] || "")
    });
  }

  // Hitung Rekap Per Layanan
  let by_service = [];
  let topSvc = { layanan: "-", count: 0 };

  for (let key in svcMap) {
    let count = svcMap[key];
    if (count > topSvc.count) topSvc = { layanan: key, count: count };
    
    by_service.push({
      layanan: key,
      count: count,
      pct: totalSelesai > 0 ? (count / totalSelesai) * 100 : 0
    });
  }

  return { 
    ok: true, 
    total: totalSelesai, 
    top_service: topSvc, 
    by_service: by_service, 
    rows: rows 
  };
}
