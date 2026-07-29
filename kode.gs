/**
 * @OnlyCurrentDoc
 */

const SPREADSHEET_ID = '1nLdFSrK1dvfrYu_db8tGEFVYuBILmNL-3X1eC1nu8YM'; 
const ADMIN_EMAILS = ['byu.project.comp@gmail.com', 'ajeefes@gmail.com']; 
const FOLDER_FOTO_ID = '13gHdT1Gt7S4uCeMskbIOTUy9FhHYzVBw'; 

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Laporan Harian Sekolah Rakyat')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getEmailUser() {
  try {
    const email = Session.getActiveUser().getEmail();
    return email ? email.toLowerCase().trim() : "user_anonim";
  } catch(e) {
    return "user_anonim";
  }
}

function getUserStats(email) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Data_Kegiatan');
  
  let emptyStats = { jamKerja: '0', laporan: '0', entri: '0', foto: '0' };
  let result = { bulanIni: { ...emptyStats }, semua: { ...emptyStats } };
  
  if (!sheet || sheet.getLastRow() < 2) return result;
  
  const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 12).getValues();
  let lapSetSemua = new Set(), lapSetBulanIni = new Set();
  let totEntriSemua = 0, totEntriBulan = 0;
  let totDiffSemua = 0, totDiffBulan = 0;
  let totFotoSemua = 0, totFotoBulan = 0;
  
  let now = new Date();
  let currentMonth = now.getMonth();
  let currentYear = now.getFullYear();
  
  for(let row of data) {
    if(String(row[1]).toLowerCase().trim() === email) {
      let isBulanIni = false;
      if (row[7]) {
          let tglData = new Date(row[7]);
          if (tglData.getMonth() === currentMonth && tglData.getFullYear() === currentYear) {
              isBulanIni = true;
          }
      }
      
      lapSetSemua.add(row[0]); 
      totEntriSemua++;
      if (row[10] && String(row[10]).trim() !== '') totFotoSemua++; 
      
      let diffMs = 0;
      try {
         const startStr = row[8] instanceof Date ? Utilities.formatDate(row[8], "GMT+7", "HH:mm") : String(row[8]);
         const endStr = row[9] instanceof Date ? Utilities.formatDate(row[9], "GMT+7", "HH:mm") : String(row[9]);
         if (startStr && endStr && startStr.includes(':') && endStr.includes(':')) {
             const t1 = new Date("1970-01-01T" + startStr + ":00Z");
             const t2 = new Date("1970-01-01T" + endStr + ":00Z");
             if(t2 > t1) diffMs = (t2 - t1);
         }
      } catch(e){}
      
      totDiffSemua += diffMs;
      
      if(isBulanIni) {
          lapSetBulanIni.add(row[0]);
          totEntriBulan++;
          if (row[10] && String(row[10]).trim() !== '') totFotoBulan++;
          totDiffBulan += diffMs;
      }
    }
  }
  
  result.semua = { jamKerja: Math.floor(totDiffSemua / 3600000).toString(), laporan: lapSetSemua.size.toString(), entri: totEntriSemua.toString(), foto: totFotoSemua.toString() };
  result.bulanIni = { jamKerja: Math.floor(totDiffBulan / 3600000).toString(), laporan: lapSetBulanIni.size.toString(), entri: totEntriBulan.toString(), foto: totFotoBulan.toString() };
  
  return result;
}

function getPengaturanSurat() {
  const email = Session.getActiveUser().getEmail();
  const ADMIN_EMAIL = ADMIN_EMAILS[0]; 
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  let sheetKop = ss.getSheetByName("Pengaturan_Surat");
  if (!sheetKop) sheetKop = ss.insertSheet("Pengaturan_Surat");
  
  let sheetTtd = ss.getSheetByName("Pengaturan_TTD");
  if (!sheetTtd) sheetTtd = ss.insertSheet("Pengaturan_TTD");

  let dataKop = getConfigByEmail(sheetKop, email, 'kop') || getConfigByEmail(sheetKop, ADMIN_EMAIL, 'kop') || {};
  let dataTtd = getConfigByEmail(sheetTtd, email, 'ttd') || {};
  
  return { kop: dataKop, ttd: dataTtd };
}

function getConfigByEmail(sheet, email, type) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  const data = sheet.getDataRange().getValues();
  const target = String(email).toLowerCase().trim();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === target) {
      if (type === 'kop') {
        return {
          logoKiri: getFileBase64FromDrive(data[i][1]), 
          logoKanan: getFileBase64FromDrive(data[i][2]), 
          judul: data[i][3],
          fontJudul: data[i][4], ukuranLogo: data[i][5], fontTabel: data[i][6],
          fotoTabel: data[i][7], barisKop: data[i][8]
        };
      } else if (type === 'ttd') {
        return {
          kota: data[i][1], jabatanTtd: data[i][2], ukuranTtd: data[i][3],
          namaTtd: data[i][4], nipTtd: data[i][5], 
          tandaTangan: getFileBase64FromDrive(data[i][6])
        };
      }
    }
  }
  return null;
}

function getFileBase64FromDrive(fileUrlOrId) {
  if (!fileUrlOrId) return "";
  if (String(fileUrlOrId).length > 200) return fileUrlOrId;
  
  let fileId = fileUrlOrId;
  if (fileId.includes('drive.google.com')) {
    const match = fileId.match(/id=([^&]+)/) || fileId.match(/\/d\/([^\/]+)/);
    if (match && match[1]) fileId = match[1];
  }
  
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return "data:image/jpeg;base64," + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    return "";
  }
}

// ----------------- SISTEM CEK STATUS & TOKEN (1 BULAN) -----------------
function checkUserStatus() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheetUser = ss.getSheetByName('Users');
  let dataUsers = sheetUser.getDataRange().getValues();
  let headers = dataUsers[0].map(h => String(h).trim());
  
  let userFoundIndex = -1, roleUser = "Tamu", isActive = true, expiryDate = null;
  let rawFotoProfil = "";

  for (let i = 1; i < dataUsers.length; i++) {
    let rowEmail = dataUsers[i][headers.indexOf('Email')];
    if (rowEmail && String(rowEmail).toLowerCase().trim() === email.toLowerCase().trim()) {
      userFoundIndex = i + 1;
      isActive = dataUsers[i][headers.indexOf('IsActive')];
      
      let roleIndex = headers.findIndex(h => String(h).trim().toLowerCase() === 'role');
      if (roleIndex === -1) roleIndex = 10;
      let dbRole = dataUsers[i][roleIndex];
      if (dbRole && String(dbRole).trim() !== "") roleUser = String(dbRole).trim();

      let expiryIndex = headers.indexOf('ExpiryDate') > -1 ? headers.indexOf('ExpiryDate') : headers.indexOf('Expiry');
      if (expiryIndex > -1) expiryDate = dataUsers[i][expiryIndex];

      // Ambil data foto profil dari kolom
      let fotoIndex = headers.indexOf('FotoProfil');
      if (fotoIndex > -1) {
          rawFotoProfil = dataUsers[i][fotoIndex] || "";
      }
      break;
    }
  }

  if (userFoundIndex === -1) return { status: "REQUIRE_TOKEN", message: "Akun Anda belum terdaftar. Silakan masukkan token akses." };
  if (expiryDate && new Date() > new Date(expiryDate)) return { status: "EXPIRED_TOKEN", message: "Masa aktif habis." };
  if (isActive === false || isActive === "FALSE") return { status: "SUSPENDED", message: "Akun dinonaktifkan." };

  const liveStats = getUserStats(email.toLowerCase().trim());
  let userData = { 
    nama: dataUsers[userFoundIndex - 1][headers.indexOf('Nama')] || "User", 
    nip: dataUsers[userFoundIndex - 1][headers.indexOf('NIP')] || "-",
    jabatan: dataUsers[userFoundIndex - 1][headers.indexOf('Jabatan')] || "-", 
    instansi: dataUsers[userFoundIndex - 1][headers.indexOf('Instansi')] || "-", 
    email: email, 
    fotoProfil: getFileBase64FromDrive(rawFotoProfil) || rawFotoProfil, // Ubah ID Drive ke Base64/URL
    statsBulanIni: liveStats.bulanIni,
    statsSemua: liveStats.semua
  };

  return { status: "ACTIVE", isAdmin: (roleUser === "Admin"), role: roleUser, userData: userData };
}

function verifyAndActivateToken(token) {
  const email = getEmailUser();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  let tokenSheet = ss.getSheetByName('Tokens');
  if (!tokenSheet) return { success: false, message: "Sistem token tidak ditemukan." };

  const data = tokenSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(token).trim()) {
      if (String(data[i][1]).trim().toLowerCase() === 'unused') {
        tokenSheet.getRange(i + 1, 2).setValue('used');
        tokenSheet.getRange(i + 1, 3).setValue(email);

        let userSheet = ss.getSheetByName('Users');
        let expiry = new Date();
        expiry.setMonth(expiry.getMonth() + 1); // Masa aktif awal 1 bulan
        
        userSheet.appendRow([email, true, expiry, 'User Baru', '-', '-', '-', String(token).trim(), '']);
        return { success: true };
      } else {
        return { success: false, message: 'Token sudah pernah digunakan.' };
      }
    }
  }
  return { success: false, message: 'Token tidak valid / salah ketik.' };
}

function verifyAndRenewToken(token) {
  const email = getEmailUser();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  let tokenSheet = ss.getSheetByName('Tokens');
  if (!tokenSheet) return { success: false, message: "Sistem token tidak ditemukan." };

  const data = tokenSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(token).trim()) {
      if (String(data[i][1]).trim().toLowerCase() === 'unused') {
        tokenSheet.getRange(i + 1, 2).setValue('used');
        tokenSheet.getRange(i + 1, 3).setValue(email);

        let userSheet = ss.getSheetByName('Users');
        let usersData = userSheet.getDataRange().getValues();
        let headers = usersData[0];

        let newExpiry = new Date();
        newExpiry.setMonth(newExpiry.getMonth() + 1); // Perpanjangan 1 bulan ke depan

        for (let j = 1; j < usersData.length; j++) {
          if (String(usersData[j][headers.indexOf('Email')]).toLowerCase().trim() === email) {
            userSheet.getRange(j + 1, headers.indexOf('Expiry') + 1).setValue(newExpiry);
            userSheet.getRange(j + 1, headers.indexOf('Token') + 1).setValue(String(token).trim());
            break;
          }
        }
        return { success: true };
      } else {
        return { success: false, message: 'Token sudah pernah digunakan.' };
      }
    }
  }
  return { success: false, message: 'Token tidak valid / salah ketik.' };
}

function getUnusedToken() {
  if(!ADMIN_EMAILS.includes(getEmailUser())) return null; 
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Tokens');
  if(!sheet) {
    sheet = ss.insertSheet('Tokens');
    sheet.appendRow(['Token', 'Status', 'Used_By']);
  }
  
  const token = Math.random().toString(36).substr(2, 6).toUpperCase();
  sheet.appendRow([token, 'unused', '']);
  return token;
}

function getAllUsersData() {
  if(!ADMIN_EMAILS.includes(getEmailUser())) return []; 
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let users = [];
  
  for(let i=1; i<data.length; i++) {
    let email = data[i][0];
    let stats = getUserStats(email);
    users.push({
      email: email, isActive: data[i][1], 
      expiry: data[i][2] ? Utilities.formatDate(new Date(data[i][2]), "GMT+7", "dd MMM yyyy") : '-',
      nama: data[i][3], nip: data[i][4], jabatan: data[i][5], instansi: data[i][6], token: data[i][7] || '-', telepon: data[i][8] || '', 
      jamKerja: stats.semua.jamKerja, laporan: stats.semua.laporan, entri: stats.semua.entri, foto: stats.semua.foto
    });
  }
  return users;
}

function toggleUserStatus(email, currentStatus) {
  if(!ADMIN_EMAILS.includes(getEmailUser())) return { success: false }; 
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(data[i][0] === email) {
      sheet.getRange(i+1, 2).setValue(!currentStatus);
      return { success: true };
    }
  }
  return { success: false };
}

function saveUserPhone(email, phone) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === email.toLowerCase()) {
      sheet.getRange(i + 1, 9).setValue(phone);
      return true;
    }
  }
  return false;
}

function updateUserProfile(formData) {
  const email = getEmailUser();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const stats = getUserStats(email);

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase().trim() === email) {
      sheet.getRange(i + 1, headers.indexOf('Nama') + 1).setValue(formData.nama);
      sheet.getRange(i + 1, headers.indexOf('NIP') + 1).setValue(formData.nip);
      sheet.getRange(i + 1, headers.indexOf('Jabatan') + 1).setValue(formData.jabatan);
      sheet.getRange(i + 1, headers.indexOf('Instansi') + 1).setValue(formData.instansi);
      
      let fotoIndex = headers.indexOf('FotoProfil');
      let fotoProfilUrl = data[i][fotoIndex] || ""; // Ambil foto lama dulu

      // Jika user mengunggah foto baru (berbentuk base64)
      if (formData.fotoProfil && formData.fotoProfil.startsWith("data:image")) {
          fotoProfilUrl = processBase64Image(formData.fotoProfil, email + "_profile");
          if (fotoIndex > -1) {
              sheet.getRange(i + 1, fotoIndex + 1).setValue(fotoProfilUrl);
          }
      }

      const updatedData = {
        email: email, nama: formData.nama, nip: formData.nip, jabatan: formData.jabatan, instansi: formData.instansi,
        fotoProfil: getFileBase64FromDrive(fotoProfilUrl) || fotoProfilUrl,
        statsBulanIni: stats.bulanIni, statsSemua: stats.semua
      };
      return { success: true, message: 'Profil berhasil diperbarui!', userData: updatedData };
    }
  }
  return { success: false, message: 'Data user tidak ditemukan.' };
}

function savePengaturanSurat(payload) {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const logoKiriUrl = processBase64Image(payload.kop.logoKiri, email + "_logo_kiri");
  const logoKananUrl = processBase64Image(payload.kop.logoKanan, email + "_logo_kanan");
  const ttdUrl = processBase64Image(payload.ttd.tandaTangan, email + "_ttd");

  const sheetKop = ss.getSheetByName("Pengaturan_Surat");
  const dataKop = sheetKop.getDataRange().getValues();
  let rowKop = -1;
  for(let i = 1; i < dataKop.length; i++) {
    if(dataKop[i][0] === email) { rowKop = i + 1; break; }
  }

  const valuesKop = [email, logoKiriUrl, logoKananUrl, payload.kop.judul, 
                     payload.kop.fontJudul, payload.kop.ukuranLogo, payload.kop.fontTabel, 
                     payload.kop.fotoTabel, payload.kop.barisKop];
  if (rowKop > -1) sheetKop.getRange(rowKop, 1, 1, valuesKop.length).setValues([valuesKop]);
  else sheetKop.appendRow(valuesKop);

  const sheetTtd = ss.getSheetByName("Pengaturan_TTD");
  const dataTtd = sheetTtd.getDataRange().getValues();
  let rowTtd = -1;
  for(let i = 1; i < dataTtd.length; i++) {
    if(dataTtd[i][0] === email) { rowTtd = i + 1; break; }
  }

  const valuesTtd = [email, payload.ttd.kota, payload.ttd.jabatanTtd, payload.ttd.ukuranTtd, 
                     payload.ttd.namaTtd, payload.ttd.nipTtd, ttdUrl]; 
                     
  if (rowTtd > -1) {
    sheetTtd.getRange(rowTtd, 1, 1, valuesTtd.length).setValues([valuesTtd]);
  } else {
    sheetTtd.appendRow(valuesTtd);
  }
}

function processBase64Image(base64Str, fileNamePrefix) {
  if (!base64Str) return "";
  if (base64Str.length > 1000) { 
    try {
      let base64Data = base64Str;
      let mimeType = 'image/png'; 
      if (base64Str.includes('base64,')) {
        let parts = base64Str.split('base64,');
        base64Data = parts[1];
        let header = parts[0];
        if (header.includes('data:')) mimeType = header.split('data:')[1].split(';')[0];
      }
      const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileNamePrefix + ".png");
      return saveImageToDrive(blob, fileNamePrefix + ".png");
    } catch(e) {
      throw new Error("Gagal mengupload gambar " + fileNamePrefix + " ke Google Drive. Pesan error: " + e.message);
    }
  }
  return base64Str;
}

function getAllRHKData(jabatan) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Deteksi sheet berdasarkan Role pengguna
  let sheetName = 'RHK_Master_Wali_Asuh';
  if (jabatan === 'Wali Asrama') {
    sheetName = 'RHK_Master_Wali_Asrama';
  }
  
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['ID', 'Tipe', 'RHK_Nama', 'Aktivitas', 'Keterangan_Default']);
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => ({
    id: row[0], tipe: row[1], parentId: row[2],
    nama: (row[1] === 'Global') ? (row[2] || 'RHK Tanpa Nama') : (row[3] || 'Aktivitas Tanpa Nama'),
    aktivitas: row[3] || '', keterangan: row[4] || ''
  }));
}
function getSemuaRHKAdmin() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    asuh: getRhkDataFromSheet(ss, 'RHK_Master_Wali_Asuh'),
    asrama: getRhkDataFromSheet(ss, 'RHK_Master_Wali_Asrama')
  };
}
function getRhkDataFromSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['ID', 'Tipe', 'RHK_Nama', 'Aktivitas', 'Keterangan_Default']);
      return [];
  }
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => ({
    id: row[0], tipe: row[1], parentId: row[2],
    nama: (row[1] === 'Global') ? (row[2] || 'RHK Tanpa Nama') : (row[3] || 'Aktivitas Tanpa Nama'),
    aktivitas: row[3] || '', keterangan: row[4] || ''
  }));
}

function saveAktivitasBaru(rhkId, namaAktivitas, keterangan, jenis) {
  const sheetName = jenis === 'Wali Asrama' ? 'RHK_Master_Wali_Asrama' : 'RHK_Master_Wali_Asuh';
  let sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(sheetName);
  sheet.appendRow([new Date().getTime(), 'Aktivitas', rhkId, namaAktivitas, keterangan]);
  return { success: true };
}

function saveRHKBaru(namaRHK, jenis) {
  const newId = new Date().getTime();
  const sheetName = jenis === 'Wali Asrama' ? 'RHK_Master_Wali_Asrama' : 'RHK_Master_Wali_Asuh';
  let sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(sheetName);
  sheet.appendRow([newId, 'Global', namaRHK, '', '']);
  return { success: true, newId: newId };
}

function updateData(id, namaBaru, ketBaru, jenis) {
  const sheetName = jenis === 'Wali Asrama' ? 'RHK_Master_Wali_Asrama' : 'RHK_Master_Wali_Asuh';
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return { success: false };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      if(data[i][1] === 'Global') sheet.getRange(i + 1, 3).setValue(namaBaru); 
      else {
        sheet.getRange(i + 1, 4).setValue(namaBaru); 
        sheet.getRange(i + 1, 5).setValue(ketBaru || "");
      }
      return { success: true };
    }
  }
  return { success: false, message: 'ID Data tidak ditemukan.' };
}

function hapusData(id, jenis) {
  const sheetName = jenis === 'Wali Asrama' ? 'RHK_Master_Wali_Asrama' : 'RHK_Master_Wali_Asuh';
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return { success: false };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

function getSemuaPendahuluan() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Pendahuluan_Master');
  if (!sheet) {
    sheet = ss.insertSheet('Pendahuluan_Master');
    sheet.appendRow(['ID', 'Nama', 'Isi', 'RHK_Link']);
    return [];
  }
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => ({ id: row[0], nama: row[1], isi: row[2], rhkLink: row[3] }));
}

function savePendahuluan(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Pendahuluan_Master');
  if (data.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] == data.id) {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[data.nama, data.isi, data.rhkLink]]);
        return { success: true };
      }
    }
  } else {
    sheet.appendRow([new Date().getTime(), data.nama, data.isi, data.rhkLink]);
  }
  return { success: true };
}

function hapusPendahuluan(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Pendahuluan_Master');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function hapusLaporanTotal(reportId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Kegiatan');
  const data = sheet.getDataRange().getValues();
  let deleted = false;
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(reportId)) {
      const urlFoto = data[i][10];
      const fileId = extractDriveId(urlFoto);
      if (fileId) {
        try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {}
      }
      sheet.deleteRow(i + 1);
      deleted = true;
    }
  }
  return { success: deleted, message: "Laporan beserta foto terkait berhasil dihapus permanen." };
}

function simpanLaporanBatch(payload) {
  if (!payload || !payload.kegiatan) return { success: false, message: "Data kosong" };
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Data_Kegiatan') || ss.insertSheet('Data_Kegiatan');
  const folder = DriveApp.getFolderById(FOLDER_FOTO_ID);
  
  const reportId = payload.reportId || "LPR-" + new Date().getTime(); 
  let existingRowIndices = [];
  
  if (payload.reportId) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.reportId)) {
        existingRowIndices.push(i + 1);
      }
    }
  }

  payload.kegiatan.forEach((keg, index) => {
    let fileUrl = "";
    if (keg.foto && keg.foto.startsWith("data:image")) {
      try {
        const base64Data = keg.foto.split(',')[1];
        const contentType = keg.foto.split(';')[0].replace('data:image/', '');
        const dataDecode = Utilities.base64Decode(base64Data);
        const fileName = "Foto_" + new Date().getTime() + "_" + index + "." + contentType;
        const blob = Utilities.newBlob(dataDecode, "image/" + contentType, fileName);
        const file = folder.createFile(blob);
        Utilities.sleep(500);
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
        fileUrl = file.getUrl(); 
      } catch (e) { fileUrl = "Error: " + e.toString(); }
    } else {
      fileUrl = keg.foto || ""; 
    }

    // [PERBAIKAN] Menambahkan status 'Draft' pada kolom ke-13 (M)
    let rowData = [
      reportId, payload.email || "", keg.rhkId || "", keg.rhkNama || "", keg.aktivitasId || "", 
      keg.aktivitasNama || "", keg.keterangan || "", keg.tanggal || "", keg.jamMulai || "", keg.jamAkhir || "", 
      fileUrl, (Array.isArray(payload.pendahuluan) ? payload.pendahuluan.join(",") : ""),
      "Draft" 
    ];

    if (existingRowIndices.length > 0) {
      let rowIndex = existingRowIndices.shift();
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
  });

  existingRowIndices.reverse().forEach(rowIndex => {
    sheet.deleteRow(rowIndex);
  });
  
  return { success: true, message: "Laporan berhasil disimpan" };
}

function getRiwayatLaporan() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Data_Kegiatan');
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  const emailUser = getEmailUser();
  const isAdmin = ADMIN_EMAILS.includes(emailUser);
  
  const sheetUsers = ss.getSheetByName('Users');
  let mapNama = {};
  let isKepalaSekolah = false;
  
  if (sheetUsers) {
    const dataU = sheetUsers.getDataRange().getValues();
    const headers = dataU[0];
    const roleIndex = headers.indexOf('Role'); 
    
    for (let i = 1; i < dataU.length; i++) {
      let dbEmail = String(dataU[i][0]).toLowerCase().trim();
      mapNama[dbEmail] = dataU[i][3];
      if (dbEmail === emailUser && roleIndex > -1 && dataU[i][roleIndex] === 'Kepala Sekolah') {
          isKepalaSekolah = true;
      }
    }
  }
  
  const canViewAll = isAdmin || isKepalaSekolah; 
  // [PERBAIKAN] Mengubah pembacaan range dari 12 menjadi 13 kolom
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  const riwayat = {};

  for (let i = 0; i < data.length; i++) {
    let rowEmail = String(data[i][1]).toLowerCase().trim();
    
    if (rowEmail !== emailUser && !canViewAll) continue; 
    
    let reportId = data[i][0]; 
    if (!reportId) continue;
    if (!riwayat[reportId]) {
       let tglObj = data[i][7] ? new Date(data[i][7]) : null;
       let statusDB = data[i][12] || 'Draft'; // [PERBAIKAN] Baca data status dari kolom ke-13
       
       riwayat[reportId] = {
        id: reportId, 
        email: data[i][1], 
        nama: mapNama[rowEmail] || 'User', 
        jumlahKegiatan: 0,
        tanggalMulai: tglObj ? Utilities.formatDate(tglObj, "GMT+7", "dd MMM yy") : '-',
        bulanTahun: tglObj ? Utilities.formatDate(tglObj, "GMT+7", "MMMM yyyy") : '-',
        status: statusDB // [PERBAIKAN] Tampilkan status asli, bukan Selesai
      };
    }
    riwayat[reportId].jumlahKegiatan += 1;
  }
  return Object.values(riwayat).sort((a,b) => b.id.localeCompare(a.id));
}

function getDetailLaporan(reportId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Kegiatan');
  const data = sheet.getDataRange().getValues();
  const searchId = String(reportId).trim();
  const hasil = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === searchId) {
      let rawFoto = data[i][10] || '';
      let pendRaw = data[i][11] || "";
      let base64Foto = getFileBase64FromDrive(rawFoto);

      hasil.push({
        email: data[i][1], 
        rhkId: data[i][2], rhkNama: data[i][3], aktivitasId: data[i][4], aktivitasNama: data[i][5],
        keterangan: data[i][6], 
        tanggal: data[i][7] ? Utilities.formatDate(new Date(data[i][7]), "GMT+7", "yyyy-MM-dd") : '', 
        jamMulai: data[i][8] instanceof Date ? Utilities.formatDate(data[i][8], "GMT+7", "HH:mm") : data[i][8],
        jamAkhir: data[i][9] instanceof Date ? Utilities.formatDate(data[i][9], "GMT+7", "HH:mm") : data[i][9],
        foto: base64Foto, 
        pendahuluan: String(pendRaw)
      });
    }
  }
  return hasil;
}

function getFolderByNameOrCreate(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

function saveFileWithOverwrite(folder, fileName, b64Data) {
  const existingFiles = folder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true); 
  }
  const blob = Utilities.newBlob(Utilities.base64Decode(b64Data), MimeType.PDF, fileName);
  folder.createFile(blob);
}

function simpanKeDriveNative(payload) {
  try {
    const root = DriveApp.getRootFolder();
    const folderUtama = getFolderByNameOrCreate(root, "Laporan Harian");
    const folderUser = getFolderByNameOrCreate(folderUtama, payload.namaUser); 
    const tglFile = payload.tglLaporan ? payload.tglLaporan : payload.bulanTahun;

    if (payload.pdfGabunganB64) {
      const folderBulan = getFolderByNameOrCreate(folderUser, payload.bulanTahun);
      const fileName = `Laporan_Gabungan_${tglFile}.pdf`;
      saveFileWithOverwrite(folderBulan, fileName, payload.pdfGabunganB64);
    }

   if (payload.dataPerRhk && payload.dataPerRhk.length > 0) {
      payload.dataPerRhk.forEach(item => {
        const amanRhkNama = item.rhkNama.replace(/[\\/:*?"<>|]/g, "_");
        const tgl = item.tanggal ? item.tanggal : "TanpaTanggal";
        
        const folderRhk = getFolderByNameOrCreate(folderUser, amanRhkNama);
        const folderBulanRhk = getFolderByNameOrCreate(folderRhk, payload.bulanTahun);
        
        const fileName = `Laporan_${amanRhkNama}_${tgl}.pdf`;
        saveFileWithOverwrite(folderBulanRhk, fileName, item.pdfB64);
      });
    }
    
    // [PERBAIKAN] Update status di Spreadsheet menjadi "Generate" setelah PDF selesai dibuat
    if (payload.reportId) {
       const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
       const sheet = ss.getSheetByName('Data_Kegiatan');
       const data = sheet.getDataRange().getValues();
       for (let i = 1; i < data.length; i++) {
         if (String(data[i][0]) === String(payload.reportId)) {
           sheet.getRange(i + 1, 13).setValue("Generate"); // Update kolom M (13)
         }
       }
    }

    return { success: true, message: "Laporan PDF berhasil di-generate dan tersimpan di Drive!" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function exportExcelBulanan(bulanTahunInput, namaUser) {
  const emailUser = getEmailUser(); 
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetKegiatan = ss.getSheetByName('Data_Kegiatan');
  const data = sheetKegiatan.getDataRange().getValues();
  
  const tempSS = SpreadsheetApp.create("Rekap_Laporan_" + namaUser + "_" + bulanTahunInput);
  const sheet = tempSS.getSheets()[0];
  sheet.setName("Laporan_Harian");
  
  const folderLinksCache = {};
  try {
    const root = DriveApp.getRootFolder();
    const fUtamaIt = root.getFoldersByName("Laporan Harian");
    if(fUtamaIt.hasNext()) {
      const fUserIt = fUtamaIt.next().getFoldersByName(namaUser);
      if(fUserIt.hasNext()) {
        const fUser = fUserIt.next();
        const rhkFolders = fUser.getFolders();
        while(rhkFolders.hasNext()) {
          const rhkFolder = rhkFolders.next();
          const blnFolders = rhkFolder.getFoldersByName(bulanTahunInput);
          if(blnFolders.hasNext()) {
            folderLinksCache[rhkFolder.getName()] = blnFolders.next().getUrl();
          }
        }
      }
    }
  } catch(e) {}

  sheet.getRange("A1").setValue("Bulan " + bulanTahunInput);
  sheet.getRange("C1").setValue("Pilih Rencana Aksi");
  sheet.getRange("E1").setValue("Input Kinerja Harian");
  sheet.getRange("F1").setValue("TOTAL JAM");
  
  const headers = ["RHK", "No", "Tanggal", "Rencana Aksi", "Kinerja Harian", "Jam Mulai", "Jam Akhir", "Total Jam", "Hasil/Output", "Link Folder Drive", "Created By"];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  
  let rowIdx = 4;
  let no = 1;
  let totalJamSemua = 0; 

  for (let i = 1; i < data.length; i++) {
    const tglData = data[i][7] ? new Date(data[i][7]) : null;
    if(!tglData) continue;
    
    const rowEmail = String(data[i][1]).toLowerCase().trim();
    if(rowEmail !== emailUser) continue;
    
    const namaBulan = tglData.toLocaleString('id-ID', { month: 'long', year: 'numeric' });
    if (namaBulan === bulanTahunInput) {
       let tMulai = data[i][8];
       let tAkhir = data[i][9];
       let selisihText = "00:00:00";
       
       try {
           const d1 = new Date("1970-01-01T" + (tMulai instanceof Date ? Utilities.formatDate(tMulai, "GMT+7", "HH:mm") : tMulai) + ":00");
           const d2 = new Date("1970-01-01T" + (tAkhir instanceof Date ? Utilities.formatDate(tAkhir, "GMT+7", "HH:mm") : tAkhir) + ":00");
           let diffMs = d2 - d1;
           if(diffMs > 0) {
              totalJamSemua += diffMs;
              let h = Math.floor(diffMs / 3600000);
              let m = Math.floor((diffMs % 3600000) / 60000);
              selisihText = (h<10?'0':'')+h + ":" + (m<10?'0':'')+m + ":00";
           }
       } catch(e){}

       const rhkName = data[i][3];
       const amanRhkNama = rhkName.replace(/[\\/:*?"<>|]/g, "_");
       const linkFolderRHK = folderLinksCache[amanRhkNama] || "Belum ada folder";

       const rowData = [
         rhkName, no++, Utilities.formatDate(tglData, "GMT+7", "dd-MM-yyyy"),
         data[i][5], data[i][6], 
         tMulai instanceof Date ? Utilities.formatDate(tMulai, "GMT+7", "HH:mm:ss") : tMulai,
         tAkhir instanceof Date ? Utilities.formatDate(tAkhir, "GMT+7", "HH:mm:ss") : tAkhir,
         selisihText, "1", linkFolderRHK, namaUser
       ];
       
       sheet.getRange(rowIdx, 1, 1, headers.length).setValues([rowData]);
       rowIdx++;
    }
  }

  let th = Math.floor(totalJamSemua / 3600000);
  let tm = Math.floor((totalJamSemua % 3600000) / 60000);
  sheet.getRange("G1").setValue((th<10?'0':'')+th + ":" + (tm<10?'0':'')+tm + ":00");

  SpreadsheetApp.flush();
  return { url: "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx" };
}

function getFolderDriveUrl(namaUser, bulanTahun) {
  try {
    const root = DriveApp.getRootFolder();
    const foldersUtama = root.getFoldersByName("Laporan Harian");
    if (!foldersUtama.hasNext()) return "";
    const fUtama = foldersUtama.next();

    const foldersUser = fUtama.getFoldersByName(namaUser);
    if (!foldersUser.hasNext()) return "";
    const fUser = foldersUser.next();

    return fUser.getUrl();
  } catch(e) {
    return "";
  }
}

function saveImageToDrive(blob, fileName) {
  try {
    const folderName = "Assets_App";
    const folders = DriveApp.getFoldersByName(folderName);
    let folder;
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
    
    const existingFiles = folder.getFilesByName(fileName);
    while(existingFiles.hasNext()) {
      try { existingFiles.next().setTrashed(true); } catch(e) {}
    }
    
    const file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(shareErr) {}
    
    return file.getId(); 
  } catch (e) {
    throw new Error("Detail Error Drive: " + e.message);
  }
}

function extractDriveId(url) {
  if (!url) return null;
  let match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  match = url.match(/id=([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  if (!url.includes('http') && url.length > 20) return url.trim();
  return null;
}

function runGarbageCollector() {
  if(!ADMIN_EMAILS.includes(getEmailUser())) return { success: false, message: "Akses Ditolak. Anda bukan Admin." };
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let activeIds = new Set();
  
  const sheetKeg = ss.getSheetByName('Data_Kegiatan');
  if (sheetKeg && sheetKeg.getLastRow() > 1) {
    const dataKeg = sheetKeg.getDataRange().getValues();
    for (let i = 1; i < dataKeg.length; i++) {
      const id = extractDriveId(dataKeg[i][10]);
      if (id) activeIds.add(id);
    }
  }
  
  const sheetKop = ss.getSheetByName('Pengaturan_Surat');
  if (sheetKop && sheetKop.getLastRow() > 1) {
    const dataKop = sheetKop.getDataRange().getValues();
    for (let i = 1; i < dataKop.length; i++) {
      let idKiri = extractDriveId(dataKop[i][1]);
      let idKanan = extractDriveId(dataKop[i][2]);
      if (idKiri) activeIds.add(idKiri);
      if (idKanan) activeIds.add(idKanan);
    }
  }

  const sheetTtd = ss.getSheetByName('Pengaturan_TTD');
  if (sheetTtd && sheetTtd.getLastRow() > 1) {
    const dataTtd = sheetTtd.getDataRange().getValues();
    for (let i = 1; i < dataTtd.length; i++) {
      let idTtd = extractDriveId(dataTtd[i][6]);
      if (idTtd) activeIds.add(idTtd);
    }
  }

  let deletedCount = 0;
  try {
    const folderFoto = DriveApp.getFolderById(FOLDER_FOTO_ID);
    const files = folderFoto.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (!activeIds.has(file.getId())) {
        file.setTrashed(true);
        deletedCount++;
      }
    }
  } catch(e) {}
  
  try {
    const foldersAsset = DriveApp.getFoldersByName("Assets_App");
    while (foldersAsset.hasNext()) {
      const folderAsset = foldersAsset.next();
      const files = folderAsset.getFiles();
      while(files.hasNext()) {
        const file = files.next();
        if(!activeIds.has(file.getId())) {
          file.setTrashed(true);
          deletedCount++;
        }
      }
    }
  } catch(e) {}
  
  return { success: true, message: `Pembersihan berhasil!\nSebanyak ${deletedCount} file sampah telah dihapus dari Google Drive.` };
}

function getPengaturanUser(targetEmail) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetKop = ss.getSheetByName("Pengaturan_Surat");
  const sheetTtd = ss.getSheetByName("Pengaturan_TTD");
  const sheetUsers = ss.getSheetByName("Users");
  const ADMIN_EMAIL = ADMIN_EMAILS[0];

  let dataKop = getConfigByEmail(sheetKop, targetEmail, 'kop') || getConfigByEmail(sheetKop, ADMIN_EMAIL, 'kop') || {};
  let dataTtd = getConfigByEmail(sheetTtd, targetEmail, 'ttd') || {};
  
  let namaPemilik = "User";
  let nipPemilik = "-";
  let jabatanPemilik = "Jabatan";
  
  if (sheetUsers) {
    let usersData = sheetUsers.getDataRange().getValues();
    let target = String(targetEmail).toLowerCase().trim();
    for(let i = 1; i < usersData.length; i++) {
       if(String(usersData[i][0]).toLowerCase().trim() === target) {
          namaPemilik = usersData[i][3] || namaPemilik; 
          nipPemilik = usersData[i][4] || nipPemilik;   
          jabatanPemilik = usersData[i][5] || jabatanPemilik; 
          break;
       }
    }
  }

  if (!dataTtd.namaTtd) {
      dataTtd.namaTtd = namaPemilik;
      dataTtd.nipTtd = nipPemilik;
      dataTtd.jabatanTtd = jabatanPemilik !== "-" ? jabatanPemilik : "Wali Asuh";
      dataTtd.kota = "Temanggung"; 
      dataTtd.tandaTangan = ""; 
  }

  return { kop: dataKop, ttd: dataTtd, nama: namaPemilik };
}

function getSheetDataMap() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('DataSiswa');
    if (!sheet) return { error: "Sheet dengan nama 'DataSiswa' tidak ditemukan." };
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { headers: [], rows: [] };
    
    const headers = data[0];
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      let obj = {};
      for (let j = 0; j < headers.length; j++) {
        obj[headers[j]] = data[i][j];
      }
      result.push(obj);
    }
    return { headers: headers, rows: result };
  } catch (e) {
    return { error: "Gagal mengakses Spreadsheet Peta." };
  }
}

function getExportUrlMap() {
  return "https://docs.google.com/spreadsheets/d/" + SPREADSHEET_ID + "/export?format=xlsx";
}

function importCSVMap(csvString) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('DataSiswa');
    if(!sheet) sheet = ss.insertSheet('DataSiswa');
    
    const data = Utilities.parseCsv(csvString);
    if (data.length > 0) {
      sheet.clear(); 
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      return "Data lokasi siswa berhasil diimpor!";
    }
    return "Gagal memproses data.";
  } catch(e) {
    return "Error saat import: " + e.toString();
  }
}

function getHakAkses() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Hak_Akses');
  
  if (!sheet) {
    sheet = ss.insertSheet('Hak_Akses');
    sheet.appendRow(['Role', 'Akses_JSON']);
    return {};
  }
  
  const data = sheet.getDataRange().getValues();
  let result = {};
  
  for (let i = 1; i < data.length; i++) {
    try {
      result[data[i][0]] = JSON.parse(data[i][1]);
    } catch(e) {}
  }
  return result;
}

function saveHakAkses(payload) {
  if(!ADMIN_EMAILS.includes(getEmailUser())) return false; 
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Hak_Akses');
  if(!sheet) sheet = ss.insertSheet('Hak_Akses');
  
  sheet.clear();
  sheet.appendRow(['Role', 'Akses_JSON']);
  
  for (const [role, akses] of Object.entries(payload)) {
    sheet.appendRow([role, JSON.stringify(akses)]);
  }
  return true;
}

function simpanDataPelanggaran(payload) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName("Data_Pelanggaran");
        if (!sheet) return { success: false, message: "Sheet 'Data_Pelanggaran' tidak ditemukan!" };

        let urlFoto = "";
        if (payload.fotoBase64) {
            const base64Data = payload.fotoBase64.split(",")[1];
            const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.JPEG, `Bukti_${payload.siswa}_${payload.tanggal}.jpg`);
            
            const folderName = "Bukti Pelanggaran Siswa";
            let folders = DriveApp.getFoldersByName(folderName);
            let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

            const file = folder.createFile(blob);
            try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
            
            urlFoto = file.getUrl();
        }

        sheet.appendRow([
            new Date(), payload.siswa, payload.jenis,
            payload.poin, payload.tanggal, payload.keterangan, urlFoto, payload.dilaporkanOleh
        ]);

        // Trigger Notifikasi Otomatis
        let sheetNotif = ss.getSheetByName('Notifikasi');
        if (!sheetNotif) {
            sheetNotif = ss.insertSheet('Notifikasi');
            sheetNotif.appendRow(['ID', 'Target', 'Pesan', 'Tipe', 'Data Tambahan', 'Status Baca']);
        }
        
        // PERBAIKAN 3: Gunakan nama lengkap di notifikasi
        const labelNama = payload.namaSiswaLabel ? payload.namaSiswaLabel : payload.siswa;
        const pesanNotif = `Pelanggaran Baru: ${labelNama} - ${payload.jenis}`;
        
        sheetNotif.appendRow([new Date().getTime(), 'All', pesanNotif, 'Pelanggaran', payload.siswa, '']);

        return { success: true, message: "Pelanggaran berhasil dicatat!" };
    } catch (e) {
        return { success: false, message: "Error: " + e.message };
    }
}

function getHistoryPelanggaranSiswa(namaSiswa) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName("Data_Pelanggaran");
        if (!sheet) return [];

        const data = sheet.getDataRange().getValues();
        const history = [];
        const targetNama = String(namaSiswa).trim();

        for (let i = 1; i < data.length; i++) {
            let rowNama = String(data[i][1]).trim();
            if (rowNama === targetNama) {
                let tglData = data[i][4];
                let tglFormat = (tglData instanceof Date) ? Utilities.formatDate(tglData, "GMT+7", "dd/MM/yyyy") : String(tglData);

                // PERBAIKAN: Ubah URL Drive menjadi Base64 sebelum dikirim ke Frontend
                let b64Foto = "";
                if(data[i][6]) {
                    b64Foto = getFileBase64FromDrive(data[i][6]);
                }

                history.push({
                    tanggal: tglFormat,     
                    jenis: data[i][2],
                    poin: data[i][3],
                    keterangan: data[i][5],
                    urlFoto: b64Foto, // Sekarang berisi Base64, bukan link
                    dilaporkanOleh: data[i][7] || 'Sistem'
                });
            }
        }
        return history.reverse();
    } catch (e) {
        return [];
    }
}

function getDaftarSiswa() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName("DataSiswa"); 
        if (!sheet) return [];

        const data = sheet.getDataRange().getValues();
        const listSiswa = [];

        for (let i = 1; i < data.length; i++) {
            const nisn = data[i][0]; 
            const namaSiswa = data[i][2]; 
            if (namaSiswa) { 
                listSiswa.push(nisn ? `${namaSiswa} (${nisn})` : namaSiswa);
            }
        }
        return [...new Set(listSiswa)].sort();
    } catch (e) {
        return [];
    }
}

// Statistik Card Pelanggaran
function getStatistikPelanggaran() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheetSiswa = ss.getSheetByName("DataSiswa");
        let totalSiswa = 0;
        if (sheetSiswa) {
            const dataS = sheetSiswa.getDataRange().getValues();
            const uniqueSiswa = new Set();
            for (let i = 1; i < dataS.length; i++) {
                let nisn = String(dataS[i][0]).trim();
                let nama = String(dataS[i][2]).trim();
                
                // PERBAIKAN 4: Filter baris merah agar tidak dihitung sebagai siswa
                if (nisn && nama && nisn !== "" && nama !== "SEKOLAH RAKYAT") {
                    uniqueSiswa.add(nama);
                }
            }
            totalSiswa = uniqueSiswa.size;
        }

        const sheetPel = ss.getSheetByName("Data_Pelanggaran");
        let totalPelanggaran = 0;
        let totalPoin = 0;
        let setSiswaBermasalah = new Set();

        if (sheetPel) {
            const dataP = sheetPel.getDataRange().getValues();
            for (let i = 1; i < dataP.length; i++) {
                let namaSiswa = dataP[i][1];
                let poin = parseInt(dataP[i][3]) || 0;
                if (namaSiswa) { 
                    totalPelanggaran++;
                    totalPoin += poin;
                    setSiswaBermasalah.add(String(namaSiswa).trim());
                }
            }
        }

        return {
            success: true,
            totalSiswa: totalSiswa,
            totalPelanggaran: totalPelanggaran,
            totalPoin: totalPoin,
            siswaBermasalah: setSiswaBermasalah.size
        };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

// Fungsi Hapus Semua Data Pelanggaran & Foto di Drive (Hanya Admin)
function deleteAllPelanggaran() {
    try {
        const emailUser = getEmailUser();
        if (!ADMIN_EMAILS.includes(emailUser)) {
             return { success: false, message: "Akses Ditolak: Hanya Admin yang dapat menghapus seluruh data pelanggaran." };
        }

        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName("Data_Pelanggaran");
        if (!sheet) return { success: false, message: "Sheet 'Data_Pelanggaran' tidak ditemukan!" };

        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
            const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues(); 
            let fotoTerhapus = 0;

            for (let i = 0; i < data.length; i++) {
                const urlFoto = data[i][6]; 
                if (urlFoto) {
                    const fileId = extractDriveId(urlFoto);
                    if (fileId) {
                        try {
                            DriveApp.getFileById(fileId).setTrashed(true);
                            fotoTerhapus++;
                        } catch (errDrive) {}
                    }
                }
            }
            sheet.deleteRows(2, lastRow - 1);
            return { success: true, message: `Berhasil! Seluruh data pelanggaran beserta ${fotoTerhapus} file foto di Drive telah dikosongkan.` };
        } else {
            return { success: true, message: "Database pelanggaran sudah kosong." };
        }
    } catch (e) {
        return { success: false, message: "Gagal menghapus data: " + e.message };
    }
}

// Mengambil daftar jenis pelanggaran dari Spreadsheet (atau buat sheet baru otomatis jika belum ada)
function getJenisPelanggaranMaster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MasterPelanggaran");
  
  if (!sheet) {
    sheet = ss.insertSheet("MasterPelanggaran");
    sheet.appendRow(["ID", "Nama Pelanggaran", "Poin", "Kategori"]);
    // Data default
    sheet.appendRow([1, "Terlambat Masuk Kelas", 10, "Disiplin"]);
    sheet.appendRow([2, "Tidak Mengerjakan Tugas", 5, "Akademik"]);
  }
  
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    result.push({
      id: data[i][0],
      nama: data[i][1],
      poin: data[i][2],
      kategori: data[i][3]
    });
  }
  return result;
}

// Menyimpan jenis pelanggaran baru ke Spreadsheet
function simpanJenisPelanggaranMaster(item) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MasterPelanggaran");
  if (!sheet) {
    sheet = ss.insertSheet("MasterPelanggaran");
    sheet.appendRow(["ID", "Nama Pelanggaran", "Poin", "Kategori"]);
  }
  sheet.appendRow([item.id, item.nama, item.poin, item.kategori]);
  return { success: true };
}

// Menghapus jenis pelanggaran dari Spreadsheet berdasarkan ID
function hapusJenisPelanggaranMaster(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MasterPelanggaran");
  if (!sheet) return { success: false };
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { success: true };
}

// Fungsi untuk mengakses spreadsheet utama menggunakan konstanta SPREADSHEET_ID yang sudah ada
function getSpreadsheetMaster_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getJenisPelanggaranMaster() {
  var ss = getSpreadsheetMaster_();
  var sheet = ss.getSheetByName("MasterPelanggaran");
  
  if (!sheet) {
    sheet = ss.insertSheet("MasterPelanggaran");
    sheet.appendRow(["ID", "Nama Pelanggaran", "Poin", "Kategori"]);
    sheet.appendRow([1, "Terlambat Masuk Kelas", 10, "Disiplin"]);
    sheet.appendRow([2, "Tidak Mengerjakan Tugas", 5, "Akademik"]);
  }
  
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    result.push({
      id: data[i][0],
      nama: data[i][1],
      poin: data[i][2],
      kategori: data[i][3]
    });
  }
  return result;
}

function simpanJenisPelanggaranMaster(item) {
  var ss = getSpreadsheetMaster_();
  var sheet = ss.getSheetByName("MasterPelanggaran");
  if (!sheet) {
    sheet = ss.insertSheet("MasterPelanggaran");
    sheet.appendRow(["ID", "Nama Pelanggaran", "Poin", "Kategori"]);
  }
  sheet.appendRow([item.id, item.nama, item.poin, item.kategori]);
  return { success: true };
}

function hapusJenisPelanggaranMaster(id) {
  var ss = getSpreadsheetMaster_();
  var sheet = ss.getSheetByName("MasterPelanggaran");
  if (!sheet) return { success: false };
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { success: true };
}





// ==========================================
// MODUL BACKEND: PENILAIAN ASRAMA SISWA
// ==========================================

function getAppData(bulan, tahun) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. Ambil Data Siswa (Kolom A: NISN, Kolom C: Nama Siswa)
  const sheetSiswa = ss.getSheetByName('DataSiswa');
  let students = [];
  if (sheetSiswa) {
    const dataSiswa = sheetSiswa.getDataRange().getValues();
    for (let i = 1; i < dataSiswa.length; i++) {
      if(dataSiswa[i][0]) {
        students.push({ nisn: dataSiswa[i][0].toString(), nama: dataSiswa[i][2] });
      }
    }
  }

  // 2. Ambil Referensi Indikator
  const sheetRef = ss.getSheetByName('Referensi');
  let indicators = [];
  if (sheetRef) {
    const dataRef = sheetRef.getDataRange().getValues();
    let currentCategory = "";
    for (let i = 1; i < dataRef.length; i++) {
      const id = dataRef[i][0];
      if (!id) continue;
      
      const cat = dataRef[i][2];
      if (cat && cat.toString().trim() !== "") currentCategory = cat.toString().trim();
      
      indicators.push({
        id: id.toString(),
        category: currentCategory,
        indicator: dataRef[i][3]
      });
    }
  }

  // 3. Ambil Penilaian Tersimpan (Filter by Periode)
  const sheetPenilaian = ss.getSheetByName('PenilaianSiswa');
  let assessments = [];
  if (sheetPenilaian) {
    const dataPenilaian = sheetPenilaian.getDataRange().getValues();
    const paramBulan = String(bulan || '').trim().toLowerCase();
    const paramTahun = String(tahun || '').trim();

    for (let i = 1; i < dataPenilaian.length; i++) {
      let dbBulan = String(dataPenilaian[i][2] || '').trim().toLowerCase();
      let dbTahun = String(dataPenilaian[i][3] || '').trim();

      if (dbBulan === paramBulan && dbTahun === paramTahun) {
        try {
          assessments.push({
            id: dataPenilaian[i][0],
            bulan: dataPenilaian[i][2],
            tahun: dataPenilaian[i][3],
            nisn: dataPenilaian[i][4].toString(),
            nama: dataPenilaian[i][5],
            penilai: dataPenilaian[i][6],
            rataRata: parseFloat(dataPenilaian[i][7] || 0).toFixed(2),
            scores: JSON.parse(dataPenilaian[i][8] || '{}')
          });
        } catch (e) {
          continue;
        }
      }
    }
  }

  let user = Session.getActiveUser().getEmail() || 'User Penilai';

  return {
    students: students,
    indicators: indicators,
    assessments: assessments,
    currentUser: user
  };
}

function saveAssessment(payload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  let sheetPenilaian = ss.getSheetByName('PenilaianSiswa');
  if (!sheetPenilaian) {
    sheetPenilaian = ss.insertSheet('PenilaianSiswa');
    sheetPenilaian.appendRow(['ID', 'Timestamp', 'Bulan', 'Tahun', 'NISN', 'Nama Siswa', 'Penilai', 'Rata-rata', 'Detail JSON']);
  }
  
  const id = Utilities.getUuid();
  const timestamp = new Date();
  const rataRataUser = calculateAverage(payload.scores);
  
  const data = sheetPenilaian.getDataRange().getValues();
  let rowToUpdate = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] == payload.bulan && data[i][3] == payload.tahun && data[i][4] == payload.nisn && data[i][6] == payload.user) {
      rowToUpdate = i + 1;
      break;
    }
  }

  const rowData = [id, timestamp, payload.bulan, payload.tahun, payload.nisn, payload.nama, payload.user, rataRataUser, JSON.stringify(payload.scores)];
  
  if (rowToUpdate > -1) {
    sheetPenilaian.getRange(rowToUpdate, 1, 1, 9).setValues([rowData]);
  } else {
    sheetPenilaian.appendRow(rowData);
  }
  
  updateRataRataSemuaPenilai(ss, payload.bulan, payload.tahun, payload.nisn, payload.nama);
  return getAppData(payload.bulan, payload.tahun);
}

function updateRataRataSemuaPenilai(ss, bulan, tahun, nisn, nama) {
  const sheetPenilaian = ss.getSheetByName('PenilaianSiswa');
  const data = sheetPenilaian.getDataRange().getValues();
  
  let totalScore = 0;
  let countPenilai = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] == bulan && data[i][3] == tahun && data[i][4] == nisn) {
      totalScore += parseFloat(data[i][7] || 0);
      countPenilai++;
    }
  }
  const rataRataAkhir = countPenilai > 0 ? (totalScore / countPenilai).toFixed(2) : 0;
  
  let sheetRekap = ss.getSheetByName('RataRataSiswa');
  if (!sheetRekap) {
    sheetRekap = ss.insertSheet('RataRataSiswa');
    sheetRekap.appendRow(['Bulan', 'Tahun', 'NISN', 'Nama Siswa', 'Rata-rata Semua Penilai']);
  }
  
  const dataRekap = sheetRekap.getDataRange().getValues();
  let rowRekap = -1;
  for (let i = 1; i < dataRekap.length; i++) {
    if (dataRekap[i][0] == bulan && dataRekap[i][1] == tahun && dataRekap[i][2] == nisn) {
      rowRekap = i + 1;
      break;
    }
  }
  
  if (rowRekap > -1) {
    sheetRekap.getRange(rowRekap, 5).setValue(rataRataAkhir);
  } else {
    sheetRekap.appendRow([bulan, tahun, nisn, nama, rataRataAkhir]);
  }
}

function deleteAssessment(id, bulan, tahun) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('PenilaianSiswa');
  const data = sheet.getDataRange().getValues();
  let nisn = '';

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      nisn = data[i][4];
      sheet.deleteRow(i + 1);
      break;
    }
  }
  
  if (nisn) updateRataRataSemuaPenilai(ss, bulan, tahun, nisn, '');
  return getAppData(bulan, tahun);
}

function calculateAverage(scoresObj) {
  let total = 0, count = 0;
  for (const key in scoresObj) {
    if (scoresObj[key] !== '' && scoresObj[key] !== null) {
      total += parseFloat(scoresObj[key]);
      count++;
    }
  }
  return count > 0 ? (total / count).toFixed(2) : 0;
}

function generateSheetHasil(bulan, tahun) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const sheetSiswa = ss.getSheetByName('DataSiswa');
  let dataSiswaMap = {};
  if (sheetSiswa) {
    const dataSiswa = sheetSiswa.getDataRange().getValues();
    for (let i = 1; i < dataSiswa.length; i++) {
      let nisn = String(dataSiswa[i][0]).trim();
      let nik = String(dataSiswa[i][1]).trim();
      let nama = String(dataSiswa[i][2]).trim();
      if (nisn) dataSiswaMap[nisn] = { nik: nik, nama: nama };
    }
  }

  const sheetRef = ss.getSheetByName('Referensi');
  let refMap = {};
  let orderedIds = []; 
  if (sheetRef) {
    const dataRef = sheetRef.getDataRange().getValues();
    for (let i = 1; i < dataRef.length; i++) {
      let idAsli = String(dataRef[i][0] || '').trim();
      if (idAsli !== "") {
        orderedIds.push(idAsli);
        let angkaId = idAsli.replace(/[^0-9]/g, ''); 
        refMap[idAsli] = {
          headerNomor: angkaId,
          prefix: String(dataRef[i][4] || '').trim(),  
          verb: String(dataRef[i][6] || '').trim(),    
          context: String(dataRef[i][7] || '').trim(), 
          suffix: String(dataRef[i][8] || '').trim()   
        };
      }
    }
  }

  const sheetPenilaian = ss.getSheetByName('PenilaianSiswa');
  if (!sheetPenilaian) return "Sheet PenilaianSiswa tidak ditemukan.";
  
  const dataPen = sheetPenilaian.getDataRange().getValues();
  let paramBulan = String(bulan || '').trim().toLowerCase();
  let paramTahun = String(tahun || '').trim();
  
  let groupedSiswa = {};
  for (let i = 1; i < dataPen.length; i++) {
    let dbBulan = String(dataPen[i][2] || '').trim().toLowerCase();
    let dbTahun = String(dataPen[i][3] || '').trim();
    let nisn = String(dataPen[i][4] || '').trim();
    
    if (dbBulan === paramBulan && dbTahun === paramTahun && nisn) {
      if (!groupedSiswa[nisn]) groupedSiswa[nisn] = { scores: {} };
      try {
        let scoresObj = JSON.parse(dataPen[i][8] || '{}'); 
        for (let key in scoresObj) {
          let val = scoresObj[key];
          if (val !== '' && val !== null && !isNaN(val)) {
            let numVal = parseFloat(val);
            if (!groupedSiswa[nisn].scores[key]) {
              groupedSiswa[nisn].scores[key] = { sum: 0, count: 0 };
            }
            groupedSiswa[nisn].scores[key].sum += numVal;
            groupedSiswa[nisn].scores[key].count++;
          }
        }
      } catch (e) { continue; }
    }
  }

  let outputData = [];
  let header = ['NISN', 'NIK', 'nama_siswa'];
  for (let i = 0; i < orderedIds.length; i++) {
    let num = refMap[orderedIds[i]].headerNomor;
    header.push(`${num}_nilai`, ``, `${num}_deskripsi`); 
  }
  outputData.push(header);

  const getPredikat = (grade) => {
    switch(grade) {
      case 'A': return 'sangat baik';
      case 'B': return 'baik';
      case 'C': return 'cukup';
      case 'D': return 'kurang';
      case 'E': return 'sangat kurang';
      default: return '';
    }
  };

  for (let nisn in groupedSiswa) {
    let row = [
      nisn, 
      dataSiswaMap[nisn] ? dataSiswaMap[nisn].nik : '', 
      dataSiswaMap[nisn] ? dataSiswaMap[nisn].nama : ''
    ];
    
    for (let i = 0; i < orderedIds.length; i++) {
      let idAsli = orderedIds[i];
      let avgScore = '';
      let grade = '';
      let deskripsi = '';
      
      let sData = groupedSiswa[nisn].scores[idAsli]; 
      if (sData && sData.count > 0) {
        avgScore = Math.round(sData.sum / sData.count);
        if (avgScore >= 81) grade = 'A';
        else if (avgScore >= 61) grade = 'B';
        else if (avgScore >= 41) grade = 'C';
        else if (avgScore >= 21) grade = 'D';
        else grade = 'E';
        
        if (refMap[idAsli]) {
          let predikat = getPredikat(grade);
          deskripsi = `${refMap[idAsli].prefix} ${predikat} ${refMap[idAsli].verb} ${refMap[idAsli].context}${refMap[idAsli].suffix}`;
        }
      }
      row.push(avgScore, grade, deskripsi);
    }
    if (row.length > 3) outputData.push(row);
  }

  let targetSheetName = `Hasil_${bulan}_${tahun}`;
  let targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) targetSheet = ss.insertSheet(targetSheetName);
  else targetSheet.clear();
  
  if (outputData.length > 1) {
    targetSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
    targetSheet.getRange(1, 1, 1, outputData[0].length).setFontWeight("bold");
    targetSheet.setFrozenRows(1);
    targetSheet.setFrozenColumns(3);
  }

  return `Berhasil mencetak hasil ke sheet: ${targetSheetName}`;
}

// Tambahkan di bagian paling bawah Kode.gs
function getPenilaianHtml() {
  return HtmlService.createHtmlOutputFromFile('penilaian').getContent();
}

function getDashboardMainHtml() {
  return HtmlService.createTemplateFromFile('Index').evaluate().getContent();
}

// ==========================================
// NOTIFIKASI SISTEM & MANUAL
// ==========================================

function getNotifikasi() {
  const emailUser = getEmailUser();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheetNotif = ss.getSheetByName('Notifikasi');
  
  if (!sheetNotif) {
    sheetNotif = ss.insertSheet('Notifikasi');
    sheetNotif.appendRow(['ID', 'Target', 'Pesan', 'Tipe', 'Data Tambahan', 'Status Baca']);
    return [];
  }
  
  const data = sheetNotif.getDataRange().getValues();
  let notifList = [];
  
  for (let i = data.length - 1; i >= 1; i--) {
    let target = data[i][1];
    let statusBaca = data[i][5] ? data[i][5].toString() : '';
    
    // Tampilkan jika target 'All' atau email spesifik, dan belum dibaca oleh user ini
    if ((target === 'All' || target === emailUser) && !statusBaca.includes(emailUser)) {
      notifList.push({
        id: data[i][0],
        pesan: data[i][2],
        tipe: data[i][3],
        dataTambahan: data[i][4]
      });
    }
  }
  return notifList;
}

function tandaiNotifikasiDibaca(notifId) {
  const emailUser = getEmailUser();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetNotif = ss.getSheetByName('Notifikasi');
  if (!sheetNotif) return { success: false };
  
  const data = sheetNotif.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === notifId.toString()) {
      let statusBaca = data[i][5] ? data[i][5].toString() : '';
      if (!statusBaca.includes(emailUser)) {
        statusBaca += (statusBaca === '' ? '' : ',') + emailUser;
        sheetNotif.getRange(i + 1, 6).setValue(statusBaca);
      }
      return { success: true };
    }
  }
  return { success: false };
}

function simpanNotifikasiManual(pesan) {
  if(!ADMIN_EMAILS.includes(getEmailUser())) return { success: false, message: "Hanya Admin" };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheetNotif = ss.getSheetByName('Notifikasi');
  sheetNotif.appendRow([new Date().getTime(), 'All', pesan, 'Manual', '', '']);
  return { success: true, message: "Notifikasi berhasil dikirim!" };
}
// ==========================================
// FUNGSI MANAJEMEN NOTIFIKASI (ADMIN ONLY)
// ==========================================

function getSemuaNotifikasiAdmin() {
  if(!ADMIN_EMAILS.includes(getEmailUser())) return [];
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheetNotif = ss.getSheetByName('Notifikasi');
  if(!sheetNotif) return [];
  
  const data = sheetNotif.getDataRange().getValues();
  let notifList = [];
  
  // Looping dari bawah ke atas agar notifikasi terbaru muncul paling atas
  for(let i = data.length - 1; i >= 1; i--) { 
    notifList.push({
      id: data[i][0],
      target: data[i][1],
      pesan: data[i][2],
      tipe: data[i][3],
      dataTambahan: data[i][4]
    });
  }
  return notifList;
}

function hapusNotifikasiAdmin(id) {
  if(!ADMIN_EMAILS.includes(getEmailUser())) return { success: false };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetNotif = ss.getSheetByName('Notifikasi');
  if(!sheetNotif) return { success: false };
  
  const data = sheetNotif.getDataRange().getValues();
  for(let i = 1; i < data.length; i++) {
    if(data[i][0].toString() === id.toString()) {
      sheetNotif.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

function editNotifikasiAdmin(id, pesanBaru) {
  if(!ADMIN_EMAILS.includes(getEmailUser())) return { success: false };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetNotif = ss.getSheetByName('Notifikasi');
  if(!sheetNotif) return { success: false };
  
  const data = sheetNotif.getDataRange().getValues();
  for(let i = 1; i < data.length; i++) {
    if(data[i][0].toString() === id.toString()) {
      sheetNotif.getRange(i + 1, 3).setValue(pesanBaru); // Update kolom Pesan
      // Opsional: Mereset status baca agar notifikasi muncul lagi jika diedit
      sheetNotif.getRange(i + 1, 6).setValue(''); 
      return { success: true };
    }
  }
  return { success: false };
}
function getRoleAndAkses() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let roles = new Set(['Wali Asuh', 'Wali Asrama', 'Kepala Sekolah', 'Guru', 'Tamu', 'Tendik']); // Default role utama
  
  // 1. Ambil role tambahan dari Sheet Users jika ada
  try {
    const sheetUser = ss.getSheetByName('Users');
    if (sheetUser && sheetUser.getLastRow() > 1) {
      const data = sheetUser.getDataRange().getValues();
      const headers = data[0].map(h => String(h).trim().toLowerCase());
      let roleIndex = headers.indexOf('role');
      if (roleIndex === -1) roleIndex = 10; 
      
      for(let i=1; i<data.length; i++) {
         let r = data[i][roleIndex];
         if(r && String(r).trim() !== "") roles.add(String(r).trim());
      }
    }
  } catch(e) {}
  
  // 2. Ambil juga role yang sudah tersimpan di sheet Hak_Akses agar tidak terlewat
  let aksesData = {};
  try {
    let sheetAkses = ss.getSheetByName('Hak_Akses');
    if (sheetAkses && sheetAkses.getLastRow() > 1) {
      const dataAkses = sheetAkses.getDataRange().getValues();
      for (let i = 1; i < dataAkses.length; i++) {
        let roleName = dataAkses[i][0];
        if(roleName) roles.add(roleName.trim());
        
        let rawVal = dataAkses[i][1];
        try {
          let parsed = JSON.parse(rawVal);
          if (!Array.isArray(parsed) && typeof parsed === 'object' && parsed !== null) {
            let activeMenus = [];
            for (let [menuKey, isAllowed] of Object.entries(parsed)) {
              if (isAllowed === true || isAllowed === "true") activeMenus.push(menuKey);
            }
            aksesData[roleName] = activeMenus;
          } else if (Array.isArray(parsed)) {
            aksesData[roleName] = parsed;
          } else {
            aksesData[roleName] = [];
          }
        } catch(e) {
          aksesData[roleName] = [];
        }
      }
    }
  } catch(e) {}
  
  return { roles: Array.from(roles), akses: aksesData };
}
function getSiswaList() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName("DataSiswa");
        if (!sheet) return [];

        const data = sheet.getDataRange().getValues();
        const listSiswa = [];

        // Looping data mulai dari baris ke-2 (index 1)
        for (let i = 1; i < data.length; i++) {
            const nisn = String(data[i][0]).trim(); // Kolom A (NISN)
            const nama = String(data[i][2]).trim(); // Kolom C (Nama Siswa)

            // Filter: Pastikan NISN tidak kosong dan bukan baris header pembatas (misal: SEKOLAH RAKYAT)
            if (nisn && nama && nisn !== "" && nama !== "SEKOLAH RAKYAT") {
                listSiswa.push({
                    id: nisn,     // Menggunakan NISN sebagai value di HTML
                    nama: nama,
                    nisn: nisn
                });
            }
        }
        
        // Urutkan siswa berdasarkan abjad agar rapi
        listSiswa.sort((a, b) => a.nama.localeCompare(b.nama));

        return listSiswa;
    } catch (e) {
        return [];
    }
}