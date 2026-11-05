/**
 * @OnlyCurrentDoc
 */

const SPREADSHEET_ID = '1nLdFSrK1dvfrYu_db8tGEFVYuBILmNL-3X1eC1nu8YM'; 
const ADMIN_EMAILS = ['byu.project.comp@gmail.com', 'ajeefes@gmail.com']; 
const FOLDER_FOTO_RHK_ID = '13gHdT1Gt7S4uCeMskbIOTUy9FhHYzVBw'; 
const FOLDER_FOTO_MAKAN_ID = '1nZWOfWDcBTo_amD2W5N3RXcz6YH6HMeJ'; 
const FOLDER_FOTO_EKSTRA_ID = '1WPkuzpz-jI575WMfxQfWilV5mBAQMmu7'; 
const FOLDER_ASSESMEN_SISWA_ID = '1YeZugs9o4N5T4S1A8ydyQE_YH8cSfRLL'; 
const FOLDER_FOTO_CSDU_ID = '1NLSuoaLj6dAtjAQpO8BDsijTi2ujpLfU'; 
const FOLDER_PDF_CSDU_ID  = '1yLnmxHpUm57OXVSXBGnia5dAmDjxxvky'; 
const SHEET_CSDU_NAME     = 'Data_DU_CS_Security'; 

// --- TAMBAHAN BARU UNTUK FOLDER BERKAS SISWA ---
const FOLDER_KK_ID = '1p40IUSqbp0RMc6zQbjtCUVhX5KVPGfgw';
const FOLDER_AKTA_ID = '1YkJD3hNxYBYLqogyrpoSpxDtX_ITzCeT';

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

  const isReallyAdmin = ADMIN_EMAILS.includes(email.toLowerCase().trim()) || (roleUser === "Admin");
  return { status: "ACTIVE", isAdmin: isReallyAdmin, role: roleUser, userData: userData };
}

function verifyAndActivateToken(token) {
  const email = getEmailUser();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  let tokenSheet = ss.getSheetByName('Tokens');
  if (!tokenSheet) return { success: false, message: "Sistem token tidak ditemukan." };

  const tokenData = tokenSheet.getDataRange().getValues();
  for (let i = 1; i < tokenData.length; i++) {
    if (String(tokenData[i][0]).trim() === String(token).trim()) {
      if (String(tokenData[i][1]).trim().toLowerCase() === 'unused') {
        
        // 1. Tandai token sebagai 'used'
        tokenSheet.getRange(i + 1, 2).setValue('used');
        tokenSheet.getRange(i + 1, 3).setValue(email);

        let userSheet = ss.getSheetByName('Users');
        let usersData = userSheet.getDataRange().getValues();
        let headers = usersData[0];
        
        let expiry = new Date();
        // Atur perpanjangan masa aktif (Saat ini diset +1 bulan, ubah setMonth ke setFullYear jika ingin +1 tahun)
        expiry.setMonth(expiry.getMonth() + 1); 

        let userExists = false;
        
        // 2. Cek apakah user sudah terdaftar berdasarkan email
        for (let j = 1; j < usersData.length; j++) {
          let emailIndex = headers.indexOf('Email') > -1 ? headers.indexOf('Email') : 0;
          
          if (String(usersData[j][emailIndex]).toLowerCase().trim() === email) {
            // JIKA USER SUDAH ADA: Update Status, Expiry, dan Token-nya saja
            userSheet.getRange(j + 1, headers.indexOf('IsActive') + 1).setValue(true);
            
            let expIndex = headers.indexOf('ExpiryDate') > -1 ? headers.indexOf('ExpiryDate') : headers.indexOf('Expiry');
            userSheet.getRange(j + 1, expIndex + 1).setValue(expiry);
            
            let tokenIndex = headers.indexOf('Token');
            if (tokenIndex > -1) {
                userSheet.getRange(j + 1, tokenIndex + 1).setValue(String(token).trim());
            }
            
            userExists = true;
            break;
          }
        }

        // 3. JIKA USER BELUM ADA: Buat baris baru
        if (!userExists) {
          userSheet.appendRow([email, true, expiry, 'User Baru', '-', '-', '-', String(token).trim(), '']);
        }

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
  
  let sheetName = ''; // Default kosong

  if (jabatan === 'Wali Asuh') {
    sheetName = 'RHK_Master_Wali_Asuh';
  } else if (jabatan === 'Wali Asrama') {
    sheetName = 'RHK_Master_Wali_Asrama';
  } else if (jabatan === 'Guru') {
    sheetName = 'RHK_Master_Guru';
  } else {
    return []; 
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

// Fungsi pembantu untuk membaca seluruh baris data master (Global & Aktivitas)
function getRhkDataFromSheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const rawData = sheet.getDataRange().getValues();
  let result = [];
  
  // Asumsi Kolom: A (0) = ID, B (1) = Tipe, C (2) = RHK_Nama/Aktivitas, D (3) = ParentId/Keterangan
  for (let i = 1; i < rawData.length; i++) {
    if (rawData[i][0]) { // Pastikan baris tidak kosong
      result.push({
        id: rawData[i][0],
        tipe: rawData[i][1],          // 'Global' atau 'Aktivitas'
        nama: rawData[i][2],          // Nama RHK atau Nama Aktivitas
        keterangan: rawData[i][3] || '' // Keterangan atau Parent ID
      });
    }
  }
  return result;
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

function savePendahuluan(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Pendahuluan_Master');
  
  if (!sheet) {
    sheet = ss.insertSheet('Pendahuluan_Master');
    // Tambahkan header Jenis_Master di kolom E
    sheet.appendRow(['ID', 'Nama', 'Isi', 'RHK_Link', 'Jenis_Master']);
  }

  if (data.id && data.id !== '') {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.id.toString()) {
        // Lebarkan jangkauan update menjadi 4 kolom (2, 1, 4) untuk mencakup jenisMaster
        sheet.getRange(i + 1, 2, 1, 4).setValues([[data.nama, data.isi, data.rhkLink, data.jenisMaster || '']]);
        return { success: true, message: "Template berhasil diperbarui!" };
      }
    }
  } 
  
  // Tambahkan data.jenisMaster saat membuat baris baru
  sheet.appendRow([new Date().getTime(), data.nama, data.isi, data.rhkLink, data.jenisMaster || '']);
  return { success: true, message: "Template berhasil disimpan ke Database!" };
}

function getSemuaPendahuluan() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Pendahuluan_Master');
  
  if (!sheet) {
    sheet = ss.insertSheet('Pendahuluan_Master');
    sheet.appendRow(['ID', 'Nama', 'Isi', 'RHK_Link', 'Jenis_Master']);
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  // Tambahkan jenisMaster: row[4] pada mapping
  return data.slice(1).map(row => ({ 
    id: row[0], 
    nama: row[1], 
    isi: row[2], 
    rhkLink: row[3], 
    jenisMaster: row[4] || '' 
  }));
}

function getPendahuluanById(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
  const sheet = ss.getSheetByName('Pendahuluan_Master');
  
  if (!sheet) {
    return { success: false, message: "Sheet Pendahuluan_Master tidak ditemukan!" };
  }

  const data = sheet.getDataRange().getValues();
  let dataPendahuluan = null;

  // Looping untuk mencari baris dengan ID yang cocok
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      dataPendahuluan = {
        id: data[i][0],
        nama: data[i][1],
        isi: data[i][2],
        rhkLink: data[i][3],
        jenisMaster: data[i][4] || '' // Tangkap Jenis Master dari Kolom E
      };
      break; 
    }
  }

  if (dataPendahuluan) {
    let kopHtml = ""; 
    // const sheetKop = ss.getSheetByName('Pengaturan_KOP');
    // kopHtml = sheetKop.getRange("A2").getValue(); 

    return { 
      success: true, 
      dataPendahuluan: dataPendahuluan, 
      kopHtml: kopHtml 
    };
  } else {
    return { success: false, message: "Data template pendahuluan tidak ditemukan." };
  }
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
  const folder = DriveApp.getFolderById(FOLDER_FOTO_RHK_ID);
  
  const reportId = payload.reportId || "LPR-" + new Date().getTime(); 
  let existingRowIndices = [];
  
  if (payload.reportId) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(payload.reportId).trim()) {
        existingRowIndices.push(i + 1);
      }
    }
  }

  payload.kegiatan.forEach((keg, index) => {
    let fileUrl = "";
    
    if (keg.rawFoto && keg.rawFoto.trim() !== "") {
      fileUrl = keg.rawFoto;
    } 
    else if (keg.foto && keg.foto.startsWith("data:image")) {
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

    // [PERBAIKAN] Menambahkan petik tunggal (') di depan jamMulai dan jamAkhir
    let rowData = [
      reportId, payload.email || "", keg.rhkId || "", keg.rhkNama || "", keg.aktivitasId || "", 
      keg.aktivitasNama || "", keg.keterangan || "", keg.tanggal || "", 
      keg.jamMulai ? "'" + keg.jamMulai : "", 
      keg.jamAkhir ? "'" + keg.jamAkhir : "", 
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
  
  SpreadsheetApp.flush();
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
    const roleIndex = headers.indexOf('Jabatan'); 
    
    for (let i = 1; i < dataU.length; i++) {
      let dbEmail = String(dataU[i][0]).toLowerCase().trim();
      mapNama[dbEmail] = dataU[i][3];
      if (dbEmail === emailUser && roleIndex > -1 && dataU[i][roleIndex] === 'Kepala Sekolah') {
          isKepalaSekolah = true;
      }
    }
  }
  
  const canViewAll = isAdmin || isKepalaSekolah; 
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  const riwayat = {};
  
  // --- PENAMBAHAN: Array Nama Bulan Bahasa Indonesia ---
  const namaBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  const namaBulanPendek = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Ags", "Sep", "Okt", "Nov", "Des"];
  // -----------------------------------------------------

  for (let i = 0; i < data.length; i++) {
    let rowEmail = String(data[i][1]).toLowerCase().trim();
    
    if (rowEmail !== emailUser && !canViewAll) continue; 
    
    let reportId = data[i][0]; 
    if (!reportId) continue;
    if (!riwayat[reportId]) {
       let tglObj = data[i][7] ? new Date(data[i][7]) : null;
       let statusDB = data[i][12] || 'Draft'; 
       
       // --- PERBAIKAN: Format Tanggal ke Bahasa Indonesia ---
       let tglMulaiIndo = '-';
       let bulanTahunIndo = '-';
       
       if (tglObj) {
           let d = String(tglObj.getDate()).padStart(2, '0');
           let m = tglObj.getMonth(); // Index 0-11
           let yyyy = tglObj.getFullYear();
           let yy = String(yyyy).slice(-2);
           
           tglMulaiIndo = `${d} ${namaBulanPendek[m]} ${yy}`; // Contoh: 03 Ags 26
           bulanTahunIndo = `${namaBulan[m]} ${yyyy}`;        // Contoh: Agustus 2026
       }
       // -----------------------------------------------------
       
       riwayat[reportId] = {
        id: reportId, 
        email: data[i][1], 
        nama: mapNama[rowEmail] || 'User', 
        jumlahKegiatan: 0,
        tanggalMulai: tglMulaiIndo,   // Menggunakan string bahasa Indonesia
        bulanTahun: bulanTahunIndo,   // Menggunakan string bahasa Indonesia ("Agustus 2026")
        status: statusDB 
      };
    }
    riwayat[reportId].jumlahKegiatan += 1;
  }
  return Object.values(riwayat).sort((a,b) => b.id.localeCompare(a.id));
}

function getDetailLaporan(reportId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data_Kegiatan');
  const data = sheet.getDataRange().getValues();
  const displayData = sheet.getDataRange().getDisplayValues(); // [PERBAIKAN] Ambil format murni layar
  const searchId = String(reportId).trim();
  const hasil = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === searchId) {
      let rawFoto = data[i][10] || '';
      let pendRaw = data[i][11] || "";
      let base64Foto = getFileBase64FromDrive(rawFoto);
      
      // [PERBAIKAN] Ambil string dari DisplayData dan potong ke "HH:mm"
      let valMulai = displayData[i][8] || '';
      let valAkhir = displayData[i][9] || '';
      if(valMulai.includes(':')) valMulai = valMulai.substring(0, 5);
      if(valAkhir.includes(':')) valAkhir = valAkhir.substring(0, 5);

      hasil.push({
        email: data[i][1], 
        rhkId: data[i][2], rhkNama: data[i][3], aktivitasId: data[i][4], aktivitasNama: data[i][5],
        keterangan: data[i][6], 
        tanggal: data[i][7] ? Utilities.formatDate(new Date(data[i][7]), "GMT+7", "yyyy-MM-dd") : '', 
        jamMulai: valMulai,
        jamAkhir: valAkhir,
        foto: base64Foto, 
        rawFoto: rawFoto,
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
    // ---- 1. SETUP ROOT DRIVE USER (FUNGSI ASLI) ----
    const root = DriveApp.getRootFolder();
    const folderUtama = getFolderByNameOrCreate(root, "Laporan Harian");
    const folderUser = getFolderByNameOrCreate(folderUtama, payload.namaUser); 
    const tglFile = payload.tglLaporan ? payload.tglLaporan : payload.bulanTahun;

    // ---- 2. SETUP DRIVE PARALEL (TAMBAHAN) ----
    const idBackup = '1quTf_wLAPnzta9EY_7dBsRE9RvmmHfAc';
    const rootBackup = DriveApp.getFolderById(idBackup);
    const folderUtamaBackup = getFolderByNameOrCreate(rootBackup, "Laporan Harian");
    const folderUserBackup = getFolderByNameOrCreate(folderUtamaBackup, payload.namaUser); 

    // ---- 3. EKSEKUSI SIMPAN (GABUNGAN) ----
    if (payload.pdfGabunganB64) {
      const fileName = `Laporan_Gabungan_${tglFile}.pdf`;
      
      // Simpan Asli
      const folderBulan = getFolderByNameOrCreate(folderUser, payload.bulanTahun);
      saveFileWithOverwrite(folderBulan, fileName, payload.pdfGabunganB64);
      
      // Simpan Paralel
      const folderBulanBackup = getFolderByNameOrCreate(folderUserBackup, payload.bulanTahun);
      saveFileWithOverwrite(folderBulanBackup, fileName, payload.pdfGabunganB64);
    }

    // ---- 4. EKSEKUSI SIMPAN (PER RHK) ----
    if (payload.dataPerRhk && payload.dataPerRhk.length > 0) {
      payload.dataPerRhk.forEach(item => {
        const amanRhkNama = item.rhkNama.replace(/[\\/:*?"<>|]/g, "_");
        const tgl = item.tanggal ? item.tanggal : "TanpaTanggal";
        const fileName = `Laporan_${amanRhkNama}_${tgl}.pdf`;
        
        // Simpan Asli
        const folderRhk = getFolderByNameOrCreate(folderUser, amanRhkNama);
        const folderBulanRhk = getFolderByNameOrCreate(folderRhk, payload.bulanTahun);
        saveFileWithOverwrite(folderBulanRhk, fileName, item.pdfB64);

        // Simpan Paralel
        const folderRhkBackup = getFolderByNameOrCreate(folderUserBackup, amanRhkNama);
        const folderBulanRhkBackup = getFolderByNameOrCreate(folderRhkBackup, payload.bulanTahun);
        saveFileWithOverwrite(folderBulanRhkBackup, fileName, item.pdfB64);
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

    return { success: true, message: "Laporan PDF berhasil di-generate dan tersimpan secara paralel di Drive!" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function hapusDataBulanPilihan(bulanString, arraySheetDipilih) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // --- KONFIGURASI KOLOM PER SHEET BERDASARKAN GAMBAR ---
  // Catatan: Index kolom dimulai dari 0 (A=0, B=1, C=2, dst)
  // colFoto dibuat dalam bentuk array [] agar bisa menghapus >1 foto per baris
  const configKolom = {
    "Data_Kegiatan":       { colTanggal: 7, colFoto: [10] },       // Tanggal: H(7), Foto: K(10)
    "Data_DU_CS_Security": { colTanggal: 1, colFoto: [6, 7, 8] },  // Tanggal: B(1), Foto1: G(6), Foto2: H(7), Foto3: I(8)
    "Data_PJ_PotongFC":    { colTanggal: 4, colFoto: [6] },        // Tanggal: E(4), Foto: G(6)
    "Data_PJ_Pembelian":   { colTanggal: 4, colFoto: [6, 7, 8] },  // Tanggal: E(4), Foto1: G(6), Foto2: H(7), Foto3: I(8)
    "Data_PJ_Ekstra":      { colTanggal: 5, colFoto: [7] },        // Tanggal: F(5), URL_Foto1: H(7)
    "Data_PJ_Makan":       { colTanggal: 5, colFoto: [7] },        // Tanggal: F(5), URL_Foto_Makanan: H(7)
    "LapKepsek":           { colTanggal: 2, colFoto: [] }          // Tanggal Laporan: C(2), Tidak ada kolom link foto langsung
  };
  
  let summaryDihapus = {};
  let totalKeseluruhan = 0;

  // Lakukan perulangan untuk setiap sheet yang diceklis
  arraySheetDipilih.forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    let config = configKolom[sheetName];
    
    // Lewati jika sheet tidak ada atau konfigurasinya belum dibuat
    if (!sheet || !config) {
      summaryDihapus[sheetName] = "Gagal (Sheet/Konfigurasi tidak ditemukan)";
      return;
    }

    const data = sheet.getDataRange().getValues();
    let jumlahDihapus = 0;

    // Perulangan dari bawah ke atas agar indeks baris tidak berantakan saat dihapus
    for (let i = data.length - 1; i > 0; i--) { 
      let baris = data[i];
      let tanggalCell = baris[config.colTanggal];
      let formatSheet = "";
      
      // Ambil dan konversi format tanggal menjadi "YYYY-MM"
      if (tanggalCell instanceof Date) {
        let thn = tanggalCell.getFullYear();
        let bln = String(tanggalCell.getMonth() + 1).padStart(2, '0');
        formatSheet = `${thn}-${bln}`;
      } else if (typeof tanggalCell === "string" && tanggalCell.length >= 7) {
        formatSheet = tanggalCell.substring(0, 7); 
      }
      
      // Jika bulan di baris tersebut cocok dengan bulan pilihan user
      if (formatSheet === bulanString) {
        
        // 1. Ekstrak dan hapus file foto (bisa lebih dari 1 kolom foto)
        if (config.colFoto && config.colFoto.length > 0) {
          config.colFoto.forEach(colIndex => {
            let urlFoto = baris[colIndex];
            if (urlFoto && typeof urlFoto === 'string' && urlFoto.includes('/d/')) {
              try {
                let idFile = urlFoto.split('/d/')[1].split('/')[0];
                if (idFile) DriveApp.getFileById(idFile).setTrashed(true);
              } catch (e) {
                Logger.log(`Gagal menghapus file di ${sheetName} baris ${i+1}`);
              }
            }
          });
        }
        
        // 2. Hapus baris di Spreadsheet
        sheet.deleteRow(i + 1);
        jumlahDihapus++;
      }
    }
    
    summaryDihapus[sheetName] = `${jumlahDihapus} baris dihapus`;
    totalKeseluruhan += jumlahDihapus;
  });
  
  // Menyusun pesan summary untuk ditampilkan di pop-up HTML
  let pesanSummary = `<b>Total: ${totalKeseluruhan} data berhasil dihapus.</b><br><br><ul style="text-align:left; margin-left: 20px;">`;
  for (const [sheet, hasil] of Object.entries(summaryDihapus)) {
    pesanSummary += `<li><b>${sheet}</b>: ${hasil}</li>`;
  }
  pesanSummary += `</ul>`;

  return { 
    status: "success", 
    pesanHTML: pesanSummary 
  };
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
    // Array berisi ke-3 ID Folder baru
    const folderIds = [FOLDER_FOTO_RHK_ID, FOLDER_FOTO_MAKAN_ID, FOLDER_FOTO_EKSTRA_ID];
    
    // Looping untuk membersihkan sampah di ketiga folder
    folderIds.forEach(idFolder => {
      const folderFoto = DriveApp.getFolderById(idFolder);
      const files = folderFoto.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        if (!activeIds.has(file.getId())) {
          file.setTrashed(true);
          deletedCount++;
        }
      }
    });
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
    // PERBAIKAN: Ubah header menjadi Jabatan agar seragam
    sheet.appendRow(['Jabatan', 'Akses_JSON']);
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
  // PERBAIKAN: Ubah header menjadi Jabatan agar seragam
  sheet.appendRow(['Jabatan', 'Akses_JSON']);
  
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

// Menyimpan jenis pelanggaran baru ke Spreadsheet
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

// Fungsi untuk mengakses spreadsheet utama menggunakan konstanta SPREADSHEET_ID yang sudah ada
function getSpreadsheetMaster_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
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
  
  let roles = new Set(['Wali Asuh', 'Wali Asrama', 'Kepala Sekolah', 'Guru', 'Tamu', 'Tendik', 'Perawat', 'DU', 'CS', 'Satpam']); 
  
  try {
    const sheetUser = ss.getSheetByName('Users');
    if (sheetUser && sheetUser.getLastRow() > 1) {
      const data = sheetUser.getDataRange().getValues();
      const headers = data[0].map(h => String(h).trim().toLowerCase());
      // PERBAIKAN 2: Mencari index 'jabatan' (karena sudah di-toLowerCase)
      let roleIndex = headers.indexOf('jabatan');
      if (roleIndex === -1) roleIndex = 10; 
      
      for(let i=1; i<data.length; i++) {
         let r = data[i][roleIndex];
         if(r && String(r).trim() !== "") roles.add(String(r).trim());
      }
    }
  } catch(e) {}
  
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
// ==========================================
// MODUL BACKEND: RIWAYAT KESEHATAN
// ==========================================

function simpanDataKesehatan(payload) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName("Data_Kesehatan");
        if (!sheet) return { success: false, message: "Sheet 'Data_Kesehatan' tidak ditemukan!" };

        // Struktur penulisan baru: NISN dan Nama dipisah
        sheet.appendRow([
            new Date(), 
            payload.nisn,          // Kolom B (NISN)
            payload.namaSiswa,     // Kolom C (Nama)
            payload.kelas,         // Kolom D (Kelas)
            payload.jenis,         // Kolom E
            payload.tanggal,       // Kolom F
            payload.ttv,           // Kolom G
            payload.keterangan,    // Kolom H
            payload.rujukan,       // Kolom I
            payload.dilaporkanOleh // Kolom J
        ]);

        let sheetNotif = ss.getSheetByName('Notifikasi');
        if (sheetNotif) {
            const pesanNotif = `Info Kesehatan Baru: ${payload.namaSiswa} - ${payload.jenis}`;
            sheetNotif.appendRow([new Date().getTime(), 'All', pesanNotif, 'Kesehatan', payload.nisn, '']);
        }

        return { success: true, message: "Riwayat kesehatan berhasil dicatat!" };
    } catch (e) {
        return { success: false, message: "Error: " + e.message };
    }
}

function getHistoryKesehatanSiswa(nisnSiswa) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName("Data_Kesehatan");
        if (!sheet) return [];

        const data = sheet.getDataRange().getValues();
        const history = [];
        const targetNisn = String(nisnSiswa).trim();

        for (let i = 1; i < data.length; i++) {
            let rowNisn = String(data[i][1]).trim(); // Pengecekan dilakukan berdasarkan NISN di Kolom B
            if (rowNisn === targetNisn) {
                let tglData = data[i][5]; // Bergeser ke index 5
                let tglFormat = (tglData instanceof Date) ? Utilities.formatDate(tglData, "GMT+7", "dd/MM/yyyy") : String(tglData);
                
                let rawTglInput = data[i][0];
                let rowId = (rawTglInput instanceof Date) ? rawTglInput.getTime() : new Date(rawTglInput).getTime();
                
                let tanggapanData = [];
                try {
                    if(data[i][10]) tanggapanData = JSON.parse(data[i][10]); // Kolom K (10)
                } catch(e) {}

                history.push({
                    idRow: rowId, 
                    tanggal: tglFormat,
                    kelas: data[i][3] || '-',     // Kolom D
                    jenis: data[i][4],            // Kolom E
                    ttv: data[i][6] || '-',       // Kolom G
                    keterangan: data[i][7],       // Kolom H
                    rujukan: data[i][8] || '-',   // Kolom I
                    dilaporkanOleh: data[i][9] || 'Sistem', // Kolom J
                    tanggapan: tanggapanData
                });
            }
        }
        return history.reverse();
    } catch (e) {
        return [];
    }
}

function getStatistikKesehatan() {
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
                if (nisn && nama && nisn !== "" && nama !== "SEKOLAH RAKYAT") {
                    uniqueSiswa.add(nama);
                }
            }
            totalSiswa = uniqueSiswa.size;
        }

        const sheetKes = ss.getSheetByName("Data_Kesehatan");
        let totalKeluhan = 0;
        let setSiswaSakit = new Set();

        if (sheetKes) {
            const dataK = sheetKes.getDataRange().getValues();
            for (let i = 1; i < dataK.length; i++) {
                let nisnSiswa = dataK[i][1]; // Ambil patokan dari NISN
                if (nisnSiswa) { 
                    totalKeluhan++;
                    setSiswaSakit.add(String(nisnSiswa).trim());
                }
            }
        }

        return {
            success: true,
            totalSiswa: totalSiswa,
            totalKeluhan: totalKeluhan,
            siswaSakit: setSiswaSakit.size
        };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

function deleteAllKesehatan() {
    try {
        const emailUser = getEmailUser();
        if (!ADMIN_EMAILS.includes(emailUser)) {
             return { success: false, message: "Akses Ditolak: Hanya Admin yang dapat menghapus data." };
        }

        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName("Data_Kesehatan");
        if (!sheet) return { success: false, message: "Sheet 'Data_Kesehatan' tidak ditemukan!" };

        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
            sheet.deleteRows(2, lastRow - 1);
            return { success: true, message: "Seluruh data riwayat kesehatan telah dikosongkan." };
        }
        return { success: true, message: "Database riwayat kesehatan sudah kosong." };
    } catch (e) {
        return { success: false, message: "Gagal menghapus data: " + e.message };
    }
}

// Master Jenis Keluhan
function getJenisKeluhanMaster() {
  // PERBAIKAN: Gunakan openById, bukan getActiveSpreadsheet
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
  var sheet = ss.getSheetByName("MasterKeluhan");
  
  if (!sheet) {
    sheet = ss.insertSheet("MasterKeluhan");
    sheet.appendRow(["ID", "Nama Keluhan"]);
    sheet.appendRow([1, "Demam"]);
    sheet.appendRow([2, "Sakit Perut"]);
  }
  
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    result.push({
      id: data[i][0],
      nama: data[i][1]
    });
  }
  return result;
}

function simpanJenisKeluhanMaster(item) {
  // PERBAIKAN: Gunakan openById
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("MasterKeluhan");
  if (!sheet) {
    sheet = ss.insertSheet("MasterKeluhan");
    sheet.appendRow(["ID", "Nama Keluhan"]);
  }
  sheet.appendRow([item.id, item.nama]);
  return { success: true };
}

function hapusJenisKeluhanMaster(id) {
  // PERBAIKAN: Gunakan openById
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("MasterKeluhan");
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
// MODUL BACKEND: PJ MAKAN & PJ EKSTRA
// ==========================================

function simpanPJMakanBatch(payload) {
  if (!payload || !payload.kegiatan || payload.kegiatan.length === 0) return { success: false, message: "Data kosong" };
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Data_PJ_Makan');
  
  const reportId = payload.reportId || "PJM-" + new Date().getTime(); 

  // --- LOGIKA KEAMANAN EDIT ---
  if (payload.reportId) {
     const emailUserAktif = Session.getActiveUser().getEmail().toLowerCase().trim();
     const isAdmin = ADMIN_EMAILS.includes(emailUserAktif);
     const data = sheet.getDataRange().getValues();
     
     let isAuthorized = false;
     for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === payload.reportId) {
           const emailPenginput = String(data[i][2]).toLowerCase().trim();
           if (emailUserAktif === emailPenginput || isAdmin) isAuthorized = true;
           break;
        }
     }
     
     if (!isAuthorized) return { success: false, message: "AKSES DITOLAK: Anda hanya dapat mengedit data yang Anda input sendiri." };

     for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === payload.reportId) sheet.deleteRow(i + 1);
     }
  }
  // ----------------------------

  const folder = DriveApp.getFolderById(FOLDER_FOTO_MAKAN_ID);
  payload.kegiatan.forEach((keg, index) => {
    let url1 = uploadB64ToDriveSafe(keg.foto1, folder, "Makan1_" + index);
    let url2 = uploadB64ToDriveSafe(keg.foto2, folder, "Makan2_" + index);
    let url3 = uploadB64ToDriveSafe(keg.foto3, folder, "Makan3_" + index);

    sheet.appendRow([
      reportId, new Date(), payload.email, 
      keg.jenis, keg.makan, keg.tanggal, keg.keterangan, 
      url1, url2, url3, keg.kategori || 'Umum'
    ]);
  });
  return { success: true, message: "Laporan PJ Makan berhasil disimpan!" };
}

function simpanPJMakanPuasaBatch(payload) {
  if (!payload || !payload.kegiatan || payload.kegiatan.length === 0) return { success: false, message: "Data kosong" };
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Data_PJ_Makan_Puasa'); 
  
  const reportId = payload.reportId || "PJMP-" + new Date().getTime(); 

  // --- LOGIKA KEAMANAN EDIT ---
  if (payload.reportId) {
     const emailUserAktif = Session.getActiveUser().getEmail().toLowerCase().trim();
     const isAdmin = ADMIN_EMAILS.includes(emailUserAktif);
     const data = sheet.getDataRange().getValues();
     
     let isAuthorized = false;
     for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === payload.reportId) {
           const emailPenginput = String(data[i][2]).toLowerCase().trim();
           if (emailUserAktif === emailPenginput || isAdmin) isAuthorized = true;
           break;
        }
     }
     
     if (!isAuthorized) return { success: false, message: "AKSES DITOLAK: Anda hanya dapat mengedit data yang Anda input sendiri." };

     for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === payload.reportId) sheet.deleteRow(i + 1);
     }
  }
  // ----------------------------

  const folder = DriveApp.getFolderById(FOLDER_FOTO_MAKAN_ID); 
  payload.kegiatan.forEach((keg, index) => {
    let url1 = uploadB64ToDriveSafe(keg.foto1, folder, "MakanPuasa1_" + index);
    let url2 = uploadB64ToDriveSafe(keg.foto2, folder, "MakanPuasa2_" + index);
    let url3 = uploadB64ToDriveSafe(keg.foto3, folder, "MakanPuasa3_" + index);

    sheet.appendRow([
      reportId, new Date(), payload.email, 
      keg.jenis, keg.makan, keg.tanggal, keg.keterangan, 
      url1, url2, url3, keg.kategori || 'Umum'
    ]);
  });
  
  return { success: true, message: "Laporan PJ Makan Puasa berhasil disimpan!" };
}

function simpanPJEkstraBatch(payload) {
  if (!payload || !payload.kegiatan || payload.kegiatan.length === 0) return { success: false, message: "Data kosong" };
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Data_PJ_Ekstra');
  
  if(!sheet) {
      sheet = ss.insertSheet('Data_PJ_Ekstra');
      sheet.appendRow(['ID', 'Timestamp', 'Email', 'Jenis', 'Pendamping', 'Tanggal', 'Keterangan', 'Foto1', 'Foto2', 'Foto3', 'Foto4', 'Foto5', 'Kategori']);
  }
  
  const reportId = payload.reportId || "PJE-" + new Date().getTime();

  // --- LOGIKA KEAMANAN EDIT ---
  if (payload.reportId) {
     const emailUserAktif = Session.getActiveUser().getEmail().toLowerCase().trim();
     const isAdmin = ADMIN_EMAILS.includes(emailUserAktif);
     const data = sheet.getDataRange().getValues();
     
     let isAuthorized = false;
     for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === payload.reportId) {
           const emailPenginput = String(data[i][2]).toLowerCase().trim();
           if (emailUserAktif === emailPenginput || isAdmin) isAuthorized = true;
           break;
        }
     }
     
     if (!isAuthorized) return { success: false, message: "AKSES DITOLAK: Anda hanya dapat mengedit data yang Anda input sendiri." };

     for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === payload.reportId) sheet.deleteRow(i + 1);
     }
  }
  // ----------------------------

 const folder = DriveApp.getFolderById(FOLDER_FOTO_EKSTRA_ID);
  payload.kegiatan.forEach((keg, index) => {
    let url1 = uploadB64ToDriveSafe(keg.foto1 || keg.foto, folder, "Ekstra1_" + index);
    let url2 = uploadB64ToDriveSafe(keg.foto2, folder, "Ekstra2_" + index);
    let url3 = uploadB64ToDriveSafe(keg.foto3, folder, "Ekstra3_" + index);
    let url4 = uploadB64ToDriveSafe(keg.foto4, folder, "Ekstra4_" + index);
    let url5 = uploadB64ToDriveSafe(keg.foto5, folder, "Ekstra5_" + index);
    
    sheet.appendRow([
      reportId, new Date(), payload.email, 
      keg.jenis, keg.pendamping, keg.tanggal, keg.keterangan, url1, url2, url3, url4, url5, keg.kategori || 'Umum'
    ]);
  });
  return { success: true, message: "Laporan PJ Ekstrakurikuler berhasil disimpan!" };
}

// Helper untuk upload aman
function uploadB64ToDriveSafe(b64, folder, prefixName) {
  if (!b64) return "";
  if (b64.startsWith("http")) return b64; // Bypass jika data berupa URL dari mode Edit (sudah pernah diupload)
  if (!b64.startsWith("data:image")) return "";
  try {
    const base64Data = b64.split(',')[1];
    const contentType = b64.split(';')[0].replace('data:image/', '');
    const dataDecode = Utilities.base64Decode(base64Data);
    const fileName = prefixName + "_" + new Date().getTime() + "." + contentType;
    const blob = Utilities.newBlob(dataDecode, "image/" + contentType, fileName);
    const file = folder.createFile(blob);
    Utilities.sleep(300);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
    return file.getUrl(); 
  } catch (e) { return ""; }
}
// ==========================================
// MODUL BACKEND: RIWAYAT & PDF PJ MAKAN/EKSTRA
// ==========================================

function getRiwayatPJMaster(jenis) {
  try {
    // 1. Penentuan nama sheet berdasarkan jenis laporan [PERBAIKAN]
    let sheetName = '';
    if (jenis === 'MakanPuasa') {
        sheetName = 'Data_PJ_Makan_Puasa';
    } else if (jenis === 'Makan' || jenis === 'Snack') {
        sheetName = 'Data_PJ_Makan';
    } else if (jenis === 'Pembelian') {
        sheetName = 'Data_PJ_Pembelian';
    } else if (jenis === 'PotongFC') {
        sheetName = 'Data_PJ_PotongFC';
    } else {
        sheetName = 'Data_PJ_Ekstra';
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    data.shift(); 
    
    const grouped = {};
    data.forEach(row => {
      const id = row[0];
      if (!id) return;

      // Filter khusus agar Makan dan Snack tidak tercampur di Riwayat
      // (Untuk MakanPuasa kita lewati filter ini karena sheet-nya sudah khusus)
      if ((jenis === 'Makan' || jenis === 'Snack') && row[3] !== jenis) return;

      let rawTimestamp = row[1];
      let safeTimestamp = (rawTimestamp instanceof Date) ? rawTimestamp.toISOString() : String(rawTimestamp);

      if (!grouped[id]) {
        grouped[id] = { id: id, timestamp: safeTimestamp, email: row[2], items: [] };
      }
      
      let item = { jenis: row[3] };
      
      // 2. Pemetaan index kolom [PERBAIKAN]
      if (jenis === 'Makan' || jenis === 'Snack' || jenis === 'MakanPuasa') {
          item.makan = row[4];
          item.tanggal = (row[5] instanceof Date) ? Utilities.formatDate(row[5], "GMT+7", "yyyy-MM-dd") : String(row[5]).substring(0,10);
          item.keterangan = row[6];
          item.foto1 = row[7];
          item.foto2 = row[8];
          item.foto3 = row[9];
          item.kategori = row[10] || '-';
      } 
      else if (jenis === 'Pembelian') {
          item.tanggal = (row[4] instanceof Date) ? Utilities.formatDate(row[4], "GMT+7", "yyyy-MM-dd") : String(row[4]).substring(0,10);
          item.keterangan = row[5];
          item.foto1 = row[6];
          item.foto2 = row[7];
          item.foto3 = row[8];
          item.kategori = '-';
      } 
      else if (jenis === 'PotongFC') {
          item.tanggal = (row[4] instanceof Date) ? Utilities.formatDate(row[4], "GMT+7", "yyyy-MM-dd") : String(row[4]).substring(0,10);
          item.keterangan = row[5];
          item.foto = row[6];
          item.kategori = '-';
      } 
      else {
          // Jenis Ekstra
          item.pendamping = row[4];
          item.tanggal = (row[5] instanceof Date) ? Utilities.formatDate(row[5], "GMT+7", "yyyy-MM-dd") : String(row[5]).substring(0,10);
          item.keterangan = row[6];
          item.foto1 = row[7];
          item.foto2 = row[8];
          item.foto3 = row[9];
          item.foto4 = row[10]; // Tambahan Foto 4
          item.foto5 = row[11]; // Tambahan Foto 5
          item.kategori = row[12] || '-'; // Geser ke index 12
      }
      
      grouped[id].items.push(item);
    });
    
    return Object.values(grouped).filter(g => g.items.length > 0).sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));
  } catch(e) {
    return [];
  }
}

function hapusRiwayatPJMaster(idReport, jenis) {
  try {
    // 1. TANGKAP EMAIL USER YANG SEDANG AKTIF & CEK ADMIN
    const emailUserAktif = Session.getActiveUser().getEmail().toLowerCase().trim();
    const isAdmin = ADMIN_EMAILS.includes(emailUserAktif);

    // [PERBAIKAN] Penentuan Sheet
    let sheetName = 'Data_PJ_Ekstra';
    if (jenis === 'MakanPuasa') {
        sheetName = 'Data_PJ_Makan_Puasa';
    } else if (jenis === 'Makan' || jenis === 'Snack') {
        sheetName = 'Data_PJ_Makan';
    } else if (jenis === 'Pembelian') {
        sheetName = 'Data_PJ_Pembelian';
    } else if (jenis === 'PotongFC') {
        sheetName = 'Data_PJ_PotongFC';
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return {success: false, message: 'Sheet tidak ditemukan'};
    
    const data = sheet.getDataRange().getValues();
    
    // Looping dari bawah ke atas untuk mencari data
    for (let i = data.length - 1; i > 0; i--) {
      if (data[i][0] === idReport) {
        
        // 2. CEK OTORISASI (KEAMANAN BACKEND)
        const emailPenginput = String(data[i][2]).toLowerCase().trim(); // Email ada di index 2 (Kolom C)
        
        if (emailUserAktif !== emailPenginput && !isAdmin) {
            throw new Error("AKSES DITOLAK: Anda hanya dapat menghapus data yang Anda input sendiri.");
        }
        
        // Jika aman, eksekusi hapus baris
        sheet.deleteRow(i + 1);
      }
    }
    return {success: true, message: 'Data laporan berhasil dihapus!'};
  } catch(e) {
    return {success: false, message: 'Error: ' + e.message};
  }
}

function getDataPJForPDF(jenis, tglAwal, tglAkhir, kategoriFilter = 'Semua') {
  try {
    // [PERBAIKAN] Tambahkan MakanPuasa ke dalam grup isMakanTab
    const isMakanTab = (jenis === 'Makan' || jenis === 'Snack' || jenis === 'MakanPuasa');
    
    // 1. Tentukan Nama Sheet [PERBAIKAN]
    let sheetName = 'Data_PJ_Ekstra';
    if (jenis === 'MakanPuasa') {
        sheetName = 'Data_PJ_Makan_Puasa';
    } else if (jenis === 'Makan' || jenis === 'Snack') {
        sheetName = 'Data_PJ_Makan';
    } else if (jenis === 'Pembelian') {
        sheetName = 'Data_PJ_Pembelian';
    } else if (jenis === 'PotongFC') {
        sheetName = 'Data_PJ_PotongFC';
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    data.shift(); 
    
    let result = [];
    
    const safeGetB64 = (url) => {
        if (!url || String(url).trim() === '' || String(url).trim().toLowerCase() === '#n/a') return '';
        try {
            return getBase64FromDriveUrlSafe(url);
        } catch (e) {
            return ''; 
        }
    };

    data.forEach(row => {
      if (!row || row.join('').trim() === '') return;
      
      let idxTgl = (jenis === 'Pembelian' || jenis === 'PotongFC') ? 4 : 5;
      let tglData = row[idxTgl]; 
      if (!tglData) return;
      
      let rowDateStr = "";
      if (tglData instanceof Date) {
        let y = tglData.getFullYear();
        let m = String(tglData.getMonth() + 1).padStart(2, '0');
        let d = String(tglData.getDate()).padStart(2, '0');
        rowDateStr = `${y}-${m}-${d}`;
      } else {
        let strTgl = String(tglData).trim();
        if (strTgl.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)) {
            let parts = strTgl.split(/[\/\-]/);
            rowDateStr = `${parts[2]}-${parts[1]}-${parts[0]}`; 
        } else {
            let matchIso = strTgl.match(/^(\d{4}-\d{2}-\d{2})/);
            if(matchIso) rowDateStr = matchIso[1];
            else rowDateStr = strTgl.substring(0, 10);
        }
      }
      
      let rawKat = 'Umum';
      if (jenis === 'Ekstra') { 
          // Ekstra memiliki 5 foto, jadi kategori ada di index 12 (Kolom M)
          rawKat = row[12] !== undefined && row[12] !== '' ? row[12] : 'Umum';
      } else if (isMakanTab) {
          // Makan (Pagi, Siang, Malam, Puasa) hanya punya 3 foto, kategori ada di index 10 (Kolom K)
          rawKat = row[10] !== undefined && row[10] !== '' ? row[10] : 'Umum';
      } else {
          rawKat = row[8] || 'Umum';
      }

      let dataKategori = String(rawKat).trim().toLowerCase();
      let filterKat = String(kategoriFilter).trim().toLowerCase();
      
      if (dataKategori === 'putra' || dataKategori.includes('putra')) {
          rawKat = 'Putra';
          dataKategori = 'putra';
      } else if (dataKategori === 'putri' || dataKategori.includes('putri')) {
          rawKat = 'Putri';
          dataKategori = 'putri';
      }

      if (filterKat === 'putra') filterKat = 'putra';
      if (filterKat === 'putri') filterKat = 'putri';
      
      let matchKategori = (filterKat === 'semua' || filterKat === '' || dataKategori === filterKat);
      
      if (jenis === 'PotongFC' || jenis === 'Pembelian' || jenis === 'Ekstra') {
          matchKategori = true; 
      }

      // Filter Jenis [PERBAIKAN LOGIKA MAKAN PUASA]
      let matchJenis = true;
      if (isMakanTab) {
          if (jenis === 'MakanPuasa') {
              matchJenis = true; // Langsung true karena khusus sheet puasa
          } else {
              let jenisDiSheet = String(row[3] || '').trim().toLowerCase();
              let waktuDiSheet = String(row[4] || '').trim().toLowerCase(); 
              let jenisDiFilter = String(jenis).trim().toLowerCase();
              
              if (!jenisDiSheet.includes(jenisDiFilter) && !waktuDiSheet.includes(jenisDiFilter)) {
                  matchJenis = false;
              }
          }
      } else {
          let jenisSub = String(row[3] || '').trim().toLowerCase();
          if (filterKat !== 'semua' && filterKat !== '') {
              if (!jenisSub.includes(filterKat)) {
                  matchJenis = false;
              }
          }
      }
      
      if (rowDateStr >= tglAwal && rowDateStr <= tglAkhir && matchKategori && matchJenis) {
        let item = {
            Tanggal: rowDateStr,
            Kategori: String(rawKat)
        };
        
        if (isMakanTab) {
            item.Keterangan = String(row[6] || ''); 
            item.Makan = String(row[4] || '');
            item.b64_1 = safeGetB64(row[7]);
            item.b64_2 = safeGetB64(row[8]);
            item.b64_3 = safeGetB64(row[9]);
        } else if (jenis === 'Pembelian') {
            item.Keterangan = String(row[5] || ''); 
            item.Jenis_Ekstra = String(row[3] || '');
            item.b64_1 = safeGetB64(row[6]);
            item.b64_2 = safeGetB64(row[7]);
            item.b64_3 = safeGetB64(row[8]);
        } else if (jenis === 'PotongFC') {
            item.Keterangan = String(row[5] || ''); 
            item.Jenis_Ekstra = String(row[3] || '');
            item.b64 = safeGetB64(row[6]);          
        } else {
            item.Keterangan = String(row[6] || ''); 
            item.Jenis_Ekstra = String(row[3] || '');
            item.Pendamping = String(row[4] || '');
            item.b64_1 = safeGetB64(row[7]);
            item.b64_2 = safeGetB64(row[8]);
            item.b64_3 = safeGetB64(row[9]);
            item.b64_4 = safeGetB64(row[10]);
            item.b64_5 = safeGetB64(row[11]);
        }
        result.push(item);
      }
    });
    
    return result;
  } catch(e) {
    Logger.log("Error getDataPJForPDF: " + e.message);
    return [];
  }
}

// Helper untuk konversi gambar aman
function getBase64FromDriveUrlSafe(url) {
  if (!url) return '';
  try {
    let id = '';
    if (url.includes('/d/')) id = url.match(/\/d\/([a-zA-Z0-9_-]+)/)[1];
    else if (url.includes('id=')) id = url.match(/id=([a-zA-Z0-9_-]+)/)[1];
    if (!id) return '';
    
    let blob = DriveApp.getFileById(id).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch(e) { 
    return ''; 
  }
}
function get5DataPelanggaranTerakhir() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        
        // --- TAMBAHAN KODE: Membuat Kamus NISN ke Nama Siswa ---
        const sheetSiswa = ss.getSheetByName("DataSiswa");
        let mapSiswa = {};
        if (sheetSiswa) {
            const dataSiswa = sheetSiswa.getDataRange().getValues();
            // Mulai dari i = 2 (Baris ke-3) melewati header dan teks statis
            for (let i = 2; i < dataSiswa.length; i++) {
                let nisn = dataSiswa[i][0]; // Kolom A (NISN)
                let nama = dataSiswa[i][2]; // Kolom C (Nama Siswa)
                if (nisn) {
                    mapSiswa[String(nisn).trim()] = nama; 
                }
            }
        }
        // --------------------------------------------------------

        const sheet = ss.getSheetByName("Data_Pelanggaran");
        if (!sheet) return [];
        const data = sheet.getDataRange().getValues();
        let result = [];
        
        // Loop dari baris paling bawah
        for (let i = data.length - 1; i >= 1; i--) {
            let tglData = data[i][4];
            let tglFormat = (tglData instanceof Date) ? Utilities.formatDate(tglData, "GMT+7", "dd/MM/yyyy") : String(tglData);
            
            // --- Menerjemahkan NISN menjadi Nama ---
            let nisnPelanggar = String(data[i][1]).trim();
            let namaSiswa = mapSiswa[nisnPelanggar] || nisnPelanggar; 

            result.push({
                siswa: namaSiswa, // <-- Sekarang mengirimkan Nama, bukan NISN
                jenis: data[i][2],
                poin: data[i][3],
                tanggal: tglFormat,
                pelapor: data[i][7]
            });
            
            if (result.length >= 5) break;
        }
        return result;
    } catch(e) { return []; }
}

function get5DataKesehatanTerakhir() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        // Tetap menggunakan kamus dataSiswa sebagai fallback jika data nama belum masuk di row lama
        const sheetSiswa = ss.getSheetByName("DataSiswa");
        let mapSiswa = {};
        if (sheetSiswa) {
            const dataSiswa = sheetSiswa.getDataRange().getValues();
            for (let i = 2; i < dataSiswa.length; i++) {
                let nisn = dataSiswa[i][0]; 
                let nama = dataSiswa[i][2]; 
                if (nisn) mapSiswa[String(nisn).trim()] = nama; 
            }
        }

        const sheet = ss.getSheetByName("Data_Kesehatan");
        if (!sheet) return [];
        const data = sheet.getDataRange().getValues();
        let result = [];
        
        for (let i = data.length - 1; i >= 1; i--) {
            let tglData = data[i][5]; // Kolom F
            let tglFormat = (tglData instanceof Date) ? Utilities.formatDate(tglData, "GMT+7", "dd/MM/yyyy") : String(tglData);
            
            let nisnPasien = String(data[i][1]).trim();
            let namaDariSheet = String(data[i][2]).trim(); // Ambil nama langsung dari Kolom C
            let namaSiswa = namaDariSheet || mapSiswa[nisnPasien] || nisnPasien; 

            result.push({
                siswa: namaSiswa, 
                jenis: data[i][4],     // Kolom E
                tanggal: tglFormat,
                pelapor: data[i][9]    // Kolom J
            });
            
            if (result.length >= 5) break;
        }
        return result;
    } catch(e) { return []; }
}
function simpanTanggapanKesehatan(rowId, nisnSiswa, tanggapanBaru, namaPenanggap) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName("Data_Kesehatan");
        if (!sheet) return { success: false, message: "Sheet tidak ditemukan" };
        
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            let rId = (data[i][0] instanceof Date) ? data[i][0].getTime() : new Date(data[i][0]).getTime();
            let rNisn = String(data[i][1]).trim();
            
            if (rId.toString() === String(rowId) && rNisn === String(nisnSiswa).trim()) {
                let arrTanggapan = [];
                try {
                    if (data[i][10]) arrTanggapan = JSON.parse(data[i][10]);
                } catch(e) {}
                
                arrTanggapan.push({
                    nama: namaPenanggap,
                    isi: tanggapanBaru.isi,
                    tanggal: tanggapanBaru.tanggal
                });
                
                // Simpan kembali array JSON ke Kolom 11 (K)
                sheet.getRange(i + 1, 11).setValue(JSON.stringify(arrTanggapan));
                return { success: true };
            }
        }
        return { success: false, message: "Data keluhan spesifik tidak ditemukan." };
    } catch(e) {
        return { success: false, message: e.message };
    }
}
// ==========================================
// FUNGSI EDIT & HAPUS KESEHATAN
// ==========================================

function hapusDataKesehatan(idRow, nisn, userNama, isAdmin) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName("Data_Kesehatan");
        if (!sheet) return { success: false, message: "Sheet tidak ditemukan" };
        
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            let rId = (data[i][0] instanceof Date) ? data[i][0].getTime() : new Date(data[i][0]).getTime();
            let rNisn = String(data[i][1]).trim();
            
            if (rId.toString() === String(idRow) && rNisn === String(nisn).trim()) {
                let pelapor = data[i][9]; // Kolom J (Dilaporkan Oleh)
                if (isAdmin || pelapor === userNama) {
                    sheet.deleteRow(i + 1);
                    return { success: true, message: "Data riwayat kesehatan berhasil dihapus." };
                } else {
                    return { success: false, message: "Akses ditolak. Anda bukan pembuat data ini." };
                }
            }
        }
        return { success: false, message: "Data tidak ditemukan." };
    } catch(e) {
        return { success: false, message: e.message };
    }
}

function updateDataKesehatan(payload, userNama, isAdmin) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName("Data_Kesehatan");
        if (!sheet) return { success: false, message: "Sheet tidak ditemukan" };
        
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            let rId = (data[i][0] instanceof Date) ? data[i][0].getTime() : new Date(data[i][0]).getTime();
            let rNisn = String(data[i][1]).trim();
            
            if (rId.toString() === String(payload.idRow) && rNisn === String(payload.nisn).trim()) {
                let pelapor = data[i][9];
                if (isAdmin || pelapor === userNama) {
                    sheet.getRange(i + 1, 4).setValue(payload.kelas);      // Kolom D
                    sheet.getRange(i + 1, 5).setValue(payload.jenis);      // Kolom E
                    sheet.getRange(i + 1, 6).setValue(payload.tanggal);    // Kolom F
                    sheet.getRange(i + 1, 7).setValue(payload.ttv);        // Kolom G
                    sheet.getRange(i + 1, 8).setValue(payload.keterangan); // Kolom H
                    sheet.getRange(i + 1, 9).setValue(payload.rujukan);    // Kolom I
                    return { success: true, message: "Data riwayat kesehatan berhasil diupdate." };
                } else {
                    return { success: false, message: "Akses ditolak. Anda bukan pembuat data ini." };
                }
            }
        }
        return { success: false, message: "Data tidak ditemukan." };
    } catch(e) {
        return { success: false, message: e.message };
    }
}

function getJadwalPiketHariIni() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName("Jadwal_Piket");
        if (!sheet) return null;

        const data = sheet.getDataRange().getValues();
        
        // Ambil nama bulan dari sel B1 (Baris 1, Kolom B pada array index adalah [0][1])
        const bulanSheet = String(data[0][2] || "").trim();

        // 1. Dapatkan tanggal hari ini dan besok (WIB GMT+7)
        const hariIni = new Date();
        const tglHariIni = Number(Utilities.formatDate(hariIni, "GMT+7", "d")); 
        
        const besok = new Date(hariIni.getTime() + 24 * 60 * 60 * 1000);
        const tglBesok = Number(Utilities.formatDate(besok, "GMT+7", "d"));

        // 2. Cari indeks kolom untuk tanggal hari ini dan besok
        const barisTanggal = data[1]; 
        let indexKolomHariIni = -1;
        let indexKolomBesok = -1;
        
        for (let c = 2; c < barisTanggal.length; c++) {
            if (barisTanggal[c] == tglHariIni) indexKolomHariIni = c;
            if (barisTanggal[c] == tglBesok) indexKolomBesok = c;
        }

        // 3. Tambahkan properti bulan ke dalam objek jadwal
        let jadwal = {
            hariIni: { P: [], S: [], M: [], L: [], LP: [], DL: [], C: [] },
            besok: { P: [], S: [], M: [], L: [], LP: [], DL: [], C: [] },
            bulan: bulanSheet // <--- Data bulan dimasukkan di sini
        };

        // 4. Looping data personil
        for (let r = 2; r < data.length; r++) {
            let nama = data[r][1]; // Kolom B (Nama)
            if (!nama) continue;

            // Masukkan shift Hari Ini
            if (indexKolomHariIni !== -1 && data[r][indexKolomHariIni]) {
                let kodeIni = String(data[r][indexKolomHariIni]).trim().toUpperCase();
                if (jadwal.hariIni[kodeIni] !== undefined) jadwal.hariIni[kodeIni].push(nama);
            }
            
            // Masukkan shift Besok
            if (indexKolomBesok !== -1 && data[r][indexKolomBesok]) {
                let kodeBesok = String(data[r][indexKolomBesok]).trim().toUpperCase();
                if (jadwal.besok[kodeBesok] !== undefined) jadwal.besok[kodeBesok].push(nama);
            }
        }

        return jadwal;

    } catch (e) {
        return { error: e.message };
    }
}
// ==========================================
// FUNGSI SIMPAN GENERATE PDF PJ KE DRIVE
// ==========================================
function simpanPDFPJKeDrive(base64Data, fileName, jenis, start, end, namaUser) {
  try {
    // ---- 1. SIMPAN KE ROOT DRIVE USER (FUNGSI ASLI) ----
    const root = DriveApp.getRootFolder();
    const folderUtama = getFolderByNameOrCreate(root, "Laporan Harian");
    const folderUser = getFolderByNameOrCreate(folderUtama, namaUser);
    const folderLaporanPJ = getFolderByNameOrCreate(folderUser, "Laporan PJ");
    
    let subFolderName = jenis; 
    if (jenis === 'Ekstra') {
        subFolderName = `Ekstra_${start}_SD_${end}`; 
    }
    const folderKegiatan = getFolderByNameOrCreate(folderLaporanPJ, subFolderName);
    
    saveFileWithOverwrite(folderKegiatan, fileName, base64Data);
    
    // ---- 2. SIMPAN PARALEL KE DRIVE ID SPESIFIK (TAMBAHAN) ----
    const idBackup = '1quTf_wLAPnzta9EY_7dBsRE9RvmmHfAc';
    const rootBackup = DriveApp.getFolderById(idBackup);
    
    // Membuat struktur folder yang persis sama di dalam Folder Backup
    const folderUtamaBackup = getFolderByNameOrCreate(rootBackup, "Laporan Harian");
    const folderUserBackup = getFolderByNameOrCreate(folderUtamaBackup, namaUser);
    const folderLaporanPJBackup = getFolderByNameOrCreate(folderUserBackup, "Laporan PJ");
    const folderKegiatanBackup = getFolderByNameOrCreate(folderLaporanPJBackup, subFolderName);
    
    // Simpan file ke folder Backup
    saveFileWithOverwrite(folderKegiatanBackup, fileName, base64Data);
    
    return { 
      success: true, 
      message: `Folder: Laporan Harian > ${namaUser} > Laporan PJ > ${subFolderName} (Tersimpan Paralel)` 
    };
  } catch (e) {
    throw new Error(e.toString());
  }
}
// ==========================================
// BACKEND PJ PEMBELIAN & POTONG/FC
// ==========================================

function simpanPJPembelianBatch(payload) {
  if (!payload || !payload.kegiatan || payload.kegiatan.length === 0) return { success: false, message: "Data kosong" };
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Data_PJ_Pembelian');
  if(!sheet) {
      sheet = ss.insertSheet('Data_PJ_Pembelian');
      sheet.appendRow(['ID', 'Timestamp', 'Email', 'Jenis', 'Tanggal', 'Keterangan', 'Foto1', 'Foto2', 'Foto3']);
  }
  
  const reportId = payload.reportId || "PJP-" + new Date().getTime(); 

  // --- LOGIKA KEAMANAN EDIT ---
  if (payload.reportId) {
     const emailUserAktif = Session.getActiveUser().getEmail().toLowerCase().trim();
     const isAdmin = ADMIN_EMAILS.includes(emailUserAktif);
     const data = sheet.getDataRange().getValues();
     
     let isAuthorized = false;
     for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === payload.reportId) {
           const emailPenginput = String(data[i][2]).toLowerCase().trim();
           if (emailUserAktif === emailPenginput || isAdmin) isAuthorized = true;
           break;
        }
     }
     
     if (!isAuthorized) return { success: false, message: "AKSES DITOLAK: Anda hanya dapat mengedit data yang Anda input sendiri." };

     for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === payload.reportId) sheet.deleteRow(i + 1);
     }
  }
  // ----------------------------
  
  const folder = DriveApp.getFolderById(FOLDER_FOTO_EKSTRA_ID); 
  
  payload.kegiatan.forEach((keg, index) => {
    let url1 = uploadB64ToDriveSafe(keg.foto1, folder, "Pembelian1_" + index);
    let url2 = uploadB64ToDriveSafe(keg.foto2, folder, "Pembelian2_" + index);
    let url3 = uploadB64ToDriveSafe(keg.foto3, folder, "Pembelian3_" + index);

    sheet.appendRow([reportId, new Date(), payload.email, keg.jenis, keg.tanggal, keg.keterangan, url1, url2, url3]);
  });
  return { success: true, message: "Laporan Pembelian berhasil disimpan!" };
}

function simpanPJPotongFCBatch(payload) {
  if (!payload || !payload.kegiatan || payload.kegiatan.length === 0) return { success: false, message: "Data kosong" };
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Data_PJ_PotongFC');
  if(!sheet) {
      sheet = ss.insertSheet('Data_PJ_PotongFC');
      sheet.appendRow(['ID', 'Timestamp', 'Email', 'Jenis', 'Tanggal', 'Keterangan', 'Foto']);
  }
  
  const reportId = payload.reportId || "PJC-" + new Date().getTime(); 

  // --- LOGIKA KEAMANAN EDIT ---
  if (payload.reportId) {
     const emailUserAktif = Session.getActiveUser().getEmail().toLowerCase().trim();
     const isAdmin = ADMIN_EMAILS.includes(emailUserAktif);
     const data = sheet.getDataRange().getValues();
     
     let isAuthorized = false;
     for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === payload.reportId) {
           const emailPenginput = String(data[i][2]).toLowerCase().trim();
           if (emailUserAktif === emailPenginput || isAdmin) isAuthorized = true;
           break;
        }
     }
     
     if (!isAuthorized) return { success: false, message: "AKSES DITOLAK: Anda hanya dapat mengedit data yang Anda input sendiri." };

     for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === payload.reportId) sheet.deleteRow(i + 1);
     }
  }
  // ----------------------------
  
  const folder = DriveApp.getFolderById(FOLDER_FOTO_EKSTRA_ID); 
  
  payload.kegiatan.forEach((keg, index) => {
    let url = uploadB64ToDriveSafe(keg.foto, folder, "PotongFC_" + index);
    sheet.appendRow([reportId, new Date(), payload.email, keg.jenis, keg.tanggal, keg.keterangan, url]);
  });
  return { success: true, message: "Laporan Potong/FC berhasil disimpan!" };
}

// ==========================================
// FUNGSI EXPORT EXCEL KESEHATAN
// ==========================================
function exportExcelKesehatan(bulanInput) { // Format: YYYY-MM
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetKes = ss.getSheetByName('Data_Kesehatan');
    if (!sheetKes) return { success: false, message: "Sheet Data_Kesehatan tidak ditemukan!" };
    
    const data = sheetKes.getDataRange().getValues();
    
    // Buat spreadsheet sementara untuk diexport
    const tempSS = SpreadsheetApp.create("Rekap_Kesehatan_" + bulanInput);
    const sheet = tempSS.getSheets()[0];
    sheet.setName("Laporan_Kesehatan");
    
    // Header KOP Excel
    sheet.getRange("A1").setValue("REKAP DATA KESEHATAN SISWA");
    sheet.getRange("A2").setValue("Bulan: " + bulanInput);
    sheet.getRange("A1:A2").setFontWeight("bold");
    
    // Setup Tabel Header
    const headers = ["No", "Tanggal", "NISN", "Nama Siswa", "Kelas", "Jenis Keluhan", "TTV", "Catatan/Tindakan", "Rujukan", "Dilaporkan Oleh"];
    sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(4, 1, 1, headers.length).setFontWeight("bold").setBackground("#e2e8f0");
    
    let rowIdx = 5;
    let no = 1;
    let outputData = [];
    
    // Pecah YYYY-MM
    let [fYear, fMonth] = bulanInput.split('-');

    for (let i = 1; i < data.length; i++) {
      let tglData = data[i][5]; // Kolom F (Tanggal)
      if (!tglData) continue;
      
      // Amankan format tanggal
      let tglFormat = (tglData instanceof Date) ? Utilities.formatDate(tglData, "GMT+7", "yyyy-MM-dd") : String(tglData);
      
      // Jika formatnya DD/MM/YYYY ubah ke YYYY-MM-DD
      if(tglFormat.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)) {
          let parts = tglFormat.split(/[\/\-]/);
          tglFormat = `${parts[2]}-${parts[1]}-${parts[0]}`; 
      }
      
      // Filter berdasarkan Tahun dan Bulan
      if (tglFormat.startsWith(`${fYear}-${fMonth}`)) {
         let displayDate = (tglData instanceof Date) ? Utilities.formatDate(tglData, "GMT+7", "dd/MM/yyyy") : String(tglData);
         
         let rowData = [
           no++, 
           displayDate,
           data[i][1], // NISN
           data[i][2], // Nama
           data[i][3], // Kelas
           data[i][4], // Jenis Keluhan
           data[i][6], // TTV
           data[i][7], // Keterangan/Catatan
           data[i][8], // Rujukan
           data[i][9]  // Dilaporkan Oleh
         ];
         outputData.push(rowData);
      }
    }
    
    // Inject Data
    if(outputData.length > 0) {
        sheet.getRange(rowIdx, 1, outputData.length, headers.length).setValues(outputData);
        
        // Mempercantik border dan auto-resize kolom
        sheet.getRange(4, 1, outputData.length + 1, headers.length).setBorder(true, true, true, true, true, true);
        for(let c = 1; c <= headers.length; c++) {
          sheet.autoResizeColumn(c);
        }
    } else {
        sheet.getRange(rowIdx, 1).setValue("Tidak ada data keluhan kesehatan pada bulan ini.");
    }
    
    SpreadsheetApp.flush();
    
    // Hasilkan link download xlsx langsung dari ID temp spreadsheet
    return { 
      url: "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx" 
    };
    
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Dipakai khusus untuk menarik data RHK utama (yang bertipe 'Global') saja untuk dropdown di menu Pendahuluan.
function getRHKDataDariSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);    
  let dataRHK = [];

  // 1. Ambil data Wali Asuh 
  const sheetAsuh = ss.getSheetByName('RHK_Master_Wali_Asuh');
  if(sheetAsuh && sheetAsuh.getLastRow() > 1) {
    const dataAsuh = sheetAsuh.getDataRange().getValues();
    for(let i = 1; i < dataAsuh.length; i++) {
       if(dataAsuh[i][0] && dataAsuh[i][1] === 'Global') { 
         dataRHK.push({
           id: dataAsuh[i][0],
           tipe: dataAsuh[i][1],
           nama: dataAsuh[i][2],
           jenisMaster: 'Wali Asuh'
         });
       }
    }
  }

  // 2. Ambil data Wali Asrama
  const sheetAsrama = ss.getSheetByName('RHK_Master_Wali_Asrama');
  if(sheetAsrama && sheetAsrama.getLastRow() > 1) {
    const dataAsrama = sheetAsrama.getDataRange().getValues();
    for(let i = 1; i < dataAsrama.length; i++) {
       if(dataAsrama[i][0] && dataAsrama[i][1] === 'Global') {
         dataRHK.push({
           id: dataAsrama[i][0],
           tipe: dataAsrama[i][1],
           nama: dataAsrama[i][2],
           jenisMaster: 'Wali Asrama'
         });
       }
    }
  }

  return dataRHK;
}
function prosesDuplikatPendahuluan(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Pendahuluan_Master');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      const newId = new Date().getTime();
      const newNama = data[i][1] + " (Copy)";
      const isi = data[i][2];
      const rhkLink = data[i][3];
      
      sheet.appendRow([newId, newNama, isi, rhkLink]);
      return { success: true, message: "Template berhasil diduplikat!" };
    }
  }
  return { success: false, message: "Gagal menduplikat, data asli tidak ditemukan." };
}

// FUNGSI HAPUS SAMPAH DI DRIVE UNTUK ID YANG KEPAKAI
function cleanUpDriveFiles() {
  try {
    // ID Folder yang akan dibersihkan
    var folderId = "1bNll1RtEYNRkSXdlu4tHB-yIAYHzZisw"; 
    var folder = DriveApp.getFolderById(folderId);
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var usedFileIds = new Set();
    
    // 1. TAMBAHKAN NAMA-NAMA SHEET YANG INGIN DICEK DI SINI
    var daftarSheet = [
      'Laporan_PJ',
      'Data_Kegiatan',
      'Data_PJ_Makan',
      'Data_PJ_Ekstra',
      'Data_PJ_Pembelian',
      'Data_PJ_PotongFC',
      'Data_Pelanggaran',
      'Data_DU_CS_Security' // Sesuaikan/tambahkan nama sheet lainnya jika ada
    ];
    
    // 2. Kumpulkan semua ID file/URL dari semua sheet yang terdaftar di atas
    daftarSheet.forEach(function(namaSheet) {
      var sheet = ss.getSheetByName(namaSheet);
      if (sheet) {
        var data = sheet.getDataRange().getValues();
        // Mulai dari baris ke-2 (index 1) untuk melewati Header
        for (var i = 1; i < data.length; i++) {
          for (var j = 0; j < data[i].length; j++) {
            var val = String(data[i][j]);
            // Jika val terlihat seperti URL / ID Google Drive
            if (val.length > 20 && !val.includes(" ")) {
              usedFileIds.add(val);
            }
          }
        }
      }
    });

    // 3. Scan file di dalam folder Drive
    var files = folder.getFiles();
    var jumlahTerhapus = 0;

    while (files.hasNext()) {
      var file = files.next();
      var fileId = file.getId();
      var fileUrl = file.getUrl();
      var mimeType = file.getMimeType();

      // PROTEKSI: Lewati file Google Sheets, Apps Script, dan Google Sites
      if (mimeType === MimeType.GOOGLE_SHEETS || 
          mimeType === MimeType.GOOGLE_APPS_SCRIPT || 
          mimeType === MimeType.GOOGLE_SITES) {
        continue; 
      }

      // Jika ID atau URL file TIDAK ADA di satupun sheet, pindahkan ke sampah
      if (!usedFileIds.has(fileId) && !usedFileIds.has(fileUrl)) {
        file.setTrashed(true);
        jumlahTerhapus++;
      }
    }

    return { 
      success: true, 
      message: "Pembersihan selesai! Berhasil menghapus " + jumlahTerhapus + " file foto sampah yang tidak digunakan." 
    };

  } catch (error) {
    return { 
      success: false, 
      message: "Gagal membersihkan file: " + error.message 
    };
  }
}
function hapusSemuaDataNotifikasi() {
  try {
    // Membuka spreadsheet secara spesifik menggunakan ID
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Notifikasi"); 
    
    // Validasi apakah sheet ditemukan
    if (!sheet) {
      return { success: false, message: "Sheet dengan nama 'Notifikasi' tidak ditemukan." };
    }
    
    const lastRow = sheet.getLastRow();
    
    // Menghapus data jika ada baris selain header (baris 1)
    if (lastRow > 1) {
      // Menghapus mulai dari baris ke-2, sebanyak (lastRow - 1) baris
      sheet.deleteRows(2, lastRow - 1);
    } else {
       return { success: true, message: "Data sudah kosong, tidak ada yang perlu dihapus." };
    }
    
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Fungsi Assesmen boooy v
function getStudentData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('DataSiswa'); 
  
  // Mengambil data mulai dari baris ke-3 (karena baris 1 header, baris 2 kosong/warna)
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  
  var data = sheet.getRange(3, 1, lastRow - 2, 14).getValues();
  
  var students = data.map(function(row) {
    var namaSiswa = row[2] ? row[2].toString().trim() : ""; // Kolom C (Nama)
    var nisn = row[0] ? row[0].toString().trim() : "";      // Kolom A (NISN)
    
    // Gabungkan Nama dan NISN. Jika NISN kosong, tampilkan nama saja.
    var namaLengkap = nisn ? namaSiswa + " (" + nisn + ")" : namaSiswa;
    
    return {
      nama: namaLengkap, // Tampil sebagai "NAMA (NISN)"
      kelas: row[11],    // Kolom L (Kelas)
      wali: row[12],     // Kolom M (Wali Asuh)
      kamar: row[13]     // Kolom N (Kamar)
    };
  }).filter(function(s) { 
    // Filter agar siswa yang namanya kosong tidak masuk ke daftar
    return s.nama && s.nama.toString().trim() !== ""; 
  });
  
  return students;
}

// Fungsi untuk menyimpan data asesmen lengkap ke Google Sheets
function submitForm(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('HasilAssesment');
  
  if (!sheet) {
    sheet = ss.insertSheet('HasilAssesment');
    var headers = [
      "Tanggal Input", "Nama Siswa", "Kelas", "Asrama/Kamar", "Wali Asuh", "Tanggal Asesmen",
      "ASPEK BIOLOGIS", "ASPEK PSIKOLOGIS", "ASPEK SOSIAL", "ASPEK SPIRITUAL", "ASPEK KELUARGA", "Penilai"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9ead3");
  }

  var data = sheet.getDataRange().getValues();
  var targetRow = -1;

  // Format bulan & tahun dari tanggal asesmen baru untuk pengecekan periode
  var tglInputBaru = new Date(payload.tanggal);
  var bulanBaru = tglInputBaru.getMonth();
  var tahunBaru = tglInputBaru.getFullYear();

  // 1. Cek jika dikirim rowIndex dari frontend (mode edit langsung)
  if (payload.rowIndex && payload.rowIndex > 1) {
    targetRow = payload.rowIndex;
  } else {
    // 2. Cek otomatis berdasarkan Nama Siswa dan Periode Bulan/Tahun yang sama
    for (var i = 1; i < data.length; i++) {
      var namaSheet = String(data[i][1]).trim().toLowerCase();
      var namaPayload = String(payload.nama).trim().toLowerCase();
      
      if (namaSheet === namaPayload && data[i][5]) {
        var tglSheet = new Date(data[i][5]);
        if (tglSheet.getMonth() === bulanBaru && tglSheet.getFullYear() === tahunBaru) {
          targetRow = i + 1; // Baris ditemukan di sheet (indeks mulai dari 1)
          break;
        }
      }
    }
  }

  var rowData = [
    new Date(),             // Tanggal Input / Update
    payload.nama,           
    payload.kelas,          
    payload.kamar,          
    payload.wali,           
    payload.tanggal,        
    payload.biologis,       
    payload.psikologis,     
    payload.sosial,         
    payload.spiritual,      
    payload.keluarga,       
    payload.catatan_wali,
    payload.user            // Penilai (bisa memperbarui penilai jika diedit user lain pada bulan yang sama)
  ];

  if (targetRow > 1) {
    // Timpa/Update data pada baris yang sama (mencegah duplikasi bulan yang sama)
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    return { success: true, message: "Data asesmen berhasil diperbarui (update)!" };
  } else {
    // Tambah baris baru jika belum ada data di periode bulan tersebut
    sheet.appendRow(rowData);
    return { success: true, message: "Data asesmen berhasil disimpan!" };
  }
}
// Fungsi untuk mengambil riwayat asesmen yang disimpan ke UI
function getHistoryAssesment() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('HasilAssesment');
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  let result = [];
  
  data.forEach((row, index) => {
    let tglFormat = (row[5] instanceof Date) ? Utilities.formatDate(row[5], "GMT+7", "yyyy-MM-dd") : String(row[5] || '');
    
    result.push({
      rowIndex: index + 2, // Baris asli di Google Sheets
      tanggal: tglFormat,
      nama: row[1],
      kelas: row[2],
      kamar: row[3],
      wali: row[4],
      biologis: row[6],
      psikologis: row[7],
      sosial: row[8],
      spiritual: row[9],
      keluarga: row[10],
      catatan_wali: row[11],
      penilai: row[12]
    });
  });
  
  return result.reverse();
}

function simpanPDFAsesmenKeDrive(base64Data, fileName) {
  var parentFolderId = '1YeZugs9o4N5T4S1A8ydyQE_YH8cSfRLL';
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  
  // Buat nama folder berdasarkan Bulan dan Tahun saat ini (Contoh: Agustus_2026_assesmen)
  var monthsIndo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  var now = new Date();
  var folderName = monthsIndo[now.getMonth()] + "_" + now.getFullYear() + "_assesmen";
  
  var targetFolder;
  var folders = parentFolder.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    targetFolder = folders.next();
  } else {
    targetFolder = parentFolder.createFolder(folderName);
  }
  
  // Konversi Base64 ke Blob file PDF
  var decodedBytes = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(decodedBytes, 'application/pdf', fileName);
  
  // Cek jika file dengan nama sama sudah ada di dalam folder, hapus atau timpa agar tidak duplikat menumpuk
  var existingFiles = targetFolder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }
  
  var savedFile = targetFolder.createFile(blob);
  
  return {
    success: true,
    message: "File tersimpan di folder: " + folderName + " / " + fileName,
    fileUrl: savedFile.getUrl()
  };
}
function deleteAssessmentByRow(rowIndex) {
  try {

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('HasilAssesment');
    
    if (!sheet) {
      throw new Error("Sheet dengan nama '" + sheetName + "' tidak ditemukan.");
    }
    
    sheet.deleteRow(rowIndex);
    
    return { status: 'success', message: 'Data berhasil dihapus.' };
  } catch (error) {
    throw new Error(error.message);
  }
}
function getUniqueJenisEkstra() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Data_PJ_Ekstra");
  
  if (!sheet) return []; // Mencegah error jika nama sheet salah/tidak ada
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // Mencegah error jika sheet masih kosong
  
  // Ambil data dari baris 2 ke bawah pada kolom D (Indeks kolom 4)
  const data = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  
  // Filter data kosong dan ambil nilai uniknya saja
  const uniqueJenis = [...new Set(data.map(row => row[0]).filter(val => val !== ""))];
  return uniqueJenis.sort();
}


// Konfigurasi ID Folder & Sheet untuk CS, DU, & Security

/**
 * Menyimpan data laporan baru dan menyimpan foto ke Drive
 */
function simpanLaporanCSDUSecurity(payload) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_CSDU_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_CSDU_NAME);
      // Kolom 12 (L) & 13 (M) untuk Laporan Ke & Jam
      sheet.appendRow(['ID', 'Tanggal', 'Kategori', 'Kegiatan', 'Nama Pelapor', 'Keterangan', 'Foto 1', 'Foto 2', 'Foto 3', 'Timestamp', 'Status', 'Laporan Ke', 'Jam']);
    }

    let photoUrls = ['', '', ''];
    if (payload.fotos && payload.fotos.length > 0) {
      const folderFoto = DriveApp.getFolderById(FOLDER_FOTO_CSDU_ID);
      payload.fotos.forEach((base64Str, idx) => {
        if (idx < 3) {
          const splitData = base64Str.split(',');
          const contentType = splitData[0].match(/:(.*?);/)[1];
          const fileData = Utilities.base64Decode(splitData[1]);
          const blob = Utilities.newBlob(fileData, contentType, `CSDU_${new Date().getTime()}_${idx+1}.jpg`);
          const savedFile = folderFoto.createFile(blob);
          try { savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
          photoUrls[idx] = savedFile.getUrl();
        }
      });
    }

    const reportId = 'CSDU-' + new Date().getTime();
    sheet.appendRow([
      reportId, payload.tanggal, payload.kategori, payload.kegiatan, payload.namaPelapor, 
      payload.keterangan || '', photoUrls[0], photoUrls[1], photoUrls[2], new Date(), 
      'Draft', payload.laporanKe, payload.jam
    ]);
    return { success: true, message: 'Data berhasil disimpan!' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Mengambil seluruh data laporan CS, DU, Security
 */
function getLaporanCSDUSecurity() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CSDU_NAME);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const displayData = sheet.getDataRange().getDisplayValues(); 
  if (data.length <= 1) return [];

  let result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let fotos = [];
    if (row[6]) fotos.push(row[6]);
    if (row[7]) fotos.push(row[7]);
    if (row[8]) fotos.push(row[8]);

    // Format Jam menjadi HH:mm agar dapat dibaca oleh input type="time"
    let jamVal = displayData[i][12] || row[12] || '';
    if (jamVal) {
      if (jamVal.includes(':')) {
        let parts = jamVal.split(':');
        let h = parts[0].padStart(2, '0');
        let m = parts[1].padStart(2, '0');
        jamVal = `${h}:${m}`;
      }
    }

    result.push({
      rowIndex: i + 1,
      id: row[0],
      tanggal: row[1] ? Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "yyyy-MM-dd") : '',
      kategori: row[2],
      kegiatan: row[3],
      namaPelapor: row[4],
      keterangan: row[5],
      fotoUrls: fotos,
      status: row[10] ? String(row[10]).trim() : 'Draft',
      laporanKe: row[11] || '1',
      jam: jamVal
    });
  }

  return result.reverse();
}

/**
 * Memperbarui data laporan
 */
function updateLaporanCSDUSecurity(payload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CSDU_NAME);
  const rIndex = parseInt(payload.rowIndex);

  // 1. Update data teks biasa termasuk Laporan Ke (Kolom L) & Jam (Kolom M)
  sheet.getRange(rIndex, 2).setValue(payload.tanggal);
  sheet.getRange(rIndex, 3).setValue(payload.kategori);
  sheet.getRange(rIndex, 4).setValue(payload.kegiatan);
  sheet.getRange(rIndex, 5).setValue(payload.namaPelapor);
  sheet.getRange(rIndex, 6).setValue(payload.keterangan);
  sheet.getRange(rIndex, 11).setValue('Draft'); // Reset status jadi Draft saat diedit
  sheet.getRange(rIndex, 12).setValue(payload.laporanKe);
  sheet.getRange(rIndex, 13).setValue("'" + payload.jam); // Tambah petik agar format jam aman

  // 2. Proses dan update foto jika ada data fotos dari payload
  if (payload.fotos) {
    let fotoUrls = ['', '', ''];

    payload.fotos.forEach((item, idx) => {
      if (idx < 3) {
        if (item && item.startsWith('http')) {
          fotoUrls[idx] = item;
        } else if (item && item.startsWith('data:image')) {
          try {
            const parts = item.split(',');
            const contentType = parts[0].match(/:(.*?);/)[1];
            const decodedData = Utilities.base64Decode(parts[1]);
            const fileName = `foto_edit_${payload.kategori}_${new Date().getTime()}_${idx + 1}.jpg`;
            const blob = Utilities.newBlob(decodedData, contentType, fileName);
            
            const folder = DriveApp.getFolderById(FOLDER_FOTO_CSDU_ID); 
            const file = folder.createFile(blob);
            try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
            
            fotoUrls[idx] = file.getUrl();
          } catch (err) {}
        }
      }
    });

    sheet.getRange(rIndex, 7).setValue(fotoUrls[0]);
    sheet.getRange(rIndex, 8).setValue(fotoUrls[1]);
    sheet.getRange(rIndex, 9).setValue(fotoUrls[2]);
  }

  return { success: true };
}

/**
 * Mengambil nama pelapor unik untuk Autotext
 */
function getUniquePelaporCSDUSecurity() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CSDU_NAME);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let names = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][4]) names.add(data[i][4].toString().trim());
  }
  return Array.from(names);
}

/**
 * Mengambil nama kegiatan unik untuk Autotext
 */
function getUniqueKegiatanCSDUSecurity() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CSDU_NAME);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let kegs = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][3]) kegs.add(data[i][3].toString().trim());
  }
  return Array.from(kegs);
}

/**
 * Generate PDF dan Menyimpan ke Folder Drive PDF serta Update Status
 */
function generatePDFCSDUSecurity(dataItem) {
  try {
    const emailPembuat = Session.getActiveUser().getEmail();
    const userRoleInfo = checkUserStatus(); 
    const userConfig = getPengaturanUser(emailPembuat);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_CSDU_NAME);
    const data = sheet.getDataRange().getValues();
    const displayData = sheet.getDataRange().getDisplayValues();
    
    let kumpulanLaporan = [];

    const bulanIndo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    let tglObj = new Date(dataItem.tanggal);
    let namaBulan = bulanIndo[tglObj.getMonth()];
    let tahun = tglObj.getFullYear();
    
    for (let i = 1; i < data.length; i++) {
      let tglSheet = data[i][1] ? Utilities.formatDate(new Date(data[i][1]), "GMT+7", "yyyy-MM-dd") : '';
      
      if (tglSheet === dataItem.tanggal && data[i][2] === dataItem.kategori) {
        let jamStr = displayData[i][12] || data[i][12] || '-';
        if (jamStr.includes(':')) jamStr = jamStr.substring(0, 5).replace(':', '.');

        let petugasRaw = data[i][4] ? String(data[i][4]).split('\n') : [];
        let petugasHtml = petugasRaw.map((p, idx) => petugasRaw.length > 1 ? `${idx + 1}. ${p.trim()}` : p.trim()).join('<br/>');

        let kegiatanRaw = data[i][3] ? String(data[i][3]) : '';
        let ketRaw = data[i][5] ? String(data[i][5]) : '';
        let gabungKegiatan = [kegiatanRaw, ketRaw].filter(Boolean).join('\n').split('\n');
        let kegiatanHtml = gabungKegiatan.map(k => `- ${k.trim()}`).join('<br/>');

        kumpulanLaporan.push({
          laporanKe: parseInt(data[i][11]) || 99,
          jam: jamStr,
          petugas: petugasHtml,
          kegiatan: kegiatanHtml,
          foto1: data[i][6] ? getBase64FromDriveUrlSafe(data[i][6]) : '',
          foto2: data[i][7] ? getBase64FromDriveUrlSafe(data[i][7]) : '',
          foto3: data[i][8] ? getBase64FromDriveUrlSafe(data[i][8]) : '',
          rowIndex: i + 1
        });
      }
    }
    
    kumpulanLaporan.sort((a, b) => a.laporanKe - b.laporanKe);

    let tableRows = '';
    kumpulanLaporan.forEach((item, index) => {
      let fotosHtml = '';
      [item.foto1, item.foto2, item.foto3].forEach(b64 => {
        if(b64) fotosHtml += `<img src="${b64}" style="width: 130px; height: 95px; object-fit: cover; border: 1px solid #444; margin: 2px;">`;
      });

      let tanggalFormatted = `${tglObj.getDate()} ${namaBulan} ${tahun}`;

      tableRows += `
        <tr>
          <td style="text-align: center;">${index + 1}</td>
          <td style="text-align: center;">${tanggalFormatted}</td>
          <td style="text-align: center;">${item.jam}</td>
          <td>${item.petugas}</td>
          <td>${item.kegiatan}</td>
          <td style="text-align: center; vertical-align: middle;">${fotosHtml}</td>
        </tr>
      `;
    });

    // Perbaikan KOP agar bersih dari string JSON mentah
    let kopText = userConfig.kop.barisKop || '';
    if (kopText.startsWith('[')) {
        try {
            let parsed = JSON.parse(kopText);
            kopText = parsed.map(p => `<span style="font-size:${p.ukuran || 12}px; ${p.isBold ? 'font-weight:bold;' : ''}">${p.teks}</span>`).join('<br/>');
        } catch(e) {
            kopText = `<b>KEMENTERIAN SOSIAL REPUBLIK INDONESIA</b><br/><b>PUSDIKLATBANGPROFKESOS</b><br/><b>SEKOLAH RAKYAT MENENGAH ATAS 16 TEMANGGUNG</b><br/>Jalan Kartini No.1-2 Bendo Kertosari, Temanggung`;
        }
    } else if (!kopText) {
        kopText = `<b>KEMENTERIAN SOSIAL REPUBLIK INDONESIA</b><br/><b>PUSDIKLATBANGPROFKESOS</b><br/><b>SEKOLAH RAKYAT MENENGAH ATAS 16 TEMANGGUNG</b><br/>Jalan Kartini No.1-2 Bendo Kertosari, Temanggung`;
    }

    let kopHtml = `
      <table style="width:100%; border:none; margin-bottom:10px;">
        <tr>
          <td style="width:15%; border:none; text-align:center; vertical-align:middle;">
            ${userConfig.kop.logoKiri ? `<img src="${userConfig.kop.logoKiri}" style="max-height:85px;">` : ''}
          </td>
          <td style="width:70%; border:none; text-align:center; line-height:1.2; font-family:Arial, sans-serif;">
             ${kopText}
          </td>
          <td style="width:15%; border:none; text-align:center; vertical-align:middle;">
            ${userConfig.kop.logoKanan ? `<img src="${userConfig.kop.logoKanan}" style="max-height:85px;">` : ''}
          </td>
        </tr>
      </table>
      <div style="border-bottom:3px solid #000; margin-bottom:15px; margin-top:5px;"></div>
      <div style="text-align:center; margin-bottom: 15px;">
        <h3 style="margin:0; font-size:14px; font-weight:bold;">
           Laporan Kegiatan ${dataItem.kategori} Bulan ${namaBulan} Tahun ${tahun}
        </h3>
      </div>`;

    let ttdHtml = '';
    if(userConfig.ttd && userConfig.ttd.namaTtd) {
      let imgTtd = userConfig.ttd.tandaTangan ? `<img src="${userConfig.ttd.tandaTangan}" style="max-height:75px; margin: 5px 0;">` : '<br/><br/><br/>';
      ttdHtml = `
      <div style="width: 280px; float: right; text-align: center; margin-top: 25px; font-size: 12px; font-family:Arial, sans-serif;">
        Kepala SRMA 16 Temanggung<br/>
        ${imgTtd}<br/>
        <b><u>${userConfig.ttd.namaTtd}</u></b><br/>
        NIP. ${userConfig.ttd.nipTtd || '-'}
      </div>
      <div style="clear:both;"></div>
      `;
    }

    const htmlContent = `
      <html>
        <head>
          <style>
            @page { size: A4 landscape; margin: 1cm; }
            body { font-family: 'Times New Roman', Times, serif; font-size: 12px; }
            table.data-table { width: 100%; border-collapse: collapse; margin-bottom: 10px;}
            table.data-table th, table.data-table td { border: 1px solid #000; padding: 5px; vertical-align: top; }
            table.data-table th { background-color: #f1f5f9; text-align: center; font-weight: bold; }
          </style>
        </head>
        <body>
          ${kopHtml}
          <table class="data-table">
            <thead>
              <tr>
                <th style="width: 5%;">No</th>
                <th style="width: 10%;">Tanggal</th>
                <th style="width: 8%;">Waktu</th>
                <th style="width: 22%;">Petugas</th>
                <th style="width: 25%;">Aktivitas / Kegiatan</th>
                <th style="width: 30%;">Dokumentasi</th>
              </tr>
            </thead>
            <tbody>${tableRows}</tbody>
          </table>
          ${ttdHtml}
        </body>
      </html>
    `;

    const blob = Utilities.newBlob(htmlContent, 'text/html', 'temp.html');
    const fileName = `Laporan_${dataItem.kategori}_${dataItem.tanggal}.pdf`;
    const pdfBlob = blob.getAs('application/pdf').setName(fileName);
    
    // --- MENGGUNAKAN ID FOLDER SPESIFIK ---
    const baseFolder = DriveApp.getFolderById(FOLDER_PDF_CSDU_ID);
    let targetFolder;

    if (userRoleInfo.role === 'Kepala Sekolah') {
       let folderKepsek = getFolderByNameOrCreate(baseFolder, "Laporan Kepala Sekolah");
       let folderBulan = getFolderByNameOrCreate(folderKepsek, Utilities.formatDate(tglObj, "GMT+7", "MMMM_yyyy"));
       targetFolder = getFolderByNameOrCreate(folderBulan, dataItem.kategori);
    } else {
       // Jika admin/user biasa, masukkan ke subfolder CS_DU_Security di dalam folder ID
       targetFolder = getFolderByNameOrCreate(baseFolder, "CS_DU_Security");
    }

    saveFileWithOverwrite(targetFolder, fileName, Utilities.base64Encode(pdfBlob.getBytes()));

    kumpulanLaporan.forEach(item => {
      sheet.getRange(item.rowIndex, 11).setValue('Generated'); 
    });
    
    return { success: true, message: "PDF Gabungan A4 Landscape berhasil disimpan ke Drive." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}
/**
 * Menghapus baris data
 */
function deleteLaporanCSDUSecurity(rowIndex) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CSDU_NAME);
  sheet.deleteRow(parseInt(rowIndex));
  return { success: true };
}

/**
 * Generate PDF Berdasarkan Kategori dan Rentang Tanggal CS DU SECURITY yo brooooo
 */
function generatePDFCSDUSecurityRentang(filterConfig) {
  try {
    const emailPembuat = Session.getActiveUser().getEmail();
    const userRoleInfo = checkUserStatus(); 
    const userConfig = getPengaturanUser(emailPembuat);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_CSDU_NAME);
    if (!sheet) return { success: false, message: "Sheet laporan tidak ditemukan." };
    
    const data = sheet.getDataRange().getValues();
    const displayData = sheet.getDataRange().getDisplayValues();
    
    let kumpulanLaporan = [];
    let katLabel = filterConfig.kategori === 'Semua' ? 'Gabungan' : filterConfig.kategori;

    for (let i = 1; i < data.length; i++) {
      let tglRaw = data[i][1];
      if (!tglRaw) continue;
      let tglSheet = Utilities.formatDate(new Date(tglRaw), "GMT+7", "yyyy-MM-dd");
      let katSheet = String(data[i][2]).trim();

      let matchKat = (filterConfig.kategori === 'Semua' || katSheet === filterConfig.kategori);
      let matchTgl = (tglSheet >= filterConfig.tglAwal && tglSheet <= filterConfig.tglAkhir);

      if (matchKat && matchTgl) {
        let jamStr = displayData[i][12] || data[i][12] || '-';
        if (jamStr.includes(':')) jamStr = jamStr.substring(0, 5).replace(':', '.');

        let petugasRaw = data[i][4] ? String(data[i][4]).split('\n') : [];
        let petugasHtml = petugasRaw.map((p, idx) => petugasRaw.length > 1 ? `${idx + 1}. ${p.trim()}` : p.trim()).join('<br/>');

        let kegiatanRaw = data[i][3] ? String(data[i][3]) : '';
        let ketRaw = data[i][5] ? String(data[i][5]) : '';
        let gabungKegiatan = [kegiatanRaw, ketRaw].filter(Boolean).join('\n').split('\n');
        let kegiatanHtml = gabungKegiatan.map(k => `- ${k.trim()}`).join('<br/>');

        kumpulanLaporan.push({
          tanggal: tglSheet,
          laporanKe: parseInt(data[i][11]) || 99,
          jam: jamStr,
          petugas: petugasHtml,
          kegiatan: kegiatanHtml,
          foto1: data[i][6] ? getBase64FromDriveUrlSafe(data[i][6]) : '',
          foto2: data[i][7] ? getBase64FromDriveUrlSafe(data[i][7]) : '',
          foto3: data[i][8] ? getBase64FromDriveUrlSafe(data[i][8]) : '',
          rowIndex: i + 1
        });
      }
    }

    if (kumpulanLaporan.length === 0) {
      return { success: false, message: "Tidak ada data laporan yang cocok pada rentang tanggal dan kategori tersebut." };
    }

    // Urutkan berdasarkan Tanggal lalu Laporan Ke
    kumpulanLaporan.sort((a, b) => {
      if (a.tanggal !== b.tanggal) return a.tanggal.localeCompare(b.tanggal);
      return a.laporanKe - b.laporanKe;
    });

    const bulanIndo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    let tglObj = new Date(filterConfig.tglAwal);
    let namaBulan = bulanIndo[tglObj.getMonth()];
    let tahun = tglObj.getFullYear();

    let tableRows = '';
    kumpulanLaporan.forEach((item, index) => {
      let fotosHtml = '';
      [item.foto1, item.foto2, item.foto3].forEach(b64 => {
        if(b64) fotosHtml += `<img src="${b64}" style="width: 120px; height: 90px; object-fit: cover; border: 1px solid #444; margin: 2px;">`;
      });

      let partsTgl = item.tanggal.split('-');
      let formattedDate = `${partsTgl[2]} ${bulanIndo[parseInt(partsTgl[1])-1]} ${partsTgl[0]}`;

      tableRows += `
        <tr>
          <td style="text-align: center;">${index + 1}</td>
          <td style="text-align: center;">${formattedDate}</td>
          <td style="text-align: center;">${item.jam}</td>
          <td>${item.petugas}</td>
          <td>${item.kegiatan}</td>
          <td style="text-align: center; vertical-align: middle;">${fotosHtml}</td>
        </tr>
      `;
    });

    let kopText = userConfig.kop.barisKop || '';
    if (kopText.startsWith('[')) {
        try {
            let parsed = JSON.parse(kopText);
            kopText = parsed.map(p => `<span style="font-size:${p.ukuran || 12}px; ${p.isBold ? 'font-weight:bold;' : ''}">${p.teks}</span>`).join('<br/>');
        } catch(e) {
            kopText = `<b>KEMENTERIAN SOSIAL REPUBLIK INDONESIA</b><br/><b>PUSDIKLATBANGPROFKESOS</b><br/><b>SEKOLAH RAKYAT MENENGAH ATAS 16 TEMANGGUNG</b><br/>Jalan Kartini No.1-2 Bendo Kertosari, Temanggung`;
        }
    } else if (!kopText) {
        kopText = `<b>KEMENTERIAN SOSIAL REPUBLIK INDONESIA</b><br/><b>PUSDIKLATBANGPROFKESOS</b><br/><b>SEKOLAH RAKYAT MENENGAH ATAS 16 TEMANGGUNG</b><br/>Jalan Kartini No.1-2 Bendo Kertosari, Temanggung`;
    }

    let kopHtml = `
      <table style="width:100%; border:none; margin-bottom:10px;">
        <tr>
          <td style="width:15%; border:none; text-align:center; vertical-align:middle;">
            ${userConfig.kop.logoKiri ? `<img src="${userConfig.kop.logoKiri}" style="max-height:85px;">` : ''}
          </td>
          <td style="width:70%; border:none; text-align:center; line-height:1.2; font-family:Arial, sans-serif;">
             ${kopText}
          </td>
          <td style="width:15%; border:none; text-align:center; vertical-align:middle;">
            ${userConfig.kop.logoKanan ? `<img src="${userConfig.kop.logoKanan}" style="max-height:85px;">` : ''}
          </td>
        </tr>
      </table>
      <div style="border-bottom:3px solid #000; margin-bottom:15px; margin-top:5px;"></div>
      <div style="text-align:center; margin-bottom: 15px;">
        <h3 style="margin:0; font-size:14px; font-weight:bold;">
           Laporan Rekap ${katLabel} Periode ${filterConfig.tglAwal} s.d. ${filterConfig.tglAkhir}
        </h3>
      </div>`;

    let ttdHtml = '';
    if(userConfig.ttd && userConfig.ttd.namaTtd) {
      let imgTtd = userConfig.ttd.tandaTangan ? `<img src="${userConfig.ttd.tandaTangan}" style="max-height:75px; margin: 5px 0;">` : '<br/><br/><br/>';
      ttdHtml = `
      <div style="width: 280px; float: right; text-align: center; margin-top: 25px; font-size: 12px; font-family:Arial, sans-serif;">
        Kepala SRMA 16 Temanggung<br/>
        ${imgTtd}<br/>
        <b><u>${userConfig.ttd.namaTtd}</u></b><br/>
        NIP. ${userConfig.ttd.nipTtd || '-'}
      </div>
      <div style="clear:both;"></div>
      `;
    }

    const htmlContent = `
      <html>
        <head>
          <style>
            @page { size: A4 landscape; margin: 1cm; }
            body { font-family: 'Times New Roman', Times, serif; font-size: 12px; }
            table.data-table { width: 100%; border-collapse: collapse; margin-bottom: 10px;}
            table.data-table th, table.data-table td { border: 1px solid #000; padding: 5px; vertical-align: top; }
            table.data-table th { background-color: #f1f5f9; text-align: center; font-weight: bold; }
          </style>
        </head>
        <body>
          ${kopHtml}
          <table class="data-table">
            <thead>
              <tr>
                <th style="width: 5%;">No</th>
                <th style="width: 10%;">Tanggal</th>
                <th style="width: 8%;">Waktu</th>
                <th style="width: 22%;">Petugas</th>
                <th style="width: 25%;">Aktivitas / Kegiatan</th>
                <th style="width: 30%;">Dokumentasi</th>
              </tr>
            </thead>
            <tbody>${tableRows}</tbody>
          </table>
          ${ttdHtml}
        </body>
      </html>
    `;

    const blob = Utilities.newBlob(htmlContent, 'text/html', 'temp.html');
    const fileName = `Rekap_Laporan_${katLabel}_${filterConfig.tglAwal}_sd_${filterConfig.tglAkhir}.pdf`;
    const pdfBlob = blob.getAs('application/pdf').setName(fileName);
    
    // --- MENGGUNAKAN ID FOLDER SPESIFIK ---
    const baseFolder = DriveApp.getFolderById(FOLDER_PDF_CSDU_ID);
    let targetFolder;

    if (userRoleInfo.role === 'Kepala Sekolah') {
       let folderKepsek = getFolderByNameOrCreate(baseFolder, "Laporan Kepala Sekolah");
       let folderBulan = getFolderByNameOrCreate(folderKepsek, Utilities.formatDate(tglObj, "GMT+7", "MMMM_yyyy"));
       targetFolder = getFolderByNameOrCreate(folderBulan, katLabel);
    } else {
       // Jika admin/user biasa, masukkan ke subfolder CS_DU_Security di dalam folder ID
       targetFolder = getFolderByNameOrCreate(baseFolder, "CS_DU_Security");
    }

    saveFileWithOverwrite(targetFolder, fileName, Utilities.base64Encode(pdfBlob.getBytes()));

    // Ubah status baris yang ikut ter-generate menjadi 'Generated'
    kumpulanLaporan.forEach(item => {
      sheet.getRange(item.rowIndex, 11).setValue('Generated'); 
    });
    
    return { success: true, message: "PDF Rekap Rentang Tanggal berhasil di-generate dan tersimpan di Google Drive." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

//LAPORAN KEPSEK
const SHEET_NAME = 'LapKepsek';
const MASTER_SHEET_NAME = 'Master_Aktivitas'; // Tambahkan nama sheet master
const FOTO_FOLDER_ID = '1bEjthe5qif_2q3L7gzBmVz2MUwlvsbx_'; // Pastikan Folder ID benar
const PDF_FOLDER_ID = '1JZW7_ShwiE1Zr4do7kTu8aHqLJ5ggMDo';

// 1. Mengambil semua data saat aplikasi pertama kali dimuat
function getInitialDataLapKesek() {
  try {
    const props = PropertiesService.getScriptProperties();
    let settings = {};
    try { 
      const setStr = props.getProperty('lapkesek_settings');
      if (setStr) settings = JSON.parse(setStr);
    } catch(e) {}

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    
    // --- BUAT SHEET MASTER OTOMATIS JIKA BELUM ADA ---
    let sheetMaster = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!sheetMaster) {
      sheetMaster = ss.insertSheet(MASTER_SHEET_NAME);
      sheetMaster.appendRow(["Jenis Aktivitas"]);
      sheetMaster.appendRow(["Melaksanakan home visit"]);
      sheetMaster.appendRow(["Mengikuti zoom"]);
      sheetMaster.appendRow(["Pemantauan secara langsung"]);
    }
    
    // Ambil daftar aktivitas dari Master Sheet
    const dataMaster = sheetMaster.getDataRange().getValues();
    let masterAktivitas = [];
    for (let i = 1; i < dataMaster.length; i++) {
      if (dataMaster[i][0]) masterAktivitas.push(dataMaster[i][0]);
    }

    // --- BUAT SHEET LAPORAN OTOMATIS JIKA BELUM ADA ---
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["Laporan ID", "Tanggal Input", "Tanggal Laporan", "Data JSON", "Status", "PDF URL"]);
    }

    const data = sheet.getDataRange().getValues();
    const reports = [];
    
    if (data && data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[0] || row[0] === "") continue; 

        let jsonKegiatan = []; 
        try {
          if (row[3] && row[3] !== "") {
            jsonKegiatan = JSON.parse(row[3]);
          }
        } catch(e) {}

        // KOREKSI UTAMA: Ubah format Date menjadi String agar tidak membuat google.script.run error
        let rawTglInput = row[1];
        let rawTglLaporan = row[2];
        
        let stringTglInput = (rawTglInput instanceof Date) ? rawTglInput.toISOString() : String(rawTglInput);
        let stringTglLaporan = (rawTglLaporan instanceof Date) ? rawTglLaporan.toISOString() : String(rawTglLaporan);

        reports.push({
          id: row[0],
          tanggalInput: stringTglInput, // Menggunakan string
          tanggalLaporan: stringTglLaporan, // Menggunakan string
          kegiatan: jsonKegiatan, 
          status: row[4],
          pdfUrl: row[5] || ''
        });
      }
    }
    
    return { 
      settings: settings, 
      masterAktivitas: masterAktivitas, 
      reports: reports 
    };

  } catch(error) {
    throw new Error("Backend Data Error: " + error.message);
  }
}

// 2. FUNGSI BARU: Menyimpan pengaturan dan update Sheet Master_Aktivitas
function saveSettingsLapKesek(settings, masterArr) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('lapkesek_settings', JSON.stringify(settings));

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheetMaster = ss.getSheetByName(MASTER_SHEET_NAME);
  
  // Jika sheet belum ada, buat otomatis
  if (!sheetMaster) {
    sheetMaster = ss.insertSheet(MASTER_SHEET_NAME);
  }
  
  // Bersihkan data lama dan masukkan data baru
  sheetMaster.clear();
  sheetMaster.appendRow(["Jenis Aktivitas"]);
  
  if (masterArr && masterArr.length > 0) {
    const dataToWrite = masterArr.map(item => [item]);
    sheetMaster.getRange(2, 1, dataToWrite.length, 1).setValues(dataToWrite);
  }
  
  return true;
}

// 3. FUNGSI BARU: Mengunggah foto ke Google Drive
function uploadFotoToDrive(base64Data, fileName) {
  try {
    const folder = DriveApp.getFolderById(FOTO_FOLDER_ID);
    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = folder.createFile(blob);
    
    // PELINDUNG WAJIB: Cegah error simpan jika admin sekolah memblokir sharing publik
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(sharingErr) {
      // Abaikan error sharing, biarkan proses berlanjut
    }
    
    return `https://drive.google.com/file/d/${file.getId()}/view?usp=drivesdk`;
  } catch (e) {
    throw new Error("Upload Error: " + e.message);
  }
}

// 4. FUNGSI BARU: Menyimpan data Laporan ke Sheet
function saveReportLapKesek(reportData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Laporan ID", "Tanggal Input", "Tanggal Laporan", "Data JSON", "Status", "PDF URL"]);
  }
  
  sheet.appendRow([
    reportData.id,
    reportData.tanggalInput,
    reportData.tanggalLaporan,
    JSON.stringify(reportData.kegiatan),
    reportData.status,
    reportData.pdfUrl
  ]);
  
  return true;
}

// 5. FUNGSI BARU: Menghapus data laporan dari Sheet
function deleteReportLapKesek(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1); // +1 karena index row dimulai dari 1
      break;
    }
  }
  return true;
}
// 6. FUNGSI BARU: Update Data Laporan di Sheet
function updateReportLapKesek(reportData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reportData.id) {
      sheet.getRange(i + 1, 2).setValue(reportData.tanggalInput);
      sheet.getRange(i + 1, 3).setValue(reportData.tanggalLaporan);
      sheet.getRange(i + 1, 4).setValue(JSON.stringify(reportData.kegiatan));
      sheet.getRange(i + 1, 5).setValue(reportData.status);
      return true;
    }
  }
  return false;
}

// 7. FUNGSI PERBAIKAN FINAL: Generate PDF (Autoheight Fix) & Auto Replace Drive
function generatePDFLapKesek(id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    let reportRowIndex = -1;
    let reportData = null;
    let emailPembuat = Session.getActiveUser().getEmail(); 
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        reportRowIndex = i + 1;
        let kegObj = {};
        try { kegObj = JSON.parse(data[i][3]); } catch(e) {}
        reportData = { id: data[i][0], tanggalInput: data[i][1], tglLaporan: data[i][2], kegiatan: kegObj };
        break;
      }
    }
    if (!reportData) throw new Error("Data tidak ditemukan.");
    
    // --- 1. AMBIL KOP DARI PENGATURAN USER ---
    let pengaturanUser = {};
    try { pengaturanUser = getPengaturanUser(emailPembuat); } catch(e) {}
    const kop = pengaturanUser.kop || {};

    // --- 2. AMBIL ISIAN DARI PENGATURAN LAPKEPSEK ---
    const props = PropertiesService.getScriptProperties();
    let lapkesekSettings = {};
    try { lapkesekSettings = JSON.parse(props.getProperty('lapkesek_settings')) || {}; } catch(e){}
    
    const namaKepsek = lapkesekSettings.kepsek || '-';
    const koordinator = lapkesekSettings.koordinator || '-';
    const pic = lapkesekSettings.pic || '-';
    const lokasi = lapkesekSettings.lokasi || '-';
    const ppk = lapkesekSettings.ppk || '-';
    const nipKoordinator = lapkesekSettings.nipKoordinator || '';
    const nipPic = lapkesekSettings.nipPic || '';
    const nipKepsek = lapkesekSettings.nipKepsek || '';
    
    const tglFormatted = reportData.tglLaporan ? new Date(reportData.tglLaporan).toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }) : '-';
    const tglInputFormatted = reportData.tanggalInput ? new Date(reportData.tanggalInput).toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }) : '-';

    // --- PERBAIKAN AUTOHEIGHT: UBAH ENTER (\n) MENJADI <br> SECARA NATIVE ---
    let daftarUraianHtml = '';
    if (reportData.kegiatan.rincian && reportData.kegiatan.rincian.length > 0) {
        reportData.kegiatan.rincian.forEach((rin, idx) => {
            let textUraian = (rin.nama || '').toString().replace(/\n/g, '<br>');
            daftarUraianHtml += `${idx + 1}. ${textUraian}<br>`;
        });
    } else {
        daftarUraianHtml = (reportData.kegiatan.uraianUtama || '-').toString().replace(/\n/g, '<br>');
    }
    
    // Ganti enter jadi <br> untuk Kendala & Tindak Lanjut
    let kendalaHtml = (reportData.kegiatan.kendalaUtama || '-').toString().replace(/\n/g, '<br>');
    let tindakLanjutHtml = (reportData.kegiatan.tindakLanjutUtama || '-').toString().replace(/\n/g, '<br>');

    function getDriveThumb(url) {
        if (!url) return '';
        if (url.startsWith('data:image')) return url;
        let fileId = '';
        if (url.includes('/file/d/')) {
            let match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
            if (match) fileId = match[1];
        } else if (url.includes('id=')) {
            let match = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
            if (match) fileId = match[1];
        }
        return fileId ? `https://drive.google.com/thumbnail?id=${fileId}&sz=w800` : url;
    }

    // Susun Baris KOP
    let barisKopHtml = '';
    if (kop.barisKop) {
        try {
            let arrKop = typeof kop.barisKop === 'string' ? JSON.parse(kop.barisKop) : kop.barisKop;
            arrKop.forEach(b => {
                let isBold = b.B || b.isBold || b.b ? 'font-weight:bold;' : ''; 
                barisKopHtml += `<div style="${isBold} font-size:${b.ukuran || 12}pt; margin-bottom: 2px;">${b.teks}</div>`;
            });
        } catch(e) {}
    }

    let html = `
      <!DOCTYPE html>
      <html>
      <head>
      <style>
         @page { size: A4; margin: 15mm; }
         body { font-family: 'Times New Roman', Times, serif; font-size: 11pt; color: #000; }
         
         /* KUNCI AUTOHEIGHT TABEL: table-layout: fixed dan membatasi kolom */
         .main-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; table-layout: fixed; }
         .main-table td { 
             border: 1px solid black; 
             padding: 8px; 
             vertical-align: top; 
             word-wrap: break-word; /* Memaksa teks panjang wrap ke bawah */
         }
         
         .bg-gray { background-color: #f2f2f2; font-weight: bold; width: 35%; }
         .content-td { width: 65%; line-height: 1.4; text-align: justify; }
         
         .photo-box { page-break-inside: avoid; text-align: center; margin-bottom: 15px; }
      </style>
      </head>
      <body>
      
      <div>
        <!-- KOP SURAT DINAMIS -->
        <table style="width: 100%; border-collapse: collapse; border: none; margin-bottom: 10px;">
          <tr>
            <td style="width: 15%; text-align: center; border: none; vertical-align: middle;">
              ${kop.logoKiri ? `<img src="${getDriveThumb(kop.logoKiri)}" style="max-height: 80px; max-width: 100%; object-fit: contain;">` : ''}
            </td>
            <td style="width: 70%; text-align: center; border: none; vertical-align: middle;">
              ${barisKopHtml}
            </td>
            <td style="width: 15%; text-align: center; border: none; vertical-align: middle;">
              ${kop.logoKanan ? `<img src="${getDriveThumb(kop.logoKanan)}" style="max-height: 80px; max-width: 100%; object-fit: contain;">` : ''}
            </td>
          </tr>
        </table>
        
        <hr style="border: none; border-top: 2px solid black; margin-bottom: 2px;">
        <hr style="border: none; border-top: 1px solid black; margin-bottom: 15px;">

        <!-- TABEL UTAMA FORM LAPORAN -->
        <table class="main-table">
          <tr>
            <td class="bg-gray">TANGGAL</td>
            <td class="content-td">${tglInputFormatted}</td>
          </tr>
          <tr>
            <td class="bg-gray">NAMA KEPALA SEKOLAH</td>
            <td class="content-td">${namaKepsek}</td>
          </tr>
          <tr>
            <td class="bg-gray">NAMA KOORDINATOR</td>
            <td class="content-td">${koordinator}</td>
          </tr>
          <tr>
            <td class="bg-gray">NAMA PIC</td>
            <td class="content-td">${pic}</td>
          </tr>
          <tr>
            <td class="bg-gray">LOKASI SR</td>
            <td class="content-td">${lokasi}</td>
          </tr>
          <tr>
            <td class="bg-gray">NAMA PPK</td>
            <td class="content-td">${ppk}</td>
          </tr>
          <tr>
            <td class="bg-gray">TANGGAL LAPORAN</td>
            <td class="content-td">${tglFormatted}</td>
          </tr>
          <tr>
            <td class="bg-gray">URAIAN KEGIATAN/AKTIVITAS HARIAN</td>
            <td class="content-td">${daftarUraianHtml}</td>
          </tr>
          <tr>
            <td class="bg-gray">PERMASALAHAN DAN KENDALA</td>
            <td class="content-td">${kendalaHtml}</td>
          </tr>
          <tr>
            <td class="bg-gray">TINDAK LANJUT YANG DILAKUKAN (JIKA ADA)</td>
            <td class="content-td">${tindakLanjutHtml}</td>
          </tr>
          <tr>
            <td class="bg-gray">I. DOKUMENTASI KEGIATAN</td>
            <td class="content-td" style="padding: 10px;">
    `;

    // Render Dokumentasi Foto
    if (reportData.kegiatan.rincian && reportData.kegiatan.rincian.length > 0) {
        reportData.kegiatan.rincian.forEach((rin, rIdx) => {
            if (rin.foto && rin.foto.length > 0) {
                rin.foto.forEach((fUrl, fIdx) => {
                    let thumb = getDriveThumb(fUrl);
                    let isLastPhotoOfItem = (fIdx === rin.foto.length - 1);
                    
                    if (thumb) {
                        html += `
                          <div class="photo-box">
                            <img src="${thumb}" style="max-width: 100%; max-height: 250px; object-fit: contain; border: 1px solid #ccc; padding: 3px; background: #fff;">
                        `;
                        if (isLastPhotoOfItem) {
                            html += `<span style="font-weight: bold; font-size: 10pt; display: block; margin-top: 5px; margin-bottom: 8px;">${rin.nama}</span>`;
                        }
                        html += `</div>`;
                    }
                });

                if (rIdx < reportData.kegiatan.rincian.length - 1) {
                    html += `<hr style="border: none; border-top: 1px dashed #aaa; margin: 12px 0;">`;
                }
            }
        });
    } else {
        html += `<i>Tidak ada foto dokumentasi.</i>`;
    }

    html += `
            </td>
          </tr>
        </table>

        <!-- TANDA TANGAN -->
        <table style="width: 100%; text-align: center; border: none; margin-top: 20px; page-break-inside: avoid;">
          <tr>
            <td style="width: 50%; border: none; vertical-align: top; padding-bottom: 40px;">
              KOORDINATOR,<br><br><br><br>
              <strong>${koordinator}</strong>
              ${nipKoordinator ? '<br>NIP. ' + nipKoordinator : ''}
            </td>
            <td style="width: 50%; border: none; vertical-align: top; padding-bottom: 40px;">
              PIC,<br><br><br><br>
              <strong>${pic}</strong>
              ${nipPic ? '<br>NIP. ' + nipPic : ''}
            </td>
          </tr>
          <tr>
            <td colspan="2" style="border: none; vertical-align: top; padding-top: 10px;">
              Mengetahui,<br>
              <strong>KEPALA SEKOLAH</strong><br><br><br><br>
              <strong>${namaKepsek}</strong>
              ${nipKepsek ? '<br>NIP. ' + nipKepsek : ''}
            </td>
          </tr>
        </table>
        
      </div>
      </body>
      </html>
    `;

    const folder = DriveApp.getFolderById(PDF_FOLDER_ID);
    const fileName = `Laporan_${id}.pdf`;

    // --- FITUR BARU: AUTO REPLACE JIKA FILE LAMA SUDAH ADA ---
    const existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
        let fileLama = existingFiles.next();
        fileLama.setTrashed(true); // Pindahkan file PDF lama ke folder sampah
    }

    // Buat file PDF baru
    const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF).setName(fileName);
    const file = folder.createFile(blob);
    
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    
    const pdfUrl = file.getUrl();
    sheet.getRange(reportRowIndex, 6).setValue(pdfUrl);
    sheet.getRange(reportRowIndex, 5).setValue('PDF Generated');
    
    return { url: pdfUrl };
  } catch (e) {
    throw new Error("Gagal sistem backend: " + e.message);
  }
}

//--END LAPORAN KEPSEK

// ==========================================
// MODUL MANAJEMEN DATA SISWA (TAMBAH, EDIT, KELUAR)
// ==========================================

// Helper Upload ke Folder Spesifik
function uploadPdfToDrive(base64Data, fileName, folderId) {
  try {
    var splitBase = base64Data.split(',');
    var type = splitBase[0].split(';')[0].replace('data:', '');
    var byteCharacters = Utilities.base64Decode(splitBase[1]);
    var blob = Utilities.newBlob(byteCharacters, type, fileName);
    
    var folder = DriveApp.getFolderById(folderId);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); 
    return file.getUrl();
  } catch (e) {
    return "";
  }
}

function getDataSiswaAll() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('DataSiswa');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  let result = [];
  
  for (let i = 2; i < data.length; i++) {
    let tglLahirRaw = data[i][15]; // Kolom P
    let tglLahirFormatted = (tglLahirRaw instanceof Date) ? Utilities.formatDate(tglLahirRaw, Session.getScriptTimeZone(), "yyyy-MM-dd") : tglLahirRaw;

    if (data[i][0] && data[i][2] && data[i][2] !== "SEKOLAH RAKYAT") {
      result.push({
        rowIndex: i + 1, 
        nisn: data[i][0],        // A (0)
        nik: data[i][1],         // B (1)
        nama: data[i][2],        // C (2)
        alamat: data[i][3],      // D (3)
        hpSiswa: data[i][4],     // E (4)
        hpOrtu: data[i][5],      // F (5)
        lat: data[i][6],         // G (6)
        lng: data[i][7],         // H (7)
        bapak: data[i][8],       // I (8)
        ibu: data[i][9],         // J (9)
        catatan: data[i][10],    // K (10)
        kelas: data[i][11],      // L (11)
        wali: data[i][12],       // M (12)
        kamar: data[i][13],      // N (13)
        tmptLahir: data[i][14],  // O (14)
        tglLahir: tglLahirFormatted, // P (15)
        jk: data[i][16],         // Q (16)
        fileKK: data[i][17],     // R (17)
        fileAkta: data[i][18]    // S (18)
      });
    }
  }
  return result;
}

function simpanDataSiswa(payload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('DataSiswa');
  
  // 1. Paksa konversi tipe data agar tidak meleset
  let isEditMode = (payload.isEdit === true || String(payload.isEdit) === "true");
  let rowIndex = parseInt(payload.rowIndex);
  
  // 2. SISTEM PENGAMAN: Jika disuruh Edit tapi nomor baris (rowIndex) hilang/error, blokir prosesnya!
  if (isEditMode && (!rowIndex || isNaN(rowIndex) || rowIndex < 2)) {
     return { success: false, message: "GAGAL EDIT: Nomor baris di Sheet tidak terdeteksi. Silakan muat ulang (refresh) halaman dan coba lagi." };
  }
  
  // 3. Proses Upload File
  let urlKK = payload.fileKK_lama || "";
  let urlAkta = payload.fileAkta_lama || "";
  
  if (payload.fileKK_base64) {
    urlKK = uploadPdfToDrive(payload.fileKK_base64, payload.nisn + "_KK.pdf", FOLDER_KK_ID);
  }
  if (payload.fileAkta_base64) {
    urlAkta = uploadPdfToDrive(payload.fileAkta_base64, payload.nisn + "_AKTA.pdf", FOLDER_AKTA_ID);
  }
  
  // 4. Siapkan Data (19 Kolom)
  let rowData = [
    payload.nisn, payload.nik, payload.nama, payload.alamat, payload.hpSiswa, 
    payload.hpOrtu, payload.lat, payload.lng, payload.bapak, payload.ibu, 
    payload.catatan, payload.kelas, payload.wali, payload.kamar, 
    payload.tmptLahir, payload.tglLahir, payload.jk, urlKK, urlAkta
  ];

  // 5. Eksekusi Simpan
  if (isEditMode) {
    sheet.getRange(rowIndex, 1, 1, 19).setValues([rowData]);
    return { success: true, message: "Data siswa berhasil diperbarui!" };
  } else {
    sheet.appendRow(rowData);
    return { success: true, message: "Data siswa baru berhasil ditambahkan!" };
  }
}

function keluarkanSiswa(nisn) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetIn = ss.getSheetByName('DataSiswa');
  let sheetOut = ss.getSheetByName('DataSiswaOut');
  
  if (!sheetOut) {
    sheetOut = ss.insertSheet('DataSiswaOut');
    const headers = sheetIn.getRange(1, 1, 1, sheetIn.getLastColumn()).getValues();
    sheetOut.getRange(1, 1, 1, headers[0].length).setValues(headers); 
  }
  
  const data = sheetIn.getDataRange().getValues();
  
  for (let i = 2; i < data.length; i++) { 
    if (String(data[i][0]).trim() === String(nisn).trim()) {
      sheetOut.appendRow(data[i]);
      sheetIn.deleteRow(i + 1);
      return { success: true, message: "Siswa berhasil dikeluarkan dan datanya dipindah ke sheet DataSiswaOut." };
    }
  }
  return { success: false, message: "Data siswa tidak ditemukan!" };
}

// Fungsi untuk mengekstrak ID Folder dari URL Google Drive
function dapatkanIdDariUrl(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

// Fungsi utama untuk memindahkan/copy folder bulanan ke folder drive tiap bulan yoo...
function prosesCopyBackupKeDriveBaru(urlDriveTujuan, bulanTahunTarget) {
  try {
    // 1. Dapatkan ID Tujuan
    const idTujuan = dapatkanIdDariUrl(urlDriveTujuan);
    if (!idTujuan) throw new Error("Format Link Google Drive tidak valid.");

    // 2. Setup Folder Tujuan (Langsung ke folder root dari link yang dimasukkan)
    const folderTujuanRoot = DriveApp.getFolderById(idTujuan);

    // 3. Setup Sumber Backup Lama
    const idBackupLama = '1quTf_wLAPnzta9EY_7dBsRE9RvmmHfAc';
    const rootBackupLama = DriveApp.getFolderById(idBackupLama);
    const folderUtamaLama = rootBackupLama.getFoldersByName("Laporan Harian");

    if (!folderUtamaLama.hasNext()) {
      throw new Error("Folder sumber 'Laporan Harian' tidak ditemukan di backup.");
    }
    const folderHarian = folderUtamaLama.next();

    // 4. Looping Data Per User
    const folderUsers = folderHarian.getFolders();
    let totalFileDicopy = 0;
    let totalFolderUser = 0;

    while (folderUsers.hasNext()) {
      const userFolder = folderUsers.next();
      const namaUser = userFolder.getName();
      const bulanFolders = userFolder.getFoldersByName(bulanTahunTarget);

      // Jika user memiliki folder laporan pada bulan tersebut
      if (bulanFolders.hasNext()) {
        const sourceBulanFolder = bulanFolders.next();
        
        // Mengecek apakah folder nama user sudah ada di folder tujuan
        let newUserFolder;
        const checkFolder = folderTujuanRoot.getFoldersByName(namaUser);
        if (checkFolder.hasNext()) {
            newUserFolder = checkFolder.next();
        } else {
            // Buat folder atas nama user baru
            newUserFolder = folderTujuanRoot.createFolder(namaUser);
        }
        
        totalFolderUser++;

        // Copy seluruh file PDF dengan logika REPLACE
        const files = sourceBulanFolder.getFiles();
        while (files.hasNext()) {
          const sourceFile = files.next();
          const namaFile = sourceFile.getName();

          // Cek apakah file dengan nama yang sama persis sudah ada
          const existingFiles = newUserFolder.getFilesByName(namaFile);
          
          // Jika ada, hapus file lama tersebut (masukkan ke tong sampah Drive)
          while (existingFiles.hasNext()) {
             existingFiles.next().setTrashed(true);
          }

          // Setelah file lama bersih, salin (copy) file baru ke folder tersebut
          sourceFile.makeCopy(namaFile, newUserFolder);
          totalFileDicopy++;
        }
      }
    }

    return { 
      status: true, 
      pesan: `Berhasil! ${totalFileDicopy} file (replace) dipindahkan ke dalam ${totalFolderUser} folder pengguna.` 
    };

  } catch (error) {
    return { 
      status: false, 
      pesan: `Terjadi kesalahan: ${error.message}` 
    };
  }
}
