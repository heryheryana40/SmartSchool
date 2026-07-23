// Code.gs

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Portal Akademik & CBT Sekolah')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// FUNGSI LOGIN BARU: Berdasarkan Username & Password Manual
function prosesLoginManual(usernameInput, passwordInput) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUser = ss.getSheetByName("Users");
  var dataUser = sheetUser.getDataRange().getValues();
  
  for(var i = 1; i < dataUser.length; i++) {
    // Bersihkan sisa tanda kutip (') di depan data lama (jika ada, akibat bug import massal sebelumnya)
    var usernameSheet = dataUser[i][0].toString().trim().replace(/^'/, "");
    var passwordSheet = dataUser[i][1].toString().trim().replace(/^'/, "");
    // Mencocokkan username (kolom A) dan password (kolom B)
    if(usernameSheet === usernameInput.trim() && passwordSheet === passwordInput.trim()) {
      return { 
        status: "SUCCESS", 
        role: dataUser[i][3], // Kolom D (Role)
        data: {
          username: dataUser[i][0],
          nama: dataUser[i][2], // Kolom C (Nama)
          role: dataUser[i][3],
          kelas: dataUser[i].length > 4 ? dataUser[i][4] : "" // Kolom E (Kelas) - dibutuhkan utk filter Materi & Tugas siswa
        }
      };
    }
  }
  return { status: "ERROR", message: "Username atau Password salah!" };
}

// FUNGSI PEMBERSIH DATA (jalankan SEKALI SAJA dari editor Apps Script, lalu boleh dihapus)
// Membersihkan sisa karakter ' di depan Username/Password akibat bug import massal sebelumnya
function bersihkanTandaKutipUserLama() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  var jumlahDiperbaiki = 0;

  for (var i = 1; i < data.length; i++) {
    var usernameAsli = data[i][0].toString();
    var passwordAsli = data[i][1].toString();
    var usernameBersih = usernameAsli.replace(/^'/, "").trim();
    var passwordBersih = passwordAsli.replace(/^'/, "").trim();

    if (usernameAsli !== usernameBersih || passwordAsli !== passwordBersih) {
      sheet.getRange(i + 1, 1, 1, 2).setNumberFormat("@").setValues([[usernameBersih, passwordBersih]]);
      jumlahDiperbaiki++;
    }
  }
  Logger.log("Selesai. Total baris diperbaiki: " + jumlahDiperbaiki);
  return "Selesai. Total baris diperbaiki: " + jumlahDiperbaiki;
}

// ==================== KHUSUS FITUR ADMIN ====================

// Ambil semua user untuk ditampilkan di dashboard Admin
function adminAmbilSemuaUser() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  var listUser = [];
  
  for(var i = 1; i < data.length; i++) {
    listUser.push({
      username: data[i][0],
      nama: data[i][2],
      role: data[i][3],
      kelas: data[i].length > 4 ? data[i][4] : "" // Kolom E (Kelas) - sebelumnya tidak ikut terkirim
    });
  }
  return listUser;
}

// Tambah User Baru oleh Admin
// Fungsi Admin Tambah User Baru lengkap dengan Kolom Kelas (Kolom E)
function adminTambahUserBaruDenganKelas(username, password, nama, role, kelas) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  
  // Cek apakah username sudah terdaftar sebelumnya
  for(var i = 1; i < data.length; i++) {
    if(data[i][0].toString().trim() === username.trim()) {
      return { status: "FAILED", message: "Username/NISN sudah terdaftar di sistem!" };
    }
  }
  
  // Masukkan data ke baris baru: [Username, Password, Nama, Role, Kelas]
  sheet.appendRow([
    username.trim(),
    password.trim(),
    nama.trim(),
    role,
    role === "Siswa" ? kelas.trim().toUpperCase() : "" // Hanya simpan kelas jika dia Siswa
  ]);
  
  catatLogAktivitas("Admin", "Tambah Akun", "Membuat akun baru " + username + " (" + nama + ") sebagai " + role);
  
  return { status: "SUCCESS", message: "Akun " + role + " atas nama " + nama + " berhasil dibuat!" };
}

// Hapus User oleh Admin
function adminHapusUser(username) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  
  for(var i = data.length - 1; i >= 1; i--) {
    if(data[i][0].toString().trim() === username.trim()) {
      var namaUser = data[i][2];
      sheet.deleteRow(i + 1);
      catatLogAktivitas("Admin", "Hapus Akun", "Menghapus akun " + username + " (" + namaUser + ")");
      return { status: "SUCCESS", message: "User berhasil dihapus." };
    }
  }
  return { status: "ERROR", message: "User Tidak Ditemukan." };
}

// ==================== KHUSUS FITUR GURU ====================
// 1. Fungsi Ambil Daftar Siswa Berdasarkan Kelas
function ambilDaftarSiswaBerdasarkanKelas(kelasDipilih) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  var siswaList = [];
  
  // Ambil data dan buat list kelas unik untuk dikembalikan ke filter dropdown guru
  var semuaKelas = [];
  
  for(var i = 1; i < data.length; i++) {
    var role = data[i][3].toString().trim();
    
    if(role === "Siswa") {
      // Kita asumsikan format username atau nama mengandung kelas, 
      // ATAU jika Anda punya kolom Kelas tersendiri (misal Kolom E / indeks ke-4).
      // Untuk amannya, jika belum ada kolom kelas, kita buat fallback default "Semua Kelas" atau deteksi kolom ke-5 jika ada.
      var kelasSiswa = (data[i].length > 4 && data[i][4]) ? data[i][4].toString().trim() : "Reguler";
      
      if (!semuaKelas.includes(kelasSiswa)) {
        semuaKelas.push(kelasSiswa);
      }
      
      // Jika guru memilih kelas tertentu, filter di sini
      if(kelasDipilih === "SEMUA" || kelasSiswa === kelasDipilih) {
        siswaList.push({
          username: data[i][0].toString(), // NISN
          nama: data[i][2].toString(),      // Nama
          kelas: kelasSiswa
        });
      }
    }
  }
  
  return { siswa: siswaList, daftarKelas: semuaKelas };
}

// 2. Fungsi Simpan Presensi dengan Tanggal Kustom (Bisa Tanggal Lampau)
// DIPERBARUI: sekarang bersifat UPSERT (update baris yang sudah ada utk siswa+tanggal yg sama, bukan selalu menambah baris baru)
// agar saat guru mengoreksi presensi yang sudah tersimpan, datanya diperbarui, bukan menumpuk jadi duplikat.
function simpanPresensiSiswaKustom(dataPresensi, infoGuru, tanggalPilihan) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Presensi");
  var HEADER_BENAR = ["Tanggal Presensi", "Waktu Input", "NISN/Username", "Nama Siswa", "Kelas", "Status Presensi", "Guru Pengajar"];
  
  if (!sheet) {
    sheet = ss.insertSheet("Presensi");
    sheet.appendRow(HEADER_BENAR);
  } else {
    // PENTING: sheet "Presensi" mungkin sudah ada duluan (dari versi lama/manual) dengan header berbeda.
    // Paksa selaraskan baris header dengan skema yang benar-benar dipakai kode ini, agar tidak membingungkan saat dibuka manual.
    var headerSaatIni = sheet.getRange(1, 1, 1, HEADER_BENAR.length).getValues()[0];
    var headerCocok = HEADER_BENAR.every(function(h, idx) { return headerSaatIni[idx] === h; });
    if (!headerCocok) {
      sheet.getRange(1, 1, 1, HEADER_BENAR.length).setValues([HEADER_BENAR]);
    }
  }
  
  var tglSekarang = new Date();
  var waktuFormated = Utilities.formatDate(tglSekarang, Session.getScriptTimeZone(), "HH:mm:ss");
  var dataSheet = sheet.getDataRange().getValues();
  
  // Buat peta baris eksisting: kunci "tanggal|username" -> nomor baris (agar pencarian cepat, bukan looping berulang)
  var petaBarisEksisting = {};
  for (var i = 1; i < dataSheet.length; i++) {
    var tglBaris = dataSheet[i][0];
    var tglString = (tglBaris instanceof Date) ? Utilities.formatDate(tglBaris, Session.getScriptTimeZone(), "yyyy-MM-dd") : tglBaris.toString().trim();
    var kunci = tglString + "|" + dataSheet[i][2].toString().trim();
    petaBarisEksisting[kunci] = i + 1; // Nomor baris di sheet (1-indexed)
  }
  
  dataPresensi.forEach(function(item) {
    var kunciCari = tanggalPilihan + "|" + item.username.toString().trim();
    
    if (petaBarisEksisting[kunciCari]) {
      // SUDAH ADA -> update baris tsb (koreksi presensi)
      var nomorBaris = petaBarisEksisting[kunciCari];
      // PENTING: kunci kolom Username (C) sebagai TEKS agar NISN berawalan 0 tidak hilang
      sheet.getRange(nomorBaris, 3).setNumberFormat("@");
      sheet.getRange(nomorBaris, 2, 1, 6).setValues([[
        waktuFormated, item.username, item.nama, item.kelas, item.status, infoGuru
      ]]);
    } else {
      // BELUM ADA -> tambah baris baru
      sheet.appendRow([
        tanggalPilihan, waktuFormated, item.username, item.nama, item.kelas, item.status, infoGuru
      ]);
      // PENTING: paksa kolom A (Tanggal) sebagai TEKS murni, bukan Date otomatis dari Google Sheets.
      // Ini mencegah ambiguitas konversi tanggal/timezone yang bisa membuat pencocokan tanggal gagal
      // (menyebabkan status presensi terlihat "reset" & dashboard Admin tidak terupdate).
      var barisTerakhir = sheet.getLastRow();
      sheet.getRange(barisTerakhir, 1).setNumberFormat("@").setValue(tanggalPilihan);
      sheet.getRange(barisTerakhir, 3).setNumberFormat("@").setValue(item.username);
    }
  });
  
  return { status: "SUCCESS", message: "Data presensi tanggal " + tanggalPilihan + " berhasil direkam!" };
}

/**
 * Mengambil presensi yang sudah tersimpan untuk kelas & tanggal tertentu,
 * dikembalikan sebagai peta {username: status} agar frontend bisa langsung pre-fill dropdown status
 */
function ambilPresensiTerpilih(tanggalDipilih, kelasDipilih, infoGuru) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Presensi");
  if (!sheet) return {};
  
  var data = sheet.getDataRange().getValues();
  var petaStatus = {};
  
  for (var i = 1; i < data.length; i++) {
    var tglBaris = data[i][0];
    var tglString = (tglBaris instanceof Date) ? Utilities.formatDate(tglBaris, Session.getScriptTimeZone(), "yyyy-MM-dd") : tglBaris.toString().trim();
    var kelasBaris = data[i][4] ? data[i][4].toString().trim() : "";
    
    if (tglString === tanggalDipilih && kelasBaris === kelasDipilih) {
      petaStatus[data[i][2].toString().trim()] = data[i][5]; // username -> status
    }
  }
  return petaStatus;
}

function ambilSoalUjian() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bank_Soal");
  var data = sheet.getDataRange().getValues();
  var listSoal = [];
  for(var i = 1; i < data.length; i++) {
    listSoal.push({
      id: data[i][0], tipe: data[i][1], pertanyaan: data[i][2],
      opsi: [data[i][3], data[i][4], data[i][5], data[i][6], data[i][7]].filter(Boolean)
    });
  }
  return listSoal; 
}

function submitJawabanServer(jawabanSiswa, token, usernameSiswa) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSoal = ss.getSheetByName("Bank_Soal");
  var dataSoal = sheetSoal.getDataRange().getValues();
  var totalSkorDiperoleh = 0; var totalSkorMaksimal = 0; var mapKunci = {};
  
  for(var i = 1; i < dataSoal.length; i++) {
    mapKunci[dataSoal[i][0]] = { kunci: dataSoal[i][8].toString().trim(), poin: Number(dataSoal[i][9]), tipe: dataSoal[i][1] };
    totalSkorMaksimal += Number(dataSoal[i][9]);
  }
  
  jawabanSiswa.forEach(function(jawab) {
    var soalObj = mapKunci[jawab.idSoal];
    if (soalObj) {
      if (soalObj.tipe === "MCMA") {
        if (jawab.jawabanUser.sort().join(",") === soalObj.kunci.split(",").sort().join(",")) totalSkorDiperoleh += soalObj.poin;
      } else {
        if (jawab.jawabanUser.toString().trim() === soalObj.kunci) totalSkorDiperoleh += soalObj.poin;
      }
    }
  });
  
  var nilaiAkhir = totalSkorMaksimal > 0 ? (totalSkorDiperoleh / totalSkorMaksimal) * 100 : 0;
  nilaiAkhir = Math.min(100, Math.round(nilaiAkhir * 100) / 100);
  ss.getSheetByName("Hasil_Ujian").appendRow([usernameSiswa, token, totalSkorDiperoleh, nilaiAkhir, "Selesai", new Date()]);
  return { status: "SUCCESS", nilai: nilaiAkhir };
}

function kunciUjianSiswa(token, usernameSiswa) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName("Log_Kecurangan").appendRow([usernameSiswa, token, 4, "TERKUNCI"]);
  ss.getSheetByName("Hasil_Ujian").appendRow([usernameSiswa, token, 0, 0, "Terkunci", new Date()]);
  return "LOCKED";
}
// ==================== TAMBAHAN FITUR ADMIN (IMPORT & FILTER) ====================

/**
 * Menerima array objek hasil parsing CSV dari frontend
 * dan memasukkannya secara massal (batch append) ke database Google Sheets
 */
// ==================== TAMBAHAN FITUR ADMIN (IMPORT & FILTER) ====================

/**
 * Memproses Impor Massal dari CSV dengan mempertahankan angka 0 di depan NISN
 * Dan mengembalikan data detail jumlah sukses/gagal untuk notifikasi SweetAlert2
 */
function adminImportUserMassal(dataArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var dataEksisting = sheet.getDataRange().getValues();
  var daftarKelasValid = ambilDaftarKelasMaster(); // Untuk validasi kelas dikenal/tidak
  
  // Ambil semua username eksisting untuk validasi duplikasi
  var setUsername = {};
  for (var i = 1; i < dataEksisting.length; i++) {
    setUsername[dataEksisting[i][0].toString().trim()] = true;
  }
  
  var barisBaru = [];
  var jumlahSukses = 0;
  var detailGagal = []; // {baris, username, nama, alasan}
  var peringatanKelas = []; // Kelas yang belum terdaftar di master data (tidak menggagalkan import)
  
  for (var j = 0; j < dataArray.length; j++) {
    var row = dataArray[j];
    var nomorBaris = j + 2; // +2 karena baris 1 adalah header di file Excel
    var username = row[0] ? row[0].toString().trim() : "";
    var password = row[1] ? row[1].toString().trim() : "";
    var nama = row[2] ? row[2].toString().trim() : "";
    var role = row[3] ? row[3].toString().trim() : "";
    var kelas = row[4] ? row[4].toString().trim().toUpperCase() : "";
    
    if (!username || !password || !role) {
      detailGagal.push({ baris: nomorBaris, username: username || "(kosong)", nama: nama || "-", alasan: "Data wajib (Username/Password/Role) ada yang kosong" });
      continue;
    }
    
    if (setUsername[username]) {
      detailGagal.push({ baris: nomorBaris, username: username, nama: nama || "-", alasan: "Username/NISN duplikat, sudah terdaftar" });
      continue;
    }
    
    if (role === "Siswa" && kelas && daftarKelasValid.indexOf(kelas) === -1) {
      peringatanKelas.push({ username: username, nama: nama, kelas: kelas });
    }
    
    barisBaru.push([
      username, 
      password, 
      nama, 
      role, 
      role === "Siswa" ? kelas : ""
    ]);
    
    setUsername[username] = true;
    jumlahSukses++;
  }
  
  if (barisBaru.length > 0) {
    var barisMulai = sheet.getLastRow() + 1;
    // SOLUSI ANGKA 0: Set format kolom Username & Password sebagai TEKS (bukan pakai tanda kutip literal)
    // agar angka 0 di depan NISN tidak hilang, dan tidak ada karakter ' yang ikut tersimpan di data
    sheet.getRange(barisMulai, 1, barisBaru.length, 2).setNumberFormat("@");
    sheet.getRange(barisMulai, 1, barisBaru.length, 5).setValues(barisBaru);
  }
  
  catatLogAktivitas("Admin", "Import Massal", jumlahSukses + " akun berhasil diimpor, " + detailGagal.length + " baris gagal dari total " + dataArray.length + " baris");
  
  // Mengembalikan response terstruktur untuk dibaca SweetAlert2 di frontend
  return {
    status: "SUCCESS",
    sukses: jumlahSukses,
    gagal: detailGagal.length,
    total: dataArray.length,
    detailGagal: detailGagal,
    peringatanKelas: peringatanKelas
  };
}
// Fungsi Admin untuk Reset Password Pengguna secara Instan
function adminResetPasswordUser(username, passwordBaru) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  
  for(var i = 1; i < data.length; i++) {
    if(data[i][0].toString().trim() === username.trim()) {
      // Kolom B adalah Password (indeks ke-1)
      sheet.getRange(i + 1, 2).setValue(passwordBaru.trim());
      catatLogAktivitas("Admin", "Reset Password", "Mereset password akun " + username);
      return { status: "SUCCESS", message: "Password untuk " + username + " berhasil direset menjadi: " + passwordBaru };
    }
  }
  return { status: "ERROR", message: "User tidak ditemukan." };
}
// 1. Fungsi untuk mengambil data semua siswa dari sheet "Users"
function ambilDaftarSiswa() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  var siswaList = [];
  
  // Ambil hanya user yang memiliki role "Siswa"
  for(var i = 1; i < data.length; i++) {
    if(data[i][3].toString().trim() === "Siswa") {
      siswaList.push({
        username: data[i][0].toString(), // NISN
        nama: data[i][2].toString()       // Nama Siswa
      });
    }
  }
  return siswaList;
}

// ==================== AMBIL KONFIGURASI TAHUN AJARAN GLOBAL ====================

/**
 * Mengambil Tahun Ajaran aktif dari Sheet Pengaturan (Kolom: Nama_Setting | Value)
 */
function ambilTahunAjaranAktif() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Pengaturan");
  if (!sheet) return "2026/2027"; // Fallback aman jika terjadi kendala sheet
  
  // Gunakan getDisplayValues() (bukan getValues()) khusus untuk kolom B,
  // supaya kalau sel tidak sengaja tersimpan sebagai tanggal/angka oleh Sheets,
  // yang terbaca tetap TEKS PERSIS SEPERTI YANG TERLIHAT DI LAYAR, bukan hasil toString() dari objek Date/Number.
  var dataTeks = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < dataTeks.length; i++) {
    // Mencocokkan Nama_Setting di Kolom A (indeks 0)
    if (dataTeks[i][0].toString().trim() === "TAHUN_AJARAN_AKTIF") {
      return dataTeks[i][1].toString().trim(); // Mengambil nilai Value di Kolom B (indeks 1), sebagai teks tampilan
    }
  }
  return "2026/2027"; // Fallback jika baris belum ditambahkan
}

// ==================== MODUL PLOTTING GURU MANDIRI ====================

/**
 * Menyimpan atau memperbarui data tugas mengajar (Mata Pelajaran & Kelas) oleh Guru
 */
function simpanPlottingGuruServer(usernameGuru, namaGuru, mataPelajaran, daftarKelasString) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Plotting_Guru");
  
  // Jika karena suatu hal sheet Plotting_Guru belum terbuat, buat otomatis
  if (!sheet) {
    sheet = ss.insertSheet("Plotting_Guru");
    sheet.appendRow(["Tahun Ajaran", "Username Guru", "Nama Guru", "Mata Pelajaran", "Daftar Kelas"]);
  }
  
  var tahunAktif = ambilTahunAjaranAktif().toString().trim();
  var usernameBersih = usernameGuru.toString().trim();
  var dataEksisting = sheet.getDataRange().getValues();
  var indexBarisDitemukan = -1;
  
  // Cari apakah guru ini sudah pernah menyimpan plotting di tahun ajaran aktif ini
  for (var i = 1; i < dataEksisting.length; i++) {
    var tahunBaris = dataEksisting[i][0] === "" ? "" : dataEksisting[i][0].toString().trim();
    var userBaris = dataEksisting[i][1] === "" ? "" : dataEksisting[i][1].toString().trim();
    if (tahunBaris === tahunAktif && userBaris === usernameBersih) {
      indexBarisDitemukan = i + 1; // Baris spreadsheet (1-indexed)
      break;
    }
  }
  
  if (indexBarisDitemukan > -1) {
    // Jika sudah ada: Overwrite/Update SEKALIGUS kolom Tahun Ajaran (A) untuk menormalkan
    // baris lama yang mungkin tersimpan sebagai tipe tanggal/angka, bukan teks murni.
    sheet.getRange(indexBarisDitemukan, 1, 1, 5)
      .setNumberFormat("@")
      .setValues([[tahunAktif, usernameBersih, namaGuru.trim(), mataPelajaran.trim(), daftarKelasString.trim()]]);
  } else {
    // Jika belum ada: Tambah baris plotting baru
    sheet.appendRow([
      tahunAktif,
      usernameBersih,
      namaGuru.trim(),
      mataPelajaran.trim(),
      daftarKelasString.trim()
    ]);
    sheet.getRange(sheet.getLastRow(), 1, 1, 5).setNumberFormat("@");
  }

  SpreadsheetApp.flush(); // Paksa semua perubahan langsung tertulis, jangan menunggu batch Apps Script selesai

  return {
    status: "SUCCESS",
    message: "Pengaturan mengajar Anda untuk Tahun Ajaran " + tahunAktif + " berhasil disimpan!",
    debug: { modeUpdate: indexBarisDitemukan > -1, barisKe: indexBarisDitemukan }
  };
}

/**
 * FUNGSI DIAGNOSTIK SEMENTARA — boleh dihapus/diabaikan setelah masalah dashboard selesai.
 * Cara pakai: buka aplikasi web-nya, tekan F12 (Console), lalu jalankan:
 *   google.script.run.withSuccessHandler(console.log).debugPlottingGuru('guru')
 * lalu screenshot hasil yang muncul di Console.
 */
function debugPlottingGuru(usernameGuru) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Plotting_Guru");
  var tahunAktif = ambilTahunAjaranAktif();
  
  if (!sheet) return { error: "Sheet Plotting_Guru tidak ditemukan", namaSpreadsheetAktif: ss.getName() };
  
  var data = sheet.getDataRange().getValues();
  var dataTeks = sheet.getDataRange().getDisplayValues();
  var baris = [];
  for (var i = 1; i < data.length; i++) {
    baris.push({
      barisKe: i + 1,
      tahunBaris_getValues: data[i][0],
      tahunBaris_getDisplayValues: dataTeks[i][0],
      tipeData_tahun: typeof data[i][0],
      userBaris: data[i][1],
      cocokTahun: dataTeks[i][0].toString().trim() === tahunAktif,
      cocokUsername: data[i][1].toString().trim() === usernameGuru.toString().trim()
    });
  }
  
  return {
    namaSpreadsheetAktif: ss.getName(),      // Untuk pastikan script terhubung ke spreadsheet yang benar
    urlSpreadsheetAktif: ss.getUrl(),
    tahunAktif_terbaca: tahunAktif,
    usernameGuru_dicari: usernameGuru,
    semuaBarisPlottingGuru: baris
  };
}

/**
 * Mengambil data mata pelajaran dan daftar kelas yang diampu guru untuk membatasi dropdown presensi
 */
/**
 * FUNGSI DIAGNOSTIK SEMENTARA #2 — untuk menelusuri kartu "Status Presensi Hari Ini"
 */
function debugStatusPresensiHariIni(usernameGuru) {
  var hariIniIndex = new Date().getDay();
  var namaHariIni = NAMA_HARI_INDONESIA[hariIniIndex];

  var jadwalGuru = guruAmbilJadwalMengajar(usernameGuru);
  var jadwalHariIni = jadwalGuru.filter(function(j) { return j.hari === namaHariIni; });
  var plotting = ambilKelasAmpoanGuru(usernameGuru);
  var hasilAkhir = guruAmbilStatusPresensiHariIni(usernameGuru);

  return {
    usernameGuru_dipakai: usernameGuru,
    hariIniIndex: hariIniIndex,
    namaHariIni_terbaca: namaHariIni,
    jumlahJadwalGuru_total: jadwalGuru.length,
    jumlahJadwalHariIni: jadwalHariIni.length,
    plotting_mataPelajaran: plotting.mataPelajaran,
    plotting_kelasList: plotting.kelasList,
    HASIL_AKHIR_guruAmbilStatusPresensiHariIni: hasilAkhir,
    panjangHasilAkhir: hasilAkhir.length
  };
}

function ambilKelasAmpoanGuru(usernameGuru) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Plotting_Guru");
  var tahunAktif = ambilTahunAjaranAktif();
  
  if (!sheet) return { mataPelajaran: "", kelasList: [] };
  
  var data = sheet.getDataRange().getValues();
  // dataTeks dipakai KHUSUS untuk membaca kolom Tahun Ajaran (kolom A) sebagai teks tampilan,
  // supaya konsisten dengan ambilTahunAjaranAktif() dan tidak salah baca kalau selnya
  // ternyata tersimpan sebagai tipe tanggal/angka, bukan teks murni.
  var dataTeks = sheet.getDataRange().getDisplayValues();
  
  for (var i = 1; i < data.length; i++) {
    var tahunBaris = dataTeks[i][0].toString().trim();
    var userBaris = data[i][1] === "" ? "" : data[i][1].toString().trim();
    
    if (tahunBaris === tahunAktif && userBaris === usernameGuru.toString().trim()) {
      
      var mapel = data[i][3].toString();
      var kelasRaw = data[i][4].toString(); // Misal formatnya: "X-1, XI-2"
      
      var kelasList = kelasRaw.split(",").map(function(item) {
        return item.trim();
      }).filter(Boolean);
      
      return {
        mataPelajaran: mapel,
        kelasList: kelasList
      };
    }
  }
  
  // Jika guru belum melakukan pengaturan, kembalikan array kosong
  return { mataPelajaran: "", kelasList: [] };
}

// ==================== MODUL LOG AKTIVITAS ADMIN ====================

/**
 * Mencatat satu baris log aktivitas ke sheet "Log_Aktivitas" (dibuat otomatis jika belum ada)
 */
function catatLogAktivitas(pelaku, aksi, keterangan) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log_Aktivitas");
    if (!sheet) {
      sheet = ss.insertSheet("Log_Aktivitas");
      sheet.appendRow(["Waktu", "Pelaku", "Aksi", "Keterangan"]);
    }
    var waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([waktu, pelaku, aksi, keterangan]);
  } catch (err) {
    Logger.log("Gagal mencatat log aktivitas: " + err.message);
  }
}

/**
 * Mengambil 200 log aktivitas terbaru untuk ditampilkan di panel Admin (urut terbaru dulu)
 */
function ambilLogAktivitas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Log_Aktivitas");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var listLog = [];
  for (var i = 1; i < data.length; i++) {
    listLog.push({
      waktu: data[i][0].toString(),
      pelaku: data[i][1],
      aksi: data[i][2],
      keterangan: data[i][3]
    });
  }
  return listLog.reverse().slice(0, 200); // Terbaru dulu, maksimal 200 baris
}

// ==================== MODUL STATISTIK DASHBOARD ADMIN ====================

function adminAmbilStatistikDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUser = ss.getSheetByName("Users");
  var dataUser = sheetUser.getDataRange().getValues();
  
  var totalSiswa = 0, totalGuru = 0, totalAdmin = 0;
  var setKelas = {};
  
  for (var i = 1; i < dataUser.length; i++) {
    var role = dataUser[i][3].toString().trim();
    if (role === "Siswa") {
      totalSiswa++;
      var kelas = dataUser[i].length > 4 ? dataUser[i][4].toString().trim() : "";
      if (kelas) setKelas[kelas] = true;
    } else if (role === "Guru") {
      totalGuru++;
    } else if (role === "Admin") {
      totalAdmin++;
    }
  }
  
  // Hitung presensi yang sudah direkam untuk tanggal hari ini
  var presensiHariIni = 0;
  var sheetPresensi = ss.getSheetByName("Presensi");
  if (sheetPresensi) {
    var dataPresensi = sheetPresensi.getDataRange().getValues();
    var tglHariIni = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    for (var j = 1; j < dataPresensi.length; j++) {
      var tglBaris = dataPresensi[j][0];
      var tglString = (tglBaris instanceof Date) ? Utilities.formatDate(tglBaris, Session.getScriptTimeZone(), "yyyy-MM-dd") : tglBaris.toString().trim();
      if (tglString === tglHariIni) presensiHariIni++;
    }
  }
  
  return {
    totalSiswa: totalSiswa,
    totalGuru: totalGuru,
    totalAdmin: totalAdmin,
    totalKelas: Object.keys(setKelas).length,
    presensiHariIni: presensiHariIni
  };
}

// ==================== MODUL MASTER DATA KELAS ====================

function ambilDaftarKelasMaster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Daftar_Kelas");
  if (!sheet) {
    sheet = ss.insertSheet("Daftar_Kelas");
    sheet.appendRow(["Nama Kelas"]);
  }
  var data = sheet.getDataRange().getValues();
  var listKelas = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) listKelas.push(data[i][0].toString().trim());
  }
  return listKelas;
}

function adminTambahKelasBaru(namaKelas) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Daftar_Kelas");
  if (!sheet) {
    sheet = ss.insertSheet("Daftar_Kelas");
    sheet.appendRow(["Nama Kelas"]);
  }
  
  var namaBersih = namaKelas.trim().toUpperCase();
  if (!namaBersih) return { status: "ERROR", message: "Nama kelas tidak boleh kosong." };
  
  var daftarSekarang = ambilDaftarKelasMaster();
  if (daftarSekarang.indexOf(namaBersih) !== -1) {
    return { status: "ERROR", message: "Kelas \"" + namaBersih + "\" sudah terdaftar." };
  }
  
  sheet.appendRow([namaBersih]);
  catatLogAktivitas("Admin", "Tambah Kelas", "Menambahkan kelas baru: " + namaBersih);
  return { status: "SUCCESS", message: "Kelas \"" + namaBersih + "\" berhasil ditambahkan." };
}

function adminHapusKelasMaster(namaKelas) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Daftar_Kelas");
  if (!sheet) return { status: "ERROR", message: "Data kelas tidak ditemukan." };
  
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0].toString().trim() === namaKelas.trim()) {
      sheet.deleteRow(i + 1);
      catatLogAktivitas("Admin", "Hapus Kelas", "Menghapus kelas: " + namaKelas);
      return { status: "SUCCESS", message: "Kelas \"" + namaKelas + "\" berhasil dihapus dari daftar master." };
    }
  }
  return { status: "ERROR", message: "Kelas tidak ditemukan." };
}

// ==================== MODUL GANTI PASSWORD SENDIRI (Admin/Guru/Siswa) ====================

function gantiPasswordSendiri(username, passwordLama, passwordBaru) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === username.trim()) {
      var passwordTersimpan = data[i][1].toString().trim().replace(/^'/, "");
      if (passwordTersimpan !== passwordLama.trim()) {
        return { status: "ERROR", message: "Password lama yang Anda masukkan salah." };
      }
      sheet.getRange(i + 1, 2).setNumberFormat("@").setValue(passwordBaru.trim());
      catatLogAktivitas(username, "Ganti Password Sendiri", "Pengguna mengganti password akunnya sendiri");
      return { status: "SUCCESS", message: "Password berhasil diperbarui. Gunakan password baru saat login berikutnya." };
    }
  }
  return { status: "ERROR", message: "Akun tidak ditemukan." };
}

// ==================== MODUL EXPORT REKAP PRESENSI ====================

/**
 * Mengambil data presensi untuk keperluan export Excel, bisa difilter per Kelas & Bulan (format "yyyy-MM")
 */
function ambilRekapPresensiUntukExport(bulanFilter, kelasFilter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Presensi");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var hasil = [];
  
  for (var i = 1; i < data.length; i++) {
    var tglBaris = data[i][0];
    var tglString = (tglBaris instanceof Date) ? Utilities.formatDate(tglBaris, Session.getScriptTimeZone(), "yyyy-MM-dd") : tglBaris.toString().trim();
    var bulanBaris = tglString.substring(0, 7); // "yyyy-MM"
    var kelasBaris = data[i][4] ? data[i][4].toString().trim() : "";
    
    if (bulanFilter && bulanFilter !== "SEMUA" && bulanBaris !== bulanFilter) continue;
    if (kelasFilter && kelasFilter !== "SEMUA" && kelasBaris !== kelasFilter) continue;
    
    hasil.push({
      tanggal: tglString,
      waktu: data[i][1],
      username: data[i][2],
      nama: data[i][3],
      kelas: kelasBaris,
      status: data[i][5],
      guru: data[i][6]
    });
  }
  return hasil;
}

// ==================== MODUL MATERI & TUGAS ====================

/**
 * Mengunggah file (base64 dari browser) ke folder Google Drive khusus aplikasi ini,
 * lalu mengatur sharing "anyone with link can view" dan mengembalikan URL-nya.
 */
// ID Folder Google Drive tujuan upload — sudah ditentukan langsung oleh Admin, tidak perlu deteksi otomatis lagi
var FOLDER_ID_MATERI = "1UMYqd45VZ9mxjXZ3IvM2IivBrKpSBc-f";
var FOLDER_ID_TUGAS = "1Otf_a2gBs6u3UdNjHCTyQFWTtnhn8S44";

function uploadFileKeDrive(base64Data, namaFile, mimeType, jenisUpload) {
  try {
    var idFolder = (jenisUpload === "Tugas") ? FOLDER_ID_TUGAS : FOLDER_ID_MATERI;
    var folder = DriveApp.getFolderById(idFolder);

    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, mimeType, namaFile);
    var file = folder.createFile(blob); // Proses inti: file SUDAH tersimpan setelah baris ini berhasil

    // Coba atur sharing supaya siapapun (guru/siswa) bisa membuka linknya.
    // Kalau domain Workspace sekolah membatasi share publik, coba fallback ke "siapapun di domain sekolah",
    // dan kalau itu juga gagal, JANGAN gagalkan keseluruhan proses — file tetap tersimpan, warisi izin folder induknya.
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (errSharing1) {
      try {
        file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (errSharing2) {
        Logger.log("Gagal mengatur sharing (diabaikan, file tetap tersimpan): " + errSharing2.message);
      }
    }
    
    return { status: "SUCCESS", url: file.getUrl(), namaFile: namaFile };
  } catch (err) {
    return { status: "ERROR", message: "Gagal mengunggah file: " + err.message };
  }
}

function _buatIdUnik(prefix) {
  return prefix + "-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);
}

// ---------- MATERI ----------

var HEADER_MATERI_BENAR = ["ID Baris", "ID Induk", "Tanggal Upload", "Guru Username", "Guru Nama", "Mapel", "Kelas", "Tanggal Tampil", "Judul", "Deskripsi", "Link File", "Link Youtube"];

/**
 * Memastikan sheet "Materi" ada dan skemanya sesuai. Kalau sheet lama sudah ada dengan skema
 * berbeda (dari versi sebelumnya), sheet lama di-backup (di-rename), lalu dibuat sheet baru yang bersih.
 */
function _pastikanSheetMateriBenar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Materi");

  if (!sheet) {
    sheet = ss.insertSheet("Materi");
    sheet.appendRow(HEADER_MATERI_BENAR);
    return sheet;
  }

  var headerSaatIni = sheet.getRange(1, 1, 1, HEADER_MATERI_BENAR.length).getValues()[0];
  var headerCocok = HEADER_MATERI_BENAR.every(function(h, idx) { return headerSaatIni[idx] === h; });

  if (!headerCocok) {
    // Cukup selaraskan header di tempat (tidak membuat sheet backup baru, agar tidak banyak sheet)
    sheet.getRange(1, 1, 1, HEADER_MATERI_BENAR.length).setValues([HEADER_MATERI_BENAR]);
  }
  return sheet;
}

/**
 * kelasTanggalArray: [{kelas: "10-A", tanggalTampil: "2026-07-10"}, {kelas: "10-B", tanggalTampil: "2026-07-12"}, ...]
 * Satu upload bisa untuk banyak kelas, MASING-MASING dengan tanggal tampil sendiri (1 baris per kelas di sheet).
 */
function guruUploadMateri(guruUsername, guruNama, mapel, kelasTanggalArray, judul, deskripsi, linkFile, linkYoutube) {
  var sheet = _pastikanSheetMateriBenar();
  var idInduk = _buatIdUnik("MAT");
  var waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  kelasTanggalArray.forEach(function(kt) {
    var idBaris = _buatIdUnik("MATROW");
    sheet.appendRow([idBaris, idInduk, waktu, guruUsername, guruNama, mapel, kt.kelas, kt.tanggalTampil, judul, deskripsi, linkFile || "", linkYoutube || ""]);
    // PENTING: kunci kolom Tanggal Upload (C) & Tanggal Tampil (H) sebagai TEKS,
    // agar Google Sheets tidak otomatis mengonversinya jadi objek Date (yang bikin .replace() error di client)
    var barisTerakhir = sheet.getLastRow();
    sheet.getRange(barisTerakhir, 3).setNumberFormat("@").setValue(waktu);
    sheet.getRange(barisTerakhir, 8).setNumberFormat("@").setValue(kt.tanggalTampil);
  });

  var daftarKelasTeks = kelasTanggalArray.map(function(kt) { return kt.kelas + " (tayang " + kt.tanggalTampil + ")"; }).join(", ");
  catatLogAktivitas(guruNama, "Upload Materi", "Mengupload materi \"" + judul + "\" untuk: " + daftarKelasTeks);

  return { status: "SUCCESS", message: "Materi \"" + judul + "\" berhasil diupload untuk " + kelasTanggalArray.length + " kelas." };
}

/**
 * Mengambil daftar materi milik guru, DIKELOMPOKKAN per ID Induk (1 upload = 1 kartu di tampilan guru,
 * walau tayang di beberapa kelas dengan tanggal berbeda-beda).
 */
function guruAmbilDaftarMateri(guruUsername) {
  var sheet = _pastikanSheetMateriBenar();
  var data = sheet.getDataRange().getValues();
  var grup = {};

  for (var i = 1; i < data.length; i++) {
    if (data[i][3].toString().trim() !== guruUsername.toString().trim()) continue;
    var idInduk = data[i][1];
    if (!grup[idInduk]) {
      grup[idInduk] = {
        idInduk: idInduk, tanggalUpload: data[i][2], mapel: data[i][5],
        judul: data[i][8], deskripsi: data[i][9], linkFile: data[i][10], linkYoutube: data[i][11],
        daftarKelas: []
      };
    }
    grup[idInduk].daftarKelas.push({ idBaris: data[i][0], kelas: data[i][6], tanggalTampil: data[i][7] });
  }

  var hasilAkhir = Object.values(grup).reverse(); // Terbaru dulu

  // DIAGNOSTIK SEMENTARA: kalau hasil kosong, sertakan info detail isi kolom Guru Username yang sebenarnya ada di sheet
  if (hasilAkhir.length === 0 && data.length > 1) {
    var daftarUsernameDitemukan = [];
    for (var j = 1; j < data.length; j++) {
      daftarUsernameDitemukan.push('baris' + (j+1) + ':"' + data[j][3] + '"(tipe:' + (typeof data[j][3]) + ')');
    }
    return { kosong: true, dicari: guruUsername, ditemukan: daftarUsernameDitemukan };
  }

  return hasilAkhir;
}

function guruHapusMateri(idInduk) {
  var sheet = _pastikanSheetMateriBenar();
  var data = sheet.getDataRange().getValues();
  var jumlahDihapus = 0;

  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][1].toString().trim() === idInduk.toString().trim()) {
      sheet.deleteRow(i + 1);
      jumlahDihapus++;
    }
  }
  return jumlahDihapus > 0
    ? { status: "SUCCESS", message: "Materi berhasil dihapus dari " + jumlahDihapus + " kelas." }
    : { status: "ERROR", message: "Materi tidak ditemukan." };
}

/**
 * Siswa hanya melihat materi yang: (a) kelasnya cocok, DAN (b) tanggal tampil-nya sudah tiba (<= hari ini)
 */
function siswaAmbilMateriKelas(kelasSiswa) {
  var sheet = _pastikanSheetMateriBenar();
  var data = sheet.getDataRange().getValues();
  var tglHariIni = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var hasil = [];

  for (var i = 1; i < data.length; i++) {
    var kelasBaris = data[i][6].toString().trim();
    var tanggalTampil = data[i][7].toString().trim();
    if (kelasBaris.toUpperCase() === kelasSiswa.toString().trim().toUpperCase() && tanggalTampil <= tglHariIni) {
      hasil.push({
        id: data[i][0], tanggalUpload: data[i][2], guruNama: data[i][4], mapel: data[i][5],
        judul: data[i][8], deskripsi: data[i][9], linkFile: data[i][10], linkYoutube: data[i][11]
      });
    }
  }
  return hasil.reverse();
}

// ---------- TUGAS ----------

var HEADER_TUGAS_BENAR = ["ID Baris", "ID Induk", "Tanggal Upload", "Guru Username", "Guru Nama", "Mapel", "Kelas", "Tanggal Tampil", "Judul", "Deskripsi", "Deadline", "Link File Lampiran", "Izinkan Terlambat", "Jenis Tugas"];

function _pastikanSheetTugasBenar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Tugas");

  if (!sheet) {
    sheet = ss.insertSheet("Tugas");
    sheet.appendRow(HEADER_TUGAS_BENAR);
    return sheet;
  }

  var headerSaatIni = sheet.getRange(1, 1, 1, HEADER_TUGAS_BENAR.length).getValues()[0];
  var headerCocok = HEADER_TUGAS_BENAR.every(function(h, idx) { return headerSaatIni[idx] === h; });

  if (!headerCocok) {
    // Cukup selaraskan header di tempat (tidak membuat sheet backup baru, agar tidak banyak sheet)
    sheet.getRange(1, 1, 1, HEADER_TUGAS_BENAR.length).setValues([HEADER_TUGAS_BENAR]);
  }
  return sheet;
}

/**
 * kelasTanggalArray: [{kelas: "10-A", tanggalTampil: "2026-07-10"}, ...] — deadline tetap sama untuk semua kelas,
 * tapi TANGGAL TAMPIL (kapan tugas mulai terlihat siswa) bisa berbeda per kelas.
 */
function guruBuatTugas(guruUsername, guruNama, mapel, kelasTanggalArray, judul, deskripsi, deadline, linkFileLampiran, izinkanTerlambat, jenisTugas) {
  var sheet = _pastikanSheetTugasBenar();
  var idInduk = _buatIdUnik("TGS");
  var waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var jenis = (jenisTugas === "Kelompok") ? "Kelompok" : "Individu";

  kelasTanggalArray.forEach(function(kt) {
    var idBaris = _buatIdUnik("TGSROW");
    sheet.appendRow([idBaris, idInduk, waktu, guruUsername, guruNama, mapel, kt.kelas, kt.tanggalTampil, judul, deskripsi, deadline, linkFileLampiran || "", izinkanTerlambat ? "Ya" : "Tidak", jenis]);
    // PENTING: kunci kolom Tanggal Upload (C), Tanggal Tampil (H), dan Deadline (K) sebagai TEKS,
    // agar Google Sheets tidak otomatis mengonversinya jadi objek Date (yang bikin .replace() error di client)
    var barisTerakhir = sheet.getLastRow();
    sheet.getRange(barisTerakhir, 3).setNumberFormat("@").setValue(waktu);
    sheet.getRange(barisTerakhir, 8).setNumberFormat("@").setValue(kt.tanggalTampil);
    sheet.getRange(barisTerakhir, 11).setNumberFormat("@").setValue(deadline);
  });

  var daftarKelasTeks = kelasTanggalArray.map(function(kt) { return kt.kelas + " (tayang " + kt.tanggalTampil + ")"; }).join(", ");
  catatLogAktivitas(guruNama, "Buat Tugas", "Membuat tugas " + jenis + " \"" + judul + "\" untuk: " + daftarKelasTeks + " (deadline: " + deadline + ")");

  return { status: "SUCCESS", message: "Tugas \"" + judul + "\" berhasil dibuat untuk " + kelasTanggalArray.length + " kelas.", idInduk: idInduk };
}

/** Daftar tugas milik guru, dikelompokkan per ID Induk (1 upload = 1 kartu, walau tayang di beberapa kelas) */
function guruAmbilDaftarTugas(guruUsername) {
  var sheet = _pastikanSheetTugasBenar();
  var data = sheet.getDataRange().getValues();
  var grup = {};

  for (var i = 1; i < data.length; i++) {
    if (data[i][3].toString().trim() !== guruUsername.toString().trim()) continue;
    var idInduk = data[i][1];
    if (!grup[idInduk]) {
      grup[idInduk] = {
        idInduk: idInduk, tanggalUpload: data[i][2], mapel: data[i][5],
        judul: data[i][8], deskripsi: data[i][9], deadline: data[i][10],
        linkFileLampiran: data[i][11], izinkanTerlambat: data[i][12],
        jenisTugas: (data[i][13] || "Individu").toString().trim() || "Individu",
        daftarKelas: []
      };
    }
    grup[idInduk].daftarKelas.push({ idBaris: data[i][0], kelas: data[i][6], tanggalTampil: data[i][7] });
  }
  return Object.values(grup).reverse();
}

// ---------- KELOMPOK TUGAS ----------

var HEADER_KELOMPOK_TUGAS = ["ID Kelompok", "ID Baris Tugas", "Nama Kelompok", "Anggota", "Metode"];

function _pastikanSheetKelompokTugasBenar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Kelompok_Tugas");
  if (!sheet) {
    sheet = ss.insertSheet("Kelompok_Tugas");
    sheet.appendRow(HEADER_KELOMPOK_TUGAS);
    return sheet;
  }
  var headerSaatIni = sheet.getRange(1, 1, 1, HEADER_KELOMPOK_TUGAS.length).getValues()[0];
  var headerCocok = HEADER_KELOMPOK_TUGAS.every(function(h, idx) { return headerSaatIni[idx] === h; });
  if (!headerCocok) sheet.getRange(1, 1, 1, HEADER_KELOMPOK_TUGAS.length).setValues([HEADER_KELOMPOK_TUGAS]);
  return sheet;
}

function _acakArray(arr) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = a[i]; a[i] = a[j]; a[j] = tmp;
  }
  return a;
}

/** Daftar siswa satu kelas berdasarkan ID Baris Tugas (dipakai untuk UI Atur Kelompok manual) */
function guruAmbilSiswaUntukKelompok(idBarisTugas) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTugas = _pastikanSheetTugasBenar();
  var dataTugas = sheetTugas.getDataRange().getValues();
  var kelasTarget = "";
  for (var i = 1; i < dataTugas.length; i++) {
    if (dataTugas[i][0].toString().trim() === idBarisTugas.toString().trim()) { kelasTarget = dataTugas[i][6].toString().trim(); break; }
  }
  var sheetUser = ss.getSheetByName("Users");
  var dataUser = sheetUser.getDataRange().getValues();
  var siswaKelas = [];
  for (var u = 1; u < dataUser.length; u++) {
    if (dataUser[u][3].toString().trim() === "Siswa") {
      var kelasSiswa = dataUser[u].length > 4 ? dataUser[u][4].toString().trim() : "";
      if (kelasSiswa.toUpperCase() === kelasTarget.toUpperCase()) {
        siswaKelas.push({ username: dataUser[u][0].toString(), nama: dataUser[u][2].toString() });
      }
    }
  }
  return siswaKelas;
}

/** Hapus semua baris kelompok milik satu baris tugas tertentu (dipakai sebelum menulis ulang) */
function _hapusKelompokLamaUntukTugas(idBarisTugas) {
  var sheet = _pastikanSheetKelompokTugasBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][1].toString().trim() === idBarisTugas.toString().trim()) sheet.deleteRow(i + 1);
  }
}

/**
 * Membuat kelompok secara OTOMATIS untuk satu baris tugas (satu kelas), berdasarkan
 * jumlah anggota per kelompok yang ditentukan guru. Siswa diacak lalu dibagi rata;
 * kelompok terakhir boleh berisi lebih sedikit anggota jika jumlah siswa tidak habis dibagi.
 */
function guruBuatKelompokOtomatis(idBarisTugas, jumlahAnggotaPerKelompok) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTugas = _pastikanSheetTugasBenar();
  var dataTugas = sheetTugas.getDataRange().getValues();
  var kelasTarget = "";
  for (var i = 1; i < dataTugas.length; i++) {
    if (dataTugas[i][0].toString().trim() === idBarisTugas.toString().trim()) { kelasTarget = dataTugas[i][6].toString().trim(); break; }
  }
  if (!kelasTarget) return { status: "ERROR", message: "Baris tugas tidak ditemukan." };

  var sheetUser = ss.getSheetByName("Users");
  var dataUser = sheetUser.getDataRange().getValues();
  var siswaKelas = [];
  for (var u = 1; u < dataUser.length; u++) {
    if (dataUser[u][3].toString().trim() === "Siswa") {
      var kelasSiswa = dataUser[u].length > 4 ? dataUser[u][4].toString().trim() : "";
      if (kelasSiswa.toUpperCase() === kelasTarget.toUpperCase()) {
        siswaKelas.push({ username: dataUser[u][0].toString(), nama: dataUser[u][2].toString() });
      }
    }
  }
  if (siswaKelas.length === 0) return { status: "ERROR", message: "Tidak ada siswa di kelas ini." };

  var jumlahPer = parseInt(jumlahAnggotaPerKelompok, 10);
  if (!jumlahPer || jumlahPer < 1) return { status: "ERROR", message: "Jumlah anggota per kelompok tidak valid." };

  var acak = _acakArray(siswaKelas);
  var kelompokKelompok = [];
  for (var k = 0; k < acak.length; k += jumlahPer) {
    kelompokKelompok.push(acak.slice(k, k + jumlahPer));
  }

  _hapusKelompokLamaUntukTugas(idBarisTugas);
  var sheet = _pastikanSheetKelompokTugasBenar();
  kelompokKelompok.forEach(function(anggota, idx) {
    var idKelompok = _buatIdUnik("KLP");
    sheet.appendRow([idKelompok, idBarisTugas, "Kelompok " + (idx + 1), JSON.stringify(anggota), "Otomatis"]);
  });

  return { status: "SUCCESS", message: "Berhasil membuat " + kelompokKelompok.length + " kelompok secara otomatis (" + jumlahPer + " anggota/kelompok)." };
}

/**
 * Menyimpan pembagian kelompok secara MANUAL.
 * daftarKelompok: [{ namaKelompok: "Kelompok A", anggota: [{username, nama}, ...] }, ...]
 */
function guruSimpanKelompokManual(idBarisTugas, daftarKelompok) {
  if (!daftarKelompok || daftarKelompok.length === 0) return { status: "ERROR", message: "Belum ada kelompok yang dibuat." };

  // Validasi: pastikan tidak ada siswa yang masuk lebih dari satu kelompok
  var sudahDipakai = {};
  for (var i = 0; i < daftarKelompok.length; i++) {
    var anggota = daftarKelompok[i].anggota || [];
    for (var j = 0; j < anggota.length; j++) {
      var uname = anggota[j].username;
      if (sudahDipakai[uname]) return { status: "ERROR", message: "Siswa " + (anggota[j].nama || uname) + " sudah masuk ke lebih dari satu kelompok." };
      sudahDipakai[uname] = true;
    }
  }

  _hapusKelompokLamaUntukTugas(idBarisTugas);
  var sheet = _pastikanSheetKelompokTugasBenar();
  daftarKelompok.forEach(function(klp) {
    var idKelompok = _buatIdUnik("KLP");
    sheet.appendRow([idKelompok, idBarisTugas, klp.namaKelompok || "Kelompok", JSON.stringify(klp.anggota || []), "Manual"]);
  });

  return { status: "SUCCESS", message: "Berhasil menyimpan " + daftarKelompok.length + " kelompok secara manual." };
}

/** Ambil daftar kelompok untuk satu baris tugas */
function guruAmbilKelompokTugas(idBarisTugas) {
  var sheet = _pastikanSheetKelompokTugasBenar();
  var data = sheet.getDataRange().getValues();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() === idBarisTugas.toString().trim()) {
      var anggota = [];
      try { anggota = JSON.parse(data[i][3]); } catch (e) { anggota = []; }
      hasil.push({ idKelompok: data[i][0], namaKelompok: data[i][2], anggota: anggota, metode: data[i][4] });
    }
  }
  return hasil;
}

/** Cari kelompok milik seorang siswa untuk satu baris tugas tertentu */
function _cariKelompokSiswa(idBarisTugas, username) {
  var daftarKelompok = guruAmbilKelompokTugas(idBarisTugas);
  for (var i = 0; i < daftarKelompok.length; i++) {
    var anggota = daftarKelompok[i].anggota || [];
    for (var j = 0; j < anggota.length; j++) {
      if (anggota[j].username && anggota[j].username.toString().trim() === username.toString().trim()) return daftarKelompok[i];
    }
  }
  return null;
}

function guruHapusTugas(idInduk) {
  var sheet = _pastikanSheetTugasBenar();
  var data = sheet.getDataRange().getValues();
  var jumlahDihapus = 0;
  var idBarisTerkait = [];

  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][1].toString().trim() === idInduk.toString().trim()) {
      idBarisTerkait.push(data[i][0].toString().trim());
      sheet.deleteRow(i + 1);
      jumlahDihapus++;
    }
  }
  idBarisTerkait.forEach(function(idBaris) { _hapusKelompokLamaUntukTugas(idBaris); });

  return jumlahDihapus > 0
    ? { status: "SUCCESS", message: "Tugas berhasil dihapus dari " + jumlahDihapus + " kelas." }
    : { status: "ERROR", message: "Tugas tidak ditemukan." };
}

/** Siswa hanya melihat tugas yang: kelasnya cocok DAN tanggal tampil-nya sudah tiba */
function siswaAmbilTugasKelas(kelasSiswa, usernameSiswa) {
  var sheet = _pastikanSheetTugasBenar();
  var dataTugas = sheet.getDataRange().getValues();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetKumpul = _pastikanSheetPengumpulanTugasBenar();
  var dataKumpul = sheetKumpul.getDataRange().getValues();
  var tglHariIni = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  var hasil = [];
  for (var i = 1; i < dataTugas.length; i++) {
    var kelasBaris = dataTugas[i][6].toString().trim();
    var tanggalTampil = dataTugas[i][7].toString().trim();
    if (kelasBaris.toUpperCase() !== kelasSiswa.toString().trim().toUpperCase() || tanggalTampil > tglHariIni) continue;

    var idBaris = dataTugas[i][0].toString().trim();
    var jenisTugas = (dataTugas[i][13] || "Individu").toString().trim() || "Individu";
    var statusKumpul = "Belum Kumpul";
    var nilai = "";
    var catatanGuru = "";
    var usernamePengirim = "";
    var riwayatSubmission = [];

    for (var j = 1; j < dataKumpul.length; j++) {
      if (dataKumpul[j][0].toString().trim() === idBaris && dataKumpul[j][1].toString().trim() === usernameSiswa.toString().trim()) {
        statusKumpul = dataKumpul[j][7].toString();
        nilai = dataKumpul[j][8];
        catatanGuru = dataKumpul[j][9];
        usernamePengirim = dataKumpul[j][11] || "";
        try { riwayatSubmission = JSON.parse(dataKumpul[j][13] || "[]"); } catch (e) { riwayatSubmission = []; }
        break;
      }
    }

    var infoKelompok = null;
    if (jenisTugas === "Kelompok") {
      var kelompokSaya = _cariKelompokSiswa(idBaris, usernameSiswa);
      if (kelompokSaya) {
        infoKelompok = { idKelompok: kelompokSaya.idKelompok, namaKelompok: kelompokSaya.namaKelompok, anggota: kelompokSaya.anggota };
      }
    }

    hasil.push({
      id: idBaris, tanggalUpload: dataTugas[i][2], guruNama: dataTugas[i][4], mapel: dataTugas[i][5],
      judul: dataTugas[i][8], deskripsi: dataTugas[i][9], deadline: dataTugas[i][10],
      linkFileLampiran: dataTugas[i][11], izinkanTerlambat: dataTugas[i][12],
      jenisTugas: jenisTugas, kelompok: infoKelompok, usernamePengirim: usernamePengirim,
      riwayatSubmission: riwayatSubmission,
      statusKumpul: statusKumpul, nilai: nilai, catatanGuru: catatanGuru
    });
  }
  return hasil.reverse();
}

var HEADER_PENGUMPULAN_TUGAS = ["ID Tugas", "Username Siswa", "Nama Siswa", "Kelas", "Waktu Kumpul", "Jenis Jawaban", "Isi Jawaban", "Status Waktu", "Nilai", "Catatan Guru", "", "Username Pengirim", "Laporan Kontribusi", "Riwayat Submission"];

function _pastikanSheetPengumpulanTugasBenar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Pengumpulan_Tugas");
  if (!sheet) {
    sheet = ss.insertSheet("Pengumpulan_Tugas");
    sheet.appendRow(["ID Tugas", "Username Siswa", "Nama Siswa", "Kelas", "Waktu Kumpul", "Jenis Jawaban", "Isi Jawaban", "Status Waktu", "Nilai", "Catatan Guru", "ID Kelompok", "Username Pengirim", "Laporan Kontribusi", "Riwayat Submission"]);
    return sheet;
  }
  var headerSaatIni = sheet.getRange(1, 1, 1, 14).getValues()[0];
  if (!headerSaatIni[10]) sheet.getRange(1, 11).setValue("ID Kelompok");
  if (!headerSaatIni[11]) sheet.getRange(1, 12).setValue("Username Pengirim");
  if (!headerSaatIni[12]) sheet.getRange(1, 13).setValue("Laporan Kontribusi");
  if (!headerSaatIni[13]) sheet.getRange(1, 14).setValue("Riwayat Submission");
  return sheet;
}

/**
 * Siswa mengumpulkan tugas (link Google Drive ATAU teks jawaban langsung).
 * idTugas di sini adalah ID BARIS (spesifik per kelas), bukan ID Induk.
 * Bersifat upsert: kalau siswa submit ulang untuk tugas yang sama, jawaban lama dipindah ke riwayat
 * (TIDAK dihapus) supaya nilai yang sudah diberikan guru sebelumnya tetap ada sebagai acuan.
 *
 * Untuk tugas KELOMPOK: cukup satu anggota yang submit, tapi baris pengumpulan otomatis
 * dibuat/diperbarui untuk SEMUA anggota kelompok (agar tiap anak bisa dinilai per-orang).
 * laporanKontribusi (opsional): [{ username, status: "Ikut"/"Tidak Ikut", catatan }, ...] —
 * hanya untuk anggota LAIN (bukan diri sendiri), dan hanya tampil ke guru, tidak ke siswa lain.
 */
function siswaKumpulkanTugas(idTugas, usernameSiswa, namaSiswa, kelasSiswa, jenisJawaban, isiJawaban, laporanKontribusi) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTugas = _pastikanSheetTugasBenar();
  var dataTugas = sheetTugas.getDataRange().getValues();

  var deadline = null, izinkanTerlambat = "Tidak", judulTugas = "", jenisTugas = "Individu";
  for (var i = 1; i < dataTugas.length; i++) {
    if (dataTugas[i][0].toString().trim() === idTugas.toString().trim()) {
      deadline = new Date(dataTugas[i][10]);
      izinkanTerlambat = dataTugas[i][12].toString().trim();
      judulTugas = dataTugas[i][8];
      jenisTugas = (dataTugas[i][13] || "Individu").toString().trim() || "Individu";
      break;
    }
  }

  var sekarang = new Date();
  var statusWaktu = "Tepat Waktu";
  if (deadline && sekarang > deadline) {
    if (izinkanTerlambat !== "Ya") {
      return { status: "ERROR", message: "Maaf, batas waktu pengumpulan tugas \"" + judulTugas + "\" sudah lewat dan guru tidak mengizinkan pengumpulan terlambat." };
    }
    statusWaktu = "Terlambat";
  }

  // Tentukan daftar username yang perlu diberi baris pengumpulan
  var daftarUsernameTarget = [{ username: usernameSiswa, nama: namaSiswa }];
  var idKelompok = "";
  if (jenisTugas === "Kelompok") {
    var kelompokSaya = _cariKelompokSiswa(idTugas, usernameSiswa);
    if (!kelompokSaya) return { status: "ERROR", message: "Anda belum terdaftar di kelompok manapun untuk tugas ini. Hubungi guru." };
    idKelompok = kelompokSaya.idKelompok;
    daftarUsernameTarget = kelompokSaya.anggota;
  }

  var sheet = _pastikanSheetPengumpulanTugasBenar();
  var waktuKumpul = Utilities.formatDate(sekarang, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var petaLaporan = {};
  (laporanKontribusi || []).forEach(function(l) { petaLaporan[l.username.toString().trim()] = { status: l.status || "", catatan: l.catatan || "" }; });

  daftarUsernameTarget.forEach(function(target) {
    var data = sheet.getDataRange().getValues();
    var laporanUntukIni = target.username.toString().trim() === usernameSiswa.toString().trim() ? "" : JSON.stringify(petaLaporan[target.username.toString().trim()] || {});
    var barisDitemukan = -1;
    for (var k = 1; k < data.length; k++) {
      if (data[k][0].toString().trim() === idTugas.toString().trim() && data[k][1].toString().trim() === target.username.toString().trim()) { barisDitemukan = k; break; }
    }

    if (barisDitemukan >= 0) {
      var lama = data[barisDitemukan];
      var isiLama = lama[6] ? lama[6].toString() : "";
      var riwayat = [];
      try { riwayat = JSON.parse(lama[13] || "[]"); } catch (e) { riwayat = []; }
      if (isiLama) {
        riwayat.push({ waktuKumpul: lama[4], jenisJawaban: lama[5], isiJawaban: lama[6], statusWaktu: lama[7], nilai: lama[8], catatanGuru: lama[9] });
      }
      var baris = barisDitemukan + 1;
      sheet.getRange(baris, 5, 1, 4).setValues([[waktuKumpul, jenisJawaban, isiJawaban, statusWaktu]]);
      sheet.getRange(baris, 5).setNumberFormat("@").setValue(waktuKumpul);
      sheet.getRange(baris, 7).setNumberFormat("@").setValue(isiJawaban);
      sheet.getRange(baris, 11).setValue(idKelompok);
      sheet.getRange(baris, 12).setValue(usernameSiswa);
      if (laporanUntukIni) sheet.getRange(baris, 13).setValue(laporanUntukIni);
      sheet.getRange(baris, 14).setValue(JSON.stringify(riwayat));
    } else {
      sheet.appendRow([idTugas, target.username, target.nama, kelasSiswa, waktuKumpul, jenisJawaban, isiJawaban, statusWaktu, "", "", idKelompok, usernameSiswa, laporanUntukIni, "[]"]);
      var barisTerakhir = sheet.getLastRow();
      sheet.getRange(barisTerakhir, 2).setNumberFormat("@").setValue(target.username);
      sheet.getRange(barisTerakhir, 5).setNumberFormat("@").setValue(waktuKumpul);
      sheet.getRange(barisTerakhir, 7).setNumberFormat("@").setValue(isiJawaban);
    }
  });

  return { status: "SUCCESS", message: "Tugas berhasil dikumpulkan." + (jenisTugas === "Kelompok" ? " (berlaku untuk seluruh anggota kelompok)" : "") + (statusWaktu === "Terlambat" ? " (Ditandai Terlambat)" : "") };
}

/**
 * Guru mengambil daftar pengumpulan siswa untuk SATU baris tugas (satu kelas spesifik).
 * idTugas = ID Baris (bukan ID Induk), karena tiap kelas sekarang punya baris/ID sendiri.
 */
function guruAmbilPengumpulanTugas(idTugas) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTugas = _pastikanSheetTugasBenar();
  var dataTugas = sheetTugas.getDataRange().getValues();

  var kelasTarget = "";
  var jenisTugas = "Individu";
  for (var i = 1; i < dataTugas.length; i++) {
    if (dataTugas[i][0].toString().trim() === idTugas.toString().trim()) {
      kelasTarget = dataTugas[i][6].toString().trim();
      jenisTugas = (dataTugas[i][13] || "Individu").toString().trim() || "Individu";
      break;
    }
  }

  var sheetUser = ss.getSheetByName("Users");
  var dataUser = sheetUser.getDataRange().getValues();
  var semuaSiswaTarget = [];
  for (var u = 1; u < dataUser.length; u++) {
    if (dataUser[u][3].toString().trim() === "Siswa") {
      var kelasSiswa = dataUser[u].length > 4 ? dataUser[u][4].toString().trim() : "";
      if (kelasSiswa.toUpperCase() === kelasTarget.toString().trim().toUpperCase()) {
        semuaSiswaTarget.push({ username: dataUser[u][0].toString(), nama: dataUser[u][2].toString(), kelas: kelasSiswa });
      }
    }
  }

  var sheetKumpul = _pastikanSheetPengumpulanTugasBenar();
  var dataKumpul = sheetKumpul.getDataRange().getValues();

  var hasil = semuaSiswaTarget.map(function(s) {
    var record = {
      username: s.username, nama: s.nama, kelas: s.kelas, statusKumpul: "Belum Kumpul", waktuKumpul: "", jenisJawaban: "", isiJawaban: "",
      nilai: "", catatanGuru: "", idKelompok: "", namaKelompok: "", usernamePengirim: "", laporanKontribusi: null, riwayatSubmission: []
    };
    for (var j = 1; j < dataKumpul.length; j++) {
      if (dataKumpul[j][0].toString().trim() === idTugas.toString().trim() && dataKumpul[j][1].toString().trim() === s.username.toString().trim()) {
        record.statusKumpul = dataKumpul[j][7];
        record.waktuKumpul = dataKumpul[j][4];
        record.jenisJawaban = dataKumpul[j][5];
        record.isiJawaban = dataKumpul[j][6];
        record.nilai = dataKumpul[j][8];
        record.catatanGuru = dataKumpul[j][9];
        record.idKelompok = dataKumpul[j][10] || "";
        record.usernamePengirim = dataKumpul[j][11] || "";
        try { record.laporanKontribusi = dataKumpul[j][12] ? JSON.parse(dataKumpul[j][12]) : null; } catch (e) { record.laporanKontribusi = null; }
        try { record.riwayatSubmission = dataKumpul[j][13] ? JSON.parse(dataKumpul[j][13]) : []; } catch (e) { record.riwayatSubmission = []; }
        break;
      }
    }
    return record;
  });

  // Jika tugas kelompok, sertakan nama kelompok masing-masing siswa & susun juga versi terkelompok untuk tampilan guru
  var kelompokList = [];
  if (jenisTugas === "Kelompok") {
    kelompokList = guruAmbilKelompokTugas(idTugas);
    var petaNamaKelompok = {};
    kelompokList.forEach(function(klp) { petaNamaKelompok[klp.idKelompok] = klp.namaKelompok; });
    hasil.forEach(function(r) { if (r.idKelompok) r.namaKelompok = petaNamaKelompok[r.idKelompok] || ""; });
  }

  return { jenisTugas: jenisTugas, daftarSiswa: hasil, daftarKelompok: kelompokList };
}

function guruBeriNilaiTugas(idTugas, usernameSiswa, nilai, catatan) {
  var sheet = _pastikanSheetPengumpulanTugasBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === idTugas.toString().trim() && data[i][1].toString().trim() === usernameSiswa.toString().trim()) {
      sheet.getRange(i + 1, 9, 1, 2).setValues([[nilai, catatan || ""]]);
      return { status: "SUCCESS", message: "Nilai berhasil disimpan." };
    }
  }
  return { status: "ERROR", message: "Siswa ini belum mengumpulkan tugas, tidak bisa diberi nilai." };
}

/** Beri nilai yang SAMA untuk seluruh anggota satu kelompok sekaligus (mode cepat, default untuk tugas kelompok) */
function guruBeriNilaiSekelompok(idTugas, idKelompok, nilai, catatan) {
  var sheet = _pastikanSheetPengumpulanTugasBenar();
  var data = sheet.getDataRange().getValues();
  var jumlah = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === idTugas.toString().trim() && data[i][10].toString().trim() === idKelompok.toString().trim()) {
      sheet.getRange(i + 1, 9, 1, 2).setValues([[nilai, catatan || ""]]);
      jumlah++;
    }
  }
  return jumlah > 0
    ? { status: "SUCCESS", message: "Nilai " + nilai + " berhasil diberikan ke " + jumlah + " anggota kelompok." }
    : { status: "ERROR", message: "Belum ada anggota kelompok yang mengumpulkan tugas." };
}

// ==================== MODUL REKAP KEHADIRAN & INDIKATOR PRESENSI (GURU) ====================

/**
 * Rekap kehadiran per siswa untuk satu kelas, opsional difilter per bulan (format "yyyy-MM").
 * Mengembalikan jumlah Hadir/Sakit/Izin/Alpa dan persentase kehadiran tiap siswa.
 */
function ambilRekapKehadiranSiswa(kelasDipilih, bulanFilter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Presensi");
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var rekap = {}; // username -> { nama, Hadir, Sakit, Izin, Alpa }

  for (var i = 1; i < data.length; i++) {
    var kelasBaris = data[i][4] ? data[i][4].toString().trim() : "";
    if (kelasBaris !== kelasDipilih) continue;

    var tglBaris = data[i][0];
    var tglString = (tglBaris instanceof Date) ? Utilities.formatDate(tglBaris, Session.getScriptTimeZone(), "yyyy-MM-dd") : tglBaris.toString().trim();
    var bulanBaris = tglString.substring(0, 7);
    if (bulanFilter && bulanFilter !== "SEMUA" && bulanBaris !== bulanFilter) continue;

    var username = data[i][2].toString().trim();
    var nama = data[i][3];
    var status = data[i][5].toString().trim();

    if (!rekap[username]) {
      rekap[username] = { username: username, nama: nama, Hadir: 0, Sakit: 0, Izin: 0, Alpa: 0 };
    }
    if (rekap[username].hasOwnProperty(status)) {
      rekap[username][status]++;
    }
  }

  var hasil = Object.keys(rekap).map(function(username) {
    var r = rekap[username];
    var totalPertemuan = r.Hadir + r.Sakit + r.Izin + r.Alpa;
    var persentase = totalPertemuan > 0 ? Math.round((r.Hadir / totalPertemuan) * 100) : 0;
    return {
      username: r.username, nama: r.nama,
      hadir: r.Hadir, sakit: r.Sakit, izin: r.Izin, alpa: r.Alpa,
      totalPertemuan: totalPertemuan, persentaseHadir: persentase
    };
  });

  // Urutkan berdasarkan nama
  hasil.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
  return hasil;
}

/**
 * Mengecek status presensi HARI INI untuk setiap kelas yang diampu guru,
 * agar guru tidak lupa mengisi presensi (indikator visual di Dasbor).
 */
// ==================== MODUL JADWAL MENGAJAR MINGGUAN ====================

var NAMA_HARI_INDONESIA = [
  "Minggu",
  "Senin",
  "Selasa",
  "Rabu",
  "Kamis",
  "Jumat",
  "Sabtu"
];

/**
 * Mengubah nilai Jam Ke menjadi teks.
 * Nilai lama seperti 7-8 mungkin sudah diubah Google Sheets
 * menjadi objek Date 8 Juli.
 */
function normalisasiJamKeJadwal_(nilaiJam) {
  if (nilaiJam instanceof Date) {
    return Utilities.formatDate(
      nilaiJam,
      Session.getScriptTimeZone(),
      "M-d"
    );
  }

  if (nilaiJam === null || typeof nilaiJam === "undefined") {
    return "";
  }

  return nilaiJam.toString().trim();
}


/**
 * Menentukan nama hari berdasarkan zona waktu project Apps Script.
 */
function ambilNamaHariHariIni_() {
  var zonaWaktu = Session.getScriptTimeZone();

  var tanggalLokal = Utilities.formatDate(
    new Date(),
    zonaWaktu,
    "yyyy-MM-dd"
  ).split("-");

  var tahun = Number(tanggalLokal[0]);
  var bulan = Number(tanggalLokal[1]) - 1;
  var tanggal = Number(tanggalLokal[2]);

  var indeksHari = new Date(tahun, bulan, tanggal).getDay();

  return NAMA_HARI_INDONESIA[indeksHari];
}


/**
 * Menyimpan jadwal baru.
 */
function guruSimpanJadwalMengajar(
  usernameGuru,
  namaGuru,
  hari,
  jamKe,
  mapel,
  kelas
) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Jadwal_Mengajar");

  if (!sheet) {
    sheet = ss.insertSheet("Jadwal_Mengajar");
    sheet.appendRow([
      "ID",
      "Guru Username",
      "Guru Nama",
      "Hari",
      "Jam Ke",
      "Mapel",
      "Kelas"
    ]);
  }

  var usernameBersih = usernameGuru.toString().trim();
  var namaBersih = namaGuru.toString().trim();
  var hariBersih = hari.toString().trim();
  var jamKeBersih = normalisasiJamKeJadwal_(jamKe);
  var mapelBersih = mapel ? mapel.toString().trim() : "";
  var kelasBersih = kelas.toString().trim();

  if (!usernameBersih || !hariBersih || !jamKeBersih || !kelasBersih) {
    return {
      status: "ERROR",
      message: "Hari, jam ke, dan kelas wajib diisi."
    };
  }

  var id = _buatIdUnik("JDW");
  var barisBaru = sheet.getLastRow() + 1;

  // Kolom E dikunci sebagai teks sebelum data dimasukkan.
  sheet.getRange(barisBaru, 5).setNumberFormat("@");

  sheet.getRange(barisBaru, 1, 1, 7).setValues([[
    id,
    usernameBersih,
    namaBersih,
    hariBersih,
    jamKeBersih,
    mapelBersih,
    kelasBersih
  ]]);

  return {
    status: "SUCCESS",
    message:
      "Jadwal " +
      hariBersih +
      " jam ke-" +
      jamKeBersih +
      " (Kelas " +
      kelasBersih +
      ") berhasil ditambahkan."
  };
}


/**
 * Mengambil seluruh jadwal guru.
 * Seluruh nilai dikonversi ke tipe data yang aman dikirim ke frontend.
 */
function guruAmbilJadwalMengajar(usernameGuru) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Jadwal_Mengajar");

  if (!sheet) {
    return [];
  }

  var usernameBersih = usernameGuru.toString().trim();
  var data = sheet.getDataRange().getValues();
  var hasil = [];

  for (var i = 1; i < data.length; i++) {
    var usernameBaris = data[i][1]
      ? data[i][1].toString().trim()
      : "";

    if (usernameBaris === usernameBersih) {
      hasil.push({
        id: data[i][0] ? data[i][0].toString() : "",
        hari: data[i][3] ? data[i][3].toString().trim() : "",
        jamKe: normalisasiJamKeJadwal_(data[i][4]),
        mapel: data[i][5] ? data[i][5].toString().trim() : "",
        kelas: data[i][6] ? data[i][6].toString().trim() : ""
      });
    }
  }

  return hasil;
}


/**
 * Menghapus jadwal berdasarkan ID.
 */
function guruHapusJadwalMengajar(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Jadwal_Mengajar");

  if (!sheet) {
    return {
      status: "ERROR",
      message: "Data jadwal tidak ditemukan."
    };
  }

  var data = sheet.getDataRange().getValues();
  var idBersih = id.toString().trim();

  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0].toString().trim() === idBersih) {
      sheet.deleteRow(i + 1);

      return {
        status: "SUCCESS",
        message: "Jadwal berhasil dihapus."
      };
    }
  }

  return {
    status: "ERROR",
    message: "Jadwal tidak ditemukan."
  };
}


/**
 * Mengambil status presensi kelas yang dijadwalkan hari ini.
 */
function guruAmbilStatusPresensiHariIni(usernameGuru) {
  var usernameBersih = usernameGuru.toString().trim();
  var namaHariIni = ambilNamaHariHariIni_();

  var jadwalGuru = guruAmbilJadwalMengajar(usernameBersih);

  var jadwalHariIni = jadwalGuru.filter(function(jadwal) {
    return jadwal.hari.toLowerCase() === namaHariIni.toLowerCase();
  });

  var daftarTampil = [];

  if (jadwalGuru.length === 0) {
    // Fallback hanya jika guru benar-benar belum membuat jadwal.
    var plotting = ambilKelasAmpoanGuru(usernameBersih);
    var kelasList = plotting.kelasList || [];

    daftarTampil = kelasList.map(function(kelas) {
      return {
        kelas: kelas,
        jamKe: "",
        mapel: plotting.mataPelajaran || ""
      };
    });

  } else if (jadwalHariIni.length === 0) {
    // Guru mempunyai jadwal, tetapi tidak mengajar hari ini.
    return [];

  } else {
    daftarTampil = jadwalHariIni.map(function(jadwal) {
      return {
        kelas: jadwal.kelas,
        jamKe: normalisasiJamKeJadwal_(jadwal.jamKe),
        mapel: jadwal.mapel
      };
    });
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPresensi = ss.getSheetByName("Presensi");

  var tanggalHariIni = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  var kelasSudahIsi = {};

  if (sheetPresensi) {
    var dataPresensi = sheetPresensi.getDataRange().getValues();

    for (var i = 1; i < dataPresensi.length; i++) {
      var tanggalBaris = dataPresensi[i][0];

      var tanggalString = tanggalBaris instanceof Date
        ? Utilities.formatDate(
            tanggalBaris,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd"
          )
        : tanggalBaris.toString().trim();

      var kelasBaris = dataPresensi[i][4]
        ? dataPresensi[i][4].toString().trim()
        : "";

      var guruBaris = dataPresensi[i][6]
        ? dataPresensi[i][6].toString()
        : "";

      // Mencegah presensi guru lain dianggap sebagai presensi guru ini.
      var milikGuruIni =
        guruBaris.indexOf("(" + usernameBersih + ")") !== -1;

      if (tanggalString === tanggalHariIni && milikGuruIni) {
        kelasSudahIsi[kelasBaris] = true;
      }
    }
  }

  return daftarTampil.map(function(dataJadwal) {
    return {
      kelas: dataJadwal.kelas,
      jamKe: normalisasiJamKeJadwal_(dataJadwal.jamKe),
      mapel: dataJadwal.mapel,
      sudahPresensi: !!kelasSudahIsi[dataJadwal.kelas]
    };
  });
}


/**
 * Jalankan satu kali untuk memperbaiki data Jam Ke lama
 * yang sudah berubah menjadi tanggal.
 */
function perbaikiFormatJamKeJadwalLama() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Jadwal_Mengajar");

  if (!sheet) {
    return "Sheet Jadwal_Mengajar tidak ditemukan.";
  }

  var data = sheet.getDataRange().getValues();
  var jumlahDiperbaiki = 0;

  for (var i = 1; i < data.length; i++) {
    var nilaiJam = data[i][4];

    if (nilaiJam instanceof Date) {
      var jamKeTeks = normalisasiJamKeJadwal_(nilaiJam);
      var selJamKe = sheet.getRange(i + 1, 5);

      selJamKe.setNumberFormat("@");
      selJamKe.setValue(jamKeTeks);

      jumlahDiperbaiki++;
    }
  }

  return (
    "Selesai. " +
    jumlahDiperbaiki +
    " data Jam Ke berhasil diubah menjadi teks."
  );
}

// FUNGSI MIGRASI (jalankan SEKALI SAJA dari editor Apps Script jika presensi lama tidak terbaca konsisten, lalu boleh dihapus)
// Mengubah semua nilai Tanggal di sheet Presensi yang masih berupa Date object menjadi TEKS murni "yyyy-MM-dd"
function perbaikiFormatTanggalPresensiLama() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Presensi");
  if (!sheet) return "Sheet Presensi tidak ditemukan.";

  var data = sheet.getDataRange().getValues();
  var jumlahDiperbaiki = 0;

  for (var i = 1; i < data.length; i++) {
    var tglBaris = data[i][0];
    if (tglBaris instanceof Date) {
      var tglString = Utilities.formatDate(tglBaris, Session.getScriptTimeZone(), "yyyy-MM-dd");
      sheet.getRange(i + 1, 1).setNumberFormat("@").setValue(tglString);
      jumlahDiperbaiki++;
    }
  }
  Logger.log("Selesai. Total baris tanggal presensi diperbaiki: " + jumlahDiperbaiki);
  return "Selesai. Total baris tanggal presensi diperbaiki: " + jumlahDiperbaiki;
}

// FUNGSI MIGRASI (jalankan SEKALI SAJA dari editor Apps Script, lalu boleh dihapus)
// Memperbaiki NISN/Username di sheet Presensi yang sudah terlanjur kehilangan angka 0 di depan
// (karena Google Sheets otomatis membacanya sebagai angka). Dicocokkan ulang dengan data asli di sheet Users
// menggunakan nilai numeriknya, lalu ditimpa dengan versi lengkap (dengan nol) dari sheet Users.
function perbaikiUsernamePresensiLama() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPresensi = ss.getSheetByName("Presensi");
  var sheetUsers = ss.getSheetByName("Users");
  if (!sheetPresensi || !sheetUsers) return "Sheet Presensi atau Users tidak ditemukan.";

  // Buat peta: versi numerik NISN -> username lengkap asli (dengan nol) dari sheet Users
  var dataUsers = sheetUsers.getDataRange().getValues();
  var petaUsernameAsli = {};
  for (var u = 1; u < dataUsers.length; u++) {
    var usernameAsli = dataUsers[u][0].toString().trim().replace(/^'/, "");
    var versiNumerik = String(Number(usernameAsli)); // "0082175042" -> "82175042"
    petaUsernameAsli[versiNumerik] = usernameAsli;
  }

  var dataPresensi = sheetPresensi.getDataRange().getValues();
  var jumlahDiperbaiki = 0;

  for (var i = 1; i < dataPresensi.length; i++) {
    var usernameSaatIni = dataPresensi[i][2].toString().trim();
    var versiNumerikPresensi = String(Number(usernameSaatIni));
    var usernameSeharusnya = petaUsernameAsli[versiNumerikPresensi];

    if (usernameSeharusnya && usernameSeharusnya !== usernameSaatIni) {
      sheetPresensi.getRange(i + 1, 3).setNumberFormat("@").setValue(usernameSeharusnya);
      jumlahDiperbaiki++;
    }
  }
  Logger.log("Selesai. Total baris username presensi diperbaiki: " + jumlahDiperbaiki);
  return "Selesai. Total baris username presensi diperbaiki: " + jumlahDiperbaiki;
}

// FUNGSI KHUSUS UNTUK MEMICU JENDELA OTORISASI GOOGLE DRIVE (VERSI TULIS/WRITE)
// Jalankan fungsi ini SEKALI dari editor Apps Script (pilih di dropdown lalu klik Run).
function testOtorisasiDrive() {
  var folder = DriveApp.createFolder("_test_otorisasi_boleh_dihapus");
  Logger.log("Otorisasi TULIS Drive berhasil! Folder test dibuat: " + folder.getName());
  folder.setTrashed(true); // Langsung hapus folder test-nya
  return "Otorisasi Drive (tulis) berhasil!";
}

// FUNGSI PEMBERSIH (jalankan SEKALI dari editor Apps Script kalau ada sheet "_LAMA_BACKUP_" tersisa dari percobaan sebelumnya)
// Menghapus semua sheet backup lama agar daftar sheet tidak membingungkan
function hapusSemuaSheetBackupLama() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var semuaSheet = ss.getSheets();
  var dihapus = [];

  semuaSheet.forEach(function(sheet) {
    if (sheet.getName().indexOf("_LAMA_BACKUP_") !== -1) {
      dihapus.push(sheet.getName());
      ss.deleteSheet(sheet);
    }
  });

  var pesan = dihapus.length > 0 ? "Sheet dihapus: " + dihapus.join(", ") : "Tidak ada sheet backup yang ditemukan.";
  Logger.log(pesan);
  return pesan;
}

// ==================== EDIT JADWAL TAMPIL MATERI & TUGAS (setelah diupload) ====================

/** Mengubah Tanggal Tampil satu baris Materi (satu kelas spesifik), tanpa perlu hapus-upload ulang */
function guruUbahTanggalTampilMateri(idBaris, tanggalTampilBaru) {
  var sheet = _pastikanSheetMateriBenar();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === idBaris.toString().trim()) {
      sheet.getRange(i + 1, 8).setNumberFormat("@").setValue(tanggalTampilBaru); // Kolom H = Tanggal Tampil (teks, anti auto-convert Date)
      catatLogAktivitas(data[i][4], "Ubah Jadwal Materi", "Mengubah tanggal tampil materi \"" + data[i][8] + "\" (Kelas " + data[i][6] + ") menjadi " + tanggalTampilBaru);
      return { status: "SUCCESS", message: "Tanggal tampil berhasil diperbarui menjadi " + tanggalTampilBaru + "." };
    }
  }
  return { status: "ERROR", message: "Data materi tidak ditemukan." };
}

/** Mengubah Tanggal Tampil DAN/ATAU Deadline satu baris Tugas (satu kelas spesifik) */
function guruUbahJadwalTugas(idBaris, tanggalTampilBaru, deadlineBaru) {
  var sheet = _pastikanSheetTugasBenar();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === idBaris.toString().trim()) {
      sheet.getRange(i + 1, 8).setNumberFormat("@").setValue(tanggalTampilBaru);  // Kolom H = Tanggal Tampil (teks)
      sheet.getRange(i + 1, 11).setNumberFormat("@").setValue(deadlineBaru);      // Kolom K = Deadline (teks)
      catatLogAktivitas(data[i][4], "Ubah Jadwal Tugas", "Mengubah jadwal tugas \"" + data[i][8] + "\" (Kelas " + data[i][6] + ") — tampil: " + tanggalTampilBaru + ", deadline: " + deadlineBaru);
      return { status: "SUCCESS", message: "Jadwal tugas berhasil diperbarui." };
    }
  }
  return { status: "ERROR", message: "Data tugas tidak ditemukan." };
}

// FUNGSI MIGRASI (jalankan SEKALI dari editor Apps Script, lalu boleh dihapus)
// Memperbaiki kolom tanggal di sheet Materi & Tugas yang sudah terlanjur berubah jadi objek Date
// (bukan teks), yang menyebabkan tampilan gagal dimuat di aplikasi.
function perbaikiFormatTanggalMateriTugasLama() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hasil = [];

  function perbaikiKolom(namaSheet, nomorKolom, labelKolom) {
    var sheet = ss.getSheetByName(namaSheet);
    if (!sheet) { hasil.push(namaSheet + ": sheet tidak ditemukan."); return; }
    var data = sheet.getDataRange().getValues();
    var jumlah = 0;
    for (var i = 1; i < data.length; i++) {
      var nilai = data[i][nomorKolom - 1];
      if (nilai instanceof Date) {
        var format = (labelKolom === "Deadline") ? "yyyy-MM-dd'T'HH:mm" : "yyyy-MM-dd";
        var teks = Utilities.formatDate(nilai, Session.getScriptTimeZone(), format);
        sheet.getRange(i + 1, nomorKolom).setNumberFormat("@").setValue(teks);
        jumlah++;
      }
    }
    hasil.push(namaSheet + " kolom " + labelKolom + ": " + jumlah + " baris diperbaiki.");
  }

  perbaikiKolom("Materi", 3, "Tanggal Upload");
  perbaikiKolom("Materi", 8, "Tanggal Tampil");
  perbaikiKolom("Tugas", 3, "Tanggal Upload");
  perbaikiKolom("Tugas", 8, "Tanggal Tampil");
  perbaikiKolom("Tugas", 11, "Deadline");

  var pesan = hasil.join("\n");
  Logger.log(pesan);
  return pesan;
}

// FUNGSI MIGRASI (jalankan SEKALI dari editor Apps Script, lalu boleh dihapus)
// Memperbaiki Username Siswa di Pengumpulan_Tugas yang sudah kehilangan angka 0 di depan NISN,
// dicocokkan ulang dengan data asli di sheet Users (sama seperti perbaikan Presensi sebelumnya).
function perbaikiUsernamePengumpulanTugasLama() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetKumpul = ss.getSheetByName("Pengumpulan_Tugas");
  var sheetUsers = ss.getSheetByName("Users");
  if (!sheetKumpul || !sheetUsers) return "Sheet Pengumpulan_Tugas atau Users tidak ditemukan.";

  var dataUsers = sheetUsers.getDataRange().getValues();
  var petaUsernameAsli = {};
  for (var u = 1; u < dataUsers.length; u++) {
    var usernameAsli = dataUsers[u][0].toString().trim().replace(/^'/, "");
    petaUsernameAsli[String(Number(usernameAsli))] = usernameAsli;
  }

  var data = sheetKumpul.getDataRange().getValues();
  var jumlahDiperbaiki = 0;

  for (var i = 1; i < data.length; i++) {
    var usernameSaatIni = data[i][1].toString().trim();
    var usernameSeharusnya = petaUsernameAsli[String(Number(usernameSaatIni))];
    if (usernameSeharusnya && usernameSeharusnya !== usernameSaatIni) {
      sheetKumpul.getRange(i + 1, 2).setNumberFormat("@").setValue(usernameSeharusnya);
      jumlahDiperbaiki++;
    }
  }
  var pesan = "Selesai. Total baris username pengumpulan tugas diperbaiki: " + jumlahDiperbaiki;
  Logger.log(pesan);
  return pesan;
}

// ==================== MODUL DASHBOARD SISWA: KEHADIRAN & PENGINGAT TUGAS ====================

/**
 * Rekap kehadiran siswa yang sedang login, dikelompokkan per Mata Pelajaran
 * (dengan menelusuri siapa "Guru Pengajar" tiap baris presensi, lalu mencocokkan mapel guru tsb).
 */
function siswaAmbilRekapKehadiranSaya(usernameSiswa) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPresensi = ss.getSheetByName("Presensi");
  if (!sheetPresensi) return [];

  // Buat peta: username guru -> mapel (dari Plotting_Guru)
  var sheetPlotting = ss.getSheetByName("Plotting_Guru");
  var petaMapelGuru = {};
  if (sheetPlotting) {
    var dataPlotting = sheetPlotting.getDataRange().getValues();
    for (var p = 1; p < dataPlotting.length; p++) {
      petaMapelGuru[dataPlotting[p][1].toString().trim()] = dataPlotting[p][3]; // Username Guru -> Mata Pelajaran
    }
  }

  var data = sheetPresensi.getDataRange().getValues();
  var rekap = {}; // mapel -> { Hadir, Sakit, Izin, Alpa }

  for (var i = 1; i < data.length; i++) {
    if (data[i][2].toString().trim() !== usernameSiswa.toString().trim()) continue;

    var infoGuru = data[i][6] ? data[i][6].toString() : ""; // Contoh: "Hery Heryana, S.T (guru)"
    var cocokUsername = infoGuru.match(/\(([^)]+)\)/); // Ambil teks dalam kurung -> username guru
    var usernameGuru = cocokUsername ? cocokUsername[1].trim() : "";
    var mapel = petaMapelGuru[usernameGuru] || "Lainnya";

    var status = data[i][5].toString().trim();
    if (!rekap[mapel]) rekap[mapel] = { mapel: mapel, hadir: 0, sakit: 0, izin: 0, alpa: 0 };
    if (status === "Hadir") rekap[mapel].hadir++;
    else if (status === "Sakit") rekap[mapel].sakit++;
    else if (status === "Izin") rekap[mapel].izin++;
    else if (status === "Alpa") rekap[mapel].alpa++;
  }

  return Object.values(rekap).map(function(r) {
    var total = r.hadir + r.sakit + r.izin + r.alpa;
    r.persentaseHadir = total > 0 ? Math.round((r.hadir / total) * 100) : 0;
    r.totalPertemuan = total;
    return r;
  });
}

/**
 * Ringkasan status tugas siswa yang login: total, sudah kumpul, belum kumpul,
 * plus daftar judul tugas yang BELUM dikumpulkan (untuk pengingat di Dashboard).
 */
function siswaAmbilRingkasanTugasSaya(usernameSiswa, kelasSiswa) {
  var daftarTugas = siswaAmbilTugasKelas(kelasSiswa, usernameSiswa);

  var sudahKumpul = 0, belumKumpul = 0;
  var daftarBelumKumpul = [];

  daftarTugas.forEach(function(t) {
    if (t.statusKumpul === "Belum Kumpul") {
      belumKumpul++;
      daftarBelumKumpul.push({ judul: t.judul, mapel: t.mapel, deadline: t.deadline });
    } else {
      sudahKumpul++;
    }
  });

  return {
    total: daftarTugas.length,
    sudahKumpul: sudahKumpul,
    belumKumpul: belumKumpul,
    daftarBelumKumpul: daftarBelumKumpul
  };
}

// ==================== EDIT LENGKAP MATERI & TUGAS (judul/deskripsi/link + tambah kelas yang terlewat) ====================

function guruAmbilDetailMateri(idInduk) {
  var sheet = _pastikanSheetMateriBenar();
  var data = sheet.getDataRange().getValues();
  var detail = null;
  var daftarKelas = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() === idInduk.toString().trim()) {
      if (!detail) {
        detail = { idInduk: idInduk, mapel: data[i][5], judul: data[i][8], deskripsi: data[i][9], linkFile: data[i][10], linkYoutube: data[i][11] };
      }
      daftarKelas.push({ kelas: data[i][6], tanggalTampil: data[i][7] });
    }
  }
  if (!detail) return { status: "ERROR", message: "Materi tidak ditemukan." };
  detail.daftarKelas = daftarKelas;
  detail.status = "SUCCESS";
  return detail;
}

/** kelasTanggalArrayBaru: SEMUA kelas yang harusnya tayang (existing + baru). Kelas yang sudah ada di-update tanggalnya, yang belum ada ditambah baris baru (idInduk sama). */
function guruUpdateMateri(idInduk, judul, deskripsi, linkFile, linkYoutube, kelasTanggalArrayBaru) {
  var sheet = _pastikanSheetMateriBenar();
  var data = sheet.getDataRange().getValues();
  var guruUsername = "", guruNama = "", mapel = "";
  var kelasSudahAda = {};

  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() === idInduk.toString().trim()) {
      guruUsername = data[i][3]; guruNama = data[i][4]; mapel = data[i][5];
      sheet.getRange(i + 1, 9).setValue(judul);
      sheet.getRange(i + 1, 10).setValue(deskripsi);
      sheet.getRange(i + 1, 11).setValue(linkFile || "");
      sheet.getRange(i + 1, 12).setValue(linkYoutube || "");
      kelasSudahAda[data[i][6].toString().trim()] = i + 1;
    }
  }
  if (!guruUsername) return { status: "ERROR", message: "Materi tidak ditemukan." };

  var waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var jumlahDitambah = 0;

  kelasTanggalArrayBaru.forEach(function(kt) {
    if (kelasSudahAda[kt.kelas]) {
      sheet.getRange(kelasSudahAda[kt.kelas], 8).setNumberFormat("@").setValue(kt.tanggalTampil);
    } else {
      var idBaris = _buatIdUnik("MATROW");
      sheet.appendRow([idBaris, idInduk, waktu, guruUsername, guruNama, mapel, kt.kelas, kt.tanggalTampil, judul, deskripsi, linkFile || "", linkYoutube || ""]);
      var barisTerakhir = sheet.getLastRow();
      sheet.getRange(barisTerakhir, 3).setNumberFormat("@").setValue(waktu);
      sheet.getRange(barisTerakhir, 8).setNumberFormat("@").setValue(kt.tanggalTampil);
      jumlahDitambah++;
    }
  });

  catatLogAktivitas(guruNama, "Edit Materi", "Mengedit materi \"" + judul + "\"" + (jumlahDitambah > 0 ? ", menambah " + jumlahDitambah + " kelas baru" : ""));
  return { status: "SUCCESS", message: "Materi berhasil diperbarui." + (jumlahDitambah > 0 ? " (" + jumlahDitambah + " kelas baru ditambahkan)" : "") };
}

function guruAmbilDetailTugas(idInduk) {
  var sheet = _pastikanSheetTugasBenar();
  var data = sheet.getDataRange().getValues();
  var detail = null;
  var daftarKelas = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() === idInduk.toString().trim()) {
      if (!detail) {
        detail = { idInduk: idInduk, mapel: data[i][5], judul: data[i][8], deskripsi: data[i][9], deadline: data[i][10], linkFileLampiran: data[i][11], izinkanTerlambat: data[i][12], jenisTugas: (data[i][13] || "Individu").toString().trim() || "Individu" };
      }
      daftarKelas.push({ idBaris: data[i][0], kelas: data[i][6], tanggalTampil: data[i][7] });
    }
  }
  if (!detail) return { status: "ERROR", message: "Tugas tidak ditemukan." };
  detail.daftarKelas = daftarKelas;
  detail.status = "SUCCESS";
  return detail;
}

function guruUpdateTugas(idInduk, judul, deskripsi, deadline, linkFileLampiran, izinkanTerlambat, kelasTanggalArrayBaru) {
  var sheet = _pastikanSheetTugasBenar();
  var data = sheet.getDataRange().getValues();
  var guruUsername = "", guruNama = "", mapel = "", jenisTugas = "Individu";
  var kelasSudahAda = {};

  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() === idInduk.toString().trim()) {
      guruUsername = data[i][3]; guruNama = data[i][4]; mapel = data[i][5];
      jenisTugas = (data[i][13] || "Individu").toString().trim() || "Individu";
      sheet.getRange(i + 1, 9).setValue(judul);
      sheet.getRange(i + 1, 10).setValue(deskripsi);
      sheet.getRange(i + 1, 11).setNumberFormat("@").setValue(deadline);
      sheet.getRange(i + 1, 12).setValue(linkFileLampiran || "");
      sheet.getRange(i + 1, 13).setValue(izinkanTerlambat ? "Ya" : "Tidak");
      kelasSudahAda[data[i][6].toString().trim()] = i + 1;
    }
  }
  if (!guruUsername) return { status: "ERROR", message: "Tugas tidak ditemukan." };

  var waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var jumlahDitambah = 0;

  kelasTanggalArrayBaru.forEach(function(kt) {
    if (kelasSudahAda[kt.kelas]) {
      sheet.getRange(kelasSudahAda[kt.kelas], 8).setNumberFormat("@").setValue(kt.tanggalTampil);
    } else {
      var idBaris = _buatIdUnik("TGSROW");
      sheet.appendRow([idBaris, idInduk, waktu, guruUsername, guruNama, mapel, kt.kelas, kt.tanggalTampil, judul, deskripsi, deadline, linkFileLampiran || "", izinkanTerlambat ? "Ya" : "Tidak", jenisTugas]);
      var barisTerakhir = sheet.getLastRow();
      sheet.getRange(barisTerakhir, 3).setNumberFormat("@").setValue(waktu);
      sheet.getRange(barisTerakhir, 8).setNumberFormat("@").setValue(kt.tanggalTampil);
      sheet.getRange(barisTerakhir, 11).setNumberFormat("@").setValue(deadline);
      jumlahDitambah++;
    }
  });

  catatLogAktivitas(guruNama, "Edit Tugas", "Mengedit tugas \"" + judul + "\"" + (jumlahDitambah > 0 ? ", menambah " + jumlahDitambah + " kelas baru" : ""));
  return { status: "SUCCESS", message: "Tugas berhasil diperbarui." + (jumlahDitambah > 0 ? " (" + jumlahDitambah + " kelas baru ditambahkan)" : "") };
}

// FUNGSI MIGRASI (jalankan SEKALI dari editor Apps Script, lalu boleh dihapus)
// Memperbaiki kolom Waktu Kumpul & Isi Jawaban di Pengumpulan_Tugas yang sudah terlanjur
// dikonversi Google Sheets jadi objek Date/Angka (bukan teks), penyebab macet saat guru
// membuka daftar pengumpulan tugas.
function perbaikiFormatPengumpulanTugasLama() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Pengumpulan_Tugas");
  if (!sheet) return "Sheet Pengumpulan_Tugas tidak ditemukan.";

  var data = sheet.getDataRange().getValues();
  var jumlah = 0;

  for (var i = 1; i < data.length; i++) {
    var waktuKumpul = data[i][4];
    var isiJawaban = data[i][6];
    var perluDiperbaiki = false;
    var teksWaktu = waktuKumpul, teksIsi = isiJawaban;

    if (waktuKumpul instanceof Date) {
      teksWaktu = Utilities.formatDate(waktuKumpul, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      perluDiperbaiki = true;
    }
    if (typeof isiJawaban === "number" || isiJawaban instanceof Date) {
      teksIsi = isiJawaban.toString();
      perluDiperbaiki = true;
    }

    if (perluDiperbaiki) {
      sheet.getRange(i + 1, 5).setNumberFormat("@").setValue(teksWaktu);
      sheet.getRange(i + 1, 7).setNumberFormat("@").setValue(teksIsi);
      jumlah++;
    }
  }
  var pesan = "Selesai. Total baris diperbaiki: " + jumlah;
  Logger.log(pesan);
  return pesan;
}

// ==================== NOTIFIKASI TUGAS PERLU DIPERIKSA (DASHBOARD GURU) ====================

/**
 * Menghitung jumlah pengumpulan tugas yang BELUM DINILAI (kolom Nilai kosong) untuk setiap
 * tugas milik guru ybs, supaya guru diingatkan mana yang perlu segera diperiksa.
 */
function guruAmbilNotifikasiPengumpulanTugas(guruUsername) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTugas = _pastikanSheetTugasBenar();
  var dataTugas = sheetTugas.getDataRange().getValues();

  // Ambil semua baris tugas (per kelas) milik guru ini: idBaris -> {judul, kelas}
  var petaTugasGuru = {};
  for (var i = 1; i < dataTugas.length; i++) {
    if (dataTugas[i][3].toString().trim() === guruUsername.toString().trim()) {
      petaTugasGuru[dataTugas[i][0].toString().trim()] = { judul: dataTugas[i][8], kelas: dataTugas[i][6] };
    }
  }

  var sheetKumpul = ss.getSheetByName("Pengumpulan_Tugas");
  if (!sheetKumpul) return { totalBelumDinilai: 0, daftarTugas: [] };

  var dataKumpul = sheetKumpul.getDataRange().getValues();
  var rekapPerTugas = {}; // idBaris -> jumlah belum dinilai

  for (var j = 1; j < dataKumpul.length; j++) {
    var idBarisTugas = dataKumpul[j][0].toString().trim();
    if (!petaTugasGuru[idBarisTugas]) continue; // Bukan tugas milik guru ini

    var nilai = dataKumpul[j][8];
    var sudahDinilai = nilai !== "" && nilai !== null && nilai !== undefined;
    if (!sudahDinilai) {
      rekapPerTugas[idBarisTugas] = (rekapPerTugas[idBarisTugas] || 0) + 1;
    }
  }

  var daftarTugas = Object.keys(rekapPerTugas).map(function(idBaris) {
    return {
      idBaris: idBaris,
      judul: petaTugasGuru[idBaris].judul,
      kelas: petaTugasGuru[idBaris].kelas,
      jumlahBelumDinilai: rekapPerTugas[idBaris]
    };
  });

  var totalBelumDinilai = daftarTugas.reduce(function(total, t) { return total + t.jumlahBelumDinilai; }, 0);

  return { totalBelumDinilai: totalBelumDinilai, daftarTugas: daftarTugas };
}

// ==================== MODUL REKAP NILAI & KEHADIRAN (GURU & SISWA) ====================

/**
 * Helper: hitung persentase kehadiran satu siswa (semua mapel digabung), dari sheet Presensi.
 */
function _hitungPersentaseKehadiran(usernameSiswa) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Presensi");
  if (!sheet) return { persentase: 0, hadir: 0, sakit: 0, izin: 0, alpa: 0, total: 0 };

  var data = sheet.getDataRange().getValues();
  var hadir = 0, sakit = 0, izin = 0, alpa = 0;

  for (var i = 1; i < data.length; i++) {
    if (data[i][2].toString().trim() !== usernameSiswa.toString().trim()) continue;
    var status = data[i][5].toString().trim();
    if (status === "Hadir") hadir++;
    else if (status === "Sakit") sakit++;
    else if (status === "Izin") izin++;
    else if (status === "Alpa") alpa++;
  }

  var total = hadir + sakit + izin + alpa;
  return { persentase: total > 0 ? Math.round((hadir / total) * 100) : 0, hadir: hadir, sakit: sakit, izin: izin, alpa: alpa, total: total };
}

/**
 * REKAP GURU: tabel nilai seluruh siswa di satu kelas, untuk semua tugas yang guru ini buat & sudah tayang di kelas tsb.
 * Siswa yang tidak mengumpulkan otomatis diberi nilai 0. Siswa yang sudah kumpul tapi belum dinilai -> "Belum Dinilai".
 */
function guruAmbilRekapNilaiKelas(guruUsername, kelas) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Ambil semua siswa di kelas ini
  var sheetUser = ss.getSheetByName("Users");
  var dataUser = sheetUser.getDataRange().getValues();
  var daftarSiswa = [];
  for (var u = 1; u < dataUser.length; u++) {
    if (dataUser[u][3].toString().trim() === "Siswa" && dataUser[u].length > 4 && dataUser[u][4].toString().trim().toUpperCase() === kelas.toString().trim().toUpperCase()) {
      daftarSiswa.push({ username: dataUser[u][0].toString(), nama: dataUser[u][2].toString() });
    }
  }

  // 2. Ambil semua baris Tugas milik guru ini yang tayang di kelas ini, dan SUDAH LEWAT tanggal tampilnya
  var sheetTugas = _pastikanSheetTugasBenar();
  var dataTugas = sheetTugas.getDataRange().getValues();
  var tglHariIni = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var daftarTugas = []; // {idBaris, judul}

  for (var t = 1; t < dataTugas.length; t++) {
    if (dataTugas[t][3].toString().trim() !== guruUsername.toString().trim()) continue;
    if (dataTugas[t][6].toString().trim().toUpperCase() !== kelas.toString().trim().toUpperCase()) continue;
    if (dataTugas[t][7].toString().trim() > tglHariIni) continue; // Belum tayang, jangan dimasukkan rekap dulu
    daftarTugas.push({ idBaris: dataTugas[t][0].toString().trim(), judul: dataTugas[t][8] });
  }

  // 3. Ambil semua data Pengumpulan_Tugas terkait
  var sheetKumpul = ss.getSheetByName("Pengumpulan_Tugas");
  var dataKumpul = sheetKumpul ? sheetKumpul.getDataRange().getValues() : [];

  // 4. Susun tabel: per siswa, per tugas -> nilai
  var hasil = daftarSiswa.map(function(s) {
    var kehadiran = _hitungPersentaseKehadiran(s.username);
    var totalNilai = 0, jumlahDinilai = 0;

    var daftarNilaiTugas = daftarTugas.map(function(tg) {
      var barisKumpul = null;
      for (var k = 1; k < dataKumpul.length; k++) {
        if (dataKumpul[k][0].toString().trim() === tg.idBaris && dataKumpul[k][1].toString().trim() === s.username.toString().trim()) {
          barisKumpul = dataKumpul[k];
          break;
        }
      }
      var nilai;
      if (!barisKumpul) {
        nilai = 0; // Tidak dikumpulkan sama sekali
      } else if (barisKumpul[8] === "" || barisKumpul[8] === null || barisKumpul[8] === undefined) {
        nilai = "Belum Dinilai"; // Sudah kumpul tapi guru belum kasih nilai
      } else {
        nilai = Number(barisKumpul[8]);
      }
      if (typeof nilai === "number") { totalNilai += nilai; jumlahDinilai++; }
      return { judul: tg.judul, nilai: nilai };
    });

    return {
      username: s.username, nama: s.nama,
      persentaseKehadiran: kehadiran.persentase,
      daftarNilaiTugas: daftarNilaiTugas,
      rataRataTugas: jumlahDinilai > 0 ? Math.round(totalNilai / jumlahDinilai) : "-"
    };
  });

  hasil.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
  return { daftarTugas: daftarTugas, daftarSiswa: hasil };
}

/**
 * REKAP SISWA: rekap kehadiran keseluruhan + nilai semua tugas yang sudah tayang untuknya (0 kalau tidak dikerjakan).
 * Nilai ujian/Ulangan Harian masih placeholder karena modul CBT belum dikembangkan.
 */
function siswaAmbilRekapNilaiSaya(usernameSiswa, kelasSiswa) {
  var kehadiran = _hitungPersentaseKehadiran(usernameSiswa);
  var daftarTugas = siswaAmbilTugasKelas(kelasSiswa, usernameSiswa);

  var totalNilai = 0, jumlahDinilai = 0;
  var daftarNilaiTugas = daftarTugas.map(function(t) {
    var nilai;
    if (t.statusKumpul === "Belum Kumpul") {
      nilai = 0;
    } else if (t.nilai === "" || t.nilai === null || t.nilai === undefined) {
      nilai = "Belum Dinilai";
    } else {
      nilai = Number(t.nilai);
    }
    if (typeof nilai === "number") { totalNilai += nilai; jumlahDinilai++; }
    return { judul: t.judul, mapel: t.mapel, nilai: nilai, statusKumpul: t.statusKumpul };
  });

  return {
    kehadiran: kehadiran,
    daftarNilaiTugas: daftarNilaiTugas,
    rataRataTugas: jumlahDinilai > 0 ? Math.round(totalNilai / jumlahDinilai) : "-",
    nilaiUjian: [] // Placeholder - akan diisi setelah modul Bank Soal & CBT (Ulangan Harian) dikembangkan
  };
}

// ==================== MODUL ASISTEN AI (Google Gemini) UNTUK GURU ====================

/**
 * Fungsi inti pemanggil Gemini API. API Key disimpan di Script Properties (Project Settings di editor),
 * BUKAN ditulis langsung di kode, demi keamanan.
 */
function _panggilGeminiAI(prompt) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error("API Key Gemini belum diatur. Silakan atur dulu di Project Settings > Script Properties dengan nama 'GEMINI_API_KEY'.");
  }

  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + apiKey;
  var payload = {
    contents: [{ parts: [{ text: prompt }] }]
  };
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var kode = response.getResponseCode();
  var hasil = JSON.parse(response.getContentText());

  if (kode !== 200) {
    var pesanError = (hasil.error && hasil.error.message) ? hasil.error.message : "Gagal menghubungi layanan AI.";
    throw new Error(pesanError);
  }

  if (!hasil.candidates || hasil.candidates.length === 0) {
    throw new Error("AI tidak memberikan respons. Coba dengan topik/kata kunci yang berbeda.");
  }

  return hasil.candidates[0].content.parts[0].text;
}

/**
 * Asisten Tulis Deskripsi Materi/Tugas: guru cukup masukkan judul/topik singkat,
 * AI akan membuatkan draf deskripsi yang bisa diedit lagi sebelum disimpan.
 */
function aiBantuTulisDeskripsi(judulTopik, jenis, mapel) {
  var instruksi = jenis === "Tugas"
    ? "Buatkan draf instruksi/deskripsi tugas sekolah yang jelas dan singkat (maksimal 4 kalimat) dalam Bahasa Indonesia untuk mata pelajaran " + (mapel || "umum") + ", dengan judul/topik: \"" + judulTopik + "\". Langsung tulis draf deskripsinya saja, tanpa basa-basi pembuka atau penutup."
    : "Buatkan draf ringkasan/deskripsi materi pembelajaran yang jelas dan menarik (maksimal 4 kalimat) dalam Bahasa Indonesia untuk mata pelajaran " + (mapel || "umum") + ", dengan judul/topik: \"" + judulTopik + "\". Langsung tulis draf deskripsinya saja, tanpa basa-basi pembuka atau penutup.";

  var hasil = _panggilGeminiAI(instruksi);
  return { status: "SUCCESS", teks: hasil.trim() };
}

/**
 * Analisis AI untuk Rekap Nilai & Kehadiran satu kelas: memberi ringkasan & rekomendasi
 * kepada guru berdasarkan data yang sudah ada (persentase kehadiran & nilai tugas siswa).
 */
function aiAnalisisRekapKelas(guruUsername, kelas) {
  var rekap = guruAmbilRekapNilaiKelas(guruUsername, kelas);

  if (!rekap.daftarSiswa || rekap.daftarSiswa.length === 0) {
    return { status: "ERROR", message: "Belum ada data rekap untuk dianalisis." };
  }

  var ringkasanData = rekap.daftarSiswa.map(function(s) {
    return s.nama + ": kehadiran " + s.persentaseKehadiran + "%, rata-rata nilai tugas " + s.rataRataTugas;
  }).join("\n");

  var prompt = "Kamu adalah asisten analisis data untuk guru di Indonesia. Berikut data kehadiran dan rata-rata nilai tugas siswa kelas " + kelas + ":\n\n" + ringkasanData +
    "\n\nBuatkan analisis singkat (maksimal 6 kalimat, format poin-poin) dalam Bahasa Indonesia yang mencakup: (1) siswa mana saja yang perlu perhatian khusus karena kehadiran atau nilainya rendah, (2) pola umum yang terlihat di kelas ini, (3) satu saran tindak lanjut praktis untuk guru. Jangan gunakan format markdown tebal/miring, cukup teks polos dengan poin bernomor.";

  var hasil = _panggilGeminiAI(prompt);
  return { status: "SUCCESS", analisis: hasil.trim() };
}
