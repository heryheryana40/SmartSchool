// Code.gs

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Portal Akademik & CBT Sekolah')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Health-check ringan untuk badge "Sistem Online" di header. Sengaja TIDAK menyentuh Sheet
 *  sama sekali supaya responsnya cepat & tidak membebani kuota baca/tulis Google Sheets —
 *  tujuannya cuma membuktikan server Apps Script masih hidup dan bisa dihubungi. */
function cekStatusServer() {
  return { status: "OK", waktuServer: new Date().toISOString() };
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
          username: dataUser[i][0].toString(),
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
      username: data[i][0] !== null && typeof data[i][0] !== "undefined" ? data[i][0].toString() : "",
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
    if(_samaUsername(data[i][0], username)) {
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
  // PENTING: kunci kolom Username (A) sebagai TEKS setelah ditulis — kalau tidak, Google Sheets akan
  // otomatis mengubah username/NISN yang berupa angka semua (mis. "0087703627") jadi tipe Number,
  // sehingga angka 0 di depan hilang ("87703627") dan bikin akun tidak bisa login/data macet loading.
  var barisBaru = sheet.getLastRow();
  sheet.getRange(barisBaru, 1).setNumberFormat("@").setValue(username.trim());
  
  catatLogAktivitas("Admin", "Tambah Akun", "Membuat akun baru " + username + " (" + nama + ") sebagai " + role);
  
  return { status: "SUCCESS", message: "Akun " + role + " atas nama " + nama + " berhasil dibuat!" };
}

// Hapus User oleh Admin
function adminHapusUser(username) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  
  for(var i = data.length - 1; i >= 1; i--) {
    if(_samaUsername(data[i][0], username)) {
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
  
  siswaList.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
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

  // Peta kunci ternormalisasi (huruf besar, tanpa spasi/tanda baca) -> ejaan resmi di master.
  // Dipakai supaya "XII 3" / "xii-3" / "XII - 3" dari file Excel otomatis disamakan penulisannya
  // dengan ejaan resmi di Daftar_Kelas, mencegah munculnya varian tulisan baru yang tercecer.
  var _kunciKelasImport = function(s) { return s.toUpperCase().replace(/[^A-Z0-9]/g, ''); };
  var petaKelasValidByKunci = {};
  daftarKelasValid.forEach(function(k) { petaKelasValidByKunci[_kunciKelasImport(k)] = k; });
  
  // Ambil semua username eksisting untuk validasi duplikasi
  var setUsername = {};
  for (var i = 1; i < dataEksisting.length; i++) {
    setUsername[dataEksisting[i][0].toString().trim()] = true;
  }
  
  var barisBaru = [];
  var jumlahSukses = 0;
  var detailGagal = []; // {baris, username, nama, alasan}
  var peringatanKelas = []; // Kelas yang benar-benar tidak dikenal (tidak menggagalkan import)
  
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
      var kelasCocokMaster = petaKelasValidByKunci[_kunciKelasImport(kelas)];
      if (kelasCocokMaster) {
        // Ejaan di Excel beda tipis (spasi/tanda hubung/huruf) dari master, tapi maksudnya kelas
        // yang sama -> otomatis disamakan ke ejaan resmi supaya tidak jadi varian baru yang tercecer.
        kelas = kelasCocokMaster;
      } else {
        // Benar-benar tidak dikenal di master sama sekali -> tetap diimpor apa adanya, tapi diberi
        // peringatan supaya admin bisa cek/tambahkan ke Daftar_Kelas kalau memang kelas baru.
        peringatanKelas.push({ username: username, nama: nama, kelas: kelas });
      }
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
    if(_samaUsername(data[i][0], username)) {
      // Kolom B adalah Password (indeks ke-1) — dikunci sebagai teks jaga-jaga kalau password berupa angka (mis. PIN berawalan 0)
      sheet.getRange(i + 1, 2).setNumberFormat("@").setValue(passwordBaru.trim());
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

/**
 * Diagnostik (read-only): memindai seluruh data siswa di sheet "Users" dan mengelompokkan
 * berdasarkan teks kelas PERSIS seperti tersimpan, lalu menandai kelompok yang kemungkinan
 * besar sebenarnya kelas YANG SAMA tapi tertulis beda (mis. "XII 3" vs "XII-3") — biasanya
 * akibat ketidakkonsistenan saat import Excel massal. Tidak mengubah data apa pun.
 */
function adminCekKonsistensiKelasSiswa() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  if (!sheet) return { status: "ERROR", message: "Sheet Users tidak ditemukan.", daftar: [], kemungkinanDuplikat: [] };

  var data = sheet.getDataRange().getValues();
  var tally = {}; // teks kelas persis -> jumlah siswa

  for (var i = 1; i < data.length; i++) {
    var role = data[i][3] ? data[i][3].toString().trim() : "";
    if (role !== "Siswa") continue;
    var kelasPersis = (data[i].length > 4 && data[i][4]) ? data[i][4].toString().trim() : "(kosong)";
    tally[kelasPersis] = (tally[kelasPersis] || 0) + 1;
  }

  var daftar = Object.keys(tally).map(function(k) {
    return { kelas: k, jumlah: tally[k] };
  }).sort(function(a, b) { return a.kelas.localeCompare(b.kelas); });

  // Kelompokkan berdasarkan kunci ternormalisasi (huruf besar semua, tanpa spasi/tanda baca)
  // untuk menemukan kelompok yang berisi LEBIH dari satu variasi tulisan berbeda.
  var kunci = function(s) { return s.toUpperCase().replace(/[^A-Z0-9]/g, ''); };
  var kelompok = {};
  daftar.forEach(function(item) {
    var kk = kunci(item.kelas);
    if (!kelompok[kk]) kelompok[kk] = [];
    kelompok[kk].push(item);
  });

  var kemungkinanDuplikat = Object.keys(kelompok)
    .map(function(kk) { return kelompok[kk]; })
    .filter(function(grup) { return grup.length > 1; });

  return { status: "SUCCESS", daftar: daftar, kemungkinanDuplikat: kemungkinanDuplikat };
}

/**
 * Perbaikan massal (menulis data): mengganti SEMUA siswa yang kelasnya tertulis "kelasLama"
 * (persis) menjadi "kelasBaru" yang sudah dipilih admin sebagai penulisan resmi/standar.
 * Dipakai setelah adminCekKonsistensiKelasSiswa menemukan variasi tulisan untuk kelas yang sama.
 */
function adminSamakanPenulisanKelasSiswa(kelasLama, kelasBaru) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  if (!sheet) return { status: "ERROR", message: "Sheet Users tidak ditemukan." };

  var lamaBersih = kelasLama.toString().trim();
  var baruBersih = kelasBaru.toString().trim();
  if (!lamaBersih || !baruBersih) return { status: "ERROR", message: "Nama kelas lama & baru wajib diisi." };
  if (lamaBersih === baruBersih) return { status: "ERROR", message: "Kelas lama dan baru sama, tidak ada yang perlu diubah." };

  var data = sheet.getDataRange().getValues();
  var jumlahDiubah = 0;

  for (var i = 1; i < data.length; i++) {
    var role = data[i][3] ? data[i][3].toString().trim() : "";
    var kelasBaris = (data[i].length > 4 && data[i][4]) ? data[i][4].toString().trim() : "";
    if (role === "Siswa" && kelasBaris === lamaBersih) {
      sheet.getRange(i + 1, 5).setValue(baruBersih);
      jumlahDiubah++;
    }
  }

  if (jumlahDiubah > 0) {
    catatLogAktivitas("Admin", "Samakan Penulisan Kelas", "Mengubah " + jumlahDiubah + " siswa dari kelas \"" + lamaBersih + "\" menjadi \"" + baruBersih + "\"");
  }

  return {
    status: jumlahDiubah > 0 ? "SUCCESS" : "ERROR",
    message: jumlahDiubah > 0
      ? jumlahDiubah + " siswa berhasil diubah dari \"" + lamaBersih + "\" menjadi \"" + baruBersih + "\"."
      : "Tidak ada siswa yang cocok dengan kelas \"" + lamaBersih + "\"."
  };
}

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

/**
 * Folder khusus gambar Bank Soal dibuat otomatis sekali saja (di root Drive akun pemilik script),
 * lalu ID-nya disimpan permanen di Script Properties supaya panggilan berikutnya memakai folder yang sama
 * (tidak membuat folder baru berulang-ulang).
 */
function _pastikanFolderGambarSoal() {
  var props = PropertiesService.getScriptProperties();
  var idTersimpan = props.getProperty("FOLDER_ID_GAMBAR_SOAL");
  if (idTersimpan) {
    try {
      var folderCek = DriveApp.getFolderById(idTersimpan); // pastikan folder masih ada & belum dihapus manual
      return folderCek.getId();
    } catch (errCek) {
      // Folder tersimpan sudah tidak valid (mis. dihapus manual) — buat ulang di bawah ini
    }
  }
  var folderBaru = DriveApp.createFolder("SmartSchool - Gambar Bank Soal");
  props.setProperty("FOLDER_ID_GAMBAR_SOAL", folderBaru.getId());
  return folderBaru.getId();
}

function uploadFileKeDrive(base64Data, namaFile, mimeType, jenisUpload) {
  try {
    var idFolder = (jenisUpload === "Tugas") ? FOLDER_ID_TUGAS : (jenisUpload === "Soal") ? _pastikanFolderGambarSoal() : FOLDER_ID_MATERI;
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
    
    return { status: "SUCCESS", url: file.getUrl(), urlGambar: "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w2000", namaFile: namaFile };
  } catch (err) {
    return { status: "ERROR", message: "Gagal mengunggah file: " + err.message };
  }
}

var _idUnikCounter = 0;
function _buatIdUnik(prefix) {
  _idUnikCounter = (_idUnikCounter + 1) % 1000000;
  return prefix + "-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000) + "-" + _idUnikCounter;
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
  semuaSiswaTarget.sort(function(a, b) { return a.nama.localeCompare(b.nama); });

  var sheetKumpul = _pastikanSheetPengumpulanTugasBenar();
  var dataKumpul = sheetKumpul.getDataRange().getValues();

  // Ambil daftar kelompok tugas ini & buat peta siswa->kelompok LEBIH DULU (dari data master Kelompok),
  // supaya siswa yang BELUM mengumpulkan tugas tetap muncul di bawah kelompoknya masing-masing,
  // bukan tercampur ke "Belum masuk kelompok" hanya karena belum ada baris submission.
  var kelompokList = [];
  var petaSiswaKelompok = {}; // username -> { idKelompok, namaKelompok }
  if (jenisTugas === "Kelompok") {
    kelompokList = guruAmbilKelompokTugas(idTugas);
    kelompokList.forEach(function(klp) {
      (klp.anggota || []).forEach(function(a) {
        petaSiswaKelompok[a.username.toString().trim()] = { idKelompok: klp.idKelompok, namaKelompok: klp.namaKelompok };
      });
    });
  }

  var hasil = semuaSiswaTarget.map(function(s) {
    var infoKelompok = petaSiswaKelompok[s.username.toString().trim()] || null;
    var record = {
      username: s.username, nama: s.nama, kelas: s.kelas, statusKumpul: "Belum Kumpul", waktuKumpul: "", jenisJawaban: "", isiJawaban: "",
      nilai: "", catatanGuru: "",
      idKelompok: infoKelompok ? infoKelompok.idKelompok : "",
      namaKelompok: infoKelompok ? infoKelompok.namaKelompok : "",
      usernamePengirim: "", laporanKontribusi: null, riwayatSubmission: []
    };
    for (var j = 1; j < dataKumpul.length; j++) {
      if (dataKumpul[j][0].toString().trim() === idTugas.toString().trim() && dataKumpul[j][1].toString().trim() === s.username.toString().trim()) {
        record.statusKumpul = dataKumpul[j][7];
        record.waktuKumpul = dataKumpul[j][4];
        record.jenisJawaban = dataKumpul[j][5];
        record.isiJawaban = dataKumpul[j][6];
        record.nilai = dataKumpul[j][8];
        record.catatanGuru = dataKumpul[j][9];
        // Utamakan idKelompok dari data kelompok master (peta) di atas. Jika siswa ini tidak
        // ditemukan di peta (mis. kelompok sudah dihapus/diubah setelah dia submit), pakai
        // idKelompok yang sempat tercatat saat submit sebagai fallback agar histori tidak hilang.
        if (!record.idKelompok) record.idKelompok = dataKumpul[j][10] || "";
        record.usernamePengirim = dataKumpul[j][11] || "";
        try { record.laporanKontribusi = dataKumpul[j][12] ? JSON.parse(dataKumpul[j][12]) : null; } catch (e) { record.laporanKontribusi = null; }
        try { record.riwayatSubmission = dataKumpul[j][13] ? JSON.parse(dataKumpul[j][13]) : []; } catch (e) { record.riwayatSubmission = []; }
        break;
      }
    }
    return record;
  });

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

// ==================== MODUL JURNAL MENGAJAR (GURU) ====================

var HEADER_JURNAL_MENGAJAR = ["ID", "Guru Username", "Guru Nama", "Tanggal", "Jam Ke", "Kelas", "Semester", "Mapel", "Materi Pembelajaran", "Tujuan Indikator", "Pencapaian Hasil", "Absensi Siswa", "Catatan Refleksi", "Waktu Disimpan"];

function _pastikanSheetJurnalMengajarBenar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Jurnal_Mengajar");
  if (!sheet) {
    sheet = ss.insertSheet("Jurnal_Mengajar");
    sheet.appendRow(HEADER_JURNAL_MENGAJAR);
    return sheet;
  }
  var headerSaatIni = sheet.getRange(1, 1, 1, HEADER_JURNAL_MENGAJAR.length).getValues()[0];
  var headerCocok = HEADER_JURNAL_MENGAJAR.every(function(h, idx) { return headerSaatIni[idx] === h; });
  if (!headerCocok) sheet.getRange(1, 1, 1, HEADER_JURNAL_MENGAJAR.length).setValues([HEADER_JURNAL_MENGAJAR]);
  return sheet;
}

/**
 * Simpan satu entri Jurnal Mengajar baru.
 * data: { tanggal, jamKe, kelas, semester, materi, tujuan, pencapaian, absensi, refleksi }
 */
function guruSimpanJurnalMengajar(guruUsername, guruNama, mapel, data) {
  var sheet = _pastikanSheetJurnalMengajarBenar();
  var id = _buatIdUnik("JRN");
  var waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  sheet.appendRow([id, guruUsername, guruNama, data.tanggal, data.jamKe || "", data.kelas, data.semester || "", mapel || "", data.materi, data.tujuan || "", data.pencapaian || "", data.absensi || "", data.refleksi || "", waktu]);
  var barisTerakhir = sheet.getLastRow();
  sheet.getRange(barisTerakhir, 4).setNumberFormat("@").setValue(data.tanggal); // Kunci Tanggal sebagai Teks

  catatLogAktivitas(guruNama, "Isi Jurnal Mengajar", "Mengisi jurnal mengajar kelas " + data.kelas + " (" + data.tanggal + "): " + data.materi);
  return { status: "SUCCESS", message: "Jurnal mengajar berhasil disimpan." };
}

/** Daftar jurnal mengajar milik guru, terbaru di atas */
function guruAmbilJurnalMengajar(guruUsername) {
  var sheet = _pastikanSheetJurnalMengajarBenar();
  var data = sheet.getDataRange().getValues();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() !== guruUsername.toString().trim()) continue;
    hasil.push({
      id: data[i][0], tanggal: data[i][3].toString().trim(), jamKe: data[i][4], kelas: data[i][5], semester: data[i][6],
      mapel: data[i][7], materi: data[i][8], tujuan: data[i][9], pencapaian: data[i][10], absensi: data[i][11], refleksi: data[i][12]
    });
  }
  hasil.sort(function(a, b) { return a.tanggal < b.tanggal ? 1 : (a.tanggal > b.tanggal ? -1 : 0); });
  return hasil;
}

function guruUpdateJurnalMengajar(id, data) {
  var sheet = _pastikanSheetJurnalMengajarBenar();
  var dataSheet = sheet.getDataRange().getValues();
  for (var i = 1; i < dataSheet.length; i++) {
    if (dataSheet[i][0].toString().trim() === id.toString().trim()) {
      sheet.getRange(i + 1, 4).setNumberFormat("@").setValue(data.tanggal);
      sheet.getRange(i + 1, 5).setValue(data.jamKe || "");
      sheet.getRange(i + 1, 9).setValue(data.materi);
      sheet.getRange(i + 1, 10).setValue(data.tujuan || "");
      sheet.getRange(i + 1, 11).setValue(data.pencapaian || "");
      sheet.getRange(i + 1, 12).setValue(data.absensi || "");
      sheet.getRange(i + 1, 13).setValue(data.refleksi || "");
      return { status: "SUCCESS", message: "Jurnal berhasil diperbarui." };
    }
  }
  return { status: "ERROR", message: "Data jurnal tidak ditemukan." };
}

function guruHapusJurnalMengajar(id) {
  var sheet = _pastikanSheetJurnalMengajarBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0].toString().trim() === id.toString().trim()) {
      sheet.deleteRow(i + 1);
      return { status: "SUCCESS", message: "Jurnal berhasil dihapus." };
    }
  }
  return { status: "ERROR", message: "Data jurnal tidak ditemukan." };
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


var HEADER_JADWAL_MENGAJAR = ["ID", "Guru Username", "Guru Nama", "Hari", "Jam Ke", "Mapel", "Kelas", "Jam Mulai", "Jam Selesai"];

function _pastikanSheetJadwalMengajarBenar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Jadwal_Mengajar");
  if (!sheet) {
    sheet = ss.insertSheet("Jadwal_Mengajar");
    sheet.appendRow(HEADER_JADWAL_MENGAJAR);
    return sheet;
  }
  var headerSaatIni = sheet.getRange(1, 1, 1, HEADER_JADWAL_MENGAJAR.length).getValues()[0];
  if (!headerSaatIni[7]) sheet.getRange(1, 8).setValue("Jam Mulai");
  if (!headerSaatIni[8]) sheet.getRange(1, 9).setValue("Jam Selesai");
  return sheet;
}

/**
 * Menyimpan jadwal baru.
 * jamMulai/jamSelesai: waktu manual (mis. "07:15"), diisi guru sendiri karena bisa
 * berubah-ubah (misal Senin ada upacara sehingga jam pertama mundur).
 */
function guruSimpanJadwalMengajar(
  usernameGuru,
  namaGuru,
  hari,
  jamKe,
  mapel,
  kelas,
  jamMulai,
  jamSelesai
) {
  var sheet = _pastikanSheetJadwalMengajarBenar();

  var usernameBersih = usernameGuru.toString().trim();
  var namaBersih = namaGuru.toString().trim();
  var hariBersih = hari.toString().trim();
  var jamKeBersih = normalisasiJamKeJadwal_(jamKe);
  var mapelBersih = mapel ? mapel.toString().trim() : "";
  var kelasBersih = kelas.toString().trim();
  var jamMulaiBersih = jamMulai ? jamMulai.toString().trim() : "";
  var jamSelesaiBersih = jamSelesai ? jamSelesai.toString().trim() : "";

  if (!usernameBersih || !hariBersih || !jamKeBersih || !kelasBersih) {
    return {
      status: "ERROR",
      message: "Hari, jam ke, dan kelas wajib diisi."
    };
  }

  var id = _buatIdUnik("JDW");
  var barisBaru = sheet.getLastRow() + 1;

  // Kolom E, H, I dikunci sebagai teks sebelum data dimasukkan (agar "07:15" tidak diubah jadi jam Date oleh Sheets).
  sheet.getRange(barisBaru, 5).setNumberFormat("@");
  sheet.getRange(barisBaru, 8).setNumberFormat("@");
  sheet.getRange(barisBaru, 9).setNumberFormat("@");

  sheet.getRange(barisBaru, 1, 1, 9).setValues([[
    id,
    usernameBersih,
    namaBersih,
    hariBersih,
    jamKeBersih,
    mapelBersih,
    kelasBersih,
    jamMulaiBersih,
    jamSelesaiBersih
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
 * Mengubah (edit) jadwal yang sudah ada berdasarkan ID.
 */
function guruEditJadwalMengajar(id, hari, jamKe, mapel, kelas, jamMulai, jamSelesai) {
  var sheet = _pastikanSheetJadwalMengajarBenar();
  var data = sheet.getDataRange().getValues();
  var idBersih = id.toString().trim();
  var jamKeBersih = normalisasiJamKeJadwal_(jamKe);
  var kelasBersih = kelas ? kelas.toString().trim() : "";
  var hariBersih = hari ? hari.toString().trim() : "";
  var mapelBersih = mapel ? mapel.toString().trim() : "";
  var jamMulaiBersih = jamMulai ? jamMulai.toString().trim() : "";
  var jamSelesaiBersih = jamSelesai ? jamSelesai.toString().trim() : "";

  if (!hariBersih || !jamKeBersih || !kelasBersih) {
    return { status: "ERROR", message: "Hari, jam ke, dan kelas wajib diisi." };
  }

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === idBersih) {
      var baris = i + 1;
      sheet.getRange(baris, 5).setNumberFormat("@");
      sheet.getRange(baris, 8).setNumberFormat("@");
      sheet.getRange(baris, 9).setNumberFormat("@");
      sheet.getRange(baris, 4, 1, 6).setValues([[hariBersih, jamKeBersih, mapelBersih, kelasBersih, jamMulaiBersih, jamSelesaiBersih]]);
      return {
        status: "SUCCESS",
        message: "Jadwal " + hariBersih + " jam ke-" + jamKeBersih + " (Kelas " + kelasBersih + ") berhasil diperbarui."
      };
    }
  }
  return { status: "ERROR", message: "Jadwal tidak ditemukan." };
}


/**
 * Mengambil seluruh jadwal guru.
 * Seluruh nilai dikonversi ke tipe data yang aman dikirim ke frontend.
 */
function guruAmbilJadwalMengajar(usernameGuru) {
  var sheet = _pastikanSheetJadwalMengajarBenar();
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
        kelas: data[i][6] ? data[i][6].toString().trim() : "",
        jamMulai: data[i][7] ? data[i][7].toString().trim() : "",
        jamSelesai: data[i][8] ? data[i][8].toString().trim() : ""
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
        mapel: jadwal.mapel,
        jamMulai: jadwal.jamMulai || "",
        jamSelesai: jadwal.jamSelesai || ""
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
      jamMulai: dataJadwal.jamMulai || "",
      jamSelesai: dataJadwal.jamSelesai || "",
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
/** Daftar siswa 1 kelas (dipakai untuk form Import Nilai Ujian Semester) */
function guruAmbilDaftarSiswaKelas(kelas) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUser = ss.getSheetByName("Users");
  var dataUser = sheetUser.getDataRange().getValues();
  var hasil = [];
  for (var u = 1; u < dataUser.length; u++) {
    if (dataUser[u][3].toString().trim() === "Siswa" && dataUser[u].length > 4 && dataUser[u][4].toString().trim().toUpperCase() === kelas.toString().trim().toUpperCase()) {
      hasil.push({ username: dataUser[u][0].toString(), nama: dataUser[u][2].toString() });
    }
  }
  hasil.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
  return hasil;
}

// ==================== MODUL BOBOT NILAI & NILAI UJIAN SEMESTER (RAPOR) ====================

var HEADER_BOBOT_NILAI = ["Guru Username", "Mapel", "Bobot Tugas", "Bobot UTS", "Bobot UAS", "Bobot CBT"];

function _pastikanSheetBobotNilaiBenar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Bobot_Nilai_Guru");
  if (!sheet) {
    sheet = ss.insertSheet("Bobot_Nilai_Guru");
    sheet.appendRow(HEADER_BOBOT_NILAI);
    return sheet;
  }
  var headerSaatIni = sheet.getRange(1, 1, 1, HEADER_BOBOT_NILAI.length).getValues()[0];
  if (!headerSaatIni[5]) sheet.getRange(1, 6).setValue("Bobot CBT");
  return sheet;
}

/** Simpan pengaturan bobot (harus total 100). Satu baris per guru+mapel (upsert). */
function guruSimpanBobotNilai(guruUsername, mapel, bobotTugas, bobotUTS, bobotUAS, bobotCBT) {
  var total = Number(bobotTugas) + Number(bobotUTS) + Number(bobotUAS) + Number(bobotCBT || 0);
  if (total !== 100) {
    return { status: "ERROR", message: "Total bobot harus 100% (saat ini " + total + "%)." };
  }
  var sheet = _pastikanSheetBobotNilaiBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === guruUsername.toString().trim() && data[i][1].toString().trim() === (mapel || "").toString().trim()) {
      sheet.getRange(i + 1, 3, 1, 4).setValues([[bobotTugas, bobotUTS, bobotUAS, bobotCBT || 0]]);
      return { status: "SUCCESS", message: "Bobot nilai berhasil disimpan." };
    }
  }
  sheet.appendRow([guruUsername, mapel || "", bobotTugas, bobotUTS, bobotUAS, bobotCBT || 0]);
  return { status: "SUCCESS", message: "Bobot nilai berhasil disimpan." };
}

/** Ambil bobot nilai guru untuk 1 mapel. Default: 100% Tugas, sisanya 0% (kalau belum pernah diatur). */
function guruAmbilBobotNilai(guruUsername, mapel) {
  var sheet = _pastikanSheetBobotNilaiBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === guruUsername.toString().trim() && data[i][1].toString().trim() === (mapel || "").toString().trim()) {
      return { bobotTugas: Number(data[i][2]), bobotUTS: Number(data[i][3]), bobotUAS: Number(data[i][4]), bobotCBT: Number(data[i][5]) || 0 };
    }
  }
  return { bobotTugas: 100, bobotUTS: 0, bobotUAS: 0, bobotCBT: 0 };
}

var HEADER_NILAI_UJIAN_SEMESTER = ["ID", "Guru Username", "Kelas", "Mapel", "Semester", "Jenis Ujian", "Username Siswa", "Nama Siswa", "Nilai", "Waktu Impor"];

function _pastikanSheetNilaiUjianSemesterBenar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Nilai_Ujian_Semester");
  if (!sheet) {
    sheet = ss.insertSheet("Nilai_Ujian_Semester");
    sheet.appendRow(HEADER_NILAI_UJIAN_SEMESTER);
  }
  return sheet;
}

/**
 * Simpan/impor nilai UTS atau UAS untuk satu kelas sekaligus (upsert per siswa).
 * daftarNilai: [{ username, nama, nilai }, ...]
 */
function guruSimpanNilaiUjianSemester(guruUsername, kelas, mapel, semester, jenisUjian, daftarNilai) {
  var sheet = _pastikanSheetNilaiUjianSemesterBenar();
  var data = sheet.getDataRange().getValues();
  var waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var jumlahDisimpan = 0;

  daftarNilai.forEach(function(item) {
    if (item.nilai === "" || item.nilai === null || typeof item.nilai === "undefined") return; // lewati yang kosong
    var ditemukan = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][1].toString().trim() === guruUsername.toString().trim() &&
          data[i][2].toString().trim().toUpperCase() === kelas.toString().trim().toUpperCase() &&
          data[i][3].toString().trim() === (mapel || "").toString().trim() &&
          data[i][4].toString().trim() === semester.toString().trim() &&
          data[i][5].toString().trim() === jenisUjian.toString().trim() &&
          data[i][6].toString().trim() === item.username.toString().trim()) {
        sheet.getRange(i + 1, 9).setValue(item.nilai);
        sheet.getRange(i + 1, 10).setValue(waktu);
        ditemukan = true;
        break;
      }
    }
    if (!ditemukan) {
      sheet.appendRow([_buatIdUnik("NUS"), guruUsername, kelas, mapel || "", semester, jenisUjian, item.username, item.nama || "", item.nilai, waktu]);
    }
    jumlahDisimpan++;
  });

  return { status: "SUCCESS", message: "Nilai " + jenisUjian + " berhasil disimpan untuk " + jumlahDisimpan + " siswa." };
}

/** Ambil nilai UTS/UAS yang sudah tersimpan untuk 1 kelas (dipakai untuk prefill form import & rekap). */
function guruAmbilNilaiUjianSemester(guruUsername, kelas, mapel, semester, jenisUjian) {
  var sheet = _pastikanSheetNilaiUjianSemesterBenar();
  var data = sheet.getDataRange().getValues();
  var peta = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() === guruUsername.toString().trim() &&
        data[i][2].toString().trim().toUpperCase() === kelas.toString().trim().toUpperCase() &&
        data[i][3].toString().trim() === (mapel || "").toString().trim() &&
        data[i][4].toString().trim() === semester.toString().trim() &&
        data[i][5].toString().trim() === jenisUjian.toString().trim()) {
      peta[data[i][6].toString().trim()] = data[i][8];
    }
  }
  return peta;
}

/**
 * Rata-rata Nilai Akhir CBT per siswa, HANYA dari sesi ujian milik guru ini di kelas+mapel terkait
 * yang ditandai "Sertakan Rekap" = TRUE, dan HANYA hitungan siswa yang statusnya "Selesai"
 * (Sedang Mengerjakan/Terkunci belum final, tidak ikut dihitung dulu).
 */
function _dataCBTUntukRekap(guruUsername, kelas, mapel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSesi = _pastikanSheetCbtSesiBenar();
  var dataSesi = sheetSesi.getDataRange().getValues();
  var daftarSesi = []; // [{idSesi, judul}] terurut sesuai tanggal dibuat

  for (var i = 1; i < dataSesi.length; i++) {
    if (dataSesi[i][1].toString().trim() !== guruUsername.toString().trim()) continue;
    if (dataSesi[i][5].toString().trim() !== (mapel || "").toString().trim()) continue;
    var daftarKelasSesi = dataSesi[i][6].toString().split(",").map(function(k) { return k.trim().toUpperCase(); });
    if (daftarKelasSesi.indexOf(kelas.toString().trim().toUpperCase()) === -1) continue;
    var sertakan = dataSesi[i][20] === true || dataSesi[i][20] === undefined || dataSesi[i][20].toString().toUpperCase() !== "FALSE";
    if (!sertakan) continue;
    daftarSesi.push({ idSesi: dataSesi[i][0].toString(), judul: dataSesi[i][3].toString() });
  }

  if (daftarSesi.length === 0) return { daftarSesi: [], nilaiPerSesi: {}, rataRata: {} };

  var idSesiValid = {};
  daftarSesi.forEach(function(s) { idSesiValid[s.idSesi] = true; });

  var sheetHasil = _pastikanSheetHasilUjianBenar();
  var dataHasil = sheetHasil.getDataRange().getValues();
  var nilaiPerSesi = {}; // { usernameKey: { idSesi: nilai } }
  var totalPerSiswa = {}, jumlahPerSiswa = {};

  for (var h = 1; h < dataHasil.length; h++) {
    var idSesiBaris = dataHasil[h][1].toString();
    if (!idSesiValid[idSesiBaris]) continue;
    if (dataHasil[h][10].toString() !== "Selesai") continue;
    if (dataHasil[h][8].toString() === "Menunggu") continue; // Soal Essay belum dikoreksi guru — jangan masuk rekap dulu, nilainya belum final
    var key = _normalisasiUsername(dataHasil[h][2]);
    var nilai = Number(dataHasil[h][9]) || 0;
    if (!nilaiPerSesi[key]) nilaiPerSesi[key] = {};
    nilaiPerSesi[key][idSesiBaris] = nilai;
    totalPerSiswa[key] = (totalPerSiswa[key] || 0) + nilai;
    jumlahPerSiswa[key] = (jumlahPerSiswa[key] || 0) + 1;
  }

  var rataRata = {};
  Object.keys(totalPerSiswa).forEach(function(key) {
    rataRata[key] = Math.round(totalPerSiswa[key] / jumlahPerSiswa[key]);
  });

  return { daftarSesi: daftarSesi, nilaiPerSesi: nilaiPerSesi, rataRata: rataRata };
}

function guruAmbilRekapNilaiKelas(guruUsername, kelas, semester) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var semesterAktif = semester || "Ganjil";

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
  var mapel = "";

  for (var t = 1; t < dataTugas.length; t++) {
    if (dataTugas[t][3].toString().trim() !== guruUsername.toString().trim()) continue;
    if (dataTugas[t][6].toString().trim().toUpperCase() !== kelas.toString().trim().toUpperCase()) continue;
    if (dataTugas[t][7].toString().trim() > tglHariIni) continue; // Belum tayang, jangan dimasukkan rekap dulu
    daftarTugas.push({ idBaris: dataTugas[t][0].toString().trim(), judul: dataTugas[t][8] });
    if (!mapel) mapel = dataTugas[t][5];
  }
  if (!mapel) {
    var plotting = ambilKelasAmpoanGuru(guruUsername);
    mapel = plotting.mataPelajaran || "";
  }

  // 3. Ambil semua data Pengumpulan_Tugas terkait
  var sheetKumpul = ss.getSheetByName("Pengumpulan_Tugas");
  var dataKumpul = sheetKumpul ? sheetKumpul.getDataRange().getValues() : [];

  // 4. Ambil bobot nilai & nilai UTS/UAS yang sudah diimpor
  var bobot = guruAmbilBobotNilai(guruUsername, mapel);
  var petaUTS = guruAmbilNilaiUjianSemester(guruUsername, kelas, mapel, semesterAktif, "UTS");
  var petaUAS = guruAmbilNilaiUjianSemester(guruUsername, kelas, mapel, semesterAktif, "UAS");
  var dataCBT = _dataCBTUntukRekap(guruUsername, kelas, mapel);

  // 5. Susun tabel: per siswa, per tugas -> nilai, lalu hitung Nilai Akhir/Rapor berbobot
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

    var keyCBT = _normalisasiUsername(s.username);
    var daftarNilaiCBT = dataCBT.daftarSesi.map(function(sesi) {
      var nilaiSesiIni = (dataCBT.nilaiPerSesi[keyCBT] && typeof dataCBT.nilaiPerSesi[keyCBT][sesi.idSesi] !== "undefined")
        ? dataCBT.nilaiPerSesi[keyCBT][sesi.idSesi] : "-"; // "-" = belum mengerjakan/belum selesai
      return { judul: sesi.judul, nilai: nilaiSesiIni };
    });

    var rataRataTugas = jumlahDinilai > 0 ? Math.round(totalNilai / jumlahDinilai) : "-";
    var nilaiUTS = (typeof petaUTS[s.username] !== "undefined") ? Number(petaUTS[s.username]) : "-";
    var nilaiUAS = (typeof petaUAS[s.username] !== "undefined") ? Number(petaUAS[s.username]) : "-";
    var nilaiCBT = (typeof dataCBT.rataRata[keyCBT] !== "undefined") ? dataCBT.rataRata[keyCBT] : "-";

    // Nilai Akhir/Rapor: hanya dihitung kalau semua komponen yang bobotnya > 0 sudah lengkap datanya
    var nilaiAkhir = "-";
    var lengkap = true;
    if (bobot.bobotTugas > 0 && rataRataTugas === "-") lengkap = false;
    if (bobot.bobotUTS > 0 && nilaiUTS === "-") lengkap = false;
    if (bobot.bobotUAS > 0 && nilaiUAS === "-") lengkap = false;
    if (bobot.bobotCBT > 0 && nilaiCBT === "-") lengkap = false;
    if (lengkap) {
      var totalBerbobot =
        (bobot.bobotTugas > 0 ? (rataRataTugas * bobot.bobotTugas / 100) : 0) +
        (bobot.bobotUTS > 0 ? (nilaiUTS * bobot.bobotUTS / 100) : 0) +
        (bobot.bobotUAS > 0 ? (nilaiUAS * bobot.bobotUAS / 100) : 0) +
        (bobot.bobotCBT > 0 ? (nilaiCBT * bobot.bobotCBT / 100) : 0);
      nilaiAkhir = Math.round(totalBerbobot);
    }

    return {
      username: s.username, nama: s.nama,
      persentaseKehadiran: kehadiran.persentase,
      daftarNilaiTugas: daftarNilaiTugas,
      rataRataTugas: rataRataTugas,
      daftarNilaiCBT: daftarNilaiCBT,
      nilaiUTS: nilaiUTS, nilaiUAS: nilaiUAS, nilaiCBT: nilaiCBT,
      nilaiAkhir: nilaiAkhir
    };
  });

  hasil.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
  return { daftarTugas: daftarTugas, daftarSesiCBT: dataCBT.daftarSesi, daftarSiswa: hasil, bobot: bobot, mapel: mapel, semester: semesterAktif };
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
    nilaiUjian: _siswaAmbilNilaiUjianSaya(usernameSiswa, kelasSiswa)
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

  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent?key=" + apiKey;
  var payload = {
    contents: [{ parts: [{ text: prompt }] }]
  };
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var MAKS_PERCOBAAN = 3;
  var pesanErrorTerakhir = "Gagal menghubungi layanan AI.";

  for (var percobaan = 1; percobaan <= MAKS_PERCOBAAN; percobaan++) {
    var response = UrlFetchApp.fetch(url, options);
    var kode = response.getResponseCode();
    var hasil = JSON.parse(response.getContentText());

    if (kode === 200) {
      if (!hasil.candidates || hasil.candidates.length === 0) {
        throw new Error("AI tidak memberikan respons. Coba dengan topik/kata kunci yang berbeda.");
      }
      return hasil.candidates[0].content.parts[0].text;
    }

    pesanErrorTerakhir = (hasil.error && hasil.error.message) ? hasil.error.message : "Gagal menghubungi layanan AI.";
    // Kode 503/429 (server sibuk/rate limit) bersifat sementara -> layak dicoba ulang. Selain itu, langsung berhenti.
    var bolehCobaLagi = (kode === 503 || kode === 429) && percobaan < MAKS_PERCOBAAN;
    if (!bolehCobaLagi) {
      throw new Error(pesanErrorTerakhir);
    }
    Utilities.sleep(1500 * percobaan); // jeda makin lama tiap percobaan (1.5s, 3s, ...)
  }

  throw new Error(pesanErrorTerakhir);
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

// ==================== MODUL BANK SOAL (GURU) ====================
// Gudang soal milik guru: tidak terikat waktu ujian, bisa dipakai ulang di banyak sesi CBT berbeda.

var HEADER_BANK_SOAL = ["ID Soal", "Guru Username", "Guru Nama", "Mapel", "Kelas", "Topik/Bab", "Tipe Soal", "Pertanyaan", "Opsi A", "Opsi B", "Opsi C", "Opsi D", "Opsi E", "Kunci Jawaban", "Poin", "Level Kognitif", "Timestamp", "Gambar URL"];

function _pastikanSheetBankSoalBenar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Bank_Soal");
  if (!sheet) {
    sheet = ss.insertSheet("Bank_Soal");
    sheet.appendRow(HEADER_BANK_SOAL);
    return sheet;
  }
  var headerSaatIni = sheet.getRange(1, 1, 1, HEADER_BANK_SOAL.length).getValues()[0];
  var headerCocok = HEADER_BANK_SOAL.every(function(h, idx) { return headerSaatIni[idx] === h; });
  if (!headerCocok) {
    sheet.getRange(1, 1, 1, HEADER_BANK_SOAL.length).setValues([HEADER_BANK_SOAL]);
  }
  return sheet;
}

function _baris2Soal(row) {
  return {
    id: row[0].toString(),
    guruUsername: row[1].toString(),
    guruNama: row[2].toString(),
    mapel: row[3].toString(),
    kelas: row[4].toString(),
    topik: row[5].toString(),
    tipe: row[6].toString(),
    pertanyaan: row[7].toString(),
    opsi: [row[8], row[9], row[10], row[11], row[12]].map(function(o) { return o === "" || o === null || o === undefined ? "" : o.toString(); }),
    kunci: row[13].toString(),
    poin: Number(row[14]) || 0,
    levelKognitif: row[15] ? row[15].toString() : "",
    timestamp: row[16] ? row[16].toString() : "",
    gambarUrl: row[17] ? row[17].toString() : ""
  };
}

/** Simpan soal baru, atau perbarui soal lama jika data.idEdit terisi */
/**
 * Import banyak soal sekaligus (dari file Excel yang sudah diparsing di client).
 * daftarSoal: [{ tipe, kelas, topik, pertanyaan, opsi:[A,B,C,D,E], kunci, poin, levelKognitif }, ...]
 */
function guruImportSoalMassal(guruUsername, guruNama, mapel, daftarSoal) {
  var sheet = _pastikanSheetBankSoalBenar();
  var waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var barisBaru = daftarSoal.map(function(data, idx) {
    var opsi = data.opsi || [];
    return [
      _buatIdUnik("SOAL") + "-" + idx, guruUsername, guruNama, mapel || "", data.kelas || "", data.topik || "",
      data.tipe, data.pertanyaan,
      opsi[0] || "", opsi[1] || "", opsi[2] || "", opsi[3] || "", opsi[4] || "",
      (data.kunci || "").toString(), Number(data.poin) || 0, data.levelKognitif || "", waktu, data.gambarUrl || ""
    ];
  });

  if (barisBaru.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, barisBaru.length, barisBaru[0].length).setValues(barisBaru);
  }

  catatLogAktivitas(guruNama, "Import Soal", "Mengimport " + barisBaru.length + " soal ke Bank Soal (mapel " + mapel + ")");
  return { status: "SUCCESS", message: barisBaru.length + " soal berhasil diimport ke Bank Soal." };
}

function guruSimpanSoal(data) {
  var sheet = _pastikanSheetBankSoalBenar();
  var opsi = data.opsi || [];
  var waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var baris = [
    "", data.guruUsername, data.guruNama, data.mapel, data.kelas, data.topik || "",
    data.tipe, data.pertanyaan,
    opsi[0] || "", opsi[1] || "", opsi[2] || "", opsi[3] || "", opsi[4] || "",
    (data.kunci || "").toString(), Number(data.poin) || 0, data.levelKognitif || "", waktu, data.gambarUrl || ""
  ];

  if (data.idEdit) {
    var dataSheet = sheet.getDataRange().getValues();
    for (var i = 1; i < dataSheet.length; i++) {
      if (dataSheet[i][0].toString() === data.idEdit.toString()) {
        baris[0] = data.idEdit;
        baris[16] = dataSheet[i][16]; // Pertahankan timestamp pembuatan asli
        sheet.getRange(i + 1, 1, 1, baris.length).setValues([baris]);
        catatLogAktivitas(data.guruNama, "Edit Soal", "Memperbarui soal " + data.tipe + " mapel " + data.mapel);
        return { status: "SUCCESS", message: "Soal berhasil diperbarui.", id: data.idEdit };
      }
    }
    return { status: "ERROR", message: "Soal tidak ditemukan untuk diperbarui." };
  }

  var idBaru = _buatIdUnik("SOAL");
  baris[0] = idBaru;
  sheet.appendRow(baris);
  catatLogAktivitas(data.guruNama, "Tambah Soal", "Menambahkan soal " + data.tipe + " mapel " + data.mapel + " (" + data.kelas + ")");
  return { status: "SUCCESS", message: "Soal baru berhasil disimpan ke Bank Soal.", id: idBaru };
}

/** Daftar soal milik guru, bisa difilter per mapel/kelas/topik/tipe/kata kunci pencarian */
function guruAmbilDaftarSoal(guruUsername, filter) {
  var sheet = _pastikanSheetBankSoalBenar();
  var data = sheet.getDataRange().getValues();
  filter = filter || {};
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (data[i][1].toString().trim() !== guruUsername.toString().trim()) continue;
    var soal = _baris2Soal(data[i]);
    if (filter.mapel && soal.mapel !== filter.mapel) continue;
    if (filter.kelas && soal.kelas !== filter.kelas) continue;
    if (filter.topik && soal.topik !== filter.topik) continue;
    if (filter.tipe && soal.tipe !== filter.tipe) continue;
    if (filter.cari) {
      var kunciCari = filter.cari.toString().toLowerCase();
      if (soal.pertanyaan.toLowerCase().indexOf(kunciCari) === -1 && soal.topik.toLowerCase().indexOf(kunciCari) === -1) continue;
    }
    hasil.push(soal);
  }
  hasil.sort(function(a, b) { return b.timestamp.localeCompare(a.timestamp); });
  return hasil;
}

function guruAmbilSatuSoal(idSoal) {
  var sheet = _pastikanSheetBankSoalBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === idSoal.toString()) return _baris2Soal(data[i]);
  }
  return null;
}

function guruHapusSoal(idSoal) {
  var sheet = _pastikanSheetBankSoalBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0].toString() === idSoal.toString()) {
      sheet.deleteRow(i + 1);
      return { status: "SUCCESS", message: "Soal berhasil dihapus dari Bank Soal." };
    }
  }
  return { status: "ERROR", message: "Soal tidak ditemukan." };
}

function guruAmbilDaftarTopikSoal(guruUsername, mapel, kelas) {
  var daftar = guruAmbilDaftarSoal(guruUsername, { mapel: mapel, kelas: kelas });
  var set = {};
  daftar.forEach(function(s) { if (s.topik) set[s.topik] = true; });
  return Object.keys(set).sort();
}

function guruStatistikBankSoal(guruUsername) {
  var daftar = guruAmbilDaftarSoal(guruUsername, {});
  var setTopik = {};
  daftar.forEach(function(s) { if (s.topik) setTopik[s.topik] = true; });
  return { totalSoal: daftar.length, totalTopik: Object.keys(setTopik).length };
}

/** Asisten AI: membuatkan draf SATU soal (PG/PGK) siap edit, berdasarkan topik yang diketik guru */
function aiBantuBuatSoal(topik, mapel, tipe) {
  var instruksiTipe = tipe === "PGK"
    ? "pilihan ganda kompleks (bisa lebih dari satu jawaban benar, dari 5 opsi berlabel A sampai E)"
    : "pilihan ganda biasa (hanya SATU jawaban benar, dari 5 opsi berlabel A sampai E)";
  var prompt = "Buatkan SATU soal " + instruksiTipe + " dalam Bahasa Indonesia untuk mata pelajaran " + (mapel || "umum") +
    " dengan topik \"" + topik + "\". Balas HANYA dengan JSON valid tanpa markdown/backtick, persis format berikut: " +
    "{\"pertanyaan\":\"...\",\"opsiA\":\"...\",\"opsiB\":\"...\",\"opsiC\":\"...\",\"opsiD\":\"...\",\"opsiE\":\"...\",\"kunci\":\"A\"}. " +
    "WAJIB isi kelima opsi (A sampai E) dengan pilihan jawaban yang masuk akal, jangan ada yang dikosongkan. Untuk kunci pilihan ganda kompleks tulis huruf jawaban benar dipisah koma, contoh \"A,C\".";

  var hasilMentah = _panggilGeminiAI(prompt).trim();
  hasilMentah = hasilMentah.replace(/^```json/i, "").replace(/^```/, "").replace(/```$/, "").trim();
  try {
    var obj = JSON.parse(hasilMentah);
    return { status: "SUCCESS", data: obj };
  } catch (e) {
    return { status: "ERROR", message: "AI memberikan format jawaban yang tidak terbaca. Silakan coba lagi." };
  }
}

// ==================== MODUL CBT / UJIAN TERJADWAL ====================
// Sesi ujian mengambil soal dari Bank Soal, punya jadwal, durasi, token, dan kelas peserta sendiri.

var HEADER_CBT_SESI = ["ID Sesi", "Guru Username", "Guru Nama", "Judul Ujian", "Jenis Ujian", "Mapel", "Kelas", "Daftar ID Soal", "Tanggal Mulai", "Jam Mulai", "Tanggal Selesai", "Jam Selesai", "Durasi Menit", "Token", "Acak Soal", "Acak Opsi", "Tampilkan Nilai", "Status", "Timestamp", "Mode Token", "Sertakan Rekap", "Tampilkan Review"];

function _pastikanSheetCbtSesiBenar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CBT_Sesi");
  if (!sheet) {
    sheet = ss.insertSheet("CBT_Sesi");
    sheet.appendRow(HEADER_CBT_SESI);
    return sheet;
  }
  var headerSaatIni = sheet.getRange(1, 1, 1, HEADER_CBT_SESI.length).getValues()[0];
  if (!headerSaatIni[19]) sheet.getRange(1, 20).setValue("Mode Token"); // Migrasi: tambah kolom baru tanpa mengubah yang lama
  if (!headerSaatIni[20]) sheet.getRange(1, 21).setValue("Sertakan Rekap");
  if (!headerSaatIni[21]) sheet.getRange(1, 22).setValue("Tampilkan Review");
  return sheet;
}

/**
 * Bandingkan 2 username dengan toleran. PENTING: username yang berupa NISN (angka semua,
 * kadang berawalan 0) bisa "kehilangan" angka nol di depan kalau Google Sheets salah mengira
 * itu angka biasa lalu menyimpannya sebagai Number bukan Text. Fallback ini membandingkan juga
 * versi tanpa angka nol di depan, supaya data lama yang sudah kadung begini tetap bisa cocok.
 */
/** Bentuk kunci lookup yang ternormalisasi (lowercase, trim, tanpa angka nol di depan) — dipakai berpasangan dengan _samaUsername */
function _normalisasiUsername(u) {
  var s = (u === null || typeof u === "undefined") ? "" : u.toString().trim().toLowerCase();
  var tanpaNolDepan = s.replace(/^0+/, "");
  return tanpaNolDepan !== "" ? tanpaNolDepan : s;
}

function _samaUsername(a, b) {
  var sa = (a === null || typeof a === "undefined") ? "" : a.toString().trim().toLowerCase();
  var sb = (b === null || typeof b === "undefined") ? "" : b.toString().trim().toLowerCase();
  if (sa === sb) return true;
  var na = sa.replace(/^0+/, "");
  var nb = sb.replace(/^0+/, "");
  return na !== "" && na === nb;
}

function _buatTokenUjian() {
  var karakter = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; // Tanpa 0/O/1/I agar tidak rancu dibaca siswa
  var token = "";
  for (var i = 0; i < 6; i++) token += karakter.charAt(Math.floor(Math.random() * karakter.length));
  return token;
}

/**
 * Menghasilkan token yang otomatis berganti tiap 10 menit, dihitung dari idSesi + blok waktu 10 menit saat ini.
 * PENTING: fungsi ini sengaja hanya ada di Code.gs (server), TIDAK PERNAH dikirim algoritmenya ke client,
 * supaya siswa tidak bisa menghitung sendiri token mendatang lewat DevTools.
 * blokKe: 0 = blok waktu saat ini, -1 = blok sebelumnya (dipakai sbg toleransi jeda saat token baru saja berganti)
 */
function _hitungTokenOtomatis(idSesi, blokKe) {
  var SEPULUH_MENIT_MS = 10 * 60 * 1000;
  var blokWaktu = Math.floor(new Date().getTime() / SEPULUH_MENIT_MS) + (blokKe || 0);
  var seed = idSesi.toString() + "|" + blokWaktu;
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, seed);
  var karakter = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  var token = "";
  for (var i = 0; i < 6; i++) {
    var b = digest[i];
    if (b < 0) b += 256;
    token += karakter.charAt(b % karakter.length);
  }
  return token;
}

/** Guru melihat token yang berlaku SAAT INI untuk sesi bermode token Otomatis, plus hitung mundur ke pergantian berikutnya */
function guruAmbilTokenOtomatisSaatIni(idSesi) {
  var sesi = guruAmbilDetailSesi(idSesi);
  if (!sesi) return { status: "ERROR", message: "Sesi tidak ditemukan." };
  if (sesi.modeToken !== "Otomatis") return { status: "ERROR", message: "Sesi ini tidak memakai mode token otomatis." };

  var SEPULUH_MENIT_MS = 10 * 60 * 1000;
  var sekarang = new Date().getTime();
  var detikTersisa = 600 - Math.floor((sekarang % SEPULUH_MENIT_MS) / 1000);

  return { status: "SUCCESS", token: _hitungTokenOtomatis(idSesi, 0), detikTersisa: detikTersisa };
}

function _kocokArray(arr) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = a[i]; a[i] = a[j]; a[j] = tmp;
  }
  return a;
}

/**
 * Menormalkan nilai jam (Jam Mulai/Jam Selesai) menjadi teks "HH:mm" yang bersih.
 * Berjaga-jaga untuk data SESI LAMA yang sempat tersimpan sebagai objek Date oleh Google Sheets
 * (sebelum kolom ini dikunci sebagai teks) — tanpa ini, sesi lama akan gagal total dibuka siswa.
 */
function _normalisasiJam(val) {
  if (!val) return "";
  if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
  }
  var s = val.toString().trim();
  if (/^[A-Za-z]{3} [A-Za-z]{3} \d{2} \d{4}/.test(s)) { // Sudah keburu jadi string hasil toString() dari Date
    var d = new Date(s);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), "HH:mm");
  }
  return s;
}

function _baris2Sesi(row) {
  return {
    id: row[0].toString(), guruUsername: row[1].toString(), guruNama: row[2].toString(),
    judul: row[3].toString(), jenis: row[4].toString(), mapel: row[5].toString(),
    kelas: row[6].toString().split(",").map(function(k) { return k.trim(); }).filter(Boolean),
    daftarIdSoal: row[7].toString().split(",").map(function(s) { return s.trim(); }).filter(Boolean),
    tglMulai: row[8].toString(), jamMulai: _normalisasiJam(row[9]) || "00:00",
    tglSelesai: row[10].toString(), jamSelesai: _normalisasiJam(row[11]) || "23:59",
    durasiMenit: Number(row[12]) || 0, token: row[13].toString(),
    acakSoal: row[14] === true || row[14].toString().toUpperCase() === "TRUE",
    acakOpsi: row[15] === true || row[15].toString().toUpperCase() === "TRUE",
    tampilkanNilai: row[16] === true || row[16].toString().toUpperCase() === "TRUE",
    statusManual: row[17] ? row[17].toString() : "Aktif",
    timestamp: row[18] ? row[18].toString() : "",
    modeToken: row[19] ? row[19].toString() : "Manual",
    sertakanRekap: row[20] ? row[20].toString() !== "FALSE" : true, // Default TRUE untuk sesi lama sebelum fitur ini ada
    tampilkanReview: row.length > 21 && (row[21] === true || row[21].toString().toUpperCase() === "TRUE")
  };
}

function _statusWaktuSesi(sesi) {
  if (sesi.statusManual === "Nonaktif") return "Nonaktif";
  var sekarang = new Date();
  var mulai = new Date(sesi.tglMulai + "T" + (sesi.jamMulai || "00:00") + ":00");
  var selesai = new Date(sesi.tglSelesai + "T" + (sesi.jamSelesai || "23:59") + ":00");
  if (sekarang < mulai) return "Terjadwal";
  if (sekarang > selesai) return "Selesai";
  return "Berlangsung";
}

/** data: {guruUsername,guruNama,judul,jenis,mapel,kelasArray,daftarIdSoal,tglMulai,jamMulai,tglSelesai,jamSelesai,durasiMenit,token,modeToken,acakSoal,acakOpsi,tampilkanNilai} */
function guruBuatSesiUjian(data) {
  if (!data.daftarIdSoal || data.daftarIdSoal.length === 0) return { status: "ERROR", message: "Pilih minimal satu soal dari Bank Soal untuk sesi ini." };
  data.daftarIdSoal = data.daftarIdSoal.filter(function(v, i, arr) { return arr.indexOf(v) === i; });
  if (!data.kelasArray || data.kelasArray.length === 0) return { status: "ERROR", message: "Pilih minimal satu kelas peserta." };

  var sheet = _pastikanSheetCbtSesiBenar();
  var waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var modeToken = data.modeToken === "Otomatis" ? "Otomatis" : "Manual";
  var token = modeToken === "Otomatis" ? "" : ((data.token && data.token.toString().trim()) ? data.token.toString().trim().toUpperCase() : _buatTokenUjian());
  var idBaru = _buatIdUnik("CBT");

  sheet.appendRow([
    idBaru, data.guruUsername, data.guruNama, data.judul, data.jenis, data.mapel,
    data.kelasArray.join(","), data.daftarIdSoal.join(","),
    data.tglMulai, data.jamMulai, data.tglSelesai, data.jamSelesai,
    Number(data.durasiMenit) || 0, token,
    data.acakSoal ? "TRUE" : "FALSE", data.acakOpsi ? "TRUE" : "FALSE", data.tampilkanNilai ? "TRUE" : "FALSE",
    "Aktif", waktu, modeToken, data.sertakanRekap === false ? "FALSE" : "TRUE", data.tampilkanReview ? "TRUE" : "FALSE"
  ]);
  var barisTerakhir = sheet.getLastRow();
  sheet.getRange(barisTerakhir, 9).setNumberFormat("@").setValue(data.tglMulai);
  sheet.getRange(barisTerakhir, 10).setNumberFormat("@").setValue(data.jamMulai);
  sheet.getRange(barisTerakhir, 11).setNumberFormat("@").setValue(data.tglSelesai);
  sheet.getRange(barisTerakhir, 12).setNumberFormat("@").setValue(data.jamSelesai);
  // Kunci & tulis ulang kolom TRUE/FALSE sebagai teks eksplisit (setNumberFormat saja tidak retroaktif memperbaiki nilai yang sudah kadung dikonversi jadi Boolean oleh appendRow)
  sheet.getRange(barisTerakhir, 15).setNumberFormat("@").setValue(data.acakSoal ? "TRUE" : "FALSE");
  sheet.getRange(barisTerakhir, 16).setNumberFormat("@").setValue(data.acakOpsi ? "TRUE" : "FALSE");
  sheet.getRange(barisTerakhir, 17).setNumberFormat("@").setValue(data.tampilkanNilai ? "TRUE" : "FALSE");
  sheet.getRange(barisTerakhir, 21).setNumberFormat("@").setValue(data.sertakanRekap === false ? "FALSE" : "TRUE");
  sheet.getRange(barisTerakhir, 22).setNumberFormat("@").setValue(data.tampilkanReview ? "TRUE" : "FALSE");

  catatLogAktivitas(data.guruNama, "Buat Sesi CBT", "Membuat sesi ujian \"" + data.judul + "\" untuk kelas " + data.kelasArray.join(", "));
  var pesanToken = modeToken === "Otomatis" ? "Token otomatis berganti tiap 10 menit, lihat di daftar sesi." : "Token: " + token;
  return { status: "SUCCESS", message: "Sesi ujian \"" + data.judul + "\" berhasil dibuat. " + pesanToken, id: idBaru, token: token };
}

/** Duplikat sesi ujian yang sudah ada untuk dijadikan Remedial — soal & pengaturan disalin, guru tinggal atur ulang jadwal/token */
function guruDuplikatSesiRemedial(idSesi) {
  var sesiAsli = guruAmbilDetailSesi(idSesi);
  if (!sesiAsli) return { status: "ERROR", message: "Sesi asal tidak ditemukan." };

  var sheet = _pastikanSheetCbtSesiBenar();
  var waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var hariIni = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var idBaru = _buatIdUnik("CBT");
  var token = _buatTokenUjian();

  sheet.appendRow([
    idBaru, sesiAsli.guruUsername, sesiAsli.guruNama, sesiAsli.judul + " (Remedial)", "Remedial", sesiAsli.mapel,
    sesiAsli.kelas.join(","), sesiAsli.daftarIdSoal.join(","),
    hariIni, "07:00", hariIni, "23:59",
    sesiAsli.durasiMenit, token, "FALSE", "FALSE", "TRUE",
    "Aktif", waktu, "Manual", "TRUE", sesiAsli.tampilkanReview ? "TRUE" : "FALSE"
  ]);
  var barisTerakhir = sheet.getLastRow();
  sheet.getRange(barisTerakhir, 9).setNumberFormat("@").setValue(hariIni);
  sheet.getRange(barisTerakhir, 10).setNumberFormat("@").setValue("07:00");
  sheet.getRange(barisTerakhir, 15).setNumberFormat("@").setValue("FALSE");
  sheet.getRange(barisTerakhir, 16).setNumberFormat("@").setValue("FALSE");
  sheet.getRange(barisTerakhir, 17).setNumberFormat("@").setValue("TRUE");
  sheet.getRange(barisTerakhir, 21).setNumberFormat("@").setValue("TRUE");
  sheet.getRange(barisTerakhir, 22).setNumberFormat("@").setValue(sesiAsli.tampilkanReview ? "TRUE" : "FALSE");
  sheet.getRange(barisTerakhir, 11).setNumberFormat("@").setValue(hariIni);
  sheet.getRange(barisTerakhir, 12).setNumberFormat("@").setValue("23:59");

  catatLogAktivitas(sesiAsli.guruNama, "Buat Sesi Remedial", "Menduplikat \"" + sesiAsli.judul + "\" jadi sesi remedial.");
  return { status: "SUCCESS", message: "Sesi remedial berhasil dibuat dari \"" + sesiAsli.judul + "\". Silakan atur jadwal & token, lalu bagikan hanya ke siswa yang perlu remedial.", idBaru: idBaru };
}

function guruEditSesiUjian(idSesi, data) {
  if (!data.daftarIdSoal || data.daftarIdSoal.length === 0) return { status: "ERROR", message: "Pilih minimal satu soal dari Bank Soal untuk sesi ini." };
  data.daftarIdSoal = data.daftarIdSoal.filter(function(v, i, arr) { return arr.indexOf(v) === i; });
  if (!data.kelasArray || data.kelasArray.length === 0) return { status: "ERROR", message: "Pilih minimal satu kelas peserta." };

  var sheet = _pastikanSheetCbtSesiBenar();
  var dataSheet = sheet.getDataRange().getValues();
  for (var i = 1; i < dataSheet.length; i++) {
    if (dataSheet[i][0].toString() === idSesi.toString()) {
      var modeToken = data.modeToken === "Otomatis" ? "Otomatis" : "Manual";
      var token = modeToken === "Otomatis" ? "" : ((data.token && data.token.toString().trim()) ? data.token.toString().trim().toUpperCase() : dataSheet[i][13].toString());
      var baris = [
        idSesi, dataSheet[i][1], dataSheet[i][2], data.judul, data.jenis, data.mapel,
        data.kelasArray.join(","), data.daftarIdSoal.join(","),
        data.tglMulai, data.jamMulai, data.tglSelesai, data.jamSelesai,
        Number(data.durasiMenit) || 0, token,
        data.acakSoal ? "TRUE" : "FALSE", data.acakOpsi ? "TRUE" : "FALSE", data.tampilkanNilai ? "TRUE" : "FALSE",
        dataSheet[i][17], dataSheet[i][18], modeToken, data.sertakanRekap === false ? "FALSE" : "TRUE", data.tampilkanReview ? "TRUE" : "FALSE"
      ];
      sheet.getRange(i + 1, 1, 1, baris.length).setValues([baris]);
      sheet.getRange(i + 1, 9).setNumberFormat("@").setValue(data.tglMulai);
      sheet.getRange(i + 1, 10).setNumberFormat("@").setValue(data.jamMulai);
      sheet.getRange(i + 1, 11).setNumberFormat("@").setValue(data.tglSelesai);
      sheet.getRange(i + 1, 12).setNumberFormat("@").setValue(data.jamSelesai);
      sheet.getRange(i + 1, 15).setNumberFormat("@").setValue(data.acakSoal ? "TRUE" : "FALSE");
      sheet.getRange(i + 1, 16).setNumberFormat("@").setValue(data.acakOpsi ? "TRUE" : "FALSE");
      sheet.getRange(i + 1, 17).setNumberFormat("@").setValue(data.tampilkanNilai ? "TRUE" : "FALSE");
      sheet.getRange(i + 1, 21).setNumberFormat("@").setValue(data.sertakanRekap === false ? "FALSE" : "TRUE");
      sheet.getRange(i + 1, 22).setNumberFormat("@").setValue(data.tampilkanReview ? "TRUE" : "FALSE");
      return { status: "SUCCESS", message: "Sesi ujian berhasil diperbarui." };
    }
  }
  return { status: "ERROR", message: "Sesi ujian tidak ditemukan." };
}

function guruUbahStatusAktifSesi(idSesi, statusBaru) {
  var sheet = _pastikanSheetCbtSesiBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === idSesi.toString()) {
      sheet.getRange(i + 1, 18).setValue(statusBaru === "Nonaktif" ? "Nonaktif" : "Aktif");
      return { status: "SUCCESS", message: "Status sesi ujian diperbarui." };
    }
  }
  return { status: "ERROR", message: "Sesi ujian tidak ditemukan." };
}

function guruHapusSesiUjian(idSesi) {
  var sheet = _pastikanSheetCbtSesiBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0].toString() === idSesi.toString()) {
      sheet.deleteRow(i + 1);
      return { status: "SUCCESS", message: "Sesi ujian berhasil dihapus." };
    }
  }
  return { status: "ERROR", message: "Sesi ujian tidak ditemukan." };
}

function guruAmbilDaftarSesi(guruUsername) {
  var sheet = _pastikanSheetCbtSesiBenar();
  var data = sheet.getDataRange().getValues();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (data[i][1].toString().trim() !== guruUsername.toString().trim()) continue;
    var sesi = _baris2Sesi(data[i]);
    sesi.jumlahSoal = sesi.daftarIdSoal.length;
    sesi.statusWaktu = _statusWaktuSesi(sesi);
    hasil.push(sesi);
  }
  hasil.sort(function(a, b) { return b.timestamp.localeCompare(a.timestamp); });
  return hasil;
}

function guruAmbilDetailSesi(idSesi) {
  var sheet = _pastikanSheetCbtSesiBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === idSesi.toString()) {
      var sesi = _baris2Sesi(data[i]);
      sesi.statusWaktu = _statusWaktuSesi(sesi);
      return sesi;
    }
  }
  return null;
}

// ---------- Hasil & Jawaban Ujian ----------

var HEADER_HASIL_UJIAN = ["ID Hasil", "ID Sesi", "Username", "Nama Siswa", "Kelas", "Skor Objektif", "Skor Maksimal", "Nilai Objektif", "Nilai Essay", "Nilai Akhir", "Status", "Waktu Mulai", "Waktu Selesai", "Jumlah Pelanggaran", "Catatan Waktu"];

function _pastikanSheetHasilUjianBenar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Hasil_Ujian");
  if (!sheet) {
    sheet = ss.insertSheet("Hasil_Ujian");
    sheet.appendRow(HEADER_HASIL_UJIAN);
    return sheet;
  }
  var headerSaatIni = sheet.getRange(1, 1, 1, HEADER_HASIL_UJIAN.length).getValues()[0];
  if (!headerSaatIni[14]) sheet.getRange(1, 15).setValue("Catatan Waktu"); // Migrasi: tambah kolom baru tanpa mengubah yang lama
  return sheet;
}

var HEADER_JAWABAN_SISWA = ["ID Hasil", "ID Sesi", "Username", "ID Soal", "Jawaban Siswa", "Benar", "Skor Didapat", "Catatan Guru"];

function _pastikanSheetJawabanSiswaBenar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Jawaban_Siswa");
  if (!sheet) {
    sheet = ss.insertSheet("Jawaban_Siswa");
    sheet.appendRow(HEADER_JAWABAN_SISWA);
    return sheet;
  }
  var headerSaatIni = sheet.getRange(1, 1, 1, HEADER_JAWABAN_SISWA.length).getValues()[0];
  var headerCocok = HEADER_JAWABAN_SISWA.every(function(h, idx) { return headerSaatIni[idx] === h; });
  if (!headerCocok) {
    sheet.getRange(1, 1, 1, HEADER_JAWABAN_SISWA.length).setValues([HEADER_JAWABAN_SISWA]);
  }
  return sheet;
}

/** Daftar ujian CBT yang tersedia untuk kelas siswa ybs, lengkap dengan status waktu & status pengerjaannya */
function siswaAmbilDaftarUjianTersedia(kelasSiswa, usernameSiswa) {
  var sheetSesi = _pastikanSheetCbtSesiBenar();
  var dataSesi = sheetSesi.getDataRange().getValues();
  var sheetHasil = _pastikanSheetHasilUjianBenar();
  var dataHasil = sheetHasil.getDataRange().getValues();

  var petaHasil = {};
  for (var h = 1; h < dataHasil.length; h++) {
    if (_samaUsername(dataHasil[h][2], usernameSiswa)) {
      var idSesiBaris = dataHasil[h][1].toString();
      var statusBaris = dataHasil[h][10].toString();
      var yangSudahAda = petaHasil[idSesiBaris];
      // Kalau ada baris duplikat (data lama dari sebelum diperbaiki), prioritaskan status FINAL
      // (Selesai/Terkunci) — jangan sampai baris "Sedang Mengerjakan" yang lebih baru menimpa hasil yang sudah final.
      if (yangSudahAda && (yangSudahAda.status === "Selesai" || yangSudahAda.status === "Terkunci")) continue;
      petaHasil[idSesiBaris] = { status: statusBaris, nilaiAkhir: dataHasil[h][9], idHasil: dataHasil[h][0].toString() };
    }
  }

  var hasil = [];
  for (var i = 1; i < dataSesi.length; i++) {
    if (!dataSesi[i][0]) continue;
    var sesi = _baris2Sesi(dataSesi[i]);
    if (sesi.kelas.indexOf(kelasSiswa.toString().trim()) === -1) continue;
    if (sesi.statusManual === "Nonaktif") continue;

    sesi.statusWaktu = _statusWaktuSesi(sesi);
    sesi.jumlahSoal = sesi.daftarIdSoal.length;
    var hasilSiswa = petaHasil[sesi.id];
    sesi.sudahDikerjakan = !!hasilSiswa && (hasilSiswa.status === "Selesai" || hasilSiswa.status === "Terkunci");
    sesi.sedangDikerjakan = !!hasilSiswa && hasilSiswa.status === "Sedang Mengerjakan";
    sesi.statusPengerjaan = hasilSiswa ? hasilSiswa.status : "";
    sesi.nilaiSaya = sesi.sudahDikerjakan ? (sesi.tampilkanNilai ? hasilSiswa.nilaiAkhir : null) : null;
    sesi.idHasilSaya = hasilSiswa ? hasilSiswa.idHasil : "";
    delete sesi.daftarIdSoal;
    delete sesi.token;
    hasil.push(sesi);
  }
  hasil.sort(function(a, b) { return (a.tglMulai + a.jamMulai).localeCompare(b.tglMulai + b.jamMulai); });
  return hasil;
}

/** Siswa membuka lembar soal: validasi jadwal & token, lalu susun soal (acak jika diaktifkan, tanpa kunci jawaban) */
function siswaMulaiUjian(idSesi, tokenInput, usernameSiswa, namaSiswa, kelasSiswa) {
  var sesi = guruAmbilDetailSesi(idSesi);
  if (!sesi) return { status: "ERROR", message: "Sesi ujian tidak ditemukan." };
  if (sesi.kelas.indexOf(kelasSiswa.toString().trim()) === -1) return { status: "ERROR", message: "Ujian ini bukan untuk kelas Anda." };

  var statusWaktu = _statusWaktuSesi(sesi);
  if (statusWaktu === "Terjadwal") return { status: "ERROR", message: "Ujian belum dimulai. Jadwal: " + sesi.tglMulai + " pukul " + sesi.jamMulai + "." };
  if (statusWaktu === "Selesai" || statusWaktu === "Nonaktif") return { status: "ERROR", message: "Waktu ujian sudah berakhir." };
  if (sesi.modeToken === "Otomatis") {
    var tokenBenar = tokenInput.toString().trim().toUpperCase();
    var cocok = tokenBenar === _hitungTokenOtomatis(idSesi, 0) || tokenBenar === _hitungTokenOtomatis(idSesi, -1);
    if (!cocok) {
      return { status: "ERROR", message: "Token yang Anda masukkan salah atau sudah kedaluwarsa (token berganti tiap 10 menit). Silakan tanyakan token TERBARU kepada guru pengawas." };
    }
  } else if (sesi.token && sesi.token.toString().trim().toUpperCase() !== tokenInput.toString().trim().toUpperCase()) {
    return { status: "ERROR", message: "Token yang Anda masukkan salah. Silakan tanyakan token kepada guru pengawas." };
  }

  // LockService: cegah race condition kalau tombol "Mulai Ujian" diklik dobel/cepat —
  // tanpa ini, dua request bersamaan bisa sama-sama menganggap "belum ada data" dan
  // menciptakan DUA baris Hasil_Ujian untuk siswa+sesi yang sama (bikin data ganda/rekap kacau).
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheetHasil = _pastikanSheetHasilUjianBenar();
    var dataHasil = sheetHasil.getDataRange().getValues();
    var idHasil = "", waktuMulaiTersimpan = "", sudahAda = false;

    for (var i = 1; i < dataHasil.length; i++) {
      if (dataHasil[i][1].toString().trim() === idSesi.toString().trim() && _samaUsername(dataHasil[i][2], usernameSiswa)) {
        idHasil = dataHasil[i][0].toString();
        if (dataHasil[i][10].toString() === "Selesai") {
          return { status: "ERROR", message: "Anda sudah menyelesaikan ujian ini sebelumnya." };
        }
        if (dataHasil[i][10].toString() === "Terkunci") {
          return { status: "ERROR", message: "Ujian Anda sedang dikunci karena terindikasi ada pelanggaran (keluar layar/pindah tab berulang kali). Silakan hubungi guru pengawas untuk membuka kembali." };
        }
        waktuMulaiTersimpan = dataHasil[i][11].toString();
        sudahAda = true;
        break;
      }
    }

    if (!sudahAda) {
      idHasil = _buatIdUnik("HSL");
      var waktuMulai = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      sheetHasil.appendRow([idHasil, idSesi, usernameSiswa, namaSiswa, kelasSiswa, 0, 0, 0, "", 0, "Sedang Mengerjakan", waktuMulai, "", 0]);
      var br = sheetHasil.getLastRow();
      sheetHasil.getRange(br, 3).setNumberFormat("@").setValue(usernameSiswa); // Kunci Username sebagai TEKS (jangan sampai NISN kehilangan angka 0 di depan)
      sheetHasil.getRange(br, 12).setNumberFormat("@").setValue(waktuMulai);
      waktuMulaiTersimpan = waktuMulai;
    }
  } finally {
    lock.releaseLock();
  }

  var sheetBankSoal = _pastikanSheetBankSoalBenar();
  var dataSoal = sheetBankSoal.getDataRange().getValues();
  var petaSoal = {};
  for (var s = 1; s < dataSoal.length; s++) { if (dataSoal[s][0]) petaSoal[dataSoal[s][0].toString()] = _baris2Soal(dataSoal[s]); }

  var daftarIdSoal = sesi.acakSoal ? _kocokArray(sesi.daftarIdSoal) : sesi.daftarIdSoal;
  var labelHuruf = ["A", "B", "C", "D", "E"];
  var soalUntukSiswa = [];
  daftarIdSoal.forEach(function(idSoal) {
    var soal = petaSoal[idSoal];
    if (!soal) return;
    var opsiBerlabel = [];
    soal.opsi.forEach(function(teks, idx) { if (teks) opsiBerlabel.push({ label: labelHuruf[idx], teks: teks }); });
    if (sesi.acakOpsi && (soal.tipe === "PG" || soal.tipe === "PGK")) opsiBerlabel = _kocokArray(opsiBerlabel);
    soalUntukSiswa.push({ id: soal.id, tipe: soal.tipe, pertanyaan: soal.pertanyaan, opsi: opsiBerlabel, poin: soal.poin, gambarUrl: soal.gambarUrl || "" });
  });

  // Sisa waktu dihitung dari waktu mulai TERSIMPAN di server (bukan waktu klik saat itu),
  // supaya siswa tidak bisa "reset" timer dengan cara refresh/reload halaman.
  var epochMulai = new Date(waktuMulaiTersimpan.replace(" ", "T")).getTime();
  var batasWaktuDurasi = epochMulai + (sesi.durasiMenit * 60000);
  var batasWaktuSesi = new Date(sesi.tglSelesai + "T" + (sesi.jamSelesai || "23:59") + ":00").getTime();
  var epochBatasAkhir = Math.min(batasWaktuDurasi, batasWaktuSesi);

  return {
    status: "SUCCESS", idHasil: idHasil, judul: sesi.judul, jenis: sesi.jenis, mapel: sesi.mapel,
    acakOpsi: sesi.acakOpsi,
    soal: soalUntukSiswa, epochMulai: epochMulai, epochBatasAkhir: epochBatasAkhir, epochSekarang: new Date().getTime(),
    jawabanTersimpan: sudahAda ? _siswaAmbilJawabanTersimpan(idHasil) : {}
  };
}

/**
 * Autosave: dipanggil setiap kali siswa pindah soal (bukan hanya saat submit akhir).
 * Bersifat upsert per (idHasil, idSoal) — TIDAK menilai benar/salah dulu (baru dinilai saat submit akhir),
 * supaya ringan & cepat dipanggil berulang kali. Kalau ujian sudah Selesai/Terkunci, diabaikan (no-op)
 * supaya tidak ada race condition menimpa hasil yang sudah final.
 */
function siswaSimpanJawabanOtomatis(idHasil, idSesi, usernameSiswa, idSoal, jawaban) {
  var sheetHasil = _pastikanSheetHasilUjianBenar();
  var dataHasil = sheetHasil.getDataRange().getValues();
  for (var h = 1; h < dataHasil.length; h++) {
    if (dataHasil[h][0].toString() === idHasil.toString()) {
      if (dataHasil[h][10].toString() !== "Sedang Mengerjakan") return { status: "IGNORED" };
      break;
    }
  }

  var jawabanTeks = Array.isArray(jawaban) ? jawaban.join(",") : (jawaban || "").toString();
  var sheet = _pastikanSheetJawabanSiswaBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === idHasil.toString() && data[i][3].toString() === idSoal.toString()) {
      sheet.getRange(i + 1, 5).setValue(jawabanTeks);
      return { status: "SUCCESS" };
    }
  }
  sheet.appendRow([idHasil, idSesi, usernameSiswa, idSoal, jawabanTeks, "", "", ""]);
  sheet.getRange(sheet.getLastRow(), 3).setNumberFormat("@").setValue(usernameSiswa);
  return { status: "SUCCESS" };
}

/** Ambil jawaban yang sudah tersimpan (dari autosave) untuk satu percobaan ujian — dipakai untuk memulihkan progres saat resume */
function _siswaAmbilJawabanTersimpan(idHasil) {
  var sheet = _pastikanSheetJawabanSiswaBenar();
  var data = sheet.getDataRange().getValues();
  var peta = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === idHasil.toString()) {
      peta[data[i][3].toString()] = data[i][4] ? data[i][4].toString() : "";
    }
  }
  return peta;
}
function siswaSubmitUjian(idHasil, idSesi, usernameSiswa, jawabanArray, paksaKarenaPelanggaran) {
  var iniAutoFinalisasi = !jawabanArray; // Dipanggil oleh jaring pengaman (waktu habis, browser tertutup), bukan klik siswa
  var sesi = guruAmbilDetailSesi(idSesi);
  if (!sesi) return { status: "ERROR", message: "Sesi ujian tidak ditemukan." };

  var sheetHasil = _pastikanSheetHasilUjianBenar();
  var dataHasil = sheetHasil.getDataRange().getValues();
  var barisHasil = -1;
  for (var i = 1; i < dataHasil.length; i++) {
    if (dataHasil[i][0].toString() === idHasil.toString()) { barisHasil = i; break; }
  }
  if (barisHasil === -1) return { status: "ERROR", message: "Data pengerjaan tidak ditemukan." };
  if (dataHasil[barisHasil][10].toString() === "Selesai" || dataHasil[barisHasil][10].toString() === "Terkunci") {
    return { status: "ERROR", message: "Ujian ini sudah pernah dikumpulkan sebelumnya." };
  }

  var sheetBankSoal = _pastikanSheetBankSoalBenar();
  var dataSoal = sheetBankSoal.getDataRange().getValues();
  var petaSoal = {};
  for (var s = 1; s < dataSoal.length; s++) { if (dataSoal[s][0]) petaSoal[dataSoal[s][0].toString()] = _baris2Soal(dataSoal[s]); }

  var sheetJawaban = _pastikanSheetJawabanSiswaBenar();
  var dataJawabanSaatIni = sheetJawaban.getDataRange().getValues();
  var skorDiperoleh = 0, skorMaksimal = 0, adaEssay = false;

  // Kalau jawabanArray tidak dikirim (dipanggil oleh jaring pengaman auto-finalisasi, bukan dari
  // klik siswa langsung), ambil jawaban terakhir yang sempat ter-autosave sebagai gantinya.
  if (!jawabanArray) {
    jawabanArray = [];
    for (var jj = 1; jj < dataJawabanSaatIni.length; jj++) {
      if (dataJawabanSaatIni[jj][0].toString() === idHasil.toString()) {
        jawabanArray.push({ idSoal: dataJawabanSaatIni[jj][3].toString(), jawaban: dataJawabanSaatIni[jj][4] });
      }
    }
  }

  sesi.daftarIdSoal.forEach(function(idSoal) {
    var soal = petaSoal[idSoal];
    if (!soal) return;
    skorMaksimal += soal.poin;

    var jawabObj = (jawabanArray || []).filter(function(j) { return j.idSoal === idSoal; })[0];
    var jawabanMentah = jawabObj ? jawabObj.jawaban : "";
    var jawabanTeks = Array.isArray(jawabanMentah) ? jawabanMentah.join(",") : (jawabanMentah || "").toString();
    var benar = "", skorSoal = 0;

    if (soal.tipe === "PG") {
      benar = (jawabanTeks.trim().toUpperCase() === soal.kunci.trim().toUpperCase() && jawabanTeks.trim() !== "") ? "TRUE" : "FALSE";
      if (benar === "TRUE") skorSoal = soal.poin;
    } else if (soal.tipe === "PGK") {
      var jawabSet = jawabanTeks.split(",").map(function(x) { return x.trim().toUpperCase(); }).filter(Boolean).sort();
      var kunciSet = soal.kunci.split(",").map(function(x) { return x.trim().toUpperCase(); }).filter(Boolean).sort();
      benar = (jawabSet.length > 0 && jawabSet.join(",") === kunciSet.join(",")) ? "TRUE" : "FALSE";
      if (benar === "TRUE") skorSoal = soal.poin;
    } else if (soal.tipe === "BS") {
      benar = (jawabanTeks.trim().toLowerCase() === soal.kunci.trim().toLowerCase() && jawabanTeks.trim() !== "") ? "TRUE" : "FALSE";
      if (benar === "TRUE") skorSoal = soal.poin;
    } else if (soal.tipe === "Isian") {
      benar = (jawabanTeks.trim().toLowerCase() === soal.kunci.trim().toLowerCase() && jawabanTeks.trim() !== "") ? "TRUE" : "FALSE";
      if (benar === "TRUE") skorSoal = soal.poin;
    } else { // Essay -> menunggu koreksi manual guru
      adaEssay = true; benar = ""; skorSoal = 0;
    }

    skorDiperoleh += skorSoal;

    // Upsert (bukan appendRow) supaya tidak duplikat dengan baris yang sudah dibuat autosave selama ujian berjalan
    var barisDitemukan = -1;
    for (var j = 1; j < dataJawabanSaatIni.length; j++) {
      if (dataJawabanSaatIni[j][0].toString() === idHasil.toString() && dataJawabanSaatIni[j][3].toString() === idSoal.toString()) { barisDitemukan = j; break; }
    }
    if (barisDitemukan >= 0) {
      sheetJawaban.getRange(barisDitemukan + 1, 5).setValue(jawabanTeks);
      sheetJawaban.getRange(barisDitemukan + 1, 6).setNumberFormat("@").setValue(benar);
      sheetJawaban.getRange(barisDitemukan + 1, 7).setValue(skorSoal);
    } else {
      sheetJawaban.appendRow([idHasil, idSesi, usernameSiswa, idSoal, jawabanTeks, benar, skorSoal, ""]);
      var brJwb = sheetJawaban.getLastRow();
      sheetJawaban.getRange(brJwb, 3).setNumberFormat("@").setValue(usernameSiswa);
      sheetJawaban.getRange(brJwb, 6).setNumberFormat("@").setValue(benar); // Kunci kolom Benar sebagai TEKS (jangan sampai "TRUE"/"FALSE" jadi Boolean asli)
    }
  });

  var nilaiObjektif = skorMaksimal > 0 ? Math.round((skorDiperoleh / skorMaksimal) * 10000) / 100 : 0;
  var nilaiAkhir = nilaiObjektif;
  var statusAkhir = paksaKarenaPelanggaran ? "Terkunci" : "Selesai";
  var waktuSelesai = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  // Validasi waktu submit di SERVER (bukan cuma percaya timer di browser) — supaya manipulasi
  // timer lewat DevTools tidak berpengaruh. Toleransi 90 detik untuk jeda jaringan/klik.
  var TOLERANSI_MS = 90 * 1000;
  var epochMulai = new Date(dataHasil[barisHasil][11].toString().replace(" ", "T")).getTime();
  var batasWaktuDurasi = epochMulai + (sesi.durasiMenit * 60000);
  var batasWaktuSesi = new Date(sesi.tglSelesai + "T" + (sesi.jamSelesai || "23:59") + ":00").getTime();
  var epochBatasAkhir = Math.min(batasWaktuDurasi, batasWaktuSesi);
  var terlambatMs = new Date(waktuSelesai.replace(" ", "T")).getTime() - epochBatasAkhir;
  var catatanWaktu = iniAutoFinalisasi
    ? "Waktu habis — dinilai otomatis oleh sistem (siswa tidak sempat klik selesai, kemungkinan koneksi/perangkat terputus)"
    : (terlambatMs > TOLERANSI_MS ? ("Submit terlambat ~" + Math.round(terlambatMs / 60000) + " menit dari batas waktu seharusnya") : "");

  sheetHasil.getRange(barisHasil + 1, 6, 1, 5).setValues([[
    skorDiperoleh, skorMaksimal, nilaiObjektif, adaEssay ? "Menunggu" : "-", nilaiAkhir
  ]]);
  sheetHasil.getRange(barisHasil + 1, 11).setValue(statusAkhir); // Kolom Status
  sheetHasil.getRange(barisHasil + 1, 13).setNumberFormat("@").setValue(waktuSelesai); // Kolom Waktu Selesai
  sheetHasil.getRange(barisHasil + 1, 15).setValue(catatanWaktu); // Kolom Catatan Waktu (SEBELUMNYA SALAH: menimpa kolom Jumlah Pelanggaran)

  return {
    status: "SUCCESS",
    message: paksaKarenaPelanggaran ? "Ujian dikunci otomatis karena pelanggaran berulang." : "Jawaban berhasil dikumpulkan.",
    nilaiSementara: nilaiAkhir, adaEssay: adaEssay, tampilkanNilai: sesi.tampilkanNilai
  };
}

/** Dipanggil dari sisi siswa saat terdeteksi pindah tab/keluar layar penuh selama ujian berlangsung.
 *  Setelah mencapai batas pelanggaran, ujian LANGSUNG DIKUNCI di server (status "Terkunci"),
 *  TANPA menilai/submit jawaban — supaya tidak buru-buru memvonis nilai rendah kalau ternyata cuma
 *  kesalahan teknis (notifikasi HP, dsb). Guru yang menentukan lewat tombol "Buka Kembali" di Rekap Hasil. */
function siswaLaporPelanggaran(idHasil) {
  var BATAS_MAKS = 3;
  var sheet = _pastikanSheetHasilUjianBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === idHasil.toString()) {
      if (data[i][10].toString() === "Selesai" || data[i][10].toString() === "Terkunci") {
        return { jumlahPelanggaran: Number(data[i][13]) || 0, terkunci: true, batasMaks: BATAS_MAKS };
      }
      var jumlah = (Number(data[i][13]) || 0) + 1;
      sheet.getRange(i + 1, 14).setValue(jumlah);
      var terkunci = jumlah >= BATAS_MAKS;
      if (terkunci) {
        sheet.getRange(i + 1, 11).setValue("Terkunci"); // Kolom Status — kunci langsung, TANPA menilai
      }
      return { jumlahPelanggaran: jumlah, terkunci: terkunci, batasMaks: BATAS_MAKS };
    }
  }
  return { jumlahPelanggaran: 0, terkunci: false, batasMaks: BATAS_MAKS };
}

/** Guru membuka kembali ujian yang terkunci karena pelanggaran, siswa mulai ulang dengan durasi penuh */
function guruBukaKuncPelanggaran(idHasil) {
  var sheet = _pastikanSheetHasilUjianBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === idHasil.toString()) {
      if (data[i][10].toString() !== "Terkunci") {
        return { status: "ERROR", message: "Ujian ini tidak dalam status terkunci." };
      }
      var waktuBaru = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      sheet.getRange(i + 1, 11).setValue("Sedang Mengerjakan"); // Status
      sheet.getRange(i + 1, 12).setNumberFormat("@").setValue(waktuBaru); // Waktu Mulai direset (durasi penuh lagi)
      sheet.getRange(i + 1, 13).setValue(""); // Waktu Selesai dikosongkan
      sheet.getRange(i + 1, 14).setValue(0); // Jumlah Pelanggaran direset
      return { status: "SUCCESS", message: "Ujian berhasil dibuka kembali. Siswa bisa melanjutkan dengan durasi penuh." };
    }
  }
  return { status: "ERROR", message: "Data hasil ujian tidak ditemukan." };
}

/** Rekap hasil satu sesi ujian: semua siswa di kelas peserta, lengkap status & nilai (guru) */
/**
 * JARING PENGAMAN: cari semua percobaan ujian di 1 sesi yang statusnya masih "Sedang Mengerjakan"
 * tapi waktunya SUDAH LEWAT batas (durasi habis / sesi sudah ditutup) — biasanya karena siswa
 * menutup browser/perangkat mati mendadak sehingga tidak sempat klik "Selesai" sendiri.
 * Nilai tetap dihitung otomatis dari jawaban terakhir yang ter-autosave.
 */
function _autoFinalisasiUjianTerlambat(idSesi) {
  var sesi = guruAmbilDetailSesi(idSesi);
  if (!sesi) return;

  var sheetHasil = _pastikanSheetHasilUjianBenar();
  var dataHasil = sheetHasil.getDataRange().getValues();
  var batasWaktuSesi = new Date(sesi.tglSelesai + "T" + (sesi.jamSelesai || "23:59") + ":00").getTime();
  var sekarang = new Date().getTime();
  var TOLERANSI_MS = 90 * 1000;
  var daftarTerlambat = [];

  for (var i = 1; i < dataHasil.length; i++) {
    if (dataHasil[i][1].toString() !== idSesi.toString()) continue;
    if (dataHasil[i][10].toString() !== "Sedang Mengerjakan") continue;
    var epochMulai = new Date(dataHasil[i][11].toString().replace(" ", "T")).getTime();
    var batasWaktuDurasi = epochMulai + (sesi.durasiMenit * 60000);
    var epochBatasAkhir = Math.min(batasWaktuDurasi, batasWaktuSesi);
    if (sekarang - epochBatasAkhir > TOLERANSI_MS) {
      daftarTerlambat.push({ idHasil: dataHasil[i][0].toString(), username: dataHasil[i][2].toString() });
    }
  }

  daftarTerlambat.forEach(function(item) {
    try { siswaSubmitUjian(item.idHasil, idSesi, item.username, null, false); } catch (e) { /* lewati kalau ada error, jangan sampai gagalkan seluruh rekap */ }
  });
}

function guruAmbilRekapHasilSesi(idSesi) {
  _autoFinalisasiUjianTerlambat(idSesi);
  var sesi = guruAmbilDetailSesi(idSesi);
  if (!sesi) return { status: "ERROR", message: "Sesi tidak ditemukan." };

  var sheetUser = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  var dataUser = sheetUser.getDataRange().getValues();
  var daftarSiswa = [];
  sesi.kelas.forEach(function(kls) {
    for (var u = 1; u < dataUser.length; u++) {
      if (dataUser[u][3].toString().trim() === "Siswa" && dataUser[u].length > 4 && dataUser[u][4].toString().trim().toUpperCase() === kls.toString().trim().toUpperCase()) {
        daftarSiswa.push({ username: dataUser[u][0].toString(), nama: dataUser[u][2].toString(), kelas: kls });
      }
    }
  });

  var sheetHasil = _pastikanSheetHasilUjianBenar();
  var dataHasil = sheetHasil.getDataRange().getValues();
  var petaHasil = {};
  for (var i = 1; i < dataHasil.length; i++) {
    if (dataHasil[i][1].toString().trim() === idSesi.toString().trim()) {
      var unameBaris = _normalisasiUsername(dataHasil[i][2]);
      var yangSudahAda = petaHasil[unameBaris];
      // Kalau ada baris duplikat (data lama dari sebelum diperbaiki), prioritaskan status FINAL
      // (Selesai/Terkunci) — jangan sampai baris "Sedang Mengerjakan" yang lebih baru menimpa hasil yang sudah final.
      if (yangSudahAda && (yangSudahAda[10].toString() === "Selesai" || yangSudahAda[10].toString() === "Terkunci")) continue;
      petaHasil[unameBaris] = dataHasil[i];
    }
  }

  var sheetBankSoal = _pastikanSheetBankSoalBenar();
  var dataSoal = sheetBankSoal.getDataRange().getValues();
  var mapTipeSoal = {};
  for (var s = 1; s < dataSoal.length; s++) { if (dataSoal[s][0]) mapTipeSoal[dataSoal[s][0].toString()] = dataSoal[s][6].toString(); }
  var adaSoalEssay = sesi.daftarIdSoal.some(function(id) { return mapTipeSoal[id] === "Essay"; });

  var hasil = daftarSiswa.map(function(sw) {
    var h = petaHasil[_normalisasiUsername(sw.username)];
    if (!h) return { username: sw.username, nama: sw.nama, kelas: sw.kelas, status: "Belum Mengerjakan", nilaiObjektif: "-", nilaiEssay: "-", nilaiAkhir: "-", idHasil: "" };
    return {
      username: sw.username, nama: sw.nama, kelas: sw.kelas, status: h[10].toString(),
      nilaiObjektif: h[7], nilaiEssay: h[8], nilaiAkhir: h[9],
      waktuMulai: h[11].toString(), waktuSelesai: h[12].toString(), jumlahPelanggaran: h[13], idHasil: h[0].toString(),
      catatanWaktu: h[14] ? h[14].toString() : ""
    };
  });
  hasil.sort(function(a, b) { return a.nama.localeCompare(b.nama); });

  return { status: "SUCCESS", sesi: sesi, daftarSiswa: hasil, adaSoalEssay: adaSoalEssay, jumlahSoal: sesi.daftarIdSoal.length };
}

/** Detail seluruh jawaban satu siswa untuk satu sesi ujian (dipakai guru untuk koreksi Essay) */
function guruAmbilDetailJawabanSiswa(idHasil) {
  var sheetJawaban = _pastikanSheetJawabanSiswaBenar();
  var dataJawaban = sheetJawaban.getDataRange().getValues();
  var sheetBankSoal = _pastikanSheetBankSoalBenar();
  var dataSoal = sheetBankSoal.getDataRange().getValues();
  var petaSoal = {};
  for (var s = 1; s < dataSoal.length; s++) { if (dataSoal[s][0]) petaSoal[dataSoal[s][0].toString()] = _baris2Soal(dataSoal[s]); }

  var hasil = [];
  for (var i = 1; i < dataJawaban.length; i++) {
    if (dataJawaban[i][0].toString() !== idHasil.toString()) continue;
    var soal = petaSoal[dataJawaban[i][3].toString()];
    if (!soal) continue;
    hasil.push({
      idSoal: soal.id, tipe: soal.tipe, pertanyaan: soal.pertanyaan, opsi: soal.opsi, kunci: soal.kunci, poin: soal.poin, gambarUrl: soal.gambarUrl || "",
      jawabanSiswa: dataJawaban[i][4].toString(), benar: dataJawaban[i][5].toString(),
      skorDidapat: Number(dataJawaban[i][6]) || 0, catatanGuru: dataJawaban[i][7] ? dataJawaban[i][7].toString() : ""
    });
  }
  return hasil;
}

/** Guru menyimpan nilai Essay (per soal) untuk satu siswa, lalu Nilai Akhir dihitung ulang otomatis */
/**
 * Review hasil pengerjaan untuk SISWA sendiri (benar/salah per soal + kunci jawaban).
 * Hanya diizinkan kalau: (1) idHasil memang milik usernameSiswa yang login, (2) status ujian sudah
 * final (Selesai/Terkunci), dan (3) guru mengaktifkan "Tampilkan Review" di pengaturan sesi.
 * Soal Essay tidak ditampilkan status benar/salah (karena sifatnya subjektif), hanya jawaban & skor kalau sudah dikoreksi.
 */
function siswaAmbilReviewHasil(idHasil, usernameSiswa) {
  var sheetHasil = _pastikanSheetHasilUjianBenar();
  var dataHasil = sheetHasil.getDataRange().getValues();
  var baris = null;
  for (var i = 1; i < dataHasil.length; i++) {
    if (dataHasil[i][0].toString() === idHasil.toString()) { baris = dataHasil[i]; break; }
  }
  if (!baris) return { status: "ERROR", message: "Data hasil ujian tidak ditemukan." };
  if (_normalisasiUsername(baris[2]) !== _normalisasiUsername(usernameSiswa)) return { status: "ERROR", message: "Anda tidak berhak melihat hasil ujian ini." };
  var status = baris[10].toString();
  if (status !== "Selesai" && status !== "Terkunci") return { status: "ERROR", message: "Ujian belum selesai dikerjakan." };

  var idSesi = baris[1].toString();
  var sesi = guruAmbilDetailSesi(idSesi);
  if (!sesi) return { status: "ERROR", message: "Sesi ujian tidak ditemukan." };
  if (!sesi.tampilkanReview) return { status: "ERROR", message: "Guru belum mengizinkan siswa melihat review hasil pengerjaan untuk ujian ini." };

  var jawabanDetail = guruAmbilDetailJawabanSiswa(idHasil);
  var HURUF_LABEL = ["A", "B", "C", "D", "E"];
  var soalUntukReview = jawabanDetail.map(function(j) {
    var opsiBerlabel = (j.opsi || []).map(function(teks, idx) { return { label: HURUF_LABEL[idx], teks: teks }; }).filter(function(o) { return o.teks !== ""; });
    return {
      tipe: j.tipe, pertanyaan: j.pertanyaan, opsi: opsiBerlabel, gambarUrl: j.gambarUrl,
      jawabanSiswa: j.jawabanSiswa,
      kunci: j.tipe === "Essay" ? "" : j.kunci,
      benar: j.tipe === "Essay" ? null : (j.benar === "TRUE"),
      poin: j.poin, skorDidapat: j.skorDidapat
    };
  });

  return {
    status: "SUCCESS", judul: sesi.judul, jenis: sesi.jenis, mapel: sesi.mapel,
    nilaiAkhir: sesi.tampilkanNilai ? baris[9] : null,
    soal: soalUntukReview
  };
}

function guruSimpanKoreksiEssay(idHasil, koreksiArray) {
  var sheetJawaban = _pastikanSheetJawabanSiswaBenar();
  var dataJawaban = sheetJawaban.getDataRange().getValues();

  (koreksiArray || []).forEach(function(k) {
    for (var i = 1; i < dataJawaban.length; i++) {
      if (dataJawaban[i][0].toString() === idHasil.toString() && dataJawaban[i][3].toString() === k.idSoal.toString()) {
        sheetJawaban.getRange(i + 1, 7, 1, 2).setValues([[Number(k.skor) || 0, k.catatan || ""]]);
        break;
      }
    }
  });

  var dataJawabanBaru = sheetJawaban.getDataRange().getValues();
  var sheetBankSoal = _pastikanSheetBankSoalBenar();
  var dataSoal = sheetBankSoal.getDataRange().getValues();
  var petaPoin = {};
  for (var s = 1; s < dataSoal.length; s++) { if (dataSoal[s][0]) petaPoin[dataSoal[s][0].toString()] = Number(dataSoal[s][14]) || 0; }

  var skorTotal = 0, skorMaksimal = 0;
  for (var j = 1; j < dataJawabanBaru.length; j++) {
    if (dataJawabanBaru[j][0].toString() !== idHasil.toString()) continue;
    skorMaksimal += petaPoin[dataJawabanBaru[j][3].toString()] || 0;
    skorTotal += Number(dataJawabanBaru[j][6]) || 0;
  }
  var nilaiAkhir = skorMaksimal > 0 ? Math.round((skorTotal / skorMaksimal) * 10000) / 100 : 0;

  var sheetHasil = _pastikanSheetHasilUjianBenar();
  var dataHasil = sheetHasil.getDataRange().getValues();
  for (var h = 1; h < dataHasil.length; h++) {
    if (dataHasil[h][0].toString() === idHasil.toString()) {
      sheetHasil.getRange(h + 1, 9, 1, 2).setValues([["Sudah Dinilai", nilaiAkhir]]);
      break;
    }
  }

  return { status: "SUCCESS", message: "Koreksi Essay berhasil disimpan.", nilaiAkhir: nilaiAkhir };
}

/** Rekap nilai ujian CBT milik satu siswa (dipakai di halaman Rekap Nilai siswa) */
function _siswaAmbilNilaiUjianSaya(usernameSiswa, kelasSiswa) {
  var sheetHasil = _pastikanSheetHasilUjianBenar();
  var dataHasil = sheetHasil.getDataRange().getValues();
  var sheetSesi = _pastikanSheetCbtSesiBenar();
  var dataSesi = sheetSesi.getDataRange().getValues();
  var petaSesi = {};
  for (var s = 1; s < dataSesi.length; s++) {
    if (!dataSesi[s][0]) continue;
    var sesiBaris = _baris2Sesi(dataSesi[s]);
    petaSesi[sesiBaris.id] = { judul: sesiBaris.judul, jenis: sesiBaris.jenis, mapel: sesiBaris.mapel, tampilkanNilai: sesiBaris.tampilkanNilai };
  }

  var hasil = [];
  for (var i = 1; i < dataHasil.length; i++) {
    if (!dataHasil[i][0]) continue;
    if (dataHasil[i][2].toString().trim() !== usernameSiswa.toString().trim()) continue;
    var status = dataHasil[i][10].toString();
    if (status !== "Selesai" && status !== "Terkunci") continue;
    var info = petaSesi[dataHasil[i][1].toString()] || { judul: "Ujian", jenis: "", mapel: "", tampilkanNilai: true };
    hasil.push({ judul: info.judul, jenis: info.jenis, mapel: info.mapel, nilaiAkhir: info.tampilkanNilai ? dataHasil[i][9] : "-", status: status });
  }
  return hasil;
}

// ==================================================================================
// ==================== MODUL GURU WALI (Permendikdasmen No. 11/2025) ====================
// ==================================================================================
// Modul baru, terpisah sepenuhnya dari modul lain. Murid direferensikan lewat username
// (konsisten dengan seluruh aplikasi), BUKAN sheet "Murid" terpisah.

// ---------- SHEET: GuruWali_Penugasan ----------
var HEADER_GW_PENUGASAN = ["ID Penugasan", "Guru Username", "Guru Nama", "Username Murid", "Nama Murid", "Kelas Saat Assign", "Tahun Ajaran", "Tanggal Mulai", "Tanggal Selesai", "Status"];

function _pastikanSheetGWPenugasanBenar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("GuruWali_Penugasan");
  if (!sheet) {
    sheet = ss.insertSheet("GuruWali_Penugasan");
    sheet.appendRow(HEADER_GW_PENUGASAN);
  }
  return sheet;
}

/** Tahun ajaran otomatis berdasarkan tanggal saat ini (Juli dianggap awal tahun ajaran baru) */
function _tahunAjaranSaatIni() {
  var now = new Date();
  var tahun = now.getFullYear();
  var bulan = now.getMonth() + 1; // 1-12
  return bulan >= 7 ? (tahun + "/" + (tahun + 1)) : ((tahun - 1) + "/" + tahun);
}

/**
 * Admin: assign murid ke guru wali. Bisa dipakai utk assign manual (daftar username spesifik)
 * ATAU assign 1 kelas penuh (kirim daftarUsernameMurid = semua siswa di kelas itu, ambil dulu
 * dari client lewat adminAmbilSiswaKelasUntukGuruWali). Tidak menduplikasi kalau murid itu
 * sudah aktif dengan guru wali yang SAMA; tapi kalau sudah aktif dengan guru wali LAIN,
 * penugasan lama otomatis diakhiri (Dipindahkan) sebelum yang baru dibuat.
 */
function adminAssignMuridKeGuruWali(guruUsername, guruNama, daftarMurid) {
  var sheet = _pastikanSheetGWPenugasanBenar();
  var data = sheet.getDataRange().getValues();
  var tahunAjaran = _tahunAjaranSaatIni();
  var hariIni = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var jumlahBaru = 0, jumlahDipindah = 0;

  daftarMurid.forEach(function(m) {
    var sudahAktifDenganGuruSama = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][9].toString() !== "Aktif") continue;
      if (!_samaUsername(data[i][3], m.username)) continue;
      if (data[i][1].toString().trim() === guruUsername.toString().trim()) {
        sudahAktifDenganGuruSama = true;
      } else {
        // Murid pindah dari guru wali lain -> akhiri penugasan lama
        sheet.getRange(i + 1, 9).setNumberFormat("@").setValue(hariIni);
        sheet.getRange(i + 1, 10).setValue("Dipindahkan");
        jumlahDipindah++;
      }
    }
    if (!sudahAktifDenganGuruSama) {
      var idBaru = _buatIdUnik("GWP");
      sheet.appendRow([idBaru, guruUsername, guruNama, m.username, m.nama, m.kelas || "", tahunAjaran, hariIni, "", "Aktif"]);
      var br = sheet.getLastRow();
      sheet.getRange(br, 4).setNumberFormat("@").setValue(m.username); // Kunci Username sbg teks (jaga NISN)
      sheet.getRange(br, 8).setNumberFormat("@").setValue(hariIni);
      jumlahBaru++;
    }
  });

  catatLogAktivitas(guruNama, "Assign Guru Wali", "Admin menugaskan " + jumlahBaru + " murid baru ke " + guruNama + (jumlahDipindah > 0 ? " (" + jumlahDipindah + " dipindahkan dari guru wali sebelumnya)" : ""));
  return { status: "SUCCESS", message: jumlahBaru + " murid berhasil ditugaskan ke " + guruNama + "." + (jumlahDipindah > 0 ? " " + jumlahDipindah + " murid dipindahkan dari guru wali sebelumnya." : "") };
}

/** Admin: akhiri 1 penugasan (murid lulus/pindah/ganti guru wali) — bukan hapus permanen, demi jejak riwayat */
function adminAkhiriPenugasanGuruWali(idPenugasan, statusBaru) {
  var sheet = _pastikanSheetGWPenugasanBenar();
  var data = sheet.getDataRange().getValues();
  var hariIni = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === idPenugasan.toString()) {
      sheet.getRange(i + 1, 9).setNumberFormat("@").setValue(hariIni);
      sheet.getRange(i + 1, 10).setValue(statusBaru || "Selesai");
      return { status: "SUCCESS", message: "Penugasan berhasil diakhiri." };
    }
  }
  return { status: "ERROR", message: "Data penugasan tidak ditemukan." };
}

/** Admin: rekap semua penugasan guru wali yang sedang aktif, dikelompokkan per guru */
function adminAmbilRekapGuruWali() {
  var sheet = _pastikanSheetGWPenugasanBenar();
  var data = sheet.getDataRange().getValues();
  var perGuru = {};

  for (var i = 1; i < data.length; i++) {
    if (data[i][9].toString() !== "Aktif") continue;
    var guruUsername = data[i][1].toString();
    if (!perGuru[guruUsername]) perGuru[guruUsername] = { guruUsername: guruUsername, guruNama: data[i][2].toString(), muridList: [] };
    perGuru[guruUsername].muridList.push({
      idPenugasan: data[i][0].toString(), username: data[i][3].toString(), nama: data[i][4].toString(),
      kelas: data[i][5].toString(), tahunAjaran: data[i][6].toString(), tanggalMulai: data[i][7].toString()
    });
  }

  var hasil = Object.values(perGuru);
  hasil.forEach(function(g) { g.muridList.sort(function(a, b) { return a.nama.localeCompare(b.nama); }); });
  hasil.sort(function(a, b) { return a.guruNama.localeCompare(b.guruNama); });
  return hasil;
}

/** Ambil daftar siswa 1 kelas (dipakai admin utk assign per-kelas ke guru wali) */
function adminAmbilSiswaKelasUntukGuruWali(kelas) {
  return guruAmbilDaftarSiswaKelas(kelas); // Reuse fungsi yang sudah ada, sudah terurut abjad
}

/** Guru: daftar murid dampingan aktif miliknya sendiri */
function guruAmbilMuridDampinganSaya(guruUsername) {
  var sheet = _pastikanSheetGWPenugasanBenar();
  var data = sheet.getDataRange().getValues();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() !== guruUsername.toString().trim()) continue;
    if (data[i][9].toString() !== "Aktif") continue;
    hasil.push({
      idPenugasan: data[i][0].toString(), username: data[i][3].toString(), nama: data[i][4].toString(),
      kelas: data[i][5].toString(), tahunAjaran: data[i][6].toString(), tanggalMulai: data[i][7].toString()
    });
  }
  hasil.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
  return hasil;
}

// ---------- SHEET: GuruWali_Jurnal ----------
var HEADER_GW_JURNAL = ["ID Jurnal", "Guru Username", "Guru Nama", "Username Murid", "Nama Murid", "Tanggal", "Kategori", "Topik Catatan", "Tindak Lanjut", "Status Tindak Lanjut", "Lampiran Link", "Waktu Dibuat"];

function _pastikanSheetGWJurnalBenar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("GuruWali_Jurnal");
  if (!sheet) {
    sheet = ss.insertSheet("GuruWali_Jurnal");
    sheet.appendRow(HEADER_GW_JURNAL);
  }
  return sheet;
}

/** Guru: catat sesi pendampingan baru */
function guruTambahJurnalPendampingan(guruUsername, guruNama, usernameMurid, namaMurid, tanggal, kategori, topikCatatan, tindakLanjut, lampiranLink) {
  if (!usernameMurid || !tanggal || !kategori || !topikCatatan) {
    return { status: "ERROR", message: "Murid, tanggal, kategori, dan catatan wajib diisi." };
  }
  var sheet = _pastikanSheetGWJurnalBenar();
  var id = _buatIdUnik("GWJ");
  var waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var statusTindakLanjut = tindakLanjut && tindakLanjut.trim() ? "Belum" : "-";

  sheet.appendRow([id, guruUsername, guruNama, usernameMurid, namaMurid, tanggal, kategori, topikCatatan, tindakLanjut || "", statusTindakLanjut, lampiranLink || "", waktu]);
  var br = sheet.getLastRow();
  sheet.getRange(br, 4).setNumberFormat("@").setValue(usernameMurid);
  sheet.getRange(br, 6).setNumberFormat("@").setValue(tanggal);
  sheet.getRange(br, 12).setNumberFormat("@").setValue(waktu);

  catatLogAktivitas(guruNama, "Jurnal Pendampingan", "Mencatat sesi " + kategori + " untuk " + namaMurid);
  return { status: "SUCCESS", message: "Jurnal pendampingan berhasil disimpan." };
}

function guruUpdateJurnalPendampingan(idJurnal, tanggal, kategori, topikCatatan, tindakLanjut, lampiranLink) {
  var sheet = _pastikanSheetGWJurnalBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === idJurnal.toString()) {
      sheet.getRange(i + 1, 6).setNumberFormat("@").setValue(tanggal);
      sheet.getRange(i + 1, 7).setValue(kategori);
      sheet.getRange(i + 1, 8).setValue(topikCatatan);
      sheet.getRange(i + 1, 9).setValue(tindakLanjut || "");
      sheet.getRange(i + 1, 11).setValue(lampiranLink || "");
      // Kalau tindak lanjut baru diisi (sebelumnya kosong "-"), set status jadi "Belum"
      if (tindakLanjut && tindakLanjut.trim() && data[i][9].toString() === "-") {
        sheet.getRange(i + 1, 10).setValue("Belum");
      }
      return { status: "SUCCESS", message: "Jurnal berhasil diperbarui." };
    }
  }
  return { status: "ERROR", message: "Jurnal tidak ditemukan." };
}

function guruUpdateStatusTindakLanjut(idJurnal, statusBaru) {
  var sheet = _pastikanSheetGWJurnalBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === idJurnal.toString()) {
      sheet.getRange(i + 1, 10).setValue(statusBaru);
      return { status: "SUCCESS", message: "Status tindak lanjut diperbarui menjadi \"" + statusBaru + "\"." };
    }
  }
  return { status: "ERROR", message: "Jurnal tidak ditemukan." };
}

function guruHapusJurnalPendampingan(idJurnal) {
  var sheet = _pastikanSheetGWJurnalBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0].toString() === idJurnal.toString()) {
      sheet.deleteRow(i + 1);
      return { status: "SUCCESS", message: "Jurnal berhasil dihapus." };
    }
  }
  return { status: "ERROR", message: "Jurnal tidak ditemukan." };
}

/** Semua jurnal milik guru wali ini (utk menu Jurnal Pendampingan, terbaru dulu) */
function guruAmbilSemuaJurnalSaya(guruUsername) {
  var sheet = _pastikanSheetGWJurnalBenar();
  var data = sheet.getDataRange().getValues();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() !== guruUsername.toString().trim()) continue;
    hasil.push(_baris2Jurnal(data[i]));
  }
  hasil.sort(function(a, b) { return b.tanggal.localeCompare(a.tanggal) || b.waktuDibuat.localeCompare(a.waktuDibuat); });
  return hasil;
}

/** Timeline jurnal utk 1 murid spesifik (utk menu Perkembangan Murid), terurut tanggal menaik (kronologis) */
function guruAmbilJurnalMurid(guruUsername, usernameMurid) {
  var sheet = _pastikanSheetGWJurnalBenar();
  var data = sheet.getDataRange().getValues();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() !== guruUsername.toString().trim()) continue;
    if (!_samaUsername(data[i][3], usernameMurid)) continue;
    hasil.push(_baris2Jurnal(data[i]));
  }
  hasil.sort(function(a, b) { return a.tanggal.localeCompare(b.tanggal) || a.waktuDibuat.localeCompare(b.waktuDibuat); });
  return hasil;
}

function _baris2Jurnal(row) {
  return {
    id: row[0].toString(), usernameMurid: row[3].toString(), namaMurid: row[4].toString(),
    tanggal: row[5].toString(), kategori: row[6].toString(), topikCatatan: row[7].toString(),
    tindakLanjut: row[8] ? row[8].toString() : "", statusTindakLanjut: row[9] ? row[9].toString() : "-",
    lampiranLink: row[10] ? row[10].toString() : "", waktuDibuat: row[11] ? row[11].toString() : ""
  };
}

// ---------- SHEET: GuruWali_Jadwal ----------
var HEADER_GW_JADWAL = ["ID Jadwal", "Guru Username", "Guru Nama", "Username Murid", "Nama Murid", "Tanggal Rencana", "Judul", "Status", "ID Jurnal Terkait"];

function _pastikanSheetGWJadwalBenar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("GuruWali_Jadwal");
  if (!sheet) {
    sheet = ss.insertSheet("GuruWali_Jadwal");
    sheet.appendRow(HEADER_GW_JADWAL);
  }
  return sheet;
}

/** usernameMurid boleh dikosongkan (string kosong) utk sesi kelompok/klasikal */
function guruTambahJadwalPendampingan(guruUsername, guruNama, usernameMurid, namaMurid, tanggalRencana, judul) {
  if (!tanggalRencana || !judul) return { status: "ERROR", message: "Tanggal rencana dan judul wajib diisi." };
  var sheet = _pastikanSheetGWJadwalBenar();
  var id = _buatIdUnik("GWD");
  sheet.appendRow([id, guruUsername, guruNama, usernameMurid || "", namaMurid || "(Sesi Kelompok)", tanggalRencana, judul, "Terjadwal", ""]);
  var br = sheet.getLastRow();
  if (usernameMurid) sheet.getRange(br, 4).setNumberFormat("@").setValue(usernameMurid);
  sheet.getRange(br, 6).setNumberFormat("@").setValue(tanggalRencana);
  return { status: "SUCCESS", message: "Jadwal berhasil ditambahkan." };
}

function guruUpdateStatusJadwal(idJadwal, statusBaru, idJurnalTerkait) {
  var sheet = _pastikanSheetGWJadwalBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === idJadwal.toString()) {
      sheet.getRange(i + 1, 8).setValue(statusBaru);
      if (idJurnalTerkait) sheet.getRange(i + 1, 9).setValue(idJurnalTerkait);
      return { status: "SUCCESS", message: "Status jadwal diperbarui." };
    }
  }
  return { status: "ERROR", message: "Jadwal tidak ditemukan." };
}

function guruHapusJadwalPendampingan(idJadwal) {
  var sheet = _pastikanSheetGWJadwalBenar();
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0].toString() === idJadwal.toString()) {
      sheet.deleteRow(i + 1);
      return { status: "SUCCESS", message: "Jadwal berhasil dihapus." };
    }
  }
  return { status: "ERROR", message: "Jadwal tidak ditemukan." };
}

function guruAmbilJadwalSaya(guruUsername) {
  var sheet = _pastikanSheetGWJadwalBenar();
  var data = sheet.getDataRange().getValues();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() !== guruUsername.toString().trim()) continue;
    hasil.push({
      id: data[i][0].toString(), usernameMurid: data[i][3].toString(), namaMurid: data[i][4].toString(),
      tanggalRencana: data[i][5].toString(), judul: data[i][6].toString(),
      status: data[i][7].toString(), idJurnalTerkait: data[i][8] ? data[i][8].toString() : ""
    });
  }
  hasil.sort(function(a, b) { return a.tanggalRencana.localeCompare(b.tanggalRencana); });
  return hasil;
}

// ---------- DASHBOARD ----------
function guruAmbilDashboardGuruWali(guruUsername) {
  var muridDampingan = guruAmbilMuridDampinganSaya(guruUsername);
  var semuaJurnal = guruAmbilSemuaJurnalSaya(guruUsername);
  var semuaJadwal = guruAmbilJadwalSaya(guruUsername);
  var hariIni = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var awalBulanIni = hariIni.substring(0, 7); // "yyyy-MM"

  var sesiBulanIni = semuaJurnal.filter(function(j) { return j.tanggal.substring(0, 7) === awalBulanIni; }).length;
  var tindakLanjutPending = semuaJurnal.filter(function(j) { return j.statusTindakLanjut === "Belum" || j.statusTindakLanjut === "Proses"; });
  var jadwalTerlambat = semuaJadwal.filter(function(j) { return j.status === "Terjadwal" && j.tanggalRencana < hariIni; });
  var jadwalMendatang = semuaJadwal.filter(function(j) { return j.status === "Terjadwal" && j.tanggalRencana >= hariIni; }).slice(0, 5);

  return {
    jumlahMurid: muridDampingan.length,
    sesiBulanIni: sesiBulanIni,
    tindakLanjutPending: tindakLanjutPending.slice(0, 10),
    jumlahTindakLanjutPending: tindakLanjutPending.length,
    jadwalTerlambat: jadwalTerlambat,
    jadwalMendatang: jadwalMendatang
  };
}
