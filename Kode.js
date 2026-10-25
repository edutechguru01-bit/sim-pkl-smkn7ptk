/**
 * SIM PKL - SMKN 7 Pontianak
 * Backend Script (Code.gs)
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('SIM PKL SMKN 7 Pontianak')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// === HELPER FORMAT TANGGAL & NO HP ===
function formatTanggalIndo(dateInput) {
  if (!dateInput) return "-";
  var d = new Date(dateInput);
  if (isNaN(d.getTime())) return dateInput.toString();
  var bulanIndo = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
  ];
  var tgl = d.getDate();
  var bulan = bulanIndo[d.getMonth()];
  var tahun = d.getFullYear();
  return (tgl < 10 ? '0' + tgl : tgl) + " " + bulan + " " + tahun;
}

function formatNoHp(noHp) {
  if (!noHp) return "-";
  var str = noHp.toString().trim();
  if (str === "") return "-";
  if (!str.startsWith("0") && !str.startsWith("+")) {
    str = "0" + str;
  }
  return str;
}

// === BACA DATA DARI SHEET ===
function getSheetData(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    result.push(obj);
  }
  return result;
}

// === AUTENTIKASI LOGIN ===
function loginUser(email, password) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) return { success: false, message: 'Sheet Users tidak ditemukan.' };

  var data = sheet.getDataRange().getValues();
  var inputEmail = email.toString().trim().toLowerCase();
  var inputPassword = password.toString().trim();

  for (var i = 1; i < data.length; i++) {
    var userEmail = data[i][1].toString().trim().toLowerCase();
    var userPassword = data[i][6] ? data[i][6].toString().trim() : ""; 

    if (userEmail === inputEmail) {
      if (userPassword === inputPassword) {
        return {
          success: true,
          user: {
            userId: data[i][0],
            email: data[i][1],
            nama: data[i][2],
            role: data[i][3],
            jurusan: data[i][4],
            kelas: data[i][5],
            password: data[i][6],
            nisn: data[i][7],
            jenisKelamin: data[i][8],
            alamatRumah: data[i][9],
            noHp: formatNoHp(data[i][10]),
            namaOrtu: data[i][11] || "-"
          }
        };
      } else {
        return { success: false, message: 'Password salah.' };
      }
    }
  }
  return { success: false, message: 'Email tidak terdaftar.' };
}

// === UPDATE AKUN / PROFIL ===
function updateProfilSiswa(dataProfil) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sheet) return { status: "Error", message: "Sheet Users tidak ditemukan." };

    var data = sheet.getDataRange().getValues();
    var userIdStr = dataProfil.userId.toString().trim();
    var formattedHp = formatNoHp(dataProfil.noHp);

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === userIdStr) {
        sheet.getRange(i + 1, 2).setValue(dataProfil.email);
        sheet.getRange(i + 1, 7).setValue(dataProfil.password);
        sheet.getRange(i + 1, 9).setValue(dataProfil.jenisKelamin);
        sheet.getRange(i + 1, 10).setValue(dataProfil.alamatRumah);
        sheet.getRange(i + 1, 11).setValue("'" + formattedHp);
        if (dataProfil.namaOrtu) {
          sheet.getRange(i + 1, 12).setValue(dataProfil.namaOrtu);
        }

        var updatedUser = {
          userId: data[i][0],
          email: dataProfil.email,
          nama: data[i][2],
          role: data[i][3],
          jurusan: data[i][4],
          kelas: data[i][5],
          password: dataProfil.password,
          nisn: data[i][7],
          jenisKelamin: dataProfil.jenisKelamin,
          alamatRumah: dataProfil.alamatRumah,
          noHp: formattedHp,
          namaOrtu: dataProfil.namaOrtu || data[i][11] || "-"
        };

        return { status: "Success", message: "Akun berhasil diperbarui.", user: updatedUser };
      }
    }
    return { status: "Error", message: "Pengguna tidak ditemukan." };
  } catch (error) {
    return { status: "Error", message: error.message };
  }
}

// === GELOMBANG PKL ===
function getGelombangList() {
  return getSheetData('Gelombang');
}

function simpanGelombang(formData, userRole) {
  try {
    if (userRole !== 'Admin' && userRole !== 'Super Admin' && userRole !== 'Panitia PKL') {
      return { status: "Error", message: "Akses Ditolak." };
    }
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gelombang');
    var id = "GLB-" + new Date().getTime();
    sheet.appendRow([
      id, formData.nama, formData.tglBuka, formData.tglTutup,
      formData.tglMulai, formData.tglSelesai, formData.targetJurusan,
      formData.targetTingkat, formData.statusAktif || "Aktif"
    ]);
    return { status: "Success", message: "Gelombang PKL berhasil disimpan." };
  } catch (error) {
    return { status: "Error", message: error.message };
  }
}

// === TEMPAT PKL & KUOTA ===
function getTempatPKLList(jurusanSiswa) {
  var data = getSheetData('Tempat_PKL');
  data = data.map(function(t, index) {
    t.tempat_id = (t.tempat_id !== undefined && t.tempat_id !== "") ? t.tempat_id.toString() : (index + 1).toString();
    return t;
  });

  if (!jurusanSiswa) return data;

  var keyJurusan = "kuota_" + jurusanSiswa.toString().toLowerCase().trim();
  return data.filter(function(t) {
    var kuota = t[keyJurusan] || t[keyJurusan.toUpperCase()] || 0;
    return parseInt(kuota) > 0;
  });
}

function simpanTempatPKL(formData) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tempat_PKL');
    var id = "TMP-" + new Date().getTime();
    sheet.appendRow([
      id, formData.namaInstansi, formData.alamat, "'" + formatNoHp(formData.kontak),
      formData.kuotaRPL || 0, formData.kuotaTKJ || 0, formData.kuotaAKL || 0,
      formData.kuotaTLAS || 0, formData.kuotaTSM || 0, formData.kuotaDKV || 0
    ]);
    return { status: "Success", message: "Master Tempat PKL berhasil disimpan." };
  } catch (error) {
    return { status: "Error", message: error.message };
  }
}

function updateTempatPKL(formData) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tempat_PKL');
    if (!sheet) return { status: "Error", message: "Sheet Tempat_PKL tidak ditemukan." };

    var data = sheet.getDataRange().getValues();
    var targetIdStr = formData.tempatId.toString().trim();

    for (var i = 1; i < data.length; i++) {
      var idInSheet = data[i][0] ? data[i][0].toString().trim() : (i).toString();
      if (idInSheet === targetIdStr) {
        sheet.getRange(i + 1, 2).setValue(formData.namaInstansi);
        sheet.getRange(i + 1, 3).setValue(formData.alamat);
        sheet.getRange(i + 1, 4).setValue("'" + formatNoHp(formData.kontak));
        sheet.getRange(i + 1, 5).setValue(formData.kuotaRPL || 0);
        sheet.getRange(i + 1, 6).setValue(formData.kuotaTKJ || 0);
        sheet.getRange(i + 1, 7).setValue(formData.kuotaAKL || 0);
        sheet.getRange(i + 1, 8).setValue(formData.kuotaTLAS || 0);
        sheet.getRange(i + 1, 9).setValue(formData.kuotaTSM || 0);
        sheet.getRange(i + 1, 10).setValue(formData.kuotaDKV || 0);

        return { status: "Success", message: "Data Tempat PKL & Kuota berhasil diperbarui." };
      }
    }
    return { status: "Error", message: "Tempat PKL tidak ditemukan." };
  } catch (error) {
    return { status: "Error", message: error.message };
  }
}

function hapusTempatPKL(tempatId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tempat_PKL');
    if (!sheet) return { status: "Error", message: "Sheet Tempat_PKL tidak ditemukan." };

    var data = sheet.getDataRange().getValues();
    var targetIdStr = tempatId.toString().trim();

    for (var i = 1; i < data.length; i++) {
      var idInSheet = data[i][0] ? data[i][0].toString().trim() : (i).toString();
      if (idInSheet === targetIdStr) {
        sheet.deleteRow(i + 1);
        return { status: "Success", message: "Tempat PKL berhasil dihapus." };
      }
    }
    return { status: "Error", message: "Data Tempat PKL tidak ditemukan." };
  } catch (error) {
    return { status: "Error", message: error.message };
  }
}

// === DETAIL LOKASI & SISWA TERDAFTAR ===
function getDetailTempatSiswa(tempatId, jurusanSiswa) {
  var tempatList = getSheetData('Tempat_PKL');
  var pengajuanList = getSheetData('Pengajuan_PKL');
  var userList = getSheetData('Users');

  if (!tempatId) return null;
  var targetIdStr = tempatId.toString().trim();

  var tempat = tempatList.find(function(t, index) { 
    var idInSheet = (t.tempat_id !== undefined && t.tempat_id !== "") ? t.tempat_id.toString().trim() : (index + 1).toString();
    return idInSheet === targetIdStr || (t.nama_instansi && t.nama_instansi.toString().trim().toLowerCase() === targetIdStr.toLowerCase()); 
  });

  if (!tempat) return null;

  var keyJurusan = "kuota_" + (jurusanSiswa ? jurusanSiswa.toString().toLowerCase().trim() : "");
  var kuotaMaksimal = parseInt(tempat[keyJurusan] || tempat[keyJurusan.toUpperCase()] || 0);

  var pendaftar = [];
  var addedSiswaMap = {};

  for (var i = 0; i < pengajuanList.length; i++) {
    var p = pengajuanList[i];
    var pTempat = p.tempat_pkl ? p.tempat_pkl.toString().trim() : (p.tempat_id ? p.tempat_id.toString().trim() : "");

    if (pTempat === targetIdStr || pTempat === tempat.tempat_id.toString().trim() || pTempat.toLowerCase() === tempat.nama_instansi.toString().trim().toLowerCase()) {
      var namaSiswa = p.nama_siswa || p.siswa_id;
      if (!namaSiswa) continue;

      var sKey = namaSiswa.toString().trim().toLowerCase();
      if (addedSiswaMap[sKey]) continue;
      addedSiswaMap[sKey] = true;

      var user = userList.find(function(u) { 
        return (u.nama && u.nama.toString().trim().toLowerCase() === sKey) ||
               (u.user_id && u.user_id.toString().trim().toLowerCase() === sKey); 
      });

      var rawHp = user ? (user.no_hp || user.noHp || '-') : '-';

      pendaftar.push({
        pengajuanId: p.pengajuan_id,
        nama: user ? user.nama : namaSiswa,
        jurusan: user ? user.jurusan : '-',
        kelas: user ? user.kelas : '-',
        jenisKelamin: user ? (user.jenis_kelamin || user.jenisKelamin || '-') : '-',
        nisn: user ? (user.nisn || user['nip/nisn'] || '-') : '-',
        alamatRumah: user ? (user.alamat_rumah || user.alamatRumah || '-') : '-',
        noHp: formatNoHp(rawHp),
        status: p.status_verifikasi
      });
    }
  }

  return {
    tempat: tempat,
    kuotaMaksimal: kuotaMaksimal,
    terisi: pendaftar.length,
    sisaKuota: kuotaMaksimal - pendaftar.length,
    pendaftar: pendaftar
  };
}

// === PENGAJUAN TEMPAT PKL SISWA ===
function getPengajuanListRich(jurusanKaprog) {
  var pengajuanList = getSheetData('Pengajuan_PKL');
  var userList = getSheetData('Users');
  var tempatList = getSheetData('Tempat_PKL');

  var targetJurusan = jurusanKaprog ? jurusanKaprog.toString().toLowerCase().trim() : null;

  var result = [];
  var addedSiswaMap = {};

  for (var i = 0; i < pengajuanList.length; i++) {
    var p = pengajuanList[i];
    var namaSiswaStr = p.nama_siswa ? p.nama_siswa.toString().trim() : (p.siswa_id ? p.siswa_id.toString().trim() : "");
    if (!namaSiswaStr) continue;

    var sKey = namaSiswaStr.toLowerCase();
    if (addedSiswaMap[sKey]) continue;
    addedSiswaMap[sKey] = true;

    var user = userList.find(function(u) {
      return (u.nama && u.nama.toString().trim().toLowerCase() === sKey) ||
             (u.user_id && u.user_id.toString().trim().toLowerCase() === sKey);
    });

    if (targetJurusan && (!user || !user.jurusan || user.jurusan.toString().toLowerCase().trim() !== targetJurusan)) {
      continue;
    }

    var tempatStr = p.tempat_pkl ? p.tempat_pkl.toString().trim() : (p.tempat_id ? p.tempat_id.toString().trim() : "");
    var tempat = tempatList.find(function(t, idx) {
      var idInSheet = (t.tempat_id !== undefined && t.tempat_id !== "") ? t.tempat_id.toString().trim() : (idx + 1).toString();
      return idInSheet === tempatStr || (t.nama_instansi && t.nama_instansi.toString().trim().toLowerCase() === tempatStr.toLowerCase());
    });

    var rawHp = user ? (user.no_hp || user.noHp || '-') : '-';

    result.push({
      pengajuanId: p.pengajuan_id,
      namaSiswa: user ? user.nama : namaSiswaStr,
      kelasSiswa: user ? (user.kelas || '-') : '-',
      jurusanSiswa: user ? user.jurusan : '-',
      noHp: formatNoHp(rawHp),
      namaTempat: tempat ? tempat.nama_instansi : tempatStr,
      alamatTempat: tempat ? tempat.alamat : '-',
      status: p.status_verifikasi
    });
  }
  return result;
}

// === DAFTAR PENGAJUAN DIKELOMPOKKAN PER TEMPAT ===
function getGroupedPengajuanForSurat(jurusanFilter) {
  try {
    var tempatList = getSheetData('Tempat_PKL');
    var pengajuanRich = getPengajuanListRich(jurusanFilter);
    var penugasanList = getSheetData('Penugasan_Pembimbing');
    var userList = getSheetData('Users');

    if (!tempatList || tempatList.length === 0) return [];

    var result = [];

    for (var i = 0; i < tempatList.length; i++) {
      var tmpt = tempatList[i];
      var tId = tmpt.tempat_id !== undefined && tmpt.tempat_id !== "" ? tmpt.tempat_id.toString().trim() : (i + 1).toString();
      var tNama = tmpt.nama_instansi ? tmpt.nama_instansi.toString().trim() : "";

      var totalKuota = 0;
      if (jurusanFilter) {
        var kKey = "kuota_" + jurusanFilter.toString().toLowerCase().trim();
        totalKuota = parseInt(tmpt[kKey] || tmpt[kKey.toUpperCase()] || 0);
        if (totalKuota === 0) continue;
      } else {
        totalKuota = parseInt(tmpt.kuota_rpl || 0) + parseInt(tmpt.kuota_tkj || 0) + 
                     parseInt(tmpt.kuota_akl || 0) + parseInt(tmpt.kuota_tlas || 0) + 
                     parseInt(tmpt.kuota_tsm || 0) + parseInt(tmpt.kuota_dkv || 0);
      }

      var guruPmbList = [];
      for (var k = 0; k < penugasanList.length; k++) {
        var p = penugasanList[k];
        var pTempat = p.tempat_pkl ? p.tempat_pkl.toString().trim() : (p.tempat_id ? p.tempat_id.toString().trim() : "");
        var pGuru = p.nama_guru ? p.nama_guru.toString().trim() : (p.guru_id ? p.guru_id.toString().trim() : "");

        if (pTempat === tId || pTempat.toLowerCase() === tNama.toLowerCase()) {
          var gObj = userList.find(function(u) {
            return (u.user_id && u.user_id.toString().trim() === pGuru) ||
                   (u.email && u.email.toString().trim().toLowerCase() === pGuru.toLowerCase()) ||
                   (u.nama && u.nama.toString().trim().toLowerCase() === pGuru.toLowerCase());
          });
          var namaFix = gObj ? gObj.nama : pGuru;
          if (namaFix && guruPmbList.indexOf(namaFix) === -1) {
            guruPmbList.push(namaFix);
          }
        }
      }

      var siswaSesuai = pengajuanRich.filter(function(item) {
        return item.namaTempat.toLowerCase() === tNama.toLowerCase();
      });

      result.push({
        tempatId: tId,
        namaTempat: tNama,
        alamatTempat: tmpt.alamat || '-',
        kontakTempat: formatNoHp(tmpt.kontak_person || tmpt.kontak),
        satuJurusan: siswaSesuai.length > 0 ? siswaSesuai[0].jurusanSiswa : (jurusanFilter || '-'),
        kuotaMaksimal: totalKuota,
        pembimbing: guruPmbList.length > 0 ? guruPmbList.join(', ') : 'Guru Pembimbing belum ditambahkan',
        siswaList: siswaSesuai
      });
    }

    return result;
  } catch (err) {
    Logger.log("Error getGroupedPengajuanForSurat: " + err.message);
    return [];
  }
}

function getStatusPengajuanSiswa(namaSiswa) {
  var list = getSheetData('Pengajuan_PKL');
  for (var i = 0; i < list.length; i++) {
    var pSiswa = list[i].nama_siswa || list[i].siswa_id || "";
    if (pSiswa.toString().trim().toLowerCase() === namaSiswa.toString().trim().toLowerCase()) {
      return list[i];
    }
  }
  return null;
}

function ajukanTempatPKL(dataPengajuan, jurusanSiswa, namaSiswa) {
  try {
    var detail = getDetailTempatSiswa(dataPengajuan.tempatId, jurusanSiswa);
    if (detail && detail.sisaKuota <= 0) {
      return { status: "Error", message: "Gagal: Kuota tempat PKL ini sudah penuh!" };
    }

    var tempatList = getSheetData('Tempat_PKL');
    var targetTempat = tempatList.find(function(t, idx) {
      var idInSheet = (t.tempat_id !== undefined && t.tempat_id !== "") ? t.tempat_id.toString().trim() : (idx + 1).toString();
      return idInSheet === dataPengajuan.tempatId.toString().trim();
    });
    var namaInstansi = targetTempat ? targetTempat.nama_instansi : dataPengajuan.tempatId;

    var gelombangList = getSheetData('Gelombang');
    var gelombangIdActive = (gelombangList.length > 0) ? gelombangList[0].gelombang_id : "GLB-DEFAULT";

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pengajuan_PKL');
    var id = "PGJ-" + new Date().getTime();
    
    sheet.appendRow([
      id,
      namaSiswa,
      gelombangIdActive,
      namaInstansi,
      "Pending",
      "",
      ""
    ]);
    return { status: "Success", message: "Pengajuan tempat PKL berhasil dikirim." };
  } catch (error) {
    return { status: "Error", message: error.message };
  }
}

function verifikasiPengajuan(pengajuanId, status) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pengajuan_PKL');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === pengajuanId) {
        sheet.getRange(i + 1, 5).setValue(status);
        return { status: "Success", message: "Status verifikasi diperbarui ke " + status };
      }
    }
    return { status: "Error", message: "Data tidak ditemukan." };
  } catch (error) {
    return { status: "Error", message: error.message };
  }
}

function hapusPengajuan(pengajuanId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pengajuan_PKL');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === pengajuanId) {
        sheet.deleteRow(i + 1);
        return { status: "Success", message: "Data pengajuan berhasil dihapus." };
      }
    }
    return { status: "Error", message: "Data pengajuan tidak ditemukan." };
  } catch (error) {
    return { status: "Error", message: error.message };
  }
}

// === CETAK SURAT PERMOHONAN PKL ===
function generateSuratPermohonanHTML(tempatId, jurusanSiswa) {
  var detail = getDetailTempatSiswa(tempatId, jurusanSiswa);
  var gelombangList = getSheetData('Gelombang');
  var gel = (gelombangList.length > 0) ? gelombangList[0] : {};

  var tglMulai = formatTanggalIndo(gel.tgl_mulai_pkl || "2027-01-11");
  var tglSelesai = formatTanggalIndo(gel.tgl_selesai_pkl || "2027-03-31");
  var todayStr = formatTanggalIndo(new Date());

  var urlLogoProv = "https://lh3.googleusercontent.com/d/1KhMw-xuuaHe6ZC3BAisB8S9GGoRSa8VW";
  var urlLogoSMK = "https://lh3.googleusercontent.com/d/1kKch_ISvsqp4mX7PXDiTKG1DBwRW9F5c";

  var rowsHtml = "";
  if (detail && detail.pendaftar) {
    for (var i = 0; i < detail.pendaftar.length; i++) {
      var p = detail.pendaftar[i];
      rowsHtml += `
        <tr>
          <td style="text-align:center; padding: 5px; border: 1px solid #000;">${i + 1}</td>
          <td style="padding: 5px; border: 1px solid #000;">${p.nama}</td>
          <td style="padding: 5px; border: 1px solid #000; text-align:center;">${p.jenisKelamin || '-'}</td>
          <td style="padding: 5px; border: 1px solid #000; text-align:center;">${p.nisn || '-'}</td>
          <td style="padding: 5px; border: 1px solid #000;">${p.alamatRumah || '-'}</td>
          <td style="padding: 5px; border: 1px solid #000; text-align:center; white-space:nowrap;">${formatNoHp(p.noHp)}</td>
          <td style="padding: 5px; border: 1px solid #000; text-align:center; white-space:nowrap;">${p.kelas || '-'}</td>
        </tr>
      `;
    }
  }

  var namaInstansi = (detail && detail.tempat) ? detail.tempat.nama_instansi : tempatId;
  var alamatInstansi = (detail && detail.tempat) ? detail.tempat.alamat : '-';

  return `
    <!DOCTYPE html>
    <html lang="id">
    <head>
      <meta charset="utf-8">
      <title>Surat Permohonan PKL - ${namaInstansi}</title>
      <style>
        @page { size: F4; margin: 15mm; }
        body { font-family: 'Times New Roman', Times, serif; font-size: 12pt; line-height: 1.3; margin: 0; padding: 10px; background-color: #f8f9fa; }
        .surat-container { background: #fff; padding: 25px; max-width: 800px; margin: 0 auto; border: 1px solid #ccc; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
        .no-print { margin-bottom: 20px; text-align: right; }
        .btn-print { background: #0d6efd; color: white; border: none; padding: 8px 16px; border-radius: 4px; font-weight: bold; cursor: pointer; margin-left: 5px; }
        .btn-word { background: #198754; color: white; border: none; padding: 8px 16px; border-radius: 4px; font-weight: bold; cursor: pointer; margin-left: 5px; }
        .kop-container { border-bottom: 3px double #000; padding-bottom: 8px; margin-bottom: 15px; }
        .header-table { width: 100%; border-collapse: collapse; }
        .header-table td { vertical-align: middle; }
        .text-header { text-align: center; }
        .text-header h3 { margin: 0; font-size: 13pt; font-weight: bold; text-transform: uppercase; }
        .text-header h2 { margin: 0; font-size: 15pt; font-weight: bold; text-transform: uppercase; }
        .text-header p { margin: 2px 0 0 0; font-size: 9.5pt; }
        table.data-table { width: 100%; border-collapse: collapse; margin-top: 12px; margin-bottom: 12px; }
        table.data-table th { background-color: #f2f2f2; border: 1px solid #000; padding: 6px; text-align: center; font-size: 10pt; font-weight: bold; }
        table.data-table td { border: 1px solid #000; padding: 5px; font-size: 9.5pt; }
        .col-kelas { white-space: nowrap; width: 1%; }
        
        @media print {
          body { background-color: #fff; padding: 0; }
          .surat-container { border: none; box-shadow: none; padding: 0; max-width: 100%; }
          .no-print { display: none !important; }
        }
      </style>
      <script>
        function downloadWord() {
          var header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' "+
            "xmlns:w='urn:schemas-microsoft-com:office:word' "+
            "xmlns='http://www.w3.org/TR/REC-html40'>"+
            "<head><meta charset='utf-8'><title>Surat Permohonan PKL</title></head><body>";
          var footer = "</body></html>";
          var sourceHTML = header + document.getElementById("surat-content").innerHTML + footer;
          
          var source = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(sourceHTML);
          var fileDownload = document.createElement("a");
          document.body.appendChild(fileDownload);
          fileDownload.href = source;
          fileDownload.download = 'Surat_Permohonan_PKL_${namaInstansi.replace(/[^a-zA-Z0-9]/g, '_')}.doc';
          fileDownload.click();
          document.body.removeChild(fileDownload);
        }
      </script>
    </head>
    <body>
      <div class="surat-container">
        <div class="no-print">
          <button class="btn-word" onclick="downloadWord()">📄 Unduh Word (.doc)</button>
          <button class="btn-print" onclick="window.print()">🖨️ Cetak / Simpan PDF</button>
        </div>

        <div id="surat-content">
          <div class="kop-container">
            <table class="header-table">
              <tr>
                <td style="width: 15%; text-align: left;">
                  <img src="${urlLogoProv}" style="width: 75px; height: auto;" alt="Logo Kalbar">
                </td>
                <td style="width: 70%;" class="text-header">
                  <h3>PEMERINTAH PROVINSI KALIMANTAN BARAT</h3>
                  <h2>SMK NEGERI 7 PONTIANAK</h2>
                  <p>Jalan Tanjung Raya II Kec. Pontianak Timur, Prov. Kalimantan Barat 78232<br>
                  Website: smkn7ptk.sch.id | Email: surat@smkn7ptk.sch.id | CP: 08115733755 (WA Only)</p>
                </td>
                <td style="width: 15%; text-align: right;">
                  <img src="${urlLogoSMK}" style="width: 80px; height: auto;" alt="Logo SMKN 7">
                </td>
              </tr>
            </table>
          </div>

          <table style="width: 100%; font-size: 10.5pt; margin-bottom: 12px;">
            <tr><td style="width: 12%;">Nomor</td><td style="width: 2%;">:</td><td>400.3.8 / . . . . . . / SMKN7PTK</td></tr>
            <tr><td>Hal</td><td>:</td><td><strong>Permohonan Praktek Kerja Lapangan (PKL)</strong></td></tr>
            <tr><td>Lampiran</td><td>:</td><td>-</td></tr>
          </table>

          <p style="font-size: 11pt; margin-bottom: 10px;">
            Kepada Yth.<br>
            <strong>Pimpinan ${namaInstansi}</strong><br>${alamatInstansi}<br>
          </p>

          <p style="font-size: 11pt;">Dengan Hormat,</p>
          <p style="text-align: justify; text-indent: 30px; font-size: 11pt; line-height: 1.4;">
            Dalam rangka pelaksanaan Pendidikan Vokasi terkait dengan program <em>link and match</em> guna meningkatkan kompetensi, peserta didik diwajibkan untuk melaksanakan Praktik Kerja Lapangan (PKL). Oleh karena itu kami mengajukan permohonan kepada Bapak/Ibu Pimpinan Perusahaan/Instansi agar dapat menerima peserta didik kami sebagai berikut:
          </p>

          <table class="data-table">
            <thead>
              <tr>
                <th style="width: 4%;">No</th>
                <th style="width: 25%;">Nama</th>
                <th style="width: 6%;">L/P</th>
                <th style="width: 13%;">NISN</th>
                <th>Alamat Rumah</th>
                <th style="width: 15%;">No HP</th>
                <th class="col-kelas">Kelas</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml}
            </tbody>
          </table>

          <p style="text-align: justify; font-size: 11pt; line-height: 1.4;">
            Melaksanakan PKL pada bagian yang sesuai di Perusahaan/Instansi yang Bapak/Ibu pimpin. Adapun pelaksanaan PKL kami rencanakan pada tanggal <strong>${tglMulai}</strong> sampai dengan <strong>${tglSelesai}</strong>.
          </p>
          <p style="text-align: justify; font-size: 11pt; line-height: 1.4;">
            Demikian surat permohonan ini kami ajukan, atas perhatian dan kerjasamanya Bapak/Ibu kami mengucapkan terima kasih.
          </p>

          <br>
          <table style="width: 100%;">
            <tr>
              <td style="width: 50%;"></td>
              <td style="text-align: center; font-size: 11pt;">
                Pontianak, ${todayStr}<br>
                a.n Kepala SMK Negeri 7 Pontianak<br>
                WAKA HUMAS<br><br><br><br><br>
                <strong>ERWIN ARIYANTA, ST</strong><br>
                NIP. 19750415 200604 1 005
              </td>
            </tr>
          </table>
        </div>
      </div>
    </body>
    </html>
  `;
}

// === FITUR PANITIA: CETAK AMPLOP SURAT (UKURAN DL ENVELOPE) ===
function generateAmplopSuratHTML(tempatId, jurusanSiswa) {
  var detail = getDetailTempatSiswa(tempatId, jurusanSiswa);
  var urlLogoProv = "https://lh3.googleusercontent.com/d/1KhMw-xuuaHe6ZC3BAisB8S9GGoRSa8VW";
  var urlLogoSMK = "https://lh3.googleusercontent.com/d/1kKch_ISvsqp4mX7PXDiTKG1DBwRW9F5c";

  var namaInstansi = (detail && detail.tempat) ? detail.tempat.nama_instansi : tempatId;
  var alamatInstansi = (detail && detail.tempat) ? detail.tempat.alamat : 'Pontianak';

  return `
    <!DOCTYPE html>
    <html lang="id">
    <head>
      <meta charset="utf-8">
      <title>Cetak Amplop Surat DL - ${namaInstansi}</title>
      <style>
        /* Standar Ukuran DL Envelope: 220mm x 110mm */
        @page { 
          size: 220mm 110mm landscape; 
          margin: 0; 
        }
        
        * {
          box-sizing: border-box;
        }

        body { 
          font-family: 'Times New Roman', Times, serif; 
          margin: 0; 
          padding: 10px; 
          background-color: #f8f9fa; 
        }

        .amplop-container { 
          background: #fff; 
          width: 220mm; 
          height: 110mm; 
          padding: 10mm 15mm 8mm 15mm; 
          margin: 0 auto; 
          border: 1px solid #ccc; 
          box-shadow: 0 0 10px rgba(0,0,0,0.1); 
          position: relative; 
          overflow: hidden;
        }

        .no-print { 
          margin-bottom: 15px; 
          text-align: right; 
          width: 220mm;
          margin-left: auto;
          margin-right: auto;
        }

        .btn-print { 
          background: #0d6efd; 
          color: white; 
          border: none; 
          padding: 8px 16px; 
          border-radius: 4px; 
          font-weight: bold; 
          cursor: pointer; 
          margin-left: 5px; 
        }

        .btn-word { 
          background: #198754; 
          color: white; 
          border: none; 
          padding: 8px 16px; 
          border-radius: 4px; 
          font-weight: bold; 
          cursor: pointer; 
          margin-left: 5px; 
        }

        /* Kop Surat Sama persis dengan Surat Permohonan */
        .kop-container { 
          border-bottom: 3px double #000; 
          padding-bottom: 5px; 
          margin-bottom: 8px; 
        }

        .header-table { 
          width: 100%; 
          border-collapse: collapse; 
        }

        .header-table td { 
          vertical-align: middle; 
        }

        .text-header { 
          text-align: center; 
        }

        .text-header h3 { 
          margin: 0; 
          font-size: 11pt; 
          font-weight: bold; 
          text-transform: uppercase; 
        }

        .text-header h2 { 
          margin: 0; 
          font-size: 13pt; 
          font-weight: bold; 
          text-transform: uppercase; 
        }

        .text-header p { 
          margin: 2px 0 0 0; 
          font-size: 8pt; 
        }

        .tujuan-area { 
          margin-left: 52%; 
          margin-top: 10px; 
          font-size: 11pt; 
          line-height: 1.35; 
        }

        @media print {
          body { 
            background-color: #fff; 
            padding: 0; 
          }
          .amplop-container { 
            border: none; 
            box-shadow: none; 
            margin: 0; 
            padding: 8mm 12mm; 
            width: 220mm; 
            height: 110mm; 
          }
          .no-print { 
            display: none !important; 
          }
        }
      </style>
      <script>
        function downloadWord() {
          var header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' "+
            "xmlns:w='urn:schemas-microsoft-com:office:word' "+
            "xmlns='http://www.w3.org/TR/REC-html40'>"+
            "<head><meta charset='utf-8'><title>Amplop Surat</title></head><body>";
          var footer = "</body></html>";
          var sourceHTML = header + document.getElementById("amplop-content").innerHTML + footer;
          
          var source = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(sourceHTML);
          var fileDownload = document.createElement("a");
          document.body.appendChild(fileDownload);
          fileDownload.href = source;
          fileDownload.download = 'Amplop_Surat_${namaInstansi.replace(/[^a-zA-Z0-9]/g, '_')}.doc';
          fileDownload.click();
          document.body.removeChild(fileDownload);
        }
      </script>
    </head>
    <body>
      <div class="no-print">
        <button class="btn-word" onclick="downloadWord()">📄 Unduh Word (.doc)</button>
        <button class="btn-print" onclick="window.print()">🖨️ Cetak Amplop</button>
      </div>

      <div class="amplop-container" id="amplop-content">
        <div class="kop-container">
          <table class="header-table">
            <tr>
              <td style="width: 12%; text-align: left;">
                <img src="${urlLogoProv}" style="width: 55px; height: auto;" alt="Logo Kalbar">
              </td>
              <td style="width: 76%;" class="text-header">
                <h3>PEMERINTAH PROVINSI KALIMANTAN BARAT</h3>
                <h2>SMK NEGERI 7 PONTIANAK</h2>
                <p>Jalan Tanjung Raya II Kec. Pontianak Timur, Prov. Kalimantan Barat 78232<br>
                Website: smkn7ptk.sch.id | Email: surat@smkn7ptk.sch.id | CP: 08115733755 (WA Only)</p>
              </td>
              <td style="width: 12%; text-align: right;">
                <img src="${urlLogoSMK}" style="width: 60px; height: auto;" alt="Logo SMKN 7">
              </td>
            </tr>
          </table>
        </div>

        <table style="width: 100%; font-size: 9.5pt; margin-top: 4px;">
          <tr><td style="width: 12%;">Nomor</td><td style="width: 2%;">:</td><td>400.3.8 / . . . . . / SMKN7PTK</td></tr>
          <tr><td>Hal</td><td>:</td><td><strong>Permohonan Praktek Kerja Lapangan (PKL)</strong></td></tr>
          <tr><td>Lampiran</td><td>:</td><td>-</td></tr>
        </table>

        <div class="tujuan-area">
          Kepada<br>
          Yth. <strong>${namaInstansi}</strong><br>
          ${alamatInstansi}<br>
          di<br>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>Pontianak</strong>
        </div>
      </div>
    </body>
    </html>
  `;
}

// === FITUR PANITIA: CETAK SURAT BALASAN DUDI ===
function generateSuratBalasanHTML(tempatId, jurusanSiswa) {
  var detail = getDetailTempatSiswa(tempatId, jurusanSiswa);
  var gelombangList = getSheetData('Gelombang');
  var gel = (gelombangList.length > 0) ? gelombangList[0] : {};

  var tglMulai = formatTanggalIndo(gel.tgl_mulai_pkl || "2027-01-11");
  var tglSelesai = formatTanggalIndo(gel.tgl_selesai_pkl || "2027-03-31");

  var rowsHtml = "";
  if (detail && detail.pendaftar) {
    for (var i = 0; i < detail.pendaftar.length; i++) {
      var p = detail.pendaftar[i];
      rowsHtml += `
        <tr>
          <td style="text-align:center; padding: 5px; border: 1px solid #000;">${i + 1}</td>
          <td style="padding: 5px; border: 1px solid #000;">${p.nama}</td>
          <td style="padding: 5px; border: 1px solid #000; text-align:center;">${p.jenisKelamin || '-'}</td>
          <td style="padding: 5px; border: 1px solid #000; text-align:center;">${p.nisn || '-'}</td>
          <td style="padding: 5px; border: 1px solid #000;">${p.alamatRumah || '-'}</td>
          <td style="padding: 5px; border: 1px solid #000; text-align:center; white-space:nowrap;">${formatNoHp(p.noHp)}</td>
          <td style="padding: 5px; border: 1px solid #000; text-align:center; white-space:nowrap;">${p.kelas || '-'}</td>
        </tr>
      `;
    }
  }

  var namaInstansi = (detail && detail.tempat) ? detail.tempat.nama_instansi : tempatId;
  var alamatInstansi = (detail && detail.tempat) ? detail.tempat.alamat : 'Pontianak';

  return `
    <!DOCTYPE html>
    <html lang="id">
    <head>
      <meta charset="utf-8">
      <title>Surat Balasan PKL - ${namaInstansi}</title>
      <style>
        @page { size: F4; margin: 15mm; }
        body { font-family: 'Times New Roman', Times, serif; font-size: 12pt; line-height: 1.3; margin: 0; padding: 10px; background-color: #f8f9fa; }
        .surat-container { background: #fff; padding: 25px; max-width: 800px; margin: 0 auto; border: 1px solid #ccc; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
        .no-print { margin-bottom: 20px; text-align: right; }
        .btn-print { background: #0d6efd; color: white; border: none; padding: 8px 16px; border-radius: 4px; font-weight: bold; cursor: pointer; margin-left: 5px; }
        .btn-word { background: #198754; color: white; border: none; padding: 8px 16px; border-radius: 4px; font-weight: bold; cursor: pointer; margin-left: 5px; }
        table.data-table { width: 100%; border-collapse: collapse; margin-top: 12px; margin-bottom: 12px; }
        table.data-table th { background-color: #f2f2f2; border: 1px solid #000; padding: 6px; text-align: center; font-size: 10pt; font-weight: bold; }
        table.data-table td { border: 1px solid #000; padding: 5px; font-size: 9.5pt; }

        @media print {
          body { background-color: #fff; padding: 0; }
          .surat-container { border: none; box-shadow: none; padding: 0; max-width: 100%; }
          .no-print { display: none !important; }
        }
      </style>
      <script>
        function downloadWord() {
          var header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' "+
            "xmlns:w='urn:schemas-microsoft-com:office:word' "+
            "xmlns='http://www.w3.org/TR/REC-html40'>"+
            "<head><meta charset='utf-8'><title>Surat Balasan PKL</title></head><body>";
          var footer = "</body></html>";
          var sourceHTML = header + document.getElementById("surat-content").innerHTML + footer;
          
          var source = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(sourceHTML);
          var fileDownload = document.createElement("a");
          document.body.appendChild(fileDownload);
          fileDownload.href = source;
          fileDownload.download = 'Surat_Balasan_PKL_${namaInstansi.replace(/[^a-zA-Z0-9]/g, '_')}.doc';
          fileDownload.click();
          document.body.removeChild(fileDownload);
        }
      </script>
    </head>
    <body>
      <div class="surat-container">
        <div class="no-print">
          <button class="btn-word" onclick="downloadWord()">📄 Unduh Word (.doc)</button>
          <button class="btn-print" onclick="window.print()">🖨️ Cetak Surat Balasan</button>
        </div>

        <div id="surat-content">
          <table style="width: 100%; font-size: 11pt; margin-bottom: 15px;">
            <tr><td style="width: 15%;">Nomor</td><td style="width: 2%;">:</td><td>...................................</td></tr>
            <tr><td>Perihal</td><td>:</td><td><strong>Surat Balasan Praktik Kerja Lapangan (PKL)</strong></td></tr>
          </table>

          <p style="font-size: 11pt; margin-bottom: 15px;">
            Yth. <strong>Kepala SMK N 7 Pontianak</strong><br>
            Di -<br>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Tempat
          </p>

          <p style="text-align: justify; font-size: 11pt; line-height: 1.4;">
            Sehubungan surat permohonan Praktek Kerja Lapangan dengan nomor surat <strong>400.3.8 / …… / SMKN7PTK</strong> dengan ini kami selaku Pimpinan <strong>${namaInstansi}</strong> di Pontianak yang beralamat di: ${alamatInstansi} menyatakan:
          </p>

          <p style="text-align: center; font-size: 13pt; font-weight: bold; margin: 15px 0;">
            MENERIMA / TIDAK MENERIMA
          </p>

          <p style="text-align: justify; font-size: 11pt; line-height: 1.4;">
            Siswa yang bersangkutan Praktek Kerja Lapangan di tempat kami. Adapun nama siswa tersebut:
          </p>

          <table class="data-table">
            <thead>
              <tr>
                <th style="width: 4%;">No</th>
                <th style="width: 25%;">Nama</th>
                <th style="width: 6%;">L/P</th>
                <th style="width: 13%;">NISN</th>
                <th>Alamat Rumah</th>
                <th style="width: 15%;">No HP</th>
                <th>Kelas</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml}
            </tbody>
          </table>

          <p style="text-align: justify; font-size: 11pt; line-height: 1.4;">
            InsyaAllah Praktek kerja lapangan dilaksanakan pada : <strong>${tglMulai} – ${tglSelesai}</strong>
          </p>
          <p style="text-align: justify; font-size: 11pt; line-height: 1.4;">
            Demikian surat balasan ini kami sampaikan, atas perhatiannya kami ucapkan terima kasih.
          </p>

          <br><br>
          <table style="width: 100%;">
            <tr>
              <td style="width: 50%;"></td>
              <td style="text-align: center; font-size: 11pt;">
                Pontianak, .................................... 20....<br>
                Pimpinan Du/Di<br><br><br><br><br>
                --------------------------------------
              </td>
            </tr>
          </table>
        </div>
      </div>
    </body>
    </html>
  `;
}

// === FITUR SISWA: CETAK SURAT PERNYATAAN BERSAMA ===
function generateSuratPernyataanSiswaHTML(userObj) {
  var gelombangList = getSheetData('Gelombang');
  var gel = (gelombangList.length > 0) ? gelombangList[0] : {};

  var tglMulai = formatTanggalIndo(gel.tgl_mulai_pkl || "2027-01-11");
  var tglSelesai = formatTanggalIndo(gel.tgl_selesai_pkl || "2027-03-31");

  var namaSiswa = userObj.nama || ".........................................";
  var nisn = userObj.nisn || ".........................................";
  var kelas = userObj.kelas || ".........................................";
  
  // Ambil lokasi tempat PKL dari userObj atau langsung dari Sheet Pengajuan_PKL jika belum ada di profil
  var tempat_pkl = userObj.tempat_pkl;
  if (!tempat_pkl || tempat_pkl.toString().trim() === "" || tempat_pkl === ".........................................") {
    var pData = getStatusPengajuanSiswa(namaSiswa);
    if (pData) {
      tempat_pkl = pData.tempat_pkl || pData.tempat_id || "";
    }
  }
  if (!tempat_pkl || tempat_pkl.toString().trim() === "") {
    tempat_pkl = ".........................................";
  }

  // Nama Orang Tua selalu diisi titik-titik jika kosong/strip
  var namaOrtu = (userObj.namaOrtu && userObj.namaOrtu !== "-" && userObj.namaOrtu.trim() !== "") 
                 ? userObj.namaOrtu 
                 : ".........................................";

  var alamat = userObj.alamatRumah || ".........................................";
  var noHp = formatNoHp(userObj.noHp);

  return `
    <!DOCTYPE html>
    <html lang="id">
    <head>
      <meta charset="utf-8">
      <title>Surat Pernyataan Bersama PKL - ${namaSiswa}</title>
      <style>
        @page { size: F4; margin: 15mm; }
        body { font-family: 'Times New Roman', Times, serif; font-size: 12pt; line-height: 1.4; margin: 0; padding: 10px; background-color: #f8f9fa; }
        .surat-container { background: #fff; padding: 30px; max-width: 800px; margin: 0 auto; border: 1px solid #ccc; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
        .no-print { margin-bottom: 20px; text-align: right; }
        .btn-print { background: #0d6efd; color: white; border: none; padding: 8px 16px; border-radius: 4px; font-weight: bold; cursor: pointer; margin-left: 5px; }
        .btn-word { background: #198754; color: white; border: none; padding: 8px 16px; border-radius: 4px; font-weight: bold; cursor: pointer; margin-left: 5px; }
        .judul-surat { text-align: center; font-weight: bold; font-size: 14pt; text-decoration: underline; text-transform: uppercase; margin-bottom: 25px; }
        table.identitas { width: 100%; margin-bottom: 15px; font-size: 11.5pt; }
        table.identitas td { padding: 4px 0; vertical-align: top; }
        .materi-box { border: 1px dashed #666; padding: 8px 12px; font-size: 8pt; display: inline-block; color: #555; text-align: center; margin-bottom: 5px; }

        @media print {
          body { background-color: #fff; padding: 0; }
          .surat-container { border: none; box-shadow: none; padding: 0; max-width: 100%; }
          .no-print { display: none !important; }
        }
      </style>
      <script>
        function downloadWord() {
          var header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' "+
            "xmlns:w='urn:schemas-microsoft-com:office:word' "+
            "xmlns='http://www.w3.org/TR/REC-html40'>"+
            "<head><meta charset='utf-8'><title>Surat Pernyataan Bersama</title></head><body>";
          var footer = "</body></html>";
          var sourceHTML = header + document.getElementById("surat-content").innerHTML + footer;
          
          var source = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(sourceHTML);
          var fileDownload = document.createElement("a");
          document.body.appendChild(fileDownload);
          fileDownload.href = source;
          fileDownload.download = 'Surat_Pernyataan_Bersama_${namaSiswa.replace(/[^a-zA-Z0-9]/g, '_')}.doc';
          fileDownload.click();
          document.body.removeChild(fileDownload);
        }
      </script>
    </head>
    <body>
      <div class="surat-container">
        <div class="no-print">
          <button class="btn-word" onclick="downloadWord()">📄 Unduh Word (.doc)</button>
          <button class="btn-print" onclick="window.print()">🖨️ Cetak Surat Pernyataan</button>
        </div>

        <div id="surat-content">
          <div class="judul-surat">SURAT PERNYATAAN BERSAMA</div>

          <p style="font-size: 11.5pt;">Yang bertanda tangan di bawah ini:</p>

          <table class="identitas">
            <tr><td style="width: 28%;">Nama Siswa</td><td style="width: 3%;">:</td><td><strong>${namaSiswa}</strong></td></tr>
            <tr><td>NISN</td><td>:</td><td>${nisn}</td></tr>
            <tr><td>Kelas</td><td>:</td><td>${kelas}</td></tr>
            <tr><td>Tempat PKL</td><td>:</td><td>${tempat_pkl}</td></tr>
            <tr><td>Nama Orang Tua/Wali</td><td>:</td><td>${namaOrtu}</td></tr>
            <tr><td>Alamat Lengkap</td><td>:</td><td>${alamat}</td></tr>
            <tr><td>Nomor HP</td><td>:</td><td>${noHp}</td></tr>
          </table>

          <p style="text-align: justify; text-indent: 30px; font-size: 11.5pt;">
            Menyatakan dengan sesungguhnya bahwa saya, selaku siswa peserta PKL, bersedia untuk mengikuti program Praktek Kerja Lapangan (PKL) yang akan dilaksanakan oleh sekolah pada tanggal <strong>${tglMulai}</strong> sampai dengan <strong>${tglSelesai}</strong>.
          </p>

          <p style="text-align: justify; font-size: 11.5pt;">
            Saya memahami dan bersedia mematuhi semua peraturan yang telah ditetapkan oleh sekolah dan instansi tempat PKL, serta bertanggung jawab penuh atas segala risiko yang mungkin terjadi selama pelaksanaan PKL. Selama proses PKL berlangsung saya akan selalu hadir tepat waktu sesuai jam operasional tempat PKL, dan jika tidak dapat hadir karena sakit, saya akan melaporkan selambat-lambatnya 1x24 jam kepada Guru Pembimbing Sekolah dan Industri. Jika saya tidak hadir selama 3 hari berturut-turut tanpa keterangan, saya bersedia di keluarkan / ditarik dari tempat PKL dan bersedia dinyatakan tidak lulus PKL tahap ini.
          </p>

          <p style="text-align: justify; font-size: 11.5pt;">
            Saya tidak akan mengajukan pindah tempat PKL terkecuali mengalami sakit parah (masuk rumah sakit), depresi (gangguan kesadaran), di siksa sampai mengalami luka. Jika saya dan orang tua memaksa pindah tempat PKL maka saya bersedia dinyatakan tidak lulus PKL tahap ini.
          </p>

          <p style="text-align: justify; font-size: 11.5pt;">
            Demikian surat pernyataan ini dibuat dengan sadar dan tanpa paksaan dari pihak mana pun.
          </p>

          <br>
          <table style="width: 100%; font-size: 11.5pt;">
            <tr>
              <td style="width: 50%;"></td>
              <td style="text-align: center;">
                Pontianak, ........................ 20....                         
              </td>
            </tr>
            <tr style="text-align: center; vertical-align: top;">
              <td style="padding-top: 10px;">
                Orang Tua / Wali<br><br><br><br><br>
                (________________________________________)
              </td>
              <td style="padding-top: 10px;">
                Siswa Peserta PKL<br><br>
                <div class="materi-box">Materai<br>Rp 10.000</div><br><br>
                ( ${namaSiswa} )
              </td>
            </tr>
          </table>
        </div>
      </div>
    </body>
    </html>
  `;
}

// === PLOTTING GURU PEMBIMBING & JURNAL ===
function getGuruList() {
  var allUsers = getSheetData('Users');
  return allUsers.filter(function(u) { return u.role === 'Guru'; });
}

function simpanPenugasanGuru(formData) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Penugasan_Pembimbing');
    if (!sheet) return { status: "Error", message: "Sheet Penugasan_Pembimbing tidak ditemukan." };

    var userList = getSheetData('Users');
    var tempatList = getSheetData('Tempat_PKL');

    var guru = userList.find(function(u) {
      return (u.user_id && u.user_id.toString().trim() === formData.guruId.toString().trim()) ||
             (u.email && u.email.toString().trim().toLowerCase() === formData.guruId.toString().trim().toLowerCase()) ||
             (u.nama && u.nama.toString().trim().toLowerCase() === formData.guruId.toString().trim().toLowerCase());
    });
    var namaGuruFix = guru ? guru.nama : formData.guruId;

    var tempat = tempatList.find(function(t, idx) {
      var tIdStr = (t.tempat_id !== undefined && t.tempat_id !== "") ? t.tempat_id.toString().trim() : (idx + 1).toString();
      return tIdStr === formData.tempatId.toString().trim() ||
             (t.nama_instansi && t.nama_instansi.toString().trim().toLowerCase() === formData.tempatId.toString().trim().toLowerCase());
    });
    var namaTempatFix = tempat ? tempat.nama_instansi : formData.tempatId;

    var id = "TGS-" + new Date().getTime();
    sheet.appendRow([id, namaGuruFix, namaTempatFix, ""]);
    
    return { status: "Success", message: "Guru pembimbing berhasil ditugaskan." };
  } catch (error) {
    return { status: "Error", message: error.message };
  }
}

function getPenugasanGroupedList(jurusanKaprog) {
  try {
    var penugasanList = getSheetData('Penugasan_Pembimbing');
    var userList = getSheetData('Users');
    var tempatList = getSheetData('Tempat_PKL');
    var pengajuanList = getSheetData('Pengajuan_PKL');

    if (!penugasanList || penugasanList.length === 0) return [];

    var result = [];

    for (var i = 0; i < penugasanList.length; i++) {
      var pg = penugasanList[i];
      var pGuru = pg.nama_guru ? pg.nama_guru.toString().trim() : (pg.guru_id ? pg.guru_id.toString().trim() : "");
      var pTempat = pg.tempat_pkl ? pg.tempat_pkl.toString().trim() : (pg.tempat_id ? pg.tempat_id.toString().trim() : "");

      var guru = userList.find(function(u) {
        return (u.nama && u.nama.toString().trim().toLowerCase() === pGuru.toLowerCase()) ||
               (u.user_id && u.user_id.toString().trim() === pGuru) ||
               (u.email && u.email.toString().trim().toLowerCase() === pGuru.toLowerCase());
      });

      var tempat = tempatList.find(function(t) {
        return (t.nama_instansi && t.nama_instansi.toString().trim().toLowerCase() === pTempat.toLowerCase()) ||
               (t.tempat_id && t.tempat_id.toString().trim() === pTempat);
      });

      var namaInstansiFix = tempat ? tempat.nama_instansi : pTempat;

      var siswaListInTempat = [];
      for (var j = 0; j < pengajuanList.length; j++) {
        var p = pengajuanList[j];
        var pSiswaTempat = p.tempat_pkl ? p.tempat_pkl.toString().trim() : (p.tempat_id ? p.tempat_id.toString().trim() : "");

        if (p.status_verifikasi === 'Disetujui' && pSiswaTempat.toLowerCase() === namaInstansiFix.toLowerCase()) {
          var namaSiswa = p.nama_siswa || p.siswa_id;
          var siswa = userList.find(function(u) {
            return (u.nama && u.nama.toString().trim().toLowerCase() === namaSiswa.toString().trim().toLowerCase()) ||
                   (u.user_id && u.user_id.toString().trim() === namaSiswa.toString().trim());
          });

          if (siswa) {
            if (!jurusanKaprog || (siswa.jurusan && siswa.jurusan.toString().toLowerCase().trim() === jurusanKaprog.toString().toLowerCase().trim())) {
              siswaListInTempat.push({
                nama: siswa.nama,
                kelas: siswa.kelas || '-',
                jurusan: siswa.jurusan || '-',
                noHp: formatNoHp(siswa.no_hp || siswa.noHp)
              });
            }
          }
        }
      }

      if (jurusanKaprog && tempat) {
        var kKey = "kuota_" + jurusanKaprog.toString().toLowerCase().trim();
        var kuotaJurus = parseInt(tempat[kKey] || tempat[kKey.toUpperCase()] || 0);
        if (kuotaJurus === 0 && siswaListInTempat.length === 0) continue;
      }

      result.push({
        penugasanId: pg.penugasan_id,
        namaGuru: guru ? guru.nama : pGuru,
        noHpGuru: guru ? formatNoHp(guru.no_hp || guru.noHp) : '-',
        namaTempat: namaInstansiFix,
        siswaList: siswaListInTempat
      });
    }

    return result;
  } catch (err) {
    Logger.log("Error getPenugasanGroupedList: " + err.message);
    return [];
  }
}

// KHUSUS GURU PEMBIMBING: MENGAMBIL DAFTAR SISWA BIMBINGAN DALAM BENTUK CARD
function getSiswaBimbinganGuru(guruNama) {
  try {
    var penugasanList = getSheetData('Penugasan_Pembimbing');
    var userList = getSheetData('Users');
    var tempatList = getSheetData('Tempat_PKL');
    var pengajuanList = getSheetData('Pengajuan_PKL');

    if (!guruNama) return [];
    var gNamaStr = guruNama.toString().trim().toLowerCase();

    var penugasanGuru = penugasanList.filter(function(pg) {
      var nGuru = pg.nama_guru ? pg.nama_guru.toString().trim().toLowerCase() : (pg.guru_id ? pg.guru_id.toString().trim().toLowerCase() : "");
      return nGuru === gNamaStr;
    });

    var result = [];

    for (var i = 0; i < penugasanGuru.length; i++) {
      var pg = penugasanGuru[i];
      var pTempat = pg.tempat_pkl ? pg.tempat_pkl.toString().trim() : (pg.tempat_id ? pg.tempat_id.toString().trim() : "");

      var tempat = tempatList.find(function(t) {
        return (t.nama_instansi && t.nama_instansi.toString().trim().toLowerCase() === pTempat.toLowerCase()) ||
               (t.tempat_id && t.tempat_id.toString().trim() === pTempat);
      });

      var namaInstansiFix = tempat ? tempat.nama_instansi : pTempat;

      var siswaList = [];
      for (var j = 0; j < pengajuanList.length; j++) {
        var p = pengajuanList[j];
        var pSiswaTempat = p.tempat_pkl ? p.tempat_pkl.toString().trim() : (p.tempat_id ? p.tempat_id.toString().trim() : "");

        if (p.status_verifikasi === 'Disetujui' && pSiswaTempat.toLowerCase() === namaInstansiFix.toLowerCase()) {
          var namaSiswa = p.nama_siswa || p.siswa_id;
          var siswa = userList.find(function(u) {
            return (u.nama && u.nama.toString().trim().toLowerCase() === namaSiswa.toString().trim().toLowerCase()) ||
                   (u.user_id && u.user_id.toString().trim() === namaSiswa.toString().trim());
          });

          if (siswa) {
            siswaList.push({
              nama: siswa.nama,
              kelas: siswa.kelas || '-',
              jurusan: siswa.jurusan || '-',
              noHp: formatNoHp(siswa.no_hp || siswa.noHp),
              alamat: siswa.alamat_rumah || siswa.alamatRumah || '-'
            });
          }
        }
      }

      result.push({
        namaTempat: namaInstansiFix,
        alamatTempat: tempat ? tempat.alamat : '-',
        kontakTempat: tempat ? formatNoHp(tempat.kontak_person || tempat.kontak) : '-',
        siswaList: siswaList
      });
    }

    return result;
  } catch (err) {
    Logger.log("Error getSiswaBimbinganGuru: " + err.message);
    return [];
  }
}

function hapusPenugasanGuru(penugasanId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Penugasan_Pembimbing');
    if (!sheet) return { status: "Error", message: "Sheet Penugasan_Pembimbing tidak ditemukan." };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === penugasanId.toString().trim()) {
        sheet.deleteRow(i + 1);
        return { status: "Success", message: "Penugasan pembimbing berhasil dihapus." };
      }
    }
    return { status: "Error", message: "Data penugasan tidak ditemukan." };
  } catch (error) {
    return { status: "Error", message: error.message };
  }
}

function getPenugasanPembimbingList(jurusanKaprog) {
  var penugasan = getSheetData('Penugasan_Pembimbing');
  var userList = getSheetData('Users');
  var tempatList = getSheetData('Tempat_PKL');
  var pengajuanList = getSheetData('Pengajuan_PKL');

  return penugasan.map(function(pg) {
    var pGuru = pg.nama_guru ? pg.nama_guru.toString().trim() : (pg.guru_id ? pg.guru_id.toString().trim() : "");
    var guru = userList.find(function(u) { 
      return (u.user_id && u.user_id.toString().trim() === pGuru) || 
             (u.email && u.email.toString().trim().toLowerCase() === pGuru.toLowerCase()) ||
             (u.nama && u.nama.toString().trim().toLowerCase() === pGuru.toLowerCase()); 
    });

    var pTempat = pg.tempat_pkl ? pg.tempat_pkl.toString().trim() : (pg.tempat_id ? pg.tempat_id.toString().trim() : "");
    var tempat = tempatList.find(function(t, idx) { 
      var tIdStr = (t.tempat_id !== undefined && t.tempat_id !== "") ? t.tempat_id.toString().trim() : (idx + 1).toString();
      return tIdStr === pTempat || (t.nama_instansi && t.nama_instansi.toString().trim().toLowerCase() === pTempat.toLowerCase()); 
    });

    var siswaListInTempat = [];
    for (var i = 0; i < pengajuanList.length; i++) {
      var p = pengajuanList[i];
      var pSiswaTempat = p.tempat_pkl ? p.tempat_pkl.toString().trim() : (p.tempat_id ? p.tempat_id.toString().trim() : "");
      var tNama = tempat ? tempat.nama_instansi.toString().trim() : "";

      if (p.status_verifikasi === 'Disetujui' && (pSiswaTempat === pTempat || pSiswaTempat.toLowerCase() === tNama.toLowerCase())) {
        var namaSiswa = p.nama_siswa || p.siswa_id;
        var siswa = userList.find(function(u) {
          return (u.nama && u.nama.toString().trim().toLowerCase() === namaSiswa.toString().trim().toLowerCase()) ||
                 (u.user_id && u.user_id.toString().trim() === namaSiswa.toString().trim());
        });
        if (siswa) {
          if (!jurusanKaprog || (siswa.jurusan && siswa.jurusan.toString().toLowerCase().trim() === jurusanKaprog.toString().toLowerCase().trim())) {
            siswaListInTempat.push(siswa.nama + " (" + (siswa.kelas || '-') + ")");
          }
        }
      }
    }

    return {
      penugasanId: pg.penugasan_id,
      namaGuru: guru ? guru.nama : pGuru,
      namaTempat: tempat ? tempat.nama_instansi : pTempat,
      daftarSiswa: siswaListInTempat.length > 0 ? siswaListInTempat.join(', ') : 'Belum ada siswa disetujui',
      jurusanTempat: tempat ? tempat : {}
    };
  }).filter(function(item) {
    if (!jurusanKaprog) return true;
    var key = "kuota_" + jurusanKaprog.toString().toLowerCase().trim();
    return (item.jurusanTempat[key] && parseInt(item.jurusanTempat[key]) > 0) || item.daftarSiswa !== 'Belum ada siswa disetujui';
  });
}

function getSheetNameForStudent(siswaId) {
  var userList = getSheetData('Users');
  var user = userList.find(function(u) {
    return (u.nama && u.nama.toString().trim().toLowerCase() === siswaId.toString().trim().toLowerCase()) ||
           (u.user_id && u.user_id.toString().trim() === siswaId.toString().trim());
  });
  var kelas = user && user.kelas ? user.kelas.toString().trim() : "Umum";
  var sheetName = "Jurnal_" + kelas.replace(/[^a-zA-Z0-9]/g, "_");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["timestamp", "jurnal_id", "nama_siswa", "kegiatan", "kehadiran", "catatan_guru", "status_validasi"]);
  }
  return sheetName;
}

function simpanJurnal(dataJurnal) {
  try {
    var sheetName = getSheetNameForStudent(dataJurnal.siswaId);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var id = "JRN-" + new Date().getTime();
    var timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    
    sheet.appendRow([
      timestamp,
      id, 
      dataJurnal.siswaId, 
      dataJurnal.kegiatan, 
      dataJurnal.kehadiran, 
      "", 
      "Pending"
    ]);
    return { status: "Success", message: "Jurnal harian berhasil disimpan." };
  } catch (error) {
    return { status: "Error", message: error.message };
  }
}

function getJurnalSiswa(siswaId) {
  try {
    var sheetName = getSheetNameForStudent(siswaId);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        obj[headers[j]] = row[j];
      }
      var sId = obj.nama_siswa || obj.siswa_id || "";
      if (sId.toString().trim().toLowerCase() === siswaId.toString().trim().toLowerCase()) {
        result.push(obj);
      }
    }
    return result;
  } catch (error) {
    return [];
  }
}

function getJurnalBimbinganGuru(guruId) {
  var penugasanList = getSheetData('Penugasan_Pembimbing');
  var pengajuanList = getSheetData('Pengajuan_PKL');
  var userList = getSheetData('Users');
  var tempatList = getSheetData('Tempat_PKL');

  if (!guruId) return [];
  var guruIdStr = guruId.toString().trim().toLowerCase();

  var tempatDiampu = penugasanList.filter(function(pg) {
    var gId = pg.nama_guru ? pg.nama_guru.toString().trim().toLowerCase() : (pg.guru_id ? pg.guru_id.toString().trim().toLowerCase() : "");
    return gId === guruIdStr;
  }).map(function(pg) {
    return pg.tempat_pkl ? pg.tempat_pkl.toString().trim() : (pg.tempat_id ? pg.tempat_id.toString().trim() : "");
  });

  if (tempatDiampu.length === 0) return [];

  var siswaBimbinganMap = {};
  for (var i = 0; i < pengajuanList.length; i++) {
    var p = pengajuanList[i];
    var pTempat = p.tempat_pkl ? p.tempat_pkl.toString().trim() : (p.tempat_id ? p.tempat_id.toString().trim() : "");

    var cocokTempat = tempatDiampu.some(function(tId) {
      var tmpt = tempatList.find(function(t) { return t.tempat_id.toString().trim() === tId; });
      var namaInstansi = tmpt ? tmpt.nama_instansi.toString().trim().toLowerCase() : "";
      return pTempat === tId || pTempat.toLowerCase() === namaInstansi;
    });

    if (p.status_verifikasi === 'Disetujui' && cocokTempat) {
      var namaSiswa = p.nama_siswa || p.siswa_id;
      var user = userList.find(function(u) {
        return (u.nama && u.nama.toString().trim().toLowerCase() === namaSiswa.toString().trim().toLowerCase()) ||
               (u.user_id && u.user_id.toString().trim() === namaSiswa.toString().trim());
      });
      
      var tObj = tempatList.find(function(t) {
        return t.tempat_id.toString().trim() === pTempat || t.nama_instansi.toString().trim().toLowerCase() === pTempat.toLowerCase();
      });

      var sKey = namaSiswa.toString().trim().toLowerCase();
      siswaBimbinganMap[sKey] = {
        nama: user ? user.nama : namaSiswa,
        kelas: user ? (user.kelas || '-') : '-',
        tempat: tObj ? tObj.nama_instansi : pTempat
      };
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var result = [];

  for (var s = 0; s < allSheets.length; s++) {
    var sheet = allSheets[s];
    if (sheet.getName().indexOf("Jurnal_") === 0) {
      var data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        var headers = data[0];
        for (var j = 1; j < data.length; j++) {
          var row = data[j];
          var jrn = {};
          for (var k = 0; k < headers.length; k++) {
            jrn[headers[k]] = row[k];
          }
          var jSiswa = jrn.nama_siswa || jrn.siswa_id || "";
          var sKey = jSiswa.toString().trim().toLowerCase();
          if (siswaBimbinganMap[sKey]) {
            var infoSiswa = siswaBimbinganMap[sKey];
            result.push({
              jurnal_id: jrn.jurnal_id,
              namaSiswa: infoSiswa.nama,
              kelasSiswa: infoSiswa.kelas,
              namaTempat: infoSiswa.tempat,
              tanggal: jrn.timestamp ? jrn.timestamp.toString().split(' ')[0] : (jrn.tanggal || '-'),
              kegiatan: jrn.kegiatan,
              kehadiran: jrn.kehadiran,
              status_validasi: jrn.status_validasi,
              catatan_guru: jrn.catatan_guru
            });
          }
        }
      }
    }
  }
  return result;
}

function validasiJurnal(jurnalId, catatan, status) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var allSheets = ss.getSheets();
    for (var s = 0; s < allSheets.length; s++) {
      var sheet = allSheets[s];
      if (sheet.getName().indexOf("Jurnal_") === 0) {
        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          if (data[i][1] === jurnalId) {
            sheet.getRange(i + 1, 6).setValue(catatan);
            sheet.getRange(i + 1, 7).setValue(status);
            return { status: "Success", message: "Jurnal berhasil divalidasi." };
          }
        }
      }
    }
    return { status: "Error", message: "Data jurnal tidak ditemukan." };
  } catch (error) {
    return { status: "Error", message: error.message };
  }
}

// === LAPORAN MASALAH ===
function simpanLaporanMasalah(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Laporan_Masalah') || ss.getSheetByName('LaporanMasalah');
    if (!sheet) {
      sheet = ss.insertSheet('Laporan_Masalah');
      sheet.appendRow(["timestamp", "laporan_id", "pelapor_nama", "pelapor_role", "nama_siswa", "tempat_pkl", "deskripsi_masalah"]);
    }
    var id = "LPR-" + new Date().getTime();
    var timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    
    sheet.appendRow([
      timestamp,
      id,
      data.pelaporNama,
      data.pelaporRole,
      data.namaSiswa,
      data.tempatPkl,
      data.deskripsiMasalah
    ]);
    return { status: "Success", message: "Laporan masalah siswa berhasil dikirim." };
  } catch (error) {
    return { status: "Error", message: error.message };
  }
}

function getLaporanMasalahList(user) {
  try {
    var list = getSheetData('Laporan_Masalah');
    if (!list || list.length === 0) {
      list = getSheetData('LaporanMasalah');
    }
    if (!list || list.length === 0) return [];

    var userRole = user ? user.role : '';
    var userNamaStr = user && user.nama ? user.nama.toString().trim().toLowerCase() : '';

    var result = [];

    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      var nSiswa = row.nama_siswa || row.namaSiswa || "";
      if (!nSiswa) continue;

      var pelaporNama = row.pelapor_nama || row.pelaporNama || "";
      var tmptPkl = row.tempat_pkl || row.tempatPkl || "";

      var laporanObj = {
        laporan_id: row.laporan_id || row.laporanId || "",
        timestamp: row.timestamp ? formatTanggalIndo(row.timestamp) : "",
        pelapor_nama: pelaporNama,
        pelapor_role: row.pelapor_role || row.pelaporRole || "",
        nama_siswa: nSiswa,
        tempat_pkl: tmptPkl,
        deskripsi_masalah: row.deskripsi_masalah || row.deskripsiMasalah || ""
      };

      if (userRole === 'Admin' || userRole === 'Super Admin' || userRole === 'Waka' || userRole === 'Panitia PKL' || userRole === 'Kaprog' || userRole === 'Guru') {
        result.push(laporanObj);
      } else if (userRole === 'DUDI') {
        var pNamaLower = pelaporNama.toString().trim().toLowerCase();
        var tPklLower = tmptPkl.toString().trim().toLowerCase();

        if (pNamaLower === userNamaStr || tPklLower === userNamaStr || pNamaLower.indexOf(userNamaStr) !== -1 || userNamaStr.indexOf(tPklLower) !== -1) {
          result.push(laporanObj);
        }
      } else {
        result.push(laporanObj);
      }
    }

    return result;
  } catch (e) {
    Logger.log("Error getLaporanMasalahList: " + e.message);
    return [];
  }
}

function getSiswaDudiList(userRole, userNama) {
  var pengajuanList = getSheetData('Pengajuan_PKL');
  var userList = getSheetData('Users');
  var tempatList = getSheetData('Tempat_PKL');
  var penugasanList = getSheetData('Penugasan_Pembimbing');

  var result = [];
  var addedMap = {};

  if (userRole === 'DUDI') {
    var dudiStr = userNama ? userNama.toString().trim().toLowerCase() : "";
    for (var i = 0; i < pengajuanList.length; i++) {
      var p = pengajuanList[i];
      if (p.status_verifikasi !== 'Disetujui') continue;

      var pTempat = p.tempat_pkl ? p.tempat_pkl.toString().trim() : (p.tempat_id ? p.tempat_id.toString().trim() : "");
      var tObj = tempatList.find(function(t) {
        return t.tempat_id.toString().trim() === pTempat || (t.nama_instansi && t.nama_instansi.toString().trim().toLowerCase() === pTempat.toLowerCase());
      });
      var namaInstansi = tObj ? tObj.nama_instansi : pTempat;

      if (namaInstansi.toString().trim().toLowerCase().indexOf(dudiStr) !== -1 || dudiStr.indexOf(namaInstansi.toString().trim().toLowerCase()) !== -1) {
        var namaSiswa = p.nama_siswa || p.siswa_id;
        if (!namaSiswa) continue;
        var sKey = namaSiswa.toString().trim().toLowerCase();
        if (addedMap[sKey]) continue;
        addedMap[sKey] = true;

        var user = userList.find(function(u) {
          return (u.nama && u.nama.toString().trim().toLowerCase() === sKey) ||
                 (u.user_id && u.user_id.toString().trim().toLowerCase() === sKey);
        });

        result.push({
          namaSiswa: user ? user.nama : namaSiswa,
          tempatPkl: namaInstansi
        });
      }
    }
  } else if (userRole === 'Guru') {
    var gNamaStr = userNama ? userNama.toString().trim().toLowerCase() : "";
    var penugasanGuru = penugasanList.filter(function(pg) {
      var nGuru = pg.nama_guru ? pg.nama_guru.toString().trim().toLowerCase() : (pg.guru_id ? pg.guru_id.toString().trim().toLowerCase() : "");
      return nGuru === gNamaStr;
    });

    var tempatListDiampu = penugasanGuru.map(function(pg) {
      var pT = pg.tempat_pkl ? pg.tempat_pkl.toString().trim() : (pg.tempat_id ? pg.tempat_id.toString().trim() : "");
      var t = tempatList.find(function(tmp) { return tmp.tempat_id.toString().trim() === pT || (tmp.nama_instansi && tmp.nama_instansi.toString().trim().toLowerCase() === pT.toLowerCase()); });
      return t ? t.nama_instansi.toString().trim().toLowerCase() : pT.toLowerCase();
    });

    for (var j = 0; j < pengajuanList.length; j++) {
      var p2 = pengajuanList[j];
      if (p2.status_verifikasi !== 'Disetujui') continue;

      var pTempat2 = p2.tempat_pkl ? p2.tempat_pkl.toString().trim() : (p2.tempat_id ? p2.tempat_id.toString().trim() : "");
      var tObj2 = tempatList.find(function(t) {
        return t.tempat_id.toString().trim() === pTempat2 || (t.nama_instansi && t.nama_instansi.toString().trim().toLowerCase() === pTempat2.toLowerCase());
      });
      var namaInstansi2 = tObj2 ? tObj2.nama_instansi : pTempat2;

      if (tempatListDiampu.indexOf(namaInstansi2.toString().trim().toLowerCase()) !== -1) {
        var namaSiswa2 = p2.nama_siswa || p2.siswa_id;
        if (!namaSiswa2) continue;
        var sKey2 = namaSiswa2.toString().trim().toLowerCase();
        if (addedMap[sKey2]) continue;
        addedMap[sKey2] = true;

        var user2 = userList.find(function(u) {
          return (u.nama && u.nama.toString().trim().toLowerCase() === sKey2) ||
                 (u.user_id && u.user_id.toString().trim().toLowerCase() === sKey2);
        });

        result.push({
          namaSiswa: user2 ? user2.nama : namaSiswa2,
          tempatPkl: namaInstansi2
        });
      }
    }
  }

  return result;
}

function tarikSiswaPKL(laporanId, statusPenarikan, keterangan) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetLaporan = ss.getSheetByName('Laporan_Masalah') || ss.getSheetByName('LaporanMasalah');
    if (!sheetLaporan) return { status: "Error", message: "Sheet Laporan_Masalah tidak ditemukan." };

    var dataLpr = sheetLaporan.getDataRange().getValues();
    var targetSiswa = "";
    var targetTempat = "";

    for (var i = 1; i < dataLpr.length; i++) {
      if (dataLpr[i][1].toString().trim() === laporanId.toString().trim()) {
        targetSiswa = dataLpr[i][4];
        targetTempat = dataLpr[i][5];
        sheetLaporan.getRange(i + 1, 8).setValue(statusPenarikan + " (" + keterangan + ")");
        break;
      }
    }

    if (!targetSiswa) return { status: "Error", message: "Laporan tidak ditemukan." };

    var sheetPengajuan = ss.getSheetByName('Pengajuan_PKL');
    if (sheetPengajuan) {
      var dataPgj = sheetPengajuan.getDataRange().getValues();
      for (var j = 1; j < dataPgj.length; j++) {
        var sNama = dataPgj[j][1] ? dataPgj[j][1].toString().trim().toLowerCase() : "";
        if (sNama === targetSiswa.toString().trim().toLowerCase()) {
          if (statusPenarikan === 'Pindah Tempat') {
            sheetPengajuan.deleteRow(j + 1);
          } else if (statusPenarikan === 'Tidak LULUS') {
            sheetPengajuan.getRange(j + 1, 5).setValue("Dibatalkan / Tidak LULUS");
          }
          break;
        }
      }
    }

    return { status: "Success", message: "Tindakan penarikan siswa (" + statusPenarikan + ") berhasil diproses." };
  } catch (error) {
    return { status: "Error", message: error.message };
  }
}