function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Pembuat Tagihan LHT - SBB')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function cleanText(val) {
  if (val === undefined || val === null) return '';
  return String(val).trim();
}

function sortServerItemsByCustomerAZ(itemsArray) {
  if (!Array.isArray(itemsArray)) return [];
  return itemsArray.slice().sort(function(a, b) {
    if (a.isRowEmpty && !b.isRowEmpty) return 1;
    if (!a.isRowEmpty && b.isRowEmpty) return -1;
    if (a.isRowEmpty && b.isRowEmpty) return 0;

    var nameA = cleanText(a.namaPelanggan).toLowerCase();
    var nameB = cleanText(b.namaPelanggan).toLowerCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;

    var fakA = cleanText(a.noFaktur).toLowerCase();
    var fakB = cleanText(b.noFaktur).toLowerCase();
    return fakA.localeCompare(fakB);
  });
}

function getConfData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var confSheet = ss.getSheetByName('conf');
    if (!confSheet) return { success: false, message: 'Sheet conf tidak ditemukan!' };

    var values = confSheet.getDataRange().getValues();
    if (values.length <= 1) return { success: false, message: 'Data konfigurasi pada sheet conf masih kosong!' };

    var admins = [], penagihs = [], depos = [], perusahaans = [], divisis = [];
    var tautanLHT = '';
    var lhtSheet = '';

    var headers = values[0];
    var colTautanIdx = 9;
    var colSheetIdx = 10;

    for (var h = 0; h < headers.length; h++) {
      var headerText = cleanText(headers[h]).toLowerCase();
      if (headerText === 'tautan lht') {
        colTautanIdx = h;
      }
      if (headerText === 'lht sheet') {
        colSheetIdx = h;
      }
    }

    for (var i = 1; i < values.length; i++) {
      if (values[i][0]) admins.push(cleanText(values[i][0]));
      if (values[i][1]) penagihs.push(cleanText(values[i][1]));
      if (values[i][2]) depos.push(cleanText(values[i][2]));
      if (values[i][3]) perusahaans.push(cleanText(values[i][3]));
      if (values[i][4]) divisis.push(cleanText(values[i][4]));
      
      if (!tautanLHT && values[i][colTautanIdx]) {
        tautanLHT = cleanText(values[i][colTautanIdx]);
      }

      if (!lhtSheet && values[i][colSheetIdx]) {
        lhtSheet = cleanText(values[i][colSheetIdx]);
      }
    }

    function unique(arr) { return Array.from(new Set(arr)).filter(Boolean); }

    var smartFillRules = [];
    if (typeof parseSmartFillFromConfValues === 'function') {
      smartFillRules = parseSmartFillFromConfValues(values);
    }

    return {
      success: true,
      admins: unique(admins),
      penagihs: unique(penagihs),
      depos: unique(depos),
      perusahaans: unique(perusahaans),
      divisis: unique(divisis),
      tautanLHT: tautanLHT,
      lhtSheet: lhtSheet,
      smartFillRules: smartFillRules
    };
  } catch (err) {
    return { success: false, message: 'Gagal memuat konfigurasi: ' + err.message };
  }
}

function getPiutangData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Piutang');
    if (!sheet) return { success: false, message: 'Sheet Piutang tidak ditemukan!' };

    var values = sheet.getDataRange().getValues();
    if (values.length <= 1) return { success: true, items: [], waktuPembaruan: '-' };

    var waktuPembaruan = '-';
    var headers = values[0];
    var colWaktuIdx = -1;
    for (var h = 0; h < headers.length; h++) {
      if (cleanText(headers[h]).toLowerCase() === 'waktu pembaruan') {
        colWaktuIdx = h;
        break;
      }
    }

    if (colWaktuIdx !== -1 && values.length > 1) {
      var dispWaktu = sheet.getRange(2, colWaktuIdx + 1).getDisplayValue();
      waktuPembaruan = dispWaktu ? dispWaktu : '-';
    } else {
      var dispL2 = sheet.getRange("L2").getDisplayValue();
      waktuPembaruan = dispL2 ? dispL2 : '-';
    }

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var items = [];
    var rawRows = values.slice(1);

    rawRows.forEach(function(row) {
      var noPelanggan = cleanText(row[0]);
      var noFaktur = cleanText(row[1]);
      if (!noPelanggan && !noFaktur) return;

      var tglFakturRaw = row[2];
      var tglFakturDate = parseIndoDate(tglFakturRaw);

      var umurHari = 0;
      var formattedTglFaktur = '';
      if (tglFakturDate && !isNaN(tglFakturDate.getTime())) {
        tglFakturDate.setHours(0, 0, 0, 0);
        var diffTime = today.getTime() - tglFakturDate.getTime();
        umurHari = Math.floor(diffTime / (1000 * 60 * 60 * 24));
        
        var dd = ('0' + tglFakturDate.getDate()).slice(-2);
        var mm = ('0' + (tglFakturDate.getMonth() + 1)).slice(-2);
        var yyyy = tglFakturDate.getFullYear();
        formattedTglFaktur = dd + '/' + mm + '/' + yyyy;
      } else {
        formattedTglFaktur = formatDateDMY(tglFakturRaw);
      }

      var nilaiFaktur = Number(row[4]) ? Number(row[4]) : 0;
      var sisaPiutang = Number(row[5]) ? Number(row[5]) : 0;
      var terbayar = nilaiFaktur - sisaPiutang;

      items.push({
        noPelanggan: noPelanggan,
        noFaktur: noFaktur,
        tglFaktur: formattedTglFaktur,
        jatuhTempo: formatDateDMY(row[3]),
        nilaiFaktur: nilaiFaktur,
        sisaPiutang: sisaPiutang,
        terbayar: terbayar > 0 ? terbayar : 0,
        umur: umurHari,
        umurHariNum: umurHari,
        namaPelanggan: cleanText(row[7]),
        namaPenjual: cleanText(row[8]),
        namaKontak: cleanText(row[9])
      });
    });

    var sortedItems = sortServerItemsByCustomerAZ(items);

    return { 
      success: true, 
      items: sortedItems, 
      waktuPembaruan: waktuPembaruan 
    };
  } catch (err) {
    return { success: false, message: 'Gagal memuat data piutang: ' + err.message };
  }
}

function simpanKeSpreadsheetTarget(headerConfig, itemsData) {
  try {
    var confRes = getConfData();
    var targetUrl = confRes.tautanLHT;
    if (!targetUrl) return { success: false, message: 'Tautan LHT pada sheet conf belum diisi!' };

    if (!itemsData || !Array.isArray(itemsData) || itemsData.length === 0) {
      return { success: false, message: 'Tidak ada data atau baris yang dipilih!' };
    }

    var validItems = itemsData.filter(function(item) {
      if (!item) return false;
      if (item.isRowEmpty) return true;
      var hasFaktur = Boolean(item.noFaktur);
      var hasNo = Boolean(item.noPelanggan);
      var hasNama = Boolean(item.namaPelanggan);
      return hasFaktur || hasNo || hasNama;
    });

    if (validItems.length === 0) {
      return { success: false, message: 'Tidak ada baris yang valid untuk disimpan!' };
    }

    validItems = sortServerItemsByCustomerAZ(validItems);

    var targetSS = SpreadsheetApp.openByUrl(targetUrl);
    var targetSheetName = confRes.lhtSheet;
    var targetSheet = null;

    if (targetSheetName) {
      targetSheet = targetSS.getSheetByName(targetSheetName);
      if (!targetSheet) {
        return { 
          success: false, 
          message: 'Sheet dengan nama "' + targetSheetName + '" tidak ditemukan pada Spreadsheet tujuan! Pastikan nama sheet di kolom "LHT Sheet" sudah sesuai.' 
        };
      }
    } else {
      targetSheet = targetSS.getSheets()[0];
    }

    var lastRow = targetSheet.getLastRow();

    var insertRowIndex = lastRow;
    var finder = targetSheet.createTextFinder("TOTAL")
      .matchCase(false)
      .useRegularExpression(false);
    var allMatches = finder.findAll();
    if (allMatches && allMatches.length > 0) {
      var lastTotalCell = allMatches[allMatches.length - 1];
      insertRowIndex = lastTotalCell.getRow() - 1;
      if (insertRowIndex < 3) insertRowIndex = lastRow;
    }

    var GUNAKAN_NOMOR_KANVAS = true;
    var ISI_IDENTITAS_PADA_BARIS_KOSONG = true;

    var startNo = 1;
    if (!GUNAKAN_NOMOR_KANVAS && insertRowIndex >= 4) {
      var prevNoVal = targetSheet.getRange(insertRowIndex, 6).getValue();
      if (!isNaN(prevNoVal) && prevNoVal > 0) {
        startNo = parseInt(prevNoVal, 10) + 1;
      }
    }

    var formattedTanggal = formatDateDMY(headerConfig.tanggal);

    var rowsToInsert = [];
    validItems.forEach(function(item, index) {
      var nomorBaris = GUNAKAN_NOMOR_KANVAS
        ? (index + 1)
        : (startNo + index);

      if (item.isRowEmpty) {
        rowsToInsert.push([
          ISI_IDENTITAS_PADA_BARIS_KOSONG ? cleanText(headerConfig.perusahaan) : '',
          ISI_IDENTITAS_PADA_BARIS_KOSONG ? cleanText(headerConfig.penagih) : '',
          ISI_IDENTITAS_PADA_BARIS_KOSONG ? cleanText(headerConfig.divisi) : '',
          ISI_IDENTITAS_PADA_BARIS_KOSONG ? formattedTanggal : '',
          ISI_IDENTITAS_PADA_BARIS_KOSONG ? cleanText(headerConfig.admin) : '',
          nomorBaris,
          '', '', '', '', '', '', '', '',
          '', '', '', '', '', ''
        ]);
      } else {
        rowsToInsert.push([
          cleanText(headerConfig.perusahaan),
          cleanText(headerConfig.penagih),
          cleanText(headerConfig.divisi),
          formattedTanggal,
          cleanText(headerConfig.admin),
          nomorBaris,
          cleanText(item.noPelanggan),
          cleanText(item.namaPelanggan),
          item.umurHariNum ? item.umurHariNum : 0,
          cleanText(item.noFaktur),
          cleanText(item.tglFaktur),
          item.nilaiFaktur ? item.nilaiFaktur : 0,
          item.terbayar ? item.terbayar : 0,
          item.sisaPiutang ? item.sisaPiutang : 0,
          '', '', '', '', '', ''
        ]);
      }
    });

    targetSheet.insertRowsAfter(insertRowIndex, rowsToInsert.length);
    var destRange = targetSheet.getRange(insertRowIndex + 1, 1, rowsToInsert.length, rowsToInsert[0].length);

    if (insertRowIndex >= 3) {
      var sourceRange = targetSheet.getRange(insertRowIndex, 1, 1, rowsToInsert[0].length);
      sourceRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }

    var HAPUS_VALIDASI_BARIS_BARU = true;
    if (HAPUS_VALIDASI_BARIS_BARU) {
      destRange.clearDataValidations();
    } else {
      var rules = destRange.getDataValidations();
      for (var r = 0; r < rules.length; r++) {
        for (var c = 0; c < rules[r].length; c++) {
          if (rules[r][c] != null) {
            rules[r][c] = rules[r][c].copy().setAllowInvalid(true).build();
          }
        }
      }
      destRange.setDataValidations(rules);
    }

    destRange.setValues(rowsToInsert);
    SpreadsheetApp.flush();

    return { 
      success: true, 
      message: 'Berhasil menyimpan ' + validItems.length + ' baris ke sheet [' + targetSheet.getName() + ']!' 
    };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan ke Spreadsheet: ' + err.message };
  }
}

function parseIndoDate(val) {
  if (!val) return null;
  if (val instanceof Date) {
    return new Date(val.getTime());
  }
  var str = cleanText(val);
  if (!str) return null;

  var indoMonths = {
    'jan': 'Jan', 'feb': 'Feb', 'mar': 'Mar', 'apr': 'Apr',
    'mei': 'May', 'jun': 'Jun', 'jul': 'Jul', 'agu': 'Aug', 'agt': 'Aug',
    'sep': 'Sep', 'okt': 'Oct', 'nov': 'Nov', 'des': 'Dec'
  };

  var normalized = str.replace(/\b([a-zA-Z]{3})\b/g, function(match) {
    var lower = match.toLowerCase();
    return indoMonths[lower] ? indoMonths[lower] : match;
  });

  var parsed = new Date(normalized);
  if (!isNaN(parsed.getTime())) return parsed;

  var slashParts = str.split('/');
  if (slashParts.length === 3) {
    return new Date(Number(slashParts[2]), Number(slashParts[1]) - 1, Number(slashParts[0]));
  }
  return null;
}

function formatDateDMY(val) {
  if (!val) return '';
  var parsed = parseIndoDate(val);
  if (parsed && !isNaN(parsed.getTime())) {
    var dd = ('0' + parsed.getDate()).slice(-2);
    var mm = ('0' + (parsed.getMonth() + 1)).slice(-2);
    var yyyy = parsed.getFullYear();
    return dd + '/' + mm + '/' + yyyy;
  }
  var str = cleanText(val);
  if (str.indexOf('-') !== -1) {
    var parts = str.split('T')[0].split('-');
    if (parts.length === 3) return parts[2] + '/' + parts[1] + '/' + parts[0];
  }
  return str;
}