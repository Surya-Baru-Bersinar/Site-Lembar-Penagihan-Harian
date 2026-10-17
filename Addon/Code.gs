function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Pembuat Tagihan LHT - PTM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getConfData() {
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
    var headerText = String(headers[h] || '').trim().toLowerCase();
    if (headerText === 'tautan lht') {
      colTautanIdx = h;
    }
    if (headerText === 'lht sheet') {
      colSheetIdx = h;
    }
  }

  for (var i = 1; i < values.length; i++) {
    if (values[i][0]) admins.push(String(values[i][0]).trim());
    if (values[i][1]) penagihs.push(String(values[i][1]).trim());
    if (values[i][2]) depos.push(String(values[i][2]).trim());
    if (values[i][3]) perusahaans.push(String(values[i][3]).trim());
    if (values[i][4]) divisis.push(String(values[i][4]).trim());
    
    if (!tautanLHT && values[i][colTautanIdx]) {
      tautanLHT = String(values[i][colTautanIdx]).trim();
    }

    if (!lhtSheet && values[i][colSheetIdx]) {
      lhtSheet = String(values[i][colSheetIdx]).trim();
    }
  }

  function unique(arr) { return Array.from(new Set(arr)).filter(Boolean); }

  return {
    success: true,
    admins: unique(admins),
    penagihs: unique(penagihs),
    depos: unique(depos),
    perusahaans: unique(perusahaans),
    divisis: unique(divisis),
    tautanLHT: tautanLHT,
    lhtSheet: lhtSheet
  };
}

function getPiutangData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Piutang');
  if (!sheet) return { success: false, message: 'Sheet Piutang tidak ditemukan!' };

  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return { success: true, items: [], waktuPembaruan: '-' };

  var waktuPembaruan = '-';
  var headers = values[0];
  var colWaktuIdx = -1;
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).trim().toLowerCase() === 'waktu pembaruan') {
      colWaktuIdx = h;
      break;
    }
  }

  if (colWaktuIdx !== -1 && values.length > 1) {
    waktuPembaruan = sheet.getRange(2, colWaktuIdx + 1).getDisplayValue() || '-';
  } else {
    waktuPembaruan = sheet.getRange("L2").getDisplayValue() || '-';
  }

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var items = [];
  var rawRows = values.slice(1);

  rawRows.forEach(function(row) {
    var noPelanggan = String(row[0] || '').trim();
    var noFaktur = String(row[1] || '').trim();
    if (!noPelanggan && !noFaktur) return;

    var tglFakturRaw = row[2];
    var tglFakturDate = null;
    if (tglFakturRaw instanceof Date) {
      tglFakturDate = new Date(tglFakturRaw.getTime());
    } else if (typeof tglFakturRaw === 'string' && tglFakturRaw.trim() !== '') {
      tglFakturDate = new Date(tglFakturRaw);
    }

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

    var nilaiFaktur = Number(row[4]) || 0;
    var sisaPiutang = Number(row[5]) || 0;
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
      namaPelanggan: String(row[7] || '').trim(),
      namaPenjual: String(row[8] || '').trim(),
      namaKontak: String(row[9] || '').trim()
    });
  });

  items.sort(function(a, b) {
    var nameA = a.namaPelanggan.toLowerCase();
    var nameB = b.namaPelanggan.toLowerCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    return a.noFaktur.localeCompare(b.noFaktur);
  });

  return { 
    success: true, 
    items: items, 
    waktuPembaruan: waktuPembaruan 
  };
}

function simpanKeSpreadsheetTarget(headerConfig, itemsData) {
  try {
    var confRes = getConfData();
    var targetUrl = confRes.tautanLHT;
    if (!targetUrl) return { success: false, message: 'Tautan LHT pada sheet conf belum diisi!' };

    if (!itemsData || !Array.isArray(itemsData) || itemsData.length === 0) {
      return { success: false, message: 'Tidak ada data piutang yang dipilih!' };
    }

    var validItems = itemsData.filter(function(item) {
      return item && !item.isRowEmpty && (item.noFaktur || item.noPelanggan || item.namaPelanggan);
    });

    if (validItems.length === 0) {
      return { success: false, message: 'Tidak ada data faktur valid yang dipilih!' };
    }

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
      .searchByRegex(false);
    var allMatches = finder.findAll();
    if (allMatches.length > 0) {
      var lastTotalCell = allMatches[allMatches.length - 1];
      insertRowIndex = lastTotalCell.getRow() - 1;
      if (insertRowIndex < 3) insertRowIndex = lastRow;
    }

    var startNo = 1;
    if (insertRowIndex >= 4) {
      var prevNoVal = targetSheet.getRange(insertRowIndex, 6).getValue();
      if (!isNaN(prevNoVal) && prevNoVal > 0) {
        startNo = parseInt(prevNoVal) + 1;
      }
    }

    var formattedTanggal = formatDateDMY(headerConfig.tanggal);

    var rowsToInsert = [];
    validItems.forEach(function(item, index) {
      rowsToInsert.push([
        headerConfig.perusahaan || '',
        headerConfig.penagih || '',
        headerConfig.divisi || '',
        formattedTanggal,
        headerConfig.admin || '',
        startNo + index,
        item.noPelanggan || '',
        item.namaPelanggan || '',
        item.umurHariNum || 0,
        item.noFaktur || '',
        item.tglFaktur || '',
        item.nilaiFaktur || 0,
        item.terbayar || 0,
        item.sisaPiutang || 0,
        '', '', '', '', '', ''
      ]);
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
      message: 'Berhasil menyimpan ' + validItems.length + ' data ke sheet [' + targetSheet.getName() + ']!' 
    };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan ke Spreadsheet: ' + err.message };
  }
}

function formatDateDMY(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var dd = ('0' + val.getDate()).slice(-2);
    var mm = ('0' + (val.getMonth() + 1)).slice(-2);
    var yyyy = val.getFullYear();
    return dd + '/' + mm + '/' + yyyy;
  }
  var str = String(val).trim();
  if (str.indexOf('-') !== -1) {
    var parts = str.split('T')[0].split('-');
    if (parts.length === 3) return parts[2] + '/' + parts[1] + '/' + parts[0];
  }
  return str;
}