function getSmartFillData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var confSheet = ss.getSheetByName('conf');
    if (!confSheet) {
      return { success: false, message: 'Sheet conf tidak ditemukan!', rules: [] };
    }

    var values = confSheet.getDataRange().getValues();
    var rules = parseSmartFillFromConfValues(values);

    return {
      success: true,
      rules: rules
    };
  } catch (err) {
    return {
      success: false,
      message: 'Gagal memuat data SmartFill: ' + err.message,
      rules: []
    };
  }
}

function safeTrimSF(val) {
  if (val === undefined || val === null) return '';
  return String(val).trim();
}

function parseSmartFillFromConfValues(values) {
  if (!values || values.length <= 1) return [];

  var headers = values[0];
  var idxPenagih = -1;
  var idxMinggu = -1;
  var idxHari = -1;
  var idxNoPelanggan = -1;

  for (var h = 0; h < headers.length; h++) {
    var colName = safeTrimSF(headers[h]).toLowerCase();
    if (colName === 'nama penagih') {
      idxPenagih = h;
    } else if (colName === 'minggu') {
      idxMinggu = h;
    } else if (colName === 'hari') {
      idxHari = h;
    } else if (colName === 'no. pelanggan' || colName === 'no pelanggan') {
      idxNoPelanggan = h;
    }
  }

  if (idxPenagih === -1 || idxMinggu === -1 || idxHari === -1 || idxNoPelanggan === -1) {
    return [];
  }

  var rules = [];

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var namaPenagih = safeTrimSF(row[idxPenagih]);
    var minggu = safeTrimSF(row[idxMinggu]);
    var hari = normalizeHariName(safeTrimSF(row[idxHari]));
    var rawNoPelanggan = safeTrimSF(row[idxNoPelanggan]);

    if (!namaPenagih || !minggu || !hari) continue;

    var customerCodes = [];
    if (rawNoPelanggan !== '') {
      var splitParts = rawNoPelanggan.split('|');
      for (var c = 0; c < splitParts.length; c++) {
        var cleanCode = safeTrimSF(splitParts[c]);
        if (cleanCode) customerCodes.push(cleanCode);
      }
    }

    rules.push({
      penagih: namaPenagih,
      minggu: minggu,
      hari: hari,
      customerCodes: customerCodes
    });
  }

  return rules;
}

function normalizeHariName(hariStr) {
  if (!hariStr) return '';
  var clean = safeTrimSF(hariStr).toLowerCase().replace(/['’`]/g, '');
  if (clean === 'jumat') return "Jum'at";
  if (clean === 'senin') return 'Senin';
  if (clean === 'selasa') return 'Selasa';
  if (clean === 'rabu') return 'Rabu';
  if (clean === 'kamis') return 'Kamis';
  if (clean === 'sabtu') return 'Sabtu';
  if (clean === 'minggu' || clean === 'ahad') return 'Minggu';
  return hariStr;
}