Main();

function Main() {
  showInitialDialog();
  var folderPath = Folder.selectDialog("Pilih folder yang mengandung file Excel");
  if (!folderPath) {
    alert("Folder tidak dipilih.");
    return;
  }
  var excelFiles = folderPath.getFiles("*.xlsx");
  if (excelFiles.length === 0) {
    alert(excelFiles.length);
    alert("Tidak ada file Excel di folder yang dipilih.");
    return;
  }
  var splitChar = ";";
  var indexTableCSV = 0;
  var doc = app.activeDocument;

  var tablesIndex = [];

  for (var i = 0; i < doc.pages.length; i++) {
    var page = doc.pages[i];
    var allTextFrames = page.textFrames.everyItem().getElements();

    for (var tf = 0; tf < allTextFrames.length; tf++) {
      var textFrame = allTextFrames[tf];
      if (textFrame.tables.length > 0) {
        var tables = textFrame.tables;
        for (var j = 0; j < tables.length; j++) {
          var table = tables[j];

          if (table.rows.length > 0 && table.columns.length > 0) {
            var firstCellContent = table.rows[0].cells[0].contents;
            var keywords = ["Tabel", "Gambar", "Jenis Kelamin", "Sex", "Golongan", "Subsektor", "Kelompok","Lampiran"];

            function containsKeyword(content) {
              for (var i = 0; i < keywords.length; i++) {
                if (content.indexOf(keywords[i]) !== -1) {
                  return true;
                }
              }
              return false;
            }

            if (!containsKeyword(firstCellContent)) {
            
              firstCellContent = table.rows[3].cells[0].contents;
              tablesIndex.push({
                page: i + 1,
                textFrameIndex: tf + 1,
                tableIndex: j + 1,
                table: table,
                firstCellContent: firstCellContent
              });
            }

          }
        }
      }
    }
  }

  tablesIndex.splice(13, 2);
  alert("banyak tabel di indesign : " + tablesIndex.length);

  var dictKec = {
    "1806010": "Bukit Kemuning",
    "1806011": "Abung Tinggi",
    "1806020": "Tanjung Raja",
    "1806030": "Abung Barat",
    "1806031": "Abung Tengah",
    "1806032": "Abung Kunang",
    "1806033": "Abung Pekurun",
    "1806040": "Kotabumi",
    "1806041": "Kotabumi Utara",
    "1806042": "Kotabumi Selatan",
    "1806050": "Abung Selatan",
    "1806051": "Abung Semuli",
    "1806052": "Blambangan Pagar",
    "1806060": "Abung Timur",
    "1806061": "Abung Surakarta",
    "1806070": "Sungkai Selatan",
    "1806071": "Muara Sungkai",
    "1806072": "Bunga Mayang",
    "1806073": "Sungkai Barat",
    "1806074": "Sungkai Jaya",
    "1806080": "Sungkai Utara",
    "1806081": "Hulusungkai",
    "1806082": "Sungkai Tengah",
    "1806": "Lampung Utara"
  };

  alert("banyak file : " + excelFiles.length);
  for (var i = 0; i < excelFiles.length; i++) {
    var filePath = excelFiles[i].fsName;
    // alert(filePath);
    var data = GetDataFromExcelPC(filePath, splitChar, 2);
    var total = GetDataFromExcelPC(filePath, splitChar, 1);

    var kabIndex = -1;
    for (var k = 0; k < data[0].length; k++) {
      if (data[0][k] === "kab") {
        kabIndex = k;
        break;
      }
    }

    if (kabIndex === -1) {
      alert("Kolom 'kab' tidak ditemukan. Silakan periksa kembali data header Anda.");
      return;
    }

    var totalIndex = -1;
    for (var l = 0; l < total[0].length; l++) {
      if (total[0][l] === "kab") {
        totalIndex = l;
        break;
      }
    }

    if (totalIndex === -1) {
      alert("Kolom 'kab' tidak ditemukan. Silakan periksa kembali data header Anda.");
      return;
    }

    var idKomoditasIndexData = -1;
    for (var k = 0; k < data[0].length; k++) {
      if (data[0][k] === "id_komoditas") {
        idKomoditasIndexData = k;
        break;
      }
    }
    if (idKomoditasIndexData !== -1) {
      for (var k = 0; k < data.length; k++) {
        data[k].splice(idKomoditasIndexData, 1);
      }
    }

    var idKomoditasIndexTotal = -1;
    for (var l = 0; l < total[0].length; l++) {
      if (total[0][l] === "id_komoditas") {
        idKomoditasIndexTotal = l;
        break;
      }
    }

    if (idKomoditasIndexTotal !== -1) {
      for (var l = 0; l < total.length; l++) {
        total[l].splice(idKomoditasIndexTotal, 1);
      }
    }

    var filteredData = [];
    for (var m = 1; m < data.length; m++) {
      if (data[m][kabIndex] === "1806") {
        filteredData.push(data[m]);
      }
    }

    for (var n = 0; n < filteredData.length; n++) {
      filteredData[n].splice(0, kabIndex + 1);
    }

    if (filteredData.length > 0) {
      var headers = [];
      for (var o = 1; o <= filteredData[0].length; o++) {
        headers.push("(" + o + ")");
      }
      filteredData.unshift(headers);
    }

    var filteredTotal = [];
    for (var p = 1; p < total.length; p++) {
      if (total[p][totalIndex] === "1806") {
        filteredTotal.push(total[p]);
      }
    }

    for (var q = 0; q < filteredTotal.length; q++) {
      filteredTotal[q].splice(0, totalIndex);
    }

    if (filteredTotal.length > 0) {
      var headers = [];
      for (var r = 1; r <= filteredTotal[0].length; r++) {
        headers.push("(" + r + ")");
      }
      filteredTotal.unshift(headers);
    }

    filteredData.push(filteredTotal[1]);
    for (var s = 0; s < filteredData.length; s++) {
      var kode = filteredData[s][0];
      if (dictKec[kode]) {
        filteredData[s][0] = dictKec[kode];
      }
    }

    indexTableCSV = replaceTableColumnDataFromHeader(
      filteredData,
      indexTableCSV,
      tablesIndex
    );
  }
}

function GetDataFromExcelPC(excelFilePath, splitChar, sheetNumber) {
  if (typeof splitChar === "undefined") var splitChar = ";";
  if (typeof sheetNumber === "undefined") var sheetNumber = "1";
  var appVersionNum = Number(String(app.version).split(".")[0]);

  var vbs = 'Public s\r';
  vbs += 'Function ReadFromExcel()\r';
  vbs += 'Set objExcel = CreateObject("Excel.Application")\r';
  vbs += 'Set objBook = objExcel.Workbooks.Open("' + excelFilePath + '")\r';
  vbs += 'Set objSheet =  objExcel.ActiveWorkbook.WorkSheets(' + sheetNumber + ')\r';
  vbs += 'objExcel.Visible = False\r';
  vbs += 'matrix = objSheet.UsedRange\r';
  vbs += 'maxDim0 = UBound(matrix, 1)\r';
  vbs += 'maxDim1 = UBound(matrix, 2)\r';
  vbs += 'For i = 1 To maxDim0\r';
  vbs += 'For j = 1 To maxDim1\r';
  vbs += 'If j = maxDim1 Then\r';
  vbs += 's = s & matrix(i, j)\r';
  vbs += 'Else\r';
  vbs += 's = s & matrix(i, j) & "' + splitChar + '"\r';
  vbs += 'End If\r';
  vbs += 'Next\r';
  vbs += 's = s & vbCr\r';
  vbs += 'Next\r';
  vbs += 'objBook.Close\r';
  vbs += 'Set objSheet = Nothing\r';
  vbs += 'Set objBook = Nothing\r';
  vbs += 'Set objExcel = Nothing\r';
  vbs += 'End Function\r';
  vbs += 'Function SetArgValue()\r';
  vbs += 'Set objInDesign = CreateObject("InDesign.Application")\r';
  vbs += 'objInDesign.ScriptArgs.SetValue "excelData", s\r';
  vbs += 'End Function\r';
  vbs += 'ReadFromExcel()\r';
  vbs += 'SetArgValue()\r';

  if (appVersionNum > 5) {
    app.doScript(vbs, ScriptLanguage.VISUAL_BASIC, undefined, UndoModes.FAST_ENTIRE_SCRIPT);
  }
  else {
    app.doScript(vbs, ScriptLanguage.VISUAL_BASIC);
  }

  var str = app.scriptArgs.getValue("excelData");
  app.scriptArgs.clear();

  var tempArrLine, line,
    data = [],
    tempArrData = str.split("\r");

  for (var i = 0; i < tempArrData.length; i++) {
    line = tempArrData[i];
    if (line == "") continue;
    tempArrLine = line.split(splitChar);
    data.push(tempArrLine);
  }

  return data;
}

function contains(array, element) {
  for (var i = 0; i < array.length; i++) {
    if (array[i] === element) {
      return true;
    }
  }
  return false;
}

function replaceTableColumnDataFromHeader(csvData, indexTable, tables) {
  // progressInput(indexTable, tables.length);
  var table = tables[indexTable].table;
  if (table.parent instanceof TextFrame) {
    var textFrame = table.parent;
    app.activeWindow.activePage = textFrame.parentPage;
    app.activeWindow.zoom(ZoomOptions.FIT_PAGE);
    table.cells[0].insertionPoints[0].select();
  }
  if (table.rows[2].cells[0].contents === "(1)") {
    var headers = table.rows[2].cells;
  } else {
    var headers = table.rows[1].cells;
  }
  var lastChangedRow = -1;
  var searchText = csvData[0];
  table.rows[0].cells[0].contents = "Kecamatan\rDistrict";

  // alert("sekarang tabel ke - " + indexTable);
  var columnsProcessed = [];
  var remainingCSVData = [];

  function updateTableFromCSV(headers, table) {
    if (table.rows[2].cells[0].contents === "(1)") {
      var headers = table.rows[2].cells;
      for (var j = 0; j < headers.length; j++) {
        var headerText = headers[j].contents;
        for (var k = 0; k < searchText.length; k++) {
          // alert("header : "+ headerText);
          // alert("search : " + searchText[k]);
          if (headerText == searchText[k]) {

            for (var m = 3; m < table.rows.length; m++) {
              var csvIndex = m - 2;
              if (csvData[csvIndex] && csvData[csvIndex][k] !== undefined) {

                var cellContent = csvData[csvIndex][k];
                var normalizedContent = cellContent;
                if (!isNaN(normalizedContent)) {
                  var parts = normalizedContent.split('.');
                  var integerPart = parts[0];
                  var decimalPart = parts.length > 1 ? parts[1] : '';

                  integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
                  cellContent = decimalPart ? integerPart + ',' + decimalPart : integerPart;
                }
                if (cellContent === "0" || cellContent === "0,00") {
                  cellContent = "-"
                }
                table.rows[m].cells[j].contents = cellContent;

                if (j != 0) {
                  var textObj = table.rows[m].cells[j].texts[0].paragraphs[0];
                  textObj.appliedParagraphStyle = app.activeDocument.paragraphStyles.itemByName("isi tabel");

                  textObj.justification = Justification.LEFT_ALIGN;
                  textObj.justification = Justification.CENTER_ALIGN;
                }

                lastChangedRow = Math.max(lastChangedRow, m);
              }

            }
            columnsProcessed.push(k);
            break;
          }
        }
      }
    } else {
      var headers = table.rows[1].cells;
      for (var j = 0; j < headers.length; j++) {
        var headerText = headers[j].contents;
        for (var k = 0; k < searchText.length; k++) {
          // alert("header : "+ headerText);
          // alert("search : " + searchText[k]);
          if (headerText == searchText[k]) {
            for (var m = 2; m < table.rows.length; m++) {
              var csvIndex = m - 1;
              if (csvData[csvIndex] && csvData[csvIndex][k] !== undefined) {
                var cellContent = csvData[csvIndex][k];
                var normalizedContent = cellContent;
                if (!isNaN(normalizedContent)) {
                  var parts = normalizedContent.split('.');
                  var integerPart = parts[0];
                  var decimalPart = parts.length > 1 ? parts[1] : '';

                  integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
                  cellContent = decimalPart ? integerPart + ',' + decimalPart : integerPart;
                }
                if (cellContent === "0" || cellContent === "0,00") {
                  cellContent = "-"
                }
                table.rows[m].cells[j].contents = cellContent;

                if (j != 0) {
                  var textObj = table.rows[m].cells[j].texts[0].paragraphs[0];
                  textObj.appliedParagraphStyle = app.activeDocument.paragraphStyles.itemByName("isi tabel");

                  textObj.justification = Justification.LEFT_ALIGN;
                  textObj.justification = Justification.CENTER_ALIGN;
                }

                lastChangedRow = Math.max(lastChangedRow, m);
              }
            }
            columnsProcessed.push(k);
            break;
          }
        }
      }
    }

  }

  updateTableFromCSV(headers, table);

  if (columnsProcessed.length < searchText.length) {
    for (var i = 0; i < csvData.length; i++) {
      remainingCSVData[i] = [];
      remainingCSVData[i].push(csvData[i][0]);
      for (var k = 0; k < searchText.length; k++) {
        if (!contains(columnsProcessed, k) && k !== 0) {
          remainingCSVData[i].push(csvData[i][k]);
        }
      }
    }
    if (lastChangedRow !== -1) {
      changeRowColor(table, lastChangedRow);
      removeRowsAfter(table, lastChangedRow);
    }

    if (indexTable + 1 < tables.length) {
      return replaceTableColumnDataFromHeader(
        remainingCSVData,
        indexTable + 1,
        tables
      );
    } else {
      alert("Tidak ada tabel yang cukup untuk memuat semua data CSV.");
    }
  }
  if (lastChangedRow !== -1) {
    changeRowColor(table, lastChangedRow);
    removeRowsAfter(table, lastChangedRow);
  }
  return indexTable + 1;
}

function contains(array, value) {
  for (var i = 0; i < array.length; i++) {
    if (array[i] === value) {
      return true;
    }
  }
  return false;
}


function changeRowColor(table, lastChangedRow) {
  var doc = app.activeDocument;
  var colorName = "CustomColor";
  var customColor;

  try {
    customColor = doc.colors.itemByName(colorName);
    customColor.name;
  } catch (e) {
    customColor = doc.colors.add({
      name: colorName,
      model: ColorModel.PROCESS,
      space: ColorSpace.CMYK,
      colorValue: [60, 0, 100, 10],
    });
  }

  var lastRowCells = table.rows[lastChangedRow].cells;
  for (var n = 0; n < lastRowCells.length; n++) {
    lastRowCells[n].fillColor = "CustomColor";
    lastRowCells[n].fillTint = 100;
  }
}

function removeRowsAfter(table, lastChangedRow) {
  for (var m = table.rows.length - 1; m > lastChangedRow; m--) {
    table.rows[m].remove();
  }
}
function showInitialDialog() {
  var panelWidth = 800;
  var initialDialog = new Window("dialog", "Automasi Input dari Excel ke Indesign", undefined, { resizeable: false });
  initialDialog.orientation = "column";
  initialDialog.alignChildren = ["center", "top"];
  initialDialog.spacing = 10;
  initialDialog.margins = 16;
  var infoPanel = initialDialog.add("panel", undefined, "Brief Explanation");
  infoPanel.orientation = "column";
  infoPanel.alignChildren = ["left", "top"];
  infoPanel.margins = 10;
  infoPanel.spacing = 4;
  infoPanel.preferredSize.width = panelWidth;
  infoPanel.add("statictext", undefined, "Script ini dibuat untuk meningkatkan kolaborasi antar satker dalam menyusun publikasi");
  var versiPanel = initialDialog.add("panel", undefined, "Versi");
  versiPanel.orientation = "column";
  versiPanel.alignChildren = ["left", "top"];
  versiPanel.margins = 10;
  versiPanel.preferredSize.width = panelWidth;

  var versiGroup = versiPanel.add("group");
  versiGroup.orientation = "column";
  versiGroup.alignChildren = ["left", "top"];

  versiGroup.add("statictext", undefined, "Developer : BPS Kabupaten Lampung Utara");
  versiGroup.add("statictext", undefined, "Latest Version cek di github.com/gerynastiar");

  var buttonGroup = initialDialog.add("group");
  buttonGroup.orientation = "row";
  buttonGroup.alignChildren = ["center", "center"];
  buttonGroup.spacing = 10;

  var okButton = buttonGroup.add("button", undefined, "Asiap Yay!");
  okButton.onClick = function () {
    initialDialog.close();
  }
  initialDialog.show();
}
// function progressInput(onProgress, Total) {
//   var dialog = new Window("dialog", "Progress Input");
//   dialog.orientation = "column";
//   dialog.alignChildren = ["fill", "top"];
//   dialog.margins = 20;

//   var messagePanel = dialog.add("panel", undefined, "Progress");
//   messagePanel.orientation = "column";
//   messagePanel.alignChildren = ["fill", "top"];
//   messagePanel.margins = 15;

//   messagePanel.add("statictext", undefined, "Progress Sedang di Tabel Ke-" + onProgress + " dari Total " + Total + " tabel.");

//   dialog.show();
// }

alert("Selesai.");
