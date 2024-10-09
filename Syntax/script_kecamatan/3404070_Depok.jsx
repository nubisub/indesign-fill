
Main();

function Main() {
	showInitialDialog();
	var folderPath = Folder.selectDialog(
		"Pilih folder yang mengandung file Excel"
	);
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
						var keywords = [
							"Tabel",
							"Gambar",
							"Jenis Kelamin",
							"Sex",
							"Golongan",
							"Subsektor",
							"Kelompok",
						];

						function containsKeyword(content) {
							for (var i = 0; i < keywords.length; i++) {
								if (content.indexOf(keywords[i]) !== -1) {
									return true;
								}
							}
							return false;
						}

						if (!containsKeyword(firstCellContent)) {
							var bodyRowCount = table.bodyRowCount;
							var lastIndex = bodyRowCount - 1;
							var jmlrow = 25 - lastIndex;

							for (var n = 0; n < jmlrow; n++) {
								table.rows.add(LocationOptions.AFTER, table.rows[lastIndex]);
								table.recompose();
							}
							firstCellContent = table.rows[3].cells[0].contents;
							tablesIndex.push({
								page: i + 1,
								textFrameIndex: tf + 1,
								tableIndex: j + 1,
								table: table,
								firstCellContent: firstCellContent,
							});
						}
					}
				}
			}
		}
	}

	alert("banyak tabel di indesign : " + tablesIndex.length);

	var dictKec = {
		3404010001: "Sumberrahayu",
		3404010002: "Sumbersari",
		3404010003: "Sumber Agung",
		3404010004: "Sumberarum",
		3404020001: "Sendang Mulyo",
		3404020002: "Sendang Arum",
		3404020003: "Sendang Rejo",
		3404020004: "Sendangsari",
		3404020005: "Sendangagung",
		3404030001: "Margoluwih",
		3404030002: "Margodadi",
		3404030003: "Margomulyo",
		3404030004: "Margoagung",
		3404030005: "Margokaton",
		3404040001: "Sidorejo",
		3404040002: "Sidoluhur",
		3404040003: "Sidomulyo",
		3404040004: "Sidoagung",
		3404040005: "Sidokarto",
		3404040006: "Sidoarum",
		3404040007: "Sidomoyo",
		3404050001: "Balecatur",
		3404050002: "Ambarketawang",
		3404050003: "Banyuraden",
		3404050004: "Nogotirto",
		3404050005: "Trihanggo",
		3404060001: "Tirtoadi",
		3404060002: "Sumberadi",
		3404060003: "Tlogoadi",
		3404060004: "Sendangadi",
		3404060005: "Sinduadi",
		3404070001: "Catur Tunggal",
		3404070002: "Maguwoharjo",
		3404070003: "Condong Catur",
		3404080001: "Sendang Tirto",
		3404080002: "Tegal Tirto",
		3404080003: "Jogo Tirto",
		3404080004: "Kali Tirto",
		3404090001: "Sumber Harjo",
		3404090002: "Wukir Harjo",
		3404090003: "Gayam Harjo",
		3404090004: "Sambi Rejo",
		3404090005: "Madu Rejo",
		3404090006: "Boko Harjo",
		3404100001: "Purwo Martani",
		3404100002: "Tirto Martani",
		3404100003: "Taman Martani",
		3404100004: "Selo Martani",
		3404110001: "Wedomartani",
		3404110002: "Umbulmartani",
		3404110003: "Widodo Martani",
		3404110004: "Bimo Martani",
		3404110005: "Sindumartani",
		3404120001: "Sari Harjo",
		3404120002: "Sinduharjo",
		3404120003: "Minomartani",
		3404120004: "Suko Harjo",
		3404120005: "Sardonoharjo",
		3404120006: "Donoharjo",
		3404130001: "Catur Harjo",
		3404130002: "Triharjo",
		3404130003: "Tridadi",
		3404130004: "Pandowo Harjo",
		3404130005: "Tri Mulyo",
		3404140001: "Banyu Rejo",
		3404140002: "Tambak Rejo",
		3404140003: "Sumber Rejo",
		3404140004: "Pondok Rejo",
		3404140005: "Moro Rejo",
		3404140006: "Margo Rejo",
		3404140007: "Lumbung Rejo",
		3404140008: "Merdiko Rejo",
		3404150001: "Bangun Kerto",
		3404150002: "Donokerto",
		3404150003: "Giri Kerto",
		3404150004: "Wono Kerto",
		3404160001: "Purwo Binangun",
		3404160002: "Candi Binangun",
		3404160003: "Harjo Binangun",
		3404160004: "Pakem Binangun",
		3404160005: "Hargo Binangun",
		3404170001: "Wukir Sari",
		3404170002: "Argo Mulyo",
		3404170003: "Glagah Harjo",
		3404170004: "Kepuh Harjo",
		3404170005: "Umbul Harjo",
		3404070: "Depok",
	};

	alert("banyak file : " + excelFiles.length);
	for (var i = 0; i < excelFiles.length; i++) {
		var filePath = excelFiles[i].fsName;
		// alert(filePath);
		var data = GetDataFromExcelPC(filePath, splitChar, 2);
		var total = GetDataFromExcelPC(filePath, splitChar, 1);

		var kabIndex = -1;
		for (var k = 0; k < data[0].length; k++) {
			if (data[0][k] === "kec") {
				kabIndex = k;
				break;
			}
		}

		if (kabIndex === -1) {
			alert(
				"Kolom 'kab' tidak ditemukan. Silakan periksa kembali data header Anda."
			);
			return;
		}

		var totalIndex = -1;
		for (var l = 0; l < total[0].length; l++) {
			if (total[0][l] === "kec") {
				totalIndex = l;
				break;
			}
		}

		if (totalIndex === -1) {
			alert(
				"Kolom 'kab' tidak ditemukan. Silakan periksa kembali data header Anda."
			);
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
			if (data[m][kabIndex] === "3404070") {
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
			if (total[p][totalIndex] === "3404070") {
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

	var vbs = "Public s\r";
	vbs += "Function ReadFromExcel()\r";
	vbs += 'Set objExcel = CreateObject("Excel.Application")\r';
	vbs += 'Set objBook = objExcel.Workbooks.Open("' + excelFilePath + '")\r';
	vbs +=
		"Set objSheet =  objExcel.ActiveWorkbook.WorkSheets(" + sheetNumber + ")\r";
	vbs += "objExcel.Visible = False\r";
	vbs += "matrix = objSheet.UsedRange\r";
	vbs += "maxDim0 = UBound(matrix, 1)\r";
	vbs += "maxDim1 = UBound(matrix, 2)\r";
	vbs += "For i = 1 To maxDim0\r";
	vbs += "For j = 1 To maxDim1\r";
	vbs += "If j = maxDim1 Then\r";
	vbs += "s = s & matrix(i, j)\r";
	vbs += "Else\r";
	vbs += 's = s & matrix(i, j) & "' + splitChar + '"\r';
	vbs += "End If\r";
	vbs += "Next\r";
	vbs += "s = s & vbCr\r";
	vbs += "Next\r";
	vbs += "objBook.Close\r";
	vbs += "Set objSheet = Nothing\r";
	vbs += "Set objBook = Nothing\r";
	vbs += "Set objExcel = Nothing\r";
	vbs += "End Function\r";
	vbs += "Function SetArgValue()\r";
	vbs += 'Set objInDesign = CreateObject("InDesign.Application")\r';
	vbs += 'objInDesign.ScriptArgs.SetValue "excelData", s\r';
	vbs += "End Function\r";
	vbs += "ReadFromExcel()\r";
	vbs += "SetArgValue()\r";

	if (appVersionNum > 5) {
		app.doScript(
			vbs,
			ScriptLanguage.VISUAL_BASIC,
			undefined,
			UndoModes.FAST_ENTIRE_SCRIPT
		);
	} else {
		app.doScript(vbs, ScriptLanguage.VISUAL_BASIC);
	}

	var str = app.scriptArgs.getValue("excelData");
	app.scriptArgs.clear();

	var tempArrLine,
		line,
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
	// alert("Tabel ke-" + indexTable);
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
	// table.rows[0].cells[0].contents = "Desa/Kelurahan\rVillage/Urban Village";

	// alert("sekarang tabel ke - " + indexTable);
	var columnsProcessed = [];
	var remainingCSVData = [];

	function updateTableFromCSV(headers, table) {
		if (table.rows[2].cells[0].contents === "(1)") {
			var headers = table.rows[2].cells;
			for (var j = 0; j < headers.length; j++) {
				var headerText = headers[j].contents;
				// alert("header : " + headerText);
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
									var parts = normalizedContent.split(".");
									var integerPart = parts[0];
									var decimalPart = parts.length > 1 ? parts[1] : "";

									integerPart = integerPart.replace(
										/\B(?=(\d{3})+(?!\d))/g,
										"."
									);
									cellContent = decimalPart
										? integerPart + "," + decimalPart
										: integerPart;
								}
								if (cellContent === "0" || cellContent === "0,00") {
									cellContent = "-";
								}
								if (cellContent === "") {
									cellContent = "-";
								}
								table.rows[m].cells[j].contents = cellContent;

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
									var parts = normalizedContent.split(".");
									var integerPart = parts[0];
									var decimalPart = parts.length > 1 ? parts[1] : "";

									integerPart = integerPart.replace(
										/\B(?=(\d{3})+(?!\d))/g,
										"."
									);
									cellContent = decimalPart
										? integerPart + "," + decimalPart
										: integerPart;
								}
								if (cellContent === "0" || cellContent === "0,00") {
									cellContent = "-";
								}
								if (cellContent === "") {
									cellContent = "-";
								}
								table.rows[m].cells[j].contents = cellContent;

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
		// alert("Sisa data : " + remainingCSVData.join(","));
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
	var initialDialog = new Window(
		"dialog",
		"Automasi Input dari Excel ke Indesign",
		undefined,
		{ resizeable: false }
	);
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
	infoPanel.add(
		"statictext",
		undefined,
		"Script ini dibuat untuk meningkatkan kolaborasi antar satker dalam menyusun publikasi"
	);
	var versiPanel = initialDialog.add("panel", undefined, "Versi");
	versiPanel.orientation = "column";
	versiPanel.alignChildren = ["left", "top"];
	versiPanel.margins = 10;
	versiPanel.preferredSize.width = panelWidth;

	var versiGroup = versiPanel.add("group");
	versiGroup.orientation = "column";
	versiGroup.alignChildren = ["left", "top"];

	versiGroup.add(
		"statictext",
		undefined,
		"Developer : BPS Kabupaten Lampung Utara"
	);
	versiGroup.add(
		"statictext",
		undefined,
		"Latest Version cek di github.com/gerynastiar"
	);

	var buttonGroup = initialDialog.add("group");
	buttonGroup.orientation = "row";
	buttonGroup.alignChildren = ["center", "center"];
	buttonGroup.spacing = 10;

	var okButton = buttonGroup.add("button", undefined, "Asiap Yay!");
	okButton.onClick = function () {
		initialDialog.close();
	};
	initialDialog.show();
}
alert("Selesai.");


