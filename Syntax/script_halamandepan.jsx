Main();

function Main() {
	app.findTextPreferences = NothingEnum.NOTHING;
	app.changeTextPreferences = NothingEnum.NOTHING;

	// Melakukan penggantian teks untuk setiap pasangan find/replace
	app.findTextPreferences.findWhat = "no_katalog";
	app.changeTextPreferences.changeTo = "no_katalog";
	document.changeText();

    app.findTextPreferences.findWhat = "nama_kecamatan";
    app.changeTextPreferences.changeTo = "nama_kecamatan";
    document.changeText();

    app.findTextPreferences.findWhat = "no_publikasi";
    app.changeTextPreferences.changeTo = "no_publikasi";
    document.changeText();

	// Menghapus preferensi pencarian dan penggantian setelah selesai
	app.findTextPreferences = NothingEnum.NOTHING;
	app.changeTextPreferences = NothingEnum.NOTHING;

	
}

alert("Selesai.");
