
Main();

function Main() {{
	app.findTextPreferences = NothingEnum.NOTHING;
	app.changeTextPreferences = NothingEnum.NOTHING;

	// Melakukan penggantian teks untuk setiap pasangan find/replace
	app.findTextPreferences.findWhat = "no_katalog";
	app.changeTextPreferences.changeTo = "5106045.3404130";
	document.changeText();

    app.findTextPreferences.findWhat = "nama_kecamatan";
    app.changeTextPreferences.changeTo = "Sleman";
    document.changeText();

    app.findTextPreferences.findWhat = "no_publikasi";
    app.changeTextPreferences.changeTo = "34040.24041";
    document.changeText();

	// Menghapus preferensi pencarian dan penggantian setelah selesai
	app.findTextPreferences = NothingEnum.NOTHING;
	app.changeTextPreferences = NothingEnum.NOTHING;

	
}}

alert("Selesai.");


