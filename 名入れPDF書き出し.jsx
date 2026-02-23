var doc = app.activeDocument;

for (var i=0;i<doc.artboards.length;i++) {
    ab = doc.artboards[i];
    var ab_name = ab.name;
    var options = new PDFSaveOptions();
    options.artboardRange = (i+1).toString();
    var file = new File(doc.path + "/" + ab_name + ".pdf");
    doc.saveAs(file, options);
}
alert("終了しました");