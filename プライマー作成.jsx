var doc = app.activeDocument;
var selection = app.selection;

try {
    var sw_primer = doc.swatches.getByName("RDG_PRIMER");
} catch (e) {
    alert("RDG_PRIMERスウォッチがドキュメント内に見つかりません。");
}

if (selection.length!==0){
    for(var i=0;i<selection.length;i++){
        if (selection[i].typename==="TextFrame") {
            var src = selection[i];
            primer_tf = selection[i].duplicate(src, ElementPlacement.PLACEAFTER);
            primer_tf.textRange.characterAttributes.fillColor = sw_primer.color;
        }
    }
} else {
    alert("テキストを選択してください");
}