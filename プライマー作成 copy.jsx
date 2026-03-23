var doc = app.activeDocument;
var selection = app.selection;

const alert_message = "テキストまたはアウトライン化したテキストを選択してください";

try {
    var sw_primer = doc.swatches.getByName("RDG_PRIMER");
} catch (e) {
    alert("RDG_PRIMERスウォッチがドキュメント内に見つかりません。");
}

function setColorRecursive(item, color) {
    alert(item.typename);
    switch(item.typename){
        case "pageItem":
            setColorRecursive(item.compoundPathItems, color);
            break;
        
        case "coumpoundPathItems":
            setColorRecursive(item.pathItems, color);
            break;
        
        case "pathItem":
            item.filled = true;
            item.fillColor = color;
            break;
    }
}

if (selection.length!==0){
    for(var i=0;i<selection.length;i++){
        var src = selection[i];
        if (src.typename==="TextFrame") {
            primer_tf = selection[i].duplicate(src, ElementPlacement.PLACEAFTER);
            primer_tf.textRange.characterAttributes.fillColor = sw_primer.color;
        } else if(src.typename === "GroupItem") {
            primer_group = src.duplicate(src, ElementPlacement.PLACEAFTER);
            for (var j = 0; j < primer_group.pageItems.length; j++) {
                compoundPathItems = primer_group.pageItems[j];
                setColorRecursive(compoundPathItems, sw_primer.color);
                /*
                for(var k = 0; k < compoundPathItems.pathItems.length; k++) {
                    pathItem = compoundPathItems.pathItems[k];
                    pathItem.filled = true;
                    pathItem.fillColor = sw_primer.color;
                }
                */
            }
        } else {
            alert(alert_message);
        }
    }
} else {
    alert(alert_message);
}