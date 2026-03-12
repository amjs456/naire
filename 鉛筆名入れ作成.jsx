var doc = app.activeDocument;
var selection = app.selection;

if (selection.length !== 1) {
    alert("テキストボックスを1つだけ選択してください");
}

if (selection[0].typename!=="TextFrame") {
    alert("テキストオブジェクトを選択してください");
}

var selected_tf = selection[0];

x_coordinate = 1000;
y_coordinate = 0;
x_margin = 100;

var outlined_group = selected_tf.createOutline();

var outlined_group_width = outlined_group.geometricBounds[2] - outlined_group.geometricBounds[0];
outlined_group.position = [x_coordinate - (outlined_group_width / 2), y_coordinate];

for(var i=0;i<11;i++){
    x_coordinate-=x_margin;
    var duplicated_outlined_group = outlined_group.duplicate();
    duplicated_outlined_group.position = [x_coordinate - (outlined_group_width / 2), y_coordinate];
}