const black_c = new CMYKColor();
black_c.cyan = 63.99;
black_c.magenta = 69.27;
black_c.yellow = 68.81;
black_c.black = 78.82;

const blue_c = new CMYKColor();
blue_c.cyan = 94.51;
blue_c.magenta = 69.1;
blue_c.yellow = 0;
blue_c.black = 0;

const green_c = new CMYKColor();
green_c.cyan = 63.99;
green_c.magenta = 69.27;
green_c.yellow = 68.81;
green_c.black = 78.82;

const pink_c = new CMYKColor();
pink_c.cyan = 63.99;
pink_c.magenta = 69.27;
pink_c.yellow = 68.81;
pink_c.black = 78.82;

const orange_c = new CMYKColor();
orange_c.cyan = 63.99;
orange_c.magenta = 69.27;
orange_c.yellow = 68.81;
orange_c.black = 78.82;

const gold_c = new CMYKColor();
gold_c.cyan = 63.99;
gold_c.magenta = 69.27;
gold_c.yellow = 68.81;
gold_c.black = 78.82;

const brown_c = new CMYKColor();
brown_c.cyan = 63.99;
brown_c.magenta = 69.27;
brown_c.yellow = 68.81;
brown_c.black = 78.82;

const COLOR = {
    "黒" : black_c,
    "青" : blue_c,
    "緑" : green_c,
    "ピンク" : pink_c,
    "オレンジ" : orange_c,
    "金" : gold_c,
    "茶" : brown_c
}

//CSVをロード
function LoadCSV(){
    var file = File.openDialog("CSVを選択してください", "*.csv");
    if (!file){
        alert("キャンセルされました");
    } else {
        file.encoding = "UTF-8";
        if (file.open("r")){
            var text = file.read();
            file.close();

            if(text.charCodeAt(0)===0xFEFF){
                text = text.substring(1);
            }
            
            var name_with_info_list = text.split(/\r\n|\r|\n/);
            return name_with_info_list;
        }
    }
}

//CSVの解析
function CreateNameListAndInfoListDict(name_with_info_list){
    const info_line_num = 4;
    var name_list_and_info_list_dict= {};
    var prefix= "__AB__:";
    name_list_and_info_list_dict["font"] = name_with_info_list[0];
    name_list_and_info_list_dict["color"] = name_with_info_list[1];

    if (name_with_info_list[2].match(/^[+-]?(?:\d+\.?\d*|\.\d+)$/)){
        size = Number(name_with_info_list[2]);
        if (size){
            name_list_and_info_list_dict["size"] = size;
        } else {
            alert("サイズは数値のみを入力してください");
        }  
    } else {
        alert("サイズは半角で入力してください");
    }

    var head_x_margin_mm = name_with_info_list[3];
    var head_x_margin_px = UnitValue(head_x_margin_mm, "mm").as("px");
    name_list_and_info_list_dict["head_x_margin"] = head_x_margin_px;
    
    name_list_and_info_list_dict["classes"] = {};
    for(i=0;i<name_with_info_list.length-info_line_num;i++){
        if(name_with_info_list[i+info_line_num].indexOf(prefix)===0){
            var name_list = []
            var class_name = name_with_info_list[i+info_line_num].substring(prefix.length);
        } else {
            name_list.push(name_with_info_list[i+info_line_num]);
        }
        name_list_and_info_list_dict["classes"][class_name] = name_list;
    }
    return name_list_and_info_list_dict;
}

function CreateTextFrame(name_list_and_info_list_dict){
    var doc = app.activeDocument;

    function CreateArtboard(){
            //アートボード5枚で折り返す
            if((doc.artboards.length)%5==0) {
                ab_top_side = ab_bottom_side - ab_margin;
                ab_bottom_side = ab_top_side - ab_height;
                ab_left_side = base_left_side;
                ab_right_side = ab_left_side + ab_width;
                right_x_coordinate = right_x_coordinate_base;
                y_coordinate = y_coordinate_base = y_coordinate_base - ab_height - ab_margin;
           } else {
                ab_left_side = ab_right_side + ab_margin;
                ab_right_side = ab_left_side + ab_width;
                right_x_coordinate = right_x_coordinate + ab_width + ab_margin;
                y_coordinate = y_coordinate_base;
            }
            var rect = [ab_left_side, ab_top_side, ab_right_side, ab_bottom_side];
            ab = doc.artboards.add(rect);
            return ab;
    }

    //プライマーと白のスウォッチを取得
    try {
        var sw_primer = doc.swatches.getByName("RDG_PRIMER");
        var sw_white = doc.swatches.getByName("RDG_WHITE");
        COLOR["白"] = sw_white.color;
    } catch (e) {
        alert("RDG_PRIMERまたはRDG_WHITEスウォッチがドキュメント内に見つかりません。");
        return;
    }

    //アートボードの横幅と縦幅を取得
    var ab = doc.artboards[0];
    var r = ab.artboardRect;
    var ab_left_side = base_left_side = r[0];
    var ab_top_side = r[1];
    var ab_right_side = r[2];
    var ab_bottom_side = r[3];
    var ab_width = ab_right_side - ab_left_side;
    var ab_height = ab_top_side - ab_bottom_side;

    //CSVから取得した情報を展開
    var font = name_list_and_info_list_dict["font"];
    var color = name_list_and_info_list_dict["color"];
    var size = name_list_and_info_list_dict["size"];
    var head_x_margin = name_list_and_info_list_dict["head_x_margin"];
    var classes = name_list_and_info_list_dict["classes"];

    //textFrameを生成する位置とマージンを指定
    const head_x_coordinate = 380.409190390055;
    //名入れ位置の右辺のX座標を計算
    var right_x_coordinate = right_x_coordinate_base = head_x_coordinate - head_x_margin;//right_x_coordinateは名入れの右

    //中心のY座標。positionで使うときは、y_coordinate + height/2
    var y_coordinate = y_coordinate_base = 857.197028345381;
    var ab_margin = 50;
    const y_coordinate_margin = 62.3620984246008;
    

    is_first_loop = true;
    //名前のリストからtextFrameを生成
    for (var class_name in classes) {
        if(!is_first_loop){
            ab = CreateArtboard();
        }
        is_first_loop = false

        var ab_count = 1;
        name_list = classes[class_name];
        ab.name = class_name + "_" + ab_count.toString();
        for(i=0;i<name_list.length;i++){
            var primer_tf = doc.textFrames.add();
            primer_tf.contents = name_list[i];
            //primer_tf.textRange.characterAttributes.font = ;
            primer_tf.textRange.characterAttributes.size = size;
            var color_tf = primer_tf.duplicate();
            primer_tf.textRange.characterAttributes.fillColor = sw_primer.color;
            color_tf.textRange.characterAttributes.fillColor = COLOR["黒"];
            var outlined_primer_group = primer_tf.createOutline();
            var outlined_color_group = color_tf.createOutline();
            var group_width = outlined_primer_group.geometricBounds[2] - outlined_primer_group.geometricBounds[0];
            var group_height = outlined_primer_group.geometricBounds[1] - outlined_primer_group.geometricBounds[3];
            var x_coordinate = right_x_coordinate - group_width;
            outlined_primer_group.position = outlined_color_group.position = [x_coordinate, y_coordinate+(group_height/2)];
            y_coordinate = y_coordinate - y_coordinate_margin;

            //10本ずつアートボードを分ける
            if((i+1)%10==0){
                //10本の区切りの時
                //現在の要素がname_listの一番最後ならbreak
                if((i+1)==name_list.length){
                    break;
                }
                //違えばアートボードを作成
                ab = CreateArtboard();
                ab_count++;
                ab.name = class_name + "_" + ab_count.toString();
            }
        }
    }
}

name_with_info_list = LoadCSV();
name_list_and_info_list_dict = CreateNameListAndInfoListDict(name_with_info_list);
CreateTextFrame(name_list_and_info_list_dict);
