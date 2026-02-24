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
green_c.cyan = 82.06;
green_c.magenta = 5.67;
green_c.yellow = 100;
green_c.black = 0.35;

const pink_c = new CMYKColor();
pink_c.cyan = 2.54;
pink_c.magenta = 77.4;
pink_c.yellow = 3.68;
pink_c.black = 0;

const orange_c = new CMYKColor();
orange_c.cyan = 2.24;
orange_c.magenta = 49.27;
orange_c.yellow = 100;
orange_c.black = 0;

const gold_c = new CMYKColor();
gold_c.cyan = 32.57;
gold_c.magenta = 38.12;
gold_c.yellow = 100;
gold_c.black = 5.98;

const brown_c = new CMYKColor();
brown_c.cyan = 44.06;
brown_c.magenta = 69.09;
brown_c.yellow = 95.34;
brown_c.black = 42.26;

const COLOR = {
    "黒" : black_c,
    "青" : blue_c,
    "緑" : green_c,
    "ピンク" : pink_c,
    "オレンジ" : orange_c,
    "金" : gold_c,
    "茶" : brown_c
}

const FONT_NAME = {
    "筆記体":"ShelleyAllegroBT Regular",
    "角ゴシック体":"DFHSGothic W3-WINP-RKSJ-H",
    "丸ゴシック体":"HGMaruGothicMPRO",
    "楷書体":"FGKaishoNT M",
    "明朝体":"KozMinPro Regular-90ms-RKSJ-H"
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
            
            var parsed_csv_list=[];
            var row_list = text.split(/\r\n|\r|\n/);
            for(var i=0;i<row_list.length;i++){
                var col_list = row_list[i].split(",")
                parsed_csv_list.push(col_list);
            }
            return parsed_csv_list;
        }
    }
}

//CSVの解析
function CreateNameListAndInfoListDict(parsed_csv_list){
    const info_line_num = 1;
    var name_list_and_info_list_dict= {};
    var class_prefix= "__AB__:";

    //var head_x_margin_mm = name_with_info_list[3];
    //var head_x_margin_px = UnitValue(head_x_margin_mm, "mm").as("px");
    //name_list_and_info_list_dict["head_x_margin"] = head_x_margin_px;
    
    //class別にクラスの名簿を作る
    var class_list = [];

    //クラス内の名前を入れるリスト
    var name_and_info_list = [];

    //クラス名の変数
    var class_name = "";


    var is_first_loop = true;
    //全クラス全員forで回す
    for(i=0;i<parsed_csv_list.length-info_line_num;i++){
        //__AB__:がない場合name_and_info_listへ、__AB__:がある場合name_and_info_listをclass_listへ
        if(parsed_csv_list[i+info_line_num][0].indexOf(class_prefix)===0){
            //最初の__AB__:でclass_listへ入れるのを避ける。2回目の__AB__:以降はclass_listへ
            if(!is_first_loop){
                class_list.push([class_name,name_and_info_list]);
            }
            //クラス名の取得
            class_name = parsed_csv_list[i+info_line_num][0].substring(class_prefix.length);
            //name_and_info_listを空に
            name_and_info_list = [];
        } else {
            //名前をname_and_info_listへ
            name_and_info_list.push(parsed_csv_list[i+info_line_num]);

            //最後の名前の場合、name_and_info_listをclass_listへ
            if(i===parsed_csv_list.length-info_line_num-1){
                class_list.push([class_name,name_and_info_list]);
            }
        }
        is_first_loop = false;
    }
    return class_list;
}

function CreateTextFrame(class_list){
    var doc = app.activeDocument;

    function CreateArtboard(){
            //アートボード5枚で折り返す
            if((doc.artboards.length)%5==0) {
                ab_top_side = ab_bottom_side - ab_margin;
                ab_bottom_side = ab_top_side - ab_height;
                ab_left_side = base_left_side;
                ab_right_side = ab_left_side + ab_width;
                head_x_coordinate = head_x_coordinate_base;
                y_coordinate = y_coordinate_base = y_coordinate_base - ab_height - ab_margin;
           } else {
                ab_left_side = ab_right_side + ab_margin;
                ab_right_side = ab_left_side + ab_width;
                head_x_coordinate = head_x_coordinate + ab_width + ab_margin;
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

    //textFrameを生成する位置とマージンを指定
    var head_x_coordinate = head_x_coordinate_base = 380.409190390055;
    
    //中心のY座標。positionで使うときは、y_coordinate + height/2
    var y_coordinate = y_coordinate_base = 857.197028345381;

    var ab_margin = 50;
    const y_coordinate_margin = 62.3620984246008;
    

    is_first_loop = true;
    //名前のリストからtextFrameを生成
    for(var class_i=0;class_i<class_list.length;class_i++){
        if(!is_first_loop){
            ab = CreateArtboard();
        }
        is_first_loop = false

        var ab_count = 1;
        ab.name = class_list[class_i][0]+ab_count.toString();
        
        for(var name_i=0;name_i<class_list[class_i][1].length;name_i++){
            var name = class_list[class_i][1][name_i][0];
            var font = class_list[class_i][1][name_i][1];
            var color = class_list[class_i][1][name_i][2];
            var size = parseInt(class_list[class_i][1][name_i][3]);
            var head_x_margin = UnitValue(parseInt(class_list[class_i][1][name_i][4]), "mm").as("px");

            var primer_tf = doc.textFrames.add();
            primer_tf.contents = name;
            //primer_tf.textRange.characterAttributes.textFont = FONT_NAME[font];
            primer_tf.textRange.characterAttributes.size = size;
            var color_tf = primer_tf.duplicate();
            primer_tf.textRange.characterAttributes.fillColor = sw_primer.color;
            color_tf.textRange.characterAttributes.fillColor = COLOR[color];
            var outlined_primer_group = primer_tf.createOutline();
            var outlined_color_group = color_tf.createOutline();
            var group_width = outlined_primer_group.geometricBounds[2] - outlined_primer_group.geometricBounds[0];
            var group_height = outlined_primer_group.geometricBounds[1] - outlined_primer_group.geometricBounds[3];
            var x_coordinate = head_x_coordinate - head_x_margin - group_width;
            outlined_primer_group.position = outlined_color_group.position = [x_coordinate, y_coordinate+(group_height/2)];
            y_coordinate = y_coordinate - y_coordinate_margin;

            if((name_i+1)%10===0){
                if((name_i+1)===class_list[class_i][1].length){
                    break;
                }
                ab = CreateArtboard();
                ab_count++;
                ab.name = class_list[class_i][0];
            }
        }
    }
}

parsed_csv_list = LoadCSV();
name_list_and_info_list_dict = CreateNameListAndInfoListDict(parsed_csv_list);
CreateTextFrame(name_list_and_info_list_dict);
