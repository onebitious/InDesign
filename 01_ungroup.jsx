/*
  ●グループ解除スクリプト
    仕様
    ・レイヤーをロックを解除。
    ・アイテムのロックを解除。
    ・グループ化を解除。
    　　このとき、テキストフレーム内にアンカーオブジェクトがある場合と、空欄のテキストフレームの場合とでエラーになります。
    　　この場合のテキストフレームにロックをかけます。
    ・再度、ロックをかけた以外のフレームでグループ化を解除。
    ・最後に、全てのアイテムのロックを解除します。
    
    2018/7/3　大澤作成
*/

var doc = app.activeDocument; //アクティブドキュメント
var items = doc.allPageItems; //ページアイテム
var grpItem = doc.groups; //グループアイテム
var txtFrames = doc.textFrames; //テキストフレーム
var layer = doc.layers; //レイヤー

//▼テキストフレーム内アンカーオブジェクトの検索設定
app.findGrepPreferences = NothingEnum.nothing; //検索文字列初期化
app.findGrepPreferences.findWhat = "~a"; //検索文字をアンカーに設定→アンカーは正規表現で"~a"

//▼ダイアログを表示
var win = new Window("dialog", "警告！", [200, 100, 450, 210]);
win.add("statictext", [10, -50, 250, 80], "レイヤーのロック、オブジェクトのロック、\r\nグループ化を解除します。\r\n宜しいですか？");
win.cancelBtn = win.add("button", [35, 70, 115, 95], "キャンセル", {
    name: "cancel"
});
win.okBtn = win.add("button", [140, 70, 220, 95], "OK", {
    name: "ok"
});
win.center();
var btnFlg = win.show();
if (btnFlg == 2) {
    alert("処理を中断します。");
    exit();
}

//▼レイヤーのロック解除
for (var i = 0, layLen = layer.length; i < layLen; i++) {
    if (layer[i].locked == true) {
        layer[i].locked = false;
    }
}

//▼一度全てのアイテムロックを解除
for (var j = 0, itemsLen = items.length; j < itemsLen; j++) {
    if (items[j].locked == true) {
        items[j].locked = false;
    }
}

//▼グループ解除
try {
    for (var k = 0, itemsLen = items.length; k < itemsLen; k++) { //一度全てのグループ解除を試みる
        if (items[k].constructor.name == "Group") {
            items[k].ungroup(); //グループ解除
        }
    }
} catch (e) { //エラー処理 → アンカーオブジェクト、空欄のテキストフレームがあった場合ロックをかける
    for (var l = 0, txtFramesLen = txtFrames.length; l < txtFramesLen; l++) {
        if (txtFrames[l].contents.length == 0 || txtFrames[l].findGrep().length > 0) {
            txtFrames[l].locked = true; //ロックをかける
        }
    }
    for (var m = 0, itemsLen = items.length; m < itemsLen; m++) { //エラー対象以外で、再度グループ解除
        if (items[m].constructor.name == "Group") {
            doc.selection = grpItem; //グループアイテムを選択
            var sel = doc.selection;
            for (var n = 0, selLen = sel.length; n < selLen; n++) {
                sel[n].ungroup(); //グループ解除
            }
        }
    }
}

//▼最後に全てのアイテムロックを解除
var items = doc.allPageItems; //ページアイテム（再度宣言しないとエラーになる。上部でのグローバル宣言での不可）
for (var j = 0, itemsLen = items.length; j < itemsLen; j++) {
    if (items[j].locked == true) {
        items[j].locked = false;
    }
}

alert("処理が終わりました。\r\nテキストフレーム内のアンカーオブジェクトのグループは解除できません。");
