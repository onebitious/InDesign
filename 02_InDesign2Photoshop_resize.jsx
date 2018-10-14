#target 'indesign'
#targetengine 'imgResize'

var doc = app.activeDocument, //最前面のドキュメント
    sel = doc.selection, //選択しているオブジェクト
    docName = doc.name; //ドキュメント名

//▼エラー処理
if (sel.length < 1) {
    alert("画像フレームを選択してから処理を実行してください。");
    exit();
}

for (var i = 0, selLen = sel.length; i < selLen; i++) { //選択しているオブジェクト分繰り返す
    try {
        if (sel[i].contentType != 1735553140 || sel[i].contentType == 1952412773 || sel[i].contentType == 1970168179) {
            alert("画像フレーム以外が選択されています。処理を中断します。");
            exit();
        } else if (sel[i].allGraphics[0] == undefined) {
            alert("空の画像フレームを選択している可能性があります。\r\n処理を中断します。");
            exit();
        }
    } catch (e) {
        alert("画像フレーム以外が選択されています。処理を中断します。");
        exit();
    }
}

process(endCall); //処理実行

//▼処理
function process(callback) {
for (var i = 0, selLen = sel.length; i < selLen; i++) { //選択しているオブジェクト分繰り返す
    var filePath = sel[i].allGraphics[0].itemLink.filePath; //ファイルパス
    var replacePath = filePath.replace(/\\/g, '\/'); // '\' を '/' に置換

    var horiSize = (sel[i].allGraphics[0].absoluteHorizontalScale) / 100; //幅を取得
    var verSize = (sel[i].allGraphics[0].absoluteVerticalScale) / 100; //高さを取得

    //▼BridgeTalk　Photoshopへ
    var bridgeTalkPH = new BridgeTalk();
    bridgeTalkPH.target = 'photoshop';
    bridgeTalkPH.body = uneval(imgOpen) + '(' + '"' + replacePath + '"' + ',' + horiSize + ',' + verSize + ');'; //Photoshop リサイズ～上書き保存の関数実行
    bridgeTalkPH.send();
}
callback(); //コールバック関数
}

//▼処理終了の表示関数
function endCall() {
    var bridgeTalkPH2 = new BridgeTalk();
    bridgeTalkPH2.target = 'photoshop';
    bridgeTalkPH2.body = 'alert("リサイズ処理が終わりました。");';
    bridgeTalkPH2.send();
}

//▼Photoshop リサイズ～上書き保存の関数
function imgOpen(replacePath, horiSize, verSize) {
    var openFile = new File(replacePath);
    open(openFile);
    var doc = app.activeDocument;
    var docName = doc.name,
        docWidth = doc.width.value,
        docHeight = doc.height.value;
    doc.resizeImage(docWidth * horiSize, docHeight * verSize);
    doc.save();
    doc.close();
}