var doc = app.activeDocument; //アクティブドキュメント
var img = doc.allGraphics; //ドキュメント上の全てのグラフィックフレーム

var imgPropertiesArray = []; //結果の配列を準備

for (var i = 0, imgLen = img.length; i < imgLen; i++) { //グラフィックフレームの数だけ繰り返す
    if (doc.allGraphics[i].itemLink.status == 1819242340) { //更新が必要なフレームだった場合
        doc.allGraphics[i].itemLink.update(); //更新
        var imgName = doc.allGraphics[i].itemLink.name, //画像名の取得
            imgHoriSize = doc.allGraphics[i].absoluteHorizontalScale, //幅の取得
            imgVerSize = doc.allGraphics[i].absoluteVerticalScale; //高さの取得

        if (imgHoriSize == imgVerSize) {
            var comment = "縦横比：○"; //縦横比が同じ場合のコメント
        } else {
            var comment = "縦横比：×"; //縦横比が異なる場合のコメント
        }

        var imgResult = "ファイル名：" + imgName + "\r\n幅：" + imgHoriSize + "%\r\n高：" + imgVerSize + "%\r\n" + comment + "\r\n";
        imgPropertiesArray.push(imgResult);
    }
}

var myArraymyStyleText = imgPropertiesArray.toString(); //スタイルを格納した配列を文字列に変換
var repmyStyleText = myArraymyStyleText.replace(/,/g, "\r\n"); //文字列化したCSVのカンマを改行に置換

//▼ファイル名にする日付を取得
var myDate = new Date(),
    myYear = myDate.getFullYear(), //年を取得
    myMonth = myDate.getMonth() + 1, //月を取得
    myDay = myDate.getDate(), //日を取得
    myHours = myDate.getHours(), //時を取得
    myMinutes = myDate.getMinutes(), //分を取得
    mySeconds = myDate.getSeconds(), //秒を取得
    nowTime = myYear + "年" + myMonth + "月" + myDay + "日" + myHours + "時" + myMinutes + "分" + mySeconds + "秒" + "_log";

/*==============================================================
=======================ファイル保存================================
================================================================*/
(function () { //即時関数を定義し実行
    var fileObj = new File("~/Desktop/" + nowTime + ".txt"); //ファイル名を入力して保存
    var inputFileName = fileObj.open("w");
    if (inputFileName == true) {
        fileObj.encoding = "UTF-8"; //UTF-8を指定。これを入れないと異体字依存で書き出されない場合がある。
        fileObj.writeln(repmyStyleText); //書き出すテキストを連結
        fileObj.close();
        alert("処理が終わりました。\r\nデスクトップにログファイルを保存しました。");
    }
})();

