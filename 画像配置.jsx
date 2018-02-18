//▼サイズ指定ウィンドウ表示
var win = app.dialogs.add({
    name: 'ドキュメントサイズを選択'
});
var dlgClmn = win.dialogColumns.add();
var btnGrp = dlgClmn.radiobuttonGroups.add();
var a4Btn = btnGrp.radiobuttonControls.add({
    staticLabel: 'A4 (210 x 297mm)',
    minWidth: 150,
    checkedState: true
});
var a3Btn = btnGrp.radiobuttonControls.add({
    staticLabel: 'A3 (297 x 420mm)',
    minWidth: 150
});
var btnFlg = win.show();
if (btnFlg == false) {
    alert("キャンセルしました。");
    exit();
}
if (a4Btn.checkedState == true) {
    var docWidth = 210;
    var docHeight = 297;
} else {
    var docWidth = 297;
    var docHeight = 420;
}

//▼新規ドキュメント作成
var doc = app.documents.add(); //ドキュメント
doc.documentPreferences.pageWidth = docWidth; //横サイズ
doc.documentPreferences.pageHeight = docHeight; //縦サイズ
doc.documentPreferences.facingPages = false; //単ページ
doc.viewPreferences.horizontalMeasurementUnits = 2053991795; //水平方向の定規の単位をミリに指定
doc.viewPreferences.verticalMeasurementUnits = 2053991795; //水垂直方向の定規の単位をミリに指定

try {
    flag = false;
    while (flag == false) {
        var selFolder = Folder.selectDialog("画像の入ったフォルダを選択してください。"); //フォルダを選択
        if (!selFolder) {
            alert("キャンセルしました。");
            doc.close(SaveOptions.no); //保存しないで閉じる
            exit();
        } else {
            var myFiles = selFolder.getFiles(); //ファイルを取得
            var myFilesLen = myFiles.length; //総ファイル数
            flag = true;
        }
    }

    var myFilesLen = myFiles.length; //総ファイル数

    doc.documentPreferences.pagesPerDocument = myFilesLen; //ファイル数分ページを作成
    var myPage = doc.pages; //ドキュメント上のページ

//▼プログレッシブバー表示
var myProDialog = new Window('palette', '処理中...', [800, 480, 1200, 560]);
myProDialog.center();
myProDialog.myProgressBar = myProDialog.add("progressbar", [10, 30, 394, 45], 0, 100);
myProDialog.show();

    //▼ページ数分、グラフィックフレーム作成
    for (var i = 0, myPageLen = myPage.length; i < myPageLen; i++) {
        var myRect = myPage[i].rectangles.add(); //長方形を追加
        myRect.contentType = 1735553140; //グラフィックフレーム
        myRect.visibleBounds = ['5mm', '5mm', docHeight - 5, docWidth - 5];
        var myImg = myRect.place(myFiles[i]);
        myRect.fit(FitOptions.proportionally); //内容をフレーム内に収める
    }
} catch (e) {
    alert("何かしらのエラーが発生しました。\r\n処理するフォルダを選択していない可能性があります。\r\n処理を中断します。");
    doc.close(SaveOptions.no); //保存しないで閉じる
    exit();
}



//▼ページ数分、画像配置
for (var j = 0; j < myPageLen; j++) {
            //▼プログレッシブバーの値
        var processLength = myFilesLen; //総処理画像数
        myProDialog.myProgressBar.value += 100 / processLength;
        
    var pageObj = myPage[j].pageItems[0].graphics[0]; //ページ（フレーム）内画像
    var imgX1 = pageObj.visibleBounds[1]; //X1座標
    var imgY1 = pageObj.visibleBounds[0]; //Y1座標
    var imgX2 = pageObj.visibleBounds[3]; //X2座標
    var imgY2 = pageObj.visibleBounds[2]; //Y2座標

    var imgWidth = imgX2 - imgX1; //画像の幅を取得
    var imgHeighth = imgY2 - imgY1; //画像の高さを取得

    //▼長辺が縦になるよう画像を回転させる
    try {
        if (imgWidth > imgHeighth) {
            pageObj.rotationAngle = 90; //回転
            pageObj.fit(FitOptions.proportionally); //内容をフレーム内に収める
        } else {
            pageObj.fit(FitOptions.proportionally); //内容をフレーム内に収め
        }

        
    } catch (e) {
        alert("何かしらのエラーが発生しました。\r\n処理を中断します。\r\n「この値を入力すると、 1つ以上のオブジェクトがペーストボードから失われる可能性があります。」のエラーかも。。。？");
        exit;
    }

        
}
//▼プログレスバーを閉じる
myProDialog.close();
alert("処理が終わりました。");