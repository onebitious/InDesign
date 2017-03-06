/*////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
変数定義とエラー処理
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////*/
var myDoc = app.activeDocument; //ドキュメント
var mySel = myDoc.selection; //選択しているテキストフレーム

MAIN: {
    try { //エラー処理
        if (mySel.length == 0 || mySel[0].constructor.name != "TextFrame") {
            alert("テキストフレームを選択してください。");
            break MAIN;
        }
    } catch (e) {
        break MAIN;
    }

    var myCont = mySel[0].contents; //コンテンツ 確認用
    var myPara = mySel[0].paragraphs; //段落
    var myParaStyle = myDoc.paragraphStyles; //ドキュメントの段落スタイル

    /*////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    配列定義
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////*/
    var myHTMLArray = []; //HTMLタグ用配列
    var myCSSArray = []; //CSS用配列
    var myPropArray = []; //CSSプロパティ用配列
    var mySizeArray = []; //サイズ用配列
    var myFontFamilyArray = []; //フォントファミリー用配列
    var myFontWeightArray = []; //フォントウェイト用配列
    var myColorArray = []; //カラー用配列
    var myLeadingArray = []; //行送り用配列
    var myJustificationArray = []; //行揃え用配列

    /*////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    関数定義
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////*/
    //▼ゼロパディング
    function zeroPad(myNum) {
        var myNum = ('00' + myNum).slice(-2);
        return myNum;
    }

    //▼四捨五入
    function myRound(number, n) { //
        var myPow = Math.pow(10, n);
        return Math.round(number * myPow) / myPow;
    }

    /*////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    単位取得
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////*/
    //▼環境設定
    var myTypoMeasUni = myDoc.viewPreferences.typographicMeasurementUnits; //環境設定 -> 単位と増減値 -> 組版
    var myTextSizeMeasUni = myDoc.viewPreferences.textSizeMeasurementUnits; //環境設定 -> 単位と増減値 -> テキストサイズ

    //▼環境設定 -> 単位と増減値 -> 組版 に設定されている単位を変数に格納
    switch (myTypoMeasUni) {
    case 2054188905:
        myTypoMeasUni = 2054188905; //ポイント
        break;
    case 1516790048:
        myTypoMeasUni = 1516790048; //歯
        break;
    case 1514238068:
        myTypoMeasUni = 1514238068; //アメリカ式ポイント
        break;
    case 2051691808:
        myTypoMeasUni = 1514238068; //U
        break;
    case 2051170665:
        myTypoMeasUni = 2051170665; //Bai
        break;
    case 2051893612:
        myTypoMeasUni = 2051170665; //Mils
        break;
    }

    //▼環境設定 -> 単位と増減値 -> テキストサイズ に設定されている単位を変数に格納
    switch (myTextSizeMeasUni) {
    case 2054188905:
        myTextSizeMeasUni = 2054188905; //ポイント
        break;
    case 2054255973:
        myTextSizeMeasUni = 2054255973; //級
        break;
    case 1514238068:
        myTextSizeMeasUni = 1514238068; //アメリカ式ポイント
    }

    /*////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    InDesign2HTML
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////*/
    //▼一度、単位をポイントに変更
    myDoc.viewPreferences.typographicMeasurementUnits = 2054188905; //環境設定 -> 単位と増減値 -> 組版 をポイントに変更
    myDoc.viewPreferences.textSizeMeasurementUnits = 2054188905; //環境設定 -> 単位と増減値 -> テキストサイズ をポイントに変更

    for (var i = 0, myParaLen = myPara.length; myParaLen > i; i++) {
        var myPraContents = myPara[i].contents; //段落ごとのコンテンツ
        var myParaTagName = myPara[i].appliedParagraphStyle.name; //段落ごとの段落スタイル名
        var myStarTag = "<" + myParaTagName + ">"; //開始タグ
        var myEndTag = "</" + myParaTagName + ">"; //終了タグ
        var myHTMLSentence = myStarTag + myPraContents + myEndTag; //開始タグ＋コンテンツ＋終了タグ
        myHTMLArray.push(myHTMLSentence); //配列に追加
    }

    for (var j = 2, myParaStyleLen = myParaStyle.length; myParaStyleLen > j; j++) {
        var myParaStyleName = myParaStyle[j].name; //段落スタイル名
        myCSSArray.push(myParaStyleName);

        var myFontSize = myParaStyle[j].pointSize; //サイズ
        var myLeading = myParaStyle[j].leading; //行送り
        var myAutoLeading = myParaStyle[j].autoLeading; //自動行送り

        if (myLeading == 1635019116) { //行送りが自動の場合
            var myLeadingResult = "line-height: " + myRound(myAutoLeading, 2) / 100 + ";";
        } else {
            var myLeadingResult = "line-height: " + myRound(myLeading, 3) + "pt;"; //単位はポイントに固定
        }
        myLeadingArray.push(myLeadingResult); //配列に追加

        var myFontSizeResult = "font-size: " + myRound(myFontSize, 3) + "pt;"; //単位はポイントに固定
        mySizeArray.push(myFontSizeResult); //配列に追加   

        var myFontFamily = myParaStyle[j].appliedFont.name //フォントファミリー
        var myFontFamily = "font-family: " + myFontFamily + ";";
        myFontFamilyArray.push(myFontFamily); //配列に追加

        var myFontWeight = myParaStyle[j].appliedFont.fontStyleName //フォントウェイト
        var myFontWeight = "font-weight: " + myFontWeight + ";";
        myFontWeightArray.push(myFontWeight); //配列に追加

        var myColor = myParaStyle[j].fillColor.colorValue; //カラー
        var myColorRed = myColor[0].toString(16).toUpperCase(); //red値を取得（大文字のHEX値）
        var myColorRed = zeroPad(myColorRed); //ゼロパディング関数実行
        var myColorGreen = myColor[1].toString(16).toUpperCase(); //red値を取得（大文字のHEX値）
        var myColorGreen = zeroPad(myColorGreen); //ゼロパディング関数実行
        var myColorBlue = myColor[2].toString(16).toUpperCase(); //red値を取得（大文字のHEX値）
        var myColorBlue = zeroPad(myColorBlue); //ゼロパディング関数実行
        var myColor = "color: #" + myColorRed + myColorGreen + myColorBlue + ";"; //rgb値として取得
        myColorArray.push(myColor); //配列に追加

        var myJustification = myParaStyle[j].justification; //行揃え
        switch (myJustification) {
        case Justification.LEFT_ALIGN:
            myJustification = "text-align: left;";
            break;
        case Justification.CENTER_ALIGN:
            myJustification = "text-align: center;";
            break;
        case Justification.RIGHT_ALIGN:
            myJustification = "text-align: right;";
            break;
        case Justification.LEFT_JUSTIFIED:
            myJustification = "text-align: justify;";
            break;
        }
        myJustificationArray.push(myJustification); //配列に追加
        //alert(myJustification);//確認用
    }
    for (var k = 0, myCSSArrayLen = myCSSArray.length; k < myCSSArrayLen; k++) {
        var myCSSResult = myCSSArray[k] + "{" + mySizeArray[k] + myFontFamilyArray[k] + myFontWeightArray[k] + myColorArray[k] + myLeadingArray[k] + myJustificationArray[k] + "margin: 0;" + "}"; //CSS生成
        myPropArray.push(myCSSResult); //配列に追加
        //alert(myPropArray);//確認用
    }

    var myHTMLArray = myHTMLArray.toString(); //HTMLタグ用配列を文字列に変換　※変数上書き
    var myHTMLArray = myHTMLArray.replace(/,/g, ""); //文字列化したCSVのカンマを削除
    var myPropArray = myPropArray.toString(); //CSS用配列を文字列に変換　※変数上書き
    var myPropArray = myPropArray.replace(/,/g, ""); //文字列化したCSVのカンマを削除

    var MyHTMLRessult = '<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8"><title>Document</title><style>' + myPropArray + '</style></head><body>' + myHTMLArray + '</body></html>'; //HTML生成

    //▼単位を元に戻す
    myDoc.viewPreferences.typographicMeasurementUnits = myTypoMeasUni; //環境設定 -> 単位と増減値 -> 組版 をポイントに変更
    myDoc.viewPreferences.textSizeMeasurementUnits = myTextSizeMeasUni; //環境設定 -> 単位と増減値 -> テキストサイズ をポイントに変更

    /*////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ファイル保存
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////*/
    (function () { //無名関数を定義し実行
        var inputFileName = File.saveDialog("保存先とファイル名を指定して下さい");
        if (inputFileName == null) { //キャンセル処理
            alert("キャンセルしました。");
            return;
        }

        var inputFileName = inputFileName.toString(); //ファイル名に拡張子がないときの処理
        var myNewText = inputFileName.match(/.html/); //拡張子".txt"を付与
        if (myNewText) {
            var inputFileName = inputFileName; //入力したファイル名に拡張子が含まれる場合はそのまま
        } else {
            var inputFileName = inputFileName + ".html"; //入力したファイル名に拡張子が含まれない場合は付与
        }

        var fileObj = new File(inputFileName); //ファイル名を入力して保存
        var inputFileName = fileObj.open("w");
        if (inputFileName == true) {
            fileObj.encoding = "UTF-8"; //UTF-8を指定。これを入れないと異体字依存で書き出されない場合がある。
            fileObj.writeln(MyHTMLRessult); //書き出すテキストを連結
            fileObj.close();
            alert("処理が終わりました。");
        }
    })();
}