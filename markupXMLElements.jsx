/*
大澤作成

入れ子になったフレームは処理しません。
複数のフレームをグループ化し、入れ子にしたフレームを処理すると「何かしらのエラーが発生しました。」と表示されます。
例：Illustratorで作成したアウトライン化したロゴをグラフィックフレームに配置しているもの。

★20160309-01
矩形フレームのオブジェクトの属性が「グラフィック」と「割り当てなし」の場合、
「グラフィック」→"ItemImage"
「割り当てなし」→"ItemText"
とタグを付与するように改修。
★20160309-02
多角形フレームにも"ItemImage"が付与されるように改修。
★20160309-03
多角形フレームツール、楕円形フレームツールには、"ItemImage"を付与、
多角形ツール、楕円形ツールには、"ItemText"を付与するように改修。
★20160310-01
入れ子になったグループがある場合のエラーメッセージを変更。
処理を中断するように改修。
★20160311-01
処理後、追加したフレームにも処理をするようにエラー処理をコメントアウトしました。
*/

var myDoc = app.activeDocument;

var myAllPagesItemsLen = myDoc.allPageItems.length; //グループ化しているものも含め総フレーム数を取得
var myAllPagesItems = myDoc.allPageItems; //グループ化しているものも含め全てのフレームを取得
//alert(myAllPagesItemsLen);//確認用

try {
    //▼グループ解除処理
    try {
        for (var i = 0; i < myAllPagesItemsLen; i++) {
            if (myAllPagesItems[i].constructor.name == "Group") {
                myAllPagesItems[i].ungroup();
            }
        }
    } catch (e) {
        alert("入れ子になったグループオブジェクトが存在していると思われます。" + "\r" + "\r" + "予期せぬ結果になるので、処理を中断します。");
        exit();
    }

    //▼タグを作成
    try {
        myDoc.xmlTags.add("PageText", UIColors.MAGENTA); //PegeTextはマゼンタ
        myDoc.xmlTags.add("ItemImage", UIColors.BLUE); //ItemTextは青
        myDoc.xmlTags.add("ItemText", UIColors.GREEN); //ItemTextは緑
        myDoc.xmlTags.add("Spec", UIColors.RED); //Specは赤
    } catch (e) {
        //alert("すでに同じ名前のタグが存在します。");
        // exit();
    }

    //▼フレームを定義
    var myTextFramesLen = myDoc.textFrames.length; //テキストフレームの総数
    var myRectanglesLen = myDoc.rectangles.length; //長方形フレームの総数
    var myOvalsLen = myDoc.ovals.length; //楕円形フレームの総数
    var myPolygonLen = myDoc.polygons.length; //多角形フレームの総数
    var myTextFrames = myDoc.textFrames; //テキストフレーム
    var myRectangles = myDoc.rectangles; //長方形フレーム
    var myOvals = myDoc.ovals; //楕円形フレーム
    var myPolygon = myDoc.polygons; //多角形フレーム

    //フレームの種類によって付与する要素の定義
    var myRoot = myDoc.xmlElements.item("Root"); //Root
    var myItemText = myDoc.xmlTags.itemByName("ItemText"); //テキストフレームにItemTextを定義
    var myItemImage = myDoc.xmlTags.itemByName("ItemImage"); //長方形フレームにItemImageを定義

    //▼フレームにタグ付与処理
    try {
        for (i = 0; i < myTextFramesLen; i++) { //テキストフレームの処理
            myRoot.xmlElements.add(myItemText, myTextFrames[i]); //テキストフレームには"ItemText"を付与
        }
    } catch (e) {
        //alert("このフレームは既にタグ付されています。");
        //exit();
    }

    //▼長方形フレームツール、長方形ツールの処理
    try {
        for (j = 0; j < myRectanglesLen; j++) {
            if (myRectangles[j].contentType == ContentType.unassigned) { //「割り当てなし」フレーム
                myRoot.xmlElements.add(myItemText, myRectangles[j]); //「割り当てなし」フレームには"ItemText"を付与
            } else {
                myRoot.xmlElements.add(myItemImage, myRectangles[j]); //その他は"ItemImage"を付与
            }
        }
    } catch (e) {
        //alert("このフレームは既にタグ付されています。");
        //exit();
    }

    //▼楕円形フレームツール、楕円形ツールの処理
    try {
        for (k = 0; k < myOvalsLen; k++) {
            if (myOvals[k].contentType == ContentType.unassigned) { //「割り当てなし」フレーム           
                myRoot.xmlElements.add(myItemText, myOvals[k]); //楕円形ツール（「割り当てなし」）に"ItemText"を付与
            } else {
                myRoot.xmlElements.add(myItemImage, myOvals[k]); //楕円形フレームに"ItemImage"を付与
            }
        }
    } catch (e) {
        //alert("このフレームは既にタグ付されています。");
        //exit();
    }

    //▼多角形フレームツール、多角形ツールの処理
    try {
        for (l = 0; l < myPolygonLen; l++) {
            if (myPolygon[l].contentType == ContentType.unassigned) { //「割り当てなし」フレーム          
                myRoot.xmlElements.add(myItemText, myPolygon[l]); //多角形ツール（「割り当てなし」）に"ItemText"を付与
            } else {
                myRoot.xmlElements.add(myItemImage, myPolygon[l]); //多角形フレームツールに"ItemImage"を付与
            }
        }
    } catch (e) {
        //alert("このフレームは既にタグ付されています。");
        //exit();
    }

    //▼要素に属性名を付与
    for (var i = 0; i < myRoot.xmlElements.length; i++) {
        var myElementsName = myRoot.xmlElements[i].markupTag.name.toString();
        var myElementsObj = myRoot.xmlElements.item(i);
        try {
            myElementsObj.xmlAttributes.add("id", "");
        } catch (e) {
            //alert("属性はすでに付与されています。");
            //break;
        }
    }
} catch (e) {
    //alert("何かしらのエラーが発生しました。" + "\r" + "処理を中断します。");
    //exit();
} finally {
    alert("処理が終了しました。" + "\r" + "\r" + "グループ化は解除されてます。" + "\r" + "ご確認ください。");
}
