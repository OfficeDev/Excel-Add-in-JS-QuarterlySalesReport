# Excel 2016 用の四半期売上レポート作業ウィンドウ アドインのサンプル

_適用対象: Excel 2016_

これは、Excel 2016 でデータをワークシートに読み込んでから、基本的なグラフを作成する簡単な Excel 作業ウィンドウ アドインです。コード エディターと Visual Studio のいずれかを選択できます。

![四半期売上レポートのサンプル](../images/QuarterlySalesReport_report.PNG)

## お試しください。
### コード エディターのバージョン

アドインを展開してテストする最も簡単な方法は、ファイルをネットワーク共有にコピーすることです。

1.  ネットワーク共有にフォルダー (たとえば \\\MyShare\QuarterlySalesReport) を作成し、コード エディター フォルダー内のすべてのファイルをコピーします。 
2.  マニフェスト ファイルの <SourceLocation> 要素を編集して、手順 1 の共有場所を指すようにします。 
3.  マニフェスト (QuarterlySalesReportManifest.xml) をネットワーク共有 (たとえば \\\MyShare\\MyManifests) にコピーします。
4.  マニフェストを格納する共有の場所を、Excel で信頼されるアプリ カタログとして追加します。

    a.Excel を起動し、空のスプレッドシートを開きます。  
    
    b.**[ファイル]** タブを選択し、**[オプション]** を選択します。
    
    c.**[セキュリティ センター]** を選択し、**[セキュリティ センターの設定]** ボタンを選択します。
    
    d.**[信頼できるアドイン カタログ]** を選択します。
    
    e.**[カタログの URL]** ボックスで、手順 3 で作成したネットワーク共有のパスを入力し、**[カタログの追加]** を選択します。
    
   f. **[メニューに表示する]** チェック ボックスをオンにしてから **[OK]** を選択します。これらの設定は Office を次回起動したときに適用されることを示すメッセージが表示されます。 
        
5.  アドインをテストし、実行します。 

    a.Excel 2016 の **[挿入]** タブで、**[個人用アドイン]** を選択します。
    
    b.**[Office アドイン]** ダイアログ ボックスで、**[共有フォルダー]** を選択します。
    
    c.**[四半期売上レポートのサンプル]** > **[挿入]** の順に選択します。 次の図に示すように、現在のワークシートの右側の作業ウィンドウでアドインが開きます。 
        
  ![四半期売上レポートのサンプル](../images/QuarterlySalesReport_taskpane.PNG))

    d.**[ここをクリック]** ボタンをクリックして、次の図に示されているように、ワークシート内にデータとグラフを表示します。  グラフが動的に更新されることを確認するには、範囲に含まれるデータを変更します。
        
  ![四半期売上レポートのサンプル](../images/QuarterlySalesReport_report.PNG)

### Visual Studio のバージョン
1.  プロジェクトをローカル フォルダーにコピーし、Excel-Add-in-JS-QuarterlySalesReport.sln を開きます。
2.  F5 キーを押して、サンプル アドインをビルドおよび展開します。Excel が起動し、次の図に示すように、空白のワークシートの右側の作業ウィンドウでアドインが開きます。 
        
  ![四半期売上レポートのサンプル](../images/QuarterlySalesReport_taskpane.PNG)

3. **[ここをクリック]** ボタンをクリックして、次の図に示されているように、ワークシート内のデータとグラフを表示します。グラフが動的に更新されるようにするには、範囲に含まれるデータを変更します。 
        
  ![四半期売上レポートのサンプル](../images/QuarterlySalesReport_report.PNG)
        
## コード化する

このアドインをご自分で最初から作成する場合は、「[初めての Excel アドインを作成する](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)」を参照してください。


### 詳細を見る


1.  [Excel アドインのプログラミングの概要](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Excel のスニペット エクスプローラー](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Excel アドインのコード サンプル](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [Excel アドインの JavaScript API リファレンス](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [初めての Excel アドインを作成する](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)
