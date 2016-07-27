# Excel 2016 的每季銷售報表工作窗格增益集範例

_適用版本：Excel 2016_

這個簡單的 Excel 工作窗格增益集可在 Excel 2016 中將一些資料載入工作表中，並建立一個基本圖表。共有兩種型態︰程式碼編輯器和 Visual Studio。

![每季銷售報表範例](../images/QuarterlySalesReport_report.PNG)

## 進行測試
### 程式碼編輯器版本

部署及測試增益集的最簡單方式，是將檔案複製到網路共用。

1.  在網路共用中建立資料夾 (例如，\\\MyShare\QuarterlySalesReport)，然後複製 [程式碼編輯器] 資料夾中的所有檔案。 
2.  編輯資訊清單檔案的 <SourceLocation> 元素，讓它指向步驟 1 的共用位置。 
3.  將資訊清單 (QuarterlySalesReportManifest.xml) 複製到網路共用 (例如 \\\MyShare\MyManifests)。
4.  在 Excel 中，將包含資訊清單的共用位置新增為受信任的應用程式目錄。

    a.啟動 Excel，並開啟空白的試算表。  
    
    b.選擇 **[檔案]** 索引標籤，然後選擇 **[選項]**。
    
    c.選擇 **[信任中心]**，然後選擇 **[信任中心設定]** 按鈕。
    
    d.選擇 **[受信任的增益集目錄]**。
    
    e.在 **[目錄 URL]** 方塊中，輸入您在步驟 3 建立的網路共用路徑，然後選擇 **[新增目錄]**。
    
   f. 選取 [在功能表中顯示]<e /> 核取方塊，然後選擇 [確定]<e />。接著會顯示訊息，通知您下次啟動 Office 時就會套用您的設定。 
        
5.  測試並執行增益集。 

    a.在 Excel 2016 的 **[插入]** 索引標籤上，選擇 **[我的增益集]**。
    
    b.在 **[Office 增益集]** 對話方塊中，選擇 **[共用資料夾]**。
    
    c.選擇 **[每季銷售報告範例]** > **[插入]**。 增益集會在目前的工作表右側的工作窗格中開啟，如下圖所示。 
        
  ![每季銷售報表範例](../images/QuarterlySalesReport_taskpane.PNG))

    d.按一下 **[按我!]** 按鈕來轉譯工作表內的資料和圖表，如下圖所示。  若要查看動態更新的圖表，只要變更範圍內的資料。
        
  ![每季銷售報表範例](../images/QuarterlySalesReport_report.PNG)

### Visual Studio 版本
1.  將專案複製到本機資料夾，並在 Visual Studio 中開啟Excel-Add-in-JS-QuarterlySalesReport.sln。
2.  按 F5 建置及部署範例增益集。Excel 會啟動，且增益集會在工作表右側空白部分的工作窗格中開啟，如下圖所示。 
        
  ![每季銷售報表範例](../images/QuarterlySalesReport_taskpane.PNG)

3. 按一下 [按我!]<e /> 按鈕來轉譯工作表內的資料和圖表，如下圖所示。若要查看動態更新的圖表，只要將範圍內的資料變更。 
        
  ![每季銷售報表範例](../images/QuarterlySalesReport_report.PNG)
        
## 編碼

如果您想要自行從頭寫入這個增益集，請參閱[建立第一個 Excel 增益集](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)。


### 深入了解


1.  [Excel 增益集程式設計概觀](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Excel 的程式碼片段總管](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Excel 增益集程式碼範例](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [Excel 增益集 JavaScript API 參考](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [建立第一個 Excel 增益集](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)
