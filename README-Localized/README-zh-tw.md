# 愚蠢的故事 Word 增益集範例︰載入檔案並使用內容控制項

此 Word 增益集將會為您示範如何：

1. 從服務載入 docx 檔案的清單，並填入下拉式方塊控制項，以檔案名稱做為選項。
2. 從服務載入 docx 檔案，並將它插入 Word 文件。
3. 載入內容控制項集合，並且根據內容控制項建立輸入方塊。
4. 根據輸入方塊中的值，更新內容控制項集合的文字值。
5. 使用 Office UI Fabric，建立順暢的 Word 使用者經驗。

> 附註：咯咯笑是執行這個範例的副作用。

## 必要條件

若要使用愚蠢的故事 Word 增益集範例，需要有下列各項。

* [node.js](https://nodejs.org) 以提供 docx 檔案。
* [npm](https://www.npmjs.com/) 以安裝相依性。
* JQuery，針對 Office UI Fabric [下拉式](dev.office.com/fabric/components/dropdown)元件。
* Word 2016 或支援 Word Javascript API 的任何用戶端。這個範例會執行必要檢查，以查看它是否在支援的主機中執行。

## 啟動 Web 應用程式

1. 在命令列上執行專案根目錄中的 ```npm install```，隨著 Node 的封裝管理員 (npm) 一起安裝專案相依項目。
2. 執行專案根目錄中的 ```node server.js``` 以啟動開發伺服器。增益集會在 127.0.0.1:8080 執行。

### 在 Word for Mac 2016 上設定和執行

1. 在 Users/Library/Containers/com.microsoft.word/Data/Documents/ 中建立一個稱為 "wef" 的資料夾
2. 在 wef 資料夾中放置資訊清單 (Users/Library/Containers/com.microsoft.word/Data/Documents/wef)
3. 在 Mac 上開啟 Word 2016，然後按一下 [插入] 索引標籤 > [我的增益集] 下拉式清單。您應該會看到下拉式清單列出增益集。選取它，它將載入增益集。

### 在 Word for Windows 2016 上設定和執行

1. 建立網路共用，或[共用網路資料夾](https://technet.microsoft.com/zh-tw/library/cc770880.aspx)，並將 [word-add-in-sillystories.xml](word-add-in-sillystories.xml) 資訊清單檔案放置在其中。您已經在這個時候部署您的增益集。現在，您需要讓 Word 知道哪裡可以找到此增益集。
2. 啟動 Word 並開啟一個文件。
3. 選擇 [檔案]<e /> 索引標籤，然後選擇 [選項]<e />。
4. 選擇 [信任中心]<e />，然後選擇 [信任中心設定]<e /> 按鈕。
5. 選擇 [受信任的增益集目錄]<e />。
6. 在 [目錄 URL]<e /> 方塊中，輸入包含 word-add-in-sillystories.xml 的資料夾共用的網路路徑，然後選擇 [新增目錄]<e />。
7. 選取 [顯示於功能表中]<e /> 核取方塊，然後選擇 [確定]<e />。
8. 接著會顯示訊息，通知您下次啟動 Office 時就會套用您的設定。關閉並重新啟動 Word。 

現在您已經準備好在 Word 中執行。 

1. 開啟 Word 文件。 
2. 在 Word 2016 的 [插入]<e /> 索引標籤上，選擇 [我的增益集]<e />。 
3. 選取 [共用資料夾]<e /> 索引標籤。
4. 選擇 [愚蠢的故事增益集]<e />，然後選取 [插入]<e />。
5. 增益集會在一個工作窗格中載入。請參閱圖 1，了解其載入後的外觀。
6. 選取故事，即可將重複使用文字輸入 Word 文件中。

**圖 1. 載入至 Word 的愚蠢的故事增益集**

![已載入愚蠢的故事增益集的 Word 應用程式的圖片](../readme-images/sillystoriesUI.PNG)

## 問題和建議

我們很樂於收到您對於愚蠢的故事 Word 增益集範例的意見反應。您可以在此儲存機制的[問題](https://github.com/OfficeDev/Word-Add-in-SIllyStories/issues)區段中，將您的問題及建議傳送給我們。

請在 [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API) 提出有關增益集開發的一般問題。務必以 [office-js]、[word] 和 [API] 標記您的問題或意見。

## 深入了解

這裡有更多的資源，可協助您建立以 Word Javascript API 為基礎的增益集︰

* [在 iPad 和 Mac 上側載 Office 增益集](http://dev.office.com/docs/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac)
* [Office 增益集平台概觀](https://msdn.microsoft.com/zh-tw/library/office/jj220082.aspx)
* [Word 增益集](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins.md)
* [Word 增益集程式設計概觀](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins-programming-guide.md)
* [Word 的程式碼片段總管](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
* [Word 增益集 JavaScript API 參考](https://github.com/OfficeDev/office-js-docs/tree/master/word/word-add-ins-javascript-reference)

## 著作權
Copyright (c) 2015 Microsoft.著作權所有，並保留一切權利。
