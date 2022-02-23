# GASExcelDownloader
GAS Download Spreadsheet as Excel format(Only Value)  
スプレッドシート上のボタン等から実行、指定したシートを値のみの*.xlsxでダウンロード  

## Setting

### 1 :First, set myURL to PropertiesService  
  取り込みたいシートのURLをset
  function setUrlProperty()内myURLの値を取り込みたいシートURLに変更()  
  
  ####    const myURL = "https://docs.google.com/spreadsheets/d/****************/edit#gid=********"; <-change here  
  
    function setUrlProperty(){  
        const myURL = "https://docs.google.com/spreadsheets/d/****************/edit#gid=********";  
        PropertiesService.getScriptProperties().setProperty("url",myURL);  
        Logger.log(PropertiesService.getScriptProperties().getProperty("url"));  
    }

### 2 :run setUrlProperty()  
  setUrlProperty()を実行しPropertiesServiceに書き込む

## Usage
  run getExcelSheet()  
  スプレッドシートの画像ボタンにgetExcelSheetを設定し実行  
