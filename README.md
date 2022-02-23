# GASExcelDownloader
GAS Download Spreadsheet as Excel format(Only Value)  
スプレッドシート上のボタン等から実行、指定したシートを値のみの*.xlsxでダウンロード  

# Setting

1 :First, set myURL to PropertiesService  
//myURLの値を取り込みたいシートURLに変更  
    function setUrlProperty(){  
      const myURL = "https://docs.google.com/spreadsheets/d/****************/edit#gid=********";  
      PropertiesService.getScriptProperties().setProperty("url",myURL);  
      Logger.log(PropertiesService.getScriptProperties().getProperty("url"));  
    }

2 :run setUrlProperty()


# Usage
  run getExcelSheet()  
  スプレッドシートの画像ボタンにgetExcelSheetを設定し実行  
