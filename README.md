# GASExcelDownloader
GAS Download Spreadsheet as Excel format(Only Value)  
指定したスプレッドシートを（セルの数式を値に置き換え）*.xlsx形式でダウンロード  

スプレッドシートのメニューからダウンロードした場合に、数式があると正しく取得できない為に作成
値への置換えは一時シートを作成して行うため、現本には影響を与えない。

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
  エディターからsetUrlProperty()を実行しPropertiesServiceに書き込む

## Usage
  run getExcelSheet()  
  スプレッドシートの画像ボタンにgetExcelSheetを設定し実行  
