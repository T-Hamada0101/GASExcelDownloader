# GASExcelDownloader
GAS Download Spreadsheet as Excel format(Only Value)
スプレッドシートのボタン等から実行し指定したシートを値のみの*.xlsxでダウンロード


1 :First, set URL to PropertiesService
//myURLの値を取り込みたいシートURLに変更
function setUrlProperty(){
  const myURL = "https://docs.google.com/spreadsheets/d/****************/edit#gid=********";
  PropertiesService.getScriptProperties().setProperty("url",myURL);
  Logger.log(PropertiesService.getScriptProperties().getProperty("url"));
}

2 :run setUrlProperty()

3 :execution 
  run getExcelSheet()
