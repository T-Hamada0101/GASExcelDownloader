//setProperty　読み込むシートのURLをsetして１度だけ実行
//"https://docs.google.com/spreadsheets/d/****************/edit#gid=********";
function setUrlProperty(){
  PropertiesService.getScriptProperties().setProperty("url","https:***");
  Logger.log(PropertiesService.getScriptProperties().getProperty("url"));
}

class TargetSheet {
  constructor() {
    this.url = PropertiesService.getScriptProperties().getProperty("url");
    this.book = null;
    this.sheet = null;
    this.sheetName = "";
    this.ui = SpreadsheetApp.getUi();
  }
  //functions
  getBook(){
    try{
      if(this.book == null)this.book = SpreadsheetApp.openByUrl(this.url);
    }catch{
      //this.ui.alert("URLのスプレッドシートが取得できませんでした");
    }
    return this.book;
  }
  tryGetSheet() {
    if(this.sheet != null){return true;}
    if(this.book == null)this.getBook();
    try{
        //const sheets  = this.book.getSheets();
        const sheetId = Number(this.url.split('#gid=')[1]);
        for (const sheet of this.book.getSheets()) {
          if (sheetId === sheet.getSheetId()) {
            this.sheet = sheet;
            this.sheetName = this.sheet.getSheetName();
            return true;
          }
        }
    }catch{}
    return false;
  }
}
class Timer {
  constructor() {
    this.startTime = new Date();
  }
  elapsedTime(){
    const now = new Date();
    //実行時間の計算
    return (now - this.startTime) / 1000;
  }
}

function getExcelSheet(){
  const Time = new Timer();
  const Source = new TargetSheet();
  Logger.log("tryGetSheet" + Time.elapsedTime());
  if(!Source.tryGetSheet()){
    Source.ui.alert("シートが取得できませんでした");
    return;
  }
  Logger.log("SheetGet" + Time.elapsedTime());
  const sheetName = Source.sheet.getSheetName();
  const result = Browser.msgBox("「" + sheetName + "」をExcelに変換します",Browser.Buttons.OK_CANCEL);
  if(result === "cancel"){
    return;
  }
 
  //数式のままでは正確にExcel化出来ないので一時シートに値のみを入れる
  const lastRow    = Source.sheet.getLastRow();
  const lastColumn = Source.sheet.getLastColumn();
  const copyValue = Source.sheet.getRange(1,1,lastRow,lastColumn).getValues(); 
  const thisBook = SpreadsheetApp.getActiveSpreadsheet();
  const destinationSh = Source.sheet.copyTo(thisBook);
  const tempSheetName = "_temp_"+ sheetName;
  try{
    destinationSh.setName("_temp_"+ sheetName);
  }catch{
    thisBook.deleteSheet(thisBook.getSheetByName("_temp_"+ sheetName));
    destinationSh.setName("_temp_"+ sheetName);
  }
  destinationSh.getRange(1,1,lastRow,lastColumn).setValues(copyValue);//値のみペースト
  Logger.log("showDialog" + Time.elapsedTime());
  showDialog(thisBook,destinationSh,sheetName);
  Logger.log("end" + Time.elapsedTime());
}


//エクセルでダウンロードダイアログ
function showDialog(thisBook,destinationSh,sheetName,tempSheetName) {
  // 選択中のスプレッドシート
  //Logger.log("showDialog");
  const spreadSheetId = thisBook.getId();
  const sheetId       = destinationSh.getSheetId();
  //const tempSheetName =destinationSh.getSheetName();
  const url = 'https://docs.google.com/spreadsheets/d/' + spreadSheetId + '/export?format=xlsx&gid=' + sheetId;
　//Browser.msgBox(tempSheetName);
  //button onclick JavaScript : tempSheet削除;ダイアログclose;
  //const jsCode0 = 'google.script.run.deleatSheet('+ tempSheetName +');google.script.host.close();';
  //const jsCode1 = 'google.script.run.deleatSheet('+ tempSheetName +');';
  const jsCode = 'google.script.host.close();';
  const html = HtmlService.createHtmlOutput('<p>リンクをクリックしてダウンロードしてください。</p>')
    .append('<p style="text-align:center;"><a href="' + url +'" style="text-align: center;">Excelファイル<a/></p>')
    .append('<div style="text-align:right;"><input type="button" value="閉じる" onclick="'+ jsCode +'" /></div>')
    .setWidth(500).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, '「'+sheetName+'」ダウンロード');
}

function deleatSheet(tempSheetName){
  Logger.log()
  const thisBook = SpreadsheetApp.getActiveSpreadsheet();
  const sheets  = thisBook.getSheets();
  let terget;
  for (const sheet of sheets) {
   if (tempSheetName === sheet.getSheetName()){
     terget = sheet;
     break;
   }
 }
  if(terget != undefined){
    thisBook.deleteSheet(terget);
  }
}
