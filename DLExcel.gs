//setProperty　読み込むシートのURLをsetして１度だけ実行
//"https://docs.google.com/spreadsheets/d/****************/edit#gid=********";
function setUrlProperty(){
  const myURL = "https://docs.google.com/spreadsheets/d/****************/edit#gid=********";
  PropertiesService.getScriptProperties().setProperty("url",myURL);
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

class Ui{
  constructor() {
    this.ui = new Date();
  }
}


function getExcelSheet(){
  //const Time = new Timer();
  const ui = SpreadsheetApp.getUi();
  const Source = new TargetSheet();
  //Logger.log("tryGetSheet:" + Time.elapsedTime());
  if(!Source.tryGetSheet()){
    ui.alert("シートが取得できませんでした");
    return;
  }
  //Logger.log("SheetGet:" + Time.elapsedTime());
  const sheetName = Source.sheet.getSheetName();
  const result = ui.alert("「" + sheetName + "」をExcelに変換します",ui.ButtonSet.OK_CANCEL);
  if(result == "CANCEL")return;
 
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
  //Logger.log("showDialog:" + Time.elapsedTime());
  showDialog(thisBook,destinationSh,sheetName);
  //Logger.log("end:" + Time.elapsedTime());
}


//エクセルでダウンロードダイアログ
function showDialog(thisBook,destinationSh,sheetName,tempSheetName) {
  // 選択中のスプレッドシート
  //Logger.log("showDialog");
  const spreadSheetId = thisBook.getId();
  const sheetId = destinationSh.getSheetId();
  const dlFileName =thisBook.getName() +  ".xlsx";
  const url = 'https://docs.google.com/spreadsheets/d/' + spreadSheetId + '/export?format=xlsx&gid=' + sheetId;
  const jsCode = 'google.script.host.close();';
  const html = HtmlService.createHtmlOutput()
    .append('<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.5.3/dist/css/bootstrap.min.css" integrity="sha384-TX8t27EcRE3e/ihU7zmQxVncDAy5uIKz4rEkgIXeMed4M0jlfIDPvg6uqKI2xXr2" crossorigin="anonymous">')
    .append('<br><div class="card"><div class="card-header"><h5><p>ダウンロードリンク</p></h5></div>')
    .append('<br><div class="card-body"><h4><p style="text-align:center;"><a href="' + url +'">'+ dlFileName + '<a/></p></h4></div><br></div>')
    .append('<br><div style="text-align:right;"><input type="button"class="btn btn-primary btn-block" value="閉じる" onclick="'+ jsCode +'" /></div>')
    .setWidth(600).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'シート「'+sheetName+'」のダウンロード');
}
