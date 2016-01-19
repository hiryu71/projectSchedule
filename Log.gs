/* Logクラス */
Log = function(){
  this.logSheet("Log");
}

/* Log設計：http://qiita.com/nanasess/items/350e59b29cceb2f122b3 */

/* Logシート */
Log.prototype.logSheet = function(sheetName){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  this.sheet = ss.getSheetByName(sheetName);
  /* Logシートが存在しない場合は作成 */
  if(this.sheet == null){
    var activeSheet = ss.getActiveSheet();
    var sheetNumber = ss.getSheets().length;
    this.sheet = ss.insertSheet(sheetName, sheetNumber);
    this.sheet.getRange("A1:C1").setValues([["Timestamp", "Level", "Message"]]);
    this.sheet.getRange("A2:C2").setValues([[new Date(), "INFO", sheetName + " has been created"]]);
    ss.setActiveSheet(activeSheet);
  }
}

/* Logging */
Log.prototype.logging = function(level, message){
  var sheet = this.sheet;
  var date = new Date();
  var lastRow = sheet.getLastRow();
  sheet.insertRowAfter(lastRow).getRange(lastRow + 1, 1, 1, 3).setValues([[date, level, message]]);
}

/* fatal log */
/* 致命的なエラー：プログラムの異常終了を伴うようなエラー */
Log.prototype.fatal = function(message){
  this.logging("FATAL", message);
}

/* error log */
/* エラー：予期しないその他の実行時エラー */
Log.prototype.error = function(message){
  this.logging("ERROR", message);
}

/* warn log */
/* 警告：エラーに近い事象など。異常とは言い切れないが正常とも異なる何等かの予期しない問題 */
Log.prototype.warn = function(message){
  this.logging("WARN", message);
}

/* info log */
/* 情報：実行時の何らかの注目すべき事象。メッセージ内容は簡潔に止めるべき */
Log.prototype.info = function(message){
  this.logging("INFO", message);
}

/* degug log */
/* デバッグ情報：システムの動作状況に関する詳細な情報 */
Log.prototype.debug = function(message){
  this.logging("DEBUG", message);
}

/* test Log */
function LogTest(){
  var log = new Log();
  try{
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("log2");
    sh.getRange("A1:C1").setValues([["Timestamp", "Level", "Message"]]);
  } catch(error){
    log.scriptErrorLog(error);
  }
}