/* Project Logクラス(Logクラスの子クラス) */
ProjectLog = function(){
  Log.apply(this);
}

inherit = (function(){
  /* 空のプロキシオブジェクトを作成 */
  var F = function(){};
  return function(C, P){
    F.prototype = P.prototype;
    C.prototype = new F();
    
    /* スーパークラスを格納 */
    C.uber = P.prototyoe;
    
    /* コンストラクタを指定 */
    C.prototype.constructor = C;
  }
}());
inherit(ProjectLog, Log);

/* Google Apps Scriptのエラーをログ */
Log.prototype.scriptErrorLog = function(error){
  this.fatal("[Google Apps Scriptのエラー] ファイル名：" + error.fileName + ", 行：" + error.lineNumber + ", メッセージ："+ error.message);
}

/* test */
function PtojectLogTest(){
  var log = new ProjectLog();
  try{
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("log2");
    sh.getRange("A1:C1").setValues([["Timestamp", "Level", "Message"]]);
  } catch(error){
    log.scriptErrorLog(error);
  }
}


