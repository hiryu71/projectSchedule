/* ProjectRecordクラス */
ProjectRecord = function() {
  var sheetName = "シート1";
//  var sheetName = "PJ情報";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  this.sheet = ss.getSheetByName(sheetName);
  if(this.sheet == null){
    this.status = Status.fatal;
  }
  else{
    this.row = this.sheet.getLastRow() + 1;
    this.status = Status.running;
  }
}

/* Status */
Status = {
  notNewPJ: 0,  /* PJ発行なし */
  running: 1,  /* 実行中 */
  fatal: 2,  /* 致命的なエラー */
  error: 3,  /* エラー */
  done: 4,  /* 終了 */
  _sizeof: 5
}

/* 実行 */
ProjectRecord.prototype.running = function(){
  var gmail = new Gmail();
  var project = new Project();
  
  var msgs = gmail.getInboxMail(/^【PJ発行】/);
  if(0 == msgs){  
    this.status = Status.notNewPJ;
  }
  else{
    for(var m = msgs.length - 1; 0 <= m ; m--){
      var msg = msgs[m];
      var body = msg.getPlainBody();
      var strings = body.split("\n");
      var Items = project.getItems(strings);
      this.checkPJNumber(Items[PJInfo.ProjectNumber]);
      if(Status.notNewPJ != this.status){
        this.record(Items);
      }
    }
  }
  
}

/* PJ番号確認*/
ProjectRecord.prototype.checkPJNumber = function(ProjectNumber){
  for(var row = 1; row < this.row; row++){
    var OldProjectNumber = this.sheet.getRange(row, PJInfo.ProjectNumber + 1).getValue();
    
    /* PJ番号重複確認 */
    if(ProjectNumber == OldProjectNumber){
      this.status = Status.notNewPJ;
    }
    
    /* PJ番号抜け確認 */
    
    if(ProjectNumber == OldProjectNumber){
    
    }
  }
}
  
/* シートへ記録 */
ProjectRecord.prototype.record = function(Items){
//  var sheet = SpreadsheetApp.getActive().getSheetByName(this.sheetName);
  for(var m = 0; m < Items.length; m++){
//    sheet.getRange(this.row, m + 1).setValue(Items[m]);
    this.sheet.getRange(this.row, m + 1).setValue(Items[m]);
  }
  this.row++;
}

/* test */
function ProjectRecordTest(){
  var projectRecord = new ProjectRecord();
  projectRecord.running();
}

function testWrite(word){
  var sheet = SpreadsheetApp.getActive().getSheetByName("シート1");
  sheet.getRange(this.row, 1).setValue(word);
}