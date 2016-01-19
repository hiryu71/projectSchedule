/* Gmail�N���X */
Gmail = function() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("config");
  for (var Cnt = 1; Cnt <= sheet.getLastRow(); Cnt++){
    if(sheet.getRange(Cnt, 1).getValue() == "mail address"){ /* A��ɍ��ڂ����邱�Ƃ��O�� */
      this.mailAddress = sheet.getRange(Cnt, 2).getValue();
    }
  }
}

/* ��M�g���C����w�肵�������̃��[����擾���� */
Gmail.prototype.getInboxMail = function(subjectWord){
  var result = new Array();
  var thds = GmailApp.getInboxThreads();
  var cnt = 0;
  for(var n in thds){
    var thd = thds[n];
    var msgs = thd.getMessages();
    for(var m in msgs){
      var msg = msgs[m];
      var subject = msg.getSubject();
      if(0 == subject.search(subjectWord)){
        result[cnt] = msg;
        cnt++;

        /* ラベルを適用 */
        //this.setLabel(thd, "PJ発行");
          
        /* メールを既読に変更 */
        //msg.markRead();
          
        /* メールをアーカイブへ移動 */
        //GmailApp.moveThreadToArchive(thd);            
      }
    }
  }
  return result;
}

/* ラベルを適用 */
Gmail.prototype.setLabel = function(thread, labelName){
  labelName = String(labelName);
  var label = GmailApp.getUserLabelByName(labelName);
  if(label == null){
    label = GmailApp.createLabel(labelName);
    var logMassage = "creage " + labelName + " label";
    Logger.log(logMassage);
  }
  label.addToThread(thread);
}

/* エラー発生メールを送信する */
Gmail.prototype.sendErrorMessage = function(errorMessage){
  var scriptAddress = "https://script.google.com/macros/d/MtgVi2tDku2DFFOZXjFL5nGY3YA6J_whq/edit?uiv=2&mid=ACjPJvGgQrNrAQhF6Ql-U6xjTq2Rvx-LnkeWU_rvMKbi45-QtDabOAQO-l6Us7ULS2PwiauBOfQJYAH1F_Y-duP93kNSGjvVvamYcV9SGsBjlcIPrJtZ4X6idzX-dhCd78NOdXsUlxE_F7c";
  var errorMessage = "「PJ別スケジュール」で”" + String(errorMessage) + "”エラー発生 \n" + scriptAddress;
  MailApp.sendEmail(String(this.mailAddress), "【エラー】「PJ別スケジュール」でエラー発生", errorMessage); 
}

/* test */
function GmailTest(){
  var gmail = new Gmail();
  gmail.sendErrorMessage("test");

  //Browser.msgBox(this.mailAddress);
}
