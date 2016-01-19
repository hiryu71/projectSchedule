/* Projectクラス */
Project = function() {
  this.ProjectStartWords = /^【PJ発行】/g;
  this.ProjectNumberWords = /^【\d{6}-\d{2}】|^【\d{6}】/g;
  this.ProjectNameWords = /^・(ＰＪ名|プロジェクト名)(　*)( *)/g;
  this.CustomerWords = /^・お客様(　*)( *)/g;
  this.PLNameWords = /^・(ＰＬ名|PL)(　*)( *)/g;
  this.OrderVolumeWords = /^・受注額(　*)( *)/g;
  this.LeadTimeWords = /^・(納期|予定納期)(　*)( *)/g;
  this.ScheduledDateOfDR1Words = /^・(ＤＲ１予定日|DR1実施予定日)(　*)( *)/g;
}

/* PJ項目 */
PJInfo = {
  ProjectNumber: 0,
  ProjectName: 1,
  Customer: 2,
  PLName: 3,
  OrderVolume: 4,
  LeadTime: 5,
  ScheduledDateOfDR1: 6,
  _sizeof: 7
};

/* PJ発行を検索 */
//Project.prototype.searchProjectStart = function(strings){
//  strings = String(strings);
//  return strings.search(this.ProjectStartWords);//本番のコード
//  return strings.search(/^testCC/);//testコード
//}

/* 項目を取得 */
Project.prototype.getItems = function(strings){
  var lists = ["getProjectNumber", "getProjectName", "getCustomer", "getPLName", "getOrderVolume", "getLeadTime", "getScheduledDateOfDR1"];
  var items = new Array();
  for(var n = 0; n < lists.length; n++){
    var func = this[lists[n]];
    for(var m in strings){
      var string = strings[m];
      var result = func.call(this, string)
      if(null != result){
        items[n] = result;
      }
    }
  }
  return items;
}

/* PJ番号を取得 */
Project.prototype.getProjectNumber = function(strings){
  strings = String(strings);
  var result = String(strings.match(this.ProjectNumberWords));
  if(null != result){
    result = String(result.replace(/【/, "【'"));
    result = result.match(/'\d{6}-\d{2}|'\d{6}/g);
  }
  return result;
}

/* 項目を抽出 */
Project.prototype.extractItem = function(strings, condition){
  var result = null;
  strings = String(strings);
  if(0 == strings.search(condition)){
    var parts = strings.split("：");
    result = parts[1];
  }
  return result;
}

/* PJ名を取得 */
Project.prototype.getProjectName = function(strings){
  return this.extractItem(strings, this.ProjectNameWords);
}

/* お客様を取得 */
Project.prototype.getCustomer = function(strings){
  return this.extractItem(strings, this.CustomerWords);
}

/* PL名を取得 */
Project.prototype.getPLName = function(strings){
  return this.extractItem(strings, this.PLNameWords);
}

/* 受注額を取得 */
Project.prototype.getOrderVolume = function(strings){
  var result =  this.extractItem(strings, this.OrderVolumeWords);
  if(null != result){
    result = result.match(/(\d{1,3}(,\d{3})*)\b/g);
  }
  return result;
}

/* 納期を取得 */
Project.prototype.getLeadTime = function(strings){
  var result =  this.extractItem(strings, this.LeadTimeWords);
  if(null != result){
    result = result.match(/\d{4}\/\d{1,2}\/\d{1,2}/g);
  }
  return result;
}

/* DR1予定日を取得 */
Project.prototype.getScheduledDateOfDR1 = function(strings){
  var result =  this.extractItem(strings, this.ScheduledDateOfDR1Words);
  if(null != result){
    result = result.match(/\d{4}\/\d{1,2}\/\d{1,2}/g);
  }
  return result;
}

/* Project test */
function ProjectTest(){
  var strings = ["【151263】", "・プロジェクト名：AAAシステム", "・お客様：株式会社BBB", "・PL    ：CCC", "・受注額：\\1,000,000（税抜き）", "・予定納期：5V版*2式　2016/1/29", "・DR1実施予定日：2015/12/7"];
  
  var project = new Project();
  var Items = project.getItems(strings);
  for(var n = 0; n < strings.length; n++){
      var sheet = SpreadsheetApp.getActive().getSheetByName('シート1');
      var string = strings[n];
      var Item = Items[n];
      sheet.getRange(n + 1, 1).setValue(string);
      sheet.getRange(n + 1, 2).setValue(Item);
  }
}
