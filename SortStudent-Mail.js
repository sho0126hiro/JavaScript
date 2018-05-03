//参考：https://qiita.com/matsuhandy/items/ddb7ed092b29c5b8b3c5

var logID    = "**************";//log2ID
var ssID1    = "**************";//フォーム入力で作成されたSpreadsheetのID
var ssID3    = "**************"//database ssID

function submitForm(){//フォームが送信されたら呼び出される関数，重複処理を避ける
  log = new Doc(logID);
  log.clear();
  log.print('\n'+getDateAndTime(0)+" スクリプト開始\n");
  var lock = LockService.getScriptLock();//ロックサービスのオブジェクトを生成
  try{
    lock.waitLock(30000);//複数のフォーム送信がほぼ同時にあった時，遅い方に最大30秒待ってもらう
    log.print("他のスクリプト実行要求をロック完了，最大60秒\n");
    main(log);
  }catch(err){
    log.print("発生したエラー："+err+'\n');
  }finally{
    lock.releaseLock();//次の送信のためにロック解除
    log.print(getDateAndTime(0)+" ロック解除，次のスクリプト要求を受け付け開始\n");
  } 
}

function main(log){
  var ssSrc = new Ssheet(ssID1);//フォーム入力されたデータが入っているSpreadsheet
  var todayNow =getDateAndTime(0);//現在日時の取得

  //getLastRow(sheet); | getLastColumn(0);
  var maxRow = ssSrc.getLastRow(0);
  var maxColumn = ssSrc.getLastColumn(0);
  log.print("最後の行"+maxRow+" 最後の列"+maxColumn+"\n");

  for(var i=1;i<=maxColumn;i++){
    /*index：タイムスタンプ・題目*/
    //getValue : spreadsheetのsheetからrow行のcal列のデータをもらってくるメソッド:getValue(sheet,row,col) 
    var index = ssSrc.getValue(0, 1, i); //index : 0sheet 1行 i列   
    //data :　一番下の行のi列目のデータをコピー
    var data = ssSrc.getValue(0, maxRow, i);
    log.print("i = " + i + " index : " + index + "\n");
    log.print("data="+data+"\n");
    if(index == "回答者 or 未回答者を選択して下さい。"){
      var studentsdata=data.split(', '); //spaceがないと1番以外動作しない
    }
    if(index == "回答者か未回答者かを選択してください"){
      var checker = data; //data : "回答者" "未回答者"
    }
  }
  // database
  log.print("database\n");
  var ssData = new Ssheet(ssID3);//students database Spreadsheet ID
  var MaxRow_ssData = ssData.getLastRow(0);
  var maxColumn_ssData = ssData.getLastColumn(0);
  //get data
  var studentsdata_2 = new Array();
  var slackID = new Array();
  for(var j=1;j<=studentsdata.length;j++)log.print(studentsdata[j-1]);
  log.print("\n");
  //Slack ID data >> row1:index row2~:studentsdata//
  for(var i=2;i<=MaxRow_ssData;i++){
    studentsdata_2[i-2] = ssData.getValue(0, i, 1);
    slackID[i-2]   = "@"+ssData.getValue(0, i, 2) + ssData.getValue(0, i, 3);
    //log.print("students data_2: "+studentsdata_2[i-2]+ "  slackID  "+slackID[i-2]+"\n");
  }
  //Comparioson data
  var stChecker = new Array();
  for(var j=1;j<=studentsdata.length;j++){
    for(var i=2;i<MaxRow_ssData;i++){
      if(studentsdata[j-1]==studentsdata_2[i-2]){//sheets input data == database
        stChecker[i-2]=1;
        j++;
      }else{
        stChecker[i-2]=0;
      }
      if(stChecker[i-2]==1)log.print("check:"+ stChecker[i-2] +" student : "+studentsdata_2[i-2]+" slackID : "+slackID[i-2]+"\n");
    }
  }
  //mail
  toAdr = "**************" //送信先メールアドレス
  var finished = new Array(); //回答者格納配列
  var still    = new Array(); //未回答者格納配列
  if(checker=="回答者"){
    for(var i=0;i<stChecker.length;i++){
      if(stChecker[i]) finished.push(slackID[i]);
      else             still   .push(slackID[i]);
    }
    /*
    log.print("finish\n");
    for(var i=0;i<finished.length;i++)log.print(finished[i]+"\n");
    log.print("still\n");
    for(var i=0;i<still.length;i++)log.print(still[i]+"\n");
    */
  }else{
    for(var i=0;i<stChecker.length;i++){
      if(stChecker[i])still   .push (slackID[i]);
      else            finished.push (slackID[i]);
    }
    /*
    log.print("finish\n");
    for(var i=0;i<finished.length;i++)log.print(finished[i]+"\n");
    log.print("still\n");
    for(var i=0;i<still.length;i++)log.print(still[i]+"\n");
    */
  }
  var finishedbody=finished.join("\n")
  var stillbody=still.join("\n")
  //send mail
  MailApp.sendEmail({// メール送信
    to: toAdr,
    subject: "[mail test]",
    name: "回答者/未回答者 SlackID自動送信サービス",
    body: "送信日時："+getDateAndTime(0)+"\n"
    +"回答者のSlackIDは以下の通りです。\n"+finishedbody+"\n未回答者のSlackIDは以下の通りです。\n"+stillbody
  });
  
}

//-----------------------日付や時間のための関数群
function getSerial(date,time){//日付と時間からシリアル値をゲット
  var serial = new Date(date.toString().slice(0,16)+time.toString().slice(16));
  return serial;
}

function getDateAndTime(data){// 引数がゼロなら現在日時/そうでなければ指定日時をyyyymmdd_hhmmssで返す
  if(data==0) var now = new Date();
  else        var now = new Date(data);
  var year = now.getYear();
  var month = now.getMonth() + 1;
  var day = now.getDate();
  var hour = now.getHours();
  var min = now.getMinutes();
  var sec = now.getSeconds();
  return (""+year).slice(0,4) + ("0"+month).slice(-2) + ("0"+day).slice(-2) +'_'+ 
    ("0"+hour).slice(-2) + ("0"+min).slice(-2) + ("0"+sec).slice(-2);
}

//-----------------------ファイルの新規作成等の関数群
//フォルダID，ファイルID，ファイル名を受け取り，ファイルIDのコピーをフォルダ内に作成して，ファイルのIDを返す
function copyFileInFolder(folderID, srcID, fileName) {
  var originalFile = DriveApp.getFileById(srcID);
  var folder = DriveApp.getFolderById(folderID);
  var copiedFile = originalFile.makeCopy(fileName, folder);
  copiedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);//リンクからアクセスできる人は編集可能にする
  var copiedFileId = copiedFile.getId();//コピーのファイルIDをゲット
  return copiedFileId;
}
//フォルダID，ファイル名を受け取り，スプレッドシートを指定フォルダ内に新規作成しそのファイルIDを返す
function createSpreadsheetInFolder(folderID, fileName) {
  var folder = DriveApp.getFolderById(folderID);
  var newSS=SpreadsheetApp.create(fileName);
  var originalFile=DriveApp.getFileById(newSS.getId());
  var copiedFile = originalFile.makeCopy(fileName, folder);
  DriveApp.getRootFolder().removeFile(originalFile);
  var copiedFileId = copiedFile.getId();//コピーのファイルIDをゲット
  return copiedFileId;
}

//////////Ssheetクラスの定義開始（コンストラクタとメンバ関数で構成）
//Ssheetクラスのコンストラクタの記述
Ssheet = function(id){
  this.ssFile = SpreadsheetApp.openById(id);
  this.ssFileName = this.ssFile.getName();
  SpreadsheetApp.setActiveSpreadsheet(this.ssFile);//値を返さない
  this.activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
}
//Ssheetクラスのメンバ関数（メソッド）の定義開始
//spreadsheetのファイル名を返すメソッド
Ssheet.prototype.getFileName = function(){
  return this.ssFileName;
}
//spreadsheetのファイル名を変更するメソッド
Ssheet.prototype.rename = function(newName){
  this.ssFile.rename(newName);
}
//spreadsheetのsheetでrow行cal列にデータを入れるメソッド
Ssheet.prototype.setValue = function(sheet,row,col,value){
  this.activeSheet = this.activeSpreadsheet.getSheets()[sheet];
  var cell = this.activeSheet.getRange(row,col);
  cell.setValue(value);
}
//spreadsheetのsheetでrow行cal列をクリアするメソッド
Ssheet.prototype.clear = function(sheet,row,col,value){
  this.activeSheet = this.activeSpreadsheet.getSheets()[sheet];
  var cell = this.activeSheet.getRange(row,col);
  cell.clear(value);
}
//spreadsheetのsheetからrow行のcol列のデータをもらってくるメソッド
Ssheet.prototype.getValue = function(sheet,row,col) {
  this.activeSheet = this.activeSpreadsheet.getSheets()[sheet];
  var value = this.activeSheet.getRange(row, col).getValue();
  return value;
}
//背景の色を設定するメソッド
Ssheet.prototype.setBackgroundColor = function(sheet,row,col, r,g,b) {
  this.activeSheet = this.activeSpreadsheet.getSheets()[sheet];
  var cell = this.activeSheet.getRange(row,col);
  cell.setBackgroundRGB(r,g,b);
}
//spreadsheetのsheet数を指定の数増やすメソッド
Ssheet.prototype.insertSheet = function(num){
  var sheetNum = this.activeSpreadsheet.getNumSheets();
  while(num>sheetNum){
    this.activeSpreadsheet.insertSheet();
    sheetNum++;
  }
}
//spreadsheetの指定sheetを削除
Ssheet.prototype.deleteSheet = function(sheet){
  this.activeSpreadsheet.deleteSheet(this.activeSpreadsheet.getSheets()[sheet])
}
//spreadsheetのsheetの名前をセットするメソッド
Ssheet.prototype.renameSheet = function(sheet,newName){
  this.activeSheet = this.activeSpreadsheet.getSheets()[sheet];
  this.activeSheet.setName(newName);
}
//spreadsheetの指定sheetの全データを取得
Ssheet.prototype.getValues = function(sheet){
  this.activeSheet = this.activeSpreadsheet.getSheets()[sheet];
  return this.activeSheet.getDataRange().getValues();//シートの全データを取得
}  
//spreadsheetの指定sheetを取得
Ssheet.prototype.getSheet = function(sheet){
  this.activeSheet = this.activeSpreadsheet.getSheets()[sheet];
  return this.activeSheet;//シートを返す
}
//spreadsheetの指定sheetの最後の行番号を取得
Ssheet.prototype.getLastRow = function(sheet){
  this.activeSheet = this.activeSpreadsheet.getSheets()[sheet];
  return this.activeSheet.getLastRow();//最後の行番号を取得
}
//spreadsheetの指定sheetの最後の列番号を取得
Ssheet.prototype.getLastColumn = function(sheet){
  this.activeSheet = this.activeSpreadsheet.getSheets()[sheet];
  return this.activeSheet.getLastColumn();//最後の列番号を取得
}
//spreadsheetの指定行rowを削除
Ssheet.prototype.deleteRow = function(sheet,row){
  this.activeSheet = this.activeSpreadsheet.getSheets()[sheet];
  this.activeSheet.deleteRow(row);//行を削除
}
//spreadsheetの指定行rowを挿入
Ssheet.prototype.insertRow = function(sheet,row){
  this.activeSheet = this.activeSpreadsheet.getSheets()[sheet];
  this.activeSheet.insertRows(row);//行を挿入
}
//spreadsheetの指定列colを削除
Ssheet.prototype.deleteColumn = function(sheet,col){
  this.activeSheet = this.activeSpreadsheet.getSheets()[sheet];
  this.activeSheet.deleteColumn(col);//列を削除
}
//spreadsheetの指定列colを挿入
Ssheet.prototype.insertColumn = function(sheet,col){
  this.activeSheet = this.activeSpreadsheet.getSheets()[sheet];
  this.activeSheet.insertColumns(col);//列を挿入
}
//////////Ssheetクラスの定義終了

//////////Docクラスの定義開始（コンストラクタとメンバ関数で構成）
//Docクラスのコンストラクタの記述
Doc = function(id){
  this.ID = id;
  this.doc = DocumentApp.openById(this.ID); 
  this.body = this.doc.getBody();
  this.docText = this.body.editAsText();
}
//Docクラスのメンバ関数の定義開始
//メソッドprintの定義，テキスト追加
Doc.prototype.print = function(str){
  this.docText.appendText(str);
}
//メソッドreplaceの定義，文字列置き換え
Doc.prototype.replace = function(src,dst){
  this.body.replaceText(src,dst);
}
//メソッドclearの定義，全消去
Doc.prototype.clear = function(){
  this.body.clear();
}
//メソッドgetIDの定義，ファイルIDを返す
Doc.prototype.getID = function(){
  return this.ID;
}
//指定秒数のウェイト，表示動作を遅らせたい時などに使用
Doc.prototype.waitSec = function(sec){
  var start = new Date().getSeconds();
  while((new Date().getSeconds()-start) < sec);
}
//指定ミリ秒のウェイト，表示動作を遅らせたい時などに使用
Doc.prototype.waitMiliSec = function(msec){
  var start = new Date(); //new Date()は，「1970年1月1日午前0時」からの通算ミリ秒を返す
  while((new Date()-start) < msec);
}
//今現在の日時を表示
Doc.prototype.printTodayNow = function(){
  var now = new Date();
  var year = now.getYear();
  var month = now.getMonth() + 1;
  var day = now.getDate();
  var hour = now.getHours();
  var min = now.getMinutes();
  var sec = now.getSeconds();
  this.docText.appendText(year +'_'+ ("0"+month).slice(-2) +'_'+ ("0"+day).slice(-2) +' '+ 
    ("0"+hour).slice(-2) +'-'+ ("0"+min).slice(-2) +'-'+ ("0"+sec).slice(-2));
}
/////////Docクラスの定義終了
