function onFormSubmit(event) {
  var number = searchUser(event.namedValues['Eメールアドレス']);
  if(number >= 0) {
    // 見つかった
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("フォームによるレスポンス");
    
    var lastrow = sheet.getLastRow();
    Logger.log(lastrow);
    Logger.log(2 + number);
    if(lastrow != (2 + number)) {
      // 以前登録されたものであれば削除
      sheet.deleteRow(2 + number);
    }
    
  } else {
    // 見つからなかった。
    Logger.log("対象が見つからない");
  }
  
  // 登録
  Logger.log("メール送信");
  sendMailSuccess(event.namedValues['お名前（漢字）'], event.namedValues['ご所属'], event.namedValues['Eメールアドレス'], event.namedValues['懇親会への参加']);
  
}

function searchUser(email) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // メールアドレスの収集
  var sheet = spreadsheet.getSheetByName("フォームによるレスポンス");
  var lastrow = sheet.getLastRow();
  
  if(lastrow == 1) {
    // 最後の行 = 1の場合、ラベルのみ。メール送信できないので終了。
    Logger.log("申し込み者 = 0");
    return;
  }
  
  var values = sheet.getRange(2, 5, lastrow).getValues();        // 1行目はラベルなので2行目から取得する。

  for(var i = 0; i < values.length; i++) {
    if(values[i][0] == email) {
      // 対象のユーザ      
      return i;
    }
  }
  
  return -1;
}

function sendMailSuccess(name, assign, address, party) {
  var body = assign + " " + name + "様" + "\n\n";
  body += "【イベント名】への参加申し込みありがとうございました。\n";
  body += "【イベント名】の開催日は2013年5月25日 13:00からです。\n";
  body += "何かありましたら、主催者までご連絡下さい。\n\n";
  body += "申し込み内容\n";
  body += "ご所属： " + assign + "\n";
  body += "お名前： " + name + "\n";
  body += "メールアドレス: " + address + "\n";
  body += "懇親会: " + party + "\n";
  
  Logger.log("%s %s %s", name, address, party);
  MailApp.sendEmail(address, "申し込みありがとうございました。", body);
}

function remainder() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // メールアドレスの収集
  var sheet = spreadsheet.getSheetByName("フォームによるレスポンス");
  var lastrow = sheet.getLastRow();
  
  if(lastrow == 1) {
    // 最後の行 = 1の場合、ラベルのみ。メール送信できないので終了。
    Logger.log("申し込み者 = 0");
    return;
  }
  
  var values = sheet.getRange(2, 3, lastrow).getValues();        // 1行目はラベルなので2行目から取得する。
  
  var recipient = "【メールアドレス】";                       // 自分（主催者）のメールアドレス（要編集）
  var bcc = "";
  var row = 0;
  while(values[row][0] != "") {
    bcc += values[row][0] + ",";
    row++;
  }
  
  var subject = "【イベント名】";
  var body = "参加者の皆さん\n\n";
  body += "イベント名の開催は、明日の13:00からですが\n";
  body += "午前中から会場を空ける予定です。午前中からワイワイやりましょう！\n";
  body += "明日は、【会場名】でお待ちしております。\n";
  
  MailApp.sendEmail(recipient, subject, body, {bcc:bcc});
}