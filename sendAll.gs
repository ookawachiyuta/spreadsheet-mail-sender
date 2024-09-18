function sendAllTest() {
  const COL_NO_TO        = 0;  // to(宛先)列
  const COL_TITLE        = 1;  // title(メールのタイトル)列
  const COL_NO_BCC       = 2;  // bcc(宛先)列
  const COL_UNIRINK      = 3;  //　　ユニークリンク
  const COL_MEETRINK     = 4;  //　　　打ち合わせリンク
  const COL_OFFICE       = 5;  //　　事務所名
  const COL_LAWYER       = 6;  // 弁護士名
  const COL_TEMNUM       = 7;  //　　テンプレートナンバー
  const COL_NOTSEND      = 8;  //　　送信しなくていいメール
  const COL_RESULT       = 9;  // 送信結果
  const COL_TEXT         = 12; // 本文テンプレート

  var dtLimit = new Date();
  // 開始行数
  const START_ROWS = 2;  // 2行目から
  // 開始列数
  const START_COLS = 1;  // A列から
 
  // シートの取得
  var activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadSheet.getSheetByName('シート1');
  var sheetData = sheet.getSheetValues(START_ROWS, START_COLS, sheet.getLastRow(), sheet.getLastColumn());

  var ui = SpreadsheetApp.getUi();
  var count = 1;
  var response = ui.alert('メールを送信します。', '続行しますか？', ui.ButtonSet.YES_NO);
  var isNull = null;
  console.log("発火前")
  if (response == ui.Button.YES) {
    //設定項目確認
    console.log("確認前")
    sheetData.forEach(function(value, index) {
      count++
      if(value[COL_NOTSEND] == "送信する"){
        if (value[COL_TEMNUM] == ""){
          ui.alert(count+"行目の設定でテンプレート番号が指定されていません！指定してからもう一度実行してください。")
        }else if(value[COL_NO_TO] == "" || value[COL_TITLE] == "" || value[COL_UNIRINK] == "" || value[COL_MEETRINK] == "" || value[COL_OFFICE] == "" || value[COL_LAWYER] == ""){
          isNull = ui.alert(count+'行目の設定に空の項目があります。', '続行しますか？', ui.ButtonSet.YES_NO);
        }
      }
    })
    //メール送信
    console.log("送信前")
    console.log(isNull)
    sheetData.forEach(function(value, index) {
      if(value[COL_NOTSEND] == "送信する" && value[COL_NO_TO] != "" && value[COL_TEMNUM] != "" && isNull != ui.Button.NO){
        let temText = sheetData[Number(value[COL_TEMNUM] -2)][COL_TEXT]
        let sendText = temText.replace("**会社名**",value[COL_OFFICE].trim()).replace("**部署名**",value[COL_LAWYER].trim()).replace("**ユニークリンク**",value[COL_UNIRINK].trim()).replace("**打ち合わせリンク**",value[COL_MEETRINK].trim())
        try{
          console.log(value[COL_NO_BCC])
          MailApp.sendEmail(
            value[COL_NO_TO], 
            value[COL_TITLE], 
            sendText,
            {"bcc": value[COL_NO_BCC]}
            // {bcc:"spyspysee55@gmail.com"}
          );
            // 送信成功の文字列を挿入
          sheet.getRange(START_ROWS + index, COL_RESULT + 1).setValue("送信完了しました\n送信日：" + dtLimit);
          console.log("送信完了")
        } catch(e){
          // 送信失敗の文字列を挿入
          console.log(e);
          sheet.getRange(START_ROWS + index, COL_RESULT + 1).setValue("送信失敗しました\n送信日：" + dtLimit);
          console.log("送信失敗")
        }
      }else if(value[COL_NOTSEND] == "送信しない" && value[COL_NO_TO] != ""){
        sheet.getRange(START_ROWS + index, COL_RESULT + 1).setValue("送信しませんでした");
        console.log("送信せず")
      }
    });
  } else {
    ui.alert('中止します');
    console.log("中止")
  }
  console.log("end")
}
