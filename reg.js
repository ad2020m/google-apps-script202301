// リソース追加 moment.js
// MHMchiX6c1bwSqGM1PZiW_PxhMjh3Sh48
// スプレッドシート作成し、公開。URL中にID表示されているため取得

// gasを共有　→　このリンクを知っている〜
// 公開　ウェブアプリケーションとして導入　https://note.com/nauly/n/ne16df02573ca

const url = "スクリプトの共有";
const shId = "スプレッドシート🆔";
const chkId = "管理者ページ用パスワード";

const m = new Moment.moment();
const minDate = m.format('YYYY/MM/DD');
const startDay = m.day();
const maxDate = m.add(1,'months').format('YYYY/MM/DD');
const thisMonth = new Moment.moment().startOf('month');
const nextMonth = new Moment.moment().add(1, 'months').startOf('month');

const days = ["日","月","火","水","木","金","土"];

const spreadsheet = SpreadsheetApp.openById(shId);

function doGet(e) {


  if (e.pathInfo === 'json') {
    let params = JSON.stringify(e);
    return HtmlService.createHtmlOutput(params);
  }

  // 予約日指定された場合の処理→情報入力画面へ
  if(e.parameter.date){
    const inputHtml = HtmlService.createTemplateFromFile('input');
    inputHtml.date = e.parameter.date;
    inputHtml.title = 'ボランティア予約-入力画面';
    inputHtml.setTitle = 'ボランティア予約-入力画面';
    inputHtml.url = url;
    return inputHtml.evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 管理画面から予約変更リクエスト
  if(e.parameter.yoyaku){
    const editHtml = HtmlService.createTemplateFromFile('edit');
    editHtml.yoyaku = e.parameter.yoyaku;
    editHtml.youbi = e.parameter.youbi;
    editHtml.zan = e.parameter.zan;
    editHtml.user = e.parameter.user;
    editHtml.title = 'ボランティア予約-変更画面';
    editHtml.url = url;
    return editHtml.evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const data1 = getYoyaku(thisMonth);
  const data2 = getYoyaku(nextMonth);

  const indexHtml = HtmlService.createTemplateFromFile('index');
  indexHtml.err = '';
  indexHtml.startDay = startDay;
  indexHtml.makeMonth = [data1, data2];
  indexHtml.title = 'ボランティア予約';
  indexHtml.url = url;
  return indexHtml.evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// 初期画面カレンダー表示
function getYoyaku(month){

  // シート名取得
  const shName = month.format('YYYY/MM/DD').replace('/', '').slice(0, 6);
  const dataSh = set_sheet(shName, month);

  // 登録済みデータをスプレッドシートより取得（当月は当日の日付以降）
  const chkToday = dataSh.getRange(1, 1, dataSh.getLastRow() -1, 1).getDisplayValues();
  const minD = chkToday.findIndex( d => d[0] == minDate) >= 0 ?  chkToday.findIndex( d => d[0] == minDate) : 1;
  const maxD = chkToday.findIndex( d => d[0] == maxDate) >= 0 ?  30 - chkToday.findIndex( d => d[0] == maxDate) : 0;
  const data = dataSh.getRange(minD + 1, 1, dataSh.getLastRow() - minD - maxD, 6).getDisplayValues();

  return data;
}

// 編集処理
function editYoyaku(yoyaku, zan, user){

  // シート名取得
  const name = yoyaku.replace('/', '').slice(0, 6);
  const dataSheet = set_sheet(name);

  // 登録済みデータをスプレッドシートより取得
  const data = dataSheet.getRange(1, 1, dataSheet.getLastRow(), 6).getDisplayValues();

  // スプレッドシートに反映
  const row = data.findIndex( d => d[0] == yoyaku);
  const rowData = data.find( d => d[0] == yoyaku);

  if(zan){
    dataSheet.getRange(row + 1,3).setValue(zan);
  }else if(user){
  console.log(rowData);
    const cell = rowData.findIndex( d => d == user);
    dataSheet.getRange(row + 1, cell + 1).deleteCells(SpreadsheetApp.Dimension.COLUMNS);
    const nowZan = dataSheet.getRange(row + 1,3).getDisplayValues();
    dataSheet.getRange(row + 1,3).setValue(parseInt(nowZan) + 1);

  }


}


// 予約入力
function doPost(e) {

  // 編集処理へ
  if(e.parameter.edit){
    editYoyaku(e.parameter.yoyaku, e.parameter.zan, e.parameter.user);

    const data1 = getYoyaku(thisMonth);
    const data2 = getYoyaku(nextMonth);

      // 管理画面表示処理
      const adminHtml = HtmlService.createTemplateFromFile('admin');
      adminHtml.makeMonth = [data1, data2];
      adminHtml.title = 'ボランティア予約';
      adminHtml.err = '';
      adminHtml.url = url;
      return adminHtml.evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 管理者画面
  if(e.parameter.admin){

    const data1 = getYoyaku(thisMonth);
    const data2 = getYoyaku(nextMonth);

    if(e.parameter.user == chkId || e.parameter.modoru == "modoru"){
      // 管理画面表示処理
      const adminHtml = HtmlService.createTemplateFromFile('admin');
      adminHtml.makeMonth = [data1, data2];
      adminHtml.title = 'ボランティア予約';
      adminHtml.err = '';
      adminHtml.url = url;
      return adminHtml.evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    // 認証通らない場合はトップ表示
    const indexHtml = HtmlService.createTemplateFromFile('index');
    indexHtml.err = '';
    indexHtml.startDay = startDay;
    indexHtml.makeMonth = [data1, data2];
    indexHtml.setTitle = 'ボランティア予約';
    indexHtml.title = 'ボランティア予約';
    indexHtml.url = url;
    return indexHtml.evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 予約登録へ
  const param = {
    date: e.parameter.date,
    grade: e.parameter.grade,
    cl: e.parameter.cl,
    num: e.parameter.num,
    mail: e.parameter.mail
  }

  const resultHtml = HtmlService.createTemplateFromFile('result');
  resultHtml.title ="予約結果";

  // シート名取得
  const name = param.date.replace('/', '').slice(0, 6);
  const dataSheet = set_sheet(name);

  // 登録済みデータをスプレッドシートより取得
  const data = dataSheet.getRange(1, 1, dataSheet.getLastRow(), 6).getDisplayValues();
  // メールアドレスから先頭6文字取得
  const chkMail = param.mail.slice(0, 6)
  // 重複予約チェック(日付、メールアドレス)
  if(mailCheck(data, param.date,chkMail) ){
    resultHtml.err = 'すでにご登録いただいたメールアドレスで登録されています';
    return resultHtml.evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // スプレッドシートに登録
  const row = data.findIndex( d => d[0] == param.date)
  dataSheet.getRange(row + 1,4).insertCells(SpreadsheetApp.Dimension.COLUMNS);

  dataSheet.getRange(row + 1,4).setValue(param.grade + "年" +  param.cl + "組" + param.num + "番" + chkMail);
  dataSheet.getRange(row + 1,3).setValue(dataSheet.getRange(row + 1,3).getDisplayValues() - 1);

  okMail(param);

  resultHtml.err = param.date + 'の予約を受付ました。<br />メールをご確認ください。<br />予約変更の際は、届いたメールに返信でご連絡してください。<br />ありがとうございました。';
  resultHtml.url = url;

  return resultHtml.evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1');
}



// ------------------------------------------------

// 同じ日に同じメールアドレスからの予約がないかチェック
function mailCheck(data, date, mail){
  const chkDate = data.filter( chk => chk[0] == date);
  const chkMail = chkDate.flat().filter( chk => chk.includes(mail));
  return chkMail.length > 0 ;
}


// 予約日のシート準備
function set_sheet(shName, makeDate){

  let sheet = spreadsheet.getSheetByName(shName)
  // 当月シートがなければ作成
  if(!sheet){
    sheet = spreadsheet.insertSheet(shName);
    sheet.getRange(1, 1).setValue('日付');
    sheet.getRange(1, 2).setValue('曜日');
    sheet.getRange(1, 3).setValue('残数');
    sheet.getRange(1, 4).setValue('予約者1');
    sheet.getRange(1, 5).setValue('予約者2');
    sheet.getRange(1, 6).setValue('予約者3');

    let getDay = makeDate;
    let chkMonth = makeDate.month();

    let i = 2;
    while (getDay.month() == chkMonth){

      sheet.getRange(i, 1).setValue(getDay.format('YYYY/MM/DD'));
      sheet.getRange(i, 2).setValue(days[getDay.day()]);
      if (getDay.day() == 0 || getDay.day() == 6){
        sheet.getRange(i, 3).setValue("x");
      }else{
        sheet.getRange(i, 3).setValue("3");
      }
      getDay = getDay.add(1, 'days');
      i += 1;
    }
  }
  return sheet;

}

// メール送信
function okMail(p){
  let recipient = p.mail;
  let subject = p.date + "_予約受付しました";
  let body = "ボランティア申し込みありがとうございます。\n\n";
  body += "予約日：" + p.date + "\n\n";
  body += p.grade + "年" + p.cl + "組" + p.num + "番" + "\n\n";
  body += "予約キャンセルされる場合は、このメールへの返信でご連絡ください。\n\n";
  body += "本メールにお心当たりのない場合、破棄して頂けますようお願いいたします。";

  GmailApp.sendEmail(recipient, subject, body);

}


// 外部ファイル　インクルード（CSS,JSファイル）
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
