/**
 * デイリーダイアリーアプリ用のGoogle Apps Scriptコード
 *
 * 【設定手順】
 * 1. Googleアカウントで新しいスプレッドシートを1つ作成する
 * 2. メニューの「拡張機能」→「Apps Script」を開く
 * 3. デフォルトで書かれているコードを全部消して、このファイルの内容を全部貼り付ける
 * 4. 上部の「デプロイ」→「新しいデプロイ」をクリック
 * 5. 種類の選択(歯車アイコン)で「ウェブアプリ」を選ぶ
 * 6. 「次のユーザーとして実行」→ 自分
 *    「アクセスできるユーザー」→ 全員(Anyone)
 *    にして「デプロイ」をクリック
 * 7. 初回は権限の承認を求められるので許可する
 * 8. 発行された「ウェブアプリURL」をコピーして、
 *    デイリーダイアリーアプリの「履歴・出力」タブ→設定 の
 *    「Googleスプレッドシート連携URL」欄に貼り付ける
 *
 * これで、アプリの保存ボタンを押すたびに、ローカル保存に加えて
 * このスプレッドシートにも自動で1行(その日の記録)が追記/上書きされる。
 * 日付が重複する場合は同じ行を上書きする(1日1行になる)。
 */

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = JSON.parse(e.postData.contents);

  // 既存のヘッダー行を取得
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(String);

  // ヘッダーが空(初回)なら、送られてきたデータのキーからヘッダーを作る
  if (headers.length === 0) {
    headers = Object.keys(data);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // 送られてきたデータに、まだヘッダーにない項目があれば列を追加する
  const missing = Object.keys(data).filter(function (k) {
    return headers.indexOf(k) === -1;
  });
  if (missing.length > 0) {
    headers = headers.concat(missing);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const row = headers.map(function (h) {
    return data[h] !== undefined && data[h] !== null ? data[h] : '';
  });

  // 同じ日付(date列)の行が既にあれば上書き、なければ追記
  const dateColIdx = headers.indexOf('date');
  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  if (lastRow > 1 && dateColIdx !== -1) {
    const dates = sheet.getRange(2, dateColIdx + 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < dates.length; i++) {
      if (dates[i][0] === data.date) {
        targetRow = i + 2;
        break;
      }
    }
  }

  if (targetRow === -1) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ブラウザで直接URLを開いた時の簡易確認用(POST以外のアクセス)
function doGet(e) {
  return ContentService.createTextOutput('デイリーダイアリー連携用エンドポイントは正常に動作しています。');
}
