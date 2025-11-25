/* ==============================================
   Code.gs: バックエンドロジック (署名追加版)
   ============================================== */

// ページルーティング
function doGet(e) {
  if (e.parameter && e.parameter.p === 'day2') {
    return HtmlService.createHtmlOutputFromFile('day2')
        .setTitle('事後実験 (Day 2)')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('記憶実験 (Day 1)')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

// データ保存
function saveData(data) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const timestamp = new Date();
    let rowData = [];

    const wmcFor = Number(data.scoreForward) || 0;
    const wmcBack = Number(data.scoreBackward) || 0;
    const wmcTotal = wmcFor + wmcBack;
    
    let tlxTotal = 0;
    if (data.nasaTlx) {
      for (let key in data.nasaTlx) {
        tlxTotal += Number(data.nasaTlx[key]) || 0;
      }
    }

    if (data.dataType === 'wmc') {
      rowData = [timestamp, data.participantId, 'wmc', '', wmcFor, wmcBack, wmcTotal, '', '', '', '', '', '', '', ''];
    } else if (data.dataType === 'experiment') {
      rowData = [
        timestamp,
        data.participantId,
        'experiment',
        data.condition,
        '', '', '', 
        data.recallText,
        data.quizScore,
        tlxTotal,
        data.memoText || '',
        JSON.stringify(data.nasaTlx),
        data.email || '',
        data.materialIdx,
        '' // 送信フラグ
      ];
    } else if (data.dataType === 'day2') {
      rowData = [timestamp, data.participantId, 'day2', data.condition, '', '', '', data.recallText, data.quizScore, '', '', '', '', '', ''];
    }

    sheet.appendRow(rowData);
    return { result: "success" };

  } catch (e) {
    return { result: "error", error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// Day 1履歴取得
function getParticipantData(participantId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const logs = data.slice(1).filter(row => row[1] === participantId && row[2] === 'experiment');
  logs.sort((a, b) => new Date(a[0]) - new Date(b[0]));
  return logs.map(row => ({ condition: row[3], materialIdx: row[13] }));
}

// メール送信トリガー関数（署名付き）
function sendFollowUpEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  const ID_COL = 1;     // B列
  const TYPE_COL = 2;   // C列
  const TIME_COL = 0;   // A列
  const EMAIL_COL = 12; // M列
  const SENT_FLAG_COL = 14; // O列

  const now = new Date();
  const WAIT_TIME = 5 * 60 * 1000; // ★テスト用: 5分 (本番は 24 * 60 * 60 * 1000)

  // ★あなたのWebアプリURLに書き換えてください
  const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzPNP0tiDrkvDkRlQguL_5L2L6RW-L1LcF5P0dZrjfNg9pKhCvVnoek4NIEo-MCmst0/exec"; 
  const DAY2_URL = SCRIPT_URL + "?p=day2";

  // 1. データを参加者IDごとにまとめる
  const participants = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const type = row[TYPE_COL];
    const id = row[ID_COL];

    if (type !== 'experiment') continue;

    if (!participants[id]) {
      participants[id] = {
        email: row[EMAIL_COL],
        rows: [],      
        isSent: false, 
        timestamp: new Date(row[TIME_COL])
      };
    }
    participants[id].rows.push(i + 1);
    
    if (row[EMAIL_COL] && String(row[EMAIL_COL]).includes('@')) {
      participants[id].email = row[EMAIL_COL];
    }
    
    if (row[SENT_FLAG_COL] === "Sent") {
      participants[id].isSent = true;
    }
  }

  // 2. 参加者ごとに判定してメール送信
  for (const id in participants) {
    const p = participants[id];

    if (p.isSent) continue;
    if (!p.email || !String(p.email).includes('@')) continue;

    if (now.getTime() - p.timestamp.getTime() >= WAIT_TIME) {
      try {
        MailApp.sendEmail({
          to: p.email,
          subject: "【記憶実験】事後テストのお願い",
          body: `
実験にご協力いただきありがとうございます。
所定の時間が経過しましたので、事後テストへの回答をお願いいたします。

【重要：URLが開けない場合】
リンクを開いた際にエラーが表示される場合は、URLをコピーし、
ブラウザの「シークレットウィンドウ」または「プライベートブラウズ」で開いてください。

URL: ${DAY2_URL}

IDはDay1と同じものを入力してください。
所要時間は10分程度です。

--------------------------------------------------
【お問い合わせ先】
〇〇大学 〇〇学部 〇〇研究室
実験責任者：[name]
連絡先：[⬜︎⬜︎.com]
--------------------------------------------------
          `
        });
        
        console.log(`Email sent to ${p.email} (ID: ${id})`);

        p.rows.forEach(rowNum => {
          sheet.getRange(rowNum, SENT_FLAG_COL + 1).setValue("Sent");
        });
        
      } catch (e) {
        console.error(`Failed to send email to ${p.email}: ${e}`);
      }
    }
  }
}
