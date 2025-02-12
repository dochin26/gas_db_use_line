
// 利用しているシート
var SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
// 利用しているSSのシート名（※変えるとみえなくなる）
var SHEET_NAME = 'db';
// 利用しているもしかしてSSのシート名（※変えるとみえなくなる）
var SHEET_NAME_MAYBE = 'maybe';

// LINE Message API アクセストークン
var ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('ACCESS_TOKEN');
// 通知URL
var PUSH = "https://api.line.me/v2/bot/message/push";
// リプライ時URL
var REPLY = "https://api.line.me/v2/bot/message/reply";
// プロフィール取得URL
var PROFILE = "https://api.line.me/v2/profile";

/**
 * doPOST
 * POSTリクエストのハンドリング
 */
function doPost(e) {
  var json = JSON.parse(e.postData.contents);
  reply(json);
}

/** 
 * doGet
 * GETリクエストのハンドリング
 */
function doGet(e) {
    return ContentService.createTextOutput("SUCCESS");
}

/** 
 * reply
 * ユーザからのアクションに返信する
 */
function reply(data) {
  // POST情報から必要データを抽出
  var lineUserId = data.events[0].source.userId;
  var postMsg    = data.events[0].message.text;
  var replyToken = data.events[0].replyToken;
  var action    = data.events[0].message.action;
  // 記録用に検索語とuserIdを記録
  debug(postMsg, lineUserId);
//  debug(action, lineUserId);

  // 検索語に対しての回答をSSから取得
  var answers = findResponseArray(postMsg);

  // 回答メッセージを作成
  var replyText = '「' + postMsg + '」の回答です。';
  replyText += '\n\n＝＝＝＝＝＝＝＝＝＝＝＝＝\n\n';

  if (answers.length >= 3) {
    var addResult = 0;
    var subtractResult = [];
  
    for (var i = 0; i < answers.length; i++) {
      var answer = parseFloat(answers[i].value); // 変数の宣言修正
      subtractResult.push(answer); // 配列に値を追加する修正
      addResult += answer;
      replyText += '回答' + (i +1) + '：' + answers[i].key + '　' + answers[i].value + 'ページ\n';
    }
    
    var sort = subtractResult.sort((a, b) => b - a); // ソートして最大値を取得
    var times = subtractResult[0] * 2; // 最大値に2倍をする

    replyText += '\n全合計値：' + addResult + '\n';

    if(answers.length === 2) {
      var subtract = sort[0] - sort[1];
      replyText += '引算結果：' + subtract + '\n';
    } else {
    replyText += '最大値 ' + subtractResult[0] + ' の2倍値：' + times;
    }

  } 
  // else if (answers.length === 2) {

  //   var answer1 = parseFloat(answers[0].value);
  //   var answer2 = parseFloat(answers[1].value);
  //   var addResult = answer1 + answer2;
  //   var subtractResult1 = answer1 - answer2;
  //   var subtractResult2 = answer2 - answer1;
  //   var subtractResult3 = answer1 * 2;
  //   var subtractResult4 = answer2 * 2;

  //   replyText += '\n\n＝＝＝＝＝＝＝＝＝＝＝＝＝\n\n';
  //   replyText += '1つ目の回答：' + answers[0].key + ' (' + answers[0].value + ')\n';
  //   replyText += '2つ目の回答：' + answers[1].key + ' (' + answers[1].value + ')\n';
  //   replyText += '\n足し算結果：' + addResult + '\n';
  //   replyText += '引き算結果1：' + subtractResult1 + '\n';
  //   replyText += '引き算結果2：' + subtractResult2 + '\n';
  //   replyText += '　倍数結果1：' + subtractResult3 + '\n';
  //   replyText += '　倍数結果2：' + subtractResult4 + '\n';

  // } 
  else if (answers.length === 1) {
    // 1つだけの場合の処理
    replyText += '回答：' + answers[0].key + '　' + answers[0].value + 'ページ\n';
  } else {
    // その他の場合の処理
    replyText += '答えが見つかりませんでした。別のキーワードで質問してみてください。';
  }

  sendMessage(replyToken, replyText);
}
  


//   // 回答の有無に応じて分岐
//   if (answers.length === 0) {
//     // 「類似の検索キーワード」がないかチェック
//     var mayBeWord = findMaybe(postMsg);
//     if (typeof mayBeWord === "undefined") {
//       // 回答がない場合の定型文
//       sendMessage(replyToken, '答えが見つかりませんでした。別のキーワードで質問してみてください。');        
//     } else {
//       sendMayBe(replyToken, mayBeWord);
//     }
//   } else {
//     // 回答がある場合のメッセージ生成
//     answers.forEach(function(answer) {
//       replyText = replyText + "\n\n＝＝＝＝＝＝＝＝＝＝＝＝＝\n\nQ：" + answer.key + "\n\nA：" + answer.value + " ページ";
//     });

//     // 1000文字を超える場合は途中で切る
//     if (replyText.length > 1000) {
//       replyText = replyText.slice(0,1000) + "……\n\n＝＝＝＝＝＝＝＝＝＝＝＝＝\n\n回答文字数オーバーです。詳細に検索キーワードを絞ってください。";
//     }
//     // メッセージAPI送信
//     sendMessage(replyToken, replyText);
//   }
// }

// SSからデータを取得
function getData() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  return data.map(function(row) { return {key: row[0], value: row[1], type: row[2]}; });
}

// SSから「もしかして」データを取得
function getMayBeData() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_MAYBE);
  var data = sheet.getDataRange().getValues();
  return data.map(function(row) { return {key: row[0], value: row[1], type: row[2]}; });
}

// // 単語が一致したセルの回答を配列で返す
// function findResponseArray(word) {
//   // スペース検索用のスペースを半角に統一
//   word = word.replace('　',' ');
//   // 単語ごとに配列に分割
//   var wordArray = word.split(' ');
//   return getData().reduce(function(memo, row) {
//     // 値が入っているか
//     if (row.value) {
//       // AND検索ですべての単語を含んでいるか
//       var matchCnt = 0;
//       wordArray.forEach(function(wordUnit) {
//         // 単語を含んでいればtrue
//         if (row.key.indexOf(wordUnit) > -1) {
//           matchCnt = matchCnt + 1;
//         }
//       });
//       if (wordArray.length === matchCnt) {
//         memo.push(row);
//       }
//     }
//     return memo;
//   }, []) || [];
// }

function findResponseArray(words) {
  // 入力文字列をスペースで分割し、半角スペースに統一
  var wordArray = words.replace(/　/g, ' ').split(' ');
  var data = getData();
  var responses = [];

  wordArray.forEach(function(word) {
    // 各単語に対してデータを検索
    var response = data.find(function(row) {
      // row.keyを小文字に変換
      var keyLower = row.key.toLowerCase();
      // 入力された単語がすべて含まれているか確認するためのフラグ
      var containsAllWords = true;
      for (var i = 0; i < wordArray.length; i++) {
        if (!keyLower.includes(wordArray[i])) {
          containsAllWords = false;
          break;  // 一致しない単語があればループを抜ける
        }
      }
      return containsAllWords;
    });
    if (response) {
      responses.push(response);
    }
  });
  return responses;
}


// function findResponseArray(word) {
//   // スペース検索用のスペースを半角に統一
//   word = word.replace('　',' ');
//   // 単語ごとに配列に分割
//   var wordArray = word.split(' ');
//   return getData().filter(function(row) {
//     // 値が入っているかつ部分一致するかを検索
//     return row.value && wordArray.some(function(wordUnit) {
//       return row.key.includes(wordUnit);
//     });
//   });
// }

// 単語が一致したセルの回答を「もしかして」を返す
function findMaybe(word) {
  return getMayBeData().reduce(function(memo, row) { return memo || (row.key === word && row.value); }, false) || undefined;
}

// 画像形式でAPI送信
function sendMessageImage(replyToken, imageUrl) {
  // replyするメッセージの定義
  var postData = {
    "replyToken" : replyToken,
    "messages" : [
      {
        "type": "image",
        "originalContentUrl": imageUrl
      }
    ]
  };
  return postMessage(postData);
}

// LINE messaging apiにJSON形式でデータをPOST
function sendMessage(replyToken, replyText) {  
  // replyするメッセージの定義
  var postData = {
    "replyToken" : replyToken,
    "messages" : [
      {
        "type" : "text",
        "text" : replyText
      }
    ]
  };
  return postMessage(postData);
}

// LINE messaging apiにJSON形式で確認をPOST
function sendMayBe(replyToken, mayBeWord) {  
  // replyするメッセージの定義
  var postData = {
    "replyToken" : replyToken,
    "messages" : [
      {
        "type" : "template",
        "altText" : "もしかして検索キーワードは「" + mayBeWord + "」ですか？",
        "template": {
          "type": "confirm",
          "actions": [
            {
                "type":"postback",
                "label":"はい",
                "data":"action=detail",
            },
            {
                "type": "message",
                "label": "いいえ",
                "text": "いいえ、違います。"
            }
          ],
          "text": "答えが見つかりませんでした。もしかして検索キーワードは「" + mayBeWord + "」ですか？"
        }

      }
    ]
  };
  return postMessage(postData);
}

// LINE messaging apiにJSON形式でデータをPOST
function postMessage(postData) {  
  // リクエストヘッダ
  var headers = {
    "Content-Type" : "application/json; charset=UTF-8",
    "Authorization" : "Bearer " + ACCESS_TOKEN
  };
  // POSTオプション作成
  var options = {
    "method" : "POST",
    "headers" : headers,
    "payload" : JSON.stringify(postData)
  };
  return UrlFetchApp.fetch(REPLY, options);      
}

/** ユーザーのアカウント名を取得
 */
function getUserDisplayName(userId) {
  var url = 'https://api.line.me/v2/bot/profile/' + userId;
  var userProfile = UrlFetchApp.fetch(url,{
    'headers': {
      'Authorization' :  'Bearer ' + ACCESS_TOKEN,
    },
  })
  return JSON.parse(userProfile).displayName;
}

// userIdシートに記載
function lineUserId(userId) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('userId');
  sheet.appendRow([userId]);
}

// debugシートに値を記載
function debug(text, userId) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('debug');
  var date = new Date();
  var userName = getUserDisplayName(userId);
  sheet.appendRow([userId, userName, text, Utilities.formatDate( date, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')]);
}

