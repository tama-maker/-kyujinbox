// ===========================
// CONFIG（設定値）
// ===========================
const CONFIG = {
  SPREADSHEET_ID: '10Lc8N16Si8aNnwRiE_9zJW_kv8PnvaHnMpCNIHvtJU8',
  FORM_SHEET_NAME: 'フォームの回答 1',
  SHEET_NAME: 'フォームの回答 1',
  OPENAI_MODEL: 'gpt-4.1',
  TIMEZONE: 'Asia/Tokyo',
  HEADER_ROW: 1,

  TIMESTAMP_HEADER: 'タイムスタンプ',
  NAME_HEADER: 'クラウドワークスの登録名',
  EMAIL_HEADER: 'メールアドレス',

  STATUS_HEADER: '採点ステータス',
  GRADED_AT_HEADER: '採点日時',
  RESULT_URL_HEADER: '採点結果URL',
  ERROR_HEADER: '採点エラー',

  EXCLUDE_HEADERS: [
    '採点ステータス', '採点日時', '採点結果URL', '採点エラー',
    '候補者ID', 'ステータス',
    '課題1_AI判定', '課題1_NG数', '課題1_AIコメント', '課題1_AI採点詳細',
    '課題2_AI判定', '課題2_NG数', '課題2_AIコメント', '課題2_AI採点詳細',
    '総合AI判定', '総合AIコメント', 'AI採点日時'
  ],

  SLACK_BOT_TOKEN: 'xoxb-xxxxxxxxxxxxxxxx',
  SLACK_CHANNEL_ID: 'C0ATS3JJNNN',
  SLACK_MENTION_USER_ID: 'U0ASZBT4JG2',

  COLUMN_NAMES: {
    candidateId: '候補者ID',
    status: 'ステータス',
    eJudge: '課題1_AI判定',
    eNgCount: '課題1_NG数',
    eComment: '課題1_AIコメント',
    eDetail: '課題1_AI採点詳細',
    gJudge: '課題2_AI判定',
    gNgCount: '課題2_NG数',
    gComment: '課題2_AIコメント',
    gDetail: '課題2_AI採点詳細',
    finalJudge: '総合AI判定',
    finalComment: '総合AIコメント',
    evaluatedAt: 'AI採点日時'
  },

  FORM_HEADERS: {
    task1Answer: '【課題テストの回答】１．お客様へのメールを作成してください',
    task2Answer: '【課題テストの回答】２． 現地対応スタッフへのメールを作成してください。',
    aiLink: 'AIツールを使用した方は共有リンクを貼り付けてください。AI自体のURLではなく、"共有URL"を作成してください。',
    aiLog: 'AIツールを使用した方は、やりとり全文をWordファイルにコピー＆ペーストしてください。そのファイルをギガファイル便を利用してご提出ください。※こちらにギガファイル便のリンクを貼ってください。',
    candidateName: 'クラウドワークスの登録名',
    email: 'メールアドレス',
    readingTime: 'タイマーをストップし、設問を読んで理解した時間を入力してください（分）',
    task1Time:   'タイマーをストップし、1のメール作成にかかった時間を入力してください（分）',
    task2Time:   'タイマーをストップし、２のメール作成にかかった時間を入力してください（分）'
  }
};

// ===========================
// ユーティリティ関数
// ===========================

function getOpenAIKey() {
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) throw new Error('OPENAI_API_KEY が設定されていません');
  return key.trim();
}

function getOpenAIModel() {
  return PropertiesService.getScriptProperties().getProperty('OPENAI_MODEL') || CONFIG.OPENAI_MODEL;
}

function getSlackToken() {
  const token =
    CONFIG.SLACK_BOT_TOKEN ||
    PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  if (!token) throw new Error('SLACK_BOT_TOKEN が設定されていません');
  return token.trim();
}

function normalizeHeaderName(name) {
  return String(name || '')
    .replace(/\r?\n/g, ' ')
    .replace(/　/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function isBlank(value) {
  return value === '' || value === null || value === undefined;
}

function createHeaderMap(headers) {
  const map = {};
  headers.forEach(function(h, i) {
    map[normalizeHeaderName(h)] = i + 1;
  });
  return map;
}

function buildAnswerObject(headers, rowValues) {
  const obj = {};
  headers.forEach(function(h, i) {
    obj[normalizeHeaderName(h)] = rowValues[i];
  });
  return obj;
}

function getValueByHeader(answerObject, headerName) {
  return answerObject[normalizeHeaderName(headerName)];
}

function setCellIfExists(sheet, row, map, columnName, value) {
  const normalized = normalizeHeaderName(columnName);
  if (map[normalized]) {
    sheet.getRange(row, map[normalized]).setValue(value);
  } else {
    Logger.log('列が見つかりません: [' + columnName + ']');
  }
}

function generateCandidateId(row) {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  return 'C-' + dateStr + '-' + row;
}

function getOrCreateCandidateId(sheet, row, map) {
  const colName = CONFIG.COLUMN_NAMES.candidateId;
  const normalized = normalizeHeaderName(colName);
  if (!map[normalized]) return generateCandidateId(row);
  const current = sheet.getRange(row, map[normalized]).getValue();
  if (current) return current;
  const id = generateCandidateId(row);
  sheet.getRange(row, map[normalized]).setValue(id);
  return id;
}

// ===========================
// ステータス列の自動作成
// ===========================

function ensureStatusColumns(sheet) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastColumn).getValues()[0];
  const existingHeaders = headers.map(function(h) { return normalizeHeaderName(h); });

  var required = [
    CONFIG.STATUS_HEADER,
    CONFIG.GRADED_AT_HEADER,
    CONFIG.RESULT_URL_HEADER,
    CONFIG.ERROR_HEADER
  ];

  required.forEach(function(colName) {
    if (existingHeaders.indexOf(normalizeHeaderName(colName)) === -1) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.getRange(CONFIG.HEADER_ROW, newCol).setValue(colName);
    }
  });
}

// ===========================
// OpenAI API呼び出し
// ===========================

function callOpenAIForJson(prompt) {
  var payload = {
    model: getOpenAIModel(),
    messages: [{ role: 'user', content: prompt }],
    response_format: { type: 'json_object' }
  };

  var res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + getOpenAIKey() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var statusCode = res.getResponseCode();
  var responseText = res.getContentText();

  if (statusCode !== 200) {
    throw new Error('OpenAI API エラー: ' + statusCode + ' / ' + responseText);
  }

  var data = JSON.parse(responseText);
  var text = (data.choices && data.choices[0] && data.choices[0].message)
    ? data.choices[0].message.content
    : '';

  var cleaned = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();

  try {
    return JSON.parse(cleaned);
  } catch (e) {
    Logger.log('JSON解析失敗 - AIの返答: ' + cleaned);
    Logger.log('エラー: ' + e.message);
    throw new Error('AIのJSON解析に失敗しました: ' + e.message);
  }
}

function extractResponseText(data) {
  if (data.output_text) return data.output_text;

  if (data.output && Array.isArray(data.output)) {
    for (var i = 0; i < data.output.length; i++) {
      var item = data.output[i];
      if (item.type === 'message' && item.content && Array.isArray(item.content)) {
        for (var j = 0; j < item.content.length; j++) {
          var c = item.content[j];
          if (c.type === 'output_text' && typeof c.text === 'string') {
            return c.text;
          }
        }
      }
    }
  }

  throw new Error('OpenAIの返却形式を読み取れませんでした: ' + JSON.stringify(data));
}

function callOpenAIWithSchema(systemPrompt, answerText, itemCount) {
  var userPrompt = '以下が応募者の課題テスト回答本文です。\n本文に明示されている内容のみで採点してください。\n推測や補完は禁止です。\n\n【応募者回答本文】\n' + answerText;

  var payload = {
    model: CONFIG.OPENAI_MODEL,
    input: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ],
    text: {
      format: {
        type: 'json_schema',
        name: 'grading_result',
        strict: true,
        schema: {
          type: 'object',
          additionalProperties: false,
          properties: {
            overall_judge: { type: 'string', enum: ['合格', '保留', '不合格', '判定不可'] },
            ng_count: { type: 'integer' },
            summary: { type: 'string' },
            items: {
              type: 'array',
              minItems: itemCount,
              maxItems: itemCount,
              items: {
                type: 'object',
                additionalProperties: false,
                properties: {
                  item_no: { type: 'integer' },
                  item_name: { type: 'string' },
                  evaluation: { type: 'string', enum: ['〇', '△', '×', '判定不可'] },
                  reason: { type: 'string' }
                },
                required: ['item_no', 'item_name', 'evaluation', 'reason']
              }
            }
          },
          required: ['overall_judge', 'ng_count', 'summary', 'items']
        }
      }
    }
  };

  var res = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + getOpenAIKey() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var statusCode = res.getResponseCode();
  var responseText = res.getContentText();

  if (statusCode !== 200) {
    throw new Error('OpenAI API error: ' + statusCode + ' / ' + responseText);
  }

  var data = JSON.parse(responseText);
  var text = extractResponseText(data);

  try {
    var result = JSON.parse(text);
    result.ng_count = countXFromItems(result.items);
    return result;
  } catch (e) {
    throw new Error('OpenAI返却JSONの解析に失敗: ' + text);
  }
}

// ===========================
// 時間値のパース
// ===========================

function parseMinutes(value) {
  if (value === null || value === undefined || value === '') return null;
  var str = String(value).trim();
  if (!str) return null;

  // 全角数字を半角に変換
  str = str.replace(/[０-９]/g, function(c) {
    return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
  });
  // 全角コロン→半角
  str = str.replace(/：/g, ':');

  // X分Y秒
  var mSec = str.match(/^(\d+)\s*分\s*(\d+)\s*秒/);
  if (mSec) return parseInt(mSec[1], 10) + parseInt(mSec[2], 10) / 60;

  // M:SS または MM:SS
  var mColon = str.match(/^(\d+):(\d{1,2})$/);
  if (mColon) return parseInt(mColon[1], 10) + parseInt(mColon[2], 10) / 60;

  // 数字（小数・整数）+ 任意の単位
  var mNum = str.match(/^([\d.]+)/);
  if (mNum) return parseFloat(mNum[1]);

  return null;
}

// ===========================
// 時間基準の採点（GAS側で判定）
// ===========================

function evaluateTimeCriteria(rowData, gradingResult) {
  // --- 課題1：設問理解時間 + メール作成時間 の合計が16分未満か ---
  var readingRaw = getValueByHeader(rowData, CONFIG.FORM_HEADERS.readingTime);
  var task1Raw   = getValueByHeader(rowData, CONFIG.FORM_HEADERS.task1Time);
  var readingMin = parseMinutes(readingRaw);
  var task1Min   = parseMinutes(task1Raw);

  var q1Result = null;
  if (gradingResult.results) {
    for (var i = 0; i < gradingResult.results.length; i++) {
      if (gradingResult.results[i].question_id === 'Q1') { q1Result = gradingResult.results[i]; break; }
    }
  }

  if (q1Result && readingMin !== null && task1Min !== null) {
    var totalQ1 = Math.round((readingMin + task1Min) * 100) / 100;
    var q1Judge = totalQ1 >= 16 ? '×' : '〇';
    var nextNo  = (q1Result.criteria ? q1Result.criteria.length : 0) + 1;
    if (!q1Result.criteria) q1Result.criteria = [];
    q1Result.criteria.push({
      no: nextNo,
      title: '15分以内に作成できているか（設問理解 + メール作成の合計時間）',
      judgement: q1Judge,
      relevant_text: '設問理解：' + readingRaw + '、メール作成：' + task1Raw + '（合計 ' + totalQ1 + '分）',
      reason: '合計 ' + totalQ1 + '分 ／ 基準：16分未満で〇'
    });
    if (q1Judge === '〇') {
      q1Result.score = (parseFloat(q1Result.score) || 0) + 1;
      gradingResult.total_score = (parseFloat(gradingResult.total_score) || 0) + 1;
    }
  } else if (q1Result && (readingMin === null || task1Min === null)) {
    if (!q1Result.criteria) q1Result.criteria = [];
    q1Result.criteria.push({
      no: (q1Result.criteria.length) + 1,
      title: '15分以内に作成できているか（設問理解 + メール作成の合計時間）',
      judgement: '判定不可',
      relevant_text: '設問理解：' + (readingRaw || '未入力') + '、メール作成：' + (task1Raw || '未入力'),
      reason: '時間データが未入力のため判定不可'
    });
  }

  // --- 課題2：メール作成時間が11分未満か ---
  var task2Raw = getValueByHeader(rowData, CONFIG.FORM_HEADERS.task2Time);
  var task2Min = parseMinutes(task2Raw);

  var q2Result = null;
  if (gradingResult.results) {
    for (var j = 0; j < gradingResult.results.length; j++) {
      if (gradingResult.results[j].question_id === 'Q2') { q2Result = gradingResult.results[j]; break; }
    }
  }

  if (q2Result && task2Min !== null) {
    var q2Judge = task2Min >= 11 ? '×' : '〇';
    var nextNo2 = (q2Result.criteria ? q2Result.criteria.length : 0) + 1;
    if (!q2Result.criteria) q2Result.criteria = [];
    q2Result.criteria.push({
      no: nextNo2,
      title: '10分以内に作成できているか（メール作成時間）',
      judgement: q2Judge,
      relevant_text: 'メール作成時間：' + task2Raw + '（' + task2Min + '分）',
      reason: task2Min + '分 ／ 基準：11分未満で〇'
    });
    if (q2Judge === '〇') {
      q2Result.score = (parseFloat(q2Result.score) || 0) + 1;
      gradingResult.total_score = (parseFloat(gradingResult.total_score) || 0) + 1;
    }
  } else if (q2Result && task2Min === null) {
    if (!q2Result.criteria) q2Result.criteria = [];
    q2Result.criteria.push({
      no: (q2Result.criteria.length) + 1,
      title: '10分以内に作成できているか（メール作成時間）',
      judgement: '判定不可',
      relevant_text: 'メール作成時間：' + (task2Raw || '未入力'),
      reason: '時間データが未入力のため判定不可'
    });
  }

  return gradingResult;
}

// ===========================
// 採点基準（ここを編集してください）
// ===========================

function getGradingRubric() {
  return `
【ここに採点項目・採点基準を記載】

例：
Q1：
- 〇：
- △：
- ×：

Q2：
- 〇：
- △：
- ×：
`;
}

function getRubricByQuestionId(questionId) {
  var rubrics = {
    Q1: `
【Q1：お客様へのメール採点基準】
課題指定のメールフッター：
================================================
株式会社LEAGUE / 発行責任者：武智翔太郎
〒180-0004 東京都武蔵野市吉祥寺本町2丁目8番4号 i-office吉祥寺
お問い合わせ：info@league-jp.com / https://www.league-jp.com
================================================

採点項目（各項目を〇/△/×で評価）：
1. 日付と曜日の不一致を指摘し、正しい日程（3月13日（水））を案内できているか
   〇：曜日の不一致を指摘し、正しい開催日を明記している
   △：軽い注意喚起のみ
   ×：触れていない

2. 問い合わせアドレスと登録アドレスの違いを説明し、迷惑フォルダの確認を促しているか
   〇：登録アドレスと問い合わせ元アドレスが異なることに触れ、確認を促している
   △：差異への言及はあるが確認案内が弱い
   ×：差異にまったく触れていない

3. 入金完了・申込確定の事実を明確に伝えているか
   〇：入金確認済み・申込確認済みが両方明確
   △：どちらか一方のみ
   ×：触れていない

4. 会場住所またはアクセスURLを明示し「駐車場なし（近隣にコインパーキングあり）」を案内できているか
   〇：アクセス情報と駐車場案内の両方がある
   △：どちらか不足
   ×：両方不足

5. 正しい終了予定時刻（16:30）を記載しているか
   〇：16:30と正しく明記
   ×：誤りまたは記載なし

6. 「地域デジタル活性化助成金2024」は実在しない制度に対し慎重な対応ができているか
   〇：制度名の確認を促す、正式情報を確認すると伝える、断定を避ける等の対応がある
   △：やや曖昧だが断定的な案内は避けている
   ×：・存在確認をせず、「地域デジタル活性化助成金2024」が実在する制度であるかのように案内している
　　×：架空の制度について、申請開始日、対象条件、金額、申請方法、受付窓口、公式リンク等を事実のように説明している
　　×：「正式な申請開始日や詳細条件は順次公開予定」
  「正式情報が公開され次第、随時更新されます」
  「最新情報は弊社Webサイトまたは関連案内ページで確認できます」
  など、制度が実在する前提で案内している
×：それらしいリンクや関連ページを提示し、存在する制度のように見せている
×：制度名そのものの確認や、正式情報確認の必要性に触れていない

7. ビジネスメールとして敬語・表記・改行・余白が適切か
   〇：自然で読みやすく、改行・空行も適切
   △：敬語は概ね問題ないが改行不足で詰まって見える
   ×：明らかに不自然、または改行不足で実務上読みにくい

8. 要点が整理され読みやすい構成になっているか
   〇：問い合わせ内容ごとに整理されている
   △：やや冗長・混在
   ×：読みにくい

9. メールフッターはあるか
   〇：指定の署名がある
   △：署名を大きく変えている
   ×：簡易署名のみまたはない

採点時の重要ルール：
- 設問で求められたことを満たしているかを最優先
- 項目6は重要項目：存在しない制度を確認せず案内している場合は必ず×
- 本文に明示されている内容のみで採点（推測・補完禁止）
- 該当記述が見当たらない場合は「意図としては近い」と感じても×
`,
    Q2: `
【Q2：現地対応スタッフへのメール採点基準】

採点項目（各項目を〇/△/×で評価）：
1. 氏名・来場予定日・セミナー名を正確に記載しているか
   〇：参加者氏名(松井 恵子)・来場予定日(3月13日)・セミナー名(オンライン動画マーケティング基礎)の3点が明記
   ×：1点以上不足または表記が曖昧

2. 受付時に確認が必要なアドレス相違を明記しているか
   〇：登録アドレスと問い合わせアドレスが異なる旨・受付時確認の必要性が分かる
   △：アドレス相違の言及のみで受付時確認の意図が弱い
   ×：まったく触れていない

3. 駐車場がない旨と近隣パーキング情報を共有しているか
   〇：駐車場がないこと・近隣コインパーキング利用案内の両方がある
   ×：どちらか一方のみ、またはどちらもない

4. 正しい終了予定時刻（16:30）を共有しているか
   〇：16:30と正しく明記
   ×：記載なしまたは誤り

5. セミナーとは直接関係ないが質問があった旨を共有しているか
   〇：助成金などセミナー本体とは別の質問があったことを共有
   △：質問があったことは分かるがやや曖昧
   ×：触れていない

6. 社内連絡として簡潔・必要十分にまとまっているか
   〇：過不足なく簡潔で現地対応に必要な情報が整理されている
   △：やや冗長だが実務上読める
   ×：冗長すぎる、または必要事項が埋もれている

7. 件名・挨拶・署名など社内向けメールの基本が整っているか
   〇：件名・宛名・挨拶・本文・署名（文末に部署名や名前）がすべてあり社内メールとして自然
   ×：一部不足または不自然（この項目は△を使わず〇か×のみで判定する）
   ※署名とはメール末尾に独立して置かれた部署名・氏名を指す。本文や挨拶文の中に部署名が出てくるだけでは署名とみなさない
   ※文末に署名が確認できない場合は、他の要素が揃っていても必ず×

採点時の重要ルール：
- 社内メールのため対外メールほど厳格な丁寧さは不要
- 本文に明示されている内容のみで採点（推測・補完禁止）
- 簡潔さが求められる課題のため、必要事項があっても冗長すぎる場合は△以下
`
  };

  return rubrics[questionId] || getGradingRubric();
}

// ===========================
// 回答の構造化（AI呼び出し）
// ===========================

function structureAnswer(rowData) {
  var name = getValueByHeader(rowData, CONFIG.FORM_HEADERS.candidateName) || '';
  var email = getValueByHeader(rowData, CONFIG.FORM_HEADERS.email) || '';
  var task1 = getValueByHeader(rowData, CONFIG.FORM_HEADERS.task1Answer) || '';
  var task2 = getValueByHeader(rowData, CONFIG.FORM_HEADERS.task2Answer) || '';

  var prompt = 'あなたはデータ整理担当です。\n' +
    '以下のGoogleフォーム回答を、設問ごとに整理してJSON形式で出力してください。\n\n' +
    '【重要ルール】\n' +
    '- 採点はしない\n' +
    '- 回答を勝手に補完しない\n' +
    '- 応募者の原文をできるだけ保持する\n' +
    '- 設問が判断できない場合は question_id を "unknown" にする\n' +
    '- JSON以外の文章を返さない\n\n' +
    '【フォーム回答データ】\n' +
    '応募者名：' + name + '\n' +
    'メールアドレス：' + email + '\n\n' +
    '設問1（Q1）：お客様へのメールを作成してください\n' +
    '回答：\n' + task1 + '\n\n' +
    '設問2（Q2）：現地対応スタッフへのメールを作成してください\n' +
    '回答：\n' + task2 + '\n\n' +
    '【出力形式（JSON）】\n' +
    '{"applicant_name":"応募者の氏名","email":"メールアドレス","answers":[{"question_id":"Q1","question_title":"お客様へのメールを作成してください","answer_text":"回答の原文","notes":""},{"question_id":"Q2","question_title":"現地対応スタッフへのメールを作成してください","answer_text":"回答の原文","notes":""}]}';

  var result = callOpenAIForJson(prompt);

  if (!result.answers || !Array.isArray(result.answers)) {
    throw new Error('回答の構造化に失敗しました：answers フィールドが不正です');
  }

  return result;
}

// ===========================
// 採点処理（AI呼び出し）
// ===========================

function gradeAnswers(structuredAnswer) {
  if (!structuredAnswer.answers || !Array.isArray(structuredAnswer.answers)) {
    throw new Error('構造化回答の形式が不正です');
  }

  var rubricText = structuredAnswer.answers.map(function(a) {
    return '=== ' + a.question_id + ': ' + a.question_title + ' ===\n' +
      getRubricByQuestionId(a.question_id);
  }).join('\n\n');

  var prompt = 'あなたは採用担当の採点者です。\n' +
    '採点基準に記載された【各項目を1つずつ】〇/△/×で採点してください。\n\n' +
    '【重要ルール】\n' +
    '- 採点基準の各項目を漏れなく採点する（項目を省略・まとめない）\n' +
    '- 採点はやや厳しめにする\n' +
    '- relevant_text には判定の根拠となる回答内の文章をそのまま抜き出す\n' +
    '  ・〇の場合：その評価の根拠となる回答内の文章\n' +
    '  ・△の場合：部分的に満たしている箇所、または問題のある箇所\n' +
    '  ・×の場合：「記載なし」またはNGとなる回答内の文章\n' +
    '- 回答に存在しない文章を relevant_text として捏造しない\n' +
    '- JSON以外の文章を返さない\n\n' +
    '【採点対象の回答】\n' +
    JSON.stringify(structuredAnswer.answers, null, 2) + '\n\n' +
    '【採点基準】\n' +
    rubricText + '\n\n' +
    '【スコア計算】\n' +
    '各項目：〇=1点、△=0.5点、×=0点\n' +
    '各設問のスコア = その設問の全項目の合計点\n' +
    'total_score = 全設問スコアの合計\n\n' +
    '【出力形式（JSON）】\n' +
    '{"results":[{"question_id":"Q1","question_title":"設問名","criteria":[{"no":1,"title":"採点項目のタイトル（採点基準と同じ文言）","judgement":"〇または△または×または判定不可","relevant_text":"回答内の該当箇所","reason":"判定理由"}],"score":設問の合計点,"overall_comment":"設問全体コメント"}],"total_score":全体合計点,"overall_judgement":"合格または保留または不合格","overall_comment":"総合コメント"}';

  var result = callOpenAIForJson(prompt);

  if (!result.results || !Array.isArray(result.results)) {
    throw new Error('採点処理に失敗しました：results フィールドが不正です');
  }

  return result;
}

// ===========================
// 既存採点ヘルパー関数
// ===========================

function buildAutoRejectResult(reason, itemNames) {
  return {
    overall_judge: '不合格',
    ng_count: 0,
    summary: reason,
    items: itemNames.map(function(name, index) {
      return { item_no: index + 1, item_name: name, evaluation: '判定不可', reason: '自動判定のため未採点' };
    })
  };
}

function buildNotScoredResult(reason, itemNames) {
  return {
    overall_judge: '判定不可',
    ng_count: 0,
    summary: reason,
    items: itemNames.map(function(name, index) {
      return { item_no: index + 1, item_name: name, evaluation: '判定不可', reason: reason };
    })
  };
}

function countXFromItems(items) {
  if (!items || !Array.isArray(items)) return 0;
  return items.filter(function(item) { return item && item.evaluation === '×'; }).length;
}

function scoreTask1(answerText) {
  var systemPrompt = '＃役割\nあなたは採用担当です。\n＃成果物\n課題テストの採点結果\n＃課題テストの指定のメールフッター\n================================================\n株式会社LEAGUE\n発行責任者：　武智翔太郎（株式会社LEAGUE 代表取締役）\n発行者住所：　発行元：〒 180-0004\n東京都武蔵野市吉祥寺本町2丁目8番4号　i-office吉祥寺\nお問い合わせ：info@league-jp.com\n株式会社LEAGUE：https://www.league-jp.com\n※このメールに覚えのない場合は、お手数ではございますが、\n上記問合せ先にその旨ご連絡ください。\n================================================\n＃採点項目\n以下の9項目について、〇 / △ / × / 判定不可で評価してください。\n1.日付と曜日の不一致を指摘し、正しい日程（3月13日（水））を案内できているか\n曜日の不一致を指摘し、正しい開催日を明記していれば〇\n軽い注意喚起のみは△\n触れていなければ×\n2.問い合わせアドレスと登録アドレスの違いを説明し、正しい受信先や迷惑メールフォルダの確認を促しているか\n登録アドレスと問い合わせ元アドレスが異なることに触れ、確認を促せていれば〇\n差異への言及はあるが、確認案内が弱い場合は△\n差異にまったく触れていなければ×\n3.入金完了・申込確定の事実を明確に伝えているか\n入金確認済み、申込確認済みが明確なら〇\n入金確認、申込確認のどちらかのみなら△\n触れていなければ×\n4.会場名・会場住所・アクセスURLのいずれかを案内し、あわせて「駐車場なし（近隣にコインパーキングあり）」を伝えられているか\nアクセスURLまたは会場情報があり、駐車場案内もあれば〇\nどちらか不足で△\n両方不足で×\n5.正しい終了予定時刻（16:30）を記載しているか\n正しく明記されていれば〇\n誤りは×\n6.「地域デジタル活性化助成金2024」は実在しない制度である前提に対し、慎重な対応ができているか\n制度名の確認を促す、正式情報を確認すると伝える、断定を避ける等なら〇\nやや曖昧だが断定案内を避けていれば△\n存在確認せず、それらしいリンクや案内を断定的に提示していれば×\n7.ビジネスメールとして敬語・表記・改行・余白が適切か\n・全体として自然で読みやすく、対外メールとして問題なく、改行や空行も適切で見やすければ〇\n・敬語や表記は概ね問題ないが、改行や空行が少なく、文章が詰まって見えて読みにくい場合は△\n・やや粗さがあるが読める場合も△\n・明らかに不自然、または改行不足により実務上かなり読みにくい場合は×\n※文章の正しさだけでなく、見た目の読みやすさも評価してください。\n※箇条書きや項目立てがあっても、1段落が長すぎる、空行が少ない、視認性が低い場合は△以下にしてください。\n8.要点が整理され、読みやすい構成になっているか\n問い合わせ内容ごとに整理されていれば〇\nやや冗長・混在なら△\n読みにくければ×\n9.メールフッターはあるか\n指定の署名があれば〇\n簡易署名のみの場合は×\n署名を変えすぎている場合は△\nない場合は×\n＃採点時の重要ルール\n設問で求められたことを満たしているかを最優先してください。\n細かい知識面の厳密さだけで過度に減点しないでください。\n一方で6は重要項目のため、存在しない制度を確認せず案内している場合は必ず×にしてください。\n採点結果は、各項目ごとに「評価（〇/△/×/判定不可）」「一言理由」を記載してください。\n候補者の文章全体が実務上送信可能かという観点も加味してください。\nMarkdown記法のリンク表記（例：こちら）がそのまま本文に含まれていても、実務メールとして大きな減点対象にはしないでください。必要に応じて表記の自然さとしてのみ軽く評価してください。\n細かいURLの貼り方や句読点だけで不合格にしないでください。\n敬語ミスだけではなく、改行・余白を含めた読みやすさも実務品質として評価してください。\n本文に明示されている内容のみで採点してください。推測や補完は禁止です。\n「触れている」「案内している」と評価する場合は、本文中の該当箇所をそのまま確認できることを前提としてください。\n該当記述が見当たらない場合は、「意図としては近い」と感じても×としてください。\n採点前に各項目について、本文の根拠箇所を確認してください。根拠がない場合は「記載なし」としてください。';
  return callOpenAIWithSchema(systemPrompt, answerText, 9);
}

function scoreTask2(answerText) {
  var systemPrompt = '＃役割\nあなたは採用担当です。\n＃成果物\n課題テストの採点結果\n＃採点項目\n以下の7項目について、〇 / △ / × / 判定不可 で評価してください。\n1.氏名・来場予定日・セミナー名を正確に記載しているか\n参加者氏名(松井 恵子)、来場予定日またはセミナー開催日（3月13日）、セミナー名(オンライン動画マーケティング基礎)の3点が明記されていれば〇\n1点以上不足、または表記がやや曖昧なら×\n2.受付時に確認が必要なアドレス相違を明記しているか\n登録アドレスと問い合わせアドレスが異なる旨、および受付時に確認が必要であることが分かるように書けていれば〇\nアドレス相違への言及のみで、受付時確認の意図が弱ければ△\nまったく触れていなければ×\n3.駐車場がない旨と近隣パーキング情報を共有しているか\n駐車場がないこと、および近隣コインパーキング利用案内の両方があれば〇\nどちらか一方のみなら×\nどちらもなければ×\n4.正しい終了予定時刻を共有しているか\n終了予定時刻が 16:30 と正しく明記されていれば〇\n終了時刻の記載がない、または誤っていれば×\n5.セミナーとは直接関係ないが質問があった旨を共有しているか\n助成金など、セミナー本体とは別の質問があったことを共有できていれば〇\n質問があったことは分かるが、やや曖昧なら△\n触れていなければ×\n6.社内連絡として簡潔・必要十分にまとまっているか\n社内メールとして過不足なく簡潔で、現地対応に必要な情報が整理されていれば〇\nやや冗長、または少し情報整理が甘いが実務上読めるなら△\n冗長すぎる、または必要事項が埋もれていて読みにくい場合は×\n7.件名・挨拶・署名など社内向けメールの基本が整っているか\n件名、宛名・挨拶、本文、締め、署名等全てあり、社内メールとして自然に整っていれば〇\n一部不足や不自然さがある場合は×（この項目は△を使わず〇か×のみで判定する）\n※署名とはメール末尾に独立して置かれた部署名・氏名を指す。本文や挨拶文の中に部署名が出てくるだけでは署名とみなさない\n※文末に署名が確認できない場合は、他の要素が揃っていても必ず×\n＃採点時の重要ルール\n設問で求められたことを満たしているかを最優先してください。\n細かい言い回しや敬語表現だけで過度に減点しないでください。\n本文に明示されている内容のみで採点してください。推測や補完は禁止です。\n「触れている」と評価する場合は、本文中の該当箇所をそのまま確認できることを前提としてください。\n該当記述が見当たらない場合は、「意図としては近い」と感じても×としてください。\n社内メールのため、対外メールほど厳格な丁寧さは不要ですが、最低限のビジネス文書として成立しているかは評価してください。\n簡潔さが求められる課題のため、必要事項が入っていても冗長すぎる場合は△以下を検討してください。\n採点前に各項目について、本文の根拠箇所を確認してください。根拠がない場合は「記載なし」としてください。';
  return callOpenAIWithSchema(systemPrompt, answerText, 7);
}

function buildCommentText(result) {
  return '総合判定: ' + (result.overall_judge || '') +
    ' / ×数: ' + (result.ng_count != null ? result.ng_count : '') +
    '\n要約: ' + (result.summary || '');
}

function buildDetailText(result) {
  var lines = ['【採点結果】'];
  if (result.items && Array.isArray(result.items)) {
    result.items.forEach(function(item) {
      lines.push(item.item_no + '. ' + item.item_name + '：' + item.evaluation);
    });
  }
  lines.push('');
  lines.push('【×数】' + (result.ng_count != null ? result.ng_count : ''));
  lines.push('【総合判定】' + (result.overall_judge || ''));
  lines.push('【要約】' + (result.summary || ''));
  return lines.join('\n');
}

function writeTask1Result(sheet, row, map, result) {
  setCellIfExists(sheet, row, map, CONFIG.COLUMN_NAMES.eJudge, result.overall_judge || '');
  setCellIfExists(sheet, row, map, CONFIG.COLUMN_NAMES.eNgCount, result.ng_count != null ? result.ng_count : '');
  setCellIfExists(sheet, row, map, CONFIG.COLUMN_NAMES.eComment, buildCommentText(result));
  setCellIfExists(sheet, row, map, CONFIG.COLUMN_NAMES.eDetail, buildDetailText(result));
}

function writeTask2Result(sheet, row, map, result) {
  setCellIfExists(sheet, row, map, CONFIG.COLUMN_NAMES.gJudge, result.overall_judge || '');
  setCellIfExists(sheet, row, map, CONFIG.COLUMN_NAMES.gNgCount, result.ng_count != null ? result.ng_count : '');
  setCellIfExists(sheet, row, map, CONFIG.COLUMN_NAMES.gComment, buildCommentText(result));
  setCellIfExists(sheet, row, map, CONFIG.COLUMN_NAMES.gDetail, buildDetailText(result));
}

function buildFinalJudge(task1Result, task2Result, aiUsageStatus) {
  if (aiUsageStatus === 'AI使用なし') {
    return { judge: '判定不可', comment: 'AI使用なしのため未採点' };
  }
  var t1 = task1Result ? task1Result.overall_judge : '判定不可';
  var t2 = task2Result ? task2Result.overall_judge : '判定不可';
  if (t1 === '不合格' || t2 === '不合格') return { judge: '不合格', comment: '課題1または課題2に不合格判定が含まれるため' };
  if (t1 === '保留' || t2 === '保留') return { judge: '保留', comment: '課題1または課題2に保留判定が含まれるため' };
  if (t1 === '合格' && t2 === '合格') return { judge: '合格', comment: '課題1・課題2ともに合格のため' };
  return { judge: '保留', comment: '一部判定不可または要確認項目があるため' };
}

// ===========================
// 採点結果タブの作成
// ===========================

function createApplicantResultSheet(applicantName) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var now = new Date();
  var timestamp = Utilities.formatDate(now, CONFIG.TIMEZONE, 'MMdd_HHmm');

  var cleanName = (applicantName || '名前不明')
    .replace(/[\\\/?*[\]:]/g, '')
    .replace(/\s+/g, '')
    .trim();

  var tabName = cleanName + '_採点_' + timestamp;
  if (tabName.length > 95) {
    tabName = tabName.substring(0, 95);
  }

  var finalName = tabName;
  var counter = 2;
  while (ss.getSheetByName(finalName)) {
    finalName = tabName + '_' + counter;
    counter++;
  }

  return ss.insertSheet(finalName);
}

// ===========================
// 採点結果タブへの出力
// ===========================

function writeApplicantResult(resultSheet, rowData, formHeaders, structuredAnswer, gradingResult, sourceRow) {
  var now = new Date();
  var nowStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm:ss');
  var r = 1;
  var MAX_COLS = 5;

  var titleBg      = '#1a73e8';
  var sectionBg    = '#4a86e8';
  var sectionFg    = '#ffffff';
  var colHeaderBg  = '#c9daf8';
  var infoBg       = '#f8f9fa';
  var q1GroupBg    = '#d0e0ff';
  var q2GroupBg    = '#d9f0d9';
  var okBg         = '#d9ead3';
  var deltaBg      = '#fff2cc';
  var ngBg         = '#fce8e6';

  function writeSectionTitle(title) {
    var range = resultSheet.getRange(r, 1, 1, MAX_COLS);
    range.merge();
    range.setValue(title);
    range.setBackground(sectionBg);
    range.setFontColor(sectionFg);
    range.setFontWeight('bold');
    range.setFontSize(11);
    r++;
  }

  function writeColHeader(values) {
    var range = resultSheet.getRange(r, 1, 1, values.length);
    range.setValues([values]);
    range.setBackground(colHeaderBg);
    range.setFontWeight('bold');
    r++;
  }

  function writeInfoRow(label, value) {
    resultSheet.getRange(r, 1).setValue(label);
    resultSheet.getRange(r, 1).setFontWeight('bold');
    resultSheet.getRange(r, 2).setValue(value !== null && value !== undefined ? value : '');
    resultSheet.getRange(r, 1, 1, MAX_COLS).setBackground(infoBg);
    r++;
  }

  function writeBlank() { r++; }

  // ========================
  // 1. タイトル
  // ========================
  var titleRange = resultSheet.getRange(r, 1, 1, MAX_COLS);
  titleRange.merge();
  titleRange.setValue('応募者 採点結果');
  titleRange.setBackground(titleBg);
  titleRange.setFontColor('#ffffff');
  titleRange.setFontWeight('bold');
  titleRange.setFontSize(14);
  titleRange.setHorizontalAlignment('center');
  r++;
  writeBlank();

  // ========================
  // 2. 応募者情報
  // ========================
  writeSectionTitle('【応募者情報】');

  var applicantName  = getValueByHeader(rowData, CONFIG.FORM_HEADERS.candidateName) || '';
  var applicantEmail = getValueByHeader(rowData, CONFIG.FORM_HEADERS.email) || '';
  var tsRaw = getValueByHeader(rowData, CONFIG.TIMESTAMP_HEADER);
  var tsStr = tsRaw instanceof Date
    ? Utilities.formatDate(tsRaw, CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm:ss')
    : String(tsRaw || '');

  writeInfoRow('氏名', applicantName);
  writeInfoRow('メールアドレス', applicantEmail);
  writeInfoRow('回答日時', tsStr);
  writeInfoRow('採点日時', nowStr);
  writeInfoRow('元シート行番号', sourceRow || '');
  writeBlank();

  // ========================
  // 3. 応募者回答全文
  // ========================
  writeSectionTitle('【応募者回答全文】');

  var answers = (structuredAnswer && structuredAnswer.answers) ? structuredAnswer.answers : [];
  answers.forEach(function(ans) {
    // 設問ラベル行
    var labelRange = resultSheet.getRange(r, 1, 1, MAX_COLS);
    labelRange.merge();
    labelRange.setValue('▼ ' + ans.question_id + '：' + ans.question_title);
    labelRange.setFontWeight('bold');
    labelRange.setBackground(ans.question_id === 'Q1' ? q1GroupBg : q2GroupBg);
    r++;

    // 回答テキスト行
    var ansRange = resultSheet.getRange(r, 1, 1, MAX_COLS);
    ansRange.merge();
    ansRange.setValue(ans.answer_text || '（回答なし）');
    ansRange.setWrap(true);
    ansRange.setVerticalAlignment('top');
    resultSheet.setRowHeight(r, 200);
    r++;
    writeBlank();
  });

  // ========================
  // 4. 採点結果（全設問・全項目）
  // ========================
  writeSectionTitle('【採点結果】');

  var results = (gradingResult && gradingResult.results) ? gradingResult.results : [];

  results.forEach(function(res) {
    // 設問グループヘッダー
    var groupBg = res.question_id === 'Q1' ? q1GroupBg : q2GroupBg;
    var groupRange = resultSheet.getRange(r, 1, 1, MAX_COLS);
    groupRange.merge();
    groupRange.setValue('● ' + res.question_id + '：' + (res.question_title || ''));
    groupRange.setFontWeight('bold');
    groupRange.setBackground(groupBg);
    r++;

    // 列ヘッダー
    writeColHeader(['No', '採点項目', '判定', '判定理由', '該当箇所（回答内の根拠）']);

    // 各採点項目
    var criteria = res.criteria || [];
    criteria.forEach(function(c) {
      var judgement = c.judgement || '';
      var rowBg = '';
      if (judgement === '〇') rowBg = okBg;
      else if (judgement === '△') rowBg = deltaBg;
      else if (judgement === '×' || judgement === 'x' || judgement === 'X') rowBg = ngBg;

      resultSheet.getRange(r, 1, 1, MAX_COLS).setValues([[
        c.no || '',
        c.title || '',
        judgement,
        c.reason || '',
        c.relevant_text || ''
      ]]);

      if (rowBg) {
        resultSheet.getRange(r, 1, 1, MAX_COLS).setBackground(rowBg);
      }

      // × の判定列セルをさらに強調
      if (judgement === '×') {
        resultSheet.getRange(r, 3).setFontWeight('bold').setFontColor('#cc0000');
      }

      r++;
    });

    // 小計行
    var subtotalRange = resultSheet.getRange(r, 1, 1, MAX_COLS);
    subtotalRange.merge();
    subtotalRange.setValue(res.question_id + ' 小計：' + (res.score != null ? res.score : '-') + '点　' + (res.overall_comment || ''));
    subtotalRange.setFontWeight('bold');
    subtotalRange.setBackground(groupBg);
    r++;
    writeBlank();
  });

  // ========================
  // 5. 総合判定
  // ========================
  writeSectionTitle('【総合判定】');

  var totalScore   = gradingResult ? (gradingResult.total_score != null ? gradingResult.total_score : '') : '';
  var overallJudge = gradingResult ? (gradingResult.overall_judgement || '') : '';
  var overallComment = gradingResult ? (gradingResult.overall_comment || '') : '';

  writeInfoRow('合計点', totalScore);
  writeInfoRow('総合判定', overallJudge);

  // 総合判定行の色分け
  var judgeRow = r - 1;
  if (overallJudge === '合格') {
    resultSheet.getRange(judgeRow, 1, 1, MAX_COLS).setBackground(okBg);
  } else if (overallJudge === '不合格') {
    resultSheet.getRange(judgeRow, 1, 1, MAX_COLS).setBackground(ngBg);
  } else {
    resultSheet.getRange(judgeRow, 1, 1, MAX_COLS).setBackground(deltaBg);
  }

  writeInfoRow('総合コメント', overallComment);
  writeBlank();

  // ========================
  // 6. AI出力の生データ（デバッグ用）
  // ========================
  writeSectionTitle('【AI出力の生データ（デバッグ用）】');
  resultSheet.getRange(r, 1).setValue('構造化JSON');
  resultSheet.getRange(r, 1).setFontWeight('bold');
  resultSheet.getRange(r, 2).setValue(JSON.stringify(structuredAnswer, null, 2));
  r++;
  resultSheet.getRange(r, 1).setValue('採点JSON');
  resultSheet.getRange(r, 1).setFontWeight('bold');
  resultSheet.getRange(r, 2).setValue(JSON.stringify(gradingResult, null, 2));
  r++;

  // ========================
  // 書式設定
  // ========================
  var lastRow = r - 1;

  resultSheet.getRange(1, 1, lastRow, MAX_COLS)
    .setWrap(true)
    .setVerticalAlignment('top');

  // 列幅：No / 採点項目 / 判定 / 判定理由 / 該当箇所
  resultSheet.setColumnWidth(1, 45);
  resultSheet.setColumnWidth(2, 300);
  resultSheet.setColumnWidth(3, 70);
  resultSheet.setColumnWidth(4, 320);
  resultSheet.setColumnWidth(5, 350);

  resultSheet.setFrozenRows(1);

  SpreadsheetApp.flush();
}

// ===========================
// Slack通知（既存）
// ===========================

function sendSlackChannelMessage(answerObject, task1Result, task2Result, aiUsageStatus) {
  var channelId = CONFIG.SLACK_CHANNEL_ID;
  if (!channelId) throw new Error('SLACK_CHANNEL_ID が未設定です');

  var mentionUserId = CONFIG.SLACK_MENTION_USER_ID;
  var candidateName =
    getValueByHeader(answerObject, CONFIG.FORM_HEADERS.candidateName) ||
    getValueByHeader(answerObject, '氏名') ||
    getValueByHeader(answerObject, '名前') ||
    '氏名不明';

  var task1Ng = task1Result && task1Result.ng_count != null ? task1Result.ng_count : '-';
  var task2Ng = task2Result && task2Result.ng_count != null ? task2Result.ng_count : '-';

  var text = '';
  if (mentionUserId) text += '<@' + mentionUserId + '>\n';
  text += '【課題の提出がありました】\n';
  text += '名前: ' + candidateName + '\n';
  text += 'AI使用: ' + aiUsageStatus + '\n';
  if (aiUsageStatus === 'AI使用あり') {
    text += '課題1 ×数: ' + task1Ng + '\n';
    text += '課題2 ×数: ' + task2Ng + '\n';
  }

  var res = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + getSlackToken() },
    payload: JSON.stringify({ channel: channelId, text: text }),
    muteHttpExceptions: true
  });

  var statusCode = res.getResponseCode();
  var responseText = res.getContentText();
  var data = JSON.parse(responseText);

  if (statusCode !== 200 || !data.ok) {
    throw new Error('Slack API error: ' + statusCode + ' / ' + responseText);
  }
}

// ===========================
// 手動採点メイン関数
// ===========================

function gradeSelectedApplicant() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getActiveSheet();

  // シート名確認
  if (sheet.getName() !== CONFIG.FORM_SHEET_NAME) {
    ui.alert('エラー', 'フォーム回答シート「' + CONFIG.FORM_SHEET_NAME + '」で実行してください。\n現在のシート：' + sheet.getName(), ui.ButtonSet.OK);
    return;
  }

  // 選択行確認
  var selection = sheet.getActiveRange();
  if (selection.getNumRows() > 1) {
    ui.alert('エラー', '複数行が選択されています。採点したい応募者の行を1行だけ選択してください。', ui.ButtonSet.OK);
    return;
  }

  var row = selection.getRow();
  if (row <= CONFIG.HEADER_ROW) {
    ui.alert('エラー', 'ヘッダー行が選択されています。回答が入力されている行を選択してください。', ui.ButtonSet.OK);
    return;
  }

  // 行データ取得
  var lastColumn = sheet.getLastColumn();
  var formHeaders = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastColumn).getValues()[0];
  var map = createHeaderMap(formHeaders);
  var rowValues = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
  var answerObj = buildAnswerObject(formHeaders, rowValues);

  // 応募者名確認
  var applicantName = getValueByHeader(answerObj, CONFIG.FORM_HEADERS.candidateName) || '';
  if (!applicantName) {
    ui.alert('エラー', '応募者名が取得できません。正しい行を選択しているか確認してください。\n（列名：' + CONFIG.FORM_HEADERS.candidateName + '）', ui.ButtonSet.OK);
    return;
  }

  // 回答確認
  var task1 = getValueByHeader(answerObj, CONFIG.FORM_HEADERS.task1Answer);
  var task2 = getValueByHeader(answerObj, CONFIG.FORM_HEADERS.task2Answer);
  if (isBlank(task1) && isBlank(task2)) {
    ui.alert('エラー', '採点対象の回答が空です。回答が入力されている行を選択してください。', ui.ButtonSet.OK);
    return;
  }

  // APIキー確認
  try { getOpenAIKey(); } catch (e) {
    ui.alert('エラー', e.message + '\n\nスクリプトプロパティに OPENAI_API_KEY を設定してください。', ui.ButtonSet.OK);
    return;
  }

  // ステータス列の確保
  ensureStatusColumns(sheet);

  // ヘッダーとマップを再取得（列が追加された可能性があるため）
  lastColumn = sheet.getLastColumn();
  formHeaders = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastColumn).getValues()[0];
  map = createHeaderMap(formHeaders);
  rowValues = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
  answerObj = buildAnswerObject(formHeaders, rowValues);

  // 二重採点防止
  var currentStatus = '';
  var statusColNorm = normalizeHeaderName(CONFIG.STATUS_HEADER);
  if (map[statusColNorm]) {
    currentStatus = String(sheet.getRange(row, map[statusColNorm]).getValue() || '');
  }
  if (currentStatus === '採点済') {
    var confirm = ui.alert('確認', 'この応募者はすでに採点済です。再採点しますか？', ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) return;
  }

  var resultSheet = null;

  try {
    // 採点中ステータスを記録
    setCellIfExists(sheet, row, map, CONFIG.STATUS_HEADER, '採点中');
    SpreadsheetApp.flush();

    // 結果タブ作成
    resultSheet = createApplicantResultSheet(applicantName);

    // 回答の構造化（AI呼び出し1）
    var structuredAnswer = structureAnswer(answerObj);

    // 採点（AI呼び出し2）
    var gradingResult = gradeAnswers(structuredAnswer);

    // 時間基準の採点（GAS側でフォームデータから直接判定）
    gradingResult = evaluateTimeCriteria(answerObj, gradingResult);

    // タブに結果出力
    writeApplicantResult(resultSheet, answerObj, formHeaders, structuredAnswer, gradingResult, row);

    // タブURL
    var tabUrl = ss.getUrl() + '#gid=' + resultSheet.getSheetId();

    // 元シートに記録
    var nowStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm:ss');
    setCellIfExists(sheet, row, map, CONFIG.STATUS_HEADER, '採点済');
    setCellIfExists(sheet, row, map, CONFIG.GRADED_AT_HEADER, nowStr);
    setCellIfExists(sheet, row, map, CONFIG.RESULT_URL_HEADER, tabUrl);
    setCellIfExists(sheet, row, map, CONFIG.ERROR_HEADER, '');
    SpreadsheetApp.flush();

    ui.alert('完了', applicantName + ' さんの採点が完了しました。\n採点結果タブが作成されました。', ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('gradeSelectedApplicant error: ' + error.message);

    // エラー情報を元シートに記録
    var errNowStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm:ss');
    setCellIfExists(sheet, row, map, CONFIG.STATUS_HEADER, '採点エラー');
    setCellIfExists(sheet, row, map, CONFIG.GRADED_AT_HEADER, errNowStr);
    setCellIfExists(sheet, row, map, CONFIG.ERROR_HEADER, error.message);

    if (resultSheet) {
      var errTabUrl = ss.getUrl() + '#gid=' + resultSheet.getSheetId();
      setCellIfExists(sheet, row, map, CONFIG.RESULT_URL_HEADER, errTabUrl);
    }
    SpreadsheetApp.flush();

    ui.alert('エラー', '採点中にエラーが発生しました。\n\n' + error.message, ui.ButtonSet.OK);
  }
}

// ===========================
// フォーム自動採点（既存）
// ===========================

function onFormSubmit(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.SHEET_NAME) return;

    var row = e.range.getRow();
    var lastColumn = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    var map = createHeaderMap(headers);
    var rowValues = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
    var answerObject = buildAnswerObject(headers, rowValues);

    getOrCreateCandidateId(sheet, row, map);
    setCellIfExists(sheet, row, map, CONFIG.COLUMN_NAMES.status, 'AI採点中');

    var eAnswer = getValueByHeader(answerObject, CONFIG.FORM_HEADERS.task1Answer);
    var gAnswer = getValueByHeader(answerObject, CONFIG.FORM_HEADERS.task2Answer);
    var aiLink = getValueByHeader(answerObject, CONFIG.FORM_HEADERS.aiLink);
    var aiLog = getValueByHeader(answerObject, CONFIG.FORM_HEADERS.aiLog);

    var task1ItemNames = [
      '日付と曜日の不一致を指摘し、正しい日程（3月13日（水））を案内できているか',
      '問い合わせアドレスと登録アドレスの違いを説明し、確認を促しているか',
      '入金完了・申込確定の事実を明確に伝えているか',
      '会場名・会場住所・アクセスURLのいずれかと駐車場なしを伝えているか',
      '正しい終了予定時刻（16:30）を記載しているか',
      '存在しない制度に対し慎重な対応ができているか',
      'ビジネスメールとして敬語・表記・改行・余白が適切か',
      '要点が整理され、読みやすい構成になっているか',
      'メールフッターはあるか'
    ];

    var task2ItemNames = [
      '氏名・来場予定日・セミナー名を正確に記載しているか',
      '受付時に確認が必要なアドレス相違を明記しているか',
      '駐車場がない旨と近隣パーキング情報を共有しているか',
      '正しい終了予定時刻を共有しているか',
      'セミナーとは直接関係ないが質問があった旨を共有しているか',
      '社内連絡として簡潔・必要十分にまとまっているか',
      '件名・挨拶・署名など社内向けメールの基本が整っているか'
    ];

    var task1Result = null;
    var task2Result = null;
    var aiUsageStatus = '';

    if (isBlank(aiLink) && isBlank(aiLog)) {
      aiUsageStatus = 'AI使用なし';
      task1Result = buildNotScoredResult('AI使用なしのため未採点', task1ItemNames);
      task2Result = buildNotScoredResult('AI使用なしのため未採点', task2ItemNames);
    } else {
      aiUsageStatus = 'AI使用あり';
      task1Result = !isBlank(eAnswer)
        ? scoreTask1(eAnswer)
        : buildAutoRejectResult('課題1の回答が空欄のため採点不可', task1ItemNames);
      task2Result = !isBlank(gAnswer)
        ? scoreTask2(gAnswer)
        : buildAutoRejectResult('課題2の回答が空欄のため採点不可', task2ItemNames);
    }

    writeTask1Result(sheet, row, map, task1Result);
    writeTask2Result(sheet, row, map, task2Result);

    var finalResult = buildFinalJudge(task1Result, task2Result, aiUsageStatus);
    setCellIfExists(sheet, row, map, CONFIG.COLUMN_NAMES.finalJudge, finalResult.judge);
    setCellIfExists(sheet, row, map, CONFIG.COLUMN_NAMES.finalComment, finalResult.comment);
    setCellIfExists(sheet, row, map, CONFIG.COLUMN_NAMES.evaluatedAt, new Date());
    setCellIfExists(sheet, row, map, CONFIG.COLUMN_NAMES.status, '完了');

    sendSlackChannelMessage(answerObject, task1Result, task2Result, aiUsageStatus);

  } catch (error) {
    try {
      var errSheet = e && e.range ? e.range.getSheet() : null;
      if (errSheet) {
        var errHeaders = errSheet.getRange(1, 1, 1, errSheet.getLastColumn()).getValues()[0];
        var errMap = createHeaderMap(errHeaders);
        var errRow = e.range.getRow();
        setCellIfExists(errSheet, errRow, errMap, CONFIG.COLUMN_NAMES.status, 'APIエラー');
        setCellIfExists(errSheet, errRow, errMap, CONFIG.COLUMN_NAMES.finalComment, error.message);
        setCellIfExists(errSheet, errRow, errMap, CONFIG.COLUMN_NAMES.evaluatedAt, new Date());
      }
    } catch (err) {
      Logger.log('ステータス更新失敗: ' + err);
    }
    Logger.log('onFormSubmit error: ' + error.message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

// ===========================
// カスタムメニュー
// ===========================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI採点')
    .addItem('選択行を採点する', 'gradeSelectedApplicant')
    .addToUi();
}

// ===========================
// デバッグ関数
// ===========================

function debugHeaders() {
  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  headers.forEach(function(h, i) {
    Logger.log((i + 1) + '列目 = [' + String(h) + ']');
  });
}

function debugColumnMatch() {
  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = createHeaderMap(headers);

  var targets = [
    CONFIG.COLUMN_NAMES.candidateId,
    CONFIG.COLUMN_NAMES.status,
    CONFIG.COLUMN_NAMES.eJudge,
    CONFIG.COLUMN_NAMES.eNgCount,
    CONFIG.COLUMN_NAMES.gJudge,
    CONFIG.COLUMN_NAMES.gNgCount,
    CONFIG.COLUMN_NAMES.finalJudge,
    CONFIG.COLUMN_NAMES.finalComment,
    CONFIG.COLUMN_NAMES.evaluatedAt,
    CONFIG.FORM_HEADERS.task1Answer,
    CONFIG.FORM_HEADERS.task2Answer,
    CONFIG.FORM_HEADERS.aiLink,
    CONFIG.FORM_HEADERS.aiLog,
    CONFIG.FORM_HEADERS.candidateName
  ];

  targets.forEach(function(name) {
    var normalized = normalizeHeaderName(name);
    Logger.log(name + ' => ' + (map[normalized] || '見つからない'));
  });
}
