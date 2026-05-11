import OpenAI from 'openai';
import type {
  ApplicantRow,
  GradingResult,
  JobType,
  QuestionResult,
  StructuredAnswer,
} from './types';

function getOpenAI() {
  return new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
}

function getModel() {
  return process.env.OPENAI_MODEL ?? 'gpt-4.1';
}

// ===========================
// OpenAI ユーティリティ
// ===========================

async function callOpenAIForJson(prompt: string): Promise<Record<string, unknown>> {
  const res = await getOpenAI().chat.completions.create({
    model: getModel(),
    messages: [{ role: 'user', content: prompt }],
    response_format: { type: 'json_object' },
  });
  const text = res.choices[0].message.content ?? '{}';
  return JSON.parse(text) as Record<string, unknown>;
}

// ===========================
// 時間パース
// ===========================

function parseMinutes(value: string | null | undefined): number | null {
  if (!value) return null;
  let s = String(value).trim();
  if (!s) return null;

  s = s.replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));
  s = s.replace(/：/g, ':');

  const mSec = s.match(/^(\d+)\s*分\s*(\d+)\s*秒/);
  if (mSec) return parseInt(mSec[1]) + parseInt(mSec[2]) / 60;

  const mColon = s.match(/^(\d+):(\d{1,2})$/);
  if (mColon) return parseInt(mColon[1]) + parseInt(mColon[2]) / 60;

  const mNum = s.match(/^([\d.]+)/);
  if (mNum) return parseFloat(mNum[1]);

  return null;
}

// ===========================
// 採点ルーブリック
// ===========================

const Q1_RUBRIC = `
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
   ×：存在確認をせず、実在する制度のように案内している

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
`;

const Q2_RUBRIC = `
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
   〇：件名・宛名・挨拶・本文・署名がすべてあり社内メールとして自然
   ×：一部不足または不自然（この項目は△を使わず〇か×のみ）
   ※文末に署名が確認できない場合は必ず×

採点時の重要ルール：
- 社内メールのため対外メールほど厳格な丁寧さは不要
- 本文に明示されている内容のみで採点（推測・補完禁止）
- 簡潔さが求められる課題のため、必要事項があっても冗長すぎる場合は△以下
`;

function getRubric(questionId: string): string {
  if (questionId === 'Q1') return Q1_RUBRIC;
  if (questionId === 'Q2') return Q2_RUBRIC;
  return '';
}

// ===========================
// 回答の構造化
// ===========================

async function structureAnswer(applicant: ApplicantRow): Promise<StructuredAnswer> {
  const prompt =
    'あなたはデータ整理担当です。\n' +
    '以下のGoogleフォーム回答を、設問ごとに整理してJSON形式で出力してください。\n\n' +
    '【重要ルール】\n' +
    '- 採点はしない\n' +
    '- 回答を勝手に補完しない\n' +
    '- 応募者の原文をできるだけ保持する\n' +
    '- 設問が判断できない場合は question_id を "unknown" にする\n' +
    '- JSON以外の文章を返さない\n\n' +
    '【フォーム回答データ】\n' +
    '応募者名：' + applicant.name + '\n' +
    'メールアドレス：' + applicant.email + '\n\n' +
    '設問1（Q1）：お客様へのメールを作成してください\n' +
    '回答：\n' + (applicant.task1Answer || '（回答なし）') + '\n\n' +
    '設問2（Q2）：現地対応スタッフへのメールを作成してください\n' +
    '回答：\n' + (applicant.task2Answer || '（回答なし）') + '\n\n' +
    '【出力形式（JSON）】\n' +
    '{"applicant_name":"応募者の氏名","email":"メールアドレス","answers":[' +
    '{"question_id":"Q1","question_title":"お客様へのメールを作成してください","answer_text":"回答の原文","notes":""},' +
    '{"question_id":"Q2","question_title":"現地対応スタッフへのメールを作成してください","answer_text":"回答の原文","notes":""}' +
    ']}';

  const raw = await callOpenAIForJson(prompt);
  if (!raw.answers || !Array.isArray(raw.answers)) {
    throw new Error('回答の構造化に失敗しました（answers フィールドが不正）');
  }
  return raw as unknown as StructuredAnswer;
}

// ===========================
// 採点
// ===========================

async function gradeAnswers(structured: StructuredAnswer): Promise<GradingResult> {
  const rubricText = structured.answers
    .map((a) => `=== ${a.question_id}: ${a.question_title} ===\n${getRubric(a.question_id)}`)
    .join('\n\n');

  const prompt =
    'あなたは採用担当の採点者です。\n' +
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
    JSON.stringify(structured.answers, null, 2) + '\n\n' +
    '【採点基準】\n' +
    rubricText + '\n\n' +
    '【スコア計算】\n' +
    '各項目：〇=1点、△=0.5点、×=0点\n' +
    '各設問のスコア = その設問の全項目の合計点\n' +
    'total_score = 全設問スコアの合計\n\n' +
    '【出力形式（JSON）】\n' +
    '{"results":[{"question_id":"Q1","question_title":"設問名","criteria":[{"no":1,"title":"採点項目のタイトル","judgement":"〇または△または×または判定不可","relevant_text":"回答内の該当箇所","reason":"判定理由"}],"score":設問の合計点,"overall_comment":"設問全体コメント"}],"total_score":全体合計点,"overall_judgement":"合格または保留または不合格","overall_comment":"総合コメント"}';

  const raw = await callOpenAIForJson(prompt);
  if (!raw.results || !Array.isArray(raw.results)) {
    throw new Error('採点処理に失敗しました（results フィールドが不正）');
  }
  return raw as unknown as GradingResult;
}

// ===========================
// 時間基準の採点（ロジック側で判定）
// ===========================

function evaluateTimeCriteria(applicant: ApplicantRow, result: GradingResult): GradingResult {
  const q1 = result.results.find((r: QuestionResult) => r.question_id === 'Q1');
  const q2 = result.results.find((r: QuestionResult) => r.question_id === 'Q2');

  // Q1：設問理解 + メール作成 合計 < 16分
  const readingMin = parseMinutes(applicant.readingTime);
  const task1Min   = parseMinutes(applicant.task1Time);
  if (q1) {
    if (readingMin !== null && task1Min !== null) {
      const total   = Math.round((readingMin + task1Min) * 100) / 100;
      const judge   = total >= 16 ? '×' : '〇';
      const nextNo  = (q1.criteria?.length ?? 0) + 1;
      q1.criteria   = q1.criteria ?? [];
      q1.criteria.push({
        no: nextNo,
        title: '15分以内に作成できているか（設問理解 + メール作成の合計時間）',
        judgement: judge,
        relevant_text: `設問理解：${applicant.readingTime}、メール作成：${applicant.task1Time}（合計 ${total}分）`,
        reason: `合計 ${total}分 ／ 基準：16分未満で〇`,
      });
      if (judge === '〇') {
        q1.score = (q1.score ?? 0) + 1;
        result.total_score = (result.total_score ?? 0) + 1;
      }
    } else {
      q1.criteria = q1.criteria ?? [];
      q1.criteria.push({
        no: (q1.criteria.length) + 1,
        title: '15分以内に作成できているか（設問理解 + メール作成の合計時間）',
        judgement: '判定不可',
        relevant_text: `設問理解：${applicant.readingTime || '未入力'}、メール作成：${applicant.task1Time || '未入力'}`,
        reason: '時間データが未入力のため判定不可',
      });
    }
  }

  // Q2：メール作成時間 < 11分
  const task2Min = parseMinutes(applicant.task2Time);
  if (q2) {
    if (task2Min !== null) {
      const judge  = task2Min >= 11 ? '×' : '〇';
      const nextNo = (q2.criteria?.length ?? 0) + 1;
      q2.criteria  = q2.criteria ?? [];
      q2.criteria.push({
        no: nextNo,
        title: '10分以内に作成できているか（メール作成時間）',
        judgement: judge,
        relevant_text: `メール作成時間：${applicant.task2Time}（${task2Min}分）`,
        reason: `${task2Min}分 ／ 基準：11分未満で〇`,
      });
      if (judge === '〇') {
        q2.score = (q2.score ?? 0) + 1;
        result.total_score = (result.total_score ?? 0) + 1;
      }
    } else {
      q2.criteria = q2.criteria ?? [];
      q2.criteria.push({
        no: (q2.criteria.length) + 1,
        title: '10分以内に作成できているか（メール作成時間）',
        judgement: '判定不可',
        relevant_text: `メール作成時間：${applicant.task2Time || '未入力'}`,
        reason: '時間データが未入力のため判定不可',
      });
    }
  }

  return result;
}

// ===========================
// NG閾値による合否判定
// ===========================

function applyNgThreshold(result: GradingResult, jobType: JobType): GradingResult {
  const totalNg = result.results?.reduce(
    (sum, qr) => sum + (qr.criteria?.filter((c) => c.judgement === '×').length ?? 0), 0
  ) ?? 0;
  const threshold = jobType === '秘書' ? 6 : 7;
  result.overall_judgement = totalNg >= threshold ? '不合格' : '合格';
  return result;
}

// ===========================
// メイン採点関数
// ===========================

export async function gradeApplicant(applicant: ApplicantRow, jobType: JobType): Promise<GradingResult> {
  const structured = await structureAnswer(applicant);
  let result       = await gradeAnswers(structured);
  result           = evaluateTimeCriteria(applicant, result);
  result           = applyNgThreshold(result, jobType);
  result.structured_answer = structured;
  return result;
}
