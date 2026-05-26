import { google } from 'googleapis';
import type { ApplicantRow } from './types';

const SPREADSHEET_ID = process.env.SPREADSHEET_ID!;
const SHEET_NAME = 'フォームの回答 1';

const COL = {
  TIMESTAMP:    'タイムスタンプ',
  NAME:         '氏名（応募時のお名前）',
  EMAIL:        'メールアドレス（応募時に記載したもの）',
  TASK1:        '【課題テストの回答】１．お客様へのメールを作成してください',
  TASK2:        '【課題テストの回答】２． 現地対応スタッフへのメールを作成してください。',
  AI_LINK:      'AIツールを使用した方は共有リンクを貼り付けてください。AI自体のURLではなく、“共有URL”を作成してください。',
  AI_LOG:       'AIツールを使用した方は、やりとり全文をWordファイルにコピー＆ペーストしてください。そのファイルをギガファイル便を利用してご提出ください。※こちらにギガファイル便のリンクを貼ってください。',
  READING_TIME:       'タイマーをストップし、設問を読んで理解した時間を入力してください（分）',
  TASK1_TIME:         'タイマーをストップし、1のメール作成にかかった時間を入力してください（分）',
  TASK2_TIME:         'タイマーをストップし、２のメール作成にかかった時間を入力してください（分）',
  DESIRED_WORK_HOURS: '採用時の希望勤務時間（週○日、1日○時〜○時の何時間、勤務可能な曜日 ※土日歓迎！）',
  FINAL_JUDGE:  '合否',
  NOTES:        '備考',
  STATUS:       '採点ステータス',
  GRADED_AT:    '採点日時',
  ERROR:        '採点エラー',
  RESULT_JSON:  '採点結果JSON',
  PROMPT_CHECKS: 'プロンプト評価',
  JOB_TYPE:      '職種',
};

function normalize(h: unknown): string {
  return String(h ?? '')
    .replace(/\r?\n/g, ' ')
    .replace(/　/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function colLetter(n: number): string {
  let r = '';
  while (n > 0) {
    r = String.fromCharCode(65 + ((n - 1) % 26)) + r;
    n = Math.floor((n - 1) / 26);
  }
  return r;
}

function getAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON!);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

async function getSheetsClient() {
  const auth = getAuth();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return google.sheets({ version: 'v4', auth: (await auth.getClient()) as any });
}

// シートの列数が足りない場合に拡張する
async function expandColumnsIfNeeded(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  sheets: any,
  neededColumns: number,
): Promise<void> {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets(properties(sheetId,title,gridProperties))',
  });

  const sheetMeta = meta.data.sheets?.find(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (s: any) => s.properties?.title === SHEET_NAME,
  );
  if (!sheetMeta) return;

  const currentCols = sheetMeta.properties?.gridProperties?.columnCount ?? 0;
  if (neededColumns <= currentCols) return;

  const addCount = neededColumns - currentCols + 5; // 余裕を持って追加
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [{
        appendDimension: {
          sheetId: sheetMeta.properties.sheetId,
          dimension: 'COLUMNS',
          length: addCount,
        },
      }],
    },
  });
}

export async function getApplicants(): Promise<ApplicantRow[]> {
  const sheets = await getSheetsClient();

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'`,
  });

  const rows = res.data.values ?? [];
  if (rows.length < 2) return [];

  const headers = rows[0].map(normalize);
  const idx = (name: string) => headers.indexOf(normalize(name));

  return rows.slice(1)
    .map((row, i) => {
      const get = (name: string) => {
        const j = idx(name);
        return j >= 0 && j < row.length ? String(row[j] ?? '') : '';
      };
      return {
        row: i + 2,
        timestamp:         get(COL.TIMESTAMP),
        name:              get(COL.NAME),
        email:             get(COL.EMAIL),
        task1Answer:       get(COL.TASK1),
        task2Answer:       get(COL.TASK2),
        aiLink:            get(COL.AI_LINK),
        aiLog:             get(COL.AI_LOG),
        readingTime:       get(COL.READING_TIME),
        task1Time:         get(COL.TASK1_TIME),
        task2Time:         get(COL.TASK2_TIME),
        desiredWorkHours:  get(COL.DESIRED_WORK_HOURS),
        finalJudge:        get(COL.FINAL_JUDGE),
        notes:             get(COL.NOTES),
        gradingStatus:     get(COL.STATUS),
        gradedAt:          get(COL.GRADED_AT),
        gradingResultJson: get(COL.RESULT_JSON),
        gradingError:      get(COL.ERROR),
        promptChecks:      get(COL.PROMPT_CHECKS),
        jobType:           get(COL.JOB_TYPE),
      } satisfies ApplicantRow;
    })
    .filter((a) => a.name.trim() !== '');
}

export async function getApplicantByRow(rowNumber: number): Promise<ApplicantRow | null> {
  const all = await getApplicants();
  return all.find((a) => a.row === rowNumber) ?? null;
}

export async function writeGradingResult(
  rowNumber: number,
  result: object | null,
  error?: string,
  finalJudge?: string,
): Promise<void> {
  const sheets = await getSheetsClient();

  const headersRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!1:1`,
  });
  const headers = (headersRes.data.values?.[0] ?? []).map(String);

  // Ensure a column exists; returns its 0-based index
  const ensureCol = async (name: string): Promise<number> => {
    const normed = normalize(name);
    let i = headers.findIndex((h) => normalize(h) === normed);
    if (i < 0) {
      i = headers.length;
      headers.push(name);
      // 列数が上限を超える場合は先に拡張
      await expandColumnsIfNeeded(sheets, i + 1);
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${SHEET_NAME}'!${colLetter(i + 1)}1`,
        valueInputOption: 'RAW',
        requestBody: { values: [[name]] },
      });
    }
    return i;
  };

  const statusIdx      = await ensureCol(COL.STATUS);
  const dateIdx        = await ensureCol(COL.GRADED_AT);
  const jsonIdx        = await ensureCol(COL.RESULT_JSON);
  const errorIdx       = await ensureCol(COL.ERROR);
  const finalJudgeIdx  = await ensureCol(COL.FINAL_JUDGE);

  const now    = new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });
  const status = error ? '採点エラー' : '採点済';

  const writes: { i: number; v: string }[] = [
    { i: statusIdx,     v: status },
    { i: dateIdx,       v: now },
    { i: jsonIdx,       v: result ? JSON.stringify(result) : '' },
    { i: errorIdx,      v: error ?? '' },
  ];
  if (!error && finalJudge) writes.push({ i: finalJudgeIdx, v: finalJudge });

  await Promise.all(writes.map(({ i, v }) =>
    sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${SHEET_NAME}'!${colLetter(i + 1)}${rowNumber}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[v]] },
    }),
  ));
}

export async function updateApplicantFields(
  rowNumber: number,
  fields: { finalJudge?: string; notes?: string; promptChecks?: string; jobType?: string },
): Promise<void> {
  const sheets = await getSheetsClient();

  const headersRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!1:1`,
  });
  const headers = (headersRes.data.values?.[0] ?? []).map(String);

  const ensureCol = async (name: string): Promise<number> => {
    const normed = normalize(name);
    let i = headers.findIndex((h) => normalize(h) === normed);
    if (i < 0) {
      i = headers.length;
      headers.push(name);
      await expandColumnsIfNeeded(sheets, i + 1);
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${SHEET_NAME}'!${colLetter(i + 1)}1`,
        valueInputOption: 'RAW',
        requestBody: { values: [[name]] },
      });
    }
    return i;
  };

  const writes: { i: number; v: string }[] = [];
  if (fields.finalJudge !== undefined) {
    writes.push({ i: await ensureCol(COL.FINAL_JUDGE), v: fields.finalJudge });
  }
  if (fields.notes !== undefined) {
    writes.push({ i: await ensureCol(COL.NOTES), v: fields.notes });
  }
  if (fields.promptChecks !== undefined) {
    writes.push({ i: await ensureCol(COL.PROMPT_CHECKS), v: fields.promptChecks });
  }
  if (fields.jobType !== undefined) {
    writes.push({ i: await ensureCol(COL.JOB_TYPE), v: fields.jobType });
  }

  await Promise.all(writes.map(({ i, v }) =>
    sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${SHEET_NAME}'!${colLetter(i + 1)}${rowNumber}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[v]] },
    }),
  ));
}

export async function updateGradingStatus(rowNumber: number, status: string): Promise<void> {
  const sheets = await getSheetsClient();

  const headersRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!1:1`,
  });
  const headers = (headersRes.data.values?.[0] ?? []).map(String);

  const normed = normalize(COL.STATUS);
  let i = headers.findIndex((h) => normalize(h) === normed);
  if (i < 0) {
    i = headers.length;
    await expandColumnsIfNeeded(sheets, i + 1);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${SHEET_NAME}'!${colLetter(i + 1)}1`,
      valueInputOption: 'RAW',
      requestBody: { values: [[COL.STATUS]] },
    });
  }

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!${colLetter(i + 1)}${rowNumber}`,
    valueInputOption: 'RAW',
    requestBody: { values: [[status]] },
  });
}
