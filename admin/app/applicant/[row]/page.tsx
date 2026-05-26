'use client';

import { useSession } from 'next-auth/react';
import { useRouter, useParams } from 'next/navigation';
import { useEffect, useState, useCallback } from 'react';
import type { ApplicantRow, GradingResult, CriteriaItem, JobType } from '@/lib/types';

const JUDGE_STYLE: Record<string, string> = {
  '〇':    'bg-green-50 text-green-700',
  '△':    'bg-yellow-50 text-yellow-700',
  '×':    'bg-red-50 text-red-700 font-bold',
  '判定不可': 'bg-gray-50 text-gray-500',
};

const OVERALL_STYLE: Record<string, string> = {
  '合格':    'bg-green-100 text-green-800 border-green-300',
  '不合格':  'bg-red-100 text-red-800 border-red-300',
  '保留':    'bg-yellow-100 text-yellow-800 border-yellow-300',
  '判定不可': 'bg-gray-100 text-gray-600 border-gray-300',
};

function CriteriaRow({ c }: { c: CriteriaItem }) {
  return (
    <tr className="border-b border-gray-100 last:border-0 align-top">
      <td className="px-3 py-2 text-gray-400 text-xs w-8">{c.no}</td>
      <td className="px-3 py-2 text-sm">{c.title}</td>
      <td className={`px-3 py-2 text-sm text-center w-16 ${JUDGE_STYLE[c.judgement] ?? ''}`}>
        {c.judgement}
      </td>
      <td className="px-3 py-2 text-xs text-gray-600 leading-relaxed">{c.reason}</td>
      <td className="px-3 py-2 text-xs text-gray-500 leading-relaxed italic">{c.relevant_text}</td>
    </tr>
  );
}

export default function ApplicantPage() {
  const { status } = useSession();
  const router     = useRouter();
  const { row }    = useParams<{ row: string }>();

  const [applicant, setApplicant] = useState<ApplicantRow | null>(null);
  const [result, setResult]       = useState<GradingResult | null>(null);
  const [loading, setLoading]     = useState(true);
  const [grading, setGrading]     = useState(false);
  const [error, setError]         = useState('');
  const [jobType, setJobType]     = useState<JobType>('秘書');
  const [finalJudge, setFinalJudge] = useState('');
  const [notes, setNotes]           = useState('');
  const [saving, setSaving]         = useState(false);
  const [saveMsg, setSaveMsg]       = useState('');
  const [promptChecks, setPromptChecks] = useState<boolean[]>(Array(7).fill(false));

  const toggleCheck = (i: number) =>
    setPromptChecks((prev) => prev.map((v, idx) => (idx === i ? !v : v)));

  const parsePromptChecks = (raw: string): boolean[] => {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length === 7) return parsed.map(Boolean);
    } catch { /* ignore */ }
    return Array(7).fill(false);
  };

  useEffect(() => {
    if (status === 'unauthenticated') router.push('/login');
  }, [status, router]);

  const load = useCallback(async () => {
    setLoading(true);
    setError('');
    try {
      const res = await fetch(`/api/applicants/${row}`);
      if (!res.ok) throw new Error(await res.text());
      const data: ApplicantRow = await res.json();
      setApplicant(data);
      setFinalJudge(data.finalJudge ?? '');
      setNotes(data.notes ?? '');
      setPromptChecks(parsePromptChecks(data.promptChecks ?? ''));
      if (data.jobType) setJobType(data.jobType as JobType);
      if (data.gradingResultJson) {
        setResult(JSON.parse(data.gradingResultJson) as GradingResult);
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }, [row]);

  useEffect(() => {
    if (status === 'authenticated') load();
  }, [status, load]);

  const handleGrade = async () => {
    if (!applicant) return;
    const alreadyGraded = applicant.gradingStatus === '採点済';
    if (alreadyGraded) {
      if (!window.confirm('すでに採点済みです。再採点しますか？')) return;
    }
    setGrading(true);
    setError('');
    try {
      const res = await fetch(`/api/grade/${row}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ jobType }),
      });
      const text = await res.text();
      let json: { error?: string; result?: GradingResult } = {};
      try {
        json = JSON.parse(text);
      } catch {
        throw new Error(`サーバーエラー (${res.status}): ${text.slice(0, 200)}`);
      }
      if (!res.ok) throw new Error(json.error ?? '採点に失敗しました');
      if (json.result) setResult(json.result);
      await load();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setGrading(false);
    }
  };

  const handleSaveMeta = async () => {
    setSaving(true);
    setSaveMsg('');
    try {
      const res = await fetch(`/api/applicants/${row}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ finalJudge, notes, promptChecks: JSON.stringify(promptChecks), jobType }),
      });
      if (!res.ok) throw new Error(await res.text());
      setSaveMsg('保存しました');
    } catch (e) {
      setSaveMsg(e instanceof Error ? e.message : '保存に失敗しました');
    } finally {
      setSaving(false);
    }
  };

  if (status === 'loading' || loading) {
    return <div className="flex items-center justify-center min-h-screen text-gray-400">読み込み中...</div>;
  }
  if (status === 'unauthenticated') return null;
  if (!applicant) return <div className="p-8 text-gray-500">応募者が見つかりません</div>;

  const aiUsed = applicant.aiLink || applicant.aiLog;

  return (
    <div className="min-h-screen">
      {/* ヘッダー */}
      <header className="bg-white border-b border-gray-200 px-6 py-4 flex items-center gap-4">
        <button onClick={() => router.push('/')} className="text-sm text-gray-500 hover:text-gray-700">
          ← 一覧に戻る
        </button>
        <h1 className="text-lg font-bold">{applicant.name} さんの採点</h1>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-8 space-y-6">
        {/* 応募者情報 */}
        <section className="bg-white rounded-xl border border-gray-200 p-5">
          <h2 className="font-semibold mb-3 text-gray-700">応募者情報</h2>
          <dl className="grid grid-cols-2 gap-x-8 gap-y-2 text-sm">
            <dt className="text-gray-500">氏名</dt>       <dd>{applicant.name}</dd>
            <dt className="text-gray-500">メール</dt>     <dd>{applicant.email}</dd>
            <dt className="text-gray-500">回答日時</dt>   <dd>{applicant.timestamp}</dd>
            <dt className="text-gray-500">AI使用</dt>
            <dd>
              {aiUsed
                ? <span className="text-purple-700 bg-purple-50 px-2 py-0.5 rounded text-xs">あり</span>
                : <span className="text-gray-500 bg-gray-50 px-2 py-0.5 rounded text-xs">なし</span>}
            </dd>
          </dl>
        </section>

        {/* 回答 */}
        {['Q1', 'Q2'].map((qid) => {
          const text  = qid === 'Q1' ? applicant.task1Answer : applicant.task2Answer;
          const label = qid === 'Q1' ? '課題1：お客様へのメール' : '課題2：現地対応スタッフへのメール';
          return (
            <section key={qid} className="bg-white rounded-xl border border-gray-200 p-5">
              <h2 className="font-semibold mb-3 text-gray-700">{label}</h2>
              <pre className="text-sm text-gray-700 whitespace-pre-wrap bg-gray-50 rounded-lg p-4 leading-relaxed">
                {text || '（回答なし）'}
              </pre>
            </section>
          );
        })}

        {/* その他の回答 */}
        <section className="bg-white rounded-xl border border-gray-200 p-5">
          <h2 className="font-semibold mb-3 text-gray-700 text-sm">その他の回答</h2>
          <dl className="grid grid-cols-3 gap-x-6 gap-y-1.5 text-xs mb-3">
            <dt className="text-gray-500">設問読解時間</dt>
            <dd className="col-span-2">{applicant.readingTime ? `${applicant.readingTime}分` : '—'}</dd>
            <dt className="text-gray-500">課題1 作成時間</dt>
            <dd className="col-span-2">{applicant.task1Time ? `${applicant.task1Time}分` : '—'}</dd>
            <dt className="text-gray-500">課題2 作成時間</dt>
            <dd className="col-span-2">{applicant.task2Time ? `${applicant.task2Time}分` : '—'}</dd>
            {applicant.desiredWorkHours && (
              <>
                <dt className="text-gray-500">希望勤務時間</dt>
                <dd className="col-span-2">{applicant.desiredWorkHours}</dd>
              </>
            )}
          </dl>
          <div className="space-y-1 text-xs border-t border-gray-100 pt-2">
            <div className="flex gap-2">
              <span className="text-gray-500 shrink-0">AI共有リンク:</span>
              {applicant.aiLink
                ? <a href={applicant.aiLink} target="_blank" rel="noopener noreferrer"
                    className="text-blue-600 hover:underline break-all">{applicant.aiLink}</a>
                : <span className="text-gray-400">—</span>}
            </div>
            <div className="flex gap-2">
              <span className="text-gray-500 shrink-0">AIログ (ギガファイル):</span>
              {applicant.aiLog
                ? <a href={applicant.aiLog} target="_blank" rel="noopener noreferrer"
                    className="text-blue-600 hover:underline break-all">{applicant.aiLog}</a>
                : <span className="text-gray-400">—</span>}
            </div>
          </div>
        </section>

        {/* プロンプト評価チェックリスト */}
        <section className="bg-white rounded-xl border border-gray-200 p-5">
          <h2 className="font-semibold mb-3 text-gray-700 text-sm">プロンプト評価チェックリスト</h2>
          <div className="space-y-2">
            {[
              'プロンプト内に「顧客を気遣う」表現や配慮に関する指示が含まれているか（例：「丁寧な文面」「安心感を与える」「不安を和らげる」など）',
              'プロンプト内に情報の正確性を確認・検証する旨の指示があるか（例：「日付や情報の正誤を確認」「事実確認を行う」など）',
              '必要情報や条件が具体的に指示されているか（例：「会場情報を含める」「終了時刻を明記する」など）',
              '誤情報や関係ない情報を含めない旨の指示があるか',
              'プロンプト内に最終成果物の目的が明確に示されているか（例：「顧客が安心できる文面を作成」「誤解のない事実確認を行う」など）',
              '「顧客に送るため」「社内共有資料として使うため」など、使うシーンが説明されているか',
              '結果を「売上に結びつける」視点（顧客獲得・効率化・改善提案など）がある',
            ].map((label, i) => (
              <label key={i} className="flex items-start gap-3 cursor-pointer group">
                <input
                  type="checkbox"
                  checked={promptChecks[i]}
                  onChange={() => toggleCheck(i)}
                  className="mt-0.5 w-4 h-4 accent-blue-600 shrink-0"
                />
                <span className={`text-sm leading-relaxed ${promptChecks[i] ? 'text-gray-800' : 'text-gray-500'}`}>
                  {label}
                </span>
              </label>
            ))}
          </div>
          <p className="mt-3 text-xs text-gray-400">
            チェック数: {promptChecks.filter(Boolean).length} / 7
          </p>
        </section>

        {/* 採点ボタン */}
        <div className="flex items-center gap-4">
          <select
            value={jobType}
            onChange={(e) => setJobType(e.target.value as JobType)}
            disabled={grading}
            className="px-3 py-3 rounded-xl border border-gray-300 bg-white text-sm text-gray-700 disabled:opacity-50"
          >
            <option>秘書</option>
            <option>海外営業</option>
            <option>オンラインセールス</option>
            <option>コールスタッフ</option>
          </select>
          <button
            onClick={handleGrade}
            disabled={grading}
            className="px-6 py-3 rounded-xl bg-blue-600 text-white font-semibold hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed transition text-sm"
          >
            {grading ? '採点中... (30〜60秒かかります)' : applicant.gradingStatus === '採点済' ? '再採点する' : '採点を実行する'}
          </button>
          {applicant.gradedAt && (
            <span className="text-xs text-gray-400">最終採点: {applicant.gradedAt}</span>
          )}
        </div>

        {error && (
          <div className="bg-red-50 border border-red-200 text-red-700 rounded-lg p-3 text-sm">
            {error}
          </div>
        )}

        {/* 採点結果 */}
        {result && (
          <section className="space-y-4">
            {/* 総合判定 */}
            <div className={`rounded-xl border p-5 ${OVERALL_STYLE[result.overall_judgement] ?? 'bg-gray-50 border-gray-200'}`}>
              <div className="flex items-center gap-4">
                <span className="text-2xl font-bold">{result.overall_judgement}</span>
                <span className="text-sm">合計点: <strong>{result.total_score}</strong> 点</span>
                <span className="text-sm">
                  NG数: <strong>{result.results?.reduce((sum, qr) => sum + (qr.criteria?.filter(c => c.judgement === '×').length ?? 0), 0)}</strong> 個
                </span>
              </div>
              {result.overall_comment && (
                <p className="mt-2 text-sm leading-relaxed">{result.overall_comment}</p>
              )}
            </div>

            {/* 設問ごとの結果 */}
            {result.results?.map((qr, i) => (
              <div key={`${qr.question_id}-${i}`} className="bg-white rounded-xl border border-gray-200 overflow-hidden">
                <div className={`px-5 py-3 font-semibold flex items-center justify-between ${qr.question_id === 'Q1' ? 'bg-blue-50' : 'bg-teal-50'}`}>
                  <span>{qr.question_id}：{qr.question_title}</span>
                  <div className="flex items-center gap-3 text-sm font-normal text-gray-600">
                    <span>NG数: <strong>{qr.criteria?.filter(c => c.judgement === '×').length ?? 0}</strong> 個</span>
                    <span>{qr.score} 点</span>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-gray-50 border-b border-gray-100">
                      <tr>
                        <th className="px-3 py-2 text-left text-xs text-gray-500 w-8">No</th>
                        <th className="px-3 py-2 text-left text-xs text-gray-500">採点項目</th>
                        <th className="px-3 py-2 text-left text-xs text-gray-500 w-16">判定</th>
                        <th className="px-3 py-2 text-left text-xs text-gray-500">判定理由</th>
                        <th className="px-3 py-2 text-left text-xs text-gray-500">根拠箇所</th>
                      </tr>
                    </thead>
                    <tbody>
                      {qr.criteria?.map((c) => <CriteriaRow key={c.no} c={c} />)}
                    </tbody>
                  </table>
                </div>
                {qr.overall_comment && (
                  <div className="px-5 py-3 border-t border-gray-100 text-sm text-gray-600 bg-gray-50">
                    {qr.overall_comment}
                  </div>
                )}
              </div>
            ))}
          </section>
        )}

        {/* 採用判断・備考 */}
        <section className="bg-white rounded-xl border border-gray-200 p-5 space-y-4">
          <h2 className="font-semibold text-gray-700">採用判断・備考</h2>
          <div className="flex items-center gap-4">
            <label className="text-sm text-gray-500 w-16 shrink-0">合否</label>
            <select
              value={finalJudge}
              onChange={(e) => setFinalJudge(e.target.value)}
              className="px-3 py-2 rounded-lg border border-gray-300 bg-white text-sm text-gray-700"
            >
              <option value="">未判定</option>
              <option value="合格">合格</option>
              <option value="不合格">不合格</option>
              <option value="保留">保留</option>
            </select>
          </div>
          <div className="flex gap-4">
            <label className="text-sm text-gray-500 w-16 shrink-0 pt-2">備考</label>
            <textarea
              value={notes}
              onChange={(e) => setNotes(e.target.value)}
              rows={4}
              className="flex-1 px-3 py-2 rounded-lg border border-gray-300 text-sm text-gray-700 resize-y"
              placeholder="備考・メモを入力"
            />
          </div>
          <div className="flex items-center gap-3">
            <button
              onClick={handleSaveMeta}
              disabled={saving}
              className="px-5 py-2 rounded-lg bg-gray-800 text-white text-sm font-medium hover:bg-gray-900 disabled:opacity-50 transition"
            >
              {saving ? '保存中...' : '保存する'}
            </button>
            {saveMsg && <span className="text-sm text-gray-500">{saveMsg}</span>}
          </div>
        </section>
      </main>
    </div>
  );
}
