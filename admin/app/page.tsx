'use client';

import { useSession, signOut } from 'next-auth/react';
import { useRouter } from 'next/navigation';
import { useEffect, useState, useCallback } from 'react';
import type { ApplicantRow } from '@/lib/types';

type Filter = 'all' | 'pending' | 'graded' | 'error';

const STATUS_LABEL: Record<string, { label: string; className: string }> = {
  '採点済':    { label: '採点済',  className: 'bg-green-100 text-green-800' },
  '採点中':    { label: '採点中',  className: 'bg-blue-100 text-blue-800' },
  '採点エラー': { label: 'エラー',  className: 'bg-red-100 text-red-800' },
};

function finalJudgeBadge(judge: string) {
  if (judge === '合格')  return <span className="px-2 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800">合格</span>;
  if (judge === '不合格') return <span className="px-2 py-0.5 rounded-full text-xs font-medium bg-red-100 text-red-800">不合格</span>;
  if (judge === '保留')  return <span className="px-2 py-0.5 rounded-full text-xs font-medium bg-yellow-100 text-yellow-800">保留</span>;
  return <span className="px-2 py-0.5 rounded-full text-xs font-medium bg-gray-100 text-gray-400">未判定</span>;
}

function statusBadge(status: string) {
  const s = STATUS_LABEL[status];
  if (s) return <span className={`px-2 py-0.5 rounded-full text-xs font-medium ${s.className}`}>{s.label}</span>;
  return <span className="px-2 py-0.5 rounded-full text-xs font-medium bg-gray-100 text-gray-600">未採点</span>;
}

export default function DashboardPage() {
  const { data: session, status } = useSession();
  const router = useRouter();
  const [applicants, setApplicants] = useState<ApplicantRow[]>([]);
  const [loading, setLoading]       = useState(true);
  const [error, setError]           = useState('');
  const [filter, setFilter]         = useState<Filter>('all');

  useEffect(() => {
    if (status === 'unauthenticated') router.push('/login');
  }, [status, router]);

  const load = useCallback(async () => {
    setLoading(true);
    setError('');
    try {
      const res = await fetch('/api/applicants');
      if (!res.ok) throw new Error(await res.text());
      setApplicants(await res.json());
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    if (status === 'authenticated') load();
  }, [status, load]);

  if (status === 'loading') return <div className="flex items-center justify-center min-h-screen text-gray-400">読み込み中...</div>;
  if (status === 'unauthenticated') return null;

  const filtered = applicants
    .filter((a) => {
      if (filter === 'pending') return !a.gradingStatus || a.gradingStatus === '';
      if (filter === 'graded')  return a.gradingStatus === '採点済';
      if (filter === 'error')   return a.gradingStatus === '採点エラー';
      return true;
    })
    .sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());

  const total   = applicants.length;
  const graded  = applicants.filter((a) => a.gradingStatus === '採点済').length;
  const pending = applicants.filter((a) => !a.gradingStatus || a.gradingStatus === '').length;
  const errors  = applicants.filter((a) => a.gradingStatus === '採点エラー').length;

  return (
    <div className="min-h-screen">
      {/* ヘッダー */}
      <header className="bg-white border-b border-gray-200 px-6 py-4 flex items-center justify-between">
        <h1 className="text-lg font-bold">採点管理システム</h1>
        <div className="flex items-center gap-4">
          <span className="text-sm text-gray-500">{session?.user?.email}</span>
          <button
            onClick={() => signOut({ callbackUrl: '/login' })}
            className="text-sm text-gray-500 hover:text-gray-700"
          >
            ログアウト
          </button>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-8">
        {/* サマリー */}
        <div className="grid grid-cols-4 gap-4 mb-6">
          {[
            { label: '合計',   value: total,   color: 'text-gray-900' },
            { label: '未採点', value: pending, color: 'text-yellow-600' },
            { label: '採点済', value: graded,  color: 'text-green-600' },
            { label: 'エラー', value: errors,  color: 'text-red-600' },
          ].map(({ label, value, color }) => (
            <div key={label} className="bg-white rounded-xl border border-gray-200 p-4">
              <p className="text-xs text-gray-500 mb-1">{label}</p>
              <p className={`text-2xl font-bold ${color}`}>{value}</p>
            </div>
          ))}
        </div>

        {/* フィルター + リフレッシュ */}
        <div className="flex items-center justify-between mb-4">
          <div className="flex gap-2">
            {(['all', 'pending', 'graded', 'error'] as Filter[]).map((f) => (
              <button
                key={f}
                onClick={() => setFilter(f)}
                className={`px-3 py-1.5 rounded-lg text-sm font-medium transition ${
                  filter === f ? 'bg-blue-600 text-white' : 'bg-white border border-gray-200 text-gray-600 hover:bg-gray-50'
                }`}
              >
                {{ all: '全件', pending: '未採点', graded: '採点済', error: 'エラー' }[f]}
              </button>
            ))}
          </div>
          <button
            onClick={load}
            disabled={loading}
            className="text-sm text-blue-600 hover:text-blue-800 disabled:opacity-40"
          >
            {loading ? '読み込み中...' : '更新'}
          </button>
        </div>

        {error && (
          <div className="bg-red-50 border border-red-200 text-red-700 rounded-lg p-3 mb-4 text-sm">
            {error}
          </div>
        )}

        {/* テーブル */}
        <div className="bg-white rounded-xl border border-gray-200 overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200">
              <tr>
                <th className="px-4 py-3 text-left font-medium text-gray-600">行</th>
                <th className="px-4 py-3 text-left font-medium text-gray-600">氏名</th>
                <th className="px-4 py-3 text-left font-medium text-gray-600">メール</th>
                <th className="px-4 py-3 text-left font-medium text-gray-600">回答日時</th>
                <th className="px-4 py-3 text-left font-medium text-gray-600">AI使用</th>
                <th className="px-4 py-3 text-left font-medium text-gray-600">採点状況</th>
                <th className="px-4 py-3 text-left font-medium text-gray-600">合否</th>
                <th className="px-4 py-3 text-left font-medium text-gray-600">採点日時</th>
                <th className="px-4 py-3"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {loading && (
                <tr>
                  <td colSpan={9} className="px-4 py-8 text-center text-gray-400">読み込み中...</td>
                </tr>
              )}
              {!loading && filtered.length === 0 && (
                <tr>
                  <td colSpan={9} className="px-4 py-8 text-center text-gray-400">データがありません</td>
                </tr>
              )}
              {filtered.map((a) => (
                <tr
                  key={a.row}
                  className="hover:bg-gray-50 cursor-pointer"
                  onClick={() => router.push(`/applicant/${a.row}`)}
                >
                  <td className="px-4 py-3 text-gray-400">{a.row}</td>
                  <td className="px-4 py-3 font-medium">{a.name}</td>
                  <td className="px-4 py-3 text-gray-500">{a.email}</td>
                  <td className="px-4 py-3 text-gray-500 whitespace-nowrap">{a.timestamp}</td>
                  <td className="px-4 py-3">
                    {(a.aiLink || a.aiLog)
                      ? <span className="text-xs bg-purple-100 text-purple-700 px-2 py-0.5 rounded-full">あり</span>
                      : <span className="text-xs bg-gray-100 text-gray-500 px-2 py-0.5 rounded-full">なし</span>}
                  </td>
                  <td className="px-4 py-3">{statusBadge(a.gradingStatus)}</td>
                  <td className="px-4 py-3">{finalJudgeBadge(a.finalJudge)}</td>
                  <td className="px-4 py-3 text-gray-500 text-xs whitespace-nowrap">{a.gradedAt}</td>
                  <td className="px-4 py-3">
                    <span className="text-blue-600 hover:text-blue-800 text-xs font-medium">詳細 →</span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </main>
    </div>
  );
}
