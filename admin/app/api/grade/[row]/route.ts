import { NextResponse } from 'next/server';
import { getServerSession } from 'next-auth';
import { authOptions } from '@/lib/auth';
import { getApplicantByRow, writeGradingResult, updateGradingStatus } from '@/lib/sheets';
import { gradeApplicant } from '@/lib/grading';
import type { JobType } from '@/lib/types';

export const maxDuration = 60;

export async function POST(
  _req: Request,
  { params }: { params: Promise<{ row: string }> },
) {
  try {
    const session = await getServerSession(authOptions);
    if (!session) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    const { row } = await params;
    const rowNum = parseInt(row, 10);
    if (isNaN(rowNum)) {
      return NextResponse.json({ error: 'Invalid row' }, { status: 400 });
    }

    const applicant = await getApplicantByRow(rowNum);
    if (!applicant) {
      return NextResponse.json({ error: `行 ${rowNum} の応募者が見つかりません` }, { status: 404 });
    }

    const body = await _req.json().catch(() => ({})) as { jobType?: JobType };
    const jobType: JobType = body.jobType ?? '秘書';

    // 採点中ステータスをセット（失敗しても続行）
    try {
      await updateGradingStatus(rowNum, '採点中');
    } catch (e) {
      console.error('updateGradingStatus error:', e);
    }

    try {
      const result = await gradeApplicant(applicant, jobType);
      await writeGradingResult(rowNum, result, undefined, result.overall_judgement);
      return NextResponse.json({ success: true, result });
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      console.error('gradeApplicant error:', message);
      try {
        await writeGradingResult(rowNum, null, message);
      } catch (writeErr) {
        console.error('writeGradingResult error:', writeErr);
      }
      return NextResponse.json({ error: message }, { status: 500 });
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error('grade route unexpected error:', message);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
