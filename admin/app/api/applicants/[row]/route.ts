import { NextResponse } from 'next/server';
import { getServerSession } from 'next-auth';
import { authOptions } from '@/lib/auth';
import { getApplicantByRow, updateApplicantFields } from '@/lib/sheets';

export const dynamic = 'force-dynamic';

export async function GET(
  _req: Request,
  { params }: { params: Promise<{ row: string }> },
) {
  const session = await getServerSession(authOptions);
  if (!session) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  const { row } = await params;
  const rowNum = parseInt(row, 10);
  if (isNaN(rowNum)) {
    return NextResponse.json({ error: 'Invalid row' }, { status: 400 });
  }

  try {
    const applicant = await getApplicantByRow(rowNum);
    if (!applicant) {
      return NextResponse.json({ error: 'Not found' }, { status: 404 });
    }
    return NextResponse.json(applicant);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function PATCH(
  req: Request,
  { params }: { params: Promise<{ row: string }> },
) {
  const session = await getServerSession(authOptions);
  if (!session) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  const { row } = await params;
  const rowNum = parseInt(row, 10);
  if (isNaN(rowNum)) {
    return NextResponse.json({ error: 'Invalid row' }, { status: 400 });
  }

  try {
    const body = await req.json() as { finalJudge?: string; notes?: string; promptChecks?: string; jobType?: string; gradingResultJson?: string };
    await updateApplicantFields(rowNum, body);
    return NextResponse.json({ success: true });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
