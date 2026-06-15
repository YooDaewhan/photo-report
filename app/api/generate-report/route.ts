import { NextRequest, NextResponse } from 'next/server';
import { buildReport } from '@/lib/docx-builder';

export const runtime = 'nodejs';

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();

    const docxFile = formData.get('docx') as File | null;
    const photoFiles = formData.getAll('photos') as File[];

    if (!docxFile || docxFile.size === 0) {
      return NextResponse.json({ message: 'Word 파일(.docx)이 필요합니다.' }, { status: 400 });
    }
    if (photoFiles.length === 0) {
      return NextResponse.json({ message: '사진이 최소 1장 필요합니다.' }, { status: 400 });
    }

    const docxBuffer = Buffer.from(await docxFile.arrayBuffer());

    const photos = await Promise.all(
      photoFiles.map(async file => ({
        name: file.name,
        buffer: Buffer.from(await file.arrayBuffer()),
      }))
    );

    const columns = parseInt(formData.get('columns') as string ?? '3', 10) || 3;

    const resultBuffer = await buildReport(docxBuffer, photos, columns);

    return new NextResponse(new Uint8Array(resultBuffer), {
      status: 200,
      headers: {
        'Content-Type':
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': 'attachment; filename="report.docx"',
        'Content-Length': String(resultBuffer.length),
      },
    });
  } catch (err) {
    console.error('[generate-report]', err);
    const message =
      err instanceof Error ? err.message : '서버에서 오류가 발생했습니다.';
    return NextResponse.json({ message }, { status: 500 });
  }
}
