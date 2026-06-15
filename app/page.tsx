'use client';

import { useState } from 'react';

type Status = { type: 'error' | 'success'; message: string } | null;

export default function Home() {
  const [docxFile, setDocxFile] = useState<File | null>(null);
  const [photoFiles, setPhotoFiles] = useState<File[]>([]);
  const [columns, setColumns] = useState(3);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState<Status>(null);

  async function handleSubmit(e: React.FormEvent<HTMLFormElement>) {
    e.preventDefault();
    setStatus(null);

    if (!docxFile) {
      setStatus({ type: 'error', message: '기본 Word 파일(.docx)을 선택해주세요.' });
      return;
    }
    if (photoFiles.length === 0) {
      setStatus({ type: 'error', message: '사진을 최소 1장 선택해주세요.' });
      return;
    }

    setLoading(true);
    try {
      const formData = new FormData();
      formData.append('docx', docxFile);
      formData.append('columns', String(columns));
      for (const photo of photoFiles) {
        formData.append('photos', photo);
      }

      console.log('[generate] 요청 시작:', docxFile.name, photoFiles.length + '장');

      const res = await fetch('/api/generate-report', {
        method: 'POST',
        body: formData,
      });

      console.log('[generate] 응답 상태:', res.status);

      if (!res.ok) {
        const json = await res.json().catch(() => ({ message: '알 수 없는 오류' }));
        throw new Error(json.message ?? '서버 오류가 발생했습니다.');
      }

      const blob = await res.blob();
      console.log('[generate] blob 크기:', blob.size, 'bytes');

      const url = URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = `report_${Date.now()}.docx`;
      anchor.style.display = 'none';
      document.body.appendChild(anchor);
      anchor.click();
      document.body.removeChild(anchor);
      setTimeout(() => URL.revokeObjectURL(url), 1000);

      setStatus({ type: 'success', message: '보고서 생성 완료! 다운로드가 시작됩니다.' });
    } catch (err) {
      const msg = err instanceof Error ? err.message : '오류가 발생했습니다.';
      console.error('[generate] 오류:', err);
      setStatus({ type: 'error', message: msg });
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="min-h-screen bg-gray-100 flex items-center justify-center p-6">
      <div className="bg-white rounded-2xl shadow-md w-full max-w-lg p-8 space-y-6">
        <div>
          <h1 className="text-2xl font-bold text-gray-800">사진 첨부 보고서 생성기</h1>
          <p className="mt-1 text-sm text-gray-500">
            Word 파일 맨 아래에 사진을 가로 3열 표로 자동 첨부합니다.
          </p>
        </div>

        <form onSubmit={handleSubmit} className="space-y-5">
          {/* DOCX upload */}
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              기본 Word 파일 <span className="text-gray-400">(.docx)</span>
            </label>
            <input
              type="file"
              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              onChange={e => setDocxFile(e.target.files?.[0] ?? null)}
              className="block w-full text-sm text-gray-500
                file:mr-3 file:py-2 file:px-4
                file:rounded-lg file:border-0
                file:text-sm file:font-medium
                file:bg-blue-50 file:text-blue-700
                hover:file:bg-blue-100 cursor-pointer"
            />
            {docxFile && (
              <p className="mt-1 text-xs text-gray-400 truncate">선택됨: {docxFile.name}</p>
            )}
          </div>

          {/* Photo upload + column select */}
          <div>
            <div className="flex items-center justify-between mb-1">
              <label className="text-sm font-medium text-gray-700">
                사진 파일{' '}
                <span className="text-gray-400">(jpg / jpeg / png / webp, 여러 장 가능)</span>
              </label>
              <select
                value={columns}
                onChange={e => setColumns(Number(e.target.value))}
                className="text-xs text-gray-600 border border-gray-300 rounded px-2 py-1 bg-white"
              >
                {[2, 3, 4, 5].map(n => (
                  <option key={n} value={n}>{n}열</option>
                ))}
              </select>
            </div>
            <input
              type="file"
              accept=".jpg,.jpeg,.png,.webp,image/jpeg,image/png,image/webp"
              multiple
              onChange={e => setPhotoFiles(e.target.files ? Array.from(e.target.files) : [])}
              className="block w-full text-sm text-gray-500
                file:mr-3 file:py-2 file:px-4
                file:rounded-lg file:border-0
                file:text-sm file:font-medium
                file:bg-green-50 file:text-green-700
                hover:file:bg-green-100 cursor-pointer"
            />
            {photoFiles.length > 0 && (
              <p className="mt-1 text-xs text-gray-400">{photoFiles.length}개 파일 선택됨</p>
            )}
          </div>

          {/* Submit */}
          <button
            type="submit"
            disabled={loading}
            className="w-full py-3 rounded-xl bg-blue-600 text-white text-sm font-semibold
              hover:bg-blue-700 transition-colors
              disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {loading ? '생성 중...' : '보고서 생성'}
          </button>

          {/* Status message */}
          {status && (
            <div
              className={`rounded-lg px-4 py-3 text-sm font-medium ${
                status.type === 'success'
                  ? 'bg-green-50 text-green-700 border border-green-200'
                  : 'bg-red-50 text-red-700 border border-red-200'
              }`}
            >
              {status.message}
            </div>
          )}
        </form>

        <p className="text-xs text-gray-400 text-center">
          업로드 파일은 서버에 저장되지 않으며 요청 처리 후 즉시 폐기됩니다.
        </p>
      </div>
    </main>
  );
}
