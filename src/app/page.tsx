"use client";

import React, { useState, useRef } from "react";

interface CmMaterial {
  name: string;
  abbr: string;
  material: string;
  status: string;
  period: string;
  seconds: number;
}

interface FileResult {
  name: string;
  status: "waiting" | "processing" | "done" | "error";
  message?: string;
}

const defaultMaterial = (): CmMaterial => ({
  name: "", abbr: "", material: "OL", status: "送", period: "", seconds: 15,
});

export default function Home() {
  const [meisaiFiles, setMeisaiFiles] = useState<File[]>([]);
  const [station, setStation]         = useState("");
  const [person, setPerson]           = useState("");
  const [docDate, setDocDate]         = useState("");
  const [materials, setMaterials]     = useState<CmMaterial[]>([defaultMaterial()]);
  const [loading, setLoading]         = useState(false);
  const [results, setResults]         = useState<FileResult[]>([]);

  const meisaiRef = useRef<HTMLInputElement>(null);

  const addMaterial    = () => setMaterials([...materials, defaultMaterial()]);
  const removeMaterial = (i: number) => setMaterials(materials.filter((_, idx) => idx !== i));
  const updateMaterial = (i: number, key: keyof CmMaterial, val: string | number) => {
    setMaterials(materials.map((m, idx) => idx === i ? { ...m, [key]: val } : m));
  };

  const handleFiles = (files: FileList | null) => {
    if (!files) return;
    setMeisaiFiles(Array.from(files));
    setResults([]);
  };

  const removeFile = (i: number) => {
    setMeisaiFiles(meisaiFiles.filter((_, idx) => idx !== i));
    if (meisaiRef.current) meisaiRef.current.value = "";
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (meisaiFiles.length === 0) return;

    setLoading(true);
    const initial: FileResult[] = meisaiFiles.map(f => ({ name: f.name, status: "waiting" }));
    setResults(initial);

    for (let i = 0; i < meisaiFiles.length; i++) {
      setResults(prev => prev.map((r, idx) => idx === i ? { ...r, status: "processing" } : r));

      const form = new FormData();
      form.append("meisai",    meisaiFiles[i]);
      form.append("station",   station);
      form.append("person",    person);
      form.append("date",      docDate);
      form.append("materials", JSON.stringify(materials));

      try {
        const res = await fetch("/api/generate", { method: "POST", body: form });
        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          throw new Error((err as { error?: string }).error ?? `エラー: ${res.status}`);
        }
        const blob = await res.blob();
        const url  = URL.createObjectURL(blob);
        const a    = document.createElement("a");
        const cd   = res.headers.get("Content-Disposition") ?? "";
        const fnm  = cd.match(/filename\*?=(?:UTF-8'')?([^;]+)/)?.[1] ?? "CMスケ表.xlsx";
        a.href     = url;
        a.download = decodeURIComponent(fnm.replace(/"/g, ""));
        a.click();
        URL.revokeObjectURL(url);
        setResults(prev => prev.map((r, idx) => idx === i ? { ...r, status: "done" } : r));
      } catch (err) {
        const msg = err instanceof Error ? err.message : "不明なエラー";
        setResults(prev => prev.map((r, idx) => idx === i ? { ...r, status: "error", message: msg } : r));
      }
    }

    setLoading(false);
  };

  const allDone  = results.length > 0 && results.every(r => r.status === "done");
  const hasError = results.some(r => r.status === "error");

  return (
    <main className="min-h-screen bg-slate-50 py-10 px-4">
      <div className="max-w-3xl mx-auto">
        <div className="mb-8 text-center">
          <h1 className="text-3xl font-bold text-slate-800 tracking-tight">
            📺 CMスケ表
          </h1>
          <p className="mt-2 text-slate-500 text-sm">
            TV局の明細リストPDFを読み取り、スケジュール表Excelを自動生成します
          </p>
        </div>

        <form onSubmit={handleSubmit} className="space-y-6">
          {/* PDF */}
          <section className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
            <h2 className="text-lg font-semibold text-slate-700 mb-4">📄 明細リスト PDF（複数選択可）</h2>

            {/* ドロップゾーン */}
            <label className="flex flex-col items-center justify-center gap-2 cursor-pointer border-2 border-dashed border-slate-300 hover:border-blue-400 rounded-xl px-4 py-6 transition bg-slate-50 hover:bg-blue-50">
              <span className="text-3xl">📂</span>
              <span className="text-sm text-slate-500">
                クリックしてPDFを選択（複数選択OK）
              </span>
              <input
                ref={meisaiRef}
                type="file"
                accept=".pdf"
                multiple
                required={meisaiFiles.length === 0}
                className="hidden"
                onChange={e => handleFiles(e.target.files)}
              />
            </label>

            {/* 選択済みファイル一覧 */}
            {meisaiFiles.length > 0 && (
              <ul className="mt-3 space-y-1">
                {meisaiFiles.map((f, i) => (
                  <li key={i} className="flex items-center gap-2 text-sm bg-slate-100 rounded-lg px-3 py-2">
                    <span className="text-base">📄</span>
                    <span className="flex-1 truncate text-slate-700">{f.name}</span>
                    <button type="button" onClick={() => removeFile(i)}
                      className="text-slate-400 hover:text-red-400 text-lg leading-none">×</button>
                  </li>
                ))}
              </ul>
            )}
          </section>

          {/* 局情報 */}
          <section className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
            <h2 className="text-lg font-semibold text-slate-700 mb-4">📡 放送局情報</h2>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <Field label="放送局名" placeholder="例: FCT"
                value={station} onChange={setStation} />
              <Field label="局担当者名" placeholder="例: 岩沢"
                value={person} onChange={setPerson} />
              <div>
                <label className="block text-sm font-medium text-slate-600 mb-1">作成日</label>
                <input
                  type="date"
                  value={docDate.replace(/\//g, "-")}
                  onChange={e => setDocDate(e.target.value.replace(/-/g, "/"))}
                  className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
              </div>
            </div>
          </section>

          {/* CM素材 */}
          <section className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-lg font-semibold text-slate-700">🎬 CM素材</h2>
              <button type="button" onClick={addMaterial}
                className="text-sm bg-blue-50 text-blue-600 hover:bg-blue-100 px-3 py-1 rounded-lg font-medium transition">
                ＋ 素材を追加
              </button>
            </div>
            <div className="space-y-4">
              {materials.map((mat, i) => (
                <div key={i} className="border border-slate-200 rounded-xl p-4 bg-slate-50 relative">
                  {materials.length > 1 && (
                    <button type="button" onClick={() => removeMaterial(i)}
                      className="absolute top-3 right-3 text-slate-400 hover:text-red-400 text-xl leading-none">×</button>
                  )}
                  <p className="text-xs font-semibold text-slate-400 mb-3">素材 {i + 1}</p>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <div className="sm:col-span-2">
                      <Field label="素材名" placeholder="例: ふくしまの酒まつり 前売りチケット篇"
                        value={mat.name} onChange={v => updateMaterial(i, "name", v)} />
                    </div>
                    <Field label="略号" placeholder="例: S, A, B"
                      value={mat.abbr} onChange={v => updateMaterial(i, "abbr", v.toUpperCase())} />
                    <Field label="使用期間" placeholder="例: 2/10〜2/23"
                      value={mat.period} onChange={v => updateMaterial(i, "period", v)} />
                    <SelectField label="素材種別" value={mat.material}
                      options={["OL", "XD", "F", "S"]} onChange={v => updateMaterial(i, "material", v)} />
                    <SelectField label="在送使廻" value={mat.status}
                      options={["送", "在", "廻", "使"]} onChange={v => updateMaterial(i, "status", v)} />
                  </div>
                </div>
              ))}
            </div>
          </section>

          {/* 処理結果 */}
          {results.length > 0 && (
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 space-y-2">
              <h2 className="text-sm font-semibold text-slate-600 mb-3">処理状況</h2>
              {results.map((r, i) => (
                <div key={i} className={`flex items-center gap-3 rounded-lg px-3 py-2 text-sm
                  ${r.status === "done"       ? "bg-green-50 text-green-700"
                  : r.status === "error"      ? "bg-red-50 text-red-700"
                  : r.status === "processing" ? "bg-blue-50 text-blue-700"
                  : "bg-slate-50 text-slate-400"}`}>
                  <span>
                    {r.status === "done"       ? "✅"
                    : r.status === "error"     ? "⚠️"
                    : r.status === "processing" ? "⏳"
                    : "⬜"}
                  </span>
                  <span className="flex-1 truncate">{r.name}</span>
                  {r.message && <span className="text-xs">{r.message}</span>}
                </div>
              ))}
              {allDone && (
                <p className="text-green-600 text-sm font-medium pt-1">
                  全ファイルの生成が完了しました！
                </p>
              )}
              {hasError && !loading && (
                <p className="text-red-600 text-sm pt-1">
                  一部のファイルでエラーが発生しました。
                </p>
              )}
            </div>
          )}

          <button type="submit" disabled={loading || meisaiFiles.length === 0}
            className="w-full py-3 rounded-xl font-semibold text-white text-base transition
              bg-blue-600 hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed shadow-sm">
            {loading
              ? `⏳ 処理中... (${results.filter(r => r.status === "done" || r.status === "error").length} / ${results.length})`
              : `📥 スケジュール表を生成・ダウンロード${meisaiFiles.length > 1 ? `（${meisaiFiles.length}局）` : ""}`}
          </button>
        </form>

        <p className="mt-6 text-center text-xs text-slate-400">
          CMスケ表 — 明細リストPDF → Excelスケジュール表 自動生成システム
        </p>
      </div>
    </main>
  );
}

// ── 共通コンポーネント ──────────────────────────────

function Field({ label, placeholder, value, onChange }: {
  label: string; placeholder: string; value: string; onChange: (v: string) => void;
}) {
  return (
    <div>
      <label className="block text-sm font-medium text-slate-600 mb-1">{label}</label>
      <input type="text" placeholder={placeholder} value={value}
        onChange={e => onChange(e.target.value)}
        className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
    </div>
  );
}

function SelectField({ label, value, options, onChange }: {
  label: string; value: string; options: string[]; onChange: (v: string) => void;
}) {
  return (
    <div>
      <label className="block text-sm font-medium text-slate-600 mb-1">{label}</label>
      <select value={value} onChange={e => onChange(e.target.value)}
        className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white">
        {options.map(o => <option key={o} value={o}>{o}</option>)}
      </select>
    </div>
  );
}
