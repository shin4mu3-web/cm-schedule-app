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

interface FileEntry {
  file: File;
  station: string;
  person: string;
}

interface FileResult {
  name: string;
  status: "processing" | "done" | "error";
  message?: string;
}

const defaultMaterial = (): CmMaterial => ({
  name: "", abbr: "", material: "OL", status: "送", period: "", seconds: 15,
});

export default function Home() {
  const [entries, setEntries]       = useState<FileEntry[]>([]);
  const [docDate, setDocDate]       = useState("");
  const [materials, setMaterials]   = useState<CmMaterial[]>([defaultMaterial()]);
  const [loading, setLoading]       = useState(false);
  const [results, setResults]       = useState<FileResult[]>([]);

  const meisaiRef = useRef<HTMLInputElement>(null);

  const addMaterial    = () => setMaterials([...materials, defaultMaterial()]);
  const removeMaterial = (i: number) => setMaterials(materials.filter((_, idx) => idx !== i));
  const updateMaterial = (i: number, key: keyof CmMaterial, val: string | number) =>
    setMaterials(materials.map((m, idx) => idx === i ? { ...m, [key]: val } : m));

  const handleFiles = (files: FileList | null) => {
    if (!files) return;
    setEntries(prev => {
      const existing = new Set(prev.map(e => e.file.name));
      const added = Array.from(files)
        .filter(f => !existing.has(f.name))
        .map(f => ({ file: f, station: "", person: "" }));
      return [...prev, ...added];
    });
    setResults([]);
    if (meisaiRef.current) meisaiRef.current.value = "";
  };

  const removeEntry = (i: number) => setEntries(entries.filter((_, idx) => idx !== i));
  const updateEntry = (i: number, key: "station" | "person", val: string) =>
    setEntries(entries.map((e, idx) => idx === i ? { ...e, [key]: val } : e));

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (entries.length === 0) return;

    setLoading(true);
    setResults(entries.map(e => ({ name: e.file.name, status: "processing" as const })));

    const form = new FormData();
    entries.forEach(e => form.append("meisai", e.file));
    form.append("stations_info", JSON.stringify(entries.map(e => ({
      station: e.station,
      person:  e.person,
    }))));
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
      setResults(entries.map(e => ({ name: e.file.name, status: "done" as const })));
    } catch (err) {
      const msg = err instanceof Error ? err.message : "不明なエラー";
      setResults(entries.map(e => ({ name: e.file.name, status: "error" as const, message: msg })));
    }

    setLoading(false);
  };

  const allDone  = results.length > 0 && results.every(r => r.status === "done");
  const hasError = results.some(r => r.status === "error");

  return (
    <main className="min-h-screen bg-slate-50 py-10 px-4">
      <div className="max-w-4xl mx-auto">
        <div className="mb-8 text-center">
          <h1 className="text-3xl font-bold text-slate-800 tracking-tight">📺 CMスケ表</h1>
          <p className="mt-2 text-slate-500 text-sm">
            TV局の明細リストPDFを読み取り、スケジュール表Excelを自動生成します
          </p>
        </div>

        <form onSubmit={handleSubmit} className="space-y-6">

          {/* PDF + 局情報 */}
          <section className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
            <h2 className="text-lg font-semibold text-slate-700 mb-4">📄 明細リスト PDF・放送局情報</h2>

            <label className="flex flex-col items-center justify-center gap-2 cursor-pointer border-2 border-dashed border-slate-300 hover:border-blue-400 rounded-xl px-4 py-5 transition bg-slate-50 hover:bg-blue-50">
              <span className="text-3xl">📂</span>
              <span className="text-sm text-slate-500">クリックしてPDFを選択（複数選択OK）</span>
              <input ref={meisaiRef} type="file" accept=".pdf" multiple
                required={entries.length === 0} className="hidden"
                onChange={e => handleFiles(e.target.files)} />
            </label>

            {entries.length > 0 && (
              <div className="mt-4 space-y-2">
                {/* ヘッダー */}
                <div className="hidden sm:grid grid-cols-12 gap-2 px-2 text-xs font-medium text-slate-400">
                  <span className="col-span-5">ファイル名</span>
                  <span className="col-span-3">放送局名</span>
                  <span className="col-span-3">担当者名</span>
                  <span className="col-span-1"></span>
                </div>
                {entries.map((entry, i) => (
                  <div key={i} className="grid grid-cols-12 gap-2 items-center bg-slate-50 border border-slate-200 rounded-xl px-3 py-2">
                    <div className="col-span-12 sm:col-span-5 flex items-center gap-2 min-w-0">
                      <span className="text-base shrink-0">📄</span>
                      <span className="text-sm text-slate-700 truncate">{entry.file.name}</span>
                    </div>
                    <div className="col-span-5 sm:col-span-3">
                      <input type="text" placeholder="例: FCT" value={entry.station}
                        onChange={e => updateEntry(i, "station", e.target.value)}
                        className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                    </div>
                    <div className="col-span-5 sm:col-span-3">
                      <input type="text" placeholder="例: 岩沢" value={entry.person}
                        onChange={e => updateEntry(i, "person", e.target.value)}
                        className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
                    </div>
                    <div className="col-span-2 sm:col-span-1 text-right">
                      <button type="button" onClick={() => removeEntry(i)}
                        className="text-slate-400 hover:text-red-400 text-xl leading-none">×</button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </section>

          {/* 作成日 */}
          <section className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
            <h2 className="text-lg font-semibold text-slate-700 mb-4">📅 作成日</h2>
            <input type="date" value={docDate.replace(/\//g, "-")}
              onChange={e => setDocDate(e.target.value.replace(/-/g, "/"))}
              className="rounded-lg border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
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
                  : "bg-blue-50 text-blue-700"}`}>
                  <span>{r.status === "done" ? "✅" : r.status === "error" ? "⚠️" : "⏳"}</span>
                  <span className="flex-1 truncate">{r.name}</span>
                  {r.message && <span className="text-xs">{r.message}</span>}
                </div>
              ))}
              {allDone && <p className="text-green-600 text-sm font-medium pt-1">全ファイルの生成が完了しました！</p>}
              {hasError && !loading && <p className="text-red-600 text-sm pt-1">エラーが発生しました。</p>}
            </div>
          )}

          <button type="submit" disabled={loading || entries.length === 0}
            className="w-full py-3 rounded-xl font-semibold text-white text-base transition
              bg-blue-600 hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed shadow-sm">
            {loading ? "⏳ 生成中..."
              : `📥 スケジュール表を生成・ダウンロード${entries.length > 1 ? `（${entries.length}局）` : ""}`}
          </button>
        </form>

        <p className="mt-6 text-center text-xs text-slate-400">
          CMスケ表 — 明細リストPDF → Excelスケジュール表 自動生成システム
        </p>
      </div>
    </main>
  );
}

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
