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

const defaultMaterial = (): CmMaterial => ({
  name: "", abbr: "", material: "OL", status: "送", period: "", seconds: 15,
});

export default function Home() {
  const [meisai, setMeisai]       = useState<File | null>(null);
  const [timetable, setTimetable] = useState<File | null>(null);
  const [station, setStation]     = useState("");
  const [person, setPerson]       = useState("");
  const [docDate, setDocDate]     = useState("");
  const [materials, setMaterials] = useState<CmMaterial[]>([defaultMaterial()]);
  const [loading, setLoading]     = useState(false);
  const [error, setError]         = useState("");
  const [success, setSuccess]     = useState(false);

  const meisaiRef    = useRef<HTMLInputElement>(null);
  const timetableRef = useRef<HTMLInputElement>(null);

  const addMaterial    = () => setMaterials([...materials, defaultMaterial()]);
  const removeMaterial = (i: number) => setMaterials(materials.filter((_, idx) => idx !== i));
  const updateMaterial = (i: number, key: keyof CmMaterial, val: string | number) => {
    setMaterials(materials.map((m, idx) =>
      idx === i ? { ...m, [key]: val } : m
    ));
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!meisai) { setError("明細リストPDFを選択してください"); return; }

    setLoading(true);
    setError("");
    setSuccess(false);

    const form = new FormData();
    form.append("meisai", meisai);
    if (timetable) form.append("timetable", timetable);
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
      setSuccess(true);
    } catch (err) {
      setError(err instanceof Error ? err.message : "不明なエラー");
    } finally {
      setLoading(false);
    }
  };

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
            <h2 className="text-lg font-semibold text-slate-700 mb-4">📄 TV局資料 (PDF)</h2>
            <div className="space-y-4">
              <FileInput label="明細リスト（必須）" accept=".pdf" ref={meisaiRef}
                file={meisai} onChange={(f) => { setMeisai(f); setError(""); }} required />
              <FileInput label="タイムテーブル（任意）" accept=".pdf" ref={timetableRef}
                file={timetable} onChange={setTimetable} />
            </div>
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
                  onChange={(e) => setDocDate(e.target.value.replace(/-/g, "/"))}
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
                        value={mat.name} onChange={(v) => updateMaterial(i, "name", v)} />
                    </div>
                    <Field label="略号" placeholder="例: S, A, B"
                      value={mat.abbr} onChange={(v) => updateMaterial(i, "abbr", v.toUpperCase())} />
                    <Field label="使用期間" placeholder="例: 2/10〜2/23"
                      value={mat.period} onChange={(v) => updateMaterial(i, "period", v)} />
                    <SelectField label="素材種別" value={mat.material}
                      options={["OL", "XD", "F", "S"]} onChange={(v) => updateMaterial(i, "material", v)} />
                    <SelectField label="在送使廻" value={mat.status}
                      options={["送", "在", "廻", "使"]} onChange={(v) => updateMaterial(i, "status", v)} />
                  </div>
                </div>
              ))}
            </div>
          </section>

          {error && (
            <div className="bg-red-50 border border-red-200 text-red-700 rounded-xl px-4 py-3 text-sm">
              ⚠️ {error}
            </div>
          )}
          {success && (
            <div className="bg-green-50 border border-green-200 text-green-700 rounded-xl px-4 py-3 text-sm">
              ✅ スケジュール表を生成しました！ダウンロードを確認してください。
            </div>
          )}

          <button type="submit" disabled={loading}
            className="w-full py-3 rounded-xl font-semibold text-white text-base transition
              bg-blue-600 hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed shadow-sm">
            {loading ? "⏳ 生成中..." : "📥 スケジュール表を生成・ダウンロード"}
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

const FileInput = React.forwardRef<
  HTMLInputElement,
  { label: string; accept: string; file: File | null; onChange: (f: File | null) => void; required?: boolean }
>(({ label, accept, file, onChange, required }, ref) => (
  <div>
    <label className="block text-sm font-medium text-slate-600 mb-1">{label}</label>
    <label className="flex items-center gap-3 cursor-pointer border border-dashed border-slate-300 hover:border-blue-400 rounded-xl px-4 py-3 transition bg-slate-50 hover:bg-blue-50">
      <span className="text-2xl">📄</span>
      <span className="text-sm text-slate-500 truncate flex-1">
        {file ? file.name : "クリックしてPDFを選択"}
      </span>
      {file && (
        <span className="text-slate-400 hover:text-red-400 text-lg"
          onClick={(e) => {
            e.preventDefault();
            onChange(null);
            if (ref && "current" in ref && ref.current) ref.current.value = "";
          }}>×</span>
      )}
      <input ref={ref} type="file" accept={accept} required={required}
        className="hidden" onChange={(e) => onChange(e.target.files?.[0] ?? null)} />
    </label>
  </div>
));
FileInput.displayName = "FileInput";

function Field({ label, placeholder, value, onChange }: {
  label: string; placeholder: string; value: string; onChange: (v: string) => void;
}) {
  return (
    <div>
      <label className="block text-sm font-medium text-slate-600 mb-1">{label}</label>
      <input type="text" placeholder={placeholder} value={value}
        onChange={(e) => onChange(e.target.value)}
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
      <select value={value} onChange={(e) => onChange(e.target.value)}
        className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white">
        {options.map((o) => <option key={o} value={o}>{o}</option>)}
      </select>
    </div>
  );
}
