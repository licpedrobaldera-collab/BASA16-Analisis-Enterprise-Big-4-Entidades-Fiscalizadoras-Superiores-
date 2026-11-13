import React, { useState } from 'react';
import {
  FileUp, ShieldCheck, Scale, FileSpreadsheet, FileText,
  AlertTriangle, CheckCircle2, History, Search, Trash2, Loader2, ExternalLink
} from 'lucide-react';
import { parseRealFile, ParsedDocument } from './utils/fileParser';
import { exportForensicExcel } from './utils/excelExport';

const PHASES = [
  '1. Revisión Previa Acta de Lectura (Art. 68 Ley 18-24)',
  '2. Evaluación Informe Preliminar de Auditoría',
  '3. Evaluación de Réplicas / Escritos de Descargo',
  '4. Evaluación de Contraréplicas y Cierre',
  '5. Control de Calidad Informe Final Definitivo'
];

export default function App() {
  const [documents, setDocuments] = useState<ParsedDocument[]>([]);
  const [phase, setPhase] = useState<string>(PHASES[1]);
  const [entityName, setEntityName] = useState<string>('');
  const [auditPeriod, setAuditPeriod] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<any | null>(null);
  const [webSources, setWebSources] = useState<any[]>([]);
  const [iterations, setIterations] = useState<any[]>([]);
  const [activeTab, setActiveTab] = useState<'matriz' | 'capitulos' | 'directores' | 'ejecutivo' | 'historial'>('matriz');

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    setError(null);
    const files = e.target.files;
    if (!files || files.length === 0) return;

    try {
      const parsedList: ParsedDocument[] = [];
      for (let i = 0; i < files.length; i++) {
        const parsed = await parseRealFile(files[i]);
        parsedList.push(parsed);
      }
      setDocuments((prev) => [...prev, ...parsedList]);
    } catch (err: any) {
      setError(err.message || 'Error al extraer el contenido real del archivo.');
    }
    e.target.value = '';
  };

  const removeDoc = (id: string) => {
    setDocuments((prev) => prev.filter((d) => d.id !== id));
  };

  const runRealForensicAudit = async () => {
    if (documents.length === 0) {
      setError('Debe cargar al menos un archivo real (PDF, DOCX, XLSX, CSV o TXT) antes de ejecutar.');
      return;
    }
    setLoading(true);
    setError(null);

    try {
      const previousContext = iterations.length > 0 ? iterations[iterations.length - 1].analysis : null;
      const response = await fetch('/api/analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          documents,
          phase,
          entityName,
          auditPeriod,
          previousContext
        })
      });

      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.error || 'Error ejecutando el análisis forense.');
      }

      setResult(data.analysis);
      setWebSources(data.webSources || []);
      setIterations((prev) => [
        ...prev,
        {
          id: crypto.randomUUID(),
          timestamp: new Date().toLocaleString(),
          phase,
          files: documents.map((d) => ({ name: d.name, sha256: d.sha256 })),
          analysis: data.analysis
        }
      ]);
      setActiveTab('matriz');
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const downloadExecutiveWord = () => {
    if (!result) return;
    const html = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word'>
      <head><meta charset='utf-8'><title>Informe Ejecutivo Forense CCRD</title></head>
      <body style="font-family: Arial, sans-serif; line-height: 1.5;">
        <h1>DICTAMEN EJECUTIVO DE AUDITORÍA FORENSE Y SUPERVISIÓN (LEY 18-24 CCRD)</h1>
        <p><strong>Entidad Auditada:</strong> ${result.entidad}</p>
        <p><strong>Fase Procesal Evaluada:</strong> ${result.faseEvaluada}</p>
        <p><strong>Calificación Actual:</strong> ${result.puntuacionGlobal10}/10 (${result.porcentajeBlindaje}% de blindaje técnico-legal)</p>
        <h2>1. Resumen Ejecutivo</h2>
        <p>${result.resumenEjecutivo}</p>
        <h2>2. Evaluación Punto por Punto / Capítulos (Lo que Existe vs. Estándar 10/10)</h2>
        ${(result.evaluacionPorCapitulos || []).map((c: any) => `
          <h3>${c.capitulo} — Calificación: ${c.puntuacionActual10}/10 (${c.porcentajeActual}%)</h3>
          <p><strong>Lo que existe:</strong> ${c.loQueExiste}</p>
          <p><strong>Recomendación 10/10:</strong> ${c.recomendacion10de10}</p>
          <p><strong>Base Legal / Justificación:</strong> ${c.justificacionLegal}</p>
          <p><strong>Procedimiento Aplicado:</strong> ${c.procedimientoDeterminacion}</p>
        `).join('')}
        <h2>3. Segregación de Responsabilidad por Titular y Decreto Oficial</h2>
        ${(result.segregacionResponsabilidad || []).map((s: any) => `
          <p><strong>${s.director} (${s.cargo})</strong> | Período: ${s.periodoGestion} | Decreto y Fuente: ${s.decretoYFuente} | Imputado: RD$ ${Number(s.montoTotalImputadoRD || 0).toLocaleString()} | Descargado: RD$ ${Number(s.montoDescargadoRD || 0).toLocaleString()} | Persistente: RD$ ${Number(s.montoPersistenteRD || 0).toLocaleString()}</p>
        `).join('')}
      </body></html>`;
    const blob = new Blob(['\ufeff', html], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Informe_Ejecutivo_CCRD_${(result.entidad || 'Auditoria').replace(/[^a-zA-Z0-9]/g, '_')}.doc`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100">
      <header className="border-b border-slate-800 bg-slate-900/90 px-6 py-4">
        <div className="mx-auto flex max-w-7xl flex-wrap items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <ShieldCheck className="h-8 w-8 text-emerald-400" />
            <div>
              <h1 className="text-lg font-bold tracking-tight text-white">
                Sistema de Supervisión y Auditoría Forense — Ley 18-24 CCRD
              </h1>
              <p className="text-xs text-slate-400">
                Entorno Real de Producción · Sin Datos Simulados · Trazabilidad SHA-256 y Búsqueda Oficial de Decretos
              </p>
            </div>
          </div>
          {result && (
            <div className="flex items-center gap-3">
              <button
                onClick={() => exportForensicExcel(result, webSources, documents)}
                className="flex items-center gap-2 rounded-lg bg-emerald-600 px-4 py-2 text-sm font-medium text-white hover:bg-emerald-500"
              >
                <FileSpreadsheet className="h-4 w-4" />
                Exportar Matriz Excel (1 Col/Hallazgo)
              </button>
              <button
                onClick={downloadExecutiveWord}
                className="flex items-center gap-2 rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-500"
              >
                <FileText className="h-4 w-4" />
                Exportar Informe Word
              </button>
            </div>
          )}
        </div>
      </header>

      <main className="mx-auto max-w-7xl space-y-6 p-6">
        {/* PANEL DE CARGA REAL Y CONFIGURACIÓN PROCESAL */}
        <section className="rounded-xl border border-slate-800 bg-slate-900 p-6 shadow-lg">
          <h2 className="mb-4 flex items-center gap-2 text-base font-semibold text-emerald-400">
            <Scale className="h-5 w-5" />
            1. Carga de Expedientes Reales y Parámetros del Proceso (Multiformato)
          </h2>

          <div className="grid gap-4 md:grid-cols-3">
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-300">Fase Procesal a Evaluar</label>
              <select
                value={phase}
                onChange={(e) => setPhase(e.target.value)}
                className="w-full rounded-lg border border-slate-700 bg-slate-800 px-3 py-2 text-sm text-white"
              >
                {PHASES.map((p) => (
                  <option key={p} value={p}>{p}</option>
                ))}
              </select>
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-300">Entidad Auditada (Opcional: se extrae del archivo)</label>
              <input
                type="text"
                value={entityName}
                onChange={(e) => setEntityName(e.target.value)}
                placeholder="Ej. EDEESTE / Ministerio / Ayuntamiento..."
                className="w-full rounded-lg border border-slate-700 bg-slate-800 px-3 py-2 text-sm text-white"
              />
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-slate-300">Período(s) Auditado(s)</label>
              <input
                type="text"
                value={auditPeriod}
                onChange={(e) => setAuditPeriod(e.target.value)}
                placeholder="Ej. 2020-2024 (Multi-período)"
                className="w-full rounded-lg border border-slate-700 bg-slate-800 px-3 py-2 text-sm text-white"
              />
            </div>
          </div>

          <div className="mt-4">
            <label className="flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed border-slate-700 bg-slate-800/50 p-6 transition hover:border-emerald-500">
              <FileUp className="mb-2 h-8 w-8 text-emerald-400" />
              <span className="text-sm font-medium text-slate-200">
                Cargar archivos reales (PDF, DOCX, XLSX, CSV, TXT, JSON)
              </span>
              <span className="mt-1 text-xs text-slate-400">
                Soporta Acta de Lectura, Informe Preliminar, Réplicas, Contraréplicas y Anexos Financieros
              </span>
              <input
                type="file"
                multiple
                accept=".pdf,.docx,.xlsx,.xls,.csv,.txt,.json"
                onChange={handleFileUpload}
                className="hidden"
              />
            </label>
          </div>

          {documents.length > 0 && (
            <div className="mt-4 space-y-2">
              <p className="text-xs font-semibold uppercase tracking-wider text-slate-400">
                Expediente en Cadena de Custodia ({documents.length} archivo(s) real(es))
              </p>
              {documents.map((doc) => (
                <div key={doc.id} className="flex items-center justify-between rounded-lg border border-slate-800 bg-slate-950 px-4 py-2 text-xs">
                  <div>
                    <span className="font-semibold text-white">{doc.name}</span>
                    <span className="ml-2 rounded bg-slate-800 px-2 py-0.5 text-emerald-400">{doc.type}</span>
                    <span className="ml-2 text-slate-400">({doc.content.length.toLocaleString()} caracteres extraídos)</span>
                    <div className="mt-0.5 font-mono text-[11px] text-slate-500">SHA-256: {doc.sha256}</div>
                  </div>
                  <button onClick={() => removeDoc(doc.id)} className="text-rose-400 hover:text-rose-300">
                    <Trash2 className="h-4 w-4" />
                  </button>
                </div>
              ))}
            </div>
          )}

          {error && (
            <div className="mt-4 flex items-center gap-2 rounded-lg border border-rose-800 bg-rose-950/60 p-3 text-sm text-rose-200">
              <AlertTriangle className="h-5 w-5 shrink-0 text-rose-400" />
              <span>{error}</span>
            </div>
          )}

          <div className="mt-5 flex justify-end">
            <button
              onClick={runRealForensicAudit}
              disabled={loading || documents.length === 0}
              className="flex items-center gap-2 rounded-lg bg-emerald-600 px-6 py-2.5 text-sm font-semibold text-white shadow hover:bg-emerald-500 disabled:cursor-not-allowed disabled:opacity-50"
            >
              {loading ? <Loader2 className="h-4 w-4 animate-spin" /> : <Search className="h-4 w-4" />}
              {loading ? 'Ejecutando Análisis Forense y Búsqueda Oficial...' : 'Ejecutar Evaluación Forense Real (1/10 -> 10/10)'}
            </button>
          </div>
        </section>

        {/* RESULTADOS EN TIEMPO REAL */}
        {result && (
          <section className="space-y-6">
            <div className="grid gap-4 md:grid-cols-4">
              <div className="rounded-xl border border-slate-800 bg-slate-900 p-4">
                <p className="text-xs text-slate-400">Entidad Auditada</p>
                <p className="mt-1 text-base font-bold text-white">{result.entidad}</p>
              </div>
              <div className="rounded-xl border border-slate-800 bg-slate-900 p-4">
                <p className="text-xs text-slate-400">Calificación Actual</p>
                <p className="mt-1 text-2xl font-bold text-amber-400">{result.puntuacionGlobal10} / 10</p>
              </div>
              <div className="rounded-xl border border-slate-800 bg-slate-900 p-4">
                <p className="text-xs text-slate-400">Blindaje Técnico-Legal</p>
                <p className="mt-1 text-2xl font-bold text-emerald-400">{result.porcentajeBlindaje}%</p>
              </div>
              <div className="rounded-xl border border-slate-800 bg-slate-900 p-4">
                <p className="text-xs text-slate-400">Hallazgos Reales Extraídos</p>
                <p className="mt-1 text-2xl font-bold text-blue-400">{(result.hallazgos || []).length}</p>
              </div>
            </div>

            <div className="flex flex-wrap gap-2 border-b border-slate-800 pb-2">
              {[
                { id: 'matriz', label: 'Matriz Transpuesta (1 Columna por Hallazgo)' },
                { id: 'capitulos', label: 'Evaluación Punto por Punto (Existe vs 10/10)' },
                { id: 'directores', label: 'Segregación por Directores y Decretos' },
                { id: 'ejecutivo', label: 'Informe Ejecutivo y Fuentes Oficiales' },
                { id: 'historial', label: `Seguimiento Multi-Carga (${iterations.length})` }
              ].map((t) => (
                <button
                  key={t.id}
                  onClick={() => setActiveTab(t.id as any)}
                  className={`rounded-lg px-4 py-2 text-xs font-semibold transition ${
                    activeTab === t.id ? 'bg-emerald-600 text-white' : 'bg-slate-900 text-slate-400 hover:text-white'
                  }`}
                >
                  {t.label}
                </button>
              ))}
            </div>

            {activeTab === 'matriz' && (
              <div className="overflow-x-auto rounded-xl border border-slate-800 bg-slate-900">
                <table className="w-full border-collapse text-left text-xs">
                  <thead>
                    <tr className="border-b border-slate-800 bg-slate-950">
                      <th className="min-w-[230px] p-3 font-bold text-emerald-400">Criterio Forense</th>
                      {(result.hallazgos || []).map((h: any, idx: number) => (
                        <th key={idx} className="min-w-[320px] border-l border-slate-800 p-3 font-bold text-white">
                          {h.codigo || `Hallazgo ${idx + 1}`}: {h.titulo}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800">
                    {[
                      ['Período de la Incidencia', 'periodoIncidencia'],
                      ['Director / Titular Responsable', 'directorResponsable'],
                      ['Decreto de Nombramiento', 'decretoNombramiento'],
                      ['Fuente Oficial / Ubicación Decreto', 'fuenteOficialDecreto'],
                      ['Puntaje Actual (1/10)', 'puntuacionActual10'],
                      ['Porcentaje Actual (%)', 'porcentajeActual'],
                      ['Lo que Existe en el Informe', 'loQueExisteEnInforme'],
                      ['Brecha Técnica / Legal', 'brechaDetectada'],
                      ['Recomendación Exacta 10/10', 'recomendacionAjuste10de10'],
                      ['Justificación / Base Legal (Ley 18-24)', 'baseLegalNormativa'],
                      ['Procedimiento Forense', 'procedimientoForense'],
                      ['Monto Original (RD$)', 'montoOriginalRD'],
                      ['Monto Descargado Validado (RD$)', 'montoDescargadoRD'],
                      ['Monto Persistente (RD$)', 'montoPersistenteRD'],
                      ['Estado Procesal / Réplica', 'estadoReplica']
                    ].map(([label, key]) => (
                      <tr key={key} className="align-top hover:bg-slate-800/40">
                        <td className="bg-slate-950/60 p-3 font-semibold text-slate-300">{label}</td>
                        {(result.hallazgos || []).map((h: any, idx: number) => (
                          <td key={idx} className="border-l border-slate-800 p-3 text-slate-200">
                            {typeof h[key] === 'number' && key.includes('monto')
                              ? `RD$ ${h[key].toLocaleString()}`
                              : String(h[key] ?? '')}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {activeTab === 'capitulos' && (
              <div className="space-y-4">
                {(result.evaluacionPorCapitulos || []).map((cap: any, i: number) => (
                  <div key={i} className="rounded-xl border border-slate-800 bg-slate-900 p-5">
                    <div className="flex items-center justify-between">
                      <h3 className="text-sm font-bold text-white">{cap.capitulo}</h3>
                      <span className="rounded-full bg-amber-500/20 px-3 py-1 text-xs font-bold text-amber-300">
                        {cap.puntuacionActual10}/10 ({cap.porcentajeActual}%)
                      </span>
                    </div>
                    <div className="mt-3 grid gap-4 md:grid-cols-2">
                      <div className="rounded-lg border border-slate-800 bg-slate-950 p-3 text-xs">
                        <p className="font-semibold text-rose-400">Lo que Existe Actualmente:</p>
                        <p className="mt-1 text-slate-300">{cap.loQueExiste}</p>
                      </div>
                      <div className="rounded-lg border border-emerald-900/50 bg-emerald-950/20 p-3 text-xs">
                        <p className="font-semibold text-emerald-400">Recomendación Exacta para Nivel 10/10:</p>
                        <p className="mt-1 text-slate-200">{cap.recomendacion10de10}</p>
                      </div>
                    </div>
                    <div className="mt-3 grid gap-4 md:grid-cols-2 text-xs">
                      <div>
                        <span className="font-semibold text-blue-400">Justificación / Base Legal: </span>
                        <span className="text-slate-300">{cap.justificacionLegal}</span>
                      </div>
                      <div>
                        <span className="font-semibold text-purple-400">Procedimiento Determinación: </span>
                        <span className="text-slate-300">{cap.procedimientoDeterminacion}</span>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {activeTab === 'directores' && (
              <div className="overflow-x-auto rounded-xl border border-slate-800 bg-slate-900">
                <table className="w-full text-left text-xs">
                  <thead className="border-b border-slate-800 bg-slate-950 text-slate-300">
                    <tr>
                      <th className="p-3">Director / Titular</th>
                      <th className="p-3">Cargo</th>
                      <th className="p-3">Decreto y Fuente Oficial</th>
                      <th className="p-3">Período</th>
                      <th className="p-3">Hallazgos</th>
                      <th className="p-3">Imputado (RD$)</th>
                      <th className="p-3">Descargado (RD$)</th>
                      <th className="p-3">Persistente (RD$)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800">
                    {(result.segregacionResponsabilidad || []).map((s: any, i: number) => (
                      <tr key={i}>
                        <td className="p-3 font-semibold text-white">{s.director}</td>
                        <td className="p-3 text-slate-300">{s.cargo}</td>
                        <td className="p-3 text-emerald-400">{s.decretoYFuente}</td>
                        <td className="p-3 text-slate-300">{s.periodoGestion}</td>
                        <td className="p-3 text-center">{s.cantidadHallazgos}</td>
                        <td className="p-3">RD$ {Number(s.montoTotalImputadoRD || 0).toLocaleString()}</td>
                        <td className="p-3 text-emerald-400">RD$ {Number(s.montoDescargadoRD || 0).toLocaleString()}</td>
                        <td className="p-3 font-bold text-rose-400">RD$ {Number(s.montoPersistenteRD || 0).toLocaleString()}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {activeTab === 'ejecutivo' && (
              <div className="space-y-4 rounded-xl border border-slate-800 bg-slate-900 p-6 text-sm">
                <h3 className="text-base font-bold text-white">Informe Ejecutivo de Supervisión Forense</h3>
                <p className="whitespace-pre-line text-slate-300">{result.resumenEjecutivo}</p>

                {webSources.length > 0 && (
                  <div className="mt-4 border-t border-slate-800 pt-4">
                    <h4 className="mb-2 text-xs font-bold uppercase text-emerald-400">
                      Fuentes Oficiales Consultadas en Línea (Decretos / Gacetas)
                    </h4>
                    <ul className="space-y-1 text-xs">
                      {webSources.map((src, i) => (
                        <li key={i}>
                          <a href={src.uri} target="_blank" rel="noreferrer" className="inline-flex items-center gap-1 text-blue-400 hover:underline">
                            <ExternalLink className="h-3 w-3" />
                            {src.title || src.uri}
                          </a>
                        </li>
                      ))}
                    </ul>
                  </div>
                )}
              </div>
            )}

            {activeTab === 'historial' && (
              <div className="space-y-3">
                {iterations.map((it, idx) => (
                  <div key={it.id} className="rounded-xl border border-slate-800 bg-slate-900 p-4 text-xs">
                    <div className="flex items-center justify-between">
                      <span className="font-bold text-emerald-400">Iteración #{idx + 1} — {it.phase}</span>
                      <span className="text-slate-400">{it.timestamp}</span>
                    </div>
                    <p className="mt-1 text-slate-300">
                      Calificación: <strong>{it.analysis.puntuacionGlobal10}/10</strong> ({it.analysis.porcentajeBlindaje}%) | Hallazgos: {(it.analysis.hallazgos || []).length}
                    </p>
                  </div>
                ))}
              </div>
            )}
          </section>
        )}
      </main>
    </div>
  );
}
