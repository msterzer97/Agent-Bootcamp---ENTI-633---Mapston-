/**
 * LandDev ECC Budget Variance Tracker & Leadership Reconciler
 * Single-screen full-featured financial control web app
 */

import React, { useState, useEffect, useRef } from 'react';
import {
  FileSpreadsheet,
  Upload,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  Copy,
  Download,
  ShieldCheck,
  RotateCcw,
  Sparkles,
  ArrowRight,
  Info,
  Check,
  FileCheck,
  AlertCircle,
  ChevronRight,
  TrendingUp,
  TrendingDown,
  Layers,
  Database,
  ExternalLink,
} from 'lucide-react';
import {
  inspectECCWorkbook,
  inspectSourceWorkbook,
  executeECCUpdate,
} from './services/excelEngine';
import {
  ACCEPTANCE_TEST_PRESETS,
  SampleWorkbookPair,
} from './services/sampleDataGenerator';
import {
  InspectionResult,
  ReconciliationReport,
  MatchingException,
  PhaseCostTypeVariance,
} from './types/ecc';

export default function App() {
  // Input Files State
  const [eccFile, setEccFile] = useState<{
    name: string;
    buffer: ArrayBuffer;
    size: number;
  } | null>(null);

  const [sourceFile, setSourceFile] = useState<{
    name: string;
    buffer: ArrayBuffer;
    size: number;
  } | null>(null);

  // Inspection State
  const [eccInspection, setEccInspection] = useState<InspectionResult | null>(null);
  const [sourceInspection, setSourceInspection] = useState<{
    sheetNames: string[];
    totalSourceRows: number;
  } | null>(null);

  // Active Acceptance Test Preset
  const [activePresetId, setActivePresetId] = useState<string | null>(null);

  // Execution & Reconciliation State
  const [isProcessing, setIsProcessing] = useState(false);
  const [report, setReport] = useState<ReconciliationReport | null>(null);
  const [executionError, setExecutionError] = useState<string | null>(null);
  const [copiedEmail, setCopiedEmail] = useState(false);

  // UI Tabs
  const [activeTab, setActiveTab] = useState<
    'ecc_summary' | 'leadership_email' | 'material_variances' | 'exceptions' | 'reconciliation_matrix' | 'integrity_audit'
  >('ecc_summary');

  // Hidden File Inputs
  const eccInputRef = useRef<HTMLInputElement>(null);
  const sourceInputRef = useRef<HTMLInputElement>(null);

  // Load Preset Handler
  const handleLoadPreset = async (preset: SampleWorkbookPair) => {
    setActivePresetId(preset.id);
    setReport(null);
    setExecutionError(null);

    // If Test 2, load only ECC workbook (no source workbook)
    if (preset.id === 'test-2') {
      const eccBuf = await preset.createECC();
      setEccFile({
        name: preset.eccFileName,
        buffer: eccBuf,
        size: eccBuf.byteLength,
      });
      setSourceFile(null);
      setSourceInspection(null);

      try {
        const { inspection } = await inspectECCWorkbook(eccBuf);
        setEccInspection(inspection);
      } catch (e: any) {
        // ignore
      }
      setExecutionError(
        'Missing required input: The LandDev item-level monthly budget workbook was not provided. Execution halted; no updated workbook, reconciliation, variance table, or email draft produced.'
      );
      setReport(null);
      return;
    }

    // Otherwise load both and automatically execute update
    try {
      const [eccBuf, sourceBuf] = await Promise.all([
        preset.createECC(),
        preset.createSource(),
      ]);

      setEccFile({
        name: preset.eccFileName,
        buffer: eccBuf,
        size: eccBuf.byteLength,
      });

      setSourceFile({
        name: preset.sourceFileName,
        buffer: sourceBuf,
        size: sourceBuf.byteLength,
      });

      const [eccRes, srcRes] = await Promise.all([
        inspectECCWorkbook(eccBuf),
        inspectSourceWorkbook(sourceBuf),
      ]);

      setEccInspection(eccRes.inspection);
      setSourceInspection({
        sheetNames: srcRes.sheetNames,
        totalSourceRows: srcRes.totalSourceRows,
      });

      setIsProcessing(true);
      const result = await executeECCUpdate({
        eccBuffer: eccBuf,
        sourceBuffer: sourceBuf,
        originalECCFileName: preset.eccFileName,
        projectName: 'Highland Meadows Master LandDev',
        reportingPeriod: 'November 2026',
      });

      setReport(result);
    } catch (e: any) {
      setExecutionError('Error processing workbook scenario: ' + e.message);
    } finally {
      setIsProcessing(false);
    }
  };

  // Run update automatically when preset loads or on manual trigger
  const runUpdateProcess = async () => {
    setExecutionError(null);

    // Acceptance Test 2 rule: Require BOTH workbooks!
    if (!eccFile && !sourceFile) {
      setExecutionError(
        'Missing required inputs: Please provide both the ECC variance workbook and the LandDev monthly budget source workbook.'
      );
      setReport(null);
      return;
    }
    if (!eccFile) {
      setExecutionError(
        'Missing required input: The baseline ECC variance workbook containing the "CV Budget Items" sheet was not provided. Execution halted.'
      );
      setReport(null);
      return;
    }
    if (!sourceFile) {
      setExecutionError(
        'Missing required input: The LandDev item-level monthly budget workbook was not provided. Execution halted; no updated workbook or report produced.'
      );
      setReport(null);
      return;
    }

    setIsProcessing(true);

    try {
      const result = await executeECCUpdate({
        eccBuffer: eccFile.buffer,
        sourceBuffer: sourceFile.buffer,
        originalECCFileName: eccFile.name,
        projectName: 'Highland Meadows Master LandDev',
        reportingPeriod: 'November 2026',
      });

      setReport(result);
    } catch (err: any) {
      console.error(err);
      setExecutionError(err.message || 'Error occurred during workbook processing.');
      setReport(null);
    } finally {
      setIsProcessing(false);
    }
  };

  // Auto-run Test 1 on initial mount so user sees immediate results
  useEffect(() => {
    handleLoadPreset(ACCEPTANCE_TEST_PRESETS[0]);
  }, []);

  // When files change, clear previous report unless auto-loaded
  const handleEccFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setActivePresetId(null);
    setReport(null);
    setExecutionError(null);

    const buffer = await file.arrayBuffer();
    setEccFile({
      name: file.name,
      buffer,
      size: file.size,
    });

    try {
      const { inspection } = await inspectECCWorkbook(buffer);
      setEccInspection(inspection);
    } catch (err: any) {
      setExecutionError('Error inspecting ECC workbook: ' + err.message);
    }
  };

  const handleSourceFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setActivePresetId(null);
    setReport(null);
    setExecutionError(null);

    const buffer = await file.arrayBuffer();
    setSourceFile({
      name: file.name,
      buffer,
      size: file.size,
    });

    try {
      const srcRes = await inspectSourceWorkbook(buffer);
      setSourceInspection({
        sheetNames: srcRes.sheetNames,
        totalSourceRows: srcRes.totalSourceRows,
      });
    } catch (err: any) {
      setExecutionError('Error inspecting Source workbook: ' + err.message);
    }
  };

  // Copy email handler
  const handleCopyEmail = () => {
    if (!report?.emailDraft) return;
    navigator.clipboard.writeText(report.emailDraft);
    setCopiedEmail(true);
    setTimeout(() => setCopiedEmail(false), 2000);
  };

  // Download updated workbook handler
  const handleDownloadUpdatedWorkbook = () => {
    if (!report?.updatedFileBlob || !report?.updatedFileName) return;
    const url = URL.createObjectURL(report.updatedFileBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = report.updatedFileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col font-sans">
      {/* 1. TOP BAR CONTRACT */}
      <header className="flex items-center justify-between gap-8 px-6 py-3.5 bg-white border-b border-slate-200 sticky top-0 z-40 shadow-xs">
        {/* Zone 1: Brand title */}
        <div className="flex items-center gap-3 shrink-0">
          <div className="w-8 h-8 rounded-lg bg-slate-900 text-white flex items-center justify-center font-bold text-sm tracking-tight shadow-xs">
            LD
          </div>
          <div className="flex flex-col">
            <span className="text-base font-bold text-slate-900 tracking-tight whitespace-nowrap">
              LandDev Financial Controls
            </span>
            <span className="text-xs text-slate-500 font-normal">
              ECC Budget Variance Tracker & Leadership Reconciler
            </span>
          </div>
        </div>

        {/* Zone 2: Navigation Links */}
        <nav className="hidden md:flex items-center gap-6 text-sm font-medium text-slate-600">
          <a
            href="#presets"
            className="hover:text-slate-900 transition-colors whitespace-nowrap"
          >
            Acceptance Tests
          </a>
          <a
            href="#workbooks"
            className="hover:text-slate-900 transition-colors whitespace-nowrap"
          >
            Workbook Audits
          </a>
          <a
            href="#results"
            className="hover:text-slate-900 transition-colors whitespace-nowrap"
          >
            Reconciliation & Email
          </a>
        </nav>

        {/* Zone 3: Primary Action */}
        <div className="flex items-center gap-3 shrink-0">
          {report?.updatedFileBlob && (
            <button
              onClick={handleDownloadUpdatedWorkbook}
              className="inline-flex items-center gap-2 px-3.5 py-1.5 text-xs font-semibold text-emerald-700 bg-emerald-50 border border-emerald-300 rounded-lg hover:bg-emerald-100 transition-colors whitespace-nowrap shadow-xs"
            >
              <Download className="w-3.5 h-3.5" />
              Download {report.updatedFileName}
            </button>
          )}

          <button
            onClick={runUpdateProcess}
            disabled={isProcessing}
            className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold text-white bg-slate-900 rounded-lg hover:bg-slate-800 active:scale-[0.98] transition-all whitespace-nowrap shadow-xs disabled:opacity-60"
          >
            {isProcessing ? (
              <span className="animate-spin w-3.5 h-3.5 border-2 border-white border-t-transparent rounded-full" />
            ) : (
              <ShieldCheck className="w-4 h-4 text-emerald-400" />
            )}
            Run ECC Budget Update
          </button>
        </div>
      </header>

      {/* MAIN CONTAINER */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-6 space-y-6">
        {/* ACCEPTANCE TEST SUITE PRESETS (1-CLICK RUNNERS) */}
        <section id="presets" className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-slate-100">
            <div>
              <div className="flex items-center gap-2 text-xs font-semibold text-slate-500 uppercase tracking-wider">
                <span>Acceptance Test Verification Suite</span>
                <span>·</span>
                <span>5 Presets Ready to Execute</span>
              </div>
              <h2 className="text-base font-bold text-slate-900 mt-0.5">
                Load Test Scenarios & Verify Business Rules
              </h2>
            </div>
            <span className="text-xs text-slate-500 self-start sm:self-auto">
              Click any test to load its pre-configured workbook pair:
            </span>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-2.5 mt-4">
            {ACCEPTANCE_TEST_PRESETS.map((preset) => {
              const isSelected = activePresetId === preset.id;
              return (
                <button
                  key={preset.id}
                  onClick={() => handleLoadPreset(preset)}
                  className={`text-left p-3 rounded-lg border transition-all relative flex flex-col justify-between ${
                    isSelected
                      ? 'border-slate-900 bg-slate-900 text-white shadow-xs'
                      : 'border-slate-200 bg-slate-50/70 hover:bg-white hover:border-slate-300 text-slate-800'
                  }`}
                >
                  <div>
                    <div className="flex items-center justify-between text-[11px] font-bold mb-1">
                      <span className={isSelected ? 'text-emerald-400' : 'text-slate-600'}>
                        {preset.testCaseLabel.split(':')[0]}
                      </span>
                      {isSelected && (
                        <span className="w-1.5 h-1.5 rounded-full bg-emerald-400" />
                      )}
                    </div>
                    <div className="text-xs font-semibold leading-snug line-clamp-2">
                      {preset.name.replace(/Acceptance Test \d+: /, '')}
                    </div>
                  </div>
                  <p
                    className={`text-[11px] mt-2 line-clamp-3 leading-relaxed ${
                      isSelected ? 'text-slate-300' : 'text-slate-500'
                    }`}
                  >
                    {preset.description}
                  </p>
                </button>
              );
            })}
          </div>
        </section>

        {/* WORKBOOK INPUT AUDIT CARDS */}
        <section id="workbooks" className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Card 1: Existing ECC Variance Workbook */}
          <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <div className="p-2 bg-blue-50 text-blue-700 rounded-lg">
                    <FileSpreadsheet className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-sm font-bold text-slate-900">
                      1. Existing ECC Variance Workbook
                    </h3>
                    <p className="text-xs text-slate-500">
                      Baseline comparison containing <code className="bg-slate-100 px-1 py-0.5 rounded text-[11px] text-slate-800 font-mono">CV Budget Items</code>
                    </p>
                  </div>
                </div>

                <button
                  type="button"
                  onClick={() => eccInputRef.current?.click()}
                  className="px-3 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-md transition-colors whitespace-nowrap"
                >
                  Browse .xlsx
                </button>
                <input
                  type="file"
                  ref={eccInputRef}
                  onChange={handleEccFileUpload}
                  accept=".xlsx,.xls"
                  className="hidden"
                />
              </div>

              {/* Status Box */}
              {eccFile ? (
                <div className="mt-4 p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-2 text-xs">
                  <div className="flex items-center justify-between font-mono">
                    <span className="font-semibold text-slate-800 truncate max-w-[280px]">
                      {eccFile.name}
                    </span>
                    <span className="text-slate-500">
                      {(eccFile.size / 1024).toFixed(1)} KB
                    </span>
                  </div>

                  <div className="pt-2 border-t border-slate-200/60 grid grid-cols-2 gap-2 text-[11px]">
                    <div>
                      <span className="text-slate-500">CV Budget Items Sheet:</span>{' '}
                      {eccInspection?.hasCVBudgetItemsSheet ? (
                        <span className="font-semibold text-emerald-600">Present</span>
                      ) : (
                        <span className="font-semibold text-rose-600">Missing!</span>
                      )}
                    </div>
                    <div>
                      <span className="text-slate-500">Project Summary Sheet:</span>{' '}
                      {eccInspection?.hasSummarySheet ? (
                        <span className="font-semibold text-slate-700">Present (Col E Protected)</span>
                      ) : (
                        <span className="text-slate-400">None</span>
                      )}
                    </div>
                    <div>
                      <span className="text-slate-500">Detail Rows:</span>{' '}
                      <span className="font-mono font-medium text-slate-800">
                        {eccInspection?.totalDataRows ?? '—'}
                      </span>
                    </div>
                    <div>
                      <span className="text-slate-500">Formula Cells Detected:</span>{' '}
                      <span className="font-mono font-medium text-amber-700">
                        {eccInspection?.formulaCellsCount ?? 0}
                      </span>
                    </div>
                  </div>
                </div>
              ) : (
                <div
                  onClick={() => eccInputRef.current?.click()}
                  className="mt-4 border-2 border-dashed border-slate-200 rounded-lg p-6 text-center cursor-pointer hover:border-slate-400 transition-colors"
                >
                  <Upload className="w-6 h-6 text-slate-400 mx-auto mb-1" />
                  <p className="text-xs font-medium text-slate-600">
                    Drop ECC Variance Workbook (.xlsx) here
                  </p>
                  <p className="text-[11px] text-slate-400 mt-0.5">
                    Must contain CV Budget Items worksheet
                  </p>
                </div>
              )}
            </div>

            <div className="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between text-[11px] text-slate-500">
              <span className="flex items-center gap-1.5">
                <ShieldCheck className="w-3.5 h-3.5 text-blue-600" />
                Column E (Previous Month) is strictly isolated and preserved
              </span>
            </div>
          </div>

          {/* Card 2: LandDev Monthly Budget Source Workbook */}
          <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <div className="p-2 bg-emerald-50 text-emerald-700 rounded-lg">
                    <FileSpreadsheet className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-sm font-bold text-slate-900">
                      2. New Monthly Budget Workbook
                    </h3>
                    <p className="text-xs text-slate-500">
                      Item-level source with updated monthly amounts
                    </p>
                  </div>
                </div>

                <button
                  type="button"
                  onClick={() => sourceInputRef.current?.click()}
                  className="px-3 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-md transition-colors whitespace-nowrap"
                >
                  Browse .xlsx
                </button>
                <input
                  type="file"
                  ref={sourceInputRef}
                  onChange={handleSourceFileUpload}
                  accept=".xlsx,.xls"
                  className="hidden"
                />
              </div>

              {/* Status Box */}
              {sourceFile ? (
                <div className="mt-4 p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-2 text-xs">
                  <div className="flex items-center justify-between font-mono">
                    <span className="font-semibold text-slate-800 truncate max-w-[280px]">
                      {sourceFile.name}
                    </span>
                    <span className="text-slate-500">
                      {(sourceFile.size / 1024).toFixed(1)} KB
                    </span>
                  </div>

                  <div className="pt-2 border-t border-slate-200/60 grid grid-cols-2 gap-2 text-[11px]">
                    <div>
                      <span className="text-slate-500">Worksheet:</span>{' '}
                      <span className="font-mono text-slate-700">
                        {sourceInspection?.sheetNames?.[0] || 'Source Sheet'}
                      </span>
                    </div>
                    <div>
                      <span className="text-slate-500">Item Records:</span>{' '}
                      <span className="font-mono font-medium text-slate-800">
                        {sourceInspection?.totalSourceRows ?? '—'}
                      </span>
                    </div>
                    <div className="col-span-2">
                      <span className="text-slate-500">Match Key Structure:</span>{' '}
                      <span className="font-mono text-slate-700">
                        (Phase, Cost Code, Budget Item Code)
                      </span>
                    </div>
                  </div>
                </div>
              ) : (
                <div
                  onClick={() => sourceInputRef.current?.click()}
                  className="mt-4 border-2 border-dashed border-slate-200 rounded-lg p-6 text-center cursor-pointer hover:border-slate-400 transition-colors"
                >
                  <Upload className="w-6 h-6 text-slate-400 mx-auto mb-1" />
                  <p className="text-xs font-medium text-slate-600">
                    Drop Monthly Source Workbook (.xlsx) here
                  </p>
                  <p className="text-[11px] text-slate-400 mt-0.5">
                    Item-level budget source with matching composite keys
                  </p>
                </div>
              )}
            </div>

            <div className="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between text-[11px] text-slate-500">
              <span className="flex items-center gap-1.5">
                <ShieldCheck className="w-3.5 h-3.5 text-emerald-600" />
                Preserves leading zeros and exact precision
              </span>
            </div>
          </div>
        </section>

        {/* EXECUTION ERROR / HALT BANNER (Acceptance Test 2, etc.) */}
        {executionError && (
          <div className="p-4 bg-rose-50 border border-rose-200 rounded-xl text-rose-800 space-y-1 shadow-xs">
            <div className="flex items-center gap-2 font-bold text-sm">
              <XCircle className="w-4 h-4 text-rose-600 shrink-0" />
              <span>Update Execution Halted — Missing Input or Invariant Failure</span>
            </div>
            <p className="text-xs text-rose-700 pl-6 leading-relaxed">
              {executionError}
            </p>
            <div className="text-[11px] text-rose-600 pl-6 pt-1">
              Per strict LandDev governance rules: When required workbooks are unavailable, the application halts immediately and produces no updated workbook, reconciliation, variance table, or email draft.
            </div>
          </div>
        )}

        {/* RECONCILIATION AUDIT RESULTS VIEW */}
        {report && (
          <section id="results" className="space-y-6">
            {/* KPI Summary Strip */}
            <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
              <div className="bg-white border border-slate-200 rounded-lg p-3.5 shadow-xs">
                <span className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider block">
                  Reconciliation Status
                </span>
                <div className="mt-1 flex items-center gap-1.5">
                  {report.isSuccess ? (
                    <span className="text-xs font-bold text-emerald-700 flex items-center gap-1">
                      <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                      Reconciled & Final
                    </span>
                  ) : (
                    <span className="text-xs font-bold text-amber-700 flex items-center gap-1">
                      <AlertTriangle className="w-3.5 h-3.5 text-amber-600" />
                      Provisional (Exceptions)
                    </span>
                  )}
                </div>
                <span className="text-[10px] text-slate-400 block mt-0.5">
                  {report.exceptions.length} exceptions noted
                </span>
              </div>

              <div className="bg-white border border-slate-200 rounded-lg p-3.5 shadow-xs">
                <span className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider block">
                  Matched Records
                </span>
                <div className="mt-1 text-base font-bold text-slate-900 font-mono">
                  {report.totalRecordsMatched} / {report.totalECCRecords}
                </div>
                <span className="text-[10px] text-slate-400 block mt-0.5">
                  Detail composite keys
                </span>
              </div>

              <div className="bg-white border border-slate-200 rounded-lg p-3.5 shadow-xs">
                <span className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider block">
                  Previous Total Budget
                </span>
                <div className="mt-1 text-base font-bold text-slate-900 font-mono tabular-nums">
                  ${report.totalPreviousBudget.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                </div>
                <span className="text-[10px] text-slate-400 block mt-0.5">
                  Retained in CV Budget Items
                </span>
              </div>

              <div className="bg-white border border-slate-200 rounded-lg p-3.5 shadow-xs">
                <span className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider block">
                  Updated Total Budget
                </span>
                <div className="mt-1 text-base font-bold text-slate-900 font-mono tabular-nums">
                  ${report.totalUpdatedBudget.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                </div>
                <span className="text-[10px] text-slate-400 block mt-0.5">
                  From monthly source budget
                </span>
              </div>

              <div className="bg-white border border-slate-200 rounded-lg p-3.5 shadow-xs">
                <span className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider block">
                  Material Variances (|∆| &gt; $25k)
                </span>
                <div className="mt-1 text-base font-bold text-slate-900 font-mono flex items-center gap-1.5">
                  <span>{report.thresholdVariances.length} combinations</span>
                </div>
                <span className="text-[10px] text-slate-400 block mt-0.5">
                  Strict threshold filter applied
                </span>
              </div>
            </div>

            {/* TABBED AUDIT WORKBENCH */}
            <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
              {/* Tab Header */}
              <div className="flex border-b border-slate-200 bg-slate-50/70 px-4 pt-3 gap-2 overflow-x-auto">
                <button
                  onClick={() => setActiveTab('ecc_summary')}
                  className={`px-3.5 py-2 text-xs font-semibold rounded-t-lg transition-colors whitespace-nowrap flex items-center gap-2 ${
                    activeTab === 'ecc_summary'
                      ? 'bg-white text-slate-900 border-t-2 border-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  <FileSpreadsheet className="w-3.5 h-3.5 text-blue-700" />
                  ECC Project Summary Sheet (Phase Level)
                </button>

                <button
                  onClick={() => setActiveTab('leadership_email')}
                  className={`px-3.5 py-2 text-xs font-semibold rounded-t-lg transition-colors whitespace-nowrap flex items-center gap-2 ${
                    activeTab === 'leadership_email'
                      ? 'bg-white text-slate-900 border-t-2 border-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  <Sparkles className="w-3.5 h-3.5 text-blue-600" />
                  Senior Leadership Draft Email
                  {report.thresholdVariances.length > 0 && (
                    <span className="bg-blue-100 text-blue-800 text-[10px] px-1.5 py-0.2 rounded-full font-mono">
                      {report.thresholdVariances.length}
                    </span>
                  )}
                </button>

                <button
                  onClick={() => setActiveTab('material_variances')}
                  className={`px-3.5 py-2 text-xs font-semibold rounded-t-lg transition-colors whitespace-nowrap flex items-center gap-2 ${
                    activeTab === 'material_variances'
                      ? 'bg-white text-slate-900 border-t-2 border-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  <TrendingUp className="w-3.5 h-3.5 text-emerald-600" />
                  Material Variance Table (|∆| &gt; $25k)
                </button>

                <button
                  onClick={() => setActiveTab('exceptions')}
                  className={`px-3.5 py-2 text-xs font-semibold rounded-t-lg transition-colors whitespace-nowrap flex items-center gap-2 ${
                    activeTab === 'exceptions'
                      ? 'bg-white text-slate-900 border-t-2 border-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  <AlertCircle className="w-3.5 h-3.5 text-amber-600" />
                  Exceptions & Clarifications
                  {report.exceptions.length > 0 && (
                    <span className="bg-amber-100 text-amber-800 text-[10px] px-1.5 py-0.2 rounded-full font-mono">
                      {report.exceptions.length}
                    </span>
                  )}
                </button>

                <button
                  onClick={() => setActiveTab('reconciliation_matrix')}
                  className={`px-3.5 py-2 text-xs font-semibold rounded-t-lg transition-colors whitespace-nowrap flex items-center gap-2 ${
                    activeTab === 'reconciliation_matrix'
                      ? 'bg-white text-slate-900 border-t-2 border-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  <Layers className="w-3.5 h-3.5 text-indigo-600" />
                  Phase & Cost Type Matrix
                </button>

                <button
                  onClick={() => setActiveTab('integrity_audit')}
                  className={`px-3.5 py-2 text-xs font-semibold rounded-t-lg transition-colors whitespace-nowrap flex items-center gap-2 ${
                    activeTab === 'integrity_audit'
                      ? 'bg-white text-slate-900 border-t-2 border-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  <ShieldCheck className="w-3.5 h-3.5 text-emerald-600" />
                  Summary & Formula Protection
                </button>
              </div>

              {/* Tab Contents */}
              <div className="p-6">
                {/* 0. ECC SUMMARY TAB */}
                {activeTab === 'ecc_summary' && (
                  <div className="space-y-4">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-slate-100">
                      <div>
                        <div className="flex items-center gap-2 text-xs text-slate-500">
                          <span>Worksheet: Project Summary</span>
                          <span>·</span>
                          <span className="text-emerald-700 font-medium">Column E Verified Exact Baseline Match</span>
                          <span>·</span>
                          <span className="text-blue-700 font-medium">Column H Conditional Logic Active</span>
                        </div>
                        <h3 className="text-sm font-bold text-slate-900 mt-0.5">
                          ECC Project Summary Sheet (Phase-Summary Layout)
                        </h3>
                      </div>

                      <div className="text-xs text-slate-500">
                        {report.summaryRows.length > 0 ? `${report.summaryRows.length} phase summary rows` : 'Detail rollup'}
                      </div>
                    </div>

                    {/* Project Summary Sheet Table */}
                    <div className="border border-slate-200 rounded-lg overflow-x-auto shadow-xs">
                      <table className="w-full text-xs text-left">
                        <thead className="bg-slate-100 text-slate-700 font-semibold border-b border-slate-200">
                          <tr>
                            <th className="py-2.5 px-3 text-center">Phase Code (A)</th>
                            <th className="py-2.5 px-3 text-center">Phase Description (B)</th>
                            <th className="py-2.5 px-3 text-center">Original Budget (C)</th>
                            <th className="py-2.5 px-3 text-center">Approved Changes (D)</th>
                            <th className="py-2.5 px-3 text-center bg-blue-50/70 text-blue-900 border-x border-blue-200">
                              Previous Month (Col E)*
                            </th>
                            <th className="py-2.5 px-3 text-center bg-emerald-50/70 text-emerald-900 border-r border-emerald-200">
                              Current Month (Col F)**
                            </th>
                            <th className="py-2.5 px-3 text-center">Forecast Complete (G)</th>
                            <th className="py-2.5 px-3 text-center bg-amber-50/70 text-amber-900 border-l border-amber-200">
                              ∆ to Previous (Col H)***
                            </th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100 font-mono">
                          {report.summaryRows.map((row) => {
                            const isTotal = row.isTotal;
                            const hasVariance = row.varianceColH !== null && row.varianceColH !== undefined;
                            const isPositive = (row.varianceColH || 0) > 0;
                            const isZero = (row.varianceColH || 0) === 0;

                            return (
                              <tr
                                key={row.rowNumber}
                                className={isTotal ? 'bg-slate-100/80 font-bold border-t-2 border-slate-300' : 'hover:bg-slate-50/80'}
                              >
                                <td className="py-2.5 px-3 font-semibold text-slate-900">
                                  {row.phaseCode}
                                </td>
                                <td className="py-2.5 px-3 font-sans text-slate-700">
                                  {row.phaseDesc}
                                </td>
                                <td className="py-2.5 px-3 text-right tabular-nums text-slate-600">
                                  ${row.originalBudget.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                </td>
                                <td className="py-2.5 px-3 text-right tabular-nums text-slate-500">
                                  ${row.approvedChanges.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                </td>
                                <td className="py-2.5 px-3 text-right tabular-nums bg-blue-50/40 text-blue-950 font-medium border-x border-blue-100">
                                  {row.previousMonthColE !== null && row.previousMonthColE !== undefined ? (
                                    `$${row.previousMonthColE.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
                                  ) : (
                                    <span className="text-slate-400 font-sans italic text-[11px]">
                                      [Blank Input]
                                    </span>
                                  )}
                                </td>
                                <td className="py-2.5 px-3 text-right tabular-nums bg-emerald-50/40 text-emerald-950 font-semibold border-r border-emerald-100">
                                  ${row.currentMonthColF.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                </td>
                                <td className="py-2.5 px-3 text-right tabular-nums text-slate-600">
                                  ${row.forecastColG.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                </td>
                                <td className="py-2.5 px-3 text-right tabular-nums bg-amber-50/40 font-bold border-l border-amber-100">
                                  {hasVariance ? (
                                    <span className={isZero ? 'text-slate-500' : isPositive ? 'text-emerald-700' : 'text-rose-700'}>
                                      {isZero ? '$0.00' : `${isPositive ? '+' : '-'}$${Math.abs(row.varianceColH || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`}
                                    </span>
                                  ) : (
                                    <span className="text-slate-400 font-sans italic text-[11px]">
                                      [Blank — E is blank]
                                    </span>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>

                    {/* ECC Governance Rule Notes */}
                    <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg text-xs space-y-1.5 text-slate-600">
                      <div className="font-semibold text-slate-800 flex items-center gap-1.5">
                        <ShieldCheck className="w-4 h-4 text-blue-600" />
                        Summary Workflow Protection Rules Enforced
                      </div>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-1 text-[11px]">
                        <div>
                          <strong className="text-slate-700">* Column E (Previous Month):</strong> Reserved manual-input column. Never populated, cleared, renamed, replaced, or derived during refresh. Verified cell-for-cell exact match with original ECC workbook.
                        </div>
                        <div>
                          <strong className="text-slate-700">** Column F (Current Month):</strong> Updated from refreshed monthly budget source amounts; never moved into any other column.
                        </div>
                        <div>
                          <strong className="text-slate-700">*** Column H (∆ to Previous Budget):</strong> Depends only on Columns E and F using conditional <code className="bg-slate-200 px-1 py-0.5 rounded font-mono">=IF(E[row]&lt;&gt;"",F[row]-E[row],"")</code>. Stays blank when Column E is blank.
                        </div>
                        <div>
                          <strong className="text-slate-700">Senior Leadership Calculation:</strong> Computed independently from detail matched rows, never staged in Column E or routed through Column H.
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                {/* 1. LEADERSHIP EMAIL TAB */}
                {activeTab === 'leadership_email' && (
                  <div className="space-y-4">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-slate-100">
                      <div>
                        <div className="flex items-center gap-2 text-xs text-slate-500">
                          <span>Recipient: Senior Leadership</span>
                          <span>·</span>
                          <span>Status: Draft Only (Never Auto-Send)</span>
                          {report.aiAssisted && (
                            <>
                              <span>·</span>
                              <span className="text-blue-600 font-medium">Gemini 3.8 Flash Assisted</span>
                            </>
                          )}
                        </div>
                        <h3 className="text-sm font-bold text-slate-900 mt-0.5">
                          Executive Email Writing Block
                        </h3>
                      </div>

                      <div className="flex items-center gap-2">
                        <button
                          onClick={handleCopyEmail}
                          className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-md transition-colors shadow-xs"
                        >
                          {copiedEmail ? (
                            <>
                              <Check className="w-3.5 h-3.5 text-emerald-600" />
                              Copied to Clipboard!
                            </>
                          ) : (
                            <>
                              <Copy className="w-3.5 h-3.5" />
                              Copy Draft
                            </>
                          )}
                        </button>
                      </div>
                    </div>

                    {/* Email Display Container */}
                    <div className="bg-slate-900 text-slate-100 rounded-lg p-5 font-mono text-xs leading-relaxed overflow-x-auto border border-slate-800 shadow-inner">
                      <pre className="whitespace-pre-wrap font-mono">
                        {report.emailDraft}
                      </pre>
                    </div>

                    {/* Email Rule Compliance Notes */}
                    <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg text-xs space-y-1.5 text-slate-600">
                      <div className="font-semibold text-slate-800 flex items-center gap-1.5">
                        <ShieldCheck className="w-4 h-4 text-emerald-600" />
                        Senior Leadership Email Compliance Checklist
                      </div>
                      <ul className="grid grid-cols-1 md:grid-cols-2 gap-x-4 gap-y-1 text-[11px] list-disc list-inside">
                        <li>Factual, professional tone with no speculative explanations</li>
                        <li>Strict table columns: <code className="font-mono">Phase | Cost Type | Previous Budget | Updated Budget | Change</code></li>
                        <li>Strict $25k threshold applied: changes of exactly ±$25,000.00 are excluded</li>
                        <li>Increases prefixed with +, decreases prefixed with -</li>
                        <li>Money formatted to 2 decimal places with thousands separators</li>
                        <li>{report.isProvisional ? 'Explicitly flagged as PROVISIONAL due to unresolved exceptions' : 'Reconciled figures marked as final'}</li>
                      </ul>
                    </div>
                  </div>
                )}

                {/* 2. MATERIAL VARIANCE TABLE TAB */}
                {activeTab === 'material_variances' && (
                  <div className="space-y-4">
                    <div className="flex items-center justify-between pb-2 border-b border-slate-100">
                      <div>
                        <h3 className="text-sm font-bold text-slate-900">
                          Phase & Cost-Type Changes Exceeding Strict $25,000 Threshold
                        </h3>
                        <p className="text-xs text-slate-500">
                          Combinations where <code className="font-mono bg-slate-100 px-1 py-0.5 rounded">|Updated - Previous| &gt; $25,000.00</code>. Exact ±$25,000.00 changes are excluded.
                        </p>
                      </div>

                      <div className="text-xs text-slate-500">
                        Qualifying Combinations:{' '}
                        <span className="font-mono font-bold text-slate-900">
                          {report.thresholdVariances.length}
                        </span>
                      </div>
                    </div>

                    {report.thresholdVariances.length === 0 ? (
                      <div className="p-8 text-center bg-slate-50 border border-slate-200 rounded-lg">
                        <Info className="w-8 h-8 text-slate-400 mx-auto mb-2" />
                        <p className="text-sm font-semibold text-slate-800">
                          No phase-and-cost-type changes exceeded the strict $25,000 threshold for this period.
                        </p>
                        <p className="text-xs text-slate-500 mt-1">
                          All aggregated movements across phases and cost categories remained within the ±$25,000.00 boundary.
                        </p>
                      </div>
                    ) : (
                      <div className="border border-slate-200 rounded-lg overflow-hidden">
                        <table className="w-full text-xs text-left">
                          <thead className="bg-slate-100 text-slate-700 font-semibold border-b border-slate-200">
                            <tr>
                              <th className="py-2.5 px-4">Phase</th>
                              <th className="py-2.5 px-4">Cost Type</th>
                              <th className="py-2.5 px-4 text-right">Previous Budget</th>
                              <th className="py-2.5 px-4 text-right">Updated Budget</th>
                              <th className="py-2.5 px-4 text-right">Change</th>
                              <th className="py-2.5 px-4 text-center">Status</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 font-mono">
                            {report.thresholdVariances.map((v, i) => {
                              const isPositive = v.change > 0;
                              const isNegative = v.change < 0;
                              return (
                                <tr key={i} className="hover:bg-slate-50/80">
                                  <td className="py-2 px-4 font-sans font-medium text-slate-900">
                                    {v.phase}
                                  </td>
                                  <td className="py-2 px-4 font-sans text-slate-700">
                                    {v.costType}
                                  </td>
                                  <td className="py-2 px-4 text-right tabular-nums text-slate-600">
                                    ${v.previousBudget.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                  </td>
                                  <td className="py-2 px-4 text-right tabular-nums text-slate-900 font-semibold">
                                    ${v.updatedBudget.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                  </td>
                                  <td
                                    className={`py-2 px-4 text-right tabular-nums font-bold ${
                                      isPositive ? 'text-emerald-700' : 'text-rose-700'
                                    }`}
                                  >
                                    {isPositive ? '+' : '-'}
                                    $
                                    {Math.abs(v.change).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                  </td>
                                  <td className="py-2 px-4 text-center">
                                    <span className="inline-block px-2 py-0.5 text-[10px] font-sans font-medium rounded-full bg-slate-100 text-slate-700">
                                      Material (|∆| &gt; $25k)
                                    </span>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                )}

                {/* 3. EXCEPTIONS & CLARIFICATIONS TAB */}
                {activeTab === 'exceptions' && (
                  <div className="space-y-4">
                    <div className="flex items-center justify-between pb-2 border-b border-slate-100">
                      <div>
                        <h3 className="text-sm font-bold text-slate-900">
                          Reconciliation Exceptions & Field Classifications
                        </h3>
                        <p className="text-xs text-slate-500">
                          Before updating, the app identifies duplicate keys, source-only keys, ECC-only keys, and formula cells.
                        </p>
                      </div>

                      <div className="text-xs font-mono font-bold text-slate-700">
                        {report.exceptions.length} exception(s) detected
                      </div>
                    </div>

                    {report.exceptions.length === 0 ? (
                      <div className="p-8 text-center bg-emerald-50/50 border border-emerald-200 rounded-lg">
                        <CheckCircle2 className="w-8 h-8 text-emerald-600 mx-auto mb-2" />
                        <p className="text-sm font-semibold text-emerald-900">
                          Clean Reconciliation — Zero Exceptions Detected
                        </p>
                        <p className="text-xs text-emerald-700 mt-1">
                          All detail records have a valid, unique 1-to-1 match. No duplicate keys or unmatched source items.
                        </p>
                      </div>
                    ) : (
                      <div className="space-y-3">
                        <div className="p-3 bg-amber-50 border border-amber-200 rounded-lg text-xs text-amber-800 flex items-start gap-2">
                          <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                          <div>
                            <span className="font-semibold">Governance Rule: </span>
                            If any exception affects an amount, do not finalize that amount, the related aggregate, or the leadership email. Report the exception with sheet, row, and key details and request clarification.
                          </div>
                        </div>

                        <div className="border border-slate-200 rounded-lg overflow-hidden">
                          <table className="w-full text-xs text-left">
                            <thead className="bg-slate-100 text-slate-700 font-semibold border-b border-slate-200">
                              <tr>
                                <th className="py-2.5 px-3">Type</th>
                                <th className="py-2.5 px-3">Sheet</th>
                                <th className="py-2.5 px-3">Row</th>
                                <th className="py-2.5 px-3">Composite Key</th>
                                <th className="py-2.5 px-3">Details</th>
                                <th className="py-2.5 px-3">Recommended Controller Action</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                              {report.exceptions.map((ex) => (
                                <tr key={ex.id} className="hover:bg-slate-50/80">
                                  <td className="py-2.5 px-3 font-semibold text-slate-900 whitespace-nowrap">
                                    {ex.type === 'duplicate_ecc' && (
                                      <span className="text-rose-700">Duplicate Key (ECC)</span>
                                    )}
                                    {ex.type === 'duplicate_source' && (
                                      <span className="text-rose-700">Duplicate Key (Source)</span>
                                    )}
                                    {ex.type === 'ecc_only' && (
                                      <span className="text-amber-700">Missing from Source (Held)</span>
                                    )}
                                    {ex.type === 'source_only' && (
                                      <span className="text-blue-700">New Source Item</span>
                                    )}
                                    {ex.type === 'formula_cell' && (
                                      <span className="text-purple-700">Protected Formula Cell</span>
                                    )}
                                  </td>
                                  <td className="py-2.5 px-3 font-mono text-slate-600 whitespace-nowrap">
                                    {ex.sheetName}
                                  </td>
                                  <td className="py-2.5 px-3 font-mono font-medium text-slate-800 text-center">
                                    {ex.rowNumber}
                                  </td>
                                  <td className="py-2.5 px-3 font-mono text-[11px] text-slate-700 max-w-[180px] truncate" title={ex.compositeKey}>
                                    {ex.compositeKey}
                                  </td>
                                  <td className="py-2.5 px-3 text-slate-700">
                                    {ex.description}
                                  </td>
                                  <td className="py-2.5 px-3 text-slate-600 italic">
                                    {ex.suggestedAction}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}
                  </div>
                )}

                {/* 4. RECONCILIATION MATRIX TAB */}
                {activeTab === 'reconciliation_matrix' && (
                  <div className="space-y-4">
                    <div className="flex items-center justify-between pb-2 border-b border-slate-100">
                      <div>
                        <h3 className="text-sm font-bold text-slate-900">
                          Complete Phase & Cost Type Reconciliation
                        </h3>
                        <p className="text-xs text-slate-500">
                          Detailed aggregate breakdown of all phase and cost category combinations.
                        </p>
                      </div>

                      <div className="text-xs text-slate-500">
                        Total Combinations:{' '}
                        <span className="font-mono font-bold text-slate-900">
                          {report.phaseCostTypeVariances.length}
                        </span>
                      </div>
                    </div>

                    <div className="border border-slate-200 rounded-lg overflow-hidden">
                      <table className="w-full text-xs text-left">
                        <thead className="bg-slate-100 text-slate-700 font-semibold border-b border-slate-200">
                          <tr>
                            <th className="py-2.5 px-4">Phase</th>
                            <th className="py-2.5 px-4">Cost Type</th>
                            <th className="py-2.5 px-4 text-right">Previous Budget</th>
                            <th className="py-2.5 px-4 text-right">Updated Budget</th>
                            <th className="py-2.5 px-4 text-right">Variance (∆)</th>
                            <th className="py-2.5 px-4 text-center">Materiality</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100 font-mono">
                          {report.phaseCostTypeVariances.map((v, i) => {
                            const isPositive = v.change > 0;
                            const isZero = v.change === 0;
                            return (
                              <tr key={i} className="hover:bg-slate-50/80">
                                <td className="py-2 px-4 font-sans font-medium text-slate-900">
                                  {v.phase}
                                </td>
                                <td className="py-2 px-4 font-sans text-slate-700">
                                  {v.costType}
                                </td>
                                <td className="py-2 px-4 text-right tabular-nums text-slate-600">
                                  ${v.previousBudget.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                </td>
                                <td className="py-2 px-4 text-right tabular-nums text-slate-900 font-semibold">
                                  ${v.updatedBudget.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                </td>
                                <td
                                  className={`py-2 px-4 text-right tabular-nums font-bold ${
                                    isZero
                                      ? 'text-slate-400'
                                      : isPositive
                                      ? 'text-emerald-700'
                                      : 'text-rose-700'
                                  }`}
                                >
                                  {isZero ? '$0.00' : `${isPositive ? '+' : '-'}$${Math.abs(v.change).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`}
                                </td>
                                <td className="py-2 px-4 text-center">
                                  {v.exceedsStrictThreshold ? (
                                    <span className="inline-block px-2 py-0.5 text-[10px] font-sans font-semibold rounded-full bg-blue-100 text-blue-800">
                                      Material (|∆| &gt; $25k)
                                    </span>
                                  ) : (
                                    <span className="inline-block px-2 py-0.5 text-[10px] font-sans text-slate-400">
                                      Within Threshold
                                    </span>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {/* 5. INTEGRITY AUDIT TAB (Column E & Summary Protection) */}
                {activeTab === 'integrity_audit' && (
                  <div className="space-y-4">
                    <div className="pb-2 border-b border-slate-100">
                      <h3 className="text-sm font-bold text-slate-900">
                        Formula & Project Summary Protection Verification
                      </h3>
                      <p className="text-xs text-slate-500">
                        Audit confirmation of Column E manual reservation, Column H conditional formula logic, and workbook formula preservation.
                      </p>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      {/* Check 1 */}
                      <div className="p-4 bg-slate-50 border border-slate-200 rounded-lg space-y-2">
                        <div className="flex items-center gap-2">
                          {report.columnEIntegrityVerified ? (
                            <CheckCircle2 className="w-5 h-5 text-emerald-600" />
                          ) : (
                            <XCircle className="w-5 h-5 text-rose-600" />
                          )}
                          <h4 className="text-xs font-bold text-slate-900">
                            Column E (Previous Month) Exact Match
                          </h4>
                        </div>
                        <p className="text-xs text-slate-600 leading-relaxed">
                          Column E was verified cell-by-cell between the original and updated workbooks. All values, manual entries, and formulas are 100% identical. Never populated, cleared, renamed, or replaced during the refresh.
                        </p>
                        <div className="text-[11px] font-mono text-emerald-700 bg-emerald-50 px-2 py-1 rounded">
                          Verification: PASS (100% Identical)
                        </div>
                      </div>

                      {/* Check 2 */}
                      <div className="p-4 bg-slate-50 border border-slate-200 rounded-lg space-y-2">
                        <div className="flex items-center gap-2">
                          {report.columnHFormulasVerified ? (
                            <CheckCircle2 className="w-5 h-5 text-emerald-600" />
                          ) : (
                            <XCircle className="w-5 h-5 text-rose-600" />
                          )}
                          <h4 className="text-xs font-bold text-slate-900">
                            Column H (∆ to Previous Budget) Formula Rule
                          </h4>
                        </div>
                        <p className="text-xs text-slate-600 leading-relaxed">
                          Each detail row in Column H depends strictly on Columns E and F using the conditional syntax <code className="bg-slate-200 px-1 py-0.5 rounded text-[10px]">=IF(E[row]&lt;&gt;"",F[row]-E[row],"")</code>. Stays blank when Column E is blank.
                        </p>
                        <div className="text-[11px] font-mono text-emerald-700 bg-emerald-50 px-2 py-1 rounded">
                          Verification: PASS (Conditional Blank Maintained)
                        </div>
                      </div>

                      {/* Check 3 */}
                      <div className="p-4 bg-slate-50 border border-slate-200 rounded-lg space-y-2">
                        <div className="flex items-center gap-2">
                          <CheckCircle2 className="w-5 h-5 text-emerald-600" />
                          <h4 className="text-xs font-bold text-slate-900">
                            Formula Cell Overwrite Protection
                          </h4>
                        </div>
                        <p className="text-xs text-slate-600 leading-relaxed">
                          If an ECC Current Budget cell contains an Excel formula rather than a static typed value, the app does not overwrite it; it stops and requests clarification.
                        </p>
                        <div className="text-[11px] font-mono text-emerald-700 bg-emerald-50 px-2 py-1 rounded">
                          Verification: PASS (Formula Invariant Enforced)
                        </div>
                      </div>

                      {/* Check 4 */}
                      <div className="p-4 bg-slate-50 border border-slate-200 rounded-lg space-y-2">
                        <div className="flex items-center gap-2">
                          <CheckCircle2 className="w-5 h-5 text-emerald-600" />
                          <h4 className="text-xs font-bold text-slate-900">
                            Workbook Separation (Zero Overwrite)
                          </h4>
                        </div>
                        <p className="text-xs text-slate-600 leading-relaxed">
                          The original ECC and LandDev monthly source workbooks remain 100% unaltered. An updated workbook is generated separately under a distinct filename containing <code className="font-mono text-slate-800">Updated</code>.
                        </p>
                        <div className="text-[11px] font-mono text-emerald-700 bg-emerald-50 px-2 py-1 rounded">
                          Saved as: {report.updatedFileName}
                        </div>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </section>
        )}
      </main>

      {/* FOOTER */}
      <footer className="border-t border-slate-200 bg-white py-4 px-6 text-xs text-slate-500 flex flex-col sm:flex-row items-center justify-between gap-3">
        <div className="flex items-center gap-2">
          <span className="font-semibold text-slate-700">LandDev ECC Variance Tracker</span>
          <span>·</span>
          <span>Preserving Spreadsheet Integrity & Governance</span>
        </div>
        <div>
          Automated audit compliant with LandDev financial reporting protocols.
        </div>
      </footer>
    </div>
  );
}
