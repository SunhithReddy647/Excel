"use client";

import { useState, useEffect, Suspense } from "react";
import { useSearchParams } from "next/navigation";
import Link from "next/link";
import { getWorkbookSummary, getDownloadUrl, type WorkbookSummary } from "@/lib/api";

function ResultContent() {
  const searchParams = useSearchParams();
  const workbookId = searchParams.get("workbook_id") || "";

  const [summary, setSummary] = useState<WorkbookSummary | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  useEffect(() => {
    if (!workbookId) {
      setError("No workbook ID provided");
      setLoading(false);
      return;
    }

    getWorkbookSummary(workbookId)
      .then((res) => {
        setSummary(res);
        setLoading(false);
      })
      .catch((err) => {
        setError(err.message);
        setLoading(false);
      });
  }, [workbookId]);

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="w-12 h-12 border-4 border-[var(--color-primary)] border-t-transparent rounded-full animate-spin" />
      </div>
    );
  }

  if (error || !summary) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="glass-card-static p-8 max-w-lg text-center">
          <span className="text-4xl mb-4 block">⚠️</span>
          <h2 className="text-xl font-semibold mb-2">Error</h2>
          <p className="text-[var(--color-error)]">{error || "Not found"}</p>
          <Link href="/upload" className="btn-primary mt-4 inline-block no-underline">
            Start Over
          </Link>
        </div>
      </div>
    );
  }

  const validation = summary.validation;
  const allPassed =
    validation && validation.overall_status === "passed";

  return (
    <div className="min-h-screen">
      <nav className="sticky top-0 z-50 backdrop-blur-xl bg-[var(--color-bg)]/80 border-b border-[var(--color-border-subtle)]">
        <div className="max-w-7xl mx-auto px-6 h-16 flex items-center justify-between">
          <Link href="/" className="flex items-center gap-2 no-underline">
            <span className="text-2xl">📊</span>
            <span className="text-xl font-bold gradient-text">ExcelFlow</span>
          </Link>
          <span className="text-sm text-[var(--color-text-muted)]">
            Step 4 of 4 — Download
          </span>
        </div>
      </nav>

      <div className="max-w-5xl mx-auto px-6 py-10">
        {/* Success Header */}
        <div className="text-center mb-10 animate-fade-in-up">
          <div className="text-5xl mb-4">🎉</div>
          <h1 className="text-3xl font-bold mb-2">Workbook Generated!</h1>
          <p className="text-[var(--color-text-muted)]">
            Your Excel workbook is ready for download.
          </p>
        </div>

        {/* Native Pivot Highlight Badge */}
        {summary.native_pivots && (
          <div className="mb-8 p-4 rounded-xl border border-emerald-500/30 bg-emerald-500/10 backdrop-blur-md flex flex-wrap items-center justify-between gap-4 animate-fade-in-up">
            <div className="flex items-center gap-3">
              <span className="text-2xl">✨</span>
              <div>
                <h3 className="font-semibold text-emerald-400">Authentic Native PivotTables & Slicers Active</h3>
                <p className="text-xs text-[var(--color-text-muted)]">
                  Generated with genuine Excel PivotCache, dynamic field lists, tabular executive styles, and interactive Slicers.
                </p>
              </div>
            </div>
            <div className="flex items-center gap-2">
              <span className="px-3 py-1 rounded-full text-xs font-semibold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                {summary.pivots_created ?? 0} PivotTables
              </span>
              <span className="px-3 py-1 rounded-full text-xs font-semibold bg-cyan-500/20 text-cyan-300 border border-cyan-500/30">
                {summary.slicers_created ?? 0} Slicers
              </span>
            </div>
          </div>
        )}

        {/* Stats Grid */}
        <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-4 mb-8 animate-fade-in-up-delay-1">
          <div className="stat-card">
            <div className="stat-value">{summary.sheets.length}</div>
            <div className="stat-label">Sheets</div>
          </div>
          <div className="stat-card">
            <div className="stat-value">{summary.row_count}</div>
            <div className="stat-label">Rows</div>
          </div>
          <div className="stat-card">
            <div className="stat-value">{summary.formula_count}</div>
            <div className="stat-label">Formulas</div>
          </div>
          <div className="stat-card">
            <div className="stat-value">{summary.chart_count}</div>
            <div className="stat-label">Charts</div>
          </div>
          <div className="stat-card">
            <div className="stat-value">{summary.analysis_count}</div>
            <div className="stat-label">Analyses</div>
          </div>
          <div className="stat-card">
            <div className="stat-value">{summary.filter_count}</div>
            <div className="stat-label">Filters</div>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 animate-fade-in-up-delay-2">
          {/* Validation Results */}
          <div className="glass-card-static p-5">
            <h3 className="text-sm font-semibold uppercase tracking-wider text-[var(--color-text-secondary)] mb-4">
              Validation Results
            </h3>

            {allPassed && (
              <div className="p-3 rounded-lg bg-[var(--color-success)]/10 border border-[var(--color-success)]/20 mb-4">
                <p className="text-[var(--color-success)] font-semibold text-sm">
                  ✓ All validation checks passed
                </p>
              </div>
            )}

            <div className="space-y-2">
              {validation?.checks.map((check, i) => (
                <div key={i} className="flex items-center gap-3 text-sm py-2 px-3 rounded bg-[var(--color-bg-input)] border border-[var(--color-border-subtle)]">
                  <span className={check.passed ? "badge-passed" : "badge-failed"}>
                    {check.passed ? "✓ PASS" : "✗ FAIL"}
                  </span>
                  <div className="flex-1">
                    <p className="font-medium">{check.check_name}</p>
                    <p className="text-xs text-[var(--color-text-muted)]">
                      {check.message}
                    </p>
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Sheets List & Download */}
          <div className="space-y-6">
            <div className="glass-card-static p-5">
              <h3 className="text-sm font-semibold uppercase tracking-wider text-[var(--color-text-secondary)] mb-4">
                Workbook Sheets
              </h3>
              <div className="space-y-1.5">
                {summary.sheets.map((sheet) => (
                  <div
                    key={sheet}
                    className="flex items-center gap-2 text-sm py-2 px-3 rounded bg-[var(--color-bg-input)] border border-[var(--color-border-subtle)]"
                  >
                    <span className="text-[var(--color-secondary)]">📋</span>
                    <span>{sheet}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Download */}
            <div className="glass-card-static p-6 text-center">
              <p className="text-lg font-semibold mb-1">{summary.filename}</p>
              <p className="text-sm text-[var(--color-text-muted)] mb-6">
                Professional Excel workbook with real formulas, charts, and validation
              </p>
              <a
                href={getDownloadUrl(workbookId)}
                className="btn-primary inline-block no-underline text-lg px-10 py-3.5"
                download
              >
                ⬇ Download Excel
              </a>
              <div className="mt-4">
                <Link
                  href="/upload"
                  className="text-sm text-[var(--color-text-muted)] hover:text-[var(--color-primary)] transition-colors no-underline"
                >
                  ← Create another workbook
                </Link>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default function ResultPage() {
  return (
    <Suspense
      fallback={
        <div className="min-h-screen flex items-center justify-center">
          <div className="w-12 h-12 border-4 border-[var(--color-primary)] border-t-transparent rounded-full animate-spin" />
        </div>
      }
    >
      <ResultContent />
    </Suspense>
  );
}
