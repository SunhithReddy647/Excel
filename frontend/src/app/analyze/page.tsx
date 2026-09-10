"use client";

import { useState, useEffect, Suspense } from "react";
import { useRouter, useSearchParams } from "next/navigation";
import Link from "next/link";
import { analyze, generateWorkbook, type AnalyzeResponse, type ExcelPlan } from "@/lib/api";

function AnalyzeContent() {
  const router = useRouter();
  const searchParams = useSearchParams();
  const datasetId = searchParams.get("dataset_id") || "";
  const assignmentId = searchParams.get("assignment_id") || "";

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [analyzeRes, setAnalyzeRes] = useState<AnalyzeResponse | null>(null);
  const [generating, setGenerating] = useState(false);

  useEffect(() => {
    if (!datasetId || !assignmentId) {
      setError("Missing dataset or assignment ID");
      setLoading(false);
      return;
    }

    analyze(datasetId, assignmentId)
      .then((res) => {
        setAnalyzeRes(res);
        setLoading(false);
      })
      .catch((err) => {
        setError(err.message);
        setLoading(false);
      });
  }, [datasetId, assignmentId]);

  const handleGenerate = async () => {
    if (!analyzeRes) return;
    setGenerating(true);

    try {
      await generateWorkbook(analyzeRes.generation_id, analyzeRes.plan);
      router.push(`/generate?generation_id=${analyzeRes.generation_id}`);
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : "Generation failed");
      setGenerating(false);
    }
  };

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="text-center space-y-6">
          <div className="w-16 h-16 border-4 border-[var(--color-primary)] border-t-transparent rounded-full animate-spin mx-auto" />
          <div>
            <h2 className="text-xl font-semibold mb-2">Analyzing your assignment</h2>
            <p className="text-[var(--color-text-muted)]">
              The AI is interpreting your question and creating a generation plan...
            </p>
          </div>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="glass-card-static p-8 max-w-lg text-center">
          <span className="text-4xl mb-4 block">⚠️</span>
          <h2 className="text-xl font-semibold mb-2">Analysis Failed</h2>
          <p className="text-[var(--color-error)] mb-4">{error}</p>
          <Link href="/upload" className="btn-primary no-underline">
            Try Again
          </Link>
        </div>
      </div>
    );
  }

  const plan = analyzeRes!.plan;
  const profile = analyzeRes!.dataset_profile;

  return (
    <div className="min-h-screen">
      <nav className="sticky top-0 z-50 backdrop-blur-xl bg-[var(--color-bg)]/80 border-b border-[var(--color-border-subtle)]">
        <div className="max-w-7xl mx-auto px-6 h-16 flex items-center justify-between">
          <Link href="/" className="flex items-center gap-2 no-underline">
            <span className="text-2xl">📊</span>
            <span className="text-xl font-bold gradient-text">ExcelFlow</span>
          </Link>
          <span className="text-sm text-[var(--color-text-muted)]">
            Step 2 of 4 — Review Plan
          </span>
        </div>
      </nav>

      <div className="max-w-6xl mx-auto px-6 py-10">
        <h1 className="text-3xl font-bold mb-2">Generation Plan</h1>
        <p className="text-[var(--color-text-muted)] mb-8">
          Review the AI-generated plan before creating your workbook.
        </p>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* ─── Left: Plan Summary ──────────────────────── */}
          <div className="lg:col-span-2 space-y-6">
            {/* Dataset summary */}
            <div className="glass-card-static p-5">
              <h3 className="text-sm font-semibold uppercase tracking-wider text-[var(--color-text-secondary)] mb-3">
                Dataset
              </h3>
              <div className="grid grid-cols-3 gap-3">
                <div className="stat-card">
                  <div className="stat-value">{profile.row_count}</div>
                  <div className="stat-label">Rows</div>
                </div>
                <div className="stat-card">
                  <div className="stat-value">{profile.column_count}</div>
                  <div className="stat-label">Columns</div>
                </div>
                <div className="stat-card">
                  <div className="stat-value">{profile.dimensions.length}</div>
                  <div className="stat-label">Dimensions</div>
                </div>
              </div>
            </div>

            {/* Calculations */}
            {plan.calculations.length > 0 && (
              <div className="glass-card-static p-5">
                <h3 className="text-sm font-semibold uppercase tracking-wider text-[var(--color-text-secondary)] mb-3">
                  Calculations ({plan.calculations.length})
                </h3>
                <div className="space-y-2">
                  {plan.calculations.map((calc) => (
                    <div key={calc.name} className="plan-item">
                      <span className="text-[var(--color-success)]">✓</span>
                      <div className="flex-1">
                        <p className="font-medium text-sm">{calc.name}</p>
                        <p className="text-xs text-[var(--color-text-muted)] font-mono">
                          {calc.formula_template}
                        </p>
                      </div>
                      <span className="text-xs px-2 py-0.5 rounded bg-blue-500/10 text-blue-400 border border-blue-500/20">
                        {calc.format}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Analyses */}
            {plan.analyses.length > 0 && (
              <div className="glass-card-static p-5">
                <h3 className="text-sm font-semibold uppercase tracking-wider text-[var(--color-text-secondary)] mb-3">
                  Analysis ({plan.analyses.length})
                </h3>
                <div className="space-y-2">
                  {plan.analyses.map((a) => (
                    <div key={a.name} className="plan-item">
                      <span className="text-[var(--color-success)]">✓</span>
                      <div className="flex-1">
                        <p className="font-medium text-sm">{a.name}</p>
                        <p className="text-xs text-[var(--color-text-muted)]">
                          {a.aggregation}({a.metric}) by{" "}
                          {a.group_by.join(", ")}
                        </p>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Charts */}
            {plan.charts.length > 0 && (
              <div className="glass-card-static p-5">
                <h3 className="text-sm font-semibold uppercase tracking-wider text-[var(--color-text-secondary)] mb-3">
                  Visualizations ({plan.charts.length})
                </h3>
                <div className="space-y-2">
                  {plan.charts.map((c) => (
                    <div key={c.name} className="plan-item">
                      <span className="text-[var(--color-success)]">✓</span>
                      <div className="flex-1">
                        <p className="font-medium text-sm">{c.title}</p>
                        <p className="text-xs text-[var(--color-text-muted)]">
                          {c.chart_type} chart
                        </p>
                      </div>
                      <span className="text-xs px-2 py-0.5 rounded bg-purple-500/10 text-purple-400 border border-purple-500/20">
                        {c.chart_type}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Filters */}
            {plan.filters.length > 0 && (
              <div className="glass-card-static p-5">
                <h3 className="text-sm font-semibold uppercase tracking-wider text-[var(--color-text-secondary)] mb-3">
                  Filters ({plan.filters.length})
                </h3>
                <div className="space-y-2">
                  {plan.filters.map((f) => (
                    <div key={f.field} className="plan-item">
                      <span className="text-[var(--color-success)]">✓</span>
                      <div className="flex-1">
                        <p className="font-medium text-sm">{f.field}</p>
                        <p className="text-xs text-[var(--color-text-muted)]">
                          {f.filter_type}
                        </p>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* ─── Right: Sheets & Generate ──────────────────── */}
          <div className="space-y-6">
            {/* Workbook structure */}
            <div className="glass-card-static p-5">
              <h3 className="text-sm font-semibold uppercase tracking-wider text-[var(--color-text-secondary)] mb-3">
                Workbook Structure
              </h3>
              <div className="space-y-1.5">
                {plan.sheets.map((s) => (
                  <div
                    key={s.name}
                    className="flex items-center gap-2 text-sm py-1.5 px-3 rounded bg-[var(--color-bg-input)] border border-[var(--color-border-subtle)]"
                  >
                    <span className="text-[var(--color-secondary)]">📋</span>
                    <span>{s.name}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Dashboard KPIs */}
            {plan.dashboard && plan.dashboard.kpis.length > 0 && (
              <div className="glass-card-static p-5">
                <h3 className="text-sm font-semibold uppercase tracking-wider text-[var(--color-text-secondary)] mb-3">
                  Dashboard KPIs
                </h3>
                <div className="space-y-1.5">
                  {plan.dashboard.kpis.map((k) => (
                    <div
                      key={k.name}
                      className="flex items-center gap-2 text-sm py-1.5 px-3 rounded bg-[var(--color-bg-input)] border border-[var(--color-border-subtle)]"
                    >
                      <span className="text-[var(--color-warning)]">🎯</span>
                      <span>{k.label || k.name}</span>
                      <span className="text-xs text-[var(--color-text-muted)] ml-auto">
                        {k.aggregation}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Generate Button */}
            <div className="glass-card-static p-5 text-center">
              <p className="text-sm text-[var(--color-text-secondary)] mb-4">
                {plan.sheets.length} sheets •{" "}
                {plan.calculations.length} calculations •{" "}
                {plan.charts.length} charts
              </p>
              <button
                className="btn-primary w-full"
                disabled={generating}
                onClick={handleGenerate}
              >
                {generating ? (
                  <span className="flex items-center justify-center gap-2">
                    <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                    Starting...
                  </span>
                ) : (
                  "Generate Workbook"
                )}
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default function AnalyzePage() {
  return (
    <Suspense
      fallback={
        <div className="min-h-screen flex items-center justify-center">
          <div className="w-12 h-12 border-4 border-[var(--color-primary)] border-t-transparent rounded-full animate-spin" />
        </div>
      }
    >
      <AnalyzeContent />
    </Suspense>
  );
}
