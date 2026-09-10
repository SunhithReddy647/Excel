"use client";

import { useState, useEffect, Suspense } from "react";
import { useRouter, useSearchParams } from "next/navigation";
import Link from "next/link";
import { getGenerationStatus, type GenerationStatus } from "@/lib/api";

const GENERATION_STEPS = [
  "Analyzing dataset",
  "Loading data",
  "Building calculation plan",
  "Generating workbook",
  "Validating workbook",
];

function GenerateContent() {
  const router = useRouter();
  const searchParams = useSearchParams();
  const generationId = searchParams.get("generation_id") || "";

  const [status, setStatus] = useState<GenerationStatus | null>(null);
  const [error, setError] = useState("");

  useEffect(() => {
    if (!generationId) return;

    const poll = setInterval(async () => {
      try {
        const res = await getGenerationStatus(generationId);
        setStatus(res);

        if (res.status === "completed" || res.status === "failed") {
          clearInterval(poll);
          if (res.status === "completed" && res.workbook_id) {
            setTimeout(() => {
              router.push(`/result?workbook_id=${res.workbook_id}`);
            }, 1500);
          }
          if (res.status === "failed") {
            setError(res.error || "Generation failed");
          }
        }
      } catch {
        // Retry silently
      }
    }, 1000);

    return () => clearInterval(poll);
  }, [generationId, router]);

  const completedSteps = new Set(
    status?.progress
      ?.filter((p) => p.status === "completed")
      .map((p) => p.step) || []
  );
  const activeSteps = new Set(
    status?.progress
      ?.filter((p) => p.status === "in_progress")
      .map((p) => p.step) || []
  );

  return (
    <div className="min-h-screen">
      <nav className="sticky top-0 z-50 backdrop-blur-xl bg-[var(--color-bg)]/80 border-b border-[var(--color-border-subtle)]">
        <div className="max-w-7xl mx-auto px-6 h-16 flex items-center justify-between">
          <Link href="/" className="flex items-center gap-2 no-underline">
            <span className="text-2xl">📊</span>
            <span className="text-xl font-bold gradient-text">ExcelFlow</span>
          </Link>
          <span className="text-sm text-[var(--color-text-muted)]">
            Step 3 of 4 — Generating
          </span>
        </div>
      </nav>

      <div className="max-w-xl mx-auto px-6 py-20">
        <div className="text-center mb-12">
          <h1 className="text-3xl font-bold mb-2">Generating Your Workbook</h1>
          <p className="text-[var(--color-text-muted)]">
            Building formulas, analysis, charts, and dashboard...
          </p>
        </div>

        <div className="glass-card-static p-6 space-y-1">
          {GENERATION_STEPS.map((step) => {
            const isCompleted = completedSteps.has(step);
            const isActive = activeSteps.has(step);
            const isPending = !isCompleted && !isActive;

            return (
              <div
                key={step}
                className={`progress-step ${isCompleted ? "completed" : ""} ${isActive ? "active" : ""}`}
              >
                <div
                  className={`step-icon ${isCompleted ? "completed" : isActive ? "active" : "pending"}`}
                >
                  {isCompleted ? "✓" : isActive ? "●" : "○"}
                </div>
                <span
                  className={`text-sm ${
                    isPending
                      ? "text-[var(--color-text-muted)]"
                      : "text-[var(--color-text)]"
                  }`}
                >
                  {step}
                </span>
                {isActive && (
                  <span className="ml-auto text-xs text-[var(--color-primary)]">
                    In progress...
                  </span>
                )}
              </div>
            );
          })}
        </div>

        {error && (
          <div className="mt-6 p-4 rounded-lg bg-[var(--color-error)]/10 border border-[var(--color-error)]/20 text-center">
            <p className="text-[var(--color-error)] font-medium">{error}</p>
            <Link
              href="/upload"
              className="btn-secondary text-sm mt-3 inline-block no-underline"
            >
              Try Again
            </Link>
          </div>
        )}

        {status?.status === "completed" && (
          <div className="mt-6 text-center animate-fade-in-up">
            <div className="text-4xl mb-4">🎉</div>
            <p className="text-[var(--color-success)] font-semibold">
              Workbook generated successfully!
            </p>
            <p className="text-sm text-[var(--color-text-muted)] mt-1">
              Redirecting to preview...
            </p>
          </div>
        )}
      </div>
    </div>
  );
}

export default function GeneratePage() {
  return (
    <Suspense
      fallback={
        <div className="min-h-screen flex items-center justify-center">
          <div className="w-12 h-12 border-4 border-[var(--color-primary)] border-t-transparent rounded-full animate-spin" />
        </div>
      }
    >
      <GenerateContent />
    </Suspense>
  );
}
