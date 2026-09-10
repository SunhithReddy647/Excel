"use client";

import Link from "next/link";

const FEATURES = [
  {
    icon: "⚡",
    title: "Real Formulas",
    description:
      "Every calculation uses actual Excel formulas — structured references, SUMIFS, IF — not hardcoded values.",
  },
  {
    icon: "📊",
    title: "Dynamic Charts",
    description:
      "Auto-generates real Excel chart objects — bar, line, pie, doughnut — linked directly to your data.",
  },
  {
    icon: "🎯",
    title: "Smart Dashboards",
    description:
      "KPI cards, multiple chart layouts, and filters — all dynamically arranged based on your assignment.",
  },
  {
    icon: "✅",
    title: "Validated Output",
    description:
      "10-point automated validation ensures no broken references, no hardcoded values, and complete coverage.",
  },
  {
    icon: "🧠",
    title: "AI-Powered Analysis",
    description:
      "Natural language understanding of your assignment. The AI plans, deterministic code builds.",
  },
  {
    icon: "📋",
    title: "Any Dataset",
    description:
      "Sales, HR, Finance, Inventory, Marketing — works with any Excel/CSV dataset without code changes.",
  },
];

const STEPS = [
  { num: "01", label: "Upload Dataset", desc: "Excel or CSV file" },
  { num: "02", label: "Add Assignment", desc: "Paste or upload question" },
  { num: "03", label: "Review Plan", desc: "Verify before generating" },
  { num: "04", label: "Download Excel", desc: "Professional workbook" },
];

export default function LandingPage() {
  return (
    <div className="min-h-screen">
      {/* ─── Navbar ─────────────────────────────────────────────── */}
      <nav className="fixed top-0 left-0 right-0 z-50 backdrop-blur-xl bg-[var(--color-bg)]/80 border-b border-[var(--color-border-subtle)]">
        <div className="max-w-7xl mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <span className="text-2xl">📊</span>
            <span className="text-xl font-bold gradient-text">ExcelFlow</span>
          </div>
          <Link href="/upload" className="btn-primary text-sm py-2 px-6 no-underline">
            Get Started
          </Link>
        </div>
      </nav>

      {/* ─── Hero Section ───────────────────────────────────────── */}
      <section className="hero-gradient pt-32 pb-24 px-6 relative overflow-hidden">
        {/* Ambient glow effects */}
        <div className="absolute top-20 left-1/4 w-96 h-96 bg-blue-500/10 rounded-full blur-3xl" />
        <div className="absolute bottom-10 right-1/4 w-80 h-80 bg-purple-500/10 rounded-full blur-3xl" />

        <div className="max-w-4xl mx-auto text-center relative z-10">
          <div className="animate-fade-in-up">
            <span className="inline-block px-4 py-1.5 text-xs font-semibold uppercase tracking-widest text-[var(--color-secondary)] border border-[var(--color-secondary)]/30 rounded-full mb-6 bg-[var(--color-secondary)]/5">
              Intelligent Excel Automation
            </span>
          </div>

          <h1 className="text-5xl md:text-6xl lg:text-7xl font-black leading-tight mb-6 animate-fade-in-up-delay-1">
            Turn Raw Data Into a{" "}
            <span className="gradient-text">Complete Excel Assignment</span>
          </h1>

          <p className="text-lg md:text-xl text-[var(--color-text-secondary)] max-w-2xl mx-auto mb-10 leading-relaxed animate-fade-in-up-delay-2">
            Upload your dataset and assignment. The system creates the formulas,
            tables, charts, analysis and dashboard automatically.
          </p>

          <div className="flex flex-col sm:flex-row gap-4 justify-center animate-fade-in-up-delay-3">
            <Link href="/upload" className="btn-primary text-base px-10 py-3.5 no-underline">
              Create Excel Workbook
            </Link>
            <button
              className="btn-secondary text-base px-10 py-3.5"
              onClick={() => {
                document
                  .getElementById("features")
                  ?.scrollIntoView({ behavior: "smooth" });
              }}
            >
              View Example
            </button>
          </div>
        </div>
      </section>

      {/* ─── How It Works ───────────────────────────────────────── */}
      <section className="py-20 px-6 border-t border-[var(--color-border-subtle)]">
        <div className="max-w-6xl mx-auto">
          <h2 className="text-3xl font-bold text-center mb-14">
            How It Works
          </h2>
          <div className="grid grid-cols-1 md:grid-cols-4 gap-8">
            {STEPS.map((step) => (
              <div key={step.num} className="text-center group">
                <div className="w-16 h-16 rounded-2xl bg-[var(--color-primary-glow)] border border-[var(--color-primary)]/30 flex items-center justify-center mx-auto mb-4 group-hover:scale-110 transition-transform">
                  <span className="text-xl font-bold gradient-text">
                    {step.num}
                  </span>
                </div>
                <h3 className="text-lg font-semibold mb-1">{step.label}</h3>
                <p className="text-sm text-[var(--color-text-muted)]">
                  {step.desc}
                </p>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* ─── Features ───────────────────────────────────────────── */}
      <section
        id="features"
        className="py-20 px-6 border-t border-[var(--color-border-subtle)]"
      >
        <div className="max-w-6xl mx-auto">
          <h2 className="text-3xl font-bold text-center mb-4">
            Everything You Need
          </h2>
          <p className="text-center text-[var(--color-text-muted)] mb-14 max-w-xl mx-auto">
            Real Excel formulas, real charts, real validation — not a
            spreadsheet-shaped JPEG.
          </p>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {FEATURES.map((feature) => (
              <div key={feature.title} className="glass-card p-6">
                <div className="text-3xl mb-4">{feature.icon}</div>
                <h3 className="text-lg font-semibold mb-2">{feature.title}</h3>
                <p className="text-sm text-[var(--color-text-secondary)] leading-relaxed">
                  {feature.description}
                </p>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* ─── CTA ────────────────────────────────────────────────── */}
      <section className="py-20 px-6 border-t border-[var(--color-border-subtle)]">
        <div className="max-w-3xl mx-auto text-center glass-card-static p-12">
          <h2 className="text-3xl font-bold mb-4">Ready to automate?</h2>
          <p className="text-[var(--color-text-secondary)] mb-8">
            Upload your dataset and assignment — get a professional Excel
            workbook in minutes.
          </p>
          <Link href="/upload" className="btn-primary text-base px-10 py-3.5 no-underline">
            Start Now →
          </Link>
        </div>
      </section>

      {/* ─── Footer ─────────────────────────────────────────────── */}
      <footer className="py-8 px-6 border-t border-[var(--color-border-subtle)] text-center text-sm text-[var(--color-text-muted)]">
        ExcelFlow — Intelligent Excel Assignment Automation
      </footer>
    </div>
  );
}
