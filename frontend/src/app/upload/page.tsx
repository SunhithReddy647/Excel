"use client";

import { useState, useCallback, useRef, useEffect } from "react";
import { useRouter } from "next/navigation";
import Link from "next/link";
import {
  uploadDataset,
  pasteQuestion,
  uploadQuestion,
  getExamples,
  type DatasetProfile,
} from "@/lib/api";

type QuestionTab = "paste" | "upload" | "examples";

export default function UploadPage() {
  const router = useRouter();

  // Dataset state
  const [dataFile, setDataFile] = useState<File | null>(null);
  const [dataProfile, setDataProfile] = useState<DatasetProfile | null>(null);
  const [datasetId, setDatasetId] = useState<string>("");
  const [uploadingData, setUploadingData] = useState(false);
  const [dataError, setDataError] = useState("");

  // Question state
  const [questionTab, setQuestionTab] = useState<QuestionTab>("paste");
  const [questionText, setQuestionText] = useState("");
  const [questionFile, setQuestionFile] = useState<File | null>(null);
  const [assignmentId, setAssignmentId] = useState("");
  const [uploadingQuestion, setUploadingQuestion] = useState(false);
  const [questionError, setQuestionError] = useState("");

  // Examples
  const [examples, setExamples] = useState<
    Record<string, { title: string; question: string }>
  >({});
  const [dragActiveData, setDragActiveData] = useState(false);

  const dataInputRef = useRef<HTMLInputElement>(null);
  const questionInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    getExamples()
      .then(setExamples)
      .catch(() => {});
  }, []);

  // ─── Dataset Upload ────────────────────────────────────────────

  const handleDataDrop = useCallback(
    async (e: React.DragEvent) => {
      e.preventDefault();
      setDragActiveData(false);
      const file = e.dataTransfer.files?.[0];
      if (file) await handleDataFile(file);
    },
    []
  );

  const handleDataFile = async (file: File) => {
    const ext = file.name.split(".").pop()?.toLowerCase();
    if (!["xlsx", "xls", "csv"].includes(ext || "")) {
      setDataError("Unsupported format. Use .xlsx, .xls, or .csv");
      return;
    }
    if (file.size > 25 * 1024 * 1024) {
      setDataError("File too large. Maximum 25 MB.");
      return;
    }
    setDataFile(file);
    setDataError("");
    setUploadingData(true);

    try {
      const res = await uploadDataset(file);
      setDatasetId(res.dataset_id);
      setDataProfile(res.profile);
    } catch (err: unknown) {
      setDataError(err instanceof Error ? err.message : "Upload failed");
      setDataFile(null);
    } finally {
      setUploadingData(false);
    }
  };

  // ─── Question Submit ───────────────────────────────────────────

  const handleQuestionSubmit = async () => {
    if (!datasetId) {
      setQuestionError("Upload a dataset first");
      return;
    }

    setUploadingQuestion(true);
    setQuestionError("");

    try {
      if (questionTab === "paste" && questionText.trim()) {
        const res = await pasteQuestion(datasetId, questionText.trim());
        setAssignmentId(res.assignment_id);
      } else if (
        (questionTab === "upload" || questionTab === "examples") &&
        questionFile
      ) {
        const res = await uploadQuestion(datasetId, questionFile);
        setAssignmentId(res.assignment_id);
      } else if (questionTab === "paste" && !questionText.trim()) {
        setQuestionError("Please enter a question");
        setUploadingQuestion(false);
        return;
      } else {
        setQuestionError("Please provide a question");
        setUploadingQuestion(false);
        return;
      }
    } catch (err: unknown) {
      setQuestionError(err instanceof Error ? err.message : "Submit failed");
      setUploadingQuestion(false);
      return;
    }

    setUploadingQuestion(false);
  };

  // ─── Navigate to Analyze ───────────────────────────────────────

  const handleAnalyze = () => {
    if (!datasetId || !assignmentId) return;
    const params = new URLSearchParams({
      dataset_id: datasetId,
      assignment_id: assignmentId,
    });
    router.push(`/analyze?${params.toString()}`);
  };

  // ─── Render ────────────────────────────────────────────────────

  return (
    <div className="min-h-screen">
      {/* Navbar */}
      <nav className="sticky top-0 z-50 backdrop-blur-xl bg-[var(--color-bg)]/80 border-b border-[var(--color-border-subtle)]">
        <div className="max-w-7xl mx-auto px-6 h-16 flex items-center justify-between">
          <Link
            href="/"
            className="flex items-center gap-2 no-underline"
          >
            <span className="text-2xl">📊</span>
            <span className="text-xl font-bold gradient-text">ExcelFlow</span>
          </Link>
          <span className="text-sm text-[var(--color-text-muted)]">
            Step 1 of 4 — Upload
          </span>
        </div>
      </nav>

      <div className="max-w-7xl mx-auto px-6 py-10">
        <h1 className="text-3xl font-bold mb-2">Upload Your Data & Assignment</h1>
        <p className="text-[var(--color-text-muted)] mb-10">
          Provide your raw dataset and the assignment question to get started.
        </p>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
          {/* ─── Left: Dataset ──────────────────────────── */}
          <div className="glass-card-static p-6">
            <h2 className="text-lg font-semibold mb-4 flex items-center gap-2">
              <span className="w-7 h-7 rounded-full bg-[var(--color-primary)] text-white text-xs flex items-center justify-center font-bold">
                1
              </span>
              Upload Dataset
            </h2>

            <div
              className={`dropzone ${dragActiveData ? "active" : ""} ${dataProfile ? "has-file" : ""}`}
              onDragOver={(e) => {
                e.preventDefault();
                setDragActiveData(true);
              }}
              onDragLeave={() => setDragActiveData(false)}
              onDrop={handleDataDrop}
              onClick={() => dataInputRef.current?.click()}
            >
              <input
                ref={dataInputRef}
                type="file"
                className="hidden"
                accept=".xlsx,.xls,.csv"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) handleDataFile(f);
                }}
              />

              {uploadingData ? (
                <div className="flex flex-col items-center gap-3">
                  <div className="w-10 h-10 border-3 border-[var(--color-primary)] border-t-transparent rounded-full animate-spin" />
                  <p className="text-sm text-[var(--color-text-secondary)]">
                    Analyzing dataset...
                  </p>
                </div>
              ) : dataProfile ? (
                <div className="text-left space-y-2">
                  <div className="flex items-center gap-2 text-[var(--color-success)]">
                    <span>✓</span>
                    <span className="font-semibold">
                      {dataProfile.filename}
                    </span>
                  </div>
                  <div className="grid grid-cols-2 gap-2 text-sm text-[var(--color-text-secondary)]">
                    <span>
                      📄 {dataFile?.size ? (dataFile.size / 1024).toFixed(1) + " KB" : ""}
                    </span>
                    <span>📊 {dataProfile.row_count} rows</span>
                    <span>📐 {dataProfile.column_count} columns</span>
                    <span>🔑 {dataProfile.dimensions.length} dimensions</span>
                  </div>
                  <p className="text-xs text-[var(--color-text-muted)] mt-2">
                    Click or drop to replace
                  </p>
                </div>
              ) : (
                <div className="flex flex-col items-center gap-3">
                  <span className="text-4xl opacity-40">📁</span>
                  <p className="font-medium">
                    Drag & Drop Excel / CSV
                  </p>
                  <p className="text-sm text-[var(--color-text-muted)]">
                    .xlsx, .xls, .csv — max 25 MB
                  </p>
                  <button className="btn-secondary text-sm py-2 px-6 mt-2">
                    Choose File
                  </button>
                </div>
              )}
            </div>

            {dataError && (
              <p className="text-[var(--color-error)] text-sm mt-3">
                {dataError}
              </p>
            )}

            {/* Dataset Profile Preview */}
            {dataProfile && (
              <div className="mt-6 space-y-4">
                <h3 className="text-sm font-semibold text-[var(--color-text-secondary)] uppercase tracking-wider">
                  Detected Schema
                </h3>

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <p className="text-xs text-[var(--color-text-muted)] mb-1">
                      Dimensions
                    </p>
                    {dataProfile.dimensions.map((d) => (
                      <span
                        key={d}
                        className="inline-block text-xs px-2 py-1 rounded bg-blue-500/10 text-blue-400 border border-blue-500/20 mr-1 mb-1"
                      >
                        {d}
                      </span>
                    ))}
                  </div>
                  <div>
                    <p className="text-xs text-[var(--color-text-muted)] mb-1">
                      Measures
                    </p>
                    {dataProfile.measures.map((m) => (
                      <span
                        key={m}
                        className="inline-block text-xs px-2 py-1 rounded bg-green-500/10 text-green-400 border border-green-500/20 mr-1 mb-1"
                      >
                        {m}
                      </span>
                    ))}
                  </div>
                </div>

                {dataProfile.potential_calculations.length > 0 && (
                  <div>
                    <p className="text-xs text-[var(--color-text-muted)] mb-1">
                      Potential Calculations
                    </p>
                    {dataProfile.potential_calculations.map((c) => (
                      <span
                        key={c}
                        className="inline-block text-xs px-2 py-1 rounded bg-purple-500/10 text-purple-400 border border-purple-500/20 mr-1 mb-1"
                      >
                        {c}
                      </span>
                    ))}
                  </div>
                )}
              </div>
            )}
          </div>

          {/* ─── Right: Question ──────────────────────────── */}
          <div className="glass-card-static p-6">
            <h2 className="text-lg font-semibold mb-4 flex items-center gap-2">
              <span className="w-7 h-7 rounded-full bg-[var(--color-accent)] text-white text-xs flex items-center justify-center font-bold">
                2
              </span>
              Assignment / Question
            </h2>

            {/* Tabs */}
            <div className="tab-group mb-4">
              {(["paste", "upload", "examples"] as QuestionTab[]).map((tab) => (
                <button
                  key={tab}
                  className={`tab-btn ${questionTab === tab ? "active" : ""}`}
                  onClick={() => setQuestionTab(tab)}
                >
                  {tab === "paste"
                    ? "Paste Text"
                    : tab === "upload"
                      ? "Upload File"
                      : "Examples"}
                </button>
              ))}
            </div>

            {/* Paste */}
            {questionTab === "paste" && (
              <div>
                <textarea
                  className="input-field"
                  placeholder="Paste your assignment/question here...&#10;&#10;Example:&#10;Create a Sales Dashboard using the given dataset.&#10;1. Revenue by Region&#10;2. Revenue by Product&#10;3. Revenue by Month"
                  rows={8}
                  value={questionText}
                  onChange={(e) => setQuestionText(e.target.value)}
                />
              </div>
            )}

            {/* Upload */}
            {questionTab === "upload" && (
              <div
                className={`dropzone ${questionFile ? "has-file" : ""}`}
                onClick={() => questionInputRef.current?.click()}
              >
                <input
                  ref={questionInputRef}
                  type="file"
                  className="hidden"
                  accept=".pdf,.docx,.txt,.png,.jpg,.jpeg,.webp"
                  onChange={(e) => {
                    const f = e.target.files?.[0];
                    if (f) setQuestionFile(f);
                  }}
                />
                {questionFile ? (
                  <div className="flex items-center gap-2 text-[var(--color-success)]">
                    <span>✓</span>
                    <span className="font-semibold">{questionFile.name}</span>
                    <span className="text-xs text-[var(--color-text-muted)]">
                      ({(questionFile.size / 1024).toFixed(1)} KB)
                    </span>
                  </div>
                ) : (
                  <div className="flex flex-col items-center gap-3">
                    <span className="text-4xl opacity-40">📄</span>
                    <p className="font-medium">Upload Question File</p>
                    <p className="text-sm text-[var(--color-text-muted)]">
                      PDF, DOCX, TXT, PNG, JPG, WEBP
                    </p>
                  </div>
                )}
              </div>
            )}

            {/* Examples */}
            {questionTab === "examples" && (
              <div className="grid grid-cols-1 gap-2 max-h-[300px] overflow-y-auto">
                {Object.entries(examples).map(([key, ex]) => (
                  <button
                    key={key}
                    className={`plan-item text-left ${
                      questionText === ex.question
                        ? "border-[var(--color-primary)] bg-[var(--color-primary-glow)]"
                        : ""
                    }`}
                    onClick={() => {
                      setQuestionText(ex.question);
                      setQuestionTab("paste");
                    }}
                  >
                    <span className="text-lg">
                      {key.includes("sales")
                        ? "💰"
                        : key.includes("employee")
                          ? "👤"
                          : key.includes("financial")
                            ? "💵"
                            : key.includes("inventory")
                              ? "📦"
                              : key.includes("marketing")
                                ? "📣"
                                : "📊"}
                    </span>
                    <div>
                      <p className="font-medium text-sm">{ex.title}</p>
                      <p className="text-xs text-[var(--color-text-muted)] truncate max-w-[250px]">
                        {ex.question.split("\n")[0]}
                      </p>
                    </div>
                  </button>
                ))}
              </div>
            )}

            {questionError && (
              <p className="text-[var(--color-error)] text-sm mt-3">
                {questionError}
              </p>
            )}

            {/* Submit Question */}
            {!assignmentId && (
              <button
                className="btn-primary w-full mt-4"
                disabled={
                  !datasetId ||
                  uploadingQuestion ||
                  (questionTab === "paste" && !questionText.trim()) ||
                  (questionTab === "upload" && !questionFile)
                }
                onClick={handleQuestionSubmit}
              >
                {uploadingQuestion ? "Processing..." : "Submit Question"}
              </button>
            )}

            {assignmentId && (
              <div className="mt-4 p-3 rounded-lg bg-[var(--color-success)]/10 border border-[var(--color-success)]/20">
                <p className="text-[var(--color-success)] text-sm font-medium">
                  ✓ Question submitted successfully
                </p>
              </div>
            )}
          </div>
        </div>

        {/* ─── Analyze CTA ──────────────────────────────────── */}
        {datasetId && assignmentId && (
          <div className="mt-10 text-center animate-fade-in-up">
            <button
              className="btn-primary text-lg px-12 py-4"
              onClick={handleAnalyze}
            >
              Analyze & Generate Plan →
            </button>
            <p className="text-sm text-[var(--color-text-muted)] mt-3">
              The AI will analyze your dataset and assignment to create a
              generation plan.
            </p>
          </div>
        )}
      </div>
    </div>
  );
}
