/**
 * ExcelFlow API Client — typed HTTP helpers for all backend endpoints.
 */

const getApiBase = (): string => {
  if (process.env.NEXT_PUBLIC_API_URL) {
    return process.env.NEXT_PUBLIC_API_URL.replace(/\/$/, "");
  }
  // In browser on production (e.g. Vercel), default to relative path so Next.js rewrites forward to backend
  if (typeof window !== "undefined" && window.location.hostname !== "localhost" && window.location.hostname !== "127.0.0.1") {
    return "";
  }
  return "http://localhost:8000";
};

const API_BASE = getApiBase();

// ─── Types ──────────────────────────────────────────────────────────

export interface ColumnProfile {
  name: string;
  data_type: string;
  semantic_type: string;
  classification: string;
  unique_count: number;
  missing_count: number;
  sample_values: unknown[];
  min_value: string | null;
  max_value: string | null;
  mean_value: number | null;
}

export interface DatasetProfile {
  dataset_id: string;
  filename: string;
  row_count: number;
  column_count: number;
  headers: string[];
  columns: ColumnProfile[];
  dimensions: string[];
  measures: string[];
  potential_calculations: string[];
  duplicate_rows: number;
}

export interface DatasetUploadResponse {
  dataset_id: string;
  filename: string;
  row_count: number;
  column_count: number;
  headers: string[];
  profile: DatasetProfile;
}

export interface QuestionUploadResponse {
  assignment_id: string;
  question_text: string;
  source_type: string;
}

export interface CalculationSpec {
  name: string;
  formula_template: string;
  format: string;
  description: string;
}

export interface AnalysisSpec {
  name: string;
  metric: string;
  aggregation: string;
  group_by: string[];
  sort: string;
  description: string;
}

export interface ChartSpec {
  name: string;
  chart_type: string;
  title: string;
  data_source: string;
  category_field: string;
  value_field: string;
  x_axis_title: string;
  y_axis_title: string;
  width: number;
  height: number;
}

export interface KPISpec {
  name: string;
  metric: string;
  aggregation: string;
  format: string;
  label: string;
}

export interface FilterSpec {
  field: string;
  filter_type: string;
  description: string;
}

export interface DashboardSpec {
  title: string;
  kpis: KPISpec[];
  chart_refs: string[];
  filter_refs: string[];
}

export interface SheetSpec {
  name: string;
  sheet_type: string;
  description: string;
}

export interface ExcelPlan {
  workbook_title: string;
  calculations: CalculationSpec[];
  analyses: AnalysisSpec[];
  charts: ChartSpec[];
  dashboard: DashboardSpec | null;
  filters: FilterSpec[];
  sheets: SheetSpec[];
}

export interface AnalyzeResponse {
  generation_id: string;
  plan: ExcelPlan;
  dataset_profile: DatasetProfile;
}

export interface GenerationProgress {
  step: string;
  status: string;
  message: string;
}

export interface ValidationCheck {
  check_name: string;
  passed: boolean;
  message: string;
  details: string[];
}

export interface ValidationResult {
  workbook_id: string;
  overall_status: string;
  checks: ValidationCheck[];
  formula_count: number;
  chart_count: number;
  sheet_count: number;
  error_count: number;
}

export interface WorkbookSummary {
  workbook_id: string;
  filename: string;
  sheets: string[];
  row_count: number;
  formula_count: number;
  chart_count: number;
  analysis_count: number;
  filter_count: number;
  native_pivots?: boolean;
  pivots_created?: number;
  slicers_created?: number;
  validation: ValidationResult | null;
}

export interface GenerationStatus {
  generation_id: string;
  status: string;
  progress: GenerationProgress[];
  error: string | null;
  workbook_id: string | null;
}

// ─── API Functions ──────────────────────────────────────────────────

export async function uploadDataset(file: File): Promise<DatasetUploadResponse> {
  const formData = new FormData();
  formData.append("file", file);

  const res = await fetch(`${API_BASE}/api/datasets/upload`, {
    method: "POST",
    body: formData,
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: res.statusText }));
    throw new Error(err.detail || "Upload failed");
  }

  return res.json();
}

export async function uploadQuestion(datasetId: string, file: File): Promise<QuestionUploadResponse> {
  const formData = new FormData();
  formData.append("dataset_id", datasetId);
  formData.append("file", file);

  const res = await fetch(`${API_BASE}/api/questions/upload`, {
    method: "POST",
    body: formData,
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: res.statusText }));
    throw new Error(err.detail || "Upload failed");
  }

  return res.json();
}

export async function pasteQuestion(datasetId: string, questionText: string): Promise<QuestionUploadResponse> {
  const res = await fetch(`${API_BASE}/api/questions/paste`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ dataset_id: datasetId, question_text: questionText }),
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: res.statusText }));
    throw new Error(err.detail || "Paste failed");
  }

  return res.json();
}

export async function analyze(datasetId: string, assignmentId: string): Promise<AnalyzeResponse> {
  const res = await fetch(`${API_BASE}/api/analyze`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ dataset_id: datasetId, assignment_id: assignmentId }),
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: res.statusText }));
    throw new Error(err.detail || "Analysis failed");
  }

  return res.json();
}

export async function generateWorkbook(generationId: string, plan: ExcelPlan): Promise<{ generation_id: string; status: string }> {
  const res = await fetch(`${API_BASE}/api/workbooks/generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ generation_id: generationId, plan }),
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: res.statusText }));
    throw new Error(err.detail || "Generation failed");
  }

  return res.json();
}

export async function getGenerationStatus(generationId: string): Promise<GenerationStatus> {
  const res = await fetch(`${API_BASE}/api/workbooks/${generationId}/status`);

  if (!res.ok) {
    throw new Error("Failed to get status");
  }

  return res.json();
}

export async function getWorkbookSummary(workbookId: string): Promise<WorkbookSummary> {
  const res = await fetch(`${API_BASE}/api/workbooks/${workbookId}/summary`);

  if (!res.ok) {
    throw new Error("Failed to get summary");
  }

  return res.json();
}

export function getDownloadUrl(workbookId: string): string {
  return `${API_BASE}/api/workbooks/${workbookId}/download`;
}

export async function getExamples(): Promise<Record<string, { title: string; question: string }>> {
  const res = await fetch(`${API_BASE}/api/examples`);
  if (!res.ok) throw new Error("Failed to load examples");
  return res.json();
}
