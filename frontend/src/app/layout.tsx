import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "ExcelFlow — Intelligent Excel Assignment Automation Platform",
  description:
    "Upload raw data and an assignment prompt. ExcelFlow automatically generates real formulas, authentic native PivotTables, interactive Slicers, charts, and executive dashboards in an Excel workbook.",
  keywords: [
    "Excel Automation",
    "Pivot Tables",
    "Data Analysis",
    "AI Excel Generator",
    "Interactive Dashboards",
    "Slicers",
    "FastAPI",
    "Next.js",
  ],
  authors: [{ name: "ExcelFlow Team" }],
  openGraph: {
    title: "ExcelFlow — Intelligent Excel Assignment Automation Platform",
    description:
      "Convert raw CSV/Excel datasets and questions into fully calculated workbooks with native interactive PivotTables and executive dashboards.",
    type: "website",
    locale: "en_US",
    siteName: "ExcelFlow",
  },
  twitter: {
    card: "summary_large_image",
    title: "ExcelFlow — Turn Raw Data Into Complete Excel Workbooks",
    description:
      "Automate calculations, PivotTables, Slicers, and dashboards with AI-driven intelligence.",
  },
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <head>
        <link rel="preconnect" href="https://fonts.googleapis.com" />
        <link
          rel="preconnect"
          href="https://fonts.gstatic.com"
          crossOrigin="anonymous"
        />
      </head>
      <body className="min-h-screen bg-[var(--color-bg)]">{children}</body>
    </html>
  );
}
