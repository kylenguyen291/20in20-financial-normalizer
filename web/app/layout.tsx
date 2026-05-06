import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Financial Normalizer — Vietnamese Stock Reports",
  description:
    "Automatically download, extract, and normalize Vietnamese company financial statements into structured Excel workbooks. Enter a stock ticker, get a clean Excel file.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body suppressHydrationWarning>{children}</body>
    </html>
  );
}
