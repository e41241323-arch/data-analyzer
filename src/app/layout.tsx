import type { Metadata, Viewport } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: "Excel Analyzer | Analisis Excel Otomatis",
  description: "Analisis data Excel secara otomatis dengan filter, statistik, dan grafik yang rapi untuk kebutuhan reporting.",
  keywords: ["excel analyzer", "analisis excel", "dashboard excel", "vercel app"],
  metadataBase: new URL("https://excel-analyzer.vercel.app"),
  openGraph: {
    title: "Excel Analyzer",
    description: "Analisis data Excel secara otomatis dengan filter, statistik, dan grafik yang rapi.",
    type: "website",
  },
};

export const viewport: Viewport = {
  width: "device-width",
  initialScale: 1,
  maximumScale: 1,
  viewportFit: "cover",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html
      lang="en"
      className={`${geistSans.variable} ${geistMono.variable} h-full antialiased`}
    >
      <body className="min-h-full flex flex-col">{children}</body>
    </html>
  );
}
