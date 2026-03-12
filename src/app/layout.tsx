import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "PPT 生成",
  description: "根据固定模板生成 PPT。",
  icons: {
    icon: "/favicon.svg"
  }
};

export default function RootLayout({
  children
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
