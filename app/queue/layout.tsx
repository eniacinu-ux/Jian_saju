import type { Viewport } from "next";
export const viewport: Viewport = { width: "device-width", initialScale: 1, maximumScale: 5 };
export default function QueueLayout({ children }: { children: React.ReactNode }) { return children; }
