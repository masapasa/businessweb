"use client";
import { useMemo, useState } from "react";
import { Download } from "lucide-react";
import { Button } from "@/components/ui/button";
import { usePresentationSlides } from "@/hooks/presentation/usePresentationSlides";
import { usePresentationState } from "@/states/presentation-state";
import { useTheme } from "next-themes";
import { themes, type ThemeProperties } from "@/lib/presentation/themes";

export function DownloadPPT() {
  const { items } = usePresentationSlides();
  const [loading, setLoading] = useState(false);
  const { presentationInput, theme, customThemeData } = usePresentationState();
  const { resolvedTheme } = useTheme();

  const exportTheme = useMemo(() => {
    let data: ThemeProperties | null = null;
    if (customThemeData) {
      data = customThemeData;
    } else if (typeof theme === "string" && theme in themes) {
      data = themes[theme as keyof typeof themes];
    }
    if (!data) return null;
    const isDark = resolvedTheme === "dark";
    return {
      colors: isDark ? data.colors.dark : data.colors.light,
      fonts: data.fonts,
    };
  }, [customThemeData, theme, resolvedTheme]);

  const handleDownload = async () => {
    setLoading(true);
    try {
      // Debug: inspect the payload we're about to send
      // Logs the number of sections and the first section for quick inspection
      // Remove once verified
      // eslint-disable-next-line no-console
      console.log("[DownloadPPT] payload", {
        sectionsCount: items.length,
        firstSection: items[0],
        theme: exportTheme,
      });

      const res = await fetch("/api/presentation/download", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sections: items, theme: exportTheme }),
      });
      if (res.ok) {
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;

        const contentDisposition = res.headers.get("content-disposition");
        let fileName = `${presentationInput || "presentation"}.pptx`;
        if (contentDisposition) {
          const match = contentDisposition.match(/filename="?([^"]+)"?/i);
          if (match?.[1]) {
            fileName = match[1];
          }
        }
        link.download = fileName;

        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
      } else {
        alert("Download failed. Please try again later.");
      }
    } catch (e) {
      alert("Download failed. Please try again later.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <Button
      variant="ghost"
      size="sm"
      className="text-muted-foreground hover:text-foreground"
      onClick={handleDownload}
      disabled={loading}
    >
      <Download className="mr-1 h-4 w-4" />
      {loading ? "Downloading..." : "Download PPT"}
    </Button>
  );
}