import { NextResponse } from "next/server";
import PptxGenJS from "pptxgenjs";
import { db } from "@/server/db";
import {
  themes,
  type ThemeName,
  type ThemeProperties,
} from "@/lib/presentation/themes";
import { type TDescendant, type TElement } from "@udecode/plate-common";

// --- START: Type definitions from your project ---
type LayoutType = "left" | "right" | "vertical";

type PlateNode = TElement & {
  type: string;
  children: TDescendant[];
  [key: string]: any;
};

type PlateSlide = {
  id: string;
  content: PlateNode[];
  rootImage?: {
    query: string;
    url?: string;
  };
  layoutType?: LayoutType;
  alignment?: "start" | "center" | "end";
  bgColor?: string;
  width?: "L" | "M";
};

interface Theme {
  name: string;
  properties: ThemeProperties;
}

interface ContentArea {
  x: PptxGenJS.Coord;
  y: PptxGenJS.Coord;
  w: PptxGenJS.Coord;
  h: PptxGenJS.Coord;
}
// --- END: Type definitions ---

function resolveTheme(
  themeData: string | ThemeProperties | Theme | undefined,
): Theme | null {
  if (typeof themeData === "object" && themeData !== null) {
    if ("properties" in themeData && "name" in themeData) {
      return themeData as Theme;
    }
    return { name: "Custom", properties: themeData as ThemeProperties };
  }
  if (typeof themeData === "string") {
    const themeName = themeData as ThemeName;
    if (themes[themeName]) {
      return { name: themeName, properties: themes[themeName] };
    }
  }
  return null;
}

function getRichTextFromNode(
  node: TElement,
  baseOptions: PptxGenJS.TextPropsOptions,
): PptxGenJS.TextProps[] {
  const richText: PptxGenJS.TextProps[] = [];

  function traverse(
    children: TDescendant[],
    currentOptions: PptxGenJS.TextPropsOptions,
  ) {
    for (const child of children) {
      const newOptions: PptxGenJS.TextPropsOptions = { ...currentOptions };
      if ((child as any).bold) newOptions.bold = true;
      if ((child as any).italic) newOptions.italic = true;
      if ((child as any).underline) newOptions.underline = true;
      if ((child as any).strikethrough) newOptions.strike = true;

      if ((child as any).text) {
        if ((child as any).text.trim()) {
             richText.push({ text: (child as any).text, options: newOptions });
        }
      } else if ((child as TElement).children) {
        traverse((child as TElement).children as TDescendant[], newOptions);
      }
    }
  }

  if (node.children) {
    traverse(node.children as TDescendant[], baseOptions);
  }

  return richText;
}

function getTextNodeStyle(
  node: PlateNode,
  theme: Theme,
): PptxGenJS.TextPropsOptions {
  const colors = theme.properties.colors.light;
  const fonts = theme.properties.fonts;
  const baseOptions: PptxGenJS.TextPropsOptions = {
    color: colors.text.replace("#", ""),
    fontFace: fonts.body,
    fontSize: 16,
    valign: "top",
  };

  switch (node.type) {
    case "h1":
      return { ...baseOptions, color: colors.heading.replace("#", ""), fontFace: fonts.heading, fontSize: 36, bold: true };
    case "h2":
      return { ...baseOptions, color: colors.heading.replace("#", ""), fontFace: fonts.heading, fontSize: 28, bold: true };
    case "h3":
      return { ...baseOptions, color: colors.heading.replace("#", ""), fontFace: fonts.heading, fontSize: 24, bold: true };
    case "h4":
      return { ...baseOptions, color: colors.heading.replace("#", ""), fontFace: fonts.heading, fontSize: 20, bold: true };
    case "p":
      return { ...baseOptions, fontSize: 18, lineSpacing: 28 };
    case "bullet":
      return { ...baseOptions, fontSize: 16, bullet: {type: 'bullet'} };
    default:
      return baseOptions;
  }
}

function processNode(
    node: PlateNode,
    slide: PptxGenJS.Slide,
    theme: Theme,
    area: ContentArea,
    y: number,
): number {
    const style = getTextNodeStyle(node, theme);
    const richText = getRichTextFromNode(node, style);

    if (!richText.some(t => t.text?.trim())) return y;
    
    // Better height estimation
    const textContent = richText.map(t => t.text).join('');
    const lineCount = (textContent.match(/\n/g) || []).length + 1;
    const estimatedHeight = (lineCount * (style.fontSize ?? 16) * 1.5) / 72; // pixels to inches
    const elemHeight = Math.max(estimatedHeight, 0.5);

    switch (node.type) {
        case "h1": case "h2": case "h3": case "h4": case "p":
            slide.addText(richText, {
                x: area.x,
                y: y,
                w: area.w,
                h: elemHeight,
                autoFit: true,
                align: node.type.startsWith('h') ? 'center' : 'left'
            });
            return y + elemHeight + 0.1;

        case "img":
            if (node.url) {
                slide.addImage({ path: node.url, x: area.x, y, w: area.w, h: 2, sizing: { type: 'contain', w: area.w, h: 2 } });
                return y + 2.2;
            }
            return y;
            
        case "column":
            const columns = node.children.filter(child => child.type === 'column_item');
            if (!columns.length) return y;
            
            const colW = ((area.w as number) - (columns.length -1) * 0.2) / columns.length;
            let maxColY = y;

            columns.forEach((col, index) => {
                const colArea: ContentArea = {
                    ...area,
                    x: (area.x as number) + index * (colW + 0.2),
                    w: colW,
                };
                let currentYinCol = y;
                (col.children as PlateNode[]).forEach(childNode => {
                    currentYinCol = processNode(childNode, slide, theme, colArea, currentYinCol);
                });
                maxColY = Math.max(maxColY, currentYinCol);
            });
            return maxColY + 0.2;
            
        case "bullets":
            let currentYBullets = y;
            (node.children as PlateNode[]).forEach(childNode => {
                currentYBullets = processNode(childNode, slide, theme, area, currentYBullets);
            });
            return currentYBullets;

        case "bullet":
            slide.addText(richText, {
                x: (area.x as number) + 0.25,
                y,
                w: (area.w as number) - 0.25,
                h: elemHeight,
                autoFit: true,
                bullet: true
            });
            return y + elemHeight + 0.05;

        default:
            return y;
    }
}


function addSlideContent(
    pptx: PptxGenJS,
    slide: PptxGenJS.Slide,
    slideData: PlateSlide,
    theme: Theme,
) {
  // PptxGenJS 16:9 layout is 10" x 5.625"
  let contentArea: ContentArea = { x: 0.5, y: 0.5, w: 9.0, h: 4.625 };

  if (slideData.rootImage?.url) {
    const imagePath = slideData.rootImage.url;
    try {
        switch (slideData.layoutType) {
            case "left":
              slide.addImage({ path: imagePath, x: 0, y: 0, w: "50%", h: "100%", sizing: { type: 'cover', w: '50%', h: '100%'} });
              contentArea = { x: 5.25, y: 0.5, w: 4.5, h: 4.625 };
              break;
            case "right":
              slide.addImage({ path: imagePath, x: "50%", y: 0, w: "50%", h: "100%", sizing: { type: 'cover', w: '50%', h: '100%'} });
              contentArea = { x: 0.25, y: 0.5, w: 4.5, h: 4.625 };
              break;
            case "vertical":
              slide.addImage({ path: imagePath, x: 0, y: 0, w: "100%", h: "50%", sizing: { type: 'cover', w: '100%', h: '50%'} });
              contentArea = { x: 0.5, y: 3.0, w: 9.0, h: 2.125 };
              break;
            default: // Full background image with overlay
              slide.addImage({ path: imagePath, x: 0, y: 0, w: "100%", h: "100%", sizing: { type: 'cover', w: '100%', h: '100%'} });
              // Add a semi-transparent overlay to ensure text is readable
              slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: '000000', transparency: 30 } });
              // Adjust theme for dark background
              theme.properties.colors.light.text = "#FFFFFF";
              theme.properties.colors.light.heading = "#FFFFFF";
              break;
        }
    } catch(e) {
        console.error(`Error adding root image. It might be an invalid URL or path: ${imagePath}`, e);
        slide.addText(`[Image not found]`, {x: 0, y:0, w: '50%', h: '100%'});
    }
  }

  let currentY = contentArea.y as number;

  for (const node of slideData.content) {
    currentY = processNode(
        node,
        slide,
        theme,
        contentArea,
        currentY,
    );
  }
}


export async function POST(req: Request) {
  try {
    const { id } = await req.json();
    if (!id) {
      return NextResponse.json(
        { error: "Invalid or missing presentation ID" },
        { status: 400 },
      );
    }
    const presentation = await db.baseDocument.findUnique({
      where: { id },
      include: {
        presentation: true,
      },
    });

    if (!presentation?.presentation) {
      return NextResponse.json(
        { error: "Presentation not found" },
        { status: 404 },
      );
    }

    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_16x9";

    const baseTheme = resolveTheme(presentation.presentation.theme) ?? {
      name: "daktilo",
      properties: themes.daktilo,
    };
    const slides = (
      presentation.presentation.content as { slides: PlateSlide[] }
    ).slides;

    for (const slideData of slides) {
      // Deep clone the theme for each slide to avoid mutation issues
      const slideTheme = JSON.parse(JSON.stringify(baseTheme));
      const slide = pptx.addSlide();
      
      const themeColors = slideTheme.properties.colors.light;
      
      let finalBgColor = slideData.bgColor ?? themeColors.background;
      
      slide.background = { color: finalBgColor.replace("#", "") };
      
      // If the background is dark, we should use light text.
      // A simple heuristic for darkness.
      const bgHex = finalBgColor.replace("#", "");
      const r = parseInt(bgHex.substring(0,2), 16);
      const g = parseInt(bgHex.substring(2,4), 16);
      const b = parseInt(bgHex.substring(4,6), 16);
      const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;

      if (luminance < 0.5) {
          slideTheme.properties.colors.light.text = "#FFFFFF";
          slideTheme.properties.colors.light.heading = "#FFFFFF";
          slideTheme.properties.colors.light.muted = "#E5E7EB";
      }

      addSlideContent(pptx, slide, slideData, slideTheme);
    }

    const pptxBuffer = await pptx.write({ outputType: "nodebuffer" });
    const fileName = `${presentation.title || "presentation"}.pptx`;

    return new NextResponse(pptxBuffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": `attachment; filename="${fileName}"`,
      },
    });
  } catch (e) {
    console.error("PPT Download Error:", e);
    const errorMessage =
      e instanceof Error ? e.message : "An unknown error occurred";
    return NextResponse.json(
      { error: "Internal server error", details: errorMessage },
      { status: 500 },
    );
  }
}