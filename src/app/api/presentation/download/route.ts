import { NextResponse } from "next/server";
import PptxGenJS from "pptxgenjs";
import { db } from "@/server/db";
import {
  themes,
  type ThemeName,
  type ThemeProperties,
} from "@/lib/presentation/themes";
import { type TDescendant, type TElement } from "@udecode/plate-common";

// --- START: Type definitions ---
// These types are based on your project structure to ensure the logic here matches your data.

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

// This is a simplified version of the Theme type from your project
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

interface ProcessNodeParams {
  node: PlateNode;
  slide: PptxGenJS.Slide;
  theme: Theme;
  area: ContentArea;
  y: number;
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

function getRichTextFromNode(node: TElement): PptxGenJS.TextProps[] {
  const richText: PptxGenJS.TextProps[] = [];

  function traverse(
    children: TDescendant[],
    options: PptxGenJS.TextPropsOptions,
  ) {
    for (const child of children) {
      const newOptions: PptxGenJS.TextPropsOptions = { ...options };
      if ((child as TElement).bold) newOptions.bold = true;
      if ((child as TElement).italic) newOptions.italic = true;
      if ((child as TElement).underline) newOptions.underline = true;
      if ((child as TElement).strikethrough) newOptions.strike = true;

      if ((child as any).text) {
        richText.push({ text: (child as any).text, options: newOptions });
      } else if ((child as TElement).children) {
        traverse((child as TElement).children as TDescendant[], newOptions);
      }
    }
  }

  if (node.children) {
    traverse(node.children as TDescendant[], {});
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
      return { ...baseOptions, color: colors.heading.replace("#", ""), fontFace: fonts.heading, fontSize: 36, bold: true, autoFit: true };
    case "h2":
      return { ...baseOptions, color: colors.heading.replace("#", ""), fontFace: fonts.heading, fontSize: 28, bold: true, autoFit: true };
    case "h3":
      return { ...baseOptions, color: colors.heading.replace("#", ""), fontFace: fonts.heading, fontSize: 24, bold: true, autoFit: true };
    case "h4":
      return { ...baseOptions, color: colors.heading.replace("#", ""), fontFace: fonts.heading, fontSize: 20, bold: true, autoFit: true };
    case "h5":
      return { ...baseOptions, color: colors.heading.replace("#", ""), fontFace: fonts.heading, fontSize: 18, autoFit: true };
    case "h6":
      return { ...baseOptions, color: colors.muted.replace("#", ""), fontSize: 16, autoFit: true };
    case "p":
      return { ...baseOptions, fontSize: 14, autoFit: true };
    case "bullet":
        return { ...baseOptions, fontSize: 14, bullet: {type: 'bullet'}, autoFit: true };
    default:
      return baseOptions;
  }
}


function addTextElement(
  node: PlateNode,
  slide: PptxGenJS.Slide,
  theme: Theme,
  area: ContentArea,
  y: number,
  slideData: PlateSlide,
): number {
  const richText = getRichTextFromNode(node);
  if (richText.every(t => !t.text?.trim())) return y;

  const style = getTextNodeStyle(node, theme);
  const textH = (style.fontSize ?? 14) / 72 * 1.5 * richText.length; // Approximate height

  if (node.type.startsWith('h')) {
      style.align = 'center';
  }

  const options: PptxGenJS.TextPropsOptions = {
      ...style,
      x: area.x,
      y: y,
      w: area.w,
      h: textH,
  }

  if (node.type === 'bullet') {
      options.x = (area.x as number) + 0.25;
      options.w = (area.w as number) - 0.25;
  }

  slide.addText(richText, options);
  return y + textH + 0.1;
}

function addImageElement(
  node: PlateNode,
  slide: PptxGenJS.Slide,
  area: ContentArea,
  y: number,
): number {
  if (node.url) {
    const imgH = 2; // Fixed height for inline images
    slide.addImage({
      path: node.url,
      x: area.x,
      y,
      w: area.w,
      h: imgH,
      sizing: { type: "contain", w: area.w, h: imgH },
    });
    return y + imgH + 0.2;
  }
  return y;
}

function addColumnsElement(
    node: PlateNode,
    slide: PptxGenJS.Slide,
    theme: Theme,
    area: ContentArea,
    y: number,
    slideData: PlateSlide,
): number {
    const columns = node.children.filter(child => child.type === 'column_item');
    if (columns.length === 0) return y;

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
            currentYinCol = processNode({
                node: childNode,
                slide,
                theme,
                area: colArea,
                y: currentYinCol,
                slideData,
            });
        });
        maxColY = Math.max(maxColY, currentYinCol);
    });

    return maxColY + 0.2;
}

function addBulletsElement(
    node: PlateNode,
    slide: PptxGenJS.Slide,
    theme: Theme,
    area: ContentArea,
    y: number,
    slideData: PlateSlide,
): number {
    let currentY = y;
    (node.children as PlateNode[]).forEach(childNode => {
        currentY = processNode({
            node: childNode,
            slide,
            theme,
            area,
            y: currentY,
            slideData,
        });
    });
    return currentY;
}


function processNode({ node, slide, theme, area, y, slideData }: ProcessNodeParams & { slideData: PlateSlide }): number {
  switch (node.type) {
    case 'h1':
    case 'h2':
    case 'h3':
    case 'h4':
    case 'h5':
    case 'h6':
    case 'p':
        return addTextElement(node, slide, theme, area, y, slideData);
    case 'img':
      return addImageElement(node, slide, area, y);
    case 'column':
      return addColumnsElement(node, slide, theme, area, y, slideData);
    case 'bullets':
        return addBulletsElement(node, slide, theme, area, y, slideData);
    case 'bullet':
        // Individual bullets are handled inside `addBulletsElement`'s processNode call.
        return addTextElement(node, slide, theme, area, y, slideData);
    case 'chart': // Placeholder for complex elements
        slide.addText(`[Chart: ${node.chartType}]`, { x: area.x, y, w: area.w, h: 0.5 });
        return y + 0.6;
    case 'visualization-list': // Placeholder for complex elements
        slide.addText(`[Visualization: ${node.visualizationType}]`, { x: area.x, y, w: area.w, h: 0.5 });
        return y + 0.6;
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
  // 16:9 aspect ratio is 10" x 5.625" in pptxgenjs
  let contentArea: ContentArea = { x: 0.5, y: 0.5, w: 9.0, h: 4.625 };

  if (slideData.rootImage?.url) {
    const imagePath = slideData.rootImage.url;
    try {
        switch (slideData.layoutType) {
            case "left":
              slide.addImage({ path: imagePath, x: 0, y: 0, w: "50%", h: "100%" });
              contentArea = { x: 5.25, y: 0.5, w: 4.5, h: 4.625 };
              break;
            case "right":
              slide.addImage({ path: imagePath, x: "50%", y: 0, w: "50%", h: "100%" });
              contentArea = { x: 0.25, y: 0.5, w: 4.5, h: 4.625 };
              break;
            case "vertical":
              slide.addImage({ path: imagePath, x: 0, y: 0, w: "100%", h: "40%" });
              contentArea = { x: 0.5, y: 2.35, w: 9.0, h: 2.775 };
              break;
            default:
              slide.addImage({ path: imagePath, x: 0, y: 0, w: "100%", h: "100%" });
              slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: '000000', transparency: 40 } });
              break;
        }
    } catch(e) {
        console.error("Error adding root image. It might be an invalid URL or path.", e);
        slide.addText(`[Image not found: ${imagePath}]`, {x: 0, y:0, w: '50%', h: '100%'});
    }
  }

  let currentY = contentArea.y as number;

  for (const node of slideData.content) {
    currentY = processNode({
        node,
        slide,
        theme,
        area: contentArea,
        y: currentY,
        slideData
    });
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

    const theme = resolveTheme(presentation.presentation.theme) ?? {
      name: "daktilo",
      properties: themes.daktilo,
    };
    const slides = (
      presentation.presentation.content as { slides: PlateSlide[] }
    ).slides;

    for (const slideData of slides) {
      const slide = pptx.addSlide();
      const themeColors = theme.properties.colors.light;
      
      let finalBgColor = themeColors.background;
      if (slideData.bgColor) {
        finalBgColor = slideData.bgColor;
      }
      slide.background = { color: finalBgColor.replace("#", "") };

      addSlideContent(pptx, slide, slideData, theme);
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