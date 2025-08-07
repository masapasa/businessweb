import { NextResponse } from "next/server";
import PptxGenJS from "pptxgenjs";
import { db } from "@/server/db";
import {
  themes,
  type ThemeName,
  type ThemeProperties,
} from "@/lib/presentation/themes";
import { type PlateSlide } from "@/components/presentation/utils/parser";

// Define the structure for a slide element
interface SlideElement {
  type:
    | "title"
    | "subtitle"
    | "text"
    | "bullet"
    | "image"
    | "headline"
    | "description";
  content: string;
  level?: number;
  options?: PptxGenJS.TextPropsOptions;
  imagePath?: string; // For images
}

// Define the structure for theme properties
interface Theme {
  name: string;
  properties: ThemeProperties;
}

// Function to resolve the theme from the request
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

// Main function to handle the POST request
export async function POST(req: Request) {
  try {
    const { id } = await req.json();

    if (!id) {
      return NextResponse.json(
        { error: "Invalid or missing presentation ID" },
        { status: 400 }
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
        { status: 404 }
      );
    }

    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_16x9";

    const theme = resolveTheme(presentation.presentation.theme) ?? {
      name: "daktilo",
      properties: themes.daktilo,
    };

    const slides = (presentation.presentation.content as { slides: PlateSlide[] })
      .slides;

    for (const slideData of slides) {
      const slide = pptx.addSlide();
      slide.background = { color: theme.properties.colors.light.background };
      addSlideContent(slide, slideData, theme);
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
      { status: 500 }
    );
  }
}

function getStyleForElement(
  type: SlideElement["type"],
  theme: Theme,
): PptxGenJS.TextPropsOptions {
  const colors = theme.properties.colors.light;
  const fonts = theme.properties.fonts;

  switch (type) {
    case "title":
      return {
        color: colors.heading,
        fontFace: fonts.heading,
        fontSize: 32,
        bold: true,
      };
    case "subtitle":
      return { color: colors.muted, fontFace: fonts.body, fontSize: 18 };
    case "headline":
      return {
        color: colors.heading,
        fontFace: fonts.heading,
        fontSize: 24,
        bold: true,
      };
    case "description":
      return { color: colors.text, fontFace: fonts.body, fontSize: 14 };
    case "bullet":
      return {
        color: colors.text,
        fontFace: fonts.body,
        fontSize: 14,
        bullet: true,
      };
    default:
      return { color: colors.text, fontFace: fonts.body, fontSize: 14 };
  }
}

function addSlideContent(
  slide: PptxGenJS.Slide,
  slideData: PlateSlide,
  theme: Theme,
) {
  const elements = parseElements(slideData.children);

  let y = 1.0; // Starting Y position

  elements.forEach((element) => {
    let options: PptxGenJS.TextPropsOptions = {
      x: 0.5,
      y,
      w: "90%",
      h: 0.5,
      ...getStyleForElement(element.type, theme),
      ...element.options,
    };

    switch (element.type) {
      case "title":
        options.y = 0.5;
        options.h = 1;
        break;
      case "subtitle":
        options.y = 1.5;
        options.h = 0.75;
        break;
      case "image":
        if (element.imagePath) {
          slide.addImage({
            path: element.imagePath,
            x: "25%",
            y: "25%",
            w: "50%",
            h: "50%",
          });
        }
        return; // Skip text rendering for images
    }

    slide.addText(element.content, options);
    y += 0.6; // Increment Y for the next element
  });
}

// Parse raw content into structured elements
function parseElements(content: any[]): SlideElement[] {
  const elements: SlideElement[] = [];
  content.forEach((block) => {
    const text = block.children.map((c: any) => c.text).join("");
    switch (block.type) {
      case "h1":
        elements.push({ type: "title", content: text });
        break;
      case "h2":
        elements.push({ type: "subtitle", content: text });
        break;
      case "h3":
        elements.push({ type: "headline", content: text });
        break;
      case "p":
        elements.push({ type: "description", content: text });
        break;
      case "bullet":
        elements.push({ type: "bullet", content: text, level: 1 });
        break;
      case "img":
        elements.push({ type: "image", content: "", imagePath: block.url });
        break;
    }
  });
  return elements;
}
