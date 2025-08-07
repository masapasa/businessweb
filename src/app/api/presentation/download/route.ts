import { NextResponse } from "next/server";
import PptxGenJS from "pptxgenjs";
import {
  themes,
  type ThemeName,
  type ThemeProperties,
} from "@/lib/presentation/themes";

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

// Define the structure for a single slide
interface SlideData {
  id: string;
  content: ContentBlock[];
  layoutType?: "left" | "right" | "center" | "top" | "bottom" | "custom";
  alignment?: "start" | "center" | "end";
  rootImage?: {
    url: string;
    query?: string;
  };
  layout?: {
    name: string;
    elements: Partial<Record<SlideElement["type"], PptxGenJS.TextPropsOptions>>;
  };
  elements?: SlideElement[]; // Use a more structured elements format
  theme?: Theme;
}

// Define the structure for a content block within a slide
interface ContentBlock {
  type: string;
  children: TextNode[];
}

// Define the structure for a text node within a content block
interface TextNode {
  text: string;
  bold?: boolean;
  italic?: boolean;
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
      return { name: themeName, properties: themes[themeName].properties };
    }
  }
  return null;
}

// Main function to handle the POST request
export async function POST(req: Request) {
  try {
    const { sections, theme: themeData } = await req.json();

    if (!sections || !Array.isArray(sections) || sections.length === 0) {
      return NextResponse.json(
        { error: "Invalid or missing sections data" },
        { status: 400 }
      );
    }

    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_16x9";
    const defaultTheme = {
      name: "Default",
      properties: themes.Default.properties,
    };

    for (const section of sections) {
      const theme = resolveTheme(themeData) ?? defaultTheme;
      const slide = pptx.addSlide();
      slide.background = { color: theme.properties.backgroundColor };
      addSlideContent(slide, section as SlideData, theme);
    }

    const pptxBuffer = await pptx.write({ outputType: "nodebuffer" });
    const fileName = "presentation.pptx";

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

// More advanced layout engine
function addSlideContent(
  slide: PptxGenJS.Slide,
  section: SlideData,
  theme: Theme
) {
  const elements = section.elements ?? parseElements(section.content);

  // If a custom layout is defined, use it
  if (section.layoutType === "custom" && section.layout) {
    addCustomLayout(slide, elements, theme, section.layout);
    return;
  }

  // Fallback to default layouts
  addDefaultLayout(slide, elements, theme, section);
}

// Function to handle custom layouts
function addCustomLayout(
  slide: PptxGenJS.Slide,
  elements: SlideElement[],
  theme: Theme,
  layout: SlideData["layout"]
) {
  elements.forEach((element) => {
    const baseStyle =
      (theme.properties as any)[element.type] ?? theme.properties.text;
    const layoutStyle = layout?.elements[element.type] ?? {};

    const finalOptions: PptxGenJS.TextPropsOptions = {
      ...baseStyle,
      ...layoutStyle,
      ...element.options, // Individual element styles override everything
    };

    if (element.type === "image" && element.imagePath) {
      slide.addImage({
        path: element.imagePath,
        ...(finalOptions as PptxGenJS.ImageProps),
      });
    } else {
      slide.addText(element.content, finalOptions);
    }
  });
}

function addDefaultLayout(
  slide: PptxGenJS.Slide,
  elements: SlideElement[],
  theme: Theme,
  section: SlideData
) {
  let y = 1.0; // Starting Y position

  elements.forEach((element) => {
    let options: PptxGenJS.TextPropsOptions = {
      x: 0.5,
      y,
      w: "90%",
      h: 0.5,
      color: theme.properties.text.color, // Default text color
    };

    switch (element.type) {
      case "title":
        options = { ...theme.properties.title };
        break;
      case "subtitle":
        options = { ...theme.properties.subtitle };
        break;
      case "headline":
        options = {
          ...theme.properties.text,
          fontSize: 24,
          bold: true,
          ...element.options,
        };
        break;
      case "description":
        options = { ...theme.properties.text, ...element.options };
        break;
      case "bullet":
        options = {
          ...theme.properties.bullet,
          ...element.options,
          bullet: true,
        };
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

  if (section.rootImage?.url) {
    slide.addImage({
      path: section.rootImage.url,
      x: "70%",
      y: "20%",
      w: "25%",
      h: "60%",
    });
  }
}

// Parse raw content into structured elements
function parseElements(content: ContentBlock[]): SlideElement[] {
  const elements: SlideElement[] = [];
  content.forEach((block) => {
    const text = block.children.map((c) => c.text).join("");
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
    }
  });
  return elements;
}