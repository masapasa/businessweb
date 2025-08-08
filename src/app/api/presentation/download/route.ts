import { NextResponse } from 'next/server';
import PptxGenJS from 'pptxgenjs';

export const runtime = 'nodejs';

// --- Interfaces for incoming JSON data ---
interface RootImage {
  url: string;
  query?: string;
  background?: boolean;
  alt?: string;
}

type PlateNode = {
  text?: string;
  type?: string;
  children?: PlateNode[];
};

interface ContentBlock {
  type: string;
  children: PlateNode[];
}

type LayoutType = 'left' | 'right' | 'vertical';

interface SectionData {
  id: string;
  content: ContentBlock[];
  rootImage?: RootImage;
  layoutType?: LayoutType;
}

type ThemeColors = {
  primary: string;
  secondary: string;
  accent: string;
  background: string;
  text: string;
  heading: string;
  muted: string;
};

type ThemeFonts = {
  heading: string;
  body: string;
};

interface PPTInput {
  sections: SectionData[];
  theme?: {
    // We only need the resolved palette that should be used when exporting
    colors: ThemeColors;
    fonts: ThemeFonts;
  };
}

// --- Default Theme (used if none is provided by the client) ---
const DEFAULT_THEME = {
  backgroundColor: '1F1F1F', // Dark gray background
  textColor: 'FFFFFF',
  headingColor: 'FAFAFA',
  fontHeading: 'Arial',
  fontBody: 'Arial',
};

// --- Helper Functions ---

async function imageUrlToBase64(url: string): Promise<string | null> {
    if (!url?.startsWith('http')) {
        console.error(`Invalid or non-HTTP image URL: ${url}`);
        return null;
    }
    try {
        const response = await fetch(url);
        if (!response.ok) return null;
        const buffer = await response.arrayBuffer();
        const mimeType = response.headers.get('content-type') ?? 'image/jpeg';
        return `data:${mimeType};base64,${Buffer.from(buffer).toString('base64')}`;
    } catch (error) {
        console.error(`Failed to convert image to base64: ${String(error)}`);
        return null;
    }
}

function parseContentBlocks(contentBlocks: ContentBlock[]): { title: string; body: string } {
    let title = '';
    const bodyParts: string[] = [];
    const processNode = (node: PlateNode): string => {
        if (node.text) return node.text;
        if (node.children) return node.children.map(processNode).join('');
        return '';
    };

    for (const block of contentBlocks) {
        if (block.type === 'h1' && !title) {
            title = processNode(block);
        } else if (block.type === 'p') {
            bodyParts.push(processNode(block));
        } else if (block.type === 'bullets') {
            const items = block.children
              .map((bullet: PlateNode) => {
                const h3 = bullet.children?.find((c: PlateNode) => c.type === 'h3');
                const p = bullet.children?.find((c: PlateNode) => c.type === 'p');
                const h3Text = h3 ? processNode(h3) : '';
                const pText = p ? processNode(p) : '';
                const combined = [h3Text, pText].filter(Boolean).join('\n  ');
                return combined ? `• ${combined}` : '';
              })
              .filter((s: string) => s && s.trim().length > 1);
            if (items.length) bodyParts.push(items.join('\n\n'));
        }
    }
    return { title, body: bodyParts.join('\n\n') };
}

// --- Main API Route Handler ---

export async function POST(req: Request) {
  try {
    const pptData: PPTInput = await req.json();
    // Debug logs for incoming request
    // eslint-disable-next-line no-console
    console.log("[download] sections count:", pptData.sections?.length ?? 0);
    // eslint-disable-next-line no-console
    console.log("[download] first section:", pptData.sections?.[0]);

    if (!pptData.sections?.length) {
      return NextResponse.json({ error: 'Invalid or missing sections data' }, { status: 400 });
    }

    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_16x9";
    pptx.author = "Presentation AI";
    // Resolve theme for export
    const resolvedTheme = pptData.theme
      ? (() => {
          const bg = pptData.theme.colors.background ?? DEFAULT_THEME.backgroundColor;
          const text = pptData.theme.colors.text ?? DEFAULT_THEME.textColor;
          const heading = pptData.theme.colors.heading ?? DEFAULT_THEME.headingColor;
          const fontHeading = pptData.theme.fonts.heading ?? DEFAULT_THEME.fontHeading;
          const fontBody = pptData.theme.fonts.body ?? DEFAULT_THEME.fontBody;
          return {
            backgroundColor: bg.replace('#', ''),
            textColor: text.replace('#', ''),
            headingColor: heading.replace('#', ''),
            fontHeading,
            fontBody,
          };
        })()
      : DEFAULT_THEME;

    // Derive document title from first section h1
    const firstSection = pptData.sections[0];
    const { title: docTitle } = parseContentBlocks(firstSection?.content ?? []);
    pptx.title = docTitle || 'Generated Presentation';

    // --- Create Content Slides ---
    for (const section of pptData.sections) {
      const slide = pptx.addSlide();
      slide.background = { color: resolvedTheme.backgroundColor };

      const { title, body } = parseContentBlocks(section.content);

      const hasImage = Boolean(section.rootImage?.url);
      const layout: LayoutType = section.layoutType ?? 'vertical';

      // Preload image if present
      let base64Image: string | null = null;
      if (hasImage && section.rootImage?.url) {
        base64Image = await imageUrlToBase64(section.rootImage.url);
      }

      // Layout placements
      if (hasImage && base64Image) {
        if (layout === 'vertical') {
          slide.addImage({ data: base64Image, x: 0, y: 0, w: '100%', h: '45%' });
          slide.addText(title, {
            fontFace: resolvedTheme.fontHeading,
            color: resolvedTheme.headingColor,
            fontSize: 28,
            bold: true,
            x: '5%', y: '48%', w: '90%', h: '12%',
          });
          slide.addText(body, {
            fontFace: resolvedTheme.fontBody,
            color: resolvedTheme.textColor,
            fontSize: 16,
            x: '5%', y: '60%', w: '90%', h: '35%',
            lineSpacing: 22,
          });
        } else if (layout === 'left') {
          // Image on left, text on right
          slide.addImage({ data: base64Image, x: '0%', y: '0%', w: '45%', h: '100%' });
          slide.addText(title, {
            fontFace: resolvedTheme.fontHeading,
            color: resolvedTheme.headingColor,
            fontSize: 28,
            bold: true,
            x: '50%', y: '5%', w: '47%', h: '15%',
          });
          slide.addText(body, {
            fontFace: resolvedTheme.fontBody,
            color: resolvedTheme.textColor,
            fontSize: 16,
            x: '50%', y: '20%', w: '47%', h: '70%',
            lineSpacing: 22,
          });
        } else {
          // layout === 'right': image on right, text on left
          slide.addImage({ data: base64Image, x: '55%', y: '0%', w: '45%', h: '100%' });
          slide.addText(title, {
            fontFace: resolvedTheme.fontHeading,
            color: resolvedTheme.headingColor,
            fontSize: 28,
            bold: true,
            x: '5%', y: '5%', w: '47%', h: '15%',
          });
          slide.addText(body, {
            fontFace: resolvedTheme.fontBody,
            color: resolvedTheme.textColor,
            fontSize: 16,
            x: '5%', y: '20%', w: '47%', h: '70%',
            lineSpacing: 22,
          });
        }
      } else {
        // No image: simple content
        slide.addText(title, {
          fontFace: resolvedTheme.fontHeading,
          color: resolvedTheme.headingColor,
          fontSize: 32,
          bold: true,
          x: '5%', y: '8%', w: '90%', h: '12%',
        });
        slide.addText(body, {
          fontFace: resolvedTheme.fontBody,
          color: resolvedTheme.textColor,
          fontSize: 18,
          x: '5%', y: '22%', w: '90%', h: '70%',
          lineSpacing: 24,
        });
      }
    }
    
    // --- 3. Finalize and Send Response ---
    const pptxBuffer = await pptx.write({ outputType: 'nodebuffer' });
    const fileName = `${pptx.title.replace(/[^a-z0-9]/gi, '_').toLowerCase()}.pptx`;

    return new NextResponse(pptxBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        "Content-Disposition": `attachment; filename="${fileName}"`,
      }
    });

  } catch (e) {
    console.error("PPT Download Error:", e);
    const errorMessage = e instanceof Error ? e.message : "An unknown error occurred";
    return NextResponse.json({ error: 'Internal server error', details: errorMessage }, { status: 500 });
  }
}