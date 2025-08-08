import { NextResponse } from 'next/server';
import PptxGenJS from 'pptxgenjs';

export const runtime = 'nodejs';

// --- Interfaces for incoming JSON data structure ---
interface RootImage {
  url: string;
  background?: boolean;
}

interface ContentBlock {
  type: string;
  children: Array<{ text?: string; type?: string; children?: Array<{ text?: string; type?: string; children?: unknown[] }>; }>;
}

interface SectionData {
  id: string;
  content: ContentBlock[];
  rootImage?: RootImage;
  layoutType?: 'left' | 'right' | 'vertical';
}

type ExportTheme = {
  colors: {
    background: string;
    text: string;
    heading: string;
  };
  fonts: {
    heading: string;
    body: string;
  };
};

interface PPTInput {
  sections: SectionData[];
  theme?: ExportTheme;
}

// --- Centralized Theme fallbacks ---
const THEME = {
  backgroundColor: '1F1F1F', // Dark gray background
  textColor: 'FFFFFF',       // White text
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
        const response = await fetch(url, {
          // Some CDNs block default Node fetch; emulate a browser UA
          headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'image/avif,image/webp,image/apng,image/*,*/*;q=0.8',
          },
          redirect: 'follow',
        } as RequestInit);
        if (!response.ok) {
          console.error(`Image fetch failed: ${response.status} ${response.statusText}`);
          return null;
        }
        const buffer = await response.arrayBuffer();
        const mimeType = response.headers.get('content-type') ?? 'image/jpeg';
        return `data:${mimeType};base64,${Buffer.from(buffer).toString('base64')}`;
    } catch (error: unknown) {
        console.error(`Failed to convert image to base64: ${String(error)}`);
        return null;
    }
}

function parseContentBlocks(
  contentBlocks: ContentBlock[],
): { title: string; subtitle: string; body: string } {
  type Node = { text?: string; type?: string; children?: Node[] };

  let title = '';
  let subtitle = '';
  const bodyParts: string[] = [];
  const processNode = (node: Node): string => {
    if (node.text) return node.text;
    if (node.children) return node.children.map(processNode).join('');
    return '';
  };

  for (const block of (contentBlocks as unknown as Node[])) {
    if (block.type === 'h1' && !title) {
      title = processNode(block);
    } else if (block.type === 'h2' && !subtitle) {
      subtitle = processNode(block);
    } else if (block.type === 'p') {
      bodyParts.push(processNode(block));
    } else if (block.type === 'bullets' && Array.isArray(block.children)) {
      const items = (block.children)
        .map((bullet) => {
          const h3 = bullet.children?.find((c) => c.type === 'h3');
          const p = bullet.children?.find((c) => c.type === 'p');
          const h3Text = h3 ? processNode(h3) : '';
          const pText = p ? processNode(p) : '';
          const combined = [h3Text, pText].filter(Boolean).join('\n  ');
          return combined ? `• ${combined}` : '';
        })
        .filter((s) => s && s.trim().length > 1);
      if (items.length) bodyParts.push(items.join('\n\n'));
    }
  }
  return { title, subtitle, body: bodyParts.join('\n\n') };
}

// --- Main API Route Handler ---

export async function POST(req: Request) {
  try {
    const pptData: PPTInput = await req.json();
    if (!pptData.sections?.length) {
      return NextResponse.json({ error: 'Invalid or missing sections data' }, { status: 400 });
    }

    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_16x9";
    pptx.author = "Presentation AI";
    // --- Theme resolution ---
    const resolved = {
      bg: (pptData.theme?.colors.background ?? `#${THEME.backgroundColor}`).replace('#', ''),
      text: (pptData.theme?.colors.text ?? `#${THEME.textColor}`).replace('#', ''),
      heading: (pptData.theme?.colors.heading ?? `#${THEME.headingColor}`).replace('#', ''),
      fontHeading: pptData.theme?.fonts.heading ?? THEME.fontHeading,
      fontBody: pptData.theme?.fonts.body ?? THEME.fontBody,
    };

    // --- Generate Slides directly (no masters for simplicity/consistency) ---
    const { title: docTitle } = parseContentBlocks(pptData.sections[0]?.content ?? [] as unknown as ContentBlock[]);
    pptx.title = docTitle || 'Generated Presentation';

    const slideW = 10; // inches
    const slideH = 5.625; // inches

    for (const section of pptData.sections) {
      const { title, subtitle, body } = parseContentBlocks(section.content);
      const slide = pptx.addSlide();
      slide.background = { color: resolved.bg };

      const hasImage = Boolean(section.rootImage?.url) && !section.rootImage?.background;
      const layout = section.layoutType ?? 'vertical';

      if (hasImage) {
        const base64 = await imageUrlToBase64(section.rootImage!.url);
        if (base64) {
          if (layout === 'vertical') {
            slide.addImage({ data: base64, x: 0, y: 0, w: slideW, h: 3.1 });
          } else if (layout === 'left') {
            slide.addImage({ data: base64, x: 0, y: 0, w: 4.5, h: slideH });
          } else {
            slide.addImage({ data: base64, x: slideW - 4.5, y: 0, w: 4.5, h: slideH });
          }
        }
      }

      if (layout === 'vertical') {
        const titleY = hasImage ? 3.25 : 0.7;
        const subY = titleY + 0.6;
        const bodyY = titleY + 1.2;
        slide.addText(title, {
          fontFace: resolved.fontHeading,
          color: resolved.heading,
          fontSize: 32,
          bold: true,
          x: 0.6, y: titleY, w: slideW - 1.2, h: 0.8,
        });
        if (subtitle) {
          slide.addText(subtitle, {
            fontFace: resolved.fontHeading,
            color: resolved.text,
            fontSize: 20,
            x: 0.6, y: subY, w: slideW - 1.2, h: 0.6,
          });
        }
        if (body) {
          slide.addText(body, {
            fontFace: resolved.fontBody,
            color: resolved.text,
            fontSize: 18,
            x: 0.6, y: bodyY, w: slideW - 1.2, h: slideH - bodyY - 0.5,
          });
        }
      } else if (layout === 'left') {
        // Image left, content right
        slide.addText(title, {
          fontFace: resolved.fontHeading,
          color: resolved.heading,
          fontSize: 30,
          bold: true,
          x: 4.8, y: 0.6, w: 4.6, h: 0.8,
        });
        if (subtitle) {
          slide.addText(subtitle, {
            fontFace: resolved.fontHeading,
            color: resolved.text,
            fontSize: 18,
            x: 4.8, y: 1.4, w: 4.6, h: 0.6,
          });
        }
        if (body) {
          slide.addText(body, {
            fontFace: resolved.fontBody,
            color: resolved.text,
            fontSize: 16,
            x: 4.8, y: 2.1, w: 4.6, h: 3.2,
          });
        }
      } else {
        // Image right, content left
        slide.addText(title, {
          fontFace: resolved.fontHeading,
          color: resolved.heading,
          fontSize: 30,
          bold: true,
          x: 0.6, y: 0.6, w: 4.6, h: 0.8,
        });
        if (subtitle) {
          slide.addText(subtitle, {
            fontFace: resolved.fontHeading,
            color: resolved.text,
            fontSize: 18,
            x: 0.6, y: 1.4, w: 4.6, h: 0.6,
          });
        }
        if (body) {
          slide.addText(body, {
            fontFace: resolved.fontBody,
            color: resolved.text,
            fontSize: 16,
            x: 0.6, y: 2.1, w: 4.6, h: 3.2,
          });
        }
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