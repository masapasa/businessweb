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

interface ContentBlock {
  type: string;
  children: any[];
}

interface SectionData {
  id: string;
  content: ContentBlock[];
  rootImage?: RootImage;
  layoutType?: string;
  alignment?: string;
}

interface PPTInput {
  sections: SectionData[];
  references?: string[];
}

// --- Centralized Theme and Styles for Consistency ---
const THEME = {
  backgroundColor: '1F1F1F', // Dark background similar to the screenshot
  textColor: 'FFFFFF', // White text
  fontHeading: 'Arial',
  fontBody: 'Arial',
};

const FOOTER_STYLE: PptxGenJS.TextPropsOptions = {
  fontFace: THEME.fontBody,
  fontSize: 10,
  color: 'A9A9A9', // Light gray for footer
  align: 'center',
  y: '95%',
  w: '100%',
};

// --- Helper Functions ---

/**
 * Fetches an image from a URL and converts it to a base64 data string.
 * This is necessary for embedding images directly into the presentation.
 */
async function imageUrlToBase64(url: string): Promise<string | null> {
    if (!url || (!url.startsWith('http:') && !url.startsWith('https:'))) {
        console.error(`Invalid image URL: ${url}`);
        return null;
    }
    try {
        const response = await fetch(url);
        if (!response.ok) {
            console.error(`Failed to fetch image: ${response.status} ${response.statusText} for url: ${url}`);
            return null;
        }
        const buffer = await response.arrayBuffer();
        const base64 = Buffer.from(buffer).toString('base64');
        const mimeType = response.headers.get('content-type') || 'image/jpeg';
        return `data:${mimeType};base64,${base64}`;
    } catch (error) {
        console.error(`Error converting image URL to base64: ${error}`);
        return null;
    }
}

/**
 * Parses the complex content block structure from the JSON to extract a simple title and body string.
 */
function parseContentBlocks(contentBlocks: ContentBlock[]): { title: string; body: string } {
    let title = '';
    const bodyParts: string[] = [];

    const processNode = (node: any): string => {
        if (node.text) return node.text;
        if (node.children && Array.isArray(node.children)) {
            return node.children.map(processNode).join('');
        }
        return '';
    };

    contentBlocks.forEach(block => {
        if (block.type === 'h1' && !title) {
            title = processNode(block);
        } else if (block.type === 'p') {
            bodyParts.push(processNode(block));
        } else if (block.type === 'bullets') {
            const bulletItems = block.children.map((bullet: any) => {
                const h3 = bullet.children.find((c: any) => c.type === 'h3');
                const p = bullet.children.find((c: any) => c.type === 'p');
                const h3Text = h3 ? processNode(h3) : '';
                const pText = p ? processNode(p) : '';
                return `• ${h3Text}\n  ${pText}`;
            });
            bodyParts.push(bulletItems.join('\n\n'));
        }
    });

    return { title, body: bodyParts.join('\n\n') };
}


// --- Main API Route Handler ---

export async function POST(req: Request) {
  try {
    const pptData: PPTInput = await req.json();

    if (!pptData.sections || !Array.isArray(pptData.sections) || pptData.sections.length === 0) {
      return NextResponse.json({ error: 'Invalid or missing sections data' }, { status: 400 });
    }

    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_16x9";
    pptx.author = "Presentation AI";

    const { title: docTitle } = parseContentBlocks(pptData.sections[0]?.content || []);
    pptx.title = docTitle || "Generated Presentation";

    // --- Slide Generation Loop ---
    for (const [index, section] of pptData.sections.entries()) {
        const slide = pptx.addSlide();
        const { title, body } = parseContentBlocks(section.content);
        const hasImage = section.rootImage?.url && !section.rootImage.background;

        if (index === 0) { // Title Slide (full dark background)
            slide.background = { color: THEME.backgroundColor };
            slide.addText(title, {
                fontFace: THEME.fontHeading,
                color: THEME.textColor,
                fontSize: 44,
                bold: true,
                align: 'center',
                valign: 'middle',
                x: '5%',
                y: '0%',
                w: '90%',
                h: '100%',
            });
        } else if (hasImage) { // Content Slide with Image Banner
            // 1. Add the image banner at the top
            const base64Image = await imageUrlToBase64(section.rootImage!.url);
            if (base64Image) {
                slide.addImage({ data: base64Image, x: 0, y: 0, w: '100%', h: '50%' });
            }

            // 2. Add the dark content box at the bottom
            slide.addShape(pptx.ShapeType.rect, {
                x: 0,
                y: '50%',
                w: '100%',
                h: '50%',
                fill: { color: THEME.backgroundColor },
            });

            // 3. Add text ON TOP of the dark box
            slide.addText(title, {
                fontFace: THEME.fontHeading,
                color: THEME.textColor,
                fontSize: 32,
                bold: true,
                x: '5%',
                y: '55%',
                w: '90%',
                h: '15%',
            });
            slide.addText(body, {
                fontFace: THEME.fontBody,
                color: THEME.textColor,
                fontSize: 14,
                x: '5%',
                y: '68%',
                w: '90%',
                h: '27%',
                lineSpacing: 22,
            });
        } else { // Content-only slide (full dark background)
            slide.background = { color: THEME.backgroundColor };
            slide.addText(title, {
                fontFace: THEME.fontHeading,
                color: THEME.textColor,
                fontSize: 36,
                bold: true,
                x: '5%',
                y: '5%',
                w: '90%',
                h: '15%',
            });
            slide.addText(body, {
                fontFace: THEME.fontBody,
                color: THEME.textColor,
                fontSize: 16,
                x: '5%',
                y: '25%',
                w: '90%',
                h: '70%',
                lineSpacing: 24,
            });
        }
        slide.addText(`${index + 1}`, FOOTER_STYLE);
    }
    
    // --- Finalization: Buffer and Response ---
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