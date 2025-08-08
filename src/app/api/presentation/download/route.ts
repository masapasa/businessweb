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
}

interface PPTInput {
  sections: SectionData[];
}

// --- Centralized Theme for Easy Customization ---
const THEME = {
  backgroundColor: '1F1F1F', // Dark gray background
  textColor: 'FFFFFF',       // White text
  fontHeading: 'Arial',
  fontBody: 'Arial',
};

// --- Helper Functions ---

async function imageUrlToBase64(url: string): Promise<string | null> {
    if (!url || !url.startsWith('http')) {
        console.error(`Invalid or non-HTTP image URL: ${url}`);
        return null;
    }
    try {
        const response = await fetch(url);
        if (!response.ok) return null;
        const buffer = await response.arrayBuffer();
        const mimeType = response.headers.get('content-type') || 'image/jpeg';
        return `data:${mimeType};base64,${Buffer.from(buffer).toString('base64')}`;
    } catch (error) {
        console.error(`Failed to convert image to base64: ${error}`);
        return null;
    }
}

function parseContentBlocks(contentBlocks: ContentBlock[]): { title: string; body: string } {
    let title = '';
    let bodyParts: string[] = [];
    const processNode = (node: any): string => {
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
            const items = block.children.map((bullet: any) => {
                const h3 = bullet.children.find((c: any) => c.type === 'h3');
                const p = bullet.children.find((c: any) => c.type === 'p');
                return `• ${h3 ? processNode(h3) : ''}\n  ${p ? processNode(p) : ''}`;
            });
            bodyParts.push(items.join('\n\n'));
        }
    }
    return { title, body: bodyParts.join('\n\n') };
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

    // --- 1. Create a Dedicated Title Slide ---
    const { title: docTitle, body: docSubtitle } = parseContentBlocks(pptData.sections[0].content);
    pptx.title = docTitle || "Generated Presentation";

    const titleSlide = pptx.addSlide();
    titleSlide.background = { color: THEME.backgroundColor };
    titleSlide.addText(docTitle, {
        fontFace: THEME.fontHeading,
        color: THEME.textColor,
        fontSize: 44,
        bold: true,
        align: 'center',
        x: '5%', y: '35%', w: '90%', h: '20%'
    });
    if (docSubtitle) {
        titleSlide.addText(docSubtitle, {
            fontFace: THEME.fontBody,
            color: THEME.textColor,
            fontSize: 18,
            align: 'center',
            x: '10%', y: '55%', w: '80%', h: '15%'
        });
    }

    // --- 2. Loop Through All Sections to Create Content Slides ---
    // We loop through all sections, including the first, to generate their respective content slides.
    for (const section of pptData.sections) {
        const slide = pptx.addSlide();
        const { title, body } = parseContentBlocks(section.content);
        const hasImage = section.rootImage?.url && !section.rootImage.background;

        // RENDER LOGIC: Based on whether an image exists for the slide.
        if (hasImage) {
            // Layout: Image Banner on Top, Content Box on Bottom
            const base64Image = await imageUrlToBase64(section.rootImage!.url);
            if (base64Image) {
                slide.addImage({ data: base64Image, x: 0, y: 0, w: '100%', h: '50%' });
            }

            slide.addShape(pptx.ShapeType.rect, {
                x: 0, y: '50%', w: '100%', h: '50%',
                fill: { color: THEME.backgroundColor },
            });

            slide.addText(title, {
                fontFace: THEME.fontHeading, color: THEME.textColor,
                fontSize: 28, bold: true,
                x: '5%', y: '55%', w: '90%', h: '15%',
            });
            slide.addText(body, {
                fontFace: THEME.fontBody, color: THEME.textColor,
                fontSize: 14,
                x: '5%', y: '68%', w: '90%', h: '27%',
                lineSpacing: 22,
            });
        } else {
            // Layout: Full-Screen Content (No Image)
            slide.background = { color: THEME.backgroundColor };
            slide.addText(title, {
                fontFace: THEME.fontHeading, color: THEME.textColor,
                fontSize: 32, bold: true,
                x: '5%', y: '5%', w: '90%', h: '15%',
            });
            slide.addText(body, {
                fontFace: THEME.fontBody, color: THEME.textColor,
                fontSize: 16,
                x: '5%', y: '25%', w: '90%', h: '70%',
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