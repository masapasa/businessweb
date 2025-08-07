import { NextResponse } from 'next/server';
import PptxGenJS from 'pptxgenjs';
import { SlideBuilder } from '@/lib/presentation/SlideBuilder';
import { JSDOM } from 'jsdom';

export const runtime = 'nodejs';

export async function POST(req: Request) {
  try {
    const { title, items } = await req.json();
    if (!title || !items || !Array.isArray(items)) {
      return NextResponse.json({ error: 'Missing title or items' }, { status: 400 });
    }
    const pptx = new PptxGenJS();
    const slideBuilder = new SlideBuilder(pptx);
    pptx.layout = "LAYOUT_16x9";
    pptx.author = "Superslide";
    pptx.company = "Anqa IT Security GmbH";
    pptx.subject = title;
    pptx.title = title;
    items.forEach((item: { content: string }, index: number) => {
        const slide = slideBuilder.createSlide();
        slideBuilder.addLogo(slide);
        slideBuilder.addRoundBorder(slide);
        slideBuilder.addInBorder(slide);
        slideBuilder.addWebsiteAddress(slide);
        slideBuilder.addStandardText(slide);
        const dom = new JSDOM(`<!DOCTYPE html><body>${item.content}</body>`);
        const body = dom.window.document.body;
        const h1 = body.querySelector('H1');
        const p = body.querySelector('P');
        if(h1) {
            slideBuilder.addText(slide, h1.textContent || '', { x: 1.0, y: 1.0, w: 8.0, h: 0.5, fontSize: 36, bold: true });
        }
        if(p) {
            slideBuilder.addText(slide, p.textContent || '', { x: 1.0, y: 1.8, w: 8.0, h: 2.5, fontSize: 18 });
        }
        if (index + 1 === 7) {
            slideBuilder.seventhSlideRechtangelBorder(slide);
        }
        if (index + 1 === 13) {
            slideBuilder.thirteenthSlideTabell(slide);
        }
        if (index + 1 === 16) {
            slideBuilder.sixtenthSlideRechtangelBorder(slide);
        }
    });
    const pptxBuffer = await pptx.write({ outputType: 'nodebuffer' });
    return new NextResponse(pptxBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        "Content-Disposition": `attachment; filename="${title.replace(/ /g, '_')}.pptx"`,
      }
    });
  } catch (e) {
    console.error("PPT Download Error:", e);
    const errorMessage = e instanceof Error ? e.message : "An unknown error occurred";
    return NextResponse.json({ error: 'Internal error', details: errorMessage }, { status: 500 });
  }
}