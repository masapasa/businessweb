import type PptxGenJS from 'pptxgenjs';
import path from 'path';

const LOGO_PATH = path.join(process.cwd(), "public", "nubri.png");
const BORDER_COLOR = "#0000ff";
const BORDER_IN_COLOR = "#0000ff";
const rechtangelBorder_Color = "#b60c26";
const SLIDE_BACKGROUND = "#141414";
const TEXT_COLOR = "#FFFFFF";
const ArtikelText = "xxx";
const BottomFotterText = "Anqa IT Security GmbH";
const AnqaAddress = "anqa-itsecurity.de";

export class SlideBuilder {
  private pptx: PptxGenJS;

  constructor(pptx: PptxGenJS) {
    this.pptx = pptx;
  }

  createSlide(): PptxGenJS.Slide {
    const slide = this.pptx.addSlide();
    slide.background = { fill: SLIDE_BACKGROUND };
    return slide;
  }

  addText(slide: PptxGenJS.Slide, text: string, options: PptxGenJS.TextPropsOptions) {
    slide.addText(text, { ...options, color: TEXT_COLOR });
  }

  addLogo(slide: PptxGenJS.Slide, x = 9.0, y = 0.1, w = 0.8, h = 0.5) {
    slide.addImage({ path: LOGO_PATH, x, y, w, h });
  }

  addRoundBorder(slide: PptxGenJS.Slide, borderColor = BORDER_COLOR) {
    const lineOpts: PptxGenJS.ShapeProps = { line: { color: borderColor, width: 2 } };
    slide.addShape(this.pptx.ShapeType.line, { x: 0, y: 0, w: '100%', h: 0, ...lineOpts });
    slide.addShape(this.pptx.ShapeType.line, { x: 0, y: '100%', w: '100%', h: 0, ...lineOpts });
    slide.addShape(this.pptx.ShapeType.line, { x: 0, y: 0, w: 0, h: '100%', ...lineOpts });
    slide.addShape(this.pptx.ShapeType.line, { x: '100%', y: 0, w: 0, h: '100%', ...lineOpts });
  }

  addInBorder(slide: PptxGenJS.Slide) {
    const lineOpts: PptxGenJS.ShapeProps = { line: { color: BORDER_IN_COLOR, width: 0.5 } };
    slide.addShape(this.pptx.ShapeType.line, { x: 1, y: 5.3, w: 8.0, h: 0, ...lineOpts });
    slide.addShape(this.pptx.ShapeType.line, { x: 1, y: 0.7, w: 8.0, h: 0, ...lineOpts });
  }

  thirteenthSlideTabell(slide: PptxGenJS.Slide) {
    const data = [
      { abteilung: "IT", downloads: "150", klicks: "75" },
      { abteilung: "Marketing", downloads: "300", klicks: "150" },
      { abteilung: "Sales", downloads: "200", klicks: "100" },
    ];
    const headerOpts: PptxGenJS.TableCellProps = { color: "FFFFFF", bold: true, fill: { color: "0000FF" }, align: "center", valign: "middle", border: { pt: 1, color: "000000" } };
    const cellOpts: PptxGenJS.TableCellProps = { color: "000000", fill: {color: "CCCCCC"}, align: "center", valign: "middle", border: { pt: 1, color: "000000" } };

    const rows: PptxGenJS.TableRow[] = [
      [{ text: "Abteilung", options: headerOpts }, { text: "Downloads", options: headerOpts }, { text: "Klicks", options: headerOpts }],
    ];
    
    let totalDownloads = 0;
    let totalKlicks = 0;
    
    data.forEach(record => {
      rows.push([{ text: record.abteilung, options: cellOpts }, { text: record.downloads, options: cellOpts }, { text: record.klicks, options: cellOpts }]);
      totalDownloads += parseInt(record.downloads, 10);
      totalKlicks += parseInt(record.klicks, 10);
    });

    rows.push([{ text: "Total", options: headerOpts }, { text: totalDownloads.toString(), options: cellOpts }, { text: totalKlicks.toString(), options: cellOpts }]);

    slide.addTable(rows, { x: 1.0, y: 1.8, w: 4.5, colW: [2.2, 1.8, 1.8], rowH: 0.1, fontFace: "Arial", fontSize: 11 });
  }
  
  seventhSlideRechtangelBorder(slide: PptxGenJS.Slide) {
    const lineOpts: PptxGenJS.ShapeProps = { line: { color: rechtangelBorder_Color, width: 0.5 } };
    slide.addShape(this.pptx.ShapeType.line, { x: 1.5, y: 2.7, w: 6.7, h: 0, ...lineOpts });
    slide.addShape(this.pptx.ShapeType.line, { x: 1.5, y: 3.4, w: 6.7, h: 0, ...lineOpts });
    slide.addShape(this.pptx.ShapeType.line, { x: 1.5, y: 2.7, w: 0, h: 0.7, ...lineOpts });
    slide.addShape(this.pptx.ShapeType.line, { x: 8.2, y: 2.7, w: 0, h: 0.7, ...lineOpts });
  }

  sixtenthSlideRechtangelBorder(slide: PptxGenJS.Slide) {
    const lineOpts: PptxGenJS.ShapeProps = { line: { color: rechtangelBorder_Color, width: 0.5 } };
    slide.addShape(this.pptx.ShapeType.line, { x: 1.5, y: 3.4, w: 6.9, h: 0, ...lineOpts });
    slide.addShape(this.pptx.ShapeType.line, { x: 1.5, y: 3.0, w: 6.9, h: 0, ...lineOpts });
    slide.addShape(this.pptx.ShapeType.line, { x: 1.5, y: 3.0, w: 0, h: 0.4, ...lineOpts });
    slide.addShape(this.pptx.ShapeType.line, { x: 8.4, y: 3.0, w: 0, h: 0.4, ...lineOpts });
    slide.addShape(this.pptx.ShapeType.line, { x: 1.5, y: 4.3, w: 7.4, h: 0, ...lineOpts });
    slide.addShape(this.pptx.ShapeType.line, { x: 1.5, y: 5.0, w: 7.4, h: 0, ...lineOpts });
    slide.addShape(this.pptx.ShapeType.line, { x: 1.5, y: 4.3, w: 0, h: 0.7, ...lineOpts });
    slide.addShape(this.pptx.ShapeType.line, { x: 8.9, y: 4.3, w: 0, h: 0.7, ...lineOpts });
  }

  addWebsiteAddress(slide: PptxGenJS.Slide) {
    this.addText(slide, AnqaAddress, { x: 0.1, y: 5.5, fontSize: 8 });
  }
  
  addStandardText(slide: PptxGenJS.Slide) {
    this.addText(slide, ArtikelText, { x: 3.0, y: 0.5, fontSize: 18, bold: true });
    this.addText(slide, BottomFotterText, { x: 1.7, y: 5.4, fontSize: 8 });
  }
}