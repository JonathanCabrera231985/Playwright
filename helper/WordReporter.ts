import { Page } from '@playwright/test';
import { Document, Packer, Paragraph, TextRun, ImageRun } from 'docx';
import * as fs from 'fs';

export class WordReporter {
  private children: Paragraph[] = [];
  private stepCounter = 1;

  constructor() {
    this.children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: 'Reporte de Prueba Automatizada',
            bold: true,
            size: 32,
          }),
        ],
      })
    );
  }

  public async addStep(page: Page, description: string) {
    this.children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `Paso ${this.stepCounter}: ${description}`,
            bold: true,
          }),
        ],
      })
    );

    const screenshotBuffer = await page.screenshot();

    this.children.push(
      new Paragraph({
        children: [
          new ImageRun({
            data: screenshotBuffer,
            type: "png",
            transformation: {
              width: 600,
              height: 337.5,
            },
          }),
        ],
      })
    );

    this.stepCounter++;
  }

  /**
   * Inserta un salto de página en el documento.
   * El siguiente contenido que se agregue comenzará en una nueva página.
   */
  public addPageBreak() {
    this.children.push(
      new Paragraph({
        // Esta propiedad fuerza al párrafo a empezar en una nueva página.
        pageBreakBefore: true,
      })
    );
  }

  public async save(filePath: string) {
    const doc = new Document({
      sections: [{
        properties: {},
        children: this.children,
      }],
    });

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);
    console.log(`Reporte guardado en: ${filePath}`);
  }
}