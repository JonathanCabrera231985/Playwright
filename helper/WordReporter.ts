import { Page } from '@playwright/test';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  ShadingType,
  VerticalAlign,
  AlignmentType,
  Header,
  BorderStyle,
} from 'docx';
import * as fs from 'fs';

export class WordReporter {
  private children: (Paragraph | Table)[] = [];
  private header?: Header;
  private stepCounter = 1;

  constructor() {}

  public createHeader(imagePath: string) {
    const logoBuffer = fs.readFileSync(imagePath);
    this.header = new Header({
      children: [
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
          },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new ImageRun({
                          data: logoBuffer,
                          type: 'png',
                          transformation: { width: 150, height: 50 },
                        }),
                      ],
                    }),
                  ],
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                  children: [
                    new Paragraph({ text: 'Resultado de Pruebas Unitarias', alignment: AlignmentType.CENTER }),
                    new Paragraph({ text: 'Fábrica de software', alignment: AlignmentType.CENTER }),
                  ],
                  width: { size: 70, type: WidthType.PERCENTAGE },
                  verticalAlign: VerticalAlign.CENTER,
                }),
              ],
            }),
          ],
        }),
      ],
    });
  }

  public addTitle(text: string) {
    this.children.push(
      new Paragraph({
        children: [new TextRun({ text, bold: true, size: 32 })],
        spacing: { after: 200 },
      })
    );
  }

  public addSubtitle(text: string) {
    this.children.push(
      new Paragraph({
        children: [new TextRun({ text, bold: true, size: 28 })],
        spacing: { after: 150 },
      })
    );
  }

 public addFeatureTable(data: {
    functionality: string;
    scenario: string;
    type: string;
    preconditions: string;
    requiredData: string;
  }) {
    const table = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        // Fila 1: Cabeceras Principales
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: 'Funcionalidad', alignment: AlignmentType.CENTER })], shading: { fill: 'D3D3D3', type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ text: 'Escenario', alignment: AlignmentType.CENTER })], shading: { fill: 'D3D3D3', type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ text: 'Tipo', alignment: AlignmentType.CENTER })], shading: { fill: 'D3D3D3', type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER }),
          ],
        }),
        // Fila 2: Datos Principales
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph(data.functionality)], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph(data.scenario)], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph(data.type)], verticalAlign: VerticalAlign.CENTER }),
          ],
        }),
        // Fila 3: Cabeceras Secundarias
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: 'Precondiciones', alignment: AlignmentType.CENTER })], shading: { fill: '4F81BD', type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ text: 'Datos Requeridos', alignment: AlignmentType.CENTER })], columnSpan: 2, shading: { fill: '4F81BD', type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER }),
          ],
        }),
        // Fila 4: Datos Secundarios
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph(data.preconditions)], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph(data.requiredData)], columnSpan: 2, verticalAlign: VerticalAlign.CENTER }),
          ],
        }),
        // Fila 5: Cabecera "Pasos"
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: 'Pasos', alignment: AlignmentType.CENTER })], columnSpan: 3, shading: { fill: '4F81BD', type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER }),
          ],
        }),
      ],
    });
    this.children.push(table);
    this.stepCounter = 1;
  }

  public async addStep(page: Page, description: string) {
    this.children.push(
      new Paragraph({
        children: [new TextRun({ text: `Paso ${this.stepCounter}: ${description}`, bold: true, size: 24 })],
      })
    );
    const screenshotBuffer = await page.screenshot();
    this.children.push(
      new Paragraph({
        children: [
          new ImageRun({
            data: screenshotBuffer,
            type: 'png',
            transformation: { width: 600, height: 337.5 },
          }),
        ],
      })
    );
    this.stepCounter++;
  }

  public addPageBreak() {
    this.children.push(new Paragraph({ pageBreakBefore: true }));
  }

  /**
   * Inserta un párrafo vacío (un salto de línea) en el documento.
   */
  public addLineBreak() {
    this.children.push(new Paragraph({}));
  }

  public async save(filePath: string) {
    const doc = new Document({
      sections: [{
        // CORRECCIÓN: 'headers' debe ir al mismo nivel que la sección (no dentro de 'properties').
        headers: {
          default: this.header,
        },
        children: this.children,
      }],
    });

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);
    console.log(`Reporte guardado en: ${filePath}`);
  }
}