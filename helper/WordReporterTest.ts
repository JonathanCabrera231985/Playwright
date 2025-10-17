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
  ISectionOptions,
} from 'docx';
import * as fs from 'fs';

// Tipo para definir las opciones de cada celda en la tabla personalizada
export type TableCellOptions = {
  text: string;
  backgroundColor?: string;
  textColor?: string;
  bold?: boolean;
  columnSpan?: number;
  rowSpan?: number;
};

export class WordReporterTest {
  private children: (Paragraph | Table)[] = [];
  private header?: Header;
  private stepCounter = 1;

  constructor() {}

/**
   * Obtiene la fecha actual formateada como dd/mm/yyyy.
   * @returns La fecha formateada como un string.
   */
  public getFormattedDate(): string {
    const fecha = new Date();
    const dia = String(fecha.getDate()).padStart(2, '0');
    const mes = String(fecha.getMonth() + 1).padStart(2, '0'); // Se suma 1 porque los meses van de 0 a 11
    const anio = fecha.getFullYear();
    return `${dia}/${mes}/${anio}`;
  }

/**
   * Crea una tabla personalizada a partir de una matriz de datos.
   * @param data Una matriz 2D (filas y columnas) con las opciones para cada celda.
   */
  public createCustomTable(data: TableCellOptions[][]) {
    const table = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: data.map(
        (rowData) =>
          new TableRow({
            children: rowData.map(
              (cellData) =>
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: cellData.text,
                          bold: cellData.bold ?? false,
                          color: cellData.textColor ?? '000000',
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  shading: cellData.backgroundColor
                    ? { fill: cellData.backgroundColor, type: ShadingType.CLEAR }
                    : undefined,
                  columnSpan: cellData.columnSpan ?? 1,
                  rowSpan: cellData.rowSpan ?? 1,
                  verticalAlign: VerticalAlign.CENTER,
                })
            ),
          })
      ),
    });

    this.children.push(table);
  }

  // ... (aquí van todos tus otros métodos: createHeader, addTitle, addStep, save, etc.)
  // ... (no los repito para que la respuesta sea más corta, pero deben estar aquí)




  public createHeader(imagePath: string) {
    const logoBuffer = fs.readFileSync(imagePath);
    this.header = new Header({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new ImageRun({
              data: logoBuffer,
              type: 'png',
              transformation: { width: 150, height: 50 },
            }),
          ],
        }),
      ],
    });
  }

  public addTitle(text: string, bold: boolean = true, size: number = 32, color: string = '4F81BD',alignment: (typeof AlignmentType)[keyof typeof AlignmentType] = AlignmentType.LEFT ) {
    this.children.push(
      new Paragraph({
        children: [new TextRun({ text, bold: bold, size: size, color: color})],
        spacing: { after: 200 },
        alignment: alignment,
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
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: 'Funcionalidad', alignment: AlignmentType.CENTER })], shading: { fill: 'D3D3D3', type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ text: 'Escenario', alignment: AlignmentType.CENTER })], shading: { fill: 'D3D3D3', type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ text: 'Tipo', alignment: AlignmentType.CENTER })], shading: { fill: 'D3D3D3', type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph(data.functionality)], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph(data.scenario)], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph(data.type)], verticalAlign: VerticalAlign.CENTER }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: 'Precondiciones', alignment: AlignmentType.CENTER })], shading: { fill: '4F81BD', type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ text: 'Datos Requeridos', alignment: AlignmentType.CENTER })], columnSpan: 2, shading: { fill: '4F81BD', type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph(data.preconditions)], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph(data.requiredData)], columnSpan: 2, verticalAlign: VerticalAlign.CENTER }),
          ],
        }),
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
          new ImageRun({ data: screenshotBuffer, type: 'png', transformation: { width: 600, height: 337.5 } }),
        ],
      })
    );
    this.stepCounter++;
  }

  public addPageBreak() {
    this.children.push(new Paragraph({ pageBreakBefore: true }));
  }

  public addLineBreak(count: number = 1) {
    for (let i = 0; i < count; i++) {
      this.children.push(new Paragraph({}));
    }
  }

 /* public async save(filePath: string) {
    // Build a mutable properties object instead of assigning to a read-only property.
    const baseSection: ISectionOptions = {
        children: this.children,
    };

    // Start with existing properties if any, otherwise an empty object we can modify.
    const properties: any = baseSection.properties ? { ...baseSection.properties } : {};

    if (this.header) {
        properties.headers = {
            default: this.header,
        };
    }

    const sectionOptions: ISectionOptions = {
        ...baseSection,
        properties,
    };

    const doc = new Document({
      sections: [sectionOptions],
    });*/
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

public async measureStep(page: Page, description: string, action: () => Promise<void>) {
    const startTime = Date.now();
    
    // Ejecuta las acciones que le pasaste al método
    await action();
    
    const endTime = Date.now();
    const durationInSeconds = ((endTime - startTime) / 1000).toFixed(2);

    // Agrega la descripción y el tiempo al documento
    this.children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `Paso ${this.stepCounter}: ${description} (Duración: ${durationInSeconds}s)`,
            bold: true,
            size: 24,
          }),
        ],
      })
    );
    
    // Agrega la captura de pantalla del estado final
    const screenshotBuffer = await page.screenshot();
    this.children.push(
      new Paragraph({
        children: [
          new ImageRun({ data: screenshotBuffer, type: 'png', transformation: { width: 600, height: 337.5 } }),
        ],
      })
    );

    this.stepCounter++;
    return durationInSeconds;
  }

}