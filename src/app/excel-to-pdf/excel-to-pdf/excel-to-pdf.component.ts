import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import { PDFDocument, rgb, StandardFonts, PDFFont } from 'pdf-lib';
import { CommonModule } from '@angular/common';
import JSZip from 'jszip';

type FieldConfig = {
  x: number;
  y: number;
  width: number;
  height: number;
  fontSize?: number;
  bold?: boolean;
};

@Component({
  selector: 'app-excel-to-pdf',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './excel-to-pdf.component.html',
})
export class ExcelToPdfComponent {
  currentYear = new Date().getFullYear();
  loading = false;
  excelFile: File | null = null;
  selectedFileName = '';

  /* =========================
     FILE SELECTION
  ========================== */
  onFileSelect(event: Event) {
    const input = event.target as HTMLInputElement;
    if (!input.files?.length) return;

    this.excelFile = input.files[0];
    this.selectedFileName = this.excelFile.name;
  }

  downloadSampleExcel() {
    const a = document.createElement('a');
    a.href = 'assets/sample_invoice.xlsx';
    a.download = 'sample-excel-format.xlsx';
    a.click();
  }

  /* =========================
     PDF FIELD HELPER
  ========================== */
  drawField(
    page: any,
    text: any,
    config: FieldConfig,
    normalFont: PDFFont,
    boldFont: PDFFont
  ) {
    const font = config.bold ? boldFont : normalFont;
    const fontSize = config.fontSize ?? 10;

    page.drawRectangle({
      x: config.x,
      y: config.y-2,
      width: config.width,
      height: config.height,
      color: rgb(1, 1, 1),
    });

    page.drawText(String(text ?? ''), {
      x: config.x + 2,
      y: config.y ,
      size: fontSize,
      font,
      color: rgb(0, 0, 0),
      maxWidth: config.width - 4,
      lineHeight: fontSize + 2,
    });
  }

  /* =========================
     FIELD COORDINATES
  ========================== */
  FIELDS: Record<string, FieldConfig> = {
    Invoice_No: { x: 267, y: 743, width: 90, height: 10, fontSize: 9, bold: true },
    Invoice_Date: { x: 380, y: 742, width: 80, height: 12, fontSize: 9 },

    Buyer_Name: { x: 35, y: 680, width: 200, height: 10, fontSize: 10, bold: true },
    Buyer_Address: { x: 35, y: 670, width: 200, height: 10, fontSize: 9 },
    Buyer_City: { x: 35, y: 660, width: 200, height: 10, fontSize: 9 },
    Buyer_GSTIN: { x: 119, y: 628, width: 100, height: 10, fontSize: 9, bold: true },

    Consignee_Name: { x: 35, y: 595, width: 200, height: 10, fontSize: 10, bold: true },
    Consignee_Address: { x: 35, y: 584, width: 200, height: 12, fontSize: 9 },
    Consignee_GSTIN: { x: 119, y: 551, width: 100, height: 10, fontSize: 9 },

    Item_Description: { x: 49, y: 505, width: 180, height: 14, fontSize: 9 },
    HSN_Code: { x: 270, y: 505, width: 40, height: 14, fontSize: 9 },
    Quantity: { x: 320, y: 505, width: 40, height: 14, fontSize: 9 },
    Rate: { x: 370, y: 505, width: 35, height: 14, fontSize: 9 },
    Item_Amount: { x: 440, y: 504, width: 52, height: 15, fontSize: 9 },

    CGST_Amount: { x: 443, y: 462, width: 50, height: 10, fontSize: 9 },
    SGST_Amount: { x: 443, y: 450, width: 50, height: 10, fontSize: 9 },

    Total_Amount: { x: 430, y: 203, width: 63, height: 10, fontSize: 10, bold: true },
    Amount_In_Words: { x: 35, y: 180, width: 400, height: 12, fontSize: 9, bold: true },
    Tax_Amount_In_Words: { x: 144, y: 122, width: 348, height: 12, fontSize: 9 },
  };

  /* =========================
     MAIN PDF GENERATOR
  ========================== */
  async generatePdf() {
    if (!this.excelFile) return;
    this.loading = true;

    const buffer = await this.excelFile.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows: any[] = XLSX.utils.sheet_to_json(sheet);

    if (!rows.length) {
      this.loading = false;
      return;
    }

    const pdfTemplateBytes = await fetch('assets/tax.pdf').then(r => r.arrayBuffer());
    const zip = new JSZip();

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];

      const pdfDoc = await PDFDocument.load(pdfTemplateBytes);
      const page = pdfDoc.getPages()[0];

      const normalFont = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

      Object.keys(this.FIELDS).forEach((key) => {
        this.drawField(
          page,
          row[key],
          this.FIELDS[key],
          normalFont,
          boldFont
        );
      });

      const pdfBytes = await pdfDoc.save();
      zip.file(`Invoice_${i + 1}_${row.Buyer_Name || 'User'}.pdf`, pdfBytes);
    }

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const url = URL.createObjectURL(zipBlob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'Invoices.zip';
    a.click();

    URL.revokeObjectURL(url);
    this.loading = false;
  }
}
