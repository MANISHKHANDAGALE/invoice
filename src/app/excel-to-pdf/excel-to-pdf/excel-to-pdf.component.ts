import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import { PDFDocument, rgb } from 'pdf-lib';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'app-excel-to-pdf',
  imports:[CommonModule],
  templateUrl: './excel-to-pdf.component.html',
})
export class ExcelToPdfComponent {

  excelFile: File | null = null;
  selectedFileName = '';

  onFileSelect(event: Event) {
    const input = event.target as HTMLInputElement;
    if (!input.files?.length) return;

    this.excelFile = input.files[0];
    this.selectedFileName = this.excelFile.name;
  }
downloadSampleExcel() {
  const url = 'assets/sample.xlsx';

  const a = document.createElement('a');
  a.href = url;
  a.download = 'sample-excel-format.xlsx';
  a.click();
}


  async generatePdf() {
    if (!this.excelFile) return;

    /* 1️⃣ Read Excel */
    const buffer = await this.excelFile.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows: any[] = XLSX.utils.sheet_to_json(sheet);

    if (!rows.length) return;

    const fullName = `${rows[0].firstName} ${rows[0].lastName}`;

    /* 2️⃣ Load PDF */
    const pdfBytes = await fetch('assets/tax.pdf').then(r => r.arrayBuffer());
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const page = pdfDoc.getPages()[0];

    /* 3️⃣ Write Text */
     page.drawRectangle({
      x: 35, y: 592, width: 200, height: 12, color: rgb(1, 1, 1),
    });
    page.drawText(fullName, { x: 36, y: 595, size: 10 });

    page.drawRectangle({
      x: 35, y: 675, width: 200, height: 12, color: rgb(1, 1, 1),
    });
    page.drawText(fullName, { x: 36, y: 675, size: 10 });

    /* 4️⃣ Download */
    const bytes = await pdfDoc.save();
    const blob = new Blob([bytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'Invoice.pdf';
    a.click();

    URL.revokeObjectURL(url);

  }
}
