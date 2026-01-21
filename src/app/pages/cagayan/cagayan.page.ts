import { Component } from '@angular/core';
import { IonicModule, ToastController } from '@ionic/angular';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { addIcons } from 'ionicons';
import {
  cloudUploadOutline,
  searchOutline,
  micOutline,
  downloadOutline  // ✅ ADDED
} from 'ionicons/icons';
import * as XLSX from 'xlsx';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, UnderlineType } from 'docx';  // ✅ ADDED
import { saveAs } from 'file-saver';  // ✅ ADDED

interface MonthSummary {
  increaseCount: number;
  decreaseCount: number;
  highestIncrease: string[];
  lowestIncrease: string[];
  highestDecrease: string[];
  lowestDecrease: string[];
  totalProducts: number; 
}

interface ComparisonItem {
  name: string;
  unit: string;
  peso: number;
}

interface AnalysisItem {
  name: string;
  unit: string;
  current: number | null;
  weekPrice: number | null;
  month1Price: number | null;
  month3Price: number | null;
  vsWeek: { val: number | null; percent: number | null; status: string };
  vsMonth1: { val: number | null; percent: number | null; status: string };
  vsMonth3: { val: number | null; percent: number | null; status: string };
}

@Component({
  selector: 'app-cagayan',
  templateUrl: './cagayan.page.html',
  styleUrls: ['./cagayan.page.scss'],
  standalone: true,
  imports: [IonicModule, CommonModule, FormsModule]
})
export class CagayanPage {

  selectedFile: File | null = null;
  isGenerated = false;
  viewMode: 'raw' | 'analysis' | 'summary' = 'raw';
  showWarning = false;
  warningMessage = '';

  workbook: XLSX.WorkBook | null = null;
  sheetNames: string[] = [];
  selectedSheetIndex = 0;

  tableHeaders: string[] = [];
  fullTableHeaders: string[] = [];
  tableData: any[][] = [];
  fullTableData: any[][] = [];

  searchQuery: string = '';
  filteredTableData: any[][] = [];

  analysisData: AnalysisItem[] = [];

summary = {
  week: {
    increaseCount: 0,
    decreaseCount: 0,
    highestIncrease: [] as string[],
    lowestIncrease: [] as string[],
    highestDecrease: [] as string[],
    lowestDecrease: [] as string[],
    totalProducts: 0 // ✅
  },
  month1: {
    increaseCount: 0,
    decreaseCount: 0,
    highestIncrease: [] as string[],
    lowestIncrease: [] as string[],
    highestDecrease: [] as string[],
    lowestDecrease: [] as string[],
    totalProducts: 0 // ✅
  },
  month3: {
    increaseCount: 0,
    decreaseCount: 0,
    highestIncrease: [] as string[],
    lowestIncrease: [] as string[],
    highestDecrease: [] as string[],
    lowestDecrease: [] as string[],
    totalProducts: 0 // ✅
  }
};

  provinceName = 'Cagayan Valley';
  reportTitle: string = '';

  currentIndex = -1;
  weekPriceIndex = -1;
  month1PriceIndex = -1;
  month3PriceIndex = -1;
  productNameIndex = -1;
  unitIndex = -1;

  Math = Math;

  constructor(private toastController: ToastController) {
    addIcons({
      cloudUploadOutline,
      searchOutline,
      micOutline,
      downloadOutline  // ✅ ADDED
    });
  }

  onFileSelected(event: any) {
    this.selectedFile = event.target.files[0];
    this.isGenerated = false;
    this.showWarning = false;
  }

  async generateReport(event: Event) {
    event.stopPropagation();
    
    if (!this.selectedFile) {
      this.showWarning = true;
      this.warningMessage = 'Please upload a file before generating the report';
      setTimeout(() => {
        this.showWarning = false;
      }, 5000);
      
      return;
    }
    
    this.showWarning = false;

    const reader = new FileReader();
    reader.onload = (e: any) => {
      const data = new Uint8Array(e.target.result);
      this.workbook = XLSX.read(data, { type: 'array' });
      this.sheetNames = this.workbook.SheetNames;
      this.loadSheet(0);
      this.isGenerated = true;
      this.viewMode = 'raw';
      this.extractReportTitle();
    };
    reader.readAsArrayBuffer(this.selectedFile);
  }

  private extractReportTitle() {
    if (this.selectedFile) {
      const fileName = this.selectedFile.name;
      const nameWithoutExt = fileName.replace(/\.[^/.]+$/, '');
      let title = nameWithoutExt.replace(/\d{4}-\d{2}-\d{2}/g, '').trim();
      title = title.replace(/[-_]/g, ' ').trim();
      this.reportTitle = title || 'Price Analysis';
    } else {
      this.reportTitle = '';
    }
  }

  resetToUpload() {
    this.isGenerated = false;
    this.viewMode = 'raw';
    this.selectedFile = null;
    this.workbook = null;
    this.sheetNames = [];
    this.selectedSheetIndex = 0;
    this.tableHeaders = [];
    this.fullTableHeaders = [];
    this.tableData = [];
    this.fullTableData = [];
    this.filteredTableData = [];
    this.analysisData = [];
    this.searchQuery = '';
    this.showWarning = false;
  }

  loadSheet(index: number) {
    if (!this.workbook) return;

    const sheet = this.workbook.Sheets[this.sheetNames[index]];
    const json: any[][] = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      blankrows: false,
      raw: true,
      defval: null
    });

    let headerRowIndex = -1;
    for (let i = 0; i < json.length; i++) {
      if (
        json[i].some(c =>
          c?.toString().toLowerCase().includes('prevailing')
        )
      ) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) {
      headerRowIndex = 0;
    }

    console.log('=== RAW SHEET DATA ===');
    console.log('Header row index:', headerRowIndex);
    console.log('Raw header row:', json[headerRowIndex]);
    console.log('\n📊 ALL HEADERS WITH INDICES:');
    json[headerRowIndex].forEach((h: any, idx: number) => {
      console.log(`  [${idx}]: ${h}`);
    });
    
    console.log('\n📋 First 5 data rows (full):');
    for (let i = 0; i < Math.min(5, json.length - headerRowIndex - 1); i++) {
      console.log(`Row ${i}:`, json[headerRowIndex + 1 + i]);
    }

    this.tableHeaders = json[headerRowIndex] || [];
    this.fullTableHeaders = [...this.tableHeaders];

    this.fullTableData = json
      .slice(headerRowIndex + 1)
      .filter(r => this.isRowNotEmpty(r));

    this.tableData = [...this.fullTableData];
    this.filteredTableData = [...this.fullTableData];
    this.viewMode = 'raw';
    this.searchQuery = '';

    this.detectColumnIndexes();
  }

  isRowNotEmpty(row: any[]): boolean {
    return row.some(cell => {
      if (cell === null || cell === undefined) return false;
      const cellStr = cell.toString().trim();
      return cellStr !== '';
    });
  }

  detectColumnIndexes() {
    console.log('=== DETECTING COLUMNS ===');
    console.log('All headers:', this.fullTableHeaders);
    
    // Find product name column
    this.productNameIndex = -1;
    for (let i = 0; i < this.fullTableHeaders.length; i++) {
      const header = this.fullTableHeaders[i]?.toLowerCase() || '';
      if (
        header.includes('basic necessities') ||
        header.includes('prime commodities') ||
        header.includes('construction materials') ||
        header.includes('commodity') ||
        header.includes('item') ||
        header.includes('product')
      ) {
        this.productNameIndex = i;
        console.log('✓ Product name found at index:', i, '→', this.fullTableHeaders[i]);
        break;
      }
    }
    
    if (this.productNameIndex === -1) {
      this.productNameIndex = 0;
      console.log('⚠ Product name not found, using first column (0)');
    }

    // Find unit column
    this.unitIndex = -1;
    for (let i = 0; i < this.fullTableHeaders.length; i++) {
      if (!this.fullTableHeaders[i]) continue;
      const header = this.fullTableHeaders[i].toString().toLowerCase().trim();
      if (header === 'unit' || header === 'units') {
        this.unitIndex = i;
        console.log('✓ Unit found at index:', i, '→', this.fullTableHeaders[i]);
        break;
      }
    }

    if (this.unitIndex === -1 && this.productNameIndex !== -1) {
      const potentialUnitIndex = this.productNameIndex + 1;
      if (potentialUnitIndex < this.fullTableHeaders.length && this.looksLikeUnitColumn(potentialUnitIndex)) {
        this.unitIndex = potentialUnitIndex;
        console.log('✓ Unit found by heuristic at index:', potentialUnitIndex);
      }
    }

    // IMPROVED: Find ALL price columns first
    const allPriceColumns: Array<{index: number, header: string, priority: number}> = [];
    
    for (let i = 0; i < this.fullTableHeaders.length; i++) {
      const header = this.fullTableHeaders[i]?.toLowerCase() || '';
      
      const isPriceColumn = 
        header.includes('prevailing') ||
        header.includes('price') ||
        header.includes('retail') ||
        header.includes('peso') ||
        header.includes('₱');
      
      if (isPriceColumn && this.isPriceColumnByData(i)) {
        let priority = 0;
        
        if (header.includes('current') || header.includes('for the week') || (header.includes('prevailing') && !header.includes('ago'))) {
          priority = 1;
        } else if (header.includes('week ago') || header.includes('previous week')) {
          priority = 2;
        } else if (header.includes('1 month ago') || (header.includes('month ago') && !header.includes('3'))) {
          priority = 3;
        } else if (header.includes('3 months ago') || header.includes('three months ago')) {
          priority = 4;
        } else {
          priority = 10 + i;
        }
        
        allPriceColumns.push({ index: i, header: this.fullTableHeaders[i], priority });
        console.log(`Found price column at [${i}]: "${this.fullTableHeaders[i]}" (priority: ${priority})`);
      }
    }
    
    allPriceColumns.sort((a, b) => a.priority - b.priority);
    
    console.log('\n📊 Sorted price columns:', allPriceColumns);
    
    if (allPriceColumns.length >= 1) {
      this.currentIndex = allPriceColumns[0].index;
      console.log('✓ Current price assigned to index:', this.currentIndex, '→', allPriceColumns[0].header);
    }
    
    if (allPriceColumns.length >= 2) {
      this.weekPriceIndex = allPriceColumns[1].index;
      console.log('✓ Week price assigned to index:', this.weekPriceIndex, '→', allPriceColumns[1].header);
    }
    
    if (allPriceColumns.length >= 3) {
      this.month1PriceIndex = allPriceColumns[2].index;
      console.log('✓ Month1 price assigned to index:', this.month1PriceIndex, '→', allPriceColumns[2].header);
    }
    
    if (allPriceColumns.length >= 4) {
      this.month3PriceIndex = allPriceColumns[3].index;
      console.log('✓ Month3 price assigned to index:', this.month3PriceIndex, '→', allPriceColumns[3].header);
    }

    console.log('\n=== FINAL COLUMN INDICES ===');
    console.log('Product:', this.productNameIndex, this.fullTableHeaders[this.productNameIndex]);
    console.log('Unit:', this.unitIndex, this.unitIndex !== -1 ? this.fullTableHeaders[this.unitIndex] : 'N/A');
    console.log('Current Price:', this.currentIndex, this.currentIndex !== -1 ? this.fullTableHeaders[this.currentIndex] : 'N/A');
    console.log('Week Old Price:', this.weekPriceIndex, this.weekPriceIndex !== -1 ? this.fullTableHeaders[this.weekPriceIndex] : 'N/A');
    console.log('Month1 Old Price:', this.month1PriceIndex, this.month1PriceIndex !== -1 ? this.fullTableHeaders[this.month1PriceIndex] : 'N/A');
    console.log('Month3 Old Price:', this.month3PriceIndex, this.month3PriceIndex !== -1 ? this.fullTableHeaders[this.month3PriceIndex] : 'N/A');
  }

  isPriceColumnByData(colIndex: number): boolean {
    let numericCount = 0;
    let totalCount = 0;
    
    for (let row = 0; row < Math.min(15, this.fullTableData.length); row++) {
      const val = this.fullTableData[row][colIndex];
      if (val !== null && val !== undefined && val !== '') {
        totalCount++;
        const parsed = this.parsePrice(val);
        if (parsed !== null && parsed >= 0) {
          numericCount++;
        }
      }
    }
    
    const isPrice = totalCount > 0 && (numericCount / totalCount) > 0.5;
    console.log(`  Column ${colIndex} data check: ${numericCount}/${totalCount} numeric = ${isPrice ? 'YES' : 'NO'}`);
    return isPrice;
  }

  looksLikeUnitColumn(colIndex: number): boolean {
    if (colIndex < 0 || colIndex >= this.fullTableHeaders.length) return false;
    
    const sampleSize = Math.min(10, this.fullTableData.length);
    let unitLikeCount = 0;
    let validCellCount = 0;

    for (let i = 0; i < sampleSize; i++) {
      const cellValue = this.fullTableData[i]?.[colIndex];
      if (!cellValue) continue;
      
      const cellStr = cellValue.toString().toLowerCase().trim();
      if (cellStr === '') continue;
      
      validCellCount++;
      
      if (
        cellStr.includes('kg') ||
        cellStr.includes('kilo') ||
        cellStr.includes('pc') ||
        cellStr.includes('pcs') ||
        cellStr.includes('piece') ||
        cellStr.includes('liter') ||
        cellStr.includes('litre') ||
        cellStr.includes('pack') ||
        cellStr.includes('bag') ||
        cellStr.includes('sack') ||
        cellStr.includes('gram') ||
        cellStr.includes('box') ||
        cellStr.includes('can') ||
        cellStr.includes('bottle') ||
        cellStr.includes('ml') ||
        cellStr.includes('gal') ||
        cellStr.includes('gallon') ||
        /^\d+\s*(kg|g|l|ml|pc|pcs|pack)/i.test(cellStr) ||
        (cellStr.length < 15 && !cellStr.includes(' '))
      ) {
        unitLikeCount++;
      }
    }

    return validCellCount > 0 && (unitLikeCount / validCellCount) > 0.6;
  }

  parsePrice(val: any): number | null {
    if (val === undefined || val === null || val === '') return null;
    const str = val.toString().toUpperCase().trim();
    if (str === 'NO SRP' || str === 'N/A' || str === '-') return null;
    
    const cleaned = val.toString().replace(/[₱,$\s]/g, '');
    const num = parseFloat(cleaned);
    return isNaN(num) || num < 0 ? null : parseFloat(num.toFixed(2));
  }

  onSearchChange() {
    const query = this.searchQuery.toLowerCase().trim();

    if (!query) {
      if (this.viewMode === 'raw') {
        this.tableData = [...this.fullTableData];
      }
      return;
    }

    if (this.viewMode === 'raw') {
      this.tableData = this.fullTableData.filter(row => {
        return row.some(cell => {
          if (!cell) return false;
          return cell.toString().toLowerCase().includes(query);
        });
      });
    }
  }

  clearSearch() {
    this.searchQuery = '';
    this.onSearchChange();
  }

  toggleAnalysis() {
    if (this.viewMode === 'analysis') {
      this.viewMode = 'raw';
      this.tableData = [...this.fullTableData];
    } else {
      this.performAnalysis();
      this.viewMode = 'analysis';
    }
  }

  performAnalysis() {
    const buildComparison = (current: number | null, oldPrice: number | null): 
      { val: number | null; percent: number | null; status: string } => {
      
      if (current === null || oldPrice === null) {
        return { val: null, percent: null, status: 'no-data' };
      }
      
      if (Math.abs(current) < 0.001 || Math.abs(oldPrice) < 0.001) {
        console.log(`Zero detected - Current: ${current}, Old: ${oldPrice} - marking as zero-comparison`);
        return { val: 0, percent: 0, status: 'zero-comparison' };
      }
      
      const diff = parseFloat((current - oldPrice).toFixed(2));
      const percent = oldPrice > 0 ? parseFloat(((diff / oldPrice) * 100).toFixed(2)) : 0;
      
      let status: 'increase' | 'decrease' | 'stable' = 'stable';
      if (diff > 0.001) {
        status = 'increase';
      } else if (diff < -0.001) {
        status = 'decrease';
      }
      
      console.log(`Comparison - Current: ${current}, Old: ${oldPrice}, Diff: ${diff}, Percent: ${percent}, Status: ${status}`);
      
      return { val: diff, percent, status };
    };

    this.analysisData = this.fullTableData.map((row) => {
      const itemName = row[this.productNameIndex];
      if (!itemName || itemName.toString().trim() === '') return null;

      const currentPrice = this.parsePrice(row[this.currentIndex]);
      const weekPrice = this.parsePrice(row[this.weekPriceIndex]);
      const month1Price = this.parsePrice(row[this.month1PriceIndex]);
      const month3Price = this.parsePrice(row[this.month3PriceIndex]);

      console.log(`Item: ${itemName}`);
      console.log(`  Current: ${currentPrice}, Week: ${weekPrice}, M1: ${month1Price}, M3: ${month3Price}`);

      return {
        name: itemName,
        unit: row[this.unitIndex] || '',
        current: currentPrice,
        weekPrice: weekPrice,
        month1Price: month1Price,
        month3Price: month3Price,
        vsWeek: buildComparison(currentPrice, weekPrice),
        vsMonth1: buildComparison(currentPrice, month1Price),
        vsMonth3: buildComparison(currentPrice, month3Price)
      };
    }).filter(item => item !== null) as AnalysisItem[];
    
    console.log('Analysis Data:', this.analysisData);
  }

  generateSummaryReport() {
    if (this.viewMode === 'summary') {
      this.viewMode = 'raw';
      this.tableData = [...this.fullTableData];
      return;
    }

    this.viewMode = 'summary';
    
    if (this.analysisData.length === 0) {
      this.performAnalysis();
    }
    
    this.summary.week = this.calculateMonthSummaryFromAnalysis('week');
    this.summary.month1 = this.calculateMonthSummaryFromAnalysis('month1');
    this.summary.month3 = this.calculateMonthSummaryFromAnalysis('month3');

  }

calculateMonthSummaryFromAnalysis(
  period: 'week' | 'month1' | 'month3'
): MonthSummary {

  const increase: ComparisonItem[] = [];
  const decrease: ComparisonItem[] = [];
  let totalProducts = 0;

  for (const item of this.analysisData) {

    // 🛑 Skip category headers
    const name = item.name?.toString().trim().toLowerCase();
    if (
      !name ||
      name.includes('basic necessities') ||
      name.includes('prime commodities') ||
      name.includes('construction materials')
    ) {
      continue;
    }

    // ✅ Count product if it has ANY valid price
    const hasAnyPrice =
      item.current !== null ||
      item.weekPrice !== null ||
      item.month1Price !== null ||
      item.month3Price !== null;

    if (hasAnyPrice) {
      totalProducts++;
    }

    const comp =
      period === 'week'
        ? item.vsWeek
        : period === 'month1'
        ? item.vsMonth1
        : item.vsMonth3;

    if (
      comp.status === 'no-data' ||
      comp.status === 'stable' ||
      comp.status === 'zero-comparison' ||
      comp.val === null
    ) {
      continue;
    }

    const unit = item.unit || 'N/A';

    if (comp.status === 'increase') {
      increase.push({ name: item.name, unit, peso: comp.val });
    } else if (comp.status === 'decrease') {
      decrease.push({ name: item.name, unit, peso: Math.abs(comp.val) });
    }
  }

  increase.sort((a, b) => a.peso - b.peso);
  decrease.sort((a, b) => a.peso - b.peso);

  const getItems = (arr: ComparisonItem[], val: number) =>
    arr
      .filter(i => Math.abs(i.peso - val) < 0.001)
      .map(i => `₱${i.peso.toFixed(2)} - ${i.name} (${i.unit})`);

  let lowestIncreaseItems: string[] = ['N/A'];
  let highestIncreaseItems: string[] = ['N/A'];

  if (increase.length === 1) {
    lowestIncreaseItems = getItems(increase, increase[0].peso);
  } else if (increase.length > 1) {
    lowestIncreaseItems = getItems(increase, increase[0].peso);
    highestIncreaseItems = getItems(
      increase,
      increase[increase.length - 1].peso
    );
  }

  let lowestDecreaseItems: string[] = ['N/A'];
  let highestDecreaseItems: string[] = ['N/A'];

  if (decrease.length === 1) {
    lowestDecreaseItems = getItems(decrease, decrease[0].peso);
  } else if (decrease.length > 1) {
    lowestDecreaseItems = getItems(decrease, decrease[0].peso);
    highestDecreaseItems = getItems(
      decrease,
      decrease[decrease.length - 1].peso
    );
  }

  return {
    increaseCount: increase.length,
    decreaseCount: decrease.length,
    highestIncrease: highestIncreaseItems,
    lowestIncrease: lowestIncreaseItems,
    highestDecrease: highestDecreaseItems,
    lowestDecrease: lowestDecreaseItems,
    totalProducts // ✅ INCLUDED
  };
}


  // ✅ NEW: EXPORT FUNCTIONALITY
  async exportSummaryReport() {
    if (this.viewMode !== 'summary') {
      const toast = await this.toastController.create({
        message: 'Please generate the summary report first',
        duration: 3000,
        color: 'warning',
        position: 'top'
      });
      await toast.present();
      return;
    }

    try {
      const baseFilename = this.selectedFile?.name.replace(/\.[^/.]+$/, '') || 'Summary_Report';
      const sheetName = this.sheetNames[this.selectedSheetIndex] || 'Sheet';
      const filename = `${baseFilename}_${sheetName}_Summary.docx`;

      const createSummaryItems = (items: string[]): Paragraph[] => {
        return items.map(item => 
          new Paragraph({
            text: `  • ${item}`,
            spacing: { before: 100, after: 100 }
          })
        );
      };

      const doc = new Document({
        sections: [{
          properties: {},
          children: [
            // Title
            new Paragraph({
              text: 'CPD PRICE TRACKER',
              heading: HeadingLevel.HEADING_1,
              alignment: AlignmentType.CENTER,
              spacing: { after: 200 }
            }),
            new Paragraph({
              text: `${this.provinceName} - ${baseFilename}`,
              heading: HeadingLevel.HEADING_2,
              alignment: AlignmentType.CENTER,
              spacing: { after: 400 }
            }),
            new Paragraph({
              text: `File: ${this.selectedFile?.name || 'N/A'}`,
              alignment: AlignmentType.CENTER,
              spacing: { after: 100 }
            }),
            new Paragraph({
              text: `Sheet: ${sheetName}`,
              alignment: AlignmentType.CENTER,
              spacing: { after: 600 }
            }),

            // 1 WEEK COMPARISON SUMMARY
            new Paragraph({
              children: [
                new TextRun({
                  text: '1 WEEK COMPARISON SUMMARY',
                  bold: true,
                  underline: { type: UnderlineType.SINGLE }
                })
              ],
              spacing: { before: 400, after: 300 }
            }),

            // A. Increase (1 Week)
            new Paragraph({
              children: [
                new TextRun({
                  text: 'A. Increase',
                  bold: true,
                  size: 24
                })
              ],
              spacing: { before: 200, after: 200 }
            }),
            new Paragraph({
              text: `Total Increase: ${this.summary.week.increaseCount}`,
              spacing: { before: 100, after: 100 }
            }),
            new Paragraph({
              text: 'Highest Increase:',
              spacing: { before: 100, after: 100 }
            }),
            ...createSummaryItems(this.summary.week.highestIncrease),
            new Paragraph({
              text: 'Lowest Increase:',
              spacing: { before: 200, after: 100 }
            }),
            ...createSummaryItems(this.summary.week.lowestIncrease),

            // B. Decrease (1 Week)
            new Paragraph({
              children: [
                new TextRun({
                  text: 'B. Decrease',
                  bold: true,
                  size: 24
                })
              ],
              spacing: { before: 400, after: 200 }
            }),
            new Paragraph({
              text: `Total Decrease: ${this.summary.week.decreaseCount}`,
              spacing: { before: 100, after: 100 }
            }),
            new Paragraph({
              text: 'Highest Decrease:',
              spacing: { before: 100, after: 100 }
            }),
            ...createSummaryItems(this.summary.week.highestDecrease),
            new Paragraph({
              text: 'Lowest Decrease:',
              spacing: { before: 200, after: 100 }
            }),
            ...createSummaryItems(this.summary.week.lowestDecrease),

            // 1 MONTH COMPARISON SUMMARY
            new Paragraph({
              children: [
                new TextRun({
                  text: '1 MONTH COMPARISON SUMMARY',
                  bold: true,
                  underline: { type: UnderlineType.SINGLE }
                })
              ],
              spacing: { before: 800, after: 300 }
            }),

            // A. Increase (1 Month)
            new Paragraph({
              children: [
                new TextRun({
                  text: 'A. Increase',
                  bold: true,
                  size: 24
                })
              ],
              spacing: { before: 200, after: 200 }
            }),
            new Paragraph({
              text: `Total Increase: ${this.summary.month1.increaseCount}`,
              spacing: { before: 100, after: 100 }
            }),
            new Paragraph({
              text: 'Highest Increase:',
              spacing: { before: 100, after: 100 }
            }),
            ...createSummaryItems(this.summary.month1.highestIncrease),
            new Paragraph({
              text: 'Lowest Increase:',
              spacing: { before: 200, after: 100 }
            }),
            ...createSummaryItems(this.summary.month1.lowestIncrease),

            // B. Decrease (1 Month)
            new Paragraph({
              children: [
                new TextRun({
                  text: 'B. Decrease',
                  bold: true,
                  size: 24
                })
              ],
              spacing: { before: 400, after: 200 }
            }),
            new Paragraph({
              text: `Total Decrease: ${this.summary.month1.decreaseCount}`,
              spacing: { before: 100, after: 100 }
            }),
            new Paragraph({
              text: 'Highest Decrease:',
              spacing: { before: 100, after: 100 }
            }),
            ...createSummaryItems(this.summary.month1.highestDecrease),
            new Paragraph({
              text: 'Lowest Decrease:',
              spacing: { before: 200, after: 100 }
            }),
            ...createSummaryItems(this.summary.month1.lowestDecrease),

            // 3 MONTHS COMPARISON SUMMARY
            new Paragraph({
              children: [
                new TextRun({
                  text: '3 MONTHS COMPARISON SUMMARY',
                  bold: true,
                  underline: { type: UnderlineType.SINGLE }
                })
              ],
              spacing: { before: 800, after: 300 }
            }),

            // A. Increase (3 Months)
            new Paragraph({
              children: [
                new TextRun({
                  text: 'A. Increase',
                  bold: true,
                  size: 24
                })
              ],
              spacing: { before: 200, after: 200 }
            }),
            new Paragraph({
              text: `Total Increase: ${this.summary.month3.increaseCount}`,
              spacing: { before: 100, after: 100 }
            }),
            new Paragraph({
              text: 'Highest Increase:',
              spacing: { before: 100, after: 100 }
            }),
            ...createSummaryItems(this.summary.month3.highestIncrease),
            new Paragraph({
              text: 'Lowest Increase:',
              spacing: { before: 200, after: 100 }
            }),
            ...createSummaryItems(this.summary.month3.lowestIncrease),

            // B. Decrease (3 Months)
            new Paragraph({
              children: [
                new TextRun({
                  text: 'B. Decrease',
                  bold: true,
                  size: 24
                })
              ],
              spacing: { before: 400, after: 200 }
            }),
            new Paragraph({
              text: `Total Decrease: ${this.summary.month3.decreaseCount}`,
              spacing: { before: 100, after: 100 }
            }),
            new Paragraph({
              text: 'Highest Decrease:',
              spacing: { before: 100, after: 100 }
            }),
            ...createSummaryItems(this.summary.month3.highestDecrease),
            new Paragraph({
              text: 'Lowest Decrease:',
              spacing: { before: 200, after: 100 }
            }),
            ...createSummaryItems(this.summary.month3.lowestDecrease)
          ]
        }]
      });

      const blob = await Packer.toBlob(doc);
      setTimeout(() => {
        saveAs(blob, filename);
      }, 300);
      
      const toast = await this.toastController.create({
        message: 'Summary report exported successfully',
        duration: 3000,
        color: 'success',
        position: 'top'
      });
      await toast.present();

    } catch (error) {
      console.error('Error exporting summary report:', error);
      const toast = await this.toastController.create({
        message: 'Failed to export summary report',
        duration: 3000,
        color: 'danger',
        position: 'top'
      });
      await toast.present();
    }
  }

  parseNumber(value: any): number | null {
    if (!value) return null;
    const num = parseFloat(value.toString().replace(/[₱,$\s]/g, ''));
    return isNaN(num) ? null : num;
  }

  selectSheet(index: number) {
    this.selectedSheetIndex = index;
    this.loadSheet(index);
  }

  onSheetChange() {
    this.loadSheet(this.selectedSheetIndex);
  }

  get filteredAnalysisData() {
    const term = this.searchQuery.toLowerCase();
    return this.analysisData.filter(item => item.name?.toLowerCase().includes(term));
  }

  isNumber(value: any): boolean {
    if (value === null || value === undefined || value === '') return false;
    return typeof value === 'number' || (!isNaN(parseFloat(value)) && isFinite(value));
  }
}