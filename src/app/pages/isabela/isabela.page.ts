import { Component } from '@angular/core';
import { IonicModule, ToastController } from '@ionic/angular'; // Added ToastController
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { addIcons } from 'ionicons';
import {
  cloudUploadOutline,
  searchOutline,
  micOutline,
  arrowBackOutline,
  downloadOutline
} from 'ionicons/icons';

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  UnderlineType
} from 'docx';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-isabela',
  templateUrl: './isabela.page.html',
  styleUrls: ['./isabela.page.scss'],
  standalone: true,
  imports: [IonicModule, CommonModule, FormsModule]
})
export class IsabelaPage {
  selectedFile: File | null = null;
  isGenerated = false;
  showAnalysis = false;
  viewMode: 'raw' | 'analysis' | 'summary' = 'raw';
  searchText = '';
  workbook: any = null;
  sheetNames: string[] = [];
  tableHeaders: string[] = [];
  tableData: any[][] = [];
  rawJsonData: any[][] = [];
  headerRowIndex: number = -1;
  analysisData: any[] = [];
  reportDate: string = '';
  dateRowData: any[] = [];
  reportTitle: string = '';
  summaryData: any = null;
  showWarning = false;
  warningMessage = '';
  provinceName = 'Isabela';

  // ⬇️ CHANGE: Added ToastController to constructor
  constructor(private toastController: ToastController) {
    addIcons({
      cloudUploadOutline,
      searchOutline,
      micOutline,
      arrowBackOutline,
      downloadOutline
    });
  }

  summary: {
    week: any;
    month: any;
    threeMonth: any;
  } = {
    week: {
      totalIncrease: 0,
      totalDecrease: 0,
      totalProducts: 0,
      highestIncrease: [],
      lowestIncrease: [],
      highestDecrease: [],
      lowestDecrease: []
    },
    month: {
      totalIncrease: 0,
      totalDecrease: 0,
      totalProducts: 0,
      highestIncrease: [],
      lowestIncrease: [],
      highestDecrease: [],
      lowestDecrease: []
    },
    threeMonth: {
      totalIncrease: 0,
      totalDecrease: 0,
      totalProducts: 0,
      highestIncrease: [],
      lowestIncrease: [],
      highestDecrease: [],
      lowestDecrease: []
    }
  };


  onFileSelected(event: any) {
    const file = event.target.files[0];
    if (file) {
      this.selectedFile = file;
      this.isGenerated = false;
      this.reportDate = '';
      this.dateRowData = [];
      this.reportTitle = '';
      this.showWarning = false; 
    }
  }

  goBack() {
    this.isGenerated = false;
    this.selectedFile = null;
    this.viewMode = 'raw';
    this.searchText = '';
    this.tableData = [];
    this.rawJsonData = [];
    this.analysisData = [];
    this.summaryData = null;
    this.showWarning = false;
  }

  generateReport(event: Event) {
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
      this.showAnalysis = false;
      this.viewMode = 'raw';

      this.extractDate();
      this.extractReportTitle();
    };
    reader.readAsArrayBuffer(this.selectedFile);
  }

  loadSheet(index: number) {
    if (!this.workbook) return;
    const sheet = this.workbook.Sheets[this.sheetNames[index]];
    this.rawJsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false, raw: false });

    console.log('=== COMPLETE RAW DATA (first 10 rows) ===');
    this.rawJsonData.slice(0, 10).forEach((row, i) => {
      console.log(`Row ${i}:`, row);
    });

    // Find header row
    this.headerRowIndex = this.rawJsonData.findIndex(row => 
      row.some(cell => cell?.toString().toUpperCase().includes('CURRENT WEEK'))
    );
    
    if (this.headerRowIndex === -1) {
      this.headerRowIndex = 4;
    }

    console.log('=== HEADER INFO ===');
    console.log('Header row index:', this.headerRowIndex);
    console.log('Headers:', this.rawJsonData[this.headerRowIndex]);

    this.tableHeaders = this.rawJsonData[this.headerRowIndex] || [];
    
    // Date row
    if (this.rawJsonData[this.headerRowIndex + 1]) {
      this.dateRowData = this.rawJsonData[this.headerRowIndex + 1];
      console.log('Date row:', this.dateRowData);
    }
    
    // Process table data for display - round numbers
    const dataStartIndex = this.headerRowIndex + 2;
    this.tableData = this.rawJsonData.slice(dataStartIndex).map(row => 
      row.map(cell => {
        if (cell === undefined || cell === null || cell === '') return cell;
        const num = parseFloat(cell.toString().replace(/[^\d.-]/g, ''));
        if (!isNaN(num) && cell.toString().match(/^\d+\.?\d*$/)) {
          return parseFloat(num.toFixed(2));
        }
        return cell;
      })
    ).filter(row => row.length > 0 && row.some(cell => cell !== undefined && cell !== null && cell !== ''));

    console.log('=== PROCESSED DATA (first 5 rows) ===');
    this.tableData.slice(0, 5).forEach((row, i) => {
      console.log(`Data row ${i}:`, row);
    });

    this.showAnalysis = false;
    this.viewMode = 'raw';
  }

  private getColIndex(keyword: string): number {
    const idx = this.tableHeaders.findIndex(h => 
      h?.toString().toUpperCase().includes(keyword.toUpperCase())
    );
    return idx;
  }

  toggleAnalysis() {
    if (this.viewMode === 'analysis') {
      this.viewMode = 'raw';
      this.showAnalysis = false;
    } else {
      this.performAnalysis();
      this.viewMode = 'analysis';
      this.showAnalysis = true;
    }
  }

  showSummary() {
    this.viewMode = 'summary';
    this.showAnalysis = false;
    this.generateSummary();
  }

  performAnalysis() {
    const idxItem = 0;
    const idxSize = 1;
    
    let idxCurrent = this.getColIndex('CURRENT WEEK');
    let idxWeekAgo = this.getColIndex('A WEEK AGO');
    let idxMonthAgo = this.getColIndex('A MONTH AGO');
    let idx3MonthAgo = this.getColIndex('3 MONTHS AGO');

    console.log('=== INITIAL COLUMN INDICES ===');
    console.log('Current Week:', idxCurrent);
    console.log('Week Ago:', idxWeekAgo);
    console.log('Month Ago:', idxMonthAgo);
    console.log('3 Months Ago:', idx3MonthAgo);
    
    if (this.dateRowData[idxCurrent]?.toString().toLowerCase().includes('vs') ||
        this.dateRowData[idxCurrent]?.toString().toLowerCase().includes('current vs')) {
      idxCurrent = idxCurrent - 1;
      console.log('Adjusted Current Week to:', idxCurrent);
    }
    
    if (this.dateRowData[idxWeekAgo]?.toString().toLowerCase().includes('vs')) {
      idxWeekAgo = idxWeekAgo - 1;
      console.log('Adjusted Week Ago to:', idxWeekAgo);
    }
    
    if (this.dateRowData[idxMonthAgo]?.toString().toLowerCase().includes('vs')) {
      idxMonthAgo = idxMonthAgo - 1;
      console.log('Adjusted Month Ago to:', idxMonthAgo);
    }
    
    if (this.dateRowData[idx3MonthAgo]?.toString().toLowerCase().includes('vs')) {
      idx3MonthAgo = idx3MonthAgo - 1;
      console.log('Adjusted 3 Months Ago to:', idx3MonthAgo);
    }

    console.log('=== FINAL ANALYSIS COLUMN INDICES ===');
    console.log('Item:', idxItem, 'Size:', idxSize);
    console.log('Current Week:', idxCurrent);
    console.log('Week Ago:', idxWeekAgo);
    console.log('Month Ago:', idxMonthAgo);
    console.log('3 Months Ago:', idx3MonthAgo);

    if (idxCurrent === -1 || idxWeekAgo === -1 || idxMonthAgo === -1 || idx3MonthAgo === -1) {
      console.error('ERROR: Could not find all required columns!');
      console.log('Available headers:', this.tableHeaders);
      console.log('Date row:', this.dateRowData);
      alert('Error: Could not find all required columns in the spreadsheet. Please check the file format.');
      return;
    }

    const parsePrice = (val: any) => {
      if (val === undefined || val === null || val === '') return null;
      const str = val.toString().toUpperCase();
      if (str === 'NO SRP' || str === 'N/A' || str === '-') return null;
      
      const cleaned = val.toString().replace(/[^\d.-]/g, '');
      const num = parseFloat(cleaned);
      return isNaN(num) || num < 0 ? null : parseFloat(num.toFixed(2));
    };

    this.analysisData = this.tableData.map((row, idx) => {
      const itemName = row[idxItem];
      
      if (!itemName) {
        return null;
      }

      const currentPrice = parsePrice(row[idxCurrent]);
      const weekPrice = parsePrice(row[idxWeekAgo]);
      const monthPrice = parsePrice(row[idxMonthAgo]);
      const threeMonthPrice = parsePrice(row[idx3MonthAgo]);

      if (idx < 3) {
        console.log(`=== ITEM ${idx}: ${itemName} ===`);
        console.log('Raw values from row:', {
          current: row[idxCurrent],
          week: row[idxWeekAgo],
          month: row[idxMonthAgo],
          threeMonth: row[idx3MonthAgo]
        });
        console.log('Parsed prices:', {
          current: currentPrice,
          week: weekPrice,
          month: monthPrice,
          threeMonth: threeMonthPrice
        });
      }

      const calcDiff = (current: number | null, past: number | null) => {
        if (current === null || past === null) {
          return { val: null, percent: null, status: 'no data' };
        }
        
        if (current === 0 && past === 0) {
          return { val: 0, percent: 0, status: 'stable' };
        }
        
        if (past === 0) {
          return { val: current, percent: null, status: current > 0 ? 'increase' : 'stable' };
        }
        
        const diff = parseFloat((current - past).toFixed(2));
        const percent = parseFloat(((diff / past) * 100).toFixed(2));
        
        let status: 'increase' | 'decrease' | 'stable' = 'stable';
        if (Math.abs(diff) >= 0.01) {
          status = diff > 0 ? 'increase' : 'decrease';
        }
        return { val: diff, percent: percent, status };
      };

      return {
        name: itemName,
        unit: row[idxSize] || '',
        current: currentPrice,
        weekAgoPrice: weekPrice,
        monthAgoPrice: monthPrice,
        threeMonthPrice: threeMonthPrice,
        vsWeek: calcDiff(currentPrice, weekPrice),
        vsMonth: calcDiff(currentPrice, monthPrice),
        vs3Month: calcDiff(currentPrice, threeMonthPrice)
      };
    }).filter(item => item !== null) as any[];

    console.log('=== FINAL ANALYSIS DATA (first 5) ===');
    console.log(this.analysisData.slice(0, 5));
    console.log('Total analysis items:', this.analysisData.length);
    console.log('Total table data items:', this.tableData.length);
  }

  private extractDate() {
    if (!this.workbook) return;
    const firstSheet = this.workbook.Sheets[this.sheetNames[0]];
    
    const firstCell = firstSheet['A1']?.v;
    if (firstCell && typeof firstCell === 'string' && firstCell.match(/\d{4}-\d{2}-\d{2}/)) {
      this.reportDate = firstCell;
      return;
    }
    
    if (this.selectedFile) {
      const match = this.selectedFile.name.match(/\d{4}-\d{2}-\d{2}/);
      if (match) {
        this.reportDate = match[0];
        return;
      }
    }
    
    if (!this.reportDate) {
      const today = new Date();
      this.reportDate = today.toISOString().split('T')[0];
    }
  }

  get filteredTableData() {
    const term = this.searchText.toLowerCase();
    return this.tableData.filter(row => row.some(cell => cell?.toString().toLowerCase().includes(term)));
  }

  get filteredAnalysisData() {
    const term = this.searchText.toLowerCase();
    return this.analysisData.filter(item => item.name?.toLowerCase().includes(term));
  }

  isCategoryRow(row: any[]): boolean {
    return row[0] && typeof row[0] === 'string' && 
           (!row[2] || row[2].toString().trim() === '');
  }

  private extractReportTitle() {
    if (this.selectedFile) {
      const fileName = this.selectedFile.name;
      const nameWithoutExt = fileName.replace(/\.[^/.]+$/, '');
      
      let title = nameWithoutExt.replace(/\d{4}-\d{2}-\d{2}/g, '').trim();
      
      title = title.replace(/[-_]/g, ' ').trim();
      
      this.reportTitle = title || 'Monthly Report';
    } else {
      this.reportTitle = 'Monthly Report';
    }
  }

  generateSummary() {
    if (this.analysisData.length === 0) {
      this.performAnalysis();
    }

    const totalProducts = this.analysisData.filter(item => {
      const name = item.name?.toString().trim();
      if (!name) return false;

      const nameLower = name.toLowerCase();
      const categoryIndicators = [
        'basic necessities',
        'prime commodities',
        'construction materials',
        'category',
        'subtotal'
      ];
      if (categoryIndicators.some(ind => nameLower.includes(ind))) return false;

      const hasAnyPrice = [item.current, item.weekAgoPrice, item.monthAgoPrice, item.threeMonthPrice]
        .some(p => p !== null && p !== undefined && p !== '');
      const hasUnit = item.unit && item.unit.toString().trim() !== '';

      return hasAnyPrice || hasUnit;
    }).length;

    const weekSummary = {
      totalIncrease: 0,
      totalDecrease: 0,
      totalProducts,
      highestIncrease: [] as Array<{ item: string, unit: string, value: number }>,
      lowestIncrease: [] as Array<{ item: string, unit: string, value: number }>,
      highestDecrease: [] as Array<{ item: string, unit: string, value: number }>,
      lowestDecrease: [] as Array<{ item: string, unit: string, value: number }>
    };

    const monthSummary = {
      totalIncrease: 0,
      totalDecrease: 0,
      totalProducts,
      highestIncrease: [] as Array<{ item: string, unit: string, value: number }>,
      lowestIncrease: [] as Array<{ item: string, unit: string, value: number }>,
      highestDecrease: [] as Array<{ item: string, unit: string, value: number }>,
      lowestDecrease: [] as Array<{ item: string, unit: string, value: number }>
    };

    const threeMonthSummary = {
      totalIncrease: 0,
      totalDecrease: 0,
      totalProducts,
      highestIncrease: [] as Array<{ item: string, unit: string, value: number }>,
      lowestIncrease: [] as Array<{ item: string, unit: string, value: number }>,
      highestDecrease: [] as Array<{ item: string, unit: string, value: number }>,
      lowestDecrease: [] as Array<{ item: string, unit: string, value: number }>
    };

    const weekIncreases: Array<{ item: string, unit: string, value: number }> = [];
    const weekDecreases: Array<{ item: string, unit: string, value: number }> = [];
    const monthIncreases: Array<{ item: string, unit: string, value: number }> = [];
    const monthDecreases: Array<{ item: string, unit: string, value: number }> = [];
    const threeMonthIncreases: Array<{ item: string, unit: string, value: number }> = [];
    const threeMonthDecreases: Array<{ item: string, unit: string, value: number }> = [];

    this.analysisData.forEach(item => {
      const hasValidWeekData = item.current !== null && item.current !== 0 &&
                               item.weekAgoPrice !== null && item.weekAgoPrice !== 0;
      const hasValidMonthData = item.current !== null && item.current !== 0 &&
                                item.monthAgoPrice !== null && item.monthAgoPrice !== 0;
      const hasValid3MonthData = item.current !== null && item.current !== 0 &&
                                 item.threeMonthPrice !== null && item.threeMonthPrice !== 0;

      if (hasValidWeekData && item.vsWeek.status === 'increase' && item.vsWeek.val !== null) {
        weekSummary.totalIncrease++;
        weekIncreases.push({ item: item.name, unit: item.unit, value: item.vsWeek.val });
      } else if (hasValidWeekData && item.vsWeek.status === 'decrease' && item.vsWeek.val !== null) {
        weekSummary.totalDecrease++;
        weekDecreases.push({ item: item.name, unit: item.unit, value: Math.abs(item.vsWeek.val) });
      }

      if (hasValidMonthData && item.vsMonth.status === 'increase' && item.vsMonth.val !== null) {
        monthSummary.totalIncrease++;
        monthIncreases.push({ item: item.name, unit: item.unit, value: item.vsMonth.val });
      } else if (hasValidMonthData && item.vsMonth.status === 'decrease' && item.vsMonth.val !== null) {
        monthSummary.totalDecrease++;
        monthDecreases.push({ item: item.name, unit: item.unit, value: Math.abs(item.vsMonth.val) });
      }

      if (hasValid3MonthData && item.vs3Month.status === 'increase' && item.vs3Month.val !== null) {
        threeMonthSummary.totalIncrease++;
        threeMonthIncreases.push({ item: item.name, unit: item.unit, value: item.vs3Month.val });
      } else if (hasValid3MonthData && item.vs3Month.status === 'decrease' && item.vs3Month.val !== null) {
        threeMonthSummary.totalDecrease++;
        threeMonthDecreases.push({ item: item.name, unit: item.unit, value: Math.abs(item.vs3Month.val) });
      }
    });

    weekSummary.highestIncrease = weekIncreases.length === 1 ? [] : this.findAllWithValue(weekIncreases, Math.max(...(weekIncreases.map(i => i.value)), 0));
    weekSummary.lowestIncrease = this.findAllWithValue(weekIncreases, Math.min(...(weekIncreases.map(i => i.value)), Infinity));
    weekSummary.highestDecrease = weekDecreases.length === 1 ? [] : this.findAllWithValue(weekDecreases, Math.max(...(weekDecreases.map(i => i.value)), 0));
    weekSummary.lowestDecrease = this.findAllWithValue(weekDecreases, Math.min(...(weekDecreases.map(i => i.value)), Infinity));

    monthSummary.highestIncrease = monthIncreases.length === 1 ? [] : this.findAllWithValue(monthIncreases, Math.max(...(monthIncreases.map(i => i.value)), 0));
    monthSummary.lowestIncrease = this.findAllWithValue(monthIncreases, Math.min(...(monthIncreases.map(i => i.value)), Infinity));
    monthSummary.highestDecrease = monthDecreases.length === 1 ? [] : this.findAllWithValue(monthDecreases, Math.max(...(monthDecreases.map(i => i.value)), 0));
    monthSummary.lowestDecrease = this.findAllWithValue(monthDecreases, Math.min(...(monthDecreases.map(i => i.value)), Infinity));

    threeMonthSummary.highestIncrease = threeMonthIncreases.length === 1 ? [] : this.findAllWithValue(threeMonthIncreases, Math.max(...(threeMonthIncreases.map(i => i.value)), 0));
    threeMonthSummary.lowestIncrease = this.findAllWithValue(threeMonthIncreases, Math.min(...(threeMonthIncreases.map(i => i.value)), Infinity));
    threeMonthSummary.highestDecrease = threeMonthDecreases.length === 1 ? [] : this.findAllWithValue(threeMonthDecreases, Math.max(...(threeMonthDecreases.map(i => i.value)), 0));
    threeMonthSummary.lowestDecrease = this.findAllWithValue(threeMonthDecreases, Math.min(...(threeMonthDecreases.map(i => i.value)), Infinity));

    this.summaryData = {
      week: weekSummary,
      month: monthSummary,
      threeMonth: threeMonthSummary
    };
    this.summary = this.summaryData as any;
  }

  private findAllWithValue(items: Array<{ item: string, unit: string, value: number }>, targetValue: number): Array<{ item: string, unit: string, value: number }> {
    if (!isFinite(targetValue) || items.length === 0) return [];
    return items.filter(i => Math.abs(i.value - targetValue) < 0.001);
  }
  
  async exportSummaryReport() {
    if (!this.summaryData) {
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
      const baseFilename = this.selectedFile?.name.replace(/\.[^/.]+$/, '') || 'Isabela_Report';
      const filename = `${baseFilename}_${this.reportDate}_Summary.docx`;

      // Helper to create bullet list paragraphs or show N/A
      const makeItems = (items: any[]) => {
        return items && items.length
          ? items.map(i =>
              new Paragraph({
                text: `• ${i.item}${i.unit ? ` (${i.unit})` : ''} — ₱${i.value.toFixed(2)}`
              })
            )
          : [new Paragraph({ text: '• N/A' })];
      };

      const makeSubHeader = (text: string) =>
        new Paragraph({
          text,
          spacing: { before: 100 }
        });

      // Make section now includes Total Products, increases and decreases and both highest/lowest lists
      const makeSection = (title: string, data: any) => {
        return [
          new Paragraph({
            children: [
              new TextRun({
                text: title,
                bold: true,
                underline: { type: UnderlineType.SINGLE }
              })
            ],
            spacing: { before: 400, after: 200 }
          }),

          new Paragraph({
            text: `Total Products: ${data.totalProducts}`
          }),

          new Paragraph({
            text: `Total Increase: ${data.totalIncrease}`
          }),

          makeSubHeader('Highest Increase:'),
          ...makeItems(data.highestIncrease),

          makeSubHeader('Lowest Increase:'),
          ...makeItems(data.lowestIncrease),

          new Paragraph({
            text: `Total Decrease: ${data.totalDecrease}`,
            spacing: { before: 200 }
          }),

          makeSubHeader('Highest Decrease:'),
          ...makeItems(data.highestDecrease),

          makeSubHeader('Lowest Decrease:'),
          ...makeItems(data.lowestDecrease)
        ];
      };

      const doc = new Document({
        sections: [
          {
            children: [
              new Paragraph({
                text: 'CPD PRICE TRACKER',
                heading: HeadingLevel.HEADING_1,
                alignment: AlignmentType.CENTER
              }),
              new Paragraph({
                text: `${this.provinceName} - ${this.reportTitle || 'Report'}`,
                heading: HeadingLevel.HEADING_2,
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 }
              }),
              new Paragraph({
                text: `Date: ${this.reportDate}`,
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 }
              }),

              ...makeSection('WEEK SUMMARY', this.summaryData.week),
              ...makeSection('MONTH SUMMARY', this.summaryData.month),
              ...makeSection('3 MONTHS SUMMARY', this.summaryData.threeMonth)
            ]
          }
        ]
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
}