import { Component } from '@angular/core';
import { IonicModule, ToastController } from '@ionic/angular';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { addIcons } from 'ionicons';
import {
  cloudUploadOutline,
  searchOutline,
  micOutline
} from 'ionicons/icons';
import * as XLSX from 'xlsx';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, UnderlineType } from 'docx';
import { saveAs } from 'file-saver';

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

@Component({
  selector: 'app-quirino',
  templateUrl: './quirino.page.html',
  styleUrls: ['./quirino.page.scss'],
  standalone: true,
  imports: [IonicModule, CommonModule, FormsModule]
})
export class QuirinoPage {

  selectedFile: File | null = null;
  isGenerated = false;
  isComparativeMode = false;
  showWarning = false;
  warningMessage = '';

  workbook: XLSX.WorkBook | null = null;
  sheetNames: string[] = [];
  selectedSheetIndex = 0;

  tableHeaders: string[] = [];
  fullTableHeaders: string[] = [];
  tableData: any[][] = [];
  fullTableData: any[][] = [];

  // Search properties
  searchQuery: string = '';
  filteredTableData: any[][] = [];

  showSummary = false;

  summary = {
    month1: {
      increaseCount: 0,
      decreaseCount: 0,
      highestIncrease: [] as string[],
      lowestIncrease: [] as string[],
      highestDecrease: [] as string[],
      lowestDecrease: [] as string[],
      totalProducts: 0,
    },
    month3: {
      increaseCount: 0,
      decreaseCount: 0,
      highestIncrease: [] as string[],
      lowestIncrease: [] as string[],
      highestDecrease: [] as string[],
      lowestDecrease: [] as string[],
      totalProducts: 0
    }
  };

  provinceName = 'Quirino';
  reportTitle: string = '';

  currentIndex = -1;
  month1Index = -1;
  month3Index = -1;
  productNameIndex = -1;
  unitIndex = -1;

  peso1Index = -1;
  percent1Index = -1;
  peso3Index = -1;
  percent3Index = -1;

  alwaysShowCols: number[] = [];
  highlightColumns: number[] = [];
  visibleColumns: number[] = [];
  originalRowIndices: number[] = [];

  constructor(private toastController: ToastController) {
    addIcons({
      cloudUploadOutline,
      searchOutline,
      micOutline
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
    this.isComparativeMode = false;
    this.showSummary = false;
    this.selectedFile = null;
    this.workbook = null;
    this.sheetNames = [];
    this.selectedSheetIndex = 0;
    this.tableHeaders = [];
    this.fullTableHeaders = [];
    this.tableData = [];
    this.fullTableData = [];
    this.filteredTableData = [];
    this.searchQuery = '';
    this.visibleColumns = [];
    this.originalRowIndices = [];
    this.showWarning = false;
  }

  loadSheet(index: number) {
    if (!this.workbook) return;

    const sheet = this.workbook.Sheets[this.sheetNames[index]];
    const json: any[][] = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      blankrows: false,
      raw: false,
      defval: ''
    });

    // Search for the row that contains the actual headers
    // Looking for "PREVAILING PRICE" in the data rows
    let headerRowIndex = -1;
    for (let i = 0; i < json.length; i++) {
      const rowText = json[i].join('|').toLowerCase();
      if (
        rowText.includes('prevailing price') &&
        (rowText.includes('for the month') || rowText.includes('month ago'))
      ) {
        headerRowIndex = i;
        break;
      }
    }

    // If not found, look for any row with "basic necessities" or similar
    if (headerRowIndex === -1) {
      for (let i = 0; i < json.length; i++) {
        const rowText = json[i].join('|').toLowerCase();
        if (
          rowText.includes('basic necessities') ||
          rowText.includes('prime commodities') ||
          rowText.includes('prevailing price')
        ) {
          headerRowIndex = i;
          break;
        }
      }
    }

    // Default to first row if still not found
    if (headerRowIndex === -1) {
      headerRowIndex = 0;
    }

    // Use the found row as headers
    this.tableHeaders = json[headerRowIndex] || [];
    this.fullTableHeaders = [...this.tableHeaders];

    // Skip header row and any metadata rows, start data from the next row
    this.fullTableData = json
      .slice(headerRowIndex + 1)
      .filter(r => this.isRowNotEmpty(r));

    this.tableData = [...this.fullTableData];
    this.filteredTableData = [...this.fullTableData];
    this.isComparativeMode = false;
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
    // DEBUG: Print all headers
    console.log('=== ALL COLUMN HEADERS ===');
    this.fullTableHeaders.forEach((header, index) => {
      console.log(`Column ${index}: "${header}"`);
    });
    console.log('========================');
    
    // DEBUG: Print first 3 data rows to see structure
    console.log('=== FIRST 3 DATA ROWS ===');
    for (let i = 0; i < Math.min(3, this.fullTableData.length); i++) {
      console.log(`Row ${i}:`, this.fullTableData[i]);
    }
    console.log('========================');

    // Find current month price column - MORE FLEXIBLE
    this.currentIndex = -1;
    for (let i = 0; i < this.fullTableHeaders.length; i++) {
      const header = this.fullTableHeaders[i]?.toLowerCase() || '';
      // Remove line breaks and extra spaces for matching
      const cleanHeader = header.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
      
      if (
        (cleanHeader.includes('prevailing price') && cleanHeader.includes('for the month') && !cleanHeader.includes('month ago')) ||
        (cleanHeader.includes('current') && cleanHeader.includes('price')) ||
        (cleanHeader.includes('price') && cleanHeader.includes('month') && !cleanHeader.includes('ago') && !cleanHeader.includes('previous'))
      ) {
        this.currentIndex = i;
        break;
      }
    }

    // Find 1 month ago column - MORE FLEXIBLE
    this.month1Index = -1;
    for (let i = 0; i < this.fullTableHeaders.length; i++) {
      const header = this.fullTableHeaders[i]?.toLowerCase() || '';
      // Remove line breaks and extra spaces for matching
      const cleanHeader = header.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
      
      if (
        cleanHeader.includes('1 month ago') || 
        cleanHeader.includes('1-month ago') ||
        cleanHeader.includes('previous month') ||
        cleanHeader.includes('last month') ||
        (cleanHeader.includes('prevailing') && cleanHeader.includes('1') && cleanHeader.includes('month') && cleanHeader.includes('ago'))
      ) {
        this.month1Index = i;
        break;
      }
    }

    // Find 3 months ago column - MORE FLEXIBLE
    this.month3Index = -1;
    for (let i = 0; i < this.fullTableHeaders.length; i++) {
      const header = this.fullTableHeaders[i]?.toLowerCase() || '';
      // Remove line breaks and extra spaces for matching
      const cleanHeader = header.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
      
      if (
        cleanHeader.includes('3 months ago') || 
        cleanHeader.includes('3-months ago') ||
        (cleanHeader.includes('prevailing') && cleanHeader.includes('3') && cleanHeader.includes('month') && cleanHeader.includes('ago'))
      ) {
        this.month3Index = i;
        break;
      }
    }

    // Find product name column - MORE FLEXIBLE
    this.productNameIndex = -1;
    for (let i = 0; i < this.fullTableHeaders.length; i++) {
      const header = this.fullTableHeaders[i]?.toLowerCase() || '';
      if (
        header.includes('basic necessities') ||
        header.includes('prime commodities') ||
        header.includes('construction materials') ||
        header.includes('commodity') ||
        header.includes('product') ||
        header.includes('item') ||
        header.includes('description') ||
        header.includes('particulars')
      ) {
        this.productNameIndex = i;
        break;
      }
    }

    // If still not found, assume first column is product name
    if (this.productNameIndex === -1) {
      this.productNameIndex = 0;
      console.log('Product name column not found, defaulting to column 0');
    }

    // Find unit column - MORE FLEXIBLE
    this.unitIndex = -1;

    // First try: exact match for "unit" or "units"
    for (let i = 0; i < this.fullTableHeaders.length; i++) {
      if (!this.fullTableHeaders[i]) continue;
      
      const header = this.fullTableHeaders[i].toString().toLowerCase().trim();
      
      if (header === 'unit' || header === 'units' || header === 'unit of measure') {
        this.unitIndex = i;
        break;
      }
    }

    // Second try: look for unit-related keywords
    if (this.unitIndex === -1) {
      for (let i = 0; i < this.fullTableHeaders.length; i++) {
        const header = this.fullTableHeaders[i]?.toLowerCase() || '';
        if (header.includes('unit') || header.includes('measure') || header.includes('uom')) {
          this.unitIndex = i;
          break;
        }
      }
    }

    // Third try: look between product name and current price for unit-like column
    if (this.unitIndex === -1 && this.productNameIndex !== -1 && this.currentIndex !== -1) {
      for (let i = this.productNameIndex + 1; i < this.currentIndex; i++) {
        if (this.looksLikeUnitColumn(i)) {
          this.unitIndex = i;
          break;
        }
      }
    }

    // Fourth try: assume column right after product name
    if (this.unitIndex === -1 && this.productNameIndex !== -1) {
      const potentialUnitIndex = this.productNameIndex + 1;
      if (potentialUnitIndex < this.fullTableHeaders.length) {
        this.unitIndex = potentialUnitIndex;
        console.log('Unit column not found, defaulting to column after product name:', potentialUnitIndex);
      }
    }

    // Calculate peso and percent columns based on found comparison columns
    if (this.month1Index !== -1) {
      this.peso1Index = this.month1Index + 1;
      this.percent1Index = this.month1Index + 2;
    }

    if (this.month3Index !== -1) {
      this.peso3Index = this.month3Index + 1;
      this.percent3Index = this.month3Index + 2;
    }

    // Find remarks column
    const remarksIndex = this.fullTableHeaders.findIndex(h =>
      h?.toLowerCase().includes('remarks') || h?.toLowerCase().includes('notes')
    );

    // Build always show columns array
    this.alwaysShowCols = [];
    
    if (this.productNameIndex !== -1) {
      this.alwaysShowCols.push(this.productNameIndex);
    }
    
    if (this.unitIndex !== -1) {
      this.alwaysShowCols.push(this.unitIndex);
    }
    
    if (this.currentIndex !== -1) {
      this.alwaysShowCols.push(this.currentIndex);
    }
    
    if (remarksIndex !== -1) {
      this.alwaysShowCols.push(remarksIndex);
    }

    // Build highlight columns array
    this.highlightColumns = [];
    
    if (this.month1Index !== -1) this.highlightColumns.push(this.month1Index);
    if (this.peso1Index !== -1) this.highlightColumns.push(this.peso1Index);
    if (this.percent1Index !== -1) this.highlightColumns.push(this.percent1Index);
    if (this.month3Index !== -1) this.highlightColumns.push(this.month3Index);
    if (this.peso3Index !== -1) this.highlightColumns.push(this.peso3Index);
    if (this.percent3Index !== -1) this.highlightColumns.push(this.percent3Index);

    // Debug log to help identify issues
    console.log('Column Detection Results:', {
      currentIndex: this.currentIndex,
      month1Index: this.month1Index,
      month3Index: this.month3Index,
      productNameIndex: this.productNameIndex,
      unitIndex: this.unitIndex,
      peso1Index: this.peso1Index,
      percent1Index: this.percent1Index,
      peso3Index: this.peso3Index,
      percent3Index: this.percent3Index
    });

    // Validation warning
    if (this.currentIndex === -1 || this.month1Index === -1) {
      console.warn('⚠️ WARNING: Required columns not detected! Comparative analysis may not work.');
      console.warn('Missing columns:', {
        currentPrice: this.currentIndex === -1,
        month1Comparison: this.month1Index === -1,
        month3Comparison: this.month3Index === -1
      });
    }
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

  // Search functionality
  onSearchChange() {
    const query = this.searchQuery.toLowerCase().trim();

    if (!query) {
      this.applyCurrentView();
      return;
    }

    const filteredWithIndices = this.fullTableData
      .map((row, index) => ({ row, originalIndex: index }))
      .filter(({ row }) => {
        return row.some(cell => {
          if (!cell) return false;
          return cell.toString().toLowerCase().includes(query);
        });
      });

    this.filteredTableData = filteredWithIndices.map(item => item.row);
    this.originalRowIndices = filteredWithIndices.map(item => item.originalIndex);

    if (this.isComparativeMode) {
      this.tableData = this.filteredTableData.map(row =>
        this.visibleColumns.map(colIndex => row[colIndex] || '')
      );
    } else {
      this.tableData = this.filteredTableData;
    }
  }

  clearSearch() {
    this.searchQuery = '';
    this.onSearchChange();
  }

  applyCurrentView() {
    if (this.isComparativeMode) {
      const columnsToShow = [
        ...this.alwaysShowCols,
        ...this.highlightColumns
      ];

      this.visibleColumns = [...new Set(columnsToShow)].sort((a, b) => a - b);

      // Store original indices
      this.originalRowIndices = this.fullTableData.map((_, index) => index);

      this.tableData = this.fullTableData.map(row =>
        this.visibleColumns.map(colIndex => row[colIndex] || '')
      );

      this.tableHeaders = this.visibleColumns.map(
        colIndex => this.fullTableHeaders[colIndex] || ''
      );
    } else {
      this.tableData = [...this.fullTableData];
      this.tableHeaders = [...this.fullTableHeaders];
      this.originalRowIndices = [];
    }

    this.filteredTableData = [...this.fullTableData];
  }

  runComparativeAnalysis() {
    this.toggleComparativeAnalysis();
  }

  generateSummaryReport() {
    this.showSummary = !this.showSummary;
    if (!this.showSummary) return;

    this.summary.month1 = this.calculateMonthSummary(this.month1Index);
    this.summary.month3 = this.calculateMonthSummary(this.month3Index);
  }

  calculateMonthSummary(compareIndex: number): MonthSummary {
    const increase: ComparisonItem[] = [];
    const decrease: ComparisonItem[] = [];
    let totalProducts = 0;

    for (const row of this.fullTableData) {
      const current = this.parseNumber(row[this.currentIndex]);
      const oldPrice = this.parseNumber(row[compareIndex]);

      // Count all products that have valid product names
      const productName = row[this.productNameIndex];
      if (productName && productName.toString().trim() !== '') {
        const trimmedName = productName.toString().trim().toLowerCase();
        
        // Exclude category headers
        const isCategoryHeader = 
          trimmedName.includes('basic necessities') ||
          trimmedName.includes('prime commodities') ||
          trimmedName.includes('construction materials') ||
          trimmedName === 'basic necessities and prime commodities' ||
          trimmedName === 'construction materials';
        
        if (isCategoryHeader) {
          continue; // Skip this row, don't count it
        }
        
        // Check if at least one price column has a valid value
        const month1Price = this.month1Index !== -1 ? this.parseNumber(row[this.month1Index]) : null;
        const month3Price = this.month3Index !== -1 ? this.parseNumber(row[this.month3Index]) : null;
        
        const hasAnyPrice = current !== null || 
                           oldPrice !== null ||
                           month1Price !== null ||
                           month3Price !== null;
        
        if (hasAnyPrice) {
          totalProducts++;
        }
      }

      if (!current || !oldPrice) continue;

      const name = row[this.productNameIndex] || 'Unknown';
      let unit = '';

      if (this.unitIndex !== -1) unit = row[this.unitIndex];
      if (!unit && this.productNameIndex !== -1) {
        unit = row[this.productNameIndex + 1] || '';
      }

      unit = unit ? unit.toString().trim() : 'N/A';

      const pesoChange = current - oldPrice;

      if (pesoChange > 0) {
        increase.push({ name, unit, peso: pesoChange });
      } else if (pesoChange < 0) {
        decrease.push({ name, unit, peso: Math.abs(pesoChange) });
      }
    }

    // Sort arrays by peso value
    increase.sort((a, b) => a.peso - b.peso);
    decrease.sort((a, b) => a.peso - b.peso);

    // Helper function to find all items with a specific peso value
    const findAllWithValue = (arr: ComparisonItem[], targetValue: number, tolerance: number = 0.001): string[] => {
      const matchingItems = arr.filter(item => Math.abs(item.peso - targetValue) < tolerance);
      
      return matchingItems.map(item => 
        `₱${item.peso.toFixed(2)} - ${item.name} (${item.unit})`
      );
    };

    // MODIFIED LOGIC: If there's only one item, record it only as lowest, not highest
    let lowestIncreaseItems: string[] = ['N/A'];
    let highestIncreaseItems: string[] = ['N/A'];
    
    if (increase.length === 1) {
      // Only one increase - record as lowest only
      const lowestValue = increase[0].peso;
      lowestIncreaseItems = findAllWithValue(increase, lowestValue);
      highestIncreaseItems = ['N/A'];
    } else if (increase.length > 1) {
      // Multiple increases - record both lowest and highest
      const lowestValue = increase[0].peso;
      lowestIncreaseItems = findAllWithValue(increase, lowestValue);
      
      const highestValue = increase[increase.length - 1].peso;
      highestIncreaseItems = findAllWithValue(increase, highestValue);
    }

    let lowestDecreaseItems: string[] = ['N/A'];
    let highestDecreaseItems: string[] = ['N/A'];
    
    if (decrease.length === 1) {
      // Only one decrease - record as lowest only
      const lowestValue = decrease[0].peso;
      lowestDecreaseItems = findAllWithValue(decrease, lowestValue);
      highestDecreaseItems = ['N/A'];
    } else if (decrease.length > 1) {
      // Multiple decreases - record both lowest and highest
      const lowestValue = decrease[0].peso;
      lowestDecreaseItems = findAllWithValue(decrease, lowestValue);
      
      const highestValue = decrease[decrease.length - 1].peso;
      highestDecreaseItems = findAllWithValue(decrease, highestValue);
    }

    return {
      increaseCount: increase.length,
      decreaseCount: decrease.length,
      highestIncrease: highestIncreaseItems,
      lowestIncrease: lowestIncreaseItems,
      highestDecrease: highestDecreaseItems,
      lowestDecrease: lowestDecreaseItems,
      totalProducts: totalProducts,
    };
  }

  getComparisonClass(rowIndex: number, colIndex: number): string {
    if (!this.isComparativeMode) return '';

    const originalColIndex = this.visibleColumns[colIndex];
    const originalRowIndex = this.originalRowIndices[rowIndex];

    if (originalRowIndex === undefined || originalRowIndex === -1) return '';

    const comp1 = this.getRowComparison(
      this.fullTableData[originalRowIndex],
      this.month1Index
    );

    const comp3 = this.getRowComparison(
      this.fullTableData[originalRowIndex],
      this.month3Index
    );

    if (
      comp1 &&
      (originalColIndex === this.month1Index ||
        originalColIndex === this.peso1Index ||
        originalColIndex === this.percent1Index)
    ) {
      return comp1;
    }

    if (
      comp3 &&
      (originalColIndex === this.month3Index ||
        originalColIndex === this.peso3Index ||
        originalColIndex === this.percent3Index)
    ) {
      return comp3;
    }

    return '';
  }

  getRowComparison(row: any[], compareIndex: number): string {
    if (compareIndex === -1) return '';

    const current = this.parseNumber(row[this.currentIndex]);
    const oldPrice = this.parseNumber(row[compareIndex]);

    if (!current || !oldPrice) return '';

    if (oldPrice > current) return 'decrease';
    if (oldPrice < current) return 'increase';

    return '';
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

  toggleComparativeAnalysis() {
    if (this.showSummary) {
      this.showSummary = false;
    }

    this.isComparativeMode = !this.isComparativeMode;

    if (!this.isComparativeMode) {
      this.tableData = [...this.fullTableData];
      this.tableHeaders = [...this.fullTableHeaders];
      this.visibleColumns = [];
      this.originalRowIndices = [];
    } else {
      const columnsToShow = [
        ...this.alwaysShowCols,
        ...this.highlightColumns
      ];

      this.visibleColumns = [...new Set(columnsToShow)].sort((a, b) => a - b);

      // Store original indices
      this.originalRowIndices = this.fullTableData.map((_, index) => index);

      this.tableData = this.fullTableData.map(row =>
        this.visibleColumns.map(colIndex => row[colIndex] || '')
      );

      this.tableHeaders = this.visibleColumns.map(
        colIndex => this.fullTableHeaders[colIndex] || ''
      );
    }

    if (this.searchQuery) {
      this.onSearchChange();
    }
  }

  async exportSummaryReport() {
    if (!this.showSummary) {
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
      // Get filename without extension
      const baseFilename = this.selectedFile?.name.replace(/\.[^/.]+$/, '') || 'Summary_Report';
      const sheetName = this.sheetNames[this.selectedSheetIndex] || 'Sheet';
      const filename = `${baseFilename}_${sheetName}_Summary.docx`;

      // Helper function to create summary items
      const createSummaryItems = (items: string[]): Paragraph[] => {
        return items.map(item => 
          new Paragraph({
            text: `  • ${item}`,
            spacing: { before: 100, after: 100 }
          })
        );
      };

      // Create document sections
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

            // 1 MONTH COMPARISON SUMMARY
            new Paragraph({
              children: [
              new TextRun({
                text: '1 MONTH COMPARISON SUMMARY',
                bold: true,
                underline: { type: UnderlineType.SINGLE }
              })
            ],
            spacing: { before: 400, after: 300 }
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

            // Total Number of Products (1 Month)
            new Paragraph({
              children: [
                new TextRun({
                  text: `Total Number of Products: ${this.summary.month1.totalProducts}`,
                  bold: true,
                  size: 24
                })
              ],
              spacing: { before: 400, after: 200 }
            }),

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
            ...createSummaryItems(this.summary.month3.lowestDecrease),

            // Total Number of Products (3 Months)
            new Paragraph({
              children: [
                new TextRun({
                  text: `Total Number of Products: ${this.summary.month3.totalProducts}`,
                  bold: true,
                  size: 24
                })
              ],
              spacing: { before: 400, after: 200 }
            })
          ]
        }]
      });

      // Generate and save the document
      const blob = await Packer.toBlob(doc);
      // Give browser time to release file handle
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