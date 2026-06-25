// @ts-ignore
/* global Excel, Office */

export interface SheetData {
  hasData: boolean;
  headers: string[];
  rowCount: number;
  columnCount: number;
  values: any[][];
  formulas: string[][];
  address: string;
  tableName?: string;
  columnIndex?: number;
  rowIndex?: number;
}

export interface WorkbookData {
  activeSheetName: string;
  sheets: (SheetData & { name: string })[];
}

/**
 * Membaca data dari seluruh worksheet di workbook aktif.
 */
export async function getWorkbookData(): Promise<WorkbookData> {
  // @ts-ignore
  const sheetNames = await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items");
    await context.sync();
    return sheets.items.map(s => s.name);
  });
  
  // @ts-ignore
  const activeSheetName = await Excel.run(async (context) => {
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("name");
    await context.sync();
    return activeSheet.name;
  });

  const workbookSheetsData = [];
  
  for (const name of sheetNames) {
    try {
      // @ts-ignore
      const sheetData = await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(name);
        const range = sheet.getUsedRange(true /* valuesOnly */);
        range.load(["values", "formulas", "rowCount", "columnCount", "address", "columnIndex", "rowIndex"]);
        
        const tables = sheet.tables;
        tables.load("items");
        
        await context.sync();
        
        const values = range.values;
        const formulas = range.formulas;
        const rowCount = range.rowCount;
        const columnCount = range.columnCount;
        const columnIndex = range.columnIndex;
        const rowIndex = range.rowIndex;
        
        let tableName: string | undefined;
        if (tables.items.length > 0) {
          tableName = tables.items[0].name;
        }

        const headers = rowCount > 0 ? values[0].map(h => String(h || '').trim()) : [];
        
        return {
          name,
          hasData: true,
          headers,
          rowCount,
          columnCount,
          values,
          formulas,
          address: range.address,
          tableName,
          columnIndex,
          rowIndex
        };
      });
      workbookSheetsData.push(sheetData);
    } catch (err) {
      // Jika error (misal sheet kosong), return data sheet kosong
      workbookSheetsData.push({
        name,
        hasData: false,
        headers: [],
        rowCount: 0,
        columnCount: 0,
        values: [],
        formulas: [],
        address: "A1",
        columnIndex: 0,
        rowIndex: 0
      });
    }
  }
  
  return {
    activeSheetName,
    sheets: workbookSheetsData
  };
}

/**
 * Membaca data dari worksheet yang sedang aktif.
 */
export async function getActiveSheetData(): Promise<SheetData> {
  return new Promise((resolve, reject) => {
    // @ts-ignore
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Ambil used range (rentang sel yang berisi data/formula/format)
      const range = sheet.getUsedRange(true /* valuesOnly - ini mencegah error jika hanya format yang ada */);
      
      // Load properti yang dibutuhkan
      range.load(["values", "formulas", "rowCount", "columnCount", "address", "columnIndex", "rowIndex"]);
      
      // Cek apakah ada tabel resmi di sheet ini
      const tables = sheet.tables;
      tables.load("items");
      
      try {
        await context.sync();
        
        const values = range.values;
        const formulas = range.formulas;
        const rowCount = range.rowCount;
        const columnCount = range.columnCount;
        const columnIndex = range.columnIndex;
        const rowIndex = range.rowIndex;
        
        let tableName: string | undefined;
        if (tables.items.length > 0) {
          tableName = tables.items[0].name;
        }

        const headers = rowCount > 0 ? values[0].map(h => String(h || '').trim()) : [];

        resolve({
          hasData: true,
          headers,
          rowCount,
          columnCount,
          values,
          formulas,
          address: range.address,
          tableName,
          columnIndex,
          rowIndex
        });
      } catch (err) {
        // Jika getUsedRange error, kemungkinan besar sheet kosong
        resolve({
          hasData: false,
          headers: [],
          rowCount: 0,
          columnCount: 0,
          values: [],
          formulas: [],
          address: "A1",
          columnIndex: 0,
          rowIndex: 0
        });
      }
    }).catch(reject);
  });
}

/**
 * Menulis data formula ke dalam range sel tertentu.
 */
export async function writeFormulas(rangeAddress: string, formulas: string[][]): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(rangeAddress);
    range.formulas = formulas;
    range.format.autofitColumns();
    await context.sync();
  });
}

/**
 * Menulis data nilai biasa ke dalam range sel tertentu.
 */
export async function writeValues(rangeAddress: string, values: any[][]): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(rangeAddress);
    range.values = values;
    range.format.autofitColumns();
    await context.sync();
  });
}

/**
 * Mengubah suatu range data menjadi Excel Table resmi agar mempermudah sorting & filtering.
 */
export async function createExcelTable(rangeAddress: string, tableName: string): Promise<string> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // Hilangkan karakter non-alphanumeric untuk nama tabel yang valid
    const cleanTableName = tableName.replace(/[^a-zA-Z0-9]/g, "");
    
    // Cek apakah tabel dengan nama tersebut sudah ada, jika ya hapus dulu
    const tables = sheet.tables;
    tables.load("items");
    await context.sync();
    
    const existingTable = tables.items.find(t => t.name.toLowerCase() === cleanTableName.toLowerCase());
    if (existingTable) {
      existingTable.delete();
      await context.sync();
    }

    const table = sheet.tables.add(rangeAddress, true /* hasHeaders */);
    table.name = cleanTableName;
    table.showHeaders = true;
    table.showFilterButton = true;
    table.getRange().format.autofitColumns();
    
    await context.sync();
    return cleanTableName;
  });
}

/**
 * Mengurutkan kolom pada tabel.
 */
export async function sortTableColumn(tableName: string, columnName: string, direction: 'asc' | 'desc'): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItem(tableName);
    table.load("columns");
    await context.sync();
    
    const columns = table.columns;
    const colIndex = columns.items.findIndex(c => c.name.toLowerCase() === columnName.toLowerCase());
    
    if (colIndex === -1) {
      throw new Error(`Kolom '${columnName}' tidak ditemukan di tabel '${tableName}'`);
    }

    const sortField = {
      key: colIndex,
      ascending: direction === 'asc'
    };

    table.sort.apply([sortField], true /* matchCase */);
    await context.sync();
  });
}

/**
 * Menerapkan filter pada kolom tertentu di tabel.
 */
export async function filterTableColumn(
  tableName: string, 
  columnName: string, 
  value: string, 
  operator: 'Equal' | 'Contains' | 'GreaterThan' | 'LessThan' | 'Clear' = 'Equal'
): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItem(tableName);
    table.load("columns");
    await context.sync();
    
    const col = table.columns.items.find(c => c.name.toLowerCase() === columnName.toLowerCase());
    if (!col) {
      throw new Error(`Kolom '${columnName}' tidak ditemukan di tabel '${tableName}'`);
    }

    const filter = col.filter;
    
    if (operator === 'Clear') {
      filter.clear();
    } else if (operator === 'Contains') {
      // Custom filter untuk text containing
      filter.apply({
        filterOn: Excel.FilterOn.custom,
        criterion1: `=*${value}*`
      });
    } else if (operator === 'GreaterThan') {
      filter.apply({
        filterOn: Excel.FilterOn.custom,
        criterion1: `>${value}`
      });
    } else if (operator === 'LessThan') {
      filter.apply({
        filterOn: Excel.FilterOn.custom,
        criterion1: `<${value}`
      });
    } else {
      // Equal
      filter.apply({
        filterOn: Excel.FilterOn.values,
        values: [value]
      });
    }

    await context.sync();
  });
}

/**
 * Menyisipkan Grafik (Chart) Excel asli.
 */
export async function createExcelChart(
  chartTypeStr: string, 
  dataRangeAddress: string, 
  title: string
): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // Map string tipe grafik ke enum Excel.ChartType
    let chartType = Excel.ChartType.columnClustered; // Default
    const typeLower = (chartTypeStr || 'column').toLowerCase();
    
    if (typeLower.includes('line')) {
      chartType = Excel.ChartType.line;
    } else if (typeLower.includes('pie')) {
      chartType = Excel.ChartType.pie;
    } else if (typeLower.includes('bar')) {
      chartType = Excel.ChartType.barClustered;
    } else if (typeLower.includes('area')) {
      chartType = Excel.ChartType.area;
    }

    let dataRange: Excel.Range;
    if (dataRangeAddress) {
      dataRange = sheet.getRange(dataRangeAddress);
    } else {
      // Ambil used range secara otomatis jika dataRangeAddress kosong
      dataRange = sheet.getUsedRange();
    }
    
    const chart = sheet.charts.add(chartType, dataRange, Excel.ChartSeriesBy.auto);
    
    chart.title.text = title || "Grafik Data AI";
    chart.title.visible = true;
    
    // Posisikan grafik di sebelah kanan data (misal bergeser 2 kolom ke kanan)
    dataRange.load(["columnCount", "columnIndex"]);
    await context.sync();
    
    const targetCellIndex = dataRange.columnIndex + dataRange.columnCount + 1;
    const targetCell = sheet.getCell(0, targetCellIndex);
    targetCell.load(["left", "top"]);
    await context.sync();
    
    chart.left = targetCell.left;
    chart.top = 20;
    chart.height = 300;
    chart.width = 450;

    await context.sync();
  });
}

/**
 * Membuat Pivot Table dari range/tabel data.
 */
export async function createExcelPivotTable(
  sourceRangeAddressOrTable: string, 
  rowFields: string[], 
  dataFields: string[]
): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const workbook = context.workbook;
    const sheets = workbook.worksheets;
    sheets.load("items");
    await context.sync();

    // Buat worksheet baru khusus untuk Pivot Table
    const pivotSheetName = `Pivot_${Date.now().toString().slice(-4)}`;
    const pivotSheet = sheets.add(pivotSheetName);
    
    const sourceSheet = sheets.getActiveWorksheet();
    let sourceRange: Excel.Range;
    
    if (sourceRangeAddressOrTable.startsWith("Table") || !sourceRangeAddressOrTable.includes("!")) {
      // Jika nama tabel atau range di sheet aktif
      sourceRange = sourceSheet.getRange(sourceRangeAddressOrTable);
    } else {
      // Jika range lengkap beserta nama sheetnya (contoh: Sheet1!A1:B10)
      const parts = sourceRangeAddressOrTable.split("!");
      const sheetName = parts[0].replace(/'/g, "");
      const rangeAddress = parts[1];
      sourceRange = workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    }

    // Buat Pivot Table
    const destinationRange = pivotSheet.getRange("A3");
    const pivotTable = pivotSheet.pivotTables.add(
      `Pivot_${Date.now().toString().slice(-4)}`, 
      sourceRange, 
      destinationRange
    );

    // Tambahkan Row Fields (Dimensi baris)
    rowFields.forEach(field => {
      pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(field));
    });

    // Tambahkan Data Fields (Nilai ringkasan sum)
    dataFields.forEach(field => {
      const dataHierarchy = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(field));
      dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
    });

    pivotSheet.activate();
    await context.sync();
  });
}
