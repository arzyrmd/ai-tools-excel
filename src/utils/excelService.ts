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
      if (dataRangeAddress.includes(',')) {
        // Rentang non-kontigu terdeteksi (contoh: Sheet1!B1:B11,Sheet1!F1:F11)
        // Kita hitung bounding range kontigu yang melingkupi semua area
        const areas = dataRangeAddress.split(',');
        let minCol = Infinity;
        let maxCol = -Infinity;
        let minRow = Infinity;
        let maxRow = -Infinity;
        let sheetPrefix = "";

        const firstArea = areas[0].trim();
        if (firstArea.includes('!')) {
          sheetPrefix = firstArea.split('!')[0] + '!';
        }

        for (const area of areas) {
          const trimmedArea = area.trim();
          const cleanArea = trimmedArea.includes('!') ? trimmedArea.split('!')[1] : trimmedArea;
          const parts = cleanArea.split(':');
          const start = parts[0];
          const end = parts[1] || start;

          // Helper untuk memparsing cell seperti "B1" atau "AA12"
          const parseCell = (cellStr: string) => {
            const match = cellStr.match(/^([A-Z]+)([0-9]+)$/i);
            if (match) {
              const colLetters = match[1].toUpperCase();
              const rowNum = parseInt(match[2], 10);
              
              let colIdx = 0;
              for (let i = 0; i < colLetters.length; i++) {
                colIdx = colIdx * 26 + (colLetters.charCodeAt(i) - 64);
              }
              colIdx = colIdx - 1; // 0-indexed
              return { colIdx, rowNum };
            }
            return null;
          };

          const startCoord = parseCell(start);
          const endCoord = parseCell(end);

          if (startCoord && endCoord) {
            minCol = Math.min(minCol, startCoord.colIdx, endCoord.colIdx);
            maxCol = Math.max(maxCol, startCoord.colIdx, endCoord.colIdx);
            minRow = Math.min(minRow, startCoord.rowNum, endCoord.rowNum);
            maxRow = Math.max(maxRow, startCoord.rowNum, endCoord.rowNum);
          }
        }

        if (minCol !== Infinity) {
          const getColLetter = (idx: number): string => {
            let letter = '';
            let temp = idx;
            while (temp >= 0) {
              letter = String.fromCharCode((temp % 26) + 65) + letter;
              temp = Math.floor(temp / 26) - 1;
            }
            return letter;
          };

          const boundingAddress = `${sheetPrefix}${getColLetter(minCol)}${minRow}:${getColLetter(maxCol)}${maxRow}`;
          console.log(`Mengonversi range non-kontigu "${dataRangeAddress}" menjadi bounding range "${boundingAddress}"`);
          dataRange = sheet.getRange(boundingAddress);
        } else {
          dataRange = sheet.getRange(areas[0].trim());
        }
      } else {
        dataRange = sheet.getRange(dataRangeAddress);
      }
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

export interface CellFormatOptions {
  fillColor?: string;
  fontColor?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  fontName?: string;
  horizontalAlignment?: 'Left' | 'Center' | 'Right' | 'General';
  verticalAlignment?: 'Top' | 'Center' | 'Bottom';
  borderStyle?: 'None' | 'Thin' | 'DoubleBottom' | 'AllBorders' | 'HeaderBorders' | 'TotalBorders';
}

/**
 * Memformat format angka (currency, percent, date, dll) pada range sel tertentu.
 */
export async function formatNumbers(rangeAddress: string, formatCode: string): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(rangeAddress);
    
    // Map format code yang umum
    let resolvedFormat = formatCode;
    const lowerFormat = formatCode.toLowerCase();
    
    if (lowerFormat === 'rupiah' || lowerFormat === 'rp' || lowerFormat === 'idr' || lowerFormat.includes('rp')) {
      resolvedFormat = '"Rp"#,##0';
    } else if (lowerFormat === 'usd' || lowerFormat === 'dollar' || lowerFormat === '$') {
      resolvedFormat = '"$"#,##0';
    } else if (lowerFormat === 'percent' || lowerFormat === 'persen' || lowerFormat === '%') {
      resolvedFormat = '0.00%';
    } else if (lowerFormat === 'date' || lowerFormat === 'tanggal') {
      resolvedFormat = 'yyyy-mm-dd';
    } else if (lowerFormat === 'number' || lowerFormat === 'angka') {
      resolvedFormat = '#,##0';
    } else if (lowerFormat === 'decimal' || lowerFormat === 'desimal') {
      resolvedFormat = '#,##0.00';
    }
    
    range.numberFormat = [[resolvedFormat]];
    await context.sync();
  });
}

/**
 * Memformat gaya sel (warna, font, garis tepi, perataan) pada range sel tertentu.
 */
export async function formatCells(rangeAddress: string, options: CellFormatOptions): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(rangeAddress);
    const format = range.format;
    
    if (options.fillColor) {
      format.fill.color = options.fillColor;
    }
    if (options.fontColor) {
      format.font.color = options.fontColor;
    }
    if (options.fontSize) {
      format.font.size = options.fontSize;
    }
    if (options.bold !== undefined) {
      format.font.bold = options.bold;
    }
    if (options.italic !== undefined) {
      format.font.italic = options.italic;
    }
    if (options.fontName) {
      format.font.name = options.fontName;
    }
    if (options.horizontalAlignment) {
      format.horizontalAlignment = options.horizontalAlignment;
    }
    if (options.verticalAlignment) {
      format.verticalAlignment = options.verticalAlignment;
    }
    
    if (options.borderStyle) {
      const borders = format.borders;
      if (options.borderStyle === 'None') {
        borders.load('items');
        await context.sync();
        borders.items.forEach(border => {
          border.style = 'None';
        });
      } else if (options.borderStyle === 'Thin' || options.borderStyle === 'AllBorders') {
        const borderIndices = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'InsideHorizontal', 'InsideVertical'] as const;
        borderIndices.forEach(index => {
          const b = borders.getItem(index);
          b.style = 'Continuous';
          b.weight = 'Thin';
          b.color = '#cbd5e1'; // slate-300
        });
      } else if (options.borderStyle === 'DoubleBottom' || options.borderStyle === 'TotalBorders') {
        const topBorder = borders.getItem('EdgeTop');
        topBorder.style = 'Continuous';
        topBorder.weight = 'Thin';
        topBorder.color = '#cbd5e1';
        
        const bottomBorder = borders.getItem('EdgeBottom');
        bottomBorder.style = 'Double';
        bottomBorder.weight = 'Medium';
        bottomBorder.color = '#475569';
      } else if (options.borderStyle === 'HeaderBorders') {
        const bottomBorder = borders.getItem('EdgeBottom');
        bottomBorder.style = 'Continuous';
        bottomBorder.weight = 'Medium';
        bottomBorder.color = '#1e293b';
      }
    }
    await context.sync();
  });
}

/**
 * Menyisipkan sel/baris/kolom baru.
 */
export async function insertRange(rangeAddress: string, shift: 'Down' | 'Right'): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(rangeAddress);
    const shiftDirection = shift === 'Down' ? Excel.InsertShiftDirection.down : Excel.InsertShiftDirection.right;
    range.insert(shiftDirection);
    await context.sync();
  });
}

/**
 * Menghapus sel/baris/kolom.
 */
export async function deleteRange(rangeAddress: string, shift: 'Up' | 'Left'): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(rangeAddress);
    const shiftDirection = shift === 'Up' ? Excel.DeleteShiftDirection.up : Excel.DeleteShiftDirection.left;
    range.delete(shiftDirection);
    await context.sync();
  });
}

/**
 * Menghapus format/konten di range tertentu.
 */
export async function clearRange(rangeAddress: string, option: 'All' | 'Formats' | 'Contents'): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(rangeAddress);
    if (option === 'Formats') {
      range.clear(Excel.ClearApplyTo.formats);
    } else if (option === 'Contents') {
      range.clear(Excel.ClearApplyTo.contents);
    } else {
      range.clear(Excel.ClearApplyTo.all);
    }
    await context.sync();
  });
}

/**
 * Menggabungkan atau memisahkan sel.
 */
export async function mergeCells(rangeAddress: string, merge: boolean): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(rangeAddress);
    if (merge) {
      range.merge();
    } else {
      range.unmerge();
    }
    await context.sync();
  });
}

/**
 * Menyesuaikan lebar kolom/tinggi baris otomatis.
 */
export async function autofitRange(rangeAddress: string, option: 'columns' | 'rows' | 'both'): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(rangeAddress);
    if (option === 'columns' || option === 'both') {
      range.format.autofitColumns();
    }
    if (option === 'rows' || option === 'both') {
      range.format.autofitRows();
    }
    await context.sync();
  });
}

/**
 * Menyetel opsi tampilan sheet (Gridlines, Headings).
 */
export async function setSheetOptions(gridlines?: boolean, headings?: boolean): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
     if (gridlines !== undefined) {
      sheet.showGridlines = gridlines;
    }
    if (headings !== undefined) {
      sheet.showHeadings = headings;
    }
    await context.sync();
  });
}

/**
 * Membekukan baris/kolom (Freeze Panes).
 */
export async function freezePanes(rows?: number, columns?: number, unfreeze?: boolean): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    if (unfreeze) {
      sheet.freezePanes.unfreeze();
    } else if (rows !== undefined && columns !== undefined) {
      const cell = sheet.getCell(rows, columns);
      sheet.freezePanes.freezeAt(cell);
    } else if (rows !== undefined) {
      sheet.freezePanes.freezeRows(rows);
    } else if (columns !== undefined) {
      sheet.freezePanes.freezeColumns(columns);
    }
    await context.sync();
  });
}

/**
 * Memicu kalkulasi ulang workbook secara penuh.
 */
export async function calculateWorkbook(): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    context.application.calculate("FullRebuild");
    await context.sync();
  });
}

/**
 * Menghapus baris duplikat berdasarkan kolom tertentu.
 */
export async function removeDuplicates(rangeAddress: string, columns: number[], hasHeaders: boolean): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(rangeAddress);
    range.removeDuplicates(columns, hasHeaders);
    await context.sync();
  });
}

/**
 * Melindungi atau membuka proteksi sheet.
 */
export async function protectSheet(protect: boolean, password?: string): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
     const sheet = context.workbook.worksheets.getActiveWorksheet();
    if (protect) {
      sheet.protection.protect(undefined, password);
    } else {
      sheet.protection.unprotect(password);
    }
    await context.sync();
  });
}

/**
 * Mengelola sheet (add, delete, rename, activate).
 */
export async function manageSheet(
  action: 'add' | 'delete' | 'rename' | 'activate', 
  name: string, 
  newName?: string
): Promise<void> {
  // @ts-ignore
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    if (action === 'add') {
      sheets.add(name);
    } else {
      const sheet = sheets.getItem(name);
      if (action === 'delete') {
        sheet.delete();
      } else if (action === 'rename' && newName) {
        sheet.name = newName;
      } else if (action === 'activate') {
        sheet.activate();
      }
    }
    await context.sync();
  });
}
