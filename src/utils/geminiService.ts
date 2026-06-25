import { GoogleGenerativeAI } from '@google/generative-ai';
import { SheetData, WorkbookData } from './excelService';

export interface ActionPlan {
  explanation: string;
  actions: {
    type: 'CREATE_TABLE' | 'WRITE_FORMULAS' | 'WRITE_VALUES' | 'SORT' | 'FILTER' | 'CREATE_CHART' | 'CREATE_PIVOT';
    payload: any;
  }[];
}

/**
 * Mendapatkan daftar kolom alfabet dari indeks (0 -> A, 1 -> B, ..., 25 -> Z, 26 -> AA)
 */
function getColumnLetter(colIndex: number): string {
  let letter = '';
  let temp = colIndex;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

/**
 * Memformat seluruh sheet dalam workbook agar bisa dibaca AI, memungkinkannya melakukan lookup (VLOOKUP/HLOOKUP).
 */
function formatWorkbookDataForAI(workbookData: WorkbookData): string {
  const activeName = workbookData.activeSheetName;
  let formatted = `Nama Sheet Aktif Saat Ini: "${activeName}"\n\n`;
  
  workbookData.sheets.forEach(sheet => {
    formatted += `--- SHEET: "${sheet.name}" ${sheet.name === activeName ? '(AKTIF)' : ''} ---\n`;
    if (!sheet.hasData) {
      formatted += `Status: Lembar kerja kosong.\n\n`;
      return;
    }
    
    const colOffset = sheet.columnIndex || 0;
    const rowOffset = sheet.rowIndex || 0;
    formatted += `Lokasi Rentang: ${sheet.address}\n`;
    formatted += `Nama Tabel Terdeteksi: ${sheet.tableName || 'Tidak ada (Bukan tabel resmi)'}\n`;
    formatted += `Jumlah Baris Data: ${sheet.rowCount}\n`;
    formatted += `Jumlah Kolom: ${sheet.columnCount}\n`;
    
    // Format struktur datanya
    formatted += `Struktur & Isi Data:\n`;
    const totalRows = sheet.rowCount;
    const maxRowsToSend = 30; // Batasi per sheet agar tidak berlebihan jika sheetnya banyak
    
    if (totalRows <= maxRowsToSend) {
      for (let i = 0; i < totalRows; i++) {
        const rowNum = rowOffset + i + 1;
        const rowData = sheet.values[i] || [];
        const cellsStr = rowData.map((cell, idx) => {
          const colLetter = getColumnLetter(idx + colOffset);
          return `${colLetter}: ${cell === "" || cell === null || cell === undefined ? '""' : JSON.stringify(cell)}`;
        }).join(', ');
        formatted += `  Baris ${rowNum} -> { ${cellsStr} }\n`;
      }
    } else {
      // Kirim 15 baris pertama
      for (let i = 0; i < 15; i++) {
        const rowNum = rowOffset + i + 1;
        const rowData = sheet.values[i] || [];
        const cellsStr = rowData.map((cell, idx) => {
          const colLetter = getColumnLetter(idx + colOffset);
          return `${colLetter}: ${cell === "" || cell === null || cell === undefined ? '""' : JSON.stringify(cell)}`;
        }).join(', ');
        formatted += `  Baris ${rowNum} -> { ${cellsStr} }\n`;
      }
      
      formatted += `  ... [Baris ${rowOffset + 16} s.d. ${rowOffset + totalRows - 10} disembunyikan] ...\n`;
      
      // Kirim 10 baris terakhir
      for (let i = totalRows - 10; i < totalRows; i++) {
        const rowNum = rowOffset + i + 1;
        const rowData = sheet.values[i] || [];
        const cellsStr = rowData.map((cell, idx) => {
          const colLetter = getColumnLetter(idx + colOffset);
          return `${colLetter}: ${cell === "" || cell === null || cell === undefined ? '""' : JSON.stringify(cell)}`;
        }).join(', ');
        formatted += `  Baris ${rowNum} -> { ${cellsStr} }\n`;
      }
    }
    formatted += `\n`;
  });
  
  return formatted;
}

/**
 * Melakukan simulasi pemrosesan perintah lokal (Demo Mode) jika API Key tidak diset.
 */
export function getDemoResponse(prompt: string, sheetData: SheetData): ActionPlan {
  const query = prompt.toLowerCase();
  
  // Jika lembar kerja kosong, berikan data sampel terlebih dahulu
  if (!sheetData.hasData) {
    return {
      explanation: "Saya mendeteksi lembar kerja kosong. Pertama, saya akan mengisikan data penjualan sampel agar kita bisa mencoba operasi rumus, filter, dan grafik.",
      actions: [
        {
          type: 'WRITE_VALUES',
          payload: {
            range: "A1:E7",
            values: [
              ["Tanggal", "Produk", "Kategori", "Harga", "Jumlah"],
              ["2026-06-10", "Laptop Pro", "Elektronik", 12000000, 3],
              ["2026-06-11", "Meja Kerja", "Furnitur", 1500000, 10],
              ["2026-06-12", "Headset Wireless", "Elektronik", 800000, 15],
              ["2026-06-13", "Kursi Ergonomis", "Furnitur", 2500000, 8],
              ["2026-06-14", "Keyboard Mechanical", "Elektronik", 1200000, 5],
              ["2026-06-15", "Kabel HDMI", "Elektronik", 150000, 50]
            ]
          }
        },
        {
          type: 'CREATE_TABLE',
          payload: {
            range: "A1:E7",
            name: "SampelPenjualan"
          }
        }
      ]
    };
  }

  const tableName = sheetData.tableName || "Table1";
  const rowCount = sheetData.rowCount;
  
  // 1. PIVOT TABLE DEMO
  if (query.includes('pivot') || query.includes('ringkasan') || query.includes('rekapitulasi')) {
    // Cari kolom kategori/produk (string) dan harga/jumlah (number)
    const catColIndex = sheetData.headers.findIndex(h => h.toLowerCase().includes('kategori') || h.toLowerCase().includes('domisili') || h.toLowerCase().includes('produk'));
    const valColIndex = sheetData.headers.findIndex(h => h.toLowerCase().includes('harga') || h.toLowerCase().includes('jumlah') || h.toLowerCase().includes('pengiriman') || h.toLowerCase().includes('total'));
    
    const catField = catColIndex !== -1 ? sheetData.headers[catColIndex] : sheetData.headers[0];
    const valField = valColIndex !== -1 ? sheetData.headers[valColIndex] : sheetData.headers[sheetData.headers.length - 1];

    return {
      explanation: `Saya akan membuat **Pivot Table** di tab sheet baru untuk merangkum data berdasarkan kolom **${catField}** dan menghitung jumlah total pada kolom **${valField}**.`,
      actions: [
        {
          type: 'CREATE_PIVOT',
          payload: {
            source: sheetData.address,
            rowFields: [catField],
            dataFields: [valField]
          }
        }
      ]
    };
  }

  if (query.includes('grafik') || query.includes('diagram') || query.includes('chart')) {
    return {
      explanation: "Saya akan menyisipkan **Diagram Batang (Column Chart)** di sebelah kanan data untuk memvisualisasikan Kategori produk terhadap Jumlah barang.",
      actions: [
        {
          type: 'CREATE_CHART',
          payload: {
            type: 'bar',
            range: `${sheetData.address.split('!')[0]}!C1:E${rowCount}`,
            title: "Visualisasi Penjualan Barang"
          }
        }
      ]
    };
  }

  // 3. FILTER DEMO
  if (query.includes('filter') || query.includes('tampilkan hanya') || query.includes('cari')) {
    // Cari kolom Domisili atau Kategori
    const domisiliIndex = sheetData.headers.findIndex(h => h.toLowerCase().includes('domisili') || h.toLowerCase().includes('kategori') || h.toLowerCase().includes('customer'));
    const colName = domisiliIndex !== -1 ? sheetData.headers[domisiliIndex] : sheetData.headers[0];
    
    // Cari value filter
    let val = "Jakarta";
    if (query.includes('bandung')) val = "Bandung";
    if (query.includes('surabaya')) val = "Surabaya";
    if (query.includes('elektronik')) val = "Elektronik";
    if (query.includes('furnitur')) val = "Furnitur";
    
    return {
      explanation: `Saya mendeteksi perintah filter. Saya akan memformat area data ini menjadi Excel Table terlebih dahulu (jika belum), kemudian memfilter kolom **${colName}** dengan nilai **"${val}"**.`,
      actions: [
        {
          type: 'CREATE_TABLE',
          payload: {
            range: sheetData.address,
            name: tableName
          }
        },
        {
          type: 'FILTER',
          payload: {
            tableName: tableName,
            columnName: colName,
            value: val,
            operator: 'Equal'
          }
        }
      ]
    };
  }

  // 4. SORT DEMO
  if (query.includes('urut') || query.includes('sort')) {
    const numIndex = sheetData.headers.findIndex(h => h.toLowerCase().includes('harga') || h.toLowerCase().includes('jumlah') || h.toLowerCase().includes('lama') || h.toLowerCase().includes('tanggal'));
    const colName = numIndex !== -1 ? sheetData.headers[numIndex] : sheetData.headers[0];
    const isDesc = query.includes('besar') || query.includes('turun') || query.includes('terlama') || query.includes('desc');
    
    return {
      explanation: `Saya akan menyusun tabel ini dalam format Excel Table, lalu mengurutkan kolom **${colName}** secara **${isDesc ? 'Descending (Terbesar)' : 'Ascending (Terkecil)'}**.`,
      actions: [
        {
          type: 'CREATE_TABLE',
          payload: {
            range: sheetData.address,
            name: tableName
          }
        },
        {
          type: 'SORT',
          payload: {
            tableName: tableName,
            columnName: colName,
            direction: isDesc ? 'desc' : 'asc'
          }
        }
      ]
    };
  }

  // 5. FORMULA DEMO: Buat Kolom Baru / Rumus Penjumlahan
  if (query.includes('rumus') || query.includes('hitung') || query.includes('tambah kolom') || query.includes('perkalian') || query.includes('*') || query.includes('kali')) {
    const nextColIndex = sheetData.columnCount;
    const nextColLetter = getColumnLetter(nextColIndex);
    
    // Cari index Harga dan Jumlah
    const hargaIndex = sheetData.headers.findIndex(h => h.toLowerCase().includes('harga'));
    const jumlahIndex = sheetData.headers.findIndex(h => h.toLowerCase().includes('jumlah'));
    
    if (hargaIndex !== -1 && jumlahIndex !== -1) {
      const hargaLetter = getColumnLetter(hargaIndex);
      const jumlahLetter = getColumnLetter(jumlahIndex);
      
      const formulas: string[][] = [["Total"]];
      for (let i = 2; i <= rowCount; i++) {
        formulas.push([`=${hargaLetter}${i}*${jumlahLetter}${i}`]);
      }

      return {
        explanation: `Saya akan menambahkan kolom baru **Total** di kolom **${nextColLetter}**. Kolom ini berisi rumus perkalian **${hargaLetter} (Harga)** dikali **${jumlahLetter} (Jumlah)** untuk setiap baris.`,
        actions: [
          {
            type: 'WRITE_FORMULAS',
            payload: {
              range: `${nextColLetter}1:${nextColLetter}${rowCount}`,
              formulas: formulas
            }
          }
        ]
      };
    } else {
      // Demo rumus SUM biasa di bawah data
      const lastColLetter = getColumnLetter(sheetData.columnCount - 1);
      return {
        explanation: `Saya akan menambahkan rumus penjumlahan otomatis di bawah baris terakhir pada kolom terakhir (${lastColLetter}).`,
        actions: [
          {
            type: 'WRITE_FORMULAS',
            payload: {
              range: `${lastColLetter}${rowCount + 2}`,
              formulas: [[`=SUM(${lastColLetter}2:${lastColLetter}${rowCount})`]]
            }
          }
        ]
      };
    }
  }

  // DEFAULT FALLBACK DEMO
  return {
    explanation: "Halo! Saya adalah asisten Excel AI Anda. Ketik perintah seperti:\n- Urutkan kolom Lama Pengiriman dari terkecil\n- Filter domisili Bandung\n- Buat kolom total dari perkalian harga dan jumlah\n- Buat diagram batang\n- Buat pivot table",
    actions: []
  };
}

/**
 * Mengirim perintah pengguna dan data workbook ke Gemini API untuk menghasilkan ActionPlan.
 */
export async function getGeminiExcelAction(
  prompt: string, 
  workbookData: WorkbookData, 
  apiKey: string
): Promise<ActionPlan> {
  const activeSheet = workbookData.sheets.find(s => s.name === workbookData.activeSheetName) || workbookData.sheets[0];
  if (!apiKey) {
    return getDemoResponse(prompt, activeSheet);
  }

  const genAI = new GoogleGenerativeAI(apiKey);
  const model = genAI.getGenerativeModel({ 
    model: 'gemini-2.5-flash',
    generationConfig: {
      responseMimeType: 'application/json'
    }
  });

  const workbookStructure = formatWorkbookDataForAI(workbookData);

  const systemInstructions = `
Anda adalah pakar formula dan manipulasi Microsoft Excel tingkat lanjut (Advanced Excel Specialist & AI Agent). Tugas Anda adalah mengonversi perintah bahasa alami pengguna menjadi rencana aksi JSON terstruktur untuk dieksekusi melalui Office.js API secara akurat.

Struktur & Isi Seluruh Lembar Kerja di Workbook (Lengkap):
${workbookStructure}

PENTING UNTUK KECERDASAN DAN AKURASI:
1. REFERENSI ANTAR SHEET (VLOOKUP / HLOOKUP / FORMULA):
   - Pengguna sering kali menyimpan tabel referensi data di sheet terpisah (misalnya sheet "Data Referensi" atau "Sheet2") dan ingin melakukan lookup dari sheet aktif saat ini.
   - Pindai semua sheet di bagian "Struktur & Isi Seluruh Lembar Kerja" untuk mencari data tabel referensi yang cocok.
   - Jika melakukan pencarian lookup ke sheet eksternal, gunakan format penamaan sheet eksternal di dalam rumus (contoh: \`=VLOOKUP(A2, 'Data Referensi'!A2:C10, 2, FALSE)\`). Gunakan petik tunggal (\`'Data Referensi'\`) jika nama sheet mengandung spasi.
2. DETEKSI TEMPAT PENULISAN RUMUS/NILAI (TARGET CELL) SECARA PRESISI:
   - Pengguna sering kali menuliskan teks/label ringkasan di bagian bawah lembar kerja (misalnya di baris 13-20, kolom A/B/C) dan membiarkan sel di sebelahnya (misalnya di kolom E atau F pada baris yang sama) kosong agar diisi rumus oleh AI.
   - Pindai isi sheet aktif untuk mencari apakah sudah ada label ringkasan di bagian bawah yang sesuai dengan perintah pengguna.
   - JANGAN menulis hasil atau rumus di tempat acak (seperti di sebelah kanan di kolom L atau M) apabila di bagian bawah lembar kerja (misalnya baris 13-20) sudah disediakan label baris yang relevan. Tulislah rumus tersebut langsung di sel baris tersebut yang kosong (misalnya jika label "Banyaknya murid yang lulus" ada di A17, tulislah rumusnya langsung di E17).
3. PENDETEKSIAN DATA BARU:
   - Jika pengguna meminta pembuatan data baru yang jenisnya berbeda dari tabel aktif saat ini (misalnya meminta "Data Siswa/UTS" sedangkan sheet aktif berisi "Data Produk/Barang"), atau jika pengguna meminta data baru secara spesifik: 
     * JANGAN menyambung (append) data tersebut di bawah kolom lama.
     * Tulis ulang seluruh sheet aktif mulai dari A1 dengan judul kolom (header) yang baru dan sesuai.
     * Buat data dummy yang sangat kaya, lengkap dengan rumus perhitungan di kolom kanan.
4. PENULISAN FORMULA EXCEL:
   - Selalu tulis Rumus Excel dalam BAHASA INGGRIS (seperti SUM, AVERAGE, COUNT, MIN, MAX, VLOOKUP, HLOOKUP, IF, COUNTIF, SUMIF, dll.) diawali tanda '='. Excel akan menerjemahkannya ke bahasa Excel lokal pengguna secara otomatis.
   - Buat formula yang dinamis merujuk ke baris sel yang tepat. Perhatikan OFFSET baris! Karena baris header berada pada baris ${ (activeSheet.rowIndex || 0) + 1 } dan baris data dimulai dari baris ${ (activeSheet.rowIndex || 0) + 2 }, maka untuk baris data pertama, rujuklah indeks baris ke-${ (activeSheet.rowIndex || 0) + 2 } (contoh: \`=AVERAGE(C${ (activeSheet.rowIndex || 0) + 2 }:E${ (activeSheet.rowIndex || 0) + 2 })\`).
   - Tentukan rentang sel secara presisi. Rentang data utama selalu dimulai dari baris ${ (activeSheet.rowIndex || 0) + 2 } dan berakhir di baris ${ (activeSheet.rowIndex || 0) + activeSheet.rowCount }.

Format Output JSON Harus Tepat Seperti Ini:
{
  "explanation": "Penjelasan langkah demi langkah dalam bahasa Indonesia yang ramah mengenai apa yang akan dilakukan AI.",
  "actions": [
    {
      "type": "WRITE_VALUES",
      "payload": {
        "range": "A1:H1",
        "values": [["No", "Nama Siswa", "Matematika", "Fisika", "Kimia", "Rata-rata", "Status"]]
      }
    },
    {
      "type": "WRITE_FORMULAS",
      "payload": {
        "range": "F2:G11",
        "formulas": [
          ["=AVERAGE(C2:E2)", "=IF(F2>=75,\"LULUS\",\"REMIDI\")"],
          ["=AVERAGE(C3:E3)", "=IF(F3>=75,\"LULUS\",\"REMIDI\")"]
        ]
      }
    },
    {
      "type": "CREATE_TABLE",
      "payload": {
        "range": "A1:G11",
        "name": "DataNilaiSiswa"
      }
    }
  ]
}

Ketentuan Tipe Aksi:
- "CREATE_TABLE": mengonversi range data menjadi objek Tabel Excel tepercaya (mempermudah sort/filter). Membutuhkan "range" (rentang data tabel) dan "name" (nama tabel).
- "WRITE_FORMULAS": menulis array 2D rumus (string[][]) ke range tertentu. Baris header tidak boleh ditulis di sini, gunakan WRITE_VALUES untuk menulis teks/nilai biasa.
- "WRITE_VALUES": menulis nilai mentah biasa (string/number/array 2D) ke range sel tertentu.
- "SORT": mengurutkan kolom pada tabel. Membutuhkan "tableName", "columnName", dan "direction" ('asc' atau 'desc').
- "FILTER": memfilter kolom pada tabel. Membutuhkan "tableName", "columnName", "value", dan "operator" ('Equal' | 'Contains' | 'GreaterThan' | 'LessThan' | 'Clear').
- "CREATE_CHART": membuat diagram grafis Excel asli. Membutuhkan "type" ('bar'|'line'|'pie'|'area'), "range" (rentang data grafik, contoh: "A1:G11" atau nama tabel "DataNilaiSiswa"), dan "title".
- "CREATE_PIVOT": membuat Pivot Table di sheet baru. Membutuhkan "source" (range/tabel asal), "rowFields" (array nama kolom baris), dan "dataFields" (array nama kolom nilai numerik).

Hanya hasilkan format JSON valid tanpa markdown tambahan.
`;

  try {
    const response = await model.generateContent({
      contents: [{ role: 'user', parts: [{ text: systemInstructions + `\n\nPerintah Pengguna: ${prompt}` }] }]
    });

    const text = response.response.text();
    const cleanJson = text.replace(/```json/g, "").replace(/```/g, "").trim();
    return JSON.parse(cleanJson) as ActionPlan;
  } catch (error) {
    console.warn("Model utama gemini-2.5-flash sibuk/error, mencoba model cadangan gemini-2.0-flash...", error);
    try {
      const fallbackModel = genAI.getGenerativeModel({ 
        model: 'gemini-2.0-flash',
        generationConfig: {
          responseMimeType: 'application/json'
        }
      });
      const response = await fallbackModel.generateContent({
        contents: [{ role: 'user', parts: [{ text: systemInstructions + `\n\nPerintah Pengguna: ${prompt}` }] }]
      });
      const text = response.response.text();
      const cleanJson = text.replace(/```json/g, "").replace(/```/g, "").trim();
      return JSON.parse(cleanJson) as ActionPlan;
    } catch (fallbackError) {
      console.warn("Model cadangan gemini-2.0-flash juga gagal, mencoba gemini-1.5-flash...", fallbackError);
      try {
        const fallbackModel2 = genAI.getGenerativeModel({ 
          model: 'gemini-1.5-flash',
          generationConfig: {
            responseMimeType: 'application/json'
          }
        });
        const response = await fallbackModel2.generateContent({
          contents: [{ role: 'user', parts: [{ text: systemInstructions + `\n\nPerintah Pengguna: ${prompt}` }] }]
        });
        const text = response.response.text();
        const cleanJson = text.replace(/```json/g, "").replace(/```/g, "").trim();
        return JSON.parse(cleanJson) as ActionPlan;
      } catch (fallbackError2) {
        console.error("Semua model Gemini gagal:", fallbackError2);
        throw error;
      }
    }
  }
}

/**
 * Mengirim perintah pengguna dan data workbook ke DeepSeek API untuk menghasilkan ActionPlan.
 */
export async function getDeepSeekExcelAction(
  prompt: string, 
  workbookData: WorkbookData, 
  apiKey: string
): Promise<ActionPlan> {
  const activeSheet = workbookData.sheets.find(s => s.name === workbookData.activeSheetName) || workbookData.sheets[0];
  if (!apiKey) {
    return getDemoResponse(prompt, activeSheet);
  }

  const workbookStructure = formatWorkbookDataForAI(workbookData);

  const systemInstructions = `
Anda adalah pakar formula dan manipulasi Microsoft Excel tingkat lanjut (Advanced Excel Specialist & AI Agent). Tugas Anda adalah mengonversi perintah bahasa alami pengguna menjadi rencana aksi JSON terstruktur untuk dieksekusi melalui Office.js API secara akurat.

Struktur & Isi Seluruh Lembar Kerja di Workbook (Lengkap):
${workbookStructure}

PENTING UNTUK KECERDASAN DAN AKURASI:
1. REFERENSI ANTAR SHEET (VLOOKUP / HLOOKUP / FORMULA):
   - Pengguna sering kali menyimpan tabel referensi data di sheet terpisah (misalnya sheet "Data Referensi" atau "Sheet2") dan ingin melakukan lookup dari sheet aktif saat ini.
   - Pindai semua sheet di bagian "Struktur & Isi Seluruh Lembar Kerja" untuk mencari data tabel referensi yang cocok.
   - Jika melakukan pencarian lookup ke sheet eksternal, gunakan format penamaan sheet eksternal di dalam rumus (contoh: \`=VLOOKUP(A2, 'Data Referensi'!A2:C10, 2, FALSE)\`). Gunakan petik tunggal (\`'Data Referensi'\`) jika nama sheet mengandung spasi.
2. DETEKSI TEMPAT PENULISAN RUMUS/NILAI (TARGET CELL) SECARA PRESISI:
   - Pengguna sering kali menuliskan teks/label ringkasan di bagian bawah lembar kerja (misalnya di baris 13-20, kolom A/B/C) dan membiarkan sel di sebelahnya (misalnya di kolom E atau F pada baris yang sama) kosong agar diisi rumus oleh AI.
   - Pindai isi sheet aktif untuk mencari apakah sudah ada label ringkasan di bagian bawah yang sesuai dengan perintah pengguna.
   - JANGAN menulis hasil atau rumus di tempat acak (seperti di sebelah kanan di kolom L atau M) apabila di bagian bawah lembar kerja (misalnya baris 13-20) sudah disediakan label baris yang relevan. Tulislah rumus tersebut langsung di sel baris tersebut yang kosong (misalnya jika label "Banyaknya murid yang lulus" ada di A17, tulislah rumusnya langsung di E17).
3. PENDETEKSIAN DATA BARU:
   - Jika pengguna meminta pembuatan data baru yang jenisnya berbeda dari tabel aktif saat ini (misalnya meminta "Data Siswa/UTS" sedangkan sheet aktif berisi "Data Produk/Barang"), atau jika pengguna meminta data baru secara spesifik: 
     * JANGAN menyambung (append) data tersebut di bawah kolom lama.
     * Tulis ulang seluruh sheet aktif mulai dari A1 dengan judul kolom (header) yang baru dan sesuai.
     * Buat data dummy yang sangat kaya, lengkap dengan rumus perhitungan di kolom kanan.
4. PENULISAN FORMULA EXCEL:
   - Selalu tulis Rumus Excel dalam BAHASA INGGRIS (seperti SUM, AVERAGE, COUNT, MIN, MAX, VLOOKUP, HLOOKUP, IF, COUNTIF, SUMIF, dll.) diawali tanda '='. Excel akan menerjemahkannya ke bahasa Excel lokal pengguna secara otomatis.
   - Buat formula yang dinamis merujuk ke baris sel yang tepat. Perhatikan OFFSET baris! Karena baris header berada pada baris ${ (activeSheet.rowIndex || 0) + 1 } dan baris data dimulai dari baris ${ (activeSheet.rowIndex || 0) + 2 }, maka untuk baris data pertama, rujuklah indeks baris ke-${ (activeSheet.rowIndex || 0) + 2 } (contoh: \`=AVERAGE(C${ (activeSheet.rowIndex || 0) + 2 }:E${ (activeSheet.rowIndex || 0) + 2 })\`).
   - Tentukan rentang sel secara presisi. Rentang data utama selalu dimulai dari baris ${ (activeSheet.rowIndex || 0) + 2 } dan berakhir di baris ${ (activeSheet.rowIndex || 0) + activeSheet.rowCount }.

Format Output JSON Harus Tepat Seperti Ini (dan merupakan objek JSON valid tanpa teks lain):
{
  "explanation": "Penjelasan langkah demi langkah dalam bahasa Indonesia yang ramah mengenai apa yang akan dilakukan AI.",
  "actions": [
    {
      "type": "WRITE_VALUES",
      "payload": {
        "range": "A1:H1",
        "values": [["No", "Nama Siswa", "Matematika", "Fisika", "Kimia", "Rata-rata", "Status"]]
      }
    },
    {
      "type": "WRITE_FORMULAS",
      "payload": {
        "range": "F2:G11",
        "formulas": [
          ["=AVERAGE(C2:E2)", "=IF(F2>=75,\\"LULUS\\",\\"REMIDI\\")"],
          ["=AVERAGE(C3:E3)", "=IF(F3>=75,\\"LULUS\\",\\"REMIDI\\")"]
        ]
      }
    },
    {
      "type": "CREATE_TABLE",
      "payload": {
        "range": "A1:G11",
        "name": "DataNilaiSiswa"
      }
    }
  ]
}

Ketentuan Tipe Aksi:
- "CREATE_TABLE": mengonversi range data menjadi objek Tabel Excel tepercaya (mempermudah sort/filter).
- "WRITE_FORMULAS": menulis array 2D rumus (string[][]) ke range tertentu. Baris header tidak boleh ditulis di sini, gunakan WRITE_VALUES untuk menulis teks/nilai biasa.
- "WRITE_VALUES": menulis nilai mentah biasa (string/number/array 2D) ke range sel tertentu.
- "SORT": mengurutkan kolom pada tabel. Membutuhkan "tableName", "columnName", dan "direction" ('asc' atau 'desc').
- "FILTER": memfilter kolom pada tabel. Membutuhkan "tableName", "columnName", "value", dan "operator" ('Equal' | 'Contains' | 'GreaterThan' | 'LessThan' | 'Clear').
- "CREATE_CHART": membuat diagram grafis Excel asli. Membutuhkan "type" ('bar'|'line'|'pie'|'area'), "range" (rentang data grafik), dan "title".
- "CREATE_PIVOT": membuat Pivot Table di sheet baru. Membutuhkan "source" (range/tabel asal), "rowFields" (array nama kolom baris), dan "dataFields" (array nama kolom nilai numerik).

Hanya hasilkan format JSON valid tanpa markdown tambahan (\`\`\`json \`\`\`).
`;

  const response = await fetch("https://api.deepseek.com/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: "deepseek-chat",
      messages: [
        { role: "system", content: systemInstructions },
        { role: "user", content: `Perintah Pengguna: ${prompt}` }
      ],
      response_format: {
        type: "json_object"
      },
      temperature: 0.1
    })
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`DeepSeek API Error: ${response.status} ${response.statusText} - ${errText}`);
  }

  const resData = await response.json();
  const content = resData.choices[0].message.content;
  const cleanJson = content.replace(/```json/g, "").replace(/```/g, "").trim();
  return JSON.parse(cleanJson) as ActionPlan;
}
