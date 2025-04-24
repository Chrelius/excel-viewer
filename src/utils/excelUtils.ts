import * as XLSX from 'xlsx';
import { ExcelData } from '../types/excel.types';

export const excelUtils = {
  readExcelFile: (file: File): Promise<ExcelData> => {
    return new Promise((resolve, reject) => {
      try {
        const reader = new FileReader();
        
        reader.onload = (e) => {
          try {
            const data = e.target?.result;
            if (!data) {
              throw new Error('No data read from file');
            }

            const workbook = XLSX.read(data, { 
              type: 'binary',
              cellStyles: true,
              cellDates: true,
              cellFormula: true,
              dateNF: 'yyyy-mm-dd hh:mm:ss'
            });

            const headers: { [sheetName: string]: string[] } = {};
            const rows: { [sheetName: string]: any[][] } = {};
            const formats: { [sheetName: string]: string[][] } = {};

            // Process each sheet
            workbook.SheetNames.forEach(sheetName => {
              const worksheet = workbook.Sheets[sheetName];
              
              // Function to get cell format
              const getCellFormat = (cell: XLSX.CellObject | undefined): string => {
                if (!cell) return '';
                if (cell.t === 'd') return 'YYYY-MM-DD HH:mm:ss';
                return typeof cell.z === 'string' ? cell.z : '@';
              };

              const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
              const sheetHeaders: string[] = [];
              const sheetRows: any[][] = [];
              const sheetFormats: string[][] = [];

              // Extract headers
              for (let C = range.s.c; C <= range.e.c; ++C) {
                const cell = worksheet[XLSX.utils.encode_cell({ r: 0, c: C })];
                sheetHeaders.push(cell ? String(cell.v || '') : '');
              }

              // Extract data rows
              for (let R = 1; R <= range.e.r; ++R) {
                const row: any[] = [];
                const formatRow: string[] = [];
                
                for (let C = range.s.c; C <= range.e.c; ++C) {
                  const cell = worksheet[XLSX.utils.encode_cell({ r: R, c: C })];
                  let value = '';
                  
                  if (cell) {
                    if (cell.t === 'd' && cell.v) {
                      value = XLSX.SSF.format('yyyy-mm-dd hh:mm:ss', cell.v);
                    } else {
                      value = cell.v;
                    }
                  }
                  
                  row.push(value);
                  formatRow.push(getCellFormat(cell));
                }
                
                sheetRows.push(row);
                sheetFormats.push(formatRow);
              }

              headers[sheetName] = sheetHeaders;
              rows[sheetName] = sheetRows;
              formats[sheetName] = sheetFormats;
            });

            resolve({
              headers,
              rows,
              sheetNames: workbook.SheetNames,
              activeSheet: workbook.SheetNames[0],
              formats
            });
          } catch (error) {
            console.error('Excel parsing error:', error);
            reject(new Error(`Error parsing Excel file: ${error}`));
          }
        };

        reader.onerror = () => {
          reject(new Error('Error reading file'));
        };

        reader.readAsBinaryString(file);
      } catch (error) {
        reject(new Error(`Error processing file: ${error}`));
      }
    });
  },

  exportToExcel: (data: ExcelData): void => {
    try {
      const wb = XLSX.utils.book_new();

      // Process each sheet
      data.sheetNames.forEach(sheetName => {
        // Create worksheet for each sheet
        const ws = XLSX.utils.aoa_to_sheet([
          data.headers[sheetName], 
          ...data.rows[sheetName]
        ]);

        // Set column widths
        const colWidths = data.headers[sheetName].map((header, idx) => ({
          wch: Math.max(
            header.length,
            ...data.rows[sheetName].map(row => String(row[idx] || '').length),
            15
          )
        }));

        ws['!cols'] = colWidths;

        // Add styling
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        for (let R = range.s.r; R <= range.e.r; ++R) {
          for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[cellAddress]) continue;

            ws[cellAddress].s = {
              font: { name: 'Arial', sz: 11 },
              alignment: {
                horizontal: R === 0 ? 'center' : 'left',
                vertical: 'center',
                wrapText: true
              },
              border: {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                left: { style: 'thin' },
                right: { style: 'thin' }
              }
            };

            if (R === 0) {
              ws[cellAddress].s.font.bold = true;
              ws[cellAddress].s.fill = {
                fgColor: { rgb: "EFEFEF" },
                patternType: 'solid'
              };
            }
          }
        }

        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      });

      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const fullFileName = `excel_export_${timestamp}.xlsx`;

      XLSX.writeFile(wb, fullFileName, {
        bookType: 'xlsx',
        bookSST: false,
        type: 'binary',
        cellStyles: true,
        cellDates: true
      });
    } catch (error) {
      console.error('Export error:', error);
      throw new Error(`Error exporting to Excel: ${error}`);
    }
  }
};