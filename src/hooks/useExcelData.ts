import { useState, useCallback } from 'react';
import { ExcelData, CellEdit } from '../types/excel.types';
import { excelUtils } from '../utils/excelUtils';

export const useExcelData = () => {
  const [excelData, setExcelData] = useState<ExcelData | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState<boolean>(false);

  const handleFileSelect = useCallback(async (file: File) => {
    setLoading(true);
    setError(null);

    try {
      if (!file.name.match(/\.(xlsx|xls)$/)) {
        throw new Error('Please select a valid Excel file (.xlsx or .xls)');
      }

      const data = await excelUtils.readExcelFile(file);
      setExcelData(data);
      localStorage.setItem('excelData', JSON.stringify(data));
    } catch (err) {
      console.error('File processing error:', err);
      setError(err instanceof Error ? err.message : 'An error occurred');
      setExcelData(null);
    } finally {
      setLoading(false);
    }
  }, []);

  const updateCell = useCallback((edit: CellEdit) => {
    setExcelData((prevData) => {
      if (!prevData) return null;

      const newRows = { ...prevData.rows };
      newRows[edit.position.sheet] = [...prevData.rows[edit.position.sheet]];
      newRows[edit.position.sheet][edit.position.row][edit.position.col] = edit.value;

      const newData = {
        ...prevData,
        rows: newRows
      };

      localStorage.setItem('excelData', JSON.stringify(newData));
      return newData;
    });
  }, []);

  const setActiveSheet = useCallback((sheetName: string) => {
    setExcelData((prevData) => {
      if (!prevData) return null;
      return {
        ...prevData,
        activeSheet: sheetName
      };
    });
  }, []);

  const exportData = useCallback(() => {
    if (!excelData) {
      setError('No data to export');
      return;
    }

    try {
      excelUtils.exportToExcel(excelData);
    } catch (err) {
      console.error('Export error:', err);
      setError(err instanceof Error ? err.message : 'Export failed');
    }
  }, [excelData]);

  return {
    excelData,
    error,
    loading,
    handleFileSelect,
    updateCell,
    exportData,
    setActiveSheet
  };
};