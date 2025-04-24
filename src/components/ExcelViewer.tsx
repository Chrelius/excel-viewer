import React, { useCallback, useRef, useState } from 'react';
import { useExcelData } from '../hooks/useExcelData';
import { CellPosition } from '../types/excel.types';
import { appConfig } from '../config/appConfig';
import './ExcelViewer.css';

const ExcelViewer: React.FC = () => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const {
    excelData,
    error,
    loading,
    handleFileSelect,
    updateCell,
    exportData,
    setActiveSheet
  } = useExcelData();

  const [editCell, setEditCell] = useState<CellPosition | null>(null);
  const [editValue, setEditValue] = useState<string>('');

  const formatCellValue = (value: any, sheet: string, rowIndex?: number, colIndex?: number): string => {
    if (value === null || value === undefined) return '';
    
    if (typeof rowIndex === 'number' && typeof colIndex === 'number' && excelData?.formats?.[sheet]) {
      const format = excelData.formats[sheet][rowIndex]?.[colIndex];
      if (format === 'YYYY-MM-DD HH:mm:ss') {
        if (value instanceof Date) {
          return value.toISOString().replace('T', ' ').split('.')[0];
        }
        if (typeof value === 'string' && value.match(/^\d{4}-\d{2}-\d{2}/)) {
          return value;
        }
      }
    }
    
    return String(value);
  };

  const renderVersionInfo = () => {
    const currentTime = new Date().toISOString().replace('T', ' ').split('.')[0];
    
    return (
      <div className="version-info">
        <span>Version: {appConfig.version}</span>
        <span>Build Date: {appConfig.buildDate} {appConfig.buildTime}</span>
        <span>Current Time (UTC): {currentTime}</span>
        <span>User: {appConfig.username}</span>
      </div>
    );
  };

  const onFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      await handleFileSelect(file);
    }
  };

  const startEditing = (sheet: string, row: number, col: number, value: any) => {
    setEditCell({ row, col, sheet });
    setEditValue(formatCellValue(value, sheet, row, col));
  };

  const saveEdit = useCallback(() => {
    if (editCell) {
      updateCell({
        position: editCell,
        value: editValue
      });
      setEditCell(null);
    }
  }, [editCell, editValue, updateCell]);

  const handleKeyPress = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      saveEdit();
    } else if (e.key === 'Escape') {
      setEditCell(null);
    }
  };

  const renderSheetTabs = () => {
    if (!excelData?.sheetNames.length) return null;

    return (
      <div className="sheet-tabs">
        {excelData.sheetNames.map((sheetName) => (
          <button
            key={sheetName}
            className={`sheet-tab ${excelData.activeSheet === sheetName ? 'active' : ''}`}
            onClick={() => setActiveSheet(sheetName)}
          >
            {sheetName}
          </button>
        ))}
      </div>
    );
  };

  const renderTable = () => {
    if (!excelData?.headers || !excelData?.rows || !excelData.activeSheet) {
      return null;
    }

    const activeHeaders = excelData.headers[excelData.activeSheet];
    const activeRows = excelData.rows[excelData.activeSheet];

    return (
      <div className="excel-table-container">
        <table className="excel-table">
          <thead>
            <tr>
              {activeHeaders.map((header, index) => (
                <th key={`header-${index}`}>{header}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {activeRows.map((row, rowIndex) => (
              <tr key={`row-${rowIndex}`}>
                {row.map((cell, colIndex) => (
                  <td
                    key={`cell-${rowIndex}-${colIndex}`}
                    onClick={() => startEditing(excelData.activeSheet, rowIndex, colIndex, cell)}
                    className={`
                      ${editCell?.row === rowIndex && 
                        editCell?.col === colIndex && 
                        editCell?.sheet === excelData.activeSheet ? 'editing' : ''}
                      ${excelData.formats?.[excelData.activeSheet]?.[rowIndex]?.[colIndex] === 'YYYY-MM-DD HH:mm:ss' ? 'datetime-cell' : ''}
                    `}
                  >
                    {editCell?.row === rowIndex && 
                     editCell?.col === colIndex && 
                     editCell?.sheet === excelData.activeSheet ? (
                      <input
                        type="text"
                        value={editValue}
                        onChange={(e) => setEditValue(e.target.value)}
                        onBlur={saveEdit}
                        onKeyDown={handleKeyPress}
                        autoFocus
                      />
                    ) : (
                      formatCellValue(cell, excelData.activeSheet, rowIndex, colIndex)
                    )}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  };

  return (
    <div className="excel-viewer">
      <div className="toolbar">
        <input
          type="file"
          ref={fileInputRef}
          onChange={onFileChange}
          accept=".xlsx,.xls"
          style={{ display: 'none' }}
        />
        <button
          className="toolbar-button"
          onClick={() => fileInputRef.current?.click()}
          disabled={loading}
        >
          {loading ? 'Loading...' : 'Browse Excel File'}
        </button>
        {excelData && (
          <button 
            className="toolbar-button"
            onClick={exportData}
          >
            Export Excel
          </button>
        )}
      </div>

      {error && <div className="error-message">{error}</div>}
      {loading && <div className="loading">Loading...</div>}
      
      {excelData ? (
        <>
          {renderSheetTabs()}
          {renderTable()}
          {renderVersionInfo()}
        </>
      ) : (
        !loading && !error && (
          <div className="no-data-message">
            No data to display. Please select an Excel file.
          </div>
        )
      )}
    </div>
  );
};

export default ExcelViewer;