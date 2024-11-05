import React, { useState } from 'react';
import * as ExcelJS from 'exceljs';

const ExcelUploaderWithColors = () => {
  const [data, setData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [mergedCells, setMergedCells] = useState([]);

  const handleFileUpload = async (event) => {
    const file = event.target.files[0];

    if (!file) {
      console.warn("No file selected.");
      return;
    }

    const reader = new FileReader();

    reader.onload = async (e) => {
      const arrayBuffer = e.target.result;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);
      const worksheet = workbook.worksheets[0];

      // Get merged cells correctly
      const mergedCells = [];
      for (const [key, value] of Object.entries(worksheet._merges)) {
        const { model } = value;
        mergedCells.push({
          startRow: model.top - 1, // Convert to 0-based
          endRow: model.bottom - 1, // Convert to 0-based
          startCol: model.left - 1, // Convert to 0-based
          endCol: model.right - 1, // Convert to 0-based
        });
      }

      const jsonData = [];
      const headers = [];

      worksheet.eachRow((row, rowNumber) => {
        const rowData = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          let cellValue = cell.value;

          // Handle percentages, dates, and numbers
          if (cell.style?.numFmt && cell.style.numFmt.includes('%')) {
            cellValue = `${(cellValue * 100).toFixed(0)}%`;
          } else if (cellValue instanceof Date) {
            cellValue = cellValue.toLocaleDateString();
          } else if (typeof cellValue === 'number') {
            cellValue = cellValue.toLocaleString();
          }

          rowData.push({
            value: cellValue || '', // Ensure empty cells are represented
            style: cell.style,
          });
        });

        if (rowNumber === 1) {
          headers.push(...rowData);
        } else {
          jsonData.push(rowData);
        }
      });

      setColumns(headers);
      setData(jsonData);
      setMergedCells(mergedCells);
    };

    reader.readAsArrayBuffer(file);
  };

  const getColor = (color) => {
    if (!color) return 'black';
    if (color.argb) {
      return `#${color.argb.slice(2)}`;
    }
    if (color.theme) {
      const themeColors = [
        '#FFFFFF', '#1F497D', '#C0504D', '#9BBB59',
        '#4F81BD', '#F79646', '#C8C8C8', '#000000'
      ];
      return themeColors[color.theme - 1] || 'black';
    }
    return 'black';
  };

  const getCellStyle = (cell) => {
    const bgColor = getColor(cell.style?.fill?.fgColor) || 'transparent';
    let fontColor = getColor(cell.style?.font?.color) || 'black';
    if (fontColor === '#FFFFFF' && bgColor === 'transparent') {
      fontColor = '#000000';
    }
    const fontSize = cell.style?.font?.size ? `${cell.style.font.size}px` : '16px';
    const fontWeight = cell.style?.font?.bold ? 'bold' : 'normal';
    const fontFamily = cell.style?.font?.name || 'Arial';
    return {
      border: '1px solid black',
      padding: '8px',
      textAlign: 'center',
      verticalAlign: 'middle',
      backgroundColor: bgColor,
      color: fontColor,
      fontWeight: fontWeight,
      fontSize: fontSize,
      fontFamily: fontFamily,
    };
  };

  const renderCells = (row, rowIndex) => {
    const cellsToRender = [];
    let currentCellIndex = 0;

    while (currentCellIndex < row.length) {
      const cell = row[currentCellIndex];

      const mergedCell = mergedCells.find(merge =>
        rowIndex >= merge.startRow && rowIndex <= merge.endRow &&
        currentCellIndex >= merge.startCol && currentCellIndex <= merge.endCol
      );

      if (mergedCell && rowIndex === mergedCell.startRow && currentCellIndex === mergedCell.startCol) {
        const colSpan = mergedCell.endCol - mergedCell.startCol + 1;
        const rowSpan = mergedCell.endRow - mergedCell.startRow + 1;

        cellsToRender.push(
          <td
            key={currentCellIndex}
            colSpan={colSpan}
            rowSpan={rowSpan}
            style={getCellStyle(cell)}
          >
            {cell.value}
          </td>
        );

        currentCellIndex += colSpan; // Skip over the merged cells
      } else if (!mergedCell) {
        cellsToRender.push(
          <td key={currentCellIndex} style={getCellStyle(cell)}>
            {cell.value}
          </td>
        );
        currentCellIndex++;
      } else {
        currentCellIndex++; // Skip the merged cell part
      }
    }

    return <tr key={rowIndex}>{cellsToRender}</tr>;
  };

  const renderHeaders = () => {
    const headerCells = [];
    let currentCellIndex = 0;

    while (currentCellIndex < columns.length) {
      const col = columns[currentCellIndex];
      const mergedCell = mergedCells.find(merge =>
        0 >= merge.startRow && 0 <= merge.endRow && // Header row (0-based index)
        currentCellIndex >= merge.startCol && currentCellIndex <= merge.endCol
      );

      if (mergedCell && 0 === mergedCell.startRow && currentCellIndex === mergedCell.startCol) {
        const colSpan = mergedCell.endCol - mergedCell.startCol + 1;

        headerCells.push(
          <th
            key={currentCellIndex}
            colSpan={colSpan}
            style={getCellStyle(col)}
          >
            {col.value}
          </th>
        );

        currentCellIndex += colSpan; // Skip over merged cells
      } else if (!mergedCell) {
        headerCells.push(
          <th key={currentCellIndex} style={getCellStyle(col)}>
            {col.value}
          </th>
        );
        currentCellIndex++;
      } else {
        currentCellIndex++; // Skip merged part
      }
    }

    return <tr>{headerCells}</tr>;
  };

  return (
    <div>
      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      <table style={{ marginTop: 20, width: '100%', borderCollapse: 'collapse' }}>
        <thead>
          {renderHeaders()}
        </thead>
        <tbody>
          {data.map(renderCells)}
        </tbody>
      </table>
    </div>
  );
};

export default ExcelUploaderWithColors;
