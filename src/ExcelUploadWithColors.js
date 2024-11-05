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

      // Get merged cells
      const mergedCells = worksheet.merges.map(merge => {
        const [start, end] = merge.split(':').map(cell => ExcelJS.utils.decode_cell(cell));
        return {
          startRow: start.r + 1,
          endRow: end.r + 1,
          startCol: start.c + 1,
          endCol: end.c + 1,
        };
      });
      setMergedCells(mergedCells);

      const jsonData = [];
      const headers = [];

      worksheet.eachRow((row, rowNumber) => {
        const rowData = [];
        row.eachCell((cell) => {
          rowData.push({
            value: cell.value,
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
    };

    reader.readAsArrayBuffer(file);
  };

  const getColor = (color) => {
    if (!color) return 'black';
    if (color.argb) {
      return `#${color.argb.slice(2)}`;
    }
    if (color.theme) {
      return ['#FFFFFF', '#1F497D', '#C0504D', '#9BBB59', '#4F81BD', '#F79646', '#C8C8C8', '#000000'][color.theme - 1] || 'black';
    }
    return 'black';
  };

  const getCellStyle = (cell) => {
    const bgColor = getColor(cell.style?.fill?.fgColor) || 'transparent';
    const fontColor = getColor(cell.style?.font?.color) || 'black';

    return {
      border: '1px solid black',
      padding: '8px',
      textAlign: 'center',
      verticalAlign: 'middle',
      backgroundColor: bgColor,
      color: fontColor,
      fontWeight: cell.style?.font?.bold ? 'bold' : 'normal',
      fontSize: cell.style?.font?.size ? `${cell.style.font.size}px` : '16px',
      fontFamily: cell.style?.font?.name || 'Arial',
    };
  };

  const renderCells = (row, rowIndex) => {
    const cellsToRender = [];
    let currentCellIndex = 0;

    while (currentCellIndex < row.length) {
      const cell = row[currentCellIndex];
      const mergedCell = mergedCells.find(merge =>
        rowIndex + 1 >= merge.startRow && rowIndex + 1 <= merge.endRow &&
        currentCellIndex + 1 >= merge.startCol && currentCellIndex + 1 <= merge.endCol
      );

      if (mergedCell) {
        const colSpan = mergedCell.endCol - mergedCell.startCol + 1;
        const rowSpan = mergedCell.endRow - mergedCell.startRow + 1;

        // Render the merged cell only once
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

        // Skip the next columns that are part of the merged cell
        currentCellIndex += colSpan;
      } else {
        cellsToRender.push(
          <td key={currentCellIndex} style={getCellStyle(cell)}>
            {cell.value}
          </td>
        );
        currentCellIndex++;
      }
    }

    return <tr key={rowIndex}>{cellsToRender}</tr>;
  };

  return (
    <div>
      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      <table style={{ marginTop: 20, width: '100%', borderCollapse: 'collapse' }}>
        <thead>
          <tr>
            {columns.map((col, index) => (
              <th key={index} style={getCellStyle(col)}>
                {col.value}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {data.map(renderCells)}
        </tbody>
      </table>
    </div>
  );
};

export default ExcelUploaderWithColors;
