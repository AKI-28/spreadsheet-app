import React, { useState, useEffect, useCallback } from "react";
import './Spreadsheet.css';
import * as XLSX from "xlsx";

const parseCellRange = (rangeStr) => {
  const match = rangeStr.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
  if (!match) return [];

  const colStart = match[1].charCodeAt(0) - "A".charCodeAt(0);
  const rowStart = parseInt(match[2], 10) - 1;
  const colEnd = match[3].charCodeAt(0) - "A".charCodeAt(0);
  const rowEnd = parseInt(match[4], 10) - 1;

  let range = [];
  for (let i = rowStart; i <= rowEnd; i++) {
    for (let j = colStart; j <= colEnd; j++) {
      range.push([i, j]);
    }
  }
  return range;
};

//Mathematical Functions
const evaluateFormula = (formula, cells) => {
if (typeof formula !== "string" || !formula.startsWith("=")) return formula;

  const match = formula.slice(1).match(/(SUM|AVERAGE|MAX|MIN|COUNT)\((.*)\)/);
  if (!match) return formula;

  const func = match[1];
  const range = parseCellRange(match[2]);
  if (range.length === 0) return "ERROR";

  let values = range.map(([row, col]) => Number(cells[row]?.[col]?.value) || 0);

  switch (func) {
    case "SUM": return values.reduce((a, b) => a + b, 0);
    case "AVERAGE": return values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0;
    case "MAX": return Math.max(...values);
    case "MIN": return Math.min(...values);
    case "COUNT": return values.filter(v => v !== 0).length;
    default: return "ERROR";
  }
};

//Data Entry and Validation
const validateCellValue = (value, type) => {
  if (type === "number") {
    if (value.trim() === "") return ""; // Allow empty cells
    if (!/^-?\d*\.?\d+$/.test(value)) return "ERROR: Must be a number"; // Strict numeric check
    return { valid: true, value: Number(value) }; // Return a structured object
  }

  if (type === "date") {
    if (value.trim() === "") return "";
    const parsedDate = Date.parse(value);
    if (isNaN(parsedDate)) return "ERROR: Invalid date";
    return { valid: true, value: new Date(parsedDate).toISOString().split("T")[0] }; // Format as YYYY-MM-DD
  }

  return { valid: true, value }; // Default case for text
};

const Spreadsheet = () => {
  const [cells, setCells] = useState(Array(10).fill(null).map(() => Array(10).fill({ value: "", raw: "", bold: false, italic: false, fontSize: "14px", color: "black" })));
  const [selectedCell, setSelectedCell] = useState(null);
  const [columnWidths, setColumnWidths] = useState(Array(10).fill(80));
  const [rowHeights, setRowHeights] = useState(Array(10).fill(30));
  const [columnTypes, setColumnTypes] = useState(Array(10).fill("text"));
  const [resizing, setResizing] = useState(null);

   // Save the current spreadsheet to Excel file
   const handleSave = () => {
    const ws = XLSX.utils.aoa_to_sheet(cells.map(row => row.map(cell => cell.value)));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

    // Create an Excel file and trigger download
    XLSX.writeFile(wb, "spreadsheet.xlsx");
    alert("Spreadsheet saved as Excel!");
  };

  // Reload the spreadsheet from an uploaded Excel file
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    // Read the uploaded Excel file
    const reader = new FileReader();
    reader.onload = () => {
      const data = reader.result;
      const wb = XLSX.read(data, { type: "binary" });

      // Assuming the first sheet is the one we want
      const ws = wb.Sheets[wb.SheetNames[0]];

      // Convert sheet data to 2D array (cells)
      const cellsData = XLSX.utils.sheet_to_json(ws, { header: 1 });

      // Set the cells state with the data from the file
      const newCells = cellsData.map(row => row.map(value => ({ value, raw: value, bold: false, italic: false, fontSize: "14px", color: "black" })));

      // Assuming column widths and row heights are stored as extra data, you can extract them similarly (for now, we'll just reset them)
      setCells(newCells);
      setColumnWidths(Array(newCells[0]?.length || 10).fill(80)); // Adjust column widths if necessary
      setRowHeights(Array(newCells.length || 10).fill(30)); // Adjust row heights if necessary

      alert("Spreadsheet reloaded!");
    };

    reader.readAsBinaryString(file);
  };

  useEffect(() => {
    setCells(prevCells => {
      return prevCells.map(row => row.map(cell => ({
        ...cell,
        value: evaluateFormula(cell.raw, prevCells)
      })));
    });
  }, [cells]);
  
  const handleChange = (row, col, value) => {
    setCells(prevCells => {
      const type = columnTypes[col];
      const validationResult = validateCellValue(value, type);
  
      return prevCells.map((r, i) =>
        r.map((c, j) =>
          i === row && j === col
            ? {
                ...c,
                raw: validationResult.valid ? value : "",  // Clear raw input on error
                value: validationResult.valid ? validationResult.value : "",  // Clear displayed value on error
                error: validationResult.valid ? "" : validationResult, // Store error message
              }
            : c
        )
      );
    });
  
    // Show an alert and clear input
    const validationResult = validateCellValue(value, columnTypes[col]);
    if (!validationResult.valid) {
      alert(validationResult);
    }
  };

//Add, Delete, and Resize rows and columns.
  const handleMouseDown = (type, index, event) => {
    setResizing({ type, index, startX: event.clientX, startY: event.clientY });
  };

  const handleMouseMove = useCallback((event) => {
    if (!resizing) return;

    if (resizing.type === "col") {
      const newWidths = [...columnWidths];
      newWidths[resizing.index] += event.clientX - resizing.startX;
      setColumnWidths(newWidths);
      setResizing({ ...resizing, startX: event.clientX });
    } else if (resizing.type === "row") {
      const newHeights = [...rowHeights];
      newHeights[resizing.index] += event.clientY - resizing.startY;
      setRowHeights(newHeights);
      setResizing({ ...resizing, startY: event.clientY });
    }
  }, [resizing, columnWidths, rowHeights]); // Ensure dependencies are set correctly

  const handleMouseUp = () => setResizing(null);

useEffect(() => {
  window.addEventListener("mousemove", handleMouseMove);
  window.addEventListener("mouseup", handleMouseUp);
  return () => {
    window.removeEventListener("mousemove", handleMouseMove);
    window.removeEventListener("mouseup", handleMouseUp);
  };
}, [resizing, handleMouseMove]); // Add handleMouseMove as a dependency

  const handleFormat = (formatType, value) => {
    if (!selectedCell) return alert("Select a cell first!");
    const { row, col } = selectedCell;
    setCells(prevCells =>
      prevCells.map((r, i) =>
        r.map((c, j) => (i === row && j === col ? { ...c, [formatType]: value } : c))
      )
    );
  };

  const handleCellClick = (row, col) => setSelectedCell({ row, col });

  const handleAddRow = () => {
    const newRow = Array(cells[0].length).fill({ value: "", raw: "" });
    setCells([...cells, newRow]);
  };

  const handleDeleteRow = () => {
    if (cells.length > 1) {
      setCells(cells.slice(0, -1));
    }
  };
  const handleAddColumn = () => {
    setCells(prevCells =>
      prevCells.map(row => [...row, { value: "", raw: "", type: "text" }])
    );
    setColumnTypes(prevTypes => [...prevTypes, "text"]);
    setColumnWidths(prevWidths => [...prevWidths, 80]); // Add default column width
  };
  
  const handleDeleteColumn = () => {
    if (cells[0].length > 1) {
      setCells(prevCells => prevCells.map(row => row.slice(0, -1)));
      setColumnTypes(prevTypes => prevTypes.slice(0, -1));
      setColumnWidths(prevWidths => prevWidths.slice(0, -1)); // Remove width entry
    }
  };
  
  // Ensure column labels adjust dynamically
  const columnLabels = Array.from({ length: cells[0]?.length || 0 }, (_, i) =>
    String.fromCharCode(65 + i)
  );

//Data Quality Functions
  const handleRemoveDuplicates = () => {
    const uniqueRows = [];
    const seen = new Set();
    
    cells.forEach(row => {
      const rowString = JSON.stringify(row.map(cell => cell.raw));
      if (!seen.has(rowString)) {
        seen.add(rowString);
        uniqueRows.push(row);
      }
    });
    setCells(uniqueRows);
  };

  const handleTrim = () => {
    if (!selectedCell) return;
    const { row, col } = selectedCell;
    setCells(prevCells => prevCells.map((r, i) => r.map((c, j) => (i === row && j === col ? { ...c, value: c.value.trim(), raw: c.raw.trim() } : c))));
  };

  const handleUpper = () => {
    if (!selectedCell) return;
    const { row, col } = selectedCell;
    setCells(prevCells => prevCells.map((r, i) => r.map((c, j) => (i === row && j === col ? { ...c, value: c.value.toUpperCase(), raw: c.raw.toUpperCase() } : c))));
  };

  const handleLower = () => {
    if (!selectedCell) return;
    const { row, col } = selectedCell;
    setCells(prevCells => prevCells.map((r, i) => r.map((c, j) => (i === row && j === col ? { ...c, value: c.value.toLowerCase(), raw: c.raw.toLowerCase() } : c))));
  };

  const handleFindAndReplace = (findText, replaceText) => {
    const regex = new RegExp(findText, "gi"); // 'g' for global search, 'i' for case-insensitive search
  
    setCells(prevCells => 
      prevCells.map(row => 
        row.map(cell => ({
          ...cell,
          value: cell.value ? cell.value.replace(regex, replaceText) : "",
          raw: cell.raw ? cell.raw.replace(regex, replaceText) : "",
        }))
      )
    );
  };

 const handleColumnTypeChange = (col, type) => {
    setColumnTypes(prevTypes => {
      const newTypes = [...prevTypes];
      newTypes[col] = type;
      return newTypes;
    });
    setCells(prevCells => 
      prevCells.map(row => 
        row.map((c, j) => (j === col ? { ...c, type } : c))
      )
    );
  };

  return (
    <div style={{ padding: "10px" }}>
      <h2>Welcome to My Spreedsheet</h2>
      <h3>You can perform below set of operations</h3>
      {/* Toolbar */}
      <div>
      <button onClick={() => handleFormat("bold", !cells[selectedCell?.row]?.[selectedCell?.col]?.bold)}>Bold</button>
        <button onClick={() => handleFormat("italic", !cells[selectedCell?.row]?.[selectedCell?.col]?.italic)}>Italic</button>
        <button onClick={handleAddRow}>Add Row</button>
        <button onClick={handleDeleteRow}>Delete Row</button>
        <button onClick={handleAddColumn}>Add Column</button>
        <button onClick={handleDeleteColumn}>Delete Column</button>
        <button onClick={handleTrim}>Trim</button>
        <button onClick={handleUpper}>Uppercase</button>
        <button onClick={handleLower}>Lowercase</button>
        <button onClick={handleRemoveDuplicates}>Remove Duplicates</button>
        <input type="text" placeholder="Find" id="findText" />
        <input type="text" placeholder="Replace" id="replaceText" />
        <button onClick={() => handleFindAndReplace(document.getElementById("findText").value, document.getElementById("replaceText").value)}>Find & Replace</button>
        <input type="number" min="10" max="24" onChange={(e) => handleFormat("fontSize", e.target.value + "px")} placeholder="Font Size" />
        <input type="color" onChange={(e) => handleFormat("color", e.target.value)} />
        <div>
          <button onClick={handleSave}>Save Spreadsheet</button>
          <input type="file" onChange={handleFileUpload} />
          </div>
      </div>

      {/* Spreadsheet Table */}
      <table border="1" style={{ borderCollapse: "collapse", marginTop: "10px" }}>
      <thead>
          <tr>
            <th></th> {/* Empty cell for top-left corner */}
            {columnLabels.map((label, colIndex) => (
              <th key={colIndex}>
                {label}
                <br />
                <select onChange={(e) => handleColumnTypeChange(colIndex, e.target.value)}>
                  <option value="text">Text</option>
                  <option value="number">Number</option>
                  <option value="date">Date</option>
                </select>
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {cells.map((row, rowIndex) => (
           <tr key={rowIndex} style={{ height: `${rowHeights[rowIndex]}px` }}>
             <td>{rowIndex + 1}</td> {/* Row numbers */}
              {row.map((cell, colIndex) => (
                <td key={colIndex} onClick={() => handleCellClick(rowIndex, colIndex)} style={{ width: `${columnWidths[colIndex]}px` }}>
                  <textarea
                    type="text"
                    value={cell.value}
                    onChange={(e) => handleChange(rowIndex, colIndex, e.target.value)}
                    style={{
                      width: "100%",
                      height: "100%",
                      textAlign: "center",
                      fontWeight: cell.bold ? "bold" : "normal",
                      fontStyle: cell.italic ? "italic" : "normal",
                      fontSize: cell.fontSize,
                      color: cell.color,
                      border: "none",
                      outline: "none",
                      whiteSpace: "pre-wrap", // Allow multi-line text
                      wordWrap: "break-word", // Ensure long words wrap
                      resize: "none", // Prevent manual resizin
                    }}
                  />
                  <div className="resize-handle col-resize" onMouseDown={(e) => handleMouseDown("col", colIndex, e)} />
                </td>
              ))}
              <td>
                <div className="resize-handle row-resize" onMouseDown={(e) => handleMouseDown("row", rowIndex, e)} />
              </td>
            </tr>
          ))}
              
        </tbody>
      </table>
    </div>
  );
};

export default Spreadsheet;
