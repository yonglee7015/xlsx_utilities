/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("app-body").style.display = "flex";
    initializeAddIn();
  }
});

async function initializeAddIn() {
  try {
    document.getElementById("split-worksheet").onclick = splitWorksheet;
    document.getElementById("select-header-and-column").onclick = selectHeaderAndColumn;
    document.getElementById("confirm-yes").onclick = () => confirmCallback(true);
    document.getElementById("confirm-no").onclick = () => confirmCallback(false);
    updateStatus("Ready to split worksheet. Click 'Select Header & Split Column' to begin.", "info");
  } catch (error) {
    console.error(error);
    updateStatus(`Error initializing add-in: ${error.message}`, "error");
  }
}

// Store selected header row and column
let selectedHeaderRow: number | null = null;
let selectedColumnIndex: number | null = null;
let selectedColumnName: string | null = null;
let headerValues: string[] = [];

async function selectHeaderAndColumn() {
  try {
    updateStatus("Step 1: Please select any cell in the header row...", "info");
    
    // Step 1: Select header row
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("rowIndex");
      
      await context.sync();
      
      if (range.rowIndex !== undefined) {
        // Store header row (1-based)
        selectedHeaderRow = range.rowIndex + 1;
        
        // Get only the selected cell's value as header value
        const selectedCellRange = context.workbook.getSelectedRange();
        selectedCellRange.load("values");
        
        await context.sync();
        
        // Extract just the selected cell's value
        if (selectedCellRange.values && selectedCellRange.values[0] && selectedCellRange.values[0][0]) {
          headerValues = [String(selectedCellRange.values[0][0])];
        } else {
          headerValues = [];
        }
        
        updateSelectionSummary();
        updateStatus("Step 2: Please select any cell in the column you want to split by...", "info");
        
        // Step 2: Select split column
        await context.sync();
        
        const columnRange = context.workbook.getSelectedRange();
        columnRange.load("columnIndex");
        
        await context.sync();
        
        if (columnRange.columnIndex !== undefined) {
          // Store column information
          selectedColumnIndex = columnRange.columnIndex;
          selectedColumnName = getColumnName(columnRange.columnIndex + 1); // Convert to 1-based
          updateSelectionSummary();
          updateStatus(`Successfully selected header row ${selectedHeaderRow} and column ${selectedColumnName}`, "success");
        } else {
          updateStatus("Please select a single cell for the split column", "error");
          resetSelection();
        }
      } else {
        updateStatus("Please select a single cell in the header row", "error");
        resetSelection();
      }
    });
  } catch (error) {
    console.error(error);
    updateStatus(`Error selecting header and column: ${error.message}`, "error");
    resetSelection();
  }
}

function updateSelectionSummary() {
  // Update header row display
  const headerRowDisplay = document.getElementById("header-row-display");
  if (headerRowDisplay) {
    headerRowDisplay.textContent = selectedHeaderRow ? selectedHeaderRow.toString() : "Not selected";
  }
  
  // Update header value display (only show the first value)
  const headerValuesDisplay = document.getElementById("header-values-display");
  if (headerValuesDisplay) {
    if (headerValues.length > 0) {
      headerValuesDisplay.textContent = headerValues[0];
    } else {
      headerValuesDisplay.textContent = "No header value found";
    }
  }
  
  // Update column display
  const columnDisplay = document.getElementById("column-display");
  if (columnDisplay) {
    columnDisplay.textContent = selectedColumnName ? selectedColumnName : "Not selected";
  }
}

function resetSelection() {
  selectedHeaderRow = null;
  selectedColumnIndex = null;
  selectedColumnName = null;
  headerValues = [];
  updateSelectionSummary();
}



let confirmCallback: (result: boolean) => void;

function showConfirmDialog(message: string): Promise<boolean> {
  return new Promise((resolve) => {
    const dialog = document.getElementById("confirm-dialog");
    const messageElement = document.getElementById("confirm-message");
    
    if (dialog && messageElement) {
      messageElement.textContent = message;
      dialog.style.display = "flex";
      
      confirmCallback = (result: boolean) => {
        dialog.style.display = "none";
        resolve(result);
      };
    } else {
      resolve(false);
    }
  });
}



function getColumnName(columnIndex: number): string {
  let name = "";
  let index = columnIndex;
  
  while (index > 0) {
    const remainder = (index - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    index = Math.floor((index - 1) / 26);
  }
  
  return name;
}

async function splitWorksheet() {
  try {
    // Validate selections
    if (selectedHeaderRow === null || selectedColumnIndex === null || selectedColumnName === null) {
      updateStatus("Please select both header row and split column using the 'Select Header & Split Column' button.", "error");
      return;
    }
    
    const columnIndex = selectedColumnIndex;
    const columnName = selectedColumnName;
    const headerRow = selectedHeaderRow;
    
    const confirmed = await showConfirmDialog("Are you sure you want to split the worksheet? This will create multiple new worksheets in the current workbook.");
    if (!confirmed) {
      return;
    }
    
    updateStatus("Starting worksheet split...", "info");
    showProgress("Splitting worksheet...", 0);
    
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = worksheet.getUsedRange();
      usedRange.load("values, rowCount, columnCount");
      
      await context.sync();
      
      if (!usedRange.values || usedRange.rowCount <= 1) {
        updateStatus("No data found in the worksheet", "error");
        hideProgress();
        return;
      }
      
      // Validate header row position
      if (headerRow < 1 || headerRow > usedRange.rowCount) {
        updateStatus(`Invalid header row position. Please select a row between 1 and ${usedRange.rowCount}.`, "error");
        hideProgress();
        return;
      }
      
      // Calculate 0-based indices
      const headerRowIndex = headerRow - 1;
      const dataStartIndex = headerRow;
      
      // Get headers from specified row
      const headers = usedRange.values[headerRowIndex] || [];
      
      // Process column values and unique values (columnIndex is already 0-based)
      const columnValues = usedRange.values.map((row: any[]) => row[columnIndex]);
      const uniqueValues = [...new Set(columnValues.slice(dataStartIndex))].filter(value => value !== null && value !== undefined && value !== "");
      
      // Process all unique values in batches to reduce context.sync() calls
      const batchSize = 8; // Increased batch size for better performance
      let progress = 0;
      const totalSteps = uniqueValues.length;
      
      for (let batchStart = 0; batchStart < uniqueValues.length; batchStart += batchSize) {
        const batchEnd = Math.min(batchStart + batchSize, uniqueValues.length);
        const currentBatch = uniqueValues.slice(batchStart, batchEnd);
        
        // Process each value in the current batch
        for (const value of currentBatch) {
          const newWorksheet = context.workbook.worksheets.add();
          newWorksheet.name = sanitizeSheetName(String(value));
          
          // Track rows that match the current value
          const matchingRowIndexes: number[] = [];
          for (let i = dataStartIndex; i < usedRange.rowCount; i++) {
            if (columnValues[i] === value) {
              matchingRowIndexes.push(i);
            }
          }
          
          if (matchingRowIndexes.length > 0) {
            // Copy header with complete format
            const headerSourceRange = worksheet.getRangeByIndexes(headerRowIndex, 0, 1, usedRange.columnCount);
            const headerDestRange = newWorksheet.getRangeByIndexes(0, 0, 1, usedRange.columnCount);
            headerDestRange.copyFrom(headerSourceRange, Excel.RangeCopyType.all, false, false);
            
            // Copy each data row with its original complete format
            // This ensures 100% format fidelity with the original worksheet
            for (let i = 0; i < matchingRowIndexes.length; i++) {
              const sourceRowIndex = matchingRowIndexes[i];
              const destRowIndex = i + 1; // +1 for header row
              
              // Get source and destination ranges for this row
              const sourceRange = worksheet.getRangeByIndexes(sourceRowIndex, 0, 1, usedRange.columnCount);
              const destRange = newWorksheet.getRangeByIndexes(destRowIndex, 0, 1, usedRange.columnCount);
              
              // Copy the entire row with ALL properties (values + formats + formulas + conditional formatting + etc.)
              // This ensures exact format matching with the original worksheet
              destRange.copyFrom(sourceRange, Excel.RangeCopyType.all, false, false);
            }
            
            // Autofit columns to ensure proper display
            const newUsedRange = newWorksheet.getUsedRange();
            newUsedRange.format.autofitColumns();
          }
          
          progress++;
          updateProgress(Math.round((progress / totalSteps) * 100), `Processing ${value}...`);
        }
        
        // Sync once per batch instead of per value - this is the key performance optimization
        await context.sync();
      }
      
      hideProgress();
      updateStatus(`Worksheet split completed. Created ${uniqueValues.length} new worksheets.`, "success");
    });
  } catch (error) {
    console.error(error);
    updateStatus(`Error splitting worksheet: ${error.message}`, "error");
    hideProgress();
  }
}



function sanitizeSheetName(name: string): string {
  const invalidChars = ["\\", "/", "?", "*", "[", "]"];
  let sanitized = name;
  for (const char of invalidChars) {
    sanitized = sanitized.replace(new RegExp(`\\${char}`, "g"), "_");
  }
  return sanitized.substring(0, 31); // Excel sheet names are limited to 31 characters
}

function sanitizeFileName(name: string): string {
  const invalidChars = ["\\", "/", ":", "*", "?", "\"", "<", ">", "|"];
  let sanitized = name;
  for (const char of invalidChars) {
    sanitized = sanitized.replace(new RegExp(`\\${char}`, "g"), "_");
  }
  return sanitized;
}

function updateStatus(message: string, type: string = "info") {
  const statusElement = document.getElementById("status-message");
  if (statusElement) {
    statusElement.textContent = message;
    statusElement.className = `status-message ${type === "error" ? "error" : type === "success" ? "success" : ""}`;
  }
}

function showProgress(message: string, value: number) {
  const progressContainer = document.getElementById("progress-container");
  const progressMessage = document.getElementById("progress-message");
  const progressBar = document.getElementById("progress-bar");
  const progressStatus = document.getElementById("progress-status");
  
  if (progressContainer && progressMessage && progressBar && progressStatus) {
    progressContainer.style.display = "block";
    progressMessage.textContent = message;
    progressBar.style.width = `${value}%`;
    progressStatus.textContent = `${value}%`;
  }
}

function updateProgress(value: number, message: string) {
  const progressMessage = document.getElementById("progress-message");
  const progressBar = document.getElementById("progress-bar");
  const progressStatus = document.getElementById("progress-status");
  
  if (progressMessage && progressBar && progressStatus) {
    progressMessage.textContent = message;
    progressBar.style.width = `${value}%`;
    progressStatus.textContent = `${value}%`;
  }
}

function hideProgress() {
  const progressContainer = document.getElementById("progress-container");
  if (progressContainer) {
    progressContainer.style.display = "none";
  }
}
