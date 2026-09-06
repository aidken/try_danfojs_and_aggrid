import * as XLSX from "xlsx";
import { createGrid } from "ag-grid-community";
import { Temporal } from '@js-temporal/polyfill';

import "ag-grid-community/styles/ag-grid.css";
import "ag-grid-community/styles/ag-theme-quartz.css";

document.querySelector<HTMLDivElement>("#app")!.innerHTML = `
  <!DOCTYPE html>
  <html lang="en">
  <head>
    <style>
    * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
    }

    body {
        font-family: system-ui, -apple-system, sans-serif;
        background: #f5f5f5;
        min-height: 100vh;
        padding: 20px;
    }

    .header {
        text-align: center;
        margin-bottom: 24px;
    }

    .header h1 {
        font-size: 24px;
        color: #333;
        margin-bottom: 8px;
    }

    .header p {
        font-size: 14px;
        color: #666;
    }

    .upload-area {
        max-width: 600px;
        margin: 0 auto 24px;
        border: 2px dashed #ccc;
        border-radius: 12px;
        padding: 40px 20px;
        text-align: center;
        background: #fff;
        cursor: pointer;
        transition: border-color 0.2s, background 0.2s;
    }

    .upload-area:hover,
    .upload-area.dragover {
        border-color: #4a90d9;
        background: #f0f7ff;
    }

    .upload-area svg {
        width: 48px;
        height: 48px;
        margin-bottom: 12px;
        fill: #999;
    }

    .upload-area h2 {
        font-size: 16px;
        color: #444;
        margin-bottom: 8px;
    }

    .upload-area p {
        font-size: 13px;
        color: #888;
    }

    .upload-area input[type="file"] {
        display: none;
    }
  </style>
  </head>
  <body>
  <h1>CSV / Excel Viewer</h1>

  <div>this is a client side app. upload a csv/excel file and its contents is shown as a data grid.</div>

  <div>data grid is made available with <a href="https://www.ag-grid.com/">ag-grid</a>.</div>

  <button class="btn btn-clear" id="clearBtn">
    <svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
        <path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/>
    </svg>
    Clear
  </button>

  <div class="upload-area" id="uploadArea">
      <svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6zM14 3.5L18.5 8H14V3.5zM6 20V4h7v5h5v11H6z"/>
          <path d="M8 13h8v1H8zm0 3h8v1H8zm0-6h4v1H8z"/>
      </svg>
      <h2>Drop your file here</h2>
      <p>or click to browse (supports .xlsx, .xls, .csv)</p>
      <input type="file" id="fileInput" accept=".xlsx,.xls,.csv" />
  </div>

  <div id="tabs"></div>
  <div
    id="grid"
    class="ag-theme-quartz"
    style="height:600px;width:100%;margin-top:1rem;"
  >
  </div>
  </body>
`;

const uploadArea = document.getElementById("uploadArea") as HTMLDivElement;
const fileInput = document.getElementById("fileInput") as HTMLDivElement;
const clearButton = document.getElementById("clearBtn") as HTMLButtonElement;
const tabsDiv = document.getElementById("tabs")!;
const gridDiv = document.getElementById("grid")!;

// clear button
clearButton.addEventListener('click', () => {
  tabsDiv.innerHTML = "";
  gridDiv.innerHTML = "";
});

// click the drop area to upload
uploadArea.addEventListener('click', () => fileInput.click());

// drag and drop
uploadArea.addEventListener('dragover', (e) => {
  e.preventDefault();
  uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
  uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
  e.preventDefault();
  uploadArea.classList.remove('dragover');
  const file = e.dataTransfer?.files[0];
  if (file) showFileInGrid(file);
});

// File input change
fileInput.addEventListener('change', (e) => {
  const target = e.target as HTMLInputElement;
  const file   = target.files?.[0];
  if (file) showFileInGrid(file);
});

function convertExcelSerialToPlainDate(serial: number): Temporal.PlainDate {
  const excelEpoch   = new Date(Date.UTC(1899, 11, 30));
  const milliseconds = serial * 24 * 60 * 60 * 1000;
  const date         = new Date(excelEpoch.getTime() + milliseconds);
  return Temporal.PlainDate.from({
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate()
  });
}

function isExcelDateCell(cell: XLSX.CellObject): boolean {
  if (cell.t !== "n") {return false;}
  let format = "";
  if (cell.z) {
    format = String(cell.z).toLowerCase();
    // cell.z?.toLowerCase() ?? "";
  }

  return (
    format.includes("yy") ||
    format.includes("mm") ||
    format.includes("dd")
  );
}

// now fileInput is not visible
// fileInput.addEventListener("change", (e) => {
//   // const target = e.target as HTMLInputElement;
//   const file = e.target?.files[0];
//   showFileInGrid(file);
// });

async function showFileInGrid(file:File) {

  if (!file) return;

  tabsDiv.innerHTML = "";
  gridDiv.innerHTML = "";

  const buffer = await file.arrayBuffer();
  // read workbook
  const workbook = XLSX.read(buffer, {cellDates: false, cellNF: true});

  // read all worksheets
  const sheets: Record<string, any[]> = {};
  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];

    for (const address of Object.keys(worksheet)) {

      if (address.startsWith("!")) {
        continue;
      }

      const cell = worksheet[address];
      // here pass by reference, this does not produce copy
      // if you want a copy
      // const cellCopy = {...worksheet[address]};

      if (
        cell &&
        isExcelDateCell(cell) &&
        typeof cell.v === "number"
      ) {

        cell.v =
          convertExcelSerialToPlainDate(cell.v)
            .toString();

        cell.t = "s";
      }
    }

    sheets[sheetName] = XLSX.utils.sheet_to_json(worksheet);
    // sheet_to_json() generates an array of objects.
    // and keys are taken from the first row.
  }

  // tabsDiv
  for (const sheetName of workbook.SheetNames) {
    const button = document.createElement("button");
    button.textContent = sheetName;
    button.addEventListener("click", () => {
      showSheet(sheetName);
    });
    tabsDiv.appendChild(button);
  }

  function showSheet(sheetName: string) {
    const rows = sheets[sheetName];
    console.log('showSheet() invoked.');
    console.log(rows);

    // pick up column labels from the first record
    const firstRow = rows[0] as Record<string, unknown>;
    const columnDefs = Object.keys(firstRow).map(col => ({
      field: col
    }));

    // clear gridDiv so that createGrid() reuses it.
    gridDiv.innerHTML = "";
    createGrid(gridDiv, {
      columnDefs,
      rowData: rows,
      defaultColDef: {
        sortable : true,
        filter   : true,
        resizable: true
      }
    });

  }

}; // end async function showDataInGrid()
