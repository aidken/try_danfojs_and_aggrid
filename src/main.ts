import * as XLSX from "xlsx";
import { createGrid } from "ag-grid-community";

import "ag-grid-community/styles/ag-grid.css";
import "ag-grid-community/styles/ag-theme-quartz.css";

document.querySelector<HTMLDivElement>("#app")!.innerHTML = `
  <h1>CSV / Excel Viewer</h1>

  <div>this is a client side app. upload a csv/excel file and its contents is shown as a data grid.</div>

  <div>data grid is made available with <a href="https://www.ag-grid.com/"></a></div>

  <input type="file" id="fileInput" />

  <div id="tabs"></div>
  <div
    id="grid"
    class="ag-theme-quartz"
    style="height:600px;width:100%;margin-top:1rem;"
  >
  </div>
`;

const fileInput = document.getElementById("fileInput") as HTMLInputElement;

fileInput.addEventListener("change", async () => {
  const file = fileInput.files?.[0];

  if (!file) {
    return;
  }

  const buffer = await file.arrayBuffer();

  const workbook = XLSX.read(buffer);

  const sheets: Record<string, any[]> = {};

  // const rows =
  //   XLSX.utils.sheet_to_json(firstSheet);
  // console.log(rows);
  // // sheet_to_json() generates an array of objects.
  // // and keys are taken from the first row.

  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    sheets[sheetName] = XLSX.utils.sheet_to_json(worksheet);
    // sheet_to_json() generates an array of objects.
    // and keys are taken from the first row.
  }

  // tabsDiv
  const tabsDiv = document.getElementById("tabs")!;
  for (const sheetName of workbook.SheetNames) {
    const button = document.createElement("button");
    button.textContent = sheetName;
    button.addEventListener("click", () => {
      showSheet(sheetName);
    });
    tabsDiv.appendChild(button);
  }

  const gridDiv =
    document.getElementById("grid")!;

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

});
