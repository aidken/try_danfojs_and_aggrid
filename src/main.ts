import * as XLSX from "xlsx";
import { createGrid } from "ag-grid-community";

import "ag-grid-community/styles/ag-grid.css";
import "ag-grid-community/styles/ag-theme-quartz.css";

document.querySelector<HTMLDivElement>("#app")!.innerHTML = `
  <h1>CSV / Excel Viewer</h1>

  <input type="file" id="fileInput" />

  <div
    id="grid"
    class="ag-theme-quartz"
    style="height:600px;width:100%;margin-top:1rem;"
  >
  </div>
`;

const fileInput = document.getElementById("fileInput") as HTMLInputElement;
const output = document.getElementById("output") as HTMLPreElement;

fileInput.addEventListener("change", async () => {
  const file = fileInput.files?.[0];

  if (!file) {
    return;
  }

  const buffer = await file.arrayBuffer();

  const workbook = XLSX.read(buffer);

  const firstSheet =
    workbook.Sheets[workbook.SheetNames[0]];

  const rows =
    XLSX.utils.sheet_to_json(firstSheet);
  console.log(rows);
  // sheet_to_json() generates an array of objects.
  // and keys are taken from the first row.

  const firstRow = rows[0] as Record<string, unknown>;
  // as Record<string, unknown>: type assertion
  // this does not change value, but does change the way the compiler sees it.

  const columnDefs:Array<object> =
  Object.keys(firstRow).map(col => (
      { field: col }
    ));

  console.log(columnDefs);
  // [ {field: "date"}, {field: "order"}, ...]

  const gridDiv =
    document.getElementById("grid")!;

  createGrid(gridDiv, {
    columnDefs,
    rowData      : rows,
    defaultColDef: {
      sortable : true,
      filter   : true,
      resizable: true
    }
  });
  console.log(rows);

  output.textContent =
    JSON.stringify(rows.slice(0, 5), null, 2);
});
