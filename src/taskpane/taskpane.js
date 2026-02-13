/* global Office, Excel */
const statusEl = () => document.getElementById("status");

Office.onReady(() => {
  const applyBtn = document.getElementById("applyBtn");
  applyBtn.addEventListener("click", runNormalization);
});

async function runNormalization() {
  try {
    await Excel.run(async (context) => {
      const selected = context.workbook.getSelectedRange();
      selected.load(["values", "rowCount", "columnCount", "address"]);
      await context.sync();

      if (selected.rowCount < 2 || selected.columnCount < 2) {
        statusEl().textContent = "Selecciona al menos 2 filas y 2 columnas.";
        return;
      }

      const sourceValues = selected.values;
      const headers = sourceValues[0].map((h) => String(h));
      const body = sourceValues.slice(1);

      const normalized = body.map((row) => {
        return row.map((cell, index) => {
          if (index === 0) return String(cell);
          const numeric = Number(cell);
          return Number.isNaN(numeric) ? 0 : numeric;
        });
      });

      const outputSheet = context.workbook.worksheets.getActiveWorksheet();
      const startCell = outputSheet.getRange("H1");
      const tableData = [headers, ...normalized];
      const outRange = startCell.getResizedRange(tableData.length - 1, headers.length - 1);
      outRange.values = tableData;
      outRange.format.autofitColumns();

      const table = outputSheet.tables.add(outRange, true);
      table.name = `ThinkCellStyle_${Date.now()}`;
      table.style = "TableStyleMedium2";

      statusEl().textContent = `Datos procesados desde ${selected.address} y copiados a H1.`;
      await context.sync();
    });
  } catch (error) {
    statusEl().textContent = `Error: ${error.message}`;
  }
}
