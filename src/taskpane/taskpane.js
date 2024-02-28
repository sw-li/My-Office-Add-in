/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Assign event handlers and other initialization logic.
    document.getElementById("add-row").onclick = () => tryCatch(addRow);
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

async function addRow() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    
    const table = currentWorksheet.tables.getItemAt(0)

    if (table) {

      const column =  await table.columns.getItem(`N° d'ordre`)
      column.load('values')
      await context.sync()
      let orderValues =  column.values.slice(1)
      let currentOrder = Math.max(...orderValues)

      // Add a new row to the table
      const newRow = table.rows.add(null /* add at the end */, [
        // Provide values for each cell in the new row
        [currentOrder+1,null,null],
        // Add more arrays if you have more columns in the table
      ]);

    } else {
      console.log("No tables found on the current sheet.");
    }

    await context.sync();
  });
}



/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}