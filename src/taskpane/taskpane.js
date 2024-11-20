/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("get-table-counts").onclick = get_table_counts;
  }
});

export async function get_table_counts() {
  try {
    await Excel.run(async (context) => {
      const tables = context.workbook.tables;
      tables.load('count,items');
      await context.sync();
      
      console.log(tables.count);
      const dropdown = document.getElementById('tableDropdown');
            
      // Populate the dropdown
      tables.items.forEach((table) => {
          const option = document.createElement('option');
          option.value = table.name;
          option.textContent = table.name;
          dropdown.appendChild(option);
      });
    });
  } catch (error) {
    console.error(error);
  }
}
