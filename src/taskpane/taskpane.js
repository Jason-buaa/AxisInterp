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
        tables.load('items/name,items/worksheet');
        await context.sync();

        // Get the table tree container and clear its content
        const tableTree = document.getElementById('tableTree');
        tableTree.innerHTML = ''; // Clear existing content

        // Create a dictionary to group tables by worksheet
        const sheetTableMap = {};

        tables.items.forEach((table) => {
            const worksheetName = table.worksheet.name;
            const tableName = table.name;

            if (!sheetTableMap[worksheetName]) {
                sheetTableMap[worksheetName] = [];
            }
            sheetTableMap[worksheetName].push(tableName);
        });

        // Populate the tree structure
        for (const [sheetName, tableNames] of Object.entries(sheetTableMap)) {
            const sheetNode = document.createElement('div');
            sheetNode.innerHTML = `<span class="sheet-name">${sheetName}</span>`;
            
            const tableList = document.createElement('ul');
            tableNames.forEach((tableName) => {
                const tableNode = document.createElement('li');
                tableNode.textContent = tableName;
                tableList.appendChild(tableNode);
            });

            sheetNode.appendChild(tableList);
            tableTree.appendChild(sheetNode);
        }
    });
  } catch (error) {
    console.error(error);
  }
}
