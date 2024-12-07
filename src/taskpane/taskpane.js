/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
import * as THREE from 'three';
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("get-table-counts").onclick = get_table_counts;
    document.getElementById("setup_p5").onclick = setup_p5;
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
            sheetTableMap[worksheetName].push({
                name: tableName,
                worksheet: table.worksheet
            });
        });

        // Populate the tree structure
        for (const [sheetName, tables] of Object.entries(sheetTableMap)) {
            const sheetNode = document.createElement('div');
            sheetNode.innerHTML = `<span class="sheet-name">${sheetName}</span>`;
            
            const tableList = document.createElement('ul');
            tables.forEach(({ name: tableName, worksheet }) => {
                const tableNode = document.createElement('li');
                tableNode.textContent = tableName;

                // Add click event to navigate to the table
                tableNode.style.cursor = 'pointer';
                tableNode.addEventListener('click', async () => {
                    try {
                        await Excel.run(async (context) => {
                            const sheet = context.workbook.worksheets.getItem(worksheet.name);
                            const table = sheet.tables.getItem(tableName);

                            // Select the table in Excel
                            table.getRange().select();
                            await context.sync();
                        });
                    } catch (error) {
                        console.error(`Failed to navigate to table: ${error}`);
                    }
                });

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
function setup_p5()
{
  const x = [-2, -1, 0, 1, 2];
  const y = [4, 1, 0, 1, 4];
  
  const resultBuffer = new Float32Array(1);
  const interpolant = new THREE.LinearInterpolant(x, y, 1, resultBuffer);
  
  interpolant.evaluate(1.5);
  // 0.5
  console.log(resultBuffer[0]);
  
  interpolant.evaluate(1.5);
  // 2.5
  console.log(resultBuffer[0]);  

}