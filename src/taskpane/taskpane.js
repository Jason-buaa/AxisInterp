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
    document.getElementById("re-sample").onclick = demo_resample;
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
function reSample()
{

 // 原始数据
  const xAxis = [1, 3, 5, 7];  // 原始 x 轴点
  const yAxis = [0, 5, 10, 15]; // 原始 y 轴点
  const lookupTable = [
      [10, 15, 20, 25],  // 对应 y=0
      [30, 35, 40, 45],  // 对应 y=5
      [50, 55, 60, 65],  // 对应 y=10
      [70, 75, 80, 85],  // 对应 y=15
  ];
  
  // 新的 x 和 y 轴
  const newXAxis = [1, 4, 6, 7];  // 新的 x 轴点（数量比原来多）
  const newYAxis = [0, 9, 15];   // 新的 y 轴点（数量比原来少）
  
  // 工具函数：对单行数据在 x 轴方向插值
  function interpolateRow(row, xAxis, newX) {
      const interpolant = new THREE.LinearInterpolant(xAxis, row, 1); // 一维插值器
      return newX.map(x => interpolant.evaluate(x)[0]); // 对新 x 轴上的每个点插值
  }
  
  // 第一步：对每一行插值，生成新表（在新 x 轴上的值）
  const interpolatedRows = lookupTable.map(row => interpolateRow(row, xAxis, newXAxis));
  
  // 工具函数：对列数据在 y 轴方向插值
  function interpolateColumn(column, yAxis, newY) {
      const interpolant = new THREE.LinearInterpolant(yAxis, column, 1); // 一维插值器
      return newY.map(y => interpolant.evaluate(y)[0]); // 对新 y 轴上的每个点插值
  }
  
  // 第二步：对新 x 轴的每列进行插值，生成最终表（在新 y 轴上的值）
  const finalTable = newYAxis.map(newY => {
      // 获取对应的新列（逐列插值）
      return newXAxis.map((_, xIndex) => {
          const column = interpolatedRows.map(row => row[xIndex]); // 提取原列
          return interpolateColumn(column, yAxis, [newY])[0]; // 对新 Y 值插值
      });
  });
  
  // 输出结果
  console.log('Interpolated Table:', finalTable);
  
}

export async function demo_resample(){
  try{
    await Excel.run(async (context) => {
      // Step 1: 创建一个新的工作表
      const sheet = context.workbook.worksheets.add("Demo");
      sheet.activate();
  
      // 定义原始数据
      const xAxis = [1, 3, 5, 7].map(String); // 将 x 轴值转换为字符串
      const yAxis = [0, 5, 10, 15].map(String); // 将 y 轴值转换为字符串
      const lookupTable = [
          [10, 15, 20, 25],
          [30, 35, 40, 45],
          [50, 55, 60, 65],
          [70, 75, 80, 85],
      ].map(row => row.map(String)); // 将查表值转换为字符串
  
      // 创建表格 table_original
      const tableOriginal = sheet.tables.add(sheet.getRange("A1"), true);
      tableOriginal.name = "table_original";
  
      // 定义表头和添加行
      tableOriginal.getHeaderRowRange().values = [["Y/X", ...xAxis]]; // 表头
      tableOriginal.rows.add(null, yAxis.map((y, i) => [y, ...lookupTable[i]]));
  
      await context.sync(); // 确保表格创建完成后提取数据
  
      // Step 2: 从 table_original 提取数据
      const originalValues = tableOriginal.getRange().values;
      const extractedXAxis = originalValues[0].slice(1); // 提取第一行（去掉空单元格）
      const extractedYAxis = originalValues.slice(1).map(row => row[0]); // 提取第一列（去掉空单元格）
      const extractedLookupValues = originalValues.slice(1).map(row => row.slice(1)); // 提取查表值
  
      // 新的 x 和 y 轴
      const newXAxis = [1, 2, 4, 6, 7].map(String); // 将新 x 轴值转换为字符串
      const newYAxis = [0, 6, 12, 15].map(String); // 将新 y 轴值转换为字符串
  
      // 工具函数：插值逻辑
      function interpolateRow(row, xAxis, newX) {
          const interpolant = new THREE.LinearInterpolant(xAxis.map(Number), row, 1); // 将 x 轴转回数字
          return newX.map(x => interpolant.evaluate(Number(x))[0]);
      }
  
      function interpolateColumn(column, yAxis, newY) {
          const interpolant = new THREE.LinearInterpolant(yAxis.map(Number), column, 1); // 将 y 轴转回数字
          return newY.map(y => interpolant.evaluate(Number(y))[0]);
      }
  
      // 插值计算
      const interpolatedRows = extractedLookupValues.map(row => interpolateRow(row, extractedXAxis, newXAxis));
      const interpolatedTable = newYAxis.map(newY => {
          return newXAxis.map((_, xIndex) => {
              const column = interpolatedRows.map(row => row[xIndex]); // 提取列数据
              return interpolateColumn(column, extractedYAxis, [newY])[0];
          });
      });
  
      // Step 3: 创建 table_new 并插入插值数据
      const tableNew = sheet.tables.add(sheet.getRange("H1"), true);
      tableNew.name = "table_new";
  
      tableNew.getHeaderRowRange().values = [["Y/X", ...newXAxis]]; // 表头
      tableNew.rows.add(null, newYAxis.map((y, i) => [y, ...interpolatedTable[i].map(String)])); // 转换为字符串
  
      // 自动调整列宽和行高
      sheet.getUsedRange().format.autofitColumns();
      sheet.getUsedRange().format.autofitRows();
  
      // Step 4: 完成并同步
      await context.sync();
      console.log("Table 'table_new' with interpolated data created successfully!");
  });
  
  

  }
  catch(error){
    console.error(error);
  }
}