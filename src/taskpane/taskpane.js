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
    document.getElementById("re-sample").onclick = resampleNew;
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
    const xAxis = [1, 3, 5, 7]; // 将 x 轴值转换为字符串
    const yAxis = [0, 5, 10, 15]; // 将 y 轴值转换为字符串
    const lookupTable = [
        [10, 15, 20, 25],
        [30, 35, 40, 45],
        [50, 55, 60, 65],
        [70, 75, 80, 85],
    ]; // 将查表值转换为字符串

    // 动态计算表格范围
    const columnCount = xAxis.length + 1; // 列数 = xAxis + 表头
    const rowCount = yAxis.length+1;
    const tableRange = sheet.getRangeByIndexes(0, 0, 1, columnCount); // 起始点 (0,0)，动态大小

    // 创建 tableOriginal
    const tableOriginal = sheet.tables.add(tableRange, true);
    tableOriginal.name = "table_original";

    // 设置表头，将 xAxis 转换为文本
    tableOriginal.getHeaderRowRange().values = [["Y/X", ...xAxis]]; // 表头

    // 添加数据行，确保 yAxis 和查表值都转换为文本
    tableOriginal.rows.add(null, yAxis.map((y, i) => [y, ...lookupTable[i]])); // 数据行

    await context.sync(); // 确保表格创建完成后提取数据

    /// Step 2: 提取数据
    // 获取表头
    const headerRange = tableOriginal.getHeaderRowRange();
    headerRange.load("values");

    // 获取表体
    const bodyRange = tableOriginal.getDataBodyRange();
    bodyRange.load("values");

    // 获取第一列数据（y 轴）
    const yAxisColumnRange = tableOriginal.columns.getItemAt(0).getDataBodyRange();
    yAxisColumnRange.load("values");

    // 获取第一行数据（示例）
    const firstRowRange = tableOriginal.rows.getItemAt(0);
    firstRowRange.load("values");

    // 同步以确保数据加载到变量中
    await context.sync();

    // 提取数据
    const headerValues = headerRange.values; // 表头数据
    // 剔除表体的第一列数据
    const bodyValues = bodyRange.values.map(row => row.slice(1)); // 移除每行的第一列
    const yAxisValues = yAxisColumnRange.values; // y 轴数据
    const firstRowValues = firstRowRange.values; // 第一行数据

    console.log("Header Values:", headerValues);
    console.log("Body Values:", bodyValues);
    console.log("Y Axis Values:", yAxisValues);
    console.log("First Row Values:", firstRowValues);

    // 在工作表中写入提取的数据以供验证
    sheet.getRange("G1:G1").values = [["Extracted Results"]];
    sheet.getRange("G2").values = [["Header"]];
    sheet.getRange("G3:K3").values = headerValues;
    sheet.getRange("G5").values = [["Body"]];
    sheet.getRange(`H6:${String.fromCharCode(72 + columnCount - 2)}${5 + rowCount - 1}`).values = bodyValues;
    sheet.getRange("G12").values = [["Y Axis"]];
    sheet.getRange(`H12:H${11 + yAxisValues.length}`).values = yAxisValues;

    await context.sync();
    });
  
  

  }
  catch(error){
    console.error(error);
  }
}

export async function resampleNew(){
  try{
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
  
      // 获取当前工作表的所有表格名称
      const tables = sheet.tables.load("items/name");
      await context.sync();
  
      // 列出表格供用户选择
      const tableNames = tables.items.map(t => t.name);
      if (tableNames.length < 2) {
          console.error("Need at least two tables: original and target.");
          return;
      }
      console.log("Available Tables: ", tableNames);
  
      // 手动指定原始和目标表格（也可以通过界面实现）
      const originalTableName = tableNames[0]; // 假设第一个为原始表格
      const targetTableName = tableNames[1];   // 假设第二个为目标表格
  
      const originalTable = sheet.tables.getItem(originalTableName);
      const targetTable = sheet.tables.getItem(targetTableName);
  
      console.log("Selected Original Table: ", originalTableName);
      console.log("Selected Target Table: ", targetTableName);
  
      // 加载原始和目标表格数据
      const originalHeader = originalTable.getHeaderRowRange().load("values");
      const originalBody = originalTable.getDataBodyRange().load("values");
      const targetHeader = targetTable.getHeaderRowRange().load("values");
      const targetBody = targetTable.getDataBodyRange().load("values");
  
      await context.sync();
  
      // 提取原始和目标表格的 x/y 轴和查表值
      const originalX = originalHeader.values[0].slice(1).map(Number);
      const originalY = originalBody.values.map(row => Number(row[0]));
      const originalValues = originalBody.values.map(row => row.slice(1).map(Number));
  
      const targetX = targetHeader.values[0].slice(1).map(Number);
      const targetY = targetBody.values.map(row => Number(row[0]));
  
      console.log("Original Table - X Axis:", originalX);
      console.log("Original Table - Y Axis:", originalY);
      console.log("Target Table - X Axis:", targetX);
      console.log("Target Table - Y Axis:", targetY);
  
      // 工具函数：插值逻辑
      function interpolateRow(row, xAxis, newX) {
          const interpolant = new THREE.LinearInterpolant(xAxis.map(Number), row, 1);
          return newX.map(x => interpolant.evaluate(Number(x))[0]);
      }
  
      function interpolateColumn(column, yAxis, newY) {
          const interpolant = new THREE.LinearInterpolant(yAxis.map(Number), column, 1);
          return newY.map(y => interpolant.evaluate(Number(y))[0]);
      }
  
      // Step 1: 对原始表格进行列插值，获取中间插值结果
      const interpolatedRows = originalValues.map(row => interpolateRow(row, originalX, targetX));
      console.log("Row Interpolation Result: ", interpolatedRows);
  
      // Step 2: 转置插值结果，对每列进行插值
      const interpolatedColumns = targetX.map((_, colIndex) =>
          interpolateColumn(interpolatedRows.map(row => row[colIndex]), originalY, targetY)
      );
  
      // Step 3: 转置插值结果，恢复为行方向
      const finalValues = interpolatedColumns[0].map((_, rowIndex) =>
          interpolatedColumns.map(column => column[rowIndex])
      );
  
      console.log("Final Interpolated Values: ", finalValues);
  
      // 更新目标表格
      targetBody.values = targetY.map((y, rowIndex) => [y, ...finalValues[rowIndex]]);
      await context.sync();
  
      console.log("Interpolation complete and values updated in target table.");
  });
  
}
catch(error){
  console.error(error);
}
}