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
    document.getElementById("setup-sample").onclick = setupSample;
    document.getElementById("re-sample").onclick = resample;
  }
});

export async function setupSample(){
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

export async function resample(){
  try{
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const tables = sheet.tables.load("items/name");
      await context.sync();
  
      const sourceTableDropdown = document.getElementById("sourceTable");
      const targetTableDropdown = document.getElementById("targetTable");
      sourceTableDropdown.innerHTML = "";
      targetTableDropdown.innerHTML = "";
      // 使用 Set 追踪已添加的表名称，避免重复
      const addedTables = new Set();
      // 填充下拉菜单
      tables.items.forEach((table) => {
        if (!addedTables.has(table.name)) {
            // 添加到来源表下拉菜单
            const optionSource = document.createElement("option");
            optionSource.text = table.name;
            optionSource.value = table.name;
            sourceTableDropdown.add(optionSource);

            // 添加到目标表下拉菜单
            const optionTarget = document.createElement("option");
            optionTarget.text = table.name;
            optionTarget.value = table.name;
            targetTableDropdown.add(optionTarget);

            // 将表名称记录到 Set 中
            addedTables.add(table.name);
        }
    });
  
      // 监听按钮点击事件
      document.getElementById("interpolateBtn").addEventListener("click", async () => {
          const sourceTableName = sourceTableDropdown.value;
          const targetTableName = targetTableDropdown.value;
  
          if (!sourceTableName || !targetTableName) {
              console.error("Please select both source and target tables.");
              return;
          }
  
          await interpolateTables(sourceTableName, targetTableName);
      });
  });
  
}
catch(error){
  console.error(error);
}
}

// 插值函数
async function interpolateTables(sourceTableName, targetTableName) {
  await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const sourceTable = sheet.tables.getItem(sourceTableName);
      const targetTable = sheet.tables.getItem(targetTableName);

      // 加载表格数据
      const sourceHeader = sourceTable.getHeaderRowRange().load("values");
      const sourceBody = sourceTable.getDataBodyRange().load("values");
      const targetHeader = targetTable.getHeaderRowRange().load("values");
      const targetBody = targetTable.getDataBodyRange().load("values");

      await context.sync();

      // 提取 x/y 轴和查表值
      const sourceX = sourceHeader.values[0].slice(1).map(Number);
      const sourceY = sourceBody.values.map(row => Number(row[0]));
      const sourceValues = sourceBody.values.map(row => row.slice(1).map(Number));

      const targetX = targetHeader.values[0].slice(1).map(Number);
      const targetY = targetBody.values.map(row => Number(row[0]));

      console.log("Source Table - X Axis:", sourceX);
      console.log("Source Table - Y Axis:", sourceY);
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
      const interpolatedRows = sourceValues.map(row => interpolateRow(row, sourceX, targetX));
      console.log("Row Interpolation Result: ", interpolatedRows);

      // Step 2: 转置插值结果，对每列进行插值
      const interpolatedColumns = targetX.map((_, colIndex) =>
          interpolateColumn(interpolatedRows.map(row => row[colIndex]), sourceY, targetY)
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