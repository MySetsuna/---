import Excel from "exceljs";
import fs from "fs";
const { orgList } = require("./orgNameList.json");
const a = async () => {
  for (const orgName of orgList) {
    // const newFile = `${__dirname}/style-output/${orgName}.xlsx`;
    const newFile = `${__dirname}/style-output/${orgName}.xlsx`;
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(newFile);
    const [sheet1, sheet2, sheet3] = workbook.worksheets;
    const w2ContentRows = sheet2.getRows(0, sheet2.rowCount + 1) ?? [];
    const w3ContentRows = sheet3.getRows(0, sheet3.rowCount + 1) ?? [];
    console.log(`${orgName},过滤中`);

    for (let index = w2ContentRows.length; index > 1; index--) {
      const row = w2ContentRows[index];
      const curRowOrg = row?.getCell(1).value?.toString();

      if (curRowOrg && curRowOrg !== orgName) {
        sheet2.spliceRows(row.number, 1);
      } else if (curRowOrg && curRowOrg === orgName) {
        row.eachCell((cell) => {
          if (cell.formula) {
            cell.value = cell.result;
          }
        });
      }
    }
    console.log("sheet2 done");
    for (let index = w3ContentRows.length; index > 1; index--) {
      const row = w3ContentRows[index];
      const curRowOrg = row?.getCell(7).value?.toString();

      if (curRowOrg && curRowOrg !== orgName) {
        // 通过拼接的方式删除行
        sheet3.spliceRows(row.number, 1);
      } else if (curRowOrg && curRowOrg === orgName) {
        // 注意! 这里是遍历每个单元格, 如果有公式, 则将公式的结果赋值给单元格.如果不这样做, 单元格公式的行数参数是不更新的,导致公式报错
        row.eachCell((cell) => {
          if (cell.formula) {
            cell.value = cell.result;
          }
        });
      }
    }
    console.log("sheet3 done");
    workbook.xlsx.writeFile(`${__dirname}/formula-output/${orgName}.xlsx`);
    console.log(`${orgName},已保存`);
  }
};
a();
