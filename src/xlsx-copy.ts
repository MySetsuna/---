import Excel from "exceljs";
import { buildOrgNameSet, buildTempFilesByOrgName } from "./util/util";
import fs from "fs";
const workbook = new Excel.Workbook();
workbook.xlsx
  .readFile(
    `${__dirname}/r.xlsx`
  )
  .then(() => {
    const [sheet1, sheet2, sheet3] = workbook.worksheets;
    // console.log(sheet2.getRows(1, sheet2.rowCount - 1)?.[1].getCell(1).value);
    // workbook.xlsx.writeFile(`${__dirname}/style-output/test.xlsx`);

    const w2ContentRows = sheet2.getRows(2, sheet2.rowCount - 1) ?? [];

    const w3ContentRows = sheet3.getRows(1, sheet3.rowCount - 1) ?? [];

    const orgNameSet = new Set<string>();
    // sheet2 数据从第三行开始
    buildOrgNameSet(orgNameSet, sheet2, 2, 1);
    // sheet3 数据从第二行开始
    buildOrgNameSet(orgNameSet, sheet3, 1, 7);
    orgNameSet.forEach((orgName, index) => {
      buildTempFilesByOrgName(orgName, __dirname);
    });
    const orgListJSON = JSON.stringify({ orgList: Array.from(orgNameSet) });
    fs.writeFileSync(`${__dirname}/orgNameList.json`, orgListJSON);
  });
