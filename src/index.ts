import xlsx, { WorkSheet } from "node-xlsx";
import fs from "fs";
import { get } from "http";
// Or var xlsx = require('node-xlsx').default;

// Parse a file
const workSheetsFromFile = xlsx.parse(
  `${__dirname}/resource/01-研发效能部2024年1&2月货币化计费0315(同步版产品财管)-V1.xlsx`
);

const [sheet1, sheet2, sheet3] = workSheetsFromFile;

// const data = [
//   [1, 2, 3],
//   [true, false, null, "sheetjs"],
//   ["foo", "bar", new Date("2014-02-19T14:30Z"), "0.3"],
//   ["baz", null, "qux"],
// ];
// var buffer = xlsx.build([{ name: "mySheetName", data: data, options: {} }]); // Returns a buffer
// fs.writeFileSync(`${__dirname}/test.xlsx`, buffer);

const orgNameSet = new Set<string>();

const getSheetDataGroupMap = (sheet: Omit<WorkSheet, "options">, cellIndex) => {
  const sheetOrgMap = new Map<any, unknown[][]>();
  sheet.data.forEach((row, rowIndex) => {
    if (rowIndex) {
      const orgName = row[cellIndex];
      orgNameSet.add(orgName as string);
      const orgDetailArr = sheetOrgMap.get(orgName) || [];
      orgDetailArr.push(row);
      sheetOrgMap.set(orgName, orgDetailArr);
    }
  });
  return { sheetOrgMap, sheetHeader: sheet.data[0] };
};

const buildOrgGroupSheetMap = (
  sheetOrgMap: Map<any, unknown[][]>,
  sheetHeader: unknown[],
  sheetName: string
) => {
  const sheetMap: Map<any, WorkSheet<unknown>> = new Map();
  Array.from(sheetOrgMap.entries()).forEach(([orgName, orgDetailArr]) => {
    const sheetData: unknown[][] = [sheetHeader, ...orgDetailArr];
    sheetMap.set(orgName, { name: sheetName, data: sheetData, options: {} });
  });
  return sheetMap;
};

const buildSheetByOrgGroupMap = (
  orgNameSet: Set<string>,
  sheetOrgGroupMap2: Map<any, WorkSheet<unknown>>,
  sheetOrgGroupMap3: Map<any, WorkSheet<unknown>>
) => {
  orgNameSet.forEach((orgName) => {
    var buffer = xlsx.build([
      { ...sheet1, options: {} },
      sheetOrgGroupMap2.get(orgName) || {
        name: sheet2.name,
        data: [sheet2Header],
        options: {},
      },
      sheetOrgGroupMap3.get(orgName) || {
        name: sheet3.name,
        data: [sheet3Header],
        options: {},
      },
    ]); // Returns a buffer
    fs.writeFileSync(`${__dirname}/output/${orgName}.xlsx`, buffer);
  });
};

const { sheetOrgMap: sheet2OrgDetailMap, sheetHeader: sheet2Header } =
  getSheetDataGroupMap(sheet2, 0);
const { sheetOrgMap: sheet3OrgDetailMap, sheetHeader: sheet3Header } =
  getSheetDataGroupMap(sheet3, 6);

const sheet2OrgGroupMap = buildOrgGroupSheetMap(
  sheet2OrgDetailMap,
  sheet2Header,
  sheet2.name
);

const sheet3OrgGroupMap = buildOrgGroupSheetMap(
  sheet3OrgDetailMap,
  sheet3Header,
  sheet3.name
);


buildSheetByOrgGroupMap(orgNameSet, sheet2OrgGroupMap, sheet3OrgGroupMap);