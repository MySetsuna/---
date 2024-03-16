import Excel from "exceljs";
import fs from "fs";

const orgNameSet = new Set<string>();

export const buildOrgNameSet = (
  orgNameSet: Set<string>,
  sheet: Excel.Worksheet,
  startRow: number,
  cellIndex: number
) => {
  //   const sheetOrgMap = new Map<any, Excel.Row[]>();
  sheet.getRows(startRow, sheet.rowCount - 1)?.forEach((row) => {
    const orgName = row.getCell(cellIndex);
    const orgNameStr = orgName.value as string;
    if (orgNameStr) orgNameSet.add(orgNameStr);
    //   const orgDetailArr = sheetOrgMap.get(orgName) || [];
    //   orgDetailArr.push(row);
    //   sheetOrgMap.set(orgName, orgDetailArr);
  });
  //   return { sheetOrgMap, sheetHeader: sheet.getRows(0, 1) };
};

export const buildTempFilesByOrgName = (orgName: string, root: string) => {
  const newFile = `${root}/style-output/${orgName}.xlsx`;
  fs.copyFileSync(
    `${root}/.xlsx`,
    newFile
  );
};
