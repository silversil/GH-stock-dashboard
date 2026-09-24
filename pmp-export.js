/* global ExcelJS */
function createPmpWorkbook(matrix, sheetName, priceCol) {
  const book = new ExcelJS.Workbook();
  const sheet = book.addWorksheet(sheetName);
  sheet.views = [{ state: "frozen", xSplit: 1, ySplit: 1, topLeftCell: "B2", activeCell: "B2" }];
  for (let c = 3; c <= 6 && c <= priceCol; c++) sheet.getColumn(c).numFmt = "0.00####";
  sheet.getColumn(priceCol + 1).numFmt = "0.00####";
  sheet.getColumn(11).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
  sheet.addRows(matrix);
  sheet.getRow(1).font = { bold: true };
  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: matrix.length, column: matrix[0].length } };
  return book;
}
