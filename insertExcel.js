const ExcelJS = require("exceljs");
const ReadJson = require("./readJson");

async function manipulateExcel() {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet("sheet1");

  await workbook.xlsx.writeFile("./data/output.xlsx");
  console.log("Se creo el libro de Excel");

  const existingWorkbook = new ExcelJS.Workbook();
  await existingWorkbook.xlsx.readFile("./data/output.xlsx");
  const existingSheet = existingWorkbook.getWorksheet("sheet1");

  const readJson = new ReadJson();
  const json = readJson.getJson();

  const header = json.map((value) => addHeaders(value, ""));
  existingSheet.addRow(headersModification(header));

  await existingWorkbook.xlsx.writeFile("./data/output.xlsx");
  console.log("Se insertó el headers al libro de Excel");

  json.forEach((value) => {
    const modifiedRows = addRows(value).map((element) =>
      typeof element === "string" ? element.replace("undefined/", "") : element
    );

    const rows = addtoRowExcel(headersModification(header), modifiedRows);
    existingSheet.addRow(rows);
  });

  await existingWorkbook.xlsx.writeFile("./data/output.xlsx");
  console.log("Se insertó los statements al libro Excel");
}

manipulateExcel();
function addHeaders(resiveObject, resiveString) {
  return ReadJson.toArray(
    Object.entries(resiveObject).flatMap(([key, value]) =>
      typeof value === "object" && value !== null
        ? addHeaders(value, conditionHeadersValue(key, resiveString))
        : conditionHeadersValue(key, resiveString)
    )
  );
}
function addRows(resiveObject, resiveString) {
  return ReadJson.toArray(
    Object.entries(resiveObject).flatMap(([key, value]) =>
      typeof value === "object" && value !== null
        ? addRows(value, conditionHeadersValue(key, resiveString))
        : [conditionHeadersValue(key, resiveString), value]
    )
  );
}

function conditionHeadersValue(key, value) {
  return value == "" ? key : `${value}/${key}`;
}
function headersModification(headers) {
  const uniqueSet = new Set();
  return headers.flatMap((header) =>
    header.filter((value) => !uniqueSet.has(value) && uniqueSet.add(value))
  );
}
function addtoRowExcel(headers, rows) {
  return headers.map((element) =>
    rows.includes(element) ? rows[rows.indexOf(element) + 1] : ""
  );
}