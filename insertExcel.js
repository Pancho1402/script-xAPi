const ExcelJS = require("exceljs");
const ReadJson = require("./readJson");

manipulateExcel();

async function manipulateExcel() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("sheet1");

  const readJson = new ReadJson();
  let header = [];

  for (let index = 0; index < readJson.getJson(); index++) {
    header.push(addHeaders(readJson.getJson()[i], ""));
  }
  sheet.addRow(headersModification(header));

  for (let index = 0; index < 4; index++) {
    const rows = addRows(readJson.getJson()[i]);
    const modifiedRows = rows.map((element) => {
      if (typeof element === "string") {
        return element.replace("undefined/", "");
      } else {
        return element;
      }
    });

    const rowFinally = addtoRow(headersModification(header), modifiedRows);

    sheet.addRow(rowFinally);
  }

  console.log("Datos guardados");
  await workbook.xlsx.writeFile("./data/output.xlsx");
}

function addHeaders(myObject, myString) {
  let headers = [];
  const readJson = new ReadJson();

  for (let key in myObject) {
    if (myObject.hasOwnProperty(key)) {
      const value = myObject[key];

      if (typeof value === "object" && value !== null) {
        headers.push(addHeaders(value, conditionHeadersValue(key, myString)));
      } else {
        headers.push(conditionHeadersValue(key, myString));
      }
    }
  }

  return readJson.toArray(headers);
}

function conditionHeadersValue(key, myString) {
  const condition = myString == "";
  return condition ? key : `${myString}/${key}`;
}

function headersModification(headers) {
  let finalHeaders = [];

  for (const i in headers) {
    const header = headers[i];

    for (let j = 0; j < header.length; j++) {
      if (header.indexOf(header[j]) !== header.lastIndexOf(header[j])) {
        finalHeaders.push(header[j]);
      } else if (finalHeaders.indexOf(header[j]) === -1) {
        finalHeaders.push(header[j]);
      }
    }
  }
  return finalHeaders;
}

function addtoRow(headers, rows) {
  let rowFinal = [];

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];

    for (let j = 0; j < rows.length; j++) {
      if (header == rows[j]) {
        rowFinal.push(rows[j + 1]);
        break;
      } else if (j == rows.length - 1) {
        rowFinal.push("");
      }
    }
  }

  return rowFinal;
}

function addRows(myObject, myString) {
  let rows = [];
  const readJson = new ReadJson();

  for (let key in myObject) {
    if (myObject.hasOwnProperty(key)) {
      const value = myObject[key];

      if (typeof value === "object" && value !== null) {
        rows.push(addRows(value, conditionHeadersValue(key, myString)));
      } else {
        rows.push(conditionHeadersValue(key, myString), value);
      }
    }
  }

  return readJson.toArray(rows);
}
