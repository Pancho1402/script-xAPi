const ExcelJS = require("exceljs");
// biblioteca para la creacion, modificacion y insercion de datos hacia el excel.
const ReadJson = require("./readJson");
// permite la lectura de los datos.

async function manipulateExcel() {
  // Crear un nuevo libro de Excel y una hoja llamada "sheet1"
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet("sheet1");

  // Guardar el libro de Excel en un archivo
  await workbook.xlsx.writeFile("./data/output.xlsx");
  console.log("Se creo el libro de Excel");

  // Ahora que el archivo Excel está creado, puedes abrirlo y agregar filas si es necesario
  const existingWorkbook = new ExcelJS.Workbook();
  await existingWorkbook.xlsx.readFile("./data/output.xlsx");
  const existingSheet = existingWorkbook.getWorksheet("sheet1");

  // Leer el JSON utilizando la clase ReadJson
  const readJson = new ReadJson();
  const json = readJson.getJson(); // se obtienen los datos por medio del metodo getJson

  let header = [];

  // Crear encabezados a partir del JSON
  for (let value of json) {
    header.push(addHeaders(value, ""));
  }
  // Agregar encabezados a la hoja de cálculo
  existingSheet.addRow(headersModification(header));

  await existingWorkbook.xlsx.writeFile("./data/output.xlsx");
  console.log("Se insertó el headers al libro de Excel");

  // Agregar filas a la hoja de cálculo
  for (let value of json) {
    const rows = addRows(value);

    // Modificar las filas para eliminar las cadenas que contengan "undefined/"
    const modifiedRows = rows.map((element) => {
      if (typeof element === "string") {
        return element.replace("undefined/", "");
      } else {
        return element;
      }
    });

    // rowFinally almacena un vector, para ser agregar al excel
    const rowFinally = addtoRow(headersModification(header), modifiedRows);
    existingSheet.addRow(rowFinally);
  }

  // Guardar el archivo Excel actualizado
  await existingWorkbook.xlsx.writeFile("./data/output.xlsx");
  console.log("Se insertó los statements al libro Excel");
}

// Llama a la función principal
manipulateExcel();

function addHeaders(myObject, myString) {
  // Genera encabezados recursivamente a partir de un objeto JSON
  let headers = []; // almacena una matriz de objetos
  const readJson = new ReadJson();

  for (const [key, value] of Object.entries(myObject)) {
    if (typeof value === "object" && value !== null) {
      // se hace uso de funcion recursiva para ingresar a los objetos con el valor y obtener su key, tambien se realiza una modicacion del key
      headers.push(addHeaders(value, conditionHeadersValue(key, myString)));
    } else {
      headers.push(conditionHeadersValue(key, myString));
    }
  }
  // se realiza una conversion de matriz a vector
  return readJson.toArray(headers);
}
function addRows(myObject, myString) {
  const readJson = new ReadJson();
  let rows = [];

  for (const [key, value] of Object.entries(myObject)) {
    if (typeof value === "object" && value !== null) {
      rows.push(addRows(value, conditionHeadersValue(key, myString)));
    } else {
      rows.push(conditionHeadersValue(key, myString), value);
    }
  }

  return readJson.toArray(rows);
}

function conditionHeadersValue(key, myString) {
  // Condición para formatear los encabezados,
  const condition = myString == ""; // retorna un boolean, si esta vacio retorna true sino false

  // en caso de que el boolean sea true, la key se matiene de lo contrario se modifica.
  return condition ? key : `${myString}/${key}`;
}

function headersModification(headers) {
  // almacena una matriz de encabezados
  // crea  un nuevo encabezado apartir de los encabezados proporcionados, esto para evitar duplicados
  let finalHeaders = [];

  for (const header of headers) {
    for (const value of header) {
      // se valida sino se encuentran encabezados similares.
      if (header.indexOf(value) !== header.lastIndexOf(value)) {
        // si no hay duplicados se procede a insertar al nuevo encabezado
        finalHeaders.push(value);
      } else if (finalHeaders.indexOf(value) === -1) {
        // si se encuentran duplicados se toma uno de los duplicados.
        finalHeaders.push(value);
      }
    }
  }
  return finalHeaders;
}

function addtoRow(headers, rows) {
  // Agrega elementos a una fila según el encabezado
  const rowFinal = headers.map((element) => {
    if (rows.includes(element)) {
      // si encuenta una similitud, se captura su siguiente valor que seria el valor del objeto.

      //"actor/name" == "actor/name"
      // rowFinal.push("Francisco");
      const index = rows.indexOf(element);
      return rows[index + 1];
    } else {
      // permite validar en que posicion nos encontramos de rows, ya que con un else se multiplicaban los datos, y se tenia que obtener en que posicion nos encontramos para colacar vacio.
      // en caso de que no se alla encontrado, se ingresa vacio
      return "";
    }
  });

  return rowFinal;
}
