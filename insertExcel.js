const ExcelJS = require("exceljs");
// biblioteca para la creacion, modificacion y insercion de datos hacia el excel.
const ReadJson = require("./readJson");
// permite la lectura de los datos.

manipulateExcel();

async function manipulateExcel() {
  
  // Crear un nuevo libro de Excel y una hoja llamada "sheet1"
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("sheet1");

  // Leer el JSON utilizando la clase ReadJson
  const readJson = new ReadJson();
  const json = readJson.getJson(); // se obtienen los datos por medio del metodo getJson

  let header = [];
  
  // Crear encabezados a partir del JSON
  for (let index =0; index < json.length; index++) {
    header.push(addHeaders(json[index], ""));
  }

  // Agregar encabezados a la hoja de cálculo
  sheet.addRow(headersModification(header));
  
  // Agregar filas a la hoja de cálculo
  for (let index = 0; index < json.length; index++) {
    const rows = addRows(json[index]);
    
    //modifica los row para modificar las cadena que contengan undefined
    // ejemplo: 
    // antes: undefined/actor/objectType

    const modifiedRows = rows.map((element) => {
      if (typeof element === "string") {
        return element.replace("undefined/", "");
      } else {
        return element;
      }
    });
    
    // despues: actor/objectType

    // rowFinally almacena un vector, para ser agregar al excel
    const rowFinally = addtoRow(headersModification(header), modifiedRows);
    sheet.addRow(rowFinally); // se termina recorrer un stament, continua con el siguiente stament
  }

  // Guardar el archivo Excel
  console.log("Datos guardados");
  await workbook.xlsx.writeFile("./data/output.xlsx");
}

function addHeaders(myObject, myString) {
  // Genera encabezados recursivamente a partir de un objeto JSON
  let headers = []; // almacena una matriz de objetos
  const readJson = new ReadJson();

  for (let key in myObject) {
    if (myObject.hasOwnProperty(key)) {
      const value = myObject[key];

      if (typeof value === "object" && value !== null) {
        // se hace uso de funcion recursiva para ingresar a los objetos con el valor y obtener su key, tambien se realiza una modicacion del key
        headers.push(addHeaders(value, conditionHeadersValue(key, myString)));
      } else {

        headers.push(conditionHeadersValue(key, myString));
      }
    }
  }
  // se realiza una conversion de matriz a vector
  return readJson.toArray(headers);
}
function addRows(myObject, myString) {
  // Genera filas y encabezados recursivamente a partir de un objeto JSON
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


function conditionHeadersValue(key, myString) {
  // Condición para formatear los encabezados, 
  const condition = myString == ""; // retorna un boolean, si esta vacio retorna true sino false

  // en caso de que el boolean sea true, la key se matiene de lo contrario se modifica.
  return condition ? key : `${myString}/${key}`;
}

function headersModification(headers) { // almacena una matriz de encabezados
  // crea  un nuevo encabezado apartir de los encabezados proporcionados, esto para evitar duplicados
  let finalHeaders = [];

  for (const index in headers) {
    const header = headers[index];

    for (let j = 0; j < header.length; j++) {

      // se valida sino se encuentran encabezados similares.
      if (header.indexOf(header[j]) !== header.lastIndexOf(header[j])) {
        // si no hay duplicados se procede a insertar al nuevo encabezado 
        finalHeaders.push(header[j]);
      } else if (finalHeaders.indexOf(header[j]) === -1) {
        // si se encuentran duplicados se toma uno de los duplicados. 
        finalHeaders.push(header[j]);
      }
    }
  }
  return finalHeaders;
}

function addtoRow(headers, rows) {
  // Agrega elementos a una fila según el encabezado
  let rowFinal = [];

  for (let index = 0; index < headers.length; index++) {
    const header = headers[index];

    for (let j = 0; j < rows.length; j++) {
      // se busca en el row una cedena que sea igual a la que se encuentra en el header
      /*

        header= ["actor/name"] 
        rows = ["actor/name", "Francisco"]
      */
      if (header == rows[j]) {
        // si encuenta una similitud, se captura su siguiente valor que seria el valor del objeto.

        //"actor/name" == "actor/name"
        // rowFinal.push("Francisco");

        rowFinal.push(rows[j + 1]);
        break;
        
      } else if (j == rows.length - 1) {// permite validar en que posicion nos encontramos de rows, ya que con un else se multiplicaban los datos, y se tenia que obtener en que posicion nos encontramos para colacar vacio.
        // en caso de que no se alla encontrado, se ingresa vacio
        
        rowFinal.push("");
      }
    }
  }

  return rowFinal;
}

