const fs = require("node:fs");

class ReadJson {
  constructor() {}

  getJson() {
    // se obtiene el json como String
    const response = fs.readFileSync("./data/prueba-01.json", "utf8");
    return JSON.parse(response); // se realiza una conversion de string a objeto JSON
  }

  toArray(myArray) { // permite la conversion de una matriz a un vector
    let returnArray = [];

    for (const i in myArray) {
      const objectOf = myArray[i];
      // si encuentra que hay una matriz o un objeto se realiza un iteracion y se ingresa al vector
      if (typeof objectOf === "object" && objectOf !== null) {
        for (const j in objectOf) {
          const value = objectOf[j];
  
          returnArray.push(value);
        }
      } else {
        // sino solo se ingresa el valor al vector
        returnArray.push(objectOf);
      }
    }

    return returnArray;
  }
}

module.exports = ReadJson;
