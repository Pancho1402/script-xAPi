const fs = require("node:fs");

class ReadJson {
  constructor() {}

  getJson() {
    const response = fs.readFileSync("./data/prueba-01.json", "utf8");
    const json = JSON.parse(response);
    return json;
  }

  toArray(myArray) {
    let returnArray = [];
    for (const i in myArray) {
      const objectOf = myArray[i];

      if (typeof objectOf === "object" && objectOf !== null) {
        for (const j in objectOf) {
          const value = objectOf[j];
  
          returnArray.push(value);
        }
      } else {
        returnArray.push(objectOf);
      }
    }

    return returnArray;
  }
}

module.exports = ReadJson;
