const fs = require("node:fs");
class ReadJson {
  constructor() {}

  getJson() {
    const response = fs.readFileSync("./data/prueba-01.json", "utf8");
    return JSON.parse(response);
  }

  static toArray(resiveMatrix) {
    return resiveMatrix.flatMap(element =>
      Array.isArray(element) ? element : 
      typeof element === "object" && element !== null ? Object.values(element) : [element]
    );
  }
  
}

module.exports = ReadJson;