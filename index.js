/*************************************************************************************************** Conversion XLSX vers JSON  */

const xlsx2json = require("xlsx2json");
const jsonfile = require("jsonfile");

const convertName = (name) => {
  let map = {
    _: " ",
    "-": `'`,
    a: "á|à|ã|â|À|Á|Ã|Â",
    e: "é|è|ê|É|È|Ê",
    i: "í|ì|î|Í|Ì|Î",
    o: "ó|ò|ô|õ|Ó|Ò|Ô|Õ",
    u: "ú|ù|û|ü|Ú|Ù|Û|Ü",
    c: "ç|Ç",
    n: "ñ|Ñ",
  };
  for (let pattern in map) {
    name = name.replace(new RegExp(map[pattern], "g"), pattern);
  }
  return name.toLowerCase();
};

let outputJson = {};

xlsx2json("./uploads/listing.xlsx").then((jsonArray) => {
  jsonArray = jsonArray[0];
  jsonArray.splice(0, 10);
  jsonArray.splice(jsonArray.length - 2);

  let cat = "";

  for (let i = 0; i < jsonArray.length; i++) {
    if (
      jsonArray[i]["A"] !== "" &&
      isNaN(parseInt(jsonArray[i]["A"])) &&
      jsonArray[i]["A"] !== "annee" &&
      jsonArray[i]["A"] !== "NV"
    ) {
      cat = jsonArray[i]["A"];
      if (outputJson[cat] === undefined) outputJson[cat] = [];
    }

    if (
      jsonArray[i]["A"] !== "" &&
      (!isNaN(parseInt(jsonArray[i]["A"])) || jsonArray[i]["A"] === "NV")
    ) {
      let newLine = {};
      newLine[convertName(jsonArray[2]["A"])] = jsonArray[i]["A"];
      newLine[convertName(jsonArray[2]["B"])] = jsonArray[i]["B"];
      newLine[convertName(jsonArray[2]["C"])] = jsonArray[i]["C"];
      newLine[convertName(jsonArray[2]["D"])] = jsonArray[i]["D"];
      newLine[convertName(jsonArray[2]["E"])] = jsonArray[i]["E"];
      newLine[convertName(jsonArray[2]["F"])] = jsonArray[i]["F"];
      newLine[convertName(jsonArray[2]["G"])] = jsonArray[i]["G"];
      newLine[convertName(jsonArray[2]["H"])] = jsonArray[i]["H"];
      outputJson[cat].push(newLine);
    }
  }

  jsonfile.writeFile("./output/listing.json", outputJson, function (err) {
    if (err) console.error(err);
  });

  /********************************************************************************************************************************/

  /********************************************************************** génération de la requête SQL pour insérer tous les vins */
  /***************************************************************** résultat dans le console.log et copier-coller dans ./sql.sql */
  let idCategories = {
    "Bordeaux rouges": 1,
    "Bordeaux blancs": 2,
    "Bourgogne rouges": 3,
    "Bourgogne blancs": 4,
    "Rhone rouges": 5,
    "Rhone blancs": 6,
    "Champagne": 7,
    "Alsace": 8,
    "Loire": 9,
    "Languedoc": 10,
    "Sud-Ouest": 11,
    "Italie": 12,
    "Espagne": 13,
    "Allemagne": 14,
    "Portugal": 15,
    "USA": 16
  };

  let sql =
    'INSERT INTO `alto`.`vin` (`idCategorie`,  `annee`,  `format`,  `stock`,  `nom`,  `appellation`,  `prix` ,  `score` ,  `conditionnement`) VALUES ';

  for (let cat in outputJson) {
    for (let vin in outputJson[cat]) {
      sql +=
        '(' +
        idCategories[cat] + ', "' +
        (outputJson[cat][vin].annee === undefined ? '' : outputJson[cat][vin].annee) + '", "' +
        (outputJson[cat][vin].format === undefined ? '': outputJson[cat][vin].format) + '", ' +
        (outputJson[cat][vin].stock === undefined ? '': outputJson[cat][vin].stock) + ', "' +
        (outputJson[cat][vin].nom === undefined ? '': outputJson[cat][vin].nom) + '", "' +
        (outputJson[cat][vin].appellation === undefined ? '': outputJson[cat][vin].appellation) + '", "' +
        (outputJson[cat][vin].prix === undefined ? '': outputJson[cat][vin].prix) + '", "' +
        (outputJson[cat][vin].score === undefined ? '': outputJson[cat][vin].score) + '", "' +
        (outputJson[cat][vin].conditionnement === undefined ? '': outputJson[cat][vin].conditionnement) +
        '"), ';
    }
  }
  sql = sql.substring(0, sql.length - 2);
  sql += ";";

console.log(sql);
});
