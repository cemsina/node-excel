const { createExcel } = require("../report-core");
("use strict");
var Ihlal_Raporu = require("./../jsons/Ihlal_Raporu.json");
/**
    Hücre Parametreleri:

    type: "string" | "number",
    value: string | number,
    fontName: string,
    fontSize: number,
    bold: boolean,
    underline: boolean,
    italic: boolean,
    strike: boolean,
    outline: boolean,
    shadow: boolean,
    fontColor: "FFFFFF" (hex),
    fillColor: "FFFFFF" (hex),
    horizontal: "left" | "right" | "center",
    vertical: "left" | "right" | "center"
*/

/**
    Hücre Birleştirme:

    Key olarak "A1" denildiğinde sadece A1 hücresini belirtir.
    Key olarak "A1:A4" denildiğinde A1 den A4 e olan hücreleri birleştirir.
*/

let content = {
  name: Ihlal_Raporu.name,
  sheets: [
    {
      name: Ihlal_Raporu.sheets[0].name,
      cells: {
        "A1:B1": { value: Ihlal_Raporu.sheets[0].name }
      }
    },
    {
      name: Ihlal_Raporu.sheets[1].name,
      cells: {
        F1: { value: "Tarih" },
        G1: { value: Ihlal_Raporu.sheets[1].date },
        "A3:G3": {
          value: Ihlal_Raporu.sheets[1].tables[0].name,
          bold: true,
          vertical: "center",
          fontSize: 10
        }
      }
    }
  ]
};

function getRow(number) {
  switch (number) {
    case value:
      break;

    default:
      break;
  }
}

function IhlalRapor() {
  let table2Offset = 0;
  let table1Length = 3;

  let Sum1 = [0, 0, 0, 0, 0, 0, 0, 0, 0];
  let Sum2 = [0, 0, 0, 0, 0, 0, 0, 0, 0];
  for (
    let i = 0;
    i < Object.keys(Ihlal_Raporu.sheets[0].subheaders).length;
    i++
  ) {
    content.sheets[0].cells[`A${i + 2}`] = {
      bold: true,
      vertical: "center",
      fontSize: 10,
      value: Ihlal_Raporu.sheets[0].subheaders[i]
    };
    content.sheets[0].cells[`B${i + 2}`] = {
      vertical: "center",
      fontSize: 10,
      value: Ihlal_Raporu.sheets[0].values[i]
    };
  }
  table1Length++;
  for (let i = 65; i < 72; i++) {
    let cell = String.fromCharCode(i) + 4;
    content.sheets[1].cells[cell] = {
      bold: true,
      vertical: "center",
      fontSize: 10,
      value: Ihlal_Raporu.sheets[1].tables[0].headers[i - 65]
    };
    for (
      let j = 0;
      j < Object.keys(Ihlal_Raporu.sheets[1].tables[0].items).length;
      j++
    ) {
      cell = String.fromCharCode(i) + (j + 5);
      content.sheets[1].cells[cell] = {
        vertical: "center",
        fontSize: 10,
        value: Ihlal_Raporu.sheets[1].tables[0].items[j].values[i - 65]
      };
      table2Offset = j + 8;
      if (i - 65 > 2) {
        let number = Number(
          Ihlal_Raporu.sheets[1].tables[0].items[j].values[i - 65]
        );
        Sum1[i - 65] += number;
        cell = String.fromCharCode(i) + (j + 6);
        content.sheets[1].cells[cell] = {
          vertical: "center",
          fontSize: 10,
          value: Sum1[i - 65]
        };
      }
    }
  }
  table1Length++;
  cell = "A" + (table2Offset - 2);
  content.sheets[1].cells[cell] = {
    bold: true,
    vertical: "center",
    fontSize: 10,
    value: "Toplam :"
  };
  cell = `A${table2Offset}:G${table2Offset}`;
  content.sheets[1].cells[cell] = {
    bold: true,
    vertical: "center",
    fontSize: 10,
    value: Ihlal_Raporu.sheets[1].tables[1].name
  };
  table1Length += 2;
  console.log("TCL: table1Length", table1Length);

  for (let i = 65; i < 72; i++) {
    cell = String.fromCharCode(i) + (table2Offset + 1);
    content.sheets[1].cells[cell] = {
      bold: true,
      vertical: "center",
      fontSize: 10,
      value: Ihlal_Raporu.sheets[1].tables[1].headers[i - 65]
    };
    for (
      let j = 0;
      j < Object.keys(Ihlal_Raporu.sheets[1].tables[1].items).length;
      j++
    ) {
      cell = String.fromCharCode(i) + (table2Offset + 2 + j);
      content.sheets[1].cells[cell] = {
        vertical: "center",
        fontSize: 10,
        value: Ihlal_Raporu.sheets[1].tables[1].items[j].values[i - 65]
      };
      if (i - 65 > 2) {
        let number = Number(
          Ihlal_Raporu.sheets[1].tables[1].items[j].values[i - 65]
        );
        Sum2[i - 65] += number;
        cell = String.fromCharCode(i) + (j + table2Offset + 3);
        content.sheets[1].cells[cell] = {
          vertical: "center",
          fontSize: 10,
          value: Sum2[i - 65]
        };
      }
    }
  }
  cell = "A" + (table2Offset + 4);
  console.log("TCL: cell", cell);
  content.sheets[1].cells[cell] = {
    bold: true,
    vertical: "center",
    fontSize: 10,
    value: "Toplam :"
  };
  createExcel(content);
  console.log("TCL: content", content);
}

module.exports = IhlalRapor;
