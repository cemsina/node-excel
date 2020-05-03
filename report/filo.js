const {createExcel, toCell} = require('../report-core'); 
"use strict";

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

function FiloRapor(){
    "use strict";

    function createKapakCells(){
        const cells = {
            "A1":{value:""}
        };
        return cells;
    }

    function appendType(cells, title, startRow, startCol, types){
        cells[toCell(startRow,startCol)] = {
            value: title,bold:true, fontSize:12, horizontal: "center"
        };

        var totalCount = types.reduce((prev,cur)=> prev + cur.count, 0);

        cells[toCell(startRow,startCol+1)] = {
            value: totalCount,bold:true, fontSize:12, horizontal: "center"
        };

        var cursor = {row:startRow+1, col: startCol};
        types.forEach(t=>{
            cells[toCell(cursor.row,cursor.col)] = {
                value: t.typename, fontSize:11, horizontal: "center"
            };
            cells[toCell(cursor.row,cursor.col+1)] = {
                value: t.count, fontSize:11, horizontal: "center"
            };
            cursor.row++;
        });

        return cursor;
    }

    function createFiloCells(){
        const filoCells = {
            "A1:B1":{value: "İş Makineleri", bold:true, fontSize:14, horizontal: "center"},
            "C1:D1":{value: "Kamyonlar", bold:true, fontSize:14, horizontal: "center"},
            "G1:I1":{value: "Sürücü Listesi", bold:true, fontSize:14, horizontal: "center"},
            "A2":{value: "Makine Tipi", bold:true, fontSize:12, horizontal: "center"},
            "B2":{value: "Adet", bold:true, fontSize:12, horizontal: "center"},
            "C2":{value: "Kamyon Tipi", bold:true, fontSize:12, horizontal: "center"},
            "D2":{value: "Adet", bold:true, fontSize:12, horizontal: "center"}
        };

        var eks_types = [
            {typename: "Eks. Alt Tip 1", count: 2},
            {typename: "Eks. Alt Tip 2", count: 4},
            {typename: "Eks. Alt Tip 3", count: 5}
        ];

        var cursor = appendType(filoCells, "Ekskavatör", 3, 1, eks_types);

        var loader_types = [
            {typename: "Load. Alt Tip 1", count: 1},
            {typename: "Load. Alt Tip 2", count: 2},
            {typename: "Load. Alt Tip 3", count: 3}
        ];

        appendType(filoCells, "Loader", cursor.row, cursor.col, loader_types);

        return filoCells;
    }

    this.data = {
        name: "ornek_filo",
        sheets: [
            {
                name: "Kapak",
                cells: createKapakCells()
            },
            {
                name: "Filo",
                cells: createFiloCells()
            }
        ]
    };
    
    this.save = (filename) => {
        this.data.name = filename;
        createExcel(this.data);
    }
    
}

module.exports = FiloRapor;