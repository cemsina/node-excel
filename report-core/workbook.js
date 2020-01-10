const XLSX = require('xlsx-style');
const Sheet = require('./sheet');
"use strict";

function Workbook(name,sheets_json){
    "use strict";

    var sheets = sheets_json.reduce((prev,cur)=>prev.concat(new Sheet(cur.name,cur.cells)),[]);

    var isValid = !sheets.some(c=>!c.isValid());
    var errorKey = !isValid ? sheets.find(c=>!c.isValid()).errorKey() : null;

    this.errorKey = ()=>errorKey;

    this.isValid = ()=>isValid;

    this.name = ()=>name;

    this.sheets = ()=>sheets;

    this.json = () => ({
        SheetNames: sheets.map(s=>s.name()),
        Sheets: sheets.reduce((prev,cur)=>{
            prev[cur.name()] = cur.json();
            return prev;
        },{})
    });

    this.save = ()=>{
        XLSX.writeFile(this.json(), `src/public/${name}.xlsx`, {bookType: 'xlsx', bookSST: false});
        return `${name}.xlsx`;
    };

}

module.exports = Workbook;