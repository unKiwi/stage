/*********************************************************************
** ---------------------- Copyright notice ---------------------------
** This source code is part of the EVASoft project
** It is property of Alain Boute Ingenierie - www.abing.fr and is
** distributed under the GNU Public Licence version 2
** Commercial use is submited to licencing - contact eva@abing.fr
** -------------------------------------------------------------------
**        File : main.js
** Description : This file contain th configuration of the documents generator
**      Author : Luc TORRES
**     Created : Apr 2022
*********************************************************************/
global.config = {
    extention: ["ods", "odt"],
    path: {
        default: path.resolve('.') + "/",
        template: {
            fileName: "TableStats",
        },
        unziped: {
            dirName: "unziped_",
        }
    },
    spreadSheetTemplate: {
        tableau: {
            nbHeadRow: 3,
            nbBodyRow: 5,
            nbFootRow: 2,
            nbHeadCol: 1,
            nbBodyCol: 5,
            nbFootCol: 2,
        },
        graph: {
            pie: 1,
            bar: 2,
        },
    },
}

// define path
let globalTemplatePath = "/home/EVA/web/templates/";
let localTemplatePath = path.join(config.path.default, "../templates/");
config.extention.forEach(ext => {
    // template
    let fileName = `${config.path.template.fileName}.${ext}`;
    let filePath = path.join(localTemplatePath, fileName);
    config.path.template[ext] = fs.existsSync(filePath) ?
        filePath :  
        globalTemplatePath + fileName;

    // unziped
    config.path.unziped[ext] = config.path.default + config.path.unziped.dirName + ext + "/";
});

// parse dumpfmt.txt
let lsSheetConfig = fs.readFileSync("./dumpfmt.txt", {encoding:'latin1', flag:'r'});
lsSheetConfig = lsSheetConfig.split(os.type() == "Windows_NT" ? "\r\n\r\n" : "\n\n");
lsSheetConfig.pop();

lsSheetConfig = lsSheetConfig.map(sheetConfig => {
    let lsLine = sheetConfig.split(os.type() == "Windows_NT" ? "\r\n" : "\n");
    lsLine = lsLine.map(line => line.split("\t"));

    let dataFile = lsLine.shift();
    let displayParams = lsLine.pop();
    let lsTable = lsLine;

    dataFile = {
        fileName: dataFile[1],
        IdObj: dataFile[2],
        GRAPH_SEL: dataFile[3],
    }
    let tableauConfig = config.spreadSheetTemplate.tableau;
    let xIndexEndTable = tableauConfig.nbHeadCol;
    let yIndexEndTable = 0;
    lsTable = lsTable.map(table => {
        let result = {
            rowTitle: table[2],
            colTitle: table[8],
            dataTitle: table[14],

            xIndexSheet: table[7] * 1,
            yIndexSheet: table[1] * 1,

            nbBodyRow: table[5] * 1,
            nbRow: table[6] * 1,

            nbBodyCol: table[11] * 1,
            nbCol: table[12] * 1,
        }

        xIndexEndTable += result.nbCol;
        yIndexEndTable += tableauConfig.nbHeadRow + result.nbRow;
        result.xIndexEndTable = xIndexEndTable;
        result.yIndexEndTable = yIndexEndTable;

        return result;
    });
    displayParams = {
        XL_LBL_WIDTH: displayParams[1],
        XL_COL_WIDTH: displayParams[2],
        XL_COL_ORIENT: displayParams[3],
        XL_COL_HEIGHT: displayParams[4],
    }
    
    return {
        dataFile: dataFile,
        lsTable: lsTable,
        displayParams: displayParams,
    }
});

global.config.lsSheet = lsSheetConfig;