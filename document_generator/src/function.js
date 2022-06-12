/*********************************************************************
** ---------------------- Copyright notice ---------------------------
** This source code is part of the EVASoft project
** It is property of Alain Boute Ingenierie - www.abing.fr and is
** distributed under the GNU Public Licence version 2
** Commercial use is submited to licencing - contact eva@abing.fr
** -------------------------------------------------------------------
**        File : function.js
** Description : This file contain the functions of the documents generator
**      Author : Luc TORRES
**     Created : Apr 2022
*********************************************************************/
global.getHtmlTable = (fileName) => {
    // parse file
    let htmlStr = fs.readFileSync("./" + fileName, { encoding: 'latin1', flag: 'r' });
    let table = new jsdom.JSDOM(htmlStr).window.document.getElementsByTagName("table")[0];

    let archTable = [];
    let lsTr = table.getElementsByTagName("tr");
    for (let i = 0; i < lsTr.length; i++) {
        const tr = lsTr[i];

        let lsTd = tr.getElementsByTagName("td");
        if (lsTd.length == 0) continue;

        let rowData = [];
        for (let j = 0; j < lsTd.length; j++) {
            rowData.push(lsTd[j].textContent.trim());
        }
        archTable.push(rowData)
    }

    return archTable;
}

global.cleanSpace = () => {
    // empty directory
    config.extention.forEach(ext => {
        fs.rmSync(config.path.default + config.path.unziped.dirName + ext + "/", { recursive: true, force: true });
    });

    fs.rmSync("/home/EVA/web/docs/demo-social/32423/TableStatsRes.xlsx", { recursive: true, force: true });
    // fs.rmSync("/home/EVA/web/docs/demo-social/32423/TableStatsRes.ods", { recursive: true, force: true });
    fs.rmSync("/home/EVA/web/docs/demo-social/32423/TableStatsNormalize.ods", { recursive: true, force: true });
}

global.copydirSync = (src, dest) => {
    fs.mkdirSync(dest, { recursive: true });
    let entries = fs.readdirSync(src, { withFileTypes: true });

    for (let entry of entries) {
        let srcPath = path.join(src, entry.name);
        let destPath = path.join(dest, entry.name);

        entry.isDirectory() ? copydirSync(srcPath, destPath) : fs.copyFileSync(srcPath, destPath);
    }
}

global.zip = (action, source, destination) => {
    let lsCommand = {
        Windows_NT: {
            zip: `..\\..\\zip.exe -0r ..\\..\\output\\TableStatsRes.ods *`,
            unzip: ".\\unzip.exe .\\template\\TableStats.ods -d .\\template\\unziped\\",
        },
        Linux: {
            zip: `zip -9r ${destination} *`,
            unzip: `unzip ${source} -d ${destination}`,
        },
    };

    let condition = action == "zip";
    if (condition) process.chdir(source);
    let exec = execSync(lsCommand[os.type()][action]);
    if (condition) process.chdir(config.path.default);
    return exec;
}

global.addRow = (sheet, cloneIndex, index) => {
    let lsRow = sheet.getElementsByTagName("table:table-row");
    // clone cloneIndex line
    let beforeLastRow = lsRow[cloneIndex - 1].cloneNode(true);
    // add it to the index
    lsRow[index - 2].insertAdjacentElement('afterend', beforeLastRow);
}

global.rmRow = (sheet, index) => {
    sheet.getElementsByTagName("table:table-row")[index - 1].remove();
}

global.addCol = (sheet, cloneIndex, index) => {
    let lsRow = sheet.getElementsByTagName("table:table-row");
    for (let i = 0; i < lsRow.length; i++) {
        let repeated = lsRow[i].getElementsByTagName("table:covered-table-cell")[0];
        if (repeated) {
            // row have fusionned cells
            let space = repeated.getAttribute("table:number-columns-repeated");
            repeated.setAttribute("table:number-columns-repeated", space * 1 + 1);
            repeated.previousElementSibling.setAttribute("table:number-columns-spanned", space * 1 + 2);
        }
        else {
            // add element
            let newCol = lsRow[i].children[cloneIndex - 1].cloneNode(true);
            lsRow[i].children[index - 2].insertAdjacentElement("afterend", newCol);
        }
    }
    // update column style
    let lsColStyle = sheet.getElementsByTagName("table:table-column");
    let lastColStyle = lsColStyle[lsColStyle.length - 1];
    newWidth = lastColStyle.getAttribute("table:number-columns-repeated") * 1 + 1;
    lastColStyle.setAttribute("table:number-columns-repeated", newWidth);


    // let lsColStyle = sheet.getElementsByTagName("table:table-column");
    // lsColStyle = Array.from(lsColStyle);
    // lsColStyle = lsColStyle.map(col => {
    //     return {
    //         col: col,
    //         width: col.getAttribute("table:number-columns-repeated") * 1 || 1
    //     }
    // });

    // // find new col style
    // let width = 0;
    // let i = 0;
    // while (i < lsColStyle.length) {
    //     width += lsColStyle[i].width;
    //     if (cloneIndex - 1 < width) break;

    //     i++;
    // }

    // let currentIndex = 0;
    // for (let i = 0; i < lsColStyle.length; i++) {
    //     let width = lsColStyle[i].getAttribute("table:number-columns-repeated") * 1 || 1;
    //     currentIndex += width;
    //     if (index <= currentIndex) {
    //         lsColStyle[i].setAttribute("table:number-columns-repeated", width + 1);
    //         break;
    //     }
    // }
}

global.rmCol = (sheet, index) => {
    let lsRow = sheet.getElementsByTagName("table:table-row");
    for (let i = 0; i < lsRow.length; i++) {
        let repeated = lsRow[i].getElementsByTagName("table:covered-table-cell")[0];
        if (repeated) {
            // row have fusionned cells
            let space = repeated.getAttribute("table:number-columns-repeated");
            repeated.setAttribute("table:number-columns-repeated", space - 1);
            repeated.previousElementSibling.setAttribute("table:number-columns-spanned", space);
        }
        else {
            // delete element
            lsRow[i].childNodes[index - 1].remove();
        }
    }
    // update column style
    let lsColStyle = sheet.getElementsByTagName("table:table-column");
    let currentIndex = 0;
    for (let i = 0; i < lsColStyle.length; i++) {
        let width = lsColStyle[i].getAttribute("table:number-columns-repeated") || 1;
        currentIndex += width;
        if (index <= currentIndex) {
            lsColStyle[i].setAttribute("table:number-columns-repeated", width - 1);
            break;
        }
    }
}

global.setValue = (sheet, x, y, value) => {
    let row = sheet.getElementsByTagName("table:table-row")[y - 1];
    let col = row.getElementsByTagName("table:table-cell")[x - 1];
    // update value
    col.removeAttribute("office:value");
    col.removeAttribute("office:date-value");

    let type;
    if (value == "") {
        type = "string";
    }
    else if (!isNaN(value * 1)) {
        type = "float";
        col.setAttribute("office:value", value);
    }
    // else if (new Date(value) !== "Invalid Date") {
    //     type = "date";
    //     console.log(value)
    //     // col.setAttribute("office:date-value", value);
    // }
    else {
        type = "string";
    }

    col.setAttribute("office:value-type", type);
    col.setAttribute("calcext:value-type", type);
    col.firstElementChild.textContent = value;
}

global.colNumberToColChar = (number) => {
    // number--;
    let lsChar = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

    let div = number;
    let colLetter = "";
    let mod = 0;
 
    while (div > 0)
    {
        mod = (div - 1) % 26;
        colLetter = lsChar[mod] + colLetter;
        div = Math.floor((div - mod) / 26);
    }
    return colLetter;
}

global.getDocument = (pathToFile) => {
    let file = fs.readFileSync(pathToFile, { flag: 'r' });
    
    // Create document by parsing XML
    let document = parser.parseFromString(file, "text/xml");

    return document;
}

global.saveDocument = (document, pathToFile) => {
    let file = document.documentElement.outerHTML;
    fs.writeFileSync(pathToFile, file);
}

global.getLsGraph = (sheetIndex, sheetName) => {
    let lsGraph = [];
    let lsTableConfig = config.lsSheet[sheetIndex].lsTable;

    let sheetOrientation = "x";
    if (lsTableConfig.length > 1) {
        if (lsTableConfig[1].xIndexSheet == 0) {
            sheetOrientation = "y";
        }
    }

    lsTableConfig.forEach((tableConfig, i) => {
        if (!(sheetOrientation == "x" && i !== 0)) {
            // define title
            let title = `${tableConfig.rowTitle} - ${tableConfig.dataTitle}`;

            // define type
            let type = tableConfig.nbBodyRow > 4 ? "bar" : "pie";

            // define range
            let x1 = 1;
            let y1 = i == 0 ? 1 + config.spreadSheetTemplate.tableau.nbHeadRow : lsTableConfig[i - 1].xIndexEndTable;
            let x2 = x1 + tableConfig.nbBodyCol + 1;
            let y2 = y1 + tableConfig.nbBodyRow - 1;

            let range = `${sheetName}.${colNumberToColChar(x1)}${y1}:${sheetName}.${colNumberToColChar(x1)}${y2} ${sheetName}.${colNumberToColChar(x2)}${y1}:${sheetName}.${colNumberToColChar(x2)}${y2}`;

            lsGraph.push({
                title: title,
                type: type,
                range: range
            });
        }

        if (tableConfig.colTitle == "") return;

        if (!(sheetOrientation == "y" && i !== 0)) {
            // second table
            // define title
            let title = `${tableConfig.colTitle} - ${tableConfig.dataTitle}`;

            // define type
            let type = tableConfig.nbBodyCol > 4 ? "bar" : "pie";

            // define range
            let x1 = i == 0 ? 1 + config.spreadSheetTemplate.tableau.nbHeadCol : lsTableConfig[i - 1].yIndexEndTable - config.spreadSheetTemplate.tableau.nbFootCol;
            let y1 = config.spreadSheetTemplate.tableau.nbHeadRow;
            let x2 = x1 + tableConfig.nbBodyCol - 1;
            let y2 = y1 + tableConfig.nbBodyRow + 1;

            let range = `${sheetName}.${colNumberToColChar(x1)}${y1}:${sheetName}.${colNumberToColChar(x2)}${y1} ${sheetName}.${colNumberToColChar(x1)}${y2}:${sheetName}.${colNumberToColChar(x2)}${y2}`;

            lsGraph.push({
                title: title,
                type: type,
                range: range
            });
        }
    });

    return lsGraph;
}