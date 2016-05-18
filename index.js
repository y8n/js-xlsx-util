var xlsx = require('xlsx');
var fs = require('fs');
var path = require('path');

var LETTERS_OBJ = {
    A: 1,
    B: 2,
    C: 3,
    D: 4,
    E: 5,
    F: 6,
    G: 7,
    H: 8,
    I: 9,
    J: 10,
    K: 11,
    L: 12,
    M: 13,
    N: 14,
    O: 15,
    P: 16,
    Q: 17,
    R: 18,
    S: 19,
    T: 20,
    U: 21,
    V: 22,
    W: 23,
    X: 24,
    Y: 25,
    Z: 26
};
var LETTERS = ['0'].concat(Object.keys(LETTERS_OBJ));

/**
 * 读取文件内容,缓存
 * @param xlsxPath - xlsx路径
 * @param noCache - 是否缓存,默认缓存
 * @returns {*}
 */
exports.readFile = function readFile(xlsxPath, noCache) {
    var xlsxDir = path.dirname(xlsxPath);
    var xlsxName = path.basename(xlsxPath);
    var jsonPath = path.resolve(xlsxDir, xlsxName + '.json');
    var workbook;
    if (noCache) {
        workbook = xlsx.readFile(xlsxPath);
    } else if (fs.existsSync(jsonPath)) {
        workbook = require(jsonPath);
    } else {
        workbook = xlsx.readFile(xlsxPath);
        fs.writeFileSync(jsonPath, JSON.stringify(workbook), 'utf8');
    }
    exports.formatCell(workbook); // 格式化
    return workbook;
};
/**
 * 格式化单元格内容
 * @param data
 */
exports.formatCell = function (data) {
    if ('Directory' in data) { // workbook
        data.SheetNames.forEach(function (sheetName) {
            formatWorksheet(data.Sheets[sheetName]);
        });
    } else { // worksheet
        formatWorksheet(data);
    }
};
/**
 * 格式化workSheet,为每一个单元格添加行,列属性及获取上下左右四个方位单元格的方法
 * @param sheet
 */
function formatWorksheet(sheet) {
    var format;
    exports.each(sheet, function (key, value) {
        format = exports.formatKey(key);
        value.col = format.col;
        value.row = format.row;
        value.top = function (gap) {
            gap = gap || 1;
            return sheet[value.col + (value.row - gap)];
        };
        value.bottom = function (gap) {
            gap = gap || 1;
            return sheet[value.col + (value.row + gap)];
        };
        value.left = function (gap) {
            gap = gap || 1;
            var before,current = value.col;
            while (gap){
                before = exports.getPrevCol(current);
                current = before;
                gap -= 1;
            }
            return sheet[before + value.row];
        };
        value.right = function (gap) {
            gap = gap || 1;
            var after,current = value.col;
            while (gap){
                after = exports.getNextCol(current);
                current = after;
                gap -= 1;
            }
            return sheet[after + value.row];
        };
    });
}
/**
 * 获取某一列的前一列
 * @param colOrCell - 单元格或列名
 * @returns {*}
 */
exports.getPrevCol = function getPrevCol(colOrCell) {
    var col;
    if (typeof colOrCell === 'string') { // col
        col = colOrCell;
    } else {
        col = colOrCell.col;
    }
    col = col.toUpperCase();
    var last = col[col.length - 1];
    var lastLetterIndex = LETTERS_OBJ[last];
    if (lastLetterIndex === 1) {
        return col.length === 1 ? '' : (exports.getPrevCol(col.slice(0, -1)) + 'Z');
    }
    return col.length === 1 ? LETTERS[lastLetterIndex - 1] : (col.slice(0, -1) + LETTERS[lastLetterIndex - 1]);
};
/**
 * 获取某一列的后一列
 * @param colOrCell - 单元格或列名
 * @returns {*}
 */
exports.getNextCol = function getPrevCol(colOrCell) {
    var col;
    if (typeof colOrCell === 'string') {
        col = colOrCell;
    } else {
        col = colOrCell.col;
    }
    col = col.toUpperCase();
    var last = col[col.length - 1];
    var lastLetterIndex = LETTERS_OBJ[last];
    if (lastLetterIndex === 26) {
        return (col.length === 1 ? 'A' : exports.getNextCol(col.slice(0, -1))) + 'A';
    }
    return col.length === 1 ? LETTERS[lastLetterIndex + 1] : (col.slice(0, -1) + LETTERS[lastLetterIndex + 1]);
};
/**
 * 格式化键,输出行,列
 * @param key - 单元格的键,如A2
 * @returns {{col: number, row: number}}
 */
exports.formatKey = function formatKey(key) {
    var key_reg = /^([A-Z]+)(\d+)$/;
    var match = key.match(key_reg);
    return {
        col: match ? match[1] : -1,
        row: match ? parseInt(match[2]) : -1
    };
};
/**
 * 遍历worksheet,排除以!开头的属性
 * @param worksheet - 待遍历的worksheet
 * @param callback - 回调
 */
exports.each = function (worksheet, callback) {
    for (var key in worksheet) {
        if (worksheet.hasOwnProperty(key) && key[0] !== '!' && callback) {
            callback(key, worksheet[key], worksheet);
        }
    }
};

/**
 * 过滤器,排除以!开头的属性
 * @param worksheet
 * @param fn
 * @returns [] - 符合条件的单元格
 */
exports.filter = function filter(worksheet, fn) {
    var result = [];
    for (var key in worksheet) {
        if (worksheet.hasOwnProperty(key) && key[0] !== '!') {
            if (fn && fn(key, worksheet[key], worksheet)) {
                result.push(worksheet[key]);
            }
        }
    }
    return result;
};
/**
 * 构建ref
 * @param worksheet
 * @returns {string}
 */
exports.buildRef = function buildRef(worksheet) {
    if (worksheet['!ref']) {
        return worksheet['!ref'];
    }
    var keys = Object.keys(worksheet);
    return keys[0] + ':' + keys[keys.length - 1]; // 从第一个开始,排除!ref
};
/**
 * 构建一个单元格
 * @param val
 * @returns {*}
 */
exports.cell = function cell(val) {
    if (typeof val === 'number') {
        return {t: 'n', v: val};
    }
    return {t: 's', v: val.toString()};
};
/**
 * 输出文件
 * @param filepath
 * @param workbook
 */
exports.writeFile = function (filepath, workbook) {
    xlsx.writeFile(workbook, filepath);
};
/**
 * 判断两个值是否一样
 * @param a
 * @param b
 * @returns {boolean}
 */
exports.isEqual = function isEqual(a, b) {
    return a.toString().trim() === b.toString().trim();
};
/**
 * 判断是否是Undefined
 * @param val
 * @returns {boolean}
 */
exports.isUndefined = function (val) {
    return typeof val === 'undefined';
};
/**
 * 判断变量是否定义
 * @param val
 * @returns {boolean}
 */
exports.isDefined = function (val) {
    return typeof val !== 'undefined';
};
/**
 * 在worksheet中追加一行
 */
exports.addRow = function addRow(worksheet, rowObj) {
    var ref = worksheet['!ref'] || exports.buildRef(worksheet);
    var lastRow = exports.formatKey(ref.split(':')[1]).row;
    for (var key in rowObj) {
        if (rowObj.hasOwnProperty(key)) {
            if(typeof rowObj === 'object' && rowObj.t){
                worksheet[key + (lastRow + 1)] = rowObj[key];
            }else{
                worksheet[key + (lastRow + 1)] = exports.cell(rowObj[key]);
            }
        }
    }
};
/**
 * 添加一个工作表
 * @param workbook
 * @param sheetName
 * @param sheet
 */
exports.addWorkSheet = function addWorkSheet(workbook, sheetName, sheet) {
    if (!workbook.Sheets) {
        workbook.Sheets = {};
    }
    if (!workbook.SheetNames) {
        workbook.SheetNames = [];
    }
    workbook.Sheets[sheetName] = sheet;
    if (workbook.SheetNames.indexOf(sheetName) === -1) {
        workbook.SheetNames.push(sheetName);
    }
};