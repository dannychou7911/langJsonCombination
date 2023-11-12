const fs = require('fs');
const ExcelJS = require('exceljs');

// 假設有四個JSON檔案: en.json, zh.json, ... (其他兩個語系)
let enData = JSON.parse(fs.readFileSync('lang/en.json', 'utf8'));
let zhData = JSON.parse(fs.readFileSync('lang/zh.json', 'utf8'));
let thData = JSON.parse(fs.readFileSync('lang/th.json', 'utf8'));
let ptData = JSON.parse(fs.readFileSync('lang/pt.json', 'utf8'));

// 遞迴函數來展開巢狀結構
function expandNestedObjects(obj, prefix = '') {
    let result = {};
    for (let key in obj) {
        if (typeof obj[key] === 'object') {
            result = { ...result, ...expandNestedObjects(obj[key], `${prefix}${key}.`) };
        } else {
            result[`${prefix}${key}`] = obj[key];
        }
    }
    return result;
}

// 展開英文和中文數據
const expandedEnData = expandNestedObjects(enData);
const expandedZhData = expandNestedObjects(zhData);
const expandedThData = expandNestedObjects(thData);
const expandedPtData = expandNestedObjects(ptData);

const mergedData = [];
for (const key in expandedEnData) {
    if (expandedZhData.hasOwnProperty(key)) {
        mergedData.push({
            key,
            en: expandedEnData[key],
            zh: expandedZhData[key],
            th: expandedThData[key],
            pt: expandedPtData[key],
        });
    }
}

// 創建一個新的Excel工作簿
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Translations');

worksheet.columns = [
    { header: 'en', key: 'en', width: 30 },
    { header: 'zh', key: 'zh', width: 30 },
    { header: 'th', key: 'th', width: 30 },
    { header: 'pt', key: 'pt', width: 30 },
];

// Add merged data to the worksheet
mergedData.forEach((row) => {
    worksheet.addRow(row);
});

// Write to an Excel file
workbook.xlsx
    .writeFile('translations.xlsx')
    .then(() => {
        console.log('Excel file saved!');
    })
    .catch((error) => {
        console.error('Error writing to Excel file:', error);
    });
