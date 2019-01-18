var express = require('express');
var bodyParser = require('body-parser')
var router = express.Router();
var excel = require('excel4node');

router.use(bodyParser.json());

router.post('/', function(req, res, next) {
    let param = req.body; //获取参数
    let exceldata = param.exceldata;
    let filename = exceldata.filename || undefined;
    let sheets = exceldata.sheets || [];
    if (filename) {
        let workbook = new excel.Workbook(); // 有文件名，开始创建Excel文件
        if (sheets.length > 0) { // 有sheet
            for (let i=0; i<sheets.length; i++) {
                let sheetname = sheets[i].sheetname || ''; // sheet名
                if (sheetname !== '') {
                    let worksheet = workbook.addWorksheet(sheetname); // 创建sheet
                    let column = sheets[i].column || []; // 列信息
                    for (let cols=0; cols<column.length; cols++){ // 列数
                        let dataKey = column[cols].dataKey;
                        let ref = column[cols].ref;
                        let title = column[cols].title || '';
                        let width = column[cols].width || undefined;
                        worksheet.cell(1,ref).string(title);
                        if (width) {
                            worksheet.column(ref).setWidth(width);
                        }
                    }
                    let data = sheets[i].data || [];; // 行信息
                    for (let rows=0; rows<data.length; rows++) { // 行数
                        let row = data[rows].row;
                        for (let rowj=0; rowj<row.length; rowj++) {
                            let rowKey = row[rowj].dataKey;
                            let text = row[rowj].text || '';
                            let type = row[rowj].type || 's';
                            // 查找列
                            let colNum = undefined;
                            for (let cols=0; cols<column.length; cols++){
                                if (rowKey == column[cols].dataKey) {
                                    colNum = column[cols].ref;
                                }
                            }
                            if (colNum) {
                                let rowNum = rows + 2; // 从第二行开始 +2
                                if (typeof(text) == 'string') {
                                    if (type === 'n') { // str 强制转为 num
                                        text = Number(text);
                                        worksheet.cell(rowNum,colNum).number(text);
                                    } else {
                                        worksheet.cell(rowNum,colNum).string(text);
                                    }
                                } else if (typeof(text) == 'number') { // num 强制转为 str
                                    if (type === 's') {
                                        text = String(text);
                                        worksheet.cell(rowNum,colNum).string(text);
                                    } else {
                                        worksheet.cell(rowNum,colNum).number(text);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        workbook.write(`${filename}.xlsx`, res);
    } else {
        res.status(400);
        res.send({"status":"未设置文件名",});
    }
});

module.exports = router;
