const path = require('path');
const fse = require('fs-extra');
const XLSX = require('xlsx');

// 清空dist目录
fse.emptyDirSync(path.resolve(__dirname, '../dist'));

// 读取source目录中的excel文件
fse.readdir(path.resolve(__dirname,'../source'), function (err, files) {
    if (err) {
        console.log('读取文件出错', err);
        return;
    }
    const filenames = files.filter(filename => filename.match(/\.xlsx?$/));
    filenames.forEach(function (filename) {
        processXlsxFile(filename);
    });
});

// 处理excel
function processXlsxFile(filename) {
    console.log(`开始处理文件: ${filename}`);

    let workbook = XLSX.readFile(path.resolve(__dirname, `../source/${filename}`));
    let data = [
        ['学校', '班级', '姓名', '性别', '保费', '身份证']
    ];
    workbook.SheetNames.forEach(function (sheetName) {
        let sheet = workbook.Sheets[sheetName];
        let jsonSheet =  XLSX.utils.sheet_to_json(sheet, {header: 1});
        let left = [];
        let right = [];
        let school = '';
        let grade = '';
        let fee = 0;

        // 检查保费，每个sheet只用检查一次
        jsonSheet.some(function (row) {
            if (typeof row[0] === 'string') {
                let feeMatch = row[0].match(/保费：\s*(\d+)\s*元\/人/);
                if (feeMatch) {
                    fee = feeMatch[1];
                    return true;
                }
            }
        });

        jsonSheet.forEach(function (row, index) {
            // 检查学校和班级，每个sheet可能有多个班级
            if (typeof row[0] === 'string') {
                let classInfo = row[0].match(/学校：\s*(\S*)\s+班级：\s*(\S*)/);
                if (classInfo) {
                    school = classInfo[1];
                    grade = classInfo[2];
                }
            }
            if (row[0] && row[0] > 0 && row[0] < 60) {
                if (row[1]) {
                    left.push([school, grade, row[1], row[3], fee, row[4]]);
                }
                if (row[6]) {
                    right.push([school, grade, row[6], row[8], fee, row[9]]);
                }
            }
            if (row[0] == 30) {
                data = data.concat(left, right);
                left = [];
                right = [];
            }
        })
    })

    // console.log(data);
    let outputName = filename.substring(0, filename.lastIndexOf('.'));
    let wb = XLSX.utils.book_new();
    let ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, 'sheet1');
    XLSX.writeFile(wb, path.resolve(__dirname, `../dist/${outputName}.xlsx`));
    console.log(`文件处理成功: ${filename}`);
    console.log('----------------------------------------------------------');
}
