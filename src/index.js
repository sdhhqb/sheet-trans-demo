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

        // 指示当前遍历到的行是否是需要的行，每次遇到序号1时重新计算。
        // 正确的表格遇到序号1后，后续会有序号2,3,4,...,30。
        let validRow = false;

        // 遍历sheet的每一行
        jsonSheet.forEach(function (row, index) {
            // 检查学校和班级，每个sheet可能有多个班级
            if (typeof row[0] === 'string') {
                let classInfo = row[0].match(/学校[：:]\s*(\S*)\s+班级[：:]\s*(\S*)/);
                if (classInfo) {
                    school = classInfo[1];
                    grade = classInfo[2];
                }
            }

            // 当前行第一列是[1, 30]的序号
            if (row[0] && row[0] > 0 && row[0] < 31) {
                // 每次序号第1行时，检测当前小表头是否匹配, '序号', '学生姓名'
                if (row[0] == 1) {
                    let titleRow = jsonSheet[index - row[0]];
                    if (
                      typeof titleRow[0] === 'string' &&
                      titleRow[0].indexOf('序号') > -1 &&
                      typeof titleRow[1] === 'string' &&
                      titleRow[1].indexOf('学生姓名') > -1
                    ) {
                        validRow = true
                    } else {
                        validRow = false
                    }
                }

                if (validRow) {
                    if (typeof row[1] === 'string' && row[1].trim()) {
                        left.push([school, grade, row[1], row[3], fee, row[4]]);
                    }
                    if (typeof row[6] === 'string' && row[6].trim()) {
                        right.push([school, grade, row[6], row[8], fee, row[9]]);
                    }
                }
            }

            // 每个小表格有30行2列，在第30行时将表格左右两列的数据追加到data数组
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
    XLSX.writeFile(wb, path.resolve(__dirname, `../dist/${outputName}-1.xlsx`));
    console.log(`文件处理成功: ${outputName}-1.xlsx`);
    console.log('----------------------------------------------------------');
}
