const Excel = require('exceljs');
const fs = require('fs');

const Level = 1;
const Name = 2;
const WenJian = 3;
const YiBo = 4;
const Source = 5;
const AdditionalInfo = 6;

async function main() {
    const workbook = new Excel.Workbook();
    const worksheet = await workbook.csv.readFile('duty_4_data.csv');
    const sidebar = [];
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;

        const rowObj = {
            level: row.getCell(Level).value,
            name: row.getCell(Name).value,
            wenjian: row.getCell(WenJian).value,
            yibo: row.getCell(YiBo).value,
            source: row.getCell(Source).value,
            additionalInfo: row.getCell(AdditionalInfo).value,
        };

        // console.log(rowObj);
        processRow(rowObj);

        sidebar.push(`- [${rowObj.level}级 ${rowObj.name}](duty_4/${rowObj.name})`);
    });

    const sideBarText = sidebar.join(`\n`);
    const path = `../duty_4/side_bar_temp.md`;
    await fs.promises.writeFile(path, sideBarText, {flag: 'w+'});
}

async function processRow(row) {
    let output = `
<!-- docs/duty_4/${row.name}.md -->

# ${row.level}级 ${row.name}

> ${row.source ?? ''}
`;

    if (row.additionalInfo) {
        output += `
${row.additionalInfo ?? ''}
`;
    }

    if (row.wenjian) {
        output += `
## 稳健
![稳健拉法](../${row.wenjian})
`;
    }

    if (row.yibo) {
        output += `
## 一波
![一波拉法](../${row.yibo})
`;
    }

    // console.log(output);

    const path = `../duty_4/${row.name}.md`;
    await fs.promises.writeFile(path, output, {flag: 'w+'});
}

main();
