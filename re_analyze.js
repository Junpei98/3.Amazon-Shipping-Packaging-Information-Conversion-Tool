const xlsx = require('xlsx');

function analyzeFile(filePath) {
    console.log(`\n--- ${filePath} ---`);
    const workbook = xlsx.readFile(filePath);
    console.log("Sheets:", workbook.SheetNames);
    workbook.SheetNames.forEach(name => {
        const sheet = workbook.Sheets[name];
        const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });
        console.log(`\nSheet: ${name} (Rows: ${rows.length})`);
        rows.slice(0, 15).forEach((row, i) => {
            console.log(`${i}:`, JSON.stringify(row));
        });
    });
}

analyzeFile("c:/Users/tenni/OneDrive/デスクトップ/Amazon納品梱包情報変換ツール/Excel①（Amazonの梱包グループのシート）.xlsx");
analyzeFile("c:/Users/tenni/OneDrive/デスクトップ/Amazon納品梱包情報変換ツール/Excel②（ラクマートの輸送箱などの情報シート）.xlsx");
