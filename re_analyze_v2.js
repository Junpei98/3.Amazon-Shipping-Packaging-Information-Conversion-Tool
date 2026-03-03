const xlsx = require('xlsx');

function dumpSheet(filePath, sheetName) {
    console.log(`\n--- ${filePath} [${sheetName}] ---`);
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
        console.log("Sheet not found!");
        return;
    }
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "(empty)" });
    rows.slice(0, 30).forEach((row, i) => {
        console.log(`Row ${i}:`, JSON.stringify(row));
    });
}

dumpSheet("c:/Users/tenni/OneDrive/デスクトップ/Amazon納品梱包情報変換ツール/Excel①（Amazonの梱包グループのシート）.xlsx", "輸送箱の梱包情報");
dumpSheet("c:/Users/tenni/OneDrive/デスクトップ/Amazon納品梱包情報変換ツール/Excel②（ラクマートの輸送箱などの情報シート）.xlsx", "梱包リスト");
