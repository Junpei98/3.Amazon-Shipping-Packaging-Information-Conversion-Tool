const xlsx = require('xlsx');

function dumpSheet(filePath, sheetName) {
    console.log(`\n--- ${filePath} [${sheetName}] ---`);
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
        console.log("Sheet not found!");
        return;
    }
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "(null)" });
    rows.slice(0, 20).forEach((row, i) => {
        // Only show first 16 columns and truncate each cell
        const shortRow = row.slice(0, 16).map(c => {
            const s = (c || "").toString();
            return s.length > 20 ? s.slice(0, 20) + "..." : s;
        });
        console.log(`Row ${i}:`, JSON.stringify(shortRow));
    });
}

dumpSheet("c:/Users/tenni/OneDrive/デスクトップ/Amazon納品梱包情報変換ツール/Excel②（ラクマートの輸送箱などの情報シート）.xlsx", "梱包リスト");
