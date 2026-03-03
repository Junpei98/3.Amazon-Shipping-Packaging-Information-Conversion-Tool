const xlsx = require('xlsx');
const filePath = "c:/Users/tenni/OneDrive/デスクトップ/Amazon納品梱包情報変換ツール/Excel②（ラクマートの輸送箱などの情報シート）.xlsx";
try {
    const workbook = xlsx.readFile(filePath);
    console.log("SheetNames:", JSON.stringify(workbook.SheetNames));
} catch (e) {
    console.error(e);
}
