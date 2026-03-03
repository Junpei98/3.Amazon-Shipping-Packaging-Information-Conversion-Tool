const xlsx = require('xlsx');

function analyzeFile(filePath) {
    console.log(`\n======================================================`);
    console.log(`--- Analyzing ${filePath} ---`);
    try {
        const workbook = xlsx.readFile(filePath);
        console.log(`Sheet names: ${workbook.SheetNames.join(', ')}`);
        
        workbook.SheetNames.forEach(sheetName => {
            console.log(`\n[ Sheet: ${sheetName} ]`);
            const worksheet = workbook.Sheets[sheetName];
            // Read as 2D array
            const data = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: null });
            console.log(`Total Rows: ${data.length}`);
            if (data.length > 0) {
                console.log("First 10 rows:");
                data.slice(0, 10).forEach((row, i) => {
                    // Filter out empty arrays or too long empty trails for display
                    console.log(`Row ${i + 1}:`, JSON.stringify(row));
                });
            }
        });
    } catch (error) {
        console.error(`Error reading ${filePath}:`, error.message);
    }
}

analyzeFile("c:/Users/tenni/OneDrive/デスクトップ/Amazon納品梱包情報変換ツール/Excel①（Amazonの梱包グループのシート）.xlsx");
analyzeFile("c:/Users/tenni/OneDrive/デスクトップ/Amazon納品梱包情報変換ツール/Excel②（ラクマートの輸送箱などの情報シート）.xlsx");
