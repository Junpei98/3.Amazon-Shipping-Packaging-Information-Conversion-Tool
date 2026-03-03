lucide.createIcons();

const file1Input = document.getElementById('file-1');
const file2Input = document.getElementById('file-2');
const dropZone1 = document.getElementById('drop-zone-1');
const dropZone2 = document.getElementById('drop-zone-2');
const convertBtn = document.getElementById('convert-btn');
const statusArea = document.getElementById('status-area');
const progressInner = document.getElementById('progress-inner');
const logArea = document.getElementById('log-area');
const resultSection = document.getElementById('result-section');

let workbook1 = null;
let workbook2 = null;

function addLog(msg) {
    const div = document.createElement('div');
    div.innerHTML = `> ${msg}`;
    logArea.appendChild(div);
    logArea.scrollTop = logArea.scrollHeight;
}

function updateProgress(percent) {
    progressInner.style.width = `${percent}%`;
}

function norm(str) {
    if (str === null || str === undefined) return "";
    return str.toString().replace(/[\s\u3000\r\n\t]+/g, '');
}

function setupDropZone(zone, input, label) {
    zone.addEventListener('click', () => input.click());
    zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop', (e) => {
        e.preventDefault(); zone.classList.remove('dragover');
        if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0], zone, label);
    });
    input.addEventListener('change', (e) => {
        if (e.target.files.length) handleFile(e.target.files[0], zone, label);
    });
}

function handleFile(file, zone, label) {
    if (!file.name.endsWith('.xlsx')) { alert('Excelファイル(.xlsx)を選択してください。'); return; }
    zone.querySelector('.filename').textContent = file.name;
    zone.querySelector('.filename').style.display = 'block';
    const reader = new FileReader();
    reader.onload = (e) => {
        const workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        if (label === 1) workbook1 = workbook;
        if (label === 2) workbook2 = workbook;
        if (workbook1 && workbook2) convertBtn.disabled = false;
    };
    reader.readAsArrayBuffer(file);
}

setupDropZone(dropZone1, file1Input, 1);
setupDropZone(dropZone2, file2Input, 2);

convertBtn.addEventListener('click', async () => {
    convertBtn.disabled = true;
    statusArea.style.display = 'block';
    logArea.innerHTML = '';
    try {
        await processConversion();
    } catch (err) {
        addLog(`<span style="color:#FF5555;">ERROR: ${err.message}</span>`);
        console.error(err);
    } finally {
        convertBtn.disabled = false;
    }
});

async function processConversion() {
    addLog('解析を開始します...');

    // =============================================
    // STEP 1: Analyze Excel② (Rakumart)
    // =============================================
    updateProgress(10);
    const sheet2Name = workbook2.SheetNames.find(n => n.includes('梱包リスト') || n.includes('明細') || n.includes('箱子'));
    if (!sheet2Name) throw new Error('Excel②に「梱包リスト」シートが見つかりません。');
    const rawData2 = XLSX.utils.sheet_to_json(workbook2.Sheets[sheet2Name], { header: 1, defval: null });
    addLog(`Excel②「${sheet2Name}」: ${rawData2.length}行を読み込みました。`);

    let headerRowIdx2 = -1;
    for (let i = 0; i < Math.min(rawData2.length, 10); i++) {
        const rowStr = (rawData2[i] || []).map(c => norm(c)).join('|');
        if ((rowStr.includes('箱NO') || rowStr.includes('箱号')) && rowStr.includes('梱包数')) { headerRowIdx2 = i; break; }
    }
    if (headerRowIdx2 === -1) throw new Error('Excel②のヘッダー行が見つかりません。');

    const headerRow2 = rawData2[headerRowIdx2];
    const idxBox = headerRow2.findIndex(c => ["箱NO", "箱号", "箱番号"].includes(norm(c)));
    const idxSize = headerRow2.findIndex(c => norm(c).includes('寸法') || norm(c).includes('サイズ') || norm(c).includes('规格'));
    const idxWeight = headerRow2.findIndex(c => norm(c).includes('重量'));
    const idxQty = headerRow2.findIndex(c => ["梱包数", "商品数量", "入数"].includes(norm(c)));
    const idxFnsku = headerRow2.findIndex(c => ["ラベル番号", "FNSKU", "商品コード"].includes(norm(c)));

    if (idxBox === -1 || idxFnsku === -1 || idxQty === -1) throw new Error(`必要な列が見つかりません。`);

    const boxMap = {};
    let lastBoxNum = null;
    for (let i = headerRowIdx2 + 1; i < rawData2.length; i++) {
        const row = rawData2[i];
        if (!row) continue;
        const boxRaw = row[idxBox];
        if (boxRaw !== null && boxRaw !== undefined && boxRaw.toString().trim() !== '') {
            const parsed = parseInt(boxRaw.toString().replace(/[^0-9]/g, ''));
            if (!isNaN(parsed)) {
                lastBoxNum = parsed;
                const weight = parseFloat(row[idxWeight]) || 0;
                const sizeStr = norm(row[idxSize]);
                let l = 0, w = 0, h = 0;
                if (sizeStr && sizeStr.includes('*')) {
                    const dims = sizeStr.split('*').map(d => parseFloat(d));
                    if (dims.length === 3 && dims.every(d => !isNaN(d))) [l, w, h] = dims;
                }
                if (!boxMap[lastBoxNum]) boxMap[lastBoxNum] = { items: {}, weight, l, w, h };
            }
        }
        if (lastBoxNum === null) continue;
        const fnsku = norm(row[idxFnsku]).trim();
        const qty = parseInt(row[idxQty]);
        if (fnsku && fnsku !== 'null' && !isNaN(qty) && qty > 0) {
            boxMap[lastBoxNum].items[fnsku] = (boxMap[lastBoxNum].items[fnsku] || 0) + qty;
        }
    }

    const boxNums = Object.keys(boxMap).map(Number).sort((a, b) => a - b);
    addLog(`輸送箱 ${boxNums.length} 個を検出しました。`);
    if (boxNums.length === 0) throw new Error('輸送箱データが0件です。');

    // =============================================
    // STEP 2: Analyze Excel① (Amazon Template)
    // =============================================
    updateProgress(50);
    const sheet1Name = workbook1.SheetNames.find(n => n.includes('輸送箱の梱包情報'));
    if (!sheet1Name) throw new Error('Excel①に「輸送箱の梱包情報」シートが見つかりません。');
    const rawData1 = XLSX.utils.sheet_to_json(workbook1.Sheets[sheet1Name], { header: 1, defval: null });

    let skuHeaderRowIdx = -1;
    for (let i = 0; i < Math.min(rawData1.length, 20); i++) {
        if ((rawData1[i] || []).some(c => norm(c).includes('FNSKU'))) { skuHeaderRowIdx = i; break; }
    }
    if (skuHeaderRowIdx === -1) throw new Error('Excel①にFNSKU列が見つかりません。');
    const headerRow1 = rawData1[skuHeaderRowIdx];
    const idxFnsku1 = headerRow1.findIndex(c => norm(c).includes('FNSKU'));
    const idxSku1 = headerRow1.findIndex(c => norm(c) === 'SKU');
    // Look for planned quantity column (予定数量 or 出荷予定数量 etc.)
    const idxPlanned = headerRow1.findIndex(c => norm(c).includes('予定数量') || norm(c).includes('数量'));

    const products = [];
    const fnskuToProductIdx = {};
    for (let i = skuHeaderRowIdx + 1; i < rawData1.length; i++) {
        const fnsku = norm(rawData1[i][idxFnsku1]).trim();
        const sku = (rawData1[i][idxSku1] || '').toString().trim();
        if (!fnsku || fnsku === 'null') continue;
        const planned = idxPlanned !== -1 ? parseInt(rawData1[i][idxPlanned]) : NaN;
        fnskuToProductIdx[fnsku] = products.length;
        products.push({ fnsku, sku, planned });
    }
    addLog(`Excel①から ${products.length} 件の商品を読み込みました。`);
    updateProgress(80);

    // =============================================
    // STEP 3: Compute actual totals per FNSKU
    // =============================================
    const actualTotals = {};
    boxNums.forEach(boxNum => {
        Object.keys(boxMap[boxNum].items).forEach(fnsku => {
            actualTotals[fnsku] = (actualTotals[fnsku] || 0) + boxMap[boxNum].items[fnsku];
        });
    });

    // =============================================
    // STEP 4: Render Table 1 (Product × Box quantities)
    // =============================================
    const qtyTable = document.getElementById('qty-table');
    qtyTable.innerHTML = '';
    const thead = qtyTable.createTHead();
    const theadRow = thead.insertRow();
    ['FNSKU', 'SKU', ...boxNums.map(n => `箱 ${n}`)].forEach((h, i) => {
        const th = document.createElement('th');
        th.textContent = h;
        th.className = i < 2 ? 'left' : '';
        theadRow.appendChild(th);
    });
    const tbody = qtyTable.createTBody();
    products.forEach(prod => {
        const tr = tbody.insertRow();
        const cellFnsku = tr.insertCell(); cellFnsku.textContent = prod.fnsku; cellFnsku.className = 'left';
        const cellSku = tr.insertCell(); cellSku.textContent = prod.sku; cellSku.className = 'left';
        let hasValue = false;
        boxNums.forEach(boxNum => {
            const td = tr.insertCell();
            const qty = boxMap[boxNum]?.items[prod.fnsku];
            if (qty) { td.textContent = qty; td.className = 'value'; hasValue = true; }
            else { td.textContent = ''; td.className = 'empty'; }
        });
        if (!hasValue) tr.style.opacity = '0.35';
    });

    // =============================================
    // STEP 5: Render Table 2 (Box Size & Weight)
    // =============================================
    const sizeTable = document.getElementById('size-table');
    sizeTable.innerHTML = '';
    const sthead = sizeTable.createTHead();
    const stheadRow = sthead.insertRow();
    ['項目', ...boxNums.map(n => `輸送箱 ${n}`)].forEach((h, i) => {
        const th = document.createElement('th'); th.textContent = h; th.className = i === 0 ? 'left' : ''; stheadRow.appendChild(th);
    });
    const stbody = sizeTable.createTBody();
    const sizeRows = [
        { label: '輸送箱の重量 (kg)', key: 'weight' },
        { label: '輸送箱の長さ (cm)', key: 'l' },
        { label: '輸送箱の幅 (cm)', key: 'w' },
        { label: '輸送箱の高さ (cm)', key: 'h' },
    ];
    sizeRows.forEach(({ label, key }) => {
        const tr = stbody.insertRow();
        const th = document.createElement('td'); th.textContent = label; th.className = 'left'; tr.appendChild(th);
        boxNums.forEach(boxNum => {
            const td = tr.insertCell();
            const val = boxMap[boxNum]?.[key];
            td.textContent = val > 0 ? val : '';
            td.className = val > 0 ? 'value' : 'empty';
        });
    });

    // =============================================
    // STEP 6: Render Validation Table
    // =============================================
    const valTable = document.getElementById('validation-table');
    valTable.innerHTML = '';
    const vthead = valTable.createTHead();
    const vtheadRow = vthead.insertRow();
    ['FNSKU', 'SKU', '予定数量', '入庫数合計', '差異', '状態'].forEach((h, i) => {
        const th = document.createElement('th'); th.textContent = h; th.className = i < 2 ? 'left' : ''; vtheadRow.appendChild(th);
    });
    const vtbody = valTable.createTBody();
    let allMatch = true;
    products.forEach(prod => {
        const actual = actualTotals[prod.fnsku] || 0;
        const planned = !isNaN(prod.planned) ? prod.planned : null;
        const diff = planned !== null ? actual - planned : null;
        const isMatch = diff === null ? null : diff === 0;
        if (isMatch === false) allMatch = false;

        const tr = vtbody.insertRow();
        const f = tr.insertCell(); f.textContent = prod.fnsku; f.className = 'left';
        const s = tr.insertCell(); s.textContent = prod.sku; s.className = 'left';
        const p = tr.insertCell(); p.textContent = planned !== null ? planned : '−';
        const a = tr.insertCell(); a.textContent = actual;
        const d = tr.insertCell();
        const st = tr.insertCell();

        if (isMatch === null) {
            d.textContent = '−'; st.innerHTML = '<span class="badge badge-gray">不明</span>';
        } else if (isMatch) {
            d.textContent = '0'; st.innerHTML = '<span class="badge badge-ok">✔ 一致</span>';
            tr.classList.add('row-ok');
        } else {
            d.textContent = (diff > 0 ? '+' : '') + diff;
            d.style.color = '#FF5555'; d.style.fontWeight = '700';
            st.innerHTML = '<span class="badge badge-ng">✗ 不一致</span>';
            tr.classList.add('row-ng');
        }
    });

    // Summary badge
    document.getElementById('val-summary').innerHTML = allMatch
        ? '<span class="badge badge-ok" style="font-size:0.9rem; padding:8px 20px;">✔ 全商品の数量が一致しています</span>'
        : '<span class="badge badge-ng" style="font-size:0.9rem; padding:8px 20px;">✗ 数量が合っていない商品があります</span>';

    // Store copy data (data-only, no labels)
    // qty-table: skip first 2 cols (FNSKU, SKU)
    // size-table: skip first 1 col (label)
    updateProgress(100);
    lucide.createIcons();
    addLog('<span style="color:#AAFFAA;">✅ 変換完了！下の表をご確認ください。</span>');
}

// Copy only data columns (skip label columns)
function copyTableData(tableId, skipCols, btnId) {
    const table = document.getElementById(tableId);
    const rows = Array.from(table.querySelectorAll('tbody tr'));
    const tsv = rows.map(row => {
        const cells = Array.from(row.querySelectorAll('td'));
        return cells.slice(skipCols).map(c => c.textContent.trim()).join('\t');
    }).join('\n');

    const ta = document.createElement('textarea');
    ta.value = tsv;
    ta.style.position = 'fixed'; ta.style.top = '-9999px'; ta.style.left = '-9999px';
    document.body.appendChild(ta);
    ta.focus(); ta.select();
    let success = false;
    try { success = document.execCommand('copy'); } catch (e) { }
    document.body.removeChild(ta);

    const btn = document.getElementById(btnId);
    const originalHTML = btn.innerHTML;
    if (success) {
        btn.classList.add('copied');
        btn.innerHTML = '<i data-lucide="check" width="15" height="15"></i> コピーしました！';
        lucide.createIcons();
        setTimeout(() => { btn.classList.remove('copied'); btn.innerHTML = originalHTML; lucide.createIcons(); }, 2500);
    } else {
        alert('コピーに失敗しました。');
    }
}
