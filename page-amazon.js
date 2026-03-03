// =============================================
// page-amazon.js – Amazon納品梱包情報変換ロジック
// =============================================

(function () {
    let wb1 = null, wb2 = null;

    setupDropZone('dz1-amazon', 'file1-amazon', wb => { wb1 = wb; checkReady(); });
    setupDropZone('dz2-amazon', 'file2-amazon', wb => { wb2 = wb; checkReady(); });

    function checkReady() {
        document.getElementById('convert-btn-amazon').disabled = !(wb1 && wb2);
    }

    document.getElementById('convert-btn-amazon').addEventListener('click', async () => {
        document.getElementById('convert-btn-amazon').disabled = true;
        document.getElementById('status-amazon').style.display = 'block';
        document.getElementById('log-amazon').innerHTML = '';
        try { await run(); }
        catch (err) { addLog('log-amazon', `<span style="color:#FF5555;">ERROR: ${err.message}</span>`); console.error(err); }
        finally { document.getElementById('convert-btn-amazon').disabled = false; }
    });

    async function run() {
        addLog('log-amazon', '解析を開始します...');
        updateProgress('prog-amazon', 10);

        // --- Parse Excel② ---
        const boxMap = parseRakumart(wb2);
        const boxNums = Object.keys(boxMap).map(Number).sort((a, b) => a - b);
        addLog('log-amazon', `輸送箱 ${boxNums.length} 個を検出しました。`);
        if (!boxNums.length) throw new Error('輸送箱データが0件です。');

        // Flatten items for qty map
        const qtyByBox = {};
        boxNums.forEach(n => {
            qtyByBox[n] = {};
            boxMap[n].items.forEach(it => { qtyByBox[n][it.fnsku] = (qtyByBox[n][it.fnsku] || 0) + it.qty; });
        });

        updateProgress('prog-amazon', 40);

        // --- Parse Excel① ---
        const sheet1Name = wb1.SheetNames.find(n => n.includes('輸送箱の梱包情報'));
        if (!sheet1Name) throw new Error('Excel①に「輸送箱の梱包情報」シートが見つかりません。');
        const raw1 = XLSX.utils.sheet_to_json(wb1.Sheets[sheet1Name], { header: 1, defval: null });

        let skuRow = -1;
        for (let i = 0; i < Math.min(raw1.length, 20); i++) {
            if ((raw1[i] || []).some(c => norm(c).includes('FNSKU'))) { skuRow = i; break; }
        }
        if (skuRow === -1) throw new Error('Excel①にFNSKU列が見つかりません。');
        const h1 = raw1[skuRow];
        const iFnsku = h1.findIndex(c => norm(c).includes('FNSKU'));
        const iSku = h1.findIndex(c => norm(c) === 'SKU');
        const iPlanned = h1.findIndex(c => norm(c).includes('予定数量') || norm(c).includes('数量'));

        const products = [];
        const fnskuRow = {};
        for (let i = skuRow + 1; i < raw1.length; i++) {
            const fnsku = norm(raw1[i][iFnsku]).trim();
            if (!fnsku || fnsku === 'null') continue;
            const sku = (raw1[i][iSku] || '').toString().trim();
            const planned = iPlanned !== -1 ? parseInt(raw1[i][iPlanned]) : NaN;
            fnskuRow[fnsku] = products.length;
            products.push({ fnsku, sku, planned });
        }
        addLog('log-amazon', `Excel①から ${products.length} 件の商品を読み込みました。`);
        updateProgress('prog-amazon', 70);

        // Actual totals
        const actualTotals = {};
        boxNums.forEach(n => Object.keys(qtyByBox[n]).forEach(f => { actualTotals[f] = (actualTotals[f] || 0) + qtyByBox[n][f]; }));

        // ---------- Table ①: Validation ----------
        const vtbl = document.getElementById('validation-table');
        vtbl.innerHTML = '';
        const vth = vtbl.createTHead().insertRow();
        ['FNSKU', 'SKU', '予定数量', '実績合計', '差異', '状態'].forEach((h, i) => {
            const th = document.createElement('th'); th.textContent = h; th.className = i < 2 ? 'left' : ''; vth.appendChild(th);
        });
        const vtb = vtbl.createTBody();
        let allMatch = true;
        products.forEach(p => {
            const actual = actualTotals[p.fnsku] || 0;
            const diff = !isNaN(p.planned) ? actual - p.planned : null;
            const ok = diff === null ? null : diff === 0;
            if (ok === false) allMatch = false;
            const tr = vtb.insertRow();
            const addCell = (txt, cls) => { const td = tr.insertCell(); td.textContent = txt; td.className = cls || ''; };
            addCell(p.fnsku, 'left'); addCell(p.sku, 'left');
            addCell(p.planned != null && !isNaN(p.planned) ? p.planned : '−');
            addCell(actual);
            if (ok === null) { addCell('−'); tr.insertCell().innerHTML = '<span class="badge badge-gray">不明</span>'; }
            else if (ok) { addCell('0'); tr.insertCell().innerHTML = '<span class="badge badge-ok">✔ 一致</span>'; tr.className = 'row-ok'; }
            else { const d = tr.insertCell(); d.textContent = (diff > 0 ? '+' : '') + diff; d.style.cssText = 'color:#FF5555;font-weight:700;'; tr.insertCell().innerHTML = '<span class="badge badge-ng">✗ 不一致</span>'; tr.className = 'row-ng'; }
        });
        document.getElementById('val-summary').innerHTML = allMatch
            ? '<span class="badge badge-ok" style="font-size:.85rem;padding:6px 16px;">✔ 全商品の数量が一致しています</span>'
            : '<span class="badge badge-ng" style="font-size:.85rem;padding:6px 16px;">✗ 数量が合っていない商品があります</span>';

        // ---------- Table ②: Product × Box Quantities ----------
        const qtbl = document.getElementById('qty-table');
        qtbl.innerHTML = '';
        const qth = qtbl.createTHead().insertRow();
        ['FNSKU', 'SKU', ...boxNums.map(n => `箱${n}`)].forEach((h, i) => {
            const th = document.createElement('th'); th.textContent = h; th.className = i < 2 ? 'left' : ''; qth.appendChild(th);
        });
        const qtb = qtbl.createTBody();
        products.forEach(p => {
            const tr = qtb.insertRow();
            const f = tr.insertCell(); f.textContent = p.fnsku; f.className = 'left';
            const s = tr.insertCell(); s.textContent = p.sku; s.className = 'left';
            let hasVal = false;
            boxNums.forEach(n => {
                const td = tr.insertCell();
                const q = qtyByBox[n][p.fnsku];
                if (q) { td.textContent = q; td.className = 'value'; hasVal = true; }
                else { td.textContent = ''; td.className = 'empty'; }
            });
            if (!hasVal) tr.style.opacity = '0.3';
        });

        // ---------- Table ③: Box Size & Weight ----------
        const stbl = document.getElementById('size-table');
        stbl.innerHTML = '';
        const sth = stbl.createTHead().insertRow();
        ['項目', ...boxNums.map(n => `輸送箱${n}`)].forEach((h, i) => {
            const th = document.createElement('th'); th.textContent = h; th.className = i === 0 ? 'left' : ''; sth.appendChild(th);
        });
        const stb = stbl.createTBody();
        [{ label: '輸送箱の重量 (kg)', key: 'weight' }, { label: '輸送箱の長さ (cm)', key: 'l' },
        { label: '輸送箱の幅 (cm)', key: 'w' }, { label: '輸送箱の高さ (cm)', key: 'h' }]
            .forEach(({ label, key }) => {
                const tr = stb.insertRow();
                const th = document.createElement('td'); th.textContent = label; th.className = 'left'; tr.appendChild(th);
                boxNums.forEach(n => {
                    const td = tr.insertCell(); const v = boxMap[n][key];
                    td.textContent = v > 0 ? v : ''; td.className = v > 0 ? 'value' : 'empty';
                });
            });

        updateProgress('prog-amazon', 100);
        lucide.createIcons();
        addLog('log-amazon', '<span style="color:#AAFFAA;">✅ 変換完了！下の表をご確認ください。</span>');
    }
})();
