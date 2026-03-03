// =============================================
// page-shipping.js – 国際送料按分計算ロジック
// =============================================

(function () {
    let wb2 = null;

    setupDropZone('dz2-shipping', 'file2-shipping', wb => {
        wb2 = wb;
        checkReady();
    });

    document.getElementById('total-shipping-cost').addEventListener('input', checkReady);

    function checkReady() {
        const cost = parseFloat(document.getElementById('total-shipping-cost').value);
        document.getElementById('convert-btn-shipping').disabled = !(wb2 && !isNaN(cost) && cost > 0);
    }

    document.getElementById('convert-btn-shipping').addEventListener('click', async () => {
        document.getElementById('convert-btn-shipping').disabled = true;
        document.getElementById('status-shipping').style.display = 'block';
        document.getElementById('log-shipping').innerHTML = '';
        try { await run(); }
        catch (err) { addLog('log-shipping', `<span style="color:#FF5555;">ERROR: ${err.message}</span>`); console.error(err); }
        finally { document.getElementById('convert-btn-shipping').disabled = false; }
    });

    async function run() {
        const totalCost = parseFloat(document.getElementById('total-shipping-cost').value);
        addLog('log-shipping', `国際送料合計: ¥${totalCost.toLocaleString()} で計算します...`);
        updateProgress('prog-shipping', 20);

        // --- Parse Excel② ---
        const boxMap = parseRakumart(wb2);
        const boxNums = Object.keys(boxMap).map(Number).sort((a, b) => a - b);
        addLog('log-shipping', `輸送箱 ${boxNums.length} 個を検出しました。`);

        // 按分方式：
        // 各箱の体積（L×W×H cm³）を計算 → 全体積に対する割合で各箱の送料を按分
        // 箱内の各商品は、個数割合で箱の送料をさらに按分

        let totalVolume = 0;
        boxNums.forEach(n => { totalVolume += boxMap[n].volume; });

        if (totalVolume === 0) {
            throw new Error('サイズデータが取得できませんでした。Excel②に箱寸法が含まれているか確認してください。');
        }

        addLog('log-shipping', `総体積: ${totalVolume.toLocaleString()} cm³`);
        updateProgress('prog-shipping', 50);

        // Per-product cost accumulator: { fnsku: { asin, qty, totalCost } }
        const productMap = {};

        boxNums.forEach(n => {
            const box = boxMap[n];
            const boxVolume = box.volume;
            const boxCost = (boxVolume / totalVolume) * totalCost; // 箱の按分送料

            const totalQtyInBox = box.items.reduce((s, it) => s + it.qty, 0);
            if (totalQtyInBox === 0) return;

            // 箱のコストを個数で等分（商品ごとの体積比は考慮せず個数比で按分）
            const costPerUnit = boxCost / totalQtyInBox;

            box.items.forEach(it => {
                if (!productMap[it.fnsku]) {
                    productMap[it.fnsku] = { fnsku: it.fnsku, asin: it.asin, totalQty: 0, totalCost: 0 };
                }
                productMap[it.fnsku].totalQty += it.qty;
                productMap[it.fnsku].totalCost += costPerUnit * it.qty;
            });
        });

        updateProgress('prog-shipping', 80);

        // --- Render table ---
        const tbl = document.getElementById('shipping-table');
        tbl.innerHTML = '';

        const thead = tbl.createTHead().insertRow();
        ['ASIN', 'FNSKU', '合計個数', '送料合計', '1個あたり送料'].forEach((h, i) => {
            const th = document.createElement('th'); th.textContent = h; th.className = i < 2 ? 'left' : ''; thead.appendChild(th);
        });

        const tbody = tbl.createTBody();
        const items = Object.values(productMap).sort((a, b) => b.totalQty - a.totalQty);

        items.forEach(p => {
            const perUnit = p.totalCost / p.totalQty;
            const tr = tbody.insertRow();
            const addCell = (txt, cls) => { const td = tr.insertCell(); td.innerHTML = txt; if (cls) td.className = cls; };
            addCell(p.asin || '−', 'left');
            addCell(p.fnsku, 'left');
            addCell(p.totalQty.toLocaleString(), 'value');
            addCell(`¥${Math.round(p.totalCost).toLocaleString()}`, 'value');
            addCell(`<strong>¥${Math.round(perUnit).toLocaleString()}</strong>`, 'value');
        });

        // Summary
        const totalUnits = items.reduce((s, p) => s + p.totalQty, 0);
        document.getElementById('ship-summary').innerHTML =
            `<span class="badge badge-ok" style="font-size:.85rem;padding:6px 16px;">
                合計 ${items.length} 商品 / ${totalUnits.toLocaleString()} 個 ／ 送料総額 ¥${Math.round(totalCost).toLocaleString()}
            </span>`;

        // Formula explanation
        const formulaEl = document.getElementById('ship-formula');
        formulaEl.style.display = 'block';
        formulaEl.className = 'formula-box';
        formulaEl.innerHTML = `
            <h4>📐 計算方法・按分ロジックの説明</h4>
            <div class="step">
                <strong>① 各箱の体積を算出</strong><br>
                箱寸法（長さ × 幅 × 高さ cm）からそれぞれの体積を計算します。<br>
                例：60cm × 40cm × 50cm = 120,000 cm³
            </div>
            <div class="step">
                <strong>② 箱ごとの送料を按分</strong><br>
                その箱の体積 ÷ 全箱の合計体積 × 国際送料合計 = 箱の負担送料<br>
                体積の大きい箱ほど、より多くの送料を負担します。
            </div>
            <div class="step">
                <strong>③ 商品ごとの送料を算出</strong><br>
                箱の負担送料 ÷ 箱内の商品総個数 = 1個あたり送料<br>
                同じ箱に入っている商品は、個数に応じて均等に按分されます。
            </div>
            <div class="step">
                <strong>④ 複数箱に入っている商品</strong><br>
                2つ以上の箱に入っている商品は、各箱での送料を合算して「送料合計」を算出し、<br>
                送料合計 ÷ 合計個数 = 1個あたり送料 として表示します。
            </div>
            <div class="step" style="border-left-color:rgba(255,153,0,0.6);">
                <strong>⚠ 前提条件</strong><br>
                本計算は「体積比例」に基づく按分です。重さベースの場合はラクマートの実際の重量請求方式と異なる場合があります。
            </div>
        `;

        updateProgress('prog-shipping', 100);
        lucide.createIcons();
        addLog('log-shipping', '<span style="color:#AAFFAA;">✅ 按分計算が完了しました！</span>');
    }
})();
