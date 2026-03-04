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
        // ※ ラベル番号(FNSKU)や商品コードが無い商品(パッケージ等)は除外して計算

        // 有効な商品がある箱のみ体積計算に含める
        let totalVolume = 0;
        let skippedBoxes = [];
        let validBoxNums = [];

        boxNums.forEach(n => {
            const validItems = boxMap[n].items.filter(it => it.fnsku && it.fnsku !== 'null' && it.fnsku.trim() !== '');
            if (validItems.length > 0) {
                totalVolume += boxMap[n].volume;
                validBoxNums.push(n);
            } else {
                skippedBoxes.push(n);
            }
        });

        if (skippedBoxes.length > 0) {
            addLog('log-shipping', `⚠ 箱${skippedBoxes.join(', ')} にはラベル番号/商品コードのある商品がないため、按分対象から除外しました。`);
        }
        addLog('log-shipping', `按分対象: ${validBoxNums.length} 箱 / 全 ${boxNums.length} 箱`);

        if (totalVolume === 0) {
            throw new Error('有効な商品を含む箱のサイズデータが取得できませんでした。Excel②を確認してください。');
        }

        addLog('log-shipping', `按分対象の総体積: ${totalVolume.toLocaleString()} cm³`);
        updateProgress('prog-shipping', 50);

        // Per-product cost accumulator: { fnsku: { asin, qty, totalCost } }
        const productMap = {};

        validBoxNums.forEach(n => {
            const box = boxMap[n];
            const boxVolume = box.volume;
            const boxCost = (boxVolume / totalVolume) * totalCost; // 箱の按分送料

            // 有効な商品のみ（FNSKU有り）で個数を計算
            const validItems = box.items.filter(it => it.fnsku && it.fnsku !== 'null' && it.fnsku.trim() !== '');
            const totalQtyInBox = validItems.reduce((s, it) => s + it.qty, 0);
            if (totalQtyInBox === 0) return;

            // 箱のコストを有効商品の個数で等分
            const costPerUnit = boxCost / totalQtyInBox;

            validItems.forEach(it => {
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

        // Total row
        const totalUnits = items.reduce((s, p) => s + p.totalQty, 0);
        const totalCostSum = items.reduce((s, p) => s + p.totalCost, 0);
        const totalTr = tbody.insertRow();
        totalTr.style.cssText = 'border-top:2px solid var(--primary);background:rgba(255,153,0,0.06);';
        const addTotalCell = (txt, cls) => { const td = totalTr.insertCell(); td.innerHTML = txt; if (cls) td.className = cls; };
        addTotalCell('', '');
        addTotalCell('<strong style="color:var(--primary);">合計</strong>', 'left');
        addTotalCell(`<strong>${totalUnits.toLocaleString()}</strong>`, 'value');
        addTotalCell(`<strong>¥${Math.round(totalCostSum).toLocaleString()}</strong>`, 'value');
        addTotalCell('', '');

        // Summary
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
            <div class="step" style="border-left-color:rgba(255,153,0,0.6);">
                <strong>⓪ 前提：パッケージ等の除外</strong><br>
                ラベル番号（FNSKU）や商品コードが無い商品（パッケージ・緩衝材等）は按分対象から除外します。<br>
                その箱に有効な商品が1つもない場合、箱自体を体積計算から除外し、送料は他の箱の商品で按分します。
            </div>
            <div class="step">
                <strong>① 対象箱の体積を算出</strong><br>
                有効な商品を含む箱のみ、箱寸法（長さ × 幅 × 高さ cm）から体積を計算します。<br>
                例：60cm × 40cm × 50cm = 120,000 cm³
            </div>
            <div class="step">
                <strong>② 箱ごとの送料を按分</strong><br>
                その箱の体積 ÷ 対象箱の合計体積 × 国際送料合計 = 箱の負担送料<br>
                体積の大きい箱ほど、より多くの送料を負担します。
            </div>
            <div class="step">
                <strong>③ 商品ごとの送料を算出</strong><br>
                箱の負担送料 ÷ 箱内の有効商品の総個数 = 1個あたり送料<br>
                同じ箱に入っている有効商品は、個数に応じて均等に按分されます。
            </div>
            <div class="step">
                <strong>④ 複数箱に入っている商品</strong><br>
                2つ以上の箱に入っている商品は、各箱での送料を合算して「送料合計」を算出し、<br>
                送料合計 ÷ 合計個数 = 1個あたり送料 として表示します。
            </div>
        `;

        updateProgress('prog-shipping', 100);
        lucide.createIcons();
        addLog('log-shipping', '<span style="color:#AAFFAA;">✅ 按分計算が完了しました！</span>');
    }
})();
