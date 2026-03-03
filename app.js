// =============================================
// app.js – shared utilities and navigation
// =============================================

lucide.createIcons();

// -- Navigation --
function navigate(el) {
    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    el.classList.add('active');
    document.getElementById(el.dataset.page).classList.add('active');
    lucide.createIcons();
}

// -- Shared helpers --
function norm(str) {
    if (str === null || str === undefined) return "";
    return str.toString().replace(/[\s\u3000\r\n\t]+/g, '');
}

function addLog(logId, msg) {
    const div = document.createElement('div');
    div.innerHTML = `> ${msg}`;
    document.getElementById(logId).appendChild(div);
    document.getElementById(logId).scrollTop = 9999;
}

function updateProgress(progId, percent) {
    document.getElementById(progId).style.width = `${percent}%`;
}

// -- Drag & Drop setup --
function setupDropZone(zoneId, inputId, onLoad) {
    const zone = document.getElementById(zoneId);
    const input = document.getElementById(inputId);
    const fnameEl = zone.querySelector('.filename');

    zone.addEventListener('click', () => input.click());
    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop', e => {
        e.preventDefault(); zone.classList.remove('dragover');
        if (e.dataTransfer.files.length) processFile(e.dataTransfer.files[0]);
    });
    input.addEventListener('change', e => { if (e.target.files.length) processFile(e.target.files[0]); });

    function processFile(file) {
        if (!file.name.endsWith('.xlsx')) { alert('Excelファイル(.xlsx)を選択してください。'); return; }
        fnameEl.textContent = file.name;
        fnameEl.style.display = 'block';
        const reader = new FileReader();
        reader.onload = e => {
            const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
            onLoad(wb);
        };
        reader.readAsArrayBuffer(file);
    }
}

// -- Copy table (data only, skip N label columns) --
function copyTableData(tableId, skipCols, btnId) {
    const table = document.getElementById(tableId);
    const rows = Array.from(table.querySelectorAll('tbody tr'));
    const tsv = rows.map(row =>
        Array.from(row.querySelectorAll('td')).slice(skipCols)
            .map(c => c.textContent.trim()).join('\t')
    ).join('\n');

    const ta = document.createElement('textarea');
    ta.value = tsv;
    ta.style.cssText = 'position:fixed;top:-9999px;left:-9999px';
    document.body.appendChild(ta);
    ta.focus(); ta.select();
    let ok = false;
    try { ok = document.execCommand('copy'); } catch (_) { }
    document.body.removeChild(ta);

    const btn = document.getElementById(btnId);
    const orig = btn.innerHTML;
    if (ok) {
        btn.classList.add('copied');
        btn.innerHTML = '<i data-lucide="check" width="14" height="14"></i> コピーしました！';
        lucide.createIcons();
        setTimeout(() => { btn.classList.remove('copied'); btn.innerHTML = orig; lucide.createIcons(); }, 2500);
    } else {
        alert('コピーに失敗しました。');
    }
}

// -- Parse Rakumart Excel② helper (shared) --
function parseRakumart(wb) {
    const sheet2Name = wb.SheetNames.find(n => n.includes('梱包リスト') || n.includes('明細') || n.includes('箱子'));
    if (!sheet2Name) throw new Error('Excel②に「梱包リスト」シートが見つかりません。');
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet2Name], { header: 1, defval: null });

    // Find header row
    let hi = -1;
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
        const s = (rows[i] || []).map(c => norm(c)).join('|');
        if ((s.includes('箱NO') || s.includes('箱号')) && s.includes('梱包数')) { hi = i; break; }
    }
    if (hi === -1) throw new Error('Excel②のヘッダー行が見つかりません。');

    const h = rows[hi];
    const idx = {
        box: h.findIndex(c => ["箱NO", "箱号", "箱番号"].includes(norm(c))),
        size: h.findIndex(c => norm(c).includes('寸法') || norm(c).includes('サイズ') || norm(c).includes('规格')),
        weight: h.findIndex(c => norm(c).includes('重量')),
        qty: h.findIndex(c => ["梱包数", "商品数量", "入数"].includes(norm(c))),
        fnsku: h.findIndex(c => ["ラベル番号", "FNSKU", "商品コード"].includes(norm(c))),
        asin: h.findIndex(c => norm(c) === '商品コード' ? false : norm(c).includes('ASIN') || (c || '').toString().trim() === '商品コード'),
    };

    // Better ASIN detection: look for the column after ラベル番号 that contains B0... codes
    // From analysis: col 12 = ラベル番号 (FNSKU like X001ARG58D), col 13 = 商品コード (ASIN like B0FD331RWT)
    if (idx.asin === -1) {
        idx.asin = h.findIndex((c, i) => i > idx.fnsku && (c || '').toString().trim().length > 0 && !["梱包数", "単価", "小计", "国内運賃", "ラベル種類", "箱詰め備考"].includes(norm(c)));
    }

    const boxMap = {};
    let lastBox = null;

    for (let i = hi + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;

        const boxRaw = row[idx.box];
        if (boxRaw !== null && boxRaw !== undefined && boxRaw.toString().trim() !== '') {
            const n = parseInt(boxRaw.toString().replace(/[^0-9]/g, ''));
            if (!isNaN(n)) {
                lastBox = n;
                const weight = parseFloat(row[idx.weight]) || 0;
                const sizeStr = norm(row[idx.size]);
                let l = 0, w = 0, hh = 0;
                if (sizeStr.includes('*')) {
                    const d = sizeStr.split('*').map(v => parseFloat(v));
                    if (d.length === 3 && d.every(v => !isNaN(v))) [l, w, hh] = d;
                }
                if (!boxMap[lastBox]) boxMap[lastBox] = { items: [], weight, l, w, h: hh, volume: l * w * hh };
            }
        }
        if (lastBox === null) continue;
        const fnsku = norm(row[idx.fnsku]).trim();
        const asin = (row[idx.asin] || '').toString().trim();
        const qty = parseInt(row[idx.qty]);
        if (fnsku && fnsku !== 'null' && !isNaN(qty) && qty > 0) {
            boxMap[lastBox].items.push({ fnsku, asin, qty });
        }
    }

    return boxMap;
}
