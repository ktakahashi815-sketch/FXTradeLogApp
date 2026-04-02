/**
 * FX 取引記録アプリ - ロジック
 */

// --- Cloud Configuration ---
// 次の Client ID を、マニュアルに従って取得したものに書き換えてください。
const GOOGLE_CLIENT_ID = '330341346642-dbq8o9c5jcev15c3cr8h89729nhhnbdo.apps.googleusercontent.com';
const MS_CLIENT_ID = 'YOUR_MS_CLIENT_ID';

document.addEventListener('DOMContentLoaded', () => {
    // 要素の取得
    const tradeForm = document.getElementById('tradeForm');
    const tradeList = document.getElementById('tradeList');
    const saveBtn = document.getElementById('saveBtn');
    const resetBtn = document.getElementById('resetBtn');
    const exportBtn = document.getElementById('exportBtn');
    const editIndexField = document.getElementById('editIndex');
    const totalProfitDisplay = document.getElementById('totalProfitDisplay');

    // 設定関連の要素
    const settingsBtn = document.getElementById('settingsBtn');
    const settingsModal = document.getElementById('settingsModal');
    const closeSettings = document.getElementById('closeSettings');
    const saveSettingsBtn = document.getElementById('saveSettingsBtn');
    const resetSettingsBtn = document.getElementById('resetSettingsBtn');
    const clearAllDataBtn = document.getElementById('clearAllDataBtn');
    const settingsContainer = document.getElementById('settingsContainer');
    const cloudStatus = document.getElementById('cloudStatus');
    const googleLoginBtn = document.getElementById('googleLoginBtn');
    const onedriveLoginBtn = document.getElementById('onedriveLoginBtn');
    const cloudFolderPicker = document.getElementById('cloudFolderPicker');
    const folderList = document.getElementById('folderList');
    const confirmFolderBtn = document.getElementById('confirmFolderBtn');

    // フォームフィールド
    const fields = {
        date: document.getElementById('date'),
        no: document.getElementById('no'),
        entryTime: document.getElementById('entryTime'),
        serverTime: document.getElementById('serverTime'),
        exitTime: document.getElementById('exitTime'),
        currency: document.getElementById('currency'),
        lot: document.getElementById('lot'),
        side: document.getElementById('side'),
        entPrice: document.getElementById('entPrice'),
        extPrice: document.getElementById('extPrice'),
        pips: document.getElementById('pips'),
        profit: document.getElementById('profit'),
        entryReason: document.getElementById('entryReason'),
        exitReason: document.getElementById('exitReason'),
        drbOs: document.getElementById('drbOs'),
        shortRci: document.getElementById('shortRci'),
        midRci: document.getElementById('midRci'),
        longRci: document.getElementById('longRci'),
        totalRci: document.getElementById('totalRci'),
        remarks: document.getElementById('remarks')
    };

    const analysisKeys = ["entryReason", "exitReason", "drbOs", "shortRci", "midRci", "longRci", "totalRci"];

    const DEFAULT_SETTINGS = {
        currency: ["EUR/JPY", "USD/JPY", "GBP/JPY", "EUR/USD", "CAD/JPY"],
        lot: ["1", "2", "3", "4", "5", "10"],
        side: ["買", "売"],
        entryReason: ["短期正規ターン", "中期正規ターン", "大山小山", "ぽこぽこ", "お椀ぐつ", "分析項目なし"],
        exitReason: ["利確", "損切", "時間切れ", "分析項目なし"],
        drbOs: ["DRB OS 無", "DRB x_OS", "DRB a_OS", "DRB b_OS", "DRB c_OS", "DRB 全部 OS"],
        shortRci: ["正規ターン", "過熱圏張り付き", "ガタガタ", "分析項目なし"],
        midRci: ["正規ターン", "過熱圏張り付き", "ガタガタ", "分析項目なし"],
        longRci: ["正規ターン", "過熱圏張り付き", "ガタガタ", "分析項目なし"],
        totalRci: ["正規ターン", "過熱圏張り付き", "ガタガタ", "分析項目なし"],
        cyprusOffset: -6
    };

    let trades = JSON.parse(localStorage.getItem('fx_trades')) || [];
    let settings = JSON.parse(localStorage.getItem('fx_settings')) || DEFAULT_SETTINGS;
    let cloudConfig = JSON.parse(localStorage.getItem('fx_cloud_config')) || { service: null, folderId: null, folderName: null };
    let selectedFolder = null;

    // --- Microsoft Auth Setup ---
    let msalInstance = null;
    try {
        const msalConfig = {
            auth: {
                clientId: MS_CLIENT_ID,
                redirectUri: window.location.origin + window.location.pathname
            }
        };
        msalInstance = new msal.PublicClientApplication(msalConfig);
    } catch (e) { console.error('MSAL init failed', e); }

    /**
     * プルダウン生成
     */
    function renderSelects() {
        const selectKeys = ["currency", "lot", "side", "entryReason", "exitReason", "drbOs", "shortRci", "midRci", "longRci", "totalRci"];
        selectKeys.forEach(key => {
            const selectEl = document.getElementById(key);
            if (!selectEl) return;
            const currentValue = selectEl.value;
            selectEl.innerHTML = '';
            if (analysisKeys.includes(key)) {
                const emptyOption = document.createElement('option');
                emptyOption.value = "";
                emptyOption.textContent = "";
                selectEl.appendChild(emptyOption);
            }
            const opts = settings[key] || DEFAULT_SETTINGS[key];
            opts.forEach(opt => {
                const option = document.createElement('option');
                option.value = opt;
                option.textContent = opt;
                selectEl.appendChild(option);
            });
            if (currentValue && opts.includes(currentValue)) {
                selectEl.value = currentValue;
            } else if (analysisKeys.includes(key)) {
                selectEl.value = "";
            }
        });
    }

    /**
     * 合計損益
     */
    function updateTotalProfit() {
        const sum = trades.reduce((acc, t) => acc + (parseFloat(t.profit) || 0), 0);
        totalProfitDisplay.textContent = sum.toLocaleString('ja-JP', { minimumFractionDigits: 1, maximumFractionDigits: 4 });
        if (sum < 0) totalProfitDisplay.classList.add('negative-input');
        else totalProfitDisplay.classList.remove('negative-input');
    }

    /**
     * 設定画面のレンダリング
     */
    function renderSettingsFields() {
        settingsContainer.innerHTML = '';
        const labels = {
            currency: "通貨ペア", lot: "LOT", side: "売買",
            entryReason: "エントリー理由", exitReason: "エグジット理由",
            drbOs: "DRB OS", shortRci: "短期 RCI", midRci: "中期 RCI",
            longRci: "長期 RCI", totalRci: "全体 RCI"
        };
        const textAreaKeys = ["currency", "lot", "side", "entryReason", "exitReason", "drbOs", "shortRci", "midRci", "longRci", "totalRci"];
        textAreaKeys.forEach(key => {
            const div = document.createElement('div');
            div.className = 'setting-item';
            div.innerHTML = `<label>${labels[key] || key}</label><textarea id="set_${key}">${(settings[key] || DEFAULT_SETTINGS[key]).join('\n')}</textarea>`;
            settingsContainer.appendChild(div);
        });
        if (settings.cyprusOffset === -7) document.getElementById('offsetWinter').checked = true;
        else document.getElementById('offsetSummer').checked = true;

        updateCloudUI();
    }

    /**
     * クラウド保存 UI の更新
     */
    function updateCloudUI() {
        if (cloudConfig.service) {
            cloudStatus.textContent = `${cloudConfig.service} 連携済み: ${cloudConfig.folderName || 'フォルダ未選択'}`;
            cloudStatus.style.color = 'var(--accent-color)';
        } else {
            cloudStatus.textContent = '未連携';
            cloudStatus.style.color = 'var(--text-secondary)';
        }
    }

    // --- Google Drive Logic ---
    let googleTokenClient;
    function initGoogleAuth() {
        try {
            googleTokenClient = google.accounts.oauth2.initTokenClient({
                client_id: GOOGLE_CLIENT_ID,
                scope: 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/drive.metadata.readonly',
                callback: (response) => {
                    if (response.access_token) {
                        cloudConfig.service = 'Google Drive';
                        cloudConfig.accessToken = response.access_token;
                        listGoogleFolders();
                    }
                },
            });
        } catch (e) { console.error('Google Auth init failed', e); }
    }

    async function listGoogleFolders() {
        try {
            const response = await fetch('https://www.googleapis.com/drive/v3/files?q=mimeType=\'application/vnd.google-apps.folder\'+and+trashed=false&fields=files(id,name)&pageSize=100', {
                headers: { 'Authorization': `Bearer ${cloudConfig.accessToken}` }
            });
            const data = await response.json();
            
            // 「マイドライブ (ルート)」をリストの先頭に追加
            const folders = [{ id: 'root', name: '📁 マイドライブ (ルート)' }, ...(data.files || [])];
            renderFolderList(folders);
        } catch (e) {
            console.error('Google list folders failed', e);
            alert('フォルダ一覧の取得に失敗しました。');
        }
    }

    async function uploadToGoogleDrive(csvContent, fileName) {
        const metadata = { name: fileName, parents: [cloudConfig.folderId] };
        const form = new FormData();
        form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
        form.append('file', new Blob([csvContent], { type: 'text/csv' }));
        const response = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${cloudConfig.accessToken}` },
            body: form
        });
        return response.ok;
    }

    // --- OneDrive Logic ---
    async function handleMSLogin() {
        try {
            const loginResponse = await msalInstance.loginPopup({ scopes: ["Files.ReadWrite", "User.Read"] });
            cloudConfig.service = 'OneDrive';
            cloudConfig.accessToken = loginResponse.accessToken;
            listMSFolders();
        } catch (err) { console.error(err); }
    }

    async function listMSFolders() {
        const response = await fetch('https://graph.microsoft.com/v1.0/me/drive/root/children?$filter=folder ne null', {
            headers: { 'Authorization': `Bearer ${cloudConfig.accessToken}` }
        });
        const data = await response.json();
        const folders = (data.value || []).map(f => ({ id: f.id, name: f.name }));
        renderFolderList(folders);
    }

    async function uploadToOneDrive(csvContent, fileName) {
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${cloudConfig.folderId}:/${fileName}:/content`;
        const response = await fetch(url, {
            method: 'PUT',
            headers: { 'Authorization': `Bearer ${cloudConfig.accessToken}`, 'Content-Type': 'text/csv' },
            body: csvContent
        });
        return response.ok;
    }

    /**
     * フォルダリスト表示
     */
    function renderFolderList(folders) {
        folderList.innerHTML = '';
        folders.forEach(f => {
            const li = document.createElement('li');
            li.className = 'folder-item';
            li.textContent = f.name;
            li.onclick = () => {
                document.querySelectorAll('.folder-item').forEach(el => el.classList.remove('selected'));
                li.classList.add('selected');
                selectedFolder = f;
            };
            folderList.appendChild(li);
        });
        cloudFolderPicker.classList.remove('hidden');
    }

    confirmFolderBtn.onclick = () => {
        if (!selectedFolder) { alert('フォルダを選択してください'); return; }
        cloudConfig.folderId = selectedFolder.id;
        cloudConfig.folderName = selectedFolder.name;
        localStorage.setItem('fx_cloud_config', JSON.stringify(cloudConfig));
        updateCloudUI();
        cloudFolderPicker.classList.add('hidden');
        alert(`保存先を「${selectedFolder.name}」に決定しました。`);
    };

    googleLoginBtn.onclick = () => {
        if (!googleTokenClient) initGoogleAuth();
        if (googleTokenClient) googleTokenClient.requestAccessToken();
    };

    onedriveLoginBtn.onclick = handleMSLogin;

    // --- Core logic ---
    function updateServerTime() {
        const val = fields.entryTime.value;
        if (!val) return;
        const [h, m] = val.split(':').map(Number);
        let sh = h + (settings.cyprusOffset || -6);
        if (sh < 0) sh += 24; if (sh >= 24) sh -= 24;
        fields.serverTime.value = String(sh).padStart(2, '0') + ':' + String(m).padStart(2, '0');
    }

    function calculateAuto() {
        const ent = parseFloat(fields.entPrice.value) || 0;
        const ext = parseFloat(fields.extPrice.value) || 0;
        const lot = parseFloat(fields.lot.value) || 1;
        const side = fields.side.value;
        let pips = side === '買' ? ext - ent : ent - ext;
        let profit = pips * lot;
        fields.pips.value = parseFloat(pips.toFixed(4));
        fields.profit.value = parseFloat(profit.toFixed(4));
        if (pips < 0) fields.pips.classList.add('negative-input'); else fields.pips.classList.remove('negative-input');
        if (profit < 0) fields.profit.classList.add('negative-input'); else fields.profit.classList.remove('negative-input');
    }

    [fields.entPrice, fields.extPrice, fields.lot, fields.side].forEach(el => {
        el.addEventListener('input', calculateAuto);
        el.addEventListener('change', calculateAuto);
    });

    tradeForm.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && e.target.tagName === 'INPUT') {
            e.preventDefault();
            const now = new Date().toTimeString().slice(0, 5);
            if (e.target.id === 'entPrice') { fields.entryTime.value = now; updateServerTime(); }
            else if (e.target.id === 'extPrice') { fields.exitTime.value = now; }
        }
    });

    fields.entryTime.addEventListener('change', updateServerTime);

    function init() {
        renderSelects();
        if (!fields.date.value) fields.date.value = new Date().toISOString().split('T')[0];
        updateAutoNo(); updateTotalProfit(); renderList();
    }

    function updateAutoNo() {
        if (parseInt(editIndexField.value) === -1) {
            const next = trades.length > 0 ? Math.max(...trades.map(t => parseInt(t.no) || 0)) + 1 : 1;
            fields.no.value = next;
        }
    }

    saveSettingsBtn.onclick = () => {
        const textAreaKeys = ["currency", "lot", "side", "entryReason", "exitReason", "drbOs", "shortRci", "midRci", "longRci", "totalRci"];
        textAreaKeys.forEach(k => {
            settings[k] = document.getElementById(`set_${k}`).value.split('\n').map(s => s.trim()).filter(s => s !== '');
        });
        settings.cyprusOffset = parseInt(document.querySelector('input[name="offsetType"]:checked').value);
        localStorage.setItem('fx_settings', JSON.stringify(settings));
        renderSelects(); updateServerTime(); settingsModal.classList.add('hidden');
        alert('設定を保存しました');
    };

    resetSettingsBtn.onclick = () => {
        if (confirm('リセットしますか？')) {
            settings = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
            localStorage.setItem('fx_settings', JSON.stringify(settings));
            renderSettingsFields(); renderSelects(); updateServerTime();
        }
    };

    clearAllDataBtn.onclick = () => {
        if (confirm('全消去しますか？')) { trades = []; localStorage.setItem('fx_trades', JSON.stringify(trades)); tradeForm.reset(); init(); }
    };

    tradeForm.onsubmit = (e) => {
        e.preventDefault();
        const lastExt = fields.extPrice.value;
        const data = {}; for (let k in fields) data[k] = fields[k].value;
        const idx = parseInt(editIndexField.value);
        if (idx > -1) trades[idx] = data; else trades.push(data);
        localStorage.setItem('fx_trades', JSON.stringify(trades));
        tradeForm.reset(); fields.entPrice.value = lastExt; fields.extPrice.value = lastExt;
        analysisKeys.forEach(k => fields[k].value = ""); init(); alert('保存しました');
    };

    exportBtn.onclick = async () => {
        if (trades.length === 0) return alert('データがありません');
        const csv = generateCSV();
        const ymd = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        const fileName = `FX${ymd}DATA.csv`;

        if (cloudConfig.service && cloudConfig.folderId && confirm(`「${cloudConfig.folderName}」にクラウド保存しますか？`)) {
            try {
                let ok = false;
                if (cloudConfig.service === 'Google Drive') ok = await uploadToGoogleDrive(csv, fileName);
                else if (cloudConfig.service === 'OneDrive') ok = await uploadToOneDrive(csv, fileName);
                if (ok) { alert('クラウドに保存しました！'); return; }
                else throw new Error();
            } catch (e) { alert('クラウド保存に失敗しました。ダウンロードに切り替えます。'); }
        }

        const b = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(b); a.download = fileName; a.click();
    };

    function generateCSV() {
        const dates = trades.map(t => new Date(`${t.date} ${t.entryTime}`));
        const range = `${new Date(Math.min(...dates)).toLocaleString()} - ${new Date(Math.max(...dates)).toLocaleString()}`;
        let csv = "\ufeffFX取引記録\n開始時刻・終了時刻," + range + "\n";
        csv += "日付,No,ENTRY時刻,サーバー時刻,EXIT時刻,通貨ペア,LOT,売買,ENT値,EXT値,pips,損益,エントリー理由,エグジット理由,DRB OS,短期RCI,中期RCI,長期RCI,全体RCI,備考\n";
        trades.forEach(t => {
            const row = [t.date, t.no, t.entryTime, t.serverTime, t.exitTime, t.currency, t.lot, t.side, t.entPrice, t.extPrice, t.pips, t.profit, t.entryReason, t.exitReason, t.drbOs, t.shortRci, t.midRci, t.longRci, t.totalRci, (t.remarks || "").replace(/\n/g, " ")];
            csv += row.map(v => `"${v}"`).join(",") + "\n";
        });
        return csv;
    }

    function renderList() {
        tradeList.innerHTML = '';
        [...trades].reverse().forEach((t, i) => {
            const aidx = trades.length - 1 - i;
            const neg = parseFloat(t.profit) < 0;
            const card = document.createElement('div'); card.className = 'trade-card';
            card.innerHTML = `
                <div class="trade-header"><span>No.${t.no}</span><span>${t.date} ${t.entryTime}</span></div>
                <div class="trade-main"><div>${t.currency}</div><div class="${t.side === '買' ? 'side-buy' : 'side-sell'}">${t.side} / ${t.lot} LOT</div></div>
                <div class="trade-stats"><span>pips: <span class="${neg ? 'negative' : ''}">${t.pips}</span></span><span class="${neg ? 'negative' : 'side-buy'}">損益: ${t.profit}</span></div>
                <div class="trade-actions"><button onclick="editTrade(${aidx})">編集</button><button onclick="deleteTrade(${aidx})">削除</button></div>
            `;
            tradeList.appendChild(card);
        });
    }

    window.editTrade = (i) => {
        const t = trades[i]; for (let k in fields) if (fields[k]) fields[k].value = t[k] || '';
        editIndexField.value = i; saveBtn.textContent = '更新する'; calculateAuto();
        window.scrollTo({ top: 0, behavior: 'smooth' });
    };

    window.deleteTrade = (i) => {
        if (confirm('削除しますか？')) { trades.splice(i, 1); localStorage.setItem('fx_trades', JSON.stringify(trades)); init(); }
    };

    settingsBtn.onclick = () => { renderSettingsFields(); settingsModal.classList.remove('hidden'); };
    closeSettings.onclick = () => settingsModal.classList.add('hidden');
    init();
});
