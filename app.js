document.addEventListener('DOMContentLoaded', () => {
    const newFileInput = document.getElementById('newFile');
    const oldFileInput = document.getElementById('oldFile');
    const newSheetSelect = document.getElementById('newSheetSelect');
    const oldSheetSelect = document.getElementById('oldSheetSelect');
    const mergeBtn = document.getElementById('mergeBtn');
    const statusDiv = document.getElementById('status');
    const exportFileNameInput = document.getElementById('exportFileName');

    let newWorkbook = null;
    let oldWorkbook = null;
    
    const labelFileInput = document.getElementById('labelFile');
    const labelSheetSelect = document.getElementById('labelSheetSelect');
    const labelColumnInput = document.getElementById('labelColumn');
    const labelValueInput = document.getElementById('labelValue');
    const exportWordBtn = document.getElementById('exportWordBtn');
    let labelWorkbook = null;

    /**
     * 處理檔案上傳，使用 SheetJS 解析並產出工作表選單
     */
    async function handleFileUpload(file, setWorkbook, selectElement) {
        if (!file) return;
        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data, { type: 'array' });
            setWorkbook(workbook);
            
            selectElement.innerHTML = '';
            workbook.SheetNames.forEach(name => {
                const option = document.createElement('option');
                option.value = name;
                option.textContent = name;
                selectElement.appendChild(option);
            });
            selectElement.style.display = 'block';
        } catch (e) {
            console.error(e);
            showStatus('讀取檔案失敗：' + e.message, 'error');
        }
    }

    newFileInput.addEventListener('change', (e) => {
        handleFileUpload(e.target.files[0], (wb) => newWorkbook = wb, newSheetSelect);
    });

    oldFileInput.addEventListener('change', (e) => {
        handleFileUpload(e.target.files[0], (wb) => oldWorkbook = wb, oldSheetSelect);
    });

    labelFileInput.addEventListener('change', (e) => {
        handleFileUpload(e.target.files[0], (wb) => labelWorkbook = wb, labelSheetSelect);
    });

    mergeBtn.addEventListener('click', () => {
        if (!newWorkbook || !oldWorkbook) {
            showStatus('請同時上傳「今年度新版公版」與「上一期舊版紀錄」！', 'error');
            return;
        }

        // 先更新 UI 狀態
        showStatus('資料處理中，請稍候...', 'info');

        // 使用 setTimeout 延遲執行，讓瀏覽器有時間渲染上面的「資料處理中」畫面
        setTimeout(() => {
            try {
                const newSheetName = newSheetSelect.value;
                const oldSheetName = oldSheetSelect.value;

                if (!newSheetName || !oldSheetName) {
                    throw new Error("無法取得選擇的工作表，請確認檔案是否正確解析。");
                }

                // 取出 2D 陣列資料，保留原格式文字(raw:false)，空值補空字串(defval:"")
                const newDataRaw = XLSX.utils.sheet_to_json(newWorkbook.Sheets[newSheetName], { header: 1, raw: false, defval: "" });
                const oldDataRaw = XLSX.utils.sheet_to_json(oldWorkbook.Sheets[oldSheetName], { header: 1, raw: false, defval: "" });

                const newParsed = extractDataAndHeaders(newDataRaw);
                const oldParsed = extractDataAndHeaders(oldDataRaw);

                if (!newParsed || !oldParsed) {
                    throw new Error("無法解析資料，請檢查工作表內容是否包含標題列（必須包含「財產編號」或「物品編號」）。\n(提示: 可能是選錯了工作表，或是表頭有隱藏的特殊字元)");
                }

                const mergedResult = mergeData(newParsed, oldParsed);
                
                let exportName = exportFileNameInput.value.trim();
                if (!exportName) {
                    exportName = '合併後盤點表_輸出';
                }
                
                exportXLSX(mergedResult.data, mergedResult.headers, exportName);
                
                showStatus('合併成功！已自動下載 XLSX 結果。', 'success');
            } catch (error) {
                console.error(error);
                showStatus('處理失敗：' + error.message, 'error');
                // 加入 alert 當作保底提示，防範 CSS 沒正常顯示
                alert('發生錯誤：\n' + error.message);
            }
        }, 50);
    });

    /**
     * 尋找標題列，並將二維陣列轉換為 Object 陣列
     */
    function extractDataAndHeaders(rawData) {
        let headerRowIndex = -1;
        
        for (let i = 0; i < rawData.length; i++) {
            const row = rawData[i];
            if (!row || !Array.isArray(row)) continue;
            
            // 寬鬆比對：去除空白後檢查是否包含關鍵字，避免因為 Excel 儲存格有空白導致比對失敗
            const hasHeader = row.some(cell => {
                const str = String(cell || '').replace(/[\s\uFEFF\xA0\u200B]+/g, '');
                return str.includes('財產編號') || str.includes('物品編號');
            });

            if (hasHeader) {
                headerRowIndex = i;
                break;
            }
        }

        if (headerRowIndex === -1) return null;

        // 徹底清除表頭的隱藏空白與零寬字元，確保後續 key 取值絕對正確
        const headers = rawData[headerRowIndex].map(h => h ? String(h).replace(/[\s\uFEFF\xA0\u200B]+/g, '') : '');
        
        const data = [];
        for (let i = headerRowIndex + 1; i < rawData.length; i++) {
            const rowArray = rawData[i];
            
            if (!rowArray || !Array.isArray(rowArray)) continue;

            // 忽略完全空白的列
            if (rowArray.every(cell => !cell || String(cell).trim() === '')) continue;
            
            const rowObj = {};
            for (let j = 0; j < headers.length; j++) {
                if (headers[j]) {
                    rowObj[headers[j]] = rowArray[j] !== undefined ? String(rowArray[j]) : '';
                }
            }
            data.push(rowObj);
        }

        return { headers, data };
    }

    /**
     * 核心比對與合併邏輯
     */
    function mergeData(newParsed, oldParsed) {
        // 終極清理：移除所有常規空白、全形空白、零寬字元(Zero-width space)，並轉大寫 (解決大小寫差異)
        const sanitizeKeyPart = (val) => val ? String(val).replace(/[\s\uFEFF\xA0\u200B]+/g, '').toUpperCase() : '';
        
        const normalizeNum = (val) => {
            const str = sanitizeKeyPart(val);
            // 解決 Excel 自動去掉純數字開頭 0 的問題 (例如 00045 變 45)
            if (/^\d+$/.test(str)) {
                return str.replace(/^0+/, '') || '0';
            }
            return str;
        };

        const getValidValue = (row, cols) => {
            for (let col of cols) {
                const val = sanitizeKeyPart(row[col]);
                // 若值存在，且不是無效佔位符 (如橫線、NA)，則採用
                if (val && val !== '-' && val !== 'NA' && val !== '無') {
                    return val;
                }
            }
            return '';
        };

        // 絕對嚴格限制只取「編號」與「分號」
        const getNewKey = (row) => {
            const no = getValidValue(row, ['財產編號', '物品編號']);
            const sub = getValidValue(row, ['財產分號', '財產序號']);
            return `${no}_${normalizeNum(sub)}`;
        };
            
        const getOldKey = (row) => {
            const no = getValidValue(row, ['財產編號', '物品編號']);
            // 舊版非消耗品可能是「物品分號」或「財產序號」
            const sub = getValidValue(row, ['財產分號', '物品分號', '財產序號']);
            return `${no}_${normalizeNum(sub)}`;
        };

        const oldMap = new Map();
        const unmatchedKeys = new Set();
        oldParsed.data.forEach(row => {
            const no = getValidValue(row, ['財產編號', '物品編號']);
            if (no) { // 確保確實有編號才視為有效資料
                const key = getOldKey(row);
                oldMap.set(key, row);
                unmatchedKeys.add(key);
            }
        });

        const customCols = oldParsed.headers.filter(h => h && !newParsed.headers.includes(h));
        const overwriteColumns = ['檢查情形', '存置地點'];
        const CONFLICT_COL = '備註(衝突紀錄)';
        
        const finalHeaders = [...newParsed.headers, ...customCols, CONFLICT_COL];
        const finalData = [];

        newParsed.data.forEach(newRow => {
            const no = getValidValue(newRow, ['財產編號', '物品編號']);
            if (!no) return; // 跳過沒有編號的空白列

            const key = getNewKey(newRow);
            const oldRow = oldMap.get(key);
            
            const mergedRow = { ...newRow };
            const conflictLogs = [];
            
            if (oldRow) {
                unmatchedKeys.delete(key);

                overwriteColumns.forEach(col => {
                    if (newParsed.headers.includes(col) && oldParsed.headers.includes(col)) {
                        const oldVal = oldRow[col] ? oldRow[col].trim() : '';
                        const newVal = newRow[col] ? newRow[col].trim() : '';
                        
                        if (oldVal !== '') {
                            if (newVal !== '' && oldVal !== newVal) {
                                conflictLogs.push(`[${col}] 系統原值：${newVal}`);
                            }
                            mergedRow[col] = oldRow[col];
                        }
                    }
                });

                customCols.forEach(col => {
                    mergedRow[col] = oldRow[col] || '';
                });
            } else {
                customCols.forEach(col => {
                    mergedRow[col] = '';
                });
            }

            mergedRow[CONFLICT_COL] = conflictLogs.length > 0 ? conflictLogs.join('；') : '';
            finalData.push(mergedRow);
        });

        unmatchedKeys.forEach(key => {
            const oldRow = oldMap.get(key);
            const orphanedRow = {};
            
            finalHeaders.forEach(header => {
                let val = oldRow[header];
                // 處理非消耗品新舊表頭名稱差異：新版的「財產序號」等於舊版的「物品分號」
                if (header === '財產序號' && val === undefined) {
                    val = oldRow['物品分號'];
                }
                orphanedRow[header] = val !== undefined ? val : '';
            });

            if (finalHeaders.includes('存置地點')) orphanedRow['存置地點'] = '';
            if (finalHeaders.includes('使用人')) orphanedRow['使用人'] = '';

            if (finalHeaders.includes('檢查情形')) {
                orphanedRow['檢查情形'] = '【系統已移除，無須盤點】';
            }

            orphanedRow[CONFLICT_COL] = '此項目已不存在於本年度公版中，可能已報廢或減損。';
            finalData.push(orphanedRow);
        });

        return { headers: finalHeaders, data: finalData };
    }

    /**
     * 以原生 XLSX 格式匯出
     */
    function exportXLSX(data, headers, filename) {
        // 利用 header 參數強制約束欄位輸出順序
        const worksheet = XLSX.utils.json_to_sheet(data, { header: headers });
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, worksheet, "盤點結果");
        
        // 加上 .xlsx 副檔名
        const exportFullName = filename.endsWith('.xlsx') ? filename : filename + '.xlsx';
        XLSX.writeFile(newWorkbook, exportFullName);
    }

    /**
     * 顯示狀態訊息
     */
    function showStatus(message, type) {
        statusDiv.textContent = message;
        statusDiv.className = `status ${type}`;
        statusDiv.style.display = 'block';
    }

    // ==========================================
    // 標籤補發申請單匯出邏輯 (匯出 Word)
    // ==========================================
    exportWordBtn.addEventListener('click', () => {
        if (!labelWorkbook) {
            showStatus('請先上傳已盤點完畢的 Excel 檔案！', 'error');
            return;
        }

        const sheetName = labelSheetSelect.value;
        if (!sheetName) {
            showStatus('無法取得工作表，請重新上傳檔案！', 'error');
            return;
        }

        const filterCol = labelColumnInput.value.trim();
        if (!filterCol) {
            showStatus('請輸入判斷欄位 (例如: 標籤)！', 'error');
            return;
        }

        const filterVal = labelValueInput.value.trim();
        
        showStatus('正在產生 Word 文件，請稍候...', 'info');

        setTimeout(() => {
            try {
                // 1. 解析 Excel
                const rawData = XLSX.utils.sheet_to_json(labelWorkbook.Sheets[sheetName], { header: 1, raw: false, defval: "" });
                const parsed = extractDataAndHeaders(rawData);

                if (!parsed) {
                    throw new Error("無法解析資料，請確認檔案格式是否正確。");
                }

                // 2. 過濾資料
                const matchedItems = [];
                parsed.data.forEach(row => {
                    const cellValue = String(row[filterCol] || '').trim();
                    if (filterVal) {
                        // 有指定特定值
                        if (cellValue === filterVal) {
                            matchedItems.push(row);
                        }
                    } else {
                        // 沒有指定特定值，非空即算
                        if (cellValue !== '') {
                            matchedItems.push(row);
                        }
                    }
                });

                if (matchedItems.length === 0) {
                    throw new Error(`在「${filterCol}」欄位中找不到符合條件的項目。`);
                }

                // 3. 產生 Word HTML 結構
                generateWordHTMLAndDownload(matchedItems);

            } catch (error) {
                console.error(error);
                showStatus('產生 Word 失敗：' + error.message, 'error');
                alert('發生錯誤：\n' + error.message);
            }
        }, 50);
    });

    function generateWordHTMLAndDownload(items) {
        // 定義表格樣式與 Word 特定排版設定 (包含 1.4cm 邊界與頁尾)
        let htmlStr = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head>
            <meta charset="utf-8">
            <style>
                @page WordSection1 {
                    size: 21.0cm 29.7cm;
                    margin: 1.4cm 1.4cm 1.4cm 1.4cm;
                    mso-header-margin: 1.4cm;
                    mso-footer-margin: 1.4cm;
                    mso-paper-source: 0;
                    mso-footer: f1;
                }
                div.WordSection1 { page: WordSection1; }
                body { font-family: "標楷體", "Times New Roman", serif; }
                .title { text-align: center; font-size: 24pt; font-weight: bold; margin-bottom: 20px; font-family: "標楷體"; }
                .unit { font-size: 14pt; margin-bottom: 10px; font-family: "標楷體"; }
                table { border-collapse: collapse; width: 100%; border: 1px solid black; }
                th, td { border: 1px solid black; padding: 5px; text-align: left; font-size: 12pt; vertical-align: middle; }
                th { text-align: center; }
                .signature { margin-top: 50px; font-size: 14pt; font-family: "標楷體"; }
                .footer { text-align: center; font-size: 12pt; font-family: "Times New Roman", serif; }
            </style>
        </head>
        <body>
            <div class="WordSection1">
                <div class="title">國立中興大學財產標籤補發申請單</div>
                <div class="unit">單位：</div>
                <table>
                    <tr>
                        <th width="20%">財產編號<br>財產分號</th>
                        <th width="30%">財產別名</th>
                        <th width="15%">購置日期</th>
                        <th width="15%">價值</th>
                        <th width="20%">備註</th>
                    </tr>
        `;

        // 填入資料列 (單列)
        items.forEach(item => {
            const propNo = (item['財產編號'] || item['物品編號'] || '').trim();
            const propSub = (item['財產分號'] || item['物品分號'] || item['財產序號'] || '').trim();
            const propAlias = (item['財產別名'] || item['物品別名'] || '').trim();
            const purchaseDate = (item['取得日期'] || item['購置日期'] || '').trim();
            const value = (item['總價'] || item['價值'] || '').trim();
            const remark = (item['型式'] || item['備註'] || '').trim();

            htmlStr += `
                    <tr>
                        <td align="center" style="font-family: 'Times New Roman', serif;">${propNo}<br>${propSub}</td>
                        <td>${propAlias}</td>
                        <td align="center" style="font-family: 'Times New Roman', serif;">${purchaseDate}</td>
                        <td align="right" style="font-family: 'Times New Roman', serif;">${value}</td>
                        <td>${remark}</td>
                    </tr>
            `;
        });

        htmlStr += `
                </table>
                <div class="signature">申請人(財管人)簽章：</div>
                
                <!-- 頁尾定義 (加入頁碼) -->
                <div style="mso-element:footer" id="f1">
                    <p class="footer">
                        <!--[if supportFields]><span style="mso-element:field-begin"></span> PAGE <span style="mso-element:field-separator"></span><![endif]-->
                        <span style="mso-no-proof:yes"></span>
                        <!--[if supportFields]><span style="mso-element:field-end"></span><![endif]-->
                    </p>
                </div>
            </div>
        </body>
        </html>
        `;

        // 產生檔案並下載
        const blob = new Blob(['\\ufeff', htmlStr], { type: 'application/msword' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = '標籤補發申請單.doc';
        document.body.appendChild(a);
        a.click();
        
        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            showStatus('Word 文件已匯出！', 'success');
        }, 100);
    }
});
