const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const https = require('https');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname)));

// 從 Google Sheets URL 提取 spreadsheet ID 和 gid
function parseGoogleSheetUrl(url) {
    const spreadsheetIdMatch = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    const gidMatch = url.match(/gid=(\d+)/);

    return {
        spreadsheetId: spreadsheetIdMatch ? spreadsheetIdMatch[1] : null,
        gid: gidMatch ? gidMatch[1] : '0'
    };
}

// 取得 Google Sheets CSV 下載 URL
function getGoogleSheetCsvUrl(url) {
    const { spreadsheetId, gid } = parseGoogleSheetUrl(url);
    if (!spreadsheetId) return null;
    return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=${gid}`;
}

// 從 URL 下載內容
function fetchUrl(url, maxRedirects = 10) {
    return new Promise((resolve, reject) => {
        if (maxRedirects <= 0) {
            reject(new Error('Too many redirects'));
            return;
        }

        const urlObj = new URL(url);
        const protocol = urlObj.protocol === 'https:' ? https : require('http');

        protocol.get(url, (response) => {
            // 處理重定向 (301, 302, 303, 307, 308)
            if ([301, 302, 303, 307, 308].includes(response.statusCode)) {
                const redirectUrl = response.headers.location;
                if (redirectUrl) {
                    // 處理相對 URL
                    const absoluteUrl = redirectUrl.startsWith('http')
                        ? redirectUrl
                        : new URL(redirectUrl, url).toString();
                    return fetchUrl(absoluteUrl, maxRedirects - 1).then(resolve).catch(reject);
                }
            }

            if (response.statusCode !== 200) {
                reject(new Error(`HTTP ${response.statusCode}`));
                return;
            }

            let data = '';
            response.on('data', chunk => data += chunk);
            response.on('end', () => resolve(data));
            response.on('error', reject);
        }).on('error', reject);
    });
}

// 解析 CSV
function parseCsv(csvText) {
    const lines = csvText.split('\n');
    const result = [];

    for (const line of lines) {
        if (!line.trim()) continue;
        // 簡單 CSV 解析（處理逗號分隔）
        const row = [];
        let current = '';
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === ',' && !inQuotes) {
                row.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        row.push(current.trim());
        result.push(row);
    }

    return result;
}

// 載入員工對照表
async function loadEmployeeMapping(url) {
    try {
        const csvUrl = getGoogleSheetCsvUrl(url);
        if (!csvUrl) {
            console.error('Invalid employee sheet URL');
            return {};
        }

        const csvText = await fetchUrl(csvUrl);
        const rows = parseCsv(csvText);

        const mapping = {};
        // 假設第一欄是姓名，第二欄是工號
        for (let i = 1; i < rows.length; i++) { // 跳過標題行
            if (rows[i][0] && rows[i][1]) {
                mapping[rows[i][0].trim()] = rows[i][1].trim();
            }
        }

        console.log('Employee mapping loaded:', Object.keys(mapping).length, 'entries');
        return mapping;
    } catch (error) {
        console.error('Error loading employee mapping:', error.message);
        return {};
    }
}

// 載入任務代碼對照表
async function loadTaskCodeMapping(url) {
    try {
        const csvUrl = getGoogleSheetCsvUrl(url);
        if (!csvUrl) {
            console.error('Invalid task code sheet URL');
            return { early: null, middle: null, late: null };
        }

        const csvText = await fetchUrl(csvUrl);
        const rows = parseCsv(csvText);

        // 尋找包含「總控」的任務代碼
        // 假設格式: 代碼, 名稱, 班別(早/中/晚)
        const result = {
            early: null,   // 早班代碼
            middle: null,  // 中班代碼
            late: null     // 晚班代碼
        };

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            // 檢查每一欄是否包含「總控」
            for (let j = 0; j < row.length; j++) {
                if (row[j] && row[j].includes('總控')) {
                    // 嘗試找出班別，代碼在第二欄 (row[1])
                    const rowText = row.join(' ');
                    const taskCode = row[1] || row[0]; // 優先使用第二欄，若無則用第一欄
                    if (rowText.includes('早') || rowText.toLowerCase().includes('morning')) {
                        result.early = taskCode;
                    } else if (rowText.includes('中') || rowText.toLowerCase().includes('middle')) {
                        result.middle = taskCode;
                    } else if (rowText.includes('晚') || rowText.toLowerCase().includes('evening') || rowText.toLowerCase().includes('night')) {
                        result.late = taskCode;
                    } else {
                        // 如果沒有明確標示班別，依序填入
                        if (!result.early) result.early = taskCode;
                        else if (!result.middle) result.middle = taskCode;
                        else if (!result.late) result.late = taskCode;
                    }
                }
            }
        }

        console.log('Task codes found:', result);
        return result;
    } catch (error) {
        console.error('Error loading task code mapping:', error.message);
        return { early: '總控-早', middle: '總控-中', late: '總控-晚' }; // 預設值
    }
}

// 判斷班別（早/中/晚）
function getShiftType(timeValue, earlyCutoff, lateCutoff) {
    const time = parseInt(timeValue, 10);
    const earlyTime = parseInt(earlyCutoff, 10);
    const lateTime = parseInt(lateCutoff, 10);

    if (isNaN(time)) return null;

    if (time <= earlyTime) {
        return 'early';   // 早班: <= earlyCutoff
    } else if (time >= lateTime) {
        return 'late';    // 晚班: >= lateCutoff
    } else {
        return 'middle';  // 中班: earlyCutoff < time < lateCutoff
    }
}

// 檢查是否為有效的班別時間（4位數字）
function isValidShiftTime(value) {
    if (value === null || value === undefined) return false;
    const str = String(value).trim();
    // 檢查是否為 3-4 位數字（如 930, 1020, 1150）
    return /^\d{3,4}$/.test(str);
}

// 取得某月的天數
function getDaysInMonth(year, month) {
    return new Date(year, month, 0).getDate();
}

// 格式化日期
function formatDate(year, month, day) {
    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

// 轉換 API
app.post('/api/convert', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.json({ success: false, error: '請上傳檔案' });
        }

        const yearMonth = req.body.yearMonth || '';
        const earlyCutoff = req.body.earlyCutoff || '1050';
        const lateCutoff = req.body.lateCutoff || '1150';
        const employeeSheetUrl = req.body.employeeSheetUrl || '';
        const taskCodeSheetUrl = req.body.taskCodeSheetUrl || '';

        // 解析年月
        let year, month;
        if (yearMonth.length === 6) {
            year = parseInt(yearMonth.substring(0, 4), 10);
            month = parseInt(yearMonth.substring(4, 6), 10);
        } else {
            const now = new Date();
            year = now.getFullYear();
            month = now.getMonth() + 2;
            if (month > 12) {
                month = 1;
                year++;
            }
        }

        const daysInMonth = getDaysInMonth(year, month);
        console.log(`Processing for ${year}-${month}, ${daysInMonth} days`);

        // 讀取 Excel 檔案
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });

        // 檢查是否有「總控」表單
        if (!workbook.SheetNames.includes('總控')) {
            return res.json({ success: false, error: '找不到「總控」表單' });
        }

        // 載入對照表
        const [employeeMapping, taskCodes] = await Promise.all([
            loadEmployeeMapping(employeeSheetUrl),
            loadTaskCodeMapping(taskCodeSheetUrl)
        ]);

        console.log('Employee mapping:', employeeMapping);
        console.log('Task codes:', taskCodes);

        // 讀取「總控」表單
        const sheet = workbook.Sheets['總控'];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        const result = [];

        // 從第二行開始（第一行是標題）
        for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
            const row = data[rowIndex];
            const employeeName = String(row[0] || '').trim();

            if (!employeeName) continue;

            // 查找工號
            const employeeId = employeeMapping[employeeName] || employeeName;

            // 從第二欄開始處理每一天
            for (let colIndex = 1; colIndex <= daysInMonth && colIndex < row.length; colIndex++) {
                const cellValue = row[colIndex];
                const cellStr = String(cellValue).trim();

                // 檢查是否為有效的班別時間
                if (isValidShiftTime(cellStr)) {
                    const day = colIndex;
                    const date = formatDate(year, month, day);
                    const shiftType = getShiftType(cellStr, earlyCutoff, lateCutoff);

                    let taskCode;
                    if (shiftType === 'early') {
                        taskCode = taskCodes.early || '總控-早';
                    } else if (shiftType === 'middle') {
                        taskCode = taskCodes.middle || '總控-中';
                    } else {
                        taskCode = taskCodes.late || '總控-晚';
                    }

                    result.push({
                        date: date,
                        employeeId: employeeId,
                        taskCode: taskCode
                    });
                }
            }
        }

        console.log(`Converted ${result.length} entries`);

        res.json({
            success: true,
            data: result
        });

    } catch (error) {
        console.error('Convert error:', error);
        res.json({ success: false, error: error.message });
    }
});

// 下載 API
app.post('/api/download', (req, res) => {
    try {
        const { data } = req.body;

        if (!data || !Array.isArray(data)) {
            return res.status(400).json({ error: '無效的資料' });
        }

        // 建立 Excel 工作表
        const wsData = [
            ['日期', '工號', '任務代碼'], // 標題行
            ...data.map(row => [row.date, row.employeeId, row.taskCode])
        ];

        const ws = XLSX.utils.aoa_to_sheet(wsData);

        // 設定欄寬
        ws['!cols'] = [
            { wch: 12 }, // 日期
            { wch: 15 }, // 工號
            { wch: 20 }  // 任務代碼
        ];

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '轉換結果');

        // 生成 buffer
        const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="converted_result.xlsx"');
        res.send(buffer);

    } catch (error) {
        console.error('Download error:', error);
        res.status(500).json({ error: error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
    console.log('Open this URL in your browser to use the tool.');
});
