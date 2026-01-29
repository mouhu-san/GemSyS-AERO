/**
 * GemSyS_AERO (v14.8 CAMS-First & BLH Supported)
 * * [System Change]
 * - Frequency: Hourly execution enabled (Removed odd-hour skip).
 * - Data Source: CAMS (Open-Meteo) Only. AEROS disabled.
 * - New Metric: Added 'boundary_layer_height' (BLH) for inversion layer monitoring.
 */

// ==========================================
// 0. Configuration
// ==========================================

const AIR_CONFIG = {
    SHEET_ID: '1xxi-tMACrG8qpd1l5C-D6aDUkF12RrlUru1jM0w7mOc',
    DRIVE_FOLDER_ID: '1rXhpclYBQZYjw5RVSC7Qj7Xej04tqoG4',
    TEXT_PLACEHOLDER: "null",
    UI: {
        CHECK_UPDATE: 'B2', CHECK_AI: 'E8', STATUS_MAIN: 'B5', STATUS_AI: 'B9', OUTPUT_AI: 'B10'
    },
    AREA_OFFSET: 0.015,

    // AI Configuration (Passed to Module)
    AI_MODEL: 'gemini-2.5-flash',

    TARGETS: [
        { id: 'HOME', name: 'Home', lat: 35.1946209, lon: 136.7286856 },
        { id: 'UNIV', name: 'Univ', lat: 35.0793000, lon: 136.9057000 },
        { id: 'COMMUTE', name: 'Commute', lat: 35.1697000, lon: 136.8631000 }
    ],

    // AEROS Stations (Legacy / Reference only)
    STATIONS_DB: [
        { id: '23101010', name: '名古屋(Univ)', lat: 35.1776, lon: 136.9739 },
        { id: '23208530', name: '津島(Main)', lat: 35.1737, lon: 136.7401 },
        { id: '23203010', name: '一宮(North)', lat: 35.3151, lon: 136.8084 },
        { id: '23427510', name: '飛島(South)', lat: 35.0778, lon: 136.8027 },
        { id: '23112510', name: '元塩(South)', lat: 35.0859, lon: 136.9238 },
        { id: '23105050', name: '名楽町(Hub)', lat: 35.1697, lon: 136.8631 },
        { id: '23106018', name: '若宮(Road)', lat: 35.1625, lon: 136.9083 },
        { id: '23110050', name: '富田(West)', lat: 35.1405, lon: 136.8126 },
        { id: '23112060', name: '南陽(Res)', lat: 35.1044, lon: 136.8242 },
        { id: '23111021', name: '港陽(Port)', lat: 35.1032, lon: 136.8878 }
    ],

    CAMS_PARAMS: {
        AIR: 'pm2_5,pm10,uv_index,dust,aerosol_optical_depth,ozone,nitrogen_dioxide,sulphur_dioxide,carbon_monoxide,ammonia',
        // ★Added: boundary_layer_height
        METEO: 'precipitation,weather_code,wind_gusts_10m,cloud_cover,surface_pressure,pressure_msl,temperature_2m,relative_humidity_2m,freezing_level_height,boundary_layer_height'
    },
    URLS: {
        CAMS_AIR: 'https://air-quality-api.open-meteo.com/v1/air-quality',
        CAMS_METEO: 'https://api.open-meteo.com/v1/forecast',
        AEROS: 'https://soramame.env.go.jp/soramame/api/data_search'
    },
    SHEETS: {
        DASH: 'UI_Dashboard',
        INTEGRATED: 'Risk_Tracker',
        AEROS: 'DB_Hourly_AEROS',
        CAMS: 'DB_Hourly_CAMS',
        AI_SUMMARY: 'AI_Summary'
    },
    RETENTION: { DAYS: 32, AI_LOG_DAYS: 2, AI_LOG_MAX: 40 }
};

const AI_SYSTEM_PROMPT_TEMPLATE = `
# System Instruction: GemSyS Tactical Monitor (CAMS-First)
あなたは、好中球型喘息患者を守るための環境監視システムです。
データソースは信頼性の高い広域モデル「OpenMeteo (CAMS)」を主とし、EU指令2024/2881に基づく厳格な判定を行っています。

## 判定基準 (Strict EU 2024)
- 🔴 DANGER (RED): PM2.5 >= 25.0, NO2 >= 50.0, SO2 >= 50.0
- 🟡 CAUTION (YELLOW): 2030年目標値超過 (PM2.5 >= 10, NO2 >= 20)
- 🟢 SAFE (GREEN): 安全圏

## Output Format
\`\`\`markdown
### 🛡️ Tactical Report [ {Time} ]

| Item | Value | Unit | Status |
| :--- | :--- | :--- | :--- |
| **Risk Level** | **{Risk_Level}** | - | **{Risk_Reason}** |
| PM2.5 (CAMS) | {Home_PM25} | μg/m³ | Dust: {CAMS_Dust} |
| NO2 (CAMS) | {Home_NO2} | μg/m³ | NH3: {CAMS_NH3} |
| SO2 (CAMS) | {Home_SO2} | μg/m³ | CO: {CAMS_CO} |
| Meteo | {CAMS_Temp}℃ | - | AOD: {CAMS_AOD} |
\`\`\`
※実測値(AEROS)ではなく、CAMSモデル値を正として扱ってください。
`;

// ==========================================
// 1. Main Logic & UI Handlers
// ==========================================

function handleCheckboxOperation(e) {
    if (!e) return;
    const range = e.range;
    const sheet = range.getSheet();
    const ui = AIR_CONFIG.UI;
    if (sheet.getName() === AIR_CONFIG.SHEETS.DASH) {
        if (range.getA1Notation() === ui.CHECK_UPDATE && e.value === 'TRUE') {
            range.setValue(false);
            console.log(`[UI] Manual Update Triggered`);
            main_AirQualityUpdate("MANUAL");
        }
        if (range.getA1Notation() === ui.CHECK_AI && e.value === 'TRUE') {
            range.setValue(false);
            console.log(`[UI] AI Analysis Triggered`);
            run_AI_Analysis();
        }
    }
}

function main_AirQualityUpdate(mode) {
    console.log(`[MAIN] Starting Update... Mode: ${mode}`);

    let ss; try { ss = SpreadsheetApp.openById(AIR_CONFIG.SHEET_ID); } catch (e) { console.error("Sheet Error", e); return; }
    const sheetDash = ss.getSheetByName(AIR_CONFIG.SHEETS.DASH) || initDashboard(ss);
    const ui = AIR_CONFIG.UI;

    // ★Update: 1時間ごとの実行を許可 (奇数時のスキップを削除)
    // if (mode !== "MANUAL") {
    //     const currentHour = new Date().getHours();
    //     if (currentHour !== 0 && currentHour !== 6 && currentHour % 2 !== 0) ...
    // }

    if (mode === "MANUAL") sheetDash.getRange(ui.STATUS_MAIN).setValue("🔄 Updating (CAMS Priority)...");

    // Fetch OpenMeteo First
    console.log(`[MAIN] Fetching CAMS Data...`);
    const camsResult = fetchAllCamsData_Parallel();
    if (!camsResult) {
        console.error(`[MAIN] CAMS Fetch Failed`);
        sheetDash.getRange(ui.STATUS_MAIN).setValue("❌ CAMS Fetch Failed");
        return;
    }

    // AEROS Disabled (Commented out)
    let aerosResult = null;
    // try { 
    //     console.log(`[MAIN] Fetching AEROS Data (Optional)...`); 
    //     aerosResult = fetchAllAerosData(1000); 
    // } catch(e) { console.warn("AEROS fetch skip", e); }

    sheetDash.getRange(ui.STATUS_MAIN).setValue("📝 Processing...");

    const sheetInteg = ss.getSheetByName(AIR_CONFIG.SHEETS.INTEGRATED) || initIntegratedSheet(ss);
    const sheetAeros = ss.getSheetByName(AIR_CONFIG.SHEETS.AEROS) || initAerosSheet(ss);
    const sheetCams = ss.getSheetByName(AIR_CONFIG.SHEETS.CAMS) || initCamsSheet(ss);

    const now = new Date();
    // ★TimeZone Fixed: Asia/Tokyo
    const timeStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:00");
    console.log(`[MAIN] Target Time: ${timeStr}`);

    const tMap = camsResult.targetsMap;
    const sMap = camsResult.stationsMap;

    let snapshotStations = [];
    // AEROS integration skipped

    // Process ALL targets
    AIR_CONFIG.TARGETS.forEach(target => {
        const env = calculateEnvironment_CamsMain(target, snapshotStations, tMap[target.id] || {});
        writeLogRow_v14(sheetInteg, sheetAeros, sheetCams, timeStr, env, tMap, sMap, snapshotStations, target.name);
        if (target.id === 'HOME') {
            updateDashboard_v14(sheetDash, timeStr, env, tMap['HOME'] || {});
        }
    });

    console.log(`[MAIN] Running Archiver...`);
    run_Monthly_Archiver();

    const homeEnv = calculateEnvironment_CamsMain(AIR_CONFIG.TARGETS[0], snapshotStations, tMap['HOME'] || {});
    const finalRisk = assessRisk_EU2024_Strict(homeEnv);
    sheetDash.getRange(ui.STATUS_MAIN).setValue(`✅ Updated\n[${finalRisk.signal}] ${finalRisk.reason}`);

    if (mode === "MANUAL") run_AI_Analysis();

    try {
        if (typeof makeDailySummary_v14 === 'function') {
            console.log(`[LINK] Triggering Daily Summary Module...`);
            makeDailySummary_v14();
        }
    } catch (e) {
        console.warn(`[LINK] Daily Summary Module skipped.`, e);
    }

    console.log(`[MAIN] Completed Successfully.`);
}

// ------------------------------------------------
// Fetch Functions
// ------------------------------------------------

function fetchAllCamsData_Parallel() {
    let requests = [];
    let mapKeys = [];

    const buildReq = (lat, lon, id, type) => {
        const urlAir = `${AIR_CONFIG.URLS.CAMS_AIR}?latitude=${lat}&longitude=${lon}&hourly=${AIR_CONFIG.CAMS_PARAMS.AIR}&timezone=Asia%2FTokyo&forecast_days=1`;
        const urlMet = `${AIR_CONFIG.URLS.CAMS_METEO}?latitude=${lat}&longitude=${lon}&current=${AIR_CONFIG.CAMS_PARAMS.METEO}&timezone=Asia%2FTokyo`;
        requests.push({ url: urlAir, muteHttpExceptions: true });
        requests.push({ url: urlMet, muteHttpExceptions: true });
        mapKeys.push({ type: type, id: id });
    };

    AIR_CONFIG.TARGETS.forEach(t => buildReq(t.lat, t.lon, t.id, 'TARGET'));
    // Stations fetch removed to save resources if not used
    // AIR_CONFIG.STATIONS_DB.forEach(s => buildReq(s.lat, s.lon, s.id, 'STATION'));

    let responses;
    try { responses = UrlFetchApp.fetchAll(requests); } catch (e) { console.error("CAMS Fetch Error", e); return null; }

    let targetsMap = {};
    let stationsMap = {};

    for (let i = 0; i < mapKeys.length; i++) {
        const key = mapKeys[i];
        const resAir = responses[i * 2];
        const resMet = responses[i * 2 + 1];

        let mergedData = {};
        if (resAir.getResponseCode() === 200) {
            const j = JSON.parse(resAir.getContentText());
            if (j.hourly) {
                const nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd'T'HH:00");
                let idx = j.hourly.time.findIndex(t => t.startsWith(nowStr));
                if (idx === -1) idx = 0;
                Object.keys(j.hourly).forEach(k => { if (Array.isArray(j.hourly[k])) mergedData[k] = j.hourly[k][idx]; });
            }
        }
        if (resMet.getResponseCode() === 200) {
            const j = JSON.parse(resMet.getContentText());
            if (j.current) Object.assign(mergedData, j.current);
        }

        if (key.type === 'TARGET') targetsMap[key.id] = mergedData;
        else stationsMap[key.id] = mergedData;
    }
    return { targetsMap, stationsMap };
}

// AEROS functions removed or kept minimal if needed elsewhere
function fetchAllAerosData(sleepMs) { return null; }

// ------------------------------------------------
// Logic Functions
// ------------------------------------------------

function calculateEnvironment_CamsMain(target, aerosSources, targetCams) {
    const pm25 = targetCams.pm2_5;
    const no2 = targetCams.nitrogen_dioxide;
    const so2 = targetCams.sulphur_dioxide;
    const ox = targetCams.ozone;
    const co = targetCams.carbon_monoxide;
    const dust = targetCams.dust;
    const nh3 = targetCams.ammonia;
    const aod = targetCams.aerosol_optical_depth;

    const t = targetCams.temperature_2m;
    const ws = targetCams.wind_gusts_10m ? (targetCams.wind_gusts_10m / 3.6).toFixed(1) : "0.0";
    const press = targetCams.pressure_msl;

    return {
        PM2_5: valid(pm25),
        NO2: valid(no2),
        SO2: valid(so2),
        OX: valid(ox),
        CO: valid(co),
        DUST: valid(dust),
        NH3: valid(nh3),
        AOD: valid(aod),
        TEMP: t,
        WS: ws,
        PRESS: press
    };
}

function valid(val) {
    return (val !== undefined && val !== null) ? parseFloat(val).toFixed(2) : "0.00";
}

function assessRisk_EU2024_Strict(env) {
    let p = parseFloat(env.PM2_5);
    let n = parseFloat(env.NO2);
    let s = parseFloat(env.SO2);

    let reasons = [];
    let signal = "GREEN";

    if (p >= 25.0) { signal = "RED"; reasons.push(`PM2.5(${p})`); }
    if (n >= 50.0) { signal = "RED"; reasons.push(`NO2(${n})`); }
    if (s >= 50.0) { signal = "RED"; reasons.push(`SO2(${s})`); }

    if (signal !== "RED") {
        if (p >= 10.0) { signal = "YELLOW"; reasons.push(`PM2.5(${p})`); }
        else if (n >= 20.0) { signal = "YELLOW"; reasons.push(`NO2(${n})`); }
    }

    let reasonText = reasons.length > 0 ? "Alert: " + reasons.join(", ") : "Safe levels";

    return { signal: signal, reason: reasonText };
}


// ------------------------------------------------
// Helpers & Writing
// ------------------------------------------------

function writeLogRow_v14(sheetInteg, sheetAeros, sheetCams, timeStr, envHome, tMap, sMap, stationsData, locName) {
    const NA = AIR_CONFIG.TEXT_PLACEHOLDER;

    if (locName === 'Home') {
        let rowCams = [timeStr];
        const pushCams = (d) => {
            // ★Added: boundary_layer_height at the end
            const keys = ['pm2_5', 'pm10', 'uv_index', 'dust', 'aerosol_optical_depth', 'ozone', 'nitrogen_dioxide', 'sulphur_dioxide', 'carbon_monoxide', 'ammonia', 'precipitation', 'weather_code', 'wind_gusts_10m', 'cloud_cover', 'surface_pressure', 'pressure_msl', 'temperature_2m', 'relative_humidity_2m', 'freezing_level_height', 'boundary_layer_height'];
            keys.forEach(k => rowCams.push(d[k] !== undefined ? d[k] : NA));
        };
        AIR_CONFIG.TARGETS.forEach(t => pushCams(tMap[t.id] || {}));
        insertAtTop(sheetCams, rowCams);
    }

    const rowInteg = [
        timeStr,
        locName,
        envHome.PM2_5, envHome.OX, envHome.NO2, envHome.SO2,
        envHome.CO, envHome.NH3, envHome.DUST, envHome.AOD,
        envHome.TEMP, envHome.WS, envHome.PRESS,
        assessRisk_EU2024_Strict(envHome).signal
    ];
    insertAtTop(sheetInteg, rowInteg);
}

function updateDashboard_v14(s, t, h, rawCams) {
    const ui = AIR_CONFIG.UI;
    const risk = assessRisk_EU2024_Strict(h);

    s.getRange(ui.STATUS_MAIN).setValue(
        `【GemSyS v14.8】\n${t}\n` +
        `RISK: [${risk.signal}]\n` +
        `Reason: ${risk.reason}\n` +
        `PM2.5: ${h.PM2_5} / Dust: ${h.DUST}\n` +
        `NO2: ${h.NO2} / NH3: ${h.NH3}\n` +
        `SO2: ${h.SO2} / CO: ${h.CO}`
    );

    const bg = risk.signal === "RED" ? "#ffcccc" : (risk.signal === "YELLOW" ? "#fff4cc" : "#ccffcc");
    s.getRange(ui.STATUS_MAIN).setBackground(bg);
}

function insertAtTop(sheet, row) { sheet.insertRowBefore(2); sheet.getRange(2, 1, 1, row.length).setValues([row]); }

function initIntegratedSheet(ss) {
    const s = ss.insertSheet(AIR_CONFIG.SHEETS.INTEGRATED);
    s.appendRow(['Time', 'Location', 'Main_PM25', 'Main_OX', 'Main_NO2', 'Main_SO2', 'Main_CO', 'Main_NH3', 'Main_Dust', 'Main_AOD', 'Temp', 'WS', 'Press', 'Risk_Signal']);
    s.setFrozenRows(1); return s;
}
function initCamsSheet(ss) {
    const s = ss.insertSheet(AIR_CONFIG.SHEETS.CAMS);
    let h = ['Time'];
    // ★Updated: Added BLH header
    const buildHeaders = (prefix) => ['PM2.5', 'PM10', 'UV', 'Dust', 'AOD', 'O3', 'NO2', 'SO2', 'CO', 'NH3', 'Precip', 'Weather', 'Gust', 'Cloud', 'Press_S', 'Press_M', 'Temp', 'Hum', 'FreezeAlt', 'BLH'].map(col => `${prefix}_${col}`);
    AIR_CONFIG.TARGETS.forEach(t => h.push(...buildHeaders(t.name)));
    s.appendRow(h); s.setFrozenRows(1); return s;
}
function initAerosSheet(ss) { const s = ss.insertSheet(AIR_CONFIG.SHEETS.AEROS); s.appendRow(['Time', 'Log_Only']); return s; }
function initDashboard(ss) { const s = ss.insertSheet(AIR_CONFIG.SHEETS.DASH); s.getRange("B5").setValue("Initializing..."); return s; }
function initAISummarySheet(ss) { const s = ss.insertSheet(AIR_CONFIG.SHEETS.AI_SUMMARY); s.appendRow(['Time', 'AI_Response_Log']); s.setFrozenRows(1); return s; }

function run_Monthly_Archiver() { archiveSheet(AIR_CONFIG.SHEETS.CAMS, `GemSyS_Dump_CAMS`); }

function archiveSheet(sheetName, filePrefix) {
    console.log(`[ARCHIVE] Start processing: ${sheetName}`);
    let ss; try { ss = SpreadsheetApp.openById(AIR_CONFIG.SHEET_ID); } catch (e) { console.error(`[ARCHIVE] Failed to open spreadsheet`, e); return; }
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) { console.log(`[ARCHIVE] Sheet not found or empty: ${sheetName}`); return; }

    const retentionMs = AIR_CONFIG.RETENTION.DAYS * 24 * 60 * 60 * 1000;
    const now = new Date().getTime();

    const range = sheet.getDataRange();
    const values = range.getValues();
    const headers = values[0];
    const data = values.slice(1);

    let rowsToArchive = [];
    let rowsToKeep = [headers];

    data.forEach((row, index) => {
        const cellValue = row[0];
        if (cellValue) {
            let timeVal = 0;
            if (cellValue instanceof Date) {
                timeVal = cellValue.getTime();
            } else {
                const strVal = String(cellValue);
                timeVal = new Date(strVal.replace(/-/g, '/')).getTime();
            }
            if (now - timeVal > retentionMs) rowsToArchive.push(row);
            else rowsToKeep.push(row);
        }
    });

    if (rowsToArchive.length > 0) {
        let csv = headers.join(",") + "\n";
        rowsToArchive.forEach(r => {
            csv += r.map(c => { const s = String(c).replace(/"/g, '""'); return `"${s}"`; }).join(",") + "\n";
        });

        let folder; try { folder = DriveApp.getFolderById(AIR_CONFIG.DRIVE_FOLDER_ID); } catch (e) { console.error(`[ARCHIVE] Folder access error`, e); return; }
        const fileName = `${filePrefix}_${Utilities.formatDate(new Date(), "JST", "yyyy_'W'w_dd")}.csv`;
        folder.createFile(fileName, csv, MimeType.CSV);
        sheet.clearContents();
        if (rowsToKeep.length > 0) sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
    }
}

function run_AI_Analysis() {
    // Same as before...
    console.log(`[AI] Analysis Started`);
    const ss = SpreadsheetApp.openById(AIR_CONFIG.SHEET_ID);
    const integSheet = ss.getSheetByName(AIR_CONFIG.SHEETS.INTEGRATED);
    const dashSheet = ss.getSheetByName(AIR_CONFIG.SHEETS.DASH);
    const summarySheet = ss.getSheetByName(AIR_CONFIG.SHEETS.AI_SUMMARY) || initAISummarySheet(ss);
    const ui = AIR_CONFIG.UI;

    if (!integSheet) { console.error(`[AI] Risk_Tracker sheet not found.`); return; }

    const lastRow = integSheet.getLastRow(); if (lastRow < 2) return;
    const headers = integSheet.getRange(1, 1, 1, integSheet.getLastColumn()).getValues()[0];

    const checkRows = 6;
    const dataRange = integSheet.getRange(2, 1, checkRows, integSheet.getLastColumn()).getValues();

    const locIdx = headers.indexOf('Location');
    let rowValues = dataRange.find(r => r[locIdx] === 'Home');

    if (!rowValues) {
        console.warn("[AI] 'Home' row not found. Using top row.");
        rowValues = dataRange[0];
    }

    const idxPM25 = headers.indexOf('Main_PM25');
    const idxNO2 = headers.indexOf('Main_NO2');
    const idxSO2 = headers.indexOf('Main_SO2');

    const risk = assessRisk_EU2024_Strict({
        PM2_5: rowValues[idxPM25],
        NO2: rowValues[idxNO2],
        SO2: rowValues[idxSO2]
    });

    const csvInput = `${headers.join(",")},Risk_Level,Risk_Reason\n${rowValues.join(",")},${risk.signal},${risk.reason}`;

    dashSheet.getRange(ui.STATUS_AI).setValue("🤖 AI Processing...");

    const safePrompt = AI_SYSTEM_PROMPT_TEMPLATE.replace('{{PLACEHOLDER}}', AIR_CONFIG.TEXT_PLACEHOLDER);
    const model = AIR_CONFIG.AI_MODEL || 'gemini-2.5-flash';

    let response = null;
    try {
        if (typeof GEMINI_GenerateContent === 'function') {
            console.log(`[AI] Calling Module: ${model}`);
            response = GEMINI_GenerateContent(csvInput, safePrompt, model);
        } else {
            console.error("[AI] Module 'GEMINI_GenerateContent' not found.");
        }
    } catch (e) {
        console.error("[AI] Module Execution Error", e);
    }

    if (response) {
        dashSheet.getRange(ui.OUTPUT_AI).setValue(response);
        dashSheet.getRange(ui.STATUS_AI).setValue("✅ Done");
        const timestamp = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
        summarySheet.insertRowBefore(2);
        summarySheet.getRange(2, 1, 1, 2).setValues([[timestamp, response]]);
        console.log(`[AI] Response saved`);
    } else {
        dashSheet.getRange(ui.STATUS_AI).setValue("❌ API Error");
        console.error(`[AI] API Error or No Response`);
    }
}

function Run_OneShot() {
    main_AirQualityUpdate("MANUAL");
}