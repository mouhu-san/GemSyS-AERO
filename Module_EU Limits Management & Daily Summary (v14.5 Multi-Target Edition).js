/**
 * Module: EU Limits Management & Daily Summary (v14.5 Multi-Target Edition)
 * Target System: GemSyS_AERO v14.5
 * Dependency: AIR_CONFIG (from Main Script)
 * * [Update v14.5]
 * - Version sync with Main System.
 * - Logic: Multi-Target (Home, Univ, Commute) scanning enabled.
 * - Identifies 'Location' column automatically from Risk_Tracker.
 */

// ==========================================
// 1. Setup & Configuration
// ==========================================

function initEuLimitsSheet() {
    const ss = getSpreadsheeted();
    const sheetName = "EU_LIMITS";
    
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) {
        console.log(`[Module] Sheet '${sheetName}' already exists.`);
        return;
    }

    sheet = ss.insertSheet(sheetName);
    const headers = ["Pollutant", "Metric", "EU_Limit", "Unit", "Note"];
    const data = [
        ["PM2.5", "Daily", 25, "µg/m³", "Directive 2024/2881 - daily limit"],
        ["NO2", "Daily", 50, "µg/m³", "Directive 2024/2881 - daily limit"],
        ["SO2", "Daily", 50, "µg/m³", "Directive 2024/2881 - daily limit"],
        ["O3", "8h_max", 120, "µg/m³", "Target value"],
        ["CO", "Daily", 10000, "µg/m³", "10 mg/m³ -> 10000 µg/m³"],
        ["NH3", "Daily", 0, "µg/m³", "Monitor only"], 
        ["Dust", "Daily", 0, "µg/m³", "Monitor only"],
        ["AOD", "Daily", 0, "-", "Monitor only"]
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#EFEFEF");
    console.log(`[Module] Sheet '${sheetName}' created.`);
}

function getEuLimits() {
    const ss = getSpreadsheeted();
    const sheet = ss.getSheetByName("EU_LIMITS");
    
    if (!sheet) return { PM25: { Daily: 25 }, NO2: { Daily: 50 }, SO2: { Daily: 50 }, O3: { "8h_max": 120 }, CO: { Daily: 10000 } };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};

    const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const limits = {};

    values.forEach(row => {
        let pollutant = row[0];
        const metric = row[1];
        const limitVal = row[2];
        if (pollutant === "PM2.5") pollutant = "PM25";
        if (!limits[pollutant]) limits[pollutant] = {};
        limits[pollutant][metric] = Number(limitVal);
    });
    return limits;
}

// ==========================================
// 2. Main Summary Logic (Multi-Target Support)
// ==========================================

/**
 * ③ 最新の環境データを取得してサマリーを作成 (全地点対応)
 */
function makeDailySummary_v14() {
    const ss = getSpreadsheeted();
    
    // シート取得
    const logSheetName = (typeof AIR_CONFIG !== 'undefined') ? AIR_CONFIG.SHEETS.INTEGRATED : 'Risk_Tracker';
    const sLog = ss.getSheetByName(logSheetName);

    if (!sLog) {
        console.error(`[Module] Log sheet '${logSheetName}' not found.`);
        return;
    }

    const lastRow = sLog.getLastRow();
    if (lastRow < 2) return;

    // ヘッダー取得
    const headers = sLog.getRange(1, 1, 1, sLog.getLastColumn()).getValues()[0];
    
    // データ取得範囲: 最新の書き込み（3地点分）をカバーするため、上から10行ほど取得して解析する
    const checkRowCount = 10; 
    const dataRange = sLog.getRange(2, 1, Math.min(checkRowCount, lastRow - 1), sLog.getLastColumn()).getValues();

    // 最新の時刻（一番上の行の時間）を取得し、ターゲット時間を定める
    const latestTimeStr = String(dataRange[0][0]); // Time column is index 0
    
    // ターゲット時間のデータ行だけを抽出（Home, Univ, Commuteすべて）
    const targetRows = dataRange.filter(row => String(row[0]) === latestTimeStr);
    
    console.log(`[Module] Checking ${targetRows.length} locations for time: ${latestTimeStr}`);

    // レポート用シート準備
    const reportSheetName = "AI_Daily_Check";
    let sReport = ss.getSheetByName(reportSheetName);
    if (!sReport) {
        sReport = ss.insertSheet(reportSheetName);
        sReport.getRange("A1:C1").setValues([["CheckTime", "Status", "Details"]]);
        sReport.setFrozenRows(1);
    }

    // 基準値取得
    const limits = getEuLimits();
    const getLim = (k, m, d) => (limits[k] && limits[k][m]) ? limits[k][m] : d;

    // 各行（各地点）について判定
    targetRows.forEach(row => {
        // 値取得ヘルパー
        const getVal = (colName) => {
            const idx = headers.indexOf(colName);
            return (idx !== -1 && row[idx] !== "") ? Number(row[idx]) : 0;
        };
        
        // 地点名取得 (v14.4で追加された 'Location' 列を探す)
        let locName = "Unknown";
        const locIdx = headers.indexOf('Location');
        if (locIdx !== -1) {
            locName = row[locIdx];
        } else {
            // もしLocation列がない古い形式なら便宜上Homeとする
            locName = "Home(Legacy)"; 
        }

        const vals = {
            PM25: getVal('Main_PM25'),
            NO2:  getVal('Main_NO2'),
            SO2:  getVal('Main_SO2'),
            OX:   getVal('Main_OX'),
            CO:   getVal('Main_CO')
        };

        // 判定
        let warnings = [];
        if (vals.PM25 > getLim("PM25", "Daily", 25)) warnings.push(`PM2.5(${vals.PM25})`);
        if (vals.NO2  > getLim("NO2", "Daily", 50))  warnings.push(`NO2(${vals.NO2})`);
        if (vals.SO2  > getLim("SO2", "Daily", 50))  warnings.push(`SO2(${vals.SO2})`);

        // ログ文字列作成
        const status = warnings.length > 0 ? "⚠️ WARNING" : "🟢 OK";
        let detailText = `[${locName}] `;
        
        if (warnings.length > 0) {
            detailText += warnings.join(", ") + " > EU Limit";
        } else {
            detailText += `All clear. PM2.5:${vals.PM25}`;
        }

        // 重複書き込み防止（同じ時間・同じ地点が既に書き込まれていないか簡易チェックは省略し、最新を上に積む）
        sReport.insertRowBefore(2);
        sReport.getRange(2, 1, 1, 3).setValues([[latestTimeStr, status, detailText]]);
        
        console.log(`[Module] ${locName}: ${status}`);
    });
}

function getSpreadsheeted() {
    const id = (typeof AIR_CONFIG !== 'undefined' && AIR_CONFIG.SHEET_ID) ? AIR_CONFIG.SHEET_ID : SpreadsheetApp.getActiveSpreadsheet().getId();
    return SpreadsheetApp.openById(id);
}