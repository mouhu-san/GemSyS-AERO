/**
 * Module: Range CAMS Fetcher (v1.0)
 * 指定された期間のCAMSデータを取得し、DB_Hourly_CAMSシートに保存します。
 */

function run_Range_CAMS_Fetch(startDate, endDate) {
    console.log(`[RANGE-FETCH] Starting fetch from ${startDate} to ${endDate}...`);
    
    const ss = SpreadsheetApp.openById(AIR_CONFIG.SHEET_ID);
    const sheetName = AIR_CONFIG.SHEETS.CAMS;
    let sheet = ss.getSheetByName(sheetName);

    // シートが無ければ作成
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        // ヘッダー作成
        let headers = ['Time'];
        const headerLabels = ['PM2.5', 'PM10', 'UV', 'Dust', 'AOD', 'O3', 'NO2', 'SO2', 'CO', 'NH3', 'Precip', 'Weather', 'Gust', 'Cloud', 'Press_S', 'Press_M', 'Temp', 'Hum', 'FreezeAlt', 'BLH'];
        AIR_CONFIG.TARGETS.forEach(t => {
            headers.push(...headerLabels.map(col => `${t.name}_${col}`));
        });
        sheet.appendRow(headers);
        sheet.setFrozenRows(1);
    }

    // 既存データの最終行を取得（追記用）
    // もし「期間指定の場合は洗い替え（全消去）」が良い場合は、ここを sheet.clear() に変更してください。
    // 今回は安全のため「追記」としていますが、重複チェックは簡易的です。

    // APIリクエストの構築
    let requests = [];
    AIR_CONFIG.TARGETS.forEach(target => {
        // Air Quality API
        const urlAir = `${AIR_CONFIG.URLS.CAMS_AIR}?latitude=${target.lat}&longitude=${target.lon}&hourly=${AIR_CONFIG.CAMS_PARAMS.AIR}&timezone=Asia%2FTokyo&start_date=${startDate}&end_date=${endDate}`;
        // Meteo API
        const urlMet = `${AIR_CONFIG.URLS.CAMS_METEO}?latitude=${target.lat}&longitude=${target.lon}&hourly=${AIR_CONFIG.CAMS_PARAMS.METEO}&timezone=Asia%2FTokyo&start_date=${startDate}&end_date=${endDate}`;

        requests.push({ url: urlAir, muteHttpExceptions: true });
        requests.push({ url: urlMet, muteHttpExceptions: true });
    });

    console.log(`[RANGE-FETCH] Sending ${requests.length} requests...`);
    const responses = UrlFetchApp.fetchAll(requests);

    // データ処理
    let timeMap = {};
    const keys = ['pm2_5', 'pm10', 'uv_index', 'dust', 'aerosol_optical_depth', 'ozone', 'nitrogen_dioxide', 'sulphur_dioxide', 'carbon_monoxide', 'ammonia', 'precipitation', 'weather_code', 'wind_gusts_10m', 'cloud_cover', 'surface_pressure', 'pressure_msl', 'temperature_2m', 'relative_humidity_2m', 'freezing_level_height', 'boundary_layer_height'];

    AIR_CONFIG.TARGETS.forEach((target, index) => {
        const resAir = responses[index * 2];
        const resMet = responses[index * 2 + 1];

        if (resAir.getResponseCode() === 200 && resMet.getResponseCode() === 200) {
            const dataAir = JSON.parse(resAir.getContentText()).hourly;
            const dataMet = JSON.parse(resMet.getContentText()).hourly;

            dataAir.time.forEach((t, i) => {
                // ISO時刻を整形: 2025-12-01T00:00 -> 2025/12/01 00:00
                const timeStr = t.replace('T', ' ').replace(/-/g, '/');
                if (!timeMap[timeStr]) timeMap[timeStr] = {};
                if (!timeMap[timeStr][target.id]) timeMap[timeStr][target.id] = {};

                keys.forEach(k => {
                    if (dataAir[k] !== undefined) timeMap[timeStr][target.id][k] = dataAir[k][i];
                    if (dataMet[k] !== undefined) timeMap[timeStr][target.id][k] = dataMet[k][i];
                });
            });
        } else {
            console.error(`[RANGE-FETCH] Failed for ${target.name}. Code: ${resAir.getResponseCode()}/${resMet.getResponseCode()}`);
            throw new Error(`API Error for ${target.name}`);
        }
    });

    // 書き込み用配列の作成
    let rows = [];
    const sortedTimes = Object.keys(timeMap).sort(); // 古い順に並べる（追記する場合自然）

    sortedTimes.forEach(t => {
        let row = [t];
        AIR_CONFIG.TARGETS.forEach(target => {
            const d = timeMap[t][target.id] || {};
            keys.forEach(k => {
                row.push(d[k] !== undefined ? d[k] : "null");
            });
        });
        rows.push(row);
    });

    // シートへの書き込み
    if (rows.length > 0) {
        // 一括書き込み（分割処理は行数が多い場合のみ必要だが、数ヶ月分程度なら一度でいけるケースが多い。必要ならChunk処理を追加）
        const startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
        console.log(`[RANGE-FETCH] Completed. ${rows.length} rows appended.`);
    } else {
        console.warn(`[RANGE-FETCH] No data retrieved.`);
    }
}