/**
 * GemSyS AERO - 仕様書・ロジック可視化シート生成ツール
 * ※1回だけ実行すればOKです。
 */
function generateDocumentationSheets() {
    const ss = SpreadsheetApp.openById(AIR_CONFIG.SHEET_ID); // GemSyS_AERO.jsの定数を利用

    createAqiStandardSheet(ss);
    createPhysicsLogicSheet(ss);

    SpreadsheetApp.getUi().alert("✅ 仕様書シート（DOC_AQI_基準, DOC_気象工学ロジック）の生成が完了しました！");
}

function createAqiStandardSheet(ss) {
    const sheetName = "DOC_AQI_基準";
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
    } else {
        sheet.clear();
    }

    // 1. EEA AQI 濃度区分テーブル
    const headers = ["Index level (指数区分)", "PM2.5", "PM10", "NO2", "O3", "SO2", "背景色イメージ"];
    const data = [
        ["Good (良好)", "0 - 5", "0 - 15", "0 - 10", "0 - 60", "0 - 20", "🟢 安全圏"],
        ["Fair (普通)", "6 - 15", "16 - 45", "11 - 25", "61 - 100", "21 - 40", "🟡 脆弱層警戒ライン"],
        ["Moderate (中程度)", "16 - 50", "46 - 120", "26 - 60", "101 - 120", "41 - 125", "🟠 AI推論トリガー"],
        ["Poor (悪い)", "51 - 90", "121 - 195", "61 - 100", "121 - 160", "126 - 190", "🔴 基準超過"],
        ["Very poor (非常に悪い)", "91 - 140", "196 - 270", "101 - 150", "161 - 240", "191 - 400", "🟣 危険"],
        ["Extremely poor (極めて悪い)", "> 140", "> 270", "> 150", "> 240", "> 400", "🟤 極めて危険"]
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

    // 背景色の適用 (視覚化)
    const colors = ["#d9ead3", "#fff2cc", "#fce5cd", "#f4cccc", "#d9d2e9", "#ead1dc"];
    for (let i = 0; i < colors.length; i++) {
        sheet.getRange(i + 2, 1, 1, headers.length).setBackground(colors[i]);
    }

    // 2. EU 2030年限界値 (Directive EU 2024/2881 Strict)
    sheet.getRange(10, 1).setValue("【Directive (EU) 2024/2881 絶対基準 (2030年限界値)】").setFontWeight("bold");
    const euData = [
        ["物質", "基準値", "評価期間", "備考"],
        ["PM2.5", "25 μg/m3", "24時間", "超えた場合、即時EU Limit Violationフラグが立つ"],
        ["PM10", "45 μg/m3", "24時間", ""],
        ["NO2", "200 μg/m3", "1時間", ""],
        ["SO2", "350 μg/m3", "1時間", ""],
        ["O3", "120 μg/m3", "8時間", "情報提供閾値"]
    ];
    sheet.getRange(11, 1, euData.length, euData[0].length).setValues(euData);
    sheet.getRange(11, 1, 1, euData[0].length).setBackground("#6aa84f").setFontColor("white").setFontWeight("bold");

    sheet.autoResizeColumns(1, 7);
    sheet.setFrozenRows(1);
}

function createPhysicsLogicSheet(ss) {
    const sheetName = "DOC_気象工学ロジック";
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
    } else {
        sheet.clear();
    }

    const headers = ["リスク判定項目", "判定ロジック (JS計算条件)", "気象工学的推論・解説", "関連パラメータ"];
    const data = [
        ["Stagnation (滞留・蓄積)", "BLH < 500m ＆ Gust < 10km/h", "風が弱く境界層高度が低いため、移流による排出が停止し、局所的に汚染物質が蓄積しやすい危険な状態。", "BLH, Gust"],
        ["Wet Deposition (湿性沈着)", "Precip > 0.5mm", "降水による洗浄効果(Wash-out)が働いている状態。多くの水溶性ガスや粗大粒子が除去される。", "Precip"],
        ["Scavenging Gap (微小粒子残留)", "Precip > 0.5mm ＆ PM2.5 >= 16", "雨が降っているにも関わらずPM2.5が高い状態。PM2.5は雨滴との衝突断面積が小さく、雨で落ちにくい特性が表れている。", "Precip, PM2.5"],
        ["Photochemical (光化学O3生成)", "UV >= 5 ＆ Temp >= 25℃ ＆ NO2 >= 20", "強い紫外線と高温により、NO2などを前駆物質として光化学反応が進行し、二次的にオゾンが生成されやすい状態。", "UV, Temp, NO2"],
        ["SIA Conversion (無機エアロゾル生成)", "Hum >= 75% ＆ Dust >= 20", "高湿度下で粒子表面に水膜ができ、Dustが触媒となることで、ガス(SO2/NO2)から粒子(PM2.5)への転換が加速する状態。", "Hum, Dust"],
        ["Transboundary (越境輸送)", "PM2.5 <= 15 ＆ AOD >= 0.5", "地上のPM2.5は低いが、気柱全体のエアロゾル量(AOD)が多い状態。上空の高い位置を汚染塊が通過中と推測される。", "PM2.5, AOD"]
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground("#e69138").setFontColor("white").setFontWeight("bold");
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

    // セルの折り返し設定と幅調整
    sheet.getRange(2, 2, data.length, 2).setWrap(true);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 400);
    sheet.setColumnWidth(4, 150);
    sheet.setFrozenRows(1);
}