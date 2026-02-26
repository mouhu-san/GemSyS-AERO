/**
 * GemSyS AERO - 統合システム仕様書 (Strict / Vulnerable Mode)
 * user向けの厳格化された行動指針を展開します。
 */
function generateSystemSpecSheet() {
    const ss = SpreadsheetApp.openById(AIR_CONFIG.SHEET_ID);
    const sheetName = "DOC_SystemSpec";
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
    } else {
        sheet.clear();
    }

    let currentRow = 1;

    // ==========================================
    // タイトル
    // ==========================================
    sheet.getRange(currentRow, 1).setValue("GemSyS AERO v15.0 仕様定義書 (Strict Mode)")
        .setFontSize(16).setFontWeight("bold").setFontColor("#1155cc");
    currentRow += 2;

    // ==========================================
    // セクション1: EEA AQI 濃度区分 (脆弱層向けに最適化された行動指針)
    // ==========================================
    sheet.getRange(currentRow, 1).setValue("■ 1. 濃度区分および専用アクションプラン（脆弱性考慮）").setFontSize(12).setFontWeight("bold");
    currentRow++;

    const aqiHeaders = ["指数区分 (Index level)", "PM2.5", "PM10", "NO2", "O3", "SO2", "最適化された行動指針・防衛策"];
    const aqiData = [
        ["Good (良好)", "0 - 5", "0 - 15", "0 - 10", "0 - 60", "0 - 20", "【安全圏】\n大気は非常にクリーンです。積極的な換気を行い、室内の空気を入れ替えるのに最適なタイミングです。"],
        ["Fair (普通)", "6 - 15", "16 - 45", "11 - 25", "61 - 100", "21 - 40", "【初期警戒・換気注意】\n一般的な基準では安全ですが、気道が敏感な状態では微小な刺激になり得ます。長時間の激しい運動は避け、PM2.5が10を超え始めたら窓による換気は最小限に留めてください。"],
        ["Moderate (中程度)", "16 - 50", "46 - 120", "26 - 60", "101 - 120", "41 - 125", "【実質的リスク帯：活動制限】\n一般には中程度ですが、明確なリスク帯（Poor相当）として扱います。屋外での活動はできる限り控え、外出時は必ず高性能マスクを着用してください。常備薬の確認を推奨します。"],
        ["Poor (悪い)", "51 - 90", "121 - 195", "61 - 100", "121 - 160", "126 - 190", "【厳重警戒：屋内退避】\n直ちに屋外での活動を中止してください。窓を完全に閉め切り、室内の空気清浄機（HEPAフィルター）の出力を最大に引き上げてください。"],
        ["Very poor (非常に悪い)", "91 - 140", "196 - 270", "101 - 150", "161 - 180", "191 - 275", "【危険帯：絶対的制限】\n極めて危険な状態です。外出は完全に避け、外気が侵入しやすい換気扇の使用等も控えてください。"],
        ["Extremely poor (極悪)", "> 140", "> 270", "> 150", "> 180", "> 275", "【緊急事態】\n同上。室内の最も密閉性の高い空間で待機し、システムのアラート解除を待ってください。"]
    ];

    sheet.getRange(currentRow, 1, 1, aqiHeaders.length).setValues([aqiHeaders]).setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
    sheet.getRange(currentRow + 1, 1, aqiData.length, aqiData[0].length).setValues(aqiData);

    // 背景色の適用（Moderateを通常より強いオレンジ寄りにするなど、視覚的にも警戒度を上げる）
    const aqiColors = ["#d9ead3", "#fff2cc", "#f9cb9c", "#e06666", "#c27ba0", "#a64d79"];
    for (let i = 0; i < aqiColors.length; i++) {
        sheet.getRange(currentRow + 1 + i, 1, 1, aqiHeaders.length).setBackground(aqiColors[i]);
    }

    sheet.getRange(currentRow + 1, 7, aqiData.length, 1).setWrap(true);
    currentRow += aqiData.length + 3;

    // ==========================================
    // セクション2: Directive (EU) 2024/2881 絶対基準
    // ==========================================
    sheet.getRange(currentRow, 1).setValue("■ 2. 蓄積リスク判定限界値 (EU 2030 Strict)").setFontSize(12).setFontWeight("bold");
    currentRow++;

    const euHeaders = ["対象物質", "絶対限界値 (EU 2030)", "システム判定ロジック"];
    const euData = [
        ["PM2.5", "25 μg/m³ (24時間平均)", "24時間の移動平均がこれを超えた場合、日々のベースラインが底上げされており、慢性的な炎症リスクが極めて高いと判定。"],
        ["PM10", "45 μg/m³ (24時間平均)", "同上。粗大粒子による上気道への物理的刺激リスクとして評価。"],
        ["NO2", "200 μg/m³ (1時間値)", "光化学反応の前駆物質。一時的でもこれを超えた場合、呼吸器粘膜への直接的な急激なダメージを警戒。"],
        ["SO2", "350 μg/m³ (1時間値)", "同上。"],
        ["O3", "120 μg/m³ (8時間基準)", "酸化ストレスによる気道収縮リスク。高気温と強いUV時に重点監視。"]
    ];

    sheet.getRange(currentRow, 1, 1, euHeaders.length).setValues([euHeaders]).setBackground("#6aa84f").setFontColor("white").setFontWeight("bold");
    sheet.getRange(currentRow + 1, 1, euData.length, euData[0].length).setValues(euData);
    sheet.getRange(currentRow + 1, 3, euData.length, 1).setWrap(true);
    currentRow += euData.length + 3;

    // ==========================================
    // セクション3: 気象工学における推論・計算ロジック
    // ==========================================
    sheet.getRange(currentRow, 1).setValue("■ 3. 物理的リスク推論アルゴリズム").setFontSize(12).setFontWeight("bold");
    currentRow++;

    const logicHeaders = ["現象 / リスク名", "JS計算条件 (Logic Engine)", "気象工学的メカニズムと健康への影響"];
    const logicData = [
        ["滞留・蓄積 (Stagnation)", "BLH < 500m ＆ Gust < 10km/h", "接地逆転層の「蓋」効果と弱風により汚染が逃げ場を失う。同じ場所にいるだけで徐々に吸引量が増えるため、長時間の換気停止が必須となる。"],
        ["微小粒子残留 (Scavenging Gap)", "Precip > 0.5mm ＆ PM2.5 >= 16", "「雨が降っているから空気が綺麗になった」という錯覚を防ぐ。粗大粒子は雨で落ちるが、PM2.5は雨滴をすり抜けて浮遊し続けるため、雨天時でも警戒を緩めない。"],
        ["光化学O3生成 (Photochemical)", "UV >= 5 ＆ Temp >= 25℃ ＆ NO2 >= 20", "強い紫外線と高温下でNO2が反応し、数時間以内のO3ピーク到達を予測。気道収縮の引き金になるため、晴れた日の午後にとくに警戒。"],
        ["SIA転換 (無機エアロゾル生成)", "Hum >= 75% ＆ Dust >= 20", "高湿度とDust（鉱物）を触媒とし、ガスが粒子(PM2.5)へ急激に相転移する現象。急な濃度の跳ね上がりに直結する。"],
        ["越境輸送 (Transboundary)", "地上PM2.5 <= 15 ＆ AOD >= 0.5", "地上の濃度が低くても空の透明度(AOD)が低い場合、上空に汚染塊がある。後流や下降気流で突然地上に降りてくるリスクがあるため、事前準備のサインとなる。"]
    ];

    sheet.getRange(currentRow, 1, 1, logicHeaders.length).setValues([logicHeaders]).setBackground("#e69138").setFontColor("white").setFontWeight("bold");
    sheet.getRange(currentRow + 1, 1, logicData.length, logicData[0].length).setValues(logicData);

    // 折り返し設定
    sheet.getRange(currentRow + 1, 2, logicData.length, 2).setWrap(true);

    // ==========================================
    // 全体のレイアウト調整
    // ==========================================
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 150);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 150);
    sheet.setColumnWidth(7, 450);

    sheet.setFrozenRows(3);

    SpreadsheetApp.getUi().alert("✅ DOC_SystemSpec シートを『Strict Mode (脆弱層最適化版)』で展開しました。");
}