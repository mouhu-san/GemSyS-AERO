/**
 * Module: GemSyS Daily Reporter (v1.0)
 * 役割: 1日のLogic Engineの計算結果を集計し、気象工学的統計（24h平均・最大値・滞留時間）を算出・報告する。
 * 準拠: Directive (EU) 2024/2881 (24h Limit: 25μg/m3)
 */

const DailyReporter = {

    /**
     * 1日の統計レポートを生成
     */
    generateEndOfDayReport: function () {
        const ss = SpreadsheetApp.openById(AIR_CONFIG.SHEET_ID);
        const logSheet = ss.getSheetByName(AIR_CONFIG.SHEETS.INTEGRATED);
        const data = logSheet.getDataRange().getValues();

        // 本日の日付データのみ抽出（Locationが'Home'のものに限定）
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const todayRows = data.filter(row => {
            const rowDate = new Date(row[0]);
            return rowDate >= today && row[1] === 'Home';
        });

        if (todayRows.length === 0) return "分析データ不足：本日のログが記録されていません。";

        // --- 統計尺度の算出 ---

        // PM2.5 (Index 2), WS(風速 Index 11)
        const pm25Values = todayRows.map(r => parseFloat(r[2])).filter(v => !isNaN(v));
        const wsValues = todayRows.map(r => parseFloat(r[11])).filter(v => !isNaN(v));

        // 1. 24時間移動平均 (法的基準: 25μg/m3)
        const avgPm25 = pm25Values.reduce((a, b) => a + b, 0) / pm25Values.length;

        // 2. 日間最大値 (スパイク曝露の確認)
        const maxPm25 = Math.max(...pm25Values);

        // 3. 静穏時間（Stagnation Hours）の累積
        // 風速 2.0m/s (約 7.2km/h) 未満を「静穏」と定義
        const calmHours = wsValues.filter(v => v < 2.0).length;

        // 4. 超過判定
        const euViolation = avgPm25 > 25;

        // --- サマリー文の構築 ---

        const summaryText = `
### 📊 GemSyS Daily Environmental Report
【集計日】: ${Utilities.formatDate(today, "JST", "yyyy/MM/dd")}
【地点】: 愛西市 (Home)

#### 1. 環境統計尺度 (Air Quality Metrics)
- **24h平均 PM2.5**: ${avgPm25.toFixed(2)} μg/m³ （判定: ${euViolation ? "⚠️ 超過" : "✅ 適合"}）
- **日間最大 PM2.5**: ${maxPm25.toFixed(2)} μg/m³
- **大気静穏時間**: ${calmHours} 時間 / ${todayRows.length}h (蓄積リスク)

#### 2. 気象工学的インサイト
${this.generatePhysicsInsights(avgPm25, maxPm25, calmHours)}

#### 3. 戦術的アドバイス
${euViolation ? "⚠️ 24時間平均がEU法的限界値を超過しています。空気清浄機の稼働維持とフィルターの目詰まり確認を推奨します。" : "✅ 1日を通して法的基準内に収まりました。良好な空気質です。"}
`;

        console.log(summaryText);
        return summaryText;
    },

    /**
     * 統計値に基づく物理推論
     */
    generatePhysicsInsights: function (avg, max, calm) {
        let insights = "";

        // 蓄積型の判定
        if (calm >= 8) {
            insights += ">> 【蓄積型汚染】長時間の大気停滞（Stagnation）が観測されました。局所発生源からの汚染が拡散せず、濃度を底上げした可能性があります。\n";
        }

        // スパイク型の判定
        if (max > avg * 2.5) {
            insights += ">> 【スパイク型曝露】平均値に対して極端に高い最大値が検出されました。越境汚染の短時間通過、あるいは近隣での一時的な燃焼イベントが推測されます。\n";
        }

        // 総合評価
        if (insights === "") {
            insights = ">> 【安定型】大気の混合・移流が正常に行われ、濃度変動は安定推移しました。";
        }

        return insights;
    },

    /**
     * ダッシュボードUI表示用に、本日の統計数値のみを計算して返す
     */
    getTodayMetrics: function () {
        try {
            const ss = SpreadsheetApp.openById(AIR_CONFIG.SHEET_ID);
            const logSheet = ss.getSheetByName(AIR_CONFIG.SHEETS.INTEGRATED);
            if (!logSheet) return null;

            const data = logSheet.getDataRange().getValues();
            const today = new Date();
            today.setHours(0, 0, 0, 0);

            // 本日の Home データのみを抽出
            const todayRows = data.filter(row => {
                const rowDate = new Date(row[0]);
                return rowDate >= today && row[1] === 'Home';
            });

            if (todayRows.length === 0) return null;

            const pm25Values = todayRows.map(r => parseFloat(r[2])).filter(v => !isNaN(v));
            const wsValues = todayRows.map(r => parseFloat(r[11])).filter(v => !isNaN(v));

            return {
                avgPm25: pm25Values.length > 0 ? pm25Values.reduce((a, b) => a + b, 0) / pm25Values.length : 0,
                maxPm25: pm25Values.length > 0 ? Math.max(...pm25Values) : 0,
                calmHours: wsValues.filter(v => v < 2.0).length,
                dataCount: todayRows.length
            };
        } catch (e) {
            console.warn("getTodayMetrics Error:", e);
            return null;
        }
    }
};

/**
 * 既存インターフェース互換用
 */
function makeDailySummary_v14() {
    const report = DailyReporter.generateEndOfDayReport();

    // AI_Summaryシート等へ記録する処理を追加
    try {
        const ss = SpreadsheetApp.openById(AIR_CONFIG.SHEET_ID);
        const sumSheet = ss.getSheetByName(AIR_CONFIG.SHEETS.AI_SUMMARY);
        if (sumSheet) {
            sumSheet.insertRowBefore(2);
            sumSheet.getRange(2, 1, 1, 2).setValues([[new Date(), "[DAILY REPORT]\n" + report]]);
        }
    } catch (e) {
        console.warn("Summary logging failed", e);
    }

    return report;
}