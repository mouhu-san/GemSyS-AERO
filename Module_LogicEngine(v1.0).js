/**
 * Module: GemSyS_Logic Engine (v1.2準拠)
 * 役割: 気象・大気質の生データを入力とし、EEA AQI、EU基準超過、気象工学リスクを決定論的に算出する。
 * 出力: AI（Reasoning Engine）へ渡すための構造化データ (JSON)
 */

const GemSyS_Logic = {
    // 1. EEA AQI 濃度区分定義 (上限値: μg/m3)
    eeaBands: {
        pm25: [5, 15, 50, 90, 140, Infinity],
        pm10: [15, 45, 120, 195, 270, Infinity],
        no2: [10, 25, 60, 100, 150, Infinity],
        o3: [60, 100, 120, 160, 180, Infinity],
        so2: [20, 40, 125, 190, 275, Infinity]
    },

    bandNames: ["Good (良好)", "Fair (普通)", "Moderate (中程度)", "Poor (悪い)", "Very poor (非常に悪い)", "Extremely poor (極めて悪い)"],

    /**
     * 単一物質のAQIレベルを計算
     */
    getPollutantLevel: function (pollutant, value) {
        const bands = this.eeaBands[pollutant];
        for (let i = 0; i < bands.length; i++) {
            if (value <= bands[i]) return i;
        }
        return 5;
    },

    /**
     * EEA AQI と脆弱なグループへのアラート判定
     */
    calculateAQI: function (data) {
        const levels = {
            pm25: this.getPollutantLevel('pm25', data.pm25),
            pm10: this.getPollutantLevel('pm10', data.pm10),
            no2: this.getPollutantLevel('no2', data.no2),
            o3: this.getPollutantLevel('o3', data.o3),
            so2: this.getPollutantLevel('so2', data.so2)
        };

        // 最大値法による総合AQIの決定
        let maxLevel = 0;
        let primaryPollutant = [];

        for (const [key, level] of Object.entries(levels)) {
            if (level > maxLevel) {
                maxLevel = level;
                primaryPollutant = [key];
            } else if (level === maxLevel && maxLevel > 0) {
                primaryPollutant.push(key);
            }
        }

        // トリガーポイント判定: Moderate (Level 2) 以上で脆弱層（喘息患者等）へのアラートをアクティブ化
        const sensitiveGroupAlert = maxLevel >= 2;

        return {
            overall_aqi_level: maxLevel,
            overall_aqi_status: this.bandNames[maxLevel],
            primary_pollutant: primaryPollutant,
            sensitive_group_alert_active: sensitiveGroupAlert,
            details: levels
        };
    },

    /**
     * Directive (EU) 2024/2881 絶対基準の超過判定
     */
    checkEU2030Limits: function (data) {
        const alerts = [];
        if (data.pm25 > 25) alerts.push("PM2.5 24h基準 (25μg/m3) 超過");
        if (data.pm10 > 45) alerts.push("PM10 24h基準 (45μg/m3) 超過");
        if (data.no2 > 200) alerts.push("NO2 1h限界値 (200μg/m3) 超過");
        if (data.so2 > 350) alerts.push("SO2 1h限界値 (350μg/m3) 超過");
        if (data.o3 > 120) alerts.push("O3 8h基準 (120μg/m3) 超過");

        return {
            eu_limit_exceeded: alerts.length > 0,
            violations: alerts
        };
    },

    /**
     * 気象工学的物理推論 (滞留、沈着、二次生成など)
     */
    evaluatePhysicalRisks: function (data) {
        const risks = {};

        // 1. 滞留・蓄積リスク (Stagnation)
        // 境界層高度(BLH)が500m未満、かつ風速(Gust)が10km/h未満で移流停止と判定
        risks.stagnation = (data.blh < 500 && data.gust < 10);

        // 2. 湿性沈着リスク (Wet Deposition) と 洗浄ギャップ
        if (data.precip > 0.5) {
            risks.wet_deposition = true;
            // 雨天にもかかわらずPM2.5が16μg/m3以上の場合、微小粒子の残留（Scavenging Gap）と判定
            risks.pm25_scavenging_gap = (data.pm25 >= 16);
        } else {
            risks.wet_deposition = false;
            risks.pm25_scavenging_gap = false;
        }

        // 3. 光化学O3二次生成リスク (Photochemical Generation)
        risks.o3_generation = (data.uv >= 5 && data.temp >= 25 && data.no2 >= 20);

        // 4. SIA(二次生成無機エアロゾル)転換リスク
        risks.sia_conversion = (data.hum >= 75 && data.dust >= 20);

        // 5. 越境輸送の推論 (Transboundary via AOD)
        // 地上PM2.5がFair以下(15以下)であり、かつAODが0.5以上の場合、上空を汚染が通過中と判定
        risks.transboundary_aloft = (data.pm25 <= 15 && data.aod >= 0.5);

        return risks;
    },

    /**
    * メイン実行関数: AIプロンプトビルダーへ渡す統合コンテキストを生成
    * @param {Object} sensorData - 気象・大気質データ
    * @param {string} locationName - 測定ポイント名 (必須)
    */
    analyze: function (sensorData, locationName) {
        // 場所の指定がない場合は厳密にエラーを投げる
        if (!locationName) {
            throw new Error("GemSyS_Logic Error: locationName must be specified.");
        }

        return {
            timestamp: new Date().toISOString(),
            location: locationName,
            raw_data: sensorData,
            aqi_assessment: this.calculateAQI(sensorData),
            eu_compliance: this.checkEU2030Limits(sensorData),
            physical_risks: this.evaluatePhysicalRisks(sensorData)
        };
    }}