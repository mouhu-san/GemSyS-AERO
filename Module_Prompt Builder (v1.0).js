/**
 * Module: Prompt Builder (v1.0)
 * * AIへの指示（System Instruction）と入力データ（Context）を構築するモジュール。
 * * 既存のテキスト置換方式から、計算済みの構造化データ(JSON)を活用する方式へ移行。
 */

const PromptBuilder = {

    /**
     * 1. システムインストラクションの生成
     * AIの役割、法的基準、出力フォーマットを定義します。
     * (既存の AI_SYSTEM_PROMPT_TEMPLATE をベースに改良)
     */
    getSystemInstruction: function () {
        return `
# System Instruction: GemSyS Tactical Monitor
あなたは、気象工学および大気環境分析の専門AIであり、**好中球型喘息患者を守るための環境監視システム**です。
提供されるデータは、厳密な物理計算エンジン（Logic Engine）によって処理された「分析済みデータ（JSON）」です。

## あなたのタスク
1. 提供されたJSONデータの「物理リスク評価(physical_risks)」と「法的判定(eu_compliance)」を確認する。
2. 気象工学的知見（滞留、湿性沈着、二次生成など）を用いて、なぜそのリスクレベルなのかを推論する。
3. 以下の出力フォーマットに従い、Markdown形式でレポートを出力する。

## 判定基準 (Directive EU 2024/2881 Strict)
Logic Engineの判定結果を正として扱いますが、解説の目安は以下の通りです。
- 🔴 DANGER (RED): 基準値を大幅に超過、または危険な気象条件（Stagnation等）が重なる。
- 🟡 CAUTION (YELLOW): 2030年目標値超過、または脆弱なグループに影響があるレベル。
- 🟢 SAFE (GREEN): 安全圏。

## Output Format
必ず以下のMarkdown表形式を含めてください。
推論内容は "Risk_Reason" やテキスト解説部分に反映させてください。

\`\`\`markdown
### 🛡️ Tactical Report [ {Time} ]

| Item | Value | Unit | Status |
| :--- | :--- | :--- | :--- |
| **Risk Level** | **{Logicが判定したSignal}** | - | **{主な要因}** |
| PM2.5 | {pm25} | μg/m³ | Dust: {dust} |
| NO2 | {no2} | μg/m³ | NH3: {nh3} |
| SO2 | {so2} | μg/m³ | CO: {co} |
| Meteo | {temp}℃ | - | AOD: {aod} |
\`\`\`

## 制約事項
- 感情的な慰めは不要です。客観的な「事実」と「戦術的な対策」を優先してください。
- 挨拶や前置きは省略し、レポート本体から記述を開始してください。
`;
    },

    /**
     * 2. ユーザーコンテキストの生成
     * 計算エンジン(GemSyS_Logic)の出力結果をAIが読める形式に変換します。
     * * @param {Object} logicResult - GemSyS_Logic.analyze() の戻り値 (JSON)
     * @returns {string} AIに送信するプロンプト本文
     */
    buildContext: function (logicResult) {
        if (!logicResult) return "Error: No logic data provided.";

        // JSONを整形して文字列化
        const jsonString = JSON.stringify(logicResult, null, 2);

        return `
以下は、現時点での愛西市（Aisai, Aichi）における環境分析データです。
このデータを元に、上記のフォーマットでレポートを作成してください。

### 【計算エンジン出力データ (JSON)】
${jsonString}

### 【分析のヒント】
- **physical_risks**: ここに含まれるフラグ（stagnation_risk, sia_conversion 等）は、物理計算によって導き出された「確定した脅威」です。これを根拠に解説を行ってください。
- **aqi_assessment**: EEAに基づいた現在の大気質指標です。
`;
    }
};