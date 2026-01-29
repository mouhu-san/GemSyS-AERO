/**
 * Module: Gemini API Client (v1.0)
 * * A standardized client for calling Google's Generative Language API.
 * * Handles authentication, payload construction, and error logging.
 */

// ==========================================
// Public Interface
// ==========================================

/**
 * Sends a text generation request to Gemini API.
 * @param {string} userPrompt - The user's input or data to analyze.
 * @param {string} systemInstruction - (Optional) System-level instructions.
 * @param {string} modelName - (Optional) Model to use (default: gemini-2.5-flash).
 * @return {string|null} The generated text, or null if failed.
 */
function GEMINI_GenerateContent(userPrompt, systemInstruction, modelName) {
    const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) {
        console.error("[Gemini-Module] API Key is empty. Set 'GEMINI_API_KEY' in Script Properties.");
        return null;
    }

    const targetModel = modelName || 'gemini-2.5-flash';
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${targetModel}:generateContent?key=${apiKey}`;

    console.log(`[Gemini-Module] Target: ${targetModel}`);

    // Build Payload
    const payload = {
        "contents": [{ 
            "parts": [{ "text": userPrompt }] 
        }],
        "generationConfig": {
            "temperature": 0.0,
            "maxOutputTokens": 2048
        }
    };

    // Add System Instruction if provided
    if (systemInstruction) {
        payload.systemInstruction = {
            "parts": [{ "text": systemInstruction }]
        };
    }

    // Execute Request
    try {
        const options = {
            "method": "post",
            "contentType": "application/json",
            "payload": JSON.stringify(payload),
            "muteHttpExceptions": true
        };

        const response = UrlFetchApp.fetch(apiUrl, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        if (responseCode !== 200) {
            console.error(`[Gemini-Module] Error ${responseCode}: ${responseText}`);
            return null;
        }

        const json = JSON.parse(responseText);
        const outputText = json.candidates?.[0]?.content?.parts?.[0]?.text;

        if (!outputText) {
            console.warn(`[Gemini-Module] No text content in response.`);
            return null;
        }

        return outputText;

    } catch (e) {
        console.error("[Gemini-Module] Network/Script Error", e);
        return null;
    }
}