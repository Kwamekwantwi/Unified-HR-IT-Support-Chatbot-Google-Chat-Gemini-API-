/**
 * @file Code.gs
 * @description Google Apps Script backend for a UNIFIED HR/IT Helpdesk Google Chat bot.
 * It classifies user queries (HR, IT, General), uses domain-specific knowledge bases,
 * and provides dynamic escalation buttons for HR or IT.
 */

// --- 1️⃣ Service Account Credentials ---
const SERVICE_ACCOUNT = {
  "type": "service_account",
  "project_id": "[YOUR_PROJECT_ID]",
  "private_key_id": "[YOUR_PRIVATE_KEY_ID]",
  "private_key": "[YOUR_PRIVATE_KEY_BEGINS_WITH_BEGIN_PRIVATE_KEY]",
  "client_email": "[YOUR_SERVICE_ACCOUNT_EMAIL]",
  "client_id": "[YOUR_CLIENT_ID]",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "[YOUR_CLIENT_X509_CERT_URL]",
  "universe_domain": "googleapis.com"
};

// --- 2️⃣ Configuration & Knowledge Base ---
const GEMINI_MODEL = "gemini-2.0-flash"; 
const MAX_CONTEXT_LENGTH = 600000; 

const IT_KNOWLEDGE_BASE_DOC_ID = '[YOUR_IT_DOC_ID]'; 
const HR_KNOWLEDGE_BASE_DOC_ID = '[YOUR_HR_DOC_ID]'; 

const IT_ESCALATION_FORM_URL = '[YOUR_IT_FORM_URL]'; 
const HR_ESCALATION_FORM_URL = '[YOUR_HR_FORM_URL]'; 


/**
 * Generates an OAuth2 access token using the provided service account credentials.
 * This token is used to authenticate requests to Google APIs like Gemini API.
 * @returns {string} The access token.
 * @throws {Error} If token generation fails.
 */
function getServiceAccountAccessToken() {
  const jwtHeader = {
    "alg": "RS256",
    "typ": "JWT"
  };

  const now = Math.floor(Date.now() / 1000); // Current time in seconds
  const expiration = now + 3600; // Token expires in 1 hour (max allowed)

  const jwtClaimSet = {
    "iss": SERVICE_ACCOUNT.client_email,
    "scope": "https://www.googleapis.com/auth/generative-language", // Scope for Gemini API
    "aud": SERVICE_ACCOUNT.token_uri,
    "exp": expiration,
    "iat": now
  };

  const encodedHeader = Utilities.base64EncodeWebSafe(JSON.stringify(jwtHeader));
  const encodedClaimSet = Utilities.base64EncodeWebSafe(JSON.stringify(jwtClaimSet));

  const signatureInput = `${encodedHeader}.${encodedClaimSet}`;
  let signature;
  try {
    signature = Utilities.computeRsaSha256Signature(signatureInput, SERVICE_ACCOUNT.private_key);
  } catch (e) {
    console.error("Error computing RSA signature (private_key issue):", e);
    throw new Error(`Authentication Error: Invalid private key format. Please check SERVICE_ACCOUNT.private_key in Code.gs. Original error: ${e.message}`);
  }
  
  const encodedSignature = Utilities.base64EncodeWebSafe(signature);

  const jwt = `${signatureInput}.${encodedSignature}`;

  const options = {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    payload: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`,
    muteHttpExceptions: true // Allows inspection of error responses
  };

  const response = UrlFetchApp.fetch(SERVICE_ACCOUNT.token_uri, options);
  const result = JSON.parse(response.getContentText());

  if (result.access_token) {
    return result.access_token;
  } else {
    console.error("Failed to get access token:", result);
    const errorDetails = result.error_description || result.error || "unknown error from token endpoint";
    throw new Error(`Failed to obtain service account access token. Details: ${errorDetails}. Please check your GCP project, service account roles, and API enablement.`);
  }
}

/**
 * Retrieves and extracts content from a specified Google Drive document ID.
 * @param {string} docId The ID of the document to retrieve.
 * @returns {Object} An object containing the document's content, a status message, and its name.
 */
function getDocumentContentById(docId) {
  let fileContent = "";
  let status = "Unknown error";
  let fileName = "Not Found";

  if (!docId || docId === 'YOUR_IT_KNOWLEDGE_BASE_DOC_ID_HERE' || docId === 'YOUR_HR_KNOWLEDGE_BASE_DOC_ID_HERE') {
    status = `Error: Document ID not configured for this domain.`;
    console.error(status);
    return { content: "", status: status, fileName: "Not Configured" };
  }

  try {
    const file = DriveApp.getFileById(docId);
    fileName = file.getName();
    const mimeType = file.getMimeType();

    const allowedMimeTypes = [
      MimeType.GOOGLE_DOCS,
      MimeType.PLAIN_TEXT,
      // Add MimeType.GOOGLE_SHEETS or MimeType.PDF if you've implemented specific parsing for them
    ];

    if (!allowedMimeTypes.includes(mimeType)) {
      status = `Failed to access: Document has unsupported MIME type (${mimeType}). Only Google Docs and Plain Text are supported.`;
      console.warn(status);
      return { content: "", status: status, fileName: fileName };
    }

    if (mimeType === MimeType.GOOGLE_DOCS) {
      fileContent = DocumentApp.openById(file.getId()).getBody().getText();
    } else if (mimeType === MimeType.PLAIN_TEXT) {
      fileContent = file.getBlob().getDataAsString();
    }
    // Add specific handling for Google Sheets or PDFs here if implemented

    if (fileContent && fileContent.length > 0) {
      // Truncate content if it's too long for the LLM
      if (fileContent.length > MAX_CONTEXT_LENGTH) {
        fileContent = fileContent.substring(0, MAX_CONTEXT_LENGTH) + "\n... (truncated)";
        status = `Success: Document "${fileName}" accessed and truncated to ${MAX_CONTEXT_LENGTH} chars.`;
      } else {
        status = `Success: Document "${fileName}" accessed.`;
      }
    } else {
      status = `Failed to access: Document "${fileName}" is empty.`;
      console.warn(status);
      fileContent = ""; // Ensure content is empty if the doc is empty
    }
  } catch (e) {
    status = `Failed to access document "${fileName}" (ID: ${docId}): ${e.message}. Please check permissions and ID.`;
    console.error(status);
    fileContent = ""; // Ensure content is empty on error
  }
  return { content: fileContent, status: status, fileName: fileName };
}

/**
 * Processes a user message, classifies its domain (HR/IT), retrieves relevant document content,
 * and calls the Gemini API to get a response.
 * @param {string} userMessage The cleaned user message.
 * @returns {Object} An object containing the AI's response, domain, and document access status.
 */
function processQueryWithGemini(userMessage) {
  let classifiedDomain = "General"; // Default domain
  let selectedDocId = null;
  let docReadStatus = "";
  let docName = "N/A";
  let geminiResponse = null;

  // 1. Authenticate with Gemini for domain classification/response generation
  let accessToken;
  try {
    accessToken = getServiceAccountAccessToken();
  } catch (authError) {
    console.error("Authentication failed during query processing:", authError);
    return { response: `Authentication failed: ${authError.message}. I cannot process your request.`, domain: classifiedDomain, documentAccessStatus: "Auth Failed" };
  }

  // 2. Classify the domain of the user's query using Gemini
  try {
    const classificationPrompt = `Classify the following user query into one of these categories: "HR", "IT", or "General". Respond ONLY with the category in JSON format: {"domain": "CATEGORY"}. If unsure, classify as "General".

User Query: "${userMessage}"
Classification:`;

    const classificationPayload = {
      contents: [{ role: "user", parts: [{ text: classificationPrompt }] }],
      generationConfig: {
        responseMimeType: "application/json",
        responseSchema: {
          type: "OBJECT",
          properties: {
            "domain": { "type": "STRING", "enum": ["HR", "IT", "General"] }
          },
          "propertyOrdering": ["domain"]
        }
      }
    };

    const classificationResponse = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${accessToken}`,
      },
      payload: JSON.stringify(classificationPayload),
      muteHttpExceptions: true
    });
    const classificationResult = JSON.parse(classificationResponse.getContentText());

    if (classificationResult.candidates && classificationResult.candidates.length > 0 &&
        classificationResult.candidates[0].content && classificationResult.candidates[0].content.parts &&
        classificationResult.candidates[0].content.parts.length > 0) {
      const parsedClassification = JSON.parse(classificationResult.candidates[0].content.parts[0].text);
      if (parsedClassification.domain) {
        classifiedDomain = parsedClassification.domain;
        console.log("Query classified as:", classifiedDomain);
      }
    }
  } catch (e) {
    console.warn("Failed to classify domain, defaulting to General:", e);
    classifiedDomain = "General"; // Fallback if classification fails
  }

  // 3. Select appropriate document ID based on classified domain
  if (classifiedDomain === "IT") {
    selectedDocId = IT_KNOWLEDGE_BASE_DOC_ID;
  } else if (classifiedDomain === "HR") {
    selectedDocId = HR_KNOWLEDGE_BASE_DOC_ID;
  }
  // If General or no specific domain doc, selectedDocId remains null or is handled by prompt


  // 4. Get content from the selected document
  let documentContent = "";
  if (selectedDocId) {
    const docResult = getDocumentContentById(selectedDocId);
    documentContent = docResult.content;
    docReadStatus = docResult.status;
    docName = docResult.fileName;
  } else {
    docReadStatus = "No specific knowledge base document for this domain.";
  }

  // 5. Construct the main response prompt for Gemini
  let mainPrompt;
  if (documentContent) {
    const assistantRole = classifiedDomain === "IT" ? "IT support assistant" : "HR assistant"; // UPDATED HR role
    mainPrompt = `You are a friendly, helpful, and concise ${assistantRole}. Your primary goal is to assist users with their requests using the provided document titled "${docName}" as your knowledge base. If the language of the request is in french, reply in french.

Here are your guidelines:
- Read the document carefully and formulate a clear, helpful, and concise answer to the user's question.
- Do not quote directly from the document. Instead, paraphrase and summarize the information in your own words.
- Use natural, conversational, and approachable language.
- If the document contains the answer, synthesize the relevant parts into a user-friendly response.
- If you cannot find the answer *within the provided document*, politely state that the information isn't available in your current knowledge base. Do NOT invent information or make assumptions.
- After providing an answer, if applicable, suggest 1-2 related topics or next steps that might be helpful based on the document's content or common ${classifiedDomain} issues. Keep these suggestions concise.
- If the user's question seems vague or could have multiple interpretations, you may ask a single clarifying question.
- If the issue sounds complex or requires human intervention, gently guide the user towards seeking direct ${classifiedDomain} support.

--- Provided Knowledge Base Document ---
${documentContent}
--- End Knowledge Base Document ---

User's Question: ${userMessage}
Your Response (Friendly ${assistantRole}):`;
  } else {
    // If no specific document content or an error occurred accessing it
    mainPrompt = `You are a friendly and helpful assistant. The user asked about ${userMessage}.
    I could not find a specific knowledge base document for this query, or there was an issue accessing it.
    Please respond by politely stating that you don't have enough information to answer their specific question from your knowledge base.`;
  }


  // 6. Call Gemini again for the main response
  try {
    const responsePayload = { contents: [{ role: "user", parts: [{ text: mainPrompt }] }] };
    const response = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${accessToken}`,
      },
      payload: JSON.stringify(responsePayload),
      muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());

    if (result.candidates && result.candidates.length > 0 &&
        result.candidates[0].content && result.candidates[0].content.parts &&
        result.candidates[0].content.parts.length > 0) {
      geminiResponse = result.candidates[0].content.parts[0].text;
    } else {
      console.warn("Gemini API did not return a valid response:", result);
      geminiResponse = "I couldn't generate a helpful response at this time. (AI issue)";
    }
  } catch (apiError) {
    console.error("Error calling Gemini API for main response:", apiError);
    geminiResponse = `There was an issue processing your request: ${apiError.message || apiError}.`;
  }

  return { response: geminiResponse, domain: classifiedDomain, documentAccessStatus: docReadStatus, documentName: docName };
}


// 4️⃣ Utility Functions
/**
 * Formats a text response for Google Chat.
 * This function now accepts an optional 'card' argument to send structured messages.
 * @param {string} text The simple text to send back to Google Chat.
 * @param {Object} [card] An optional Card object for structured messages.
 * @returns {GoogleAppsScript.Content.TextOutput} A TextOutput object with JSON content.
 */

function respond(text, card = null, accessoryBtn = null) {
  let responseObj = { "text": text || "" };

  if (card) {
    responseObj.cardsV2 = [{ "cardId": "manualCard", "card": card }];
  }

  if (accessoryBtn) {
    // The accessoryWidgets field is what creates the "Chip" next to the text
    responseObj.accessoryWidgets = [{
      "buttonList": {
        "buttons": [accessoryBtn]
      }
    }];
  }

  return ContentService.createTextOutput(JSON.stringify(responseObj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Saves user-specific data using PropertiesService.
 * @param {string} userId The ID of the user.
 * @param {string} key The key for the property.
 * @param {string} value The value to store.
 */
function saveUserData(userId, key, value) {
  PropertiesService.getUserProperties().setProperty(`${userId}_${key}`, value);
}

/**
 * Retrieves user-specific data using PropertiesService.
 * @param {string} userId The ID of the user.
 * @param {string} key The key for the property.
 * @returns {string|null} The stored value, or null if not found.
 */
function getUserData(userId, key) {
  return PropertiesService.getUserProperties().getProperty(`${userId}_${key}`);
}

/**
 * Checks if two Date objects represent the same day (ignoring time).
 * @param {Date} d1 First date.
 * @param {Date} d2 Second date.
 * @returns {boolean} True if they are the same day, false otherwise.
 */
function isSameDay(d1, d2) {
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}

/**
 * Generates a time-based greeting (Good morning/af
 */
function getTimeBasedGreeting() {
  const hour = new Date().getHours();
  if (hour < 12) {
    return "Good morning";
  } else if (hour < 18) {
    return "Good afternoon";
  } else {
    return "Good evening";
  }
}

// 5️⃣ Main Bot Handler
/**
 * Handles messages sent to the Google Chat app.
 * This function is the entry point for Google Chat events.
 * @param {Object} event The event object from Google Chat.
 * @returns {Object} A JSON object representing the response to Google Chat.
 */



/**
 * IMPROVEMENT: CLOUD LOGGING UTILITY
 * Logs events to Google Cloud Logs Explorer for the Dashboard.
 */
function logBotEvent(eventType, data) {
  try {
    // This sends the data to your Google Cloud Project logs
    console.log(JSON.stringify({
      "event_type": eventType,
      "timestamp": new Date().toISOString(),
      ...data
    }));
  } catch (e) {
    // Fallback so the bot doesn't crash if logging fails
    console.warn("Logging failed: " + e.message);
  }
}


function onMessage(event) {
  // 1. EXTRACT DATA & FIX DEFINITIONS
  const userId = event.user?.name || "unknown_user";
  const userDisplayName = event.user?.displayName || "there";
  const userMessageRaw = event.message?.text || ""; 
  
  // Clean the message (remove bot mentions)
  const cleanedMessage = userMessageRaw.replace(/@\S+\s*/, '').trim();
  const lowerMessage = cleanedMessage.toLowerCase(); 
  
  console.log("Cleaned Message:", cleanedMessage);

  // 2. USER DATA & PERSONALIZATION
  saveUserData(userId, "displayName", userDisplayName);
  const preferredName = getUserData(userId, "preferredName") || userDisplayName;

  // 3. DAILY GREETING LOGIC
  let greetingMessage = "";
  const lastInteractionDateStr = getUserData(userId, "lastInteractionDate");
  const today = new Date();

  if (!lastInteractionDateStr || !isSameDay(new Date(lastInteractionDateStr), today)) {
    greetingMessage = `${getTimeBasedGreeting()}, ${preferredName}! How can I help you today?`;
    saveUserData(userId, "lastInteractionDate", today.toDateString());
  }

  // 4. "MY NAME IS" FEATURE
  if (lowerMessage.includes("my name is")) {
    const name = cleanedMessage.replace(/my name is/i, "").trim();
    if (name) {
      saveUserData(userId, "preferredName", name);
      return respond((greetingMessage ? greetingMessage + "\n\n" : "") + `Nice to meet you, ${name}!`);
    }
  }

  // 5. COMMON GREETINGS
  const commonGreetings = ["hi", "hello", "hey", "good morning", "good afternoon", "good evening"];
  if (commonGreetings.includes(lowerMessage)) {
    return respond(greetingMessage || `${getTimeBasedGreeting()}, ${preferredName}!`);
  }

  // 6. "NO INTERNET" FAQ
  const noInternetKeywords = ["no internet", "internet not working", "wifi not working", "network problem", "can't connect to internet", "no connection"];
  if (noInternetKeywords.some(keyword => lowerMessage.includes(keyword))) {
      let wifiResp = `WiFi Issues – Connected but No Internet\n\nTroubleshooting Steps:\n1. Restart WiFi: Turn WiFi off and on again on your Mac.\n2. Switch Networks: Try switching between different networks (if available).\n3. Check for VPN or Proxy Settings: Disable VPN or custom proxy settings temporarily.\n4. Run Network Diagnostics: Open System Settings > Network > Assist Me > Diagnostics and follow the prompts.\nIf Issue Persists: Contact IT for further assistance.`;
      return respond((greetingMessage ? greetingMessage + "\n\n" : "") + wifiResp);
  }

  // 7. MANUAL KEYWORD ESCALATION (CUSTOM UI: WHITE CARD, GREEN BUTTON)
  const itEscalationKeywords = ["escalate to it", "contact the it team", "it help", "tech support", "it support", "it issue"];
  const hrEscalationKeywords = ["escalate hr", "contact the hr team", "hr help", "people support", "hr issue", "human resources"];

  let escalationDomain = null;
  if (itEscalationKeywords.some(keyword => lowerMessage.includes(keyword))) {
      escalationDomain = "IT";
  } else if (hrEscalationKeywords.some(keyword => lowerMessage.includes(keyword))) {
      escalationDomain = "HR";
  }

  if (escalationDomain) {
    let formUrl = (escalationDomain === "IT") ? IT_ESCALATION_FORM_URL : HR_ESCALATION_FORM_URL;
    let department = (escalationDomain === "IT") ? "IT" : "HR";

    if (formUrl && !formUrl.includes("YOUR_")) {
      logBotEvent("escalation_manual_request", { "user": userId, "dept": department });

      const manualEscalationCard = {
        "header": {
          "title": `Support Request: ${department}`,
          "subtitle": "Click the button below to open the form",
          "imageUrl": "https://cdn-icons-png.flaticon.com/512/563/563713.png",
          "imageType": "CIRCLE"
        },
        "sections": [{
          "widgets": [
            { "textParagraph": { "text": `Please complete the ${department} escalation form to notify the team.` } },
            {
              "buttonList": {
                "buttons": [{
                  "text": `<b>Escalate to ${department} Team</b>`,
                  "onClick": { "openLink": { "url": formUrl } },
                  "type": "FILLED", 
                  "color": { "red": 0.0, "green": 0.8, "blue": 0.0, "alpha": 1.0 } 
                }]
              }
            }
          ]
        }]
      };
      return respond(null, manualEscalationCard);
    }
  }

  // 8. GEMINI KNOWLEDGE BASE PROCESSING
  const { response: geminiAnswer, domain: currentDomain, documentAccessStatus: finalDocStatus } = processQueryWithGemini(cleanedMessage);
  
  let responseText = "";
  if (geminiAnswer) {
    responseText = geminiAnswer;
  } else {
    const domainDisplay = (currentDomain === "HR") ? "HR" : currentDomain;
    responseText = `I'm here to help, ${preferredName}! I couldn't find a direct answer in the ${domainDisplay} knowledge base.`;
  }

  // 9. PREPARE THE ACCESSORY CHIP (NEED MORE HELP?)
  let fallbackFormUrl = (currentDomain === "IT") ? IT_ESCALATION_FORM_URL : (currentDomain === "HR" ? HR_ESCALATION_FORM_URL : null);
  let helpChip = null;

  if (fallbackFormUrl && !fallbackFormUrl.includes("YOUR_")) {
    helpChip = {
      "text": "Need more help? 🆘",
      "onClick": { "openLink": { "url": fallbackFormUrl } },
      "type": "FILLED",
      "color": { "red": 0.0, "green": 0.8, "blue": 0.0, "alpha": 1.0 }
    };
  }

  // Combine Greeting and Response
  let fullTextOutput = (greetingMessage ? greetingMessage + "\n\n" : "") + responseText;

  // Final Unified Return with Accessory Widget
  return respond(fullTextOutput, null, helpChip);
}

// 6️⃣ Entry Point for Google Chat
/**
 * The main entry point for Google Chat events (POST requests).
 * Parses the incoming event and dispatches it to the onMessage handler.
 * @param {Object} e The event object from Google Chat.
 * @returns {Object} A JSON object representing the response to Google Chat.
 */
function doPost(e) {
  try {
    const event = JSON.parse(e.postData.contents);

     // --- Handle ADDED_TO_SPACE event ---
    if (event.type === 'ADDED_TO_SPACE') {
      const userNameForWelcome = event.user.displayName || "there";
      const welcomeMessage = `Hello ${userNameForWelcome}! I'm yourbot name, your unified assistant for IT and hr.\n\n` +
                             `You can ask me questions about company policies or troubleshooting IT issues. ` +
                             `If I can't answer, I can help you escalate to the right team.\n\n` +
                             `Try asking me something like: "What is our leave policy?" or "My laptop is slow".`;
      return ContentService.createTextOutput(JSON.stringify({ text: welcomeMessage })).setMimeType(ContentService.MimeType.JSON);
    }
    // --- End ADDED_TO_SPACE event handling ---



    return onMessage(event);
  } catch (error) {
    console.error("Error in doPost (full error object):", error);
    return ContentService.createTextOutput(JSON.stringify({ text: `An internal server error occurred: ${error.message || error}. Please check your Apps Script execution logs for details, and ensure all Google Cloud APIs and service account permissions are correctly configured. Try again later.` })).setMimeType(ContentService.MimeType.JSON);
  }
}
