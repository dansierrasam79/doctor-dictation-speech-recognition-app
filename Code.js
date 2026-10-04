/**
 * CLINICAL VOICE-TO-SOAP SCRIBE PIPELINE
 * Powered by Gemini 3.8 Flash & Google Workspace
 */

// Configuration: Replace with your actual Google Drive Folder IDs
const CONFIG = {
  INBOX_FOLDER_ID: "string-id-here",
  AUDIO_ARCHIVE_FOLDER_ID: "string-id-here",
  DOCS_ARCHIVE_FOLDER_ID: "string-id-here",
  REGISTRY_SHEET_NAME: "Consultations_Registry",
  // 4.5 minutes safety threshold to avoid Apps Script's 6-minute hard termination
  MAX_EXECUTION_TIME_MS: 4.5 * 60 * 1000 
};

/**
 * Creates custom menu in the Google Sheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Clinical Scribe")
    .addItem("Process Voice Consultations", "processAudioConsultations")
    .addToUi();
}

/**
 * Main orchestration function.
 * Iterates through audio notes, transcribes via Gemini, generates formatted
 * Google Doc charts, writes index rows in Sheets, and archives source audio.
 */
function processAudioConsultations() {
  const startTime = new Date().getTime();
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");

  if (!apiKey) {
    SpreadsheetApp.getUi().alert("Error: GEMINI_API_KEY not configured in Script Properties.");
    return;
  }

  const inboxFolder = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
  const audioArchiveFolder = DriveApp.getFolderById(CONFIG.AUDIO_ARCHIVE_FOLDER_ID);
  const docsArchiveFolder = DriveApp.getFolderById(CONFIG.DOCS_ARCHIVE_FOLDER_ID);

  const files = inboxFolder.getFiles();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ensureRegistrySheetExists(ss);

  let processedCount = 0;
  let failureCount = 0;

  while (files.hasNext()) {
    // 1. Check elapsed runtime against safety limit
    const elapsedTime = new Date().getTime() - startTime;
    if (elapsedTime > CONFIG.MAX_EXECUTION_TIME_MS) {
      Logger.log(`Approaching execution limit after ${processedCount} files. Scheduling auto-resume...`);
      scheduleNextAudioBatch();
      ss.toast(`Batch paused at ${processedCount} consultations. Resuming automatically in 1 minute...`, "Auto-Resuming Queue", 10);
      return;
    }

    const file = files.next();
    const mime = file.getMimeType();

    // Accept common audio formats (m4a, mp3, wav, ogg, etc.)
    if (mime.includes("audio") || mime.includes("mp4") || mime.includes("m4a") || mime.includes("octet-stream")) {
      Logger.log("Transcribing: " + file.getName());
      ss.toast(`Transcribing ${file.getName()}...`, "Clinical Scribe", 5);

      try {
        // 2. Multimodal call to Gemini
        const scribeData = callGeminiClinicalScribe(file, apiKey);

        // 3. Generate patient chart in Google Docs
        const docResult = generateClinicalDoc(file.getName(), scribeData, docsArchiveFolder);

        // 4. Record entry in Google Sheets registry
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
        const statusText = scribeData.isEmergencyFlagged ? `⚠️ FLAG: ${scribeData.emergencyNotes || "Acute symptoms"}` : "Routine";
        
        logSheet.appendRow([
          timestamp,
          file.getName(),
          scribeData.consultationSummary,
          statusText,
          docResult.docUrl,
          file.getUrl()
        ]);

        // Highlight urgent rows in red
        const lastRow = logSheet.getLastRow();
        if (scribeData.isEmergencyFlagged) {
          logSheet.getRange(lastRow, 1, 1, 6).setBackground("#fef2f2");
          logSheet.getRange(lastRow, 4).setFontColor("#b91c1c").setFontWeight("bold");
        }

        // 5. Move audio file to archive
        file.moveTo(audioArchiveFolder);
        processedCount++;

      } catch (err) {
        Logger.log(`Failed processing ${file.getName()}: ${err.message}`);
        failureCount++;
      }
    }
  }

  // Clear any chained triggers once inbox is completely processed
  clearScheduledAudioTriggers();

  const summary = `Completed! Generated ${processedCount} chart(s). Failed: ${failureCount}.`;
  Logger.log(summary);
  ss.toast(summary, "Clinical Scribe Done", 6);
}

/**
 * Calls Gemini multimodal API with model fallback and automatic retry backoff.
 */
function callGeminiClinicalScribe(file, apiKey) {
  const base64Data = Utilities.base64Encode(file.getBlob().getBytes());
  const mimeType = file.getMimeType() === "application/octet-stream" ? "audio/mp4" : file.getMimeType();

  // Primary model and fallback model
  const models = ["gemini-3.8-flash", "gemini-3.5-flash-lite"];

  const schema = {
    type: "OBJECT",
    properties: {
      consultationSummary: { type: "STRING", description: "Concise 1-2 sentence clinical summary." },
      verbatimTranscript: { type: "STRING", description: "Accurate transcript with Doctor/Patient speakers." },
      soapNote: {
        type: "OBJECT",
        properties: {
          subjective: { type: "STRING", description: "HPI, complaints, patient history." },
          objective: { type: "STRING", description: "Exam findings, vitals spoken aloud." },
          assessment: { type: "STRING", description: "Working diagnoses and differentials." },
          plan: { type: "STRING", description: "Orders, advice, referrals, follow-up." }
        },
        required: ["subjective", "objective", "assessment", "plan"]
      },
      medications: {
        type: "ARRAY",
        items: {
          type: "OBJECT",
          properties: {
            drugName: { type: "STRING" },
            dosage: { type: "STRING" },
            route: { type: "STRING" },
            frequency: { type: "STRING" },
            duration: { type: "STRING" },
            specialInstructions: { type: "STRING" }
          },
          required: ["drugName", "dosage", "frequency"]
        }
      },
      isEmergencyFlagged: { type: "BOOLEAN" },
      emergencyNotes: { type: "STRING" }
    },
    required: ["consultationSummary", "verbatimTranscript", "soapNote", "medications", "isEmergencyFlagged"]
  };

  const payload = {
    contents: [
      {
        parts: [
          {
            text: "You are an expert Clinical Medical Scribe. Listen to this clinical encounter. Generate a verbatim transcript, structured SOAP note, an extracted medication list table, and flag any acute clinical emergency cues accurately."
          },
          {
            inline_data: {
              mime_type: mimeType,
              data: base64Data
            }
          }
        ]
      }
    ],
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: schema
    }
  };

  const options = {
    method: "POST",
    contentType: "application/json",
    headers: { "x-goog-api-key": apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  for (const model of models) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;
    const maxRetries = 3;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      Logger.log(`Calling ${model} (Attempt ${attempt}/${maxRetries})...`);
      const response = UrlFetchApp.fetch(endpoint, options);
      const code = response.getResponseCode();
      const text = response.getContentText();

      if (code === 200) {
        const json = JSON.parse(text);
        let rawContent = json.candidates[0].content.parts[0].text;
        
        // Sanitize code blocks if present
        rawContent = rawContent.replace(/```json\n?|```/g, "").trim();
        return JSON.parse(rawContent);
      }

      // Handle transient server busy or rate limit errors
      if (code === 503 || code === 429) {
        Logger.log(`${model} returned HTTP ${code}. Retrying in ${attempt * 2.5}s...`);
        Utilities.sleep(attempt * 2500);
        continue;
      }

      throw new Error(`Gemini API Error (HTTP ${code}): ${text}`);
    }
    Logger.log(`${model} was unavailable. Trying fallback model...`);
  }

  throw new Error("All Gemini models are currently busy. Please retry in a few moments.");
}

/**
 * Creates and formats a clinical Google Doc chart.
 */
function generateClinicalDoc(originalFileName, data, targetFolder) {
  const baseName = originalFileName.replace(/\.[^/.]+$/, "");
  const timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
  const docTitle = `Medical Chart - ${baseName} - ${timeStamp}`;
  
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();

  // Document Title
  const titleParagraph = body.appendParagraph("CLINICAL ENCOUNTER SOAP NOTE");
  titleParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  titleParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  // Metadata sub-header
  const metaParagraph = body.appendParagraph(`Source Recording: ${originalFileName}   |   Processed: ${new Date().toLocaleString()}`);
  metaParagraph.setFontSize(9).setForegroundColor("#64748b");
  body.appendHorizontalRule();

  // Emergency Red Flag Banner
  if (data.isEmergencyFlagged) {
    const alertPara = body.appendParagraph(`⚠️ CLINICAL RED FLAG DETECTED\n${data.emergencyNotes || "Immediate clinical intervention suggested."}`);
    alertPara.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    alertPara.setForegroundColor("#dc2626");
    body.appendHorizontalRule();
  }

  // 1. Executive Summary
  body.appendParagraph("1. Consultation Overview").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(data.consultationSummary);

  // 2. SOAP Note
  body.appendParagraph("2. Clinical SOAP Note").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  body.appendParagraph("Subjective (Patient Complaints & HPI):").setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph(data.soapNote.subjective);

  body.appendParagraph("Objective (Exam, Vitals & Observations):").setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph(data.soapNote.objective);

  body.appendParagraph("Assessment (Diagnosis & Clinical Impression):").setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph(data.soapNote.assessment);

  body.appendParagraph("Plan (Prescriptions, Referrals & Follow-up):").setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph(data.soapNote.plan);

  // 3. Medication Orders Table
  body.appendParagraph("3. Prescriptions & Medication Orders").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  if (data.medications && data.medications.length > 0) {
    const tableData = [
      ["Drug Name", "Dosage / Strength", "Route", "Frequency", "Duration", "Special Instructions"]
    ];

    data.medications.forEach(m => {
      tableData.push([
        m.drugName || "N/A",
        m.dosage || "N/A",
        m.route || "Oral",
        m.frequency || "N/A",
        m.duration || "N/A",
        m.specialInstructions || "None"
      ]);
    });

    const table = body.appendTable(tableData);
    const headerRow = table.getRow(0);
    headerRow.editAsText().setBold(true);
    for (let c = 0; c < headerRow.getNumCells(); c++) {
      headerRow.getCell(c).setBackgroundColor("#f1f5f9");
    }
  } else {
    body.appendParagraph("No new prescriptions or dosage alterations noted.");
  }

  // 4. Verbatim Dialogue Transcript
  body.appendParagraph("4. Verbatim Encounter Transcript").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const transcriptPara = body.appendParagraph(data.verbatimTranscript);
  transcriptPara.setFontSize(9).setForegroundColor("#334155");

  // Save, close, and place into the target Docs folder
  doc.saveAndClose();
  const docFile = DriveApp.getFileById(doc.getId());
  docFile.moveTo(targetFolder);

  return {
    docId: doc.getId(),
    docUrl: doc.getUrl()
  };
}

/**
 * Initializes the registry spreadsheet tab if not present.
 */
function ensureRegistrySheetExists(ss) {
  let sheet = ss.getSheetByName(CONFIG.REGISTRY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.REGISTRY_SHEET_NAME);
    sheet.appendRow([
      "Date / Time", "File Name", "Clinical Summary", "Emergency Status", "Google Doc Chart Link", "Source Audio Link"
    ]);
    sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#f8fafc");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(2, 180);
    sheet.setColumnWidth(3, 320);
    sheet.setColumnWidth(4, 150);
    sheet.setColumnWidth(5, 260);
    sheet.setColumnWidth(6, 220);
  }
  return sheet;
}

/**
 * Trigger management for automatic batch processing.
 */
function scheduleNextAudioBatch() {
  clearScheduledAudioTriggers();
  ScriptApp.newTrigger("processAudioConsultations")
    .timeBased()
    .after(60 * 1000)
    .create();
}

function clearScheduledAudioTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "processAudioConsultations") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}