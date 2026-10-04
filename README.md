# Clinical Voice-to-SOAP Scribe
A Google Apps Script automation that listens for audio recordings in a Google Drive inbox, transcribes the encounter with Gemini multimodal AI, creates a structured SOAP note and clinical summary in a Google Doc, records the consultation in a registry spreadsheet, and archives the audio file.

This project is designed for clinical intake, follow-up recordings, and consultation documentation workflows where a provider records a patient conversation and wants rapid, structured documentation without manual transcription.

## What the script does
- Scans an incoming Google Drive folder for audio or audio-like files
- Sends each recording to Gemini for clinical transcription and summarization
- Extracts structured fields such as:
  - consultation summary
  - verbatim transcript
  - SOAP note (subjective, objective, assessment, plan)
  - medication list
  - emergency flagging and notes
- Creates a formatted Google Doc chart for each consultation
- Writes a row into a registry spreadsheet with the patient encounter details
- Moves the original audio file to an archive folder after successful processing
- Automatically pauses and resumes long queues to respect Google Apps Script runtime limits

## Key features
- Clinical-grade handling of spoken consultation recordings
- Multi-model Gemini fallback logic with retry handling for transient API errors
- Automatic pausing and rescheduling when processing time approaches the Apps Script hard limit
- Red-flag alerting for emergency-related findings in the generated chart and registry
- One-click menu integration inside the Google Sheet

## Project structure
```text
.
├── Code.js            # Apps Script workflow and Gemini integration
├── README.md          # Project documentation
├── .gitignore         # Local file exclusions
└── (Google Drive + Sheets resources created at runtime)
```

## Requirements
- A Google account with access to Google Drive and Google Sheets
- A Google Sheet connected to a Google Apps Script project
- A Google Drive folder for inbound audio files
- A Google Drive folder for processed audio archives
- A Google Drive folder for generated clinical chart documents
- A Gemini API key stored in Apps Script project properties

## Before you run it

Update the configuration values at the top of `Code.js`:
```javascript
const CONFIG = {
  INBOX_FOLDER_ID: "string-id-here",
  AUDIO_ARCHIVE_FOLDER_ID: "string-id-here",
  DOCS_ARCHIVE_FOLDER_ID: "string-id-here",
  REGISTRY_SHEET_NAME: "Consultations_Registry",
  MAX_EXECUTION_TIME_MS: 4.5 * 60 * 1000
};
```

To find a folder ID:
1. Open the folder in Google Drive
2. Copy the ID from the URL
3. Example:

```text
https://drive.google.com/drive/folders/ABC123XYZ789
```

The folder ID is `ABC123XYZ789`.

## Setup steps
### 1. Create the spreadsheet and script
1. Open or create a Google Sheet.
2. Go to Extensions > Apps Script.
3. Replace the default script contents with the `Code.js` code from this project.
4. Save the project.

### 2. Add the API key
In the Apps Script editor:
1. Open Project Settings.
2. Select Script Properties.
3. Add a new property named:

```text
GEMINI_API_KEY
```

4. Set the value to your Gemini API key.
### 3. Configure folder IDs
Edit the `CONFIG` block in `Code.js` and set the proper Drive folder IDs:

- `INBOX_FOLDER_ID`: where incoming audio files are dropped
- `AUDIO_ARCHIVE_FOLDER_ID`: where processed recordings are moved
- `DOCS_ARCHIVE_FOLDER_ID`: where generated medical charts are stored

### 4. Authorize the script
When you run the script for the first time, allow the required permissions:
- Google Drive access
- Google Sheets access
- Google Docs creation/modification
- URL fetching for Gemini API requests

## Usage
1. Upload one or more supported audio files into the configured inbox folder.
2. Open the spreadsheet where the script is installed.
3. Refresh the sheet if needed.
4. In the custom menu, click:

```text
Clinical Scribe > Process Voice Consultations
```

The script will:
- process each audio file
- generate a clinical summary and SOAP note
- create a Google Doc chart
- write the results to the registry sheet
- move the source file to the archive folder

## Output
### Registry sheet
The script creates or uses a sheet named `Consultations_Registry` with these columns:

- Date / Time
- File Name
- Clinical Summary
- Emergency Status
- Google Doc Chart Link
- Source Audio Link

### Medical chart document
Each consultation creates a Google Doc titled like:

```text
Medical Chart - <filename> - yyyy-MM-dd_HHmm
```

The document contains:
- encounter summary
- SOAP note sections
- medication table
- verbatim transcript
- emergency red-flag banner when applicable

## Safety and privacy notes
This workflow processes potentially sensitive patient information. Before deployment, ensure that:

- access to the Google Drive folders is restricted appropriately
- the script’s API key is stored securely in Apps Script properties
- your organization complies with local medical and privacy requirements
- the generated documents are stored in an approved, authorized location

## Technical 
The script is built around these main functions:

- `onOpen()` — creates the custom menu item
- `processAudioConsultations()` — orchestrates the audio queue
- `callGeminiClinicalScribe()` — sends multimodal audio data to Gemini
- `generateClinicalDoc()` — builds the SOAP note chart in Google Docs
- `ensureRegistrySheetExists()` — sets up the tracking spreadsheet tab
- `scheduleNextAudioBatch()` / `clearScheduledAudioTriggers()` — handles long-running queue processing safely

## License
This project is provided as-is for workflow automation and documentation assistance. If you are using it in a clinical environment, confirm that your deployment meets your organization’s compliance and governance requirements.
