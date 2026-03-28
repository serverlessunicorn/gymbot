/**
 * @OnlyCurrentDoc
 */

function testSetCommand(setCommand = "lats 130 10 11 145 8 9 6") {
  sessionDate = startSession();
  addSet(parseSet(setCommand));
  return getLatestSetAsString();
}

/*
* Expects a non-null, non-empty string of set info.
* If the exercise is missing, but a session has been started and at least
* one exercise begun, then the last one is assumed.
* The parser guesses weights versus reps, allowing the weight to be skipped after
* the initial mention.
*/
function parseSet(setCommand) {
  if (setCommand == null || setCommand == "") {
    throw new Error("Fatal: Empty or null command passed to parseSet.");
  }
  words = setCommand.split(" ");
  if (words.length < 1) {
    throw new Error("Fatal: Empty set command '" + setCommand + "' (expected exercise name).");
  }
  exercise = words[0];
  console.log("[parseSet] Exercise name is '" + exercise + "'.");
  if (words.length < 2) {
    throw new Error("Fatal: Invalid set command '" + setCommand + "' (expected initial weight).");
  }
  weight1 = words[1];
  if (words.length < 3) {
    throw new Error("Fatal: Invalid set command '" + setCommand + "' (expected initial rep count).");
  }
  rep1 = words[1];
  return [exercise, weight1, rep1];
}


// TODOs:
// * Unlike Sessions and Sets, where the data rows can start empty, we can't bootstrap without exercise names, so
// * add bootstrap code that actually populates that sheet for real.
// * Add row managers for Sets and Exercises
// * Type safety misses the edge case where a required field never gets set. Fix this!!!
// * Convert the old session initializer to the new approach. Pay special attention to the "date display" problems we had before...are these any easier in the new regime?
// Parse set data better, & get an end-to-end test of a gym entry set up
// Add shortest edit distance fuzzy name matching
// Get chart retrieval working
// Get dropdown/menu support in slack working


/**
 * Notes:
 *   ==> columnIndex is computed during construction by comparing the object and spreadsheet layouts,
 *       although the bootstrap routine makes them consistent.
 *   ==> writeable is assumed true unless explicitly stated to be false
 *       !writeable doesn't mean const (formulas can change over time, but the formula can't be overridden)
 *   ==> fields without a default value are required (non-optional)
 *   ==> Attempts to violate any portion of type safety will result in a fatal error
 *   ==> getValue vs displayValue: The 'normal' fields use getValue to translate spreadsheet cell values into JS.
 *       Each field F has a "shadow", F_display, which is populated using getDisplayValue to preserve the
 *       literal on-screen representation visible in the spreadsheet.
 *   ==> Spreadsheet mappings and controls:
 *       'header' is the column header name in the spreadsheet (row 1 values, all in string format)
 *       'format' the Google Sheets format for the data values in that column (rows 2+). Must be one of:
 *       "@" string
 *       "0" integer
 *       "0.0" or "0.00" (two choices of resolution for decimals)
 *       "0%" percent
 *       "dd-mm-yyyy" US date
 *       "hh:mm" 12-hour, minute resolution time of day
 *   ==> In JS land, dates and times convert to objects, strings to strings, and numbers to numbers.
 *       JS string and number types default to "@" and "0" format, respectively, so they only need format
 *       controls if the default isn't appropriate. JS objects do not have a default and must be explicitly set.
 */
const SessionSchema = {
  // Data columns
  date: { header:"Date", type:"object", format:"dd-mm-yy"},
  startTime: { header:"Start Time", type: "object", format:"hh:mm" },
  latestTime: { header:"Latest Time", type:"object", format:"hh:mm" },
  location: { header: "Location", type: "string"},
  // Formula (computed) columns
  totalExercises: { header:"TOTAL EXERCISES", writeable:false, type:"number", default:"=COUNTIF(Sets!$A$2:$A, A2)" },
  averageReps: { header:"AVG REPS", writeable:false, type:"number", default:"=VLOOKUP(A2, Sets!$A$2:$Q, 17, FALSE)" },
  daysSinceLastVisit: { header:"DAYS SINCE LAST VISIT", writeable:false, type:"number", default:"=IF(A3, A2-A3, -1)" }
};

const SESSION_FORMAT = ({ date, startTime, latestTime, location, totalExercises, averageReps, daysSinceLastVisit }) => 
  `On ${date} from ${startTime} to ${latestTime}: at ${location} you did ${totalExercises} different exercises with ${averageReps} average reps. It's been ${daysSinceLastVisit} days since your last workout.`;
const SET_FORMAT = ({ date, exerciseName }) => 
  `${date}: ${exerciseName} ... more data to come ...`;
const EXERCISE_FORMAT = ({ name }) => 
  `${name} ... more data to come ...`;

class RowManager {
  /**
   * @param {string} sheetName 
   * @param {Object} schema
   */
  constructor(sheet, schema, formatter) {
    this.sheet = sheet;
    this.schema = schema;
    this.formatter = formatter;

    this.headerStartRow = 1;
    this.headerNumRows = 1;
    this.startColumn = 1;
    this.numColumns = this.sheet.getLastColumn() + 1 - this.startColumn;
    this.headers = _getHeaders();
    this._validateSchema();
  }

  _headerRectangle() {
    this.sheet.getRange(this.headerStartRow, this.headerNumRows, this.startColumn, this.numColumns);
  }

  _dataRectangle() {
    this.sheet.getRange(this.headerStartRow, this.headerNumRows, this.startColumn, this.numColumns);
  }

  _getHeaders() {
    return _headerRectangle().getValues[0];
  }

  _getDataValues() {
    return _dataRectangle().getValues[0];
  }

  _getDisplayedValues() {
    return _dataRectangle().getDisplayValues[0];
  }

  _validateSchema() {
    const required = Object.values(this.schema).map(s => s.header);
    const missing = required.filter(h => !this.headers.includes(h));
    if (missing.length > 0) throw new Error(`Missing columns: ${missing.join(", ")}`);
  }

  /**
   * Fetches the first data row in the sheet and returns it as a typed object.
   */
  getCurrent() {
    result = {}; // This will be an object conforming to the schema, populated from row 2.
    rowValues    = this._getDataValues();
    rowDisplayed = this._getDisplayedValues();
    Object.keys(this.schema).map(key => {
      fieldDefn = schema[key];
      spreadsheetColumnName = fieldDefn.header;
      spreadsheetColumn = fieldDefn.columnIndex;
      cellValue   = rowValues [spreadsheetColumn];
      cellDisplay = rowDisplay[spreadsheetColumn];
      // Implement default values:
      let field_val     = cellValue ?? fieldDefn.defaultValue;
      let field_display = cellDisplay;
      // Enforce construction-time type safety (see property setter below for post-construction enforcement)
      if (typeof field_val !== fieldDefn.type) {
        throw new Error(`Field ${key} expected type ${fieldDefn.type} but column ${spreadsheetColumnName} had type ${typeof field_val} and displays as ${field_display}.`);
      }
      // Implements the "writable by default" policy:
      let writable = fieldDefn.writable ?? true;
      // Add a read-only display field
      Object.defineProperty(result, key + "_display", {
        value: field_display,
        enumerable: true,
        configurable: false
      });
      // This is the normal field / real value:
      Object.defineProperty(result, key, {
        enumerable: true,
        configurable: false,
        get() { return field_val; },
        set(newVal) {
          // Enforce the writable flag set in the schema:
          if (!writable) throw new Error(`${key} is read only.`);
          // Enforce type safety when the object is mutated post construction
          if (typeof newVal !== fieldDefn.type) {
            throw new Error(`Expected ${fieldDefn.type} for ${key}.`);
          }
          field_val = newVal;
        }
      });
      result[key]
    })
    return result;
  }

  /**
   * Takes a typed object and writes it into the first data row of the spreadsheet.
   * Sets both values (based on the object's state) and formats (based on the schema definition).
   */
  setCurrent(obj) {
    const rowValues  = Object.keys(obj).map(key => {this.schema[key].columnIndex, obj[key]});
    const rectValues = [rowValues];
    const rowFormats = new Array(rowValues.length).fill("");
    Object.keys(obj).forEach((key) => {
      const cellFormat = null;
      fieldFormat = this.schema[key].format;
      fieldType = this.schema[key].type;
      if (schemaFormat) {
        cellFormat = fieldFormat;
      }
      else if (fieldType == "string") {
        cellFormat = "@";
      }
      else if (fieldType == "number") {
        cellFormat = "0";
      }
      else {
        throw new Error("No explicit format provided for field " + key + " with type " + fieldType);
      }
      // Place the computed format string in the right column index in the spreadsheet.
      rowFormats[this.schema[key].columnIndex] = cellFormat;
      rectFormats = [rowFormats];
    });
    _dataRectangle().setValues(rectangle).setNumberFormats(rowFormats);
  }

  /**
   * Adds a new, empty row by pushing all the data down, then performs a set() on the (new) row 2.
   */
  add(obj) {
    this.sheet.insertRowAfter(this.headerNumRows);
    setCurrent(obj);
  }

  /**
   * Formatting.
   * Can be used as a "class" method by supplying an explicit object and formatter.
   * With no arguments, will attempt to create an object from the most recent data row
   * and format it using the built-in formatter.
   */
  format(obj = null, templateFn = null) {
    obj = obj ?? getCurrent();
    templateFn = templateFn ?? this.formatter;
    return templateFn(obj);
  }
}

/*
* Checks to see if a new session entry is needed, and if so creates one.
* If not, updates the lastest time for the current session.
* In both cases, returns the stringified date of the (now) current session,
* which can be safely used as a unique, primary key for the session table.
* Throws exceptions for all fatal errors.
*/
function startSession(Sessions) {
  if (!Sessions) {
    throw new Error("Fatal: Unable to access the GymBot spreadsheet");
  }
  const now = new Date();
  const dateString = now.toLocaleDateString();
  const timeString = now.toLocaleTimeString();
  // TODO: REPLACE WITH SESSIONS.GET HERE
  // TODO: How do I handle getDisplayValue? How do I know it's a date?
  const lastSessionDateString = sessionsSheet.getRange(2, 1).getDisplayValue(); // Could be null if this is the first time using the app
  if (dateString != lastSessionDateString) {
    console.log("Starting a new gym session for today's date " + dateString + ".");
    // Ensure a new blank row right below the header row
    // TODO: REPLACE WITH SESSIONS.ADD HERE
    sessionsSheet.insertRowBefore(2);
    // TODO: THESE NEED TO BE ADDED TOTHE SESSION TYPE AS CONST STRINGS
    const sessionData = [dateString, timeString, timeString, "SF", SESSION_SETS_COUNT_FORMULA, SESSION_AVG_REPS_FORMULA, DAYDIFF_FORMULA];
    // Get the range for the new row (row 1, starting column 1, 1 row, number of columns in data)
    // Then populate it: A range is required, so the sessionData row gets wrapped in another array
    sessionsSheet.getRange(2, 1, 1, sessionData.length).setValues([sessionData]);
  }
  else {
    console.log("Extending session data for today's existing session.");
    // The only thing we update in the session itself is the latest timestamp
    sessionsSheet.getRange(2, 3, 1, 1).setValues([[timeString]]);
  }
  console.log("startSession is returning " + dateString + " as the Session key");
  return dateString;
}

const SESSIONS_SHEET_NAME = "Sessions";
const SETS_SHEET_NAME = "Sets";
const EXERCISES_SHEET_NAME = "Exercises";
let SpreadSheet = null;
let Sessions = null;
let Sets = null;
let Exercises = null;
function ensure_initialized(allowBootstrap = false) {
  if (!Spreadsheet) {
    Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!Spreadsheet) {
      throw new Error("unable to locate a spreadsheet");
    }
    sessionsSheet = Spreadsheet.getSheetByName(SESSIONS_SHEET_NAME);
    if (!sessionsSheet) {
      if (allowBootstrap) {
        sessionsSheet = SpreadSheet.createSheet(SESSIONS_SHEET_NAME);
        if (!sessionsSheet) {
          throw new Error("unable to open existing or create new Sessions tab");
        }
        const headerValues = Object.values(SessionSchema).map(s => s.header);
        const headerRow = sessionsSheet.getRange(1, 1, 1, headerValues.length);
        headerRow.setValues([headerValues])
          .setFontWeight("bold")
          .setBackground("#4a86e8") // Professional Blue
          .setFontColor("white")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
        sessionsSheet.setFrozenRows(1);
      }
      else {
          throw new Error("unable to open Sessions tab");
      }
    }
    setsSheet = Spreadsheet.getSheetByName(SETS_SHEET_NAME);
    if (!setsSheet) {
      if (allowBootstrap) {
        setsSheet = SpreadSheet.createSheet(SETS_SHEET_NAME);
        if (!setsSheet) {
          throw new Error("unable to open existing or create new Sets tab");
        }
        const headerValues = Object.values(SetsSchema).map(s => s.header);
        const headerRow = setsSheet.getRange(1, 1, 1, headerValues.length);
        headerRow.setValues([headerValues])
          .setFontWeight("bold")
          .setBackground("#4a86e8") // Professional Blue
          .setFontColor("white")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
        setsSheet.setFrozenRows(1);
      }
      else {
          throw new Error("unable to open Sets tab");
      }
    }
    exercisesSheet = Spreadsheet.getSheetByName(EXERCISES_SHEET_NAME);
    if (!exercisesSheet) {
      if (allowBootstrap) {
        exercisesSheet = SpreadSheet.createSheet(EXERCISES_SHEET_NAME);
        if (!exercisesSheet) {
          throw new Error("unable to open existing or create new Exercises tab");
        }
        const headerValues = Object.values(ExercisesSchema).map(s => s.header);
        const headerRow = exercisesSheet.getRange(1, 1, 1, headerValues.length);
        headerRow.setValues([headerValues])
          .setFontWeight("bold")
          .setBackground("#4a86e8") // Professional Blue
          .setFontColor("white")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
        exercisesSheet.setFrozenRows(1);
      }
      else {
          throw new Error("unable to open Exercises tab");
      }
    }
    // Now that we're sure we have all the sheets, create the row managers.
    Sessions  = new RowManager(sessionsSheet,  SessionsSchema,  SESSION_FORMAT);
    Sets      = new RowManager(setsSheet,      SetsSchema,      SET_FORMAT);
    Exercises = new RowManager(exercisesSheet, ExercisesSchema, EXERCISE_FORMAT);
  }
}

// Force reinitialization next time anything happens
function reset() {
  Spreadsheet = null;
}

// --- THE WEB APP HANDLER ---
function doPost(e) {
  try {
    console.log("WEB HANDLER INPUT RECEIVED: " + JSON.stringify(e));

    // Verify it's slack calling
    const signature = e.headers['x-slack-signature'];
    const timestamp = e.headers['x-slack-request-timestamp'];
    const body = e.postData.contents;
    if (!verifySlackRequest(signature, timestamp, body)) {
      console.error("Unauthorized request attempt blocked.");
      return ContentService.createTextOutput("Unauthorized").setMimeType(ContentService.MimeType.TEXT);
    }

    const param = e.parameter;
    const userText = param.text; // Text coming from Slack
    console.log("The GymBot web handler received this input: " + userText);
    ensure_initialized();
    setSummary = Sets.format(Sets.get(), SETS_VIEW);
    console.log("The GymBot web handler will return this output: " + setSummary);
    return ContentService.createTextOutput(setSummary);
  } catch (err) {
    console.error("CRITICAL ERROR: " + err.message);
    return ContentService.createTextOutput("Error: " + err.message);
  } finally {
    // Any last minute items go here :)
  }
}

const SLACK_SIGNING_SECRET = 'e11a83869549b018c40f1dcaa1fa7365';
function verifySlackRequest(signature, timestamp, body) {
  // 1. Prevent replay attacks (check if timestamp is older than 5 minutes)
  const fiveMinutesAgo = Math.floor(Date.now() / 1000) - (60 * 5);
  if (timestamp < fiveMinutesAgo) return false;

  // 2. Create the signature base string
  const sigBaseString = 'v0:' + timestamp + ':' + body;

  // 3. Hash it using your Signing Secret
  const hmac = Utilities.computeHmacSha256Signature(sigBaseString, SLACK_SIGNING_SECRET);
  const mySignature = 'v0=' + hmac.map(function(e) {
    return ('0' + (e & 0xFF).toString(16)).slice(-2);
  }).join('');

  return mySignature === signature;
}

/************************************************************/
/*********** POSTING GRAPHS *********************************/
/************************************************************/

/**
 * Posts the <index>th chart on <sheet> in <ss> to <channel> with <msg> as the text.
 * Caveat: The underlying upload in Slack is asynchronous. When this function returns
 *         the chance of subsequent failure is highly UNlikely, but there is no
 *         guarantee that users can see the image yet.
 * 
 * Slack requirements for this function to succeed:
 *     1. Bot needs files:write permission.
 *     2. Bot must be invited to <channel> to post there.
 */
function sendChartToSlack(sheet, index, label /* .png will be appended */, channel, msg) {
  const SLACK_TOKEN = 'xoxb-your-bot-token'; // Replace with your Bot User OAuth Token
  // Refresh data to ensure the chart is up to date
  SpreadsheetApp.flush();
  // Retrieve chart
  const charts = sheet.getCharts();
  if (charts.length < (index + 1)) {
    throw new Error("Fatal: Chart [" + i + "] doesn't exist in sheet " + sheet + ".");
  }
  // Instead of a basic as-is conversion like this:
  // const chartBlob = charts[index].getAs('image/png').setName(label)
  // we 'upscale' the PNG resolution for better renedering on high-res phone screens,
  // using 1600x1000 to ensure it stays sharp even when zoomed in.
  const chartBlob = originalChart.modify()
    .setOption('width', 1600)
    .setOption('height', 1000)
    .setOption('chartArea', {width: '90%', height: '80%'}) // Optional: Tighten margins
    .build()
    .getAs('image/png')
    .setName(label + ".png");
  try {
    // Upload and error check.
    result = uploadFileToSlack(SLACK_TOKEN, channel, chartBlob, msg);
    if (!result.ok) {
      throw new Error(result.error);
    }
  } catch (e) {
    throw new Error("Fatal: Chart upload failed: " + e.toString());
  }
}

/**
 * Uses the Slack V2 Upload Process to upload <blob> to <channelId>.
 */
function uploadFileToSlack(bearerToken, channelId, blob, initialComment) {
  const fileName = blob.getName();
  const fileSize = blob.getBytes().length;

  // STEP 1: Get the Upload URL
  const getUrlResponse = UrlFetchApp.fetch("https://slack.com/api/files.getUploadExternalUrl", {
    method: "post",
    headers: { Authorization: "Bearer " + bearerToken },
    payload: { filename: fileName, length: fileSize }
  });
  
  const uploadMetadata = JSON.parse(getUrlResponse.getContentText());
  if (!uploadMetadata.ok) throw new Error("Slack URL Error: " + uploadMetadata.error);

  const { upload_url, file_id } = uploadMetadata;

  // STEP 2: Upload the binary data to the provided URL
  UrlFetchApp.fetch(upload_url, {
    method: "post",
    contentType: "image/png", // Explicitly stating the format here
    payload: blob.getBytes()
  });

  // STEP 3: Complete the upload and share to the channel
  const completionResponse = UrlFetchApp.fetch("https://slack.com/api/files.completeUploadExternal", {
    method: "post",
    headers: { 
      Authorization: "Bearer " + bearerToken,
      "Content-Type": "application/json; charset=utf-8"
    },
    payload: JSON.stringify({
      files: [{ id: file_id, title: fileName }],
      channel_id: channelId,
      initial_comment: initialComment
    })
  });

  return JSON.parse(completionResponse.getContentText());
}