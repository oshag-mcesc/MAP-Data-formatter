/**
 * @fileoverview Consolidates student percentile data from the 'fall', 'winter',
 * and 'spring' tabs into a single 'Results' tab.
 *
 * Source tab column layout (A–F):
 *   A: Student ID
 *   B: Student Last Name
 *   C: Student First Name
 *   D: Subject
 *   E: Student Grade
 *   F: Percentile (column header varies by season, e.g. "Fall Percentile")
 *
 * Results tab columns:
 *   A: Student First Name
 *   B: Student Last Name
 *   C: Student ID
 *   D: Subject
 *   E: Student Grade
 *   F: Fall Percentile
 *   G: Winter Percentile
 *   H: Spring Percentile
 *
 * Rows are sorted by Student Last Name, then Student First Name.
 *
 * @author  You
 * @version 1.0
 */


/**
 * Main entry point. Run this function from the Apps Script editor
 * (or attach it to a button/trigger) to perform the consolidation.
 */
function consolidateStudentData() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 1. Load data from each seasonal tab ──────────────────────────────────

  const seasons = ['fall', 'winter', 'spring'];

  /**
   * seasonData holds an object keyed by season name.
   * Each value is an array of row arrays read from that tab.
   * Example: seasonData['fall'] = [['S001','Smith','John','Math','5','72'], ...]
   */
  const seasonData = {};

  for (const season of seasons) {
    const sheet = ss.getSheetByName(season);

    if (!sheet) {
      // Warn but continue — missing season tab means no data for that season
      SpreadsheetApp.getUi().alert(`Warning: Tab "${season}" was not found and will be skipped.`);
      seasonData[season] = [];
      continue;
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      // Tab exists but has no data rows (only a header or completely empty)
      seasonData[season] = [];
      continue;
    }

    // Read all data rows — skip row 1 (header), read columns A–F (6 columns)
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    seasonData[season] = data;
  }


  // ── 2. Build a consolidated map keyed by "StudentID|Subject" ─────────────

  /**
   * We use a Map so each unique Student+Subject combination gets exactly one
   * output row, with percentile slots for all three seasons.
   *
   * Map key   : "<studentId>|<subject>"  (pipe-separated, case-sensitive)
   * Map value : {
   *   studentId   : string,
   *   lastName    : string,
   *   firstName   : string,
   *   subject     : string,
   *   grade       : string,
   *   fall        : string|number,   // percentile or '' if absent
   *   winter      : string|number,
   *   spring      : string|number
   * }
   */
  const consolidatedMap = new Map();

  for (const season of seasons) {
    const rows = seasonData[season];

    for (const row of rows) {
      // Destructure the six columns for readability
      const [studentId, lastName, firstName, subject, grade, percentile] = row;

      // Skip completely blank rows (can appear at the bottom of a sheet)
      if (!studentId && !lastName && !firstName) continue;

      // Build a unique key for this student+subject combination
      const key = `${studentId}|${subject}`;

      if (!consolidatedMap.has(key)) {
        // First time we've seen this student+subject — create a fresh entry
        consolidatedMap.set(key, {
          studentId : studentId,
          lastName  : lastName,
          firstName : firstName,
          subject   : subject,
          grade     : grade,
          fall      : '',
          winter    : '',
          spring    : ''
        });
      }

      // Retrieve the existing entry and fill in this season's percentile.
      // If the student already had an entry (from another season), the name
      // and grade fields are already populated — we just add the percentile.
      const entry = consolidatedMap.get(key);
      entry[season] = percentile !== '' && percentile !== null && percentile !== undefined
                      ? percentile
                      : '';

      // Also update name/grade in case this season has more complete info
      // (guards against a season tab having a name where another left it blank)
      if (!entry.lastName  && lastName)  entry.lastName  = lastName;
      if (!entry.firstName && firstName) entry.firstName = firstName;
      if (!entry.grade     && grade)     entry.grade     = grade;
    }
  }


  // ── 3. Convert the Map to a plain array and sort ──────────────────────────

  /**
   * Sort order:
   *   Primary   → Student Last Name  (case-insensitive A→Z)
   *   Secondary → Student First Name (case-insensitive A→Z)
   */
  const sortedRows = Array.from(consolidatedMap.values()).sort((a, b) => {
    const lastCmp = a.lastName.toString().toLowerCase()
                    .localeCompare(b.lastName.toString().toLowerCase());
    if (lastCmp !== 0) return lastCmp;

    return a.firstName.toString().toLowerCase()
           .localeCompare(b.firstName.toString().toLowerCase());
  });


  // ── 4. Prepare the Results tab ────────────────────────────────────────────

  // Find or create the 'Combined' tab
  let resultsSheet = ss.getSheetByName('Combined');

  if (resultsSheet) {
    // Tab already exists — clear all content AND formatting so we start fresh
    resultsSheet.clear();
  } else {
    // Create a brand-new tab named 'Combined'
    resultsSheet = ss.insertSheet('Combined');
  }


  // ── 5. Write the header row ───────────────────────────────────────────────

  const headers = [
    'First Name',
    'Last Name',
    'Student ID',
    'Subject',
    'Grade',
    'Fall %ile',
    'Winter %ile',
    'Spring %ile'
  ];

  resultsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Bold the header row so it stands out visually
  resultsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');


  // ── 6. Write the data rows ────────────────────────────────────────────────

  if (sortedRows.length === 0) {
    // Nothing to write — inform the user and exit gracefully
    SpreadsheetApp.getUi().alert('Consolidation complete, but no student data was found in the Combined tab.');
    return;
  }

  /**
   * Build a 2-D array (array of arrays) that matches the Results column order:
   *   First Name | Last Name | Student ID | Subject | Grade | Fall | Winter | Spring
   */
  const outputData = sortedRows.map(entry => [
    entry.firstName,
    entry.lastName,
    entry.studentId,
    entry.subject,
    entry.grade,
    entry.fall,
    entry.winter,
    entry.spring
  ]);

  // Write all data rows in a single call (much faster than row-by-row)
  resultsSheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);


  // ── 7. Auto-resize columns for readability ────────────────────────────────

  for (let col = 1; col <= headers.length; col++) {
    resultsSheet.autoResizeColumn(col);
  }


  // ── 8. Done ───────────────────────────────────────────────────────────────

  SpreadsheetApp.getUi().alert(
    `✅ Consolidation complete!\n${outputData.length} student-subject rows written to the 'Combined' tab.`
  );
}