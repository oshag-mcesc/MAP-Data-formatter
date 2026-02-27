/**
 * @fileoverview Provides a custom "Student Data" top-level menu in Google Sheets
 * that allows users to refresh individual seasonal tabs (fall, winter, spring)
 * or all three at once, as well as rebuild the Combined tab.
 *
 * Each seasonal tab is populated by reading its corresponding map_ source tab
 * (map_fall, map_winter, map_spring) and applying a simple filter + map:
 *
 *   .filter() — keeps only rows where col F (ClassName) contains 'MATH',
 *               mirroring the original QUERY WHERE clause.
 *   .map()    — returns only the six needed columns (P, Q, R, G, T, X)
 *               in destination column order.
 *
 * A student will have at most one row per test (e.g. Math, ELA) because
 * the filter is on ClassName (col F), and col G (Subject) is returned as-is.
 *
 * Source map_ tab column layout used (columns F–X):
 *   Col F  → ClassName  (filter column — identifies Math students)
 *   Col G  → Subject    (actual subject name written to output col D)
 *   Col P  → Student ID
 *   Col Q  → Student Last Name
 *   Col R  → Student First Name
 *   Col T  → Student Grade
 *   Col X  → Percentile
 *
 * Seasonal tab output column layout (A–F):
 *   A: Student ID
 *   B: Student Last Name
 *   C: Student First Name
 *   D: Subject  (actual subject, e.g. 'Mathematics' or 'Language Arts')
 *   E: Student Grade
 *   F: [Season] %ile
 *
 * @author  Michael O'Shaughnessy and Claude
 * @version 1.1
 */


// ── Constants ────────────────────────────────────────────────────────────────

/**
 * Configuration for each season.
 * - sourceTab  : the map_ tab that holds the raw data for that season
 * - destTab    : the destination tab name to write results into
 * - header     : the label to use for column F (the percentile column)
 */
const SEASON_CONFIG = {
  fall: {
    sourceTab : 'map_fall',
    destTab   : 'fall',
    header    : 'Fall %ile'
  },
  winter: {
    sourceTab : 'map_winter',
    destTab   : 'winter',
    header    : 'Winter %ile'
  },
  spring: {
    sourceTab : 'map_spring',
    destTab   : 'spring',
    header    : 'Spring %ile'
  }
};

/**
 * The value checked against column F (ClassName) to determine which students
 * to include. A student qualifies if their ClassName contains this string.
 * Change this constant if the filter ever needs to be updated.
 */
const SUBJECT_FILTER = 'MATH';

/**
 * Column index constants — zero-based offsets within the data read from the
 * source tab starting at column F.
 *
 * The source range starts at column F (Sheets col 6), so:
 *   F = offset 0
 *   G = offset 1
 *   H = offset 2  ... and so on through ...
 *   X = offset 18
 */
const COL = {
  F :  0,  // ClassName   — used as the student filter (contains 'MATH')
  G :  1,  // Subject     — actual subject name written to output
  P : 10,  // Student ID
  Q : 11,  // Student Last Name
  R : 12,  // Student First Name
  T : 14,  // Student Grade
  X : 18   // Percentile
};

/** First column of the data range in the source tab (column F = col 6). */
const SOURCE_START_COL = 6;

/** Number of columns to read (F through X = 19 columns). */
const SOURCE_NUM_COLS  = 19;


// ── Menu Setup ───────────────────────────────────────────────────────────────

/**
 * Runs automatically when the spreadsheet is opened.
 * Builds the "Student Data" top-level menu with options for each season,
 * an "Update All Seasons" shortcut, and an "Update Combined Tab" option.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Student Data')
    .addItem('Update Fall Tab',      'updateFall')
    .addItem('Update Winter Tab',    'updateWinter')
    .addItem('Update Spring Tab',    'updateSpring')
    .addSeparator()
    .addItem('Update All Seasons',   'updateAllSeasons')
    .addSeparator()
    .addItem('Update Combined Tab',  'consolidateStudentData')  // defined in consolidateStudentData.gs
    .addSeparator()
    .addItem('Help & Instructions',  'showHelpDialog')           // defined in studentDataHelp.gs
    .addToUi();
}


// ── Public Menu Handlers ─────────────────────────────────────────────────────

/**
 * Menu handler — refreshes the 'fall' destination tab only.
 */
function updateFall() {
  updateSeasonTab_('fall');
}

/**
 * Menu handler — refreshes the 'winter' destination tab only.
 */
function updateWinter() {
  updateSeasonTab_('winter');
}

/**
 * Menu handler — refreshes the 'spring' destination tab only.
 */
function updateSpring() {
  updateSeasonTab_('spring');
}

/**
 * Menu handler — refreshes all three seasonal tabs in sequence.
 * Runs fall → winter → spring and shows a single summary alert when done.
 */
function updateAllSeasons() {
  updateSeasonTab_('fall',   /* silent= */ true);
  updateSeasonTab_('winter', /* silent= */ true);
  updateSeasonTab_('spring', /* silent= */ true);

  SpreadsheetApp.getUi().alert('✅ All three seasonal tabs have been updated.');
}


// ── Core Worker Function ─────────────────────────────────────────────────────

/**
 * Reads data from a season's map_ source tab, identifies qualifying students
 * via a two-pass filter, then clears and rewrites the seasonal destination tab.
 *
 * Two-pass filter logic:
 *   Pass 1 — Scan column F (ClassName). Collect Student IDs of anyone whose
 *             ClassName contains SUBJECT_FILTER (e.g. 'MATH'). This tells us
 *             which students belong in this seasonal tab.
 *   Pass 2 — Loop all rows again and keep every row whose Student ID was found
 *             in Pass 1. Column G (Subject) is written as-is to output column D,
 *             so the output correctly shows 'Mathematics', 'Language Arts', etc.
 *
 * @param  {string}  season         - One of 'fall', 'winter', 'spring'
 * @param  {boolean} [silent=false] - If true, suppresses the completion alert
 *                                    (used when called from updateAllSeasons)
 */
function updateSeasonTab_(season, silent = false) {

  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const config = SEASON_CONFIG[season];

  // ── 1. Validate the source (map_) tab exists ────────────────────────────

  const sourceSheet = ss.getSheetByName(config.sourceTab);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(
      `❌ Error: Source tab "${config.sourceTab}" was not found.\n` +
      `Please make sure the tab exists and try again.`
    );
    return;
  }

  // ── 2. Read all data from the source tab (columns F through X) ──────────

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) {
    // Source tab is empty or has only a header row — nothing to process
    SpreadsheetApp.getUi().alert(
      `⚠️ Warning: Source tab "${config.sourceTab}" appears to have no data.`
    );
    return;
  }

  // Read from row 1 so we capture everything; row 1 (index 0) is skipped below
  const rawData = sourceSheet.getRange(1, SOURCE_START_COL, lastRow, SOURCE_NUM_COLS).getValues();

  // ── 3. Filter and map to output columns ──────────────────────────────────

  /**
   * Mirrors the original QUERY logic with an added deduplication step:
   *   - .slice(1)  → skip the header row (index 0)
   *   - .filter()  → keep only rows where col F (ClassName) contains 'MATH'
   *   - .map()     → return only the six columns we need in output order
   *   - .reduce()  → deduplicate by StudentID + Subject, keeping the first
   *                  occurrence. Handles the edge case where a student has
   *                  two Math teachers and would otherwise appear twice for
   *                  the same subject. Max 2 rows per student: one per subject.
   */
  const dedupedRows = rawData
    .slice(1)                                                          // skip header row
    .filter(row => row[COL.F].toString().toUpperCase()
                              .includes(SUBJECT_FILTER))               // where F contains 'MATH'
    .map(row => [
      row[COL.P],  // A: Student ID
      row[COL.Q],  // B: Last Name
      row[COL.R],  // C: First Name
      row[COL.G],  // D: Subject  (e.g. 'Mathematics', 'Language Arts')
      row[COL.T],  // E: Grade
      row[COL.X]   // F: Percentile
    ])
    .reduce((map, row) => {
      // Deduplicate by StudentID + Subject, keeping the FIRST occurrence found.
      // Key format: "StudentID|Subject" — uniquely identifies one test per student.
      // If this key already exists in the Map we simply skip the duplicate row
      // (e.g. a student with two Math teachers will only keep the first entry).
      const key = `${row[0]}|${row[3]}`;
      if (!map.has(key)) map.set(key, row);
      return map;
    }, new Map())
    .values();                // extract the deduplicated rows from the Map

  // Convert the Map iterator to a plain array for sorting and writing
  const filteredRows = Array.from(dedupedRows);

  // ── 4. Sort results by Last Name (index 1) then First Name (index 2) ───────

  /**
   * Sort the collected rows before writing so the destination tab is always
   * in alphabetical order by last name, then first name.
   * Both comparisons are case-insensitive via .toLowerCase().
   * Array indices used:  0=StudentID, 1=LastName, 2=FirstName, 3=Subject, 4=Grade, 5=Percentile
   */
  filteredRows.sort((a, b) => {
    const lastCmp = a[1].toString().toLowerCase()
                    .localeCompare(b[1].toString().toLowerCase());
    if (lastCmp !== 0) return lastCmp;

    return a[2].toString().toLowerCase()
           .localeCompare(b[2].toString().toLowerCase());
  });

  // ── 5. Find or create the destination seasonal tab ───────────────────────

  let destSheet = ss.getSheetByName(config.destTab);

  if (destSheet) {
    // Tab exists — clear everything (content + formatting) for a clean rewrite
    destSheet.clear();
  } else {
    // Tab doesn't exist yet — create it
    destSheet = ss.insertSheet(config.destTab);
  }

  // ── 6. Write the header row ──────────────────────────────────────────────

  /**
   * Column F header uses the season-specific label (e.g. "Fall %ile").
   * Columns A–E are the same across all seasons.
   */
  const headers = [
    'Student ID',
    'Student Last Name',
    'Student First Name',
    'Subject',
    'Student Grade',
    config.header        // e.g. 'Fall %ile'
  ];

  destSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  destSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  // ── 7. Write the filtered data rows ─────────────────────────────────────

  if (filteredRows.length === 0) {
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        `⚠️ No students matching "${SUBJECT_FILTER}" found in "${config.sourceTab}". ` +
        `Header written but no data rows added.`
      );
    }
    return;
  }

  // Write all rows in one call for performance
  destSheet.getRange(2, 1, filteredRows.length, headers.length).setValues(filteredRows);

  // ── 8. Auto-resize columns ───────────────────────────────────────────────

  for (let col = 1; col <= headers.length; col++) {
    destSheet.autoResizeColumn(col);
  }

  // ── 9. Freeze the header row ────────────────────────────────────────────

  /**
   * Freezing row 1 keeps the headers visible when scrolling down through
   * large data sets.
   */
  destSheet.setFrozenRows(1);

  // ── 10. Delete columns K through the last column ─────────────────────────

  /**
   * The source tab may have many columns beyond our 6 output columns.
   * We delete everything from column K (index 11 in 1-based notation) to the
   * end of the sheet to keep the tab clean and uncluttered.
   * We only do this if there are actually columns beyond K to delete.
   */
  const maxCol = destSheet.getMaxColumns();
  if (maxCol >= 11) {
    // deleteColumns(startColumn, numColumns) — both arguments are 1-based
    destSheet.deleteColumns(11, maxCol - 10);
  }

  // ── 11. Show completion alert (unless running in silent/batch mode) ────────

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      `✅ "${config.destTab}" tab updated!\n${filteredRows.length} rows written.`
    );
  }
}