/**
 * @fileoverview Provides the "Help & Instructions" dialog for the Student Data
 * spreadsheet. Opens a styled HTML modal dialog from the "Student Data" menu.
 *
 * The dialog content lives in studentDataHelp.html (a separate HTML file in
 * this Apps Script project). Keeping the HTML separate from the .gs code makes
 * it easier to edit the content and styling without touching the script logic.
 *
 * This file works alongside studentDataMenu.gs and consolidateStudentData.gs.
 *
 * @author  You
 * @version 1.1
 */


// ── Help Dialog Launcher ─────────────────────────────────────────────────────

/**
 * Opens the Help & Instructions modal dialog.
 * Called from the "Student Data" → "Help & Instructions" menu item.
 *
 * HtmlService.createHtmlOutputFromFile() loads the content directly from
 * the studentDataHelp.html file in this Apps Script project, keeping
 * markup and script logic cleanly separated.
 *
 * Dialog is sized at 860x660px to give the content plenty of breathing room.
 */
function showHelpDialog() {

  const html = HtmlService.createHtmlOutputFromFile('studentDataHelp')
    .setWidth(860)
    .setHeight(660)
    .setTitle('Help & Instructions');

  SpreadsheetApp.getUi().showModalDialog(html, 'Help & Instructions');
}