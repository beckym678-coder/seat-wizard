// You can customize this later or even pull from a settings sheet
const PREFERRED_SEAT_RANGE = [1, 2, 3, 4, 5]; // seats allowed for "Y" students

/**
 * Called when the add-on is first opened.
 * Ensures the Seat Wizard folder and starter templates exist.
 */
function onInstall(e) {
  onOpen(e);              // Add menu for convenience
  initializeSeatWizard(); // Create folder and copy templates
}

//Populates the App Menu
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Seat Wizard')
    .addItem('Generate Seating Charts', 'showGenerateChartSidebar')
    .addSeparator()
    .addItem('Seat Wizard Help', 'showUserManualSidebar')
    .addItem('Import/Sync Google Classroom Rosters', 'showImportClassesSidebar')
    .addItem('Set Preferential Seats', 'showPreferentialSeatSidebar')
    .addToUi();
}

/** 
 * SHOW SIDEBARS and UPDATE SIDEBAR Functions
 * */
 //MAIN Chart generator sidebar
function showGenerateChartSidebar() {
  const template = HtmlService.createTemplateFromFile('GenerateChartSidebar');
  template.classSheets = getClassSheets();
  template.layoutSheets = getLayoutSheets();
  SpreadsheetApp.getUi().showSidebar(template.evaluate().setTitle('Generate Seating Chart'));
}

//Seat Wizard Help Sidebar
function showUserManualSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('UserManualSidebar')
    .setTitle('Seat Wizard User Manual');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showPreferentialSeatSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('PreferentialSeatSidebar')
    .setTitle('Preferential Seat Settings');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showImportClassesSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ImportClassesSidebar')
    .setTitle('Import from Google Classroom');
  SpreadsheetApp.getUi().showSidebar(html);
}

function updateProgress(message) {
  try {
    google.script.run.withSuccessHandler(() => {}).updateProgressSidebar(message);
  } catch (_) {
    // ignore if sidebar not open
  }
}

function updateProgressSidebar(message) {
  const htmlOutput = HtmlService.createHtmlOutput(`<script>
    google.script.host.setHeight(200);
    document.body.innerHTML = '<div style="padding:10px;font-family:Arial">${message}</div>';
  </script>`);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Lists class sheets (“Class - …”) and layout templates (Slides) in “Seat Wizard” folder.
function getDropdownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const classes = ss.getSheets()
    .map(s => s.getName())
    .filter(name => name.startsWith('Class - '));

  let templates = [];
  let driveError = '';

  try {
    const folderId = getSeatWizardFolderIdSafe();

    const q = `'${folderId}' in parents and mimeType='application/vnd.google-apps.presentation' and trashed=false`;
    let pageToken;
    let scanned = 0;

    do {
      const resp = driveListWithRetry({
        q,
        pageSize: 100,
        pageToken: pageToken || null,
        fields: 'files(id,name,mimeType),nextPageToken'
      });

      const items = (resp && resp.files) || [];
      for (const it of items) {
        scanned++;
        if (it.name && it.name.indexOf('Layout -') === 0) {
          templates.push({ id: it.id, name: it.name });
        }
      }
      pageToken = resp.nextPageToken;
    } while (pageToken);

    Logger.log(`[SeatWizard] getDropdownData: classes=${classes.length}, scanned=${scanned}, templates=${templates.length}`);

  } catch (e) {
    driveError = `Template scan failed: ${e && e.message ? e.message : String(e)}`;
    Logger.log(`[SeatWizard] getDropdownData ERROR: ${driveError}\nStack:\n${(e && e.stack) ? e.stack : '(no stack)'}`);
  }

  return { classes, templates, driveError };
}

//When the dropdown menu switches classes, the user is switched to that worksheet. 
function switchToClassSheet(className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(className);
  if (sheet) {
    ss.setActiveSheet(sheet);
    return `Switched to ${className}`;
  } else {
    throw new Error(`Class sheet not found: ${className}`);
  }
}

/** 
 * LAYOUT MANAGEMENT Functions
 */
/** 
//Installs sample layouts in google sheets. (no longer used?)
function installSampleLayouts() {
  const TEMPLATE_ID = '1KflVLusmwZ1eSZcl2vvle4erH7q8lxX0PvqOpbcU0q8'; // your master sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const template = SpreadsheetApp.openById(TEMPLATE_ID);
    const sheets = template.getSheets().filter(s => s.getName().startsWith('Layout -'));
    if (sheets.length === 0) {
      ui.alert('No layout templates found in the master template sheet.');
      return;
    }

    let createdCount = 0;
    sheets.forEach(sh => {
      const name = sh.getName();
      if (ss.getSheetByName(name)) return; // skip if layout already exists

      const newSheet = ss.insertSheet(name);
      const range = sh.getDataRange();
      const values = range.getValues();
      const backgrounds = range.getBackgrounds();
      const fontWeights = range.getFontWeights();
      const alignments = range.getHorizontalAlignments();
      const borders = { top: true, bottom: true, left: true, right: true, vertical: true, horizontal: true };

      newSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
      newSheet.getRange(1, 1, backgrounds.length, backgrounds[0].length).setBackgrounds(backgrounds);
      newSheet.getRange(1, 1, fontWeights.length, fontWeights[0].length).setFontWeights(fontWeights);
      newSheet.getRange(1, 1, alignments.length, alignments[0].length).setHorizontalAlignments(alignments);
      newSheet.getRange(1, 1, values.length, values[0].length).setBorder(
        borders.top, borders.left, borders.bottom, borders.right, borders.vertical, borders.horizontal
      );

      // copy column widths
      const cols = sh.getMaxColumns();
      for (let c = 1; c <= cols; c++) {
        const w = sh.getColumnWidth(c);
        newSheet.setColumnWidth(c, w);
      }

      // copy row heights
      const rows = sh.getMaxRows();
      for (let r = 1; r <= rows; r++) {
        const h = sh.getRowHeight(r);
        newSheet.setRowHeight(r, h);
      }

      createdCount++;
    });

    if (createdCount > 0) {
      ui.alert(`✅ Installed ${createdCount} sample layout(s) from Seat Wizard templates.`);
    } else {
      ui.alert('All sample layouts are already installed.');
    }

  } catch (err) {
    ui.alert('❌ Error installing layouts: ' + err.message);
  }
}

function createNewLayout(name){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if(ss.getSheetByName(name)) return;
  const sheet = ss.insertSheet(name);
  sheet.getRange("A1").setValue("Desk Layout");
  return;
}

function addNewLayout() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Add New Layout', 'Enter a name for your new layout:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    const name = response.getResponseText().trim();
    if (!name) {
      ui.alert('Layout name cannot be empty.');
      return;
    }

    const sheetName = 'Layout - ' + name;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Prevent duplicates
    if (ss.getSheets().some(s => s.getName() === sheetName)) {
      ui.alert('A layout with that name already exists.');
      return;
    }

    // Create the new layout sheet
    const newSheet = ss.insertSheet(sheetName);
    newSheet.clear();

    // --- Example Layout Grid ---
    // 4 rows of desks (2x2 grid), with room features
    const layoutData = [
      ['Board', '', '', ''],
      ['1', '2', 'Door', 'Window'],
      ['3', '4', '', ''],
      ['Teacher', '', '', '']
    ];

    newSheet.getRange(1, 1, layoutData.length, layoutData[0].length).setValues(layoutData);

    // Format
    newSheet.setColumnWidths(1, 4, 80);
    newSheet.setRowHeights(1, 4, 40);
    newSheet.getRange('A1:D4').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
    newSheet.getRange('A1:D4').setBorder(true, true, true, true, true, true);
    newSheet.getRange('A1:D1').setBackground('#cccccc'); // highlight the board row
    newSheet.getRange('A4:D4').setBackground('#e0ffe0'); // highlight teacher area

    // Freeze nothing (since this is a layout)
    newSheet.setFrozenRows(0);

    ui.alert(`✅ New layout created: "${sheetName}" with a sample classroom grid.`);
  } else {
    ui.alert('No layout created.');
  }
}
*/

//🔍 Scans "Seat Wizard" Drive folder for Slides templates named like "Layout - Something".
function getSlideTemplatesInFolder() {
  const folderName = "Seat Wizard";
  const folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) {
    Logger.log(`❌ No folder named '${folderName}' found.`);
    return [];
  }

  const folder = folders.next();
  const files = folder.getFilesByType(MimeType.GOOGLE_SLIDES);
  const templates = [];

  let count = 0;
  while (files.hasNext() && count < 20) { // scan up to 20 files
    const file = files.next();
    const name = file.getName();
    if (name.startsWith("Layout -")) {
      templates.push({ name, id: file.getId() });
    }
    count++;
  }

  Logger.log(`🧙 Found ${templates.length} slide templates.`);
  return templates;
}

/**
 * SHEET MANAGEMENT Functions
 */
function getAllSheets() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

// Returns the sheets that contain class rosters
function getClassSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .map(s => s.getName())
    .filter(name => name.startsWith('Class - '));
}

//Returns the sheets that contain layouts (no longer used?)
function getLayoutSheets() {
  const sheets = getAllSheets();
  return sheets.filter(name => name.toLowerCase().includes("layout"));
}



/**
 * RANDOMIZING Functions and helpers
 */
function randomizeSeats(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  if (data.length === 0) return;

  const seatNumbers = data.map((_, i) => i + 1);
  for (let i = seatNumbers.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [seatNumbers[i], seatNumbers[j]] = [seatNumbers[j], seatNumbers[i]];
  }

  for (let i = 0; i < data.length; i++) {
    data[i][1] = seatNumbers[i];
  }

  sheet.getRange(2, 1, data.length, 2).setValues(data);
}


function randomizeAllClasses() {
  const classSheets = getClassSheets();
  classSheets.forEach(name => randomizeSeats(name));
}

function randomizeSeatsForClass(className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(className);
  if (!sheet) throw new Error(`Class sheet not found: ${className}`);

  // Load preferred seat list
  const props = PropertiesService.getDocumentProperties();
  const stored = props.getProperty('PREFERRED_SEAT_RANGE');
  const preferredSeats = stored ? JSON.parse(stored) : [];

  // A=Display, B=Student, C=Seat, D=Pref(Y), E=KA1, F=KA2
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();

  // Build Keep Away map keyed by Student Name (matches dropdown values)
  const keepAwayMap = {};
  for (const row of data) {
    const stud = row[1];
    const ka1 = row[4], ka2 = row[5];
    keepAwayMap[stud] = [ka1, ka2].filter(Boolean);
  }

  const allSeatNums = Array.from({ length: data.length }, (_, i) => i + 1);

  // Separate preferential vs regular (by column D)
  const prefStudents = data.filter(r => r[3] && r[3].toString().toUpperCase() === 'Y')
    .map(r => ({ display: r[0], stud: r[1], pref: r[3], ka1: r[4], ka2: r[5] }));
  const regularStudents = data.filter(r => !r[3] || r[3].toString().toUpperCase() !== 'Y')
    .map(r => ({ display: r[0], stud: r[1], pref: r[3], ka1: r[4], ka2: r[5] }));

  // Shuffle to randomize order
  shuffleArray(preferredSeats);
  shuffleArray(allSeatNums);

  // Layout info for adjacency checks
  const layoutSheet = ss.getSheets().find(s => s.getName().startsWith('Layout -'));
  const layout = layoutSheet ? layoutSheet.getDataRange().getValues() : [[]];
  const numRows = layout.length || 1;
  const numCols = Math.max(...layout.map(r => r.length)) || Math.max(1, data.length);
  const layoutInfo = { rows: numRows, cols: numCols };

  const seatAssignments = [];
  const usedSeats = new Set();

  const isConflict = (studName, seat, assignments, info) => {
    const [r, c] = seatToGrid(seat, info.cols);
    const conflicts = keepAwayMap[studName] || [];
    for (const a of assignments) {
      if (!conflicts.includes(a.stud)) continue;
      const [ar, ac] = seatToGrid(a.seat, info.cols);
      const adjacent = (Math.abs(r - ar) <= 1 && Math.abs(c - ac) <= 1);
      if (adjacent) return true;
    }
    return false;
  };

  // 1) Preferential students FIRST → try preferred seats, else any
  for (const s of prefStudents) {
    let chosen = null;

    for (const seat of preferredSeats) {
      if (usedSeats.has(seat)) continue;
      if (isConflict(s.stud, seat, seatAssignments, layoutInfo)) continue;
      chosen = seat; break;
    }
    if (!chosen) {
      for (const seat of allSeatNums) {
        if (usedSeats.has(seat)) continue;
        if (isConflict(s.stud, seat, seatAssignments, layoutInfo)) continue;
        chosen = seat; break;
      }
    }
    if (!chosen) throw new Error(`No available seat for ${s.stud}`);

    usedSeats.add(chosen);
    seatAssignments.push({ display: s.display, stud: s.stud, seat: chosen, pref: s.pref, ka1: s.ka1, ka2: s.ka2 });
  }

  // 2) Regular students next → any remaining seat that doesn't violate keep-away
  for (const s of regularStudents) {
    let chosen = null;
    for (const seat of allSeatNums) {
      if (usedSeats.has(seat)) continue;
      if (isConflict(s.stud, seat, seatAssignments, layoutInfo)) continue;
      chosen = seat; break;
    }
    if (!chosen) chosen = allSeatNums.find(seat => !usedSeats.has(seat));
    if (!chosen) throw new Error(`No available seat for ${s.stud}`);

    usedSeats.add(chosen);
    seatAssignments.push({ display: s.display, stud: s.stud, seat: chosen, pref: s.pref, ka1: s.ka1, ka2: s.ka2 });
  }

  // Write back sorted by seat number (A..F)
  const sorted = seatAssignments
    .sort((a, b) => a.seat - b.seat)
    .map(x => [x.display, x.stud, x.seat, x.pref, x.ka1, x.ka2]);

  sheet.getRange(2, 1, sorted.length, 6).setValues(sorted);
}

function getPreferentialSeats() {
  const props = PropertiesService.getDocumentProperties();
  const data = props.getProperty('PREFERRED_SEAT_RANGE');
  return data ? JSON.parse(data) : [];
}

function checkPreferentialSeatCounts(seats) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheets = ss.getSheets().filter(s => s.getName().startsWith('Class -'));
  const warnings = [];

  for (const sheet of classSheets) {
    const data = sheet.getRange(2, 3, Math.max(0, sheet.getLastRow() - 1), 1).getValues();
    const yCount = data.filter(row => row[0] && row[0].toString().toUpperCase() === 'Y').length;

    if (yCount > seats.length) {
      warnings.push(`${sheet.getName()}: ${yCount} preferential students but only ${seats.length} available seats.`);
    }
  }

  return warnings;
}

//Keep Away Code
function setupKeepAwayDropdowns(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Student Name is column B in the canonical layout we write
  const names = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat().filter(Boolean);
  if (!names.length) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(names, true)
    .setAllowInvalid(true)
    .build();

  sheet.getRange('E1').setValue('Keep Away 1');
  sheet.getRange('F1').setValue('Keep Away 2');
  sheet.getRange(2, 5, lastRow - 1, 1).setDataValidation(rule);
  sheet.getRange(2, 6, lastRow - 1, 1).setDataValidation(rule);
}


function shuffleArray(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
}

function randomizeSeatsForAllClasses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheets = ss.getSheets().filter(s => s.getName().startsWith('Class -'));
  const ui = SpreadsheetApp.getUi();

  if (classSheets.length === 0) {
    ui.alert('No class sheets found to randomize.');
    return;
  }

  let successCount = 0;
  let errorList = [];

  classSheets.forEach(sheet => {
    try {
      randomizeSeatsForClass(sheet.getName());
      successCount++;
    } catch (err) {
      Logger.log(`❌ Error randomizing ${sheet.getName()}: ${err.message}`);
      errorList.push(`${sheet.getName()}: ${err.message}`);
    }
  });

  // ✅ Toast notification for quick feedback
  ss.toast(
    `✅ Randomized ${successCount} class(es) successfully.`,
    'Seat Wizard',
    5
  );

  // Detailed alert (optional but useful for debugging)
  let message = `✅ Randomized ${successCount} class(es) successfully.`;
  if (errorList.length > 0) {
    message += `\n⚠️ Some classes had issues:\n${errorList.join('\n')}`;
  }

  ui.alert(message);
}

/**
 * SEATING CHART GENERATION Functions
 */
function generateSeatingChartFromSlideTemplate(className, templateId, presentationId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheet = ss.getSheetByName(className);
  if (!classSheet) throw new Error(`Class sheet not found: ${className}`);

  // --- Detect columns by header (case-insensitive) ---
  const headers = classSheet.getRange(1, 1, 1, classSheet.getLastColumn())
    .getValues()[0]
    .map(h => h.toString().trim().toLowerCase());

  const displayCol = headers.findIndex(h => h.includes('display')) + 1;
  const seatCol    = headers.findIndex(h => h.includes('seat')) + 1;

  if (!displayCol || !seatCol) {
    throw new Error(`Missing headers in "${className}". Found: ${headers.join(', ')}`);
  }

  const numRows = Math.max(0, classSheet.getLastRow() - 1);
  if (numRows === 0) throw new Error(`No students found in "${className}".`);

  const values = classSheet.getRange(2, 1, numRows, classSheet.getLastColumn()).getValues();
  const seatMap = new Map();
  for (const row of values) {
    const display = row[displayCol - 1];
    const seat    = row[seatCol - 1];
    if (display && seat) seatMap.set(String(seat).trim(), display);
  }
  if (seatMap.size === 0) throw new Error(`No valid seat assignments in "${className}".`);

  // --- Open template ---
  const templatePres = SlidesApp.openById(templateId);
  const templateSlides = templatePres.getSlides();
  if (!templateSlides.length) throw new Error('Template has no slides.');

  // --- Create timestamp for filename ---
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");

  // --- Create or open destination presentation ---
  let pres, file;
  if (presentationId) {
    pres = SlidesApp.openById(presentationId);
    file = DriveApp.getFileById(presentationId);
  } else {
    const folderName = 'Seat Wizard';
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    const title = `Seating Chart - ${className} (${timestamp})`;
    const created = SlidesApp.create(title);
    pres = created;
    file = DriveApp.getFileById(created.getId());
    folder.addFile(file);
    try { DriveApp.getRootFolder().removeFile(file); } catch (_) {}
  }

  // Remove initial slide if present
  const initialSlides = pres.getSlides();
  if (initialSlides.length > 0) { try { initialSlides[0].remove(); } catch (_) {} }

  // Title slide
  const classTitle = className.replace(/^Class - /, '');
  const titleSlide = pres.appendSlide(SlidesApp.PredefinedLayout.TITLE);
  const titleElems = titleSlide.getPageElements();
  if (titleElems.length > 0) titleElems[0].asShape().getText().setText(classTitle);
  if (titleElems.length > 1) { try { titleElems[1].remove(); } catch(_) {} }

  // Helper to safely copy basic styles after setText
  function reapplyStyles(textRange, oldStyle) {
    const newStyle = textRange.getTextStyle();
    try { newStyle.setFontFamily(oldStyle.getFontFamily()); } catch(_) {}
    try { newStyle.setFontSize(oldStyle.getFontSize()); } catch(_) {}
    try { newStyle.setForegroundColor(oldStyle.getForegroundColor()); } catch(_) {}
    try { newStyle.setBold(oldStyle.isBold()); } catch(_) {}
    try { newStyle.setItalic(oldStyle.isItalic()); } catch(_) {}
    try { newStyle.setUnderline(oldStyle.isUnderline()); } catch(_) {}
    try {
      const paraOld = textRange.getParagraphStyle();
      const paraNew = textRange.getParagraphStyle();
      paraNew.setParagraphAlignment(paraOld.getParagraphAlignment());
    } catch(_) {}
  }

  // Duplicate and populate each slide
  templateSlides.forEach((tplSlide, idx) => {
    const slide = pres.appendSlide(tplSlide);

    slide.getShapes().forEach(shape => {
      if (!shape.getText) return;

      let content;
      try { content = shape.getText().asString().trim(); } catch(_) { return; }
      if (!/^\d+$/.test(content)) return; // only numeric placeholders

      const studentName = seatMap.get(content);
      const tr = shape.getText();

      if (studentName) {
        const oldStyle = tr.getTextStyle();
        tr.setText(studentName);
        reapplyStyles(tr, oldStyle);
      } else {
        tr.setText('');
      }
    });

    // Label if there are two slides (Student/Teacher View)
    if (templateSlides.length === 2) {
      const label = idx === 0 ? ' (Student View)' : ' (Teacher View)';
      slide.insertTextBox(classTitle + label)
        .setLeft(20).setTop(20).setWidth(300);
    }
  });

  Logger.log(`✅ Generated seating chart for ${classTitle} at ${timestamp}`);
  return pres.getUrl();
}


function generateSeatingChartFromSidebar(className, layoutName) {
  Logger.log(`[Sidebar] Received request: class=${className}, layout=${layoutName}`);
  return generateSeatingChart(className, layoutName);
}

function generateSeatingChart(className, layoutName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheet = ss.getSheetByName(className);
  const layoutSheet = ss.getSheetByName(layoutName);
  if (!classSheet || !layoutSheet) throw new Error(`Missing sheet: ${!classSheet ? className : layoutName}`);

  const slideTitle = `Seating Chart - ${className}`;
  const folderName = "Seating Wizard";
  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");
  const pres = SlidesApp.create(`${slideTitle} (${timestamp})`);
  const file = DriveApp.getFileById(pres.getId());
  folder.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (_) {}

  const slide = pres.getSlides()[0];

  // Remove any default elements instead of slide.clear()
  const els = slide.getPageElements();
  for (let i = els.length - 1; i >= 0; i--) {
    try { els[i].remove(); } catch (_) {}
  }

  // Geometry
  const layout = layoutSheet.getDataRange().getValues();
  const numRows = layout.length;
  const numCols = Math.max(...layout.map(r => r.length));
  const margin = 40;
  const slideW = pres.getPageWidth();
  const slideH = pres.getPageHeight();
  const cellW = (slideW - 2 * margin) / numCols;
  const cellH = (slideH - 2 * margin) / numRows;

  // seat number → student name map
  // Columns: A=Display Name, B=Student Name, C=Seat #
  const seatMap = new Map();
  const students = classSheet.getRange(2, 1, Math.max(0, classSheet.getLastRow() - 1), 3).getValues();
  for (const [displayName, , seat] of students) {
    if (displayName && seat) seatMap.set(Number(seat), displayName);
  }


  for (const [name, seat] of students) {
    if (name && seat) seatMap.set(Number(seat), name);
  }

  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {
      const raw = (layout[r][c] || '').toString().trim();
      if (!raw) continue;

      const val = raw.toLowerCase();
      let isFeature = false, featureText = '', featureFill = '#cccccc';
      let isDesk = false, studentName = '';

      if (val === 'door') { isFeature = true; featureText = 'Door'; featureFill = '#795548'; }
      else if (val === 'window') { isFeature = true; featureText = 'Window'; featureFill = '#90caf9'; }
      else if (val === 'board') { isFeature = true; featureText = 'Board'; featureFill = '#9e9e9e'; }
      else if (val.includes('teacher')) { isFeature = true; featureText = 'Teacher'; featureFill = '#a5d6a7'; }
      else if (!isNaN(Number(val))) { isDesk = true; studentName = seatMap.get(Number(val)) || ''; }
      else continue;

      // Larger desks, no overlap
      const scaleFactor = 1.15;
      const padding = 2;
      const maxW = cellW - padding, maxH = cellH - padding;
      let w = Math.min(cellW * scaleFactor, maxW);
      let h = Math.min(cellH * scaleFactor, maxH);
      // feature sizing
      if (isFeature && featureText === 'Board') { w = Math.min(cellW * 3, maxW * 3); h = Math.min(cellH * 0.6, maxH); }
      if (isFeature && featureText === 'Teacher') { w = Math.min(cellW * 2, maxW * 2); h = Math.min(cellH * 0.7, maxH); }
      if (isFeature && featureText === 'Window') { w = Math.min(cellW * 0.5, maxW); h = Math.min(cellH * 2, maxH * 2); }

      const x = margin + c * cellW + (cellW - w) / 2;
      const y = margin + r * cellH + (cellH - h) / 2;

      const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);

      if (isDesk) {
        shape.getFill().setSolidFill('#D2B48C'); // tan
        try {
          const line = shape.getLine();
          if (line) { line.setWeight(1); line.getFill().setSolidFill('#8B7765'); }
        } catch (_) {}
        if (studentName) {
          const tr = shape.getText();
          tr.setText(studentName);
          const ts = tr.getTextStyle();
          ts.setFontFamily('Arial').setFontSize(9).setBold(false).setForegroundColor('#000000');
          tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        }
      } else if (isFeature) {
        shape.getFill().setSolidFill(featureFill);
        const tr = shape.getText();
        tr.setText(featureText);
        const ts = tr.getTextStyle();
        ts.setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor('#000000');
        tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      }
    }
  }
  // ✅ Toast to confirm completion
  SpreadsheetApp.getActiveSpreadsheet().toast(
   `✅ Seating chart generated for ${className}.`,
    'Seat Wizard',
    5
  );


  return pres.getUrl();
}

//🧩 Generates seating charts for ALL classes using a Slides template.
function generateAllSeatingCharts(selectedTemplateId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheets = ss.getSheets().filter(s => s.getName().startsWith('Class -'));
  if (classSheets.length === 0) throw new Error("No class sheets found.");

  const folderName = "Seat Wizard";
  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");
  const title = `Seating Charts - All Classes (${timestamp})`;

  // Create master presentation
  const pres = SlidesApp.create(title);
  const file = DriveApp.getFileById(pres.getId());
  folder.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (_) {}

  // Remove Google’s default slide once
  try {
    const firstSlide = pres.getSlides()[0];
    if (firstSlide) firstSlide.remove();
  } catch (_) {}

  const total = classSheets.length;
  let count = 0;
  let errors = [];

  classSheets.forEach(sheet => {
    const className = sheet.getName();
    count++;
    ss.toast(`Generating chart ${count} of ${total}: ${className}`, 'Seat Wizard', 5);

    try {
      generateSeatingChartFromSlideTemplate(className, selectedTemplateId, pres);
      Logger.log(`✅ Generated slides for ${className}`);
    } catch (err) {
      Logger.log(`❌ Error: ${err.message}`);
      errors.push(`${className}: ${err.message}`);
    }
  });

  const url = pres.getUrl();
  const result = {
    success: true,
    message: `✅ Generated charts for ${total - errors.length} of ${total} class(es).`,
    url,
    errors
  };

  SpreadsheetApp.getActiveSpreadsheet().toast(`✅ All seating charts generated.`, 'Seat Wizard', 5);
  return result;
}

function generateSeatingChartFromSlideTemplate(className, templateId, presentationOrId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheet = ss.getSheetByName(className);
  if (!classSheet) throw new Error(`Class sheet not found: ${className}`);

  // === Detect columns dynamically ===
  const headers = classSheet.getRange(1, 1, 1, classSheet.getLastColumn())
    .getValues()[0]
    .map(h => h.toString().trim().toLowerCase());

  const displayCol = headers.findIndex(h => h.includes('display')) + 1;
  const seatCol = headers.findIndex(h => h.includes('seat')) + 1;

  if (!displayCol || !seatCol) {
    throw new Error(`Missing expected columns in "${className}". Found headers: ${headers.join(', ')}`);
  }

  // === Build seat map ===
  const numRows = Math.max(0, classSheet.getLastRow() - 1);
  if (numRows === 0) throw new Error(`No students found in "${className}".`);

  const values = classSheet.getRange(2, 1, numRows, classSheet.getLastColumn()).getValues();
  const seatMap = new Map();
  for (const row of values) {
    const display = row[displayCol - 1];
    const seat = row[seatCol - 1];
    if (display && seat) seatMap.set(String(seat).trim(), display);
  }
  if (seatMap.size === 0) throw new Error(`No valid seat assignments found in "${className}".`);

  // === Open template presentation ===
  const templatePres = SlidesApp.openById(templateId);
  const templateSlides = templatePres.getSlides();
  if (templateSlides.length === 0) throw new Error("Template has no slides.");

  // === Prepare or open the target presentation ===
  let pres;
  let createdNew = false;

  if (presentationOrId) {
    if (typeof presentationOrId.getId === 'function') {
      pres = presentationOrId;
    } else {
      pres = SlidesApp.openById(String(presentationOrId));
    }
  } else {
    const folderName = "Seat Wizard";
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");
    const created = SlidesApp.create(`Seating Chart - ${className} (${timestamp})`);
    pres = created;

    const file = DriveApp.getFileById(created.getId());
    folder.addFile(file);
    try { DriveApp.getRootFolder().removeFile(file); } catch (_) {}

    createdNew = true;
  }

  // === Remove default slide ONLY for new presentations ===
  if (createdNew) {
    const slidesNow = pres.getSlides();
    if (slidesNow.length > 0) {
      try { slidesNow[0].remove(); } catch (_) {}
    }
  }

  // === Create title slide ===
  const classTitle = className.replace(/^Class - /, '');
  const titleSlide = pres.appendSlide(SlidesApp.PredefinedLayout.TITLE);
  const titleElements = titleSlide.getPageElements();
  if (titleElements.length > 0) titleElements[0].asShape().getText().setText(classTitle);
  if (titleElements.length > 1) { try { titleElements[1].remove(); } catch (_) {} }

  // === Helper to reapply text style ===
  function reapplyStyles(textRange, oldStyle) {
    const newStyle = textRange.getTextStyle();
    try { newStyle.setFontFamily(oldStyle.getFontFamily()); } catch (_) {}
    try { newStyle.setFontSize(oldStyle.getFontSize()); } catch (_) {}
    try { newStyle.setForegroundColor(oldStyle.getForegroundColor()); } catch (_) {}
    try { newStyle.setBold(oldStyle.isBold()); } catch (_) {}
    try { newStyle.setItalic(oldStyle.isItalic()); } catch (_) {}
    try { newStyle.setUnderline(oldStyle.isUnderline()); } catch (_) {}
    try {
      const paraOld = textRange.getParagraphStyle();
      const paraNew = textRange.getParagraphStyle();
      paraNew.setParagraphAlignment(paraOld.getParagraphAlignment());
    } catch (_) {}
  }

  // === Duplicate and populate template slides ===
  templateSlides.forEach((tplSlide, i) => {
    const slide = pres.appendSlide(tplSlide);

    slide.getShapes().forEach(shape => {
      if (!shape.getText) return;
      let text;
      try { text = shape.getText().asString().trim(); } catch (_) { return; }

      if (!/^\d+$/.test(text)) return; // skip non-seat numbers

      const student = seatMap.get(text);
      const tr = shape.getText();

      if (student) {
        const oldStyle = tr.getTextStyle();
        tr.setText(student);
        reapplyStyles(tr, oldStyle);
      } else {
        tr.setText('');
      }
    });

  });

  Logger.log(`✅ Generated seating chart for ${classTitle}`);
  return pres.getUrl();
}

function generateAllSeatingChartsFromTemplate(templateId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheets = ss.getSheets().filter(s => s.getName().startsWith('Class - '));

  if (classSheets.length === 0) throw new Error("No class sheets found.");

  const folderName = "Seating Wizard";
  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");
  const pres = SlidesApp.create(`Seating Charts - All Classes (${timestamp})`);
  const file = DriveApp.getFileById(pres.getId());
  folder.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (_) {}

  // --- Wait a moment for Drive indexing to finish ---
  Utilities.sleep(1500);

  // Remove Google’s default title slide
  try {
    const firstSlide = pres.getSlides()[0];
    if (firstSlide) firstSlide.remove();
  } catch (e) {
    Logger.log("No initial slide to remove: " + e);
  }

  // --- Generate slides for each class ---
  classSheets.forEach(sheet => {
    const className = sheet.getName();
    try {
      generateSeatingChartFromSlideTemplate(className, templateId, pres);
      Logger.log(`✅ Added charts for ${className}`);
    } catch (err) {
      Logger.log(`❌ Error generating chart for ${className}: ${err.message}`);
    }
  });

  Logger.log(`🎉 All charts generated successfully: ${pres.getUrl()}`);
  return pres.getUrl();
}

// Helper: Convert seat # → (row, col) grid coordinates
function seatToGrid(seat, cols) {
  const r = Math.floor((seat - 1) / cols);
  const c = (seat - 1) % cols;
  return [r, c];
}

//Helper: debug code
/**
function debugSeatMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheets = ss.getSheets().filter(s => s.getName().startsWith("Class - "));
  
  if (classSheets.length === 0) {
    Logger.log("⚠️ No class sheets found.");
    return;
  }

  classSheets.forEach(sheet => {
    const className = sheet.getName();
    Logger.log(`\n🧙 Checking ${className}...`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(h => h.toString().trim().toLowerCase());
    const displayNameCol = headers.findIndex(h => h.includes('display name') || h.includes('student name')) + 1;
    const seatCol = headers.findIndex(h => h.includes('seat')) + 1;

    if (!displayNameCol || !seatCol) {
      Logger.log(`❌ Missing columns in ${className}: ${headers.join(', ')}`);
      return;
    }

    const numRows = sheet.getLastRow() - 1;
    if (numRows <= 0) {
      Logger.log(`⚠️ No student rows in ${className}`);
      return;
    }

    const values = sheet.getRange(2, 1, numRows, sheet.getLastColumn()).getValues();
    const seatMap = new Map();
    for (const row of values) {
      const display = row[displayNameCol - 1];
      const seat = row[seatCol - 1];
      if (display && seat) seatMap.set(seat.toString(), display);
    }

    if (seatMap.size === 0) {
      Logger.log(`❌ Seat map is empty for ${className}`);
    } else {
      Logger.log(`✅ ${seatMap.size} seat(s) found:`);
      seatMap.forEach((v, k) => Logger.log(`   Seat ${k} → ${v}`));
    }
  });
}*/

/**
 * IMPORT GOOGLE CLASSROOM ROSTERS Functions
 */
function importSelectedCourses(courseIds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const imported = [];

  courseIds.forEach(courseId => {
    try {
      const course = Classroom.Courses.get(courseId);
      const studentsResponse = Classroom.Courses.Students.list(courseId);
      const students = studentsResponse.students || [];
      const studentNames = students.map(s => s.profile.name.fullName);

      // Merge into sheet (preserve settings) and ensure canonical layout
      importOrSyncClassRoster(course.name, studentNames);

      // Auto-assign (and re-validate) seats respecting your rules
      randomizeSeatsForClass(`Class - ${course.name}`);

      imported.push(course.name);
    } catch (err) {
      Logger.log(`❌ Error importing course ${courseId}: ${err.message}`);
    }
  });

  ss.toast(`✅ Imported ${imported.length} class(es) and auto-assigned seats.`, 'Seat Wizard', 5);
  return { success: true, count: imported.length };
}

function importOrSyncClassRoster(courseName, studentNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `Class - ${courseName}`;
  let sheet = ss.getSheetByName(sheetName);

  // Canonical headers we will always enforce on write:
  const headersCanonical = ['Display Name', 'Student Name', 'Seat Number', 'Preferential Seating', 'Keep Away 1', 'Keep Away 2'];

  if (!sheet) {
    // Create brand new sheet in canonical layout
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headersCanonical.length).setValues([headersCanonical]);
    sheet.getRange('A1:F1').setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Seed rows: Display = Student = official name; seat = 1..N
    const rows = studentNames.map((n, i) => [n, n, i + 1, '', '', '']);
    if (rows.length) sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    setupKeepAwayDropdowns(sheet);
    return sheet;
  }

  // Build a resilient view of existing data using current headers (even if wrong order).
  const pos = getHeaderPositions_(sheet);

  // If no headers or broken, try to read best-effort with defaults
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 6);
  const bodyRows = Math.max(0, lastRow - 1);
  const current = bodyRows ? sheet.getRange(2, 1, bodyRows, lastCol).getValues() : [];

  // Extract existing rows keyed by official Student Name (to preserve settings)
  const byStudent = new Map();
  for (const row of current) {
    const display = pos.displayCol ? row[pos.displayCol - 1] : row[0]; // fallback
    const student = pos.studentCol ? row[pos.studentCol - 1] : row[1];
    const seat    = pos.seatCol    ? row[pos.seatCol - 1]    : row[2];
    const pref    = pos.prefCol    ? row[pos.prefCol - 1]    : row[3];
    const ka1     = pos.ka1Col     ? row[pos.ka1Col - 1]     : row[4];
    const ka2     = pos.ka2Col     ? row[pos.ka2Col - 1]     : row[5];

    if (student) {
      byStudent.set(student, {
        display: display || student,
        student,
        seat: Number(seat) || null,
        pref: pref || '',
        ka1: ka1 || '',
        ka2: ka2 || ''
      });
    }
  }

  // Merge with new roster (official names array)
  const N = studentNames.length;
  const allSeats = Array.from({ length: N }, (_, i) => i + 1);
  const used = new Set();

  // Preserve unique, in-range seats for returning students
  studentNames.forEach(stud => {
    const ex = byStudent.get(stud);
    if (ex && ex.seat && ex.seat >= 1 && ex.seat <= N && !used.has(ex.seat)) {
      used.add(ex.seat);
    }
  });

  const freeSeats = allSeats.filter(s => !used.has(s));

  const merged = studentNames.map(stud => {
    const ex = byStudent.get(stud);
    const display = ex?.display || stud;
    const pref    = ex?.pref || '';
    const ka1     = ex?.ka1  || '';
    const ka2     = ex?.ka2  || '';
    let seat      = ex?.seat || null;

    if (!seat || seat < 1 || seat > N || used.has(seat)) {
      seat = freeSeats.shift() || null;
    }
    used.add(seat);
    return [display, stud, seat, pref, ka1, ka2];
  });

  // Rewrite the sheet in canonical order
  sheet.clear();
  sheet.getRange(1, 1, 1, headersCanonical.length).setValues([headersCanonical]);
  sheet.getRange('A1:F1').setFontWeight('bold');
  sheet.setFrozenRows(1);
  if (merged.length) sheet.getRange(2, 1, merged.length, 6).setValues(merged);

  // Data validation for Pref + Keep-Away
  const prefRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Y'], true)
    .setAllowInvalid(true)
    .build();
  if (merged.length) sheet.getRange(2, 4, merged.length, 1).setDataValidation(prefRule);

  setupKeepAwayDropdowns(sheet); // KA dropdowns reference col B (Student Name)
  return sheet;
}


//Gets the list of Google Classroom courses for the active user.
function getClassroomCourses() {
  try {
    const response = Classroom.Courses.list({ courseStates: ['ACTIVE'] });
    const courses = response.courses || [];
    return courses.map(course => ({
      id: course.id,
      name: course.name
    }));
  } catch (error) {
    Logger.log('Error fetching courses: ' + error.message);
    throw new Error('Unable to fetch Google Classroom courses. Make sure the Classroom API is enabled.');
  }
}

//Not sure if this is for importing rosters, randomizing, or both
function savePreferentialSeats(selectedSeats) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('PREFERRED_SEAT_RANGE', JSON.stringify(selectedSeats));
}

/** Helper: return column indexes by header text (case-insensitive). */
function getHeaderPositions_(sheet) {
  if (!sheet || sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    return {};
  }
  const raw = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headers = raw.map(h => (h || '').toString().trim().toLowerCase());

  // Accept several spellings to be forgiving
  const idx = (want) => headers.findIndex(h => h === want) + 1;
  const idxLike = (frag) => headers.findIndex(h => h.includes(frag)) + 1;

  const displayCol = idx('display name') || idxLike('display') || 0;
  const studentCol = idx('student name') || idxLike('student') || 0;
  const seatCol    = idx('seat number')  || idxLike('seat')    || 0;
  const prefCol    = idx('preferential seating') || idxLike('preferential') || 0;
  const ka1Col     = idx('keep away 1') || idxLike('keep away 1') || 0;
  const ka2Col     = idx('keep away 2') || idxLike('keep away 2') || 0;

  return { displayCol, studentCol, seatCol, prefCol, ka1Col, ka2Col, headersRaw: raw };
}






/**
 * GOOGLE DRIVE MANAGEMENT Functions
 */

//Uses the Advanced Drive (v3) API to find or create the “Seat Wizard” folder.
function getSeatWizardFolderIdSafe() {
  const FOLDER_NAME = 'Seat Wizard';

  // Try to find the folder by name
  const found = driveListWithRetry({
    q: "mimeType='application/vnd.google-apps.folder' and trashed=false and name='" + FOLDER_NAME + "'",
    pageSize: 1,
    fields: 'files(id,name)'
  });
  if (found.files && found.files.length) return found.files[0].id;

  // Create folder if not found
  const created = Drive.Files.create({
    name: FOLDER_NAME,
    mimeType: 'application/vnd.google-apps.folder'
  });
  return created.id;
}

//Retry wrapper for Drive.Files.list() (Advanced Drive v3).
function driveListWithRetry(params, attempts) {
  const maxAttempts = attempts || 3;
  let delay = 300;
  for (let i = 1; i <= maxAttempts; i++) {
    try {
      return Drive.Files.list(params);
    } catch (e) {
      if (i === maxAttempts) throw e;
      Utilities.sleep(delay);
      delay *= 2;
    }
  }
  return { files: [] };
}


/**
 * ONE TIME SETUP Functions.
 */
function initializeSeatWizard() {
  const ui = SpreadsheetApp.getUi();
  const folderName = 'Seat Wizard';
  const sampleFolderId = '1lIn1Hgg77g4iBNfB7BHTZr9lWy-AKEm5'; // your shared folder ID
  const destFolder = getOrCreateSeatWizardFolder_();
  const drive = DriveApp;
  const srcFolder = drive.getFolderById(sampleFolderId);
  const srcFiles = srcFolder.getFiles();

  let copied = 0;
  while (srcFiles.hasNext()) {
    const file = srcFiles.next();
    const name = file.getName();
    const already = destFolder.getFilesByName(name);
    if (already.hasNext()) continue; // skip duplicates
    file.makeCopy(name, destFolder);
    copied++;
  }

  if (copied > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ Installed ${copied} sample layout template(s) into your Seat Wizard folder.`,
      'Seat Wizard',
      5
    );
  } else {
    Logger.log('Seat Wizard folder already contains sample templates.');
  }
}

//Utility: find or create the Seat Wizard folder.
function getOrCreateSeatWizardFolder_() {
  const folders = DriveApp.getFoldersByName('Seat Wizard');
  return folders.hasNext() ? folders.next() : DriveApp.createFolder('Seat Wizard');
}




