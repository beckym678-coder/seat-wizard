/**
 * =========================================
 *  Seat Wizard • Code.gs
 *  Cleaned & optimized version
 *  Last updated: Nov 16, 2025
 * =========================================
 */

/**
 * Builds the Sheets menu
 */
function onOpen(e) {
  createCustomMenu();
  
  // Check if the user has seen the welcome screen before
  const userProps = PropertiesService.getUserProperties();
  const seenWelcome = userProps.getProperty('seenWelcome') === 'true';
  
  if (!seenWelcome) {
    showWelcomeScreen();
    userProps.setProperty('seenWelcome', 'true');
  }
}


/** Default preferred seats if none saved */
//const PREFERRED_SEAT_RANGE = [1, 2, 3, 4, 5];
function savePreferentialSeats(selectedSeats) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('PREFERRED_SEAT_RANGE', JSON.stringify(selectedSeats));
  return true; // 👈 important so success handler gets called
}

function getPreferentialSeats() {
  const props = PropertiesService.getDocumentProperties();
  const data = props.getProperty('PREFERRED_SEAT_RANGE');
  return data ? JSON.parse(data) : [];
}

function setDefaultSeating() {
  const userProps = PropertiesService.getUserProperties();
  const seating = userProps.getProperty('preferentialSeating');

  // If no seating list exists yet, create the default
  if (!seating) {
    const defaultSeats = [1,2,3,4,5,6,7,8];
    userProps.setProperty('preferentialSeating', JSON.stringify(defaultSeats));
  }
}

function openGenerateChartsSidebar_() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('GenerateChartsSidebar')
      .setTitle("Generate Seating Charts")
  );
}

function openImportSidebar_() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('ImportClassesSidebar')
      .setTitle("Import Google Classroom Rosters")
  );
}

function openRandomizeSidebar_() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput("<h3>Randomize Seats</h3><p>Use the Generate Chart sidebar to randomize seats per class.</p>")
      .setTitle("Randomize Seats")
  );
}

function openPreferentialSidebar_() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('PreferentialSeatSidebar')
      .setTitle("Preferential Seats")
  );
}

function openManualSidebar_() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('UserManualSidebar')
      .setTitle("Seat Wizard Manual")
  );
}

function openRandomGroupsSidebar_() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('RandomGroupsSidebar')
      .setTitle("Random Groups")
  );
}


function buildSeatingToolsSection_(iconUrl) {
  const section = CardService.newCardSection().setHeader("Seating Charts");

  section.addWidget(
    CardService.newDecoratedText()
      .setText("Generate Seating Charts")
      .setIconUrl(iconUrl)
      .setOnClickAction(CardService.newAction().setFunctionName("openGenerateChartsSidebar_"))
  );

  section.addWidget(
    CardService.newDecoratedText()
      .setText("Import / Sync Google Classroom Rosters")
      .setIconUrl(iconUrl)
      .setOnClickAction(CardService.newAction().setFunctionName("openImportSidebar_"))
  );

  section.addWidget(
    CardService.newDecoratedText()
      .setText("Randomize Seats")
      .setIconUrl(iconUrl)
      .setOnClickAction(CardService.newAction().setFunctionName("openRandomizeSidebar_"))
  );

  return section;
}

function buildGroupToolsSection_(iconUrl) {
  const section = CardService.newCardSection().setHeader("Grouping Tools");

  section.addWidget(
    CardService.newDecoratedText()
      .setText("Random Groups Generator")
      .setIconUrl(iconUrl)
      .setOnClickAction(CardService.newAction().setFunctionName("openGroupsCard_"))
  );

  return section;
}

function buildHelpSection_(iconUrl) {
  const section = CardService.newCardSection().setHeader("Help & Settings");

  section.addWidget(
    CardService.newDecoratedText()
      .setText("Preferential Seating Settings")
      .setIconUrl(iconUrl)
      .setOnClickAction(CardService.newAction().setFunctionName("openPreferentialSidebar_"))
  );

  section.addWidget(
    CardService.newDecoratedText()
      .setText("User Manual")
      .setIconUrl(iconUrl)
      .setOnClickAction(CardService.newAction().setFunctionName("openManualSidebar_"))
  );

  return section;
}





/**
 * Runs on installation — builds folder and templates
 * Shows welcome message
 */
function onInstall(e) {
  try {
    initializeSeatWizard();
    onOpen();
    showWelcomeScreen();
  } catch (err) {
    Logger.log("Initialization deferred until UI available: " + err.message);
  }
}





// Toggle tool tips and rebuild menu instantly

function toggleTooltips() {
  const userProps = PropertiesService.getUserProperties();
  const current = userProps.getProperty('tooltipsEnabled') === 'true';
  
  // Toggle the value
  userProps.setProperty('tooltipsEnabled', (!current).toString());
  
  // Optionally notify the user
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Tooltips are now ' + (!current ? 'ON' : 'OFF')
  );
  
  // Refresh the menu so checkmark updates
  createCustomMenu();
}

function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  const userProps = PropertiesService.getUserProperties();
  const tooltipsEnabled = userProps.getProperty('tooltipsEnabled') === 'true';
  
  const menu = ui.createMenu('Seat Wizard');
    // --- Other existing commands ---
  menu.addItem('📋 Import/Sync Classes', 'showImportClassesSidebar');
  menu.addItem('🪑 Generate Seating Charts', 'showGenerateChartSidebar');
  menu.addSeparator();
  menu.addItem('🎯 Preferred Seat Settings', 'showPreferentialSeatSidebar');
  menu.addItem('❓ User Manual', 'showUserManualSidebar')
  
  // Add toggle with checkmark if enabled
  menu.addItem(
    (tooltipsEnabled ? '✅ ' : '') + 'Show Tooltips',
    'toggleTooltips'
  );
  
  menu.addToUi();
}


function showWelcomeScreen() {
  SpreadsheetApp.getUi().alert(
    'Welcome to Seat Wizard!',
    'To start, open extensions -> seat wizard -> Import/Sync rosters',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}





/* --------------------------------------------------------------------------
   SIDEBAR UI
-------------------------------------------------------------------------- */
function showQuickStartSidebar() {
  const html = HtmlService.createTemplateFromFile('QuickStartSidebar');
  html.classSheets = getClassSheets();
  SpreadsheetApp.getUi().showSidebar(
    html.evaluate().setTitle('Quick Start Guide')
  );
}

// Launch Generate Seats sidebar
function showGenerateChartSidebar() {
  const userProps = PropertiesService.getUserProperties();

  const tipsEnabled = userProps.getProperty('tipsEnabled') !== 'false';

  const template = HtmlService.createTemplateFromFile('GenerateChartsSidebar');

  template.showTip = userProps.getProperty('tooltipsEnabled') === 'true';

  const html = template.evaluate().setTitle('Generate Seating Charts');
  SpreadsheetApp.getUi().showSidebar(html);

}

function showRandomGroupsSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('RandomGroupsSidebar')
    .setTitle('Random Groups');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showUserManualSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('UserManualSidebar')
      .setTitle('Seat Wizard Manual')
  );
}

// Launch Import Classes sidebar
function showImportClassesSidebar() {
  const userProps = PropertiesService.getUserProperties();

  const tipsEnabled = userProps.getProperty('tipsEnabled') !== 'false'; 
  const seenTip = userProps.getProperty('seenImportTip') === 'true';

  const template = HtmlService.createTemplateFromFile('ImportClassesSidebar');

  // show tip only if user has tips enabled
  template.showTip = userProps.getProperty('tooltipsEnabled') === 'true';
  const html = template.evaluate().setTitle('Import Classes');
  SpreadsheetApp.getUi().showSidebar(html);

  if (!seenTip) {
    userProps.setProperty('seenImportTip', 'true');
  }
}

function showPreferentialSeatSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('PreferentialSeatSidebar')
      .setTitle('Preferential Seating')
  );
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* --------------------------------------------------------------------------
   SHEET + CLASS MANAGEMENT
-------------------------------------------------------------------------- */

function getClassSheets() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .filter(s => s.getName().startsWith('Class -'))
    .map(s => s.getName());
}

function switchToClassSheet(className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(className);
  if (!sheet) throw new Error(`Class sheet not found: ${className}`);
  ss.setActiveSheet(sheet);
  return true;
}

/* --------------------------------------------------------------------------
   GOOGLE DRIVE RETRIEVAL (Dropdown Lookup)
-------------------------------------------------------------------------- */

function getDropdownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get all class sheets
  const classes = ss.getSheets()
    .map(s => s.getName())
    .filter(name => name.startsWith('Class - '));

  let templates = [];
  let driveError = '';

  try {
    const folderId = getSeatWizardFolderIdSafe(); // your function to get the Seat Wizard folder

    const q = `'${folderId}' in parents and mimeType='application/vnd.google-apps.presentation' and trashed=false`;
    let pageToken;

    do {
      const resp = driveListWithRetry({
        q,
        pageSize: 100,
        pageToken: pageToken || null,
        fields: 'files(id,name,thumbnailLink),nextPageToken' // include thumbnailLink
      });

      const items = (resp && resp.files) || [];
      for (const it of items) {
        if (it.name && it.name.toLowerCase().startsWith('layout -')) {
          templates.push({
            id: it.id,
            name: it.name,
            thumbnailLink: it.thumbnailLink || '' // fallback if thumbnail missing
          });
        }
      }

      pageToken = resp.nextPageToken;
    } while (pageToken);

  } catch (e) {
    driveError = `Template scan failed: ${e.message}`;
    Logger.log('[SeatWizard] getDropdownData ERROR: ' + driveError);
  }

  return { classes, templates, driveError };
}




/* --------------------------------------------------------------------------
   RANDOMIZATION
-------------------------------------------------------------------------- */

function randomizeSeatsForAllClasses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheets = ss.getSheets().filter(s => s.getName().startsWith('Class -'));
  const ui = SpreadsheetApp.getUi();

  if (classSheets.length === 0) {
    ui.alert('No class sheets found to randomize.');
    return;
  }

  let successCount = 0;
  let skipped = [];
  let failed = [];

  ss.toast('🎲 Randomizing all class seat assignments...', 'Seat Wizard', 3);

  classSheets.forEach(sheet => {
    const className = sheet.getName();
    try {
      randomizeSeatsForClass(className);
      successCount++;
      ss.toast(`✅ Randomized ${className}`, 'Seat Wizard', 2);
    } catch (err) {
      const msg = `${className}: ${err.message}`;
      Logger.log(`❌ Error randomizing ${className}: ${err.message}`);

      if (err.message.includes('Too many preferential') || err.message.includes('No preferential seat list')) {
        skipped.push(msg);
      } else {
        failed.push(msg);
      }
    }
  });

  // --- Build result summary ---
  let summary = `✅ Successfully randomized ${successCount} of ${classSheets.length} class(es).`;
  if (skipped.length > 0) summary += `\n⚠️ Skipped (preferential issue):\n• ${skipped.join('\n• ')}`;
  if (failed.length > 0) summary += `\n❌ Failed:\n• ${failed.join('\n• ')}`;

  // --- Toast + Alert for clarity ---
  ss.toast('🎯 Randomization complete.', 'Seat Wizard', 3);
  ui.alert('Seat Wizard Results', summary, ui.ButtonSet.OK);

  Logger.log(summary);
}


function randomizeSeatsForClass(className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(className);
  if (!sheet) throw new Error(`Class sheet not found: ${className}`);

  const props = PropertiesService.getDocumentProperties();
  const stored = props.getProperty('PREFERRED_SEAT_RANGE');
  const preferredSeats = stored ? JSON.parse(stored).map(n => Number(n)) : [];

  if (preferredSeats.length === 0) {
    throw new Error("No preferential seat list found. Please set them first in 'Set Preferential Seats'.");
  }

  // A=Display, B=Student, C=Seat, D=Pref(Y), E=KA1, F=KA2
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  if (!data.length) throw new Error(`No students found in ${className}`);

  const keepAwayMap = {};
  for (const row of data) {
    const stud = row[1];
    const ka1 = row[4], ka2 = row[5];
    keepAwayMap[stud] = [ka1, ka2].filter(Boolean);
  }

  const prefStudents = data.filter(r => r[3]?.toString().toUpperCase() === 'Y');
  const regStudents  = data.filter(r => !r[3] || r[3].toString().toUpperCase() !== 'Y');

  if (prefStudents.length > preferredSeats.length) {
    throw new Error(
      `Too many preferential students (${prefStudents.length}) for available seats (${preferredSeats.length}).`
    );
  }

  const allSeatNums = Array.from({ length: data.length }, (_, i) => i + 1);
  const usedSeats = new Set();

  // Helper: seat adjacency for keep-aways
  const cols = Math.ceil(Math.sqrt(data.length));
  const seatToGrid = (seat) => [Math.floor((seat - 1) / cols), (seat - 1) % cols];
  const isConflict = (studName, seat, assignments) => {
    const [r, c] = seatToGrid(seat);
    const conflicts = keepAwayMap[studName] || [];
    return assignments.some(a => {
      if (!conflicts.includes(a.stud)) return false;
      const [ar, ac] = seatToGrid(a.seat);
      return Math.abs(r - ar) <= 1 && Math.abs(c - ac) <= 1;
    });
  };

  const assignments = [];

  // 1️⃣ Assign all preferential students FIRST — guaranteed seats
  shuffleArray(prefStudents);
  const shuffledPrefSeats = [...preferredSeats];
  shuffleArray(shuffledPrefSeats);

  prefStudents.forEach((s, i) => {
    const chosenSeat = shuffledPrefSeats[i];
    assignments.push({ stud: s[1], display: s[0], seat: chosenSeat, pref: s[3], ka1: s[4], ka2: s[5] });
    usedSeats.add(chosenSeat);
  });

  // 2️⃣ Assign remaining students
  const remainingSeats = allSeatNums.filter(s => !usedSeats.has(s));
  shuffleArray(remainingSeats);

  for (const s of regStudents) {
    let chosen = null;
    for (const seat of remainingSeats) {
      if (usedSeats.has(seat)) continue;
      if (isConflict(s[1], seat, assignments)) continue;
      chosen = seat;
      break;
    }
    if (!chosen) chosen = remainingSeats.find(seat => !usedSeats.has(seat));
    if (!chosen) throw new Error(`No available seat for ${s[1]}`);
    usedSeats.add(chosen);
    assignments.push({ stud: s[1], display: s[0], seat: chosen, pref: s[3], ka1: s[4], ka2: s[5] });
  }

  // 3️⃣ Write updated seat assignments
  const sorted = assignments.sort((a, b) => a.seat - b.seat)
    .map(x => [x.display, x.stud, x.seat, x.pref, x.ka1, x.ka2]);
  sheet.getRange(2, 1, sorted.length, 6).setValues(sorted);

  ss.toast(`✅ Seats randomized for ${className}`, 'Seat Wizard', 4);
}


function shuffleArray(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
}

/*
function getLayoutInfo_() {
  return { rows: 6, cols: 6 }; // overrides keep-away geometry logic
}*/

/* --------------------------------------------------------------------------
   GOOGLE CLASSROOM SYNC
-------------------------------------------------------------------------- */

function getClassroomCourses() {
  try {
    const response = Classroom.Courses.list({ courseStates: ['ACTIVE'] });
    const courses = response.courses || [];
    return courses.map(course => ({ id: course.id, name: course.name }));
  } catch (error) {
    Logger.log('Error fetching courses: ' + error.message);
    throw new Error('Unable to fetch Google Classroom courses. Make sure the Classroom API is enabled.');
  }
}

function importSelectedCourses(courseIds) {
  const imported = [];
  courseIds.forEach(id => {
    try {
      const course = Classroom.Courses.get(id);
      const studs = Classroom.Courses.Students.list(id).students || [];
      const names = studs.map(s => s.profile.name.fullName);
      importOrSyncClassRoster(course.name, names);
      randomizeSeatsForClass(`Class - ${course.name}`);
      imported.push(course.name);
    } catch (e) {
      Logger.log(`Error importing ${id}: ${e.message}`);
    }
  });
  return { success: true, count: imported.length };
}

function importOrSyncClassRoster(courseName, studentNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `Class - ${courseName}`;
  let sheet = ss.getSheetByName(sheetName);

  const headers = ['Display Name', 'Student Name', 'Seat Number', 'Preferential Seating', 'Keep Away 1', 'Keep Away 2'];

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, 6).setValues([headers]);
    sheet.getRange(2, 1, studentNames.length, 6)
      .setValues(studentNames.map((n, i) => [n, n, i + 1, '', '', '']));
    setupKeepAwayDropdowns(sheet);
    return;
  }

  const pos = getHeaderPositions_(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  const map = new Map(values.map(row => {
    const stud = row[pos.studentCol - 1];
    return [stud, {
      display: row[pos.displayCol - 1],
      seat: row[pos.seatCol - 1],
      pref: row[pos.prefCol - 1],
      ka1: row[pos.ka1Col - 1],
      ka2: row[pos.ka2Col - 1]
    }];
  }));

  const merged = studentNames.map((stud, i) => {
    const ex = map.get(stud);
    return [
      ex?.display || stud,
      stud,
      ex?.seat || (i + 1),
      ex?.pref || '',
      ex?.ka1 || '',
      ex?.ka2 || ''
    ];
  });

  sheet.clear();
  sheet.getRange(1, 1, 1, 6).setValues([headers]);
  sheet.getRange(2, 1, merged.length, 6).setValues(merged);
  setupKeepAwayDropdowns(sheet);
}

function getHeaderPositions_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(v => v.toString().trim().toLowerCase());

  const find = t => headers.findIndex(h => h.includes(t)) + 1;
  return {
    displayCol: find('display'),
    studentCol: find('student'),
    seatCol:    find('seat'),
    prefCol:    find('preferential'),
    ka1Col:     find('keep away 1'),
    ka2Col:     find('keep away 2')
  };
}

function setupKeepAwayDropdowns(sheet) {
  const rows = sheet.getLastRow() - 1;
  if (rows < 1) return;
  const names = sheet.getRange(2, 2, rows, 1).getValues().flat().filter(Boolean);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(names, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 5, rows, 2).setDataValidation(rule);
}

/* --------------------------------------------------------------------------
   SLIDE GENERATION (MAIN FEATURE)
-------------------------------------------------------------------------- */

/*
  // === Create title slide ===
  const classTitle = className.replace(/^Class - /, '');
  const titleSlide = pres.appendSlide(SlidesApp.PredefinedLayout.TITLE);
  const titleElements = titleSlide.getPageElements();
  if (titleElements.length > 0) titleElements[0].asShape().getText().setText(classTitle);
  if (titleElements.length > 1) { try { titleElements[1].remove(); } catch (_) {} 
*/

function generateSeatingChartFromSlideTemplate(className, templateId, presentationOrId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheet = ss.getSheetByName(className);
  if (!classSheet) throw new Error(`Class sheet not found: ${className}`);

  const headers = classSheet.getRange(1, 1, 1, classSheet.getLastColumn())
    .getValues()[0]
    .map(h => h.toString().trim().toLowerCase());

  const displayCol = headers.findIndex(h => h.includes('display')) + 1;
  const seatCol = headers.findIndex(h => h.includes('seat')) + 1;

  if (!displayCol || !seatCol) {
    throw new Error(`Missing expected columns in "${className}". Found headers: ${headers.join(', ')}`);
  }

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

  const templatePres = SlidesApp.openById(templateId);
  const templateSlides = templatePres.getSlides();
  if (!templateSlides.length) throw new Error('Template has no slides.');

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");

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

    const title = `Seating Chart - ${className} (${timestamp})`;
    const created = SlidesApp.create(title);
    pres = created;

    const file = DriveApp.getFileById(created.getId());
    folder.addFile(file);
    try { DriveApp.getRootFolder().removeFile(file); } catch (_) {}

    createdNew = true;
  }

  if (createdNew) {
    const initialSlides = pres.getSlides();
    if (initialSlides.length > 0) {
      try { initialSlides[0].remove(); } catch (_) {}
    }
  }

  const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM dd");

  // === Populate template slides ===
  templateSlides.forEach((tplSlide, i) => {
    const slide = pres.appendSlide(tplSlide);

    slide.getShapes().forEach(shape => {
      if (!shape.getText) return;
      let content;
      try { content = shape.getText().asString().trim(); } catch (_) { return; }

      const tr = shape.getText();

      // === Replace seat numbers with student names ===
      if (/^\d+$/.test(content)) {
        const student = seatMap.get(content);
        if (student) {
          const oldStyle = tr.getTextStyle();
          tr.setText(student);
          reapplyStyles(tr, oldStyle);
          autoFitName_(tr, oldStyle);
        } else {
          tr.setText('');
        }
      } else {
        // === Replace placeholders for Class Name / Date ===
        const textLower = content.toLowerCase();
        if (textLower.includes('class name') || textLower.includes('classname') || textLower.includes('class_name')) {
          tr.setText(className);
        } else if (textLower.includes('date')) {
          tr.setText("Created on:" + formattedDate);
        }
      }
    });
/*
    // Label Student / Teacher View if 2-slide template
    if (templateSlides.length === 2) {
      const label = i === 0 ? ' (Student View)' : ' (Teacher View)';
      slide.insertTextBox(className + label)
        .setLeft(20).setTop(20).setWidth(300);
    }*/
  });

  Logger.log(`✅ Seating chart generated for ${className} (${timestamp})`);
  return pres.getUrl();
}


  // === Helper — reapply styles ===
  function reapplyStyles(textRange, oldStyle) {
    if (!oldStyle) return;
    const newStyle = textRange.getTextStyle();
    try { newStyle.setFontFamily(oldStyle.getFontFamily()); } catch (_) {}
    try { newStyle.setFontSize(oldStyle.getFontSize()); } catch (_) {}
    try { newStyle.setForegroundColor(oldStyle.getForegroundColor()); } catch (_) {}
    try { newStyle.setBold(oldStyle.isBold()); } catch (_) {}
    try { newStyle.setItalic(oldStyle.isItalic()); } catch (_) {}
    try { newStyle.setUnderline(oldStyle.isUnderline()); } catch (_) {}
  }


  // === Helper — shrink name if needed ✅ ===
  function autoFitName_(textRange, oldStyle) {
    const name = textRange.asString().trim();
    if (!name) return;

    const origSize = oldStyle.getFontSize() || 12;
    let newSize = origSize;

    if (name.length > 10) newSize = Math.max(8, Math.round(origSize * 0.70));
    else if (name.length > 8) newSize = Math.max(9, Math.round(origSize * 0.80));
    else if (name.length > 6) newSize = Math.max(10, Math.round(origSize * 0.90));

    try {
      textRange.getTextStyle().setFontSize(newSize);
    } catch (_) {}
  }

/*
function autoFitName_(textRange, oldStyle) {
  const name = textRange.asString().trim();
  if (!name) return;

  const origSize = oldStyle.getFontSize() || 12;
  let newSize = origSize;

  if (name.length > 25) newSize = Math.max(8, Math.round(origSize * 0.70));
  else if (name.length > 18) newSize = Math.max(9, Math.round(origSize * 0.80));
  else if (name.length > 12) newSize = Math.max(10, Math.round(origSize * 0.90));

  try {
    textRange.getTextStyle().setFontSize(newSize);
  } catch (_) {}
}*/


function generateAllSeatingCharts(templateId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");
  const pres = SlidesApp.create(`Seating Charts (${timestamp})`);
  pres.getSlides()[0].remove();

  const errors = [];

  getClassSheets().forEach(name => {
    try {
      generateSeatingChartFromSlideTemplate(name, templateId, pres);
    } catch (e) {
      errors.push(`${name}: ${e.message}`);
    }
  });

  return {
    success: true,
    message: `Generated with ${errors.length} error(s)`,
    url: pres.getUrl(),
    errors
  };
}

/* --------------------------------------------------------------------------
   DRIVE UTILITIES
-------------------------------------------------------------------------- */

function getSeatWizardFolderIdSafe() {
  const found = driveListWithRetry({
    q: "mimeType='application/vnd.google-apps.folder' and name='Seat Wizard' and trashed=false",
    fields: 'files(id,name)'
  });
  if (found.files?.length) return found.files[0].id;

  return Drive.Files.create({
    mimeType: 'application/vnd.google-apps.folder',
    name: 'Seat Wizard'
  }).id;
}

function driveListWithRetry(params, attempts = 3) {
  for (let i = 0; i < attempts; i++) {
    try {
      return Drive.Files.list(params);
    } catch (e) {
      Utilities.sleep(200 * (i + 1));
    }
  }
  throw new Error('Drive API failed repeatedly');
}

/* --------------------------------------------------------------------------
   INITIALIZATION: Copy sample template files once
-------------------------------------------------------------------------- */

function initializeSeatWizard() {
  const src = DriveApp.getFolderById('1lIn1Hgg77g4iBNfB7BHTZr9lWy-AKEm5'); // 🪄 your folder
  const dest = DriveApp.getFolderById(getSeatWizardFolderIdSafe());

  let copied = 0;
  const files = src.getFiles();

  while (files.hasNext()) {
    const f = files.next();
    if (!dest.getFilesByName(f.getName()).hasNext()) {
      f.makeCopy(f.getName(), dest);
      copied++;
    }
  }

  if (copied) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`✅ Installed ${copied} sample templates!`, 'Seat Wizard', 5);
  }
}

/* --------------------------------------------------------------------------
   RANDOM GROUP MAKER: functions that support generating groupings
-------------------------------------------------------------------------- */
/**
 * Generate groups for the *current* class (active "Class - ..." sheet).
 * Creates ONE slide for that class, with multiple rows of groups as needed.
 */
function generateRandomGroups(mode, count, size, groupNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();

  // Use active class sheet if possible; otherwise fallback to first class sheet
  if (!sheet || !sheet.getName().startsWith('Class - ')) {
    const classSheets = ss.getSheets().filter(s => s.getName().startsWith('Class - '));
    if (!classSheets.length) {
      throw new Error('No class sheets found. Please select a "Class - ..." sheet.');
    }
    sheet = classSheets[0];
  }

  const className = sheet.getName().replace(/^Class - /, '');

  // Build student + keep-away data from this sheet
  const { studentIds, displayById, keepAwayMap } = getGroupingDataFromSheet_(sheet);
  if (!studentIds.length) {
    throw new Error(`No students found in ${sheet.getName()}.`);
  }

  // Compute groups respecting keep-away + no groups of 1
  const groupsIds = makeGroupsForStudents_(studentIds, keepAwayMap, mode, count, size);

  // Map to display names and assign group labels
  const groups = groupsIds.map((members, i) => ({
    name: groupNames[i] || `Group ${i + 1}`,
    members: members.map(id => displayById[id] || id),
  }));

  const folder = getOrCreateSeatWizardFolder_();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');

  const pres = SlidesApp.create(`Group Assignments - ${className} (${timestamp})`);
  const file = DriveApp.getFileById(pres.getId());
  folder.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (_) {}

  // Remove default title slide
  try { pres.getSlides()[0].remove(); } catch (_) {}

  // Create a single slide for this class
  const safeGroups = sanitizeGroups_(groups);
  buildGroupsSlideForClass_(pres, classTitle, safeGroups);

  return pres.getUrl();
}

/**
 * Generate group slides for ALL "Class - ..." sheets.
 * Each class gets its own slide in one master presentation.
 */
function generateRandomGroupsForAllClasses(mode, count, size, groupNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheets = ss.getSheets().filter(s => s.getName().startsWith("Class -"));
  if (classSheets.length === 0) throw new Error("No class sheets found.");

  const folder = getOrCreateSeatWizardFolder_();
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const pres = SlidesApp.create(`Group Assignments - ${timestamp}`);

  classSheets.forEach(sheet => {
    const className = sheet.getName().replace(/^Class - /, "");
    const students = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
      .flat()
      .filter(Boolean);

    const groups = makeGroupsInternal_(students, mode, count, size, groupNames);
    buildGroupsSlideForClass_(pres, className, groups);
  });

  const fileId = pres.getId();
  const file = DriveApp.getFileById(fileId);
  folder.addFile(file);

  return pres.getUrl();
}



function generateRandomGroups_ForOneClass_Internal(sheet, pres, mode, groupCount, groupSize, customNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const className = sheet.getName();

  // Read students
  const data = sheet.getDataRange().getValues().slice(1);
  const students = data
    .filter(r => r[1])
    .map(r => ({
      display: r[0] || r[1],
      student: r[1],
      ka1: r[4],
      ka2: r[5]
    }));

  if (students.length === 0) throw new Error(`No students found in ${className}`);

  // Shuffle
  for (let i = students.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [students[i], students[j]] = [students[j], students[i]];
  }

  // Determine groups
  let groups = [];
  if (mode === "count") {
    const n = Math.max(1, groupCount);
    groups = Array.from({ length: n }, () => []);
    students.forEach((s, i) => groups[i % n].push(s));
  } else {
    const size = Math.max(2, groupSize);
    for (let i = 0; i < students.length; i += size) {
      groups.push(students.slice(i, i + size));
    }
  }

  groups = fixKeepAwayConflicts_(groups);

  const groupNames = customNames || groups.map((_, i) => `Group ${i + 1}`);

  // Create presentation if needed
  let presObj = pres;
  if (!presObj) {
    const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");
    presObj = SlidesApp.create(`Group Assignments - ${className} (${time})`);

    const folder = getOrCreateSeatWizardFolder_();
    const file = DriveApp.getFileById(presObj.getId());
    folder.addFile(file);
    try { DriveApp.getRootFolder().removeFile(file); } catch (_) {}

    // Remove default slide
    try { presObj.getSlides()[0].remove(); } catch (_) {}
  }

  // === One slide per class ===
  const slide = presObj.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  const pageW = presObj.getPageWidth();
  const margin = 40;
  const colWidth = (pageW - margin * 2) / groups.length;

  slide.insertTextBox(className.replace(/^Class - /, ''))
    .setLeft(margin)
    .setTop(20)
    .setWidth(pageW - margin * 2)
    .getText().setText(className.replace(/^Class - /, ''))
    .getTextStyle().setFontSize(20).setBold(true);

  groups.forEach((g, i) => {
    const x = margin + i * colWidth;

    const titleBox = slide.insertTextBox(groupNames[i])
      .setLeft(x)
      .setTop(70)
      .setWidth(colWidth - 10);
    titleBox.getText().getTextStyle().setBold(true).setFontSize(14);

    const memberText = g.map(s => `• ${s.display}`).join("\n");
    const memberBox = slide.insertTextBox(memberText)
      .setLeft(x)
      .setTop(100)
      .setWidth(colWidth - 10);
    memberBox.getText().getTextStyle().setFontSize(10);
  });

  return presObj.getUrl();
}


function generateRandomGroupsFromSidebar(mode, groupCount, groupSize, groupNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const className = activeSheet.getName();

  if (!className.startsWith("Class - ")) {
    throw new Error("Please select a class sheet before generating groups.");
  }

  return generateRandomGroups_ForOneClass_Internal(
    activeSheet,
    null,           // no existing presentation
    mode,
    groupCount,
    groupSize,
    groupNames      // <- pass custom names through
  );
}


function fixKeepAwayConflicts_(groups) {
  // Very simple resolver:
  // If a student is in a group with someone in their KA list,
  // move them to next group with capacity.
  const maxSize = Math.max(...groups.map(g => g.length));

  groups.forEach((group, gi) => {
    group.forEach((s, si) => {
      const ka = [s.ka1, s.ka2].filter(Boolean);
      if (!ka.length) return;

      // Check conflict
      const conflict = group.some(other => other !== s && ka.includes(other.student));
      if (!conflict) return;

      // Move student
      for (let g2 = 0; g2 < groups.length; g2++) {
        if (g2 === gi) continue;
        if (groups[g2].length < maxSize) {
          groups[g2].push(s);
          group.splice(si, 1);
          return;
        }
      }
    });
  });

  return groups;
}


function getOrCreateSeatWizardFolder_() {
  const FOLDER_NAME = 'Seat Wizard';
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  // Create the folder if it doesn't exist yet
  return DriveApp.createFolder(FOLDER_NAME);
}

/**
 * Extracts:
 * - studentIds: array of Student Name (ID basis)
 * - displayById: map Student Name -> Display Name
 * - keepAwayMap: studentId -> [studentIds to avoid]
 */
function getGroupingDataFromSheet_(sheet) {
  const pos = getHeaderPositions_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) {
    return { studentIds: [], displayById: {}, keepAwayMap: {} };
  }

  const numRows = lastRow - 1;
  const values = sheet.getRange(2, 1, numRows, lastCol).getValues();

  const studentIds = [];
  const displayById = {};
  const keepAwayMap = {};

  values.forEach(row => {
    const display = pos.displayCol ? row[pos.displayCol - 1] : row[0];
    const student = pos.studentCol ? row[pos.studentCol - 1] : row[1];
    const ka1 = pos.ka1Col ? row[pos.ka1Col - 1] : '';
    const ka2 = pos.ka2Col ? row[pos.ka2Col - 1] : '';

    if (!student) return;

    const id = String(student).trim();
    studentIds.push(id);
    displayById[id] = display || id;

    const ka = [];
    if (ka1) ka.push(String(ka1).trim());
    if (ka2) ka.push(String(ka2).trim());
    keepAwayMap[id] = ka;
  });

  return { studentIds, displayById, keepAwayMap };
}



/**
 * Core grouping logic:
 * - mode = "count" → try to create that many groups
 * - mode = "size"  → enforce max people per group (no hard cap if merging needed)
 * - never leaves a group of size 1 if it can avoid it
 * - avoids putting keep-away students in the same group (both directions)
 */
function makeGroupsForStudents_(studentIds, keepAwayMap, mode, count, size) {
  const students = [...studentIds];
  shuffleArray(students); // you already have shuffleArray elsewhere

  const n = students.length;
  let groupCount;
  let maxSize;

  if (mode === 'size') {
    maxSize = Math.max(2, parseInt(size, 10) || 4);
    groupCount = Math.ceil(n / maxSize);
  } else {
    groupCount = Math.max(1, parseInt(count, 10) || 3);
    maxSize = Math.ceil(n / groupCount) + 1; // soft limit
  }

  const groups = Array.from({ length: groupCount }, () => []);

  function conflict(a, b) {
    const aList = keepAwayMap[a] || [];
    const bList = keepAwayMap[b] || [];
    return aList.includes(b) || bList.includes(a);
  }

  function canJoin(group, student, hardCap) {
    if (hardCap && group.length >= maxSize) return false;
    return group.every(m => !conflict(m, student));
  }

  // First pass: greedy placement respecting keep-aways and size
  students.forEach(s => {
    let placed = false;

    // Try groups ordered by current size (smallest first)
    const indices = groups.map((g, i) => i).sort((a, b) => groups[a].length - groups[b].length);

    for (const idx of indices) {
      if (canJoin(groups[idx], s, true)) {
        groups[idx].push(s);
        placed = true;
        break;
      }
    }

    // If no hard-cap-respecting fit, allow slightly over maxSize
    if (!placed) {
      for (const idx of indices) {
        if (canJoin(groups[idx], s, false)) {
          groups[idx].push(s);
          placed = true;
          break;
        }
      }
    }

    // Absolute fallback: ignore keep-away (should be rare)
    if (!placed) {
      const idx = indices[0];
      groups[idx].push(s);
    }
  });

  // Fix groups of size 1 by merging them into neighbor groups if possible
  for (let i = groups.length - 1; i >= 0; i--) {
    if (groups[i].length === 1) {
      const student = groups[i][0];
      let merged = false;

      for (let j = 0; j < groups.length; j++) {
        if (j === i) continue;
        if (canJoin(groups[j], student, false)) {
          groups[j].push(student);
          groups.splice(i, 1);
          merged = true;
          break;
        }
      }
      // if not merged, we leave a solo group as last resort
    }
  }

  return groups.filter(g => g.length > 0);
}

function sanitizeGroups_(groups) {
  return (groups || [])
    .filter(g => g && Array.isArray(g.members))        // must exist and have members array
    .map((g, i) => ({
      name: g.name ? String(g.name) : `Group ${i + 1}`, // ensure non-null string
      members: g.members.filter(Boolean)                // remove null/undefined/empty
    }))
    .filter(g => g.members.length > 0);                 // skip empty groups
}


/**
 * Creates one slide with ALL groups for a class.
 * - Up to 5 groups per row, minimizing number of rows.
 * - Skips empty groups.
 * - Never passes null to setText().
 */
function buildGroupsSlideForClass_(pres, classTitle, groups) {

  // --- fully sanitize incoming groups first ---
  groups = (groups || [])
    .filter(g => g && Array.isArray(g.members))
    .map((g, i) => ({
      name: (g.name ? String(g.name).trim() : `Group ${i + 1}`),
      members: g.members.filter(Boolean)
    }))
    .filter(g => g.members.length > 0);

  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const pageWidth = pres.getPageWidth();
  const pageHeight = pres.getPageHeight();

  const marginX = 30;
  const marginY = 40;

  // --- Title ---
  const titleShape = slide.insertTextBox(classTitle);
  titleShape
    .setLeft(marginX)
    .setTop(10)
    .setWidth(pageWidth - 2 * marginX)
    .setHeight(40);

  const titleText = titleShape.getText();
  titleText.setText(classTitle || "");
  titleText.getTextStyle()
    .setFontFamily('Arial')
    .setBold(true)
    .setFontSize(30);
  titleText.getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const titleBottom = titleShape.getTop() + titleShape.getHeight();

  if (groups.length === 0) return;

  // --- GRID CALC ---
  const maxCols = 5;
  const gCount = groups.length;
  const numCols = Math.min(maxCols, gCount);
  const numRows = Math.ceil(gCount / maxCols);

  const availHeight = pageHeight - titleBottom - marginY;
  const groupWidth = (pageWidth - (numCols + 1) * marginX) / numCols;
  const groupHeight = (availHeight - (numRows + 1) * marginY) / numRows;

  // --- Auto-scaler ---
  function computeFontSize(memberCount) {
    if (memberCount <= 3) return 26;
    if (memberCount <= 6) return 22;
    if (memberCount <= 10) return 18;
    if (memberCount <= 14) return 16;
    return 14;
  }

  // --- Place groups ---
  groups.forEach((group, index) => {
    const row = Math.floor(index / maxCols);
    const col = index % maxCols;

    const x = marginX + col * (groupWidth + marginX);
    const y = titleBottom + marginY + row * (groupHeight + marginY);

    const box = slide.insertTextBox('');
    box.setLeft(x).setTop(y).setWidth(groupWidth).setHeight(groupHeight);

    const tr = box.getText();

    const safeName = group.name || `Group ${index + 1}`;
    const memberLines = group.members.map(m => `• ${m}`);
    const finalText = [safeName, ...memberLines].join("\n");

    tr.setText(finalText);

    const fs = computeFontSize(memberLines.length);
    tr.getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(fs)
      .setBold(false);

    // Bold first line
    tr.getRange(0, safeName.length).getTextStyle().setBold(true);

    tr.getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.LEFT);
  });
}

/**
 * Internal helper to build groups from a list of students.
 *
 * @param {string[]} students   List of student names.
 * @param {string}   mode       "count" for number-of-groups, "size" for max-group-size.
 * @param {string}   countRaw   User input for number of groups (may be empty/string).
 * @param {string}   sizeRaw    User input for max group size (may be empty/string).
 * @param {string[]} groupNames Optional group names (one per group).
 *
 * @return {{name:string, members:string[]}[]} Array of groups.
 */
function makeGroupsInternal_(students, mode, countRaw, sizeRaw, groupNames) {
  const total = students.length;
  if (total === 0) {
    return [];
  }

  // Make a shuffled copy of the roster
  const pool = students.slice();
  if (typeof shuffleArray === 'function') {
    shuffleArray(pool);
  } else {
    // Fallback local shuffle if needed
    for (let i = pool.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [pool[i], pool[j]] = [pool[j], pool[i]];
    }
  }

  mode = mode === 'size' ? 'size' : 'count';

  let numGroups;

  if (mode === 'size') {
    // "Max people per group" mode
    let maxSize = parseInt(sizeRaw, 10);
    if (!maxSize || maxSize < 2) maxSize = 2;           // at least pairs
    numGroups = Math.ceil(total / maxSize);
    if (numGroups < 1) numGroups = 1;
  } else {
    // "Number of groups" mode (default)
    numGroups = parseInt(countRaw, 10);
    if (!numGroups || numGroups < 1) numGroups = 1;
    if (numGroups > total) numGroups = total;           // can't have more groups than students
  }

  const groupsMembers = splitIntoGroups_(pool, numGroups);

  // Apply group names (optional)
  const result = groupsMembers.map((members, i) => {
    const label = (groupNames && groupNames[i]) ? groupNames[i] : `Group ${i + 1}`;
    return { name: label, members: members };
  });

  return result;
}

/**
 * Split a shuffled list of students into N groups, fairly balanced.
 * Tries to avoid groups of size 1 by borrowing from larger groups.
 *
 * @param {string[]} students
 * @param {number}   numGroups
 * @return {string[][]} Array of groups (each is an array of student names).
 */
function splitIntoGroups_(students, numGroups) {
  numGroups = Math.max(1, Math.min(numGroups, students.length));
  const groups = Array.from({ length: numGroups }, () => []);

  // Simple round-robin distribution for balance
  students.forEach((name, idx) => {
    groups[idx % numGroups].push(name);
  });

  // Fix any groups of size 1 by borrowing from groups with size > 2
  for (let i = 0; i < groups.length; i++) {
    if (groups[i].length === 1 && students.length > 1) {
      const donorIndex = groups.findIndex(
        (g, idx) => g.length > 2 && idx !== i
      );
      if (donorIndex !== -1) {
        const moved = groups[donorIndex].pop();
        groups[i].push(moved);
      }
    }
  }

  return groups;
}


function openGroupsCard_() {
  const iconUrl = "https://raw.githubusercontent.com/beckym678-coder/seat-wizard/main/Seat_Wizard_App_Icons/A_flat-style_digital_illustration_icon_features_a_.png";

  return CardService.newCardBuilder()
    .setHeader(
      CardService.newCardHeader()
        .setTitle("Random Groups")
        .setSubtitle("Choose your grouping method")
        .setImageUrl(iconUrl)
    )
    .addSection(
      CardService.newCardSection()
        .addWidget(
          CardService.newDecoratedText()
            .setText("<b>Create Random Groups</b><br>Open the random group generator sidebar.")
            .setIconUrl(iconUrl)
            .setOnClickAction(CardService.newAction().setFunctionName("openRandomGroupsSidebar_"))
        )
    )
    .build();
}

/**Preview google slides */
function generateSlidePreview(slideFileId) {
  const url = `https://slides.googleapis.com/v1/presentations/${slideFileId}/pages/slideId/presentationThumbnail?thumbnailProperties.thumbnailSize=LARGE&mimeType=PNG`;

  const params = {
    method: "GET",
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, params);
  const data = JSON.parse(response.getContentText());

  // Returns a base64 PNG
  return data.contentUrl;
}


function getSlideThumbnailUrl(presentationId, slideId) {
  const url = `https://slides.googleapis.com/v1/presentations/${presentationId}/pages/${slideId}/thumbnail?thumbnailProperties.thumbnailSize=LARGE&mimeType=PNG`;

  const params = {
    method: "get",
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, params);
  const data = JSON.parse(response.getContentText());

  return data.contentUrl || '';
}


function getFirstSlideId(presentationId) {
  const url = `https://slides.googleapis.com/v1/presentations/${presentationId}`;
  const params = {
    method: "GET",
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }
  };

  const response = UrlFetchApp.fetch(url, params);
  const presentation = JSON.parse(response.getContentText());

  return presentation.slides[0].objectId; // first slide
}


function generatePreviewForTemplate(templateId) {
  const firstSlideId = getFirstSlideId(templateId);
  return getSlideThumbnail(templateId, firstSlideId);
}

function getSlideThumbnail(presentationId, slideId) {
  const url = `https://slides.googleapis.com/v1/presentations/${presentationId}/pages/${slideId}/thumbnail?thumbnailProperties.thumbnailSize=LARGE&mimeType=PNG`;

  const params = {
    method: "GET",
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    }
  };

  const response = UrlFetchApp.fetch(url, params);
  const data = JSON.parse(response.getContentText());

  // data.contentUrl = public thumbnail URL (expires after some time)
  return data.contentUrl;
}

function getTemplatePreviewUrl(templateId) {
  try {
    const slideId = getFirstSlideId(templateId);
    return getSlideThumbnailUrl(templateId, slideId);
  } catch (e) {
    Logger.log('Preview generation failed for ' + templateId + ': ' + e);
    return ''; // fallback gracefully
  }
}


function savePreviewImage(contentUrl, name, folderId) {
  const blob = UrlFetchApp.fetch(contentUrl).getBlob();
  blob.setName(name + ".png");
  return DriveApp.getFolderById(folderId).createFile(blob).getId();
}


function saveLastSelection(className, templateId) {
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty('LAST_CLASS', className || '');
  userProps.setProperty('LAST_TEMPLATE', templateId || '');
}

function getLastSelection() {
  const userProps = PropertiesService.getUserProperties();
  return {
    className: userProps.getProperty('LAST_CLASS') || '',
    templateId: userProps.getProperty('LAST_TEMPLATE') || ''
  };
}

/**
 * Returns a thumbnail URL for a Google Slides presentation.
 * @param {string} templateId The Drive ID of the Slides template.
 * @return {string} A URL of the thumbnail image.
 */
function getLayoutPreviewUrl(templateId) {
  try {
    // Use Drive API to get the thumbnail link
    const file = Drive.Files.get(templateId, { fields: 'thumbnailLink' });
    
    if (file && file.thumbnailLink) {
      // Google Drive returns a link with "=s220" at the end — optional: increase size
      return file.thumbnailLink.replace(/=s\d+$/, '=s320'); 
    } else {
      throw new Error('No thumbnail available for this template.');
    }
  } catch (e) {
    Logger.log('[SeatWizard] getLayoutPreviewUrl ERROR: ' + e.message);
    throw e;
  }
}



