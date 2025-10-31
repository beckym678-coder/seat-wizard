/**
 * =========================================
 *  Seat Wizard • Code.gs
 *  Cleaned & optimized version
 *  Last updated: Oct 29, 2025
 * =========================================
 */

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

/**
 * Runs on installation — builds folder and templates
 */
function onInstall(e) {
  try {
    initializeSeatWizard();
  } catch (err) {
    Logger.log("Initialization deferred until UI available: " + err.message);
  }
}

/**
 * Builds the Sheets menu
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Seat Wizard')
    .addItem('Generate Seating Charts', 'showGenerateChartSidebar')
    .addSeparator()
    .addItem('Seat Wizard Help', 'showUserManualSidebar')
    .addItem('Import/Sync Google Classroom Rosters', 'showImportClassesSidebar')
    .addItem('Set Preferential Seats', 'showPreferentialSeatSidebar')
    .addToUi();
}

/* --------------------------------------------------------------------------
   SIDEBAR UI
-------------------------------------------------------------------------- */

function showGenerateChartSidebar() {
  const html = HtmlService.createTemplateFromFile('GenerateChartSidebar');
  html.classSheets = getClassSheets();
  SpreadsheetApp.getUi().showSidebar(
    html.evaluate().setTitle('Generate Seating Charts')
  );
}

function showUserManualSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('UserManualSidebar')
      .setTitle('Seat Wizard Manual')
  );
}

function showImportClassesSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('ImportClassesSidebar')
      .setTitle('Import Class Rosters')
  );
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

  const classes = ss.getSheets()
    .map(s => s.getName())
    .filter(name => name.startsWith('Class - '));

  let templates = [];
  let driveError = '';

  try {
    const folderId = getSeatWizardFolderIdSafe();

    const q = `'${folderId}' in parents and mimeType='application/vnd.google-apps.presentation' and trashed=false`;
    let pageToken;

    do {
      const resp = driveListWithRetry({
        q,
        pageSize: 100,
        pageToken: pageToken || null,
        fields: 'files(id,name),nextPageToken'
      });

      const items = (resp && resp.files) || [];
      for (const it of items) {
        if (it.name &&
            it.name.toLowerCase().startsWith('layout -')) {
          templates.push({ id: it.id, name: it.name });
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

function getLayoutInfo_() {
  return { rows: 6, cols: 6 }; // overrides keep-away geometry logic
}

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
  if (!templateSlides.length) throw new Error('Template has no slides.');

  // === Create timestamp for filename ===
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");

  // === Create or open the target presentation ===
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

  // === Remove default slide for new files ONLY ===
  if (createdNew) {
    const initialSlides = pres.getSlides();
    if (initialSlides.length > 0) {
      try { initialSlides[0].remove(); } catch (_) {}
    }
  }

  // === Create title slide ===
  const classTitle = className.replace(/^Class - /, '');
  const titleSlide = pres.appendSlide(SlidesApp.PredefinedLayout.TITLE);
  const titleElements = titleSlide.getPageElements();
  if (titleElements.length > 0) titleElements[0].asShape().getText().setText(classTitle);
  if (titleElements.length > 1) { try { titleElements[1].remove(); } catch (_) {} }

  // === Helper — reapply styles ===
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

  // === Populate template slides ===
  templateSlides.forEach((tplSlide, i) => {
    const slide = pres.appendSlide(tplSlide);

    slide.getShapes().forEach(shape => {
      if (!shape.getText) return;
      let content;
      try { content = shape.getText().asString().trim(); } catch (_) { return; }

      if (!/^\d+$/.test(content)) return; // skip non-seat numbers

      const student = seatMap.get(content);
      const tr = shape.getText();

      if (student) {
        const oldStyle = tr.getTextStyle();
        tr.setText(student); // Replace ✅
        reapplyStyles(tr, oldStyle); // Preserve formatting ✅
        autoFitName_(tr, oldStyle); // Shrink if long ✅
      } else {
        tr.setText(''); // Blank unused seats ✅
      }
    });

    // Label Student / Teacher View if 2-slide template
    if (templateSlides.length === 2) {
      const label = i === 0 ? ' (Student View)' : ' (Teacher View)';
      slide.insertTextBox(classTitle + label)
        .setLeft(20).setTop(20).setWidth(300);
    }
  });

  Logger.log(`✅ Seating chart generated for ${classTitle} (${timestamp})`);
  return pres.getUrl();
}


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
}


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
