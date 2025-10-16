// You can customize this later or even pull from a settings sheet
const PREFERRED_SEAT_RANGE = [1, 2, 3, 4, 5]; // seats allowed for "Y" students


// ===== MENU CREATION =====
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Seat Wizard')
    .addItem('Generate Seating Charts', 'showGenerateChartSidebar')
    .addSeparator()
    .addItem('Import/Sync Google Classroom Rosters', 'showImportClassesSidebar')
    .addItem('Set Preferential Seats', 'showPreferentialSeatSidebar')
    .addItem('Install Sample Layouts', 'installSampleLayouts')
    .addItem('Add New Layout', 'addNewLayout')
    .addToUi();
}

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


function generateAllSeatingCharts(layoutName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const layoutSheet = ss.getSheetByName(layoutName);
  if (!layoutSheet) throw new Error('Layout sheet not found: ' + layoutName);

  const folderName = "Seating Wizard";
  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");
  const pres = SlidesApp.create(`Seating Charts - All Classes (${timestamp})`);
  const file = DriveApp.getFileById(pres.getId());
  folder.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (_) {}

  // Precompute geometry from layout
  const layout = layoutSheet.getDataRange().getValues();
  const numRows = layout.length;
  const numCols = Math.max(...layout.map(r => r.length));
  const margin = 40;
  const slideW = pres.getPageWidth();
  const slideH = pres.getPageHeight();
  const cellW = (slideW - 2 * margin) / numCols;
  const cellH = (slideH - 2 * margin) / numRows;

  // Start with a blank first slide: we'll remove later if unused
  const initialSlide = pres.getSlides()[0];

  const classSheets = ss.getSheets().filter(s => s.getName().startsWith('Class -'));
  classSheets.forEach((classSheet, idx) => {
    // create a truly blank slide
    const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);

    // Title
    const title = slide.insertTextBox(classSheet.getName());
    title.setLeft(40).setTop(20).setWidth(slideW - 80).setHeight(30);
    title.getText().getTextStyle().setFontSize(18).setBold(true);

    // seat number → student name map for this class
    // Columns: A=Display Name, B=Student Name, C=Seat #
    const seatMap = new Map();
    const students = classSheet.getRange(2, 1, Math.max(0, classSheet.getLastRow() - 1), 3).getValues();
    for (const [displayName, , seat] of students) {
      if (displayName && seat) seatMap.set(Number(seat), displayName);
    }


    for (const [name, seat] of students) {
      if (name && seat) seatMap.set(Number(seat), name);
    }

    // Render layout
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

        const scaleFactor = 1.15;
        const padding = 2;
        const maxW = cellW - padding, maxH = cellH - padding;
        let w = Math.min(cellW * scaleFactor, maxW);
        let h = Math.min(cellH * scaleFactor, maxH);
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
  });

  // Remove the initial slide (which may contain theme placeholders)
  try { initialSlide.remove(); } catch (_) {}
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ Seating charts generated for ${classSheets.length} classes.`,
    'Seat Wizard',
    5
  );

  return pres.getUrl();
}

// ===== SHEET MANAGEMENT =====
function getAllSheets() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

function getClassSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .map(s => s.getName())
    .filter(name => name.startsWith('Class - '));
}

function getLayoutSheets() {
  const sheets = getAllSheets();
  return sheets.filter(name => name.toLowerCase().includes("layout"));
}

// ===== SHOW SIDEBARS =====
function showGenerateChartSidebar() {
  const template = HtmlService.createTemplateFromFile('GenerateChartSidebar');
  template.classSheets = getClassSheets();
  template.layoutSheets = getLayoutSheets();
  SpreadsheetApp.getUi().showSidebar(template.evaluate().setTitle('Generate Seating Chart'));
}

function showEditLayoutSidebar() {
  const template = HtmlService.createTemplateFromFile('EditLayoutSidebar');
  template.layoutSheets = getLayoutSheets();
  SpreadsheetApp.getUi().showSidebar(template.evaluate().setTitle('Edit Layouts'));
}

/**
 * Opens the Randomize Seats sidebar.
 */
function showRandomizeSeatsSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('RandomizeSeatsSidebar')
    .setTitle('Randomize Seats');
  SpreadsheetApp.getUi().showSidebar(html);
}

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
// ===== LAYOUT MANAGEMENT =====
function createNewLayout(name){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if(ss.getSheetByName(name)) return;
  const sheet = ss.insertSheet(name);
  sheet.getRange("A1").setValue("Desk Layout");
  return;
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




/**
 * Sends progress updates to sidebar
 */
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
// ===== SIDEBAR INCLUDE =====
function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function importSelectedCourses(courseIds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const imported = [];

  courseIds.forEach(courseId => {
    try {
      const course = Classroom.Courses.get(courseId);
      const studentsResponse = Classroom.Courses.Students.list(courseId);
      const students = studentsResponse.students || [];
      const studentNames = students.map(s => s.profile.name.fullName);

      // Sync roster (preserves Display Name, Pref, Keep Away, etc.)
      importOrSyncClassRoster(course.name, studentNames);

      // Automatically assign seats
      randomizeSeatsForClass(`Class - ${course.name}`);

      imported.push(course.name);
    } catch (err) {
      Logger.log(`❌ Error importing course ${courseId}: ${err.message}`);
    }
  });
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ Imported ${imported.length} class(es) and auto-assigned seats.`,
    'Seat Wizard',
    5
  );

  return { success: true, count: imported.length };
}

function importOrSyncClassRoster(courseName, studentNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `Class - ${courseName}`;
  let sheet = ss.getSheetByName(sheetName);

  // Target columns (A..F):
  // A Display Name | B Student Name | C Seat Number | D Preferential Seating | E Keep Away 1 | F Keep Away 2
  const headers = ['Display Name', 'Student Name', 'Seat Number', 'Preferential Seating', 'Keep Away 1', 'Keep Away 2'];

  if (!sheet) {
    // Create fresh and write everything
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange('A1:F1').setFontWeight('bold');
    sheet.setFrozenRows(1);

    const rows = studentNames.map((name, i) => [name, name, i + 1, '', '', '']);
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    setupKeepAwayDropdowns(sheet);
    return sheet;
  }

  // If sheet exists, read current data and build a map keyed by Student Name (column B)
  const lastRow = sheet.getLastRow();
  let existingMap = new Map();
  if (lastRow >= 2) {
    const current = sheet.getRange(2, 1, lastRow - 1, 6).getValues(); // A..F
    for (const row of current) {
      const [disp, stud, seat, pref, ka1, ka2] = row;
      if (stud) {
        existingMap.set(stud, { disp, stud, seat: Number(seat) || null, pref, ka1, ka2 });
      }
    }
  }

  // Build new roster, preserving existing where possible
  const N = studentNames.length;
  const allSeats = Array.from({ length: N }, (_, i) => i + 1);
  const used = new Set();

  // First pass: reserve valid unique seats for continuing students
  studentNames.forEach(stud => {
    const ex = existingMap.get(stud);
    if (ex && ex.seat && ex.seat >= 1 && ex.seat <= N && !used.has(ex.seat)) {
      used.add(ex.seat);
    }
  });

  // Pool of free seats for new/adjusted students
  const freeSeats = allSeats.filter(s => !used.has(s));

  // Build final rows
  const merged = studentNames.map(stud => {
    const ex = existingMap.get(stud);
    const disp = ex?.disp || stud; // default display = official
    const pref = ex?.pref || '';
    const ka1  = ex?.ka1  || '';
    const ka2  = ex?.ka2  || '';
    let seat   = ex?.seat || null;

    if (!seat || seat < 1 || seat > N || used.has(seat)) {
      seat = freeSeats.shift() || null;
    }
    used.add(seat);

    return [disp, stud, seat, pref, ka1, ka2];
  });

  // Rewrite sheet cleanly
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange('A1:F1').setFontWeight('bold');
  sheet.setFrozenRows(1);
  if (merged.length) {
    sheet.getRange(2, 1, merged.length, headers.length).setValues(merged);
  }

  // Preferential Seating dropdown ("Y")
  const prefRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Y'], true)
    .setAllowInvalid(true)
    .build();
  if (merged.length) sheet.getRange(2, 4, merged.length, 1).setDataValidation(prefRule);

  // Refresh keep-away dropdowns (now based on Student Name in col B)
  setupKeepAwayDropdowns(sheet);
  return sheet;
}

/**
 * Opens the Import Classes sidebar.
 */
function showImportClassesSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ImportClassesSidebar')
    .setTitle('Import from Google Classroom');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getDropdownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const layouts = ss.getSheets()
    .map(s => s.getName())
    .filter(n => n.startsWith('Layout -'));

  const classes = ss.getSheets()
    .map(s => s.getName())
    .filter(n => n.startsWith('Class -'));

  return { layouts, classes };
}

/**
 * Gets the list of Google Classroom courses for the active user.
 */
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

function randomizeAllClasses() {
  const classSheets = getClassSheets();
  classSheets.forEach(name => randomizeSeats(name));
}


/**
 * Gets the list of active Google Classroom courses.
 */
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
    throw new Error(
      'Unable to fetch Google Classroom courses. Make sure the Classroom API is enabled.'
    );
  }
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


// Helper: Convert seat # → (row, col) grid coordinates
function seatToGrid(seat, cols) {
  const r = Math.floor((seat - 1) / cols);
  const c = (seat - 1) % cols;
  return [r, c];
}

function shuffleArray(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
}


/**
function randomizeSeatsForAllClasses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheets = ss.getSheets().filter(s => s.getName().startsWith('Class -'));
  classSheets.forEach(s => randomizeSeatsForClass(s.getName()));
}*/


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


function showPreferentialSeatSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('PreferentialSeatSidebar')
    .setTitle('Preferential Seat Settings');
  SpreadsheetApp.getUi().showSidebar(html);
}

function savePreferentialSeats(selectedSeats) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('PREFERRED_SEAT_RANGE', JSON.stringify(selectedSeats));
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

//Keep Away Code Below
function setupKeepAwayDropdowns(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Use Student Name column (B) for the list in dropdowns
  const names = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat().filter(Boolean);
  if (names.length === 0) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(names, true)
    .setAllowInvalid(true)
    .build();

  sheet.getRange('E1').setValue('Keep Away 1');
  sheet.getRange('F1').setValue('Keep Away 2');
  sheet.getRange(2, 5, lastRow - 1, 1).setDataValidation(rule);
  sheet.getRange(2, 6, lastRow - 1, 1).setDataValidation(rule);
}



