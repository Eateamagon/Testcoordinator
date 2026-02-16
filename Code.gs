/**
 * VDOE-Compliant Student Testing Coordinator
 * Server-side Google Apps Script
 *
 * Manages Students, Rooms, and Assignments sheets.
 * Provides the web app entry point and all data operations.
 */

// ---------------------------------------------------------------------------
// Web App Entry Point
// ---------------------------------------------------------------------------

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('VDOE Testing Coordinator')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Include partial HTML files (stylesheet.html, javascript.html). */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------------------------------------------------------------------------
// Sheet Helpers
// ---------------------------------------------------------------------------

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Return the named sheet, creating it with optional headers if missing.
 */
function getOrCreateSheet_(name, headers) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  return sheet;
}

function ensureSheets_() {
  getOrCreateSheet_('Students', [
    'StudentID', 'Name', 'Grade',
    'SmallGroup', 'ReadAloud', 'OneToOne',
    'Proximity', 'Prompting', 'OtherAccommodations'
  ]);
  getOrCreateSheet_('Rooms', [
    'RoomName', 'Rows', 'Columns', 'MaxCapacity'
  ]);
  getOrCreateSheet_('Assignments', [
    'StudentID', 'RoomName', 'Row', 'Column'
  ]);
}

// ---------------------------------------------------------------------------
// Student CRUD
// ---------------------------------------------------------------------------

function addStudent(data) {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Students');
  sheet.appendRow([
    data.studentId,
    data.name,
    data.grade,
    data.smallGroup ? 'Y' : '',
    data.readAloud ? 'Y' : '',
    data.oneToOne ? 'Y' : '',
    data.proximity ? 'Y' : '',
    data.prompting ? 'Y' : '',
    data.otherAccommodations || ''
  ]);
  return { success: true };
}

function getStudents() {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Students');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var students = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    students.push({
      studentId: String(row[0]),
      name: String(row[1]),
      grade: String(row[2]),
      smallGroup: row[3] === 'Y',
      readAloud: row[4] === 'Y',
      oneToOne: row[5] === 'Y',
      proximity: row[6] === 'Y',
      prompting: row[7] === 'Y',
      otherAccommodations: String(row[8] || '')
    });
  }
  return students;
}

function deleteStudent(studentId) {
  var sheet = getOrCreateSheet_('Students');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(studentId)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  // Also remove any assignments for that student
  removeAssignmentsForStudent_(studentId);
  return { success: true };
}

function updateStudent(data) {
  var sheet = getOrCreateSheet_('Students');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.studentId)) {
      var rowNum = i + 1;
      sheet.getRange(rowNum, 1, 1, 9).setValues([[
        data.studentId,
        data.name,
        data.grade,
        data.smallGroup ? 'Y' : '',
        data.readAloud ? 'Y' : '',
        data.oneToOne ? 'Y' : '',
        data.proximity ? 'Y' : '',
        data.prompting ? 'Y' : '',
        data.otherAccommodations || ''
      ]]);
      break;
    }
  }
  return { success: true };
}

// ---------------------------------------------------------------------------
// Room CRUD
// ---------------------------------------------------------------------------

function addRoom(data) {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Rooms');
  sheet.appendRow([
    data.roomName,
    parseInt(data.rows, 10),
    parseInt(data.columns, 10),
    parseInt(data.maxCapacity, 10) || (parseInt(data.rows, 10) * parseInt(data.columns, 10))
  ]);
  return { success: true };
}

function getRooms() {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Rooms');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var rooms = [];
  for (var i = 1; i < data.length; i++) {
    rooms.push({
      roomName: String(data[i][0]),
      rows: parseInt(data[i][1], 10),
      columns: parseInt(data[i][2], 10),
      maxCapacity: parseInt(data[i][3], 10) || 0
    });
  }
  return rooms;
}

function deleteRoom(roomName) {
  var sheet = getOrCreateSheet_('Rooms');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === roomName) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  // Remove assignments for that room
  removeAssignmentsForRoom_(roomName);
  return { success: true };
}

// ---------------------------------------------------------------------------
// Assignment CRUD
// ---------------------------------------------------------------------------

function saveAssignment(studentId, roomName, row, col) {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Assignments');
  // Remove prior assignment for this student
  removeAssignmentsForStudent_(studentId);
  sheet.appendRow([studentId, roomName, row, col]);
  return { success: true };
}

function removeAssignment(studentId) {
  removeAssignmentsForStudent_(studentId);
  return { success: true };
}

function getAssignments() {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Assignments');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var assignments = [];
  for (var i = 1; i < data.length; i++) {
    assignments.push({
      studentId: String(data[i][0]),
      roomName: String(data[i][1]),
      row: parseInt(data[i][2], 10),
      column: parseInt(data[i][3], 10)
    });
  }
  return assignments;
}

function removeAssignmentsForStudent_(studentId) {
  var sheet = getOrCreateSheet_('Assignments');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(studentId)) {
      sheet.deleteRow(i + 1);
    }
  }
}

function removeAssignmentsForRoom_(roomName) {
  var sheet = getOrCreateSheet_('Assignments');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === roomName) {
      sheet.deleteRow(i + 1);
    }
  }
}

// ---------------------------------------------------------------------------
// Auto-Suggest / Recommendation Engine
// ---------------------------------------------------------------------------

/**
 * Generate VDOE-compliant placement recommendations.
 *
 * Rules:
 *  1. 1:1 Testing students -> one student per room (pick smallest rooms first).
 *  2. Read Aloud students -> grouped together in designated rooms.
 *  3. Small Group students -> rooms with capacity <= smallGroupLimit.
 *  4. Proximity students -> placed in row 1 (closest to proctor).
 *  5. Remaining students fill remaining seats.
 *
 * Returns an array of { studentId, roomName, row, column } objects.
 */
function generateRecommendations(smallGroupLimit) {
  smallGroupLimit = parseInt(smallGroupLimit, 10) || 10;
  var students = getStudents();
  var rooms = getRooms();

  if (!students.length || !rooms.length) {
    return { error: 'Need at least one student and one room.' };
  }

  // Sort rooms by capacity ascending (smallest first for 1:1)
  rooms.sort(function (a, b) {
    return (a.rows * a.columns) - (b.rows * b.columns);
  });

  // Categorise students
  var oneToOne = [];
  var readAloud = [];
  var smallGroup = [];
  var proximity = [];
  var general = [];

  students.forEach(function (s) {
    if (s.oneToOne) { oneToOne.push(s); }
    else if (s.readAloud) { readAloud.push(s); }
    else if (s.smallGroup) { smallGroup.push(s); }
    else { general.push(s); }
    // Proximity is an overlay; tracked separately for seating preference
    if (s.proximity && !s.oneToOne) { proximity.push(s.studentId); }
  });

  var assignments = [];
  var usedSeats = {}; // roomName -> Set of "r,c"
  var roomStudentCount = {}; // roomName -> count

  rooms.forEach(function (r) {
    usedSeats[r.roomName] = {};
    roomStudentCount[r.roomName] = 0;
  });

  function placeStudent(student, roomName, preferFront) {
    var room = rooms.filter(function (r) { return r.roomName === roomName; })[0];
    if (!room) return false;
    var rows = room.rows;
    var cols = room.columns;
    // If preferFront, start from row 1; else start from first available
    for (var r = 1; r <= rows; r++) {
      for (var c = 1; c <= cols; c++) {
        var key = r + ',' + c;
        if (!usedSeats[roomName][key]) {
          usedSeats[roomName][key] = true;
          roomStudentCount[roomName] = (roomStudentCount[roomName] || 0) + 1;
          assignments.push({
            studentId: student.studentId,
            roomName: roomName,
            row: r,
            column: c
          });
          return true;
        }
      }
    }
    return false;
  }

  var roomIndex = 0;

  // 1. 1:1 Testing — one per room
  oneToOne.forEach(function (s) {
    while (roomIndex < rooms.length && roomStudentCount[rooms[roomIndex].roomName] > 0) {
      roomIndex++;
    }
    if (roomIndex < rooms.length) {
      placeStudent(s, rooms[roomIndex].roomName, true);
      roomIndex++;
    }
  });

  // 2. Read Aloud — group together in next available room(s)
  if (readAloud.length) {
    // Find a room big enough, or pack into multiple
    var raRoomIdx = roomIndex;
    readAloud.forEach(function (s) {
      if (raRoomIdx >= rooms.length) return;
      var room = rooms[raRoomIdx];
      var cap = room.rows * room.columns;
      if (roomStudentCount[room.roomName] >= cap) {
        raRoomIdx++;
        if (raRoomIdx >= rooms.length) return;
      }
      var isFront = proximity.indexOf(s.studentId) !== -1;
      placeStudent(s, rooms[raRoomIdx].roomName, isFront);
    });
    // Advance roomIndex past rooms used by read-aloud if they were after current index
    if (raRoomIdx >= roomIndex) roomIndex = raRoomIdx;
  }

  // 3. Small Group
  if (smallGroup.length) {
    var sgRoomIdx = roomIndex < rooms.length ? roomIndex : rooms.length - 1;
    smallGroup.forEach(function (s) {
      if (sgRoomIdx >= rooms.length) return;
      var room = rooms[sgRoomIdx];
      var cap = Math.min(room.rows * room.columns, smallGroupLimit);
      if (roomStudentCount[room.roomName] >= cap) {
        sgRoomIdx++;
        if (sgRoomIdx >= rooms.length) return;
      }
      var isFront = proximity.indexOf(s.studentId) !== -1;
      placeStudent(s, rooms[sgRoomIdx].roomName, isFront);
    });
    if (sgRoomIdx >= roomIndex) roomIndex = sgRoomIdx;
  }

  // 4. General students — fill remaining seats
  general.forEach(function (s) {
    for (var ri = 0; ri < rooms.length; ri++) {
      var room = rooms[ri];
      var cap = room.rows * room.columns;
      if (roomStudentCount[room.roomName] < cap) {
        var isFront = proximity.indexOf(s.studentId) !== -1;
        if (placeStudent(s, room.roomName, isFront)) break;
      }
    }
  });

  return assignments;
}

// ---------------------------------------------------------------------------
// Apply recommendations (save to Assignments sheet)
// ---------------------------------------------------------------------------

function applyRecommendations(assignments) {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Assignments');
  // Clear existing assignments (keep header)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  assignments.forEach(function (a) {
    sheet.appendRow([a.studentId, a.roomName, a.row, a.column]);
  });
  return { success: true };
}

// ---------------------------------------------------------------------------
// Finalize / Export Layout
// ---------------------------------------------------------------------------

/**
 * Create (or overwrite) a sheet per room that mirrors the visual grid.
 * Each cell contains "Name (codes)" where codes are accommodation abbreviations.
 */
function finalizeLayout() {
  ensureSheets_();
  var students = getStudents();
  var rooms = getRooms();
  var assignments = getAssignments();
  var ss = getSpreadsheet_();

  // Build lookup: studentId -> student object
  var studentMap = {};
  students.forEach(function (s) { studentMap[s.studentId] = s; });

  rooms.forEach(function (room) {
    var sheetName = 'Layout - ' + room.roomName;
    var existing = ss.getSheetByName(sheetName);
    if (existing) ss.deleteSheet(existing);

    var layoutSheet = ss.insertSheet(sheetName);
    var grid = [];
    for (var r = 0; r < room.rows; r++) {
      var row = [];
      for (var c = 0; c < room.columns; c++) {
        row.push('');
      }
      grid.push(row);
    }

    // Fill grid from assignments
    assignments.forEach(function (a) {
      if (a.roomName !== room.roomName) return;
      var s = studentMap[a.studentId];
      if (!s) return;
      var codes = buildAccommodationCodes_(s);
      var label = s.name;
      if (codes) label += ' (' + codes + ')';
      var ri = a.row - 1;
      var ci = a.column - 1;
      if (ri >= 0 && ri < room.rows && ci >= 0 && ci < room.columns) {
        grid[ri][ci] = label;
      }
    });

    if (room.rows > 0 && room.columns > 0) {
      layoutSheet.getRange(1, 1, room.rows, room.columns).setValues(grid);
    }

    // Add title row above grid
    layoutSheet.insertRowBefore(1);
    layoutSheet.getRange(1, 1).setValue(room.roomName + ' — Finalized Layout');
    layoutSheet.getRange(1, 1).setFontWeight('bold').setFontSize(12);

    // Add proctor label
    layoutSheet.insertRowAfter(1);
    layoutSheet.getRange(2, 1).setValue('[ PROCTOR STATION — Row 1 ]');
    layoutSheet.getRange(2, 1).setFontStyle('italic');
  });

  return { success: true, message: 'Layout sheets created for ' + rooms.length + ' room(s).' };
}

function buildAccommodationCodes_(student) {
  var codes = [];
  if (student.smallGroup) codes.push('SG');
  if (student.readAloud) codes.push('RA');
  if (student.oneToOne) codes.push('1:1');
  if (student.proximity) codes.push('PROX');
  if (student.prompting) codes.push('PMPT');
  if (student.otherAccommodations) codes.push(student.otherAccommodations);
  return codes.join(', ');
}
