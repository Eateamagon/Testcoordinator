/**
 * VDOE-Compliant Student Testing Coordinator
 * Server-side Google Apps Script
 *
 * Manages Students, Rooms, Teachers, Templates, Staging, and Assignments sheets.
 * Supports grade-level filtering, school layout templates, and staging groups.
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
  getOrCreateSheet_('Teachers', [
    'TeacherID', 'Name', 'RoomNumber', 'Hallway', 'Grade'
  ]);
  getOrCreateSheet_('Rooms', [
    'RoomName', 'RoomNumber', 'Hallway', 'Rows', 'Columns',
    'MaxCapacity', 'TeacherID', 'Grade'
  ]);
  getOrCreateSheet_('Assignments', [
    'StudentID', 'RoomName', 'Row', 'Column'
  ]);
  getOrCreateSheet_('Staging', [
    'GroupID', 'GroupName', 'StudentID'
  ]);
  getOrCreateSheet_('Templates', [
    'TemplateName', 'DataJSON'
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
  removeAssignmentsForStudent_(studentId);
  removeStagingForStudent_(studentId);
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
// Teacher CRUD
// ---------------------------------------------------------------------------

function addTeacher(data) {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Teachers');
  sheet.appendRow([
    data.teacherId,
    data.name,
    data.roomNumber || '',
    data.hallway || '',
    data.grade || ''
  ]);
  return { success: true };
}

function getTeachers() {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Teachers');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var teachers = [];
  for (var i = 1; i < data.length; i++) {
    teachers.push({
      teacherId: String(data[i][0]),
      name: String(data[i][1]),
      roomNumber: String(data[i][2] || ''),
      hallway: String(data[i][3] || ''),
      grade: String(data[i][4] || '')
    });
  }
  return teachers;
}

function updateTeacher(data) {
  var sheet = getOrCreateSheet_('Teachers');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.teacherId)) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[
        data.teacherId,
        data.name,
        data.roomNumber || '',
        data.hallway || '',
        data.grade || ''
      ]]);
      break;
    }
  }
  return { success: true };
}

function deleteTeacher(teacherId) {
  var sheet = getOrCreateSheet_('Teachers');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(teacherId)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { success: true };
}

// ---------------------------------------------------------------------------
// Room CRUD  (enhanced with RoomNumber, Hallway, TeacherID, Grade)
// ---------------------------------------------------------------------------

function addRoom(data) {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Rooms');
  var cap = parseInt(data.maxCapacity, 10) || (parseInt(data.rows, 10) * parseInt(data.columns, 10));
  sheet.appendRow([
    data.roomName,
    data.roomNumber || '',
    data.hallway || '',
    parseInt(data.rows, 10),
    parseInt(data.columns, 10),
    cap,
    data.teacherId || '',
    data.grade || ''
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
      roomNumber: String(data[i][1] || ''),
      hallway: String(data[i][2] || ''),
      rows: parseInt(data[i][3], 10),
      columns: parseInt(data[i][4], 10),
      maxCapacity: parseInt(data[i][5], 10) || 0,
      teacherId: String(data[i][6] || ''),
      grade: String(data[i][7] || '')
    });
  }
  return rooms;
}

function updateRoom(data) {
  var sheet = getOrCreateSheet_('Rooms');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === data.roomName) {
      var cap = parseInt(data.maxCapacity, 10) || (parseInt(data.rows, 10) * parseInt(data.columns, 10));
      sheet.getRange(i + 1, 1, 1, 8).setValues([[
        data.roomName,
        data.roomNumber || '',
        data.hallway || '',
        parseInt(data.rows, 10),
        parseInt(data.columns, 10),
        cap,
        data.teacherId || '',
        data.grade || ''
      ]]);
      break;
    }
  }
  return { success: true };
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
  removeAssignmentsForRoom_(roomName);
  return { success: true };
}

// ---------------------------------------------------------------------------
// Assignment CRUD
// ---------------------------------------------------------------------------

function saveAssignment(studentId, roomName, row, col) {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Assignments');
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
// Staging Groups  (plan combos before placing into rooms)
// ---------------------------------------------------------------------------

function getStagingGroups() {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Staging');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var gid = String(data[i][0]);
    if (!map[gid]) map[gid] = { groupId: gid, groupName: String(data[i][1]), studentIds: [] };
    if (data[i][2]) map[gid].studentIds.push(String(data[i][2]));
  }
  var groups = [];
  for (var k in map) groups.push(map[k]);
  return groups;
}

function createStagingGroup(groupName) {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Staging');
  var gid = 'G' + Date.now();
  sheet.appendRow([gid, groupName, '']);
  return { success: true, groupId: gid };
}

function addStudentToStagingGroup(groupId, studentId) {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Staging');
  // Look up group name
  var data = sheet.getDataRange().getValues();
  var gName = '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === groupId) { gName = String(data[i][1]); break; }
  }
  sheet.appendRow([groupId, gName, studentId]);
  return { success: true };
}

function removeStudentFromStagingGroup(groupId, studentId) {
  var sheet = getOrCreateSheet_('Staging');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === groupId && String(data[i][2]) === studentId) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { success: true };
}

function deleteStagingGroup(groupId) {
  var sheet = getOrCreateSheet_('Staging');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === groupId) {
      sheet.deleteRow(i + 1);
    }
  }
  return { success: true };
}

function removeStagingForStudent_(studentId) {
  var sheet = getOrCreateSheet_('Staging');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][2]) === String(studentId)) {
      sheet.deleteRow(i + 1);
    }
  }
}

// ---------------------------------------------------------------------------
// School Templates  (save/load teacher + room combos)
// ---------------------------------------------------------------------------

function getTemplates() {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Templates');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var templates = [];
  for (var i = 1; i < data.length; i++) {
    templates.push({
      templateName: String(data[i][0]),
      data: JSON.parse(data[i][1] || '{}')
    });
  }
  return templates;
}

/**
 * Save current teachers + rooms as a named template.
 */
function saveTemplate(templateName) {
  ensureSheets_();
  var teachers = getTeachers();
  var rooms = getRooms();
  var payload = JSON.stringify({ teachers: teachers, rooms: rooms });

  var sheet = getOrCreateSheet_('Templates');
  var data = sheet.getDataRange().getValues();
  // Overwrite if exists
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === templateName) {
      sheet.getRange(i + 1, 2).setValue(payload);
      return { success: true, message: 'Template "' + templateName + '" updated.' };
    }
  }
  sheet.appendRow([templateName, payload]);
  return { success: true, message: 'Template "' + templateName + '" saved.' };
}

/**
 * Load a template: replaces Teachers and Rooms sheets with stored data.
 */
function loadTemplate(templateName) {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Templates');
  var data = sheet.getDataRange().getValues();
  var payload = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === templateName) {
      payload = JSON.parse(data[i][1] || '{}');
      break;
    }
  }
  if (!payload) return { success: false, message: 'Template not found.' };

  // Clear and repopulate Teachers
  var tSheet = getOrCreateSheet_('Teachers');
  if (tSheet.getLastRow() > 1) {
    tSheet.getRange(2, 1, tSheet.getLastRow() - 1, tSheet.getLastColumn()).clearContent();
  }
  (payload.teachers || []).forEach(function (t) {
    tSheet.appendRow([t.teacherId, t.name, t.roomNumber, t.hallway, t.grade]);
  });

  // Clear and repopulate Rooms
  var rSheet = getOrCreateSheet_('Rooms');
  if (rSheet.getLastRow() > 1) {
    rSheet.getRange(2, 1, rSheet.getLastRow() - 1, rSheet.getLastColumn()).clearContent();
  }
  (payload.rooms || []).forEach(function (r) {
    rSheet.appendRow([
      r.roomName, r.roomNumber || '', r.hallway || '',
      r.rows, r.columns,
      r.maxCapacity || (r.rows * r.columns),
      r.teacherId || '', r.grade || ''
    ]);
  });

  return { success: true, message: 'Template "' + templateName + '" loaded.' };
}

function deleteTemplate(templateName) {
  var sheet = getOrCreateSheet_('Templates');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === templateName) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { success: true };
}

// ---------------------------------------------------------------------------
// Auto-Suggest / Recommendation Engine  (grade-aware)
// ---------------------------------------------------------------------------

/**
 * Generate VDOE-compliant placement recommendations.
 * If gradeFilter is set, only students+rooms for that grade are considered.
 */
function generateRecommendations(smallGroupLimit, gradeFilter) {
  smallGroupLimit = parseInt(smallGroupLimit, 10) || 10;
  var students = getStudents();
  var rooms = getRooms();

  // Apply grade filter
  if (gradeFilter) {
    students = students.filter(function (s) { return s.grade === gradeFilter; });
    rooms = rooms.filter(function (r) { return !r.grade || r.grade === gradeFilter || r.grade === ''; });
  }

  if (!students.length || !rooms.length) {
    return { error: 'Need at least one student and one room' + (gradeFilter ? ' for grade ' + gradeFilter : '') + '.' };
  }

  rooms.sort(function (a, b) {
    return (a.rows * a.columns) - (b.rows * b.columns);
  });

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
    if (s.proximity && !s.oneToOne) { proximity.push(s.studentId); }
  });

  var assignments = [];
  var usedSeats = {};
  var roomStudentCount = {};

  rooms.forEach(function (r) {
    usedSeats[r.roomName] = {};
    roomStudentCount[r.roomName] = 0;
  });

  function placeStudent(student, roomName) {
    var room = rooms.filter(function (r) { return r.roomName === roomName; })[0];
    if (!room) return false;
    var isFront = proximity.indexOf(student.studentId) !== -1;
    var rStart = isFront ? 1 : 1;
    for (var r = rStart; r <= room.rows; r++) {
      for (var c = 1; c <= room.columns; c++) {
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

  // 1:1 Testing — one per room
  oneToOne.forEach(function (s) {
    while (roomIndex < rooms.length && roomStudentCount[rooms[roomIndex].roomName] > 0) {
      roomIndex++;
    }
    if (roomIndex < rooms.length) {
      placeStudent(s, rooms[roomIndex].roomName);
      roomIndex++;
    }
  });

  // Read Aloud — grouped together
  if (readAloud.length) {
    var raRoomIdx = roomIndex;
    readAloud.forEach(function (s) {
      if (raRoomIdx >= rooms.length) return;
      var room = rooms[raRoomIdx];
      var cap = room.rows * room.columns;
      if (roomStudentCount[room.roomName] >= cap) {
        raRoomIdx++;
        if (raRoomIdx >= rooms.length) return;
      }
      placeStudent(s, rooms[raRoomIdx].roomName);
    });
    if (raRoomIdx >= roomIndex) roomIndex = raRoomIdx;
  }

  // Small Group
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
      placeStudent(s, rooms[sgRoomIdx].roomName);
    });
    if (sgRoomIdx >= roomIndex) roomIndex = sgRoomIdx;
  }

  // General students
  general.forEach(function (s) {
    for (var ri = 0; ri < rooms.length; ri++) {
      var room = rooms[ri];
      var cap = room.rows * room.columns;
      if (roomStudentCount[room.roomName] < cap) {
        if (placeStudent(s, room.roomName)) break;
      }
    }
  });

  return assignments;
}

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

/**
 * Place an entire staging group into the next available seats of a room.
 */
function placeStagingGroupInRoom(groupId, roomName) {
  ensureSheets_();
  var groups = getStagingGroups();
  var group = groups.filter(function (g) { return g.groupId === groupId; })[0];
  if (!group) return { success: false, message: 'Group not found.' };

  var rooms = getRooms();
  var room = rooms.filter(function (r) { return r.roomName === roomName; })[0];
  if (!room) return { success: false, message: 'Room not found.' };

  var currentAssignments = getAssignments();
  var seatMap = {};
  currentAssignments.forEach(function (a) {
    if (a.roomName === roomName) seatMap[a.row + ',' + a.column] = true;
  });

  var sheet = getOrCreateSheet_('Assignments');
  var placed = 0;

  group.studentIds.forEach(function (sid) {
    if (!sid) return;
    // Remove existing assignment for this student
    removeAssignmentsForStudent_(sid);
    // Find next open seat
    for (var r = 1; r <= room.rows; r++) {
      for (var c = 1; c <= room.columns; c++) {
        var key = r + ',' + c;
        if (!seatMap[key]) {
          seatMap[key] = true;
          sheet.appendRow([sid, roomName, r, c]);
          placed++;
          return;
        }
      }
    }
  });

  return { success: true, message: placed + ' student(s) placed in ' + roomName + '.' };
}

// ---------------------------------------------------------------------------
// Finalize / Export Layout
// ---------------------------------------------------------------------------

function finalizeLayout() {
  ensureSheets_();
  var students = getStudents();
  var rooms = getRooms();
  var teachers = getTeachers();
  var assignments = getAssignments();
  var ss = getSpreadsheet_();

  var studentMap = {};
  students.forEach(function (s) { studentMap[s.studentId] = s; });
  var teacherMap = {};
  teachers.forEach(function (t) { teacherMap[t.teacherId] = t; });

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

    // Header rows
    layoutSheet.insertRowBefore(1);
    var teacher = teacherMap[room.teacherId];
    var title = room.roomName;
    if (room.roomNumber) title += ' (Rm ' + room.roomNumber + ')';
    if (room.hallway) title += ' — ' + room.hallway;
    if (teacher) title += ' — ' + teacher.name;
    layoutSheet.getRange(1, 1).setValue(title + ' — Finalized Layout');
    layoutSheet.getRange(1, 1).setFontWeight('bold').setFontSize(12);

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
