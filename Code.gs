/**
 * VDOE-Compliant Student Testing Coordinator
 * Server-side Google Apps Script
 *
 * Manages Students, Rooms, Teachers, Templates, Staging, Assignments,
 * and FillerCells sheets. Supports floor designation and PDF/Word export.
 */

// ---------------------------------------------------------------------------
// Web App Entry Point
// ---------------------------------------------------------------------------

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Testing Coordinator')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

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
    'MaxCapacity', 'TeacherID', 'Grade', 'Floor'
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
  getOrCreateSheet_('FillerCells', [
    'RoomName', 'Row', 'Column'
  ]);
  getOrCreateSheet_('DesignerLayouts', [
    'LayoutName', 'DataJSON'
  ]);
  getOrCreateSheet_('Backups', [
    'BackupName', 'CreatedAt', 'DataJSON'
  ]);
  getOrCreateSheet_('AuditLog', [
    'Timestamp', 'User', 'Action', 'Details', 'Success', 'ErrorMessage'
  ]);
}

/**
 * Log an action to the AuditLog sheet.
 * @param {string} action
 * @param {string} details
 * @param {boolean} success
 * @param {string} [error]
 */
function logAudit_(action, details, success, error) {
  try {
    var sheet = getOrCreateSheet_('AuditLog');
    sheet.appendRow([
      new Date(),
      Session.getActiveUser().getEmail(),
      action,
      details || '',
      success ? 'Y' : 'N',
      error || ''
    ]);
  } catch (e) {
    Logger.log('FAILED TO LOG AUDIT: ' + e.toString());
  }
}

// ---------------------------------------------------------------------------
// Student CRUD
// ---------------------------------------------------------------------------

function addStudent(data) {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('Students');
    sheet.appendRow([
      data.studentId, data.name, data.grade,
      data.smallGroup ? 'Y' : '', data.readAloud ? 'Y' : '',
      data.oneToOne ? 'Y' : '', data.proximity ? 'Y' : '',
      data.prompting ? 'Y' : '', data.otherAccommodations || ''
    ]);
    logAudit_('Add Student', 'ID: ' + data.studentId + ', Name: ' + data.name, true);
    return { success: true };
  } catch (e) {
    logAudit_('Add Student', 'ID: ' + (data ? data.studentId : 'N/A'), false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function getStudents() {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('Students');
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var students = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      students.push({
        studentId: String(row[0]), name: String(row[1]), grade: String(row[2]),
        smallGroup: row[3] === 'Y', readAloud: row[4] === 'Y',
        oneToOne: row[5] === 'Y', proximity: row[6] === 'Y',
        prompting: row[7] === 'Y', otherAccommodations: String(row[8] || '')
      });
    }
    return students;
  } catch (e) {
    logAudit_('Get Students', '', false, e.toString());
    throw e; // Rethrow for client-side failure handler
  }
}

function deleteStudent(studentId) {
  try {
    var sheet = getOrCreateSheet_('Students');
    var data = sheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(studentId)) {
        sheet.deleteRow(i + 1);
        found = true;
        break;
      }
    }
    removeAssignmentsForStudent_(studentId);
    removeStagingForStudent_(studentId);
    logAudit_('Delete Student', 'ID: ' + studentId, true);
    return { success: true };
  } catch (e) {
    logAudit_('Delete Student', 'ID: ' + studentId, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function updateStudent(data) {
  try {
    var sheet = getOrCreateSheet_('Students');
    var rows = sheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.studentId)) {
        sheet.getRange(i + 1, 1, 1, 9).setValues([[
          data.studentId, data.name, data.grade,
          data.smallGroup ? 'Y' : '', data.readAloud ? 'Y' : '',
          data.oneToOne ? 'Y' : '', data.proximity ? 'Y' : '',
          data.prompting ? 'Y' : '', data.otherAccommodations || ''
        ]]);
        found = true;
        break;
      }
    }
    logAudit_('Update Student', 'ID: ' + data.studentId, true);
    return { success: true };
  } catch (e) {
    logAudit_('Update Student', 'ID: ' + (data ? data.studentId : 'N/A'), false, e.toString());
    return { success: false, message: e.toString() };
  }
}

// ---------------------------------------------------------------------------
// Teacher CRUD
// ---------------------------------------------------------------------------

function addTeacher(data) {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('Teachers');
    sheet.appendRow([data.teacherId, data.name, data.roomNumber || '', data.hallway || '', data.grade || '']);
    logAudit_('Add Teacher', 'ID: ' + data.teacherId + ', Name: ' + data.name, true);
    return { success: true };
  } catch (e) {
    logAudit_('Add Teacher', 'ID: ' + (data ? data.teacherId : 'N/A'), false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function getTeachers() {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('Teachers');
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var teachers = [];
    for (var i = 1; i < data.length; i++) {
      teachers.push({
        teacherId: String(data[i][0]), name: String(data[i][1]),
        roomNumber: String(data[i][2] || ''), hallway: String(data[i][3] || ''),
        grade: String(data[i][4] || '')
      });
    }
    return teachers;
  } catch (e) {
    logAudit_('Get Teachers', '', false, e.toString());
    throw e;
  }
}

function updateTeacher(data) {
  try {
    var sheet = getOrCreateSheet_('Teachers');
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.teacherId)) {
        sheet.getRange(i + 1, 1, 1, 5).setValues([[
          data.teacherId, data.name, data.roomNumber || '', data.hallway || '', data.grade || ''
        ]]);
        break;
      }
    }
    logAudit_('Update Teacher', 'ID: ' + data.teacherId, true);
    return { success: true };
  } catch (e) {
    logAudit_('Update Teacher', 'ID: ' + (data ? data.teacherId : 'N/A'), false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function deleteTeacher(teacherId) {
  try {
    var sheet = getOrCreateSheet_('Teachers');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(teacherId)) { sheet.deleteRow(i + 1); break; }
    }
    logAudit_('Delete Teacher', 'ID: ' + teacherId, true);
    return { success: true };
  } catch (e) {
    logAudit_('Delete Teacher', 'ID: ' + teacherId, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

// ---------------------------------------------------------------------------
// Room CRUD  (now includes Floor)
// ---------------------------------------------------------------------------

function addRoom(data) {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('Rooms');
    var cap = parseInt(data.maxCapacity, 10) || (parseInt(data.rows, 10) * parseInt(data.columns, 10));
    sheet.appendRow([
      data.roomName, data.roomNumber || '', data.hallway || '',
      parseInt(data.rows, 10), parseInt(data.columns, 10), cap,
      data.teacherId || '', data.grade || '', data.floor || '1'
    ]);
    logAudit_('Add Room', 'Name: ' + data.roomName + ', Cap: ' + cap, true);
    return { success: true };
  } catch (e) {
    logAudit_('Add Room', 'Name: ' + (data ? data.roomName : 'N/A'), false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function getRooms() {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('Rooms');
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var rooms = [];
    for (var i = 1; i < data.length; i++) {
      rooms.push({
        roomName: String(data[i][0]), roomNumber: String(data[i][1] || ''),
        hallway: String(data[i][2] || ''), rows: parseInt(data[i][3], 10),
        columns: parseInt(data[i][4], 10), maxCapacity: parseInt(data[i][5], 10) || 0,
        teacherId: String(data[i][6] || ''), grade: String(data[i][7] || ''),
        floor: String(data[i][8] || '1')
      });
    }
    return rooms;
  } catch (e) {
    logAudit_('Get Rooms', '', false, e.toString());
    throw e;
  }
}

function updateRoom(data) {
  try {
    var sheet = getOrCreateSheet_('Rooms');
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === data.roomName) {
        var cap = parseInt(data.maxCapacity, 10) || (parseInt(data.rows, 10) * parseInt(data.columns, 10));
        sheet.getRange(i + 1, 1, 1, 9).setValues([[
          data.roomName, data.roomNumber || '', data.hallway || '',
          parseInt(data.rows, 10), parseInt(data.columns, 10), cap,
          data.teacherId || '', data.grade || '', data.floor || '1'
        ]]);
        break;
      }
    }
    logAudit_('Update Room', 'Name: ' + data.roomName, true);
    return { success: true };
  } catch (e) {
    logAudit_('Update Room', 'Name: ' + (data ? data.roomName : 'N/A'), false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function deleteRoom(roomName) {
  try {
    var sheet = getOrCreateSheet_('Rooms');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === roomName) { sheet.deleteRow(i + 1); break; }
    }
    removeAssignmentsForRoom_(roomName);
    removeFillerCellsForRoom_(roomName);
    logAudit_('Delete Room', 'Name: ' + roomName, true);
    return { success: true };
  } catch (e) {
    logAudit_('Delete Room', 'Name: ' + roomName, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

// ---------------------------------------------------------------------------
// Filler (Blocked) Cells
// ---------------------------------------------------------------------------

function getFillerCells() {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('FillerCells');
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var cells = [];
    for (var i = 1; i < data.length; i++) {
      cells.push({
        roomName: String(data[i][0]),
        row: parseInt(data[i][1], 10),
        column: parseInt(data[i][2], 10)
      });
    }
    return cells;
  } catch (e) {
    logAudit_('Get Filler Cells', '', false, e.toString());
    throw e;
  }
}

function toggleFillerCell(roomName, row, col) {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('FillerCells');
    var data = sheet.getDataRange().getValues();
    // Check if already exists — if so, remove it
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === roomName &&
          parseInt(data[i][1], 10) === row &&
          parseInt(data[i][2], 10) === col) {
        sheet.deleteRow(i + 1);
        logAudit_('Toggle Filler', 'Room: ' + roomName + ', R' + row + 'C' + col + ' OFF', true);
        return { success: true, isFiller: false };
      }
    }
    // Add it
    sheet.appendRow([roomName, row, col]);
    logAudit_('Toggle Filler', 'Room: ' + roomName + ', R' + row + 'C' + col + ' ON', true);
    return { success: true, isFiller: true };
  } catch (e) {
    logAudit_('Toggle Filler', 'Room: ' + roomName, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function removeFillerCellsForRoom_(roomName) {
  var sheet = getOrCreateSheet_('FillerCells');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === roomName) sheet.deleteRow(i + 1);
  }
}

// ---------------------------------------------------------------------------
// Assignment CRUD
// ---------------------------------------------------------------------------

function saveAssignment(studentId, roomName, row, col) {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('Assignments');
    removeAssignmentsForStudent_(studentId);
    sheet.appendRow([studentId, roomName, row, col]);
    logAudit_('Save Assignment', 'SID: ' + studentId + ', Room: ' + roomName + ' (R' + row + 'C' + col + ')', true);
    return { success: true };
  } catch (e) {
    logAudit_('Save Assignment', 'SID: ' + studentId, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function removeAssignment(studentId) {
  try {
    removeAssignmentsForStudent_(studentId);
    logAudit_('Remove Assignment', 'SID: ' + studentId, true);
    return { success: true };
  } catch (e) {
    logAudit_('Remove Assignment', 'SID: ' + studentId, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function getAssignments() {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('Assignments');
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var assignments = [];
    for (var i = 1; i < data.length; i++) {
      assignments.push({
        studentId: String(data[i][0]), roomName: String(data[i][1]),
        row: parseInt(data[i][2], 10), column: parseInt(data[i][3], 10)
      });
    }
    return assignments;
  } catch (e) {
    logAudit_('Get Assignments', '', false, e.toString());
    throw e;
  }
}

function removeAssignmentsForStudent_(studentId) {
  var sheet = getOrCreateSheet_('Assignments');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(studentId)) sheet.deleteRow(i + 1);
  }
}

function removeAssignmentsForRoom_(roomName) {
  var sheet = getOrCreateSheet_('Assignments');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === roomName) sheet.deleteRow(i + 1);
  }
}

// ---------------------------------------------------------------------------
// Staging Groups
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
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('Staging');
    var gid = 'G' + Date.now();
    sheet.appendRow([gid, groupName, '']);
    logAudit_('Create Staging Group', 'Name: ' + groupName + ', GID: ' + gid, true);
    return { success: true, groupId: gid };
  } catch (e) {
    logAudit_('Create Staging Group', 'Name: ' + groupName, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function addStudentToStagingGroup(groupId, studentId) {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('Staging');
    var data = sheet.getDataRange().getValues();
    var gName = '';
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === groupId) { gName = String(data[i][1]); break; }
    }
    sheet.appendRow([groupId, gName, studentId]);
    logAudit_('Staging: Add Student', 'GID: ' + groupId + ', SID: ' + studentId, true);
    return { success: true };
  } catch (e) {
    logAudit_('Staging: Add Student', 'GID: ' + groupId + ', SID: ' + studentId, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function removeStudentFromStagingGroup(groupId, studentId) {
  try {
    var sheet = getOrCreateSheet_('Staging');
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === groupId && String(data[i][2]) === studentId) {
        sheet.deleteRow(i + 1); break;
      }
    }
    logAudit_('Staging: Remove Student', 'GID: ' + groupId + ', SID: ' + studentId, true);
    return { success: true };
  } catch (e) {
    logAudit_('Staging: Remove Student', 'GID: ' + groupId + ', SID: ' + studentId, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function deleteStagingGroup(groupId) {
  try {
    var sheet = getOrCreateSheet_('Staging');
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === groupId) sheet.deleteRow(i + 1);
    }
    logAudit_('Delete Staging Group', 'GID: ' + groupId, true);
    return { success: true };
  } catch (e) {
    logAudit_('Delete Staging Group', 'GID: ' + groupId, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function removeStagingForStudent_(studentId) {
  var sheet = getOrCreateSheet_('Staging');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][2]) === String(studentId)) sheet.deleteRow(i + 1);
  }
}

function placeStagingGroupInRoom(groupId, roomName) {
  try {
    ensureSheets_();
    var groups = getStagingGroups();
    var group = groups.filter(function (g) { return g.groupId === groupId; })[0];
    if (!group) return { success: false, message: 'Group not found.' };
    var rooms = getRooms();
    var room = rooms.filter(function (r) { return r.roomName === roomName; })[0];
    if (!room) return { success: false, message: 'Room not found.' };

    var currentAssignments = getAssignments();
    var fillerCells = getFillerCells();
    var seatMap = {};
    currentAssignments.forEach(function (a) {
      if (a.roomName === roomName) seatMap[a.row + ',' + a.column] = true;
    });
    fillerCells.forEach(function (f) {
      if (f.roomName === roomName) seatMap[f.row + ',' + f.column] = true;
    });

    var sheet = getOrCreateSheet_('Assignments');
    var placed = 0;
    group.studentIds.forEach(function (sid) {
      if (!sid) return;
      removeAssignmentsForStudent_(sid);
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
    var msg = placed + ' student(s) placed in ' + roomName + '.';
    logAudit_('Place Staging Group', 'GID: ' + groupId + ', Room: ' + roomName + ', Stat: ' + msg, true);
    return { success: true, message: msg };
  } catch (e) {
    logAudit_('Place Staging Group', 'GID: ' + groupId + ', Room: ' + roomName, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

// ---------------------------------------------------------------------------
// School Templates
// ---------------------------------------------------------------------------

function getTemplates() {
  ensureSheets_();
  var sheet = getOrCreateSheet_('Templates');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var templates = [];
  for (var i = 1; i < data.length; i++) {
    templates.push({ templateName: String(data[i][0]), data: JSON.parse(data[i][1] || '{}') });
  }
  return templates;
}

function saveTemplate(templateName) {
  try {
    ensureSheets_();
    var teachers = getTeachers();
    var rooms = getRooms();
    var payload = JSON.stringify({ teachers: teachers, rooms: rooms });
    var sheet = getOrCreateSheet_('Templates');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === templateName) {
        sheet.getRange(i + 1, 2).setValue(payload);
        logAudit_('Save Template', 'Name: ' + templateName + ' (Updated)', true);
        return { success: true, message: 'Template "' + templateName + '" updated.' };
      }
    }
    sheet.appendRow([templateName, payload]);
    logAudit_('Save Template', 'Name: ' + templateName + ' (New)', true);
    return { success: true, message: 'Template "' + templateName + '" saved.' };
  } catch (e) {
    logAudit_('Save Template', 'Name: ' + templateName, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function loadTemplate(templateName) {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('Templates');
    var data = sheet.getDataRange().getValues();
    var payload = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === templateName) { payload = JSON.parse(data[i][1] || '{}'); break; }
    }
    if (!payload) return { success: false, message: 'Template not found.' };

    var tSheet = getOrCreateSheet_('Teachers');
    if (tSheet.getLastRow() > 1) tSheet.getRange(2, 1, tSheet.getLastRow() - 1, tSheet.getLastColumn()).clearContent();
    (payload.teachers || []).forEach(function (t) {
      tSheet.appendRow([t.teacherId, t.name, t.roomNumber, t.hallway, t.grade]);
    });

    var rSheet = getOrCreateSheet_('Rooms');
    if (rSheet.getLastRow() > 1) rSheet.getRange(2, 1, rSheet.getLastRow() - 1, rSheet.getLastColumn()).clearContent();
    (payload.rooms || []).forEach(function (r) {
      rSheet.appendRow([
        r.roomName, r.roomNumber || '', r.hallway || '',
        r.rows, r.columns, r.maxCapacity || (r.rows * r.columns),
        r.teacherId || '', r.grade || '', r.floor || '1'
      ]);
    });
    logAudit_('Load Template', 'Name: ' + templateName, true);
    return { success: true, message: 'Template "' + templateName + '" loaded.' };
  } catch (e) {
    logAudit_('Load Template', 'Name: ' + templateName, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function deleteTemplate(templateName) {
  try {
    var sheet = getOrCreateSheet_('Templates');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === templateName) { sheet.deleteRow(i + 1); break; }
    }
    logAudit_('Delete Template', 'Name: ' + templateName, true);
    return { success: true };
  } catch (e) {
    logAudit_('Delete Template', 'Name: ' + templateName, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

// ---------------------------------------------------------------------------
// Auto-Suggest (grade-aware, filler-aware)
// ---------------------------------------------------------------------------

function generateRecommendations(smallGroupLimit, gradeFilter) {
  smallGroupLimit = parseInt(smallGroupLimit, 10) || 10;
  var students = getStudents();
  var rooms = getRooms();
  var fillerCells = getFillerCells();

  if (gradeFilter) {
    students = students.filter(function (s) { return s.grade === gradeFilter; });
    rooms = rooms.filter(function (r) { return !r.grade || r.grade === gradeFilter || r.grade === ''; });
  }

  if (!students.length || !rooms.length) {
    return { error: 'Need at least one student and one room' + (gradeFilter ? ' for grade ' + gradeFilter : '') + '.' };
  }

  // Build filler lookup
  var fillerMap = {};
  fillerCells.forEach(function (f) {
    if (!fillerMap[f.roomName]) fillerMap[f.roomName] = {};
    fillerMap[f.roomName][f.row + ',' + f.column] = true;
  });

  rooms.sort(function (a, b) { return (a.rows * a.columns) - (b.rows * b.columns); });

  var oneToOne = [], readAloud = [], smallGroup = [], general = [], proximityIds = [];
  students.forEach(function (s) {
    if (s.oneToOne) oneToOne.push(s);
    else if (s.readAloud) readAloud.push(s);
    else if (s.smallGroup) smallGroup.push(s);
    else general.push(s);
    if (s.proximity && !s.oneToOne) proximityIds.push(s.studentId);
  });

  var assignments = [];
  var usedSeats = {}, roomStudentCount = {};
  rooms.forEach(function (r) { usedSeats[r.roomName] = {}; roomStudentCount[r.roomName] = 0; });

  function placeStudent(student, roomName) {
    var room = rooms.filter(function (r) { return r.roomName === roomName; })[0];
    if (!room) return false;
    for (var r = 1; r <= room.rows; r++) {
      for (var c = 1; c <= room.columns; c++) {
        var key = r + ',' + c;
        if (usedSeats[roomName][key]) continue;
        if (fillerMap[roomName] && fillerMap[roomName][key]) continue;
        usedSeats[roomName][key] = true;
        roomStudentCount[roomName]++;
        assignments.push({ studentId: student.studentId, roomName: roomName, row: r, column: c });
        return true;
      }
    }
    return false;
  }

  var roomIndex = 0;

  oneToOne.forEach(function (s) {
    while (roomIndex < rooms.length && roomStudentCount[rooms[roomIndex].roomName] > 0) roomIndex++;
    if (roomIndex < rooms.length) { placeStudent(s, rooms[roomIndex].roomName); roomIndex++; }
  });

  if (readAloud.length) {
    var raIdx = roomIndex;
    readAloud.forEach(function (s) {
      if (raIdx >= rooms.length) return;
      var room = rooms[raIdx];
      var fillerCount = fillerMap[room.roomName] ? Object.keys(fillerMap[room.roomName]).length : 0;
      var cap = room.rows * room.columns - fillerCount;
      if (roomStudentCount[room.roomName] >= cap) { raIdx++; if (raIdx >= rooms.length) return; }
      placeStudent(s, rooms[raIdx].roomName);
    });
    if (raIdx >= roomIndex) roomIndex = raIdx;
  }

  if (smallGroup.length) {
    var sgIdx = roomIndex < rooms.length ? roomIndex : rooms.length - 1;
    smallGroup.forEach(function (s) {
      if (sgIdx >= rooms.length) return;
      var room = rooms[sgIdx];
      var fillerCount = fillerMap[room.roomName] ? Object.keys(fillerMap[room.roomName]).length : 0;
      var cap = Math.min(room.rows * room.columns - fillerCount, smallGroupLimit);
      if (roomStudentCount[room.roomName] >= cap) { sgIdx++; if (sgIdx >= rooms.length) return; }
      placeStudent(s, rooms[sgIdx].roomName);
    });
    if (sgIdx >= roomIndex) roomIndex = sgIdx;
  }

  general.forEach(function (s) {
    for (var ri = 0; ri < rooms.length; ri++) {
      var room = rooms[ri];
      var fillerCount = fillerMap[room.roomName] ? Object.keys(fillerMap[room.roomName]).length : 0;
      var cap = room.rows * room.columns - fillerCount;
      if (roomStudentCount[room.roomName] < cap) { if (placeStudent(s, room.roomName)) break; }
    }
  });

  logAudit_('Generate Recommendations', 'Limit: ' + smallGroupLimit + ', Grade: ' + (gradeFilter || 'All'), true);
  return assignments;
}

function applyRecommendations(assignments) {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('Assignments');
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    assignments.forEach(function (a) { sheet.appendRow([a.studentId, a.roomName, a.row, a.column]); });
    logAudit_('Apply Recommendations', 'Count: ' + assignments.length, true);
    return { success: true };
  } catch (e) {
    logAudit_('Apply Recommendations', '', false, e.toString());
    return { success: false, message: e.toString() };
  }
}

// ---------------------------------------------------------------------------
// Finalize / Export Layout to Sheets
// ---------------------------------------------------------------------------

function finalizeLayout() {
  ensureSheets_();
  var students = getStudents();
  var rooms = getRooms();
  var teachers = getTeachers();
  var assignments = getAssignments();
  var fillerCells = getFillerCells();
  var ss = getSpreadsheet_();

  var studentMap = {};
  students.forEach(function (s) { studentMap[s.studentId] = s; });
  var teacherMap = {};
  teachers.forEach(function (t) { teacherMap[t.teacherId] = t; });
  var fillerMap = {};
  fillerCells.forEach(function (f) {
    if (!fillerMap[f.roomName]) fillerMap[f.roomName] = {};
    fillerMap[f.roomName][f.row + ',' + f.column] = true;
  });

  rooms.forEach(function (room) {
    var sheetName = 'Layout - ' + room.roomName;
    var existing = ss.getSheetByName(sheetName);
    if (existing) ss.deleteSheet(existing);
    var layoutSheet = ss.insertSheet(sheetName);

    var grid = [];
    for (var r = 0; r < room.rows; r++) {
      var row = [];
      for (var c = 0; c < room.columns; c++) {
        var key = (r + 1) + ',' + (c + 1);
        if (fillerMap[room.roomName] && fillerMap[room.roomName][key]) {
          row.push('[BLOCKED]');
        } else {
          row.push('');
        }
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
      var ri = a.row - 1, ci = a.column - 1;
      if (ri >= 0 && ri < room.rows && ci >= 0 && ci < room.columns) grid[ri][ci] = label;
    });

    if (room.rows > 0 && room.columns > 0) {
      layoutSheet.getRange(1, 1, room.rows, room.columns).setValues(grid);
      // Color Read Aloud students red
      for (var r = 0; r < room.rows; r++) {
        for (var c = 0; c < room.columns; c++) {
          if (grid[r][c] && grid[r][c].indexOf('RA') !== -1) {
            layoutSheet.getRange(r + 1, c + 1).setFontColor('#cc0000').setFontWeight('bold');
          }
          if (grid[r][c] === '[BLOCKED]') {
            layoutSheet.getRange(r + 1, c + 1).setBackground('#d1d5db');
          }
        }
      }
    }

    layoutSheet.insertRowBefore(1);
    var teacher = teacherMap[room.teacherId];
    var title = room.roomName;
    if (room.roomNumber) title += ' (Rm ' + room.roomNumber + ')';
    if (room.hallway) title += ' — ' + room.hallway;
    if (room.floor) title += ' — Floor ' + room.floor;
    if (teacher) title += ' — ' + teacher.name;
    layoutSheet.getRange(1, 1).setValue(title);
    layoutSheet.getRange(1, 1).setFontWeight('bold').setFontSize(12).setFontColor('#111111');

    layoutSheet.insertRowAfter(1);
    layoutSheet.getRange(2, 1).setValue('[ PROCTOR STATION — Row 1 ]');
    layoutSheet.getRange(2, 1).setFontStyle('italic');
  });

  logAudit_('Finalize Layout', 'Rooms: ' + rooms.length, true);
  return { success: true, message: 'Layout sheets created for ' + rooms.length + ' room(s).' };
}

// ---------------------------------------------------------------------------
// Export as HTML (for PDF/Word download in client)
// ---------------------------------------------------------------------------

/**
 * Build a concise HTML document suitable for printing / saving as PDF or Word.
 * Read Aloud students are rendered in RED.
 * Organized by grade -> hallway -> floor -> room.
 */
function generateExportHTML() {
  ensureSheets_();
  var students = getStudents();
  var rooms = getRooms();
  var teachers = getTeachers();
  var assignments = getAssignments();
  var fillerCells = getFillerCells();

  var studentMap = {};
  students.forEach(function (s) { studentMap[s.studentId] = s; });
  var teacherMap = {};
  teachers.forEach(function (t) { teacherMap[t.teacherId] = t; });
  var fillerMap = {};
  fillerCells.forEach(function (f) {
    if (!fillerMap[f.roomName]) fillerMap[f.roomName] = {};
    fillerMap[f.roomName][f.row + ',' + f.column] = true;
  });

  // Group rooms by grade, then hallway, then floor
  var gradeOrder = ['6', '7', '8', ''];
  var roomsByGrade = {};
  rooms.forEach(function (r) {
    var g = r.grade || 'Shared';
    if (!roomsByGrade[g]) roomsByGrade[g] = [];
    roomsByGrade[g].push(r);
  });

  var html = [];
  html.push('<!DOCTYPE html><html><head><meta charset="utf-8">');
  html.push('<title>Testing Assignments</title>');
  html.push('<style>');
  html.push('body{font-family:Arial,sans-serif;font-size:11px;color:#1f2937;margin:20px;}');
  html.push('h1{color:#111;font-size:18px;border-bottom:3px solid #9ca3af;padding-bottom:4px;margin-bottom:8px;}');
  html.push('h2{color:#111;font-size:14px;margin:12px 0 4px;}');
  html.push('h3{font-size:12px;margin:8px 0 3px;color:#333;}');
  html.push('.ra{color:#cc0000;font-weight:bold;}');
  html.push('table{border-collapse:collapse;width:100%;margin-bottom:10px;}');
  html.push('td,th{border:1px solid #999;padding:3px 5px;text-align:center;font-size:10px;}');
  html.push('th{background:#f3f4f6;}');
  html.push('.blocked{background:#d1d5db;}');
  html.push('.summary-list{margin:4px 0 10px 16px;font-size:11px;line-height:1.4;}');
  html.push('.floor-label{background:#e5e7eb;padding:2px 8px;border-radius:3px;font-weight:bold;display:inline-block;margin:6px 0 3px;}');
  html.push('@media print{h1{page-break-before:avoid;}h2{page-break-before:always;}.no-break{page-break-inside:avoid;}}');
  html.push('</style></head><body>');
  html.push('<h1>SOL Testing Room Assignments</h1>');

  gradeOrder.forEach(function (g) {
    var gradeKey = g || 'Shared';
    var gradeRooms = roomsByGrade[gradeKey];
    if (!gradeRooms || !gradeRooms.length) return;

    var gradeLabel = g ? 'Grade ' + g : 'Shared / Elective Rooms';
    html.push('<h2>' + gradeLabel + '</h2>');

    // Group by floor
    var byFloor = {};
    gradeRooms.forEach(function (r) {
      var fl = r.floor || '1';
      if (!byFloor[fl]) byFloor[fl] = [];
      byFloor[fl].push(r);
    });

    ['1', '2'].forEach(function (fl) {
      if (!byFloor[fl]) return;
      html.push('<span class="floor-label">Floor ' + fl + '</span>');

      byFloor[fl].forEach(function (room) {
        var teacher = teacherMap[room.teacherId];
        var roomLabel = room.roomName;
        if (room.roomNumber) roomLabel += ' (Rm ' + room.roomNumber + ')';
        if (room.hallway) roomLabel += ' — ' + room.hallway;
        if (teacher) roomLabel += ' — ' + teacher.name;

        html.push('<div class="no-break">');
        html.push('<h3>' + roomLabel + '</h3>');

        // Grid table
        var roomAssign = assignments.filter(function (a) { return a.roomName === room.roomName; });
        var seatMap = {};
        roomAssign.forEach(function (a) { seatMap[a.row + ',' + a.column] = a; });

        html.push('<table>');
        for (var r = 1; r <= room.rows; r++) {
          html.push('<tr>');
          for (var c = 1; c <= room.columns; c++) {
            var key = r + ',' + c;
            if (fillerMap[room.roomName] && fillerMap[room.roomName][key]) {
              html.push('<td class="blocked"></td>');
            } else if (seatMap[key]) {
              var s = studentMap[seatMap[key].studentId];
              if (s) {
                var isRA = s.readAloud;
                var codes = buildAccommodationCodes_(s);
                var cellText = s.name;
                if (codes) cellText += ' (' + codes + ')';
                html.push('<td' + (isRA ? ' class="ra"' : '') + '>' + cellText + '</td>');
              } else {
                html.push('<td></td>');
              }
            } else {
              html.push('<td></td>');
            }
          }
          html.push('</tr>');
        }
        html.push('</table>');

        // Text summary for this room
        if (roomAssign.length) {
          html.push('<div class="summary-list">');
          roomAssign.forEach(function (a) {
            var s = studentMap[a.studentId];
            if (!s) return;
            var isRA = s.readAloud;
            var codes = buildAccommodationCodes_(s);
            var line = s.name + ' (Gr ' + s.grade + ')';
            if (codes) line += ' — ' + codes;
            line += ' → Seat R' + a.row + 'C' + a.column;
            if (isRA) {
              html.push('<div class="ra">' + line + '</div>');
            } else {
              html.push('<div>' + line + '</div>');
            }
          });
          html.push('</div>');
        }

        html.push('</div>');
      });
    });
  });

  html.push('</body></html>');
  logAudit_('Generate Export', '', true);
  return html.join('');
}

// ---------------------------------------------------------------------------
// Load Example Data (all views)
// ---------------------------------------------------------------------------

function loadExampleData() {
  ensureSheets_();

  // --- Example Teachers ---
  var tSheet = getOrCreateSheet_('Teachers');
  var existingTeachers = tSheet.getDataRange().getValues();
  if (existingTeachers.length <= 1) {
    var teachers = [
      ['T001', 'Mrs. Sarah Johnson',    '101', '6th Grade Hall',  '6'],
      ['T002', 'Mr. David Williams',    '102', '6th Grade Hall',  '6'],
      ['T003', 'Ms. Maria Garcia',      '103', '6th Grade Hall',  '6'],
      ['T004', 'Mr. James Brown',       '201', '7th Grade Hall',  '7'],
      ['T005', 'Mrs. Lisa Davis',        '202', '7th Grade Hall',  '7'],
      ['T006', 'Mr. Robert Miller',     '203', '7th Grade Hall',  '7'],
      ['T007', 'Mrs. Patricia Wilson',  '301', '8th Grade Hall',  '8'],
      ['T008', 'Mr. Michael Taylor',    '302', '8th Grade Hall',  '8'],
      ['T009', 'Ms. Jennifer Anderson', 'LIB', 'Main Hall',       ''],
      ['T010', 'Coach Thomas Martinez', 'GYM', 'Athletics Wing',  '']
    ];
    teachers.forEach(function (t) { tSheet.appendRow(t); });
  }

  // --- Example Rooms ---
  var rSheet = getOrCreateSheet_('Rooms');
  var existingRooms = rSheet.getDataRange().getValues();
  if (existingRooms.length <= 1) {
    var rooms = [
      // RoomName, RoomNumber, Hallway, Rows, Cols, MaxCap, TeacherID, Grade, Floor
      ['Room 101',      '101', '6th Grade Hall',  5, 6, 30, 'T001', '6', '1'],
      ['Room 102',      '102', '6th Grade Hall',  5, 6, 30, 'T002', '6', '1'],
      ['Room 103',      '103', '6th Grade Hall',  5, 6, 30, 'T003', '6', '1'],
      ['Library',       'LIB', 'Main Hall',       6, 8, 48, 'T009', '',  '1'],
      ['Computer Lab',  '110', 'Main Hall',       5, 6, 30, '',     '',  '1'],
      ['Gymnasium',     'GYM', 'Athletics Wing',  8, 10,80, 'T010', '',  '1'],
      ['Room 201',      '201', '7th Grade Hall',  5, 6, 30, 'T004', '7', '2'],
      ['Room 202',      '202', '7th Grade Hall',  5, 6, 30, 'T005', '7', '2'],
      ['Room 203',      '203', '7th Grade Hall',  5, 6, 30, 'T006', '7', '2'],
      ['Room 301',      '301', '8th Grade Hall',  6, 6, 36, 'T007', '8', '2'],
      ['Room 302',      '302', '8th Grade Hall',  6, 6, 36, 'T008', '8', '2'],
      ['Science Lab',   '310', '8th Grade Hall',  5, 6, 30, '',     '',  '2'],
      ['Art Room',      '220', 'Electives Wing',  5, 5, 25, '',     '',  '2'],
      ['Music Room',    '221', 'Electives Wing',  6, 6, 36, '',     '',  '2']
    ];
    rooms.forEach(function (r) { rSheet.appendRow(r); });
  }

  // --- Example Students (fictional names, various accommodations) ---
  var sSheet = getOrCreateSheet_('Students');
  var existingStudents = sSheet.getDataRange().getValues();
  if (existingStudents.length <= 1) {
    var students = [
      // StudentID, Name, Grade, SG, RA, 1:1, PROX, PMPT, Other
      ['S1001', 'Alex Thompson',   '6', 'Y', '',  '',  '',  '',  ''],
      ['S1002', 'Jordan Lee',      '6', '',  'Y', '',  '',  '',  ''],
      ['S1003', 'Casey Smith',     '6', '',  '',  '',  'Y', '',  ''],
      ['S1004', 'Riley Johnson',   '6', '',  '',  '',  '',  '',  ''],
      ['S1005', 'Morgan White',    '6', '',  '',  '',  '',  'Y', ''],
      ['S1006', 'Taylor Clark',    '6', 'Y', 'Y', '',  '',  '',  ''],
      ['S1007', 'Blake Adams',     '6', '',  '',  '',  '',  '',  ''],
      ['S1008', 'Avery Nelson',    '6', '',  '',  '',  '',  '',  'Extended Time'],
      ['S2001', 'Drew Martin',     '7', '',  '',  'Y', '',  '',  ''],
      ['S2002', 'Sam Rodriguez',   '7', '',  '',  '',  '',  '',  ''],
      ['S2003', 'Chris Lopez',     '7', '',  '',  '',  'Y', 'Y', ''],
      ['S2004', 'Cameron Wright',  '7', 'Y', '',  '',  '',  '',  ''],
      ['S2005', 'Quinn Harris',    '7', '',  'Y', '',  '',  '',  ''],
      ['S2006', 'Dakota Perez',    '7', '',  '',  '',  '',  '',  ''],
      ['S2007', 'Skyler Thomas',   '7', 'Y', 'Y', '',  'Y', '',  ''],
      ['S3001', 'Jamie Walker',    '8', '',  '',  '',  '',  '',  'Extended Time'],
      ['S3002', 'Pat Gonzalez',    '8', 'Y', '',  '',  '',  '',  ''],
      ['S3003', 'Reese Campbell',  '8', '',  'Y', '',  'Y', '',  ''],
      ['S3004', 'Finley Scott',    '8', '',  '',  '',  '',  '',  ''],
      ['S3005', 'Emerson Hill',    '8', '',  '',  'Y', '',  '',  ''],
      ['S3006', 'Harper Young',    '8', '',  '',  '',  '',  'Y', ''],
      ['S3007', 'Rowan King',      '8', 'Y', '',  '',  '',  '',  ''],
      ['S3008', 'Sage Allen',      '8', '',  '',  '',  '',  '',  '']
    ];
    students.forEach(function (s) { sSheet.appendRow(s); });
  }

  return { success: true, message: 'Example data loaded — 10 teachers, 14 rooms, 23 students with accommodations.' };
}

// ---------------------------------------------------------------------------
// Designer Layout CRUD
// ---------------------------------------------------------------------------

function saveDesignerLayout(name, json) {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('DesignerLayouts');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === name) {
        sheet.getRange(i + 1, 2).setValue(json);
        logAudit_('Save Designer Layout', 'Name: ' + name + ' (Updated)', true);
        return { success: true, message: 'Layout "' + name + '" updated.' };
      }
    }
    sheet.appendRow([name, json]);
    logAudit_('Save Designer Layout', 'Name: ' + name + ' (New)', true);
    return { success: true, message: 'Layout "' + name + '" saved.' };
  } catch (e) {
    logAudit_('Save Designer Layout', 'Name: ' + name, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function getDesignerLayouts() {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('DesignerLayouts');
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var layouts = [];
    for (var i = 1; i < data.length; i++) {
      var rooms = JSON.parse(data[i][1] || '[]');
      layouts.push({
        layoutName: String(data[i][0]),
        roomCount: rooms.length
      });
    }
    logAudit_('Get Designer Layouts', '', true);
    return layouts;
  } catch (e) {
    logAudit_('Get Designer Layouts', '', false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function loadDesignerLayout(name) {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('DesignerLayouts');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === name) {
        logAudit_('Load Designer Layout', 'Name: ' + name, true);
        return { success: true, rooms: JSON.parse(data[i][1] || '[]') };
      }
    }
    logAudit_('Load Designer Layout', 'Name: ' + name, false, 'Layout not found.');
    return { success: false, message: 'Layout not found.' };
  } catch (e) {
    logAudit_('Load Designer Layout', 'Name: ' + name, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function deleteDesignerLayout(name) {
  try {
    var sheet = getOrCreateSheet_('DesignerLayouts');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === name) { sheet.deleteRow(i + 1); break; }
    }
    logAudit_('Delete Designer Layout', 'Name: ' + name, true);
    return { success: true };
  } catch (e) {
    logAudit_('Delete Designer Layout', 'Name: ' + name, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

// ---------------------------------------------------------------------------
// Full Backup / Restore
// ---------------------------------------------------------------------------

function createFullBackup(name) {
  try {
    ensureSheets_();
    var students = getStudents();
    var teachers = getTeachers();
    var rooms = getRooms();
    var assignments = getAssignments();
    var staging = getStagingGroups();
    var filler = getFillerCells();

    var payload = JSON.stringify({
      students: students, teachers: teachers, rooms: rooms,
      assignments: assignments, stagingGroups: staging, fillerCells: filler
    });

    var sheet = getOrCreateSheet_('Backups');
    var data = sheet.getDataRange().getValues();
    var now = new Date().toISOString();
    var updated = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === name) {
        sheet.getRange(i + 1, 2, 1, 2).setValues([[now, payload]]);
        updated = true;
        break;
      }
    }
    if (updated) {
      logAudit_('Create Backup', 'Name: ' + name + ' (Updated)', true);
      return { success: true, message: 'Backup "' + name + '" updated.' };
    } else {
      sheet.appendRow([name, now, payload]);
      logAudit_('Create Backup', 'Name: ' + name + ' (New)', true);
      return { success: true, message: 'Backup "' + name + '" saved.' };
    }
  } catch (e) {
    logAudit_('Create Backup', 'Name: ' + name, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function getFullBackups() {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('Backups');
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var backups = [];
    for (var i = 1; i < data.length; i++) {
      var payload = JSON.parse(data[i][2] || '{}');
      backups.push({
        backupName: String(data[i][0]),
        createdAt: String(data[i][1]),
        studentCount: (payload.students || []).length,
        teacherCount: (payload.teachers || []).length,
        roomCount: (payload.rooms || []).length
      });
    }
    logAudit_('Get Backups', '', true);
    return backups;
  } catch (e) {
    logAudit_('Get Backups', '', false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function restoreFullBackup(name) {
  try {
    ensureSheets_();
    var sheet = getOrCreateSheet_('Backups');
    var data = sheet.getDataRange().getValues();
    var payload = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === name) { payload = JSON.parse(data[i][2] || '{}'); break; }
    }
    if (!payload) {
      logAudit_('Restore Backup', 'Name: ' + name, false, 'Backup not found.');
      return { success: false, message: 'Backup not found.' };
    }

    // Clear and restore each sheet
    // Students
    var sSheet = getOrCreateSheet_('Students');
    if (sSheet.getLastRow() > 1) sSheet.getRange(2, 1, sSheet.getLastRow() - 1, sSheet.getLastColumn()).clearContent();
    (payload.students || []).forEach(function (s) {
      sSheet.appendRow([
        s.studentId, s.name, s.grade,
        s.smallGroup ? 'Y' : '', s.readAloud ? 'Y' : '',
        s.oneToOne ? 'Y' : '', s.proximity ? 'Y' : '',
        s.prompting ? 'Y' : '', s.otherAccommodations || ''
      ]);
    });

    // Teachers
    var tSheet = getOrCreateSheet_('Teachers');
    if (tSheet.getLastRow() > 1) tSheet.getRange(2, 1, tSheet.getLastRow() - 1, tSheet.getLastColumn()).clearContent();
    (payload.teachers || []).forEach(function (t) {
      tSheet.appendRow([t.teacherId, t.name, t.roomNumber, t.hallway, t.grade]);
    });

    // Rooms
    var rSheet = getOrCreateSheet_('Rooms');
    if (rSheet.getLastRow() > 1) rSheet.getRange(2, 1, rSheet.getLastRow() - 1, rSheet.getLastColumn()).clearContent();
    (payload.rooms || []).forEach(function (r) {
      rSheet.appendRow([
        r.roomName, r.roomNumber || '', r.hallway || '',
        r.rows, r.columns, r.maxCapacity || (r.rows * r.columns),
        r.teacherId || '', r.grade || '', r.floor || '1'
      ]);
    });

    // Assignments
    var aSheet = getOrCreateSheet_('Assignments');
    if (aSheet.getLastRow() > 1) aSheet.getRange(2, 1, aSheet.getLastRow() - 1, aSheet.getLastColumn()).clearContent();
    (payload.assignments || []).forEach(function (a) {
      aSheet.appendRow([a.studentId, a.roomName, a.row, a.column]);
    });

    // Staging
    var stSheet = getOrCreateSheet_('Staging');
    if (stSheet.getLastRow() > 1) stSheet.getRange(2, 1, stSheet.getLastRow() - 1, stSheet.getLastColumn()).clearContent();
    (payload.stagingGroups || []).forEach(function (g) {
      if (!g.studentIds || !g.studentIds.length) {
        stSheet.appendRow([g.groupId, g.groupName, '']);
      } else {
        g.studentIds.forEach(function (sid) { stSheet.appendRow([g.groupId, g.groupName, sid]); });
      }
    });

    // Filler Cells
    var fSheet = getOrCreateSheet_('FillerCells');
    if (fSheet.getLastRow() > 1) fSheet.getRange(2, 1, fSheet.getLastRow() - 1, fSheet.getLastColumn()).clearContent();
    (payload.fillerCells || []).forEach(function (f) {
      fSheet.appendRow([f.roomName, f.row, f.column]);
    });

    logAudit_('Restore Backup', 'Name: ' + name, true);
    return { success: true, message: 'Backup "' + name + '" restored — ' +
      (payload.students || []).length + ' students, ' +
      (payload.teachers || []).length + ' teachers, ' +
      (payload.rooms || []).length + ' rooms.' };
  } catch (e) {
    logAudit_('Restore Backup', 'Name: ' + name, false, e.toString());
    return { success: false, message: e.toString() };
  }
}

function deleteFullBackup(name) {
  var sheet = getOrCreateSheet_('Backups');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === name) { sheet.deleteRow(i + 1); break; }
  }
  return { success: true };
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
