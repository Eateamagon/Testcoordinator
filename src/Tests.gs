/**
 * TEST SUITE FOR TESTCOORDINATOR
 * Run runAllTests() to execute all tests.
 * Results are logged to Logger and the AuditLog.
 */

function runAllTests() {
  Logger.log('--- STARTING TEST SUITE ---');
  
  testStudentCRUD();
  testTeacherCRUD();
  testRoomCRUD();
  testStagingAndAssignments();
  
  Logger.log('--- TEST SUITE COMPLETE ---');
}

function assert_(condition, message) {
  if (!condition) {
    throw new Error('ASSERTION FAILED: ' + message);
  }
}

function testStudentCRUD() {
  Logger.log('Testing Student CRUD...');
  var testId = 'TEST_S_001';
  
  // Cleanup if exists
  deleteStudent(testId);
  
  // Add
  var res = addStudent({
    studentId: testId, name: 'Test Student', grade: '5',
    smallGroup: true, readAloud: false
  });
  assert_(res.success, 'addStudent failed');
  
  // Get
  var students = getStudents();
  var found = students.find(s => s.studentId === testId);
  assert_(found, 'Student not found in list');
  assert_(found.smallGroup === true, 'Accommodation mismatch');
  
  // Update
  res = updateStudent({
    studentId: testId, name: 'Updated Student', grade: '5',
    smallGroup: false, readAloud: true
  });
  assert_(res.success, 'updateStudent failed');
  
  students = getStudents();
  found = students.find(s => s.studentId === testId);
  assert_(found.name === 'Updated Student', 'Update name failed');
  assert_(found.readAloud === true, 'Update accommodation failed');
  
  // Delete
  res = deleteStudent(testId);
  assert_(res.success, 'deleteStudent failed');
  
  students = getStudents();
  found = students.find(s => s.studentId === testId);
  assert_(!found, 'Student should be deleted');
  
  Logger.log('Student CRUD tests PASSED.');
}

function testTeacherCRUD() {
  Logger.log('Testing Teacher CRUD...');
  var testId = 'TEST_T_001';
  
  deleteTeacher(testId);
  
  var res = addTeacher({
    teacherId: testId, name: 'Test Teacher', roomNumber: '101', hallway: 'A', grade: '4'
  });
  assert_(res.success, 'addTeacher failed');
  
  var teachers = getTeachers();
  var found = teachers.find(t => t.teacherId === testId);
  assert_(found, 'Teacher not found');
  
  res = deleteTeacher(testId);
  assert_(res.success, 'deleteTeacher failed');
  
  Logger.log('Teacher CRUD tests PASSED.');
}

function testRoomCRUD() {
  Logger.log('Testing Room CRUD...');
  var name = 'TEST_ROOM_99';
  
  deleteRoom(name);
  
  var res = addRoom({
    roomName: name, roomNumber: '99', rows: 5, columns: 5, maxCapacity: 20
  });
  assert_(res.success, 'addRoom failed');
  
  var rooms = getRooms();
  var found = rooms.find(r => r.roomName === name);
  assert_(found, 'Room not found');
  assert_(found.maxCapacity === 20, 'Room capacity mismatch');
  
  res = deleteRoom(name);
  assert_(res.success, 'deleteRoom failed');
  
  Logger.log('Room CRUD tests PASSED.');
}

function testStagingAndAssignments() {
  Logger.log('Testing Staging and Assignments...');
  var sid = 'TEST_SA_001';
  var gid = null;
  var rname = 'TEST_SA_ROOM';
  
  // Cleanup
  deleteStudent(sid);
  deleteRoom(rname);
  
  addStudent({ studentId: sid, name: 'SA Student', grade: '3' });
  addRoom({ roomName: rname, rows: 2, columns: 2 });
  
  // Staging
  var res = createStagingGroup('Test SA Group');
  assert_(res.success, 'createStagingGroup failed');
  gid = res.groupId;
  
  res = addStudentToStagingGroup(gid, sid);
  assert_(res.success, 'addStudentToStagingGroup failed');
  
  // Placement
  res = placeStagingGroupInRoom(gid, rname);
  assert_(res.success, 'placeStagingGroupInRoom failed');
  
  var assignments = getAssignments();
  var found = assignments.find(a => a.studentId === sid && a.roomName === rname);
  assert_(found, 'Assignment not found after placement');
  
  // Cleanup
  deleteStagingGroup(gid);
  deleteStudent(sid);
  deleteRoom(rname);
  
  Logger.log('Staging and Assignment tests PASSED.');
}
