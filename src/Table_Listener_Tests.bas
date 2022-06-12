Attribute VB_Name = "Table_Listener_Tests"
'@IgnoreModule SelfAssignedDeclaration
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.PermissiveAssertClass
'@Ignore VariableNotUsed
Private Fakes As Rubberduck.FakesProvider
Private srcTable As ListObject

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    TestSheet.Reset
    Set Assert = New Rubberduck.PermissiveAssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set srcTable = TestSheet.DemoTable
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set srcTable = Nothing
    TestSheet.Reset
End Sub

'@TestMethod("Uncategorized")
Private Sub TestAddInstance()
    On Error GoTo TestFail
    Dim table As TableWatcher
    Set table = TableWatcher.Create(srcTable)
    Assert.AreSame srcTable, table.WrappedTable
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestDeleteTable()
    On Error GoTo TestFail
    Dim table As TableWatcher
    Set table = TableWatcher.Create(srcTable)
    srcTable.Delete
    Assert.IsNothing table.WrappedTable
    'Assert.IsNothing table.WrappedTable 'ensure it stays deleted
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestNullTableReference()
    On Error GoTo TestFail
    Dim table As TableWatcher
    Set table = TableWatcher.Create(srcTable)
    Set table.WrappedTable = Nothing
    Assert.IsNothing table.WrappedTable
    Assert.IsNothing table.wrappedTableParent

    'Assert.IsNothing table.WrappedTable 'ensure it stays deleted
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestDelete_RestoreTable()
    On Error GoTo TestFail
    Dim table As TableWatcher
    Set table = TableWatcher.Create(srcTable)
    srcTable.Delete
    Assert.IsNothing table.WrappedTable
    TestSheet.Reset
    Set srcTable = TestSheet.DemoTable
    Assert.IsNothing table.WrappedTable          'ensure it stays deleted
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Events")
Private Sub TestGetEventsObject()
    On Error GoTo TestFail
    Dim table As ITableEventsSource
    Set table = TableWatcher.Create(srcTable)
    Assert.IsNotNothing table
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Events")
Private Sub TestUnrelatedChangeInWorksheetHasNoEffect()
    On Error GoTo TestFail
    Dim table As TableWatcher
    Set table = TableWatcher.Create(srcTable)

    Dim counter As New EventsCounter
    Set counter.events = table

    '@Ignore IndexedDefaultMemberAccess
    srcTable.ListColumns(srcTable.ListColumns.Count).Range.Offset(0, 1).Insert

    Assert.AreEqual 0, counter.Log.Count, "too many events raised"
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Events")
Private Sub TestAddRow()
    On Error GoTo TestFail
    Dim table As TableWatcher
    Set table = TableWatcher.Create(srcTable)

    Dim counter As New EventsCounter
    Set counter.events = table

    Dim newRow As ListRow
    Set newRow = srcTable.ListRows.Add
    Assert.AreEqual idRowAdded, counter.EventClasses, "Only 1 kind of event should have been raised"
    Assert.AreEqual 1, counter.logEntry(idRowAdded).Count, "Count wrong"
    AreListRowsSame Assert, newRow, counter.logEntry(idRowAdded).Item(1)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Events")
Private Sub TestInsertRow()
    On Error GoTo TestFail
    Dim table As TableWatcher
    Set table = TableWatcher.Create(srcTable)

    Dim counter As New EventsCounter
    Set counter.events = table

    Dim newRow As ListRow
    Set newRow = srcTable.ListRows.Add(srcTable.ListRows.Count \ 2)
    Assert.AreEqual idRowAdded, counter.EventClasses, "Only 1 kind of event should have been raised"
    Assert.AreEqual 1, counter.logEntry(idRowAdded).Count, "Count wrong"
    AreListRowsSame Assert, newRow, counter.logEntry(idRowAdded).Item(1)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Events")
Private Sub TestColAdded()
    On Error GoTo TestFail
    Dim table As TableWatcher
    Set table = TableWatcher.Create(srcTable)

    Dim counter As New EventsCounter
    Set counter.events = table

    Dim newCol As ListColumn
    Set newCol = srcTable.ListColumns.Add
    Assert.AreEqual idColAdded + idColNameChange, counter.EventClasses, "Wrong event types raised"
    Assert.AreEqual 1, counter.logEntry(idColAdded).Count, " Col add count wrong"
    Assert.AreEqual 1, counter.logEntry(idColNameChange).Count, "Name change count wrong"
    AreListColumnsSame Assert, newCol, counter.logEntry(idColAdded).Item(1)
    '@Ignore IndexedDefaultMemberAccess
    AreRangesSame Assert, newCol.Range.Cells(1), counter.logEntry(idColNameChange).Item(1)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Events")
Private Sub TestInsertCol()
    On Error GoTo TestFail
    Dim table As TableWatcher
    Set table = TableWatcher.Create(srcTable)

    Dim counter As New EventsCounter
    Set counter.events = table

    Dim newCol As ListColumn
    Set newCol = srcTable.ListColumns.Add(srcTable.ListColumns.Count \ 2)
    Assert.AreEqual idColAdded + idColNameChange, counter.EventClasses, "Wrong event types raised"
    Assert.AreEqual 1, counter.logEntry(idColAdded).Count, " Col add count wrong"
    Assert.AreEqual 1, counter.logEntry(idColNameChange).Count, "Name change count wrong"
    AreListColumnsSame Assert, newCol, counter.logEntry(idColAdded).Item(1)
    '@Ignore IndexedDefaultMemberAccess
    AreRangesSame Assert, newCol.Range.Cells(1), counter.logEntry(idColNameChange).Item(1)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

