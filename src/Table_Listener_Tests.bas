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
Private watcher As EventsWatcher

Private Property Get logger() As LoggingEventSink
    Set logger = watcher.logger
End Property

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
    Set watcher = New EventsWatcher
    Set watcher.logger = New LoggingEventSink
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set watcher = Nothing
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

    Set watcher.events = table


    '@Ignore IndexedDefaultMemberAccess
    srcTable.ListColumns(srcTable.ListColumns.Count).Range.Offset(0, 1).Insert

    Assert.AreEqual 0, logger.log.Count, "too many events raised"
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

    Set watcher.events = table


    Dim newRow As ListRow
    Set newRow = srcTable.ListRows.Add
    Assert.AreEqual idRowAdded, logger.EventClasses, "Only 1 kind of event should have been raised"
    Assert.AreEqual 1, logger.logEntry(idRowAdded).Count, "Count wrong"
    AreListRowsSame Assert, newRow, logger.logEntry(idRowAdded).Item(1)

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

    Set watcher.events = table

    Dim newRow As ListRow
    Set newRow = srcTable.ListRows.Add(srcTable.ListRows.Count \ 2)
    Assert.AreEqual idRowAdded, logger.EventClasses, "Only 1 kind of event should have been raised"
    Assert.AreEqual 1, logger.logEntry(idRowAdded).Count, "Count wrong"
    AreListRowsSame Assert, newRow, logger.logEntry(idRowAdded).Item(1)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Events")
Private Sub TestImplicitAppendRowAtEndOfDatabody()
    On Error GoTo TestFail
    Dim table As TableWatcher
    Set table = TableWatcher.Create(srcTable)

    Set watcher.events = table

    Dim newRowTrigger As Range
    Set newRowTrigger = srcTable.DataBodyRange.Cells(srcTable.ListRows.Count + 1, 1)
    
    newRowTrigger.Value2 = "Foo"
    
    Dim newRow As ListRow
    Set newRow = ListObjectHelperMethods.TargetToListRow(srcTable, newRowTrigger)
    
    Assert.AreEqual idRowAdded, logger.EventClasses, "Wrong kind/ too many kinds of event raised"
    Assert.AreEqual 1, logger.logEntry(idRowAdded).Count, "Count wrong"
    AreListRowsSame Assert, newRow, logger.logEntry(idRowAdded).Item(1)

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

    Set watcher.events = table

    Dim newCol As ListColumn
    Set newCol = srcTable.ListColumns.Add
    Assert.AreEqual idColAdded + idColNameChange, logger.EventClasses, "Wrong event types raised"
    Assert.AreEqual 1, logger.logEntry(idColAdded).Count, " Col add count wrong"
    Assert.AreEqual 1, logger.logEntry(idColNameChange).Count, "Name change count wrong"
    AreListColumnsSame Assert, newCol, logger.logEntry(idColAdded).Item(1)
    '@Ignore IndexedDefaultMemberAccess
    AreRangesSame Assert, newCol.Range.Cells(1), logger.logEntry(idColNameChange).Item(1)

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

    Set watcher.events = table

    Dim newCol As ListColumn
    Set newCol = srcTable.ListColumns.Add(srcTable.ListColumns.Count \ 2)
    Assert.AreEqual idColAdded + idColNameChange, logger.EventClasses, "Wrong event types raised"
    Assert.AreEqual 1, logger.logEntry(idColAdded).Count, " Col add count wrong"
    Assert.AreEqual 1, logger.logEntry(idColNameChange).Count, "Name change count wrong"
    AreListColumnsSame Assert, newCol, logger.logEntry(idColAdded).Item(1)
    '@Ignore IndexedDefaultMemberAccess
    AreRangesSame Assert, newCol.Range.Cells(1), logger.logEntry(idColNameChange).Item(1)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Events")
Private Sub TestImplicitAddColumnToRightEdge()
    On Error GoTo TestFail
    Dim table As TableWatcher
    Set table = TableWatcher.Create(srcTable)

    Set watcher.events = table

    Dim newColTrigger As Range
    Set newColTrigger = srcTable.DataBodyRange.Cells(srcTable.ListRows.Count \ 2, srcTable.ListColumns.Count + 1)
    
    newColTrigger.Value2 = "Foo"
    
    Dim newCol As ListColumn
    Set newCol = ListObjectHelperMethods.TargetToListColumn(srcTable, newColTrigger)
    
    Assert.AreEqual idColAdded + idColNameChange, logger.EventClasses, "Wrong event types raised"
    Assert.AreEqual 1, logger.logEntry(idColAdded).Count, " Col add count wrong"
    Assert.AreEqual 1, logger.logEntry(idColNameChange).Count, "Name change count wrong"
    AreListColumnsSame Assert, newCol, logger.logEntry(idColAdded).Item(1)
    '@Ignore IndexedDefaultMemberAccess
    AreRangesSame Assert, newCol.Range.Cells(1), logger.logEntry(idColNameChange).Item(1)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

