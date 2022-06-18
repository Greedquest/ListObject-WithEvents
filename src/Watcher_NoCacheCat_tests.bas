Attribute VB_Name = "Watcher_NoCacheCat_tests"
'@IgnoreModule SelfAssignedDeclaration
Option Explicit
Option Private Module

'@TestModule
'@Folder "tests"

Private Assert As Rubberduck.PermissiveAssertClass
'@Ignore VariableNotUsed
Private Fakes As Rubberduck.FakesProvider
Private srcTable As ListObject
Private watcher As EventsWatcher
Private table As TableWatcher

Private Property Get logger() As LoggingEventSink
    Set logger = watcher.logger
End Property

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Application.ScreenUpdating = False
    TestSheet.ResetTable
    Set Assert = New Rubberduck.PermissiveAssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Application.ScreenUpdating = True
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set srcTable = TestSheet.DemoTable
    Set watcher = New EventsWatcher
    Set watcher.logger = New LoggingEventSink
    Set table = TableWatcher.Create(srcTable, New NoCacheCategoriser)
    Set watcher.events = table
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set watcher = Nothing
    Set srcTable = Nothing
    TestSheet.ResetTable
End Sub

'@TestMethod("Object")
Private Sub TestAddInstance()
    On Error GoTo TestFail
    Assert.AreSame srcTable, table.WrappedTable
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Object")
Private Sub TestDeleteTable()
    On Error GoTo TestFail
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
    srcTable.Delete
    Assert.IsNothing table.WrappedTable
    TestSheet.ResetTable
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
Private Sub TestUnrelatedChangeInWorksheetHasNoEffect()
    On Error GoTo TestFail
    '@Ignore IndexedDefaultMemberAccess
    srcTable.ListColumns(srcTable.ListColumns.Count).Range.Offset(0, 1).Insert

    Assert.AreEqual 0, logger.RawLog.Count, "too many events raised"
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Events")
Private Sub TestAppendRow()
    On Error GoTo TestFail
    Dim newRow As ListRow
    Set newRow = srcTable.ListRows.Add
    Assert.AreEqual idRowAppended, logger.EventClasses, "Only 1 kind of event should have been raised"
    Assert.AreEqual 1, logger.logEntry(idRowAppended).Count, "Count wrong"
    AreListRowsSame Assert, newRow, logger.logEntry(idRowAppended).Item(1)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Events")
Private Sub TestAppendRowTwice()
    On Error GoTo TestFail
    srcTable.ListRows.Add
    Dim newRow As ListRow
    Set newRow = srcTable.ListRows.Add
    Assert.AreEqual idRowAppended, logger.EventClasses, "Only 1 kind of event should have been raised"
    Assert.AreEqual 2, logger.logEntry(idRowAppended).Count, "Count wrong"
    AreListRowsSame Assert, newRow, logger.logEntry(idRowAppended).Item(2)

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
    'without cache it is impossible to differentiate between inserting and removing
    'so they should raise the same
    On Error GoTo TestFail
    Dim newRow As ListRow
    Set newRow = srcTable.ListRows.Add(srcTable.ListRows.Count \ 2)
    Assert.AreEqual idValueChanged, logger.EventClasses, "Only 1 kind of event should have been raised"
    Assert.AreEqual 1, logger.logEntry(idValueChanged).Count, "Count wrong"
    AreRangesSame Assert, newRow.Range, logger.logEntry(idValueChanged).Item(1)

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
    Dim newRowTrigger As Range
    Set newRowTrigger = srcTable.DataBodyRange.Cells(srcTable.ListRows.Count + 1, 1)

    newRowTrigger.Value2 = "Foo"

    Dim newRow As ListRow
    Set newRow = ListObjectHelperMethods.TargetToListRow(srcTable, newRowTrigger)

    Assert.AreEqual idRowAppended, logger.EventClasses, "Wrong kind/ too many kinds of event raised"
    Assert.AreEqual 1, logger.logEntry(idRowAppended).Count, "Count wrong"
    AreListRowsSame Assert, newRow, logger.logEntry(idRowAppended).Item(1)

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
    Assert.Fail                                  'todo me - the colname changed might help :)
    On Error GoTo TestFail
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

'@TestMethod("Events")
Private Sub TestValueChangedSingleCell()
    On Error GoTo TestFail
    Dim target As Range
    Set target = srcTable.DataBodyRange.Cells(srcTable.ListColumns.Count \ 2, srcTable.ListColumns.Count \ 2)
    target.Value2 = "foo"

    Assert.AreEqual idValueChanged, logger.EventClasses, "Wrong event types raised"
    Assert.AreEqual 1, logger.logEntry(idValueChanged).Count, "Event count wrong"
    '@Ignore IndexedDefaultMemberAccess
    AreRangesSame Assert, target, logger.logEntry(idValueChanged).Item(1)


TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Events")
Private Sub TestDeleteListRow()
    On Error GoTo TestFail
    Dim targetAddress As String
    With srcTable.ListRows.Item(srcTable.ListRows.Count \ 2)
        targetAddress = .Range.Address
        .Delete
    End With
    Dim target As Range
    Set target = table.wrappedTableParent.Range(targetAddress)

    Assert.AreEqual idValueChanged, logger.EventClasses, "Wrong event types raised"
    Assert.AreEqual 1, logger.logEntry(idValueChanged).Count, "Event count wrong"
    '@Ignore IndexedDefaultMemberAccess
    AreRangesSame Assert, target, logger.logEntry(idValueChanged).Item(1)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Object")
Private Sub TestParameterisedConstriuctor()
    On Error GoTo TestFail
    'check is instance but not the factory itself
    Assert.IsTrue TypeOf table.Categoriser Is NoCacheCategoriser

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Events")
Private Sub TestDeleteMiddleColumn()
    On Error GoTo TestFail
    Dim targetAddress As String
    With srcTable.ListColumns.Item(srcTable.ListColumns.Count \ 2)
        targetAddress = .Range.Address
        .Delete
    End With
    Dim target As Range
    Set target = table.wrappedTableParent.Range(targetAddress)
    Assert.AreEqual idValueChanged, logger.EventClasses, "Wrong event types raised"
    Assert.AreEqual 1, logger.logEntry(idValueChanged).Count, "Event count wrong"
    '@Ignore IndexedDefaultMemberAccess
    AreRangesSame Assert, target, logger.logEntry(idValueChanged).Item(1)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
