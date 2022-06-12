Attribute VB_Name = "Table_Events_Tests"
'@IgnoreModule SelfAssignedDeclaration
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private assert As Rubberduck.PermissiveAssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set assert = New Rubberduck.PermissiveAssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestMethod("Uncategorized")
Private Sub TestManuallyRaiseEvent()
    On Error GoTo TestFail

    'Arrange:
    Dim eventRange As Range
    Set eventRange = TestSheet.Range("A1")

    Dim eventsSource As ITableEventsSource
    Set eventsSource = New TableWatcher

    Dim counter As New EventsCounter
    Set counter.events = eventsSource

    'Act:
    eventsSource.RaiseColumnNameChanged eventRange

    'Assert:
    assert.AreEqual counter.EventClasses, idColNameChange, "unexpected event happened"
    With counter.logEntry(idColNameChange)
        assert.AreEqual 1, .Count
        TestUtils.AreRangesSame assert, eventRange, .Item(1)
    End With
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestManuallyRaiseEventMultipleTimes()
    On Error GoTo TestFail
    'Arrange:
    Dim eventRange As Range
    Set eventRange = TestSheet.Range("A1")

    Dim eventsSource As ITableEventsSource
    Set eventsSource = New TableWatcher

    Dim counter As New EventsCounter
    Set counter.events = eventsSource
    'Act:
    Const numberOfEvents As Long = 5
    Dim i As Long
    For i = 1 To numberOfEvents
        eventsSource.RaiseColumnNameChanged eventRange
    Next i
    'Assert:
    assert.AreEqual numberOfEvents, counter.logEntry(idColNameChange).Count
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

