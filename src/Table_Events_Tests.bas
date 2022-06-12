Attribute VB_Name = "Table_Events_Tests"
'@IgnoreModule SelfAssignedDeclaration
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.PermissiveAssertClass
'@Ignore VariableNotUsed
Private Fakes As Rubberduck.FakesProvider
Private watcher As EventsWatcher

Private Property Get logger() As LoggingEventSink
    Set logger = watcher.logger
End Property

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
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
    Set watcher = New EventsWatcher
    Set watcher.logger = New LoggingEventSink
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set watcher = Nothing
End Sub

'@TestMethod("Uncategorized")
Private Sub TestManuallyRaiseEvent()
    On Error GoTo TestFail

    'Arrange:
    Dim eventRange As Range
    Set eventRange = TestSheet.Range("A1")

    Dim eventsSource As ITableEventsSource
    Set eventsSource = New TableWatcher

    Set watcher.events = eventsSource

    'Act:
    eventsSource.RaiseColumnNameChanged eventRange

    'Assert:
    Assert.AreEqual logger.EventClasses, idColNameChange, "unexpected event happened"
    With logger.logEntry(idColNameChange)
        Assert.AreEqual 1, .Count
        TestUtils.AreRangesSame Assert, eventRange, .Item(1)
    End With
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
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

    Set watcher.events = eventsSource
    'Act:
    Const numberOfEvents As Long = 5
    Dim i As Long
    For i = 1 To numberOfEvents
        eventsSource.RaiseColumnNameChanged eventRange
    Next i
    'Assert:
    Assert.AreEqual numberOfEvents, logger.logEntry(idColNameChange).Count
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

