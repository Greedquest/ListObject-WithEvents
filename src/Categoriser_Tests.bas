Attribute VB_Name = "Categoriser_Tests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

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
    Set table = TableWatcher.Create(srcTable)
    Set watcher.events = table
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set watcher = Nothing
    Set srcTable = Nothing
    TestSheet.Reset
End Sub


'@TestMethod("DefaultCategoriser")
Private Sub TestCachedCategoriserEventsOffRaises()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    Dim initialEvents As Boolean
    initialEvents = Application.EnableEvents
    Application.EnableEvents = False
    With DefaultCategoriser.Create(srcTable)
    End With
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Application.EnableEvents = initialEvents
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub
