VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggingEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "testing"
Option Explicit

Implements ITableEventsSource

Private eventsLog As New Dictionary

Public Enum EventIDS
    idNone = 0
    idRowAdded = 2 ^ 0
    idColAdded = 2 ^ 1
    idColNameChange = 2 ^ 2
    idValueChanged = 2 ^ 3
    idRowDeleted = 2 ^ 4
    idRowAppended = 2 ^ 5
    idColDeleted = 2 ^ 6
    [_max]                                       'for static asserts
End Enum

'@Description "Bit mask for the types of events raised"
Public Property Get EventClasses() As EventIDS
Attribute EventClasses.VB_Description = "Bit mask for the types of events raised"
    Dim eventID As Variant
    Dim result As EventIDS
    result = idNone
    For Each eventID In RawLog.Keys
        result = result + CLng(eventID)
    Next eventID
    EventClasses = result
End Property

Public Property Get RawLog() As Dictionary
    Set RawLog = eventsLog
End Property

Public Property Get logEntry(ByVal eventID As EventIDS) As Collection
    If Not RawLog.Exists(eventID) Then
        Set RawLog.Item(eventID) = New Collection
    End If
    Set logEntry = RawLog.Item(eventID)
End Property

Private Sub ITableEventsSource_RaiseColDeleted(ByVal prevLocation As Range)
    logEntry(idColDeleted).Add prevLocation
End Sub

Private Sub ITableEventsSource_RaiseColumnAdded(ByVal where As ListColumn)
    logEntry(idColAdded).Add where
End Sub

Private Sub ITableEventsSource_RaiseColumnNameChanged(ByVal which As Range)
    logEntry(idColNameChange).Add which
End Sub

Private Sub ITableEventsSource_RaiseDataValueChanged(ByVal where As Range)
    logEntry(idValueChanged).Add where
End Sub

Private Sub ITableEventsSource_RaiseRowAdded(ByVal where As ListRow)
    logEntry(idRowAdded).Add where
End Sub

Private Sub ITableEventsSource_RaiseRowAppended(ByVal where As ListRow)
    logEntry(idRowAppended).Add where
End Sub

Private Sub ITableEventsSource_RaiseRowDeleted(ByVal prevLocation As Range)
    logEntry(idRowDeleted).Add prevLocation
End Sub
