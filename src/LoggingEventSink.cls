VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggingEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Testing")
Option Explicit

Implements ITableEventsSource

Private eventsLog As New Dictionary

Public Enum EventIDS
    idNone = 0
    idRowAdded = 2 ^ 0
    idColAdded = 2 ^ 1
    idColNameChange = 2 ^ 2
    idValueChanged = 2 ^ 3
End Enum

'@Description "Bit mask for the types of events raised"
Public Property Get EventClasses() As EventIDS
Attribute EventClasses.VB_Description = "Bit mask for the types of events raised"
    Dim eventID As Variant
    Dim result As EventIDS
    result = idNone
    For Each eventID In log.Keys
        result = result + CLng(eventID)
    Next eventID
    EventClasses = result
End Property

Public Property Get log() As Dictionary
    Set log = eventsLog
End Property

Public Property Get logEntry(ByVal eventID As EventIDS) As Collection
    If Not log.Exists(eventID) Then
        Set log.Item(eventID) = New Collection
    End If
    Set logEntry = log.Item(eventID)
End Property

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
