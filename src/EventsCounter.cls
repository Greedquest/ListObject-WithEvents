VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventsCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Testing")
'@IgnoreModule EncapsulatePublicField
Option Explicit

Public WithEvents events As TableWatcher
Attribute events.VB_VarHelpID = -1

Private eventsLog As New Dictionary

Public Enum EventIDS
    idNone = 0
    idRowAdded = 2 ^ 0
    idColAdded = 2 ^ 1
    idColNameChange = 2 ^ 2
End Enum

'@Description "Bit mask for the types of events raised"
Public Property Get EventClasses() As EventIDS
Attribute EventClasses.VB_Description = "Bit mask for the types of events raised"
    Dim eventID As Variant
    Dim result As EventIDS
    result = idNone
    For Each eventID In Log.Keys
        result = result + CLng(eventID)
    Next eventID
    EventClasses = result
End Property

Public Property Get Log() As Dictionary
    Set Log = eventsLog
End Property

Public Property Get logEntry(ByVal eventID As EventIDS) As Collection
    If Not Log.Exists(eventID) Then
        Set Log.Item(eventID) = New Collection
    End If
    Set logEntry = Log.Item(eventID)
End Property

Private Sub Events_ColumnAdded(ByVal where As ListColumn)
    logEntry(idColAdded).Add where
End Sub

Private Sub Events_ColumnNameChanged(ByVal which As Range)
    logEntry(idColNameChange).Add which
End Sub

Private Sub Events_RowAdded(ByVal where As ListRow)
    logEntry(idRowAdded).Add where
End Sub
