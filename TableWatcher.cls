VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBAProject")
'@PredeclaredId
Option Explicit

Implements ITableEventsSource

Private WithEvents parentSheet As Worksheet
Attribute parentSheet.VB_VarHelpID = -1

Public Event RowAdded(ByVal where As ListRow)
Public Event ColumnAdded(ByVal where As ListColumn)
Public Event ColumnNameChanged(ByVal which As Range)

Private Type TTableWatcher
    unsafeWrappedTable As ListObject             'unsafe as it may become disconnected - use Get/Set instead
    eventsCategoriser As IWorksheetChangeCategoriser
End Type

Private this As TTableWatcher

Private Sub ITableEventsSource_RaiseRowAdded(ByVal where As ListRow)
    RaiseEvent RowAdded(where)
End Sub

Private Sub ITableEventsSource_RaiseColumnAdded(ByVal where As ListColumn)
    RaiseEvent ColumnAdded(where)
End Sub

Private Sub ITableEventsSource_RaiseColumnNameChanged(ByVal which As Range)
    RaiseEvent ColumnNameChanged(which)
End Sub

Friend Property Get wrappedTableParent() As Worksheet
    Set wrappedTableParent = parentSheet
End Property

Private Function isTableConnected() As Boolean
    On Error Resume Next
    '@Ignore VariableNotUsed, AssignmentNotUsed: just trying to catch an error
    Dim temp As String: temp = this.unsafeWrappedTable.name
    isTableConnected = Err.Number = 0
    On Error GoTo 0
End Function

'@Description "Returns table if it exists otherwise Nothing"
Public Property Get WrappedTable() As ListObject
    If this.unsafeWrappedTable Is Nothing Then Exit Property
    If Not isTableConnected Then Set Me.WrappedTable = Nothing
    Set WrappedTable = this.unsafeWrappedTable
End Property

Public Property Set WrappedTable(ByVal newValue As ListObject)
    If newValue Is Nothing Then
        Set this.unsafeWrappedTable = Nothing
        Set parentSheet = Nothing
    Else
        Set this.unsafeWrappedTable = newValue
        Set parentSheet = this.unsafeWrappedTable.Parent
    End If
End Property

Friend Property Get Self() As TableWatcher
    Set Self = Me
End Property

Public Function Create(ByVal srcTable As ListObject, Optional ByVal eventsCategoriser As IWorksheetChangeCategoriser) As TableWatcher
    With New TableWatcher
        Set Create = .Self
        .Init srcTable, eventsCategoriser
    End With
End Function

Friend Sub Init(ByVal srcTable As ListObject, ByVal eventsCategoriser As IWorksheetChangeCategoriser)
    Set WrappedTable = srcTable
    Set this.eventsCategoriser = eventsCategoriser
    If eventsCategoriser Is Nothing Then
        Set this.eventsCategoriser = New WorksheetEventsCategoriser
    Else
        Set this.eventsCategoriser = eventsCategoriser
    End If
End Sub

Private Property Get events() As ITableEventsSource
    Set events = Me
End Property

Private Sub parentSheet_Change(ByVal target As Range)
    this.eventsCategoriser.CategoriseChange Me, target, WrappedTable
End Sub

