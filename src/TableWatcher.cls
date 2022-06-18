VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "src"
'@PredeclaredId
Option Explicit

Implements ITableEventsSource

Private WithEvents parentSheet As Worksheet
Attribute parentSheet.VB_VarHelpID = -1

Public Event RowAdded(ByVal where As ListRow)
Public Event ColumnAdded(ByVal where As ListColumn)
Public Event ColumnNameChanged(ByVal which As Range)
Public Event ColumnDeleted(ByVal prevLocation As Range)
Public Event DataValueChanged(ByVal where As Range)
Public Event RowDeleted(ByVal prevLocation As Range)
Public Event RowAppended(ByVal where As ListRow)


Private Type TTableWatcher
    unsafeWrappedTable As ListObject             'unsafe as it may become disconnected - use Get/Set instead
    eventsCategoriser As IWorksheetChangeCategoriser
End Type

Private this As TTableWatcher

Public Function Create(ByVal srcTable As ListObject, Optional ByVal eventsCategoriser As IWorksheetChangeCategoriser) As TableWatcher
    With New TableWatcher
        Set Create = .Self
        .Init srcTable, eventsCategoriser
    End With
End Function

Friend Property Get Self() As TableWatcher
    Set Self = Me
End Property

Friend Sub Init(ByVal srcTable As ListObject, ByVal eventsCategoriser As IWorksheetChangeCategoriser)
    Set WrappedTable = srcTable
    Set this.eventsCategoriser = eventsCategoriser
    If eventsCategoriser Is Nothing Then
        Set this.eventsCategoriser = DefaultCategoriser.Create(srcTable)
    Else
        Set this.eventsCategoriser = eventsCategoriser
    End If
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
Attribute WrappedTable.VB_Description = "Returns table if it exists otherwise Nothing"
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

Public Property Get Categoriser() As IWorksheetChangeCategoriser
    Set Categoriser = this.eventsCategoriser
End Property

Private Sub parentSheet_Change(ByVal target As Range)
    this.eventsCategoriser.CategoriseChange Me, target, WrappedTable
End Sub

Private Sub ITableEventsSource_RaiseColumnDeleted(ByVal prevLocation As Range)
    RaiseEvent ColumnDeleted(prevLocation)
End Sub

Private Sub ITableEventsSource_RaiseDataValueChanged(ByVal where As Range)
    RaiseEvent DataValueChanged(where)
End Sub

Private Sub ITableEventsSource_RaiseRowAdded(ByVal where As ListRow)
    RaiseEvent RowAdded(where)
End Sub

Private Sub ITableEventsSource_RaiseColumnAdded(ByVal where As ListColumn)
    RaiseEvent ColumnAdded(where)
End Sub

Private Sub ITableEventsSource_RaiseColumnNameChanged(ByVal which As Range)
    RaiseEvent ColumnNameChanged(which)
End Sub

Private Sub ITableEventsSource_RaiseRowDeleted(ByVal prevLocation As Range)
    RaiseEvent RowDeleted(prevLocation)
End Sub

Private Sub ITableEventsSource_RaiseRowAppended(ByVal where As ListRow)
    RaiseEvent RowAppended(where)
End Sub
