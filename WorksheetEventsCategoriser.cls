VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetEventsCategoriser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit
Implements IWorksheetChangeCategoriser

Private Sub IWorksheetChangeCategoriser_CategoriseChange(ByVal events As ITableEventsSource, ByVal target As Range, ByVal table As ListObject)
    CategoriseChange events, target, table
End Sub

Public Sub CategoriseChange(ByVal events As ITableEventsSource, ByVal target As Range, ByVal table As ListObject)
    If table Is Nothing Then Exit Sub
    If Application.Intersect(target, table.Range) Is Nothing Then Exit Sub
    If target.Cells.CountLarge = 1 Then
        events.RaiseColumnNameChanged target
    ElseIf target.Rows.CountLarge = 1 Then
        events.RaiseRowAdded TargetToListRow(table, target)
    ElseIf target.Columns.CountLarge = 1 Then
        events.RaiseColumnAdded TargetToListColumn(table, target)
    Else
        Debug.Assert False                       'not implemented
    End If
End Sub

Private Function relativeToOrigin(ByVal WrappedTable As ListObject, ByVal target As Range, ByVal byCol As Boolean) As Long
    If byCol Then
        relativeToOrigin = target.Column - WrappedTable.Range.Column
    Else
        relativeToOrigin = target.Row - WrappedTable.Range.Row
    End If
End Function

Private Function TargetToListColumn(ByVal WrappedTable As ListObject, ByVal target As Range) As ListColumn
    Dim colIndex As Long
    colIndex = relativeToOrigin(WrappedTable, target, byCol:=True) + 1
    Debug.Assert colIndex >= 0
    Debug.Assert colIndex <= WrappedTable.ListColumns.Count
    Set TargetToListColumn = WrappedTable.ListColumns(colIndex)
End Function

Private Function TargetToListRow(ByVal WrappedTable As ListObject, ByVal target As Range) As ListRow
    Dim rowIndex As Long
    rowIndex = relativeToOrigin(WrappedTable, target, byCol:=False) 'no +1 because of headers
    Debug.Assert rowIndex >= 0
    Debug.Assert rowIndex <= WrappedTable.ListRows.Count 'BUG adding a Total Row
    Set TargetToListRow = WrappedTable.ListRows(rowIndex)
End Function


