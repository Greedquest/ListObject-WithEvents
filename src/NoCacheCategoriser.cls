VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NoCacheCategoriser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "src.categorisers"
Option Explicit
Implements IWorksheetChangeCategoriser

Private Sub IWorksheetChangeCategoriser_CategoriseChange(ByVal events As ITableEventsSource, ByVal target As Range, ByVal table As ListObject)
    CategoriseChange events, target, table
End Sub

Public Sub CategoriseChange(ByVal events As ITableEventsSource, ByVal target As Range, ByVal table As ListObject)
    If table Is Nothing Then Exit Sub
    If Application.Intersect(target, table.Range) Is Nothing Then Exit Sub
    If target.Cells.CountLarge = 1 And Not Application.Intersect(target, table.HeaderRowRange) Is Nothing Then
        events.RaiseColumnNameChanged target
        Exit Sub
    End If
    'cell in final row or column means we are appending
    If IndexRelativeToOrigin(table, target, byCol:=False) = table.ListRows.Count Then
        events.RaiseRowAppended TargetToListRow(table, target)
    ElseIf IndexRelativeToOrigin(table, target, byCol:=True) = table.ListColumns.Count Then
        events.RaiseColumnAdded TargetToListColumn(table, target)
    ElseIf target.Rows.Count = 1 Or target.Columns.Count = 1 Then
        events.RaiseDataValueChanged target
    Else
        Debug.Assert False                       'not implemented yet
    End If
End Sub
