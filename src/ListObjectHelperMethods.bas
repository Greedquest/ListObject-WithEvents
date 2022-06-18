Attribute VB_Name = "ListObjectHelperMethods"
'@Folder "src.utils"
Option Explicit
Option Private Module

'@Description "1-based Index of listcol/listrow in that table, accounting for header row"
Public Function IndexRelativeToOrigin(ByVal WrappedTable As ListObject, ByVal target As Range, ByVal byCol As Boolean) As Long
Attribute IndexRelativeToOrigin.VB_Description = "1-based Index of listcol/listrow in that table, accounting for header row"
    IndexRelativeToOrigin = OffsetRelativeToOrigin(WrappedTable, target, byCol) + IIf(byCol, 1, 0)
End Function

'@Description "0-based Raw offset in rows/cols from tl cell of table"
Private Function OffsetRelativeToOrigin(ByVal WrappedTable As ListObject, ByVal target As Range, ByVal byCol As Boolean) As Long
Attribute OffsetRelativeToOrigin.VB_Description = "0-based Raw offset in rows/cols from tl cell of table"
    If byCol Then
        OffsetRelativeToOrigin = target.Column - WrappedTable.Range.Column
    Else
        OffsetRelativeToOrigin = target.Row - WrappedTable.Range.Row
    End If
End Function

Public Function TargetToListColumn(ByVal WrappedTable As ListObject, ByVal target As Range) As ListColumn
    Dim colIndex As Long
    colIndex = OffsetRelativeToOrigin(WrappedTable, target, byCol:=True) + 1
    Debug.Assert colIndex >= 0
    Debug.Assert colIndex <= WrappedTable.ListColumns.Count
    '@Ignore IndexedDefaultMemberAccess
    Set TargetToListColumn = WrappedTable.ListColumns(colIndex)
End Function

Public Function TargetToListRow(ByVal WrappedTable As ListObject, ByVal target As Range) As ListRow
    Dim rowIndex As Long
    rowIndex = OffsetRelativeToOrigin(WrappedTable, target, byCol:=False) 'no +1 because of headers
    Debug.Assert rowIndex >= 0
    Debug.Assert rowIndex <= WrappedTable.ListRows.Count 'BUG adding a Total Row
    '@Ignore IndexedDefaultMemberAccess
    Set TargetToListRow = WrappedTable.ListRows(rowIndex)
End Function
