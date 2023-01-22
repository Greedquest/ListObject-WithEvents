VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefaultCategoriser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "src.categorisers"
'@PredeclaredId
Option Explicit
Implements IWorksheetChangeCategoriser

Private Type TableCache
    listRowCount As Long
    listColCount As Long
End Type

Private Type TCategoriser
    cache As TableCache
End Type

Private this As TCategoriser

Public Function Create(ByVal tableInitialState As ListObject) As DefaultCategoriser
    If Not tableInitialState.Application.EnableEvents Then Err.Raise 5, Description:="Workbook must have events enabled in order to keep cache up-to-date"
    With New DefaultCategoriser
        Set Create = .Self
        .Init tableInitialState
    End With
End Function

Friend Property Get Self() As DefaultCategoriser
    Set Self = Me
End Property

Friend Sub Init(ByVal tableInitialState As ListObject)
    this.cache.listColCount = tableInitialState.ListColumns.Count
    this.cache.listRowCount = tableInitialState.ListRows.Count
End Sub

Private Sub IWorksheetChangeCategoriser_CategoriseChange(ByVal events As ITableEventsSource, ByVal target As Range, ByVal table As ListObject)
    CategoriseChange events, target, table
End Sub

'@Description "Updates cache and returns the change"
Private Function updateToRowCache(ByVal table As ListObject) As Long
Attribute updateToRowCache.VB_Description = "Updates cache and returns the change"
    updateToRowCache = table.ListRows.Count - this.cache.listRowCount
    this.cache.listRowCount = table.ListRows.Count
End Function

'@Description "Updates cache and returns the change"
Private Function updateToColCache(ByVal table As ListObject) As Long
Attribute updateToColCache.VB_Description = "Updates cache and returns the change"
    updateToColCache = table.ListColumns.Count - this.cache.listColCount
    this.cache.listColCount = table.ListColumns.Count
End Function

Public Sub CategoriseChange(ByVal events As ITableEventsSource, ByVal target As Range, ByVal table As ListObject)
    If table Is Nothing Then Exit Sub
    If Application.Intersect(target, table.Range) Is Nothing Then Exit Sub
    If target.Cells.CountLarge = 1 And Not Application.Intersect(target, table.HeaderRowRange) Is Nothing Then
        events.RaiseColumnNameChanged target
        Exit Sub
    End If
    'cell in final row or column

    Select Case updateToRowCache(table)
        Case 1
            Debug.Assert target.Rows.Count = 1
            If IndexRelativeToOrigin(table, target, False) = table.ListRows.Count Then
                events.RaiseRowAppended TargetToListRow(table, target)
            Else
                events.RaiseRowInserted TargetToListRow(table, target)
            End If
        Case -1
            Debug.Assert target.Rows.Count = 1
            events.RaiseRowDeleted target
        Case 0
            Select Case updateToColCache(table)
                Case 1
                    Debug.Assert target.Columns.Count = 1
                    If IndexRelativeToOrigin(table, target, True) = table.ListColumns.Count Then
                        events.RaiseColumnAppended TargetToListColumn(table, target)
                    Else
                        events.RaiseColumnInserted TargetToListColumn(table, target)
                    End If
                Case -1
                    events.RaiseColumnDeleted target
                Case 0
                    events.RaiseDataValueChanged target
                Case Else
                    Debug.Assert False           'multiple cols not handled yet
            End Select
        Case Else
            Debug.Assert False                   'multiple rows not handled yet
    End Select

End Sub
