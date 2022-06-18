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

Public Sub CategoriseChange(ByVal events As ITableEventsSource, ByVal target As Range, ByVal table As ListObject)
    If table Is Nothing Then Exit Sub
    If Application.Intersect(target, table.Range) Is Nothing Then Exit Sub
    If target.Cells.CountLarge = 1 And Not Application.Intersect(target, table.HeaderRowRange) Is Nothing Then
        events.RaiseColumnNameChanged target
        Exit Sub
    End If
    'cell in final row or column

    Select Case table.ListRows.Count - this.cache.listRowCount
        Case 1
            If IndexRelativeToOrigin(table, target, False) = table.ListRows.Count Then
                events.RaiseRowAppended TargetToListRow(table, target)
            Else
                events.RaiseRowAdded TargetToListRow(table, target)
            End If
        Case -1
            events.RaiseRowDeleted target
        Case 0
            Select Case table.ListColumns.Count - this.cache.listColCount
                Case 1
                    events.RaiseColumnAdded TargetToListColumn(table, target)
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
