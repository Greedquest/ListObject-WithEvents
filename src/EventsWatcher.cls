VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventsWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Testing")
'@IgnoreModule FunctionReturnValueAlwaysDiscarded, FunctionReturnValueDiscarded, EncapsulatePublicField
Option Explicit

Public WithEvents events As TableWatcher
Attribute events.VB_VarHelpID = -1

Public logger As LoggingEventSink

Private Function loggerAsEventSource() As ITableEventsSource
    Set loggerAsEventSource = logger
End Function

Private Sub Events_ColumnAdded(ByVal where As ListColumn)
    loggerAsEventSource.RaiseColumnAdded where
End Sub

Private Sub Events_ColumnNameChanged(ByVal which As Range)
    loggerAsEventSource.RaiseColumnNameChanged which
End Sub

Private Sub events_DataValueChanged(ByVal where As Range)
    loggerAsEventSource.RaiseDataValueChanged where
End Sub

Private Sub Events_RowAdded(ByVal where As ListRow)
    loggerAsEventSource.RaiseRowAdded where
End Sub

