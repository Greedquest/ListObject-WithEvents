VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventsWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "testing"
'@IgnoreModule FunctionReturnValueAlwaysDiscarded, FunctionReturnValueDiscarded, EncapsulatePublicField
Option Explicit

Public WithEvents events As TableWatcher
Attribute events.VB_VarHelpID = -1

Public logger As LoggingEventSink

Private Sub Class_Initialize()
    Const EVENTS_HANDLED As Long = 8             'increment when a new event handler is added
    static_assert_eventsHandled EVENTS_HANDLED, TypeName(Me)
End Sub

Private Function loggerAsEventSource() As ITableEventsSource
    Set loggerAsEventSource = logger
End Function

Private Sub events_ColumnAppended(ByVal where As ListColumn)
    loggerAsEventSource.RaiseColumnAppended where
End Sub

Private Sub events_ColumnDeleted(ByVal prevLocation As Range)
    loggerAsEventSource.RaiseColumnDeleted prevLocation
End Sub

Private Sub Events_ColumnInserted(ByVal where As ListColumn)
    loggerAsEventSource.RaiseColumnInserted where
End Sub

Private Sub Events_ColumnNameChanged(ByVal which As Range)
    loggerAsEventSource.RaiseColumnNameChanged which
End Sub

Private Sub events_DataValueChanged(ByVal where As Range)
    loggerAsEventSource.RaiseDataValueChanged where
End Sub

Private Sub Events_RowInserted(ByVal where As ListRow)
    loggerAsEventSource.RaiseRowInserted where
End Sub

Private Sub events_RowAppended(ByVal where As ListRow)
    loggerAsEventSource.RaiseRowAppended where
End Sub

Private Sub events_RowDeleted(ByVal prevLocation As Range)
    loggerAsEventSource.RaiseRowDeleted prevLocation
End Sub
