VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITableEventsSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder "src"
'@Interface

Public Sub RaiseRowAdded(ByVal where As ListRow)
End Sub

Public Sub RaiseColumnAdded(ByVal where As ListColumn)
End Sub

Public Sub RaiseColumnNameChanged(ByVal which As Range)
End Sub

Public Sub RaiseDataValueChanged(ByVal where As Range)
End Sub

Public Sub RaiseRowDeleted(ByVal prevLocation As Range)
End Sub
