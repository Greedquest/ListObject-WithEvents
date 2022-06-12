Attribute VB_Name = "Module2"
'@Folder("VBAProject")
Option Explicit

Public Enum EventIDS
    idNone = 0
    idRowAdded = 2 ^ 0
    idColAdded = 2 ^ 1
    idColNameChange = 2 ^ 2
End Enum

Public Type EventParams
    kind As EventIDS
    args As Variant
End Type
