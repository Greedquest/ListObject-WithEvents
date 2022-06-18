Attribute VB_Name = "TestUtils"
'@Folder "testing"
Option Private Module
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub AreRangesSame( _
       ByVal Assert As Rubberduck.PermissiveAssertClass, _
       ByVal expected As Range, _
       ByVal actual As Range _
       )
    Assert.AreEqual _
        expected.Address(external:=True), _
        actual.Address(external:=True), _
        printf( _
        "expected {} but got {}", _
        expected.Address(external:=True), _
        actual.Address(external:=True) _
        )
End Sub

Public Sub AreListRowsSame( _
       ByVal Assert As Rubberduck.PermissiveAssertClass, _
       ByVal expected As ListRow, _
       ByVal actual As ListRow _
       )
    Assert.AreEqual expected.Index, actual.Index, "List row index mismatch"
    AreRangesSame Assert, expected.Range, actual.Range
End Sub

Public Sub AreListColumnsSame( _
       ByVal Assert As Rubberduck.PermissiveAssertClass, _
       ByVal expected As ListColumn, _
       ByVal actual As ListColumn _
       )
    Assert.AreEqual expected.Index, actual.Index, "List column index mismatch"
    AreRangesSame Assert, expected.Range, actual.Range
End Sub

'@Ignore AssignedByValParameter: by design
Public Function printf(ByVal mask As String, ParamArray tokens() As Variant) As String
    Dim i As Long
    For i = 0 To UBound(tokens)
        Dim escapedToken As String
        escapedToken = Replace$(tokens(i), "}", "\}") 'only need to replace closing bracket since {i\} is already invalid
        If InStr(1, mask, "{}") <> 0 Then
            'use positional mode {}
            mask = Replace$(mask, "{}", escapedToken, Count:=1)

        Else
            'use indexed mode {i}
            mask = Replace$(mask, "{" & i & "}", escapedToken)

        End If
    Next
    mask = Replace$(mask, "\}", "}")
    printf = mask
End Function


