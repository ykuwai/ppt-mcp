Attribute VB_Name = "PptMcpHandles"
' ppt-mcp's VBA bridge for PowerPoint for Mac: the objects a call hands back,
' kept by handle between calls, and the decoding of the arguments a call sends.
' Built into ppt-mcp-bridge.ppam by scripts/build_mac_vba_addin.py; see
' src/backend/vba/__init__.py for the other side.
Option Explicit
Public handles As Collection
Public nextId As Long
Public Function Store(ByVal o As Object) As String
    If handles Is Nothing Then Set handles = New Collection
    nextId = nextId + 1
    handles.Add o, "h" & nextId
    Store = "h" & nextId
End Function
Public Function Obj(ByVal h As String) As Object
    If h = "app" Then
        Set Obj = Application
    Else
        Set Obj = handles(h)
    End If
End Function
Public Sub ReleaseAll()
    Set handles = New Collection
End Sub
Public Function FromHex(h As String) As String
    Dim i As Long, s As String
    For i = 1 To Len(h) Step 4
        s = s & ChrW(Val("&H" & Mid$(h, i, 4)))
    Next
    FromHex = s
End Function
Public Function ArgOf(ByVal s As String) As Variant
    Select Case Left$(s, 2)
        Case "i:": ArgOf = CLng(Mid$(s, 3))
        Case "d:": ArgOf = Val(Mid$(s, 3))
        Case "b:": ArgOf = (Mid$(s, 3) = "1")
        Case "o:": Set ArgOf = Obj(Mid$(s, 3))
        Case "a:": ArgOf = Split(Mid$(s, 3), Chr$(29))
        Case "N:"
            Dim parts() As String, j As Long, arr() As Variant
            parts = Split(Mid$(s, 3), Chr$(29))
            ReDim arr(UBound(parts))
            For j = 0 To UBound(parts)
                arr(j) = Val(parts(j))
            Next
            ArgOf = arr
        Case "m:": ArgOf = CVErr(448)
        Case Else: ArgOf = Mid$(s, 3)
    End Select
End Function
