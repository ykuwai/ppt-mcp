Attribute VB_Name = "PptMcpHandles"
' ppt-mcp's VBA bridge for PowerPoint for Mac: the objects a call hands back,
' kept by handle between calls, and the decoding of the arguments a call sends.
' Built into ppt-mcp-bridge.ppam by scripts/build_mac_vba_addin.py; see
' src/backend/vba/__init__.py for the other side.
Option Explicit
Public handles As Collection
Public nextId As Long
' Each server that uses the add-in names itself on every batch, and its
' handles are listed under that name, so that it lets go of its own only:
' several servers can share one PowerPoint (#347).
Public owner As String
Public owned As Collection
Public Function Store(ByVal o As Object) As String
    Dim h As String, keys As String
    If handles Is Nothing Then Set handles = New Collection
    If owned Is Nothing Then Set owned = New Collection
    nextId = nextId + 1
    h = "h" & nextId
    handles.Add o, h
    On Error Resume Next
    keys = owned(owner)
    owned.Remove owner
    On Error GoTo 0
    owned.Add keys & h & Chr$(29), owner
    Store = h
End Function
Public Function Obj(ByVal h As String) As Object
    If h = "app" Then
        Set Obj = Application
    Else
        Set Obj = handles(h)
    End If
End Function
Public Sub Release(ByVal who As String)
    Dim keys As String, k As Variant
    If handles Is Nothing Or owned Is Nothing Then Exit Sub
    On Error Resume Next
    keys = owned(who)
    owned.Remove who
    For Each k In Split(keys, Chr$(29))
        If Len(k) > 0 Then handles.Remove k
    Next
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
