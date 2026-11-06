Attribute VB_Name = "PptMcpBridge"
' ppt-mcp's VBA bridge for PowerPoint for Mac. ppt-mcp sends `run VB macro`
' with a batch of calls; each is run on the object model, written out where
' PptMcpCalls has it and with CallByName otherwise, and
' the answers are written to a file in PowerPoint's container, which is the
' one place both sides can reach. Every entry point traps its errors, because
' an untrapped one would stop PowerPoint on a dialog.
Option Explicit
Public Const BRIDGE_VERSION As String = "3"
Private Sub WriteResult(id As String, s As String)
    Dim b() As Byte, f As Integer, dir As String, tmp As String, fin As String
    dir = Environ("HOME") & "/Documents/"
    tmp = dir & "ppt-mcp-bridge-" & id & ".tmp"
    fin = dir & "ppt-mcp-bridge-" & id & ".txt"
    On Error Resume Next
    Kill tmp
    Kill fin
    On Error GoTo 0
    b = s
    f = FreeFile
    Open tmp For Binary Access Write As #f
    Put #f, , b
    Close #f
    Name tmp As fin
End Sub
Private Function Cell(x As Variant) As String
    If IsObject(x) Then
        Cell = "o"
    ElseIf IsNumeric(x) And VarType(x) <> vbString Then
        Cell = "n" & Trim$(Str$(x))
    Else
        Cell = "s" & CStr(x)
    End If
End Function
Private Function DescribeArray(v As Variant) As String
    Dim d2 As Long, i As Long, j As Long, s As String
    On Error Resume Next
    d2 = UBound(v, 2)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        s = "array1" & Chr$(31) & (UBound(v) - LBound(v) + 1)
        For i = LBound(v) To UBound(v)
            s = s & Chr$(31) & Cell(v(i))
        Next
    Else
        On Error GoTo 0
        s = "array2" & Chr$(31) & (UBound(v, 1) - LBound(v, 1) + 1) & Chr$(31) & (UBound(v, 2) - LBound(v, 2) + 1)
        For i = LBound(v, 1) To UBound(v, 1)
            For j = LBound(v, 2) To UBound(v, 2)
                s = s & Chr$(31) & Cell(v(i, j))
            Next
        Next
    End If
    DescribeArray = s
End Function
Private Function Describe(v As Variant) As String
    If IsObject(v) Then
        If v Is Nothing Then
            Describe = "nothing"
        Else
            Describe = "obj" & Chr$(31) & TypeName(v) & Chr$(31) & Store(v)
        End If
    ElseIf IsEmpty(v) Then
        Describe = "empty"
    ElseIf IsNull(v) Then
        Describe = "null"
    ElseIf VarType(v) = vbBoolean Then
        Describe = "bool" & Chr$(31) & IIf(v, "1", "0")
    ElseIf IsArray(v) Then
        Describe = DescribeArray(v)
    ElseIf IsNumeric(v) And VarType(v) <> vbString Then
        Describe = "num" & Chr$(31) & TypeName(v) & Chr$(31) & Trim$(Str$(v))
    Else
        Describe = "str" & Chr$(31) & TypeName(v) & Chr$(31) & CStr(v)
    End If
End Function
Sub Ping(id As String)
    On Error Resume Next
    WriteResult id, "ok" & Chr$(31) & BRIDGE_VERSION & Chr$(31) & Application.Version
End Sub
Sub Dispatch(id As String, batch As String)
    Dim reqs() As String, f() As String, o As Object, out As String
    Dim n As Integer, k As Long, r As Variant, ct As VbCallType
    On Error GoTo fail
    reqs = Split(FromHex(batch), Chr$(30))
    For k = 0 To UBound(reqs)
        f = Split(reqs(k), Chr$(31))
        If f(1) = "r" Then
            ReleaseAll
            out = out & "ok" & Chr$(31) & "empty" & Chr$(30)
        Else
            Set o = Obj(f(0))
            n = UBound(f) - 2
            r = Empty
            ' A call written out keeps PowerPoint's own error text, which
            ' CallByName replaces with 440 "Automation error"; CallByName runs
            ' whatever PptMcpCalls does not write out.
            If Not Direct(o, f(1), f(2), f, n, r) Then
                Select Case f(1)
                    Case "l": CallByName o, f(2), VbLet, ArgOf(f(3))
                    Case "s": CallByName o, f(2), VbSet, ArgOf(f(3))
                    Case Else
                        If f(1) = "g" Then ct = VbGet Else ct = VbMethod
                        Select Case n
                            Case 0: Assign r, CallByName(o, f(2), ct)
                            Case 1: Assign r, CallByName(o, f(2), ct, ArgOf(f(3)))
                            Case 2: Assign r, CallByName(o, f(2), ct, ArgOf(f(3)), ArgOf(f(4)))
                            Case 3: Assign r, CallByName(o, f(2), ct, ArgOf(f(3)), ArgOf(f(4)), ArgOf(f(5)))
                            Case 4: Assign r, CallByName(o, f(2), ct, ArgOf(f(3)), ArgOf(f(4)), ArgOf(f(5)), ArgOf(f(6)))
                            Case 5: Assign r, CallByName(o, f(2), ct, ArgOf(f(3)), ArgOf(f(4)), ArgOf(f(5)), ArgOf(f(6)), ArgOf(f(7)))
                            Case 6: Assign r, CallByName(o, f(2), ct, ArgOf(f(3)), ArgOf(f(4)), ArgOf(f(5)), ArgOf(f(6)), ArgOf(f(7)), ArgOf(f(8)))
                            Case 7: Assign r, CallByName(o, f(2), ct, ArgOf(f(3)), ArgOf(f(4)), ArgOf(f(5)), ArgOf(f(6)), ArgOf(f(7)), ArgOf(f(8)), ArgOf(f(9)))
                            Case Else: Assign r, CallByName(o, f(2), ct, ArgOf(f(3)), ArgOf(f(4)), ArgOf(f(5)), ArgOf(f(6)), ArgOf(f(7)), ArgOf(f(8)), ArgOf(f(9)), ArgOf(f(10)))
                        End Select
                End Select
            End If
            out = out & "ok" & Chr$(31) & Describe(r) & Chr$(30)
        End If
    Next
    WriteResult id, out
    Exit Sub
fail:
    out = out & "err" & Chr$(31) & k + 1 & Chr$(31) & Err.Number & Chr$(31) & Err.Source & Chr$(31) & Err.Description & Chr$(30)
    On Error Resume Next
    WriteResult id, out
End Sub
