Attribute VB_Name = "Module2"
Option Explicit
Public ValueFRM As String

Global Memory(65535) As Integer
Global flag As Boolean
Global Mode As Integer
Global opcode As Integer
Global operand As Long
Global instLen(255) As Integer

Global A As Integer
Global B As Integer
Global C As Integer
Global D As Integer
Global E As Integer
Global F As Integer
Global H As Integer
Global L As Integer
Global SP As Long
Global PC As Long

Global Const Carry = 0
Global Const Parity = 2
Global Const Auxillary = 4
Global Const Zero = 6
Global Const Sign = 7
Public Sub Mov(i As Integer, j As Integer)
    i = j
End Sub

Public Function GetBit(iNum As Integer, pos As Integer) As Integer
    If (iNum And (2 ^ pos)) = 0 Then
        GetBit = 0
    Else
        GetBit = 1
    End If
End Function

Public Sub SetBit(iNum As Integer, pos As Integer)
    iNum = iNum Or (2 ^ pos)
End Sub

Public Sub MovMem2Reg(x As Integer)
    Dim memPtr As Long
    'memPtr = Merge(H, L)
    x = Hex2Dec(GetValue(frm8085_Sim.txtH.Text, frm8085_Sim.txtL.Text), 2)
End Sub

Public Sub MovReg2Mem(x As Integer)
    Dim memPtr As Long
    'memPtr = Merge(H, L)
    'Memory(memPtr) = x
    'If output represents Hex value
    If InStr(1, UCase(Hex(x)), "A") Or _
       InStr(1, UCase(Hex(x)), "B") Or _
       InStr(1, UCase(Hex(x)), "C") Or _
       InStr(1, UCase(Hex(x)), "D") Or _
       InStr(1, UCase(Hex(x)), "E") Or _
       InStr(1, UCase(Hex(x)), "F") Then
        frm8085_Sim.txtM.Text = Hex(x)
        SetValue frm8085_Sim.txtH.Text, frm8085_Sim.txtL.Text, Hex(x)
    Else
        frm8085_Sim.txtM.Text = x
        SetValue frm8085_Sim.txtH.Text, frm8085_Sim.txtL.Text, CStr(x)
    End If
    
End Sub
Public Sub ResetBit(x As Integer, pos As Integer)
    x = x And (Not (2 ^ pos))
End Sub

Public Sub CheckFlags(x As Integer)
    Dim count As Integer, i As Integer
    If GetBit(x, Sign) = 1 Then
        Call SetBit(F, Sign)
    Else
        Call ResetBit(F, Sign)
    End If
    
    If x = 0 Then
        Call SetBit(F, Zero)
    Else
        Call ResetBit(F, Zero)
    End If
    
    count = 0
    
    For i = 0 To 7 Step 1
        If GetBit(x, i) = 1 Then
            count = count + 1
        End If
    Next i
    
    If (count Mod 2) = 0 Then
        Call SetBit(F, Parity)
    Else
        Call ResetBit(F, Parity)
    End If
    
End Sub
Public Function Get4(x As Integer, low As Boolean) As Integer
    If low = True Then
        Get4 = x And 15
    Else
        Get4 = (x And 240) \ 16
    End If
End Function

Public Sub Add(x As Integer, y As Integer)
    If (Get4(x, True) + Get4(y, True)) > 15 Then
        Call SetBit(F, Auxillary)
    Else
        Call ResetBit(F, Auxillary)
    End If
    
    x = x + y
    If x > 255 Then
        x = x - 1
        x = x Mod 255
        Call SetBit(F, Carry)
    Else
        Call ResetBit(F, Carry)
    End If
    Call CheckFlags(x)
End Sub

Public Sub Xthl()
    Dim s As Long
    s = SP
    Call Swap(L, Memory(s))
    s = s + 1
    If (s > 65535) Then
        s = 0
    End If
    Call Swap(H, Memory(s))
    
End Sub

Public Sub Cma()
    A = Not A
    A = A And 255
End Sub

Public Sub Adc(x As Integer, y As Integer)
    Dim yy As Integer
    yy = y
    If GetBit(F, Carry) = 1 Then
        yy = yy + 1
        yy = yy And 255
    End If
    Call Add(x, yy)
End Sub

Public Sub Subs(x As Integer, ByVal y As Integer)
    y = Not y
    y = y And 255
    y = y + 1
    y = y And 255
    Call Add(x, y)
    If GetBit(F, Carry) = 1 Then
        Call ResetBit(F, Carry)
    Else
        Call SetBit(F, Carry)
    End If
End Sub

Public Sub Ora(x As Integer)
    A = A Or x
    CheckFlags (A)
    Call ResetBit(F, Carry)
    Call ResetBit(F, Auxillary)
End Sub

Public Sub Ana(x As Integer)
    A = A And x
    CheckFlags (A)
    Call SetBit(F, Auxillary)
    Call ResetBit(F, Carry)
End Sub

Public Sub Kall()
    Dim high As Integer, low As Integer
    Call Split(Hex(PC), high, low)
    SP = SP - 1
    If SP < 0 Then
        SP = 65535
    End If
    Memory(SP) = high
    SP = SP - 1
    If SP < 0 Then
        SP = 65535
    End If
    Memory(SP) = low
    PC = operand
End Sub
Public Sub Inx(ByRef x As Integer, ByRef y As Integer)
    Dim ad As Long
    Dim s As String, k As Integer, s1 As String, s2 As String
    ad = Hex2Dec(CStr(Hex(y)), 2)
    ad = ad + 1
    s = UCase(Hex(ad))
    For k = 1 To 2 - Len(s) Step 1
        s = "0" + s
    Next k
    y = Hex2Dec(s, 2)
End Sub

Public Sub Dcx(ByRef x As Integer, ByRef y As Integer)
    Dim ad As Long
    Dim s As String, k As Integer
    ad = Hex2Dec(CStr(Hex(y)), 2)
    ad = ad - 1
    s = UCase(Hex(ad))
    For k = 1 To 2 - Len(s) Step 1
        s = "0" + s
    Next k
    y = Hex2Dec(s, 2)
End Sub

Public Sub Xra(x As Integer)
    A = A Xor x
    CheckFlags (A)
    Call ResetBit(F, Carry)
    Call ResetBit(F, Auxillary)
End Sub

Public Sub Inr(x As Integer)
    Dim i As Integer
    i = GetBit(F, Carry)
    Call Add(x, 1)
    If i = 1 Then
        Call SetBit(F, Carry)
    Else
        Call ResetBit(F, Carry)
    End If
End Sub

Public Sub Dcr(x As Integer)
    Dim i As Integer
    i = GetBit(F, Carry)
    Call Subs(x, 1)
    If i = 1 Then
        Call SetBit(F, Carry)
    Else
        Call ResetBit(F, Carry)
    End If
End Sub

Public Sub Push(x As Integer, y As Integer)
    SP = SP - 1
    If SP < 0 Then
        SP = 65535
    End If
    Memory(SP) = x
    SP = SP - 1
    If SP < 0 Then
        SP = 65535
    End If
    Memory(SP) = y
End Sub

Public Sub Pop(x As Integer, y As Integer)
    y = Memory(SP)
    SP = SP + 1
    If SP > 65535 Then
        SP = 0
    End If
    x = Memory(SP)
    SP = SP + 1
    If SP > 65535 Then
        SP = 0
    End If
End Sub

Public Sub Ret()
    Dim high As Integer, low As Integer
    low = Memory(SP)
    SP = SP + 1
    If SP > 65535 Then
        SP = 0
    End If
    high = Memory(SP)
    PC = Merge(high, low)
    SP = SP + 1
    If SP > 65535 Then
        SP = 0
    End If
End Sub

Public Function RotateLeft() As Integer
    Dim i As Integer
    i = GetBit(A, 7)
    A = A * 2
    A = A And 255
    If i = 1 Then
        Call SetBit(A, 0)
        Call SetBit(F, Carry)
    Else
        Call ResetBit(A, 0)
        Call ResetBit(F, Carry)
    End If
    RotateLeft = i
End Function

Public Function RotateRight() As Integer
    Dim i As Integer
    i = GetBit(A, 0)
    A = A \ 2
    If i = 1 Then
        Call SetBit(A, 7)
        Call SetBit(F, Carry)
    Else
        Call ResetBit(A, 7)
        Call ResetBit(F, Carry)
    End If
    RotateRight = i
End Function

Public Sub Rar()
    Dim i As Integer, j As Integer
    i = GetBit(F, Carry)
    j = RotateRight
    If i = 1 Then
        Call SetBit(A, 7)
    Else
        Call ResetBit(A, 7)
    End If
    If j = 1 Then
        Call SetBit(F, Carry)
    Else
        Call ResetBit(F, Carry)
    End If
End Sub

Public Sub Ral()
    Dim i As Integer, j As Integer
    i = GetBit(F, Carry)
    j = RotateLeft
    If i = 1 Then
        Call SetBit(A, 0)
    Else
        Call ResetBit(A, 0)
    End If
    If j = 1 Then
        Call SetBit(F, Carry)
    Else
        Call ResetBit(F, Carry)
    End If
End Sub

Public Sub Cmp(x As Integer)
    Dim i As Integer
    i = A
    Call Subs(i, x)
End Sub

Public Sub Dad(x As Integer, y As Integer)
    Dim l1 As Long, l2 As Long, sum As Long
    l1 = Merge(H, L)
    l2 = Merge(x, y)
    sum = l1 + l2
    If sum > 65535 Then
        Call SetBit(F, Carry)
        sum = sum And 65535
    End If
    Call Split(Hex(sum), H, L)
End Sub

Public Sub DadSP()
    Dim l1 As Long, l2 As Long, sum As Long
    l1 = Merge(H, L)
    l2 = SP
    sum = l1 + l2
    If sum > 65535 Then
        Call SetBit(F, Carry)
        sum = sum And 65535
    End If
    Call Split(Hex(sum), H, L)
End Sub

Public Sub Daa()
    Dim high As Integer, low As Integer, carryFlag As Integer
    low = Get4(A, True)
    high = Get4(A, False)
    If A > 15 Then
    If (low > 9) Or (GetBit(F, Auxillary) = 1) Then
        carryFlag = GetBit(F, Carry)
        Call Add(A, 6)
        Call SetBit(F, Carry)
    End If
    If (high > 9) Or (GetBit(F, Carry) = 1) Then
        Call Add(A, 96)
    End If
    End If
End Sub

Public Sub Sbb(x As Integer)
    Dim y As Integer
    y = x
    If GetBit(F, Carry) = 1 Then
        y = y + 1
        y = y And 255
    End If
    Call Subs(A, y)
End Sub

Public Function GetValue(Hi As String, Lo As String) As String
    Dim tmp As String, R As Integer, C As Integer, Fl_Row As Integer
    tmp = Hi & Left(Lo, 1)
    For Fl_Row = 1 To frm8085_Sim.MSFlxGrdSysMem.Rows - 1
        frm8085_Sim.MSFlxGrdSysMem.Row = Fl_Row
        frm8085_Sim.MSFlxGrdSysMem.Col = 0
        If Mid(frm8085_Sim.MSFlxGrdSysMem.Text, 1, Len(tmp)) = tmp Then
            R = frm8085_Sim.MSFlxGrdSysMem.Row
            Exit For
        End If
    Next
    frm8085_Sim.MSFlxGrdSysMem.Row = R
    If R > 0 Then
        If UCase(Lo) = "0A" Or UCase(Lo) = "AA" Or UCase(Lo) = "BA" Or _
        UCase(Lo) = "CA" Or UCase(Lo) = "DA" Or UCase(Lo) = "EA" Or UCase(Lo) = "FA" Or _
        Right(UCase(Lo), 1) = "A" Then
            frm8085_Sim.MSFlxGrdSysMem.Col = 11
        ElseIf UCase(Lo) = "0B" Or UCase(Lo) = "AB" Or UCase(Lo) = "BB" Or _
        UCase(Lo) = "CB" Or UCase(Lo) = "DB" Or UCase(Lo) = "EB" Or UCase(Lo) = "FB" Or _
        Right(UCase(Lo), 1) = "B" Then
            frm8085_Sim.MSFlxGrdSysMem.Col = 12
        ElseIf UCase(Lo) = "0C" Or UCase(Lo) = "AC" Or UCase(Lo) = "BC" Or _
        UCase(Lo) = "CC" Or UCase(Lo) = "DC" Or UCase(Lo) = "EC" Or UCase(Lo) = "FC" Or _
        Right(UCase(Lo), 1) = "C" Then
            frm8085_Sim.MSFlxGrdSysMem.Col = 13
        ElseIf UCase(Lo) = "0D" Or UCase(Lo) = "AD" Or UCase(Lo) = "BD" Or _
        UCase(Lo) = "CD" Or UCase(Lo) = "DD" Or UCase(Lo) = "ED" Or UCase(Lo) = "FD" Or _
        Right(UCase(Lo), 1) = "D" Then
            frm8085_Sim.MSFlxGrdSysMem.Col = 14
        ElseIf UCase(Lo) = "0E" Or UCase(Lo) = "AE" Or UCase(Lo) = "BE" Or _
        UCase(Lo) = "CE" Or UCase(Lo) = "DE" Or UCase(Lo) = "EE" Or UCase(Lo) = "FE" Or _
        Right(UCase(Lo), 1) = "E" Then
            frm8085_Sim.MSFlxGrdSysMem.Col = 15
        ElseIf UCase(Lo) = "0F" Or UCase(Lo) = "AF" Or UCase(Lo) = "BF" Or _
        UCase(Lo) = "CF" Or UCase(Lo) = "DF" Or UCase(Lo) = "EF" Or UCase(Lo) = "FF" Or _
        Right(UCase(Lo), 1) = "F" Then
            frm8085_Sim.MSFlxGrdSysMem.Col = 16
        ElseIf Mid(UCase(Lo), 1, 1) = "A" Or Mid(UCase(Lo), 1, 1) = "B" Or Mid(UCase(Lo), 1, 1) = "C" _
        Or Mid(UCase(Lo), 1, 1) = "D" Or Mid(UCase(Lo), 1, 1) = "E" Or Mid(UCase(Lo), 1, 1) = "F" Then
            frm8085_Sim.MSFlxGrdSysMem.Col = CInt(Right(Lo, 1)) + 1
        ElseIf Asc(Lo) >= 48 And Asc(Lo) <= 57 Then
            frm8085_Sim.MSFlxGrdSysMem.Col = CInt(Right(Lo, 1)) + 1
        End If
        GetValue = frm8085_Sim.MSFlxGrdSysMem.Text
    Else
        GetValue = 0
    End If
End Function

Public Sub SetValue(H As String, L As String, val As String)
    Dim tmp As String, R As Integer, C As Integer, Fl_Row As Integer
    tmp = H & Left(L, 1)
    For Fl_Row = 1 To frm8085_Sim.MSFlxGrdSysMem.Rows - 1
        frm8085_Sim.MSFlxGrdSysMem.Row = Fl_Row
        frm8085_Sim.MSFlxGrdSysMem.Col = 0
        If Mid(frm8085_Sim.MSFlxGrdSysMem.Text, 1, Len(tmp)) = tmp Then
            R = frm8085_Sim.MSFlxGrdSysMem.Row
            Exit For
        End If
    Next
    frm8085_Sim.MSFlxGrdSysMem.Row = R
    If R > 0 Then
        If UCase(L) = "0A" Or UCase(L) = "AA" Or UCase(L) = "BA" Or _
           UCase(L) = "CA" Or UCase(L) = "DA" Or UCase(L) = "EA" Or UCase(L) = "FA" Or _
           Right(UCase(L), 1) = "A" Then
            frm8085_Sim.MSFlxGrdSysMem.Col = 11
        ElseIf UCase(L) = "0B" Or UCase(L) = "AB" Or UCase(L) = "BB" Or _
               UCase(L) = "CB" Or UCase(L) = "DB" Or UCase(L) = "EB" Or UCase(L) = "FB" Or _
               Right(UCase(L), 1) = "B" Then
            frm8085_Sim.MSFlxGrdSysMem.Col = 12
        ElseIf UCase(L) = "0C" Or UCase(L) = "AC" Or UCase(L) = "BC" Or _
               UCase(L) = "CC" Or UCase(L) = "DC" Or UCase(L) = "EC" Or UCase(L) = "FC" Or _
               Right(UCase(L), 1) = "C" Then
            frm8085_Sim.MSFlxGrdSysMem.Col = 13
        ElseIf UCase(L) = "0D" Or UCase(L) = "AD" Or UCase(L) = "BD" Or _
               UCase(L) = "CD" Or UCase(L) = "DD" Or UCase(L) = "ED" Or UCase(L) = "FD" Or _
               Right(UCase(L), 1) = "D" Then
            frm8085_Sim.MSFlxGrdSysMem.Col = 14
        ElseIf UCase(L) = "0E" Or UCase(L) = "AE" Or UCase(L) = "BE" Or _
               UCase(L) = "CE" Or UCase(L) = "DE" Or UCase(L) = "EE" Or UCase(L) = "FE" Or _
               Right(UCase(L), 1) = "E" Then
            frm8085_Sim.MSFlxGrdSysMem.Col = 15
        ElseIf UCase(L) = "0F" Or UCase(L) = "AF" Or UCase(L) = "BF" Or _
               UCase(L) = "CF" Or UCase(L) = "DF" Or UCase(L) = "EF" Or UCase(L) = "FF" Or _
               Right(UCase(L), 1) = "F" Then
            frm8085_Sim.MSFlxGrdSysMem.Col = 16
        ElseIf Mid(UCase(L), 1, 1) = "A" Or Mid(UCase(L), 1, 1) = "B" Or Mid(UCase(L), 1, 1) = "C" _
            Or Mid(UCase(L), 1, 1) = "D" Or Mid(UCase(L), 1, 1) = "E" Or Mid(UCase(L), 1, 1) = "F" Then
            frm8085_Sim.MSFlxGrdSysMem.Col = CInt(Right(L, 1)) + 1
        ElseIf Asc(L) >= 48 And Asc(L) <= 57 Then
            frm8085_Sim.MSFlxGrdSysMem.Col = CInt(Right(L, 1)) + 1
        End If
        Dim k As Integer, s As String
        s = val
        For k = 1 To 2 - Len(s) Step 1
            s = "0" + s
        Next k
        frm8085_Sim.MSFlxGrdSysMem.Text = s
    Else
        MsgBox "Error in locating the Address " & H & L, vbCritical + vbOKOnly, "Address Error"
    End If
End Sub

Public Function GetRow(Hi As String, Lo As String) As Integer
    Dim tmp As String, R As Integer, C As Integer, Fl_Row As Integer
    tmp = Hi & Lo
    For Fl_Row = 1 To frm8085_Sim.MSFlexGrdAssemble.Rows - 1
        frm8085_Sim.MSFlexGrdAssemble.Row = Fl_Row
        frm8085_Sim.MSFlexGrdAssemble.Col = 0
        If InStr(1, frm8085_Sim.MSFlexGrdAssemble.Text, tmp) Then
            R = frm8085_Sim.MSFlexGrdAssemble.Row
            Exit For
        End If
    Next
    GetRow = R - 1
End Function


