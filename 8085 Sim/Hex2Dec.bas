Attribute VB_Name = "Module1"
Option Explicit

Public Function Hex2Dec(ByVal s As String, Length As Integer) As Long
    
    Dim i As Integer
    Dim val As Long
    Dim C As String * 1
    
    val = 0
    s = UCase(s)
    For i = 1 To Length Step 1
        
        C = Mid(s, i, 1)
        
        If Asc(C) >= 48 And Asc(C) <= 57 Then
            val = val * 16 + Asc(C) - 48
        ElseIf Asc(C) >= 65 And Asc(C) <= 70 Then
                val = val * 16 + (Asc(C) - 55)
        End If
        
    Next i
    
    Hex2Dec = val
    
End Function

Public Function Merge(x As Integer, y As Integer) As Long
    
    Dim s As String
    Dim s1 As String
    Dim s2 As String
    Dim i As Integer
    s1 = Hex(x)
    s2 = Hex(y)
    For i = 1 To 2 - Len(s1) Step 1
        s1 = "0" + s1
    Next i
    For i = 1 To 2 - Len(s2) Step 1
        s2 = "0" + s2
    Next i
    s = s1 + s2
    Merge = Hex2Dec(s, 4)
    
End Function

Public Sub Split(s As String, high As Integer, low As Integer)
    Dim i As Integer
    'MsgBox "in split"
    Debug.Print "in split "
    For i = 1 To 4 - Len(s)
        s = "0" + s
    Next i
    high = Hex2Dec(Mid(s, 1, 2), 2)
    low = Hex2Dec(Mid(s, 3, 2), 2)
End Sub

Public Sub Swap(x As Integer, y As Integer)
    Dim temp As Integer
    temp = x
    x = y
    y = temp
End Sub
