Dim sheet  As Worksheet
Dim count  As Long

Sub Main()
    Dim cell   As String
    Dim x, y   As Long
    Dim a, b, c As Long
    Dim align  As Integer
    Dim mode   As Integer
    Set sheet = ActiveSheet
    count = 1
    align = 0
    mode = 0
    x = 1
    y = 2
    Do While True
        cell = sheet.Cells(y, x)
        sheet.Range(sheet.Cells(y, x), sheet.Cells(y, x)).Activate
        
        If cell = "" Then
            GoTo Skip
        End If
        
        If mode = 1 Then
            mode = 0
            GoTo Skip
        End If
        
        If mode = 2 Then
            If cell = Chr(&H22) Then
                mode = 0
                GoTo Skip
            Else
                Push (Asc(cell))
                GoTo Skip
            End If
        End If
        
        Select Case cell
        Case "<"
            align = 1
        Case ">"
            align = 0
        Case "^"
            align = 2
        Case "v"
            align = 3
        Case "_"
            If Pop() = 0 Then
                align = 0
            Else
                align = 1
            End If
        Case "|"
            If Pop() = 0 Then
                align = 2
            Else
                align = 3
            End If
        Case "?"
            align = Int(4 * Rnd)
        Case " "
            align = align
        Case "#"
            mode = 1
        Case "@"
            Exit Do
        Case Chr(&H22)
            mode = 2
        Case "&"
            Push (Val(InputBox("Input Number")))
        Case "~"
            Push (Val(InputBox("Input Char")))
        Case "."
            MsgBox (Str(Pop()) + " ")
        Case ","
            MsgBox (Chr(Val(Pop())))
        Case 0
            Push (cell)
        Case 1
            Push (cell)
        Case 2
            Push (cell)
        Case 3
            Push (cell)
        Case 4
            Push (cell)
        Case 5
            Push (cell)
        Case 6
            Push (cell)
        Case 7
            Push (cell)
        Case 8
            Push (cell)
        Case 9
            Push (cell)
        Case "+"
            a = Pop()
            b = Pop()
            Push (Val(a) + Val(b))
        Case "-"
            a = Pop()
            b = Pop()
            Push (Val(a) - Val(b))
        Case "*"
            a = Pop()
            b = Pop()
            Push (Val(a) * Val(b))
        Case "/"
            a = Pop()
            b = Pop()
            Push (Val(a) / Val(b))
        Case "%"
            a = Pop()
            b = Pop()
            Push (Val(a) Mod Val(b))
        Case "`"
            a = Pop()
            b = Pop()
            If a > b Then
                Push (1)
            Else
                Push (0)
            End If
        Case "!"
            If Pop = 0 Then
                Push (1)
            Else
                Push (0)
            End If
        Case ":"
            a = Pop()
            Push (a)
            Push (a)
        Case "\"
            a = Pop()
            b = Pop()
            Push (b)
            Push (a)
        Case "$"
            a = Pop()
        Case "g"
            a = Pop()
            b = Pop()
            Push (Val(sheet.Cells(a, b)))
        Case "p"
            a = Pop()
            b = Pop()
            c = Pop()
            Sheet1.Cells(a, b) = Str(c)
        End Select
Skip:
        Select Case align
        Case 0
            x = (((x + 1) - 1) Mod sheet.Cells.width) + 1
        Case 1
            x = (((x - 1) - 1) Mod sheet.Cells.width) + 1
        Case 2
            y = (((y - 1) - 1) Mod sheet.Cells.Height) + 1
        Case 3
            y = (((y + 1) - 1) Mod sheet.Cells.Height) + 1
        End Select
    Loop
End Sub

Sub Push(Value As Integer)
    sheet.Cells(1, count) = Value
    count = count + 1
End Sub

Function Pop() As Integer
    count = count - 1
    Pop = sheet.Cells(1, count)
    sheet.Cells(1, count) = ""
End Function
