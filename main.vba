Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public s() As Range 'snake
Public t As Range 'tail
Public ws As Worksheet
Public ws2 As Worksheet
Public score As Long
Public l As Integer ' length of snake
Public x As Integer 'snake coordinates
Public y As Integer 'snake coordinates
Public i As Integer ' misc counter
Public play As Boolean 'play = false - GAME OVER
Public fieldSize As Long ' dimensions of playing area
Public grow As Boolean ' switches on when food is eaten to add another cell to snake
Dim food As Range ' cell with food in
Dim foodx As Integer 'food coordinates
Dim foody As Integer 'food coordinates

'This is how I get Keyboard Input
Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long

'-- Constant Values --
Const VK_LEFT = &H25, VK_UP = &H26, VK_RIGHT = &H27, VK_DOWN = &H28

Sub snake()

Set ws2 = ActiveWorkbook.Sheets("sheet2")
Set ws = ActiveWorkbook.Sheets("sheet1")
ws.Activate
ReDim s(0)
Set s(0) = ws.Cells(10, 10)

Dim speed As Long
speed = 100 - ws2.Range("C5").Value
fieldSize = ws2.Range("C8").Value

Call clearField
Call drawField
Call drawFood
l = 0
x = 1
y = 0
i = 0
play = True
grow = False

On Error Resume Next
Do While play = True




    For i = UBound(s) To (LBound(s) + 1) Step -1
          Set s(i) = s(i - 1)
    Next i
    
    Set s(0) = s(0).Offset(y, x)
    If s(0).Interior.ColorIndex = 1 Then Call youLose
    For i = LBound(s) To UBound(s)
       s(i).Interior.ColorIndex = 1
    Next i
      
    
    Sleep (speed)
    If GetAsyncKeyState(VK_UP) <> 0 Then
        x = 0
        y = -1
    ElseIf GetAsyncKeyState(VK_DOWN) <> 0 Then
        x = 0
        y = 1
    ElseIf GetAsyncKeyState(VK_LEFT) <> 0 Then
        x = -1
        y = 0
    ElseIf GetAsyncKeyState(VK_RIGHT) <> 0 Then
        x = 1
        y = 0
    End If

    s(UBound(s)).Interior.ColorIndex = xlNone
    If s(0).Column = food.Column And s(0).Row = food.Row Then
        Call eatFood
        Set s(0) = food
        Call drawFood
    End If
    DoEvents
    For i = LBound(s) To UBound(s)
        Debug.Print i & " + " & s(i).Address
    Next i
Loop

End Sub

Private Sub drawField()

For i = 1 To fieldSize
    ws.Cells(i, 1).Interior.Color = Black
    ws.Cells(1, i).Interior.Color = Black
    ws.Cells(fieldSize, i).Interior.Color = Black
    ws.Cells(i, fieldSize).Interior.Color = Black
Next i

End Sub
Private Sub clearField()
    ws.Range("A1:" & ws.Cells(fieldSize - 1, fieldSize - 1).Address).Interior.ColorIndex = xlNone
End Sub

Private Sub drawFood()

foodx = Int(((fieldSize - 1) - 2 + 1) * Rnd + 2)
foody = Int(((fieldSize - 1) - 2 + 1) * Rnd + 2)
Set food = ws.Cells(foody, foodx)
food.Interior.ColorIndex = 3

End Sub
Private Sub eatFood()

l = UBound(s) + 1
ReDim Preserve s(l)

For i = UBound(s) To LBound(s) Step -1
    Set s(i) = s(i - 1)
Next i





End Sub
Private Sub youLose()
        
        score = s(UBound(s)) * 100
        ws2.Range("F5").Value = score
        MsgBox "You lose"
        play = False
        ws2.Activate
        
End Sub

Private Function queueSnake(ByRef foodrng As Range, ByVal snakeArry As Variant)
Debug.Print snakeArry
For i = LBound(snakeArry) To UBound(snakeArry)
    l = UBound(snakeArry) + 1
    ReDim snakeArry(l)
    Set snakeArry(i + 1) = snakeArry(i)
    Set snakeArry(0) = foodrng
Next i

End Function
