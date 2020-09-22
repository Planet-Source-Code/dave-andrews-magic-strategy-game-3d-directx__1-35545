Attribute VB_Name = "modLogic"
Option Explicit
'This module contains the AI behind the game.
Const ThoughtLimit = 20

Public Type IntPt
    X As Integer
    Y As Integer
End Type

Public Type MoveSet
    FirstMove As String * 1
    P As IntPt
    Pt As Integer 'place tile ID
    PRow As Boolean 'Does placed Tile form a row?
    PWin As Boolean 'does placed tile win?
    F1 As IntPt
    F2 As IntPt
    Score As Integer
End Type

Type MoveArray
    m() As MoveSet
    Player As Integer
    SetID As Integer
    ParentSetID As Integer
    ParentArray As Integer
    Level As Integer
    Skip As Boolean
End Type

Public Thinking As Boolean
Public AnxLBL As Label
Dim TTime As Double
Dim TString As String
Function BestMove(MyBoard() As Integer, Moves() As MoveArray) As MoveSet
Dim i As Integer
Dim Max As Integer
Dim Choice As Integer
Max = -5000
ExpandMoves MyBoard(), Moves()
For i = 0 To UBound(Moves(0).m)
    If Moves(0).m(i).Score >= Max Then
        'Debug.Print Moves(0).m(i).P.X & "," & Moves(0).m(i).P.X & "," & Moves(0).m(i).Pt
        Max = Moves(0).m(i).Score
        Choice = i
    End If
Next i
BestMove = Moves(0).m(Choice)
'Dim i As Integer
'Dim j As Integer
'Dim Max As Integer
'Dim TArray As MoveArray
'Dim TMove As Integer
'Dim tBoard() As Integer
'Dim TMP As Integer
'ExpandMoves MyBoard(), moves()
'For j = 0 To UBound(moves(0).m)
'    If moves(0).m(j).PWin Then
'        BestMove = moves(0).m(j)
'        Exit Function
'    End If
'Next j
'For i = 0 To UBound(moves)
'    For j = 0 To UBound(moves(i).m)
'        'GetBoard MyBoard, tBoard(), Moves(), Moves(i)
'        'DoMove tBoard(), Moves(i).m(j)
'        'If HasWin(tBoard(), OtherPlayer(Moves(i).Player)) Then
'        '    Moves(i).m(j).Score = -5000
'        'End If
'        If moves(i).m(j).Score >= Max Then
'            If RootMove(moves(), moves(i), j, TMP).Score >= 0 Then
'                Max = moves(i).m(j).Score
'                TArray = moves(i)
'                TMove = j
'            End If
'        End If
'
'    Next j
'Next i
'BestMove = RootMove(moves(), TArray, TMove, TMP)
End Function


Function CalculateMove(ByRef MyBoard() As Integer, Player As Integer, MaxLevel As Integer, ByRef NewMove As MoveSet, ByRef Result() As Integer) As Boolean
Dim MyMoves() As MoveArray
Dim CMove As MoveSet
ReDim MyMoves(0)
Thinking = True
TString = "- T H I N K I N G "
MyMoves(0).SetID = 0
MyMoves(0).ParentSetID = -1
MyMoves(0).ParentArray = -1
MyMoves(0).Level = 0
MyMoves(0).Player = Player
Result() = MyBoard()
If SpaceCount(MyBoard()) >= 10 Then MaxLevel = 0
TTime = Timer
If GetMoves(Result(), Player, MyMoves(0).m()) Then
    RecursiveMove Result(), OtherPlayer(Player), 1, MaxLevel, MyMoves(), 0
    CMove = BestMove(Result(), MyMoves())
    DoMove Result(), CMove
    NewMove = CMove
    CalculateMove = HasWin(Result(), Player)
End If
Thinking = False
End Function

Sub KillHistory(ByRef Moves() As MoveArray, PID As Integer, PA As Integer)
Dim i As Integer
Dim j As Integer
Dim TMP As MoveSet
For i = 0 To UBound(Moves)
    If Moves(i).ParentArray = PA And Moves(i).ParentSetID = PID Then
        For j = 0 To UBound(Moves(i).m)
            ScoreRoot Moves(), i, j, -5000
        Next j
    End If
Next i
End Sub
Sub ExpandMoves(MyBoard() As Integer, ByRef Moves() As MoveArray)
'Using the GetMoves function,  we only itterate one tile placement, we have to double each array to account for the other side of the tile
On Local Error GoTo eTrap
Dim tBoard() As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim m As Integer
Dim R As Integer
Dim RES() As IntPt
Dim NewPT() As IntPt
Dim TMove As MoveSet
For i = 0 To UBound(Moves)
    k = UBound(Moves(i).m)
    For R = 0 To k
        'Next we set the flags for rows and wins
        GetBoard MyBoard(), tBoard(), Moves(), Moves(i)
        DoMove tBoard(), Moves(i).m(R)
        Erase RES()
        If InARow(tBoard(), Moves(i).m(R).P.X, Moves(i).m(R).P.Y, RES()) Then
            Moves(i).m(R).PRow = True
            If Moves(i).Player = Moves(0).Player Then
                ScoreRoot Moves(), i, R, (1 * (10 - Moves(i).Level))
                'moves(i).m(R).Score = moves(i).m(R).Score + (5 * (10 - moves(i).Level))
            Else
                ScoreRoot Moves(), i, R, -(1 * (10 - Moves(i).Level))
                'moves(i).m(R).Score = moves(i).m(R).Score - (5 * (10 - moves(i).Level))
            End If
            If HasWin(tBoard(), Moves(i).Player) Then
                Moves(i).m(R).PWin = True
            End If
            If Moves(i).m(R).PWin Then
                If Moves(i).Player = Moves(0).Player Then
                    ScoreRoot Moves(), i, R, (5 * (10 - Moves(i).Level))
                    'moves(i).m(R).Score = moves(i).m(R).Score + (50 * (10 - moves(i).Level))
                Else
                    ScoreRoot Moves(), i, R, -(5 * (100 - Moves(i).Level))
                    KillHistory Moves(), R, i
                    'moves(i).m(R).Score = moves(i).m(R).Score - (50 * (10 - moves(i).Level))
                End If
            End If
        End If
    Next R
Next i
Exit Sub
eTrap:
    k = -1
End Sub
Sub FixANX(Optional TXT As String)
TString = Right(TString, Len(TString) - 1) & Left(TString, 1)
AnxLBL.Caption = TString & TXT
frmBoard.dxBoard.Update
DoEvents
End Sub

Sub GetBoard(Root() As Integer, ByRef Result() As Integer, Moves() As MoveArray, MyArray As MoveArray)
Dim APath() As MoveArray
Dim TSet As MoveSet
Dim i As Integer
Dim j As Integer
Dim k As Integer
i = MyArray.ParentSetID
j = MyArray.ParentArray
If i = -1 Then
    Result() = Root()
    Exit Sub
End If
Do While i <> -1
    ReDim Preserve APath(k)
    APath(k) = Moves(j)
    TSet = Moves(j).m(i)
    ReDim APath(k).m(0)
    APath(k).m(0) = TSet
    i = APath(k).ParentSetID
    j = APath(k).ParentArray
    k = k + 1
Loop
'Now do the moves
Result() = Root()
For i = UBound(APath) To 0 Step -1
    DoMove Result(), APath(i).m(0)
Next i
End Sub

Function HasARow(MyBoard() As Integer, Player As Integer)
Dim X As Integer
Dim Y As Integer
Dim RES() As IntPt
For X = 0 To 3
    For Y = 0 To 3
        If MyBoard(X, Y) = Player Or MyBoard(X, Y) = Player + 2 Then
            If InARow(MyBoard(), X, Y, RES()) Then
                HasARow = True
                Exit Function
            End If
        End If
    Next Y
Next X
            
End Function

Function HasWin(MyBoard() As Integer, Player As Integer) As Boolean
Dim RES() As IntPt
Dim NewPT() As IntPt
Dim chk() As IntPt
Dim tBoard() As Integer
Dim i As Integer
Dim j As Integer
If RequiredFlips(MyBoard(), Player, RES()) Then
    HasWin = True
    For i = 0 To UBound(RES)
        Erase NewPT()
        If CanFlip(MyBoard(), RES(i).X, RES(i).Y, NewPT()) Then
            For j = 0 To UBound(NewPT)
                Erase chk()
                tBoard() = MyBoard()
                tBoard(NewPT(j).X, NewPT(j).Y) = Rev(tBoard(RES(i).X, RES(i).Y))
                tBoard(RES(i).X, RES(i).Y) = 0
                If Not RequiredFlips(tBoard(), Player, chk()) Then
                    HasWin = False
                    Exit Function
                End If
            Next j
        End If
    Next i
End If
End Function

Public Function IsFlippable(MyBoard() As Integer, Player As Integer)
Dim NewPT() As IntPt
Dim X As Integer
Dim Y As Integer
For X = 0 To 3
    For Y = 0 To 3
        If MyBoard(X, Y) = Player Or MyBoard(X, Y) = Player + 2 Then
            If CanFlip(MyBoard(), X, Y, NewPT()) Then
                IsFlippable = True
                Exit Function
            End If
        End If
    Next Y
Next X
                
End Function

Sub RemoveLeastSpaces(ByRef MyMoves() As MoveSet)
Dim TMoves() As MoveSet
Dim TMP As MoveSet
Dim i As Integer
Dim k As Integer
Dim BSpace As Boolean
Dim Pt As IntPt
TMP = MyMoves(0)
For i = 0 To UBound(MyMoves)
    BSpace = True
    Pt = MyMoves(i).P
    If TUp(Pt.Y) Then BSpace = False
    If TDown(Pt.Y) Then BSpace = False
    If TLeft(Pt.X) Then BSpace = False
    If TRight(Pt.X) Then BSpace = False
    If TUpLeft(Pt.X, Pt.Y) Then BSpace = False
    If TDownLeft(Pt.X, Pt.Y) Then BSpace = False
    If TUpRight(Pt.X, Pt.Y) Then BSpace = False
    If TDownRight(Pt.X, Pt.Y) Then BSpace = False
    If Not BSpace Then
        ReDim Preserve TMoves(k)
        TMoves(k) = MyMoves(i)
        k = k + 1
    End If
Next i
If k = 0 Then
    ReDim MyMoves(k)
    MyMoves(k) = TMP
Else
    MyMoves = TMoves()
End If
End Sub

Function RequiredFlips(MyBoard() As Integer, Player As Integer, ByRef RES() As IntPt) As Boolean
On Local Error GoTo eTrap
Dim X As Integer
Dim Y As Integer
Dim NewPT() As IntPt
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim HasDup As Boolean
For X = 0 To 3
    For Y = 0 To 3
        If MyBoard(X, Y) = Player Or MyBoard(X, Y) = Player + 2 Then
            Erase NewPT()
            If InARow(MyBoard(), X, Y, NewPT()) Then
                RequiredFlips = True
                For i = 0 To UBound(NewPT)
                    k = UBound(RES)
                    HasDup = False
                    For j = 0 To k
                        If RES(j).X = NewPT(i).X And RES(j).Y = NewPT(i).Y Then
                            HasDup = True
                            Exit For
                        End If
                    Next j
                    If Not HasDup Then
                        k = k + 1
                        ReDim Preserve RES(k)
                        RES(k) = NewPT(i)
                    End If
                Next i
            End If
        End If
    Next Y
Next X
Exit Function
eTrap:
    k = -1
    Resume Next
End Function

Function RootMove(Moves() As MoveArray, MyArray As MoveArray, MoveID As Integer, ByRef RootID As Integer) As MoveSet
Dim APath() As MoveArray
Dim TSet As MoveSet
Dim i As Integer
Dim j As Integer
Dim k As Integer
i = MyArray.ParentSetID
j = MyArray.ParentArray
If i = -1 Then
    RootMove = MyArray.m(MoveID)
    Exit Function
End If
Do While i <> -1
    ReDim Preserve APath(k)
    APath(k) = Moves(j)
    TSet = Moves(j).m(i)
    RootID = i
    ReDim APath(k).m(0)
    APath(k).m(0) = TSet
    i = APath(k).ParentSetID
    j = APath(k).ParentArray
    k = k + 1
Loop
i = UBound(APath)
RootMove = APath(i).m(0)

End Function



Function RecursiveMove(MyBoard() As Integer, Player As Integer, CurLevel As Integer, MaxLevel As Integer, ByRef Moves() As MoveArray, Start As Integer) As Boolean
On Local Error GoTo eTrap
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim tBoard() As Integer
Dim NewStart As Integer
If Timer() - TTime > ThoughtLimit * MaxLevel Then Exit Function
If CurLevel > MaxLevel Then Exit Function
i = UBound(Moves) + 1
NewStart = i
j = Start
FixANX "   " & Long2Time(Timer() - TTime)
For k = 0 To UBound(Moves(j).m)
    tBoard() = MyBoard()
    DoMove tBoard(), Moves(j).m(k)
    If Not HasWin(tBoard(), Moves(j).Player) Then
        ReDim Preserve Moves(i)
        Moves(i).SetID = i
        Moves(i).ParentSetID = k
        Moves(i).ParentArray = j
        Moves(i).Player = Player
        Moves(i).Level = CurLevel
        GetMoves tBoard(), Player, Moves(i).m()
        'RecursiveMove tBoard(), OtherPlayer(Player), CurLevel + 1, MaxLevel, Moves(), i
        i = i + 1
    Else
        Moves(i).Skip = True
    End If
Next k
''Added to make calc's faster
k = UBound(Moves)
For i = NewStart To k
    If Not Moves(i).Skip Then
        For j = 0 To UBound(Moves(i).m)
            tBoard() = MyBoard()
            DoMove tBoard(), Moves(i).m(j)
            RecursiveMove tBoard(), Player, CurLevel + 1, MaxLevel, Moves(), i
        Next j
    End If
Next i
Exit Function
eTrap:
    i = 0
End Function
Function Long2Time(MyTime As Long) As String
Dim a As Integer
Dim B As Integer
Dim c As Integer
a = MyTime 'Seconds of play
B = a
c = 0
'If b > 60 Then
'    c = 1
'Else
'    c = 0
'End If
Do While (B / 60) >= 1
    B = B - 60
    c = c + 1
Loop
Long2Time = Format(c, "#") & ":" & Format(B Mod 60, "0#")
End Function

Function CanFlip(MyBoard() As Integer, X As Integer, Y As Integer, ByRef NewPT() As IntPt) As Boolean
'This checks to see if a specified tile can be flipped
'and returns an array of values of possible flips
Dim X2 As Integer
Dim Y2 As Integer
Dim RES() As IntPt
Dim tBoard() As Integer
Dim T As Integer
Dim Count As Integer
T = MyBoard(X, Y)
'Check Flip Up
X2 = X
Y2 = Y
If TUp(Y2) Then
    If MyBoard(X2, Y2) = 0 Then
        tBoard() = MyBoard()
        tBoard(X, Y) = 0
        tBoard(X2, Y2) = Rev(T)
        If Not InARow(tBoard(), X2, Y2, RES()) Then
            ReDim Preserve NewPT(Count)
            NewPT(Count).X = X2
            NewPT(Count).Y = Y2
            Count = Count + 1
        End If
    End If
End If
'Check Flip Down
X2 = X
Y2 = Y
If TDown(Y2) Then
    If MyBoard(X2, Y2) = 0 Then
        tBoard() = MyBoard()
        tBoard(X, Y) = 0
        tBoard(X2, Y2) = Rev(T)
        If Not InARow(tBoard(), X2, Y2, RES()) Then
            ReDim Preserve NewPT(Count)
            NewPT(Count).X = X2
            NewPT(Count).Y = Y2
            Count = Count + 1
        End If
    End If
End If
'Check Flip Left
X2 = X
Y2 = Y
If TLeft(X2) Then
    If MyBoard(X2, Y2) = 0 Then
        tBoard() = MyBoard()
        tBoard(X, Y) = 0
        tBoard(X2, Y2) = Rev(T)
        If Not InARow(tBoard(), X2, Y2, RES()) Then
            ReDim Preserve NewPT(Count)
            NewPT(Count).X = X2
            NewPT(Count).Y = Y2
            Count = Count + 1
        End If
    End If
End If
'Check Flip Right
X2 = X
Y2 = Y
If TRight(X2) Then
    If MyBoard(X2, Y2) = 0 Then
        tBoard() = MyBoard()
        tBoard(X, Y) = 0
        tBoard(X2, Y2) = Rev(T)
        If Not InARow(tBoard(), X2, Y2, RES()) Then
            ReDim Preserve NewPT(Count)
            NewPT(Count).X = X2
            NewPT(Count).Y = Y2
            Count = Count + 1
        End If
    End If
End If
If Count > 0 Then
    CanFlip = True
Else
    CanFlip = False
End If
End Function

Function CheckVert(MyBoard() As Integer, X As Integer, Y As Integer, ByRef RES() As IntPt) As Boolean
'This check for a vertical 3-in-a-row using the specified board and the specified tile
Dim X2 As Integer
Dim Y2 As Integer
Dim T As Integer
Dim Count As Integer
T = MyBoard(X, Y)
If T = 0 Then Exit Function ' there's no tile at the location specified
ReDim Preserve RES(Count)
RES(Count).X = X
RES(Count).Y = Y
Count = Count + 1
'Check tile up
X2 = X
Y2 = Y
If TUp(Y2) Then
    If MyBoard(X, Y2) = T Then 'we have 2 tiles in a row with the same side
        ReDim Preserve RES(Count)
        RES(Count).X = X2
        RES(Count).Y = Y2
        Count = Count + 1
        'Check one more up
        If TUp(Y2) Then
            If MyBoard(X, Y2) = T Then 'we have 3 tiles in a row with the same side
                ReDim Preserve RES(Count)
                RES(Count).X = X2
                RES(Count).Y = Y2
                Count = Count + 1
                'Check one more up
                If TUp(Y2) Then
                    If MyBoard(X, Y2) = T Then 'we have 3 tiles in a row with the same side
                        ReDim Preserve RES(Count)
                        RES(Count).X = X2
                        RES(Count).Y = Y2
                        Count = Count + 1
                    End If
                End If
            End If
        End If
    End If
End If
'Check tile down
X2 = X
Y2 = Y
If TDown(Y2) Then
    If MyBoard(X, Y2) = T Then 'we have 2 tiles in a row with the same side
        ReDim Preserve RES(Count)
        RES(Count).X = X2
        RES(Count).Y = Y2
        Count = Count + 1
        'Check one more down
        If TDown(Y2) Then
            If MyBoard(X, Y2) = T Then 'we have 3 tiles in a row with the same side
                ReDim Preserve RES(Count)
                RES(Count).X = X2
                RES(Count).Y = Y2
                Count = Count + 1
                'Check one more down
                If TDown(Y2) Then
                    If MyBoard(X, Y2) = T Then 'we have 3 tiles in a row with the same side
                        ReDim Preserve RES(Count)
                        RES(Count).X = X2
                        RES(Count).Y = Y2
                        Count = Count + 1
                    End If
                End If
            End If
        End If
    End If
End If
If Count >= 3 Then
    CheckVert = True
Else
    CheckVert = False
    Erase RES()
End If
End Function
Function CheckHorz(MyBoard() As Integer, X As Integer, Y As Integer, ByRef RES() As IntPt) As Boolean
'This check for a horizontal 3-in-a-row using the specified board and the specified tile
Dim X2 As Integer
Dim Y2 As Integer
Dim T As Integer
Dim Count As Integer
T = MyBoard(X, Y)
If T = 0 Then Exit Function ' there's no tile at the location specified
ReDim Preserve RES(Count)
RES(Count).X = X
RES(Count).Y = Y
Count = Count + 1
'Check tile left
X2 = X
Y2 = Y
If TLeft(X2) Then
    If MyBoard(X2, Y) = T Then 'we have 2 tiles in a row with the same side
        ReDim Preserve RES(Count)
        RES(Count).X = X2
        RES(Count).Y = Y2
        Count = Count + 1
        'Check one more left
        If TLeft(X2) Then
            If MyBoard(X2, Y) = T Then 'we have 3 tiles in a row with the same side
                ReDim Preserve RES(Count)
                RES(Count).X = X2
                RES(Count).Y = Y2
                Count = Count + 1
                'Check one more left
                If TLeft(X2) Then
                    If MyBoard(X2, Y) = T Then 'we have 3 tiles in a row with the same side
                        ReDim Preserve RES(Count)
                        RES(Count).X = X2
                        RES(Count).Y = Y2
                        Count = Count + 1
                    End If
                End If
            End If
        End If
    End If
End If
'Check tile right
X2 = X
Y2 = Y
If TRight(X2) Then
    If MyBoard(X2, Y) = T Then 'we have 2 tiles in a row with the same side
        ReDim Preserve RES(Count)
        RES(Count).X = X2
        RES(Count).Y = Y2
        Count = Count + 1
        'Check one more right
        If TRight(X2) Then
            If MyBoard(X2, Y) = T Then 'we have 3 tiles in a row with the same side
                ReDim Preserve RES(Count)
                RES(Count).X = X2
                RES(Count).Y = Y2
                Count = Count + 1
                'Check one more right
                If TRight(X2) Then
                    If MyBoard(X2, Y) = T Then 'we have 3 tiles in a row with the same side
                        ReDim Preserve RES(Count)
                        RES(Count).X = X2
                        RES(Count).Y = Y2
                        Count = Count + 1
                    End If
                End If
            End If
        End If
    End If
End If
If Count >= 3 Then
    CheckHorz = True
Else
    CheckHorz = False
    Erase RES()

End If
End Function
Function CheckDiag(MyBoard() As Integer, X As Integer, Y As Integer, ByRef RES() As IntPt) As Boolean
'This check for a diagonal 3-in-a-row using the specified board and the specified tile
Dim X2 As Integer
Dim Y2 As Integer
Dim T As Integer
Dim Count As Integer
T = MyBoard(X, Y)
If T = 0 Then Exit Function ' there's no tile at the location specified
'FIRST we will check up-left to down-right
ReDim Preserve RES(Count)
RES(Count).X = X
RES(Count).Y = Y
Count = Count + 1
'Check tile upleft
Y2 = Y
X2 = X
If TUpLeft(X2, Y2) Then
    If MyBoard(X2, Y2) = T Then 'we have 2 tiles in a row with the same side
        ReDim Preserve RES(Count)
        RES(Count).X = X2
        RES(Count).Y = Y2
        Count = Count + 1
        'Check one more up left
        If TUpLeft(X2, Y2) Then
            If MyBoard(X2, Y2) = T Then 'we have 3 tiles in a row with the same side
                ReDim Preserve RES(Count)
                RES(Count).X = X2
                RES(Count).Y = Y2
                Count = Count + 1
                'Check one more up left
                If TUpLeft(X2, Y2) Then
                    If MyBoard(X2, Y2) = T Then 'we have 3 tiles in a row with the same side
                        ReDim Preserve RES(Count)
                        RES(Count).X = X2
                        RES(Count).Y = Y2
                        Count = Count + 1
                    End If
                End If
            End If
        End If
    End If
End If
'Check tile down right
X2 = X
Y2 = Y
If TDownRight(X2, Y2) Then
    If MyBoard(X2, Y2) = T Then 'we have 2 tiles in a row with the same side
        ReDim Preserve RES(Count)
        RES(Count).X = X2
        RES(Count).Y = Y2
        Count = Count + 1
        'Check one more down right
        If TDownRight(X2, Y2) Then
            If MyBoard(X2, Y2) = T Then 'we have 3 tiles in a row with the same side
                ReDim Preserve RES(Count)
                RES(Count).X = X2
                RES(Count).Y = Y2
                Count = Count + 1
                'Check one more down right
                If TDownRight(X2, Y2) Then
                    If MyBoard(X2, Y2) = T Then 'we have 3 tiles in a row with the same side
                        ReDim Preserve RES(Count)
                        RES(Count).X = X2
                        RES(Count).Y = Y2
                        Count = Count + 1
                    End If
                End If
            End If
        End If
    End If
End If
If Count >= 3 Then
    CheckDiag = True
    Exit Function
Else
    Count = 0
End If
'SECOND we will check up-Right to down-left
ReDim Preserve RES(Count)
RES(Count).X = X
RES(Count).Y = Y
Count = Count + 1
'Check tile upleft
Y2 = Y
X2 = X
If TUpRight(X2, Y2) Then
    If MyBoard(X2, Y2) = T Then 'we have 2 tiles in a row with the same side
        ReDim Preserve RES(Count)
        RES(Count).X = X2
        RES(Count).Y = Y2
        Count = Count + 1
        'Check one more up right
        If TUpRight(X2, Y2) Then
            If MyBoard(X2, Y2) = T Then 'we have 3 tiles in a row with the same side
                ReDim Preserve RES(Count)
                RES(Count).X = X2
                RES(Count).Y = Y2
                Count = Count + 1
                'Check one more up right
                If TUpRight(X2, Y2) Then
                    If MyBoard(X2, Y2) = T Then 'we have 3 tiles in a row with the same side
                        ReDim Preserve RES(Count)
                        RES(Count).X = X2
                        RES(Count).Y = Y2
                        Count = Count + 1
                    End If
                End If
            End If
        End If
    End If
End If
'Check tile down left
X2 = X
Y2 = Y
If TDownLeft(X2, Y2) Then
    If MyBoard(X2, Y2) = T Then 'we have 2 tiles in a row with the same side
        ReDim Preserve RES(Count)
        RES(Count).X = X2
        RES(Count).Y = Y2
        Count = Count + 1
        'Check one more down left
        If TDownLeft(X2, Y2) Then
            If MyBoard(X2, Y2) = T Then 'we have 3 tiles in a row with the same side
                ReDim Preserve RES(Count)
                RES(Count).X = X2
                RES(Count).Y = Y2
                Count = Count + 1
                'Check one more down left
                If TDownLeft(X2, Y2) Then
                    If MyBoard(X2, Y2) = T Then 'we have 3 tiles in a row with the same side
                        ReDim Preserve RES(Count)
                        RES(Count).X = X2
                        RES(Count).Y = Y2
                        Count = Count + 1
                    End If
                End If
            End If
        End If
    End If
End If
If Count >= 3 Then
    CheckDiag = True
Else
    CheckDiag = False
    Erase RES()
End If

End Function

Sub DoMove(ByRef MyBoard() As Integer, MyMove As MoveSet)
Select Case MyMove.FirstMove
    Case "F"
        MyBoard(MyMove.F2.X, MyMove.F2.Y) = Rev(MyBoard(MyMove.F1.X, MyMove.F1.Y))
        MyBoard(MyMove.F1.X, MyMove.F1.Y) = 0
        MyBoard(MyMove.P.X, MyMove.P.Y) = MyMove.Pt
    Case "P"
        MyBoard(MyMove.P.X, MyMove.P.Y) = MyMove.Pt
        MyBoard(MyMove.F2.X, MyMove.F2.Y) = Rev(MyBoard(MyMove.F1.X, MyMove.F1.Y))
        MyBoard(MyMove.F1.X, MyMove.F1.Y) = 0
    Case "O"
        MyBoard(MyMove.P.X, MyMove.P.Y) = MyMove.Pt
End Select
End Sub

Sub GetFlips(MyBoard() As Integer, Player As Integer, ByRef Moves() As MoveSet)
Dim T As Integer
Dim i As Integer
Dim j As Integer
Dim Root As Integer
Dim X As Integer
Dim Y As Integer
Dim RES() As IntPt
Dim Count As Integer
i = UBound(Moves)
Root = i
T = Player
For X = 0 To 3
    For Y = 0 To 3
        'For T = player to player + 2 step 2
            If MyBoard(X, Y) = T Or MyBoard(X, Y) = T + 2 Then
                Erase RES()
                If CanFlip(MyBoard(), X, Y, RES()) Then
                    For j = 0 To UBound(RES)
                        ReDim Preserve Moves(i)
                        Moves(i) = Moves(Root)
                        Moves(i).F1.X = X
                        Moves(i).F1.Y = Y
                        Moves(i).F2 = RES(j)
                        i = i + 1
                        Count = Count + 1
                    Next j
                End If
            End If
        'Next T
    Next Y
Next X
If Count = 0 Then
    'The placement that occured does not have any possible flips afterwards.
    Moves(Root).FirstMove = "O"
End If
End Sub

Function GetMoves(MyBoard() As Integer, Player As Integer, ByRef Moves() As MoveSet) As Boolean
On Local Error GoTo eTrap
Dim T As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim P As Integer
Dim X As Integer
Dim Y As Integer
Dim RES() As IntPt
Dim NewPT() As IntPt
Dim tBoard() As Integer
Dim TMoves() As MoveSet
Erase Moves()
'First we try placing tiles, then flipping
If RequiredFlips(MyBoard(), OtherPlayer(Player), RES()) Then
    For i = 0 To UBound(RES)
        Erase NewPT()
        If CanFlip(MyBoard(), RES(i).X, RES(i).Y, NewPT()) Then
            For j = 0 To UBound(NewPT)
                ReDim Preserve TMoves(k)
                TMoves(k).F1 = RES(i)
                TMoves(k).F2 = NewPT(j)
                k = k + 1
            Next j
        End If
    Next i
    k = 0
    For i = 0 To UBound(TMoves)
        tBoard() = MyBoard()
        tBoard(TMoves(i).F2.X, TMoves(i).F2.Y) = Rev(tBoard(TMoves(i).F1.X, TMoves(i).F1.Y))
        tBoard(TMoves(i).F1.X, TMoves(i).F1.Y) = 0
        For X = 0 To 3
            For Y = 0 To 3
                If tBoard(X, Y) = 0 Then
                    For T = Player To Player + 2 Step 2
                        ReDim Preserve Moves(k)
                        Moves(k) = TMoves(i)
                        Moves(k).Pt = T
                        Moves(k).P.X = X
                        Moves(k).P.Y = Y
                        Moves(k).FirstMove = "F"
                        k = k + 1
                    Next T
                End If
            Next Y
        Next X
    Next i
Else 'the opponent does not have a row that must be flipped
    For X = 0 To 3
        For Y = 0 To 3
            If MyBoard(X, Y) = 0 Then
                '-------------------
                tBoard() = MyBoard()
                tBoard(X, Y) = Player
                If Not IsFlippable(tBoard(), OtherPlayer(Player)) Then
                '*******************
                    For T = Player To Player + 2 Step 2
                        ReDim Preserve Moves(i)
                        Moves(i).Pt = T
                        Moves(i).P.X = X
                        Moves(i).P.Y = Y
                        Moves(i).FirstMove = "O"
                        tBoard() = MyBoard()
                        tBoard(X, Y) = T
                        'GetFlips tBoard(), OtherPlayer(Player), Moves()
                        i = UBound(Moves) + 1
                    Next T
                '--------------
                End If
                '**************
            End If
        Next Y
    Next X
    'next we flip tiles first, then place
    ReDim TMoves(0) ' gotta pass the function a non-empty array
    i = UBound(Moves) + 1
    GetFlips MyBoard(), OtherPlayer(Player), TMoves()
    If TMoves(0).FirstMove <> "O" Then
        For j = 0 To UBound(TMoves) 'itterate through newly flipped moves
            'flip the tile
            tBoard() = MyBoard()
            tBoard(TMoves(j).F2.X, TMoves(j).F2.Y) = Rev(tBoard(TMoves(j).F1.X, TMoves(j).F1.Y))
            tBoard(TMoves(j).F1.X, TMoves(j).F1.Y) = 0
            For X = 0 To 3
                For Y = 0 To 3
                    If tBoard(X, Y) = 0 Then
                        For T = Player To Player + 2 Step 2 ' both sides
                            ReDim Preserve Moves(i)
                            Moves(i) = TMoves(j)
                            Moves(i).Pt = T
                            Moves(i).P.X = X
                            Moves(i).P.Y = Y
                            Moves(i).FirstMove = "F"
                            i = i + 1
                        Next T
                    End If
                Next Y
            Next X
        Next j
    End If
End If
'RemoveDuplicateMoves Moves()
'RemoveLeastSpaces Moves()
GetMoves = True
Exit Function
eTrap:
    i = 0
    Resume Next
End Function
Function InARow(MyBoard() As Integer, X As Integer, Y As Integer, ByRef RES() As IntPt) As Boolean

If CheckVert(MyBoard(), X, Y, RES()) Then
    InARow = True
    Exit Function
End If
If CheckHorz(MyBoard(), X, Y, RES()) Then
    InARow = True
    Exit Function
End If
If CheckDiag(MyBoard(), X, Y, RES()) Then
    InARow = True
    Exit Function
End If
End Function

Function OtherPlayer(Player As Integer) As Integer
If Player Mod 2 = 0 Then
    OtherPlayer = 1
Else
    OtherPlayer = 2
End If
End Function

Sub RemoveDuplicateMoves(ByRef Moves() As MoveSet)
Dim TMoves() As MoveSet
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim Match As Boolean

Dim chk As Integer
chk = UBound(Moves)
For i = 0 To UBound(Moves) - 1
    Match = False
    For j = i + 1 To UBound(Moves)
        If Moves(i).F1.X = Moves(j).F1.X And Moves(i).F2.X = Moves(j).F2.X And Moves(i).F1.Y = Moves(j).F1.Y And Moves(i).F2.Y = Moves(j).F2.Y And Moves(i).P.X = Moves(j).P.X And Moves(i).P.Y = Moves(j).P.Y And Moves(i).Pt = Moves(j).Pt Then
            Match = True
            Exit For
        End If
    Next j
    If Not Match Then
        ReDim Preserve TMoves(k)
        TMoves(k) = Moves(i)
        k = k + 1
    End If
Next i
'Always get the last one
ReDim Preserve TMoves(k)
TMoves(k) = Moves(UBound(Moves))
'Copy the array
Moves() = TMoves()
If chk <> UBound(Moves) Then
    MsgBox "HEY"
End If
End Sub

Function Rev(T As Integer) As Integer
'Returns the reverse side value of a tile
If T > 2 Then
    Rev = T - 2
Else
    Rev = T + 2
End If
End Function

Sub ScoreRoot(ByRef Moves() As MoveArray, PID As Integer, SID As Integer, Score As Integer)
Dim TMP As MoveSet
Dim i As Integer
TMP = RootMove(Moves(), Moves(PID), SID, i)
Moves(0).m(i).Score = Moves(0).m(i).Score + Score
End Sub

Function SpaceCount(MyBoard() As Integer) As Integer
Dim X As Integer
Dim Y As Integer
For X = 0 To 3
    For Y = 0 To 3
        If MyBoard(X, Y) = 0 Then SpaceCount = SpaceCount + 1
    Next Y
Next X
End Function


Function TUp(ByRef Y As Integer) As Boolean
If Y > 0 Then
    TUp = True
    Y = Y - 1
End If
End Function
Function TUpLeft(ByRef X As Integer, ByRef Y As Integer) As Boolean
If Y > 0 And X > 0 Then
    TUpLeft = True
    X = X - 1
    Y = Y - 1
End If
End Function
Function TUpRight(ByRef X As Integer, ByRef Y As Integer) As Boolean
If Y > 0 And X < 3 Then
    TUpRight = True
    X = X + 1
    Y = Y - 1
End If
End Function

Function TDownRight(ByRef X As Integer, ByRef Y As Integer) As Boolean
If Y < 3 And X < 3 Then
    TDownRight = True
    X = X + 1
    Y = Y + 1
End If
End Function
Function TDownLeft(ByRef X As Integer, ByRef Y As Integer) As Boolean
If Y < 3 And X > 0 Then
    TDownLeft = True
    X = X - 1
    Y = Y + 1
End If
End Function
Function TDown(ByRef Y As Integer) As Boolean
If Y < 3 Then
    TDown = True
    Y = Y + 1
End If
End Function
Function TLeft(ByRef X As Integer) As Boolean
If X > 0 Then
    TLeft = True
    X = X - 1
End If
End Function
Function TRight(ByRef X As Integer) As Boolean
If X < 3 Then
    TRight = True
    X = X + 1
End If
End Function


