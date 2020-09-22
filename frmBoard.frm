VERSION 5.00
Begin VB.Form frmBoard 
   Caption         =   "Strategy Square"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   Icon            =   "frmBoard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShowOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      Picture         =   "frmBoard.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   255
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   4320
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
      Begin VB.CommandButton cmdPlayDown 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         Picture         =   "frmBoard.frx":0355
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmdPlayUp 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         Picture         =   "frmBoard.frx":03B3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.Frame fraLevel 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1575
         Begin VB.CommandButton cmdLevelUp 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   1080
            Picture         =   "frmBoard.frx":0412
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton cmdLevelDown 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   1080
            Picture         =   "frmBoard.frx":0471
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Level"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label lblLevel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Label lblFirst 
         BackStyle       =   0  'Transparent
         Caption         =   "Player Goes First"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   780
         Width           =   1455
      End
      Begin VB.Label lblStart 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "START"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1950
         Width           =   1215
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   120
         Y1              =   128
         Y2              =   128
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   192
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   120
         Y1              =   72
         Y2              =   72
      End
      Begin VB.Label lblPlayers 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Players"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox picANX 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   436
      TabIndex        =   1
      Top             =   0
      Width           =   6570
      Begin VB.CheckBox chkPlace 
         BackColor       =   &H00000000&
         Caption         =   "PLACE"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   5640
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkFlip 
         BackColor       =   &H00000000&
         Caption         =   "FLIP"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   5640
         TabIndex        =   16
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblANX 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4455
      End
   End
   Begin VB.Timer Timer1 
      Left            =   5040
      Top             =   840
   End
   Begin VB.PictureBox picBoard 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4320
      Left            =   0
      ScaleHeight     =   286
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   0
      Top             =   495
      Width           =   3975
   End
End
Attribute VB_Name = "frmBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const pi = 3.141592654
Const dZ = 0
Dim CompLevel As Integer
Public dxBoard As New clsDX73D
Dim TFrame As Direct3DRMFrame3
Dim SFrame As Direct3DRMFrame3
Dim CurPlayer As Integer
Dim CompNum As Integer
Dim FirstPlayer As Integer
Dim totPlayers As Integer
Dim Rotating As Boolean
Dim RBoard As Boolean
Dim Movement As String
Dim SelectedTile As Boolean
Dim Action As String * 1
Dim DidFlip As Boolean
Dim DidPlace As Boolean
Dim SelSide As Integer
Dim sX As Single
Dim sY As Single
Dim mX As Single
Dim mY As Single
Dim BX As Integer
Dim BY As Integer
Dim Board() As Integer
Dim XTime As Double
Function BTile(dx As Integer, dy As Integer) As Direct3DRMMeshBuilder3
Dim f As Direct3DRMFace2
Set BTile = dxBoard.mDrm.CreateMeshBuilder
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 1 + dx, 1 + dy, 1
f.AddVertex 1 + dx, 9 + dy, 1
f.AddVertex 9 + dx, 9 + dy, 1
f.AddVertex 9 + dx, 1 + dy, 1
f.SetTexture dxBoard.mDrm.LoadTexture("B1.bmp")
f.SetTextureCoordinates 0, 0, 0
f.SetTextureCoordinates 1, 1, 0
f.SetTextureCoordinates 2, 1, 1
f.SetTextureCoordinates 3, 0, 1

BTile.AddFace f
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 0 + dx, 0 + dy, 10
f.AddVertex 10 + dx, 0 + dy, 10
f.AddVertex 10 + dx, 10 + dy, 10
f.AddVertex 0 + dx, 10 + dy, 10
f.SetColorRGB 0, 0, 1
BTile.AddFace f

Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 9 + dx, 1 + dy, 1
f.AddVertex 10 + dx, 0 + dy, 0
f.AddVertex 0 + dx, 0 + dy, 0
f.AddVertex 1 + dx, 1 + dy, 1
f.SetColorRGB 0, 0, 0.5
BTile.AddFace f
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 1 + dx, 9 + dy, 1
f.AddVertex 0 + dx, 10 + dy, 0
f.AddVertex 10 + dx, 10 + dy, 0
f.AddVertex 9 + dx, 9 + dy, 1
f.SetColorRGB 0, 0.5, 0.5
BTile.AddFace f
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 1 + dx, 1 + dy, 1
f.AddVertex 0 + dx, 0 + dy, 0
f.AddVertex 0 + dx, 10 + dy, 0
f.AddVertex 1 + dx, 9 + dy, 1
f.SetColorRGB 0, 0, 0.5
BTile.AddFace f
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 9 + dx, 9 + dy, 1
f.AddVertex 10 + dx, 10 + dy, 0
f.AddVertex 10 + dx, 0 + dy, 0
f.AddVertex 9 + dx, 1 + dy, 1
f.SetColorRGB 0, 0.5, 0.5
BTile.AddFace f

Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 0 + dx, 0 + dy, 0
f.AddVertex 10 + dx, 0 + dy, 0
f.AddVertex 10 + dx, 0 + dy, 10
f.AddVertex 0 + dx, 0 + dy, 10
f.SetTexture dxBoard.mDrm.LoadTexture("B2.bmp")
f.SetTextureCoordinates 0, 0, 0
f.SetTextureCoordinates 1, 1, 0
f.SetTextureCoordinates 2, 1, 1
f.SetTextureCoordinates 3, 0, 1
BTile.AddFace f
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 10 + dx, 10 + dy, 0
f.AddVertex 0 + dx, 10 + dy, 0
f.AddVertex 0 + dx, 10 + dy, 10
f.AddVertex 10 + dx, 10 + dy, 10
f.SetTexture dxBoard.mDrm.LoadTexture("B2.bmp")
f.SetTextureCoordinates 0, 0, 0
f.SetTextureCoordinates 1, 1, 0
f.SetTextureCoordinates 2, 1, 1
f.SetTextureCoordinates 3, 0, 1
BTile.AddFace f
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 0 + dx, 10 + dy, 0
f.AddVertex 0 + dx, 0 + dy, 0
f.AddVertex 0 + dx, 0 + dy, 10
f.AddVertex 0 + dx, 10 + dy, 10
f.SetTexture dxBoard.mDrm.LoadTexture("B2.bmp")
f.SetTextureCoordinates 0, 0, 0
f.SetTextureCoordinates 1, 1, 0
f.SetTextureCoordinates 2, 1, 1
f.SetTextureCoordinates 3, 0, 1
BTile.AddFace f
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 10 + dx, 0 + dy, 0
f.AddVertex 10 + dx, 10 + dy, 0
f.AddVertex 10 + dx, 10 + dy, 10
f.AddVertex 10 + dx, 0 + dy, 10
f.SetTexture dxBoard.mDrm.LoadTexture("B2.bmp")
f.SetTextureCoordinates 0, 0, 0
f.SetTextureCoordinates 1, 1, 0
f.SetTextureCoordinates 2, 1, 1
f.SetTextureCoordinates 3, 0, 1
BTile.AddFace f
Set f = Nothing
BTile.SetName "B" & dx / 10 & dy / 10
BTile.SetQuality D3DRMRENDER_PHONG
End Function

Sub CheckForFlip(MyFrame As Direct3DRMFrame3, mX As Integer, mY As Integer)
On Local Error Resume Next
Dim X As Integer
Dim Y As Integer
Dim NewPT() As IntPt
Dim i As Integer
Dim MyTile As Direct3DRMMeshBuilder3
Set MyTile = MyFrame.GetVisual(0)
X = Mid(MyTile.GetName, 2, 1)
Y = 3 - Mid(MyTile.GetName, 3, 1)
If CanFlip(Board, X, Y, NewPT()) Then
    For i = 0 To UBound(NewPT)
        If mX = NewPT(i).X And mY = NewPT(i).Y Then
            If mX = X + 1 And mY = Y Then DoFlip MyFrame, "RIGHT"
            If mX = X - 1 And mY = Y Then DoFlip MyFrame, "LEFT"
            If mX = X And mY = Y + 1 Then DoFlip MyFrame, "DOWN"
            If mX = X And mY = Y - 1 Then DoFlip MyFrame, "UP"
            Exit Sub
        End If
    Next i
End If
End Sub

Sub CreateBoard()
Dim TMP As Direct3DRMFrame3
Dim MyTile As Direct3DRMMeshBuilder3
Dim X As Integer
Dim Y As Integer
Set TMP = dxBoard.mDrm.CreateFrame(Nothing)
For X = 0 To 3
    For Y = 0 To 3
        TMP.AddVisual BTile(X * 10, Y * 10)
    Next Y
Next X
TMP.SetPosition dxBoard.mFrO, -20, -20, 0
dxBoard.mFrO.AddChild TMP
'Side 1
Set TMP = dxBoard.mDrm.CreateFrame(Nothing)
Set MyTile = Tile(1, 0, 0)
MyTile.SetName "T1"
MyTile.Translate -5, -5, 0
TMP.AddVisual MyTile
TMP.SetPosition Nothing, 35, 20, 0
dxBoard.mFrs.AddChild TMP
'Side 3
Set TMP = dxBoard.mDrm.CreateFrame(Nothing)
Set MyTile = Tile(2, 0, 0)
MyTile.SetName "T2"
MyTile.Translate -5, -5, 0
TMP.AddVisual MyTile
TMP.SetPosition Nothing, 35, 10, 0
dxBoard.mFrs.AddChild TMP
'Side 2
Set TMP = dxBoard.mDrm.CreateFrame(Nothing)
Set MyTile = Tile(3, 0, 0)
MyTile.SetName "T3"
MyTile.Translate -5, -5, 0
TMP.AddVisual MyTile
TMP.SetPosition Nothing, 45, 20, 0
dxBoard.mFrs.AddChild TMP
'Side 4
Set TMP = dxBoard.mDrm.CreateFrame(Nothing)
Set MyTile = Tile(4, 0, 0)
MyTile.SetName "T4"
MyTile.Translate -5, -5, 0
TMP.AddVisual MyTile
TMP.SetPosition Nothing, 45, 10, 0
dxBoard.mFrs.AddChild TMP
End Sub

Sub DoComputerMove(Player As Integer, ByVal Level As Integer)
Dim WIN As Boolean
Dim NewMove As MoveSet
Dim MyFrame As Direct3DRMFrame3
Dim tBoard() As Integer
WIN = CalculateMove(Board(), Player, Level, NewMove, tBoard())
Select Case NewMove.FirstMove
    Case "P"
        SetTile Tile(NewMove.Pt, NewMove.P.X, 3 - NewMove.P.Y)
        Set MyFrame = SelTileFrame(OtherPlayer(Player), NewMove.F1.X, NewMove.F1.Y)
        If NewMove.F2.X = NewMove.F1.X + 1 And NewMove.F2.Y = NewMove.F1.Y Then DoFlip MyFrame, "RIGHT"
        If NewMove.F2.X = NewMove.F1.X - 1 And NewMove.F2.Y = NewMove.F1.Y Then DoFlip MyFrame, "LEFT"
        If NewMove.F2.X = NewMove.F1.X And NewMove.F2.Y = NewMove.F1.Y + 1 Then DoFlip MyFrame, "DOWN"
        If NewMove.F2.X = NewMove.F1.X And NewMove.F2.Y = NewMove.F1.Y - 1 Then DoFlip MyFrame, "UP"
    Case "F"
        Set MyFrame = SelTileFrame(OtherPlayer(Player), NewMove.F1.X, NewMove.F1.Y)
        If NewMove.F2.X = NewMove.F1.X + 1 And NewMove.F2.Y = NewMove.F1.Y Then DoFlip MyFrame, "RIGHT"
        If NewMove.F2.X = NewMove.F1.X - 1 And NewMove.F2.Y = NewMove.F1.Y Then DoFlip MyFrame, "LEFT"
        If NewMove.F2.X = NewMove.F1.X And NewMove.F2.Y = NewMove.F1.Y + 1 Then DoFlip MyFrame, "DOWN"
        If NewMove.F2.X = NewMove.F1.X And NewMove.F2.Y = NewMove.F1.Y - 1 Then DoFlip MyFrame, "UP"
        SetTile Tile(NewMove.Pt, NewMove.P.X, 3 - NewMove.P.Y)
    Case "O"
        SetTile Tile(NewMove.Pt, NewMove.P.X, 3 - NewMove.P.Y)
End Select
Board() = tBoard()
ShowRequiredFlips Player
dxBoard.Update
DidFlip = True
DidPlace = True
End Sub

Sub DoFlip(MyFrame As Direct3DRMFrame3, Direction As String)
Dim tX As Long 'this is used to make sure we dont get stuck in a feedback loop
tX = Timer()
Select Case Direction
    Case "UP"
        StartFlipUp MyFrame
        If Rotating Then
            picBoard.Enabled = False
            FlipTile180 MyFrame, 1, 0, 0
            picBoard.Enabled = True
            StopFlipUp MyFrame
        End If
    Case "DOWN"
        StartFlipDown MyFrame
        If Rotating Then
            picBoard.Enabled = False
            FlipTile180 MyFrame, -1, 0, 0
            picBoard.Enabled = True
            StopFlipDown MyFrame
        End If
    Case "LEFT"
        StartFlipLeft MyFrame
        If Rotating Then
            picBoard.Enabled = False
            FlipTile180 MyFrame, 0, 1, 0
            picBoard.Enabled = True
            StopFlipLeft MyFrame
        End If
    Case "RIGHT"
        StartFlipRight MyFrame
        If Rotating Then
            picBoard.Enabled = False
            FlipTile180 MyFrame, 0, -1, 0
            picBoard.Enabled = True
            StopFlipRight MyFrame
        End If
End Select
DidFlip = True
SelectedTile = False
ResetPlayedTiles
XTime = Timer
End Sub

Sub FlipTile180(MyFrame As Direct3DRMFrame3, X As Single, Y As Single, z)
Dim axis As D3DVECTOR
Dim a As D3DVECTOR
Dim B As Single
Dim i As Integer
'MyFrame.SetRotation dxBoard.mFrO, X, Y, z, 3 * (pi / 180)
MyFrame.GetRotation dxBoard.mFrO, a, B
If X <> 0 Then
    a.X = X
    If a.Y = 0 Then a.X = -1
    a.Y = 0
ElseIf Y <> 0 Then
    a.Y = Y
    If a.X = 0 Then a.Y = 1
    a.X = 0
End If
For i = 1 To 45
    MyFrame.AddRotation D3DRMCOMBINE_BEFORE, a.X, a.Y, a.z, 4 * (pi / 180)
    dxBoard.Update
Next i
'For i = 1 To 55
'    dxBoard.Update
'Next i
End Sub

Sub Forceboard(T As String)
Dim S As String
Dim i As Integer
Dim X As Integer
Dim Y As Integer
For i = 1 To Len(T)
    S = Mid(T, i, 1)
    If S <> 0 Then
        SetTile Tile(CInt(S), X, 3 - Y)
    End If
    X = X + 1
    If X > 3 Then
        X = 0
        Y = Y + 1
    End If
Next i
CurPlayer = 1
DidFlip = True
DidPlace = True
End Sub

Sub HideSelTiles(Player As Integer)
Dim HFrame As Direct3DRMFrame3
Dim i As Integer
Dim VName As String
For i = 3 To 6 'dxBoard.mFrs.GetChildren.GetSize - 1
    Set HFrame = dxBoard.mFrs.GetChildren.GetElement(i)
    VName = HFrame.GetVisual(0).GetName
    If VName = "T" & Player Or VName = "T" & Player + 2 Then
        HFrame.SetPosition Nothing, 100, 100, 100
    Else
        Select Case VName
            Case "T1": HFrame.SetPosition Nothing, 35, 20, 0
            Case "T2": HFrame.SetPosition Nothing, 35, 10, 0
            Case "T3": HFrame.SetPosition Nothing, 45, 20, 0
            Case "T4": HFrame.SetPosition Nothing, 45, 10, 0
        End Select
    End If
Next i
dxBoard.Update
End Sub

Sub HiLiteBoard(X As Integer, Y As Integer, R As Single, G As Single, B As Single)
Dim i As Integer
Dim j As Integer
Dim MyFrame As Direct3DRMFrame3
Dim MyTile As Direct3DRMMeshBuilder3
Set MyFrame = dxBoard.mFrO.GetChildren.GetElement(0)
For i = 0 To MyFrame.GetVisualCount - 1
    If MyFrame.GetVisual(i).GetName = "B" & X & Y Then
        Set MyTile = MyFrame.GetVisual(i)
        MyTile.GetFace(0).SetColorRGB R, G, B
        Set MyTile = Nothing
        Set MyFrame = Nothing
        Exit Sub
    End If
Next i
End Sub

Function OkToFlip(Player As Integer, X As Integer, Y As Integer)
Dim i As Integer
Dim RES() As IntPt
If RequiredFlips(Board(), Player, RES()) Then
    For i = 0 To UBound(RES)
        If RES(i).X = X And RES(i).Y = 3 - Y Then
            OkToFlip = True
        End If
    Next i
Else
    OkToFlip = True
End If
End Function

Sub ResetBoard()
Dim i As Integer
Dim j As Integer
Dim MyFrame As Direct3DRMFrame3
Dim MyTile As Direct3DRMMeshBuilder3
Set MyFrame = dxBoard.mFrO.GetChildren.GetElement(0)
For i = 0 To MyFrame.GetVisualCount - 1
    Set MyTile = MyFrame.GetVisual(i)
    MyTile.GetFace(0).SetColorRGB 1, 1, 1
    Set MyTile = Nothing
Next i
Set MyFrame = Nothing
End Sub
Sub GreyBoard()
Dim i As Integer
Dim j As Integer
Dim MyFrame As Direct3DRMFrame3
Dim MyTile As Direct3DRMMeshBuilder3
Set MyFrame = dxBoard.mFrO.GetChildren.GetElement(0)
For i = 0 To MyFrame.GetVisualCount - 1
    Set MyTile = MyFrame.GetVisual(i)
    MyTile.GetFace(0).SetColorRGB 0.5, 0.5, 0.5
    Set MyTile = Nothing
Next i
Set MyFrame = Nothing
End Sub
Sub ResetPlayedTiles()
Dim i As Integer
Dim MyFrame As Direct3DRMFrame3
Dim MyTile As Direct3DRMMeshBuilder3
For i = 0 To dxBoard.mFrO.GetChildren.GetSize - 1
    Set MyFrame = dxBoard.mFrO.GetChildren.GetElement(i)
    If IsNumeric(Left(MyFrame.GetVisual(0).GetName, 1)) Then
        Set MyTile = MyFrame.GetVisual(0)
        MyTile.GetFace(0).SetColorRGB 1, 1, 1
        MyTile.GetFace(1).SetColorRGB 1, 1, 1
        Set MyTile = Nothing
    End If
Next i
Set MyFrame = Nothing
End Sub

Sub ResetSelTiles()
On Local Error GoTo eTrap
SFrame.SetRotation Nothing, 0, 0, 0, 0
SFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 0, 0, 0
Select Case SelSide
    Case 1: SFrame.SetPosition Nothing, 35, 20, 0
    Case 2: SFrame.SetPosition Nothing, 35, 10, 0
    Case 3: SFrame.SetPosition Nothing, 45, 20, 0
    Case 4: SFrame.SetPosition Nothing, 45, 10, 0
End Select
SelSide = 0
Set SFrame = Nothing
eTrap:
End Sub

Sub RotateBoard(X As Single, Y As Single)
On Local Error Resume Next
Dim dx As Single
Dim dy As Single
Dim dZ As Single
Dim Distance As Single
Dim Theta As Single
PointToMouse dx, dy, Distance, X, Y
Theta = Distance / 1200
If Abs(X - sX) > Abs(Y - sY) Then
    If Y < picBoard.ScaleHeight / 2 Then
        dZ = dx
    Else
        dZ = -dx
    End If
Else
    If X > picBoard.ScaleWidth / 2 Then
        dZ = dy
    Else
        dZ = -dy
    End If
End If
dxBoard.mFrO.SetRotation Nothing, 0, 0, dZ, Theta
    
End Sub
Function PickTile(MyFrame As Direct3DRMFrame3, dx As Integer, dy As Integer) As Integer
On Local Error GoTo eTrap
Dim MyTile As Direct3DRMMeshBuilder3
Dim S As Integer
Dim X As Integer
Dim Y As Integer
Set MyTile = MyFrame.GetVisual(0)
S = Mid(MyTile.GetName, 1, 1)
X = Mid(MyTile.GetName, 2, 1)
Y = Mid(MyTile.GetName, 3, 1)
If X + dx < 0 Then PickTile = -1: Exit Function
If X + dx > 3 Then PickTile = -1: Exit Function
If 3 - Y + dy < 0 Then PickTile = -1: Exit Function
If 3 - Y + dy > 3 Then PickTile = -1: Exit Function
PickTile = Board(X + dx, 3 - Y + dy)
eTrap:
End Function

Public Sub PointToMouse(ByRef dx As Single, ByRef dy As Single, ByRef Distance As Single, X As Single, Y As Single)
Dim RX As Single
Dim RY As Single
dx = sX - X
dy = sY - Y
RX = (dx * dx)
RY = (dy * dy)
Distance = Sqr(RX + RY)
End Sub
Function SelTileFrame(Player As Integer, X As Integer, Y As Integer) As Direct3DRMFrame3
Dim i As Integer
Dim j As Integer
Dim VName As String
For i = 0 To dxBoard.mFrO.GetChildren.GetSize - 1
    For j = 0 To dxBoard.mFrO.GetChildren.GetElement(i).GetVisualCount - 1
        VName = dxBoard.mFrO.GetChildren.GetElement(i).GetVisual(j).GetName
        If VName = Player & X & 3 - Y Or VName = Player + 2 & X & 3 - Y Then
            Set SelTileFrame = dxBoard.mFrO.GetChildren.GetElement(i)
            Exit Function
        End If
    Next j
Next i
End Function

Sub SetTile(MyTile As Direct3DRMMeshBuilder3)
Dim S As Integer
Dim X As Integer
Dim Y As Integer
S = Mid(MyTile.GetName, 1, 1)
X = Mid(MyTile.GetName, 2, 1)
Y = Mid(MyTile.GetName, 3, 1)
If Board(X, 3 - Y) <> 0 Then Exit Sub
Set TFrame = dxBoard.mDrm.CreateFrame(dxBoard.mFrO)
Board(X, 3 - Y) = S
TFrame.AddVisual MyTile
TFrame.SetPosition dxBoard.mFrO, (X * 10) - 20, (Y * 10) - 20, dZ
TFrame.SetRotation dxBoard.mFrO, 0, 0, 0, 0
Action = ""
XTime = Timer()
If Not DidFlip Then DidFlip = Not IsFlippable(Board(), OtherPlayer(CurPlayer))
dxBoard.Update
End Sub

Sub ShowFlips(MyFrame As Direct3DRMFrame3)
Dim X As Integer
Dim Y As Integer
Dim NewPT() As IntPt
Dim i As Integer
Dim MyTile As Direct3DRMMeshBuilder3
GreyBoard
Set MyTile = MyFrame.GetVisual(0)
X = Mid(MyTile.GetName, 2, 1)
Y = 3 - Mid(MyTile.GetName, 3, 1)
If CanFlip(Board(), X, Y, NewPT()) Then
    For i = 0 To UBound(NewPT)
        HiLiteBoard NewPT(i).X, 3 - NewPT(i).Y, 1, 1, 1
    Next i
End If
End Sub

Sub ShowO(MyFrame As Direct3DRMFrame3)
Dim axis As D3DVECTOR
Dim a As D3DVECTOR
MyFrame.GetOrientation dxBoard.mFrO, a, axis
Me.Caption = Format(a.X, "0.0") & "," & Format(a.Y, "0.0") & "," & Format(a.z, "0.0")

End Sub

Sub ShowRequiredFlips(Player As Integer)
On Local Error Resume Next
Dim i As Integer
Dim j As Integer
Dim MyFrame As Direct3DRMFrame3
Dim MyTile As Direct3DRMMeshBuilder3
Dim RES() As IntPt
Dim DoGrey As Boolean
Dim TMP As String
ResetPlayedTiles

If Not RequiredFlips(Board(), Player, RES()) Then Exit Sub
For i = 0 To dxBoard.mFrO.GetChildren.GetSize - 1
    Set MyFrame = dxBoard.mFrO.GetChildren.GetElement(i)
    Set MyTile = MyFrame.GetVisual(0)
    TMP = MyFrame.GetVisual(0).GetName
    If IsNumeric(Left(TMP, 1)) Then
        DoGrey = True
        For j = 0 To UBound(RES)
            If TMP = Player & RES(j).X & 3 - RES(j).Y Or TMP = Player + 2 & RES(j).X & 3 - RES(j).Y Then
                DoGrey = False
                Exit For
            End If
        Next j
        If DoGrey Then
            MyTile.GetFace(0).SetColorRGB 0.75, 0.75, 0.75
            MyTile.GetFace(1).SetColorRGB 0.75, 0.75, 0.75
        Else
            MyTile.GetFace(0).SetColorRGB 1, 1, 1
            MyTile.GetFace(1).SetColorRGB 1, 1, 1
        End If
        Set MyTile = Nothing
    End If
Next i
Set MyFrame = Nothing
dxBoard.Update
End Sub
Sub StartNewGame()
Dim i As Integer
Dim TMP As Direct3DRMFrame3
Erase Board()
ReDim Board(3, 3)
For i = dxBoard.mFrO.GetChildren.GetSize - 1 To 1 Step -1
    Set TMP = dxBoard.mFrO.GetChildren.GetElement(i)
    If IsNumeric(Left(TMP.GetVisual(0).GetName, 1)) Then
        dxBoard.mFrO.DeleteChild TMP
    End If
Next i
Set TMP = Nothing
dxBoard.mFrO.SetRotation Nothing, 0, 0, 0, 0
CompLevel = CInt(lblLevel.Caption)
CurPlayer = FirstPlayer
HideSelTiles OtherPlayer(CurPlayer)
CompNum = 2
DidPlace = False
DidFlip = True
If lblPlayers.Caption = "DEMO" Then
    totPlayers = 0
    DoComputerMove CurPlayer, CompLevel
Else
    totPlayers = CInt(lblPlayers.Caption)
End If
Timer1.Interval = 100
End Sub

Function Tile(Side As Integer, X As Integer, Y As Integer) As Direct3DRMMeshBuilder3
Dim f As Direct3DRMFace2
Set Tile = dxBoard.mDrm.CreateMeshBuilder
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 1, 1, 1
f.AddVertex 9, 1, 1
f.AddVertex 9, 9, 1
f.AddVertex 1, 9, 1
Select Case Side
    Case 1
        f.SetTexture dxBoard.mDrm.LoadTexture("1.bmp")
    Case 2
        f.SetTexture dxBoard.mDrm.LoadTexture("2.bmp")
    Case 3
        f.SetTexture dxBoard.mDrm.LoadTexture("3.bmp")
    Case 4
        f.SetTexture dxBoard.mDrm.LoadTexture("4.bmp")
End Select
f.SetTextureCoordinates 0, 0, 1
f.SetTextureCoordinates 1, 1, 1
f.SetTextureCoordinates 2, 1, 0
f.SetTextureCoordinates 3, 0, 0
Tile.AddFace f

Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 1, 1, -1
f.AddVertex 1, 9, -1
f.AddVertex 9, 9, -1
f.AddVertex 9, 1, -1
Select Case Side
    Case 3
        f.SetTexture dxBoard.mDrm.LoadTexture("1.bmp")
    Case 4
        f.SetTexture dxBoard.mDrm.LoadTexture("2.bmp")
    Case 1
        f.SetTexture dxBoard.mDrm.LoadTexture("3.bmp")
    Case 2
        f.SetTexture dxBoard.mDrm.LoadTexture("4.bmp")
End Select
f.SetTextureCoordinates 0, 0, 0
f.SetTextureCoordinates 1, 1, 0
f.SetTextureCoordinates 2, 1, 1
f.SetTextureCoordinates 3, 0, 1
Tile.AddFace f
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 9, 1, 1
f.AddVertex 1, 1, 1
f.AddVertex 0, 0, 0
f.AddVertex 10, 0, 0
Select Case Side
    Case 1, 3: f.SetColorRGB 0.75, 0, 0
    Case 2, 4: f.SetColorRGB 0, 0.75, 0
End Select
Tile.AddFace f
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 1, 1, -1
f.AddVertex 9, 1, -1
f.AddVertex 10, 0, 0
f.AddVertex 0, 0, 0
Select Case Side
    Case 1, 3: f.SetColorRGB 0.75, 0, 0
    Case 2, 4: f.SetColorRGB 0, 0.75, 0
End Select
Tile.AddFace f
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 1, 9, 1
f.AddVertex 9, 9, 1
f.AddVertex 10, 10, 0
f.AddVertex 0, 10, 0
Select Case Side
    Case 1, 3: f.SetColorRGB 0.75, 0, 0
    Case 2, 4: f.SetColorRGB 0, 0.75, 0
End Select
Tile.AddFace f
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 9, 9, -1
f.AddVertex 1, 9, -1
f.AddVertex 0, 10, 0
f.AddVertex 10, 10, 0
Select Case Side
    Case 1, 3: f.SetColorRGB 0.75, 0, 0
    Case 2, 4: f.SetColorRGB 0, 0.75, 0
End Select
Tile.AddFace f
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 1, 1, 1
f.AddVertex 1, 9, 1
f.AddVertex 0, 10, 0
f.AddVertex 0, 0, 0
Select Case Side
    Case 1, 3: f.SetColorRGB 0.75, 0, 0
    Case 2, 4: f.SetColorRGB 0, 0.75, 0
End Select
Tile.AddFace f
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 1, 9, -1
f.AddVertex 1, 1, -1
f.AddVertex 0, 0, 0
f.AddVertex 0, 10, 0
Select Case Side
    Case 1, 3: f.SetColorRGB 0.75, 0, 0
    Case 2, 4: f.SetColorRGB 0, 0.75, 0
End Select
Tile.AddFace f
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 9, 9, 1
f.AddVertex 9, 1, 1
f.AddVertex 10, 0, 0
f.AddVertex 10, 10, 0
Select Case Side
    Case 1, 3: f.SetColorRGB 0.75, 0, 0
    Case 2, 4: f.SetColorRGB 0, 0.75, 0
End Select
Tile.AddFace f
Set f = dxBoard.mDrm.CreateFace()
f.AddVertex 9, 1, -1
f.AddVertex 9, 9, -1
f.AddVertex 10, 10, 0
f.AddVertex 10, 0, 0
Select Case Side
    Case 1, 3: f.SetColorRGB 0.75, 0, 0
    Case 2, 4: f.SetColorRGB 0, 0.75, 0
End Select
Tile.AddFace f
Set f = Nothing
Tile.SetName Side & X & Y
Tile.SetQuality D3DRMRENDER_PHONG

End Function

Sub StartFlipLeft(MyFrame As Direct3DRMFrame3)
If PickTile(MyFrame, -1, 0) <> 0 Then Exit Sub
Movement = "LEFT"
Rotating = True
End Sub
Sub StartFlipDown(MyFrame As Direct3DRMFrame3)
If PickTile(MyFrame, 0, 1) <> 0 Then Exit Sub
Rotating = True
Movement = "DOWN"
End Sub



Sub StartFlipUp(MyFrame As Direct3DRMFrame3)
If PickTile(MyFrame, 0, -1) <> 0 Then Exit Sub
Dim a As D3DVECTOR
MyFrame.GetPosition dxBoard.mFrO, a
MyFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 0, 1, 180 * (pi / 180)
MyFrame.SetPosition dxBoard.mFrO, a.X + 10, a.Y + 10, a.z
Rotating = True
Movement = "UP"
End Sub
Sub StartFlipRight(MyFrame As Direct3DRMFrame3)
If PickTile(MyFrame, 1, 0) <> 0 Then Exit Sub
Dim a As D3DVECTOR
MyFrame.GetPosition dxBoard.mFrO, a
MyFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 0, 1, 180 * (pi / 180)
MyFrame.SetPosition dxBoard.mFrO, a.X + 10, a.Y + 10, a.z
Rotating = True
Movement = "RIGHT"
End Sub
Sub StopFlipDown(MyFrame As Direct3DRMFrame3)
Dim axis As D3DVECTOR
Dim a As D3DVECTOR
Dim S As Integer
Dim X As Integer
Dim Y As Integer
Dim MyTile As Direct3DRMMeshBuilder3
Set MyTile = MyFrame.GetVisual(0)
S = Mid(MyTile.GetName, 1, 1)
X = Mid(MyTile.GetName, 2, 1)
Y = Mid(MyTile.GetName, 3, 1)
MyFrame.SetRotation dxBoard.mFrO, 0, 0, 0, 0
MyFrame.GetOrientation dxBoard.mFrO, a, axis
If a.z <= 0 Then
    MyFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 0, 0, 0
    MyFrame.DeleteVisual MyFrame.GetVisual(0)
    MyFrame.AddVisual Tile(Rev(S), X, Y - 1)
    MyFrame.SetPosition dxBoard.mFrO, (X * 10) - 20, ((Y * 10) - 10) - 20, dZ
    Board(X, 3 - Y) = 0
    Board(X, 3 - Y + 1) = Rev(S)
Else
    MyFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 0, 0, 0
    MyFrame.SetPosition dxBoard.mFrO, (X * 10) - 20, (Y * 10) - 20, dZ
End If

Rotating = False

End Sub
Sub StopFlipLeft(MyFrame As Direct3DRMFrame3)
Dim axis As D3DVECTOR
Dim a As D3DVECTOR
Dim S As Integer
Dim X As Integer
Dim Y As Integer
Dim MyTile As Direct3DRMMeshBuilder3
Set MyTile = MyFrame.GetVisual(0)
S = Mid(MyTile.GetName, 1, 1)
X = Mid(MyTile.GetName, 2, 1)
Y = Mid(MyTile.GetName, 3, 1)
MyFrame.SetRotation dxBoard.mFrO, 0, 0, 0, 0
MyFrame.GetOrientation dxBoard.mFrO, a, axis
If a.z <= 0 Then
    MyFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 0, 0, 0
    MyFrame.DeleteVisual MyFrame.GetVisual(0)
    MyFrame.AddVisual Tile(Rev(S), X - 1, Y)
    MyFrame.SetPosition dxBoard.mFrO, ((X * 10) - 10) - 20, (Y * 10) - 20, dZ
    Board(X, 3 - Y) = 0
    Board(X - 1, 3 - Y) = Rev(S)
Else
    MyFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 0, 0, 0
    MyFrame.SetPosition dxBoard.mFrO, (X * 10) - 20, (Y * 10) - 20, dZ
End If

Rotating = False

End Sub
Sub StopFlipRight(MyFrame As Direct3DRMFrame3)
Dim axis As D3DVECTOR
Dim a As D3DVECTOR
Dim S As Integer
Dim X As Integer
Dim Y As Integer
Dim MyTile As Direct3DRMMeshBuilder3
Set MyTile = MyFrame.GetVisual(0)
S = Mid(MyTile.GetName, 1, 1)
X = Mid(MyTile.GetName, 2, 1)
Y = Mid(MyTile.GetName, 3, 1)
MyFrame.SetRotation dxBoard.mFrO, 0, 0, 0, 0
MyFrame.GetOrientation dxBoard.mFrO, a, axis
If a.z <= 0 Then
    MyFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 0, 0, 0
    MyFrame.DeleteVisual MyFrame.GetVisual(0)
    MyFrame.AddVisual Tile(Rev(S), X + 1, Y)
    MyFrame.SetPosition dxBoard.mFrO, ((X * 10) + 10) - 20, (Y * 10) - 20, dZ
    Board(X, 3 - Y) = 0
    Board(X + 1, 3 - Y) = Rev(S)
Else
    MyFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 0, 0, 0
    MyFrame.SetPosition dxBoard.mFrO, (X * 10) - 20, (Y * 10) - 20, dZ
End If

Rotating = False

End Sub

Sub StopFlipUp(MyFrame As Direct3DRMFrame3)
Dim axis As D3DVECTOR
Dim a As D3DVECTOR
Dim S As Integer
Dim X As Integer
Dim Y As Integer
Dim MyTile As Direct3DRMMeshBuilder3
Set MyTile = MyFrame.GetVisual(0)
S = Mid(MyTile.GetName, 1, 1)
X = Mid(MyTile.GetName, 2, 1)
Y = Mid(MyTile.GetName, 3, 1)
MyFrame.SetRotation dxBoard.mFrO, 0, 0, 0, 0
MyFrame.GetOrientation dxBoard.mFrO, a, axis
If a.z <= 0 Then
    MyFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 0, 0, 0
    MyFrame.DeleteVisual MyFrame.GetVisual(0)
    MyFrame.AddVisual Tile(Rev(S), X, Y + 1)
    MyFrame.SetPosition dxBoard.mFrO, (X * 10) - 20, ((Y * 10) + 10) - 20, dZ
    Board(X, 3 - Y) = 0
    Board(X, 3 - Y - 1) = Rev(S)
Else
    MyFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 0, 0, 0
    MyFrame.SetPosition dxBoard.mFrO, (X * 10) - 20, (Y * 10) - 20, dZ
End If

Rotating = False

End Sub






Private Sub cmdLevelDown_Click()
Select Case lblLevel.Caption
    Case "5": lblLevel.Caption = "4"
    Case "4": lblLevel.Caption = "3"
    Case "3": lblLevel.Caption = "2"
    Case "2": lblLevel.Caption = "1"
    Case "1": lblLevel.Caption = "0"
End Select
End Sub

Private Sub cmdLevelUp_Click()
Select Case lblLevel.Caption
    Case "4": lblLevel.Caption = "5"
    Case "3": lblLevel.Caption = "4"
    Case "2": lblLevel.Caption = "3"
    Case "1": lblLevel.Caption = "2"
    Case "0": lblLevel.Caption = "1"
End Select
End Sub

Private Sub cmdPlayDown_Click()
Select Case lblPlayers.Caption
    Case "2"
        lblPlayers.Caption = "1"
        lblFirst.Visible = True
        fraLevel.Visible = True
    Case "1"
        lblPlayers.Caption = "DEMO"
        lblFirst.Visible = False
        fraLevel.Visible = True
End Select
End Sub

Private Sub cmdPlayUp_Click()
Select Case lblPlayers.Caption
    Case "1"
        lblPlayers.Caption = "2"
        lblFirst.Visible = False
        fraLevel.Visible = False
    Case "DEMO"
        lblPlayers.Caption = "1"
        lblFirst.Visible = True
        fraLevel.Visible = True
End Select
End Sub

Private Sub cmdShowOpt_Click()
If picOptions.Visible = True Then
    picOptions.Visible = False
Else
    picOptions.Visible = True
End If
End Sub

Private Sub Form_Load()
ReDim Preserve Board(3, 3)
dxBoard.InitDx picBoard
dxBoard.mFrC.SetPosition Nothing, 10, -5, -80
dxBoard.mFrC.AddRotation D3DRMCOMBINE_AFTER, -1, 0, 0, 50 * (pi / 180)

CreateBoard
dxBoard.Update
Set AnxLBL = lblANX
AnxLBL = "Strategy Square"
Timer1.Interval = 100
RotateBoard 0, 50
FirstPlayer = 1
totPlayers = -1
End Sub
Private Sub Form_Resize()
On Local Error Resume Next
picBoard.Width = Me.ScaleWidth
picOptions.Top = Me.ScaleHeight - 160
picOptions.Left = Me.ScaleWidth - 115
cmdShowOpt.Top = Me.ScaleHeight - 17
cmdShowOpt.Left = Me.ScaleWidth - 17

End Sub


Private Sub lblFirst_DblClick()
If FirstPlayer = 1 Then
    FirstPlayer = 2
    lblFirst = "Computer Goes First"
Else
    FirstPlayer = 1
    lblFirst = "Player Goes First"
End If
End Sub


Private Sub lblStart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStart.ForeColor = RGB(200, 200, 200)

End Sub


Private Sub lblStart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStart.ForeColor = vbWhite
StartNewGame
picOptions.Visible = False
End Sub


Private Sub picANX_Resize()
lblANX.Width = picANX.ScaleWidth
chkFlip.Left = Me.ScaleWidth - 55
chkPlace.Left = Me.ScaleWidth - 55
End Sub


Private Sub picBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PName As String
sX = X
sY = Y
RBoard = False
PName = dxBoard.PickName(X, Y)
Select Case Left(PName, 1)
    Case "B"
        If totPlayers = 1 And CurPlayer = CompNum Then Exit Sub
        ResetBoard
        BX = Mid(PName, 2, 1)
        BY = Mid(PName, 3, 1)
        If Action = "F" And Not DidFlip Then
            If SelectedTile Then CheckForFlip TFrame, BX, 3 - BY
        ElseIf Action = "S" And Not DidPlace Then
            SetTile Tile(SelSide, BX, BY)
            DidPlace = True
            ResetSelTiles
        End If
        
    Case "1", "2", "3", "4"
        If totPlayers = 1 And CurPlayer = CompNum Then Exit Sub
        ResetBoard
        If OtherPlayer(CurPlayer) = Left(PName, 1) Or OtherPlayer(CurPlayer) + 2 = Left(PName, 1) Or totPlayers = -1 Then
            If OkToFlip(OtherPlayer(CurPlayer), Mid(PName, 2, 1), Mid(PName, 3, 1)) Then
                Set TFrame = dxBoard.PickFrame(X, Y)
                SelectedTile = True
                ShowFlips TFrame
                Action = "F"
            End If
        End If
    Case "T"
        If totPlayers = 1 And CurPlayer = CompNum Then Exit Sub
        ResetSelTiles
        SelSide = Right(PName, 1)
        If SelSide = CurPlayer Or SelSide = CurPlayer + 2 Then
            Action = "S"
            Set SFrame = dxBoard.mFrs.GetChildren.GetElement(2 + SelSide)
            SFrame.SetRotation Nothing, 0, 0, 1, 30 * (pi / 180)
        End If
    Case Else
        SelectedTile = False
        RBoard = True
End Select
End Sub
Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If RBoard Then
    RotateBoard X, Y
End If

End Sub


Private Sub picBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If RBoard Then
    dxBoard.mFrO.SetRotation Nothing, 0, 0, 0, 0
End If
RBoard = False
Movement = ""
Exit Sub

Select Case Movement
    Case "DOWN"
        StopFlipDown TFrame
    Case "UP"
        StopFlipUp TFrame
    Case "LEFT"
        StopFlipLeft TFrame
    Case "RIGHT"
        StopFlipRight TFrame
End Select
End Sub

Private Sub picBoard_Resize()
On Local Error Resume Next
dxBoard.Resize picBoard
dxBoard.mVpt.SetBack 160
End Sub

Private Sub Timer1_Timer()
dxBoard.Update
If SpaceCount(Board()) = 0 And Not HasWin(Board(), 1) And Not HasWin(Board(), 2) Then
    lblANX.Caption = "DRAW!"
    Exit Sub
End If
If DidFlip Then chkFlip.Value = vbChecked Else chkFlip.Value = vbUnchecked
If DidPlace Then chkPlace.Value = vbChecked Else chkPlace.Value = vbUnchecked
Select Case totPlayers
    Case -1
        If DidFlip Or DidPlace Then
            DidFlip = False
            DidPlace = False
        End If
    Case 0
        If DidFlip And DidPlace Then
            DidFlip = False
            DidPlace = False
            If HasWin(Board(), CurPlayer) Then
                lblANX.Caption = "Player " & CurPlayer & " Wins"
                totPlayers = -1
                picOptions.Visible = True
                Timer1.Interval = 0
                Exit Sub
            Else
                HideSelTiles CurPlayer
            End If
            CurPlayer = OtherPlayer(CurPlayer)
            lblANX.Caption = "Player " & CurPlayer & " Turn"
            DoEvents
            DoComputerMove CurPlayer, CompLevel
        End If
    Case 1
        lblANX.Caption = "Player " & CurPlayer & " Turn"
        If DidFlip And DidPlace Then
            DidFlip = False
            DidPlace = False
            If HasWin(Board(), CurPlayer) Then
                lblANX.Caption = "Player " & CurPlayer & " Wins"
                totPlayers = -1
                picOptions.Visible = True
                Timer1.Interval = 0
                Exit Sub
            Else
                HideSelTiles CurPlayer
            End If
            ShowRequiredFlips CurPlayer
            CurPlayer = OtherPlayer(CurPlayer)
            lblANX.Caption = "Player " & CurPlayer & " Turn"
            DoEvents
            If CurPlayer = CompNum Then
                DoComputerMove 2, CompLevel
            End If
        End If
    Case 2
        lblANX.Caption = "Player " & CurPlayer & " Turn"
        If DidFlip And DidPlace Then
            DidFlip = False
            DidPlace = False
            If HasWin(Board(), CurPlayer) Then
                lblANX.Caption = "Player " & CurPlayer & " Wins"
                totPlayers = -1
                picOptions.Visible = True
                Timer1.Interval = 0
                Exit Sub
            Else
                HideSelTiles CurPlayer
            End If
            ShowRequiredFlips CurPlayer
            CurPlayer = OtherPlayer(CurPlayer)
            lblANX.Caption = "Player " & CurPlayer & " Turn"
            DoEvents
        End If
End Select
End Sub

Sub CheckMovement(MyFrame As Direct3DRMFrame3)
Dim axis As D3DVECTOR
Dim a As D3DVECTOR
TFrame.GetOrientation dxBoard.mFrO, axis, a
Select Case Movement
    Case "DOWN"
        If axis.Y < 0 Then StopFlipDown MyFrame
    Case "UP"
        If axis.Y > 0 Then StopFlipUp MyFrame
    Case "LEFT"
        If axis.X < 0 Then StopFlipLeft MyFrame
    Case "RIGHT"
        If axis.X > 0 Then StopFlipRight MyFrame
    Case Else
        Rotating = False
End Select
End Sub
