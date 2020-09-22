VERSION 5.00
Begin VB.UserControl ButtonScroll 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   465
   ClipControls    =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   31
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   495
      Top             =   1620
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   90
      Picture         =   "ButtonScroll.ctx":0000
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   0
      Top             =   2385
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Top             =   1425
      Width           =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   1140
      Width           =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   855
      Width           =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   570
      Width           =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   285
      Width           =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "ButtonScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Enum BtState
    sDown = 1
    sUp = 2
    sFlat = 3
    sDisable = 4
End Enum
    
Dim Pt As POINTAPI
Dim PrevIndex As Integer
Dim AtualIndex As Integer

Private MouseOver(19) As Boolean
Private MouseButton(19) As Integer
Private ButtonState(19) As BtState

Public Event Click(Index As Integer)
Public Property Let ToolTip(ByVal Index As Integer, ByVal vData As String)

    Label1(Index).ToolTipText = vData

End Property
Public Property Let Enabled(ByVal Index As Integer, ByVal vData As Boolean)

    Label1(Index).Enabled = vData
    If Not Label1(Index).Enabled Then
        MouseOver(Index) = False
        ButtonState(Index) = sFlat
        MouseButton(Index) = 0
        DrawButton sDisable, Index
        If PrevIndex = Index Then PrevIndex = PrevIndex + 1
        If AtualIndex = Index Then AtualIndex = AtualIndex + 1
    Else
        DrawButton sFlat, Index
    End If
    
End Property
Public Property Get Enabled(ByVal Index As Integer) As Boolean

    Enable = Label1(Index).Enabled
        
End Property
Private Sub Label1_Click(Index As Integer)

    RaiseEvent Click(Index)

End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MouseButton(Index) = Button
    If ButtonState(Index) <> sDown Then
        DrawButton sDown, Index
        ButtonState(Index) = sDown
    End If
    
End Sub
Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If PrevIndex <> Index Then
        DrawButton sFlat, PrevIndex
        MouseOver(PrevIndex) = False
        ButtonState(PrevIndex) = sFlat
        MouseButton(PrevIndex) = 0
        PrevIndex = Index
    End If
    AtualIndex = Index
    If Not Timer1.Enabled Then Timer1.Enabled = True
    Select Case Button
        Case 1
            If MouseOver(Index) Then
                If ButtonState(Index) <> sDown Then
                    DrawButton sDown, Index
                    ButtonState(Index) = sDown
                End If
            Else
                If ButtonState(Index) <> sUp Then
                    DrawButton sUp, Index
                    ButtonState(Index) = sUp
                End If
            End If
        Case 0
            MouseOver(Index) = True
            If ButtonState(Index) <> sUp Then
                DrawButton sUp, Index
                ButtonState(Index) = sUp
            End If
    End Select
    MouseButton(Index) = Button

End Sub
Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MouseButton(Index) = 0
    DoEvents
End Sub
Private Sub DrawButton(ByVal State As BtState, ByVal Index As Integer)
    
    On Error Resume Next
    Select Case State
        Case sUp
            PaintPicture Pic, 0, Label1(Index).Top, 16, 19, Index * 16, 20, 16, 19
            UserControl.Line (0, Label1(Index).Top)-Step(15, 0), vb3DHighlight
            UserControl.Line (0, Label1(Index).Top)-Step(0, 18), vb3DHighlight
            UserControl.Line (0, Label1(Index).Top + 18)-Step(16, 0), vb3DDKShadow
            UserControl.Line (Label1(Index).Left + 15, Label1(Index).Top)-Step(0, 18), vb3DDKShadow
        Case sDown
            PaintPicture Pic, 1, Label1(Index).Top + 1, 16, 19, Index * 16, 20, 16, 19
            UserControl.Line (0, Label1(Index).Top)-Step(15, 0), vb3DDKShadow
            UserControl.Line (0, Label1(Index).Top)-Step(0, 18), vb3DDKShadow
            UserControl.Line (0, Label1(Index).Top + 18)-Step(16, 0), vb3DHighlight
            UserControl.Line (Label1(Index).Left + 15, Label1(Index).Top)-Step(0, 18), vb3DHighlight
        Case sFlat
            PaintPicture Pic, 0, Label1(Index).Top, 16, 19, Index * 16, 0, 16, 19
        Case sDisable
            PaintPicture Pic, 0, Label1(Index).Top, 16, 19, Index * 16, 40, 16, 19
    End Select
    
End Sub
Private Sub Timer1_Timer()
    
    If MouseButton(AtualIndex) = 1 Then Exit Sub
    GetCursorPos Pt
    Call ScreenToClient(UserControl.hwnd, Pt)
    If Pt.X < 0 Or Pt.Y < Label1(AtualIndex).Top Or Pt.X > Label1(AtualIndex).Left + 16 Or Pt.Y > Label1(AtualIndex).Top + 20 Then
        MouseOver(AtualIndex) = False
        Timer1.Enabled = False
        If ButtonState(AtualIndex) <> sFlat Then
            DrawButton sFlat, AtualIndex
            ButtonState(AtualIndex) = sFlat
        End If
    Else
        MouseOver(AtualIndex) = True
    End If

End Sub
Private Sub UserControl_Initialize()

    UserControl.Cls
    For i = 0 To 5
        UserControl.PaintPicture Pic, 0, Label1(i).Top, 16, 19, i * 16, 0, 16, 19
    Next
    PrevIndex = 0

End Sub
Private Sub UserControl_Resize()
    
    UserControl.Cls
    For i = 0 To 5
        n = i * 27
        UserControl.PaintPicture Pic, 0, Label1(i).Top, 16, 19, i * 16, 0, 16, 19
    Next

End Sub
