VERSION 5.00
Begin VB.UserControl ButtonBar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10845
   ClipControls    =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   723
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
      Height          =   1215
      Left            =   0
      Picture         =   "ToolBar.ctx":0000
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   567
      TabIndex        =   0
      Top             =   2070
      Visible         =   0   'False
      Width           =   8505
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   5
      Left            =   2655
      TabIndex        =   21
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   20
      Left            =   10125
      TabIndex        =   20
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   19
      Left            =   9495
      TabIndex        =   19
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   18
      Left            =   9045
      TabIndex        =   18
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   17
      Left            =   8595
      TabIndex        =   17
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   16
      Left            =   8145
      TabIndex        =   16
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   15
      Left            =   7515
      TabIndex        =   15
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   14
      Left            =   7065
      TabIndex        =   14
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   13
      Left            =   6615
      TabIndex        =   13
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   12
      Left            =   6165
      TabIndex        =   12
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   11
      Left            =   5550
      TabIndex        =   11
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   10
      Left            =   5100
      TabIndex        =   10
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   9
      Left            =   4455
      TabIndex        =   9
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   8
      Left            =   4005
      TabIndex        =   8
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   7
      Left            =   3555
      TabIndex        =   7
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   6
      Left            =   3105
      TabIndex        =   6
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   4
      Left            =   2205
      TabIndex        =   5
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   3
      Left            =   1755
      TabIndex        =   4
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   2
      Left            =   1125
      TabIndex        =   3
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   1
      Left            =   675
      TabIndex        =   2
      Top             =   60
      Width           =   405
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   0
      Left            =   225
      TabIndex        =   1
      Top             =   60
      Width           =   405
   End
End
Attribute VB_Name = "ButtonBar"
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

Private MouseOver(20) As Boolean
Private MouseButton(20) As Integer
Private ButtonState(20) As BtState

Public Event Click(Index As Integer)
Public Property Let ToolTip(ByVal Index As Integer, ByVal vData As String)

    Label1(Index).ToolTipText = vData

End Property
Private Sub Label1_Click(Index As Integer)

    RaiseEvent Click(Index)

End Sub
Public Property Let Enabled(ByVal Index As Integer, ByVal vData As Boolean)

    Label1(Index).Enabled = vData
    If Not Label1(Index).Enabled Then
        MouseOver(Index) = False
        ButtonState(Index) = sFlat
        MouseButton(Index) = 0
        DrawButton sDisable, Index
    Else
        DrawButton sFlat, Index
    End If
    
End Property
Public Property Get Enabled(ByVal Index As Integer) As Boolean

    Enable = Label1(Index).Enabled
        
End Property
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
    
    UserControl.AutoRedraw = True
    Select Case State
        Case sUp
            PaintPicture Pic, Label1(Index).Left, Label1(Index).Top, 27, 27, Index * 27, 27, 27, 27
            UserControl.Line (Label1(Index).Left, Label1(Index).Top)-Step(26, 0), &HE0E0E0
            UserControl.Line (Label1(Index).Left, Label1(Index).Top)-Step(0, 26), &HE0E0E0
            UserControl.Line (Label1(Index).Left, Label1(Index).Top + 26)-Step(27, 0), &H404040
            UserControl.Line (Label1(Index).Left + 26, Label1(Index).Top)-Step(0, 26), &H404040
        Case sDown
            PaintPicture Pic, Label1(Index).Left + 1, Label1(Index).Top + 1, 27, 27, Index * 27, 27, 27, 27
            UserControl.Line (Label1(Index).Left, Label1(Index).Top)-Step(26, 0), &H404040
            UserControl.Line (Label1(Index).Left, Label1(Index).Top)-Step(0, 26), &H404040
            UserControl.Line (Label1(Index).Left, Label1(Index).Top + 26)-Step(27, 0), &HE0E0E0
            UserControl.Line (Label1(Index).Left + 26, Label1(Index).Top)-Step(0, 26), &HE0E0E0
        Case sFlat
            PaintPicture Pic, Label1(Index).Left, Label1(Index).Top, 27, 27, Index * 27, 0, 27, 27
        Case sDisable
            PaintPicture Pic, Label1(Index).Left, Label1(Index).Top, 27, 27, Index * 27, 54, 27, 27
    End Select
    UserControl.AutoRedraw = False
    
End Sub
Private Sub Timer1_Timer()
    
    If MouseButton(AtualIndex) = 1 Then Exit Sub
    GetCursorPos Pt
    Call ScreenToClient(UserControl.hwnd, Pt)
    If Pt.X < Label1(AtualIndex).Left Or Pt.Y < Label1(AtualIndex).Top Or Pt.X > Label1(AtualIndex).Left + 27 Or Pt.Y > Label1(AtualIndex).Top + 27 Then
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
    For i = 0 To 20
        UserControl.PaintPicture Pic, Label1(i).Left, Label1(i).Top, 27, 27, i * 27, 0, 27, 27
    Next
    PrevIndex = 0

End Sub
Private Sub UserControl_Resize()
    
    UserControl.Cls
    UserControl.Line (1, 1)-Step(UserControl.ScaleWidth - 5, 0), &HE0E0E0
    UserControl.Line (1, 34)-Step(UserControl.ScaleWidth - 5, 0), &H808080
    UserControl.Line (1, 1)-Step(0, 34), &HE0E0E0
    UserControl.Line (5, 4)-Step(0, 28), &HE0E0E0
    UserControl.Line (7, 4)-Step(0, 28), &H808080
    UserControl.Line (8, 4)-Step(0, 28), &HE0E0E0
    UserControl.Line (10, 4)-Step(0, 28), &H808080
    UserControl.Line (109, 4)-Step(0, 28), &H808080
    UserControl.Line (110, 4)-Step(0, 28), &HE0E0E0
    UserControl.Line (328, 4)-Step(0, 28), &H808080
    UserControl.Line (329, 4)-Step(0, 28), &HE0E0E0
    UserControl.Line (400, 4)-Step(0, 28), &H808080
    UserControl.Line (401, 4)-Step(0, 28), &HE0E0E0
    UserControl.Line (533, 4)-Step(0, 28), &H808080
    UserControl.Line (534, 4)-Step(0, 28), &HE0E0E0
    UserControl.Line (664, 4)-Step(0, 28), &H808080
    UserControl.Line (665, 4)-Step(0, 28), &HE0E0E0
    UserControl.Line (UserControl.ScaleWidth - 5, 1)-Step(0, 34), &H808080
    For i = 0 To 20
        UserControl.PaintPicture Pic, Label1(i).Left, Label1(i).Top, 27, 27, i * 27, 0, 27, 27
    Next

End Sub
