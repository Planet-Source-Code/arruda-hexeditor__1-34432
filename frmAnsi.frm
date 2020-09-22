VERSION 5.00
Begin VB.Form frmAnsi 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ANSI Table"
   ClientHeight    =   7155
   ClientLeft      =   1380
   ClientTop       =   1080
   ClientWidth     =   8430
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   477
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   562
   ShowInTaskbar   =   0   'False
   Begin Project1.Command Command1 
      Height          =   330
      Left            =   7290
      TabIndex        =   18
      Top             =   6705
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      Caption         =   "&Close"
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "128 - 255"
      Height          =   330
      Index           =   1
      Left            =   945
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6750
      Width           =   825
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "0 - 127"
      Height          =   330
      Index           =   0
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6750
      Value           =   -1  'True
      Width           =   825
   End
   Begin VB.PictureBox Picture2 
      Height          =   6120
      Left            =   90
      ScaleHeight     =   404
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   546
      TabIndex        =   0
      Top             =   450
      Width           =   8250
      Begin VB.VScrollBar VScroll1 
         Height          =   6060
         Left            =   7920
         TabIndex        =   2
         Top             =   0
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00D8CBCB&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   11520
         Left            =   0
         ScaleHeight     =   768
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   528
         TabIndex        =   1
         Top             =   0
         Width           =   7920
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   5
            X1              =   394
            X2              =   394
            Y1              =   0
            Y2              =   768
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   393
            X2              =   393
            Y1              =   0
            Y2              =   768
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   3
            X1              =   261
            X2              =   261
            Y1              =   0
            Y2              =   768
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   2
            X1              =   262
            X2              =   262
            Y1              =   0
            Y2              =   768
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   1
            X1              =   127
            X2              =   127
            Y1              =   0
            Y2              =   768
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   126
            X2              =   126
            Y1              =   0
            Y2              =   768
         End
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00D8CBCB&
      Index           =   10
      X1              =   537
      X2              =   537
      Y1              =   9
      Y2              =   30
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00655450&
      Index           =   9
      X1              =   536
      X2              =   536
      Y1              =   9
      Y2              =   30
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00D8CBCB&
      Index           =   8
      X1              =   402
      X2              =   402
      Y1              =   9
      Y2              =   30
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00655450&
      Index           =   7
      X1              =   401
      X2              =   401
      Y1              =   9
      Y2              =   30
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00655450&
      Index           =   6
      X1              =   269
      X2              =   269
      Y1              =   9
      Y2              =   30
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00D8CBCB&
      Index           =   5
      X1              =   270
      X2              =   270
      Y1              =   9
      Y2              =   30
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00D8CBCB&
      Index           =   4
      X1              =   135
      X2              =   135
      Y1              =   9
      Y2              =   30
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00655450&
      Index           =   3
      X1              =   134
      X2              =   134
      Y1              =   9
      Y2              =   30
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00655450&
      Index           =   2
      X1              =   554
      X2              =   554
      Y1              =   9
      Y2              =   31
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00DACFCF&
      Index           =   1
      X1              =   6
      X2              =   555
      Y1              =   9
      Y2              =   9
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00DACFCF&
      Index           =   0
      X1              =   6
      X2              =   6
      Y1              =   9
      Y2              =   31
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Char"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   11
      Left            =   7440
      TabIndex        =   14
      Top             =   210
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Char"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   10
      Left            =   5415
      TabIndex        =   13
      Top             =   210
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Char"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   9
      Left            =   3450
      TabIndex        =   12
      Top             =   210
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Char"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   8
      Left            =   1440
      TabIndex        =   11
      Top             =   210
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   7
      Left            =   6780
      TabIndex        =   10
      Top             =   210
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   6
      Left            =   6180
      TabIndex        =   9
      Top             =   210
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   5
      Left            =   4185
      TabIndex        =   8
      Top             =   210
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   4
      Left            =   4785
      TabIndex        =   7
      Top             =   210
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   3
      Left            =   2190
      TabIndex        =   6
      Top             =   210
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   2
      Left            =   2805
      TabIndex        =   5
      Top             =   210
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   1
      Left            =   810
      TabIndex        =   4
      Top             =   210
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   0
      Left            =   195
      TabIndex        =   3
      Top             =   210
      Width           =   360
   End
   Begin VB.Label Label2 
      BackColor       =   &H00AB8F8D&
      Height          =   330
      Left            =   90
      TabIndex        =   17
      Top             =   135
      Width           =   8220
   End
End
Attribute VB_Name = "frmAnsi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Msgs(1) As String
Private Sub DrawTable(ByVal C As Integer)
    
    Picture1.Cls
    VScroll1 = 0
    For i = 0 To 31
        Y = (i * 24) - 4
        If i Mod 2 <> 0 Then
            Picture1.Line (0, Y)-Step(Picture1.Width, 24), &HC4B6B5, BF
        End If
    Next
    For X = 0 To 3
        For Y = 0 To 31
            Picture1.CurrentY = Y * 24
            Picture1.CurrentX = 5 + (133 * X)
            Picture1.ForeColor = &H80&
            Picture1.Print GetHex(C) & "  ";
            Picture1.ForeColor = &H4000&
            Picture1.Print Format(C, "000") & "  ";
            Picture1.ForeColor = &HC00000
            Picture1.Print Chr(C)
            C = C + 1
        Next
    Next

End Sub
Private Sub LoadLPK()

    Open PathApp & SelectedLPK For Random As #3 Len = Len(LPK)
    Me.Caption = GetMsg(54)
    Msgs(0) = GetMsg(55)
    Msgs(1) = GetMsg(56)
    Command1.Caption = GetMsg(57)
    Close #3
    
End Sub
Private Function GetHex(ByVal Vle As Integer) As String

    GetHex = IIf(Len(Hex(Vle)) < 2, "0" & Hex(Vle), Hex(Vle))

End Function
Private Sub Command1_Click()
    
    Unload Me
    Set frmAnsi = Nothing

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then Command1_Click

End Sub
Private Sub Form_Load()

    LoadLPK
    CenterForm Me
    VScroll1.Max = (Picture1.Height - Picture2.Height) * -1
    VScroll1.LargeChange = Picture2.Height
    VScroll1.SmallChange = 24
    Option1_Click 0

End Sub
Private Sub Option1_Click(Index As Integer)

    If Picture1.Visible Then Picture1.SetFocus
    If Index = 0 Then
        DrawTable 0
        Caption = " " & Msgs(0)
    Else
        DrawTable 128
        Caption = " " & Msgs(1)
    End If

End Sub
Private Sub VScroll1_Change()

    Picture1.Top = VScroll1.Value

End Sub
Private Sub VScroll1_Scroll()
    
    Picture1.Top = VScroll1.Value

End Sub
