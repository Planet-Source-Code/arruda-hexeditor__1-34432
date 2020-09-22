VERSION 5.00
Begin VB.Form frmPage 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Go to Page"
   ClientHeight    =   1410
   ClientLeft      =   2220
   ClientTop       =   3510
   ClientWidth     =   2610
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   2610
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   675
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   225
      Width           =   1770
   End
   Begin Project1.Command Command1 
      Height          =   330
      Left            =   495
      TabIndex        =   3
      Top             =   990
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      Caption         =   "&OK"
   End
   Begin Project1.Command Command2 
      Height          =   330
      Left            =   1530
      TabIndex        =   4
      Top             =   990
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      Caption         =   "&Cancel"
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   630
      TabIndex        =   2
      Top             =   630
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Page:"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   270
      Width           =   555
   End
End
Attribute VB_Name = "frmPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Msgs(0) As String
Private Sub LoadLPK()

    Open PathApp & SelectedLPK For Random As #3 Len = Len(LPK)
    Me.Caption = GetMsg(84)
    Label1 = GetMsg(85)
    Command2.Caption = GetMsg(57)
    Msgs(0) = GetMsg(83)
    Close #3
    
End Sub
Private Sub Command1_Click()

    If Not IsNumeric(Text1) Then
        MsgBox Msgs(0), vbOKOnly + vbInformation, Caption
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1)
        Text1.SetFocus
        Exit Sub
    End If
    If Val(Text1) < 1 Then Text1 = 1
    Me.Hide
    
End Sub
Private Sub Command2_Click()

    Label1.Tag = "CANCEL"
    Me.Hide

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then Command2_Click

End Sub
Private Sub Form_Load()
    
    LoadLPK
    CenterForm Me

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 8
        Case 13
            Command1_Click
            KeyAscii = 0
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
    End Select

End Sub


