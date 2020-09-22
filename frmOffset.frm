VERSION 5.00
Begin VB.Form frmOffSet 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Go to Offset"
   ClientHeight    =   1620
   ClientLeft      =   1890
   ClientTop       =   2070
   ClientWidth     =   2640
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   176
   ShowInTaskbar   =   0   'False
   Begin Project1.Command Command1 
      Height          =   330
      Left            =   1530
      TabIndex        =   5
      Top             =   810
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      Caption         =   "&OK"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   780
      Left            =   90
      TabIndex        =   2
      Top             =   720
      Width           =   1320
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Hex"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   4
         Top             =   495
         Width           =   600
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Dec"
         Height          =   195
         Index           =   0
         Left            =   135
         Picture         =   "frmOffset.frx":0000
         TabIndex        =   3
         Top             =   225
         Value           =   -1  'True
         Width           =   600
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   630
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   225
      Width           =   1860
   End
   Begin Project1.Command Command2 
      Height          =   330
      Left            =   1530
      TabIndex        =   6
      Top             =   1170
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      Caption         =   "&Cancel"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Off&set:"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   270
      Width           =   465
   End
End
Attribute VB_Name = "frmOffSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Msgs(0) As String
Private Sub LoadLPK()

    Open PathApp & SelectedLPK For Random As #3 Len = Len(LPK)
    Me.Caption = GetMsg(82)
    Command2.Caption = GetMsg(57)
    Msgs(0) = GetMsg(83)
    Close #3
    
End Sub
Private Sub Command1_Click()

    If frmOffSet.Option1(1) Then Text1 = "&H" & Text1
    If Not IsNumeric(Text1) Then
        MsgBox Msgs(0), vbOKOnly + vbInformation, Caption
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1)
        Text1.SetFocus
        Exit Sub
    End If
    Label1.Tag = "OK"
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
            If Not IsNumeric(Chr(KeyAscii)) Then
                Select Case UCase(Chr(KeyAscii))
                    Case "A", "B", "C", "D", "E", "F"
                        KeyAscii = Asc(UCase(Chr(KeyAscii)))
                    Case Else
                        KeyAscii = 0
                End Select
            End If
    End Select

End Sub


