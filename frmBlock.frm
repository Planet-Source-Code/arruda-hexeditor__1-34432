VERSION 5.00
Begin VB.Form frmBlock 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selecionar Bloco"
   ClientHeight    =   4050
   ClientLeft      =   3225
   ClientTop       =   1635
   ClientWidth     =   3000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin Project1.Command Command1 
      Height          =   330
      Left            =   135
      TabIndex        =   13
      Top             =   3600
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      Caption         =   "&OK"
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Copi&ar para o Clipboard"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   1125
      Value           =   -1  'True
      Width           =   2040
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "P&reencher com:"
      Height          =   195
      Index           =   1
      Left            =   270
      TabIndex        =   5
      Top             =   2610
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   870
      Left            =   135
      TabIndex        =   8
      Top             =   2610
      Width           =   2715
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   330
         Left            =   1350
         MaxLength       =   2
         TabIndex        =   7
         Top             =   405
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valor Hex:"
         Height          =   240
         Index           =   2
         Left            =   405
         TabIndex        =   6
         Top             =   450
         Width           =   870
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1575
      TabIndex        =   3
      Top             =   585
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1575
      TabIndex        =   1
      Top             =   180
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame2"
      Height          =   1365
      Left            =   135
      TabIndex        =   9
      Top             =   1125
      Width           =   2715
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fonte &Pascal"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   12
         Top             =   990
         Width           =   1905
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Fonte C"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   11
         Top             =   675
         Width           =   1905
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valores &Hex "
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin Project1.Command Command2 
      Height          =   330
      Left            =   1890
      TabIndex        =   14
      Top             =   3600
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      Caption         =   "&Cancelar"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Off&set Final (Dec):"
      Height          =   195
      Index           =   1
      Left            =   210
      TabIndex        =   2
      Top             =   630
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Offset &Inicial (Dec):"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   225
      Width           =   1350
   End
End
Attribute VB_Name = "frmBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    If Not IsNumeric(Text1) Then
        Text1.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(Text2) Then
        Text2.SetFocus
        Exit Sub
    End If
    If Option1(1) Then
        If Trim(Text3) = "" Then
            Text3.SetFocus
            Exit Sub
        End If
    End If
    Label1(0).Tag = "OK"
    Me.Hide

End Sub
Private Sub Command2_Click()

    Label1(0).Tag = "CANCEL"
    Me.Hide

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then Command2_Click

End Sub
Private Sub LoadLPK()

    Open PathApp & SelectedLPK For Random As #3 Len = Len(LPK)
    Me.Caption = GetMsg(95)
    Label1(0) = GetMsg(96)
    Label1(1) = GetMsg(97)
    Option1(0).Caption = GetMsg(98)
    Option2(0).Caption = GetMsg(99)
    Option2(1).Caption = GetMsg(100)
    Option2(2).Caption = GetMsg(101)
    Option1(1).Caption = GetMsg(102)
    Label1(2) = GetMsg(103)
    Command1.Caption = GetMsg(104)
    Command2.Caption = GetMsg(105)
    Close #3
    
End Sub
Private Sub Form_Load()

    LoadLPK
    CenterForm Me
    
End Sub
Private Sub Option1_Click(Index As Integer)

    If Option1(1) Then
        Option2(0).Enabled = False
        Option2(1).Enabled = False
        Option2(2).Enabled = False
        Text3.BackColor = &HFFFFFF
        Frame1.Enabled = True
        Text3.Enabled = True
        Text3.SetFocus
    Else
        Text3 = ""
        Text3.Enabled = False
        Text3.BackColor = &HC0C0C0
        Frame1.Enabled = False
        Option2(0).Enabled = True
        Option2(1).Enabled = True
        Option2(2).Enabled = True
    End If

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    
    Select Case Chr(KeyAscii)
        Case Chr(8)
        Case Chr(13)
            KeyAscii = 0
        Case "0" To "9"
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select

End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    
    Select Case Chr(KeyAscii)
        Case Chr(8)
        Case Chr(13)
            KeyAscii = 0
        Case "0" To "9"
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select

End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)

    Select Case UCase(Chr(KeyAscii))
        Case Chr(8)
        Case Chr(13)
            KeyAscii = 0
        Case "A", "B", "C", "D", "E", "F", "0" To "9"
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select

End Sub
