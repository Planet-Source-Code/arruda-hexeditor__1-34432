VERSION 5.00
Begin VB.Form frmFindTxt 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find - Text Mode"
   ClientHeight    =   1665
   ClientLeft      =   3225
   ClientTop       =   2040
   ClientWidth     =   5580
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   111
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   372
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Direction: "
      Height          =   915
      Left            =   2385
      TabIndex        =   7
      Top             =   630
      Width           =   1635
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Down"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   5
         Top             =   585
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Up"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   315
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Case: "
      Height          =   915
      Left            =   90
      TabIndex        =   6
      Top             =   630
      Width           =   2265
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Insensitive"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   2
         Top             =   315
         Width           =   2040
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Sensitive"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   3
         Top             =   585
         Value           =   -1  'True
         Width           =   1950
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1035
      TabIndex        =   1
      Top             =   180
      Width           =   2985
   End
   Begin Project1.Command Command3 
      Height          =   330
      Left            =   4095
      TabIndex        =   8
      Top             =   855
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      Caption         =   "&Count"
   End
   Begin Project1.Command Command2 
      Height          =   330
      Left            =   4095
      TabIndex        =   9
      Top             =   1215
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      Caption         =   "&Close"
   End
   Begin Project1.Command Command1 
      Height          =   330
      Left            =   4095
      TabIndex        =   10
      Top             =   180
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      Caption         =   "&Find Next"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Find  &What:"
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   225
      Width           =   915
   End
End
Attribute VB_Name = "frmFindTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ret As Long, ptrFind As Long
Dim Msgs(4) As String
Private Sub Command1_Click()
    
    Dim ret As Long, t As Integer
    If Len(Trim(Text1)) = 0 Then
        Text1.SetFocus
        Exit Sub
    End If
    
    t = IIf(Option1(0) = True, 1, 2)
    MousePointer = 11
    If Option2(1) Then
        ret = frmEditor.FindDown(Text1, t)
    Else
        ret = frmEditor.FindUp(Text1, t)
    End If
    MousePointer = 0
    Select Case ret
        Case -1
            If ptrFind = 0 Then
                MsgBox Msgs(0), vbInformation, Caption
                FindNext.Search = False
            Else
                MsgBox Msgs(1), vbInformation, Caption
            End If
        Case Else
            ptrFind = ret
            FindNext.Search = True
    End Select

End Sub
Private Sub Command2_Click()

    Unload Me
    Set frmFindTxt = Nothing

End Sub
Private Sub Command3_Click()
    
    Dim t As Integer
    If Len(Text1) = 0 Then
        Text1.SetFocus
        Exit Sub
    End If
    t = IIf(Option1(0) = True, 1, 2)
    MousePointer = 11
    MsgBox Chr(34) & Trim(Text1) & Chr(34) & " " & Msgs(2) & " " & frmEditor.CountBytes(Text1, t) & " " & Msgs(3), vbInformation, Caption
    MousePointer = 0

End Sub
Private Sub Form_Activate()

    Text1.SetFocus

End Sub
Private Sub LoadLPK()

    Open PathApp & SelectedLPK For Random As #3 Len = Len(LPK)
    Me.Caption = GetMsg(81)
    Label1 = GetMsg(74)
    Frame1.Caption = GetMsg(79)
    Frame2.Caption = GetMsg(80)
    Option1(0).Caption = GetMsg(76)
    Option1(1).Caption = GetMsg(77)
    Option2(0).Caption = GetMsg(59)
    Option2(1).Caption = GetMsg(60)
    Command1.Caption = GetMsg(61)
    Command3.Caption = GetMsg(62)
    Command2.Caption = GetMsg(57)
    Msgs(0) = GetMsg(78)
    Msgs(1) = GetMsg(65)
    Msgs(2) = GetMsg(66)
    Msgs(3) = GetMsg(67)
    Close #3
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then Command2_Click

End Sub
Private Sub Form_Load()
    
    LoadLPK
    CenterForm Me

End Sub
Private Sub Text1_GotFocus()

    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)

    ptrFind = 0
    If KeyAscii = 13 Then
        KeyAscii = 0
        Command1_Click
    End If
    
End Sub
