VERSION 5.00
Begin VB.Form frmFindHex 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find - Hex Mode"
   ClientHeight    =   1755
   ClientLeft      =   3165
   ClientTop       =   2790
   ClientWidth     =   5100
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   117
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   ShowInTaskbar   =   0   'False
   Begin Project1.Command Command3 
      Height          =   330
      Left            =   3600
      TabIndex        =   7
      Top             =   945
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      Caption         =   "&Count"
   End
   Begin Project1.Command Command2 
      Height          =   330
      Left            =   3600
      TabIndex        =   6
      Top             =   1305
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      Caption         =   "&Close"
   End
   Begin Project1.Command Command1 
      Height          =   330
      Left            =   3600
      TabIndex        =   5
      Top             =   360
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      Caption         =   "&Find Next"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Direction: "
      Height          =   870
      Left            =   90
      TabIndex        =   4
      Top             =   765
      Width           =   3435
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Down"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   3
         Top             =   585
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Up"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   2
         Top             =   270
         Width           =   1365
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   3420
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Byte sequence:"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   2910
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmFindHex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ptrFind As Long
Dim Msgs(3) As String
Private Function FormatBytes(txtBox As TextBox) As Integer
    
    Dim TempString1 As String, TempString2 As String
    TempString1 = txtBox.Text
    TempString2 = ""
    For i = 1 To Len(TempString1)
        TempString2 = TempString2 & Trim(Mid(TempString1, i, 1))
    Next i
    FormatBytes = Trim(Len(TempString2))
    txtBox = ""
    If Len(TempString2) Mod 2 <> 0 Then TempString2 = TempString2 & "0"
    For i = 1 To Len(TempString2) Step 2
        txtBox = txtBox & Mid(TempString2, i, 2) & " "
    Next i
    
End Function
Private Sub Command1_Click()
    
    Dim ret As Long, n As Long
    FormatBytes Text1
    If Len(Text1) = 0 Then
        Text1.SetFocus
        Exit Sub
    End If
    MousePointer = 11
    If Option1(1) Then
        ret = frmEditor.FindDown(Text1, 0)
    Else
        ret = frmEditor.FindUp(Text1, 0)
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

    MousePointer = 0
    Unload Me
    Set frmFindHex = Nothing

End Sub
Private Sub Command3_Click()
    
    FormatBytes Text1
    If Len(Text1) = 0 Then
        Text1.SetFocus
        Exit Sub
    End If
    MousePointer = 11
    MsgBox Chr(34) & Trim(Text1) & Chr(34) & " " & Msgs(2) & " " & frmEditor.CountBytes(Text1, 0) & " " & Msgs(3), vbInformation, Caption
    MousePointer = 0

End Sub
Private Sub Form_Activate()

    Text1.SetFocus

End Sub
Private Sub LoadLPK()

    Open PathApp & SelectedLPK For Random As #3 Len = Len(LPK)
    Me.Caption = GetMsg(63)
    Label1(0) = GetMsg(58)
    Frame2.Caption = GetMsg(80)
    Option1(0).Caption = GetMsg(59)
    Option1(1).Caption = GetMsg(60)
    Command1.Caption = GetMsg(61)
    Command3.Caption = GetMsg(62)
    Command2.Caption = GetMsg(57)
    Msgs(0) = GetMsg(64)
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
    
    Select Case UCase(Chr(KeyAscii))
        Case Chr(8)
        Case Chr(13)
            KeyAscii = 0
            Command1_Click
        Case "A", "B", "C", "D", "E", "F", 0 To 9
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            ptrFind = 0
        Case Else
            KeyAscii = 0
    End Select

End Sub
