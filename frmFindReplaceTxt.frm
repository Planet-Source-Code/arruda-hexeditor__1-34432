VERSION 5.00
Begin VB.Form frmFindReplaceTxt 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find and Replace - Text Mode"
   ClientHeight    =   2295
   ClientLeft      =   3270
   ClientTop       =   1635
   ClientWidth     =   5640
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   153
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1215
      TabIndex        =   3
      Top             =   720
      Width           =   2850
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Direction: "
      Height          =   915
      Left            =   2430
      TabIndex        =   9
      Top             =   1215
      Width           =   1635
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Down"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   585
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Up"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   315
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Case: "
      Height          =   915
      Left            =   135
      TabIndex        =   8
      Top             =   1215
      Width           =   2265
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Insensitive"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   315
         Width           =   2085
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Sensitive"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   5
         Top             =   585
         Value           =   -1  'True
         Width           =   1995
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1215
      TabIndex        =   1
      Top             =   180
      Width           =   2850
   End
   Begin Project1.Command cmdCount 
      Height          =   330
      Left            =   4140
      TabIndex        =   10
      Top             =   1440
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      Caption         =   "Cou&nt"
   End
   Begin Project1.Command cmdClose 
      CausesValidation=   0   'False
      Height          =   330
      Left            =   4140
      TabIndex        =   11
      Top             =   1800
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      Caption         =   "&Close"
   End
   Begin Project1.Command cmdFind 
      Height          =   330
      Left            =   4140
      TabIndex        =   12
      Top             =   180
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      Caption         =   "&Find Next"
   End
   Begin Project1.Command cmdReplaceAll 
      Height          =   330
      Left            =   4140
      TabIndex        =   13
      Top             =   990
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      Caption         =   "Replace &All"
   End
   Begin Project1.Command cmdReplace 
      Height          =   330
      Left            =   4140
      TabIndex        =   14
      Top             =   630
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      Caption         =   "&Replace"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Replace With:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   2
      Top             =   765
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Find  &What:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   225
      Width           =   1020
   End
End
Attribute VB_Name = "frmFindReplaceTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ret As Long, ptrFind As Long
Dim Msgs(4) As String
Private Sub cmdFind_Click()
    
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
Private Function FormatBytes(ByVal StrText As String) As String
    
    If Len(StrText) Mod 2 <> 0 Then StrText = StrText & "0"
    For i = 1 To Len(StrText) Step 2
        FormatBytes = FormatBytes & Mid(StrText, i, 2) & " "
    Next i
    
End Function
Private Sub cmdClose_Click()

    Unload Me
    Set frmFindTxt = Nothing

End Sub
Private Sub cmdCount_Click()
    
    Dim t As Integer
    If Len(Text1) = 0 Then
        Text1.SetFocus
        Exit Sub
    End If
    t = IIf(Option1(0) = True, 1, 2)
    MousePointer = 11
    MsgBox Chr(34) & Trim(Text1) & Chr(34) & " " & Msgs(3) & " " & frmEditor.CountBytes(Text1, t) & " " & Msgs(4), vbInformation, Caption
    MousePointer = 0

End Sub
Private Sub cmdReplace_Click()
    
    Dim ret As Long, t As Integer
    Dim tempString As String, StrText As String
    If Len(Trim(Text1)) = 0 Then
        Text1.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Text2)) = 0 Then
        Text2.SetFocus
        Exit Sub
    End If
    
    tempString = ""
    For i = 1 To Len(Text2)
        StrText = Hex(Asc(Mid(Text2, i, 1)))
        StrText = IIf(Len(StrText) < 2, "0" & StrText, StrText)
        tempString = tempString & StrText
    Next
    StrText = FormatBytes(tempString)
    t = IIf(Option1(0) = True, 1, 2)
    MousePointer = 11
    If Option2(1) Then
        If cmdReplace.Tag <> "" Then frmEditor.ReplaceBytes StrText, ptrFind
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
                frmEditor.FillGrid
                MsgBox Msgs(1), vbInformation, Caption
            End If
        Case Else
            ptrFind = ret
            cmdReplace.Tag = "replace"
            FindNext.Search = True
    End Select

End Sub
Private Sub cmdReplaceAll_Click()
    
    Dim t As Integer
    If Len(Text1) = 0 Then
        Text1.SetFocus
        Exit Sub
    End If
    If Len(Text2) = 0 Then
        Text2.SetFocus
        Exit Sub
    End If
    
    Select Case Len(Text2)
        Case Is < Len(Text1)
            Text2 = Text2 & Right(Text1, Len(Text1) - Len(Text2))
        Case Is > Len(Text1)
            Text2 = Left(Text2, Len(Text1))
    End Select
    
    tempString = ""
    For i = 1 To Len(Text2)
        StrText = Hex(Asc(Mid(Text2, i, 1)))
        StrText = IIf(Len(StrText) < 2, "0" & StrText, StrText)
        tempString = tempString & StrText
    Next
    StrText = FormatBytes(tempString)
    t = IIf(Option1(0) = True, 1, 2)
    MousePointer = 11
    ret = frmEditor.ReplaceAll(Text1, StrText, t)
    If ret = 0 Then
        MsgBox Msgs(0), vbInformation, Caption
    Else
        MsgBox Msgs(1) & " " & ret & " " & Msgs(2), vbInformation, Caption
        ptrFind = 0
        cmdReplace.Tag = ""
    End If
    MousePointer = 0

End Sub
Private Sub Form_Activate()

    Text1.SetFocus

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then cmdClose_Click

End Sub
Private Sub Form_Load()
    
    LoadLPK
    CenterForm Me

End Sub
Private Sub LoadLPK()

    Open PathApp & SelectedLPK For Random As #3 Len = Len(LPK)
    Me.Caption = GetMsg(73)
    Label1(0) = GetMsg(74)
    Label1(1) = GetMsg(69)
    Frame1.Caption = GetMsg(79)
    Frame2.Caption = GetMsg(80)
    Option1(0).Caption = GetMsg(76)
    Option1(1).Caption = GetMsg(77)
    Option2(0).Caption = GetMsg(59)
    Option2(1).Caption = GetMsg(60)
    cmdFind.Caption = GetMsg(61)
    cmdReplace.Caption = GetMsg(70)
    cmdReplaceAll.Caption = GetMsg(71)
    cmdCount.Caption = GetMsg(62)
    cmdClose.Caption = GetMsg(57)
    Msgs(0) = GetMsg(78)
    Msgs(1) = GetMsg(65)
    Msgs(2) = GetMsg(72)
    Msgs(3) = GetMsg(66)
    Msgs(4) = GetMsg(67)
    Close #3
    
End Sub
Private Sub Text1_Change()

    Text2.MaxLength = Len(Text1)

End Sub
Private Sub Text1_GotFocus()

    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)

    ptrFind = 0
    cmdReplace.Tag = ""
    
End Sub
Private Sub Text2_GotFocus()
    
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2)

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdFind_Click
    End If

End Sub


