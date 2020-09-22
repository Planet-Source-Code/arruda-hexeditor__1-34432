VERSION 5.00
Begin VB.Form frmLPKEditor 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Language Pack Editor"
   ClientHeight    =   5220
   ClientLeft      =   2610
   ClientTop       =   2175
   ClientWidth     =   4725
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1755
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   270
      Width           =   2805
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1755
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   3
      Top             =   720
      Width           =   2805
   End
   Begin Project1.Command Command1 
      Height          =   375
      Left            =   2370
      TabIndex        =   10
      Top             =   4725
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      Caption         =   "&Save"
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1215
      MaxLength       =   100
      TabIndex        =   7
      Top             =   4005
      Width           =   3345
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      Height          =   2400
      Left            =   135
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   4425
   End
   Begin Project1.Command Command2 
      Height          =   375
      Left            =   3465
      TabIndex        =   11
      Top             =   4725
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      Caption         =   "&Close"
   End
   Begin Project1.Command Command3 
      Height          =   375
      Left            =   135
      TabIndex        =   8
      Top             =   4725
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      Caption         =   "&New"
   End
   Begin Project1.Command Command4 
      Height          =   375
      Left            =   1230
      TabIndex        =   9
      Top             =   4725
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Edit"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Language:"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   2
      Top             =   765
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "String to Translate"
      Height          =   240
      Index           =   1
      Left            =   135
      TabIndex        =   4
      Top             =   1215
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Translate to:"
      Height          =   240
      Index           =   3
      Left            =   135
      TabIndex        =   6
      Top             =   4050
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Source Language:"
      Height          =   240
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   315
      Width           =   1410
   End
End
Attribute VB_Name = "frmLPKEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempLPK() As StructL
Dim SourceLang As String
Dim EditMode As Integer
Dim Msgs(8) As String
Private Const MODE_NONE = 0
Private Const MODE_NEW = 1
Private Const MODE_EDIT = 2
Private Function TestStrings() As Boolean

    For i = 1 To UBound(TempLPK)
        If Trim(TempLPK(i).Msg) = "" Then
            TestStrings = False
            Exit Function
        End If
    Next
    TestStrings = True
    
End Function
Private Sub PopulateList()
    
    If Trim(SourceLang) = "" Then Exit Sub
    Open PathApp & SourceLang & ".lpk" For Random As #3 Len = Len(LPK)
    ReDim Preserve TempLPK(1 To LOF(3) / Len(LPK))
    List1.Clear
    For i = 1 To LOF(3) / Len(LPK)
        Get #3, i, LPK
        List1.AddItem Trim(LPK.Msg)
        List1.ItemData(List1.NewIndex) = LPK.Id
    Next
    Close #3

End Sub
Private Sub Command1_Click()

    Select Case EditMode
        Case MODE_NONE
            Exit Sub
        Case MODE_NEW
            If Trim(Text2) = "" Then
                MsgBox Msgs(4), vbInformation, Caption
                Text2.SetFocus
                Exit Sub
            End If
            If Dir(PathApp & Text2 & ".lpk", vbArchive) <> "" Then
                If MsgBox(Msgs(0), vbQuestion + vbYesNo, Caption) = vbNo Then Exit Sub
            End If
            Msg = Msgs(1)
        Case MODE_EDIT
            Msg = Msgs(2)
    End Select
    
    If Not TestStrings Then
        MsgBox Msgs(3), vbInformation, Caption
        Exit Sub
    End If
    Open PathApp & Text2 & ".lpk" For Random As #4 Len = Len(LPK)
    For i = 1 To UBound(TempLPK)
        Put #4, i, TempLPK(i)
    Next
    Close #4
    MsgBox Msg, vbInformation, Caption

End Sub
Private Sub Command2_Click()

    Unload Me

End Sub
Private Sub Command3_Click()

    frmSelect.Caption = Msgs(5)
    frmSelect.Show 1
    If frmSelect.Combo1.ListIndex = -1 Then Exit Sub
    SourceLang = frmSelect.Combo1.Text
    Text3 = SourceLang
    PopulateList
    For i = 1 To UBound(TempLPK)
        TempLPK(i).Id = 0
        TempLPK(i).Msg = ""
    Next
    Label1(2) = Msgs(8)
    Text1 = ""
    Text2.Locked = False
    Text2 = ""
    Text2.SetFocus
    EditMode = MODE_NEW
    
End Sub
Private Sub Command4_Click()
    
    Text1 = ""
    frmSelect.Caption = Msgs(7)
    frmSelect.Show 1
    If frmSelect.Combo1.ListIndex = -1 Then Exit Sub
    SourceLang = frmSelect.Combo1.Text
    Text3 = SourceLang
    Text2 = SourceLang
    Label1(2) = Msgs(6)
    PopulateList
    Text2.Locked = True
    EditMode = MODE_EDIT
    If Dir(PathApp & Text2 & ".lpk", vbArchive) <> "" Then
        Open PathApp & Text2 & ".lpk" For Random As #3 Len = Len(LPK)
        For i = 1 To LOF(3) / Len(LPK)
            Get #3, i, LPK
            TempLPK(i).Id = LPK.Id
            TempLPK(i).Msg = Trim(LPK.Msg)
        Next
        Close #3
    End If
    If List1.ListCount > 0 Then List1.ListIndex = 0

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        KeyAscii = 0
        Command2_Click
    End If
    
End Sub
Private Sub Form_Load()
    
    On Error Resume Next
    Dim F As String, n As Integer
    LoadLPK
    Load frmSelect
    CenterForm Me
    frmSelect.Combo1.Clear
    F = Dir(PathApp & "*.lpk", vbArchive)
    Do
        If Trim(F) = "" Then Exit Do
        frmSelect.Combo1.AddItem Left(F, Len(F) - 4)
        If F = GetSetting("HexEdit", "General", "Language", "English.lpk") Then n = frmSelect.Combo1.NewIndex
        F = Dir
    Loop
    If frmSelect.Combo1.ListCount > 0 Then
        frmSelect.Combo1.ListIndex = n
    End If

End Sub
Private Sub LoadLPK()

    Open PathApp & SelectedLPK For Random As #3 Len = Len(LPK)
    Me.Caption = GetMsg(106)
    Label1(0) = GetMsg(107)
    Label1(1) = GetMsg(109)
    Label1(2) = GetMsg(108)
    Label1(3) = GetMsg(110)
    Command1.Caption = GetMsg(113)
    Command2.Caption = GetMsg(57)
    Command3.Caption = GetMsg(111)
    Command4.Caption = GetMsg(112)
    Msgs(0) = GetMsg(115)
    Msgs(1) = GetMsg(116)
    Msgs(2) = GetMsg(117)
    Msgs(3) = GetMsg(118)
    Msgs(4) = GetMsg(119)
    Msgs(5) = GetMsg(120)
    Msgs(6) = GetMsg(121)
    Msgs(7) = GetMsg(122)
    Msgs(8) = GetMsg(108)
    frmSelect.Command1.Caption = GetMsg(123)
    frmSelect.Label1 = GetMsg(124)
    Close #3
    
End Sub
Private Sub List1_Click()

    Text1 = Trim(TempLPK(List1.ItemData(List1.ListIndex)).Msg)

End Sub
Private Sub List1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1.SetFocus
    End If
    
End Sub
Private Sub Text1_Change()

    If List1.ListIndex = -1 Then Exit Sub
    TempLPK(List1.ItemData(List1.ListIndex)).Msg = Trim(Text1)
    TempLPK(List1.ItemData(List1.ListIndex)).Id = List1.ItemData(List1.ListIndex)

End Sub
Private Sub Text1_GotFocus()

    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If List1.ListIndex < List1.ListCount - 1 Then List1.ListIndex = List1.ListIndex + 1
    End If
    
End Sub
Private Sub Text2_LostFocus()

    X = InStr(1, Text2, ".")
    If X > 0 Then Text2 = Left(Text2, X - 1)

End Sub
