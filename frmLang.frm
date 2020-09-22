VERSION 5.00
Begin VB.Form frmLang 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Languages"
   ClientHeight    =   1530
   ClientLeft      =   2160
   ClientTop       =   2595
   ClientWidth     =   3855
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin Project1.Command Command1 
      Height          =   330
      Left            =   1665
      TabIndex        =   2
      Top             =   1080
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
      Caption         =   "&Apply"
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   450
      Width           =   3660
   End
   Begin Project1.Command Command2 
      Height          =   330
      Left            =   2745
      TabIndex        =   3
      Top             =   1080
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
      Caption         =   "&Close"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Available &Languages"
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   225
      Width           =   3345
   End
End
Attribute VB_Name = "frmLang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    If Combo1.ListIndex = -1 Then
        Label1.Tag = "CANCEL"
    Else
        Label1.Tag = Combo1.Text & ".lpk"
    End If
    Me.Hide

End Sub
Private Sub LoadLPK()

    Open PathApp & SelectedLPK For Random As #3 Len = Len(LPK)
    Me.Caption = GetMsg(87)
    Command1.Caption = GetMsg(88)
    Command2.Caption = GetMsg(57)
    Label1 = GetMsg(86)
    Close #3
    
End Sub
Private Sub Command2_Click()

    Label1.Tag = "CANCEL"
    Me.Hide

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then Command2_Click

End Sub
Private Sub Form_Load()

    On Error Resume Next
    Dim F As String, n As Integer
    LoadLPK
    CenterForm Me
    F = Dir(PathApp & "*.lpk", vbArchive)
    Do
        If Trim(F) = "" Then Exit Do
        Combo1.AddItem Left(F, Len(F) - 4)
        If F = GetSetting("HexEdit", "General", "Language", "English.lpk") Then n = Combo1.NewIndex
        F = Dir
    Loop
    
    If Combo1.ListCount > 0 Then
        Combo1.ListIndex = n
    End If

End Sub
