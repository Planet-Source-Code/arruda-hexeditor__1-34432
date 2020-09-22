VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Language"
   ClientHeight    =   840
   ClientLeft      =   1155
   ClientTop       =   2460
   ClientWidth     =   3945
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   405
      Width           =   2715
   End
   Begin Project1.Command Command1 
      Height          =   330
      Left            =   2880
      TabIndex        =   2
      Top             =   405
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      Caption         =   "&Select"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Available &Languages"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   2385
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Command1_Click
    End If

End Sub
Private Sub Command1_Click()

    Me.Hide

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        KeyAscii = 0
        Command1_Click
    End If

End Sub
Private Sub Form_Load()

    CenterForm Me

End Sub
