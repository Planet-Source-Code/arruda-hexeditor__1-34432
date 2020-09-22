VERSION 5.00
Begin VB.Form frmConvert 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Conversion"
   ClientHeight    =   2580
   ClientLeft      =   1410
   ClientTop       =   3165
   ClientWidth     =   3975
   ControlBox      =   0   'False
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   172
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   ShowInTaskbar   =   0   'False
   Begin Project1.Command Command1 
      Height          =   375
      Left            =   2700
      TabIndex        =   8
      Top             =   2115
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      Caption         =   "&Close"
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   810
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1710
      Width           =   3030
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   810
      MaxLength       =   32
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1260
      Width           =   3030
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   810
      MaxLength       =   8
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   810
      Width           =   3030
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   810
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3030
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Octal"
      Height          =   240
      Index           =   3
      Left            =   45
      TabIndex        =   6
      Top             =   1755
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Binary"
      Height          =   240
      Index           =   2
      Left            =   45
      TabIndex        =   4
      Top             =   1305
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Decimal"
      Height          =   240
      Index           =   1
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Hex"
      Height          =   240
      Index           =   0
      Left            =   45
      TabIndex        =   2
      Top             =   855
      Width           =   645
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActiveTextBox As Integer
Public Function cvBinToDec(ByVal NumBin As String) As String
    
    NumBin = CStr(NumBin)
    For t = 1 To Len(NumBin)
        If Mid(NumBin, t, 1) = 1 Then
            vl = vl + (2 ^ (Len(NumBin) - t))
        End If
    Next
    cvBinToDec = CStr(vl)
    
End Function
Private Function cvBin(ByVal Value As String) As String
    
    On Error GoTo l1
    Dim n As String
    If Not IsNumeric(Value) Then GoTo l1
    n = IIf(Len(Hex(Value)) Mod 2 <> 0, "0" & Hex(Value), Hex(Value))
    For j = 1 To Len(n) Step 2
        For i = 7 To 0 Step -1
            If (GetVal(Mid(n, j, 2)) And 2 ^ i) Then
                s = s & "1"
            Else
                s = s & "0"
            End If
        Next
    Next
    cvBin = s
    Exit Function
l1:
    cvBin = ""
    Exit Function

End Function
Private Function cvHexToDec(ByVal Value As String) As String

    On Error GoTo l1
    cvHexToDec = CLng("&H" & Value)
    Exit Function
l1:
    cvHexToDec = 0

End Function

Private Function cvOctToDec(ByVal Value As String) As String

    On Error GoTo l1
    cvOctToDec = CLng("&O" & Value)
    Exit Function
l1:
    cvOctToDec = 0

End Function
Private Function cvHex(ByVal Value As Variant) As String

    On Error GoTo l1
    If Not IsNumeric(Value) Then
        cvHex = ""
        Exit Function
    End If
    cvHex = Hex((Value))
    Exit Function
l1:
    cvHex = ""
    Exit Function

End Function
Private Function cvOct(ByVal Value As String) As String

    On Error GoTo l1
    If Not IsNumeric(Value) Then
        cvOct = ""
        Exit Function
    End If
    cvOct = Oct(Value)
    Exit Function
l1:
    cvOct = ""

End Function
Function GetVal(ByVal HexValue As String) As Long

    If Trim(HexValue) = "" Then GetVal = 0 Else GetVal = "&H" & HexValue

End Function
Private Sub Command1_Click()

    Unload Me
    Set frmConvert = Nothing

End Sub
Private Sub Form_Activate()

    Text1.SetFocus

End Sub
Private Sub LoadLPK()

    Open PathApp & SelectedLPK For Random As #3 Len = Len(LPK)
    Me.Caption = GetMsg(51)
    Command1.Caption = GetMsg(57)
    Close #3
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then Command1_Click

End Sub
Private Sub Form_Load()

    LoadLPK
    CenterForm Me

End Sub
Private Sub Text1_Change()

    If ActiveTextBox <> 1 Then Exit Sub
    Text2 = cvHex(Text1)
    Text3 = cvBin(Text1)
    Text4 = cvOct(Text1)
    

End Sub

Private Sub Text1_GotFocus()
    
    ActiveTextBox = 1
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 8
        Case 13
            KeyAscii = 0
        Case Asc(0) To Asc(9), Asc("-")
        Case Else
            KeyAscii = 0
    End Select

End Sub
Private Sub Text2_Change()

    If ActiveTextBox <> 2 Then Exit Sub
    Value = cvHexToDec(Text2)
    Text1 = Value
    Text3 = cvBin(Value)
    Text4 = cvOct(Value)

End Sub
Private Sub Text2_GotFocus()
    
    ActiveTextBox = 2
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2)

End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
        Case 8
        Case 13
            KeyAscii = 0
        Case 48 To 57, 65 To 70
        Case Else
            KeyAscii = 0
    End Select

End Sub
Private Sub Text3_Change()
    
    If ActiveTextBox <> 3 Then Exit Sub
    Value = cvBinToDec(Text3)
    Text1 = Value
    Text2 = cvHex(Value)
    Text4 = cvOct(Value)

End Sub
Private Sub Text3_GotFocus()

    ActiveTextBox = 3
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3)

End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case 8
        Case 13
            KeyAscii = 0
        Case 48, 49
        Case Else
            KeyAscii = 0
    End Select

End Sub
Private Sub Text4_Change()

    If ActiveTextBox <> 4 Then Exit Sub
    Value = cvOctToDec(Text4)
    Text1 = Value
    Text2 = cvHex(Value)
    Text3 = cvBin(Value)

End Sub
Private Sub Text4_GotFocus()
    
    ActiveTextBox = 4
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4)

End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case 8
        Case 13
            KeyAscii = 0
        Case 48 To 55
        Case Else
            KeyAscii = 0
    End Select

End Sub
