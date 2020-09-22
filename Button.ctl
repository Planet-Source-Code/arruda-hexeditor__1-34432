VERSION 5.00
Begin VB.UserControl Command 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   ScaleHeight     =   62
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   225
      Top             =   2745
   End
End
Attribute VB_Name = "Command"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Enum BtState
    sDown = 1
    sUp = 2
    sFlat = 3
    sDisable = 4
End Enum
    
Dim Pt As POINTAPI

Private MouseOver As Boolean
Private MouseButton As Integer
Private ButtonState As BtState
Private mCaption As String
Const m_def_Caption = "Command"
Dim m_Caption As String
Dim Cap(2) As String

Event Click()
Private Sub DrawButton(ByVal State As BtState)
    
    On Error Resume Next
    UserControl.Cls
    ButtonState = State
    Select Case State
        Case sUp
            UserControl.Line (0, 0)-Step(ScaleWidth - 1, 0), vb3DHighlight
            UserControl.Line (0, 0)-Step(0, ScaleHeight - 1), vb3DHighlight
            UserControl.Line (0, ScaleHeight - 1)-Step(ScaleWidth - 1, 0), vb3DDKShadow
            UserControl.Line (ScaleWidth - 1, 0)-Step(0, ScaleHeight - 1), vb3DDKShadow
            DrawCaption 1
        Case sDown
            UserControl.Line (0, 0)-Step(ScaleWidth - 1, 0), vb3DDKShadow
            UserControl.Line (0, 0)-Step(0, ScaleHeight - 1), vb3DDKShadow
            UserControl.Line (0, ScaleHeight - 1)-Step(ScaleWidth - 1, 0), vb3DHighlight
            UserControl.Line (ScaleWidth - 1, 0)-Step(0, ScaleHeight - 1), vb3DHighlight
            DrawCaption 3
        Case sFlat
            UserControl.Line (0, 0)-Step(ScaleWidth - 1, ScaleHeight - 1), &HCFCFCF, B
            DrawCaption 1
        Case sDisable
            UserControl.Line (0, 0)-Step(ScaleWidth - 1, ScaleHeight - 1), &HCFCFCF, B
            DrawCaption 2
    End Select
    
End Sub
Private Sub DrawCaption(ByVal DrawType As Integer)
    
    cx = Int((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(Cap(0) & Cap(1) & Cap(2)) / 2))
    cy = Int((UserControl.ScaleHeight / 2) - (UserControl.TextHeight(Cap(0) & Cap(1) & Cap(2)) / 2))
    Select Case DrawType
        Case 1
            UserControl.CurrentX = cx
            UserControl.CurrentY = cy
            UserControl.FontUnderline = False
            UserControl.Print Cap(0);
            UserControl.FontUnderline = True
            UserControl.Print Cap(1);
            UserControl.FontUnderline = False
            UserControl.Print Cap(2)
        Case 2
            UserControl.CurrentX = cx + 1
            UserControl.CurrentY = cy + 1
            UserControl.ForeColor = vb3DHighlight
            UserControl.FontUnderline = False
            UserControl.Print Cap(0);
            UserControl.FontUnderline = True
            UserControl.Print Cap(1);
            UserControl.FontUnderline = False
            UserControl.Print Cap(2)
            UserControl.CurrentX = cx
            UserControl.CurrentY = cy
            UserControl.ForeColor = vb3DShadow
            UserControl.FontUnderline = False
            UserControl.Print Cap(0);
            UserControl.FontUnderline = True
            UserControl.Print Cap(1);
            UserControl.FontUnderline = False
            UserControl.Print Cap(2)
            UserControl.ForeColor = 0
        Case 3
            UserControl.CurrentX = cx + 1
            UserControl.CurrentY = cy + 1
            UserControl.FontUnderline = False
            UserControl.Print Cap(0);
            UserControl.FontUnderline = True
            UserControl.Print Cap(1);
            UserControl.FontUnderline = False
            UserControl.Print Cap(2)
    End Select
    
End Sub
Private Sub SetCaption()
    
    X = InStr(1, m_Caption, "&")
    Select Case X
        Case 0
            Cap(0) = m_Caption
            Cap(1) = ""
            Cap(2) = ""
        Case 1
            Cap(0) = ""
            Cap(1) = Mid(m_Caption, 2, 1)
            Cap(2) = Right(m_Caption, Len(m_Caption) - 2)
        Case Is > 1
            Cap(0) = Left(m_Caption, X - 1)
            Cap(1) = Mid(m_Caption, X + 1, 1)
            Cap(2) = Right(m_Caption, Len(m_Caption) - (X + 1))
    End Select
    UserControl.AccessKeys = Cap(1)

End Sub
Private Sub Timer1_Timer()
    
    If MouseButton = 1 Then Exit Sub
    GetCursorPos Pt
    Call ScreenToClient(UserControl.hwnd, Pt)
    If Pt.X < 0 Or Pt.Y < 0 Or Pt.X > ScaleWidth Or Pt.Y > ScaleHeight Then
        MouseOver = False
        Timer1.Enabled = False
        If ButtonState <> sFlat Then
            DrawButton sFlat
            ButtonState = sFlat
        End If
    Else
        MouseOver = True
    End If

End Sub
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

    RaiseEvent Click

End Sub
Private Sub UserControl_Initialize()

    UserControl.Cls
    PrevIndex = 0

End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MouseButton = Button
    If ButtonState <> sDown Then
        DrawButton sDown
        ButtonState = sDown
    End If

End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Not Timer1.Enabled Then Timer1.Enabled = True
    Select Case Button
        Case 1
            If MouseOver Then
                If ButtonState <> sDown Then
                    DrawButton sDown
                    ButtonState = sDown
                End If
            Else
                If ButtonState <> sUp Then
                    DrawButton sUp
                    ButtonState = sUp
                End If
            End If
        Case 0
            MouseOver = True
            If ButtonState <> sUp Then
                DrawButton sUp
                ButtonState = sUp
            End If
    End Select
    MouseButton = Button

End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MouseButton = 0
    DoEvents

End Sub
Private Sub UserControl_Resize()
    
    UserControl.Cls
    SetCaption
    If Not UserControl.Ambient.UserMode Then
        DrawButton sUp
    Else
        If Not UserControl.Enabled Then
            DrawButton sDisable
        Else
            DrawButton ButtonState
        End If
    End If

End Sub
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    
    Enabled = UserControl.Enabled

End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    
    UserControl.Enabled() = New_Enabled
    If Not New_Enabled Then
        ButtonState = sDisable
        MouseOver = False
        Timer1.Enabled = False
    Else
        If MouseOver Then
            If MouseButton = 1 Then
                ButtonState = sDown
            Else
                ButtonState = sUp
            End If
        Else
            ButtonState = sFlat
        End If
    End If
    If UserControl.Ambient.UserMode Then DrawButton ButtonState
    PropertyChanged "Enabled"

End Property
Private Sub UserControl_Click()
    
    RaiseEvent Click

End Sub
'MemberInfo=13,0,0,
Public Property Get Caption() As String
    
    Caption = m_Caption

End Property
Public Property Let Caption(ByVal New_Caption As String)
    
    m_Caption = New_Caption
    SetCaption
    DrawButton ButtonState
    PropertyChanged "Caption"
    
End Property
Private Sub UserControl_InitProperties()
    
    m_Caption = m_def_Caption
    
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    SetCaption
    If Not UserControl.Ambient.UserMode Then
        DrawButton sUp
    Else
        If Not UserControl.Enabled Then
            DrawButton sDisable
        Else
            DrawButton sFlat
        End If
    End If
    
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    
End Sub
