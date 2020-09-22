VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Grid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents Pic As PictureBox
Attribute Pic.VB_VarHelpID = -1
Private Shp As Shape

'General Variables
Private ColWidth As Integer
Private RowHeight As Integer
Private OldCol As Integer
Private OldRow As Integer
Private KeyPressed As Boolean
Private FireEvent As Boolean

'Grid Properties
Private mRow As Integer
Private mCol As Integer
Private mRows As Integer
Private mValue() As String
Private mPointer As Long
Private mMax As Long
Private mHasFocus As Boolean

'Grid events
Public Event EnterCell(ByVal Row As Integer, ByVal Col As Integer)
Public Event RequestNextRow()
Public Event RequestPreviousRow()
Public Event RequestNextPage()
Public Event RequestPreviousPage()
Public Event RequestBOF()
Public Event RequestEOF()
Public Event LeaveCell(OldRow As Integer, OldCol As Integer)
Public Event Click()
Private Sub DrawData(ByVal RowIndex As Integer, ByVal ColIndex As Integer)

    If isInvalidCell(RowIndex, ColIndex) Then
        mValue(RowIndex, ColIndex) = "-1"
        Exit Sub
    End If
    TextOut Pic.hdc, (ColIndex * ColWidth) + (ColWidth / 2) - (Pic.TextWidth(mValue(RowIndex, ColIndex)) / 2), (RowIndex * RowHeight) + ((RowHeight / 2) - (Pic.TextHeight("ql") / 2)), mValue(RowIndex, ColIndex), Len(mValue(RowIndex, ColIndex))

End Sub
Public Sub InitializeGrid(ByVal PicCtl As PictureBox, ByVal ChildCtl As Shape)
    
    Set Pic = PicCtl
    Set Shp = ChildCtl
    Col = 0
    Row = 0
    RowHeight = Int(Pic.Height / 20)
    ColWidth = Int(Pic.Width / 16)
    Pic.ScaleMode = vbPixels
    Pic.AutoRedraw = True
    Shp.Move 0, 0, ColWidth, RowHeight

End Sub
Private Function isInvalidCell(ByVal RowIndex As Integer, ByVal ColIndex As Integer) As Boolean
    
    isInvalidCell = ((mPointer + (RowIndex * 16) + ColIndex) > mMax)
    
End Function
Private Sub KeyDown()

    FireEvent = False
    If mRow + 1 < mRows Then
        mRow = mRow + 1
        MoveShape mRow, mCol
    Else
        If (mPointer + (mRows * 16)) <= mMax Then
            RaiseEvent LeaveCell(OldRow, OldCol)
            If mHasFocus Then RaiseEvent RequestNextRow
            FireEvent = True
            If SetLastCell(mRow, mCol) Then Exit Sub
            MoveShape mRow, mCol
        Else
            Exit Sub
        End If
    End If

End Sub
Private Sub KeyEnd(ByVal Shift As Integer)

    FireEvent = False
    If Shift = 2 Then
        mRow = mRows - 1
        mCol = 15
        RaiseEvent LeaveCell(OldRow, OldCol)
        If mHasFocus Then RaiseEvent RequestEOF
        FireEvent = True
        SetLastCell mRow, mCol
    Else
        mCol = 15
        MoveShape mRow, mCol
    End If

End Sub
Private Sub KeyLeft()
    
    FireEvent = False
    Select Case mCol - 1
        Case 0
           mCol = 0
            MoveShape mRow, mCol
        Case Is > 0
            mCol = mCol - 1
            MoveShape mRow, mCol
        Case Is < 0
            If (mRow - 1) > -1 Then
                mRow = mRow - 1
                mCol = 15
                MoveShape mRow, mCol
            Else
                If mPointer > 1 Then
                    RaiseEvent LeaveCell(OldRow, OldCol)
                    If mHasFocus Then RaiseEvent RequestPreviousRow
                    FireEvent = True
                    mCol = 15
                    mRow = 0
                    MoveShape mRow, mCol
                Else
                    Exit Sub
                End If
            End If
    End Select

End Sub
Private Sub KeyPageDown(ByVal Shift As Integer)
    
    FireEvent = False
    If Shift <> 2 Then
        RaiseEvent LeaveCell(OldRow, OldCol)
        If mHasFocus Then RaiseEvent RequestNextPage
        FireEvent = True
        If SetLastCell(mRow, mCol) Then Exit Sub
        MoveShape mRow, mCol
    Else
        mRow = mRows - 1
        MoveShape mRow, mCol
    End If

End Sub
Private Sub KeyHome(ByVal Shift As Integer)
            
    FireEvent = False
    If Shift = 2 Then
        mCol = 0
        mRow = 0
        RaiseEvent LeaveCell(OldRow, OldCol)
        If mHasFocus Then RaiseEvent RequestBOF
        FireEvent = True
        MoveShape 0, 0
    Else
        mCol = 0
        MoveShape mRow, mCol
    End If

End Sub
Private Sub KeyPageUp(ByVal Shift As Integer)
            
    FireEvent = False
    If Shift <> 2 Then
        RaiseEvent LeaveCell(OldRow, OldCol)
        If mHasFocus Then RaiseEvent RequestPreviousPage
        FireEvent = True
        MoveShape mRow, mCol
    Else
        mRow = 0
        MoveShape mRow, mCol
    End If

End Sub
Private Sub KeyRight()
    
    FireEvent = False
    Select Case mCol + 1
        Case 15
            mCol = 15
            MoveShape mRow, mCol
        Case Is < 15
            mCol = mCol + 1
            MoveShape mRow, mCol
        Case Is > 15
            If (mRow + 1) < mRows Then
                mRow = mRow + 1
                mCol = 0
                MoveShape mRow, mCol
            Else
                If mPointer + (mRows * (mCol + 1)) < mMax Then
                    RaiseEvent LeaveCell(OldRow, OldCol)
                    If mHasFocus Then RaiseEvent RequestNextRow
                    FireEvent = True
                    If SetLastCell(mRow, 0) Then Exit Sub
                    mCol = 0
                    MoveShape mRow, mCol
                Else
                    Exit Sub
                End If
            End If
    End Select

End Sub
Private Sub KeyUp()
    
    FireEvent = False
    If (mRow - 1) > -1 Then
        mRow = mRow - 1
        MoveShape mRow, mCol
    Else
        If mPointer > 1 Then
            RaiseEvent LeaveCell(OldRow, OldCol)
            If mHasFocus Then RaiseEvent RequestPreviousRow
            FireEvent = True
            MoveShape mRow, mCol
        Else
            Exit Sub
        End If
    End If

End Sub
Public Sub Refresh()

    SetLastCell mRow, mCol

End Sub
Public Sub Reset()

    Col = 0
    Row = 0
    OldCol = 0
    OldRow = 0
    MoveShape 0, 0
    
End Sub
Public Property Get Row() As Integer

    Row = mRow
    
End Property
Private Sub MoveShape(ByVal RowIndex As Integer, ColIndex As Integer)

    If isInvalidCell(RowIndex, ColIndex) Then
        mRow = OldRow
        mCol = OldCol
        Exit Sub
    End If
    If Not FireEvent Then RaiseEvent LeaveCell(IIf(OldRow <= mRows - 1, OldRow, mRows - 1), OldCol)
    Row = IIf(RowIndex <= mRows - 1, RowIndex, mRows - 1)
    Col = ColIndex
    Shp.Move Int(mCol * ColWidth), Int(mRow * RowHeight)
    RaiseEvent EnterCell(mRow, mCol)
    
End Sub
Public Sub SelectCell(ByVal RowIndex As Integer, ByVal ColIndex As Integer)
    
    MoveShape RowIndex, ColIndex

End Sub
Private Function SetLastCell(ByVal RowIndex As Integer, ByVal ColIndex As Integer) As Boolean
    
    Dim i As Integer, j As Integer
    If isInvalidCell(RowIndex, ColIndex) Then
        SetLastCell = True
        For i = RowIndex To 0 Step -1
            For j = ColIndex To 0 Step -1
                If Not isInvalidCell(i, j) Then
                    MoveShape i, j
                    Exit Function
                End If
            Next
        Next
    Else
        SetLastCell = False
    End If
    
End Function
Public Property Get Value(ByVal RowIndex As Integer, ByVal ColIndex As Integer) As String

    Value = mValue(RowIndex, ColIndex)
    
End Property
Public Property Get Col() As Integer

    Col = mCol
    
End Property
Public Property Get HasFocus() As Boolean

    HasFocus = mHasFocus
    
End Property
Public Property Get Pointer() As Long

    Pointer = mPointer
    
End Property
Public Property Get Rows() As Integer

    Rows = mRows
    
End Property
Public Property Let Row(ByVal newData As Integer)

    If OldRow <> mRow Then OldRow = mRow
    If OldRow > mRows - 1 Then OldRow = mRows - 1
    mRow = newData

End Property
Public Property Let Value(ByVal RowIndex As Integer, ByVal ColIndex As Integer, ByVal newData As String)

    mValue(RowIndex, ColIndex) = newData
    DrawData RowIndex, ColIndex
    
End Property
Public Property Let Col(ByVal newData As Integer)

    If OldCol <> mCol Then OldCol = mCol
    mCol = newData

End Property
Public Property Let Max(ByVal newData As Long)

    mMax = newData

End Property
Public Property Let Pointer(ByVal newData As Long)

    mPointer = newData

End Property
Public Property Let Rows(ByVal newData As Integer)

    mRows = newData
    ReDim mValue(mRows, 0 To 15)

End Property
Private Sub Pic_Click()

    RaiseEvent Click

End Sub
Private Sub Pic_GotFocus()

    Shp.DrawMode = 6
    mHasFocus = True

End Sub
Private Sub Pic_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyDown
            KeyDown
        Case vbKeyUp
            KeyUp
        Case vbKeyLeft
            KeyLeft
        Case vbKeyRight
            KeyRight
        Case vbKeyPageDown
            KeyPageDown Shift
        Case vbKeyPageUp
            KeyPageUp Shift
        Case vbKeyHome
            KeyHome Shift
        Case vbKeyEnd
            KeyEnd Shift
    End Select
    
End Sub
Private Sub Pic_LostFocus()
    
    Shp.DrawMode = 9
    mHasFocus = False
    
End Sub
Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim R As Integer
    mHasFocus = True
    R = Int(Y / RowHeight)
    If R > mRows - 1 Then R = mRows - 1
    If SetLastCell(R, Int(X / ColWidth)) Then Exit Sub
    If (mRow <> R) Or (mCol <> Int(X / ColWidth)) Then
        mRow = R
        mCol = Int(X / ColWidth)
        MoveShape mRow, mCol
    End If
    
End Sub
