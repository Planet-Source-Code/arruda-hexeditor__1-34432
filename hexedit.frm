VERSION 5.00
Begin VB.Form frmEditor 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   495
   ClientTop       =   1215
   ClientWidth     =   11505
   FontTransparent =   0   'False
   HasDC           =   0   'False
   Icon            =   "hexedit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   767
   Tag             =   "arruda"
   Begin Project1.ButtonScroll ButtonScroll1 
      Height          =   1710
      Left            =   9315
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   5190
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   3016
   End
   Begin Project1.ButtonBar ButtonBar1 
      Height          =   555
      Left            =   15
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   979
   End
   Begin VB.PictureBox PicScroll 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   4560
      Left            =   9315
      Picture         =   "hexedit.frx":030A
      ScaleHeight     =   4560
      ScaleWidth      =   240
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   630
      Width           =   240
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   210
         Left            =   0
         Picture         =   "hexedit.frx":1666
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.PictureBox PicOffSet 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00AB8F8D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00000000&
      Height          =   5715
      Left            =   270
      ScaleHeight     =   381
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1110
      Width           =   1140
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      HasDC           =   0   'False
      Height          =   195
      Left            =   9645
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   119
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7350
      Visible         =   0   'False
      Width           =   1785
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Openning"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   52
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   1770
      End
   End
   Begin VB.PictureBox PicText 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00000080&
      Height          =   5715
      Left            =   6975
      ScaleHeight     =   381
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   151
      TabIndex        =   3
      Top             =   1110
      Width           =   2265
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.PictureBox PicHex 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   5715
      Left            =   1515
      ScaleHeight     =   381
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   352
      TabIndex        =   0
      Top             =   1110
      Width           =   5280
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1140
      Left            =   9780
      Picture         =   "hexedit.frx":17A6
      Top             =   5730
      Width           =   1530
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   90
      TabIndex        =   49
      Top             =   7350
      Width           =   9330
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   9585
      TabIndex        =   48
      Top             =   7305
      Width           =   1905
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   0
      TabIndex        =   47
      Top             =   7305
      Width           =   9555
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      Index           =   11
      X1              =   649
      X2              =   756
      Y1              =   379
      Y2              =   379
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   10
      X1              =   649
      X2              =   756
      Y1              =   378
      Y2              =   378
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      Index           =   9
      X1              =   649
      X2              =   756
      Y1              =   328
      Y2              =   328
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   8
      X1              =   649
      X2              =   756
      Y1              =   327
      Y2              =   327
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      Index           =   7
      X1              =   649
      X2              =   756
      Y1              =   277
      Y2              =   277
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   6
      X1              =   649
      X2              =   756
      Y1              =   276
      Y2              =   276
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      Index           =   5
      X1              =   649
      X2              =   756
      Y1              =   229
      Y2              =   229
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   4
      X1              =   649
      X2              =   756
      Y1              =   228
      Y2              =   228
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      Index           =   3
      X1              =   649
      X2              =   756
      Y1              =   163
      Y2              =   163
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   649
      X2              =   756
      Y1              =   162
      Y2              =   162
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      X1              =   649
      X2              =   756
      Y1              =   97
      Y2              =   97
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   649
      X2              =   756
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00E0E0E0&
      Height          =   6270
      Index           =   1
      Left            =   9735
      Top             =   645
      Width           =   1635
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      FillColor       =   &H80000015&
      Height          =   6270
      Index           =   0
      Left            =   9720
      Top             =   630
      Width           =   1635
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   620
      X2              =   620
      Y1              =   43
      Y2              =   459
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Offset"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Index           =   1
      Left            =   9810
      TabIndex        =   45
      Top             =   2565
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   2
      Left            =   9810
      TabIndex        =   44
      Top             =   3525
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Size"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   3
      Left            =   9810
      TabIndex        =   43
      Top             =   5040
      Width           =   675
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "00 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   0
      Left            =   10260
      TabIndex        =   42
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hex:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   195
      Index           =   4
      Left            =   9900
      TabIndex        =   41
      Top             =   1800
      Width           =   465
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dec:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   195
      Index           =   5
      Left            =   9900
      TabIndex        =   40
      Top             =   2070
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hex:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   195
      Index           =   6
      Left            =   9900
      TabIndex        =   39
      Top             =   2790
      Width           =   405
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dec:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   195
      Index           =   7
      Left            =   9900
      TabIndex        =   38
      Top             =   3060
      Width           =   420
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "00 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   1
      Left            =   10260
      TabIndex        =   37
      Top             =   2070
      Width           =   1005
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "00 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   2
      Left            =   10260
      TabIndex        =   36
      Top             =   2790
      Width           =   1005
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "00 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   3
      Left            =   10260
      TabIndex        =   35
      Top             =   3060
      Width           =   1005
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0 of 0 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   4
      Left            =   9810
      TabIndex        =   34
      Top             =   3735
      Width           =   1470
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0 Bytes "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   5
      Left            =   9810
      TabIndex        =   33
      Top             =   5265
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Byte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Index           =   0
      Left            =   9810
      TabIndex        =   32
      Top             =   1575
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Index           =   8
      Left            =   9810
      TabIndex        =   31
      Top             =   810
      Width           =   825
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hex"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   6
      Left            =   9810
      TabIndex        =   30
      Top             =   1035
      Width           =   705
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   7
      Left            =   10575
      TabIndex        =   29
      Top             =   1035
      Width           =   735
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0 Bytes "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   8
      Left            =   9810
      TabIndex        =   28
      Top             =   4500
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes per Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   9
      Left            =   9810
      TabIndex        =   27
      Top             =   4275
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Text Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   16
      Left            =   7155
      TabIndex        =   26
      Top             =   810
      Width           =   1890
   End
   Begin VB.Label Sign 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00AB8F8D&
      Caption         =   "·"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   1
      Left            =   180
      TabIndex        =   25
      Top             =   1155
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   8
      X2              =   621
      Y1              =   71
      Y2              =   71
   End
   Begin VB.Label Label6 
      BackColor       =   &H00AB8F8D&
      Enabled         =   0   'False
      Height          =   5805
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   459
      X2              =   459
      Y1              =   44
      Y2              =   459
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   98
      X2              =   98
      Y1              =   44
      Y2              =   459
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   15
      Left            =   6480
      TabIndex        =   20
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   14
      Left            =   6150
      TabIndex        =   19
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   13
      Left            =   5820
      TabIndex        =   18
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   12
      Left            =   5490
      TabIndex        =   17
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   11
      Left            =   5160
      TabIndex        =   16
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   10
      Left            =   4830
      TabIndex        =   15
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "09"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   9
      Left            =   4500
      TabIndex        =   14
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "08"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   8
      Left            =   4170
      TabIndex        =   13
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "07"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   7
      Left            =   3840
      TabIndex        =   12
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "06"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   6
      Left            =   3510
      TabIndex        =   11
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "05"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   5
      Left            =   3180
      TabIndex        =   10
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "04"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   2850
      TabIndex        =   9
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "03"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   2520
      TabIndex        =   8
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "02"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   2190
      TabIndex        =   7
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   1860
      TabIndex        =   6
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   1530
      TabIndex        =   2
      Top             =   810
      Width           =   315
   End
   Begin VB.Label lblSelect 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Offset - DEC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   315
      TabIndex        =   1
      Top             =   810
      Width           =   1020
   End
   Begin VB.Label Sign 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00AB8F8D&
      BackStyle       =   0  'Transparent
      Caption         =   "·"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   0
      Left            =   1620
      TabIndex        =   24
      Top             =   615
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label6 
      BackColor       =   &H00AB8F8D&
      Enabled         =   0   'False
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   660
      Width           =   9180
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   6285
      Left            =   90
      TabIndex        =   21
      Top             =   630
      Width           =   9480
   End
   Begin VB.Menu Mn 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu sm1 
         Caption         =   "Open"
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu sm1 
         Caption         =   "Save"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu sm1 
         Caption         =   "Save As..."
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^A
      End
      Begin VB.Menu sm1 
         Caption         =   "Properties"
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   ^P
      End
      Begin VB.Menu sm1 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu sm1 
         Caption         =   "Exit"
         Index           =   5
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Mn 
      Caption         =   "&Edit"
      Index           =   1
      Begin VB.Menu Sm2 
         Caption         =   "Copy"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu Sm2 
         Caption         =   "Undo"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^Z
      End
      Begin VB.Menu Sm2 
         Caption         =   "Select Block"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^B
      End
      Begin VB.Menu Sm2 
         Caption         =   "Find Hex Values"
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   ^F
      End
      Begin VB.Menu Sm2 
         Caption         =   "Find Text"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   ^T
      End
      Begin VB.Menu Sm2 
         Caption         =   "Find and Replace Hex Values"
         Enabled         =   0   'False
         Index           =   5
         Shortcut        =   ^H
      End
      Begin VB.Menu Sm2 
         Caption         =   "Find and Replace Text"
         Enabled         =   0   'False
         Index           =   6
         Shortcut        =   ^R
      End
      Begin VB.Menu Sm2 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu Sm2 
         Caption         =   "Find Next"
         Enabled         =   0   'False
         Index           =   8
         Shortcut        =   {F3}
      End
      Begin VB.Menu Sm2 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu Sm2 
         Caption         =   "Go to Offset"
         Enabled         =   0   'False
         Index           =   10
         Shortcut        =   ^G
      End
      Begin VB.Menu Sm2 
         Caption         =   "Go to Page"
         Enabled         =   0   'False
         Index           =   11
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu Mn 
      Caption         =   "&Bookmark"
      Index           =   2
      Begin VB.Menu Sm3 
         Caption         =   "Toggle Bookmark"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   {F5}
      End
      Begin VB.Menu Sm3 
         Caption         =   "Clear All Bookmarks"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   {F6}
      End
      Begin VB.Menu Sm3 
         Caption         =   "Previous bookmark"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   {F7}
      End
      Begin VB.Menu Sm3 
         Caption         =   "Next bookmark"
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu Mn 
      Caption         =   "&Tools"
      Index           =   3
      Begin VB.Menu Sm4 
         Caption         =   "Data Conversion"
         Index           =   0
         Shortcut        =   {F9}
      End
      Begin VB.Menu Sm4 
         Caption         =   "ANSI Table"
         Index           =   1
         Shortcut        =   {F11}
      End
      Begin VB.Menu Sm4 
         Caption         =   "Calculator"
         Index           =   2
         Shortcut        =   {F12}
      End
      Begin VB.Menu Sm4 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu Sm4 
         Caption         =   "Languages"
         Index           =   4
         Shortcut        =   ^L
      End
      Begin VB.Menu Sm4 
         Caption         =   "Language Pack Editor"
         Index           =   5
      End
   End
   Begin VB.Menu Mn 
      Caption         =   "&About"
      Index           =   4
      Begin VB.Menu Sm5 
         Caption         =   "&About Hex Editor"
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents grid1 As Grid
Attribute grid1.VB_VarHelpID = -1
Private WithEvents grid2 As Grid
Attribute grid2.VB_VarHelpID = -1
Private Const MAX_LINE = 20
Private hg As Integer
Private Ptr As Long
Private ActiveGrid As Integer
Private FileChanged As Boolean
Private UndoInProgress As Boolean
Private ShowHex As Boolean
Private Drag As Boolean
Private Y1 As Single
Private FileName As String
Private FileOpened As Boolean
Private BookMark() As Long
Private PtrBookMark As Long
Private Msgs(6) As String
    
Private Type sUndo
    Offset As Long
    Row As Integer
    Col As Integer
    Value As Byte
End Type
Private Und() As sUndo
Private Sub CreateTempFile()

    Dim B() As Byte, n As Long, TempDir As String
    Picture1.Cls
    Label8 = Msgs(5)
    DoEvents
    Const nBytes = 512000
    TempDir = GetTempDir
    If Right(TempDir, 1) <> "\" Then TempDir = TempDir & "\"
    If Dir(TempDir & "hexfiletemp.tmp") <> "" Then Kill TempDir & "hexfiletemp.tmp"
    Open TempDir & "hexfiletemp.tmp" For Binary As #1
    If LOF(2) <= nBytes Then
        ReDim B(1 To LOF(2))
        Get #2, , B
        Put #1, , B
        Progress 100
        Exit Sub
    Else
        ReDim B(1 To nBytes)
        n = Int(LOF(2) / nBytes)
        For i = 1 To n
            Get #2, , B
            Put #1, , B
            Progress (i * 100) / n
        Next
        If (LOF(2) - (n * nBytes)) > 0 Then
            ReDim B(1 To LOF(2) - (n * nBytes))
            Get #2, , B
            Put #1, , B
            Progress 100
        End If
    End If
    ReDim B(0)

End Sub

Private Sub EnableControls()

    Dim ctl As Control
    For i = 0 To 20
        If Not ButtonBar1.Enabled(i) Then ButtonBar1.Enabled(i) = True
        If i <= 5 Then
            If Not ButtonScroll1.Enabled(i) Then ButtonScroll1.Enabled(i) = True
        End If
    Next
    For Each ctl In frmEditor.Controls
        If TypeOf ctl Is Menu Then
            If Not ctl.Enabled Then ctl.Enabled = True
        End If
    Next
    PicScroll.Enabled = True
    PicHex.Enabled = True
    PicText.Enabled = True
    Shape1.Visible = True
    Shape2.Visible = True
    Sign(0).Visible = True
    Sign(1).Visible = True
    lblSelect.Enabled = True
    
End Sub
Private Sub DisableControls()

    Dim ctl As Control
    
    For i = 0 To 19
        Select Case i
            Case 0, 16, 17, 18, 19, 20
            Case Else
                ButtonBar1.Enabled(i) = False
        End Select
    Next
    
    For Each ctl In frmEditor.Controls
        If TypeOf ctl Is Menu Then
            Select Case UCase(ctl.Name)
                Case "SM4", "SM5", "MN"
                Case "SM1"
                    If (ctl.Index <> 0) And (ctl.Index <> 4) And (ctl.Index <> 5) Then ctl.Enabled = False
                Case "SM2"
                    If (ctl.Index <> 7) And (ctl.Index <> 9) Then ctl.Enabled = False
                Case Else
                    ctl.Enabled = False
            End Select
        End If
    Next
    PicHex.Enabled = False
    PicText.Enabled = False
    PicScroll.Enabled = False
    ButtonScroll1.Enabled(0) = False
    ButtonScroll1.Enabled(1) = False
    ButtonScroll1.Enabled(2) = False
    ButtonScroll1.Enabled(3) = False
    ButtonScroll1.Enabled(4) = False
    ButtonScroll1.Enabled(5) = False
    Shape1.Visible = False
    Shape2.Visible = False
    Sign(0).Visible = False
    Sign(1).Visible = False
    lblSelect.Enabled = False
    
End Sub
Sub FillGrid()
    
    DoEvents
    Dim X As Long, B(15) As Byte
    Seek #1, Ptr
    grid1.Pointer = Ptr
    grid2.Pointer = Ptr
    X = Int((LOF(1) - Ptr) / 16)
    If X < MAX_LINE Then
        If X < 1 Then X = 1
        X = X + 1
        PicHex.Cls
        PicOffSet.Cls
        PicText.Cls
    Else
        X = MAX_LINE
    End If
    grid1.Rows = X
    grid2.Rows = X
    For R = 0 To X - 1
        Get #1, , B
        For C = 0 To 15
            grid1.Value(R, C) = GetHex(B(C))
            grid2.Value(R, C) = GetChar(B(C))
        Next
    Next
    PrintOffSets
    lblInfo(4) = GetPages
    SetScrollPosition
    grid1.Refresh

End Sub
Private Function GetChar(ByVal Byt As String) As String

    If Byt = "-1" Then
        GetChar = ""
    Else
        GetChar = IIf(Byt < 33, ".", Chr(Byt))
    End If
    
End Function
Private Function GetHex(ByVal Byt As String) As String

    If Byt = "-1" Then
        GetHex = 0
    Else
        GetHex = IIf(Len(Hex(Byt)) < 2, "0" & Hex(Byt), Hex(Byt))
    End If
    
End Function
Private Function GetOffset(ByVal nOffset As Long) As String
    
    If ShowHex Then
        GetOffset = String(10 - Len(Hex(nOffset)), "0") & Hex(nOffset)
    Else
        GetOffset = Format(nOffset, "0000000000")
    End If
    
End Function
Private Function GetPages() As String
    
    nPg = Int(((LOF(1) - 1) / (MAX_LINE * 16)) + 1)
    GetPages = nPg - Int((LOF(1) - Ptr) / (MAX_LINE * 16)) & " " & Msgs(0) & " " & nPg

End Function
Private Function GetValue(ByVal HexValue As String) As Byte

    If HexValue = "-1" Then
        GetValue = 0
    Else
        GetValue = "&H" & Trim(HexValue)
    End If

End Function
Private Function isInvalidCell(ByVal R As Integer, ByVal C As Integer) As Boolean

    isInvalidCell = ((Ptr + (R * 16) + C) > LOF(1))

End Function
Private Sub LoadLPK()

    Open PathApp & SelectedLPK For Random As #3 Len = Len(LPK)
    Me.Caption = GetMsg(1)
    ButtonBar1.ToolTip(0) = " " & GetMsg(2) & " "
    ButtonBar1.ToolTip(1) = " " & GetMsg(3) & " "
    ButtonBar1.ToolTip(2) = " " & GetMsg(4) & " "
    ButtonBar1.ToolTip(3) = " " & GetMsg(93) & " "
    ButtonBar1.ToolTip(4) = " " & GetMsg(5) & " "
    ButtonBar1.ToolTip(5) = " " & GetMsg(94) & " "
    ButtonBar1.ToolTip(6) = " " & GetMsg(6) & " "
    ButtonBar1.ToolTip(7) = " " & GetMsg(8) & " "
    ButtonBar1.ToolTip(8) = " " & GetMsg(7) & " "
    ButtonBar1.ToolTip(9) = " " & GetMsg(9) & " "
    ButtonBar1.ToolTip(10) = " " & GetMsg(10) & " "
    ButtonBar1.ToolTip(11) = " " & GetMsg(11) & " "
    ButtonBar1.ToolTip(12) = " " & GetMsg(12) & " "
    ButtonBar1.ToolTip(13) = " " & GetMsg(13) & " "
    ButtonBar1.ToolTip(14) = " " & GetMsg(14) & " "
    ButtonBar1.ToolTip(15) = " " & GetMsg(15) & " "
    ButtonBar1.ToolTip(16) = " " & GetMsg(19) & " "
    ButtonBar1.ToolTip(17) = " " & GetMsg(17) & " "
    ButtonBar1.ToolTip(18) = " " & GetMsg(16) & " "
    ButtonBar1.ToolTip(19) = " " & GetMsg(89) & " "
    ButtonBar1.ToolTip(20) = " " & GetMsg(18) & " "
    ButtonScroll1.ToolTip(0) = " " & GetMsg(26) & " "
    ButtonScroll1.ToolTip(1) = " " & GetMsg(27) & " "
    ButtonScroll1.ToolTip(2) = " " & GetMsg(28) & " "
    ButtonScroll1.ToolTip(3) = " " & GetMsg(29) & " "
    ButtonScroll1.ToolTip(4) = " " & GetMsg(30) & " "
    ButtonScroll1.ToolTip(5) = " " & GetMsg(31) & " "
    
    Label1(0) = GetMsg(22)
    Label1(2) = GetMsg(23)
    Label1(3) = GetMsg(25)
    Label1(8) = GetMsg(21)
    Label1(9) = GetMsg(24)
    lblInfo(4) = "0 " & GetMsg(33) & " 0"
    lblInfo(7) = GetMsg(32)
    Label2(16) = GetMsg(20)
    
    Mn(0).Caption = GetMsg(34)
        sm1(0).Caption = GetMsg(35)
        sm1(1).Caption = GetMsg(36)
        sm1(2).Caption = GetMsg(130)
        sm1(3).Caption = GetMsg(37)
        sm1(5).Caption = GetMsg(18)
        
    
    Mn(1).Caption = GetMsg(49)
        Sm2(0).Caption = GetMsg(93)
        Sm2(1).Caption = GetMsg(5)
        Sm2(2).Caption = GetMsg(94)
        Sm2(3).Caption = GetMsg(38)
        Sm2(4).Caption = GetMsg(39)
        Sm2(5).Caption = GetMsg(40)
        Sm2(6).Caption = GetMsg(41)
        Sm2(8).Caption = GetMsg(92)
        Sm2(10).Caption = GetMsg(42)
        Sm2(11).Caption = GetMsg(43)
    
    Mn(2).Caption = GetMsg(44)
        Sm3(0).Caption = GetMsg(45)
        Sm3(1).Caption = GetMsg(46)
        Sm3(2).Caption = GetMsg(47)
        Sm3(3).Caption = GetMsg(48)
        
    Mn(3).Caption = GetMsg(50)
        Sm4(0).Caption = GetMsg(51)
        Sm4(1).Caption = GetMsg(52)
        Sm4(2).Caption = GetMsg(53)
        Sm4(4).Caption = GetMsg(87)
        Sm4(5).Caption = GetMsg(125)
    
    Mn(4).Caption = GetMsg(90)
        Sm5.Caption = GetMsg(91)
    
    Msgs(0) = GetMsg(33)
    Msgs(1) = GetMsg(65)
    Msgs(2) = GetMsg(126)
    Msgs(3) = GetMsg(127)
    Msgs(4) = GetMsg(128)
    Msgs(5) = GetMsg(129)
    Msgs(6) = GetMsg(130)
    
    Close #3
    
End Sub
Private Sub MoveLastPage()
    
    Ptr = ((Int(LOF(1) / 16) * 16) + 1) - ((MAX_LINE - 1) * 16)
    If Ptr < 1 Then Ptr = 1
    FillGrid

End Sub
Private Sub MoveFirstPage()
    
    Ptr = 1
    FillGrid

End Sub
Sub MoveNextLine()

    If grid1.Rows < MAX_LINE Then Exit Sub
    If EOF(1) Then Exit Sub
    If Ptr + 16 > LOF(1) Then Exit Sub
    Ptr = Ptr + 16
    FillGrid
    
End Sub
Sub MoveNextPage()
    
    If grid1.Rows < MAX_LINE Then Exit Sub
    If EOF(1) Then Exit Sub
    If Ptr + (MAX_LINE * 16) < LOF(1) Then
        Ptr = Ptr + (MAX_LINE * 16)
    Else
        Ptr = (Int(LOF(1) / 16) * 16) - (MAX_LINE * 16) + 1
    End If
    If Ptr < 1 Then Ptr = 1
    FillGrid

End Sub
Sub MovePreviousLine()

    Ptr = Ptr - 16
    If Ptr < 1 Then Ptr = 1
    FillGrid

End Sub
Private Sub MovePreviousPage()
    
    If Ptr <= 1 Then Exit Sub
    Ptr = Ptr - (MAX_LINE * 16)
    If Ptr < 1 Then Ptr = 1
    FillGrid

End Sub
Sub OpenFile(ByVal FileName As String)
    
    Label7 = FileName
    FileOpened = False
    Open FileName For Binary As #2
    FileOpened = True
    Picture1.Visible = True
    CreateTempFile
    Close #2
    Ptr = 1
    grid1.Max = LOF(1)
    grid2.Max = LOF(1)
    Image1.Top = 0
    Image1.Visible = (LOF(1) > (MAX_LINE * 16))
    lblInfo(5) = LOF(1)
    lblInfo(8) = MAX_LINE * 16
    Picture1.Visible = False
    ReDim BookMark(0)
    
End Sub
Private Sub PrintOffSets()
   
    For i = 0 To grid1.Rows - 1
        If isInvalidCell(i, 0) Then Exit For
        tmpString = GetOffset((i * 16) + (grid1.Pointer - 1))
        TextOut PicOffSet.hdc, 3, hg * i + 3, tmpString, Len(tmpString)
    Next
    PicOffSet.Refresh

End Sub
Private Sub Progress(ByVal nPercent As Integer)

    n = Picture1.ScaleWidth / 100
    Picture1.Line (0, 0)-(n * nPercent, Picture1.ScaleHeight), &H0, BF
    DoEvents
    
End Sub
Public Function FindDown(ByVal StrWhat As String, ByVal FindType As Integer) As Long
    
    'FindType:
    '0 = Hex Mode
    '1 = Text mode Case Sensitive
    '2 = Text mode Case Insensitive
    
    Dim bWith() As Byte
    Dim bWhat() As Byte
    Dim B() As Byte
    Dim LastFind As Boolean
    Dim X As Long
    Dim PtrIni As Long
    Dim Findmode As Integer
    Dim WhatLen As Integer
    
    PtrIni = grid1.Pointer + ((grid1.Row * 16) + grid1.Col) + 1
    If PtrIni > LOF(1) Then
        FindDown = -1
        Exit Function
    End If
    
    FindNext.FindType = FindType
    FindNext.StrWhat = StrWhat
    FindNext.Direction = 0
    FindNext.LastFind = 0
    
    If FindType = 0 Then
        j = 0
        For i = 1 To Len(StrWhat) Step 3
            ReDim Preserve bWhat(j)
            bWhat(j) = "&H" & Trim(Mid(StrWhat, i, 3))
            j = j + 1
        Next
        WhatLen = UBound(bWhat)
    ElseIf FindType = 1 Then
        Findmode = 1
        WhatLen = Len(StrWhat)
    Else
        Findmode = 0
        WhatLen = Len(StrWhat)
    End If
    
    ReDim B(1 To 32000)
    Seek #1, PtrIni
    Do Until EOF(1)
        If (LOF(1) - PtrIni) > 32000 Then
            LastFind = False
        Else
            ReDim B(1 To (LOF(1) - PtrIni) + 1)
            LastFind = True
        End If
        Seek #1, PtrIni
        Get #1, , B
        If FindType = 0 Then
            X = InStr(1, StrConv(B, vbUnicode), StrConv(bWhat, vbUnicode), vbBinaryCompare)
        Else
            X = InStr(1, StrConv(B, vbUnicode), StrWhat, Findmode)
        End If
        If X > 0 Then
            Ptr = (PtrIni + (X - 1) + 15)
            Ptr = (Int(Ptr / 16) * 16) - 15
            cl = ((PtrIni - Ptr) + X) - 1
            If cl < 0 Then cl = 0
            FillGrid
            grid1.Refresh
            grid2.Refresh
            grid1.SelectCell 0, cl
            grid2.SelectCell 0, cl
            FindDown = PtrIni + (X - 1)
            Exit Function
        End If
        If LastFind Then Exit Do
        PtrIni = (PtrIni + UBound(B)) - WhatLen
    Loop
    FindDown = -1
    
End Function
Public Function CountBytes(ByVal StrWhat As String, ByVal FindType As Integer) As Long
    
    Dim bWhat() As Byte
    Dim B() As Byte
    Dim LastFind As Boolean
    Dim X As Long, n As Long
    Dim PtrIni As Long
    Dim StartPoint As Long
    Dim Findmode As Integer
    Dim WhatLen As Long
    
    If FindType = 0 Then
        j = 0
        For i = 1 To Len(StrWhat) Step 3
            ReDim Preserve bWhat(j)
            bWhat(j) = "&H" & Trim(Mid(StrWhat, i, 3))
            j = j + 1
        Next
        WhatLen = UBound(bWhat)
    ElseIf FindType = 1 Then
        Findmode = 1
        WhatLen = Len(StrWhat)
    Else
        Findmode = 0
        WhatLen = Len(StrWhat)
    End If
    
    n = 0
    ReDim B(1 To 32000)
    PtrIni = 1
    Seek #1, PtrIni
    Do Until EOF(1)
        If (LOF(1) - PtrIni) < 32000 Then
            ReDim B((LOF(1) - PtrIni))
            LastFind = True
        Else
            LastFind = False
        End If
        Seek #1, PtrIni
        Get #1, , B
        StartPoint = 1
        Do
            If FindType = 0 Then
                X = InStr(StartPoint, StrConv(B, vbUnicode), StrConv(bWhat, vbUnicode), vbBinaryCompare)
            Else
                X = InStr(StartPoint, StrConv(B, vbUnicode), StrWhat, Findmode)
            End If
            If X = 0 Then Exit Do
            n = n + 1
            StartPoint = X + 1
        Loop
        If LastFind Then Exit Do
        PtrIni = (PtrIni + UBound(B)) - WhatLen
    Loop
    CountBytes = n
    
End Function
Public Function ReplaceAll(ByVal StrWhat As String, ByVal StrWith As String, ByVal FindType As Integer) As Long
    
    Dim bWhat() As Byte
    Dim bWith() As Byte
    Dim B() As Byte
    Dim LastFind As Boolean
    Dim X As Long, n As Long
    Dim PtrIni As Long
    Dim StartPoint As Long
    Dim Findmode As Integer
    Dim WhatLen As Integer
    
    If FindType = 0 Then
        j = 0
        For i = 1 To Len(StrWhat) Step 3
            ReDim Preserve bWhat(j)
            bWhat(j) = "&H" & Trim(Mid(StrWhat, i, 3))
            j = j + 1
        Next
        WhatLen = UBound(bWhat)
    ElseIf FindType = 1 Then
        Findmode = 1
        WhatLen = Len(StrWhat)
    Else
        Findmode = 0
        WhatLen = Len(StrWhat)
    End If
    
    j = 0
    For i = 1 To Len(StrWith) Step 3
        ReDim Preserve bWith(j)
        bWith(j) = "&H" & Trim(Mid(StrWith, i, 3))
        j = j + 1
    Next
    
    n = 0
    ReDim B(1 To 32000)
    PtrIni = 1
    Seek #1, PtrIni
    Do Until EOF(1)
        If (LOF(1) - PtrIni) < 32000 Then
            ReDim B((LOF(1) - PtrIni))
            LastFind = True
        Else
            LastFind = False
        End If
        Seek #1, PtrIni
        Get #1, , B
        StartPoint = 1
        Do
            If FindType = 0 Then
                X = InStr(StartPoint, StrConv(B, vbUnicode), StrConv(bWhat, vbUnicode), vbBinaryCompare)
            Else
                X = InStr(StartPoint, StrConv(B, vbUnicode), StrWhat, Findmode)
            End If
            If X = 0 Then Exit Do
            StartPoint = X + 1
            n = n + 1
            Seek #1, (PtrIni + X) - 1
            Put #1, , bWith
        Loop
        If LastFind Then Exit Do
        PtrIni = (PtrIni + UBound(B)) - WhatLen
    Loop
    FillGrid
    ReplaceAll = n
    
End Function
Public Function FindUp(ByVal StrWhat As String, ByVal FindType As Integer) As Long
    
    'FindType:
    '0 = Hex Mode
    '1 = Text mode Case Sensitive
    '2 = Text mode Case Insensitive
    
    Dim bWith() As Byte
    Dim bWhat() As Byte
    Dim B() As Byte
    Dim LastFind As Boolean
    Dim X As Long
    Dim PtrIni As Long
    Dim Findmode As Integer
    Dim WhatLen As Integer
    
    FindNext.FindType = FindType
    FindNext.StrWhat = StrWhat
    FindNext.Direction = 1
    FindNext.LastFind = 1
    
    If FindType = 0 Then
        j = 0
        For i = 1 To Len(StrWhat) Step 3
            ReDim Preserve bWhat(j)
            bWhat(j) = "&H" & Trim(Mid(StrWhat, i, 3))
            j = j + 1
        Next
        WhatLen = UBound(bWhat)
    ElseIf FindType = 1 Then
        Findmode = 1
        WhatLen = Len(StrWhat)
    Else
        Findmode = 0
        WhatLen = Len(StrWhat)
    End If
    
    PtrIni = grid1.Pointer + ((grid1.Row * 16) + grid1.Col) - 1
    If PtrIni < 1 Then
        FindUp = -1
        Exit Function
    End If
    
    ReDim B(1 To (32000 + WhatLen))
    Do
        If (PtrIni - 32000) >= 1 Then
            PtrIni = PtrIni - 32000
            LastFind = False
        Else
            If PtrIni < WhatLen Then PtrIni = WhatLen
            ReDim B(1 To PtrIni)
            PtrIni = 1
            LastFind = True
        End If
        Seek #1, PtrIni
        Get #1, , B
        
        If FindType = 0 Then
            X = InStrRev(StrConv(B, vbUnicode), StrConv(bWhat, vbUnicode), -1, vbBinaryCompare)
        Else
            X = InStrRev(StrConv(B, vbUnicode), StrWhat, -1, Findmode)
        End If
        
        If X > 0 Then
            Ptr = (PtrIni + (X - 1) + 15)
            Ptr = (Int(Ptr / 16) * 16) - 15
            cl = ((PtrIni - Ptr) + X) - 1
            If cl < 0 Then cl = 0
            FillGrid
            FindUp = PtrIni + (X - 1)
            grid1.SelectCell 0, cl
            grid2.SelectCell 0, cl
            Exit Function
        End If
        If LastFind Then Exit Do
    Loop
    FindUp = -1

    
End Function
Public Sub ReplaceBytes(ByVal StrWith As String, ByVal Offset As Long)
    
    Dim bWith() As Byte
    j = 0
    For i = 1 To Len(StrWith) Step 3
        ReDim Preserve bWith(j)
        bWith(j) = "&H" & Trim(Mid(StrWith, i, 3))
        j = j + 1
    Next
    Seek #1, Offset
    Put #1, , bWith
    
End Sub
Private Sub ButtonBar1_Click(Index As Integer)

    Select Case Index
        Case 0: cmdOpen
        Case 1: cmdSave
        Case 2: cmdProperties
        Case 3: cmdCopy
        Case 4: cmdUndo
        Case 5: SelectBlock
        Case 6: cmdFindHex
        Case 7: cmdReplaceHex
        Case 8: cmdFindText
        Case 9: cmdReplaceText
        Case 10: cmdGoOffset
        Case 11: cmdGoPage
        Case 12: cmdBookmark
        Case 13: cmdPreviousBookmark
        Case 14: cmdNextBookmark
        Case 15: cmdClearBookmark
        Case 16: cmdConversion
        Case 17: cmdCharTable
        Case 18: cmdCalculator
        Case 19: cmdLanguage
        Case 20: cmdClose
    End Select
        
End Sub
Private Sub cmdSave()
    
    On Error GoTo ErrSave
    If Not FileOpened Then Exit Sub
    If Not FileChanged Then Exit Sub
    
    If MsgBox(Msgs(2) & " " & Chr(34) & FileName & Chr(34) & "?", vbQuestion + vbYesNo, Caption) = vbNo Then Exit Sub
    Dim B() As Byte, n As Long
    Picture1.Cls
    Label8 = Msgs(4)
    Picture1.Visible = True
    MousePointer = 11
    Const nBytes = 512000
    Seek #1, 1
    Open FileName For Binary As #2
    If LOF(1) <= nBytes Then
        ReDim B(1 To LOF(1))
        Get #1, , B
        Put #2, , B
        Progress 100
        FileChanged = False
        Picture1.Visible = False
        MousePointer = 0
        Close #2
        Exit Sub
    Else
        ReDim B(1 To nBytes)
        n = Int(LOF(1) / nBytes)
        For i = 1 To n
            Get #1, , B
            Put #2, , B
            Progress (i * 100) / n
        Next
        If (LOF(1) - (n * nBytes)) > 0 Then
            ReDim B(1 To LOF(1) - (n * nBytes))
            Get #1, , B
            Put #2, , B
            Progress 100
        End If
    End If
    FileChanged = False
    Picture1.Visible = False
    MousePointer = 0
    Close #2
    Exit Sub
    
ErrSave:
    Close #2
    MousePointer = 0
    MsgBox Err.Number & Chr(10) & Err.Description
    
End Sub
Private Sub cmdSaveAs()
    
    On Error GoTo ErrSave
    If Not FileOpened Then Exit Sub
    Dim FileAs As String
    FileAs = ShowSave(hwnd, Dir(FileName, vbArchive), "All Files (*.*)|*.*", Msgs(6), FileName)
    If Trim(FileAs) = "" Then Exit Sub
    FileAs = Left(FileAs, InStr(1, FileAs, Chr(0)) - 1)
    
    FileName = FileAs
    Label7 = FileAs
    
    Dim B() As Byte, n As Long
    Picture1.Cls
    Label8 = Msgs(4)
    Picture1.Visible = True
    MousePointer = 11
    Const nBytes = 512000
    Seek #1, 1
    Open FileAs For Binary As #2
    If LOF(1) <= nBytes Then
        ReDim B(1 To LOF(1))
        Get #1, , B
        Put #2, , B
        Progress 100
        FileChanged = False
        Picture1.Visible = False
        MousePointer = 0
        Close #2
        Exit Sub
    Else
        ReDim B(1 To nBytes)
        n = Int(LOF(1) / nBytes)
        For i = 1 To n
            Get #1, , B
            Put #2, , B
            Progress (i * 100) / n
        Next
        If (LOF(1) - (n * nBytes)) > 0 Then
            ReDim B(1 To LOF(1) - (n * nBytes))
            Get #1, , B
            Put #2, , B
            Progress 100
        End If
    End If
    FileChanged = False
    Picture1.Visible = False
    Close #2
    MousePointer = 0
    Exit Sub
    
ErrSave:
    Close #2
    MousePointer = 0
    MsgBox Err.Number & Chr(10) & Err.Description
    
End Sub

Private Sub cmdUndo()

    If Not FileOpened Then Exit Sub
    DoUndos

End Sub
Private Sub cmdGoOffset()
    
    If Not FileOpened Then Exit Sub
    frmOffSet.Show 1
    If frmOffSet.Label1.Tag = "CANCEL" Then
        Unload frmOffSet
        Set frmOffSet = Nothing
        Exit Sub
    End If
    Ptr = CLng(frmOffSet.Text1) + 1
    If Ptr > LOF(1) Then Ptr = LOF(1)
    If Ptr < 1 Then Ptr = 1
    FillGrid
    grid1.SelectCell 0, 0
    grid2.SelectCell 0, 0
    grid1.Refresh
    PicHex.SetFocus
    Unload frmOffSet
    Set frmOffSet = Nothing

End Sub
Private Sub cmdGoPage()
    
    If Not FileOpened Then Exit Sub
    Dim Pg As Long, nBytes As Integer
    Load frmPage
    frmPage.Label2 = "1  -  " & Int(LOF(1) / (16 * MAX_LINE)) + 1
    frmPage.Show 1
    If frmPage.Label1.Tag = "CANCEL" Then GoTo l1
    Pg = CLng(frmPage.Text1)
    If Pg > Int(LOF(1) / (16 * MAX_LINE)) + 1 Then Pg = Int(LOF(1) / (16 * MAX_LINE)) + 1
    nBytes = (16 * MAX_LINE)
    Ptr = Int((Pg * nBytes) - nBytes) + 1
    FillGrid
    grid1.SelectCell 0, 0
    grid2.SelectCell 0, 0
    PicHex.SetFocus
    grid1.Refresh
l1:
    Unload frmPage
    Set frmPage = Nothing

End Sub
Private Sub cmdBookmark()
    
    If Not FileOpened Then Exit Sub
    n = UBound(BookMark) + 1
    ReDim Preserve BookMark(n)
    BookMark(n) = CLng(lblInfo(3)) + 1
    PtrBookMark = n
    PicHex.SetFocus

End Sub
Private Sub cmdPreviousBookmark()
    
    If Not FileOpened Then Exit Sub
    If UBound(BookMark) = 0 Then Exit Sub
    PtrBookMark = PtrBookMark - 1
    If PtrBookMark < 1 Then PtrBookMark = 1
    Ptr = BookMark(PtrBookMark)
    FillGrid
    grid1.SelectCell 0, 0
    grid1.Refresh

End Sub
Private Sub cmdNextBookmark()
    
    If Not FileOpened Then Exit Sub
    If UBound(BookMark) = 0 Then Exit Sub
    PtrBookMark = PtrBookMark + 1
    If PtrBookMark > UBound(BookMark) Then PtrBookMark = UBound(BookMark)
    Ptr = BookMark(PtrBookMark)
    FillGrid
    grid1.SelectCell 0, 0
    grid1.Refresh

End Sub
Private Sub cmdConversion()
    
    frmConvert.Show 1

End Sub
Private Sub cmdCalculator()

    Shell "Calc.exe", vbNormalFocus

End Sub
Private Sub cmdClose()

    Unload Me

End Sub
Private Sub cmdCopy()
    
    If Not FileOpened Then Exit Sub
    Clipboard.Clear
    If ActiveGrid = 1 Then
        Clipboard.SetText grid1.Value(grid1.Row, grid1.Col)
    Else
        Clipboard.SetText Chr(GetValue(grid1.Value(grid2.Row, grid2.Col)))
    End If

End Sub
Private Sub ButtonScroll1_Click(Index As Integer)

    If Not FileOpened Then Exit Sub
    Select Case Index
        Case 0
            MoveFirstPage
        Case 1
            MovePreviousPage
        Case 2
            MovePreviousLine
        Case 3
            MoveNextLine
        Case 4
            MoveNextPage
        Case 5
            MoveLastPage
    End Select

End Sub
Private Sub cmdCharTable()

    frmAnsi.Show 1

End Sub
Private Sub cmdClearBookmark()
    
    If Not FileOpened Then Exit Sub
    ReDim BookMark(0)
    PtrBookMark = 0
    
End Sub
Private Sub cmdReplaceHex()
    
    If Not FileOpened Then Exit Sub
    frmFindReplaceHex.Show 1
    If ActiveGrid = 1 Then PicHex.SetFocus Else PicText.SetFocus

End Sub
Private Sub cmdReplaceText()
    
    If Not FileOpened Then Exit Sub
    frmFindReplaceTxt.Show 1
    If ActiveGrid = 1 Then PicHex.SetFocus Else PicText.SetFocus

End Sub
Private Sub cmdLanguage()

    frmLang.Show 1
    If frmLang.Label1.Tag <> "CANCEL" Then
        SelectedLPK = frmLang.Label1.Tag
        SaveSetting "HexEdit", "General", "Language", SelectedLPK
        LoadLPK
    End If
    Unload frmLang
    Set frmLang = Nothing

End Sub
Private Sub cmdOpen()
    
    Dim RetString As String, n As Long
    DisableControls
    LastDir = GetSetting("HexEdit", "General", "LastDir", App.Path)
    RetString = ShowOpen(hwnd, "All Files (*.*)|*.*", Caption, LastDir)
    If Trim(RetString) = "" Then
        If FileOpened Then EnableControls Else DisableControls
        Exit Sub
    End If
    
    RetString = Left(RetString, InStr(1, RetString, Chr(0)) - 1)
    n = InStr(1, RetString, Dir(RetString, vbArchive), vbTextCompare) - 1
    SaveSetting "HexEdit", "General", "LastDir", Left(RetString, n)
    
    DoEvents
    Reset
    ReDim Und(30)
    FileName = RetString
    OpenFile FileName
    FillGrid
    EnableControls
    grid1.Reset
    grid2.Reset
    PicHex.SetFocus
    
End Sub
Private Sub cmdProperties()
    
    If Not FileOpened Then Exit Sub
    Dim Prop As SHELLEXECUTEINFO
    Prop.fMask = &HC Or &H40 Or &H400
    Prop.hwnd = hwnd
    Prop.lpVerb = "PROPERTIES"
    Prop.lpFile = FileName
    Prop.lpParameters = vbNull
    Prop.lpDirectory = vbNull
    Prop.nShow = 0
    Prop.hInstApp = 0
    Prop.lpIDList = 0
    Prop.cbSize = Len(Prop)
    ShellExecuteEX Prop

End Sub
Private Sub cmdFindText()
    
    If Not FileOpened Then Exit Sub
    frmFindTxt.Show 1
    If ActiveGrid = 1 Then PicHex.SetFocus Else PicText.SetFocus

End Sub
Private Sub cmdFindHex()
    
    If Not FileOpened Then Exit Sub
    frmFindHex.Show 1
    If ActiveGrid = 1 Then PicHex.SetFocus Else PicText.SetFocus

End Sub
Private Sub SelectBlock()

    Dim A1 As Long
    Dim A2 As Long
    Dim Offset1 As Long
    Dim Offset2 As Long
    Dim Byt As Byte
    Dim B() As Byte
    Dim VarStr As String
    
    frmBlock.Show 1
    If frmBlock.Label1(0).Tag = "CANCEL" Then
        Unload frmBlock
        Set frmBlock = Nothing
        Exit Sub
    End If
    MousePointer = 11
    A1 = CLng(frmBlock.Text1)
    A2 = CLng(frmBlock.Text2)
    Offset1 = IIf(A1 < A2, A1, A2) + 1
    Offset2 = IIf(A1 >= A2, A1, A2) + 1
    If Offset1 > LOF(1) Then Offset1 = LOF(1)
    If Offset2 > LOF(1) Then Offset2 = LOF(1)
    Seek #1, Offset1
    If frmBlock.Option1(1) Then
        Byt = "&H" & frmBlock.Text3
        For i = Offset1 To Offset2
            Put #1, , Byt
        Next
    Else
        ReDim B(Offset2 - Offset1)
        Get #1, , B
        VarStr = ""
        If frmBlock.Option2(0) Then
            C = 0
            For i = Offset1 To Offset2
                C = C + 1
                Select Case C
                    Case Is < 16
                        VarStr = VarStr & IIf(Len(Hex(B(i - Offset1))) < 2, "0" & Hex(B(i - Offset1)), Hex(B(i - Offset1))) & " "
                    Case 16
                        C = 0
                        VarStr = VarStr & IIf(Len(Hex(B(i - Offset1))) < 2, "0" & Hex(B(i - Offset1)), Hex(B(i - Offset1))) & " " & Chr(13) & Chr(10)
                End Select
            Next
        End If
        
        If frmBlock.Option2(1) Then
            VarStr = "unsigned char data[" & UBound(B) + 1 & "] = {" & Chr(13) & Chr(10) & vbTab
            C = 0
            For i = Offset1 To Offset2
                C = C + 1
                Select Case C
                    Case Is < 16
                        VarStr = VarStr & "0x" & IIf(Len(Hex(B(i - Offset1))) < 2, "0" & Hex(B(i - Offset1)), Hex(B(i - Offset1))) & ", "
                    Case 16
                        C = 0
                        VarStr = VarStr & "0x" & IIf(Len(Hex(B(i - Offset1))) < 2, "0" & Hex(B(i - Offset1)), Hex(B(i - Offset1))) & ", " & Chr(13) & Chr(10) & vbTab
                End Select
            Next
            If Right(VarStr, 2) = ", " Then
                VarStr = Left(VarStr, Len(VarStr) - 2) & Chr(13) & Chr(10) & "};"
            Else
                VarStr = Left(VarStr, Len(VarStr) - 5) & Chr(13) & Chr(10) & "};"
            End If
        End If
        If frmBlock.Option2(2) Then
            VarStr = "data: array[0.." & UBound(B) & "] of byte = (" & Chr(13) & Chr(10) & vbTab
            C = 0
            For i = Offset1 To Offset2
                C = C + 1
                Select Case C
                    Case Is < 16
                        VarStr = VarStr & "$" & IIf(Len(Hex(B(i - Offset1))) < 2, "0" & Hex(B(i - Offset1)), Hex(B(i - Offset1))) & ", "
                    Case 16
                        C = 0
                        VarStr = VarStr & "$" & IIf(Len(Hex(B(i - Offset1))) < 2, "0" & Hex(B(i - Offset1)), Hex(B(i - Offset1))) & ", " & Chr(13) & Chr(10) & vbTab
                End Select
            Next
            If Right(VarStr, 2) = ", " Then
                VarStr = Left(VarStr, Len(VarStr) - 2) & Chr(13) & Chr(10) & ");"
            Else
                VarStr = Left(VarStr, Len(VarStr) - 5) & Chr(13) & Chr(10) & ");"
            End If
        End If
        Clipboard.Clear
        Clipboard.SetText VarStr, 1
    End If
    Seek #1, Ptr
    FillGrid
    Unload frmBlock
    Set frmBlock = Nothing
    MousePointer = 0
    
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyTab Then
        If FileOpened Then
            If ActiveGrid = 1 Then PicText.SetFocus Else PicHex.SetFocus
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If FindNext.Search Then
            MousePointer = 11
            Select Case FindNext.LastFind
                Case 0
                    ret = FindDown(FindNext.StrWhat, FindNext.FindType)
                Case 1
                    ret = FindUp(FindNext.StrWhat, FindNext.FindType)
            End Select
            If ret = -1 Then MsgBox Msgs(1), vbInformation, Caption
            MousePointer = 0
        End If
    End If

End Sub

Private Sub Form_Load()

    LoadLPK
    CenterForm Me
    Set grid1 = New Grid
    grid1.InitializeGrid PicHex, Shape1
    grid1.Rows = MAX_LINE
    Set grid2 = New Grid
    grid2.InitializeGrid PicText, Shape2
    grid2.Rows = MAX_LINE
    hg = Int(PicOffSet.Height / MAX_LINE)
    DisableControls
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If FileChanged Then
        If MsgBox(Chr(34) & FileName & Chr(34) & " " & Msgs(3), vbQuestion + vbYesNo, Caption) = vbYes Then
            Cancel = True
            cmdSave
        End If
    End If
    Set frmEditor = Nothing
    End

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Reset
    TempDir = GetTempDir
    If Right(TempDir, 1) <> "\" Then TempDir = TempDir & "\"
    If Dir(TempDir & "hexfiletemp.tmp") <> "" Then Kill TempDir & "hexfiletemp.tmp"
    cmdClose
    
End Sub
Private Sub grid1_EnterCell(ByVal Row As Integer, ByVal Col As Integer)
    
    If ActiveGrid = 1 Then
        If Sign(0).Left <> Label2(Col).Left + 6 Then Sign(0).Left = Label2(Col).Left + 6
        If Sign(1).Top <> (hg * Row) + (PicOffSet.Top + 2) Then Sign(1).Top = (hg * Row) + (PicOffSet.Top + 2)
        lblInfo(0) = grid1.Value(Row, Col)
        lblInfo(1) = GetValue(grid1.Value(Row, Col))
        lblInfo(2) = Hex((Ptr + (Row * 16) + Col) - 1)
        lblInfo(3) = (Ptr + (Row * 16) + Col) - 1
        grid2.SelectCell grid1.Row, grid1.Col
    End If

End Sub
Private Sub grid1_LeaveCell(OldRow As Integer, OldCol As Integer)
    
    If ActiveGrid = 1 Then
        If Len(grid1.Value(OldRow, OldCol)) < 2 Then
            grid1.Value(OldRow, OldCol) = "0" & grid1.Value(OldRow, OldCol)
            SaveData OldRow, OldCol
        End If
    End If
    
End Sub
Private Sub grid1_RequestBOF()

    MoveFirstPage
    
End Sub
Private Sub grid1_RequestEOF()

    MoveLastPage

End Sub
Private Sub grid1_RequestNextPage()

    MoveNextPage

End Sub
Private Sub grid1_RequestNextRow()

    If grid1.Rows = 20 And Ptr + (Rows * 16) < LOF(1) Then
        MoveNextLine
    End If

End Sub
Private Sub grid1_RequestPreviousPage()

    MovePreviousPage

End Sub
Private Sub grid1_RequestPreviousRow()

    If Ptr > 1 Then
        MovePreviousLine
    End If

End Sub
Private Sub grid2_RequestBOF()

    MoveFirstPage
    
End Sub
Private Sub grid2_RequestEOF()

    MoveLastPage

End Sub
Private Sub grid2_RequestNextPage()

    MoveNextPage

End Sub
Private Sub grid2_RequestNextRow()

    If grid1.Rows = 20 And Ptr + (Rows * 16) < LOF(1) Then
        MoveNextLine
    End If

End Sub
Private Sub grid2_RequestPreviousPage()

    MovePreviousPage

End Sub
Private Sub grid2_RequestPreviousRow()

    If Ptr > 1 Then
        MovePreviousLine
    End If

End Sub
Private Sub grid2_EnterCell(ByVal Row As Integer, ByVal Col As Integer)

    If ActiveGrid = 2 Then
        If Sign(0).Left <> Label2(grid2.Col).Left + 6 Then Sign(0).Left = Label2(grid2.Col).Left + 6
        If Sign(1).Top <> (hg * grid2.Row) + PicText.Top + 2 Then Sign(1).Top = (hg * grid2.Row) + PicText.Top + 2
        lblInfo(0) = grid1.Value(grid1.Row, grid1.Col)
        lblInfo(1) = GetValue(grid1.Value(grid1.Row, grid1.Col))
        lblInfo(2) = Hex((Ptr + (grid1.Row * 16) + grid1.Col) - 1)
        lblInfo(3) = (Ptr + (grid1.Row * 16) + grid1.Col) - 1
        grid1.SelectCell grid2.Row, grid2.Col
    End If
    
End Sub
Private Sub lblInfo_Click(Index As Integer)

    If Not FileOpened Then Exit Sub
    If lblInfo(6).ForeColor = &H4000& Then
        lblInfo(6).ForeColor = &HFF
        lblInfo(7).ForeColor = &H4000&
        If PicHex.Enabled Then PicHex.SetFocus
    Else
        lblInfo(6).ForeColor = &H4000&
        lblInfo(7).ForeColor = &HFF
        If PicText.Enabled Then PicText.SetFocus
    End If

End Sub
Private Sub lblSelect_Click()

    If Not FileOpened Then Exit Sub
    If lblSelect = "Offset - DEC" Then
        lblSelect = "Offset - HEX"
        ShowHex = True
    Else
        lblSelect = "Offset - DEC"
        ShowHex = False
    End If
    For i = 0 To grid1.Rows - 1
        tmpString = GetOffset((i * 16) + (grid1.Pointer - 1))
        TextOut PicOffSet.hdc, 3, hg * i + 3, tmpString, Len(tmpString)
    Next
    PicOffSet.Refresh
    For i = 0 To 15
        If ShowHex Then
            Label2(i) = "0" & Hex(i)
        Else
            Label2(i) = Format(i, "00")
        End If
    Next
    
End Sub
Private Sub PicHex_GotFocus()

    ActiveGrid = 1
    grid2.SelectCell grid1.Row, grid1.Col
    lblInfo(6).ForeColor = &HFF
    lblInfo(7).ForeColor = &H4000&

End Sub
Private Sub PicHex_KeyDown(KeyCode As Integer, Shift As Integer)

    UndoInProgress = (Shift = 2)
    
End Sub
Private Sub PicHex_KeyPress(KeyAscii As Integer)
    
    Dim txt As String
    If UndoInProgress Then Exit Sub
    txt = UCase(Chr(KeyAscii))
    Select Case txt
        Case "0" To "9", "A" To "F"
            If Len(grid1.Value(grid1.Row, grid1.Col)) < 2 Then
                grid1.Value(grid1.Row, grid1.Col) = grid1.Value(grid1.Row, grid1.Col) & txt
                SaveData grid1.Row, grid1.Col
                grid2.Value(grid1.Row, grid1.Col) = GetChar(GetValue(grid1.Value(grid1.Row, grid1.Col)))
            Else
                PicHex.Line (grid1.Col * (PicHex.Width / 16), grid1.Row * (PicHex.Height / MAX_LINE))-Step(20, 20), PicHex.BackColor, BF
                grid1.Value(grid1.Row, grid1.Col) = txt
                grid2.Value(grid1.Row, grid1.Col) = GetChar(GetValue(grid1.Value(grid1.Row, grid1.Col)))
            End If
        Case Chr(13)
            If Len(grid1.Value(grid1.Row, grid1.Col)) < 2 Then
                grid1.Value(grid1.Row, grid1.Col) = "0" & grid1.Value(grid1.Row, grid1.Col)
                SaveData grid1.Row, grid1.Col
            End If
        Case Else
            KeyAscii = 0
    End Select

End Sub
Private Sub PicHex_LostFocus()
    
    If Len(grid1.Value(grid1.Row, grid1.Col)) < 2 Then
        grid1.Value(grid1.Row, grid1.Col) = "0" & grid1.Value(grid1.Row, grid1.Col)
        SaveData grid1.Row, grid1.Col
    End If

End Sub
Private Sub PicScroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        Drag = ((Y >= Image1.Top) And (Y <= Image1.Top + Image1.Height))
        Y1 = Y - Image1.Top
    Else
        Drag = False
    End If

End Sub
Private Sub PicScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        If Drag Then
            t = Y - Y1
            If t < 0 Then t = 0
            If t > PicScroll.ScaleHeight - Image1.Height Then t = PicScroll.ScaleHeight - Image1.Height
            Image1.Top = t
        End If
    End If

End Sub
Private Sub PicScroll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Drag Then SetOffsetPosition

End Sub
Sub SetOffsetPosition()

    V = (Image1.Top * 100) / (PicScroll.ScaleHeight - Image1.Height)
    t = (Int(LOF(1) / 16) * 16) - (MAX_LINE * 16)
    If LOF(1) Mod 16 <> 0 Then t = t + 16
    Ptr = Int((V / 100) * t) + 1
    If Ptr < 1 Then Ptr = 1
    FillGrid
    lblInfo(0) = grid1.Value(grid1.Row, grid1.Col)
    lblInfo(1) = GetValue(grid1.Value(grid1.Row, grid1.Col))
    lblInfo(2) = Hex((Ptr + (grid1.Row * 16) + grid1.Col) - 1)
    lblInfo(3) = (Ptr + (grid1.Row * 16) + grid1.Col) - 1

End Sub
Sub SetScrollPosition()

    If Not Image1.Visible Then Exit Sub
    h = (PicScroll.ScaleHeight - Image1.Height)
    t = (Int(LOF(1) / 16) * 16) - (MAX_LINE * 16)
    t = (Ptr / t) * 100
    t = (t * h) / 100
    t = IIf(t > h, h, t)
    t = IIf(t < 0, 0, t)
    Image1.Top = t
    
End Sub
Private Sub PicText_GotFocus()

    ActiveGrid = 2
    grid1.SelectCell grid2.Row, grid2.Col
    If Sign(0).Left <> Label2(grid2.Col).Left + 6 Then Sign(0).Left = Label2(grid2.Col).Left + 6
    If Sign(1).Top <> (hg * grid2.Row) + PicText.Top + 2 Then Sign(1).Top = (hg * grid2.Row) + PicText.Top + 2
    lblInfo(6).ForeColor = &H4000&
    lblInfo(7).ForeColor = &HFF

End Sub
Private Sub PicText_KeyDown(KeyCode As Integer, Shift As Integer)

    UndoInProgress = (Shift = 2)

End Sub
Private Sub PicText_KeyPress(KeyAscii As Integer)

    If UndoInProgress Then Exit Sub
    If KeyAscii = 0 Then Exit Sub
    grid2.Value(grid2.Row, grid2.Col) = GetChar(KeyAscii)
    grid1.Value(grid2.Row, grid2.Col) = GetHex(KeyAscii)
    SaveData grid2.Row, grid2.Col

End Sub
Sub SaveData(ByVal R As Integer, ByVal C As Integer)

    Dim B As Byte, Offset As Long
    Offset = ((grid1.Pointer + (R * 16)) + C)
    Seek #1, Offset
    Get #1, , B
    Seek #1, Offset
    Put #1, , GetValue(grid1.Value(R, C))
    Seek #1, Ptr
    For i = 0 To UBound(Und) - 1
        Und(i) = Und(i + 1)
    Next
    nUndos = UBound(Und)
    Und(nUndos).Value = B
    Und(nUndos).Col = C
    Und(nUndos).Row = R
    Und(nUndos).Offset = Ptr
    FileChanged = True
    
End Sub
Sub DoUndos()
    
    Dim nUndo As Integer
    nUndo = UBound(Und)
    If Und(nUndo).Offset > 0 Then
        Seek #1, Und(nUndo).Offset + (((Und(nUndo).Row * 16)) + Und(nUndo).Col)
        Put #1, , Und(nUndo).Value
        If Ptr <> Und(nUndo).Offset Then
            Ptr = Und(nUndo).Offset
            FillGrid
        Else
            grid1.Value(Und(nUndo).Row, Und(nUndo).Col) = GetHex(Und(nUndo).Value)
            grid2.Value(Und(nUndo).Row, Und(nUndo).Col) = GetChar(Und(nUndo).Value)
        End If
        If ActiveGrid = 1 Then
            grid1.SelectCell Und(nUndo).Row, Und(nUndo).Col
        Else
            grid2.SelectCell Und(nUndo).Row, Und(nUndo).Col
        End If
        For i = nUndo To 1 Step -1
            Und(i) = Und(i - 1)
        Next
        Und(0).Offset = 0
    End If
    
End Sub
Private Sub sm1_Click(Index As Integer)

    Select Case Index
        Case 0: cmdOpen
        Case 1: cmdSave
        Case 2: cmdSaveAs
        Case 3: cmdProperties
        Case 5: cmdClose
    End Select

End Sub
Private Sub Sm2_Click(Index As Integer)
    
    Select Case Index
        Case 0: cmdCopy
        Case 1: cmdUndo
        Case 2: SelectBlock
        Case 3: cmdFindHex
        Case 4: cmdFindText
        Case 5: cmdReplaceHex
        Case 6: cmdReplaceText
        Case 10: cmdGoOffset
        Case 11: cmdGoPage
    End Select

End Sub
Private Sub Sm3_Click(Index As Integer)
    
    Select Case Index
        Case 0: cmdBookmark
        Case 1: cmdClearBookmark
        Case 2: cmdPreviousBookmark
        Case 3: cmdNextBookmark
    End Select

End Sub
Private Sub Sm4_Click(Index As Integer)
    
    Select Case Index
        Case 0: cmdConversion
        Case 1: cmdCharTable
        Case 2: cmdCalculator
        Case 4: cmdLanguage
        Case 5
            frmLPKEditor.Show 1
            Set frmLPKEditor = Nothing
    End Select

End Sub
Private Sub Sm5_Click()

    Load frmAbout
    frmAbout.Caption = Sm5.Caption
    frmAbout.Show 1
    
End Sub
