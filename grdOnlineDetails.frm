VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdOnlineDetailsfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ÊÞÑíÑ ãÊÇÈÚÉ ØáÈíÇÊ Ãæä áÇíä"
   ClientHeight    =   10290
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   10305
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   45
      Width           =   10050
      Begin VB.OptionButton xtype 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ãáÛíÉ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   495
         Width           =   1005
      End
      Begin VB.OptionButton xtype 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ãÓáãÉ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   2580
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   495
         Width           =   1185
      End
      Begin VB.OptionButton xtype 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ÛíÑ ãÓáãÉ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   1245
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   495
         Width           =   1185
      End
      Begin VB.OptionButton xtype 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Çáßá"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   495
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7740
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1320
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5310
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   330
         Left            =   5310
         TabIndex        =   2
         Top             =   540
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÍÊí"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6975
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   645
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ãä ÊÇÑíÎ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9135
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   225
         Width           =   660
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÇáÝÑÚ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   9135
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   585
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4500
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   315
      Width           =   5775
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "grdOnlineDetails.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1170
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
         _Version        =   196610
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "grdOnlineDetails.frx":2323
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdOnlineDetails.frx":4699
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   555
         Left            =   4635
         TabIndex        =   3
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "grdOnlineDetails.frx":681C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   2295
         TabIndex        =   5
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "grdOnlineDetails.frx":96CF
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSql 
         Height          =   555
         Left            =   3420
         TabIndex        =   4
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "grdOnlineDetails.frx":BBFA
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7170
      Left            =   90
      TabIndex        =   8
      Top             =   1080
      Width           =   20265
      _cx             =   35745
      _cy             =   12647
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   24
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   4410
      Top             =   -90
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   330
      Left            =   3420
      Top             =   -90
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   10095
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   540
      Top             =   180
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   3330
      Top             =   270
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   1890
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "grdOnlineDetailsfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearch As New Search_abd, oSearchCode As New Search_abd
Dim aHeader()
Private Sub CMD_EXEL_Click()
End Sub
Private Sub CMD_PRINT_Click()
End Sub
Private Sub cmdExcel_Click()
Me.MousePointer = 11

Dim aRow As Variant
aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aRow = AddFlag(aRow, aSub)


If grid1.Rows > 1 Then
    aSub = AddFlag(Empty, "row", 1)
    aSub = AddFlag(aSub, "bold", True)
    aSub = AddFlag(aSub, "word_wrap", False)
    aSub = AddFlag(aSub, "back_color", 40)
    aRow = AddFlag(aRow, aSub)
End If
ToFileExelNew grid1, , , aRow, , 1.2, , , , , , Me, Array(Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2), retHeader(aHeader, 4, 5))

Me.MousePointer = 0
fixGrd
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
Me.MousePointer = vbHourglass
myload
Me.MousePointer = vbNormal
End Sub

Private Sub cmdPdf_Click()
myPrint "1", True
End Sub
Private Sub cmdPrint_Click()
myPrint
End Sub

Private Sub cmdSql_Click()
myload True
End Sub
Private Sub Form_Resize()
grid1.Height = IIf(Me.Height - grid1.Top - 1000 < 3000, 3000, Me.Height - grid1.Top - 1000)
grid1.Width = IIf(Me.Width - grid1.Left - 150 < 5000, 5000, Me.Width - grid1.Left - 150)
Frame1.Left = Me.Width - Frame1.Width - 150
'Frame2.Left = IIf(300 + Me.Width - Frame2.Width - Frame1.Left < 0, 0, 300 + Me.Width - Frame2.Width - Frame1.Left)
Frame2.Left = Frame1.Left - Frame2.Width - 50
End Sub
Private Sub Form_Load()
openCon con


Set data1.Recordset = cmd("Select code ,descA From file0_40 ORDER BY CODE", con).Execute
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "CODE"

Set grid1.DataSource = DATA11

fixGrd
End Sub
Private Sub myload(Optional bString As Boolean = False)
With grid1
Dim aPrm As Variant

ReDim aHeader(2)
If xStore.MatchedWithList Then
    aPrm = AddFlag(aPrm, "STORE", addstring(xStore.BoundText))
    aHeader(0) = "ÇáÝÑÚ : " & xStore.text
End If

If IsDate(xDate1.text) Then
    aPrm = AddFlag(aPrm, "DATE1", addDate(xDate1.text))
    aHeader(1) = BetweenString(xDate1.text, xDate2.text)
End If

If IsDate(xDate1.text) Then
    aPrm = AddFlag(aPrm, "DATE2", addDate(xDate2.text))
    aHeader(1) = BetweenString(xDate1.text, xDate2.text)
End If

For i = 1 To xtype.UBound
    If xtype(i).Value Then
        aPrm = AddFlag(aPrm, "TYPE", i)
        aHeader(2) = "ØáÈíÇÊ " & xtype(i).Caption
    End If
Next

Dim cString As String
cString = myPrcString("[dbo].[sp_online_payment_detail]", aPrm)

If bString Then
    Clipboard.Clear
    Clipboard.SetText cString
    Exit Sub
End If

Set DATA11.Recordset = cmd(cString, con, , , 600).Execute
End With
fixGrd
End Sub
Sub fixGrd()
With grid1
    .RowHeight(0) = 600
    .WordWrap = True
    .FrozenCols = 4
    .TextMatrix(0, 0) = "ÇáÊÇÑíÎ"
    .TextMatrix(0, 1) = "ÑÞã ØáÈíÉ"
    .TextMatrix(0, 2) = "ÚÏÏ ÇáÞØÚ"
    .TextMatrix(0, 3) = "ÞíãÉ ÇáÇæÏÑ"
    .TextMatrix(0, 4) = "ÑÞã ÇáÝÇÊæÑÉ"
    .TextMatrix(0, 5) = "ÊÇÑíÎ ÇáÝÇÊæÑÉ"
    .TextMatrix(0, 6) = "ÚÏÏ ÇáÞØÚ"
    .TextMatrix(0, 7) = "ÞíãÉ ÇáÝÇÊæÑÉ"
    .TextMatrix(0, 8) = "ÈæáíÕÉ ÇáÔÍä"
    .TextMatrix(0, 9) = "ÔÑßÉ ÇáÔÍä"
    .TextMatrix(0, 10) = "ÊÇÑíÎ ÇáÔÍä"
    .TextMatrix(0, 11) = "ØÑíÞÉ ÇáÓÏÇÏ"

    .TextMatrix(0, 12) = "ÑÞã ãÓÊäÏ ÇáÊÍÕíá"
    .TextMatrix(0, 13) = "ÊÇÑíÎ ÇáÊÍÕíá"
    .TextMatrix(0, 14) = "ÞíãÉ ÇáÊÍÕíá"
    .TextMatrix(0, 15) = "ÊÇÑíÎ ÇáÇáÛÇÁ"
    
    .TextMatrix(0, 16) = "ÑÞã ÝÇÊæÑÉ ÇáãÑÊÌÚ"
    .TextMatrix(0, 17) = "ÚÏÏ ÞØÚ ÇáãÑÊÌÚ"
    .TextMatrix(0, 18) = "ÞíãÉ ÇáãÑÊÌÚ"
    .TextMatrix(0, 19) = "ÑÞã ÈæáíÕå ÇáãÑÊÌÚ"
    
    .TextMatrix(0, 20) = "ãÓÊäÏ ÊÍÕíá ÇáãÑÊÌÚ"
    .TextMatrix(0, 21) = "ÊÇÑíÎ ÓÏÇÏ ÇáãÑÊÌÚ"
    .TextMatrix(0, 22) = "ÞíãÉ ÓÏÇÏÇáãÑÊÌÚ"
    .TextMatrix(0, 23) = "ÅÓã ÇáãäÏæÈ"

    .ColDataType(1) = flexDTDouble
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColDataType(14) = flexDTDouble
    .ColDataType(17) = flexDTDouble
    .ColDataType(18) = flexDTDouble
    .ColDataType(22) = flexDTDouble
    
    .ColWidth(0) = 1300
    .ColWidth(1) = 1100
    .ColWidth(2) = 1100
    .ColWidth(3) = 1400
    .ColWidth(4) = 1600
    .ColWidth(5) = 1400
    .ColWidth(6) = 1100
    .ColWidth(7) = 1400
    .ColWidth(8) = 1300
    .ColWidth(9) = 1300
    .ColWidth(10) = 1300
    .ColWidth(11) = 2200
    .ColWidth(12) = 800
    .ColWidth(13) = 1300
    .ColWidth(14) = 1600
    .ColWidth(15) = 1300
    .ColWidth(16) = 1600
    .ColWidth(17) = 1000
    .ColWidth(18) = 1300
    .ColWidth(19) = 1300
    .ColWidth(20) = 800
    .ColWidth(21) = 1300
    .ColWidth(22) = 1300
    .ColWidth(23) = 1800
    
    .ColDataType(0) = flexDTDate
    .ColDataType(5) = flexDTDate
    .ColDataType(10) = flexDTDate
    .ColDataType(13) = flexDTDate
    .ColDataType(15) = flexDTDate
    .ColDataType(21) = flexDTDate
    
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignCenterCenter
        If .ColDataType(i) = flexDTDate Then
            .ColFormat(i) = "yyyy/m/d"
        End If
    Next
'    .MergeCol(1) = True
'    .MergeCol(2) = True
'    .MergeCol(3) = True
'    .MergeCol(4) = True
'    .MergeCol(5) = True
'    .MergeCol(6) = True
'    .MergeCol(7) = True
'    .MergeCol(8) = True
'    .MergeCol(9) = True
'
'    .MergeCells = flexMergeRestrictRows
'
'    '.ColComboList(0) = "..."
'    For col = 4 To 9
'    Next col
'
'
    .SubtotalPosition = flexSTAbove
'
    .Subtotal flexSTSum, -1, 2, "##,##", &HC0FFC0, vbBlack, True, "ÇáÅÌãÇáì"
    .Subtotal flexSTSum, -1, 3, "##,##", &HC0FFC0, vbBlack, True, "ÇáÅÌãÇáì"
    .Subtotal flexSTSum, -1, 6, "##,##", &HC0FFC0, vbBlack, True, "ÇáÅÌãÇáì"
    .Subtotal flexSTSum, -1, 7, "##,##", &HC0FFC0, vbBlack, True, "ÇáÅÌãÇáì"
    .Subtotal flexSTSum, -1, 14, "##,##", &HC0FFC0, vbBlack, True, "ÇáÅÌãÇáì"
    .Subtotal flexSTSum, -1, 17, "##,##", &HC0FFC0, vbBlack, True, "ÇáÅÌãÇáì"
    .Subtotal flexSTSum, -1, 18, "##,##", &HC0FFC0, vbBlack, True, "ÇáÅÌãÇáì"
    .Subtotal flexSTSum, -1, 22, "##,##", &HC0FFC0, vbBlack, True, "ÇáÅÌãÇáì"
    

''    .Subtotal flexSTSum, -1, 12 + 2, "##,##", &HC0FFC0, vbBlack, True, "ÇáÅÌãÇáì"
''    .Subtotal flexSTSum, -1, 13 + 2, "##,##", &HC0FFC0, vbBlack, True, "ÇáÅÌãÇáì"
''    .Subtotal flexSTSum, -1, 14 + 2, "##,##", &HC0FFC0, vbBlack, True, "ÇáÅÌãÇáì"
''    .Subtotal flexSTSum, -1, 17 + 2, "##,##", &HC0FFC0, vbBlack, True, "ÇáÅÌãÇáì"
''    .Subtotal flexSTSum, -1, 18 + 2, "##,##", &HC0FFC0, vbBlack, True, "ÇáÅÌãÇáì"
''    .Subtotal flexSTSum, -1, 19 + 2, "##,##", &HC0FFC0, vbBlack, True, "ÇáÅÌãÇáì"
'
'
'    For i = 0 To grid1.Cols - 1
'        .ColAlignment(i) = flexAlignRightCenter
'    Next
    
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdOrdersDeliveryfrm2 = Nothing
End Sub
Private Function MYVALID() As Boolean
MYVALID = True
End Function
Private Sub myPrint(Optional pDevice As String = "", Optional bIgPreview As Boolean = False)
If grid1.Rows < 3 Then Exit Sub
Dim aRow As Variant, aSub As Variant

aSub = AddFlag(Empty, "row", grid1.Rows - 1)
aSub = AddFlag(aSub, "col", 1)
aSub = AddFlag(aSub, "cols", 2)
aSub = AddFlag(aSub, "text", "ÅÌãÇáí")
aRow = AddFlag(aRow, aSub)

Set printGrdNew.myForm = Me
printGrdNew.doprint Me.grid1, 0.75, 0, "ÌæäíæÑ", Me.Caption, retHeader(aHeader, 0, 2), , True, True, 8, , aRow, Array(1)

If Not bIgPreview Then
    printGrdNew.Show 1
Else
    Unload printGrdNew
End If
End Sub

Private Sub xdesca_GotFocus()
myGotFocus xdesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xdesca
End Sub
Private Sub xModelFACT_GotFocus()
myGotFocus xModelFact
End Sub
Private Sub XModelFACT_LostFocus()
myLostFocus xModelFact
End Sub

Private Sub grid1_DblClick()
If grid1.Row < 2 Then Exit Sub
If grid1.col = 0 Then
    OrdersClientsfrm.sDoc_no = grid1.TextMatrix(grid1.Row, 0)
    OrdersClientsfrm.Show
Else
    grdOrdersDelivery_popfrm.sPO_NO = grid1.TextMatrix(grid1.Row, 0)
    grdOrdersDelivery_popfrm.Show
End If
End Sub

Private Sub grid1_EnterCell()
'grid1.Editable = IIf(grid1.Col = 0, flexEDKbdMouse, flexEDNone)
End Sub

Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP Me, oSearchCode
End Sub

Private Sub xPO_NO_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    OrdersLookup Me, oSearch
End If
End Sub
Private Sub xdate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub XPO_NO_GotFocus()
myGotFocus XPO_NO
End Sub
Private Sub xpo_no_LostFocus()
myLostFocus XPO_NO
End Sub
Private Sub xcode_GotFocus()
myGotFocus xcode
End Sub
Private Sub xcode_LostFocus()
myLostFocus xcode
If Not xcode.MatchedWithList Then xcode.BoundText = ""
End Sub
Public Sub myProc()
If ActiveControl.Name = xcode.Name Then
    xcode.BoundText = oSearchCode.grid1.TextMatrix(oSearchCode.grid1.Row, 0)
    oSearchCode.Hide
Else
    XPO_NO.text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    oSearch.Hide
End If
End Sub
