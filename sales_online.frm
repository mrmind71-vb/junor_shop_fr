VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form sales_online_incasefrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„»Ì⁄«  «Ê‰ ·«Ì‰"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20805
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   20805
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame fmDay 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   9945
      TabIndex        =   123
      Top             =   7695
      Width           =   1545
      Begin VB.CheckBox chkDayBranch 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÌÊ„Ì… «·›—⁄"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   270
         Value           =   1  'Checked
         Width           =   1230
      End
   End
   Begin VB.Frame fmAlert 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      TabIndex        =   120
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Image Image1 
         Height          =   450
         Left            =   90
         Picture         =   "sales_online.frx":0000
         Top             =   45
         Width           =   420
      End
      Begin VB.Label lblAlert 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " «—ÌŒ «·ÃÂ«“ „Œ ·› ⁄‰  «—ÌŒ «·„»Ì⁄« "
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   540
         TabIndex        =   121
         Top             =   90
         Width           =   4785
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   -7515
      TabIndex        =   52
      Top             =   9990
      Visible         =   0   'False
      Width           =   7620
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   945
         TabIndex        =   76
         Top             =   495
         Width           =   4065
         Begin Threed.SSCommand SSCommand2 
            CausesValidation=   0   'False
            Height          =   600
            Left            =   2205
            TabIndex        =   77
            Top             =   135
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   1058
            _Version        =   196610
            CaptionStyle    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "ÿ»«⁄… ≈Ã„«·Ï «·ÌÊ„"
            PictureAlignment=   3
         End
         Begin Threed.SSCommand cmd_tsal1 
            CausesValidation=   0   'False
            Height          =   600
            Left            =   270
            TabIndex        =   78
            Top             =   45
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1058
            _Version        =   196610
            CaptionStyle    =   1
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arabic Transparent"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "sales_online.frx":2877
            Alignment       =   1
            PictureAlignment=   3
         End
      End
      Begin VB.Frame Frame10 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   0
         TabIndex        =   67
         Top             =   2160
         Width           =   4695
         Begin Threed.SSCommand LOOKPH 
            CausesValidation=   0   'False
            Height          =   600
            Left            =   90
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   135
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   1058
            _Version        =   196610
            CaptionStyle    =   1
            ForeColor       =   0
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
            Picture         =   "sales_online.frx":542E
            Caption         =   "≈” ⁄·«„ ’Ê—"
            Alignment       =   1
            ButtonStyle     =   1
            PictureAlignment=   10
            BevelWidth      =   10
            ShapeSize       =   1
         End
         Begin Threed.SSCommand cmdOpen 
            CausesValidation=   0   'False
            Height          =   600
            Left            =   3420
            TabIndex        =   69
            Top             =   135
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   1058
            _Version        =   196610
            CaptionStyle    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "«” ⁄·«„ »Ê‰«  „› ÊÕ…"
            PictureAlignment=   3
         End
         Begin Threed.SSCommand cmdinformbarcode 
            CausesValidation=   0   'False
            Height          =   600
            Left            =   2295
            TabIndex        =   70
            Top             =   135
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1058
            _Version        =   196610
            CaptionStyle    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "«” ⁄·«„ „»Ì⁄«  ’‰›"
            PictureAlignment=   3
         End
         Begin Threed.SSCommand LOOKPH_BARCODE 
            CausesValidation=   0   'False
            Height          =   600
            Left            =   1350
            TabIndex        =   71
            Top             =   135
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   1058
            _Version        =   196610
            CaptionStyle    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "≈” ⁄·«„ »«—ﬂÊœ"
            PictureAlignment=   3
         End
      End
      Begin VB.CheckBox chkCash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»Ì⁄ ‰ﬁœÌ"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5715
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   2250
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CheckBox XISRETS 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5310
         TabIndex        =   62
         Top             =   1755
         Width           =   195
      End
      Begin VB.CheckBox XISONEST 
         Caption         =   "Check1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6525
         TabIndex        =   61
         Top             =   1575
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox xinv_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   855
         MaxLength       =   6
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   315
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CheckBox xPrinted 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   585
         TabIndex        =   53
         Top             =   1710
         Visible         =   0   'False
         Width           =   195
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   405
         TabIndex        =   64
         Top             =   1260
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdShip 
         Height          =   555
         Left            =   5625
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   225
         Width           =   1770
         _ExtentX        =   3122
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
         Picture         =   "sales_online.frx":798C
         Caption         =   " ”ÃÌ· «·‘Õ‰"
         ButtonStyle     =   2
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "sales_online.frx":A009
      End
      Begin Threed.SSCommand CMD_ONLINE 
         Height          =   390
         Left            =   180
         TabIndex        =   106
         Top             =   675
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   688
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "»Ê‰ «Ê‰ ·«Ì‰  ”·Ì„"
         ButtonStyle     =   2
      End
      Begin VB.Label xtime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1350
         TabIndex        =   72
         Top             =   1260
         Width           =   780
      End
      Begin VB.Label xIpName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   2160
         Width           =   2025
      End
      Begin VB.Label XBRANCH 
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2205
         TabIndex        =   65
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame fmMonth 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4815
      TabIndex        =   101
      Top             =   7695
      Width           =   5100
      Begin VB.CheckBox chkDay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "›Ê« Ì— «·ÌÊ„"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   225
         Width           =   1230
      End
      Begin MSDataListLib.DataCombo xYear 
         Height          =   330
         Left            =   3105
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo xMonth 
         Height          =   330
         Left            =   1530
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   180
         Width           =   870
         _ExtentX        =   1535
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
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·‘Â—"
         Height          =   330
         Left            =   2385
         TabIndex        =   105
         Tag             =   "Color"
         Top             =   225
         Width           =   555
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·”‰…"
         Height          =   285
         Left            =   4410
         TabIndex        =   104
         Tag             =   "Color"
         Top             =   225
         Width           =   510
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   11115
      RightToLeft     =   -1  'True
      TabIndex        =   97
      Top             =   -45
      Width           =   3390
      Begin Threed.SSCommand cmd_tsal 
         Height          =   510
         Left            =   1710
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   135
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   900
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
         Picture         =   "sales_online.frx":C4B5
         Caption         =   " „»Ì⁄«  «·ÌÊ„ "
         Alignment       =   4
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_online.frx":EB87
      End
      Begin Threed.SSCommand cmdPrintDay 
         Height          =   510
         Left            =   45
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   135
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   900
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
         Picture         =   "sales_online.frx":1106B
         Caption         =   " «€·«ﬁ «·ÌÊ„ "
         Alignment       =   4
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_online.frx":13701
      End
   End
   Begin VB.Frame fmClose 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   8235
      TabIndex        =   86
      Top             =   1665
      Width           =   1500
      Begin VB.CheckBox xClosed 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "„” ‰œ „€·ﬁ"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   87
         Top             =   225
         Width           =   1230
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8010
      RightToLeft     =   -1  'True
      TabIndex        =   81
      Top             =   7110
      Width           =   3480
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2610
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   135
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
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
         Picture         =   "sales_online.frx":15BE5
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "sales_online.frx":17D8C
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1710
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   135
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   741
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
         Picture         =   "sales_online.frx":19DD3
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "sales_online.frx":1BEBE
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   855
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   135
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
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
         Picture         =   "sales_online.frx":1DEB8
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "sales_online.frx":1FFC9
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
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
         Picture         =   "sales_online.frx":21FC3
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "sales_online.frx":241E7
      End
   End
   Begin VB.Frame FRAME_CUR 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Index           =   4
      Left            =   9765
      RightToLeft     =   -1  'True
      TabIndex        =   73
      Top             =   1080
      Width           =   1365
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   45
         TabIndex        =   74
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
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
         Picture         =   "sales_online.frx":262B8
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_online.frx":28BDD
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   45
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   675
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
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
         Picture         =   "sales_online.frx":2B431
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_online.frx":2D591
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   14535
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   -45
      Width           =   5775
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   4635
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
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
         Picture         =   "sales_online.frx":2F87E
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_online.frx":31C49
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   3510
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
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
         Picture         =   "sales_online.frx":33CF2
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_online.frx":35CFA
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   2340
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   135
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   900
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
         Picture         =   "sales_online.frx":37CB1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_online.frx":3A44D
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
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
         Picture         =   "sales_online.frx":3C8E1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   510
         Left            =   1170
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   135
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   900
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
         Picture         =   "sales_online.frx":3EC04
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "sales_online.frx":40F7A
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   15705
      TabIndex        =   36
      Top             =   5760
      Width           =   4560
      Begin VB.TextBox xTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00292952&
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   50
         Top             =   2160
         Width           =   2310
      End
      Begin VB.TextBox xDiscount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   495
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   49
         Top             =   1800
         Width           =   1050
      End
      Begin VB.TextBox xRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   1575
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   47
         Top             =   1800
         Width           =   825
      End
      Begin Threed.SSCommand cmd_disc 
         Height          =   330
         Left            =   90
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   1800
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "%"
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "’«›Ì «·›« Ê—…"
         Height          =   270
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   2205
         Width           =   1095
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’„ ›« Ê—…"
         Height          =   270
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1845
         Width           =   900
      End
      Begin VB.Label xTotalItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   90
         TabIndex        =   46
         Top             =   1395
         Width           =   2310
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï ›« Ê—…"
         Height          =   270
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label xTotalItemDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   90
         TabIndex        =   44
         Top             =   990
         Width           =   2310
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’„ √’‰«›"
         Height          =   270
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label xTotalItemNoDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   90
         TabIndex        =   42
         Top             =   585
         Width           =   2310
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï «·ﬁÌ„…"
         Height          =   270
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label xTotalQuant2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "„— Ã⁄ "
         Height          =   270
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   225
         Width           =   525
      End
      Begin VB.Label xTotalQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1710
         TabIndex        =   38
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ ﬁÿ⁄"
         Height          =   270
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   225
         Width           =   630
      End
   End
   Begin VB.Frame fmRc 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4230
      TabIndex        =   27
      Top             =   1665
      Width           =   3975
      Begin VB.CheckBox xSendRc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " „"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   180
         TabIndex        =   30
         Top             =   225
         Width           =   465
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   330
         Left            =   2565
         TabIndex        =   28
         Top             =   180
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Ì’«· «·Ìﬂ —Ê‰Ì"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin VB.Label xUUID_RC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   180
         Width           =   1785
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4500
      TabIndex        =   25
      Top             =   5760
      Width           =   1545
      Begin VB.CheckBox chkprint 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·€«¡ ÿ»«⁄…"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   26
         Top             =   180
         Value           =   1  'Checked
         Width           =   1275
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   3390
      Left            =   135
      TabIndex        =   9
      Top             =   2340
      Width           =   20130
      _cx             =   35507
      _cy             =   5980
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
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   22
      FixedRows       =   1
      FixedCols       =   1
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame fmDate 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   360
      TabIndex        =   19
      Top             =   7695
      Width           =   5010
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   720
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   126
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2205
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   1455
      End
      Begin Threed.SSCommand DAYALL 
         Height          =   375
         Left            =   90
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   180
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   661
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
         Picture         =   "sales_online.frx":430FD
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_online.frx":452FF
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ «·›« Ê—…"
         Height          =   285
         Left            =   3780
         TabIndex        =   108
         Tag             =   "Color"
         Top             =   225
         Width           =   1050
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -3240
      Top             =   6660
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   -855
      Top             =   6885
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin Crystal.CrystalReport REPORT1 
      Left            =   135
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowProgressCtls=   0   'False
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   -4770
      Top             =   8145
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   -3060
      Top             =   6975
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc data5 
      Height          =   330
      Left            =   -1755
      Top             =   6885
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
      Left            =   3510
      Top             =   3465
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin Crystal.CrystalReport Report_T 
      Left            =   4410
      Top             =   3195
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data12 
      Height          =   330
      Left            =   5400
      Top             =   810
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc DATA7 
      Height          =   330
      Left            =   1395
      Top             =   630
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc data8 
      Height          =   330
      Left            =   -630
      Top             =   1395
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   11160
      TabIndex        =   11
      Top             =   630
      Width           =   9150
      Begin VB.TextBox xdoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4545
         Locked          =   -1  'True
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   2310
      End
      Begin VB.TextBox xDoc_no2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6885
         Locked          =   -1  'True
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1185
      End
      Begin VB.TextBox xCode2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5355
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   2460
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4545
         MaxLength       =   200
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3525
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   6885
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   540
         Width           =   1185
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   330
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   2355
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   330
         Left            =   90
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   540
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo xBox 
         Height          =   330
         Left            =   90
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   900
         Width           =   2355
         _ExtentX        =   4154
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
      Begin MSDataListLib.DataCombo xMan 
         Height          =   330
         Left            =   4545
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   900
         Width           =   3525
         _ExtentX        =   6218
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
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "«·»«∆⁄"
         Height          =   270
         Left            =   8190
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   900
         Width           =   405
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "—ﬁ„ „” ‰œ"
         Height          =   285
         Index           =   0
         Left            =   8145
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   885
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "⁄„·«¡ ﬂ—Ê  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   2430
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "„·«ÕŸ« "
         Height          =   270
         Left            =   8145
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1305
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·Œ“‰…"
         Height          =   270
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   990
         Width           =   480
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4545
         TabIndex        =   3
         Top             =   540
         Width           =   2310
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
         Height          =   270
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   225
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„Œ“‰"
         Height          =   270
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   585
         Width           =   540
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "«·⁄„Ì·"
         Height          =   270
         Left            =   8145
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   585
         Width           =   450
      End
   End
   Begin MSAdodcLib.Adodc DATA6 
      Height          =   330
      Left            =   2520
      Top             =   585
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin VB.Frame Frame14 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   11520
      TabIndex        =   31
      Top             =   5760
      Width           =   4155
      Begin VB.TextBox xdate_Delivery 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   765
         MaxLength       =   10
         TabIndex        =   127
         Top             =   2025
         Width           =   2040
      End
      Begin VB.TextBox xship_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   765
         MaxLength       =   10
         TabIndex        =   34
         Top             =   945
         Width           =   2040
      End
      Begin MSDataListLib.DataCombo xship 
         Height          =   360
         Left            =   765
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   540
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdSaveLading 
         Height          =   510
         Left            =   90
         TabIndex        =   135
         Top             =   1890
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   900
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
         Picture         =   "sales_online.frx":4741B
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_online.frx":49D40
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·»Ê·’…"
         Height          =   270
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   130
         Top             =   1710
         Width           =   960
      End
      Begin VB.Label xlading_type 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   765
         TabIndex        =   129
         Top             =   1665
         Width           =   2040
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «· ”·Ì„"
         Height          =   270
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   128
         Top             =   2070
         Width           =   960
      End
      Begin VB.Label xLading 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   765
         TabIndex        =   115
         Top             =   1305
         Width           =   2040
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "»Ê·Ì’… „”»ﬁ…"
         Height          =   270
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   114
         Top             =   1350
         Width           =   1065
      End
      Begin VB.Label xonline_doc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   765
         TabIndex        =   80
         Top             =   180
         Width           =   2040
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·ÿ·»Ì…"
         Height          =   270
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   225
         Width           =   810
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·»Ê·Ì’…"
         Height          =   270
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "‘—ﬂ… «·‘Õ‰"
         Height          =   270
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   585
         Width           =   945
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   88
      Top             =   9630
      Width           =   20805
      _ExtentX        =   36698
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   89
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   714
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   1
         Left            =   4095
         TabIndex        =   90
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   2
         Left            =   8100
         TabIndex        =   91
         Top             =   45
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   3
         Left            =   12150
         TabIndex        =   92
         Top             =   45
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   4
         Left            =   16155
         TabIndex        =   93
         Top             =   45
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   105
      Left            =   0
      TabIndex        =   94
      Top             =   9525
      Visible         =   0   'False
      Width           =   20805
      _ExtentX        =   36698
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame99 
      BackColor       =   &H00FFFFFF&
      Height          =   1410
      Left            =   6075
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   5715
      Width           =   5415
      Begin VB.TextBox xdoc_ret 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   111
         Top             =   225
         Width           =   1950
      End
      Begin VB.CheckBox xIsRet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "”„«Õ »„— Ã⁄"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   225
         TabIndex        =   98
         Top             =   855
         Width           =   1455
      End
      Begin Threed.SSCommand cmd_showdoc 
         Height          =   465
         Left            =   90
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   225
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   820
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
         Picture         =   "sales_online.frx":4C594
         Caption         =   "⁄—÷ «·„” ‰œ"
         Alignment       =   4
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_online.frx":4EBA0
      End
      Begin VB.Label XSALES_RET 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1800
         TabIndex        =   113
         Top             =   630
         Width           =   1950
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "«œŒ«· —ﬁ„ «·„— Ã⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3795
         TabIndex        =   112
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ „” ‰œ «·„— Ã⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3825
         TabIndex        =   109
         Top             =   675
         Width           =   1485
      End
      Begin VB.Label xusername_RET 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   1800
         TabIndex        =   24
         Top             =   990
         Width           =   1950
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„” Œœ„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3825
         TabIndex        =   23
         Top             =   1035
         Width           =   1080
      End
   End
   Begin MSAdodcLib.Adodc data20 
      Height          =   330
      Left            =   5850
      Top             =   1125
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
      Caption         =   "data20"
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
   Begin MSAdodcLib.Adodc data21 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
      Caption         =   "data20"
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
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   15120
      TabIndex        =   116
      Top             =   8280
      Width           =   5145
      Begin VB.OptionButton optclosed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "›Ê« Ì— „€·ﬁ…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   180
         TabIndex        =   119
         Top             =   270
         Width           =   1500
      End
      Begin VB.OptionButton optclosed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "›Ê« Ì— €Ì— „€·ﬁ…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   1755
         TabIndex        =   118
         Top             =   270
         Width           =   1635
      End
      Begin VB.OptionButton optclosed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ﬂ· «·›Ê« Ì—"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3870
         TabIndex        =   117
         Top             =   270
         Value           =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4050
      TabIndex        =   131
      Top             =   8280
      Width           =   11040
      Begin VB.OptionButton optLading 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " »œÌ· Ê„— Ã⁄ »›« Ê—… €Ì— „€·ﬁ"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   2655
         TabIndex        =   137
         Top             =   225
         Width           =   2850
      End
      Begin VB.OptionButton optLading 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " »œÌ· Ê„— Ã⁄ »›« Ê—… „€·ﬁ"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   90
         TabIndex        =   136
         Top             =   225
         Width           =   3255
      End
      Begin VB.OptionButton optLading 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ﬂ· «·›Ê« Ì—"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   9720
         TabIndex        =   134
         Top             =   225
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optLading 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "›Ê« Ì— »œÊ‰  «—ÌŒ  ”·Ì„"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   7335
         TabIndex        =   133
         Top             =   225
         Width           =   2220
      End
      Begin VB.OptionButton optLading 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "›Ê« Ì— » «—ÌŒ  ”·Ì„"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   5490
         TabIndex        =   132
         Top             =   225
         Width           =   1770
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   " «—ÌŒ «·„” ‰œ"
      Height          =   330
      Left            =   0
      TabIndex        =   107
      Top             =   0
      Width           =   1050
   End
End
Attribute VB_Name = "sales_online_incasefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temptable As New ADODB.Recordset, cMsgDisplay As String, nDayDiscout As Double, cDescDiscout As String
Dim nCountPrint As Double
Dim bStopCell As Boolean
Dim bSuperSales As Boolean
Dim bEditRecord As Boolean
Public sDoc_no As String, lSave As Boolean, sUserDisc As String, lIsPrice_2 As Boolean, lIsdisc2   As Boolean, lDiscModelRet As Boolean
Dim cFilter As String, lIsDocRet As Boolean, cFilterLook As String
Dim bAct As Boolean
Dim sDateSales As String
Dim computerName As String
Dim tCard As ADODB.Recordset, lDiscOpenDay As Boolean
Dim nRound As Long, cList1 As String, lIsBoxOnline As Boolean
Dim CardTable As ADODB.Recordset, cId_NoDesca1 As String, cId_NoDesca2 As String
Dim cDefClient As String, cDefClientDesca As String
Dim dDiscBranch1 As Date, dDiscBranch2 As Date, nDiscBranch As Double
Dim oSearchItem As New Search3, oSearchDoc  As New Search_abd, oSearchCl1 As New Search3, bMarket As Boolean
Dim oSearchCl2 As New Search3, oSearchCl3 As New Search3
Dim cQrCode As New ClsQrCode
Dim bEdit As Boolean, lISNOGRDISC As Boolean
Dim bIgClick As Boolean
Dim bIgFilter As Boolean
Dim cFile As String, cCodeDesca As String
Dim con As New ADODB.Connection
Dim formMode, dDateLast As String
Public myPublic As Integer, lNotBranch As Boolean
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup()
    ItemsLookupAll Me, oSearchItem
End Sub
Private Function myreplace(Optional Row As Long = -1, Optional sPrinted As String = "", Optional sClosed As String = "") As Boolean
Dim aInsert As Variant
aInsert = AddFlag(aInsert, "CODE", addstring(xCode.text))
aInsert = AddFlag(aInsert, "STORE", addstring(xStore.BoundText))
aInsert = AddFlag(aInsert, "NOTES", addstring(xNotes.text))
aInsert = AddFlag(aInsert, "RATE", Val(xRate.text))
aInsert = AddFlag(aInsert, "DISCOUNT", Val(xDiscount.text))
aInsert = AddFlag(aInsert, "BOX", addstring(xBox.BoundText))
aInsert = AddFlag(aInsert, "BRANCH", addstring(XBRANCH.Caption))
aInsert = AddFlag(aInsert, "MOSM", addstring(xMosm.BoundColumn))
aInsert = AddFlag(aInsert, "[TIME]", addTime(xtime.Caption))
aInsert = AddFlag(aInsert, "VISA", "0")
aInsert = AddFlag(aInsert, "MAN", addstring(IIf(xMan.MatchedWithList, xMan.BoundText, "00")))
aInsert = AddFlag(aInsert, "[ISNEW]", "1")
aInsert = AddFlag(aInsert, "username_ret", addstring(xusername_RET.Caption))
aInsert = AddFlag(aInsert, "SALES_RET", addstring(XSALES_RET.Caption))
aInsert = AddFlag(aInsert, "ONLINE", "2")
aInsert = AddFlag(aInsert, "INV_NO", addstring(xinv_no.text))
aInsert = AddFlag(aInsert, "USER_IP", addstring(xIpName.Caption))
aInsert = AddFlag(aInsert, "SHIP_NO", addstring(xship_no.text))
aInsert = AddFlag(aInsert, "SHIP", addstring(xship.BoundText))

If sClosed <> "" Then
    aInsert = AddFlag(aInsert, "ISCLOSED", sClosed)
End If

If sPrinted <> "" Then
    aInsert = AddFlag(aInsert, "PRINTED", sPrinted)
End If



con.BeginTrans
On Error GoTo myerror
If xdoc_no.Tag = DefineMode Then
    checkDate
    
    xdoc_no.text = NewflagDoc(myFormat(xDate.text), xBox.BoundText, con)
    xDoc_no2.text = Mid(xdoc_no.text, 7, 6)
    
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xdoc_no.text))
    aInsert = AddFlag(aInsert, "DATE", addDate(xDate.text))
    
    con.Execute addInsert(aInsert, "FILE6_20H")
Else
    aInsert = AddFlag(aInsert, "USERNAME", addstring(cusername))
    con.Execute addUpdate(aInsert, "FILE6_20H", "DOC_NO = " & addstring(xdoc_no.text))
End If
myreplaceGrd Row
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub checkDate()
'Dim sDateSales  As String
'sDateSales = myFormat_p(myField("SELECT DSALES FROM DSALES WHERE BRANCH = " & MyParn(cBranch), con))
'If sDateSales <> myFormat(xDate.text) Then
'    MsgBox " „  €Ì— «· «—ÌŒ"
'End If
'xDate.text = myFormat_p(sDateSales)
End Sub
Sub myProc()
'On Error GoTo myerror
If ActiveControl.Name = grid1.Name Or ActiveControl.Name = LOOKPH.Name Or ActiveControl.Name = LOOKPH_BARCODE.Name Then
    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    If ActiveControl.Name = grid1.Name Then
        grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    End If
    GrdDesc grid1.TextMatrix(grid1.Row, 1), grid1.Row
    Grid1_AfterEdit grid1.Row, grid1.Col
ElseIf ActiveControl.Name = cmdInform.Name Or ActiveControl.Name = cmdOpen.Name Then
    openCardTable tbMode.tbFind, oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub chkDay_Click()
If chkDayBranch.Value = 1 Then Exit Sub
myUndo
End Sub

Private Sub chkDayBranch_Click()
If IsDate(xdate1.text) Or IsDate(xdate2.text) Then Exit Sub
myUndo
End Sub

Private Sub chkprint_Click()
    addSetting "print", chkprint.Value, App.Path & "\other.txt"
End Sub

Private Sub cmd_closed_Click()

End Sub

Private Sub CMD_ISRET_Click()
End Sub
Private Sub CMD_ONLINE_Click()
On Error GoTo myerror
    lSave = False
    If grid1.Rows = 2 Then Exit Sub

    con.Execute " UPDATE FILE6_20H SET DISCOUNT = " & Val(xDiscount.text) & "  , RATE = " & Val(xRate.text) & " , username_disc  = " & addstring(XUSERNAME_DISC.Caption) & " , DOC_disc  = " & addstring(XDOC_DISC.Caption) & " , ISNEW =  1  WHERE DOC_NO = " & MyParn(xdoc_no.text)
    
    Set CashOnLine.myform = Me
    CashOnLine.Show 1
    
    
    myUndo

    If xonline_doc.Caption <> "" Then
        Dim con_SERVER As New Connection
        If openCon(con_SERVER, LoadConString_OnLine) = "ok" Then
            con_SERVER.Execute " UPDATE FILE6_90H SET SHIP_NO = " & addstring(xship_no.text) & " WHERE doc_no = " & MyParn(xonline_doc.Caption), nRec
        End If
        con_SERVER.Close
    End If

'    If Not lSave Then Exit Sub
'
'    If doprint(False) Then SavePrint
'    DoEvents
'    CmdNewInv_Click
'    On Error Resume Next
'    XMAN.SetFocus
'    Err.Clear
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdSendSales_Click()
reSendSales
End Sub
Private Sub cmdPrintDay_Click()
If cBranchStore = "" Then
    MsgBox "·« ÌÊÃœ ›—⁄"
    Exit Sub
End If

Dim bOnline As Variant
bOnline = myField("SELECT online from file0_40 where code = " & MyParn(cBranchStore), con)

If IsEmpty(bOnline) Then
    MsgBox "«·›—⁄ €Ì— ’ÕÌÕ"
    Exit Sub
ElseIf Not bOnline Then
    MsgBox "«·›—⁄ ·Ì” ›—⁄ «Ê‰ ·«Ì‰"
    Exit Sub
End If

If myFormat(Date) = myFormat(sDateSales) Then
    MsgBox " «—ÌŒ «·ÃÂ«“ »‰›”  «—ÌŒ «·„»Ì⁄« "
    Exit Sub
End If

If myFormat(Date) < myFormat(sDateSales) Then
    MsgBox " «—ÌŒ «·ÃÂ«“ «ﬁ· „‰  «—ÌŒ «·„»Ì⁄« "
    Exit Sub
End If

If cBranch <> "" Then
    Dim cString As String
    cString = "SELECT COUNT(*) " & _
              " FROM FILE6_20H " & _
              " WHERE BRANCH = " & MyParn(cBranch) & _
              " AND (PRINTED = 0 OR ISCLOSED = 0)" & _
              " AND DATE = " & addDate(sDateSales)
    
    Dim nCount As Long
    nCount = myField(cString, con, , , 0)
    If nCount > 0 Then
        MsgBox nCount & " ›Ê« Ì— „»Ì⁄«  " & "€Ì— „€·ﬁ…"
    Else
        Dim nAffect As Long
        con.Execute "UPDATE DSALES " & _
                    " SET DSALES = " & DateSq(Date) & _
                    " WHERE BRANCH = " & MyParn(cBranch), nAffect
        If nAffect = 1 Then MsgBox " „ «€·«ﬁ «·ÌÊ„ »‰Ã«Õ"
    End If
End If
fixDateSales
If Not openCardTable Then myDefine
End Sub
Private Sub cmdSaveLading_Click()

If MsgBox("  ⁄œÌ·  «—ÌŒ «· ”·Ì„", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub

Dim nAffect As Long
On Error GoTo myerror
con.Execute "UPDATE file6_20h set date_delivery = " & addDate(xdate_Delivery.text) & _
            " WHERE DOC_NO = " & addstring(xdoc_no.text), nAffect
Finally:
myUndo
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
GoTo Finally
End Sub

Private Sub cmdShip_Click()
On Error GoTo myerror
Dim cAddre As String
If ((xClosed.Value <> 0 Or lMainShow Or lSupperVisor) And cBranch <> "00") Then
    On erorr GoTo myerror
    xonline_doc.Caption = InputBox(" —ﬁ„ «Ê—œ— «Ê‰ ·«Ì‰ ", , xonline_doc.Caption)
    
    If MsgBox("  ⁄œÌ· —ﬁ„ «Ê—œ— «Ê‰ ·«Ì‰ ··›« Ê—… ", vbYesNo + vbDefaultButton2) = vbYes Then
        Dim con_SERVER As New Connection
        If openCon(con_SERVER, LoadConString_OnLine) = "ok" Then
            con_SERVER.Execute " UPDATE FILE6_90H SET SHIP_NO = " & addstring(xship_no.Caption) & " , sales_doc = " & addstring(xdoc_no.text) & " , sales_date = " & addDate(xDate.text) & " WHERE doc_no = " & MyParn(xonline_doc.Caption), nRec
        
            If nRec > 0 Then
                aRet = aGetDesca("select name , Shipping_City , STREET , phone from file6_90h where doc_no = " & MyParn(xonline_doc.Caption), con_SERVER)
                If UBound(aRet) > 0 Then
                    cAddre = aRet(2) + "   " + aRet(3)
                    con.Execute " UPDATE  FILE6_20H SET file6_20h.NAME = " & addstring(aRet(1)) & " , file6_20h.ADDRESS = " & addstring(cAddre) & " ,  file6_20h.PHONE = " & addstring(aRet(4)) & "   WHERE file6_20h.doc_no = " & MyParn(xdoc_no.text)
                End If
            
            Else
                MsgBox " —ﬁ„ ÿ·»Ì… €Ì— ’ÕÌÕ "
            End If
        End If
        con_SERVER.Close
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear

con.Execute "UPDATE FILE6_20H " & _
            "SET FILE6_20H.SHIP = " & addstring(xship.BoundText) & _
            " AND FILE6_20H.SHIP_NO = " & addstring(xship_no.text) & _
            " WHERE FILE6_20H.DOC_NO = " & MyParn(xdoc_no.text)
Inform
End Sub

Private Sub CND_PRICE3_Click()
    If InputBox("", "") = "987654321" Then
        con.Execute " UPDATE FILE6_20 SET  file6_20.price = file1_10.PRICE , file6_20.price_C = file1_10.PRICE  FROM   FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM where DOC_NO = " & MyParn(xdoc_no.text)
        myLoadGrd
    End If

End Sub

Private Sub Command3_Click()
    If InputBox("", "") = "987654321" Then
        con.Execute " UPDATE FILE6_20 SET  file6_20.price_C = file1_10.PRICE  FROM   FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM where DOC_NO = " & MyParn(xdoc_no.text)
        myLoadGrd
    End If
End Sub

Private Sub Command4_Click()
    If InputBox("", "") = "987654321" Then
        con.Execute " UPDATE FILE6_20 SET  file6_20.price_C = file1_10.PRICE_2021   FROM   FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM where DOC_NO = " & MyParn(xdoc_no.text)
        con.Execute " UPDATE FILE6_20 SET  file6_20.price_C = file1_10.PRICE        FROM   FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM where ( file6_20.price_C =0 0 OR file6_20.price_C IS NULL ) AND DOC_NO = " & MyParn(xdoc_no.text)
        myLoadGrd
    End If
End Sub
Private Sub Command5_Click()
Dim loctable As New ADODB.Recordset
Dim sb As New ChilkatStringBuilder
Dim nAffect As Integer

Dim con_SERVER As New Connection
If openCon(con_SERVER, LoadConString_OnLine) <> "ok" Then
    Exit Sub
End If
Set loctable = myCmd("SELECT  * FROM FILE6_20h where branch = '09'  and codevisa is not null ", con_SERVER, , , 5000)

Dim i As Long, nRecord As Long, nRecordCount As Long, nCount As Long
Dim nUpdated As Long
Dim sCaption As String

prog1.Visible = True
prog1.Value = 0
nRecordCount = loctable.RecordCount
sCaption = Me.Caption

'Inform "÷»ÿ „·› —ﬁ„ " & pDoc_No & " ⁄œœ ”Ã·«  " & nRecordCount

Do Until loctable.EOF
    nRecord = nRecord + 1
    i = i + 1
    
    Me.Caption = sCaption & " - " & " ”Ã· " & nRecord & " „‰ " & nRecordCount
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    cString = "UPDATE FILE6_20H SET" & _
              " FILE6_20H.codevisa = " & loctable!CODEVISA & _
              " FROM FILE6_20H " & _
              " WHERE FILE6_20H.DOC_NO = " & MyParn(loctable!doc_no)
    sb.Append cString & ";"
    
    If i >= 500 Then
        con.Execute sb.GetAsString, nAffect
        nUpdated = nUpdated + nAffect
        sb.Clear
        i = 0
    End If
    loctable.MoveNext
Loop

If sb.Length > 0 Then
    con.Execute sb.GetAsString, nAffect
    nUpdated = nUpdated + nAffect
End If

prog1.Value = 0
prog1.Visible = False

Me.Caption = sCaption
MsgBox nUpdated
End Sub

Private Sub optclosed_Click(Index As Integer)
myUndo
End Sub

Private Sub optLading1_Click(Index As Integer)
myUndo
End Sub

Private Sub optLading_Click(Index As Integer)
myUndo
End Sub

Private Sub xClosed_Click()
If bIgClick Then Exit Sub
If xClosed.Value = 0 Then
    con.Execute "update file6_20h set " & _
                "isnew = 1," & _
                "isclosed = 0," & _
                "printed = 0" & _
                " where doc_no = " & MyParn(xdoc_no.text), _
                nAffect
    If nAffect <> 0 Then myInform " „ › Õ «·„” ‰œ"
Else
    con.Execute "update file6_20h set " & _
                 "isnew = 1," & _
                 "isclosed = 1," & _
                 "printed = 1" & _
                 " where doc_no = " & MyParn(xdoc_no.text), _
                 nAffect
    If nAffect <> 0 Then myInform " „ «€·«ﬁ «·„” ‰œ"
End If
myUndo
End Sub
Private Sub CMD_PRINT_Click()
    doprint_doc
End Sub
Private Sub cmd_tsal_Click()
'    Dim cPassword As String, aUser As Variant
'    ManagPass.pPassWord = ""
'    ManagPass.Show 1
'    cPassword = ManagPass.pPassWord
'    If Trim(cPassword) = "" Then Exit Sub
'    aUser = aGetDesca("SELECT CODE , DESCA FROM SUB_USER WHERE ( ( PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0 ) OR ( STOP2 = 0 AND PASSWORD2 = " & MyParn(cPassword) & "))", con)
'    If UBound(aUser) = 0 Then Exit Sub
    TDaySal.Show 1
End Sub
Private Sub CmdDel_Click()
Dim cPassword As String, aUser As Variant

If Not bEditRecord Then Exit Sub

'ManagPass.pPassWord = ""
'ManagPass.Show 1
'cPassword = ManagPass.pPassWord
'If Trim(cPassword) = "" Then Exit Sub
'aUser = aGetDesca("SELECT CODE , DESCA FROM SUB_USER WHERE ( ( PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0 ) OR ( STOP2 = 0 AND PASSWORD2 = " & MyParn(cPassword) & "))", con)

'If UBound(aUser) = 0 Then Exit Sub


    If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) <> vbOK Then Exit Sub
    
    AddLod_Data cusername, 2, " Õ–› „»Ì⁄«  ", con, xdoc_no.text, xDate.text, , xCodeDesca.Caption
    
    con.BeginTrans
        ' Õ–› «·„” ‰œ
        con.Execute "Delete  From FILE6_20 where Doc_No = " & MyParn(xdoc_no.text)
        con.Execute "Delete  From FILE6_20H where Doc_No = " & MyParn(xdoc_no.text)
        
        Dim nAffect As Integer
        
        If xonline_doc.Caption <> "" Then
            con.Execute "UPDATE FILE6_90H SET " & _
                        " SHIP_NO = NULL," & _
                        " SALES_DOC = NULL," & _
                        " SALES_DATE = NULL " & _
                        " WHERE doc_no = " & MyParn(xonline_doc.Caption), nAffect
            If nAffect = 1 Then Inform "  „ «·€«¡  —ÕÌ· «·ÿ·»  "
        End If
    con.CommitTrans

    'If bEdit Then
        CmdNewInv_Click
'    ElseIf Not openCardTable(tbMode.tbPrevious, xdoc_no.text) Then
'        If Not openCardTable(tbMode.tbFirst) Then
'            myDefine
'        End If
'    End If

Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
If mySendReceipt(xdoc_no.text) Then
    myUndo
End If
End Sub
Private Sub CmdInform_Click()
SalesLadingLookup Me, oSearchDoc, cFilter
End Sub
Private Sub cmdinformbarcode_Click()
    grdsales.Show 1
End Sub
Private Sub CmdNewInv_Click()
myDefine
'grid1.SetFocus
End Sub
Private Sub cmdSave_Click()
'On Error GoTo myerror
If mysave(, , True) Then CmdNewInv_Click
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function mySendReceipt(pDoc_No As String) As Boolean
Dim sError As String
If Not BranchReceipt Then Exit Function
If Not SendReceipt(pDoc_No, con, sError) Then
    MsgBox sError
Else
    Inform " „ «—”«· «·«Ì’«· " & pDoc_No & " »‰Ã«Õ ⁄·Ì «·„Êﬁ⁄ «·«·Ìﬂ —Ê‰Ì"
    mySendReceipt = True
End If
End Function
Private Sub cmdTransFrom_Click()
    Dim cString As String
    transManfrm.sDate = xDate.text
    transManfrm.sCaption = "”Õ» „‰ Œ“Ì‰… " & xBox.text
    transManfrm.sBox1 = xBox.BoundText
    transManfrm.Sbox2 = GetDesca("Select code from file0_50 where type = 1", con)
    transManfrm.Show 1
End Sub
Private Sub cmdTransTo_Click()
    Dim cString As String
    transManfrm.sDate = xDate.text
    transManfrm.sCaption = "«Ìœ«⁄ ›Ì Œ“Ì‰… " & xBox.text
    transManfrm.Sbox2 = xBox.BoundText
    transManfrm.sBox1 = GetDesca("Select code from file0_50 where type = 2", con)
    transManfrm.Show 1
End Sub
Private Sub CmdUndo_Click()
myUndo
End Sub
Private Sub cmdopen_Click()
CardLookup " AND PRINTED = 0"
End Sub
Private Sub DAYALL_Click()
    myUndo
End Sub
Private Sub Form_Activate()
On Error Resume Next
If Not bAct Then
    bAct = True
    If xdoc_no.Tag = LoadMode Then
        grid1.SetFocus
        'CellPos 13, 0, grid1.Cols - 1
    Else
        grid1.SetFocus
        'CellPos 13, 0, grid1.Cols - 1
    End If
End If
Err.Clear
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
End If
End Sub
Private Sub Form_Load()
openCon con

fixDateSales

myLoadSetting
HandleMain

FixAddress

Set data12.Recordset = myCmd("SELECT * FROM FILE6_25 WHERE ISSTOP = 0 AND ( CODE = '0001' OR BRANCH = " & MyParn(cBranch) & " ) ORDER BY DESCA ", con)
Set xMan.RowSource = data12
xMan.ListField = "Desca"
xMan.BoundColumn = "Code"

Set data1.Recordset = myCmd("SELECT * FROM FILE0_40 WHERE BRANCH = " & MyParn(cBranch), con)
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"
xStore.BoundText = cBranchStore

Set data2.Recordset = myCmd("SELECT * FROM MOSM ", con)
Set xMosm.RowSource = data2
xMosm.ListField = "DESCA"
xMosm.BoundColumn = "MOSM"

Set DATA6.Recordset = myCmd("SELECT CODE,DESCA FROM SHIP ORDER BY DESCA", con)
Set xship.RowSource = DATA6
xship.ListField = "DESCA"
xship.BoundColumn = "CODE"

Set data4.Recordset = myCmd("SELECT * FROM FILE0_50 WHERE BRANCH = " & MyParn(cBranch), con)
Set xBox.RowSource = data4
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"
xBox.BoundText = cBranchBox

Set grid1.DataSource = DATA11
bIgFilter = True
Set data20.Recordset = myCmd(YearsString(2023), con)
Set xYear.RowSource = data20
xYear.ListField = "CODE"
xYear.BoundColumn = "CODE"
xYear.BoundText = Year(Date)

Set data21.Recordset = myCmd(MonthString(12), con)
Set xMonth.RowSource = data21
xMonth.ListField = "CODE"
xMonth.BoundColumn = "CODE"

xMonth.BoundText = Month(Date)

If cBranch <> "" Then
    If Not myField("select online from file0_40 where code = " & MyParn(cBranchStore), con, , , False) Then
        MsgBox "«·›—⁄ «·Õ«·Ì ·Ì” ›—⁄ online"
        bEdit = False
    End If
End If


If sDoc_no = "" Then
    If Not openCardTable Then myDefine
Else
    If Not openCardTable Then myDefine
End If

Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

CardTable.Close
Set CardTable = Nothing

closeCon con

Set sales_onlinefrm = Nothing
'If sBranchCode <> "00" And bopt1 And Not lMainShow Then
'    End
'End If

'If CBRAsBranchCode <> "00" Then End
Err.Clear
End Sub
Public Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror

Dim bNew As Boolean

With grid1

If Not validRow(Row) Then
    CalcTotals
    Exit Sub
End If

If Row = grid1.Rows - 1 Then
    myAddItem
    bNew = True
End If
CalcTotals

If myreplace(Row) Then
    If xdoc_no.Tag = DefineMode Then
        openCardTable tbMode.tbFind, xdoc_no.text
    ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myLoadGrd
    End If
    If bNew Then
        bStopCell = True
        grid1.Select grid1.Rows - 1, 1
    End If
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Grid1_EnterCell()
With grid1
    If Not bEditRecord Then
        .Editable = flexEDNone
    ElseIf .Col = 1 Or .Col = 10 Or .Col = 17 Or .Col = 15 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub grid1_GotFocus()
Grid1_EnterCell
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'If grid1.Col = 12 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    'If Col = 12 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim sBarCode As String
Dim nSal As Double
Dim nRet As Double
Dim cPassword As String, aUser As Variant
With grid1
    If Col = 10 And Val(.EditText) < 0 Then
        Dim cString As String
        ManagPass.pPassWord = ""
        ManagPass.Show 1
        cPassword = ManagPass.pPassWord
        If Trim(cPassword) = "" Then
            Cancel = True
            Exit Sub
        End If
        aUser = aGetDesca("SELECT CODE , DESCA FROM SUB_USER WHERE ( isreturn = 1 and (( PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0 ) OR ( STOP2 = 0 AND PASSWORD2 = " & MyParn(cPassword) & ")))", con)
        If UBound(aUser) > 0 Then
            lRetDoc = True
            xusername_RET.Caption = aUser(2)
            MsgBox "Ì„ﬂ‰  ”ÃÌ· «·„— Ã⁄"
        Else
            Cancel = True
            Exit Sub
        End If
    ElseIf Col = 1 Then
        grid1.EditText = retItemCode(grid1.EditText, con)
        
        If Not isNum(grid1.EditText) Then
            Cancel = True
            Exit Sub
        ElseIf Not GrdDesc(grid1.EditText, Row) Then
            Cancel = True
            Exit Sub
        End If
    ElseIf Col = 10 Then
        If Val(.EditText) > 0 And Val(.TextMatrix(Row, 10)) < 0 Then
            MsgBox "ÌÃ» Õ–› «·”Ã· ﬁ»·  ⁄œÌ·… „‰ „— Ã⁄ ≈·Ï „»Ì⁄« "
            Cancel = True
            Exit Sub
        End If
    End If
End With
End Sub
Private Sub cmdPrint_Click()
If grid1.Rows = 2 Then Exit Sub
End Sub
Private Sub xdate_open_Change()
DAYALL.Enabled = True
End Sub
Private Sub xdate_open_LostFocus()
    myValidDate xdate_open
End Sub

Private Sub xDiscount_LostFocus()
myLostFocus xDiscount
CalcTotals
End Sub
Private Function myValid(Optional bIgMsg As Boolean = False, Optional bClose As Boolean = False) As Boolean

If Trim(xCode.text) = "" Then
    MsgBox "·« ÌÊÃœ ﬂÊœ ⁄„Ì· «Ê ﬂÊœ „Ê—œ"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If


If Not (xStore.MatchedWithList) Then
    If Not bIgMsg Then MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If

If (Not xBox.MatchedWithList) And cBranch <> "00" Then
    If Not bIgMsg Then MsgBox "·„ Ì „  ÕœÌœ «·Œ“‰…"
    Exit Function
End If

If xCodeDesca.Caption = "" Then
    If Not bIgMsg Then MsgBox "·« ÌÊÃœ ﬂÊœ ⁄„Ì·"
    Exit Function
End If

If Not ValidQuant Then
    If Not bIgMsg Then MsgBox "ﬂ„Ì… „ÊÃ»… Êﬂ„Ì… ”«·»… ›Ï ‰›” «·›« Ê—…"
    Exit Function
End If

If myFormat(sDateSales) > myFormat(Date) Then
    MsgBox " «—ÌŒ  «·ÃÂ«“ «ﬂ»— „‰  «—ÌŒ «·ÌÊ„"
    Exit Function
End If

If bClose Then
    If Not xship.MatchedWithList Then
        MsgBox "‘—ﬂ… «·‘Õ‰ €Ì— „”Ã·…"
        Exit Function
    End If
    
    If Trim(xship_no.text) = "" Then
        MsgBox "—ﬁ„ «·»Ê·Ì’… €Ì— „”Ã·"
        Exit Function
    End If
End If

myValid = True


End Function
Private Sub myload(Optional bLeaveBal As Boolean = False)
xdoc_no.text = CardTable!doc_no
xdoc_ret.text = ""
xUUID_RC.Caption = CardTable!UUID_RC & ""

bIgClick = True
xSendRc.Value = IIf(CardTable!sendRc, 1, 0)
xClosed.Value = IIf(CardTable!isClosed, 1, 0)
bIgClick = False

XBRANCH.Caption = CardTable!branch & ""
xCode.Enabled = False
XISRETS.Value = 0
XISONEST.Value = 0
XISONEST.Value = IIf(CardTable!ISONEST, 1, 0)
'XISNODEL.Value = IIf(CardTable!ISNODEL, 1, 0)
xIpName.Caption = CardTable!user_ip & ""
xship_no.text = CardTable!SHIP_NO & ""
xlading_type.Caption = CardTable!lading_type_desca & ""
xlading_type.Tag = CardTable!lading_type & ""



xinv_no.text = CardTable!INV_NO & ""
xDoc_no2.text = CardTable!Doc_no2
xDate.text = myFormat_p(CardTable!Date)
xStore.BoundText = CardTable!store & ""
xMan.BoundText = CardTable!MAN & ""
xBox.BoundText = CardTable!BOX & ""
xNotes.text = CardTable!NOTES & ""

xship.BoundText = CardTable!SHIP & ""

xonline_doc.Caption = CardTable!ONLINE_DOC & ""
XSALES_RET.Caption = CardTable!SALES_RET & ""
xusername_RET.Caption = CardTable!username_ret & ""
xdate_Delivery.text = myFormat_p(CardTable!DATE_DELIVERY)
'xNOTE_DISC.Caption = CardTable!NOTE_DISC & ""
'XDOC_DISC.Caption = CardTable!DOC_DISC & ""

xCode.text = CardTable!code & ""
xCodeDesca.Caption = CardTable!CLIENTDESCA & ""

panel1(1).Caption = CardTable!UserName & ""
panel1(2).Caption = CardTable!username_disc & ""
panel1(3).Caption = CardTable!username_ret & ""

'xusername.Caption = CardTable!UserName & ""
'XUSERNAME_DISC.Caption = CardTable!username_disc & ""

chkCash.Value = IIf(CardTable!CASH, 1, 0)
xDiscount.text = TurnValue(Val(CardTable!discount & ""), 0, "")
xRate.text = TurnValue(Val(CardTable!Rate & ""), 0, "")

xIsRet.Value = 0
xPrinted.Value = IIf(CardTable!PRINTED, 1, 0)
'xTotal.Enabled = IIf(xPrinted.Value = 0, True, False)
xtime.Caption = Format(CardTable!Time, "hh:nn")
xLading.Caption = CardTable!LADING & ""
Handlecontrols LoadMode
myLoadGrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub myDefine()
xIsRet.Value = 0
xUUID_RC.Caption = ""
'XISNODEL.Value = 0

bIgClick = True
xSendRc.Value = 0
xClosed.Value = 0
bIgClick = False

XBRANCH.Caption = cBranch
xship.BoundText = ""
xship_no.text = ""
xLading.Caption = ""
xLading.Tag = ""
xonline_doc.Caption = ""
lDiscModelRet = False
xIpName.Caption = GetComputerName
'XTRANS_DOC.Caption = ""
XISRETS.Value = 0
'XCODEVISA.Enabled = True
sUserDisc = ""
lIsDocRet = False
xMan.BoundText = IIf(Not xMan.MatchedWithList, "0001", xMan.BoundText)
'XUSERNAME_DISC.Caption = ""
xRate.Locked = True
xDiscount.Locked = True
xdoc_ret.text = ""
'cmd_closed.BackColor = &H8080FF
'cmd_closed.Caption = "› Õ „” ‰œ"
'xClosed.Visible = False
'xClosed.Enabled = bopt2

'XGIFT.Value = 0
xusername_RET = ""
'xNOTE_DISC.Caption = ""
'XDOC_DISC.Caption = ""

'xcard_disc.text = ""
'xcard_desca.Caption = ""
'xcard_value.Caption = ""
'xcard_doc.Caption = ""

XSALES_RET.Caption = ""
xdoc_no.text = ""
xdoc_no.Tag = DefineMode
xDoc_no2.text = ""
xinv_no.text = ""
panel1(1).Caption = cusername

'xusername.Caption = cusername
    
xStore.BoundText = cBranchStore
xBox.BoundText = cBranchBox
xCode.text = "0000"
xCodeDesca.Caption = cDefClientDesca
xDate.text = myFormat_p(sDateSales)
'xBalance.Caption = ""
xMosm.BoundText = cPMosm

xDiscount.text = ""
chkCash.Value = 1
xTotalQuant.Caption = ""
xTotalQuant2.Caption = ""
xTotalItemNoDiscount.Caption = ""
xTotalItem.Caption = ""
xTotalItemDiscount.Caption = ""
xRate.text = ""
xTotal.text = ""
xDiscount.text = ""
xTotal.text = ""


xPrinted.Value = 0
xNotes.text = ""
xtime.Caption = Format(Time, "hh:nn")


grid1.Rows = 1
myAddItem

fixGrd

Handlecontrols DefineMode

CalcTotals
On Error Resume Next
'grid1.SetFocus
'Err.Clear
End Sub
Private Sub Handlecontrols(nMode)
xdoc_no.Tag = nMode
cmdNewInv.Enabled = bEdit
bEditRecord = (bEdit And xPrinted.Value = 0) And xLading.Caption = "" And xlading_type.Tag = ""

cmddel.Enabled = bEditRecord And nMode = LoadMode

cmdSave.Enabled = bEditRecord
cmdShip.Enabled = bEditRecord

xClosed.Enabled = bSuperSales And nMode = LoadMode And xLading.Caption = "" And xlading_type.Tag = ""

xIsRet.Enabled = bSuperSales And nMode = LoadMode And XSALES_RET.Caption = "" And Val(xTotal.text) > 0
xIsRet.Enabled = xIsRet.Enabled And xPrinted.Value = 1 And xClosed.Value = 1

xdoc_ret.Enabled = (xPrinted.Value = 0)
cmd_showdoc.Enabled = (xPrinted.Value = 0)

cmdSaveLading.Enabled = nMode = LoadMode And _
                        (XBRANCH.Caption = cBranch) And xClosed.Value = 1 _
                        And Val(xlading_type.Tag) > 2 _
                        And xship.MatchedWithList And xship_no.text <> "" _
                        And (xdate_Delivery.text = "" Or bOpt9)

xBox.Enabled = False
xdoc_no.Tag = nMode
xCode.Enabled = xPrinted.Value = 0 And bEdit = True
xIsRet.Enabled = bOpt9 And nMode = LoadMode

cmdGo.Enabled = (xUUID_RC.Caption = "" Or xSendRc.Value = 0) And nMode = LoadMode And xPrinted.Value = 1
xSendRc.Enabled = nMode = LoadMode And xUUID_RC.Caption <> "" And xPrinted.Value = 1

Dim nRecord As Long, nRecords As Long
retRecords xdoc_no.text, nRecords, nRecord

cmdNext.Enabled = nRecord < nRecords And nRecords <> 0 And nMode = LoadMode
cmdPrevious.Enabled = nRecord <> 1 And nRecords <> 0 And nMode = LoadMode
cmdLast.Enabled = nRecord < nRecords And nRecords <> 0 And nMode = LoadMode
cmdFirst.Enabled = nRecord <> 1 And nRecords <> 0 And nMode = LoadMode

If nMode = LoadMode Then
    panel1(0).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
    panel1(0).Caption = "”Ã· ÃœÌœ (" & (nRecords + 1) & ")"
End If
End Sub
Private Sub HandleMain()
computerName = GetComputerName
fmRc.Visible = BranchReceipt
fmDate.Visible = (bSuperSales) Or bOpt9
fmDay.Visible = (bSuperSales) Or bOpt9
fmClose.Visible = (bSuperSales) Or bOpt9
fmMonth.Visible = (bSuperSales) Or bOpt9
End Sub
Private Sub xDoc_No_LostFocus()
If Trim(xdoc_no.text) = "" Then
    myDefine
    Exit Sub
End If
If Not openCardTable(tbMode.tbFind, xdoc_no.text) Then
   If xCode.Tag = LoadMode Then
        myDefine
   End If
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
'If IsPrinted Then Exit Sub
'If (XSALES_RET.Caption <> "" And grid1.ValueMatrix(grid1.Row, 10) < 0) Then Exit Sub

'If KeyCode = 46 And grid1.Row = grid1.Rows - 1 And xPrinted.Value = 0 And bEdit And grid1.Rows > 3 Then
'    grid1.Select grid1.Rows - 2, 0
'End If
If Not bEditRecord Then Exit Sub
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And bEditRecord And grid1.Rows > 3 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        RemoveItem grid1.Row
        CalcTotals
        MakeSerial grid1.Row
        bStopCell = True
        grid1.SetFocus

        
'        grid1.Select grid1.Rows - 1, 1
'        grid1.ShowCell grid1.Rows - 1, 1
        
    End If
ElseIf KeyCode = 112 Then
'    Grid1.Row = Grid1.Rows - 1
'    Grid1.Col = 1
    ItemsLookupAll Me, oSearchItem
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub
Public Function GrdDesc(sitem As String, Row As Long) As Boolean
On Error GoTo myerror
With grid1
    .TextMatrix(Row, 1) = ""
    .TextMatrix(Row, 2) = ""
    .TextMatrix(Row, 3) = ""
    .TextMatrix(Row, 4) = ""
    .TextMatrix(Row, 5) = ""
    .TextMatrix(Row, 6) = ""
    .TextMatrix(Row, 7) = ""
    .TextMatrix(Row, 8) = ""
    .TextMatrix(Row, 9) = ""
    .TextMatrix(Row, 10) = ""
    .TextMatrix(Row, 11) = ""
    .TextMatrix(Row, 12) = ""
    .TextMatrix(Row, 13) = ""
    .TextMatrix(Row, 14) = ""
    .TextMatrix(Row, 17) = ""

    If Trim(sitem) = "" Then Exit Function
    
    Set loctable = ItemFind(sitem, con)
    If Not loctable.EOF Then
        .TextMatrix(Row, 1) = loctable!Item
        .TextMatrix(Row, 2) = loctable!MOSM
    
        .TextMatrix(Row, 3) = loctable!Fact
        .TextMatrix(Row, 4) = loctable!SUPP & ""
        .TextMatrix(Row, 5) = loctable!modelfact0 & ""
        .TextMatrix(Row, 6) = loctable!DESCA
        .TextMatrix(Row, 7) = loctable!SCAL
        .TextMatrix(Row, 8) = loctable!color
        
        grid1.TextMatrix(Row, 9) = Format(loctable!price)
        
        If Trim(.TextMatrix(Row, 10)) = "" Then
            grid1.TextMatrix(Row, 10) = "1"
        End If
        
        grid1.TextMatrix(Row, 11) = Format(loctable!price)
        grid1.TextMatrix(Row, 19) = Format(loctable!price)
        grid1.TextMatrix(Row, 12) = 0
        grid1.TextMatrix(Row, 13) = Val(grid1.TextMatrix(Row, 10)) * Val(grid1.TextMatrix(Row, 11))
        grid1.TextMatrix(Row, 17) = loctable!MODELNO & ""
        GrdDesc = True
    End If
End With
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function CalcTotals(Optional nMode As Integer = 0)
Dim nTotalQuant As Double, nTotalQuant2 As Double, nTotalItem As Double, nTotalItemNoDiscount As Double, nTotalDiscount As Double, nTotalMin As Double
Dim nTotal2 As Double
With grid1
For i = 1 To grid1.Rows - 2
    nTotal2 = nTotal2 + (Val(.TextMatrix(i, 10)) * Val(.TextMatrix(i, 9)))
    .TextMatrix(i, 13) = Val(.TextMatrix(i, 10)) * Val(.TextMatrix(i, 11))
    If cBranch = "00" Then
        If Val(.TextMatrix(i, 9)) <> Val(.TextMatrix(i, 11)) And xStore.BoundText <> "000" Then
            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
            .Cell(flexcpFontStrikethru, i, 0, i, .Cols - 1) = True
        End If
    End If
    
    If Val(grid1.TextMatrix(i, 10)) > 0 Then
        nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(i, 10))
    Else
        nTotalQuant2 = nTotalQuant2 + Abs(Val(grid1.TextMatrix(i, 10)))
    End If
    nTotalItem = nTotalItem + Val(.TextMatrix(i, 10)) * Val(.TextMatrix(i, 11))
    If cBranch = "00" Then
        nTotalItemNoDiscount = nTotalItemNoDiscount + (Val(.TextMatrix(i, 11)) * Val(.TextMatrix(i, 10)))
        nTotalItemDiscount = nTotalItemDiscount + (Val(.TextMatrix(i, 11)) * Val(.TextMatrix(i, 10))) - (Val(.TextMatrix(i, 11)) * Val(.TextMatrix(i, 10)))
    Else
        nTotalItemNoDiscount = nTotalItemNoDiscount + (Val(.TextMatrix(i, 9)) * Val(.TextMatrix(i, 10)))
        nTotalItemDiscount = nTotalItemDiscount + (Val(.TextMatrix(i, 9)) * Val(.TextMatrix(i, 10))) - (Val(.TextMatrix(i, 11)) * Val(.TextMatrix(i, 10)))
    End If
Next
'XTOTAL2.Caption = Format(nTotal2, "#0.00")
xTotalQuant.Caption = Myvalue(nTotalQuant, "#0")
xTotalQuant2.Caption = Myvalue(nTotalQuant2, "#0")
xTotalItemNoDiscount.Caption = Format(nTotalItemNoDiscount, "#0.00")
xTotalItemDiscount.Caption = Format(nTotalItemDiscount, "#0.00")
xTotalItem.Caption = Format(nTotalItem, "#0.00")

If Val(xTotalItem.Caption) <> 0 Then
    If Round(Val(xRate.text), nRound) <> Round(Val(xDiscount.text) / Val(xTotalItem.Caption) * 100, nRound) Then
        xRate.text = Myvalue(Round((Val(xDiscount.text) / Val(xTotalItem.Caption)) * 100, nRound))
    End If
Else
    xRate.text = ""
End If

xTotal.text = Round(nTotalItem - Val(xDiscount.text), 2)
End With
End Function
Private Sub CardLookup(Optional pFilter As String = "")
Dim Generalarray(5)
Dim listarray
Dim GrdArray(9, 1)
Set Generalarray(0) = Me
If cBranch = "00" Then
    Generalarray(1) = "SELECT   FILE6_20H.DOC_NO , FILE6_20H.DOC_NO2, FILE6_20H.TRANS_DOC, CONVERT(VARCHAR(10), FILE6_20H.date, 111) , FILE0_50.DESCA, CONVERT(VARCHAR(10), FILE6_20H.time, 108) AS Expr2, FILE3_10.DESCA  , SUM(FILE6_20.TOTAL) AS total , file0_40.desca  , file6_20h.ONLINE_DOC" & _
                      " FROM    FILE6_20H LEFT JOIN  FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO INNER JOIN FILE3_10 ON FILE6_20H.CODE = FILE3_10.CODE LEFT OUTER JOIN    FILE0_50 ON FILE6_20H.BOX = FILE0_50.CODE  INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE WHERE FILE6_20H.ISINVOICE = 0 "
Else
'                                   0               1                       2                   3                                       4                   5                   6
    Generalarray(1) = "SELECT   FILE6_20H.DOC_NO , FILE6_20H.DOC_NO2, FILE6_20H.INV_NO, CONVERT(VARCHAR(10), FILE6_20H.date, 111) , FILE0_50.DESCA, CONVERT(VARCHAR(10), FILE6_20H.time, 108) AS Expr2, FILE3_10.DESCA  , SUM(FILE6_20.TOTAL) AS total , file0_40.desca , file6_20h.ONLINE_DOC" & _
                      " FROM    FILE6_20H LEFT JOIN  FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO INNER JOIN FILE3_10 ON FILE6_20H.CODE = FILE3_10.CODE LEFT OUTER JOIN    FILE0_50 ON FILE6_20H.BOX = FILE0_50.CODE INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE WHERE FILE6_20H.ISINVOICE = 0 "
End If
If lIsBranchStore Then Generalarray(1) = Generalarray(1) & " AND FILE6_20H.BRANCH = " & MyParn(cBranch)
If cBranch = "00" Then
    If lNotBranch Then
'        Generalarray(1) = Generalarray(1) & " AND FILE0_40.ISSTOREBRANCH = 1 "
        Generalarray(1) = Generalarray(1) & " AND FILE6_20H.branch <> '00' "
    Else
        Generalarray(1) = Generalarray(1) & " AND FILE6_20H.branch = '00' "
    End If
End If

If cBranch <> "00" Or lNotBranch Then
    If IsDate(xdate_open.text) Then
        cFilter = cFilter & " AND  FILE6_20H.date = " & DateSq(Format(xdate_open.text, "DD-MM-YYYY"))
        cFilterLook = cFilterLook & " AND FILE6_20H.date = " & DateSq(Format(xdate_open.text, "DD-MM-YYYY"))
    ElseIf IsDate(sDateSales) Then
        cFilter = cFilter & " AND FILE6_20H.date = " & DateSq(sDateSales)
        cFilterLook = cFilterLook & " AND FILE6_20H.date = " & DateSq(sDateSales)
    End If
    If pFilter <> "" Then Generalarray(1) = Generalarray(1) & pFilter
    If cFilterLook <> "" Then Generalarray(1) = Generalarray(1) & cFilterLook
End If

If cBranch = "00" Then
    Generalarray(2) = "GROUP BY FILE6_20H.DOC_NO, FILE3_10.DESCA , FILE6_20H.DOC_NO2, FILE6_20H.TRANS_DOC, CONVERT(VARCHAR(10), FILE6_20H.date, 111), FILE0_50.DESCA, CONVERT(VARCHAR(10),  FILE6_20H.time, 108) , file0_40.desca , file6_20h.ONLINE_DOC Order by doc_no2 DESC "
Else
    Generalarray(2) = "GROUP BY FILE6_20H.DOC_NO, FILE3_10.DESCA , FILE6_20H.DOC_NO2, FILE6_20H.INV_NO, CONVERT(VARCHAR(10), FILE6_20H.date, 111), FILE0_50.DESCA, CONVERT(VARCHAR(10),  FILE6_20H.time, 108) , file0_40.desca , file6_20h.ONLINE_DOC Order by doc_no2 DESC "
End If
Generalarray(3) = 6000
Generalarray(5) = False

If bopt3 Then
   ReDim listarray(3, 5)
   listarray(0, 0) = "«·—ﬁ„-«·Œ“‰…"
   listarray(0, 1) = "(@@Doc_No2@@5 or %%file0_50.DESCA%% or %%file3_10.DESCA%%) "

   listarray(1, 0) = "«· «—ÌŒ"
   listarray(1, 1) = "(##DATE##)"


    If cBranch = "00" Then
        listarray(2, 0) = "„” ‰œ «· ÕÊÌ·"
        listarray(2, 1) = "(%%TRANS_DOC%%)"
    Else
        listarray(2, 0) = "—ﬁ„ «·«Ì’«·"
        listarray(2, 1) = "(%%INV_NO%%)"
    End If

   listarray(3, 0) = "«·„Œ“‰"
   listarray(3, 1) = "%%file0_40.DESCA%%"

Else
   ReDim listarray(2, 5)

    listarray(0, 0) = "«·—ﬁ„"
    listarray(0, 1) = "(@@Doc_No2@@4 or %%file3_10.DESCA%%)"
   
    If cBranch = "00" Then
        listarray(1, 0) = "„” ‰œ «· ÕÊÌ·"
        listarray(1, 1) = "(%%TRANS_DOC%%)"
    Else
        listarray(1, 0) = "—ﬁ„ «·«Ì’«·"
        listarray(1, 1) = "(%%INV_NO%%)"
    End If
   
   listarray(2, 0) = "«·„Œ“‰"
   listarray(2, 1) = "%%file0_40.DESCA%%"

End If

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(1, 1) = 1200

GrdArray(2, 0) = "„” ‰œ  ÕÊÌ· "
If cBranch = "00" Then
    GrdArray(2, 1) = 1000
Else
    GrdArray(2, 1) = 0
End If
GrdArray(3, 0) = "«· «—ÌŒ"
GrdArray(3, 1) = 1500

GrdArray(4, 0) = "«·Õ“‰…"
GrdArray(4, 1) = 2000

GrdArray(5, 0) = "«·Êﬁ "
GrdArray(5, 1) = 1300

GrdArray(6, 0) = "«·⁄„Ì·"
GrdArray(6, 1) = 2500

GrdArray(7, 0) = "«·ﬁÌ„…"
GrdArray(7, 1) = 1300

GrdArray(8, 0) = "«·„Œ“‰"
GrdArray(8, 1) = 2000

GrdArray(9, 0) = "ONLINE"
GrdArray(9, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.sid = "sales" & IIf(bopt3, "1", "0")
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Sub xdoc_ret_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then cmd_showdoc_Click
End Sub

Private Sub xIsRet_Click()
If bIgClick Then Exit Sub

If MsgBox(IIf(xIsRet.Value = 1, "«·”„«Õ »„— Ã⁄ ··›« Ê—…", "«·€«¡ „— Ã⁄ «·›« Ê—…"), vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub

Dim nAffect As Integer
con.Execute "UPDATE FILE6_20H SET " & _
            " ISRET = " & xIsRet.Value & "," & _
            " ISNEW = 1" & _
            " WHERE DOC_NO = " & MyParn(xdoc_no.text), nAffect

If nAffect = 1 Then myInform " „  ⁄œÌ· «·„” ‰œ"
End Sub
Private Sub XPHONE_Click()
    PhoneCust.Show 1
End Sub

Private Sub xMAN_Validate(Cancel As Boolean)
If Not xMan.MatchedWithList Then
    If Not IsNumeric(xMan.text) Then Exit Sub
    xMan.BoundText = RetZero(xMan.text, 4)
    If Not xMan.MatchedWithList Then
        xMan.BoundText = ""
    End If
End If
End Sub

Private Sub xMonth_Change()
If chkDay.Value = 1 Or chkDayBranch.Value = 1 Or IsDate(xdate1.text) Or IsDate(xdate2.text) Then Exit Sub
myUndo
End Sub

Private Sub xRate_LostFocus()
myLostFocus xRate
If Val(xTotalItem.Caption) <> 0 Then
    If Round(Val(xRate.text), nRound) <> Round(Val(xDiscount.text) / Val(xTotalItem.Caption) * 100, nRound) Or xDiscount.Locked Then
        xDiscount.text = Round((Val(xRate.text) * Val(xTotalItem.Caption)) / 100, 0)
    End If
Else
    xDiscount.text = ""
End If
CalcTotals
End Sub
Private Function RetItemBalance(cItem, cStore, dDate) As Double
If cItem = "" Then Exit Function
movetable.Seek Array(cItem, cStore), adSeekFirstEQ
Do Until movetable.EOF
    If IsNull(movetable!Date) Then Exit Do
    If Trim(movetable!Item) <> cItem Or cStore <> movetable!store Or DateValue(movetable!Date) > DateValue(Format(dDate, "dd-mm-yyyy")) Then Exit Do
    RetItemBalance = RetItemBalance + TurnValue(movetable!In, Null, 0) - TurnValue(movetable!out, Null, 0)
    movetable.MoveNext
Loop
End Function
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub fixGrd()

With grid1
.Cols = 22
'                0        1           2          3          4           5           6           7       8              9           10          11           12            13          14                    15              16          17
.FormatString = "„.|" & "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "”⁄— „” Â·ﬂ|" & "⁄œœ|" & "”⁄— «·»Ì⁄|" & "‰”»… Œ’„|" & "«·≈Ã„«·Ï|" & "”⁄— √Êﬂ«“ÌÊ‰|" & "ﬂÊœ «·»«∆⁄|" & "«·»«∆⁄|" & "—ﬁ„ «·„ÊœÌ·|"
.RowHeight(0) = 900
.WordWrap = True
.ColHidden(.Cols - 1) = True
.ColHidden(14) = True
.ColHidden(18) = True
'.ColHidden(16) = True
.ColHidden(17) = True
.ColWidth(17) = 2000
.ColWidth(0) = 600
.ColWidth(1) = 1000
.ColWidth(2) = 0
.ColWidth(3) = 1300
.ColWidth(4) = 0
.ColWidth(5) = 1500
.ColWidth(6) = 2800
.ColWidth(7) = 700
.ColWidth(8) = 1000
.ColWidth(9) = 1000
.ColWidth(10) = 800
.ColWidth(11) = 1000
.ColWidth(12) = 500
.ColWidth(13) = 1200
.ColWidth(15) = 1700
.ColWidth(16) = 0

.ColHidden(19) = True
.ColHidden(20) = True
If cBranch = "00" And Not lNotBranch Then
    .ColHidden(12) = True
    .ColHidden(14) = True
'   .ColHidden(9) = True
End If

'.ColFormat(12) = "#.##%"
'.ColComboList(3) = clist1

If cBranch > "60" Then
    .ColComboList(15) = StrList("SELECT * FROM FILE6_25 WHERE ISSTOP = 0 ORDER BY DESCA ", con)
Else
    If lNotBranch Then
        .ColComboList(15) = StrList("SELECT * FROM FILE6_25  ORDER BY DESCA ", con)
    Else
        .ColComboList(15) = StrList("SELECT * FROM FILE6_25 WHERE ISSTOP = 0 AND ( BRANCH = " & MyParn(cBranch) & " OR CODE = '0001' ) ORDER BY DESCA ", con)
    End If
End If
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
.Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter

.ShowCell .Rows - 1, 0


'''''''If cBranch = "00" And XISONEST.Value <> 0 Then
'''''''    For nRow = 1 To .Rows - 1
'''''''        If .ValueMatrix(nRow, 9) <> .ValueMatrix(nRow, 11) Then
'''''''            .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
'''''''        End If
'''''''    Next nRow
'''''''End If

'If Not bopt3 And DateValue(Date) <> DateValue(xDate.Text) Then
'    grid1.Visible = False
'    Picture3.Visible = False
'    Picture1.Visible = False
'Else
'    grid1.Visible = True
'    SSCommand2.Visible = True
'    Picture3.Visible = True
'    Picture1.Visible = True
'End If
End With
End Sub
Private Sub CLIENTLOOKUP()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From file3_10"
Generalarray(2) = "Order by file3_10.Desca"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchCl1.Caption = "«” ⁄·«„"
oSearchCl1.Show 1
End Sub
Private Sub CLIENTLOOKUP2()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select BARCODE , DescA,phone1,fax,ADDRESS From file3_20 "
Generalarray(2) = " Order by file3_20.Desca"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·«”„-«· ·Ì›Ê‰-«·„Ê»Ì·-«·⁄‰Ê«‰"
listarray(0, 1) = "(%%DESCA%% OR %%PHONE1%% OR %%FAX%% OR %%ADDRESS%%)"

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = "«· ·Ì›Ê‰"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«·›«ﬂ”"
GrdArray(3, 1) = 1300

GrdArray(4, 0) = "«·⁄‰Ê«‰"
GrdArray(4, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchCl2.Caption = "«” ⁄·«„ ⁄„·«¡ «·‰ﬁœÌ"
oSearchCl2.Show 1
End Sub
Private Sub CLIENTLOOKUP3()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From file4_10"
Generalarray(2) = "Order by file4_10.Desca"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·„Ê—œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·„Ê—œ"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchCl3.Caption = "«” ⁄·«„ «·„Ê—œÌ‰"
oSearchCl3.Show 1
End Sub


Private Sub xRateDis_Lostfocus()
    xDiscount.text = Fix((Val(xTotalItem.Caption) * Val(xRateDis.text) / 100))
End Sub

Private Sub xRet1_Change()

End Sub

Private Sub xSendRc_Click()
If bIgClick Then Exit Sub
If xdoc_no.text <> "" Then
    On Error GoTo myerror
    If xUUID_RC.Caption <> "" Then
        If Not IsEmpty(myField("select doc_no from file6_20h where PREVIOUS_UUID = " & MyParn(xUUID_RC.Caption), con)) Then
            MsgBox "—ﬁ„ „”·”· „—Ã⁄Ì ”«»ﬁ ··›« Ê—… —ﬁ„ " & MyParn(xdoc_no.text)
            'xSendRc.Value = IIf(xSendRc.Value = 1, 0, 1)
        End If
    End If
    con.Execute "UPDATE FILE6_20H SET FILE6_20H.sendRc = " & xSendRc.Value & " FROM FILE6_20H WHERE FILE6_20H.DOC_NO = " & MyParn(xdoc_no.text)
    Inform " „ «· ⁄œÌ· »‰Ã«Õ"
End If
finaly:
Handlecontrols xdoc_no.Tag
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
bIgClick = True
xSendRc.Value = IIf(xSendRc.Value = 1, 0, 1)
bIgClick = False
GoTo finaly
End Sub
Private Function RemoveItem(nRow) As Boolean
Dim cPassword As String, aUser As Variant
On Error GoTo myerror
'con.BeginTrans
If grid1.TextMatrix(nRow, grid1.Cols - 1) <> "" Then

'    ManagPass.pPassWord = ""
'    ManagPass.Show 1
'    cPassword = ManagPass.pPassWord
'
'    If Trim(cPassword) = "" Then Exit Function
'    aUser = aGetDesca("SELECT CODE , DESCA FROM SUB_USER WHERE ( ( PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0 ) OR ( STOP2 = 0 AND PASSWORD2 = " & MyParn(cPassword) & "))", con)
'    If UBound(aUser) = 0 Then Exit Function
    
    con.Execute "Delete  From FILE6_20 where id = " & grid1.TextMatrix(nRow, grid1.Cols - 1)
End If
grid1.RemoveItem nRow
'con.CommitTrans
Exit Function
myerror:
'con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Sub myproc2(nDoc_no)

xdoc_no = nDoc_no
myUndo
'CardTable.Find "Doc_no = " & MyParn(nDoc_no), , adSearchForward, adBookmarkFirst
'If Not CardTable.EOF Then
'    myload
'Else
'    MsgBox "—ﬁ„ «·›« Ê—… €Ì— ’ÕÌÕ"
'    Unload Me
'End If
End Sub
Function itemPrice(cItem) As Single
    itemPrice = GetDesca("select PRICE from file1_10 where item = " & MyParn(cItem), con)
End Function
Private Function myreplaceGrd(nRow) As Boolean
Dim aInsert(12, 1)
With grid1
    For i = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        If Val(.TextMatrix(i, 10)) = 0 Or (Val(.TextMatrix(i, 10)) > 100 And cBranch <> "00") Then
            .TextMatrix(i, 10) = ""
            MsgBox "—«Ã⁄ ⁄œœ «·„»Ì⁄« "
            Exit Function
        End If
        If Len(grid1.TextMatrix(i, 1)) <= 2 Then
            MsgBox " ’‰› ⁄Ì— ’ÕÌÕ"
            Exit Function
        End If
        If Val(.TextMatrix(i, 11)) = 0 Then
            .TextMatrix(i, 11) = Val(.TextMatrix(i, 9))
        End If
        
        If Val(grid1.TextMatrix(i, 11)) = 0 Then
            MsgBox " ”⁄— ’‰› ’›— "
            Exit Function
        End If
        
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xdoc_no.text)
        
        aInsert(1, 0) = "item"
        aInsert(1, 1) = addstring(grid1.TextMatrix(i, 1))
        
        aInsert(2, 0) = "quant"
        aInsert(2, 1) = Val(.TextMatrix(i, 10))

        aInsert(3, 0) = "Price"
        aInsert(3, 1) = Val(.TextMatrix(i, 11))

        aInsert(4, 0) = "Discount"
        aInsert(4, 1) = 0

        aInsert(5, 0) = "PRICE_C"
        aInsert(5, 1) = Val(.TextMatrix(i, 9))

        aInsert(6, 0) = "row"
        aInsert(6, 1) = i
        
        aInsert(7, 0) = "COST"
        aInsert(7, 1) = Item_cost(.TextMatrix(i, 1), con)

        aInsert(8, 0) = "MAN"
        aInsert(8, 1) = addstring(.TextMatrix(i, 15))

        aInsert(9, 0) = "S_OKAZ"
        aInsert(9, 1) = Val(.TextMatrix(i, 12))


        aInsert(11, 0) = "Price_C2"
        aInsert(11, 1) = Val(.TextMatrix(i, 19))

        aInsert(12, 0) = "S_OKAZ_2"
        aInsert(12, 1) = Val(.TextMatrix(i, 20))

        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            aInsert(10, 0) = "USER_IP"
            aInsert(10, 1) = addstring(computerName)
            con.Execute CreateInsert(aInsert, "FILE6_20")
        Else
            con.Execute CreateUpdate(aInsert, "FILE6_20", " where ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myLoadGrd()
With grid1
    Dim cString As String
    Dim cPrice As String
    cField1 = "case when file6_20.Discount = 0 then Null else file6_20.Discount end "
    cPrice = " FILE6_20.PRICE_C"
    
'                           1          2                3                4                   5               6             7               8                9                  10              11               12               13                 14                  15                  16                  17              18                      19                  20          21
    cString = "SELECT FILE6_20.ITEM, FILE1_10.MOSM, FACT.DESCA , FILE1_10.SUPP, FILE1_10.MODELFACT0, FILE1_10.DESCA, FILE1_10.SCAL , FILE1_10.COLOR , " & cPrice & " , file6_20.Quant, file6_20.Price, FILE6_20.S_OKAZ ,FILE6_20.TOTAL  ,   FILE6_20.PRICE_C2 ,   FILE6_20.MAN    ,FILE6_25.DESCA ,   FILE1_10.MODELNO   , FILE1_10.BARCODE13,  FILE6_20.PRICE_C2 , FILE6_20.S_OKAZ_2,ID " & _
          " FROM (FILE6_20 LEFT JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM) LEFT JOIN FACT ON FACT.CODE = FILE1_10.FACT LEFT JOIN FILE6_25 ON FILE6_20.MAN  = FILE6_25.CODE "
    cString = cString & turn(cString) & " DOC_NO = " & MyParn(xdoc_no.text)
    cString = cString & " ORDER BY FILE6_20.ID"
    Set DATA11.Recordset = myRecordSet(cString, con)
    myAddItem
End With
CalcTotals
fixGrd
End Sub
Private Function mysave(Optional bEnd As Boolean = True, Optional bPrint As Boolean = True, Optional bSendReceipt As Boolean) As Boolean
If Not myValid(, True) Then Exit Function

CalcTotals

If Not myreplace(-1, "1", "1") Then Exit Function

myUndo
Exit Function
myerror:
    myUndo
End Function
Private Sub myLoadSetting()
bEdit = cUserBox <> ""

nCountPrint = 1
nRound = 2

bSuperSales = (cBranch = "00" And bopt2) Or bSupermode
cDefClient = "0000"
cDefClientDesca = "⁄„Ì· ‰ﬁœÌ"

computerName = GetComputerName
chkprint.Value = Val(RetSetting("print", App.Path & "\other.txt"))
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = True) As Boolean
With grid1
If Not myValid(bIgMsg) Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 1)) Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 10)) Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 11)) Then Exit Function
If Trim(grid1.TextMatrix(Row, 15)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf Col < 10 Then
    grid1.Col = 10
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 10)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub myAddItem()
grid1.AddItem ""
If grid1.Rows > 2 Then
    grid1.TextMatrix(grid1.Rows - 1, 15) = grid1.TextMatrix(grid1.Rows - 2, 15)
Else
    grid1.TextMatrix(grid1.Rows - 1, 15) = xMan.BoundText
End If
MakeSerial
End Sub
Private Sub xTotal_LostFocus()
myLostFocus xTotal
xDiscount.text = Round(Val(xTotalItem.Caption) - Val(xTotal.text), 2)
CalcTotals
End Sub
Public Sub PrintTDay()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
ReDim aHeader(1)
Dim pDate As Date, pstore As String
Dim nFBalItem As Double
Dim nTInlItem As Double
Dim nSallItem As Double
Dim nBalItem As Double
Dim nOverCash As Double, nVisa2 As Double
Dim VisaTable As New ADODB.Recordset
pDate = xDate.text
pstore = xStore.BoundText
contemp.Execute "DELETE * FROM TEMP"


    temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

    cString = " SELECT [SECTION] ,  SUM(QUANT) AS t_q, SUM(TOTAL * ((100 -rate )/ 100 )) AS t_tot FROM SALES_MODEL "
    cString = cString & " where date = " & DateSq(xDate.text)
    cString = cString & " GROUP BY [SECTION] "
    aHeader(0) = "[" & BetweenString(Format(xDate.text, "d-m-yyyy"), Format(xDate.text, "d-m-yyyy")) & "]"
    
    cString = "SELECT  SUM(T_QUANT) AS T_Q, SUM(T_QUANT1) AS T_Q1, SUM(T_QUANT2) AS T_Q2, SUM(t_total) AS T_Item, SUM(t_total1) AS T_Item1, SUM(t_total2) AS T_Item2, SUM(discount) AS T_Disc, SUM(CASH) AS T_Cash, SUM(VISA+VISA3) AS T_Visa FROM   T_SALESDOC "
    cString = cString & " where date = " & DateSq(xDate.text)
    aHeader(0) = "[" & BetweenString(Format(xDate.text, "d-m-yyyy"), Format(xDate.text, "d-m-yyyy")) & "]"
    
    
    If cManBox <> "" Then
        cString = cString & " AND BOX = " & MyParn(cManBox)
        aHeader(1) = "[" & salesfrm.xBox.text & "]"
    End If
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

cStr1 = " SELECT SUM(VALUE) FROM FILE8_60 INNER JOIN FILE8_60H ON FILE8_60.DOC_NO = FILE8_60H.DOC_NO WHERE DATE = " & DateSq(xDate.text)
If cManBox <> "" Then
    cStr1 = cStr1 & " AND BOX = " & MyParn(cManBox)
End If
nOverCash = Val(GetDesca(cStr1, con) & "")
If Not (sourcetable.EOF And sourcetable.BOF) Then
With sourcetable
    Do While Not sourcetable.EOF
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "⁄œœ „»Ì⁄« "
    
    temptable!val3 = !t_q1
    temptable.Update
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "⁄œœ „— Ã⁄« "
    
    temptable!val3 = !t_q2
    temptable.Update
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "’«›Ï ⁄œœ „»Ì⁄« "
    temptable!val3 = !t_q
    temptable.Update

    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "ﬁÌ„… „»Ì⁄« "
    temptable!val3 = !T_Item1
    temptable.Update
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "ﬁÌ„… „— Ã⁄« "
    temptable!val3 = !T_Item2
    temptable.Update
    
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "’«›Ï ﬁÌ„… „ÊœÌ·« "
    temptable!val3 = !T_Item
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "ﬁÌ„… «·Œ’„"
    temptable!val3 = !T_Disc
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "’«›Ï „»Ì⁄« "
    
    temptable!val3 = Val(!T_Item & "") - Val(!T_Disc & "")
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "”œ«œ ‰ﬁœÏ"
    temptable!val3 = Val(!T_CASH & "")
    temptable.Update
    
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "”œ«œ ›Ì“« "
    temptable!val3 = Val(!T_VISA & "")
    temptable.Update
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "√Ã· "
    temptable!val3 = Val(!T_Item & "") - Val(!T_Disc & "") - Val(!T_CASH & "") - Val(!T_VISA & "")
    temptable.Update
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "OVER CASH"
    temptable!val3 = nOverCash
    temptable.Update
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "≈Ã„«·Ï ‰ﬁœÌ…"
    temptable!val3 = nOverCash + Val(!T_CASH & "")
    temptable.Update
    
''  «Ã„«·Ï ›Ì“«
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "⁄„Ê·…  ﬁ”Ìÿ ›Ì“«"
    cString = "select sum(visa2+visa3_2) from file6_20h where date = " & DateSq(xDate.text)
    If cManBox <> "" Then cString = cString & " AND BOX = " & MyParn(cManBox)
    nVisa2 = Val(GetDesca(cString, con) & "")
    temptable!val3 = nVisa2
    temptable.Update
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "’«›Ï ›Ì“« "
    temptable!val3 = nVisa2 + Val(!T_VISA & "")
    temptable.Update
    
''  «Ã„«·Ï ›Ì“«
    
    .MoveNext
    
    Loop
End With
End If
'
cString = "SELECT   FILE0_50.DESCA AS BOXDESCA , SUM(BOXMOVE.PLUS ) AS T_IN, SUM(BOXMOVE.MINUS ) AS T_OUT, BOXMOVE.BOX, BOXMOVE.DESCA as desca, BOXMOVE.FLAG FROM         BOXMOVE INNER JOIN FILE0_50 ON FILE0_50.CODE = BOXMOVE.BOX WHERE  DATE = " & DateSq(salesfrm.xDate.text)
If cManBox <> "" Then cString = cString & " AND BOX = " & MyParn(cManBox)
cString = cString & " GROUP BY FILE0_50.DESCA ,BOXMOVE.BOX, BOXMOVE.DESCA , BOXMOVE.FLAG ORDER BY BOXMOVE.FLAG desc "
sourcetable.Close
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If Not (sourcetable.EOF And sourcetable.BOF) Then
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str11 = 6
    temptable!str12 = "≈Ã„«·Ï «·Õ“‰…"
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    temptable!str1 = !BOX
    temptable!str2 = TurnValue(!BOXDESCA)
    temptable!str5 = TurnValue(!DESCA)
    
    temptable!val3 = !T_IN + !T_OUT
    temptable!Val10 = 2
    temptable.Update
    sourcetable.MoveNext
Loop
End With
End If
    
'   ›’Ì·Ï «·›Ì“«

cString = "SELECT  SUM(FILE6_20H.VISA) AS TVISA ,SUM(FILE6_20H.VISA3) AS TVISA3 , FILE6_20H.CODEVISA, FILE6_20H.CODEVISA3, VISA.DESCA , VISA3.DESCA AS DESCA3 FROM    FILE6_20H LEFT JOIN VISA ON FILE6_20H.CODEVISA = VISA.CODE  LEFT JOIN VISA AS VISA3 ON FILE6_20H.CODEVISA3 = VISA3.CODE " & _
            " WHERE FILE6_20H.VISA <> 0  AND FILE6_20H.date = " & DateSq(xDate.text)
If cBranchStore <> "" Then cString = cString & " AND STORE = " & MyParn(cBranchStore)
If cManBox <> "" Then cString = cString & " AND BOX = " & MyParn(cManBox)
cString = cString & " GROUP BY FILE6_20H.CODEVISA, VISA.DESCA ,FILE6_20H.CODEVISA3, VISA3.DESCA"
VisaTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (VisaTable.EOF And VisaTable.BOF) Then
    With VisaTable
    Do Until .EOF
        temptable.AddNew
        temptable!str11 = 7
        temptable!str12 = "≈Ã„«·Ï ›Ì“«"
        temptable!str3 = " ÌÊ„Ì… " & xDate.text
        temptable!str1 = cManBox
'       temptable!str2 = TurnValue(!BOXDESCA)
        temptable!str5 = TurnValue(!DESCA)
        temptable!val3 = !TVISA
        temptable!Val10 = 2
        temptable.Update
        
        If !TVISA3 <> 0 Then
        temptable.AddNew
        temptable!str11 = 7
        temptable!str12 = "≈Ã„«·Ï ›Ì“«"
        temptable!str3 = " ÌÊ„Ì… " & xDate.text
        temptable!str1 = cManBox
'       temptable!str2 = TurnValue(!BOXDESCA)
        temptable!str5 = TurnValue(!DESCA3)
        temptable!val3 = !TVISA3
        temptable!Val10 = 2
        temptable.Update
        End If
        
        .MoveNext
    Loop
    End With
End If

    
    
cString = "SELECT   FILE0_50.DESCA AS BOXDESCA , BOXMOVE.BOX , SUM(BOXMOVE.PLUS - BOXMOVE.MINUS ) AS BALBOX FROM         BOXMOVE INNER JOIN FILE0_50 ON FILE0_50.CODE = BOXMOVE.BOX WHERE  DATE = " & DateSq(salesfrm.xDate.text)
If cManBox <> "" Then cString = cString & " AND BOX = " & MyParn(cManBox)
cString = cString & " GROUP BY FILE0_50.DESCA ,BOXMOVE.BOX ORDER BY BOXMOVE.BOX "
sourcetable.Close
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If Not (sourcetable.EOF And sourcetable.BOF) Then
    With sourcetable
        Do Until sourcetable.EOF
            temptable.AddNew
            temptable!str11 = 6
            temptable!str12 = "≈Ã„«·Ï «·Õ“‰…"
            temptable!str3 = " ÌÊ„Ì… " & xDate.text
            temptable!str1 = !BOX
            temptable!str2 = !BOXDESCA
            temptable!str5 = "—’Ìœ «·Œ“‰…"
            
            temptable!val3 = !BalBox
            temptable!Val10 = 2
            temptable.Update
            sourcetable.MoveNext
        Loop
    End With
End If
    
contemp.BeginTrans
contemp.CommitTrans

REPORT1.Reset
FixPrinter REPORT1

REPORT1.ReportFileName = App.Path & "\Reports\TDAY.RPT"
REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
REPORT1.Destination = crptToWindow
REPORT1.WindowState = crptMaximized
REPORT1.Action = 1
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub FixAddress()
Dim loctable As New ADODB.Recordset
If lIsBranchStore Then
    loctable.Open "select * From Address where branch = " & MyParn(cBranch), con, adOpenStatic, adLockReadOnly
Else
    loctable.Open "select * From Address", con, adOpenStatic, adLockReadOnly
End If
If Not (loctable.EOF And loctable.BOF) Then
    cComp_Name = loctable!DESCA & ""
    cComp_address = loctable!Address & ""
    cComp_Phone = loctable!phone & ""
    cComp_Head1 = loctable!HEAD1 & ""
    cComp_Head2 = loctable!HEAD2 & ""
    nCountPrint = Val(loctable!COUNTPRINT & "")
    nLang_Boon = Val(loctable!lang & "")
End If
loctable.Close
Set loctable = Nothing
End Sub
Private Sub cmd_showdoc_Click_AMR()
    Dim pDocRet As String
    Dim dDateRet As Date
    Dim cPassword As String, aUser As Variant
    
    pDocRet = xdoc_ret.text
    xdoc_ret.text = pDocRet
    
    
    If Len(xdoc_ret.text) >= 11 Then
    
    
        xIsRet.Value = IIf(GetBoolean("SELECT ISRET FROM FILE6_20H WHERE DOC_NO = " & MyParn(xdoc_ret.text), con), 1, 0)
        dDateRet = DateValue(Format(Mid(xdoc_ret.text, 1, 2) & "-" & Mid(xdoc_ret.text, 3, 2) & "-" & Mid(xdoc_ret.text, 5, 2), "DD-MM-YYYY"))
        If DateDiff("D", dDateRet, sDateSales) > 30 And xIsRet.Value = 0 Then
            MsgBox " »Ê‰ „»Ì⁄«   «—ÌŒ… «ﬂ»— „‰ 30 ÌÊ„"
            Exit Sub
        End If

         
        ManagPass.pPassWord = ""
        ManagPass.Show 1
        cPassword = ManagPass.pPassWord
        
        If Trim(cPassword) = "" Then Exit Sub
        aUser = aGetDesca("SELECT CODE , DESCA FROM SUB_USER WHERE ( ( PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0 ) OR ( STOP2 = 0 AND PASSWORD2 = " & MyParn(cPassword) & "))", con)
        If UBound(aUser) > 0 Then
            lRetDoc = True
            xIsRet.Value = 1
            xusername_RET.Caption = aUser(2)
        Else
            Exit Sub
        End If
        SalesDoc.sDoc_no = xdoc_ret.text
        SalesDoc.Show 1
        If xinv_no.text <> "" Then
            cmdSave_Click
        End If
    End If
End Sub
Private Sub cmd_disc_Click()
Dim cPassword As String, aUser As Variant
If xPrinted.Value <> 0 Then Exit Sub

'    ManagPass.pPassWord = ""
'    ManagPass.Show 1
'    cPassword = Trim(ManagPass.pPassWord)
'    If cPassword = "" Then Exit Sub
'    aUser = aGetDesca("SELECT CODE , DESCA FROM SUB_USER WHERE ( ( PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0 ) OR ( STOP2 = 0 AND PASSWORD2 = " & MyParn(cPassword) & "))", con)
'    If UBound(aUser) > 0 Then
        'MsgBox "Ì„ﬂ‰  ”ÃÌ· «·Œ’„ "
        xDiscount.Locked = False
        xRate.Locked = False
        XUSERNAME_DISC.Caption = ""
        'xNOTE_DISC.Caption = ""
        'XDOC_DISC.Caption = ""
'    Else
'        Exit Sub
'    End If
End Sub
Private Sub myUndo()
If xdoc_no.Tag = DefineMode Then
    If Not openCardTable Then
        CmdNewInv_Click
    End If
Else
    If Not openCardTable(tbMode.tbFind, xdoc_no.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdNext_Click()
If Not openCardTable(tbMode.tbNext, xdoc_no.text) Then
    If Not openCardTable(tbMode.tblast) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdPrevious_Click()
If Not openCardTable(tbMode.tbPrevious, xdoc_no.text) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdFirst_Click()
If Not openCardTable(tbMode.tbFirst) Then
    myDefine
End If
End Sub
Private Sub CmdLast_Click()
If Not openCardTable Then
    myDefine
End If
End Sub
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pDoc_No As String = "", Optional bDefine As Boolean = True) As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
Dim nRecords As Long, nRecord As Long

If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pDoc_No) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass

'On Error GoTo myerror
cFilter = ""

If Not (cBranch = "00" And bopt2) Then
    If cBranchBox <> "" Then
        cFilter = cFilter & Tr(cFilter) & "FILE6_20H.BOX = " & MyParn(cBranchBox)
    ElseIf cBranch <> "" Then
        cFilter = cFilter & Tr(cFilter) & "FILE6_20H.BRANCH = " & MyParn(cBranch)
    End If
End If

If optclosed(1).Value Then
    cFilter = cFilter & Tr(cFilter) & "(FILE6_20H.isClosed = 0 OR FILE6_20H.PRINTED = 0)"
ElseIf optclosed(2).Value Then
    cFilter = cFilter & Tr(cFilter) & "(FILE6_20H.isClosed = 1 AND FILE6_20H.PRINTED = 1)"
End If

If bSuperSales Then
    If IsDate(xdate1.text) Or IsDate(xdate2.text) Then
        If IsDate(xdate1.text) Then
            cFilter = cFilter & Tr(cFilter) & "FILE6_20H.DATE " & IIf(IsDate(xdate2.text), " >= ", " = ") & DateSq(xdate1.text)
        End If
        
        If IsDate(xdate2.text) Then
            cFilter = cFilter & Tr(cFilter) & "FILE6_20H.DATE <= " & DateSq(xdate2.text)
        End If
    ElseIf chkDayBranch.Value = 1 Then
        cFilter = cFilter & Tr(cFilter) & "FILE6_20H.DATE = " & DateSq(sDateSales)
    ElseIf chkDay.Value = 1 Then
        cFilter = cFilter & Tr(cFilter) & "FILE6_20H.DATE = " & DateSq(Date)
    Else
        If xMonth.BoundText <> "" And xYear.BoundText <> "" Then
            cFilter = cFilter & Tr(cFilter) & "FILE6_20H.YEAR_MONTH = " & MyParn(xYear.BoundText & "-" & RetZero(xMonth.BoundText, 2))
        ElseIf xYear.BoundText <> "" Then
            cFilter = cFilter & Tr(cFilter) & "FILE6_20H.YEAR_MONTH LIKE " & MyParn(xYear.BoundText & "%")
        End If
    End If
Else
    cFilter = cFilter & Tr(cFilter) & "FILE6_20H.DATE = " & DateSq(sDateSales)
End If

If optLading(1).Value Then
    cFilter = "LADING_BILL.TYPE > 2 AND FILE6_20H.DATE_DELIVERY IS NULL"
ElseIf optLading(2).Value Then
    cFilter = cFilter & Tr(cFilter) & "LADING_BILL.TYPE > 2 AND FILE6_20H.DATE_DELIVERY IS NOT NULL"
ElseIf optLading(3).Value Then
    cFilter = "LADING_BILL.TYPE <=  2 AND (SELECT DOC_NO FROM FILE6_20H AS F WHERE F.SALES_RET = FILE6_20H.DOC_NO) IS NULL"
ElseIf optLading(3).Value Then
    cFilter = "LADING_BILL.TYPE <=  2 AND FILE6_20H.DOC_NO IN (SELECT F.SALES_RET FROM FILE6_20H AS F)"
End If

If sDoc_no <> "" Then
    cFilter = "FILE6_20H.DOC_NO = " & MyParn(sDoc_no)
End If

cFilter = cFilter & Tr(cFilter) & "FILE6_20H.ONLINE = 2"
cFilter = cFilter & Tr(cFilter) & "FILE6_20H.ISINVOICE = 0"

Set CardTable = New ADODB.Recordset
cString = "SELECT TOP 1 FILE6_20H.*," & _
          "FILE3_10.DESCA AS CLIENTDESCA," & _
          "FILE3_10.ISONEST," & _
          " LADING_BILL.DOC_NO AS LADING," & _
          " LADING_BILL.TYPE AS LADING_TYPE," & _
          " LADING_CODES.DESCA AS LADING_TYPE_DESCA" & _
          " FROM FILE6_20H " & _
          " inner join file3_10 on file3_10.code = FILE6_20H.code" & _
          " INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE" & _
          " LEFT JOIN LADING_BILL ON FILE6_20H.DOC_NO = LADING_BILL.DOC_NO" & _
          " LEFT JOIN LADING_CODES ON LADING_CODES.CODE = LADING_BILL.TYPE"

If pMode = tbMode.tbFirst Then
    cOrder = "Order by FILE6_20H.DOC_NO"
ElseIf pMode = tbMode.tblast Then
    cOrder = "Order by FILE6_20H.DOC_NO DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "FILE6_20H.Doc_no = " & MyParn(pDoc_No)
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "FILE6_20H.DOC_NO < " & MyParn(pDoc_No)
    cOrder = "Order by FILE6_20H.DOC_NO DESC"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "FILE6_20H.DOC_NO > " & MyParn(pDoc_No)
    cOrder = "Order by FILE6_20H.DOC_NO"
End If

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

If cFilter <> "" Then
    cString = cString & IIf(cWhere = "", " WHERE ", " AND ") & cFilter
End If

cString = cString & " " & cOrder

Set CardTable = myCmd(cString, con)

If (Not CardTable.EOF) Then
    myload
    openCardTable = True
End If
Me.MousePointer = vbNormal
Exit Function
myerror:
Me.MousePointer = vbNormal
MsgBox Err.Description
Err.Clear
End Function
Private Sub cmd_showdoc_Click()
    Dim pDocRet As String, cRetBranch As String
    Dim dDateRet As Date
    Dim cPassword As String, aUser As Variant
    
    If grid1.Rows > 2 Then
        MsgBox " ÌÊÃœ «’‰«› „”Ã·… ›Ï «·»Ê‰"
        Exit Sub
    End If
    
    
        
        
        
        
'        pDocRet = xdoc_ret.text
'        xdoc_ret.text = pDocRet
'        ManagPass.pPassWord = ""
'        ManagPass.Show 1
'        cPassword = ManagPass.pPassWord
'        If Trim(cPassword) = "" Then Exit Sub
'        aUser = aGetDesca("SELECT CODE , DESCA FROM SUB_USER WHERE ( ( PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0 ) OR ( STOP2 = 0 AND PASSWORD2 = " & MyParn(cPassword) & "))", con)
'
'        AddLod_Data cusername, 0, " › Õ »Ê‰ ··„— Ã⁄  ", con, xdoc_no.text, xDate.text, , xCodeDesca.Caption
'
'
'        If UBound(aUser) > 0 Then
'            lRetDoc = True
'            XISRET.Value = 1
'            xusername_RET.Caption = aUser(2)
'        Else
'            Exit Sub
'        End If
        
'        SalesDoc.sDoc_no = xdoc_ret.text
'        SalesDoc.Show 1
'        If xinv_no.text <> "" Then
'            cmdSave_Click
'        End If

Dim cmdLaning As ADODB.Command


sales_returnsfrm.sDoc_no_ret = xdoc_ret.text
sales_returnsfrm.sDoc_no = xdoc_no.text
sales_returnsfrm.Show
End Sub
Private Function doprint_online() As Boolean
Dim aHeader(2)
If Not myValid Then Exit Function
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

Dim loctable As New ADODB.Recordset, cString As String
cString = "SELECT FILE6_20.ITEM,FILE6_20.QUANT,FILE6_20.SKU  ,FILE6_20.PRICE,FILE1_10.DESCA AS ITEM_DESCA,FILE6_20.DESCA2 , FILE6_20.TOTAL,FILE6_20H.TIME, file1_10.modelfact0 , FACT.DESCA AS FACTDESCA , FILE1_10.COLOR , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.MODEL   , file6_20h.[name] , file6_20h.phone , file6_20h.address , file6_20h.street , file6_20h.Payment_Method , CITY , Shipping_City , ONLINE_DOC  , CHARGE1 , CHARGE2 , ship_no FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM LEFT JOIN FACT ON FACT.CODE = FILE1_10.[FACT]  "
cString = cString & turn(cString) & "FILE6_20.DOC_NO = " & MyParn(xdoc_no.text)
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
Do While Not loctable.EOF
    temptable.AddNew
    temptable!str16 = xDoc_no2.text
    temptable!str6 = loctable!ONLINE_DOC
    temptable!str12 = xStore.text
    temptable!Date1 = xDate.text
    
    temptable!STR7 = loctable!Name
    temptable!str8 = loctable!phone
    temptable!str9 = loctable!Shipping_City
    temptable!str9 = loctable!Shipping_City
    temptable!str14 = loctable!city
    temptable!str15 = loctable!Street
    
    temptable!str2 = loctable!modelfact0
    temptable!str3 = IIf(loctable!Desca2 <> "", loctable!Desca2, loctable!ITEM_DESCA & " " & loctable!color & " " & loctable!SCAL)
    
    temptable!val3 = loctable!Quant
    temptable!VAL1 = loctable!price
    temptable!val2 = loctable!TOTAL
    
    temptable!val5 = Val(xDiscount.text)
    temptable!VAL7 = Val(xTotal.text)
    temptable!val11 = Val(loctable!CHARGE1 & "") + Val(loctable!charge2 & "")
    temptable!val12 = Val(xTotal.text) + Val(loctable!CHARGE1 & "") + Val(loctable!charge2 & "")
    temptable!VAL6 = Val(xTotalQuant.Caption) - Val(xTotalQuant2.Caption)
    temptable!str10 = loctable!SHIP_NO
    temptable!STR20 = xdoc_no.text
    
    If Val(xTotal.text) + Val(loctable!CHARGE1 & "") + Val(loctable!charge2 & "") > 0 Then
        temptable!str11 = MyOnly(Val(xTotal.text) + Val(loctable!CHARGE1 & "") + Val(loctable!charge2 & ""))
    End If
    
    temptable!str4 = TurnValue(cId_NoDesca1)
    temptable!str5 = TurnValue(cId_NoDesca2)
    temptable!str18 = TurnValue(cHeadBoon)
    temptable!str17 = loctable!SKU
    
    temptable!VAL14 = i
    temptable.Update
    loctable.MoveNext
Loop

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If

contemp.BeginTrans
contemp.CommitTrans

Main.REPORT1.ReportFileName = App.Path & "\Reports\Print_SALES_ONLINE.rpt"
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.REPORT1.Action = 1

temptable.Close
Set temptable = Nothing
doprint_online = True
End Function
Private Function ValidQuant() As Boolean
Dim nType As Integer
For i = 1 To grid1.Rows - 1
    If grid1.ValueMatrix(i, 10) = 0 Then
    ElseIf nType = 0 Then
        nType = IIf(grid1.ValueMatrix(i, 10) > 0, 1, 2)
    Else
        If nType <> IIf(grid1.ValueMatrix(i, 10) > 0, 1, 2) Then Exit Function
    End If
Next
ValidQuant = True
End Function
Private Function ValidQuantValue(nValue As Double, Row As Long) As Boolean
Dim nType As Integer
For i = 1 To grid1.Rows - 1
    If IIf(i = Row, nValue, grid1.ValueMatrix(i, 10)) = 0 Then
    ElseIf nType = 0 Then
        If i <> Row Then
            nType = IIf(grid1.ValueMatrix(i, 10) > 0, 1, 2)
        Else
            nType = IIf(nValue > 0, 1, 2)
        End If
    Else
        If i <> Row Then
            If nType <> IIf(grid1.ValueMatrix(i, 10) > 0, 1, 2) Then Exit Function
        Else
            If nType <> IIf(nValue > 0, 1, 2) Then Exit Function
        End If
    End If
Next
ValidQuantValue = True
End Function
Private Function retRecords(pDoc_No, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pDoc_No <> "" Then
    cString = "SELECT Count(*) AS records," & _
               "SUM(CASE WHEN FILE6_20H.DOC_NO <= " & MyParn(pDoc_No) & " THEN 1 ELSE 0 END) AS record"
Else
    cString = "SELECT Count(*) AS records," & _
               "0 as record"
End If

cString = cString & " FROM FILE6_20H " & _
                    " INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE" & _
                    " LEFT JOIN LADING_BILL ON FILE6_20H.DOC_NO = LADING_BILL.DOC_NO"
If cFilter <> "" Then
    cString = cString & " WHERE " & cFilter
End If
Set loctable = myCmd(cString, con)
If Not loctable.EOF Then
    nRecords = loctable!RECORDS
    nRecord = Val(loctable!Record & "")
End If
End Function
Private Sub fixDateSales()
'If (Not bSuperSales) Then
Dim cString As New ChilkatStringBuilder
cString.Append "SELECT DSALES FROM DSALES"
cString.Append " WHERE BRANCH = " & MyParn(cBranch)
sDateSales = myFormat(myField(cString.GetAsString(), con))
'End If

If myFormat(sDateSales) > myFormat(Date) Then
    MsgBox " «—ÌŒ  «·ÃÂ«“ «ﬂ»— „‰  «—ÌŒ «·ÌÊ„"
    lblAlert.Caption = " «—ÌŒ  «·ÃÂ«“ «ﬂ»— „‰  «—ÌŒ «·ÌÊ„-€Ì— „”„ÊÕ »«· ⁄œÌ·"
    fmAlert.Visible = True
    bEdit = False
ElseIf myFormat(Date) <> sDateSales Then
    MsgBox " «—ÌŒ «·ÃÂ«“ „Œ ·› ⁄‰  «—ÌŒ «·„»Ì⁄«  ,  »—Ã«¡ „—«Ã⁄…  «—ÌŒ «·»Ì⁄"
    fmAlert.Visible = True
End If
End Sub
Private Sub xYear_Change()
If chkDay.Value = 1 Or chkDayBranch.Value = 1 Or IsDate(xdate1.text) Or IsDate(xdate2.text) Then Exit Sub
myUndo
End Sub
Private Sub xInv_No_GotFocus()
myGotFocus xinv_no
End Sub
Private Sub xInv_No_LostFocus()
myLostFocus xinv_no
End Sub
Private Sub xMosm_GotFocus()
myGotFocus xMosm
End Sub
Private Sub xMosm_LostFocus()
myLostFocus xMosm
If Not xMosm.MatchedWithList Then xMosm.BoundText = ""
End Sub
Private Sub xYear_GotFocus()
myGotFocus xYear
End Sub
Private Sub xYear_LostFocus()
myLostFocus xYear
If Not xYear.MatchedWithList Then xYear.BoundText = ""
End Sub
Private Sub xMonth_GotFocus()
myGotFocus xMonth
End Sub
Private Sub xMonth_LostFocus()
myLostFocus xMonth
If Not xMonth.MatchedWithList Then xMonth.BoundText = ""
End Sub
Private Sub xTotal_GotFocus()
myGotFocus xTotal
End Sub
Private Sub xDiscount_GotFocus()
myGotFocus xDiscount
End Sub
Private Sub xRate_GotFocus()
myGotFocus xRate
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xdoc_no
End Sub
Private Sub xDoc_no2_GotFocus()
myGotFocus xDoc_no2
End Sub
Private Sub xDoc_no2_LostFocus()
myLostFocus xDoc_no2
End Sub
Private Sub XCode2_GotFocus()
myGotFocus xCode2
End Sub
Private Sub xCODE2_LostFocus()
myLostFocus xCode2
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub Xcode_GotFocus()
myGotFocus xCode
End Sub
Private Sub Xcode_LostFocus()
myLostFocus xCode
End Sub
Private Sub xdate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub xDate_Delivery_GotFocus()
myGotFocus xdate_Delivery
End Sub
Private Sub xDate_Delivery_LostFocus()
myLostFocus xdate_Delivery
myValidDate xdate_Delivery
End Sub

Private Sub XSTORE_GotFocus()
myGotFocus xStore
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub xBox_GotFocus()
myGotFocus xBox
End Sub
Private Sub xBox_LostFocus()
myLostFocus xBox
If Not xBox.MatchedWithList Then xBox.BoundText = ""
End Sub
Private Sub xMAN_GotFocus()
myGotFocus xMan
End Sub
Private Sub xMan_LostFocus()
myLostFocus xMan
If Not xMan.MatchedWithList Then xMan.BoundText = ""
End Sub
Private Sub xship_no_GotFocus()
myGotFocus xship_no
End Sub
Private Sub xship_no_LostFocus()
myLostFocus xship_no
End Sub
Private Sub xship_GotFocus()
myGotFocus xship
End Sub
Private Sub xship_LostFocus()
myLostFocus xship
If Not xship.MatchedWithList Then xship.BoundText = ""
End Sub
Private Sub xdoc_ret_GotFocus()
myGotFocus xdoc_ret
End Sub
Private Sub xdoc_ret_LostFocus()
myLostFocus xdoc_ret
End Sub
