VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form sales_abdfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20355
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
   ScaleHeight     =   11085
   ScaleWidth      =   20355
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   900
      TabIndex        =   171
      Top             =   -45
      Width           =   3795
      Begin Threed.SSCommand SSCommand1 
         Height          =   510
         Left            =   1890
         TabIndex        =   172
         TabStop         =   0   'False
         Top             =   180
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   900
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
         Caption         =   "«” ⁄·«„ »«· ·Ì›Ê‰"
         ButtonStyle     =   3
         PictureAlignment=   12
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   510
         Left            =   45
         TabIndex        =   173
         TabStop         =   0   'False
         Top             =   180
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   900
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
         Caption         =   "«” ⁄·«„ »«·—’Ìœ"
         ButtonStyle     =   3
         PictureAlignment=   12
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame8 
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
      Left            =   4995
      TabIndex        =   80
      Top             =   1215
      Width           =   1320
      Begin VB.CheckBox xGift 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "»Ê‰ Âœ«Ì«"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   158
         TabStop         =   0   'False
         Top             =   180
         Width           =   1050
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
      Height          =   690
      Left            =   450
      TabIndex        =   161
      Top             =   1710
      Width           =   2940
      Begin VB.CheckBox chkSendServer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰ﬁ·  ·ﬁ«∆Ì"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   162
         TabStop         =   0   'False
         Top             =   270
         Value           =   1  'Checked
         Width           =   1140
      End
      Begin Threed.SSCommand cmdFixMobileSetting 
         Height          =   420
         Left            =   90
         TabIndex        =   169
         Top             =   180
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
         BackColor       =   16777215
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«—”«· «·„»Ì⁄« "
         ButtonStyle     =   3
      End
   End
   Begin VB.Frame fmPrint 
      BackColor       =   &H00FFFFFF&
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
      Height          =   645
      Left            =   -450
      TabIndex        =   50
      Top             =   495
      Visible         =   0   'False
      Width           =   1590
      Begin VB.CheckBox chkprint 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·€«¡ ÿ»«⁄…"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   180
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   225
         Value           =   1  'Checked
         Width           =   1275
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "Frame14"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Left            =   16155
      TabIndex        =   26
      Top             =   315
      Visible         =   0   'False
      Width           =   9285
      Begin VB.CheckBox xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "„€·ﬁ"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1665
         TabIndex        =   170
         TabStop         =   0   'False
         Top             =   1980
         Width           =   690
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
         Height          =   555
         Left            =   0
         TabIndex        =   164
         Top             =   0
         Visible         =   0   'False
         Width           =   3615
         Begin VB.CheckBox xSendRc 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   " „"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   135
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   180
            Width           =   465
         End
         Begin Threed.SSCommand cmdGo 
            Height          =   330
            Left            =   2160
            TabIndex        =   166
            TabStop         =   0   'False
            Top             =   135
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
            Left            =   630
            RightToLeft     =   -1  'True
            TabIndex        =   167
            Top             =   135
            Width           =   1470
         End
      End
      Begin VB.CheckBox xisbankahly 
         Alignment       =   1  'Right Justify
         Caption         =   "»‰ﬂ «Â·Ï"
         Height          =   240
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   159
         Top             =   4095
         Visible         =   0   'False
         Width           =   1005
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
         Left            =   -450
         TabIndex        =   139
         Top             =   -90
         Width           =   3570
         Begin VB.TextBox xdate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   945
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   180
            Width           =   1185
         End
         Begin Threed.SSCommand DAYALL1 
            Height          =   375
            Left            =   90
            TabIndex        =   141
            TabStop         =   0   'False
            Top             =   180
            Width           =   825
            _ExtentX        =   1455
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
            Picture         =   "sales_abd.frx":0000
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            PictureDisabledFrames=   1
            ShapeSize       =   1
            PictureDisabled =   "sales_abd.frx":2202
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Left            =   2610
            TabIndex        =   142
            Tag             =   "Color"
            Top             =   225
            Width           =   555
         End
      End
      Begin VB.TextBox xinv_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2025
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   990
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox xIsOffer 
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
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   129
         Top             =   3645
         Width           =   1005
      End
      Begin VB.Frame Frame4 
         Caption         =   "—’Ìœ «·’‰›"
         Height          =   690
         Left            =   4680
         TabIndex        =   83
         Top             =   3105
         Width           =   1365
         Begin VB.Label xBalance 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   45
            TabIndex        =   84
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "»ÕÀ ›Ï «·„” ‰œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   3645
         Width           =   6090
         Begin VB.TextBox xfilter 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3510
            RightToLeft     =   -1  'True
            TabIndex        =   65
            TabStop         =   0   'False
            ToolTipText     =   "»ÕÀ"
            Top             =   180
            Width           =   1410
         End
         Begin VB.TextBox xfilter2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   64
            TabStop         =   0   'False
            ToolTipText     =   "»ÕÀ"
            Top             =   180
            Width           =   1770
         End
         Begin VB.TextBox xfilter3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   63
            TabStop         =   0   'False
            ToolTipText     =   "»ÕÀ"
            Top             =   585
            Width           =   1770
         End
         Begin VB.TextBox xfilter4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2565
            RightToLeft     =   -1  'True
            TabIndex        =   62
            TabStop         =   0   'False
            ToolTipText     =   "»ÕÀ"
            Top             =   585
            Width           =   2355
         End
         Begin VB.Label Label8 
            Caption         =   "»«—ﬂÊœ"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   270
            Width           =   885
         End
         Begin VB.Label Label31 
            Caption         =   "—ﬁ„ „ÊœÌ· "
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1890
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   180
            Width           =   885
         End
         Begin VB.Label Label32 
            Caption         =   "„’‰⁄"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1980
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   630
            Width           =   885
         End
         Begin VB.Label Label33 
            Caption         =   "«”„ «·’‰› "
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4995
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   630
            Width           =   885
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
         Left            =   4320
         TabIndex        =   53
         Top             =   1980
         Width           =   4695
         Begin Threed.SSCommand LOOKPH 
            CausesValidation=   0   'False
            Height          =   600
            Left            =   90
            TabIndex        =   54
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
            Picture         =   "sales_abd.frx":431E
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
            TabIndex        =   55
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
            TabIndex        =   56
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
            TabIndex        =   57
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
      Begin VB.CheckBox XISNODEL 
         Caption         =   "„— Ã⁄ ⁄—÷"
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
         Height          =   330
         Left            =   810
         TabIndex        =   52
         Top             =   2835
         Width           =   1230
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
         Left            =   2745
         TabIndex        =   47
         Top             =   1170
         Visible         =   0   'False
         Width           =   195
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
         Left            =   3555
         TabIndex        =   45
         Top             =   2070
         Width           =   195
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
         Left            =   405
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   2070
         Width           =   1005
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2430
         MaxLength       =   10
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   765
         Width           =   1185
      End
      Begin VB.PictureBox pctBarCode 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Left            =   6525
         ScaleHeight     =   118
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   151
         TabIndex        =   28
         Top             =   1305
         Visible         =   0   'False
         Width           =   2325
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   2970
         TabIndex        =   48
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
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
      Begin MSAdodcLib.Adodc data1 
         Height          =   330
         Left            =   315
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
      Begin MSAdodcLib.Adodc DATA3 
         Height          =   330
         Left            =   2475
         Top             =   1755
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
      Begin MSAdodcLib.Adodc data4 
         Height          =   330
         Left            =   405
         Top             =   1575
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
         Left            =   1665
         Top             =   1530
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
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   675
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   3195
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo xBoxVISA 
         Height          =   315
         Left            =   5805
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   315
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
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
      Begin MSAdodcLib.Adodc data12 
         Height          =   330
         Left            =   5040
         Top             =   1080
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
         Left            =   675
         Top             =   1350
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
         Left            =   4860
         Top             =   900
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
      Begin Threed.SSCommand cmdShowPhones 
         Height          =   510
         Left            =   0
         TabIndex        =   168
         TabStop         =   0   'False
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "sales_abd.frx":687C
         Caption         =   "⁄—÷ «·⁄„·«¡"
         ButtonStyle     =   2
         PictureAlignment=   9
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin VB.Label XVISA2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   0
         TabIndex        =   137
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label XPOINT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   90
         TabIndex        =   136
         Top             =   1170
         Width           =   915
      End
      Begin VB.Label xdoc_disc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   0
         TabIndex        =   122
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label xNote_disc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   4005
         TabIndex        =   121
         Top             =   270
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label XBRANCH 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„” Œœ„"
         Height          =   285
         Left            =   0
         TabIndex        =   86
         Top             =   0
         Width           =   555
      End
      Begin VB.Label xlate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   360
         TabIndex        =   73
         Top             =   2385
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„Œ“‰"
         Height          =   270
         Left            =   2805
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   3285
         Width           =   540
      End
      Begin VB.Label XTOTAL2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   555
         Left            =   4950
         TabIndex        =   60
         Top             =   855
         Width           =   2535
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "«·≈Ã„«·Ï »«·„” Â·ﬂ"
         Height          =   540
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   90
         Width           =   1050
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì·"
         Height          =   270
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   825
         Width           =   450
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   90
         TabIndex        =   42
         Top             =   780
         Width           =   2310
      End
      Begin VB.Label xcard_doc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   1440
         TabIndex        =   27
         Top             =   405
         Visible         =   0   'False
         Width           =   675
      End
   End
   Begin VB.TextBox xDoc_No 
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
      Height          =   330
      Left            =   2610
      MaxLength       =   12
      TabIndex        =   128
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4695
      Left            =   135
      TabIndex        =   6
      Top             =   2430
      Width           =   15135
      _cx             =   26696
      _cy             =   8281
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3210
      Left            =   0
      ScaleHeight     =   3210
      ScaleWidth      =   20355
      TabIndex        =   88
      Top             =   7410
      Width           =   20355
      Begin Threed.SSCommand cmdAddOffer 
         Height          =   555
         Left            =   3645
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   2115
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "sales_abd.frx":8D85
         Caption         =   "«÷«›… ⁄—÷"
         ButtonStyle     =   2
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_abd.frx":B1E5
      End
      Begin VB.Frame frPayment 
         BackColor       =   &H00FFFFFF&
         Height          =   1995
         Left            =   7560
         TabIndex        =   111
         Top             =   0
         Width           =   2850
         Begin Threed.SSCommand xcash 
            Height          =   375
            Left            =   135
            TabIndex        =   112
            Top             =   225
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   661
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TagVariant      =   "«Œ «— «·Œ«„…"
            Alignment       =   4
            ButtonStyle     =   3
         End
         Begin VB.Label xvisa 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000040&
            Height          =   375
            Left            =   135
            TabIndex        =   119
            Top             =   630
            Width           =   1275
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "”œ«œ ‰ﬁœÏ"
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
            Left            =   1485
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Top             =   270
            Width           =   780
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "”œ«œ ›Ì“«"
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
            Left            =   1485
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   675
            Width           =   720
         End
         Begin VB.Label xPay 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000040&
            Height          =   375
            Left            =   135
            TabIndex        =   116
            Top             =   1035
            Width           =   1275
         End
         Begin VB.Label xRest 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000040&
            Height          =   375
            Left            =   135
            TabIndex        =   115
            Top             =   1440
            Width           =   1275
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„œ›Ê⁄ ‰ﬁœÌ"
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
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   1080
            Width           =   1035
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "«·»«ﬁÌ «·‰ﬁœÌ"
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
            Left            =   1485
            RightToLeft     =   -1  'True
            TabIndex        =   113
            Top             =   1485
            Width           =   990
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid GrdOffer 
         Height          =   1860
         Left            =   3645
         TabIndex        =   120
         Top             =   180
         Width           =   3840
         _cx             =   6773
         _cy             =   3281
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
         Cols            =   2
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
         WordWrap        =   -1  'True
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
      Begin VB.Frame Frame6 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   123
         Top             =   1440
         Width           =   3525
         Begin Threed.SSCommand cmdFirst 
            Height          =   420
            Left            =   2610
            TabIndex        =   124
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
            Picture         =   "sales_abd.frx":D447
            Caption         =   "√Ê·"
            ButtonStyle     =   3
            PictureAlignment=   10
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "sales_abd.frx":F5EE
         End
         Begin Threed.SSCommand cmdPrevious 
            Height          =   420
            Left            =   1710
            TabIndex        =   125
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
            Picture         =   "sales_abd.frx":11635
            Caption         =   "”«»ﬁ"
            ButtonStyle     =   3
            PictureAlignment=   10
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "sales_abd.frx":13720
         End
         Begin Threed.SSCommand cmdNext 
            Height          =   420
            Left            =   855
            TabIndex        =   126
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
            Picture         =   "sales_abd.frx":1571A
            Caption         =   "·«Õﬁ"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "sales_abd.frx":1782B
         End
         Begin Threed.SSCommand cmdLast 
            Height          =   420
            Left            =   45
            TabIndex        =   127
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
            Picture         =   "sales_abd.frx":19825
            Caption         =   "√ŒÌ—"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "sales_abd.frx":1BA49
         End
      End
      Begin VB.Frame FrOffer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„— Ã⁄ Ê«” »œ«· ⁄—Ê÷"
         Height          =   1365
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   90
         Width           =   3480
         Begin VB.TextBox xdoc_ret 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   90
            MaxLength       =   12
            TabIndex        =   90
            Top             =   315
            Width           =   2850
         End
         Begin Threed.SSCommand cmdRefund 
            Height          =   510
            Left            =   1575
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   765
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   900
            _Version        =   196610
            ForeColor       =   0
            BackColor       =   16777215
            PictureFrames   =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "sales_abd.frx":1DB1A
            Caption         =   "„— Ã⁄"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            ShapeSize       =   1
            PictureDisabled =   "sales_abd.frx":1FFC5
         End
         Begin Threed.SSCommand cmdReplace 
            Height          =   510
            Left            =   90
            TabIndex        =   132
            TabStop         =   0   'False
            Top             =   765
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   900
            _Version        =   196610
            ForeColor       =   0
            BackColor       =   16777215
            PictureFrames   =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "sales_abd.frx":222AC
            Caption         =   "«” »œ«·"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            ShapeSize       =   1
            PictureDisabled =   "sales_abd.frx":24757
         End
         Begin VB.Label Label14 
            BackColor       =   &H00FFFFFF&
            Caption         =   "—ﬁ„"
            Height          =   285
            Left            =   3060
            TabIndex        =   135
            Top             =   360
            Width           =   375
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
         Height          =   2535
         Left            =   10440
         TabIndex        =   100
         Top             =   135
         Width           =   4785
         Begin VB.CheckBox xIsRet 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "”„«Õ »„— Ã⁄"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   174
            Top             =   1980
            Width           =   1320
         End
         Begin VB.TextBox xDiscount_add_rate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00400000&
            Height          =   330
            Left            =   2070
            MaxLength       =   10
            TabIndex        =   154
            Top             =   1260
            Width           =   555
         End
         Begin VB.TextBox xRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00400000&
            Height          =   330
            Left            =   2070
            MaxLength       =   10
            TabIndex        =   101
            Top             =   1620
            Width           =   555
         End
         Begin VB.TextBox xDiscount_add 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00400000&
            Height          =   330
            Left            =   2655
            MaxLength       =   15
            TabIndex        =   148
            Top             =   1260
            Width           =   870
         End
         Begin VB.TextBox xDiscount 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00400000&
            Height          =   330
            Left            =   2655
            MaxLength       =   15
            TabIndex        =   102
            Top             =   1620
            Width           =   870
         End
         Begin Threed.SSCommand cmd_disc 
            Height          =   330
            Left            =   1665
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   1620
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   582
            _Version        =   196610
            ForeColor       =   0
            BackColor       =   16777215
            Enabled         =   0   'False
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
         Begin Threed.SSCommand cmdDiscount_rate_add 
            Height          =   330
            Left            =   1665
            TabIndex        =   155
            TabStop         =   0   'False
            Top             =   1260
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   582
            _Version        =   196610
            ForeColor       =   0
            BackColor       =   16777215
            Enabled         =   0   'False
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
         Begin VB.Shape Shape1 
            Height          =   330
            Left            =   2070
            Top             =   1980
            Width           =   1455
         End
         Begin VB.Label lblOffer 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "«·’«›Ì"
            Height          =   270
            Left            =   1350
            RightToLeft     =   -1  'True
            TabIndex        =   157
            Top             =   945
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label xtotalitem_offer 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   90
            TabIndex        =   156
            Top             =   900
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Œ’„ «÷«›Ì"
            Height          =   270
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   153
            Top             =   1305
            Width           =   945
         End
         Begin VB.Label xDiscount_offer_rate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2070
            TabIndex        =   152
            Top             =   900
            Width           =   555
         End
         Begin VB.Label xDiscount_offer 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2655
            TabIndex        =   151
            Top             =   900
            Width           =   870
         End
         Begin VB.Label xtotalitem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   90
            TabIndex        =   150
            Top             =   540
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label lblTotalitem 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "«·’«›Ì"
            Height          =   270
            Left            =   1350
            RightToLeft     =   -1  'True
            TabIndex        =   149
            Top             =   585
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label lblDiscount_offer 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Œ’„ ⁄—Ê÷"
            Height          =   270
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   945
            Width           =   990
         End
         Begin VB.Label xTotalItemDiscount_rate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2070
            TabIndex        =   144
            Top             =   540
            Width           =   555
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Œ’„ «’‰«›"
            Height          =   270
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   143
            Top             =   585
            Width           =   960
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "’«›Ì «·›« Ê—…"
            Height          =   270
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   2025
            Width           =   1095
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Œ’„ ›« Ê—…"
            Height          =   270
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   1665
            Width           =   900
         End
         Begin VB.Label xTotalItemDiscount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2655
            TabIndex        =   108
            Top             =   540
            Width           =   870
         End
         Begin VB.Label xTotalItemNoDiscount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2070
            TabIndex        =   107
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «’‰«›"
            Height          =   270
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   225
            Width           =   900
         End
         Begin VB.Label xTotalQuant 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   90
            TabIndex        =   105
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label lblQuant 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ ﬁÿ⁄"
            Height          =   270
            Left            =   1350
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   225
            Width           =   630
         End
         Begin VB.Label xTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2070
            TabIndex        =   163
            Top             =   1980
            Width           =   1455
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Index           =   0
         Left            =   90
         TabIndex        =   92
         Top             =   1980
         Width           =   3525
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "„— Ã⁄« "
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   135
            TabIndex        =   95
            Top             =   270
            Width           =   1050
         End
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "„»Ì⁄« "
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1260
            TabIndex        =   94
            Top             =   270
            Width           =   915
         End
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«·ﬂ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2565
            TabIndex        =   93
            Top             =   270
            Value           =   -1  'True
            Width           =   690
         End
      End
      Begin VB.Frame frClosed 
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
         Height          =   690
         Left            =   7515
         TabIndex        =   96
         Top             =   1980
         Width           =   2895
         Begin VB.OptionButton optclosed 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "„› ÊÕ…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   990
            TabIndex        =   99
            Top             =   270
            Width           =   915
         End
         Begin VB.OptionButton optclosed 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "„€·ﬁ…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   135
            TabIndex        =   98
            Top             =   270
            Width           =   735
         End
         Begin VB.OptionButton optclosed 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«·ﬂ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2115
            TabIndex        =   97
            Top             =   270
            Value           =   -1  'True
            Width           =   690
         End
      End
   End
   Begin VB.Frame Frame99 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   17280
      TabIndex        =   16
      Top             =   5715
      Visible         =   0   'False
      Width           =   3300
      Begin VB.CommandButton cmd_showdoc 
         Caption         =   "⁄—÷  ›’Ì·Ï «·»Ê‰"
         Height          =   450
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   270
         Width           =   2310
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "„— Ã⁄ "
         Height          =   270
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   147
         Top             =   1080
         Width           =   525
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   146
         Top             =   1035
         Width           =   1005
      End
      Begin VB.Label xusername_RET 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   2610
         TabIndex        =   21
         Top             =   990
         Width           =   1950
      End
      Begin VB.Label XUSERNAME_DISC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   2610
         TabIndex        =   20
         Top             =   1350
         Width           =   1950
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "»Ê‰ «·„»Ì⁄« "
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
         Left            =   675
         TabIndex        =   19
         Top             =   765
         Width           =   960
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„— Ã⁄"
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
         Left            =   3195
         TabIndex        =   18
         Top             =   720
         Width           =   735
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
      Height          =   600
      Left            =   3465
      TabIndex        =   58
      Top             =   1215
      Width           =   1500
      Begin VB.CheckBox xPrinted 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "»Ê‰ „ÿ»Ê⁄"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   180
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   225
         Width           =   1185
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
      Height          =   1185
      Index           =   4
      Left            =   6345
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   1215
      Width           =   1275
      Begin Threed.SSCommand cmdSave 
         Height          =   465
         Left            =   45
         TabIndex        =   39
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   820
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "sales_abd.frx":26A48
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_abd.frx":2936D
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   465
         Left            =   45
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   630
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   820
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
         Picture         =   "sales_abd.frx":2BBC1
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_abd.frx":2DD21
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
      Left            =   10170
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   0
      Width           =   5055
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   4050
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   135
         Width           =   960
         _ExtentX        =   1693
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
         Picture         =   "sales_abd.frx":3000E
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_abd.frx":323D9
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   3060
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   135
         Width           =   960
         _ExtentX        =   1693
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
         Picture         =   "sales_abd.frx":34482
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_abd.frx":3648A
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   2055
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   135
         Width           =   960
         _ExtentX        =   1693
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
         Picture         =   "sales_abd.frx":38441
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_abd.frx":3ABDD
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   135
         Width           =   960
         _ExtentX        =   1693
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
         Picture         =   "sales_abd.frx":3D071
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   510
         Left            =   1035
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   135
         Width           =   960
         _ExtentX        =   1693
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
         Picture         =   "sales_abd.frx":3F394
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "sales_abd.frx":4170A
      End
   End
   Begin VB.Frame Frame17 
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
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   0
      Width           =   5415
      Begin Threed.SSCommand cmd_tsal 
         Height          =   510
         Left            =   1665
         TabIndex        =   30
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
         Picture         =   "sales_abd.frx":4388D
         Caption         =   " „»Ì⁄«  «·ÌÊ„ "
         Alignment       =   4
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_abd.frx":45F5F
      End
      Begin Threed.SSCommand cmdPrintDay 
         Height          =   510
         Left            =   45
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   135
         Width           =   1590
         _ExtentX        =   2805
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
         Picture         =   "sales_abd.frx":48443
         Caption         =   " «€·«ﬁ «·ÌÊ„ "
         Alignment       =   4
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_abd.frx":4AAD9
      End
      Begin Threed.SSCommand cmdPrintDayTotal 
         Height          =   510
         Left            =   3240
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   135
         Width           =   2130
         _ExtentX        =   3757
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
         Picture         =   "sales_abd.frx":4CFBD
         Caption         =   "ÿ»«⁄… ≈Ã„«·Ì «·ÌÊ„ "
         Alignment       =   4
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_abd.frx":4F61F
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   225
      Top             =   135
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   -765
      Top             =   -225
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
      Left            =   3015
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
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1080
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   2
      RTSEnable       =   -1  'True
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   1710
      Top             =   450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7650
      TabIndex        =   9
      Top             =   720
      Width           =   7575
      Begin VB.TextBox xcard_disc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5040
         MaxLength       =   200
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1500
      End
      Begin VB.TextBox xDoc_no2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5625
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   915
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
         TabIndex        =   7
         Top             =   2460
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3150
         MaxLength       =   200
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   900
         Width           =   3390
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   900
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   1230
      End
      Begin MSDataListLib.DataCombo xBox 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   540
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xMan 
         Height          =   315
         Left            =   3150
         TabIndex        =   2
         Top             =   540
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label xType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3150
         TabIndex        =   134
         Top             =   180
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„” Œœ„"
         Height          =   270
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   945
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬁÌ„… «·ﬂ«— "
         Height          =   270
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1305
         Width           =   870
      End
      Begin VB.Label xcard_value 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   90
         TabIndex        =   25
         Top             =   1260
         Width           =   2025
      End
      Begin VB.Label xcard_desca 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   3150
         TabIndex        =   24
         Top             =   1260
         Width           =   1860
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬂ«—  Œ’„"
         Height          =   270
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1305
         Width           =   810
      End
      Begin VB.Label xUserName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   900
         Width           =   2025
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·»«∆⁄"
         Height          =   270
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   585
         Width           =   405
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ „” ‰œ"
         Height          =   285
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   10
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
         TabIndex        =   14
         Top             =   2430
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label xtime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·«ÕŸ« "
         Height          =   270
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   945
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·Œ“‰…"
         Height          =   270
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   585
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
         Height          =   270
         Left            =   2265
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   270
         Width           =   510
      End
      Begin VB.Label xdoc_replace 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   4050
         TabIndex        =   133
         Top             =   180
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label XSALES_RET 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   4050
         TabIndex        =   130
         Top             =   180
         Visible         =   0   'False
         Width           =   1545
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   74
      Top             =   10620
      Width           =   20355
      _ExtentX        =   35904
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
         TabIndex        =   75
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
         TabIndex        =   76
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
         TabIndex        =   77
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
         TabIndex        =   78
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
         TabIndex        =   79
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
   Begin MSAdodcLib.Adodc DATA13 
      Height          =   330
      Left            =   810
      Top             =   -135
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   160
      Top             =   7215
      Visible         =   0   'False
      Width           =   20355
      _ExtentX        =   35904
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Frame Frame9 
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
      Left            =   3420
      TabIndex        =   81
      Top             =   1755
      Width           =   2895
      Begin Threed.SSCommand xphone 
         Height          =   465
         Left            =   45
         TabIndex        =   82
         Top             =   135
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   820
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   -2147483641
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
         Picture         =   "sales_abd.frx":51B03
         ButtonStyle     =   2
         PictureAlignment=   1
         BevelWidth      =   0
      End
   End
End
Attribute VB_Name = "sales_abdfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temptable As New ADODB.Recordset, cMsgDisplay As String, nDayDiscout As Double, cDescDiscout As String
Dim lManSales As Boolean, nCountPrint As Double
Dim bAct As Boolean
Dim cFilter As String
Dim bIg As Boolean
Dim bStopCell As Boolean
Public sDoc_no As String, lSave As Boolean, sUserDisc As String, lIsPrice_2 As Boolean, lIsdisc2   As Boolean, lDiscModelRet As Boolean
Public subUserCode As String, subUserName As String
Dim cFilterLook As String
Dim nRound As Long, cList1 As String, lIsBoxOnline As Boolean
Dim CardTable As ADODB.Recordset, cId_NoDesca1 As String, cId_NoDesca2 As String
Dim cDefClient As String, cDefClientDesca As String
Dim oSearchItem As New Search_abd, oSearchDoc   As New Search_abd
Dim cQrCode As New ClsQrCode
Dim bedit As Boolean
Dim bEditRecord As Boolean
Dim bIgClick As Boolean
Dim frmOffer As New offerAddfrm
Dim frmOfferReplace As New offerReplacefrm
Dim oSalesRefund As New sales_refundfrm

Dim con As New ADODB.Connection
Dim formMode
Const LoadMode = 0, DefineMode = 1

Dim nUser As Integer
Private Enum enUser
Admin = 3
MainUser = 2
Super = 1
User = 0
End Enum
Sub ItemsLookup()
ItemsLookupAll Me, oSearchItem
End Sub
Function myreplace(Optional Row As Long = -1, Optional bOffer As Boolean, Optional bReplace As Boolean) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(aInsert, "CODE", addstring(cDefClient))
aInsert = AddFlag(aInsert, "STORE", addstring(xStore.BoundText))
aInsert = AddFlag(aInsert, "[NOTES]", addstring(xNotes.text))
aInsert = AddFlag(aInsert, "[RATE]", Val(xRate.text))
aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(xDiscount.text))
aInsert = AddFlag(aInsert, "[DISCOUNT_ADD]", Val(xDiscount_add.text))
aInsert = AddFlag(aInsert, "[CASH]", Val(xcash.Caption))
aInsert = AddFlag(aInsert, "[BOX]", addstring(xBox.BoundText))
aInsert = AddFlag(aInsert, "[PAY]", Val(xPay.Caption))
aInsert = AddFlag(aInsert, "[LATE]", "0")
aInsert = AddFlag(aInsert, "[REST]", Val(xRest.Caption))
aInsert = AddFlag(aInsert, "[BRANCH]", addstring(XBRANCH.Caption))

If bReplace Then
    aInsert = AddFlag(aInsert, "[DOC_REPLACE]", addstring(frmOfferReplace.sDoc_Offer))
    If grid1.Rows < 3 And frmOfferReplace.xMan.Caption <> "" Then
        aInsert = AddFlag(aInsert, "[MAN]", addstring(frmOfferReplace.xMan.Caption))
    ElseIf xMan.MatchedWithList Then
        aInsert = AddFlag(aInsert, "[MAN]", addstring(xMan.BoundText))
    End If
ElseIf xMan.MatchedWithList Then
    aInsert = AddFlag(aInsert, "[MAN]", addstring(xMan.BoundText))
End If

aInsert = AddFlag(aInsert, "[MOSM]", addstring(xMosm.BoundText))
aInsert = AddFlag(aInsert, "[TIME]", addTime(xtime.Caption))
aInsert = AddFlag(aInsert, "[PHONE]", addstring(xphone.Caption))
aInsert = AddFlag(aInsert, "[USERNAME_DISC]", addstring(XUSERNAME_DISC.Caption))
aInsert = AddFlag(aInsert, "[USERNAME_RET]", addstring(xusername_RET.Caption))
aInsert = AddFlag(aInsert, "[SALES_RET]", addstring(XSALES_RET.Caption))
aInsert = AddFlag(aInsert, "[NOTE_DISC]", addstring(xNote_disc.Caption))
aInsert = AddFlag(aInsert, "[DOC_DISC]", addstring(xdoc_disc.Caption))
aInsert = AddFlag(aInsert, "[INV_NO]", addstring(xinv_no.text))
aInsert = AddFlag(aInsert, "[ISNODEL]", XISNODEL.Value)
aInsert = AddFlag(aInsert, "[CARD_DISC]", addstring(xcard_disc.text))
aInsert = AddFlag(aInsert, "[CARD_DESCA]", addstring(xcard_desca.Caption))
aInsert = AddFlag(aInsert, "[CARD_VALUE]", addstring(xcard_value.Caption))
aInsert = AddFlag(aInsert, "[CARD_DOC]", addstring(xcard_doc.Caption))
aInsert = AddFlag(aInsert, "[ISNEW]", "1")
On Error GoTo myError
con.BeginTrans
If xDoc_No.Tag = DefineMode Then
    dSalesDate = fnDateSales(con)
    If myFormat(dSalesDate) <> myFormat(xDate.text) Then
        MsgBox " „  €Ì— «· «—ÌŒ"
    End If
    
    xDate.text = myFormat_p(dSalesDate)
    xDoc_No.text = NewflagDoc(xDate.text, cBranchBox, con)
    xDoc_no2.text = Mid(xDoc_No.text, 7, 6)
    
    aInsert = AddFlag(aInsert, "[DATE]", addDate(xDate.text))
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.text))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cusername))
    aInsert = AddFlag(aInsert, "[USER_IP]", addstring(GetComputerName))
    con.Execute addInsert(aInsert, "FILE6_20H")
Else
    con.Execute addUpdate(aInsert, "FILE6_20H", "DOC_NO = " & addstring(xDoc_No.text))
End If

If bOffer Then
    myreplaceGrdOffer
    UpdateDiscount xDoc_No.text, con
ElseIf bReplace Then
    myreplaceGrdReplace
    UpdateDiscount xDoc_No.text, con
    If Not sendInvoicesNew(, True, xDoc_No.text) Then
        MsgBox " ⁄–— ‰ﬁ· «·»Ì«‰«  ⁄·Ì «·”Ì—›—"
        Exit Function
    End If
Else
    myreplaceGrd Row
End If
con.CommitTrans
myreplace = True
Exit Function
myError:
'prog1.Visible = False
MsgBox Err.Description
If TransCount(con) > 0 Then
    con.RollbackTrans
End If
Err.Clear
End Function
Private Sub myLoadGrdOffer()

End Sub
Sub myProc()
'On Error GoTo myError
If ActiveControl.Name = grid1.Name Or ActiveControl.Name = LOOKPH.Name Or ActiveControl.Name = LOOKPH_BARCODE.Name Then
    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    If ActiveControl.Name = grid1.Name Then
        grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    ElseIf ActiveControl.Name = LOOKPH.Name Then
        grid1.TextMatrix(grid1.Row, 1) = ItemLookPh_pict.grid1.TextMatrix(ItemLookPh_pict.grid1.Row, 0)
    ElseIf ActiveControl.Name = LOOKPH_BARCODE.Name Then
        grid1.TextMatrix(grid1.Row, 1) = ItemLookPh.grid1.TextMatrix(ItemLookPh.grid1.Row, 0)
    End If
    If Not bNew Then
        If ActiveControl.Name = grid1.Name Then Unload oSearchItem Else Unload ItemLookPh
        CellPos 13, grid1.Row, grid1.col
    ElseIf ActiveControl.Name = LOOKPH_BARCODE.Name Then
        grid1.SetFocus
        grid1.Select grid1.Rows - 1, 1
        Unload ItemLookPh
    End If
    GrdDesc grid1.TextMatrix(grid1.Row, 1), grid1.Row
    Grid1_AfterEdit grid1.Row, grid1.col
ElseIf ActiveControl.Name = cmdInform.Name Or ActiveControl.Name = cmdOpen.Name Then
    xDoc_No.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
    myUndo
ElseIf ActiveControl.Name = cmdAddOffer.Name Then
    If myreplace(, True) Then
        MsgBox " „  «÷«›… «·⁄—÷ »‰Ã«Õ"
        Unload frmOffer
    
        If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
            If Not openCardTable Then myDefine
        End If
    End If
ElseIf ActiveControl.Name = cmdReplace.Name Then
    If myreplace(, , True) Then
        MsgBox " „ «” »œ«· «·⁄—÷ »‰Ã«Õ"
        If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
            If Not openCardTable Then myDefine
        End If
    End If
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Public Function myReplaceRefund(pDiscount_add As Double, pDiscount_total As Double) As Boolean
If Not MYVALID Then Exit Function
With oSalesRefund
Dim aInsert As Variant
aInsert = AddFlag(Empty, "CODE", addstring(cDefClient))
aInsert = AddFlag(aInsert, "STORE", addstring(xStore.BoundText))
aInsert = AddFlag(aInsert, "[DISCOUNT]", -1 * Val(.xdiscount_ret.Caption))
aInsert = AddFlag(aInsert, "[BOX]", addstring(xBox.BoundText))
aInsert = AddFlag(aInsert, "[BRANCH]", addstring(XBRANCH.Caption))
aInsert = AddFlag(aInsert, "[PHONE]", addstring(.xphone.Caption))
If oSalesRefund.grid1.Rows > 1 Then
    aInsert = AddFlag(aInsert, "[MAN]", addstring(oSalesRefund.grid1.TextMatrix(1, 16)))
ElseIf xMan.MatchedWithList Then
    aInsert = AddFlag(aInsert, "[MAN]", addstring(xMan.BoundText))
End If
aInsert = AddFlag(aInsert, "[MOSM]", addstring(xMosm.BoundText))
aInsert = AddFlag(aInsert, "[TIME]", addTime(xtime.Caption))
aInsert = AddFlag(aInsert, "[SALES_RET]", addstring(.xDoc_No.Caption))
dSalesDate = fnDateSales(con)
If myFormat(dSalesDate) <> myFormat(xDate.text) Then
    MsgBox " „  €Ì— «· «—ÌŒ"
End If

xDate.text = myFormat_p(dSalesDate)
xDoc_No.text = NewflagDoc(xDate.text, cBranchBox, con)
xDoc_no2.text = Mid(xDoc_No.text, 7, 6)

aInsert = AddFlag(aInsert, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.text))
aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cusername))
aInsert = AddFlag(aInsert, "[USER_IP]", addstring(GetComputerName))

con.BeginTrans
On Error GoTo myError
con.Execute addInsert(aInsert, "FILE6_20H")

End With

With oSalesRefund.grid1
    For i = 1 To .Rows - 2
        aInsert = AddFlag(Empty, "ITEM", addstring(.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "QUANT", -1 * .ValueMatrix(i, 8))
        aInsert = AddFlag(aInsert, "PRICE", .ValueMatrix(i, 9))
        aInsert = AddFlag(aInsert, "PRICE_c", .ValueMatrix(i, 7))
        aInsert = AddFlag(aInsert, "COST", .ValueMatrix(i, 15))
        
        aInsert = AddFlag(aInsert, "MAN", addstring(.TextMatrix(i, 16)))
        
        aInsert = AddFlag(aInsert, "USER_IP", addstring(cIpName))
        aInsert = AddFlag(aInsert, "OFFER_NO", addstring(.TextMatrix(i, 11)))
        aInsert = AddFlag(aInsert, "DOC_OFFER", addstring(.TextMatrix(i, 12)))
        aInsert = AddFlag(aInsert, "DISCOUNT_OFFER", -1 * .ValueMatrix(i, 13))
        aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.text))
        con.Execute addInsert(aInsert, "FILE6_20")
    Next
End With

UpdateDiscount xDoc_No.text, con, pDiscount_add, pDiscount_total
If Not sendInvoicesNew(, True, xDoc_No.text) Then
    MsgBox " ⁄–— ‰ﬁ· «·»Ì«‰«  ⁄·Ì «·”Ì—›—"
    Exit Function
End If

con.CommitTrans
myReplaceRefund = True

MsgBox " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"

Unload oSalesRefund

If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then Exit Function
Exit Function
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub chkprint_Click()
addSetting "print", chkprint.Value, App.Path & "\other.txt"
End Sub
Private Function sendCash() As Boolean
End Function
Private Sub Command6_Click()
'mySendInvoices
End Sub

Private Sub chkSendServer_Click()
If chkSendServer.Value = 0 Then
    Dim conServer As New ADODB.Connection
    If openCon(conServer, LoadConServer, 2) = "ok" Then
        closeCon conServer
        chkSendServer.Value = 1
        MsgBox "«·”Ì—›— „ ’· »‰Ã«Õ"
        Exit Sub
    End If
End If
End Sub

Private Sub cmdAddOffer_Click()
If Not MYVALID(True) Then Exit Sub
Set frmOffer.myForm = Me
frmOffer.Show 1
End Sub

Private Sub cmdDiscount_rate_add_Click()
Dim cPassword As String, aUser As Variant
If xPrinted.Value <> 0 Then Exit Sub

Dim oUser As New sub_userfrm
Set oUser.myForm = Me
oUser.Show 1
If subUserCode <> "" Then
    MsgBox "Ì„ﬂ‰  ”ÃÌ· «·Œ’„ "
    
    xDiscount_add.Enabled = True
    xDiscount_add_rate.Enabled = True
End If
End Sub

Private Sub cmdFixMobileSetting_Click()
Dim nRecords As Long
If sendInvoicesNew(nRecords) Then
    MsgBox " „ «—”«· " & nRecords & " ”Ã· „»Ì⁄«  ··—∆Ì”Ì"
Else
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰ «—”«· «·„»Ì⁄«  ··—∆Ì”Ì"
End If
End Sub

Private Sub cmdPrintDayTotal_Click()
Dim oUser As New sub_userfrm
Set oUser.myForm = Me
oUser.Show 1

If subUserCode = "" Then Exit Sub

PrintTDay
End Sub
Private Sub ISPICT_Click()
    addSetting "ISPICT", ISPICT.Value, App.Path & "\other.txt"
End Sub
Private Sub CMD_TRANS_Click()
transfrm.Show
End Sub
Private Sub cmd_tsal_Click()
Dim oUser As New sub_userfrm
Set oUser.myForm = Me
oUser.Show 1

If subUserCode = "" Then Exit Sub

TDaySal.Show 1
End Sub
Private Sub CmdDel_Click()
Dim oUser As New sub_userfrm
Set oUser.myForm = Me
oUser.Show 1

If subUserCode = "" Then Exit Sub

If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) <> vbOK Then
    Exit Sub
End If

AddLod_Data cusername, 2, " Õ–› „»Ì⁄«  ", con, xDoc_No.text, xDate.text, , xCodeDesca.Caption

If xdoc_replace.Caption <> "" Or XSALES_RET.Caption <> "" Then
    Dim conServer As New ADODB.Connection
    If openCon(conServer, LoadConServer) <> "ok" Then
         MsgBox " ⁄–— «·œŒÊ· «·Ì «·›—⁄ «·—∆Ì”Ì ·Õ–› «·„·›"
         Exit Sub
    End If
    conServer.BeginTrans
    conServer.Execute "Delete  From FR6_20 where Doc_No = " & MyParn(xDoc_No.text)
    conServer.Execute "Delete  From FR6_20H where Doc_No = " & MyParn(xDoc_No.text)
    conServer.CommitTrans
End If

con.BeginTrans
con.Execute "Delete  From FILE6_20 where Doc_No = " & MyParn(xDoc_No.text)
con.Execute "Delete  From FILE6_20H where Doc_No = " & MyParn(xDoc_No.text)
con.CommitTrans
CmdNewInv_Click
Exit Sub
myError:
If TransCount(conServer) > 0 Then
    conServer.RollbackTrans
End If
If TransCount(con) > 0 Then
    con.RollbackTrans
End If
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdGo_Click()
If mySendReceipt(xDoc_No.text) Then
    myUndo
End If
End Sub
Private Sub CmdInform_Click()
CardLookup
End Sub
Private Sub cmdinformbarcode_Click()
    grdsales.Show 1
End Sub
Private Sub CmdNewInv_Click()
myDefine
On Error Resume Next
xMan.SetFocus
grid1.Select 1, 1
Err.Clear
End Sub
Private Sub cmdPrintDay_Click()
'If Not lIsBranchStore Then
'    DoBackUp
'End If


If CheckOpen > 0 Then Exit Sub

Dim nBalance As Double
Dim cString As String
cString = "SELECT COALESCE(SUM(BOXMOVE.PLUS - BOXMOVE.MINUS),0) " & _
           " FROM  FILE0_50 " & _
           " INNER JOIN BOXMOVE ON FILE0_50.CODE = BOXMOVE.BOX" & _
           " AND FILE0_50.TYPE = 1"

nBalance = Round(myField(cString, con), 2)

If nBalance <> 0 Then
    MsgBox " —’Ìœ Œ“‰ «·ﬂ«‘Ì—  " & nBalance & " »—Ã«¡  ÕÊÌ· «·‰ﬁœÌ… ﬁ»· «·√€·«ﬁ "
    Exit Sub
End If

If Not bOpt4 Then
    cString = "Select Option6 from USERS"
    managerfrm.sString = cString
    managerfrm.sFlag = ActiveControl.Name
    managerfrm.sFilter = "option4 = 1"
    managerfrm.Show 1
    If Not IsEmpty(aUser) Then
        If Not aUser(0) Then
            MsgBox "·Ì” ·œÌﬂ ’·«ÕÌ… ⁄—÷ ≈Ã„«·Ï «·„»Ì⁄« "
            Exit Sub
        End If
    Else
        MsgBox "·Ì” ·œÌﬂ ’·«ÕÌ… «€·«ﬁ «·ÌÊ„"
        Exit Sub
    End If
End If
If MsgBox(" —ÕÌ· Ê ≈€·«ﬁ «·Õ—ﬂ«  Ê «· ”ÃÌ·« ", vbYesNo + vbDefaultButton2) = vbYes Then
    Dim nRecords As Long
    If Not sendInvoicesNew(nRecords) Then
        MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰ «—”«· „” ‰œ«  «·„»Ì⁄«  »‰Ã«Õ «·Ì «·”Ì—›—"
    Else
        MsgBox " „ «—”«· " & nRecords & " ”Ã·"
    End If
     
    If CheckIsNew > 0 Then
        Exit Sub
    End If
        
    PostData
    CmdNewInv_Click
End If
End Sub

Private Sub cmdPrintGift_Click()
doprint_cash
End Sub
Private Sub cmdRefund_Click()
Dim conServer As New ADODB.Connection
If openCon(conServer, LoadConServer, 2) <> "ok" Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ ··œŒÊ· ··›—⁄ «·—∆Ì”Ì ·⁄„· „— Ã⁄"
    Exit Sub
End If
    
Set loctable = New ADODB.Recordset
Dim cString As String
cString = "select TOP 1 FILE6_20H.BRANCH," & _
          " FILE6_20H.PRINTED," & _
          " FILE6_20H.ISINVOICE," & _
          " FILE6_20H.ISRET," & _
          " FILE6_20H.TOTAL_ITEM," & _
          " FILE6_20H.DATE " & _
          " FROM file6_20h " & _
          " WHERE  FILE6_20H.DOC_NO = " & MyParn(xdoc_ret.text) & _
          " UNION ALL " & _
          " SELECT TOP 1 FR6_20H.BRANCH," & _
          " FR6_20H.PRINTED," & _
          " CAST(0 AS BIT) AS ISINVOICE," & _
          " FR6_20H.ISRET," & _
          " FR6_20H.TOTAL_ITEM," & _
          " FR6_20H.DATE " & _
          " FROM FR6_20H " & _
          " WHERE FR6_20h.DOC_NO = " & MyParn(xdoc_ret.text)
Set loctable = cmd(cString, conServer).Execute
If loctable.EOF Then
    MsgBox "›« Ê—… €Ì— „”Ã·…"
    Exit Sub
End If

If loctable!branch & "" = "00" Then
    MsgBox "›« Ê—… ›—⁄ —∆Ì”Ì"
    Exit Sub
End If

If Not loctable!PRINTED Then
    MsgBox "›« Ê—… „› ÊÕ…"
    Exit Sub
End If

If loctable!ISINVOICE Then
    MsgBox "›« Ê—… „Ã„⁄…"
    Exit Sub
End If

If loctable!TOTAL_ITEM <= 0 Then
    If loctable!TOTAL_ITEM = 0 Then
        MsgBox "›« Ê—… »œÊ‰ ﬁÌ„…"
    End If
    If loctable!TOTAL_ITEM < 0 Then
        MsgBox "›« Ê—… „— Ã⁄"
    End If
    Exit Sub
End If

If DateDiff("d", myFormat(loctable!Date), fnDateSales(con)) > 30 Then
    If Not loctable!ISRET Then
        MsgBox "„— ⁄·Ì «·›« Ê—… «ﬂÀ— „‰ 30 ÌÊ„"
        Exit Sub
    End If
End If

If Not MYVALID Then
    Exit Sub
End If

Dim oUser As New sub_userfrm
Set oUser.myForm = Me
oUser.Show 1

If subUserCode <> "" Then
    xusername_RET.Caption = subUserName
Else
    Cancel = True
    Exit Sub
End If

Set oSalesRefund.myForm = Me
oSalesRefund.sDoc_no = xdoc_ret.text
oSalesRefund.Show 1
End Sub

Private Sub cmdReplace_Click()
Dim conServer As New ADODB.Connection
If openCon(conServer, LoadConServer, 2) <> "ok" Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ ··œŒÊ· ··›—⁄ «·—∆Ì”Ì ·⁄„· „— Ã⁄"
    Exit Sub
End If


Dim loctable As New ADODB.Recordset
Dim cString As String

cString = "SELECT TOP 1 FILE6_20H.TOTAL_ITEM," & _
          " BRANCH," & _
          " ISINVOICE " & _
          " FROM file6_20h " & _
          " WHERE  FILE6_20H.DOC_NO = " & MyParn(xdoc_ret.text) & _
          " UNION ALL " & _
          " SELECT TOP 1 FR6_20H.TOTAL_ITEM," & _
          " BRANCH," & _
          " CAST(0 AS BIT) " & _
          " from FR6_20H " & _
          " where FR6_20H.DOC_NO = " & MyParn(xdoc_ret.text)

Set loctable = cmd(cString, conServer).Execute
If loctable.EOF Then
    MsgBox "«·›« Ê—… €Ì— „”Ã·…"
    Exit Sub
End If

If loctable!TOTAL_ITEM >= 0 Then
    MsgBox "«·›« Ê—… ·Ì”  „— Ã⁄"
    Exit Sub
End If

If loctable!branch & "" = "00" Then
    MsgBox "›« Ê—… ›—⁄ —∆Ì”Ì"
    Exit Sub
End If

If loctable!ISINVOICE Then
    MsgBox "›« Ê—… „Ã„⁄…"
    Exit Sub
End If


cString = "SELECT FILE6_20.DOC_NO  " & _
          " FROM FILE6_20 " & _
          " WHERE DOC_OFFER IS NOT NULL " & _
          " AND  DOC_NO = " & MyParn(xdoc_ret.text) & _
          " UNION ALL " & _
          " SELECT FR6_20.DOC_NO" & _
          " FROM FR6_20 " & _
          " WHERE FR6_20.DOC_OFFER IS NOT NULL " & _
          " AND  FR6_20.DOC_NO = " & MyParn(xdoc_ret.text)

Set loctable = cmd(cString, conServer).Execute
If loctable.EOF Then
    MsgBox "·« ÌÊÃœ ⁄—Ê÷ »«·›« Ê—…"
    Exit Sub
End If

cString = "SELECT FILE6_20H.DOC_NO," & _
          "FILE0_40.DESCA," & _
          "FILE6_20H.DATE," & _
          "FILE6_20H.PRINTED" & _
          " FROM FILE6_20H " & _
          " INNER JOIN FILE0_40 ON FILE6_20H.STORE = FILE0_40.CODE" & _
          " WHERE FILE6_20H.DOC_REPLACE = " & MyParn(xdoc_ret.text) & _
          " UNION ALL " & _
          " SELECT FR6_20H.DOC_NO," & _
          " BRANCH_FR.DESCA," & _
          " FR6_20H.DATE," & _
          " FR6_20H.PRINTED" & _
          " FROM FR6_20H " & _
          " INNER JOIN BRANCH_FR ON FR6_20H.STORE = BRANCH_FR.CODE" & _
          " WHERE FR6_20H.DOC_REPLACE = " & MyParn(xdoc_ret.text)
           
Set loctable = cmd(cString, conServer).Execute

If Not loctable.EOF Then
    MsgBox " „ ⁄„· «” »œ«· ·›« Ê—… «·„— Ã⁄ »—ﬁ„ " & loctable!doc_no & vbCrLf & _
           "·›—⁄ " & loctable!desca & vbCrLf & _
           "» «—ÌŒ " & myFormat_p(loctable!Date) & vbCrLf & _
           IIf(loctable!PRINTED, "„€·ﬁ…", "„› ÊÕ…")
    Exit Sub

    If DateDiff("d", myFormat(loctable!Date), fnDateSales(con)) > 0 Then
        If Not loctable!ISRET Then
            MsgBox "„— ⁄·Ì «·›« Ê—… ÌÊ„"
            Exit Sub
        End If
    End If
End If

cString = "SELECT FILE6_20H.DOC_NO," & _
          "FILE0_40.DESCA," & _
          "FILE6_20H.DATE," & _
          "FILE6_20H.PRINTED" & _
          " FROM FILE6_20H " & _
          " INNER JOIN FILE0_40 ON FILE6_20H.STORE = FILE0_40.CODE" & _
          " WHERE FILE6_20H.DOC_REPLACE = " & MyParn(xdoc_ret.text)
Set loctable = cmd(cString, conServer).Execute
If Not loctable.EOF Then
    MsgBox " „ ⁄„· «” »œ«· ·›« Ê—… «·„— Ã⁄ »—ﬁ„ " & loctable!doc_no & vbCrLf & _
           "»«·›—⁄ «·Õ«·Ì " & _
           "» «—ÌŒ " & myFormat_p(loctable!Date) & vbCrLf & _
           IIf(loctable!PRINTED, "„€·ﬁ…", "„› ÊÕ…") & _
           "Ê·„  —Õ· »⁄œ ··›—⁄ «·—∆Ì”Ì"
    Exit Sub
End If

closeCon conServer

If Not MYVALID Then Exit Sub
Set frmOfferReplace.myForm = Me
frmOfferReplace.sDoc_Offer = xdoc_ret.text
frmOfferReplace.Show 1
End Sub
Private Sub cmdSave_Click()

If grid1.Rows = 2 Then Exit Sub

If Not MYVALID Then Exit Sub

myDisplay "Total Qty  :" & Format(Val(xTotalQuant.Caption) - Val(xTotalQuant2.Caption), "#0"), "Total Price:" & Format(Val(xTotal.Caption), "#0.00")

lSave = False
Set CashReg.myForm = Me
CashReg.Show 1

If Not lSave Then Exit Sub


If Not mysave(, , True) Then Exit Sub

CmdNewInv_Click
On Error Resume Next
xMan.SetFocus
Err.Clear
End Sub
Private Function mySendReceipt(pDoc_no As String) As Boolean
Dim sError As String
If Not BranchReceipt Then Exit Function
If Not SendReceipt(pDoc_no, con, sError) Then
    MsgBox sError
Else
    Inform " „ «—”«· «·«Ì’«· " & pDoc_no & " »‰Ã«Õ ⁄·Ì «·„Êﬁ⁄ «·«·Ìﬂ —Ê‰Ì"
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

Private Sub cmdShowPhones_Click()
  phoneEditfrm.Show
End Sub

Private Sub CmdUndo_Click()
    Dim cString As String
    If cBranch <> "00" And False Then
        cString = "Select Option6 from USERS"
        managerfrm.sString = cString
        managerfrm.sFlag = ActiveControl.Name
        managerfrm.sFilter = "option6 = 1"
        managerfrm.Show 1
        If Not IsEmpty(aUser) Then
            If Not aUser(0) Then
                MsgBox "·Ì” ·œÌﬂ ’·«ÕÌ… "
                Exit Sub
            Else
    '           MsgBox " OK "
                myUndo
                grid1.ShowCell grid1.Rows - 1, 1
                grid1.Select grid1.Rows - 1, 1
            End If
        Else
            Exit Sub
        End If
    Else
        myUndo
        grid1.ShowCell grid1.Rows - 1, 1
        grid1.Select grid1.Rows - 1, 1
    End If
On Error Resume Next
grid1.SetFocus
End Sub
Private Sub cmdopen_Click()
    CardLookup " AND PRINTED = 0"
End Sub
Private Sub Command2_Click()
If cBranch <> "00" Then
    If InputBox(" -- ") = "654321" Then
        If InputBox(" -- ") = "123456" Then
        If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
            con.Execute "Delete  From FILE6_20 WHERE DOC_NO IN (SELECT DOC_NO FROM FILE6_20H WHERE BRANCH = " & MyParn(cBranch) & ")"
            con.Execute "Delete  From FILE6_20H WHERE BRANCH = " & MyParn(cBranch)
            Unload Me
        End If
        End If
    End If
End If
End Sub

Private Sub command31_Click()
End Sub
Private Sub DAYALL_Click()
'    cFilterLook = " FILE6_20H.date = " & DateSq(Format(xdate_open.Text, "DD-MM-YYYY"))
    myUndo
End Sub
Private Sub FIX_COST_Click()
    On Error GoTo myError
    con.CommandTimeout = 500
    cStr1 = " UPDATE    FILE6_20 SET  cost = (SELECT TOP 1 (COST ) FROM Q_ITEMCOST WHERE FILE6_20.ITEM = Q_ITEMCOST.ITEM AND Q_ITEMCOST.DATE <= FILE6_20H.DATE ORDER BY DATE DESC ) " & _
            " FROM      FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO where year(file6_20h.date) >= " & Year(Date) - 1
    con.Execute cStr1, nRec
    Inform "  „ Ÿ»ÿ  ﬂ·›… «·„»Ì⁄«  " & nRec
    If cBranch = "00" Then
        cStr1 = " UPDATE    FR6_20 SET  cost = (SELECT TOP 1 (COST ) FROM Q_ITEMCOST_FR WHERE FR6_20.ITEM = Q_ITEMCOST_FR.ITEM  AND FR6_20H.STORE = Q_ITEMCOST_FR.STORE AND Q_ITEMCOST_FR.DATE <= FR6_20H.DATE ORDER BY DATE DESC ) " & _
                " FROM      FR6_20 INNER JOIN FR6_20H ON FR6_20.DOC_NO = FR6_20H.DOC_NO where year(fR6_20h.date) >= " & Year(Date) - 1
        con.Execute cStr1, nRec
        Inform "  „ Ÿ»ÿ  ﬂ·›… «·„»Ì⁄«  «· ÊﬂÌ·« " & nRec
    End If
    
    cStr1 = " UPDATE    FILE1_60 SET  cost = (SELECT TOP 1 (COST ) FROM Q_ITEMCOST WHERE FILE1_60.ITEM = Q_ITEMCOST.ITEM AND Q_ITEMCOST.DATE <= FILE1_60H.DATE ORDER BY DATE DESC ) " & _
            " FROM      FILE1_60 INNER JOIN FILE1_60H ON FILE1_60.DOC_NO = FILE1_60H.DOC_NO where year(file1_60h.date) >= " & Year(Date) - 1
    con.Execute cStr1, nRec
    Inform "  „ Ÿ»ÿ  ﬂ·›… «· ÕÊÌ·« " & nRec
    Exit Sub
myError:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub Form_Activate()
If Not validOpen Then
    Unload Me
    Exit Sub
End If
If Not bAct Then
    If xDoc_No.Tag = LoadMode Then
        On Error Resume Next
        grid1.SetFocus
        'CellPos 13, 0, grid1.Cols - 1
    ElseIf xMan.Enabled Then
       xMan.SetFocus
    End If
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ActiveControl.Name = xMan.Name Then Exit Sub
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then KeyAscii = 0
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
ElseIf KeyCode = 116 And cmdSave.Enabled Then
    cmdSave_Click
End If
End Sub
Private Sub Form_KeyUp2(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
'If KeyCode = 117 Then  '
'     If grid1.Row <= grid1.Rows - 2 Then
'         grid1.Select grid1.Row, 10, grid1.Row, 10
'     Else
'         grid1.Select 1, 10, 1, 10
'     End If
' End If
'
' If KeyCode = 118 Then  '  «·”⁄—
'     grid1.Select 1, 11, 1, 11
' End If
'
' If KeyCode = 114 Then  '  ‰”Ì… Œ’„
'     xDiscount.SetFocus
' End If
'
''    If ActiveControl.Name <> GRID1.Name Then
''        If KeyCode = 116 Then   '  ”œ«œ ›« Ê—…
''            If Not IsPrinted Then cmdSave_Click
''        End If
''    End If
'
''    If KeyAscii = 27 Then xTotal.SetFocus   '  «·≈Ã„«·Ï
'
'Err.Clear
End Sub
Private Sub Form_Load()
'On Error GoTo myError
openCon con

Me.Caption = "„»»⁄«  " & myFormat_p(dSalesDate)

nRound = 0
HandleInit
fixDisplay

'cmdInform.Visible = Not bopt1
'Frame6.Visible = Not bopt1

'xBox.Enabled = False
'xStore.Enabled = (cBranch = "00")

'Dim aRet As Variant
'ÿlManSales = False

cDefClient = "0000"
cDefClientDesca = "⁄„Ì· ‰ﬁœÌ"

Set data12.Recordset = cmd("SELECT * FROM FILE6_25 WHERE ISSTOP = 0 ORDER BY DESCA ", con).Execute
Set xMan.RowSource = data12
xMan.ListField = "Desca"
xMan.BoundColumn = "Code"
If data12.Recordset.RecordCount = 1 Then
    xMan.BoundText = data12.Recordset!code
End If


Set data1.Recordset = cmd("SELECT * FROM FILE0_40 WHERE BRANCH = " & MyParn(cBranch), con).Execute
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"
xStore.BoundText = cBranchStore

Set data2.Recordset = cmd("SELECT * FROM MOSM ", con).Execute
Set xMosm.RowSource = data2
xMosm.ListField = "DESCA"
xMosm.BoundColumn = "MOSM"

Set data4.Recordset = cmd("SELECT * FROM FILE0_50", con).Execute
Set xBox.RowSource = data4
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"
xBox.BoundText = cBranchBox

Set grid1.DataSource = DATA11
Set GrdOffer.DataSource = DATA13

dSalesDate = fnDateSales(con)
If Not IsDate(dSalesDate) Then
    dSalesDate = Date
End If

If myFormat(dSalesDate) <> myFormat(Date) Then
    MsgBox " «—ÌŒ «·ÃÂ«“ „Œ ·› ⁄‰  «—ÌŒ «·„»Ì⁄«  ,  »—Ã«¡ „—«Ã⁄…  «—ÌŒ «·»Ì⁄"
End If

frClosed.Visible = nUser > enUser.User

If cBranchBox <> "" Then
    cFilter = retFilter
    CmdNewInv_Click
Else
    If Not openCardTable Then myDefine
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Resize()
Dim nHeight As Long
nHeight = Me.Height - (grid1.Top + 4100)
If nHeight > 3000 Then
    grid1.Height = nHeight
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

CardTable.Close
Set CardTable = Nothing

closeCon con

'Unload Search3
'Unload oSearchDoc
'Unload search32
Set sales_abdfrm = Nothing

'If sBranchCode <> "00" And bopt1 And Not lMainShow Then
'    End
'End If

'If CBRAsBranchCode <> "00" Then End
Err.Clear
End Sub

Private Sub GrdOffer_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If GrdOffer.Row = 0 Or (GrdOffer.Row = GrdOffer.Rows - 1 And GrdOffer.Rows > 2) Or GrdOffer.TextMatrix(GrdOffer.Row, 0) = "" Then
        Exit Sub
    End If
    
    If XSALES_RET.Caption <> "" Or xdoc_replace.Caption <> "" Then
        Exit Sub
    End If
    
    Dim oUser As New sub_userfrm
    Set oUser.myForm = Me
    oUser.Show 1
    
    If subUserCode = "" Then Exit Sub
    
    If MsgBox("Õ–› «·⁄—÷ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) <> vbOK Then
        Exit Sub
    End If
    
    Dim nDiscount As Long
    con.BeginTrans
    con.Execute "delete from file6_20 " & _
                " where doc_no = " & MyParn(xDoc_No.text) & _
                " and offer_no = " & GrdOffer.TextMatrix(GrdOffer.Row, 0)
    
    UpdateDiscount xDoc_No.text, con
    con.CommitTrans
    
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then myDefine
    End If
    
    MsgBox " „ Õ–› «·⁄—÷ »«·ﬂ«„· Ê ÕœÌÀ «·Œ’„"
End If
Exit Sub
myError:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Public Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
On Error GoTo myError
Dim bNew As Boolean
With grid1
If Not MYVALID(True) Then
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
    myLoadGrd
    If Row < grid1.Rows - 1 Then
        grid1.Select Row, col
    Else
        CellPos 13, grid1.Rows - 2, grid1.Cols - 1
    End If
    Exit Sub
End If


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
    If xDoc_No.Tag = DefineMode Then
        openCardTable tbMode.tbFind, xDoc_No.text
    ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myLoadGrd
    End If
    If bNew Then
        'bStopCell = True
        grid1.Select grid1.Rows - 1, 1
    End If
End If
End With
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
grid1_EnterCell
End Sub

Private Sub grid1_EnterCell()
With grid1
    If Not bEditRecord Then
        .Editable = flexEDNone
    ElseIf XSALES_RET.Caption <> "" Or xdoc_replace.Caption <> "" Then
        .Editable = flexEDNone
    ElseIf .ValueMatrix(grid1.Row, 19) > 0 Then
        .Editable = flexEDNone
    ElseIf (.col = 1 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "") Or (.col = 10) Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub grid1_GotFocus()
grid1_EnterCell
Err.Clear
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
'If xPrinted.Value = 1 And Not bopt2 Then Exit Sub
''If KeyCode = 115 Or (KeyCode = 13 And Shift = 2) Then xDiscount.SetFocus
''If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 And validRows(grid1.Row) Then
''    grid1.AddItem "", grid1.Row
''End If
'If KeyCode = 112 And xPrinted.Value = 0 And grid1.TextMatrix(grid1.Row, 1) = "" Then
''    Grid1.Row = Grid1.Rows - 1
''    Grid1.Col = 1
'    ItemsLookupAll Me, oSearchItem
''    VsModelLook2.Show 1
'End If
End Sub
Private Sub grid1_KeyDownEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
    'If xPrinted.Value = 1 And Not bopt2 Then Exit Sub
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'If grid1.Col = 12 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    KeyAscii = 0
End If

'If KeyAscii = 13 Then
'    If grid1.Row = grid1.Rows - 1 And grid1.Row > 1 And Trim(grid1.TextMatrix(grid1.Row, 1)) = "" Then
'        KeyAscii = 0
'        If Not IsPrinted Then cmdSave_Click
'        Err.Clear
'    End If
'End If
'If KeyAscii = 43 Then
'    KeyAscii = 0
'    If grid1.Row > 1 Then
'        grid1.Row = grid1.Row - 1
'        grid1.Col = 10
'    End If
'End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    'If Col = 12 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    CellPos KeyCode, Row, col
End If
End Sub
Private Sub grid1_LostFocus()
'ImgX1.Images.Clear
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
Dim sBarCode As String
Dim nSal As Double
Dim nRet As Double
Dim cPassword As String, aUser As Variant
With grid1
    If col = 10 Then
        If Not IsNumeric(.EditText) Then
            Cancel = True
            Exit Sub
        ElseIf Val(.EditText) < 0 Then
            For i = 1 To .Rows - 1
                If i <> Row Then
                    If .ValueMatrix(i, 10) > 0 Then
                        MsgBox "ﬂ„Ì… „— Ã⁄ ›Ï ›« Ê—… »ﬂ„Ì«  „ÊÃ»…"
                        Cancel = True
                        Exit Sub
                    End If
                End If
            Next
                            
            Dim oUser As New sub_userfrm
            oUser.sField = "ISRETURN"
            Set oUser.myForm = Me
            oUser.Show 1
            
            If subUserCode <> "" Then
                xusername_RET.Caption = subUserName
                MsgBox "Ì„ﬂ‰  ”ÃÌ· «·„— Ã⁄"
            Else
                Cancel = True
                Exit Sub
            End If
        ElseIf Val(.EditText) > 0 Then
            For i = 1 To .Rows - 1
                If i <> Row Then
                    If .ValueMatrix(i, 10) < 0 Then
                        MsgBox "ﬂ„Ì… „ÊÃ»… ›Ï ›« Ê—… „— Ã⁄"
                        Cancel = True
                        Exit Sub
                    End If
                End If
            Next
        ElseIf .ValueMatrix(Row, 10) > 0 And Val(.EditText) > 0 Then
            MsgBox "ÌÃ» Õ–› «·”Ã· ﬁ»·  ⁄œÌ·… „‰ „— Ã⁄ ≈·Ï „»Ì⁄« "
            Cancel = True
        ElseIf Val(.EditText) > 0 Then
            nBalance = fnBalance(grid1.TextMatrix(grid1.Row, 1), con, xStore.BoundText, xDate.text, grid1.TextMatrix(grid1.Row, grid1.Cols - 1))
            If nBalance < Val(.EditText) Then
                Inform "«·—’Ìœ ·« Ì”„Õ"
                Exit Sub
            End If
        End If
    ElseIf col = 1 Then
        If Not GrdDesc(.EditText, Row) Then
            Cancel = True
        Else
            If .EditText <> grid1.TextMatrix(Row, 1) Then
                .EditText = grid1.TextMatrix(Row, 1)
            End If
        End If
    End If
End With
End Sub
Private Sub cmdPrint_Click()
If grid1.Rows = 2 Then Exit Sub
If Val(xTotal.Caption) <> Val(xvisa.Caption) + Val(xcash.Caption) + Val(XPOINT.Caption) Then
    MsgBox "„—«Ã⁄… ”œ«œ «·»Ê‰ "
    Exit Sub
End If
If doprint_cash(True) Then SavePrint
End Sub

Private Sub optclosed_Click(Index As Integer)
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then
        myDefine
    End If
End If
End Sub

Private Sub optType_Click(Index As Integer)
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then
        myDefine
    End If
End If
End Sub

Private Sub SSCommand1_Click()
    grdsales.Show 1
End Sub

Private Sub SSCommand2_Click()
Set ItemLookPh.myForm = sales_abdfrm
ItemLookPh.Show
End Sub

Private Sub xcard_disc_LostFocus()
    'If xPrinted.Value = 0 Then CheckCardDisc
End Sub

Private Sub xcash_Click()
lSave = False
Set CashReg.myForm = Me
CashReg.Show 1

If Not lSave Then Exit Sub


If Not mysave(, , True) Then Exit Sub

CmdNewInv_Click
On Error Resume Next
xMan.SetFocus
Err.Clear
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 112 And xPrinted.Value = 0 Then CLIENTLOOKUP
End Sub
Private Sub xcodevisa_LostFocus()
'    con.Execute " UPDATE FILE6_20H SET ISNEW = 1 , CODEVISA = " & addvalue(xcodevisa.BoundText) & " WHERE DOC_NO = " & MyParn(xdoc_no.Text), nRec
'    Inform "  „  ⁄œÌ· ⁄œœ  " & nRec
End Sub

Private Sub xdate_open_LostFocus()
    myValidDate xdate_open
End Sub

Private Sub xDate_Validate(Cancel As Boolean)
If Not IsDate(xDate.text) Then Cancel = True
End Sub

Private Sub xDiscount_add_Change()
If bIg Then Exit Sub
CalcTotals
End Sub

Private Sub xDiscount_add_rate_Change()
If bIg Then Exit Sub
bIg = True
If Val(xtotalitem.Caption) - Val(xDiscount_offer.Caption) <> 0 Then
    If Round(Val(xDiscount_add_rate.text), 0) <> Round(Val(xDiscount_add.text) / (Val(xtotalitem.Caption) - Val(xDiscount_offer.Caption)) * 100, 0) Then
        xDiscount_add.text = Round((Val(xDiscount_add_rate.text) * (Val(xtotalitem.Caption) - Val(xDiscount_offer.Caption))) / 100, 2)
    End If
Else
    xDiscount_add.text = ""
End If
CalcTotals
bIg = False
End Sub
Private Sub xDiscount_Change()
If bIg Then Exit Sub

CalcTotals
End Sub

Private Sub xDiscount_LostFocus()
myLostFocus xDiscount
'CalcTotals
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
'If xDoc_No.Text = "" Then
'    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
'    Exit Function
'End If

If Trim(xCode.text) = "" Then
    MsgBox "·« ÌÊÃœ ﬂÊœ ⁄„Ì· «Ê ﬂÊœ „Ê—œ"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Not xStore.MatchedWithList Then
    If Not bIgMsg Then MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If

If Not xBox.MatchedWithList Then
    If Not bIgMsg Then MsgBox "·„ Ì „  ÕœÌœ «·Œ“‰…"
    Exit Function
End If

If xCodeDesca.Caption = "" Then
    If Not bIgMsg Then MsgBox "·« ÌÊÃœ ﬂÊœ ⁄„Ì·"
    Exit Function
End If

'If Not xMan.MatchedWithList Then
'    MsgBox "«·»«∆⁄ €Ì— „”Ã·"
'    Exit Function
'End If
'If Abs(Val(xDiscount_offer.Caption)) > Abs(Val(xDiscount.text)) Then
'    If Not bIgMsg Then
'        If MsgBox("Œ’„ «·›« Ê—… «ﬁ· „‰ Œ’„ «·⁄—Ê÷ √⁄«œ… Õ”«» «·Œ’„", vbDefaultButton2) = vbOK Then
'            UpdateDiscount xDoc_no.text, con
'        End If
'    End If
'    Exit Function
'End If

If Not ValidQuant Then
    If Not bIgMsg Then MsgBox "ﬂ„Ì… „ÊÃ»… Êﬂ„Ì… ”«·»… ›Ï ‰›” «·›« Ê—…"
    Exit Function
End If

For i = 1 To grid1.Rows - 2
    If grid1.TextMatrix(i, 15) = "" Then
        MsgBox "«·»«∆⁄ €Ì— „”Ã·"
    End If
Next
MYVALID = True
End Function
Private Sub myload()
xdoc_ret.text = ""
xUUID_RC.Caption = CardTable!UUID_RC & ""

bIgClick = True
xSendRc.Value = IIf(CardTable!sendRc, 1, 0)
bIgClick = False

XBRANCH.Caption = CardTable!branch & ""
xCode.Enabled = False


XISRETS.Value = 0

'XISONEST.Value = IIf(CardTable!ISONEST, 1, 0)

XISNODEL.Value = IIf(CardTable!ISNODEL, 1, 0)
'xIpName.Caption = CardTable!user_ip & ""
xDoc_No.text = CardTable!doc_no
xinv_no.text = CardTable!INV_NO & ""
xDoc_no2.text = CardTable!Doc_no2
xphone.Caption = CardTable!Phone & ""
xDate.text = myFormat_p(CardTable!Date)
xStore.BoundText = CardTable!STORE & ""
xMan.BoundText = CardTable!MAN & ""
xBox.BoundText = CardTable!BOX & ""
xNotes.text = CardTable!NOTES & ""
'XPOINT.Caption = Val(CardTable!Point & "")

xIsOffer.Value = IIf(CardTable!isoffer, 1, 0)

xUserName.Caption = ""
XSALES_RET.Caption = CardTable!SALES_RET & ""
xusername_RET.Caption = CardTable!username_ret & ""
xNote_disc.Caption = CardTable!NOTE_DISC & ""
xdoc_disc.Caption = CardTable!DOC_DISC & ""

xCode.text = CardTable!code & ""
xCodeDesca.Caption = CardTable!ClientDesca & ""

xcard_disc.text = CardTable!card_disc & ""
xcard_desca.Caption = CardTable!card_desca & ""
xcard_value.Caption = CardTable!card_VALUE & ""
xcard_doc.Caption = CardTable!card_DOC & ""

panel1(0).Caption = CardTable!UserName & ""
XUSERNAME_DISC.Caption = CardTable!username_disc & ""
xphone.Caption = CardTable!Phone & ""

chkCash.Value = IIf(CardTable!CASH, 1, 0)
bIg = True
xDiscount_add.text = Myvalue(CardTable!discount_add)
xDiscount.text = Myvalue(CardTable!discount)
xIsRet.Value = IIf(CardTable!ISRET, 1, 0)
bIg = False
xDiscount_offer.Caption = Myvalue(CardTable!discount_offer)
xcash.Caption = Myvalue(CardTable!CASH)
xvisa.Caption = Myvalue(Val(CardTable!Visa & "") + Val(CardTable!VISA3 & ""))
xPay.Caption = Myvalue(CardTable!PAY)
If Val(xPay.Caption) > 0 Then
    xRest.Caption = Myvalue(CardTable!rest)
Else
    xRest.Caption = ""
End If
xlate.Caption = Myvalue(CardTable!late)
bIg = True
xPrinted.Value = IIf(CardTable!PRINTED, 1, 0)
bIg = False
xdoc_replace.Caption = CardTable!Doc_REPLACE & ""
xtime.Caption = Format(CardTable!Time, "hh:nn")
panel1(1).Caption = CardTable!user_ip & ""

myLoadGrd

Handlecontrols LoadMode

myloadgrd2
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
Exit Sub
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub myDefine()
chkSendServer.Value = 1
bIg = True
xIsRet.Value = 0
bIg = False

xUUID_RC.Caption = ""
XISNODEL.Value = 0
xIsOffer.Value = 0

bIgClick = True
xSendRc.Value = 0
bIgClick = False

XBRANCH.Caption = cBranch
'xship_no.Caption = ""
'xonline_doc.Caption = ""
'lDiscModelRet = False
'xIpName.Caption = GetComputerName
panel1(1).Caption = GetComputerName
XISRETS.Value = 0
sUserDisc = ""
lIsDocRet = False
xMan.BoundText = ""
XUSERNAME_DISC.Caption = ""
xUserName.Caption = cusername

'xRate.Locked = True
'xDiscount.Locked = True

xdoc_ret.text = ""
myDisplay cMsgDisplay, " "
xClosed.Value = 0
xType.Caption = ""
xusername_RET = ""
xcard_disc.text = ""
xcard_desca.Caption = ""
xcard_value.Caption = ""
xcard_doc.Caption = ""

XSALES_RET.Caption = ""
xphone.Caption = ""
xDoc_No.text = ""


'xDoc_No.Tag = DefineMode
xDoc_no2.text = ""

xinv_no.text = ""
'xusername.Caption = cusername
xStore.BoundText = cBranchStore
xBox.BoundText = cBranchBox
xCode.text = cDefClient
xCodeDesca.Caption = cDefClientDesca

xDate.text = myFormat_p(fnDateSales(con))
xdoc_replace.Caption = ""
xBalance.Caption = ""
xMosm.BoundText = cPMosm


chkCash.Value = 1
xTotalQuant.Caption = ""
xTotalItemNoDiscount.Caption = ""
xtotalitem.Caption = ""
xTotalItemDiscount.Caption = ""
xTotalItemDiscount_rate.Caption = ""
xtotalitem_offer.Caption = ""


bIg = True
xIsRet.Value = 0
xPrinted.Value = 0
xDiscount_add.text = ""
xDiscount.text = ""
xRate.text = ""
xDiscount_add_rate.text = ""
bIg = False

xDiscount_offer.Caption = ""
xTotal.Caption = ""

xlate.Caption = ""
xvisa.Caption = ""
xcash.Caption = ""
xNotes.text = ""
xRest.Caption = ""
xPay.Caption = ""

xtime.Caption = Format(Time, "hh:nn")

grid1.Rows = 1

myAddItem

fixGrd

GrdOffer.Rows = 1
Fixgrd2
Handlecontrols DefineMode
CalcTotals
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bedit And (cBranchBox <> "" Or nMode = LoadMode)

cmdNewInv.Enabled = bedit And cBranchBox <> ""

bEditRecord = bedit
bEditRecord = bEditRecord And xPrinted.Value = 0
bEditRecord = bEditRecord And XISNODEL.Value = 0
If (Not optType(0).Value) Then
    bEditRecord = False
End If


xphone.Enabled = bEditRecord
cmdSave.Enabled = bEditRecord
xdoc_ret.Enabled = bedit And nMode = DefineMode


cmd_showdoc.Enabled = bEditRecord
fmClose.Enabled = xPrinted.Value = 1 And nMode = LoadMode And nUser = enUser.Admin
cmd_disc.Enabled = bEditRecord And XSALES_RET.Caption = "" And xdoc_replace.Caption = "" And xIsOffer.Value = 0
cmdDiscount_rate_add.Enabled = bEditRecord And XSALES_RET.Caption = "" And xdoc_replace.Caption = "" And xIsOffer.Value = 1
cmdAddOffer.Enabled = bEditRecord And XSALES_RET.Caption = "" And xdoc_replace.Caption = ""

cmddel.Enabled = bEditRecord And nMode = LoadMode

cmdRefund.Enabled = bEditRecord And IsDgt(xdoc_ret.text) And xDoc_No.Tag = DefineMode
cmdReplace.Enabled = bEditRecord And IsDgt(xdoc_ret.text) And xDoc_No.Tag = DefineMode

xcash.Enabled = nUser >= enUser.Super And xPrinted.Value = 1

xDiscount.Enabled = False
xRate.Enabled = False
xDiscount_add.Enabled = False
xDiscount_add_rate.Enabled = False

xIsRet.Visible = Val(xTotalItemNoDiscount.Caption) > 0 And nUser > enUser.Super And nMode = LoadMode

Dim nRecord As Long, nRecords As Long
retRecords xDoc_No.text, nRecords, nRecord

cmdNext.Enabled = nRecord < nRecords And nRecords <> 0 And nMode = LoadMode
cmdPrevious.Enabled = nRecord <> 1 And nRecords <> 0 And nMode = LoadMode
cmdLast.Enabled = nRecord < nRecords And nRecords <> 0 And nMode = LoadMode
cmdFirst.Enabled = nRecord <> 1 And nRecords <> 0 And nMode = LoadMode

If nMode = LoadMode Then
    panel1(0).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
    panel1(0).Caption = "”Ã· ÃœÌœ (" & (nRecords + 1) & ")"
End If

xDoc_No.Tag = nMode

cmdGo.Enabled = (xUUID_RC.Caption = "" Or xSendRc.Value = 0) And nMode = LoadMode And xPrinted.Value = 1
xSendRc.Enabled = nMode = LoadMode And xUUID_RC.Caption <> "" And xPrinted.Value = 1
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If Not bEditRecord Then
    Exit Sub
ElseIf XSALES_RET.Caption <> "" Or xdoc_replace.Caption <> "" Then
    Exit Sub
ElseIf KeyCode = 112 And grid1.Editable = flexEDKbdMouse Then
    ItemsLook Me, oSearchItem
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 And (grid1.TextMatrix(grid1.Row, 18) = "" Or grid1.ValueMatrix(grid1.Row, 19) = 0) Then
            
    RemoveItem (grid1.Row)
    
    grid1.Select grid1.Rows - 1, 1
    grid1.ShowCell grid1.Rows - 1, 1
    MakeSerial grid1.Row
        
    CalcTotals
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
End If
'If KeyCode = 116 Then
'    cmdSave_Click
'End If
End Sub
Private Function GrdDesc(pItem As String, ByRef Row As Long, Optional ByRef sMsg As String) As Boolean
With grid1
If Trim(pItem) = "" Or Len(pItem) < 3 Or Len(pItem) > 30 Then
    sMsg = "ﬂÊœ €Ì— ’«·Õ"
    Exit Function
End If

Dim cm As New ADODB.command
Set cm = cmd("dbo.sp_item", con, adStoredProc, AddFlag(Empty, "BARCODE_FIND", pItem))
cm.Execute

If IsNull(cm.Parameters("@ITEM").Value) Then
    sMsg = "ﬂÊœ €Ì— ’«·Õ"
    Exit Function
End If

If Not cm.Parameters("@ISNOITEM").Value Then
    If cm.Parameters("@PRICE").Value = 0 Then
        MsgBox "”⁄— «·»Ì⁄ ’›— "
        Exit Function
    End If

    If fnBalance(cm.Parameters("@ITEM").Value, con, xStore.BoundText, xDate.text) <= 0 Then
        Inform "·« ÌÊÃœ —’Ìœ «·’‰›"
    End If
End If

Dim cmDiscount As New ADODB.command
aPrm = AddFlag(Empty, "ITEM", cm.Parameters("@ITEM").Value)
aPrm = AddFlag(aPrm, "DATE", myFormat_sp(xDate.text))
Set cmDiscount = cmd("dbo.sp_offer_price", con, adStoredProc, aPrm)
cmDiscount.Execute

grid1.TextMatrix(Row, 1) = cm.Parameters("@ITEM").Value
grid1.TextMatrix(Row, 2) = cm.Parameters("@MOSM").Value
grid1.TextMatrix(Row, 3) = cm.Parameters("@FACT").Value
grid1.TextMatrix(Row, 4) = cm.Parameters("@SUPP").Value & ""
grid1.TextMatrix(Row, 5) = cm.Parameters("@MODELFACT0").Value
grid1.TextMatrix(Row, 6) = cm.Parameters("@DESCA").Value
grid1.TextMatrix(Row, 7) = cm.Parameters("@SCAL").Value
grid1.TextMatrix(Row, 8) = cm.Parameters("@COLOR").Value
grid1.TextMatrix(Row, 9) = cm.Parameters("@PRICE").Value
grid1.TextMatrix(Row, 10) = "1"

If IsNull(cmDiscount.Parameters("@ID").Value) Then
    grid1.TextMatrix(Row, 11) = cm.Parameters("@PRICE").Value
End If
grid1.TextMatrix(Row, 17) = Val(cm.Parameters("@COSTITEM").Value & "")
If Not IsNull(cmDiscount.Parameters("@ID").Value) Then
    grid1.TextMatrix(Row, 11) = cmDiscount.Parameters("@PRICE").Value
    grid1.TextMatrix(Row, 18) = cmDiscount.Parameters("@DOC_NO").Value
    grid1.TextMatrix(Row, 19) = 0
    grid1.TextMatrix(Row, 20) = 0
End If

GrdDesc = True
End With
End Function
Private Function CalcTotals(Optional nMode As Integer = 0)
Dim nTotalQuant As Double, nTotalQuant2 As Double, nTotalItem As Double, nTotalItemNoDiscount As Double, nTotalDiscount As Double, nTotalMin As Double
Dim nDisountRow As Double
With grid1
Dim i As Long
For i = 1 To grid1.Rows - 2
    nDisountRow = (.ValueMatrix(i, 9) * .TextMatrix(i, 10)) - (.ValueMatrix(i, 11)) * .ValueMatrix(i, 10)
    .TextMatrix(i, 13) = .ValueMatrix(i, 10) * .ValueMatrix(i, 9)
    If nDisountRow <> 0 And .ValueMatrix(i, 13) <> 0 Then
        .TextMatrix(i, 12) = Format(Round(nDisountRow / .ValueMatrix(i, 13), 2), "0%")
    Else
        .TextMatrix(i, 12) = ""
    End If
        
    
    nTotalQuant = nTotalQuant + grid1.ValueMatrix(i, 10)
    nTotalItem = nTotalItem + .ValueMatrix(i, 10) * .ValueMatrix(i, 11)
    
    nTotalItemNoDiscount = nTotalItemNoDiscount + (.ValueMatrix(i, 9) * .ValueMatrix(i, 10))
    nTotalItemDiscount = nTotalItemDiscount + nDisountRow
Next

xTotalQuant.Caption = Myvalue(nTotalQuant)
xTotalItemNoDiscount.Caption = Myvalue(nTotalItemNoDiscount)
xTotalItemDiscount.Caption = Myvalue(nTotalItemDiscount)
xtotalitem.Caption = Myvalue(nTotalItem)
xtotalitem_offer.Caption = Myvalue(nTotalItem - Val(xDiscount_offer.Caption))


If nTotalItemNoDiscount <> 0 And nTotalItemDiscount Then
    xTotalItemDiscount_rate.Caption = Format(Round(nTotalItemDiscount / nTotalItemNoDiscount, 4), "0%")
Else
    xTotalItemDiscount_rate.Caption = ""
End If

If Val(xDiscount_offer.Caption) <> 0 And nTotalItem <> 0 Then
    xDiscount_offer_rate.Caption = Format(Val(xDiscount_offer.Caption) / nTotalItem, "0%")
Else
    xDiscount_offer_rate.Caption = ""
End If

bIg = True
If Val(xtotalitem_offer.Caption) <> 0 Then
    If Round(Val(xDiscount_add_rate.text), nRound) <> Round(Val(xDiscount_add.text) / Val(xtotalitem_offer.Caption) * 100, nRound) Then
        xDiscount_add_rate.text = Myvalue(Round((Val(xDiscount_add.text) / Val(xtotalitem_offer.Caption)) * 100, nRound))
    End If
Else
    xDiscount_add_rate.text = ""
End If

If xIsOffer.Value = 1 Then
    xDiscount.text = Myvalue(Val(xDiscount_offer.Caption) + Val(xDiscount_add.text))
ElseIf xIsOffer.Value = 0 And Val(xDiscount_add.text) <> 0 Then
    'xDiscount.text = Myvalue(xDiscount_add.text)
    xDiscount_add.text = ""
End If

If Val(xtotalitem.Caption) <> 0 Then
    If Round(Val(xRate.text), nRound) <> Round(Val(xDiscount.text) / Val(xtotalitem.Caption) * 100, nRound) Then
        xRate.text = Myvalue(Round((Val(xDiscount.text) / Val(xtotalitem.Caption)) * 100, nRound))
    End If
Else
    xRate.text = ""
End If

xTotal.Caption = Myvalue(mRound(nTotalItem - Val(xDiscount.text), 2))
bIg = False
End With
End Function
Private Sub CardLookup(Optional pFilter As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(5, 1)
Set Generalarray(0) = Me

Generalarray(1) = "SELECT FILE6_20H.DOC_NO ," & _
                  " FILE6_20H.DOC_NO2," & _
                  " FORMAT(FILE6_20H.date,'yyyy/M/d')," & _
                  " FILE0_50.DESCA," & _
                  " CONVERT(VARCHAR(10), FILE6_20H.time, 108)," & _
                  " FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT" & _
                  " FROM FILE6_20H" & _
                  " INNER JOIN FILE3_10 ON FILE6_20H.CODE = FILE3_10.CODE" & _
                  " INNER JOIN FILE0_50 ON FILE6_20H.BOX = FILE0_50.CODE" & _
                  " INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & " WHERE " & cFilter
End If

Generalarray(2) = "Order by doc_no2 DESC "
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·„” ‰œ"
listarray(0, 1) = "(FILE6_20H.DOC_NO = 'cFilter' OR FILE6_20H.DOC_NO2 = 'cFilter')"

listarray(1, 0) = "«· «—ÌŒ"
listarray(1, 1) = "(##FILE6_20H.DATE##)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(1, 1) = 1200

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·Õ“‰…"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "«·Êﬁ "
GrdArray(4, 1) = 1300

GrdArray(5, 0) = "«·ﬁÌ„…"
GrdArray(5, 1) = 1300

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub


Private Sub xDiscount_offer_Change()
lblOffer.Visible = Val(xDiscount_offer.Caption) <> 0
xtotalitem_offer.Visible = Val(xDiscount_offer.Caption) <> 0
End Sub

Private Sub xdoc_replace_Change()
xdoc_replace.Visible = xdoc_replace.Caption <> ""
If xdoc_replace.Caption <> "" Then
    xType.Caption = "«” »œ«·"
Else
    xType.Caption = ""
End If
End Sub
Private Sub xdoc_ret_Change()
cmdReplace.Enabled = IsDgt(xdoc_ret.text) And xDoc_No.Tag = DefineMode And (Len(Trim(xdoc_ret.text)) >= 11)
cmdRefund.Enabled = IsDgt(xdoc_ret.text) And xDoc_No.Tag = DefineMode And (Len(Trim(xdoc_ret.text)) >= 11)
End Sub

Private Sub xdoc_ret_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then cmd_showdoc_Click
End Sub

Private Sub xIsRet_Click()
If bIg Then Exit Sub
If MsgBox("«·”„«Õ »„— Ã⁄ ··›« Ê—…", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
    
    Dim conServer As New ADODB.Connection
    If openCon(conServer, LoadConServer, 2) <> "ok" Then
        MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ ··œŒÊ· ··›—⁄ «·—∆Ì”Ì ·«—”«· «·›Ê« Ì—"
        Exit Sub
    End If
    
    Dim nRec As Long
    On Error GoTo myError
    con.Execute " UPDATE FILE6_20H " & _
                      " SET ISRET = " & xIsRet.Value & " , ISNEW = 1 WHERE DOC_NO = " & MyParn(xDoc_No.text)
    conServer.BeginTrans
    conServer.Execute " UPDATE FR6_20H " & _
                      " SET ISRET = " & xIsRet.Value & _
                      " WHERE DOC_NO = " & MyParn(xDoc_No.text)
    conServer.CommitTrans
                      
    If nRec = 1 Then MsgBox "  „  ⁄œÌ· «·„” ‰œ "
Finally:
    closeCon conServer
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
If TransCount(conServer) > 0 Then conServer.RollbackTrans
GoTo Finally
End Sub

Private Sub xMan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (Not xMan.MatchedWithList) And IsDgt(xMan.text) Then
        Dim sCodeEmp As String
        sCodeEmp = xMan.BoundText
        xMan.BoundText = RetZero(xMan.text, 4)
        If Not xMan.MatchedWithList Then
            If Not xMan.MatchedWithList Then
                xMan.BoundText = GetField("SELECT CODE FROM FILE6_25 WHERE ID_EMP = " & MyParn(sCodeEmp), con)
            End If
        End If
    End If
End If
End Sub
Private Sub xMAN_Validate(Cancel As Boolean)
'If Not xMan.MatchedWithList Then
'    Cancel = True
'ElseIf xMan.BoundText = "" Then
'    Cancel = True
'Else
'    grid1.TextMatrix(grid1.Rows - 1, 15) = xMan.BoundText
'End If
End Sub

Private Sub XPHONE_Click()
Set Phonesfrm.myForm = Me
Set Phonesfrm.con = con
Phonesfrm.Show 1
End Sub

Private Sub xPrinted_Click()
If bIg Then Exit Sub
con.Execute "update file6_20h set printed = " & xPrinted.Value & ",closed = " & xPrinted.Value & " from file6_20h where doc_no = " & MyParn(xDoc_No.text)
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then
        myDefine
    End If
End If
End Sub
Private Sub xRate_Change()
If bIg Then Exit Sub
bIg = True
If Val(xtotalitem.Caption) <> 0 Then
    If Round(Val(xRate.text), 0) <> Round(Val(xDiscount.text) / Val(xtotalitem.Caption) * 100, 0) Then
        xDiscount.text = Round((Val(xRate.text) * Val(xtotalitem.Caption)) / 100, 2)
    End If
Else
    xDiscount.text = ""
End If
CalcTotals
bIg = False
End Sub

Private Sub xRate_LostFocus()
myLostFocus xRate
'If Val(xTotalItem.Caption) <> 0 Then
'    If Round(Val(xRate.text), nRound) <> Round(Val(xDiscount.text) / Val(xTotalItem.Caption) * 100, nRound) Or xDiscount.Locked Then
'        xDiscount.text = Round((Val(xRate.text) * Val(xTotalItem.Caption)) / 100, 0)
'    End If
'Else
'    xDiscount.text = ""
'End If
'CalcTotals
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub fixGrd()
With grid1
'                 0          1          2         3        4          5              6          7          8            9           10          11             12           13            14                15            16            17           18              19              20
.FormatString = "„.|" & "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "”⁄— „” Â·ﬂ|" & "⁄œœ|" & "”⁄— «·»Ì⁄|" & "‰”»… Œ’„|" & "«·≈Ã„«·Ï|" & "”⁄— √Êﬂ«“ÌÊ‰|" & "ﬂÊœ «·»«∆⁄|" & "«·»«∆⁄|" & "«· ﬂ·›…|" & "„” ‰œ «·⁄—÷|" & "—ﬁ„ «·⁄—÷|" & "Œ’„ «·⁄—÷|"
.RowHeight(0) = 600
.WordWrap = True


.ColHidden(14) = True

'.ColHidden(18) = True

.ColWidth(0) = 600
.ColWidth(1) = 1000
.ColWidth(2) = 0
.ColWidth(3) = 1200
.ColWidth(4) = 0
.ColWidth(5) = 1300
.ColWidth(6) = 2500
.ColWidth(7) = 700
.ColWidth(8) = 1000
.ColWidth(9) = 1000
.ColWidth(10) = 800
.ColWidth(11) = 1000
.ColWidth(12) = 600
.ColWidth(13) = 1200
.ColWidth(16) = 1700
'.ColHidden(12) = True
.ColHidden(15) = True
.ColHidden(17) = True
.ColHidden(18) = True
.ColHidden(19) = True
.ColHidden(20) = True

'.ColHidden(.Cols - 2) = True
.ColHidden(.Cols - 1) = True
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
.Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter

Dim bGroup As Boolean
Dim sGroup As String
For i = 1 To grid1.Rows - 2
    If .ValueMatrix(i, 9) <> .ValueMatrix(i, 11) Then
        .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
        .Cell(flexcpFontUnderline, i, 1, i, .Cols - 1) = vbRed
    End If
'    If grid1.TextMatrix(i, 18) <> "" And .ValueMatrix(i, 19) > 0 Then
'        grid1.Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &HF0F0F0
'    Else
'        grid1.Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &H80000005
'    End If
    If .ValueMatrix(i, 19) > 0 Then
        If sGroup <> .TextMatrix(i, 19) Then
            sGroup = .TextMatrix(i, 19)
            bGroup = Not bGroup
        End If
        .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = RGB(IIf(bGroup, 255, 200), 255, 210)
    Else
        .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = vbWhite
    End If
Next
'.ShowCell .Rows - 1, 0
End With
End Sub
Private Sub Fixgrd2()
With GrdOffer '                 0          1          2
.FormatString = "„|" & "«·„” ‰œ|" & "«·⁄—÷|" & "«·Œ’„"
.ColHidden(0) = True
.ColHidden(1) = True
.ColWidth(0) = 500
.ColWidth(1) = 1000
.ColWidth(2) = 2300
.ColWidth(3) = 1100
'.ColComboList(2) = "..."

For i = 0 To GrdOffer.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
If .Rows > 2 Then
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 3, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    .TextMatrix(.Rows - 1, 2) = "«·≈Ã„«·Ì"
End If
End With
End Sub
Private Sub XSALES_RET_Change()
XSALES_RET.Visible = XSALES_RET.Caption <> ""
End Sub
Private Sub xSendRc_Click()
If bIgClick Then Exit Sub
If xDoc_No.text <> "" Then
    On Error GoTo myError
    If xUUID_RC.Caption <> "" Then
        If Not IsEmpty(myField("select doc_no from file6_20h where PREVIOUS_UUID = " & MyParn(xUUID_RC.Caption), con)) Then
            MsgBox "—ﬁ„ „”·”· „—Ã⁄Ì ”«»ﬁ ··›« Ê—… —ﬁ„ " & MyParn(xDoc_No.text)
            'xSendRc.Value = IIf(xSendRc.Value = 1, 0, 1)
        End If
    End If
    con.Execute "UPDATE FILE6_20H SET FILE6_20H.sendRc = " & xSendRc.Value & " FROM FILE6_20H WHERE FILE6_20H.DOC_NO = " & MyParn(xDoc_No.text)
    Inform " „ «· ⁄œÌ· »‰Ã«Õ"
End If
Finaly:
Handlecontrols xDoc_No.Tag
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
bIgClick = True
xSendRc.Value = IIf(xSendRc.Value = 1, 0, 1)
bIgClick = False
GoTo Finaly
End Sub
Private Sub xStore_Validate(Cancel As Boolean)
    If Trim(xStore.BoundText) = "" Then Cancel = True
End Sub
Private Sub xTax_LostFocus()
    myLostFocus xTax
    CalcTotals
End Sub
Private Function RemoveItem(Row As Long) As Boolean
If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) <> vbOK Then
    Exit Function
End If

If grid1.TextMatrix(Row, grid1.Cols - 1) <> "" Then
    Dim oUser As New sub_userfrm
    Set oUser.myForm = Me
    oUser.Show 1
    
    If subUserCode = "" Then
        Exit Function
    End If
    
    If subUserCode <> "" Then
        On Error GoTo myError
        con.Execute "Delete  From FILE6_20 where id = " & grid1.TextMatrix(Row, grid1.Cols - 1)
    End If
End If
grid1.RemoveItem Row
RemoveItem = True
Exit Function
myError:
MsgBox Err.Description
Err.Clear
End Function
Private Function myreplaceGrd(Row) As Boolean
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "ITEM", addstring(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "QUANT", .ValueMatrix(i, 10))
        aInsert = AddFlag(aInsert, "PRICE", .ValueMatrix(i, 11))
        aInsert = AddFlag(aInsert, "PRICE_c", .ValueMatrix(i, 9))
        aInsert = AddFlag(aInsert, "COST", .ValueMatrix(i, 17))
        aInsert = AddFlag(aInsert, "MAN", addstring(.TextMatrix(i, 15)))
        aInsert = AddFlag(aInsert, "S_OKAZ", .ValueMatrix(i, 12))
        aInsert = AddFlag(aInsert, "USER_IP", addstring(cIpName))
        
        aInsert = AddFlag(aInsert, "DOC_OFFER", addstring(grid1.TextMatrix(i, 18)))
        aInsert = AddFlag(aInsert, "OFFER_NO", addstring(grid1.TextMatrix(i, 19)))
        
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.text))
            con.Execute addInsert(aInsert, "FILE6_20")
        Else
            con.Execute addUpdate(aInsert, "FILE6_20", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Function myreplaceGrdReplace() As Boolean
Dim aInsert As Variant
Dim cString As String
cString = "select Max(Offer_no) " & _
          " from file6_20 " & _
          " where doc_no = " & MyParn(xDoc_No.text)
Dim nOffer As Long
If Val(frmOfferReplace.xQty2.Caption) > 0 Then
    nOffer = Val(myField(cString, con) & "") + 1
End If
                                    
With frmOfferReplace.grid1
    For i = 1 To .Rows - 2
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.text))
        aInsert = AddFlag(aInsert, "ITEM", addstring(.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "QUANT", .ValueMatrix(i, 10))
        aInsert = AddFlag(aInsert, "PRICE", .ValueMatrix(i, 9))
        aInsert = AddFlag(aInsert, "PRICE_c", .ValueMatrix(i, 14))
        aInsert = AddFlag(aInsert, "COST", .ValueMatrix(i, 13))
        If frmOfferReplace.xMan.Caption <> "" Then
            aInsert = AddFlag(aInsert, "MAN", addstring(frmOfferReplace.xMan.Caption))
        ElseIf xMan.MatchedWithList Then
            aInsert = AddFlag(aInsert, "MAN", addstring(xMan.BoundText))
        End If
        aInsert = AddFlag(aInsert, "OFFER_NO", nOffer)
        aInsert = AddFlag(aInsert, "DOC_OFFER", addstring(frmOfferReplace.xDoc_No.Caption))
        aInsert = AddFlag(aInsert, "DISCOUNT_OFFER", Val(frmOfferReplace.xDiscount.Caption))
        aInsert = AddFlag(aInsert, "DISCOUNT_OFFER_RATE", Val(frmOfferReplace.xRate.Caption))
        aInsert = AddFlag(aInsert, "USER_IP", addstring(cIpName))
        con.Execute addInsert(aInsert, "FILE6_20")
    Next
End With
myreplaceGrdReplace = True
End Function
Private Function myreplaceGrdOffer() As Boolean
Dim aInsert As Variant
Dim cString As String
cString = "select Max(Offer_no) " & _
          " from file6_20 " & _
          " where doc_no = " & MyParn(xDoc_No.text)
Dim nOffer As Long
nOffer = Val(myField(cString, con) & "") + 1
                                    
With frmOffer.grid1
    For i = 1 To frmOffer.grid1.Rows - 2
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.text))
        aInsert = AddFlag(aInsert, "ITEM", addstring(.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "QUANT", .ValueMatrix(i, 10))
        aInsert = AddFlag(aInsert, "PRICE", .ValueMatrix(i, 9))
        aInsert = AddFlag(aInsert, "PRICE_c", .ValueMatrix(i, 9))
        aInsert = AddFlag(aInsert, "COST", .ValueMatrix(i, 13))
        aInsert = AddFlag(aInsert, "MAN", addstring(xMan.BoundText))
        aInsert = AddFlag(aInsert, "OFFER_NO", nOffer)
        aInsert = AddFlag(aInsert, "DOC_OFFER", addstring(frmOffer.xDoc_No.Caption))
        aInsert = AddFlag(aInsert, "DISCOUNT_OFFER", Val(frmOffer.xDiscount.Caption))
        aInsert = AddFlag(aInsert, "DISCOUNT_OFFER_RATE", Val(frmOffer.xRate.Caption))
        aInsert = AddFlag(aInsert, "USER_IP", addstring(cIpName))
        con.Execute addInsert(aInsert, "FILE6_20")
    Next
End With
myreplaceGrdOffer = True
End Function
Private Sub myLoadGrd()
With grid1
Dim cString As String
    cString = "SELECT FILE6_20.ITEM," & _
              "FILE1_10.MOSM," & _
              "FACT.DESCA," & _
              "FILE1_10.SUPP," & _
              "FILE1_10.MODELFACT0," & _
              "FILE1_10.DESCA," & _
              "FILE1_10.SCAL," & _
              "FILE1_10.COLOR," & _
              "FILE6_20.PRICE_C," & _
              "FILE6_20.QUANT," & _
              "FILE6_20.PRICE," & _
              "FILE6_20.S_OKAZ," & _
              "FILE6_20.TOTAL," & _
              "FILE6_20.PRICE_C2," & _
              "FILE6_20.MAN," & _
              "FILE6_25.DESCA," & _
              "FILE6_20.COST," & _
              "FILE6_20.DOC_OFFER," & _
              "FILE6_20.OFFER_NO," & _
              "FILE6_20.DISCOUNT_OFFER," & _
              "FILE6_20.ID " & _
              " FROM FILE6_20 " & _
              " INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM" & _
              " LEFT JOIN FACT ON FACT.CODE = FILE1_10.FACT" & _
              " LEFT JOIN FILE6_25 ON FILE6_20.MAN  = FILE6_25.CODE "
    cString = cString & " WHERE FILE6_20.DOC_NO = " & MyParn(xDoc_No.text)
    cString = cString & " ORDER BY FILE6_20.ID"
    Set DATA11.Recordset = cmd(cString, con).Execute
    myAddItem
End With
CalcTotals
fixGrd
End Sub
Private Function mysave(Optional bEnd As Boolean = True, Optional bPrint As Boolean = True, Optional bSendReceipt As Boolean = False) As Boolean

If Not MYVALID Then Exit Function

CalcTotals

If Not myreplace Then Exit Function
             
         
Inform_OK " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ", " «·»«ﬁÏ " & Format(xRest.Caption, "#0.00")

If Val(xTotal.Caption) <> Val(xvisa.Caption) + Val(xcash.Caption) Then
    MsgBox " »—Ã«¡ „—«Ã⁄…  ”ÃÌ· «·”œ«œ "
    Exit Function
End If

SavePrint
                                                  
doprint_cash (False)

If chkSendServer.Value = 1 Or XSALES_RET.Caption <> "" Or xdoc_replace.Caption <> "" Then
    Dim nRecords As Long
    If sendInvoicesNew(nRecords, , xDoc_No.text) Then
        If nRecords > 1 Then
            MsgBox " „ «—”«· " & nRecords & " ›« Ê—… ··›—⁄ «·—∆Ì”Ì"
        ElseIf nRecords = 0 Then
            MsgBox " „ «·« ’«· »«·”Ì—›— Ê·„ Ì „ «÷«›… «·›« Ê—…"
        End If
    ElseIf chkSendServer.Value = 1 Then
        MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰ «—”«· «·›« Ê—… ··›—⁄ «·—∆Ì”Ì"
        Exit Function
    End If
End If
                                                                             
mysave = True
End Function
Private Sub SavePrint()
On Error GoTo myError
'con.BeginTrans
con.Execute "update file6_20h set FILE6_20H.PRINTED = 1 , ISCLOSED = 1  WHERE DOC_NO = " & MyParn(xDoc_No.text)
'con.CommitTrans
bIg = True
xPrinted.Value = 1
bIg = False
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Function doprint_cash(Optional bCopy As Boolean = False)
Dim loctable As New ADODB.Recordset
Set loctable = cmd("dbo.sp_invoice_print", con, adStoredProc, AddFlag(Empty, "DOC_NO", xDoc_No.text)).Execute

contemp.Execute "DELETE * FROM TEMP"
If temptable.State = adStateOpen Then temptable.Close

temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

With loctable
Do Until loctable.EOF
    temptable.AddNew
    temptable!STR20 = loctable!doc_no
    temptable!STR21 = loctable!company_name
    
    If bCopy Then
        temptable!str22 = "C  O  P  Y"
    End If
    
    'temptable!STR23 = loctable!HEAD1
    temptable!STR24 = loctable!Company_address
    If Not IsNull(loctable!COMPANY_Phone) Then
        temptable!STR25 = "Phone : " & loctable!COMPANY_Phone
    End If
    
    If Not IsNull(loctable!ID_NO1) Then
        temptable!STR19 = "—ﬁ„  ”ÃÌ· ÷—Ì»Ì : " & loctable!ID_NO1
    End If
    
    If Not IsNull(loctable!ID_NO2) Then
        temptable!str13 = "”Ã·  Ã«—Ì : " & loctable!ID_NO2
    End If
    
       
    temptable!str1 = myFormat_p(loctable!Date)
    temptable!str2 = Format(loctable!Time, "hh:nn")
    temptable!str3 = loctable!Doc_no2
    temptable!str4 = loctable!Casher
    
    temptable!str10 = loctable!Item
    temptable!str11 = loctable!desca
    temptable!VAL1 = loctable!Quant
    temptable!val2 = loctable!PRICE_C
    temptable!val3 = loctable!total_c
    temptable!VAL4 = loctable!price
    temptable!val5 = loctable!total
    
    temptable!VAL6 = loctable!total_inv_c
    temptable!VAL7 = loctable!items_discount
    temptable!Val8 = loctable!TOTAL_inv
    
    temptable!val9 = loctable!discount_offer
    temptable!Val10 = loctable!discount_add
    temptable!val11 = loctable!discount
    temptable!val12 = loctable!discount + loctable!items_discount
    temptable!VAL13 = loctable!total_net
    
    temptable!val15 = loctable!CASH
    temptable!val16 = loctable!Visa
    temptable!val17 = loctable!Point
    temptable!str4 = loctable!BOX
    temptable!str5 = loctable!MAN
    
    temptable!Val21 = IIf(IsNull(loctable!offer_no), -1, loctable!offer_no)
    If Not IsNull(loctable!offer_no) Then
        If loctable!offer_no = 0 Then
            temptable!str9 = "Direct Discount"
            temptable!val18 = loctable!total_c - loctable!total
            temptable!VAL19 = loctable!discount_offer_rate_direct
        Else
            temptable!str9 = loctable!offer_desca
            temptable!val18 = loctable!discount_offer_Sub
            temptable!VAL19 = loctable!discount_offer_Rate_Sub
        End If
    End If
        
    'If bCopy Then temptable!str1 = "C  O  P  Y"
    'temptable!STR4 = loctable!UserName
    
    temptable.Update
    loctable.MoveNext
Loop
End With

contemp.BeginTrans
contemp.CommitTrans
On Error Resume Next
If Not lNotBranch Then
    REPORT1.Reset
    FixPrinter REPORT1
End If

REPORT1.ReportFileName = App.Path & "\Reports\sales_cash.rpt" '
REPORT1.CopiesToPrinter = 1
REPORT1.Destination = crptToPrinter
REPORT1.ProgressDialog = False
REPORT1.DataFiles(0) = tempFile
iSubreports = REPORT1.GetNSubreports

If (iSubreports <> 0) Then
    For i = 0 To iSubreports - 1
        sSubreportName = REPORT1.GetNthSubreportName(i)
        REPORT1.SubreportToChange = sSubreportName
        REPORT1.DataFiles(0) = tempFile
    Next
End If
REPORT1.Action = 1


If xGift.Value = 1 Then
    REPORT1.ReportFileName = App.Path & "\Reports\sales_cash_G.rpt"
    REPORT1.CopiesToPrinter = 1
    'REPORT1.Destination = crptToWindow
    REPORT1.Destination = crptToPrinter
    REPORT1.ProgressDialog = False
    REPORT1.DataFiles(0) = tempFile
    iSubreports = REPORT1.GetNSubreports
    
    If (iSubreports <> 0) Then
        For i = 0 To iSubreports - 1
            sSubreportName = REPORT1.GetNthSubreportName(i)
            REPORT1.SubreportToChange = sSubreportName
            REPORT1.DataFiles(0) = tempFile
        Next
    End If
    REPORT1.Action = 1
End If

doprin_cash = True
Finllay:
temptable.Close
Set temptable = Nothing
Exit Function
myError:
MsgBox Err.Description
Err.Clear
GoTo Finllay
End Function
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = True) As Boolean
With grid1
'If Not myValid(bIgMsg) Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 1)) Then Exit Function
If .ValueMatrix(Row, 10) = 0 Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 11)) Then Exit Function
If Not Trim(.TextMatrix(Row, 15)) <> "" Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf col = 1 And grid1.TextMatrix(Row, col) = "" Then
ElseIf col < 10 Then
    grid1.col = 10
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 10)
    grid1.ShowCell grid1.Row, 0
End If
End Sub
Private Sub myAddItem()
grid1.AddItem ""
If grid1.Rows > 2 Then
    grid1.TextMatrix(grid1.Rows - 1, 15) = grid1.TextMatrix(grid1.Rows - 2, 15)
    grid1.TextMatrix(grid1.Rows - 1, 16) = grid1.TextMatrix(grid1.Rows - 2, 16)
End If
If grid1.TextMatrix(grid1.Rows - 1, 15) = "" Then
    grid1.TextMatrix(grid1.Rows - 1, 15) = xMan.BoundText
    grid1.TextMatrix(grid1.Rows - 1, 16) = xMan.text
End If
MakeSerial
End Sub
Private Sub xTotal_Change()
If Not bIg Then
    bIg = True
    xDiscount.text = Round(Val(xtotalitem.Caption) - Val(xTotal.Caption), 2)
    CalcTotals
    bIg = False
End If
End Sub

Private Sub xDoc_no2_GotFocus()
myGotFocus xDoc_no2
End Sub
Sub PostData()
'If DateValue(salesfrm.xDate.text) = DateValue(dSalesDate) Then
              
    Dim sDate As String
    
    'dSalesDate = fnDateSales(con)
    sDate = myFormat(DateAdd("d", 1, dSalesDate))
    
    If DateDiff("d", fnDateSales(con), myFormat(Date)) < -1 Then
        MsgBox " «—ÌŒ «·ÃÂ«“ Ì ⁄œÌ  «—ÌŒ «·„»Ì⁄«  »√ﬂÀ— „‰ ÌÊ„" & vbCrLf & _
                "«· «—ÌŒ ·« Ì”„Õ » —ÕÌ· «·„»Ì⁄« "
        Exit Sub
    End If
    
    
    If IsDate(sDate) Then
        sMsg = "”Ê› Ì „  €Ì— «· «—ÌŒ «·Õ«·Ï ≈·Ï " & myFormat_p(sDate)
        If MsgBox(sMsg, vbOKCancel + vbDefaultButton2) = vbOK Then
            con.BeginTrans
            
            On Error GoTo myError
            con.Execute " UPDATE FILE0_51  SET ISCLOSED = 1 WHERE ISCLOSED = 0 "
            con.Execute " UPDATE FILE8_20H SET ISCLOSED = 1 WHERE ISCLOSED = 0 "
            con.Execute " UPDATE FILE8_70H SET ISCLOSED = 1 WHERE ISCLOSED = 0 AND BRANCH = " & MyParn(cBranch)
            con.Execute " UPDATE DSALES SET DSALES = " & addDate(sDate)
            con.CommitTrans
             
            dSalesDate = fnDateSales(con)
            Firsttitle = cComp_Name & myFormat_p(dSalesDate)
            Main.Caption = Firsttitle
            Me.Caption = Format(dSalesDate, "DD-MM-YYYY")
            
            MsgBox " „  €Ì— «· «—ÌŒ"
        End If
    End If
    Exit Sub
myError:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Function CheckOpen() As Long
Dim cString As String, nCount As Long
cString = "Select count(*) from file6_20h"
cString = cString & turn(cString) & "File6_20h.printed = 0"
cString = cString & turn(cString) & " [DATE] = " & DateSq(dSalesDate)
nCount = Val(GetField(cString, con) & "")
If nCount > 0 Then
    If lMsg Then
        MsgBox "Â‰«ﬂ ⁄œœ " & nCount & " »Ê‰«  »Ì⁄ „› ÊÕ…!!«·—Ã«¡ «·Õ–› «Ê «· ”ÃÌ·", vbCritical
    End If
    CheckOpen = nCount
End If
End Function
Private Function CheckIsNew() As Long
Dim cString As String, nCount As Long
cString = "Select count(*) from file6_20h"
cString = cString & turn(cString) & "File6_20h.isNew = 1"
cString = cString & turn(cString) & " [DATE] = " & DateSq(dSalesDate)
nCount = Val(myField(cString, con) & "")
If nCount > 0 Then
    If lMsg Then
        MsgBox "Â‰«ﬂ ⁄œœ " & nCount & " »Ê‰«  €Ì— „—Õ·…", vbCritical
    End If
    CheckIsNew = nCount
End If
End Function
Sub opencom()
On Error GoTo myError
Dim i As Integer
Dim comFlag As Boolean
If cComPort = "" Then Exit Sub
  comFlag = False
    mOpenUSBpd
  IntrCh23$ = Chr$(&H23)
  IntrCh24$ = Chr$(&H24)
  IntrCh40$ = Chr$(&H40)
  IntrCh5B$ = Chr$(&H5B)
  IntrCh5C$ = Chr$(&H5C)
  IntrCh5D$ = Chr$(&H5D)
  IntrCh5E$ = Chr$(&H5E)
  IntrCh60$ = Chr$(&H60)
  IntrCh7B$ = Chr$(&H7B)
  IntrCh7C$ = Chr$(&H7C)
  IntrCh7D$ = Chr$(&H7D)
  IntrCh7E$ = Chr$(&H7E)
   lDisplayCom = True
  Exit Sub
myError:
lDisplayCom = False
  Err.Clear
End Sub
Sub ClearDisplay()
    On Error GoTo myError
    Dim i, Pagei, nn As Integer
    Dim ss As String
    Dim TmpTime$
    Dim TmpMM, TmpHH, TmpSS
    If cComPort = "" Then Exit Sub
    If Not lDisplayCom Then Exit Sub
    DoEvents
    Call Epson_Init
    ss = "CLEAR LINE1" + Chr$(&HA)
    ss = ss + "and LINE2"
    DoEvents
'    Sleep (1000)
    ss = Chr$(&HC)
    DoEvents
    Exit Sub
myError:
    Err.Clear
End Sub
Private Sub Epson_Init()
Dim ss As String
    ss = Chr$(&H1B) + Chr$(&H40)
    DoEvents
    mWritePD ss, Len(ss)
    DoEvents
End Sub
Function mWritePD(mdata As String, ByVal mlength As Long) As Long
    If cComPort = "USB" Then
       mWritePD = WritePD(mdata, mlength)
    ElseIf Left(cComPort, 3) = "COM" Then
       MSComm1.Output = mdata
    End If
End Function
Function mOpenUSBpd() As Long
    MSComm1.CommPort = Mid$(cComPort, 4, 2)
    MSComm1.PortOpen = True
End Function
Sub myDisplay(pLine1, pLine2)
    Dim ss As String
    If cBranch = "00" Then Exit Sub
    If Not lDisplayCom Then Exit Sub
    ClearDisplay
    ss = pLine1
    ss = ss + Chr$(&H1F) + Chr$(&H43) + Chr$(&H0)
    mWritePD ss, Len(ss)
    DoEvents
    ss = vbCrLf + pLine2
    ss = ss + Chr$(&H1F) + Chr$(&H43) + Chr$(&H1)
    mWritePD ss, Len(ss)
    DoEvents
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
    aHeader(1) = "[" & xBox.text & "]"
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
cString = "SELECT   FILE0_50.DESCA AS BOXDESCA , SUM(BOXMOVE.PLUS ) AS T_IN, SUM(BOXMOVE.MINUS ) AS T_OUT, BOXMOVE.BOX, BOXMOVE.DESCA as desca, BOXMOVE.FLAG FROM         BOXMOVE INNER JOIN FILE0_50 ON FILE0_50.CODE = BOXMOVE.BOX WHERE  DATE = " & DateSq(xDate.text)
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
    temptable!str5 = TurnValue(!desca)
    
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
        temptable!str5 = TurnValue(!desca)
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

    
    
cString = "SELECT   FILE0_50.DESCA AS BOXDESCA , BOXMOVE.BOX , SUM(BOXMOVE.PLUS - BOXMOVE.MINUS ) AS BALBOX FROM         BOXMOVE INNER JOIN FILE0_50 ON FILE0_50.CODE = BOXMOVE.BOX WHERE  DATE = " & DateSq(xDate.text)
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
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Function IsPrinted() As Boolean
'IsPrinted = TurnValue(GetField("SELECT PRINTED FROM FILE6_20H WHERE DOC_NO = " & MyParn(xDoc_No.text), con), Null, False)
End Function
Private Sub cmd_disc_Click()
Dim cPassword As String, aUser As Variant
If xPrinted.Value <> 0 Then Exit Sub

Dim oUser As New sub_userfrm
Set oUser.myForm = Me
oUser.Show 1
If subUserCode <> "" Then
    MsgBox "Ì„ﬂ‰  ”ÃÌ· «·Œ’„ "
    
    xDiscount.Enabled = True
    xRate.Enabled = True
    
    XUSERNAME_DISC.Caption = subUserName
    
    xNote_disc.Caption = ""
    xdoc_disc.Caption = ""
End If
End Sub
Private Sub myUndo()
If xDoc_No.Tag = DefineMode Then
    If Not openCardTable Then
        CmdNewInv_Click
    End If
Else
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdNext_Click()
If Not openCardTable(tbMode.tbNext, xDoc_No.text) Then
    If Not openCardTable(tbMode.tblast) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdPrevious_Click()
If Not openCardTable(tbMode.tbPrevious, xDoc_No.text) Then
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
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pDoc_no As String = "", Optional pDoc_no_Filter As String = "") As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
Dim nRecords As Long, nRecord As Long

If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pDoc_no) = "" And Trim(pDoc_no_Filter) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass


On Error GoTo myError
cFilter = retFilter

Set CardTable = New ADODB.Recordset
cString = "SELECT TOP 1 FILE6_20H.*," & _
          " FILE3_10.DESCA AS CLIENTDESCA, " & _
          "dbo.fn_offer(FILE6_20H.DOC_NO) AS isOffer" & _
          " FROM FILE6_20H " & _
          " INNER JOIN FILE3_10 ON FILE3_10.CODE = FILE6_20H.code" & _
          " INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE"
  
If pMode = tbMode.tbFirst Then
    cOrder = "Order by FILE6_20H.DOC_NO"
ElseIf pMode = tbMode.tblast Then
    cOrder = "Order by FILE6_20H.DOC_NO DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "FILE6_20H.Doc_no = " & MyParn(pDoc_no)
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "FILE6_20H.DOC_NO < " & MyParn(pDoc_no)
    cOrder = "Order by FILE6_20H.DOC_NO DESC"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "FILE6_20H.DOC_NO > " & MyParn(pDoc_no)
    cOrder = "Order by FILE6_20H.DOC_NO"
End If

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

If cFilter <> "" Then
    cString = cString & IIf(cWhere = "", " WHERE ", " AND ") & cFilter
End If

cString = cString & " " & cOrder

Set CardTable = cmd(cString, con).Execute

If (Not CardTable.EOF) Then
    myload
    openCardTable = True
End If
Me.MousePointer = vbNormal
Exit Function
myError:
Me.MousePointer = vbNormal
MsgBox Err.Description
Err.Clear
End Function
Private Function DoBackUp() As Boolean
On Error GoTo myError
Dim fs As FileSystemObject, cDir As String, cFileName As String, aDrive As Variant, cBackDrive  As String
Set fs = CreateObject("Scripting.FileSystemObject")
cBackDrive = GetDesca("SELECT BACK FROM ADDRESS", con)
If cBackDrive = "" Then cBackDrive = "D"
'aDrive = aLastDrive(False)
'cDir = retFlag(aDrive, "LETTER") & ":\DataBackup"
cFileName = cBackDrive & ":\" & sCatalog & "_BACK.bak"
'createBackUp (cFileName)
Exit Function
myError:
MsgBox Err.Description
Err.Clear
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
Private Function validOpen() As Boolean
If nUser < enUser.Super Then
    If cBranchBox = "" Then
        MsgBox "·«  ÊÃœ Œ“Ì… „”Ã·…"
        Exit Function
    End If
End If
validOpen = True
End Function
Private Sub fixDisplay()
If cBranch = "00" Then Exit Sub
    
Dim loctable As New ADODB.Recordset
Set loctable = cmd("SELECT DISPLAY , PORT , CountPrint FROM COMP ", con).Execute
If Not loctable.EOF Then
    cMsgDisplay = loctable!display & ""
    cComPort = loctable!Port & ""
    nCountPrint = loctable!COUNTPRINT
End If
If nCountPrint = 0 Then nCountPrint = 1
    
opencom
ClearDisplay
myDisplay cMsgDisplay, " "
End Sub
Private Function retRecords(pDoc_no, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pDoc_no <> "" Then
    cString = "SELECT Count(FILE6_20H.DOC_NO) AS records,COUNT(CASE WHEN FILE6_20H.DOC_NO <= " & MyParn(pDoc_no) & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(FILE6_20H.DOC_NO) AS records,0 as record"
End If

cString = cString & " FROM FILE6_20H " & _
                    " INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE"

If cFilter <> "" Then
    cString = cString & " WHERE " & cFilter
End If

Set loctable = cmd(cString, con).Execute
If Not loctable.EOF Then
    nRecords = loctable!RECORDS
    nRecord = Val(loctable!Record & "")
End If
End Function
Private Function retFilter()
retFilter = "FILE6_20H.ONLINE = 0"
retFilter = retFilter & " AND FILE6_20H.BRANCH = " & MyParn(cBranch)

If nUser < enUser.Admin Then
    retFilter = retFilter & " AND FILE6_20H.DATE = " & DateSq(fnDateSales(con))
End If

If nUser = enUser.User Then
    retFilter = retFilter & " AND FILE6_20H.BOX = " & MyParn(cBranchBox)
    retFilter = retFilter & " AND FILE6_20H.PRINTED = 0"
ElseIf optclosed(1).Value Then
    retFilter = retFilter & " AND FILE6_20H.PRINTED = 0"
ElseIf optclosed(2).Value Then
    retFilter = retFilter & " AND FILE6_20H.PRINTED = 1"
End If

If optType(1).Value Then
    retFilter = retFilter & " AND FILE6_20H.IS_RETURN = 0"
ElseIf optType(2).Value Then
    retFilter = retFilter & " AND FILE6_20H.IS_RETURN = 1"
End If

End Function
Private Sub xdoc_ret_GotFocus()
myGotFocus xdoc_ret
End Sub
Private Sub xdoc_ret_LostFocus()
myLostFocus xdoc_ret
End Sub
Private Sub XDATE1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub XDATE1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
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
Private Sub xInv_No_GotFocus()
myGotFocus xinv_no
End Sub
Private Sub xInv_No_LostFocus()
myLostFocus xinv_no
End Sub
Private Sub xcard_disc_GotFocus()
myGotFocus xcard_disc
End Sub
Private Sub xDoc_no2_LostFocus()
myLostFocus xDoc_no2
End Sub
Private Sub xCode2_GotFocus()
myGotFocus xCode2
End Sub
Private Sub xCode2_LostFocus()
myLostFocus xCode2
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xdate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub xbox_GotFocus()
myGotFocus xBox
End Sub
Private Sub xbox_LostFocus()
myLostFocus xBox
If Not xBox.MatchedWithList Then xBox.BoundText = ""
End Sub
Private Sub xMAN_GotFocus()
myGotFocus xMan
End Sub
Private Sub xMan_LostFocus()
myLostFocus xMan
If xMan.MatchedWithList And bEditRecord And grid1.TextMatrix(grid1.Rows - 1, 15) = "" Then
    grid1.TextMatrix(grid1.Rows - 1, 15) = xMan.BoundText
    grid1.TextMatrix(grid1.Rows - 1, 16) = xMan.text
End If
End Sub
Private Sub xfilter_GotFocus()
myGotFocus xfilter
End Sub
Private Sub xfilter_LostFocus()
myLostFocus xfilter
End Sub
Private Sub xfilter2_GotFocus()
myGotFocus xfilter2
End Sub
Private Sub xfilter2_LostFocus()
myLostFocus xfilter2
End Sub
Private Sub xfilter3_GotFocus()
myGotFocus xfilter3
End Sub
Private Sub xfilter3_LostFocus()
myLostFocus xfilter3
End Sub
Private Sub xfilter4_GotFocus()
myGotFocus xfilter4
End Sub
Private Sub xfilter4_LostFocus()
myLostFocus xfilter4
End Sub
Private Sub xcode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xcode_LostFocus()
myLostFocus xCode
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
End Sub
Private Sub xMosm_GotFocus()
myGotFocus xMosm
End Sub
Private Sub xMosm_LostFocus()
myLostFocus xMosm
If Not xMosm.MatchedWithList Then xMosm.BoundText = ""
End Sub
Private Sub XSTORE_GotFocus()
myGotFocus xStore
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub xBoxVISA_GotFocus()
myGotFocus xBoxVISA
End Sub
Private Sub xBoxVISA_LostFocus()
myLostFocus xBoxVISA
If Not xBoxVISA.MatchedWithList Then xBoxVISA.BoundText = ""
End Sub
Private Sub myloadgrd2()
With GrdOffer
Dim cString As String
cString = "SELECT FILE6_20.OFFER_NO," & _
          "FILE6_20.DOC_OFFER," & _
          "FILE0_90H.DESCA," & _
          "FILE6_20.DISCOUNT_OFFER" & _
          " FROM FILE6_20" & _
          " INNER JOIN FILE0_90H ON FILE6_20.DOC_OFFER = FILE0_90H.DOC_NO" & _
          " WHERE FILE6_20.DOC_NO = " & MyParn(xDoc_No.text) & _
          " AND FILE6_20.OFFER_NO > 0" & _
          " GROUP BY FILE6_20.DOC_OFFER," & _
          " FILE0_90H.DESCA," & _
          " FILE6_20.DISCOUNT_OFFER," & _
          " FILE6_20.OFFER_NO" & _
          " ORDER BY FILE6_20.OFFER_NO"

    Set DATA13.Recordset = cmd(cString, con).Execute
'    GrdOffer.AddItem ""
End With
CalcTotals2
Fixgrd2
End Sub
Private Sub CalcTotals2()
Dim i As Long
Dim nTotal As Double
For i = 1 To GrdOffer.Rows - 2
    nTotal = nTotal + GrdOffer.ValueMatrix(i, 3)
Next
'xDiscount.text = nTotal
End Sub
Private Sub HandleInit()
If lSupperVisor Then
    nUser = enUser.Admin
ElseIf cBranch = "00" Then
    nUser = enUser.Admin
ElseIf bopt2 Then
    nUser = enUser.Super
Else
    nUser = enUser.User
End If

bedit = cBranchBox <> "" Or (nUser = enUser.Admin)
cmdPrintDay.Enabled = nUser > enUser.User
'fmClose.Visible = nUser >= enUser.Super
'fmPrint.Visible = nUser >= enUser.Admin
'cmdShowPhones.Enabled = nUser = enUser.Admin
End Sub

Private Sub xTotal_LostFocus()
    myLostFocus xTotal
End Sub

Private Sub xTotalItemDiscount_Change()
xtotalitem.Visible = Val(xTotalItemDiscount) <> 0
lblTotalitem.Visible = Val(xTotalItemDiscount) <> 0
End Sub

Private Sub xTotalQuant_Change()
xTotalQuant.ForeColor = IIf(Val(xTotalQuant.Caption) < 0, vbRed, vbBlack)
lblQuant.ForeColor = IIf(Val(xTotalQuant.Caption) < 0, vbRed, vbBlack)
End Sub

Private Sub xType_Change()
xType.Visible = xType.Caption <> ""
End Sub
Public Function sendInvoicesNew(Optional ByRef nRecords As Long, Optional bNoPrint As Boolean = False, Optional pDoc_no As String = "") As Boolean
Dim conServer As New ADODB.Connection
If openCon(conServer, LoadConServer, 2) <> "ok" Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ ··œŒÊ· ··›—⁄ «·—∆Ì”Ì ·«—”«· «·›Ê« Ì—"
    Exit Function
End If

Dim loctable As New ADODB.Recordset
Dim cString As String
cString = "SELECT FILE6_20H.* " & _
          " FROM FILE6_20H"

If bNoPrint Then
    cWhere = " FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
Else
    cWhere = " (FILE6_20H.PRINTED = 1 and FILE6_20H.ISNEW = 1)"
    If pDoc_no <> "" Then
        cWhere = cWhere & " OR FILE6_20H.doc_no = " & MyParn(pDoc_no)
    End If
End If

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

Set loctable = cmd(cString, con).Execute

If loctable.EOF Then
    loctable.Close
    sendInvoicesNew = True
    Exit Function
End If

Dim strTab As New ChilkatStringTable
Dim sbList As New ChilkatStringBuilder

Dim aInsert As Variant
sCaption = Me.Caption
Do Until loctable.EOF
    Me.Caption = sCaption & " - " & loctable.AbsolutePosition & " „‰ " & loctable.RecordCount
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "DATE", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!code))
    aInsert = AddFlag(aInsert, "STORE", addstring(loctable!STORE))
    aInsert = AddFlag(aInsert, "[NOTES]", addstring(loctable!NOTES))
    aInsert = AddFlag(aInsert, "[RATE]", Val(loctable!Rate & ""))
    
    aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount))
    aInsert = AddFlag(aInsert, "[DISCOUNT_ADD]", Val(loctable!discount_add))
    aInsert = AddFlag(aInsert, "[DISCOUNT_OFFER]", Val(loctable!discount_offer))
    aInsert = AddFlag(aInsert, "[TOTAL_ITEM]", Val(loctable!TOTAL_ITEM))
    aInsert = AddFlag(aInsert, "[DOC_REPLACE]", addstring(loctable!Doc_REPLACE))
        
    aInsert = AddFlag(aInsert, "[CASH]", Val(loctable!CASH))
    aInsert = AddFlag(aInsert, "[BOX]", addstring(loctable!BOX))
    aInsert = AddFlag(aInsert, "[PAY]", Val(loctable!PAY))
    aInsert = AddFlag(aInsert, "[LATE]", "0")
    aInsert = AddFlag(aInsert, "[REST]", Val(loctable!rest & ""))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(loctable!branch))
    aInsert = AddFlag(aInsert, "[MAN]", addstring(loctable!MAN))
    aInsert = AddFlag(aInsert, "[TIME]", addTimeString(loctable!Time & ""))
    aInsert = AddFlag(aInsert, "[PHONE]", addstring(loctable!Phone))
    aInsert = AddFlag(aInsert, "[USERNAME_DISC]", addstring(loctable!username_disc))
    aInsert = AddFlag(aInsert, "[USERNAME_RET]", addstring(loctable!username_ret))
    aInsert = AddFlag(aInsert, "[SALES_RET]", addstring(loctable!SALES_RET))
    aInsert = AddFlag(aInsert, "[NOTE_DISC]", addstring(loctable!NOTE_DISC))
    aInsert = AddFlag(aInsert, "[DOC_DISC]", addstring(loctable!DOC_DISC))
    aInsert = AddFlag(aInsert, "[INV_NO]", addstring(loctable!INV_NO))
    aInsert = AddFlag(aInsert, "[CARD_DISC]", addstring(loctable!card_disc))
    aInsert = AddFlag(aInsert, "[CARD_DESCA]", addstring(loctable!card_desca))
    aInsert = AddFlag(aInsert, "[CARD_VALUE]", addstring(loctable!card_VALUE))
    aInsert = AddFlag(aInsert, "[CARD_DOC]", addstring(loctable!card_DOC))
    aInsert = AddFlag(aInsert, "[PRINTED]", IIf(loctable!PRINTED, 1, 0))
    aInsert = AddFlag(aInsert, "[CLOSED]", IIf(loctable!CLOSED, 1, 0))
    
    strTab.Append "DELETE FROM FR6_20 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    strTab.Append "DELETE FROM FR6_20H WHERE DOC_NO = " & MyParn(loctable!doc_no)
    strTab.Append addInsert(aInsert, "FR6_20H")
    
    If sbList.Length > 0 Then sbList.Append ","
    sbList.Append MyParn(loctable!doc_no)
    
    nRecords = nRecords + 1
    loctable.MoveNext
Loop


cString = "SELECT FILE6_20.* " & _
          " FROM FILE6_20 " & _
          " INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO"

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

Set loctable = cmd(cString, con).Execute
Do Until loctable.EOF
    Me.Caption = sCaption & " - " & loctable.AbsolutePosition & " „‰ " & loctable.RecordCount
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "ITEM", addstring(loctable!Item))
    aInsert = AddFlag(aInsert, "QUANT", Val(loctable!Quant))
    aInsert = AddFlag(aInsert, "PRICE", loctable!price)
    aInsert = AddFlag(aInsert, "PRICE_c", loctable!PRICE_C)
    aInsert = AddFlag(aInsert, "COST", Val(loctable!cost & ""))
    aInsert = AddFlag(aInsert, "MAN", addstring(loctable!MAN))
    aInsert = AddFlag(aInsert, "S_OKAZ", Val(loctable!S_OKAZ & ""))
    
    aInsert = AddFlag(aInsert, "OFFER_NO", addstring(loctable!offer_no))
    aInsert = AddFlag(aInsert, "DOC_OFFER", addstring(loctable!DOC_OFFER))
    aInsert = AddFlag(aInsert, "DISCOUNT_OFFER", loctable!discount_offer)
    aInsert = AddFlag(aInsert, "DISCOUNT_OFFER_RATE", loctable!discount_offer_rate)
    
    strTab.Append addInsert(aInsert, "FR6_20") & ";"
    loctable.MoveNext
Loop
Me.Caption = sCaption

Dim i As Long
prog1.Visible = True
conServer.BeginTrans
For i = 0 To strTab.Count - 1
    Me.Caption = sCaption & " - " & (i) & " „‰ " & ((strTab.Count))
    prog1.Value = Round(i / (strTab.Count - 1), 2) * 100
    conServer.Execute strTab.GetStrings(i, 1, 0), nAffect
Next
con.Execute "UPDATE FILE6_20H SET FILE6_20H.ISNEW = 0 WHERE FILE6_20H.DOC_NO IN (" & sbList.GetAsString & ")", nAffect
conServer.CommitTrans
sendInvoicesNew = True
Me.Caption = sCaption
prog1.Value = 0
prog1.Visible = False
Finally:
loctable.Close
closeCon conServer
Exit Function
MsgBox Err.Description
Err.Clear
If TransCount(conServer) > 0 Then
   conServer.RollbackTrans
End If
GoTo Finally
End Function
Private Function sendPhones(ByRef nRecords As Long) As Boolean
Dim conServer As New ADODB.Connection
If openCon(conServer, LoadConServer, 2) <> "ok" Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ ··œŒÊ· ··›—⁄ «·—∆Ì”Ì ·«—”«· «· ·Ì›Ê‰« "
    Exit Function
End If

Dim loctable As New ADODB.Recordset
Set loctable = cmd("select * from subcust where isNew = 1", con).Execute

If loctable.EOF Then
    loctable.Close
    Exit Function
End If

Dim strTab As New ChilkatStringTable
Dim sbList As New ChilkatStringBuilder

Dim aInsert As Variant
sCaption = Me.Caption
Do Until loctable.EOF
    Me.Caption = sCaption & " - " & loctable.AbsolutePosition & " „‰ " & loctable.RecordCount
    If sbList.Length > 0 Then sbList.Append ","
    sbList.Append MyParn(loctable!Phone)
    If UpdatePhones(locable!Phone, locatable!Name, myFormat(loctable!F_DATE), conServer) Then
        sbList.Append MyParn(loctable!Phone)
        nRecords = nRecords + 1
    End If
    loctable.MoveNext
Loop
sendPhontes = True
Me.Caption = sCaption
prog1.Value = 0
prog1.Visible = False
Finally:
loctable.Close
closeCon conServer
Exit Function
MsgBox Err.Description
Err.Clear
If TransCount(conServer) > 0 Then
   conServer.RollbackTrans
End If
GoTo Finally
End Function

