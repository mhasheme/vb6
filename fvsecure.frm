VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSECURE 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Security"
   ClientHeight    =   10650
   ClientLeft      =   105
   ClientTop       =   -10200
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10650
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   50
      SmallChange     =   4
      TabIndex        =   450
      Top             =   10680
      Width           =   12375
   End
   Begin VB.VScrollBar scrControl 
      Height          =   6855
      LargeChange     =   300
      Left            =   12720
      Max             =   4000
      SmallChange     =   300
      TabIndex        =   397
      Top             =   3120
      Value           =   50
      Width           =   255
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   89
      Top             =   10095
      Visible         =   0   'False
      Width           =   15120
      _Version        =   65536
      _ExtentX        =   26670
      _ExtentY        =   979
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Alignment       =   6
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8790
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   405
         Left            =   3000
         Top             =   30
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   714
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   87
      Top             =   0
      Width           =   15120
      _Version        =   65536
      _ExtentX        =   26670
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BevelInner      =   2
      Font3D          =   1
      Begin VB.CommandButton cmdFindUser 
         Caption         =   "Find User"
         Height          =   330
         Left            =   10800
         TabIndex        =   100
         Top             =   80
         Width           =   1095
      End
      Begin VB.CommandButton cmdCopySecuritys 
         Caption         =   "Copy Security Settings"
         Height          =   330
         Left            =   8160
         TabIndex        =   99
         Top             =   80
         Width           =   2415
      End
      Begin Threed.SSPanel Panel3D2 
         Height          =   1170
         Left            =   0
         TabIndex        =   90
         Top             =   525
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         Caption         =   "Panel3D2"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         BevelInner      =   1
         Font3D          =   1
         Alignment       =   1
      End
      Begin VB.Label lblUSERID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "USERID"
         DataField       =   "USERID"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   750
         TabIndex        =   96
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "USERNAME"
         DataField       =   "USERName"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2280
         TabIndex        =   88
         Top             =   120
         Width           =   1770
      End
   End
   Begin Threed.SSCheck chkUISecurity 
      Height          =   225
      Index           =   21
      Left            =   870
      TabIndex        =   97
      Top             =   0
      Visible         =   0   'False
      Width           =   2565
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Product Line/Operation"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSCheck chkUMSecurity 
      Height          =   225
      Index           =   21
      Left            =   0
      TabIndex        =   98
      Top             =   0
      Visible         =   0   'False
      Width           =   435
      _Version        =   65536
      _ExtentX        =   767
      _ExtentY        =   397
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   2520
      Left            =   0
      TabIndex        =   91
      Top             =   495
      Width           =   15120
      _Version        =   65536
      _ExtentX        =   26670
      _ExtentY        =   4445
      _StockProps     =   15
      ForeColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Begin VB.CommandButton cmdScreenRight 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11190
         Picture         =   "fvsecure.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   547
         Tag             =   "Next Security Screen"
         Top             =   2115
         Width           =   705
      End
      Begin VB.CommandButton cmdScreenLeft 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10320
         Picture         =   "fvsecure.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   545
         Tag             =   "Previous Security Screen"
         Top             =   2115
         Width           =   705
      End
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   285
         Left            =   3420
         TabIndex        =   1
         Top             =   2130
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         RefreshDescriptionWhen=   2
      End
      Begin VB.TextBox txtEEName 
         Appearance      =   0  'Flat
         DataField       =   "USERNAME"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6000
         TabIndex        =   2
         Tag             =   "00-Name"
         Text            =   "txtEEID"
         Top             =   2130
         Width           =   3135
      End
      Begin VB.TextBox txtUSERID 
         Appearance      =   0  'Flat
         DataField       =   "USERID"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   720
         MaxLength       =   25
         TabIndex        =   0
         Tag             =   "00-User ID"
         Text            =   "txtUSERID"
         Top             =   2130
         Width           =   1425
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
         Bindings        =   "fvsecure.frx":0884
         Height          =   2010
         Left            =   0
         Negotiate       =   -1  'True
         OleObjectBlob   =   "fvsecure.frx":0898
         TabIndex        =   546
         Tag             =   "Listing of Security Records"
         Top             =   0
         Width           =   11895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5280
         TabIndex        =   95
         Top             =   2175
         Width           =   615
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   94
         Top             =   2175
         Width           =   660
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         DataField       =   "EMPNBR"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3720
         TabIndex        =   93
         Top             =   2520
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblEENumber 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2400
         TabIndex        =   92
         Top             =   2175
         Width           =   840
      End
   End
   Begin Threed.SSPanel panWindow 
      Height          =   7680
      Left            =   0
      TabIndex        =   101
      Top             =   3120
      Width           =   12600
      _Version        =   65536
      _ExtentX        =   22225
      _ExtentY        =   13547
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.PictureBox panDetails 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   5
         Left            =   120
         ScaleHeight     =   2385
         ScaleWidth      =   2985
         TabIndex        =   329
         Top             =   2880
         Width           =   3015
         Begin VB.CommandButton cmdPageRight 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   5
            Left            =   9960
            Picture         =   "fvsecure.frx":B9E8
            Style           =   1  'Graphical
            TabIndex        =   334
            Tag             =   "Grant All Basic"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdPageLeft 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   5
            Left            =   9120
            Picture         =   "fvsecure.frx":BE2A
            Style           =   1  'Graphical
            TabIndex        =   333
            Tag             =   "Grant All Basic"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdGrandInquAT 
            Appearance      =   0  'Flat
            Caption         =   "Grant Inquire"
            Height          =   450
            Left            =   5760
            TabIndex        =   332
            Tag             =   "Grant All Applicant Tracker"
            Top             =   2400
            Width           =   2865
         End
         Begin VB.CommandButton cmdRemoveAllAT 
            Appearance      =   0  'Flat
            Caption         =   "Remove All "
            Height          =   450
            Left            =   5760
            TabIndex        =   331
            Tag             =   "Grant All Applicant Tracker"
            Top             =   2880
            Width           =   2865
         End
         Begin VB.CommandButton cmdGrantAllA 
            Appearance      =   0  'Flat
            Caption         =   "Grant All for Applicant Tracker"
            Height          =   450
            Left            =   5760
            TabIndex        =   330
            Tag             =   "Grant All Applicant Tracker"
            Top             =   1920
            Width           =   2865
         End
         Begin Threed.SSCheck chkAMSecurity 
            Height          =   225
            Index           =   5
            Left            =   0
            TabIndex        =   335
            Top             =   1700
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAISecurity 
            Height          =   225
            Index           =   5
            Left            =   1080
            TabIndex        =   336
            Top             =   1700
            Width           =   2325
            _Version        =   65536
            _ExtentX        =   4101
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Applicant Associations"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAMSecurity 
            Height          =   225
            Index           =   7
            Left            =   0
            TabIndex        =   337
            Top             =   2115
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAISecurity 
            Height          =   225
            Index           =   7
            Left            =   1080
            TabIndex        =   338
            Top             =   2115
            Width           =   2205
            _Version        =   65536
            _ExtentX        =   3889
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Applicant Follow Ups"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAISecurity 
            Height          =   225
            Index           =   4
            Left            =   1080
            TabIndex        =   339
            Top             =   1492
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Applicant Interviews"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAISecurity 
            Height          =   225
            Index           =   8
            Left            =   6960
            TabIndex        =   340
            Top             =   660
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Requisitions"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAISecurity 
            Height          =   225
            Index           =   6
            Left            =   1080
            TabIndex        =   341
            Top             =   1908
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Applicant References"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAISecurity 
            Height          =   225
            Index           =   9
            Left            =   6960
            TabIndex        =   342
            Top             =   868
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   3254
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Recruitment Cost"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAISecurity 
            Height          =   225
            Index           =   0
            Left            =   1080
            TabIndex        =   343
            Top             =   660
            Width           =   4485
            _Version        =   65536
            _ExtentX        =   7911
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Applicant Demographics/Dates/References/Hire"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAISecurity 
            Height          =   225
            Index           =   1
            Left            =   1080
            TabIndex        =   344
            Top             =   868
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Applicant Skills"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAISecurity 
            Height          =   225
            Index           =   2
            Left            =   1080
            TabIndex        =   345
            Top             =   1076
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4736
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Applicant Formal Education"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAISecurity 
            Height          =   225
            Index           =   3
            Left            =   1080
            TabIndex        =   346
            Top             =   1284
            Width           =   2925
            _Version        =   65536
            _ExtentX        =   5159
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Applicant Continuing Education"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAMSecurity 
            Height          =   225
            Index           =   4
            Left            =   0
            TabIndex        =   347
            Top             =   1492
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAMSecurity 
            Height          =   225
            Index           =   9
            Left            =   5880
            TabIndex        =   348
            Top             =   868
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAMSecurity 
            Height          =   225
            Index           =   8
            Left            =   5880
            TabIndex        =   349
            Top             =   660
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAMSecurity 
            Height          =   225
            Index           =   6
            Left            =   0
            TabIndex        =   350
            Top             =   1908
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAMSecurity 
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   351
            Top             =   660
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAMSecurity 
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   352
            Top             =   868
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAMSecurity 
            Height          =   225
            Index           =   2
            Left            =   0
            TabIndex        =   353
            Top             =   1076
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAMSecurity 
            Height          =   225
            Index           =   3
            Left            =   0
            TabIndex        =   354
            Top             =   1284
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAISecurity 
            Height          =   225
            Index           =   10
            Left            =   6960
            TabIndex        =   492
            Top             =   1076
            Width           =   2205
            _Version        =   65536
            _ExtentX        =   3889
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Employment History"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAMSecurity 
            Height          =   228
            Index           =   10
            Left            =   5880
            TabIndex        =   493
            Top             =   1074
            Width           =   432
            _Version        =   65536
            _ExtentX        =   762
            _ExtentY        =   402
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAMSecurity 
            Height          =   225
            Index           =   11
            Left            =   0
            TabIndex        =   556
            Top             =   3000
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAISecurity 
            Height          =   225
            Index           =   11
            Left            =   1080
            TabIndex        =   557
            Top             =   3000
            Width           =   2445
            _Version        =   65536
            _ExtentX        =   4313
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Letters by Position Type"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAMSecurity 
            Height          =   225
            Index           =   12
            Left            =   0
            TabIndex        =   558
            Top             =   2760
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAISecurity 
            Height          =   225
            Index           =   12
            Left            =   1080
            TabIndex        =   559
            Top             =   2760
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4736
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Application Form Workflow"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAMSecurity 
            Height          =   225
            Index           =   13
            Left            =   0
            TabIndex        =   560
            Top             =   2520
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkAISecurity 
            Height          =   225
            Index           =   13
            Left            =   1080
            TabIndex        =   561
            Top             =   2520
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4736
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Application Form Defaults"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Applicant Tracker"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   21
            Left            =   0
            TabIndex        =   359
            Top             =   0
            Width           =   1860
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inquire"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   9
            Left            =   6960
            TabIndex        =   358
            Top             =   360
            Width           =   600
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Maintain"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   8
            Left            =   5880
            TabIndex        =   357
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Maintain"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   356
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inquire"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   1080
            TabIndex        =   355
            Top             =   360
            Width           =   600
         End
      End
      Begin VB.PictureBox panDetails 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2295
         Index           =   0
         Left            =   120
         ScaleHeight     =   2265
         ScaleWidth      =   1995
         TabIndex        =   314
         Top             =   120
         Width           =   2025
         Begin Threed.SSCheck chkViewOwnCounsel 
            DataSource      =   "Data1"
            Height          =   225
            Left            =   8400
            TabIndex        =   55
            Tag             =   "40-View Own"
            Top             =   1806
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "View Own"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkViewOwnOthInfo 
            DataSource      =   "Data1"
            Height          =   225
            Left            =   8400
            TabIndex        =   58
            Tag             =   "40-View Own"
            Top             =   2027
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "View Own"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin VB.ComboBox cmbTemplate 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Tag             =   "40-Country Code"
            Top             =   4680
            Width           =   1600
         End
         Begin VB.ComboBox cmbSecTemplate 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Tag             =   "40-Security Template"
            Top             =   5080
            Width           =   1600
         End
         Begin INFOHR_Controls.DateLookup ExpireDate 
            DataField       =   "PS_EXPIR_DATE"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   6240
            TabIndex        =   84
            Tag             =   "40-Division Effective Date"
            Top             =   5520
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            TextBoxWidth    =   1215
            Enabled         =   0   'False
         End
         Begin INFOHR_Controls.DateLookup LastExpireDate 
            DataField       =   "PS_CHGDATE"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   9120
            TabIndex        =   85
            Tag             =   "40-Last Change Date"
            Top             =   5520
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            TextBoxWidth    =   1215
            Enabled         =   0   'False
         End
         Begin VB.TextBox txtExpireDays 
            Appearance      =   0  'Flat
            DataField       =   "PS_EXPIR_DAYS"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   3960
            TabIndex        =   83
            Top             =   5520
            Width           =   735
         End
         Begin VB.TextBox Updstats 
            Appearance      =   0  'Flat
            DataField       =   "LDATE"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   10920
            MaxLength       =   25
            TabIndex        =   319
            TabStop         =   0   'False
            Top             =   840
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox Updstats 
            Appearance      =   0  'Flat
            DataField       =   "LTIME"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   10920
            MaxLength       =   25
            TabIndex        =   318
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox Updstats 
            Appearance      =   0  'Flat
            DataField       =   "LUSER"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   10920
            MaxLength       =   25
            TabIndex        =   317
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton cmdGrantAllB 
            Appearance      =   0  'Flat
            Caption         =   "&Grant All Basic 1"
            Height          =   330
            Left            =   7680
            TabIndex        =   73
            Tag             =   "Grant All Basic"
            Top             =   3870
            Width           =   1800
         End
         Begin VB.TextBox txtPWord 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   960
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   82
            Tag             =   "00-Password"
            Top             =   5520
            Width           =   1305
         End
         Begin VB.TextBox txtSecPassword 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            DataField       =   "PassWord"
            DataSource      =   "Data1"
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   9240
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   315
            Tag             =   "00-Password"
            Top             =   4320
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton cmdGrantInqu 
            Appearance      =   0  'Flat
            Caption         =   "Grant All &Inquire"
            Height          =   330
            Left            =   5700
            TabIndex        =   72
            Tag             =   "Grant All Basic"
            Top             =   3870
            Width           =   1800
         End
         Begin VB.CommandButton cmdRemoveAll 
            Appearance      =   0  'Flat
            Caption         =   "&Remove All Basic 1"
            Height          =   330
            Left            =   3600
            TabIndex        =   71
            Tag             =   "Grant All Basic"
            Top             =   3870
            Width           =   1920
         End
         Begin VB.ComboBox cmbCountry 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   74
            Tag             =   "40-Country Code"
            Text            =   "cmbCountry"
            Top             =   4290
            Width           =   1600
         End
         Begin VB.CommandButton cmdPageRight 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   9960
            Picture         =   "fvsecure.frx":C26C
            Style           =   1  'Graphical
            TabIndex        =   86
            Tag             =   "Grant All Basic"
            Top             =   0
            Width           =   705
         End
         Begin Threed.SSCheck chkMSecurity 
            Height          =   225
            Index           =   26
            Left            =   5535
            TabIndex        =   56
            Top             =   2027
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkSecurity 
            Height          =   225
            Index           =   26
            Left            =   6510
            TabIndex        =   57
            Top             =   2027
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Other Infomation"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Height          =   225
            Index           =   24
            Left            =   5535
            TabIndex        =   67
            Top             =   3132
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkSecurity 
            Height          =   225
            Index           =   24
            Left            =   6510
            TabIndex        =   68
            Top             =   3132
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Salary Grids"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C6AE
            Height          =   225
            Index           =   21
            Left            =   6510
            TabIndex        =   66
            Top             =   2910
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Position Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C6B9
            Height          =   225
            Index           =   21
            Left            =   5535
            TabIndex        =   65
            Top             =   2910
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C6C4
            Height          =   225
            Index           =   20
            Left            =   5535
            TabIndex        =   63
            Top             =   2690
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C6CF
            Height          =   225
            Index           =   19
            Left            =   6510
            TabIndex        =   62
            Top             =   2469
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Position Skills"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C6DA
            Height          =   225
            Index           =   19
            Left            =   5535
            TabIndex        =   61
            Top             =   2469
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C6E5
            Height          =   225
            Index           =   18
            Left            =   6510
            TabIndex        =   52
            Top             =   1585
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   4048
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Continuing Education"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C6F0
            Height          =   225
            Index           =   18
            Left            =   5535
            TabIndex        =   51
            Top             =   1585
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C6FB
            Height          =   225
            Index           =   17
            Left            =   5535
            TabIndex        =   59
            Top             =   2248
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C706
            Height          =   225
            Index           =   16
            Left            =   5535
            TabIndex        =   49
            Top             =   1364
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C711
            Height          =   225
            Index           =   15
            Left            =   5535
            TabIndex        =   47
            Top             =   1143
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C71C
            Height          =   225
            Index           =   14
            Left            =   5535
            TabIndex        =   45
            Top             =   922
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C727
            Height          =   225
            Index           =   13
            Left            =   5535
            TabIndex        =   40
            Top             =   480
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C732
            Height          =   195
            Index           =   12
            Left            =   15
            TabIndex        =   32
            Top             =   3392
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   344
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C73D
            Height          =   225
            Index           =   11
            Left            =   15
            TabIndex        =   29
            Top             =   3168
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C748
            Height          =   225
            Index           =   10
            Left            =   15
            TabIndex        =   27
            Top             =   2944
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C753
            Height          =   225
            Index           =   9
            Left            =   15
            TabIndex        =   25
            Top             =   2720
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Height          =   225
            Index           =   8
            Left            =   15
            TabIndex        =   21
            Top             =   2272
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Height          =   225
            Index           =   7
            Left            =   15
            TabIndex        =   19
            Top             =   2048
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Height          =   225
            Index           =   6
            Left            =   15
            TabIndex        =   17
            Top             =   1824
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Height          =   225
            Index           =   5
            Left            =   15
            TabIndex        =   15
            Top             =   1600
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Height          =   225
            Index           =   4
            Left            =   15
            TabIndex        =   13
            Top             =   1376
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Height          =   225
            Index           =   3
            Left            =   15
            TabIndex        =   11
            Top             =   1152
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Height          =   225
            Index           =   2
            Left            =   15
            TabIndex        =   8
            Top             =   928
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Height          =   225
            Index           =   1
            Left            =   15
            TabIndex        =   6
            Top             =   704
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C75E
            Height          =   225
            Index           =   15
            Left            =   6510
            TabIndex        =   48
            Top             =   1143
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Other Earnings"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C769
            Height          =   225
            Index           =   17
            Left            =   6510
            TabIndex        =   60
            Top             =   2248
            Width           =   3075
            _Version        =   65536
            _ExtentX        =   5424
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "National Occupational Class'ns"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C774
            Height          =   225
            Index           =   16
            Left            =   6510
            TabIndex        =   50
            Top             =   1364
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Terminations"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C77F
            Height          =   225
            Index           =   14
            Left            =   6510
            TabIndex        =   46
            Top             =   922
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Dollar Entitlements"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C78A
            Height          =   225
            Index           =   1
            Left            =   990
            TabIndex        =   7
            Top             =   701
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Banking"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C795
            Height          =   225
            Index           =   12
            Left            =   990
            TabIndex        =   33
            Top             =   3353
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Health and Safety"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C7A0
            Height          =   225
            Index           =   11
            Left            =   990
            TabIndex        =   30
            Top             =   3132
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Follow-Ups"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C7AB
            Height          =   225
            Index           =   10
            Left            =   990
            TabIndex        =   28
            Top             =   2910
            Width           =   1965
            _Version        =   65536
            _ExtentX        =   3466
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Associations"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C7B6
            Height          =   225
            Index           =   9
            Left            =   990
            TabIndex        =   26
            Top             =   2690
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Sick/Vacation Entitlements"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C7C1
            Height          =   225
            Index           =   8
            Left            =   990
            TabIndex        =   22
            Top             =   2248
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Benefits"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C7CC
            Height          =   225
            Index           =   7
            Left            =   990
            TabIndex        =   20
            Top             =   2027
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Position"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C7D7
            Height          =   225
            Index           =   6
            Left            =   990
            TabIndex        =   18
            Top             =   1806
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Performance"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C7E2
            Height          =   225
            Index           =   5
            Left            =   990
            TabIndex        =   16
            Top             =   1585
            Width           =   765
            _Version        =   65536
            _ExtentX        =   1349
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Salary"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C7ED
            Height          =   225
            Index           =   4
            Left            =   990
            TabIndex        =   14
            Top             =   1364
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Formal Education"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C7F8
            Height          =   225
            Index           =   2
            Left            =   990
            TabIndex        =   9
            Top             =   922
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Dependents"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C803
            Height          =   225
            Index           =   0
            Left            =   990
            TabIndex        =   4
            Top             =   480
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Employee Demographics / Dates"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Height          =   225
            Index           =   0
            Left            =   15
            TabIndex        =   3
            Top             =   480
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkEESecurity 
            DataField       =   "EmpNBR_Based"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   3840
            TabIndex        =   77
            Tag             =   "40-Employee Number Based Security -y/n"
            Top             =   4290
            Width           =   3435
            _Version        =   65536
            _ExtentX        =   6059
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "Employee Number Based Security"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C80E
            Height          =   225
            Index           =   20
            Left            =   6510
            TabIndex        =   64
            Top             =   2690
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Position Evaluation"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkEESIN 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   7320
            TabIndex        =   80
            Tag             =   "40-Show SIN/SSN -y/n"
            Top             =   4290
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "Show SIN/SSN "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C819
            Height          =   225
            Index           =   23
            Left            =   6510
            TabIndex        =   54
            Top             =   1806
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Counseling"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C824
            Height          =   225
            Index           =   23
            Left            =   5535
            TabIndex        =   53
            Top             =   1806
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C82F
            Height          =   195
            Index           =   25
            Left            =   15
            TabIndex        =   34
            Top             =   3585
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   344
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C83A
            Height          =   225
            Index           =   25
            Left            =   990
            TabIndex        =   35
            Top             =   3570
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Comments"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C845
            Height          =   225
            Index           =   13
            Left            =   6510
            TabIndex        =   41
            Top             =   480
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Attendance"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkASecurity 
            DataSource      =   "Data1"
            Height          =   225
            Left            =   7980
            TabIndex        =   42
            Tag             =   "40-Add Attendance"
            Top             =   480
            Visible         =   0   'False
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Add Attendance"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Height          =   225
            Index           =   27
            Left            =   8400
            TabIndex        =   69
            Top             =   3132
            Visible         =   0   'False
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkSecurity 
            Height          =   225
            Index           =   27
            Left            =   9375
            TabIndex        =   70
            Top             =   3132
            Visible         =   0   'False
            Width           =   2880
            _Version        =   65536
            _ExtentX        =   5080
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Versatility Chart - Production"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C850
            Height          =   225
            Index           =   28
            Left            =   5535
            TabIndex        =   43
            Top             =   701
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C85B
            Height          =   225
            Index           =   28
            Left            =   6510
            TabIndex        =   44
            Top             =   701
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Attendance History"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkEEMarital 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   3840
            TabIndex        =   79
            Tag             =   "40-Show Marital Status"
            Top             =   4800
            Width           =   2475
            _Version        =   65536
            _ExtentX        =   4366
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "Show Marital Status"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C866
            Height          =   225
            Index           =   3
            Left            =   990
            TabIndex        =   12
            Top             =   1143
            Width           =   705
            _Version        =   65536
            _ExtentX        =   1244
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Skills"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkPswdLocked 
            DataField       =   "LOCK_PASSWORD"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   0
            TabIndex        =   316
            Tag             =   "40-Check to lock the login into info:HR"
            Top             =   5880
            Visible         =   0   'False
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "Lock Password   "
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Height          =   225
            Index           =   29
            Left            =   15
            TabIndex        =   23
            Top             =   2496
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C871
            Height          =   225
            Index           =   29
            Left            =   990
            TabIndex        =   24
            Top             =   2469
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Beneficiary"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin VB.TextBox txtSecTemplate 
            Appearance      =   0  'Flat
            DataField       =   "SECURE_TEMPLATE"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   3480
            TabIndex        =   500
            Top             =   5095
            Visible         =   0   'False
            Width           =   255
         End
         Begin Threed.SSCheck chkEEDOB 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   3840
            TabIndex        =   78
            Tag             =   "40-Show Birth Date"
            Top             =   4550
            Width           =   3435
            _Version        =   65536
            _ExtentX        =   6059
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "Show Birth Date"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkEEADDRESS 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   7320
            TabIndex        =   81
            Tag             =   "40-Show Address"
            Top             =   4550
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "Show Address"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkDSecurity 
            DataSource      =   "Data1"
            Height          =   225
            Left            =   2400
            TabIndex        =   10
            Tag             =   "40-Add Attendance"
            Top             =   922
            Visible         =   0   'False
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Delete Dependents"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkNHireSecurity 
            DataSource      =   "Data1"
            Height          =   225
            Left            =   4200
            TabIndex        =   5
            Tag             =   "40-Add New Hire"
            Top             =   480
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "New Hire"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkViewOwnComm 
            DataSource      =   "Data1"
            Height          =   225
            Left            =   2760
            TabIndex        =   36
            Tag             =   "40-View Own"
            Top             =   3570
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "View Own"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkViewOwnFollUp 
            DataSource      =   "Data1"
            Height          =   225
            Left            =   2760
            TabIndex        =   31
            Tag             =   "40-View Own"
            Top             =   3132
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "View Own"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSecurity 
            Bindings        =   "fvsecure.frx":C87C
            DataSource      =   "Data2"
            Height          =   225
            Index           =   22
            Left            =   6510
            TabIndex        =   39
            Top             =   255
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Hourly Entitlements"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMSecurity 
            Bindings        =   "fvsecure.frx":C887
            Height          =   195
            Index           =   22
            Left            =   5535
            TabIndex        =   38
            Top             =   270
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   344
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkACommentSecurity 
            DataSource      =   "Data1"
            Height          =   225
            Left            =   3960
            TabIndex        =   37
            Tag             =   "40-Add Comments"
            Top             =   3570
            Visible         =   0   'False
            Width           =   2385
            _Version        =   65536
            _ExtentX        =   4207
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Add Comments"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkViewOwnPerform 
            DataSource      =   "Data1"
            Height          =   225
            Left            =   2760
            TabIndex        =   554
            Tag             =   "40-View Own"
            Top             =   1800
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "View Own"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inquire"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   1020
            TabIndex        =   518
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Security Template"
            Height          =   195
            Left            =   0
            TabIndex        =   499
            Top             =   5140
            Width           =   1545
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Timesheet Template"
            Height          =   195
            Left            =   0
            TabIndex        =   498
            Top             =   4740
            Width           =   1725
         End
         Begin VB.Label lblTemplate 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Template"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3240
            TabIndex        =   497
            Top             =   4710
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label lblPWord 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Expiration Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   4800
            TabIndex        =   413
            Top             =   5550
            Width           =   1320
         End
         Begin VB.Label lblPWord 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Last Change"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   7920
            TabIndex        =   412
            Top             =   5550
            Width           =   1080
         End
         Begin VB.Label lblPWord 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Expiration Days"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   2520
            TabIndex        =   411
            Top             =   5550
            Width           =   1335
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Maintain"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   15
            TabIndex        =   328
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inquire"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   6780
            TabIndex        =   327
            Top             =   30
            Width           =   600
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Maintain"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   5520
            TabIndex        =   326
            Top             =   30
            Width           =   735
         End
         Begin VB.Label lblCNum 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Label20"
            DataField       =   "COMPNO"
            DataSource      =   "Data1"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   10080
            TabIndex        =   325
            Top             =   480
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Label1"
            DataField       =   "ID"
            DataSource      =   "Data1"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   9720
            TabIndex        =   324
            Top             =   3960
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label lblPWord 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   15
            TabIndex        =   323
            Top             =   5550
            Width           =   825
         End
         Begin VB.Label lblCountry 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Country"
            DataField       =   "COUNTRY"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3240
            TabIndex        =   322
            Top             =   4320
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lblCntryCode 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Country"
            Height          =   195
            Left            =   0
            TabIndex        =   321
            Top             =   4350
            Width           =   660
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Basic 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   18
            Left            =   0
            TabIndex        =   320
            Top             =   0
            Width           =   780
         End
      End
      Begin VB.PictureBox panDetails 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   7
         Left            =   3240
         ScaleHeight     =   2385
         ScaleWidth      =   3825
         TabIndex        =   455
         Top             =   2880
         Width           =   3855
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   18
            Left            =   30
            TabIndex        =   458
            Top             =   360
            Width           =   2640
            _Version        =   65536
            _ExtentX        =   4657
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Time Request"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   26
            Left            =   5040
            TabIndex        =   481
            Top             =   1800
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4741
            _ExtentY        =   402
            _StockProps     =   78
            Caption         =   "Approve Timesheet"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   20
            Left            =   30
            TabIndex        =   460
            Top             =   840
            Width           =   2640
            _Version        =   65536
            _ExtentX        =   4657
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Request Approval"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin VB.CommandButton cmdPageLeft 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   7
            Left            =   9120
            Picture         =   "fvsecure.frx":C892
            Style           =   1  'Graphical
            TabIndex        =   487
            Tag             =   "Grant All Basic"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdPageRight 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   7
            Left            =   9960
            Picture         =   "fvsecure.frx":CCD4
            Style           =   1  'Graphical
            TabIndex        =   486
            Tag             =   "Grant All Basic"
            Top             =   0
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.CommandButton cmdRemoveAllWeb 
            Appearance      =   0  'Flat
            Caption         =   "Remove All"
            Height          =   450
            Left            =   6720
            TabIndex        =   457
            Tag             =   "Grant All Utilities"
            Top             =   4800
            Width           =   2505
         End
         Begin VB.CommandButton cmdGrantAllWeb 
            Appearance      =   0  'Flat
            Caption         =   "Grant All ESS && Timesheet"
            Height          =   450
            Left            =   6720
            TabIndex        =   456
            Tag             =   "Grant All Utilities"
            Top             =   4320
            Width           =   2505
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   6
            Left            =   5040
            TabIndex        =   475
            Top             =   360
            Width           =   2640
            _Version        =   65536
            _ExtentX        =   4657
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Timesheet Template Setup"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   7
            Left            =   5040
            TabIndex        =   476
            Top             =   600
            Width           =   3165
            _Version        =   65536
            _ExtentX        =   5583
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Timesheet User Template Setup"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   8
            Left            =   5040
            TabIndex        =   477
            Top             =   840
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4736
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Delete Timesheets"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   14
            Left            =   5040
            TabIndex        =   479
            Top             =   1320
            Width           =   4605
            _Version        =   65536
            _ExtentX        =   8123
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Override Employee # Based Security in Timesheet"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   15
            Left            =   5040
            TabIndex        =   478
            Top             =   1080
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4736
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Delete Approved Timesheet"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   19
            Left            =   30
            TabIndex        =   459
            Top             =   600
            Width           =   2640
            _Version        =   65536
            _ExtentX        =   4657
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Vacation Request"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   21
            Left            =   30
            TabIndex        =   461
            Top             =   1080
            Width           =   2640
            _Version        =   65536
            _ExtentX        =   4657
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Request Approval Report"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   22
            Left            =   30
            TabIndex        =   462
            Top             =   1320
            Width           =   2640
            _Version        =   65536
            _ExtentX        =   4657
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Print Archived Report"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   23
            Left            =   30
            TabIndex        =   463
            Top             =   1560
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Archive Vacation/Time Request"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   24
            Left            =   30
            TabIndex        =   464
            Top             =   1800
            Width           =   2640
            _Version        =   65536
            _ExtentX        =   4657
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Mass Delete Request"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   17
            Left            =   5040
            TabIndex        =   482
            Top             =   2040
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Application Settings"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   25
            Left            =   5040
            TabIndex        =   480
            Top             =   1560
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4741
            _ExtentY        =   402
            _StockProps     =   78
            Caption         =   "Archive Timesheets"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   27
            Left            =   30
            TabIndex        =   465
            Top             =   2040
            Width           =   2640
            _Version        =   65536
            _ExtentX        =   4657
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Show All Requests"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   28
            Left            =   5040
            TabIndex        =   483
            Top             =   2280
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Punch In/Out"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   29
            Left            =   30
            TabIndex        =   469
            Top             =   3000
            Width           =   2985
            _Version        =   65536
            _ExtentX        =   5265
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Enable Supervisor Drop-down"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   30
            Left            =   30
            TabIndex        =   466
            Top             =   2280
            Width           =   2985
            _Version        =   65536
            _ExtentX        =   5265
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Delete Approved Time Requests"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   31
            Left            =   30
            TabIndex        =   467
            Top             =   2520
            Width           =   3585
            _Version        =   65536
            _ExtentX        =   6324
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Delete Approved Vacation Requests"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   32
            Left            =   30
            TabIndex        =   468
            Top             =   2760
            Width           =   3705
            _Version        =   65536
            _ExtentX        =   6535
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Enable Request Cancel Button"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   33
            Left            =   30
            TabIndex        =   470
            Top             =   3240
            Width           =   4065
            _Version        =   65536
            _ExtentX        =   7170
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Show Employee's Reporting Authority only"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   34
            Left            =   30
            TabIndex        =   471
            Top             =   3480
            Width           =   4065
            _Version        =   65536
            _ExtentX        =   7170
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Delete Future Dated Vacation Requests"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   35
            Left            =   30
            TabIndex        =   472
            Top             =   3720
            Width           =   4065
            _Version        =   65536
            _ExtentX        =   7170
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Delete Future Dated Time Requests"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   36
            Left            =   30
            TabIndex        =   473
            Top             =   3960
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Calendar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   37
            Left            =   30
            TabIndex        =   474
            Top             =   4200
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Dashboards"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   38
            Left            =   30
            TabIndex        =   519
            Top             =   4440
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Quick Info"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   39
            Left            =   30
            TabIndex        =   524
            Top             =   4680
            Width           =   2640
            _Version        =   65536
            _ExtentX        =   4657
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Maintain Demographic Data"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   40
            Left            =   30
            TabIndex        =   527
            Top             =   4920
            Width           =   3600
            _Version        =   65536
            _ExtentX        =   6350
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Show All Approved/Rejected Requests"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   41
            Left            =   5040
            TabIndex        =   528
            Top             =   2520
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Timesheet Submission"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   42
            Left            =   30
            TabIndex        =   531
            Top             =   5160
            Width           =   5000
            _Version        =   65536
            _ExtentX        =   8819
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Use Department Security on Request Authorization"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   43
            Left            =   2880
            TabIndex        =   535
            Top             =   840
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "View Own"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   44
            Left            =   30
            TabIndex        =   538
            Top             =   5400
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8819
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "My Co-Workers"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   45
            Left            =   5040
            TabIndex        =   542
            Top             =   2760
            Width           =   2985
            _Version        =   65536
            _ExtentX        =   5265
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Enable Supervisor Drop-down"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   46
            Left            =   5040
            TabIndex        =   543
            Top             =   3000
            Width           =   4065
            _Version        =   65536
            _ExtentX        =   7170
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Show Employee's Reporting Authority only"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   47
            Left            =   30
            TabIndex        =   555
            Top             =   5640
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8819
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Check ESS Setup"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Timesheet Web Module"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   25
            Left            =   5040
            TabIndex        =   485
            Top             =   0
            Width           =   2475
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ESS Web Module"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   24
            Left            =   0
            TabIndex        =   484
            Top             =   0
            Width           =   1830
         End
      End
      Begin VB.PictureBox panDetails 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   2
         Left            =   4560
         ScaleHeight     =   2385
         ScaleWidth      =   2025
         TabIndex        =   174
         Top             =   120
         Width           =   2055
         Begin VB.CommandButton cmdGrantAll 
            Appearance      =   0  'Flat
            Caption         =   "Grant All Utilities"
            Height          =   450
            Left            =   7560
            TabIndex        =   179
            Tag             =   "Grant All Utilities"
            Top             =   4440
            Width           =   1785
         End
         Begin VB.CommandButton cmdGrantInquire 
            Appearance      =   0  'Flat
            Caption         =   "Grant All Inquire"
            Height          =   450
            Left            =   7560
            TabIndex        =   178
            Tag             =   "Grant All Utilities"
            Top             =   4920
            Width           =   1785
         End
         Begin VB.CommandButton cmdRemove 
            Appearance      =   0  'Flat
            Caption         =   "Remove All"
            Height          =   450
            Left            =   7560
            TabIndex        =   177
            Tag             =   "Grant All Utilities"
            Top             =   5400
            Width           =   1785
         End
         Begin VB.CommandButton cmdPageRight 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   9960
            Picture         =   "fvsecure.frx":D116
            Style           =   1  'Graphical
            TabIndex        =   176
            Tag             =   "Grant All Basic"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdPageLeft 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   9120
            Picture         =   "fvsecure.frx":D558
            Style           =   1  'Graphical
            TabIndex        =   175
            Tag             =   "Grant All Basic"
            Top             =   0
            Width           =   720
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   255
            Index           =   0
            Left            =   3495
            TabIndex        =   180
            Top             =   2940
            Width           =   2160
            _Version        =   65536
            _ExtentX        =   3810
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Province"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   9
            Left            =   7470
            TabIndex        =   205
            Top             =   2280
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Table Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   8
            Left            =   6645
            TabIndex        =   204
            Top             =   2280
            Width           =   435
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   " "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   7
            Left            =   900
            TabIndex        =   181
            Top             =   1680
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "General Ledger Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   7
            Left            =   30
            TabIndex        =   182
            Top             =   1680
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   6
            Left            =   30
            TabIndex        =   183
            Top             =   1050
            Width           =   435
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   5
            Left            =   30
            TabIndex        =   184
            Top             =   1470
            Width           =   435
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   4
            Left            =   30
            TabIndex        =   185
            Top             =   1260
            Width           =   435
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   3
            Left            =   30
            TabIndex        =   186
            Top             =   840
            Width           =   435
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   1
            Left            =   30
            TabIndex        =   187
            Top             =   630
            Width           =   435
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   0
            Left            =   30
            TabIndex        =   188
            Top             =   420
            Width           =   435
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   0
            Left            =   900
            TabIndex        =   189
            Top             =   420
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Company Details"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   3
            Left            =   900
            TabIndex        =   190
            Top             =   840
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Department"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   4
            Left            =   900
            TabIndex        =   191
            Top             =   1260
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Audit Table"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   3
            Left            =   3500
            TabIndex        =   192
            Top             =   1050
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Payroll 'Matrix'"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   1
            Left            =   3495
            TabIndex        =   193
            Top             =   1500
            Visible         =   0   'False
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Entitlements 'Matrix'"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   5
            Left            =   900
            TabIndex        =   194
            Top             =   1470
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Employment Equity "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   1
            Left            =   900
            TabIndex        =   195
            Top             =   630
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Security Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   0
            Left            =   6645
            TabIndex        =   196
            Top             =   600
            Width           =   435
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   " "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   1
            Left            =   7470
            TabIndex        =   197
            Top             =   600
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Attendance Data"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   2
            Left            =   6645
            TabIndex        =   198
            Top             =   810
            Width           =   435
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   " "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   3
            Left            =   7470
            TabIndex        =   199
            Top             =   810
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Benefits Data "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   4
            Left            =   6645
            TabIndex        =   200
            Top             =   1230
            Width           =   435
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   " "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   5
            Left            =   7470
            TabIndex        =   201
            Top             =   1230
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Employee Data "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   6
            Left            =   6645
            TabIndex        =   202
            Top             =   2070
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   " "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   7
            Left            =   7470
            TabIndex        =   203
            Top             =   2070
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Salary Data"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   2
            Left            =   3500
            TabIndex        =   212
            Top             =   420
            Width           =   2325
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Compress/Fix Database"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   6
            Left            =   900
            TabIndex        =   213
            Top             =   1050
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Division"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   10
            Left            =   900
            TabIndex        =   214
            Top             =   2130
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Holiday Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   10
            Left            =   30
            TabIndex        =   215
            Top             =   2130
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   4
            Left            =   3495
            TabIndex        =   216
            Top             =   2730
            Visible         =   0   'False
            Width           =   2325
            _Version        =   65536
            _ExtentX        =   4101
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Door Name Setup"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   9
            Left            =   900
            TabIndex        =   217
            Top             =   1890
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Custom Report Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   9
            Left            =   30
            TabIndex        =   218
            Top             =   1890
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   5
            Left            =   3495
            TabIndex        =   219
            Top             =   2520
            Visible         =   0   'False
            Width           =   3045
            _Version        =   65536
            _ExtentX        =   5371
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Summarize Attendance Records    "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   11
            Left            =   900
            TabIndex        =   220
            Top             =   2340
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "New Hire Procedure"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   11
            Left            =   30
            TabIndex        =   221
            Top             =   2340
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   12
            Left            =   900
            TabIndex        =   222
            Top             =   2550
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Label Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   12
            Left            =   30
            TabIndex        =   223
            Top             =   2550
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   8
            Left            =   900
            TabIndex        =   224
            Top             =   5655
            Visible         =   0   'False
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Door Access"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   8
            Left            =   30
            TabIndex        =   225
            Top             =   5655
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   13
            Left            =   30
            TabIndex        =   226
            Top             =   2760
            Width           =   435
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   13
            Left            =   900
            TabIndex        =   227
            Top             =   2760
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Salary Distribution"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   14
            Left            =   30
            TabIndex        =   228
            Top             =   5430
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   14
            Left            =   900
            TabIndex        =   229
            Top             =   5430
            Visible         =   0   'False
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Payroll Category"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   17
            Left            =   900
            TabIndex        =   230
            Top             =   3600
            Width           =   2325
            _Version        =   65536
            _ExtentX        =   4101
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Machine #"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   17
            Left            =   30
            TabIndex        =   231
            Top             =   3600
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   15
            Left            =   900
            TabIndex        =   232
            Top             =   5880
            Visible         =   0   'False
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Charge Code"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   15
            Left            =   30
            TabIndex        =   233
            Top             =   5880
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   16
            Left            =   900
            TabIndex        =   234
            Top             =   3390
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Account Code"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   16
            Left            =   30
            TabIndex        =   235
            Top             =   3390
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   19
            Left            =   30
            TabIndex        =   236
            Top             =   2970
            Width           =   435
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   18
            Left            =   30
            TabIndex        =   237
            Top             =   3180
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   19
            Left            =   900
            TabIndex        =   238
            Top             =   2970
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Pay Period Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   18
            Left            =   900
            TabIndex        =   239
            Top             =   3180
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Email Setup"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   20
            Left            =   900
            TabIndex        =   240
            Top             =   6090
            Visible         =   0   'False
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Product Line/Operation"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   20
            Left            =   30
            TabIndex        =   241
            Top             =   6090
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   9
            Left            =   3500
            TabIndex        =   242
            Top             =   630
            Width           =   2325
            _Version        =   65536
            _ExtentX        =   4101
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Company Preference"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   10
            Left            =   3500
            TabIndex        =   243
            Top             =   840
            Width           =   2445
            _Version        =   65536
            _ExtentX        =   4313
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Employee Flags Setup"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   11
            Left            =   3495
            TabIndex        =   244
            Top             =   1280
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Multiple Data Source Setup"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   12
            Left            =   3495
            TabIndex        =   406
            Top             =   3165
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Help Description Setup"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   38
            Left            =   900
            TabIndex        =   407
            Top             =   3810
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Course Code Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   38
            Left            =   30
            TabIndex        =   410
            Top             =   3810
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   13
            Left            =   3495
            TabIndex        =   414
            Top             =   3375
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Benefit Group Setup"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   41
            Left            =   30
            TabIndex        =   417
            Top             =   6300
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   762
            _ExtentY        =   402
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   41
            Left            =   900
            TabIndex        =   418
            Top             =   6300
            Visible         =   0   'False
            Width           =   3405
            _Version        =   65536
            _ExtentX        =   6006
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Attendance Group Code Matrix"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUSecurity 
            Height          =   225
            Index           =   16
            Left            =   3495
            TabIndex        =   432
            Top             =   3600
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Change Your Password"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   47
            Left            =   900
            TabIndex        =   444
            Top             =   4260
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Budgeted Manpower"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   47
            Left            =   30
            TabIndex        =   445
            Top             =   4260
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   59
            Left            =   900
            TabIndex        =   409
            Top             =   4030
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Attendance Code Matrix"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   59
            Left            =   30
            TabIndex        =   408
            Top             =   4030
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   11
            Left            =   7470
            TabIndex        =   207
            Top             =   2490
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "YTD Data"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   10
            Left            =   6645
            TabIndex        =   206
            Top             =   2490
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   " "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   13
            Left            =   7470
            TabIndex        =   209
            Top             =   1650
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Payroll Transaction"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   12
            Left            =   6645
            TabIndex        =   208
            Top             =   1650
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   " "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   15
            Left            =   7470
            TabIndex        =   211
            Top             =   1020
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Continuing Education Data"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   14
            Left            =   6645
            TabIndex        =   210
            Top             =   1020
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   " "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   17
            Left            =   7470
            TabIndex        =   488
            Top             =   1860
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Performance Review"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   16
            Left            =   6645
            TabIndex        =   489
            Top             =   1860
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   " "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   19
            Left            =   7470
            TabIndex        =   494
            Top             =   1440
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Employment Equity"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkIESecurity 
            Height          =   225
            Index           =   18
            Left            =   6645
            TabIndex        =   495
            Top             =   1440
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   " "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   64
            Left            =   900
            TabIndex        =   508
            Top             =   4500
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Work Schedule Rule"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   64
            Left            =   30
            TabIndex        =   509
            Top             =   4500
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   65
            Left            =   900
            TabIndex        =   511
            Top             =   4720
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Dashboard Setup"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   65
            Left            =   30
            TabIndex        =   512
            Top             =   4720
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   66
            Left            =   30
            TabIndex        =   513
            Top             =   6525
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   66
            Left            =   900
            TabIndex        =   514
            Top             =   6525
            Visible         =   0   'False
            Width           =   2325
            _Version        =   65536
            _ExtentX        =   4101
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Counseling Audit Table"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   67
            Left            =   30
            TabIndex        =   522
            Top             =   6750
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   67
            Left            =   900
            TabIndex        =   523
            Top             =   6750
            Visible         =   0   'False
            Width           =   2325
            _Version        =   65536
            _ExtentX        =   4101
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "On Call Hours"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   69
            Left            =   900
            TabIndex        =   529
            Top             =   4950
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Follow Up Code Email Matrix"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   69
            Left            =   30
            TabIndex        =   530
            Top             =   4955
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   70
            Left            =   30
            TabIndex        =   536
            Top             =   6960
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   70
            Left            =   900
            TabIndex        =   537
            Top             =   6960
            Visible         =   0   'False
            Width           =   3525
            _Version        =   65536
            _ExtentX        =   6218
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Department/GL Matrix"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   71
            Left            =   900
            TabIndex        =   540
            Top             =   5190
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "OHRS Department"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   71
            Left            =   30
            TabIndex        =   541
            Top             =   5190
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   72
            Left            =   7470
            TabIndex        =   548
            Top             =   3240
            Visible         =   0   'False
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Database Setup"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   72
            Left            =   6645
            TabIndex        =   549
            Top             =   3240
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   73
            Left            =   7470
            TabIndex        =   550
            Top             =   3480
            Visible         =   0   'False
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Integration Setup"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   73
            Left            =   6645
            TabIndex        =   551
            Top             =   3480
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Maintain"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   30
            Left            =   6615
            TabIndex        =   553
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inquire"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   29
            Left            =   7470
            TabIndex        =   552
            Top             =   3000
            Width           =   600
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inquire"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   11
            Left            =   900
            TabIndex        =   249
            Top             =   210
            Width           =   600
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Maintain"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   10
            Left            =   20
            TabIndex        =   248
            Top             =   210
            Width           =   735
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Export"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   7470
            TabIndex        =   247
            Top             =   360
            Width           =   555
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Import"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   6
            Left            =   6615
            TabIndex        =   246
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Utilities"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   19
            Left            =   0
            TabIndex        =   245
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.PictureBox panDetails 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2175
         Index           =   3
         Left            =   6720
         ScaleHeight     =   2145
         ScaleWidth      =   2025
         TabIndex        =   128
         Top             =   120
         Width           =   2055
         Begin VB.CommandButton cmdGrantAllR 
            Appearance      =   0  'Flat
            Caption         =   "&Grant All Reports"
            Height          =   450
            Left            =   10200
            TabIndex        =   132
            Tag             =   "Grant All Reports"
            Top             =   360
            Width           =   1900
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            Caption         =   "&Remove All Reports"
            Height          =   450
            Left            =   10200
            TabIndex        =   131
            Tag             =   "Grant All Reports"
            Top             =   840
            Width           =   1900
         End
         Begin VB.CommandButton cmdPageLeft 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   9120
            Picture         =   "fvsecure.frx":D99A
            Style           =   1  'Graphical
            TabIndex        =   130
            Tag             =   "Grant All Basic"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdPageRight 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   9960
            Picture         =   "fvsecure.frx":DDDC
            Style           =   1  'Graphical
            TabIndex        =   129
            Tag             =   "Grant All Basic"
            Top             =   0
            Width           =   705
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   255
            Index           =   35
            Left            =   3000
            TabIndex        =   149
            Top             =   360
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Emergency Leave"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   7
            Left            =   3000
            TabIndex        =   133
            Top             =   2040
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Entitlements"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   195
            Index           =   32
            Left            =   3000
            TabIndex        =   134
            Top             =   1335
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Employee Turnover"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   29
            Left            =   15
            TabIndex        =   135
            Top             =   3480
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Dependents"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   30
            Left            =   3000
            TabIndex        =   136
            Top             =   1560
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Employee Skills"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   28
            Left            =   3000
            TabIndex        =   137
            Top             =   1800
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Employment Equity"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   27
            Left            =   6720
            TabIndex        =   138
            Top             =   600
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Other Earnings"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   26
            Left            =   3000
            TabIndex        =   139
            Top             =   4200
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Hourly Entitlements"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   25
            Left            =   15
            TabIndex        =   140
            Top             =   3960
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Dollar Entitlements"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   24
            Left            =   3000
            TabIndex        =   148
            Top             =   3720
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Health and Safety "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   23
            Left            =   6720
            TabIndex        =   153
            Top             =   3720
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Table Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   22
            Left            =   15
            TabIndex        =   141
            Top             =   3240
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Continuing Education Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   21
            Left            =   6720
            TabIndex        =   154
            Top             =   2280
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Salary History Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   3
            Left            =   15
            TabIndex        =   144
            Top             =   4920
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Emergency Contacts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   4
            Left            =   15
            TabIndex        =   147
            Top             =   5640
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Employee Labels"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   6
            Left            =   3000
            TabIndex        =   155
            Top             =   840
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Employee Profile"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   8
            Left            =   3000
            TabIndex        =   156
            Top             =   2520
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Follow-up"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   9
            Left            =   3000
            TabIndex        =   157
            Top             =   3960
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Home Address/Phone List"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   10
            Left            =   6720
            TabIndex        =   158
            Top             =   2520
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Salary/Performance Review"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   11
            Left            =   6720
            TabIndex        =   159
            Top             =   2760
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Seniority"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   1
            Left            =   15
            TabIndex        =   160
            Top             =   2040
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Compensatory Time"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   16
            Left            =   15
            TabIndex        =   161
            Top             =   3720
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Division Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   17
            Left            =   3000
            TabIndex        =   162
            Top             =   1080
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Terminations"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   18
            Left            =   3000
            TabIndex        =   163
            Top             =   2760
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Formal Education Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   19
            Left            =   6720
            TabIndex        =   164
            Top             =   2040
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Position/Skills/Evaluation"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   33
            Left            =   15
            TabIndex        =   165
            Top             =   7800
            Visible         =   0   'False
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Door Access"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   34
            Left            =   15
            TabIndex        =   166
            Top             =   3000
            Width           =   2025
            _Version        =   65536
            _ExtentX        =   3572
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Counselling"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   2
            Left            =   15
            TabIndex        =   167
            Top             =   2520
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Cost of Employment"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   5
            Left            =   3000
            TabIndex        =   150
            Top             =   600
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Employee/Position List"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   20
            Left            =   6720
            TabIndex        =   168
            Top             =   1560
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Security Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   12
            Left            =   6720
            TabIndex        =   169
            Top             =   3960
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Telephone Extension List"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   31
            Left            =   3000
            TabIndex        =   170
            Top             =   4920
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Languages"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   46
            Left            =   6720
            TabIndex        =   171
            Top             =   840
            Width           =   2445
            _Version        =   65536
            _ExtentX        =   4313
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Overtime Bank"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   47
            Left            =   6720
            TabIndex        =   172
            Top             =   1080
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Overtime Bank Lost Hours"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   48
            Left            =   3000
            TabIndex        =   398
            Top             =   2280
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "External Hire Rate"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   49
            Left            =   3000
            TabIndex        =   399
            Top             =   4440
            Width           =   3585
            _Version        =   65536
            _ExtentX        =   6324
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Internal Transfers to Total Hires Ratio"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   50
            Left            =   3000
            TabIndex        =   400
            Top             =   4680
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Key Workforce Demographic"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   51
            Left            =   6720
            TabIndex        =   401
            Top             =   360
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Manpower Plan"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   52
            Left            =   6720
            TabIndex        =   402
            Top             =   3240
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Staff/Management Ratios"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   53
            Left            =   6720
            TabIndex        =   403
            Top             =   4680
            Width           =   5265
            _Version        =   65536
            _ExtentX        =   9287
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Workers Compensation (WC) Lost Time Incident Rate"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   54
            Left            =   6720
            TabIndex        =   404
            Top             =   4920
            Width           =   4785
            _Version        =   65536
            _ExtentX        =   8440
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Workers Compensation (WC) Lost Work Hours Rate"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   55
            Left            =   6720
            TabIndex        =   405
            Top             =   1320
            Width           =   3705
            _Version        =   65536
            _ExtentX        =   6535
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Paid Sick Hours Per Eligible Employee"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   56
            Left            =   6720
            TabIndex        =   415
            Top             =   4440
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "User Defined Table"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   57
            Left            =   3000
            TabIndex        =   416
            Top             =   3000
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Future Entitlements"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   58
            Left            =   15
            TabIndex        =   145
            Top             =   5160
            Width           =   2385
            _Version        =   65536
            _ExtentX        =   4207
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Employee Flags"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   59
            Left            =   3000
            TabIndex        =   422
            Top             =   7560
            Visible         =   0   'False
            Width           =   3585
            _Version        =   65536
            _ExtentX        =   6324
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Temporary/Cross Training Assigments"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   61
            Left            =   15
            TabIndex        =   424
            Top             =   6840
            Visible         =   0   'False
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "I Want You to Know..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   62
            Left            =   15
            TabIndex        =   425
            Top             =   7080
            Visible         =   0   'False
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "IT Hire Form"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   63
            Left            =   15
            TabIndex        =   426
            Top             =   7320
            Visible         =   0   'False
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "IT Notice of Change Form"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   64
            Left            =   15
            TabIndex        =   427
            Top             =   7560
            Visible         =   0   'False
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Notice of Change Form"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   65
            Left            =   3000
            TabIndex        =   419
            Top             =   6840
            Visible         =   0   'False
            Width           =   3645
            _Version        =   65536
            _ExtentX        =   6429
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Performance Improvement Action Plan"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   66
            Left            =   3000
            TabIndex        =   420
            Top             =   7080
            Visible         =   0   'False
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Performance Review Form"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   69
            Left            =   6720
            TabIndex        =   428
            Top             =   6840
            Visible         =   0   'False
            Width           =   2325
            _Version        =   65536
            _ExtentX        =   4101
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Update Meeting Form"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   70
            Left            =   6720
            TabIndex        =   429
            Top             =   7080
            Visible         =   0   'False
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Warning Form"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   60
            Left            =   6720
            TabIndex        =   431
            Top             =   5400
            Width           =   3585
            _Version        =   65536
            _ExtentX        =   6324
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Required Courses History"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   67
            Left            =   3000
            TabIndex        =   423
            Top             =   7800
            Visible         =   0   'False
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Termination Form"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   68
            Left            =   3000
            TabIndex        =   421
            Top             =   7320
            Visible         =   0   'False
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Separation Form"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   71
            Left            =   15
            TabIndex        =   143
            Top             =   4680
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Email Address"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   72
            Left            =   3000
            TabIndex        =   151
            Top             =   3240
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Gap Analysis"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   73
            Left            =   3000
            TabIndex        =   433
            Top             =   5160
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Leave of Absence"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   74
            Left            =   3000
            TabIndex        =   434
            Top             =   5640
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Plan of Establishment"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   75
            Left            =   6720
            TabIndex        =   435
            Top             =   3000
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "S.I.N./S.S.N."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   76
            Left            =   6720
            TabIndex        =   436
            Top             =   3480
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Succession"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   78
            Left            =   15
            TabIndex        =   437
            Top             =   2280
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Comments"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   0
            Left            =   15
            TabIndex        =   438
            Top             =   1800
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Birthday/Age"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   15
            Left            =   15
            TabIndex        =   439
            Top             =   1560
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Benefits Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   13
            Left            =   15
            TabIndex        =   440
            Top             =   360
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Associations File"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   14
            Left            =   15
            TabIndex        =   441
            Top             =   840
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Attendance Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   77
            Left            =   15
            TabIndex        =   442
            Top             =   600
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Attendance History"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   79
            Left            =   15
            TabIndex        =   146
            Top             =   5400
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Employee History"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   80
            Left            =   6720
            TabIndex        =   443
            Top             =   1800
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Payroll Transactions"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   81
            Left            =   15
            TabIndex        =   142
            Top             =   4440
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "EEO Reports"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   82
            Left            =   6720
            TabIndex        =   451
            Top             =   5160
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Work Schedule"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   83
            Left            =   15
            TabIndex        =   452
            Top             =   1080
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Attendance Bonus Points"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   84
            Left            =   15
            TabIndex        =   453
            Top             =   1320
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Attendance Calendar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   85
            Left            =   15
            TabIndex        =   454
            Top             =   2760
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Costed Attendance"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   86
            Left            =   3000
            TabIndex        =   152
            Top             =   3480
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "G/L Distribution"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   88
            Left            =   15
            TabIndex        =   490
            Top             =   6480
            Width           =   2325
            _Version        =   65536
            _ExtentX        =   4101
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Applicant Education"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   87
            Left            =   15
            TabIndex        =   491
            Top             =   6240
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Applicant Profile"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   89
            Left            =   6720
            TabIndex        =   503
            Top             =   4200
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Training Plan"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   90
            Left            =   6720
            TabIndex        =   506
            Top             =   7560
            Width           =   3885
            _Version        =   65536
            _ExtentX        =   6853
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Attendance/Work Schedule Discrepancy"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   91
            Left            =   6720
            TabIndex        =   507
            Top             =   7800
            Width           =   3885
            _Version        =   65536
            _ExtentX        =   6853
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Wellington Terrace Attendance"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   92
            Left            =   6720
            TabIndex        =   510
            Top             =   5640
            Width           =   3585
            _Version        =   65536
            _ExtentX        =   6324
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "ESS Requests - Transaction Audit"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   93
            Left            =   15
            TabIndex        =   520
            Top             =   5880
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Employee Dates"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   94
            Left            =   3000
            TabIndex        =   521
            Top             =   5400
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Length of Service"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   95
            Left            =   3000
            TabIndex        =   532
            Top             =   5880
            Width           =   2505
            _Version        =   65536
            _ExtentX        =   4419
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Attendance Sign In Form"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   96
            Left            =   6720
            TabIndex        =   533
            Top             =   5880
            Visible         =   0   'False
            Width           =   2385
            _Version        =   65536
            _ExtentX        =   4207
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "ATT Discipline Form"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   97
            Left            =   6720
            TabIndex        =   534
            Top             =   6120
            Visible         =   0   'False
            Width           =   2385
            _Version        =   65536
            _ExtentX        =   4207
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "COC Discipline Form"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   98
            Left            =   3000
            TabIndex        =   539
            Top             =   6120
            Visible         =   0   'False
            Width           =   3645
            _Version        =   65536
            _ExtentX        =   6429
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Flex Time"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   99
            Left            =   15
            TabIndex        =   544
            Top             =   4200
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Document Type"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   225
            Index           =   100
            Left            =   6720
            TabIndex        =   430
            Top             =   7320
            Visible         =   0   'False
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Staff Profile"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Reports "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   20
            Left            =   0
            TabIndex        =   173
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.PictureBox panDetails 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2655
         Index           =   6
         Left            =   7560
         ScaleHeight     =   2625
         ScaleWidth      =   4185
         TabIndex        =   360
         Top             =   2760
         Width           =   4215
         Begin VB.CommandButton cmdPageLeft 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   6
            Left            =   9120
            Picture         =   "fvsecure.frx":E21E
            Style           =   1  'Graphical
            TabIndex        =   364
            Tag             =   "Grant All Basic"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdRemoveCourseAdmin 
            Appearance      =   0  'Flat
            Caption         =   "Remove All"
            Height          =   330
            Left            =   6960
            TabIndex        =   363
            Tag             =   "Grant All Basic"
            Top             =   1440
            Width           =   2160
         End
         Begin VB.CommandButton cmdGrantInquCourseAdmin 
            Appearance      =   0  'Flat
            Caption         =   "Grant All Inquire"
            Height          =   330
            Left            =   6960
            TabIndex        =   362
            Tag             =   "Grant All Basic"
            Top             =   1080
            Width           =   2160
         End
         Begin VB.CommandButton cmdGrantCourseAdmin 
            Appearance      =   0  'Flat
            Caption         =   "&Grant All Course Admin"
            Height          =   330
            Left            =   6960
            TabIndex        =   361
            Tag             =   "Grant All Basic"
            Top             =   720
            Width           =   2160
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   27
            Left            =   870
            TabIndex        =   365
            Top             =   600
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Organizations/Contacts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   27
            Left            =   0
            TabIndex        =   366
            Top             =   600
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   28
            Left            =   870
            TabIndex        =   367
            Top             =   840
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Course Catalog"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   28
            Left            =   0
            TabIndex        =   368
            Top             =   840
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   29
            Left            =   870
            TabIndex        =   369
            Top             =   1080
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Training Location"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   29
            Left            =   0
            TabIndex        =   370
            Top             =   1080
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   30
            Left            =   870
            TabIndex        =   371
            Top             =   1320
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Scheduling"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   30
            Left            =   0
            TabIndex        =   372
            Top             =   1320
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   31
            Left            =   870
            TabIndex        =   373
            Top             =   1560
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Enrollment Request"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   31
            Left            =   0
            TabIndex        =   374
            Top             =   1560
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   32
            Left            =   870
            TabIndex        =   375
            Top             =   1800
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Enrollment Approval"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   32
            Left            =   0
            TabIndex        =   376
            Top             =   1800
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   33
            Left            =   870
            TabIndex        =   377
            Top             =   2040
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Enrollment"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   33
            Left            =   0
            TabIndex        =   378
            Top             =   2040
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   34
            Left            =   870
            TabIndex        =   379
            Top             =   2280
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Waiting List"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   34
            Left            =   0
            TabIndex        =   380
            Top             =   2280
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   255
            Index           =   36
            Left            =   3840
            TabIndex        =   381
            Top             =   600
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Calendar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   255
            Index           =   37
            Left            =   3840
            TabIndex        =   382
            Top             =   840
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Class List"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   255
            Index           =   38
            Left            =   3840
            TabIndex        =   383
            Top             =   1080
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Waiting List"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   255
            Index           =   39
            Left            =   3840
            TabIndex        =   384
            Top             =   1320
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Conflict"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   255
            Index           =   40
            Left            =   3840
            TabIndex        =   385
            Top             =   1560
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Course Catalog"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   255
            Index           =   41
            Left            =   3840
            TabIndex        =   386
            Top             =   1800
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Required Courses Per Position"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   255
            Index           =   42
            Left            =   3840
            TabIndex        =   387
            Top             =   2040
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Label"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   255
            Index           =   43
            Left            =   3840
            TabIndex        =   388
            Top             =   2280
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Prerequisite Exception"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   255
            Index           =   44
            Left            =   3840
            TabIndex        =   389
            Top             =   2520
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Required Course Not Completed or Enrolled"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkSSecurity 
            Height          =   255
            Index           =   45
            Left            =   3840
            TabIndex        =   390
            Top             =   2760
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Training Summary"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   2
            Left            =   0
            TabIndex        =   391
            Top             =   2520
            Width           =   435
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   2
            Left            =   870
            TabIndex        =   392
            Top             =   2520
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   3254
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Table Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Course Admin"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   22
            Left            =   0
            TabIndex        =   396
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Reports"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   17
            Left            =   3840
            TabIndex        =   395
            Top             =   360
            Width           =   675
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Maintain"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   14
            Left            =   0
            TabIndex        =   394
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inquire"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   12
            Left            =   1080
            TabIndex        =   393
            Top             =   360
            Width           =   600
         End
      End
      Begin VB.PictureBox panDetails 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2295
         Index           =   4
         Left            =   8880
         ScaleHeight     =   2265
         ScaleWidth      =   2385
         TabIndex        =   102
         Top             =   120
         Width           =   2415
         Begin VB.CommandButton cmdRemoveAllMC 
            Appearance      =   0  'Flat
            Caption         =   "Remove All Mass Changes"
            Height          =   450
            Left            =   6630
            TabIndex        =   106
            Tag             =   "Grant All Utilities"
            Top             =   1440
            Width           =   2445
         End
         Begin VB.CommandButton cmdGrantAllMC 
            Appearance      =   0  'Flat
            Caption         =   "Grant All Mass Changes"
            Height          =   450
            Left            =   6630
            TabIndex        =   105
            Tag             =   "Grant All Utilities"
            Top             =   960
            Width           =   2445
         End
         Begin VB.CommandButton cmdPageLeft 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   9120
            Picture         =   "fvsecure.frx":E660
            Style           =   1  'Graphical
            TabIndex        =   104
            Tag             =   "Grant All Basic"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdPageRight 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   9960
            Picture         =   "fvsecure.frx":EAA2
            Style           =   1  'Graphical
            TabIndex        =   103
            Tag             =   "Grant All Basic"
            Top             =   0
            Width           =   705
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   107
            Top             =   300
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Attendance History"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   108
            Top             =   540
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Attendance Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   2
            Left            =   90
            TabIndex        =   109
            Top             =   780
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Benefit Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   3
            Left            =   90
            TabIndex        =   110
            Top             =   1020
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Codes"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   4
            Left            =   90
            TabIndex        =   111
            Top             =   1260
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Continuing Education"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   5
            Left            =   90
            TabIndex        =   112
            Top             =   1500
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Dollar Entitlements"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   6
            Left            =   90
            TabIndex        =   114
            Top             =   1980
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Entitlements"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   7
            Left            =   90
            TabIndex        =   116
            Top             =   2460
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Follow - Ups"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   9
            Left            =   3090
            TabIndex        =   121
            Top             =   300
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Other Earings"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   10
            Left            =   3090
            TabIndex        =   123
            Top             =   780
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Position Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   11
            Left            =   3090
            TabIndex        =   124
            Top             =   1020
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Salary Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   8
            Left            =   90
            TabIndex        =   117
            Top             =   2700
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Hourly Entitlements"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   12
            Left            =   90
            TabIndex        =   115
            Top             =   2220
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Employee Number"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   13
            Left            =   3090
            TabIndex        =   122
            Top             =   540
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Overtime Master"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   14
            Left            =   90
            TabIndex        =   113
            Top             =   1740
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Emergency Leave"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   15
            Left            =   90
            TabIndex        =   120
            Top             =   3420
            Width           =   2715
            _Version        =   65536
            _ExtentX        =   4789
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Maintain Photo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   16
            Left            =   3090
            TabIndex        =   125
            Top             =   1260
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Work Schedule"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   17
            Left            =   3090
            TabIndex        =   126
            Top             =   1800
            Visible         =   0   'False
            Width           =   2835
            _Version        =   65536
            _ExtentX        =   5001
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Anniversary Month Year End "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   18
            Left            =   90
            TabIndex        =   118
            Top             =   2940
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Import Email Address"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkMCSecurity 
            Height          =   225
            Index           =   19
            Left            =   90
            TabIndex        =   119
            Top             =   3180
            Width           =   2715
            _Version        =   65536
            _ExtentX        =   4789
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Import Attachment"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Mass Changes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   13
            Left            =   0
            TabIndex        =   127
            Top             =   0
            Width           =   1545
         End
      End
      Begin VB.PictureBox panDetails 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2295
         Index           =   1
         Left            =   2280
         ScaleHeight     =   2265
         ScaleWidth      =   2145
         TabIndex        =   250
         Top             =   120
         Width           =   2175
         Begin VB.CommandButton cmdPageLeft 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   9120
            Picture         =   "fvsecure.frx":EEE4
            Style           =   1  'Graphical
            TabIndex        =   255
            Tag             =   "Grant All Basic"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdRemoveAll2 
            Appearance      =   0  'Flat
            Caption         =   "Remove All"
            Height          =   330
            Left            =   7200
            TabIndex        =   254
            Tag             =   "Grant All Basic"
            Top             =   4920
            Width           =   1800
         End
         Begin VB.CommandButton cmdGrantInqu2 
            Appearance      =   0  'Flat
            Caption         =   "Grant All Inquire"
            Height          =   330
            Left            =   7200
            TabIndex        =   253
            Tag             =   "Grant All Basic"
            Top             =   4560
            Width           =   1800
         End
         Begin VB.CommandButton cmdGrantInquire2 
            Appearance      =   0  'Flat
            Caption         =   "&Grant All Basic 2"
            Height          =   330
            Left            =   7200
            TabIndex        =   252
            Tag             =   "Grant All Basic"
            Top             =   4200
            Width           =   1800
         End
         Begin VB.CommandButton cmdPageRight 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   9960
            Picture         =   "fvsecure.frx":F326
            Style           =   1  'Graphical
            TabIndex        =   251
            Tag             =   "Grant All Basic"
            Top             =   0
            Width           =   705
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   22
            Left            =   870
            TabIndex        =   262
            Top             =   1680
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   3254
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Employee Flags"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   22
            Left            =   0
            TabIndex        =   263
            Top             =   1680
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   23
            Left            =   870
            TabIndex        =   264
            Top             =   1920
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   3254
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Employee History"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   23
            Left            =   0
            TabIndex        =   265
            Top             =   1920
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   24
            Left            =   870
            TabIndex        =   268
            Top             =   2400
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "G/L Distribution"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   24
            Left            =   0
            TabIndex        =   269
            Top             =   2400
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   25
            Left            =   870
            TabIndex        =   280
            Top             =   3840
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Languages"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   25
            Left            =   0
            TabIndex        =   281
            Top             =   3840
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   26
            Left            =   6390
            TabIndex        =   290
            Top             =   720
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Succession Plan"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   26
            Left            =   5520
            TabIndex        =   291
            Top             =   720
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   36
            Left            =   870
            TabIndex        =   258
            Top             =   1200
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Employee ADP Data"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   36
            Left            =   0
            TabIndex        =   259
            Top             =   1200
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   37
            Left            =   0
            TabIndex        =   271
            Top             =   2640
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   37
            Left            =   870
            TabIndex        =   270
            Top             =   2640
            Width           =   4005
            _Version        =   65536
            _ExtentX        =   7064
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Health && Safety Claim/Medical Information"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   39
            Left            =   5520
            TabIndex        =   296
            Top             =   1200
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   39
            Left            =   6390
            TabIndex        =   295
            Top             =   1200
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "User Defined Table"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   40
            Left            =   0
            TabIndex        =   283
            Top             =   4080
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   40
            Left            =   870
            TabIndex        =   282
            Top             =   4080
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Payroll Transactions"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   42
            Left            =   5520
            TabIndex        =   300
            Top             =   1680
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   42
            Left            =   6390
            TabIndex        =   299
            Top             =   1680
            Visible         =   0   'False
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Job Files Attachment"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   43
            Left            =   5520
            TabIndex        =   302
            Top             =   1920
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   43
            Left            =   6390
            TabIndex        =   301
            Top             =   1920
            Visible         =   0   'False
            Width           =   3165
            _Version        =   65536
            _ExtentX        =   5583
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Temporary/Cross Training Position"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   44
            Left            =   5520
            TabIndex        =   294
            Top             =   960
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   44
            Left            =   6390
            TabIndex        =   293
            Top             =   960
            Width           =   3165
            _Version        =   65536
            _ExtentX        =   5583
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Training List"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   45
            Left            =   5520
            TabIndex        =   304
            Top             =   2160
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   45
            Left            =   6390
            TabIndex        =   303
            Top             =   2160
            Visible         =   0   'False
            Width           =   3165
            _Version        =   65536
            _ExtentX        =   5583
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Retirement Process"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   46
            Left            =   5520
            TabIndex        =   306
            Top             =   2400
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   46
            Left            =   6390
            TabIndex        =   305
            Top             =   2400
            Visible         =   0   'False
            Width           =   3165
            _Version        =   65536
            _ExtentX        =   5583
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Death Process"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   35
            Left            =   870
            TabIndex        =   260
            Top             =   1440
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Emergency Contacts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   35
            Left            =   0
            TabIndex        =   261
            Top             =   1440
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   48
            Left            =   870
            TabIndex        =   256
            Top             =   960
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Budgeted Position"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   48
            Left            =   0
            TabIndex        =   257
            Top             =   960
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   49
            Left            =   870
            TabIndex        =   284
            Top             =   4320
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4736
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Position Application Process"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   49
            Left            =   0
            TabIndex        =   285
            Top             =   4320
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   50
            Left            =   870
            TabIndex        =   286
            Top             =   4560
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Position Required Courses"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   50
            Left            =   0
            TabIndex        =   287
            Top             =   4560
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   51
            Left            =   6390
            TabIndex        =   288
            Top             =   480
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Rehire"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   51
            Left            =   5520
            TabIndex        =   289
            Top             =   480
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   52
            Left            =   870
            TabIndex        =   266
            Top             =   2160
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Enter a Leave"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   52
            Left            =   0
            TabIndex        =   267
            Top             =   2160
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   53
            Left            =   870
            TabIndex        =   272
            Top             =   2880
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Health && Safety Contacts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   53
            Left            =   0
            TabIndex        =   273
            Top             =   2880
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   54
            Left            =   870
            TabIndex        =   276
            Top             =   3360
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Health && Safety Cost"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   54
            Left            =   0
            TabIndex        =   277
            Top             =   3360
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   55
            Left            =   870
            TabIndex        =   274
            Top             =   3120
            Width           =   3165
            _Version        =   65536
            _ExtentX        =   5583
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Health && Safety Corrective Action"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   55
            Left            =   0
            TabIndex        =   275
            Top             =   3120
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   56
            Left            =   870
            TabIndex        =   278
            Top             =   3600
            Width           =   3045
            _Version        =   65536
            _ExtentX        =   5371
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Health && Safety Root Cause"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   56
            Left            =   0
            TabIndex        =   279
            Top             =   3600
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   58
            Left            =   0
            TabIndex        =   446
            Top             =   720
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   58
            Left            =   870
            TabIndex        =   447
            Top             =   720
            Width           =   3645
            _Version        =   65536
            _ExtentX        =   6429
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "EEO - Purge Applicants Records"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   57
            Left            =   0
            TabIndex        =   448
            Top             =   480
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   57
            Left            =   870
            TabIndex        =   449
            Top             =   480
            Width           =   4365
            _Version        =   65536
            _ExtentX        =   7699
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "EEO - Data Maintenance"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   60
            Left            =   6390
            TabIndex        =   297
            Top             =   1440
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Work Schedule"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   60
            Left            =   5520
            TabIndex        =   298
            Top             =   1440
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   195
            Index           =   61
            Left            =   6390
            TabIndex        =   307
            Top             =   3375
            Visible         =   0   'False
            Width           =   3525
            _Version        =   65536
            _ExtentX        =   6218
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Form 7 Employer Information"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   61
            Left            =   5520
            TabIndex        =   308
            Top             =   3360
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   195
            Index           =   63
            Left            =   6390
            TabIndex        =   309
            Top             =   3615
            Visible         =   0   'False
            Width           =   3525
            _Version        =   65536
            _ExtentX        =   6218
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Health && Safety Injury WSIB Form 7"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   63
            Left            =   5520
            TabIndex        =   310
            Top             =   3600
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   195
            Index           =   62
            Left            =   6390
            TabIndex        =   504
            Top             =   3855
            Visible         =   0   'False
            Width           =   3525
            _Version        =   65536
            _ExtentX        =   6218
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Health && Safety WSIB Form 9"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   62
            Left            =   5520
            TabIndex        =   505
            Top             =   3840
            Visible         =   0   'False
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkViewOwnSuccPlan 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   8280
            TabIndex        =   292
            Tag             =   "40-View Own"
            Top             =   675
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "View Own"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkViewOwnEmpFlags 
            DataSource      =   "Data1"
            Height          =   225
            Left            =   2880
            TabIndex        =   515
            Tag             =   "40-View Own"
            Top             =   1680
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "View Own"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkViewOwnEmpHis 
            DataSource      =   "Data1"
            Height          =   225
            Left            =   2880
            TabIndex        =   516
            Tag             =   "40-View Own"
            Top             =   1920
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "View Own"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkViewOwnGLDist 
            DataSource      =   "Data1"
            Height          =   225
            Left            =   2880
            TabIndex        =   517
            Tag             =   "40-View Own"
            Top             =   2400
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "View Own"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUISecurity 
            Height          =   225
            Index           =   68
            Left            =   870
            TabIndex        =   525
            Top             =   4800
            Width           =   2565
            _Version        =   65536
            _ExtentX        =   4524
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Additional Payroll ID Data"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCheck chkUMSecurity 
            Height          =   225
            Index           =   68
            Left            =   0
            TabIndex        =   526
            Top             =   4800
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   397
            _StockProps     =   78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inquire"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   28
            Left            =   6405
            TabIndex        =   502
            Top             =   240
            Width           =   600
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Maintain"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   27
            Left            =   5510
            TabIndex        =   501
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WSIB Form 7 && 9 Security"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   26
            Left            =   5520
            TabIndex        =   496
            Top             =   3120
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Maintain"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   15
            Left            =   0
            TabIndex        =   313
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inquire"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   16
            Left            =   880
            TabIndex        =   312
            Top             =   240
            Width           =   600
         End
         Begin VB.Label lblHeading 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Basic 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   23
            Left            =   0
            TabIndex        =   311
            Top             =   0
            Width           =   780
         End
      End
   End
   Begin VB.Menu mnu_File 
      Caption         =   "File"
      Begin VB.Menu mnu_File_ESecure 
         Caption         =   "Exit Security"
      End
      Begin VB.Menu mnu_F_Sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Exit 
         Caption         =   "Exit info:HR"
      End
   End
   Begin VB.Menu mnu_security 
      Caption         =   "More Security"
      Begin VB.Menu mnu_Sec 
         Caption         =   "&Basic 1"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnu_Sec 
         Caption         =   "B&asic 2"
         Index           =   1
      End
      Begin VB.Menu mnu_Sec 
         Caption         =   "&Utilities"
         Index           =   2
      End
      Begin VB.Menu mnu_Sec 
         Caption         =   "&Reports"
         Index           =   3
      End
      Begin VB.Menu mnu_Sec 
         Caption         =   "&Mass Changes"
         Index           =   4
      End
      Begin VB.Menu mnu_Sec 
         Caption         =   "&Applicant Tracker"
         Index           =   5
      End
      Begin VB.Menu mnu_Sec 
         Caption         =   "&Course Admin"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_Sec 
         Caption         =   "&ESS && Timesheet Web Modules"
         Index           =   7
      End
      Begin VB.Menu mnu_Sec_Sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Attend_Reason 
         Caption         =   "&Attendance Reason"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_Codes 
         Caption         =   "&Codes"
      End
      Begin VB.Menu mnu_Comments 
         Caption         =   "C&omments"
      End
      Begin VB.Menu mnu_Door 
         Caption         =   "&Door Access"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_Dept 
         Caption         =   "Depar&tments"
      End
      Begin VB.Menu mnu_DocumentType 
         Caption         =   "Document Type"
      End
      Begin VB.Menu mnu_FollowUp 
         Caption         =   "&Follow Up"
      End
      Begin VB.Menu mnu_Cus 
         Caption         =   "C&ustom Features"
      End
      Begin VB.Menu mnu_CusRpt 
         Caption         =   "Cu&stom Report"
      End
      Begin VB.Menu mnu_Pension 
         Caption         =   "Pension System"
      End
      Begin VB.Menu mnu_Sec_ApplT 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Requisition 
         Caption         =   "&Requisition"
      End
   End
End
Attribute VB_Name = "frmSECURE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew%, fglbView%
Dim newNew As Boolean ' double check security
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim OUserID
Dim X%
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim tPass As String
Dim ChkPass
Dim ChangeCBox
Dim IDBack
Dim Qu
Dim OldExpireDays As Double
Dim OldPwd As String
Dim iCunt As Integer
Dim oldSecTemplate As String
Dim xEmployeeNoMissing As Boolean

Private Function chkSecureOk()

Dim SQLQ As String, Msg$, ID As Long
Dim snapSec As New ADODB.Recordset

Screen.MousePointer = HOURGLASS
chkSecureOk = False

On Error GoTo chkSecureOk_Err
txtUSERID = Trim(txtUSERID)
If Len(txtUSERID) <= 0 Then
    MsgBox "User ID is required"
    GoTo chkExit
End If
If Len(txtEEName) <= 0 Then
    MsgBox "User Name is required"
    GoTo chkExit
End If

'Ticket #22682 - Release 8.0
If chkNHireSecurity.Value = True And (chkMSecurity(0).Value = False Or chkSecurity(0).Value = False) Then
    MsgBox "You need 'Maintain' and 'Inquire' rights on 'Employee Demographics / Dates' as well to be able to 'Add New Hire'."
    Call HideAllShowOnePanDetail(0)
    chkMSecurity(0).SetFocus
    GoTo chkExit
End If

'Ticket #28635 - Add View Own security
If chkViewOwnPerform.Value = True And (chkMSecurity(6).Value = False And chkSecurity(6).Value = False) Then
    MsgBox "You need 'Maintain' and/or 'Inquire' rights on '" & lStr("Performance") & "' to be able to 'View Own'."
    Call HideAllShowOnePanDetail(0)
    chkMSecurity(6).SetFocus
    GoTo chkExit
End If

'Ticket #23923 - Release 8.0 - View Own Succession Planning
If chkViewOwnSuccPlan.Value = True And (chkUMSecurity(26).Value = False And chkUISecurity(26).Value = False) Then
    MsgBox "You need 'Maintain' and/or 'Inquire' rights on 'Succession Plan' to be able to 'View Own'."
    Call HideAllShowOnePanDetail(0)
    chkUMSecurity(26).SetFocus
    GoTo chkExit
End If

If chkViewOwnComm.Value = True And (chkMSecurity(25).Value = False And chkSecurity(25).Value = False) Then
    MsgBox "You need 'Maintain' and/or 'Inquire' rights on '" & lStr("Comments") & "' to be able to 'View Own'."
    Call HideAllShowOnePanDetail(0)
    chkMSecurity(25).SetFocus
    GoTo chkExit
End If

If chkViewOwnCounsel.Value = True And (chkMSecurity(23).Value = False And chkSecurity(23).Value = False) Then
    MsgBox "You need 'Maintain' and/or 'Inquire' rights on '" & lStr("Counseling") & "' to be able to 'View Own'."
    Call HideAllShowOnePanDetail(0)
    chkMSecurity(23).SetFocus
    GoTo chkExit
End If

If chkViewOwnFollUp.Value = True And (chkMSecurity(11).Value = False And chkSecurity(11).Value = False) Then
    MsgBox "You need 'Maintain' and/or 'Inquire' rights on '" & lStr("Follow-ups") & "' to be able to 'View Own'."
    Call HideAllShowOnePanDetail(0)
    chkMSecurity(11).SetFocus
    GoTo chkExit
End If

If chkViewOwnOthInfo.Value = True And (chkMSecurity(26).Value = False And chkSecurity(26).Value = False) Then
    MsgBox "You need 'Maintain' and/or 'Inquire' rights on 'Other Information' to be able to 'View Own'."
    Call HideAllShowOnePanDetail(0)
    chkMSecurity(26).SetFocus
    GoTo chkExit
End If

If chkViewOwnEmpFlags.Value = True And (chkUMSecurity(22).Value = False And chkUISecurity(22).Value = False) Then
    MsgBox "You need 'Maintain' and/or 'Inquire' rights on 'Employee Flags' to be able to 'View Own'."
    Call HideAllShowOnePanDetail(1)
    chkUMSecurity(22).SetFocus
    GoTo chkExit
End If

If chkViewOwnEmpHis.Value = True And (chkUMSecurity(23).Value = False And chkUISecurity(23).Value = False) Then
    MsgBox "You need 'Maintain' and/or 'Inquire' rights on 'Employee History' to be able to 'View Own'."
    Call HideAllShowOnePanDetail(1)
    chkUMSecurity(23).SetFocus
    GoTo chkExit
End If

If chkViewOwnGLDist.Value = True And (chkUMSecurity(24).Value = False And chkUISecurity(24).Value = False) Then
    MsgBox "You need 'Maintain' and/or 'Inquire' rights on '" & lStr("G/L") & " Distribution' to be able to 'View Own'."
    Call HideAllShowOnePanDetail(1)
    chkUMSecurity(24).SetFocus
    GoTo chkExit
End If


ID& = Val(lblID.Caption)
SQLQ = "SELECT * FROM HR_SECURE_BASIC "
SQLQ = SQLQ & "Where (USERID = '" & Replace(txtUSERID, "'", "''") & "'"
SQLQ = SQLQ & " AND ID <> " & ID & ") "
snapSec.Open SQLQ, gdbAdoIhr001, adOpenStatic
If snapSec.BOF And snapSec.EOF Then
   Rem everything is ok
Else
    MsgBox "This user already has a security record"
    GoTo chkExit
End If
snapSec.Close

'Ticket #24031 - Employee # required if Employee # Based Security
If chkEESecurity.Value = True And Len(elpEEID) = 0 And cmbSecTemplate <> "TEMPLATE" Then
    MsgBox "'Employee #' is required when 'Employee Number Based Security' is checked", vbOKOnly + vbExclamation, "info:HR Security"
    elpEEID.SetFocus
    GoTo chkExit
End If

'Ticket #24320 - Check if the Templated has Employee # Based Security checked. If so this User must have Employee # assigned.
If cmbSecTemplate <> "TEMPLATE" And cmbSecTemplate <> "" And Len(elpEEID) = 0 Then
    'Check if the Template has Employee # Based Security checked
    If EmployeeNoBasedSecurity(cmbSecTemplate) = 2 Then
        'No Security Profile found for this template
        cmbSecTemplate.SetFocus
        GoTo chkExit
    ElseIf EmployeeNoBasedSecurity(cmbSecTemplate) = 1 Then
        MsgBox "'Employee #' is mandatory. This User's Template has 'Employee Number Based Security' checked.", vbOKOnly + vbExclamation, "Employee # required: Employee Number Based Security"
        elpEEID.SetFocus
        GoTo chkExit
    End If
End If

If chkMSecurity(0).Value = True And chkEEADDRESS.Value = False Then
    Call mnu_Sec_Click(0)
    MsgBox "Cannot hide Address (Show Address), if 'Maintain' is checked for 'Employee Demographics / Dates' Security.", vbOKOnly + vbExclamation, "info:HR Security"
    chkMSecurity(0).SetFocus
    GoTo chkExit
'ElseIf chkMSecurity(0).Value = False And chkEEADDRESS.Value = True Then
'    Call mnu_Sec_Click(0)
'    MsgBox "Cannot 'Maintain' Employee Demographics / Dates' Security, if 'Show Address' is not checked.", vbOKOnly + vbExclamation, "info:HR Security"
'    chkMSecurity(0).SetFocus
'    GoTo chkExit
End If

If Len(txtExpireDays.Text) > 0 Then
    If Not IsNumeric(txtExpireDays.Text) Then
        MsgBox "Invalid Expiration Days"
        txtExpireDays.SetFocus
        GoTo chkExit
    End If
End If

If gsSECURED_PSW Then
    'Ticket #21685 Franks 03/06/2012 - begin
    If Len(txtExpireDays.Text) = 0 Then
        MsgBox "Expiration Days is required."
        txtExpireDays.SetFocus
        GoTo chkExit
    Else
        If Val(txtExpireDays.Text) = 0 Then
            MsgBox "Expiration Days is required."
            txtExpireDays.SetFocus
            GoTo chkExit
        End If
    End If
    'Ticket #21685 Franks 03/06/2012 - end
    
    If Not (OldPwd = txtSecPassword.Text) Then 'Password changed
        SQLQ = "SELECT * FROM HR_SECURE_BASIC "
        SQLQ = SQLQ & "Where (USERID = '" & Replace(txtUSERID, "'", "''") & "')"
        snapSec.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not (snapSec.BOF And snapSec.EOF) Then
            If txtSecPassword.Text = snapSec("PS_OLDPW") Or txtSecPassword.Text = snapSec("PS_OLDPW2") Or txtSecPassword.Text = snapSec("PS_OLDPW3") Then
                MsgBox "This password already has been used before."
                GoTo chkExit
            End If
        End If
        snapSec.Close
    End If
End If

If Len(elpEEID) > 0 Then
    SQLQ = "SELECT ED_EMPNBR FROM HREMP "
    SQLQ = SQLQ & "Where ED_EMPNBR = " & getEmpnbr(elpEEID)

    snapSec.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If snapSec.BOF And snapSec.EOF Then 'not a valid ee
        MsgBox "This employee does not exist."
        GoTo chkExit
    End If
    snapSec.Close
End If


If Len(txtPWord) < 1 Or Len(txtPWord) > 15 Then
    MsgBox "Invalid Password (must be between 1 and 15 characters).'"
    Call panVisible '10June99 js - selects panel to be shown
    txtPWord.SetFocus
    GoTo chkExit
End If

If gsSECURED_PSW Then
    If Len(txtPWord) < 8 Or Len(txtPWord) > 15 Then
        MsgBox "Invalid Password (must be between 8 and 15 characters).'"
        Call panVisible
        txtPWord.SetFocus
        GoTo chkExit
    End If
End If
chkSecureOk = True

chkExit:
Screen.MousePointer = DEFAULT
Exit Function

chkSecureOk_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HR_OCC_HEALTH_SAFETY", "edit/Add")
Call RollBack   '10June99 js

End Function

Private Sub chkACommentSecurity_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkACommentSecurity_KeyUp(KeyCode As Integer, Shift As Integer)
    If chkACommentSecurity.Value = True Then
        chkSecurity(25).Value = True
    End If
End Sub

Private Sub chkACommentSecurity_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkACommentSecurity.Value = True Then
        chkSecurity(25).Value = True
    End If
End Sub

Private Sub chkAISecurity_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If chkAISecurity(Index).Value = False Then
    chkAMSecurity(Index).Value = False
End If
End Sub

Private Sub chkAISecurity_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkAISecurity(Index).Value = False Then
    chkAMSecurity(Index).Value = False
End If
End Sub

Private Sub chkAMSecurity_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If chkAMSecurity(Index).Value = True Then
    chkAISecurity(Index).Value = True
End If
End Sub

Private Sub chkAMSecurity_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkAMSecurity(Index).Value = True Then
    chkAISecurity(Index).Value = True
End If
End Sub

Private Sub chkASecurity_KeyUp(KeyCode As Integer, Shift As Integer)
    'tkt10423 jerry said make it available to everyone
    'If glbCompSerial = "S/N - 2173W" Then
        If chkASecurity.Value = True Then
            chkSecurity(13).Value = True
        End If
    'End If
End Sub

Private Sub chkASecurity_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'tkt10423 jerry said make it available to everyone
  '  If glbCompSerial = "S/N - 2173W" Then
        If chkASecurity.Value = True Then
            chkSecurity(13).Value = True
        End If
   ' End If
End Sub

Private Sub chkEEADDRESS_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkEEDOB_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkEEMarital_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkMSecurity_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'If chkSecurity(Index).Value = False Then
'    chkMSecurity(Index).Value = False
'End If
If chkMSecurity(Index).Value = True Then
    chkSecurity(Index).Value = True
    
    '7.9 Enhancement Benefit/Beneficiary
    If chkMSecurity(8).Value = True Then
        chkMSecurity(29).Enabled = True
        chkSecurity(29).Enabled = True
    End If
Else
    '7.9 Enhancement - Benefits/Beneficiary
    If chkMSecurity(8).Value = False And chkSecurity(8).Value = False Then
        chkMSecurity(29).Value = False
        chkSecurity(29).Value = False
    End If
End If

End Sub

Private Sub chkMSecurity_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If chkSecurity(Index).Value Then
'    chkMSecurity(Index).Value = False
'End If
If chkMSecurity(Index).Value = True Then
    chkSecurity(Index).Value = True
    
    '7.9 Enhancement Benefit/Beneficiary
    If chkMSecurity(8).Value = True Then
        chkMSecurity(29).Enabled = True
        chkSecurity(29).Enabled = True
    End If
Else
    If chkMSecurity(8).Value = False And chkSecurity(8).Value = False Then
        chkMSecurity(29).Value = False
        chkSecurity(29).Value = False
    End If
End If

End Sub

Private Sub chkNHireSecurity_Click(Value As Integer)
    'Ticket #22682 - Release 8.0
    If chkNHireSecurity.Value = True Then
        chkMSecurity(0).Value = True
        chkSecurity(0).Value = True
    End If
End Sub

Private Sub chkNHireSecurity_GotFocus()
    'Ticket #22682 - Release 8.0
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkNHireSecurity_KeyUp(KeyCode As Integer, Shift As Integer)
    'Ticket #22682 - Release 8.0
    If chkNHireSecurity.Value = True Then
        chkSecurity(0).Value = True
    End If
End Sub

Private Sub chkNHireSecurity_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ticket #22682 - Release 8.0
    If chkNHireSecurity.Value = True Then
        chkSecurity(0).Value = True
    End If
End Sub

Private Sub chkPswdLocked_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkSecurity_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If chkSecurity(Index).Value = False Then
    chkMSecurity(Index).Value = False
    'tkt10423 jerry said make it available to everyone
    'If glbCompSerial = "S/N - 2173W" Then
        If chkSecurity(13).Value = False Then
            chkASecurity.Value = False
        End If
    'End If
    
    'Release 8.1
    If chkSecurity(25).Value = False Then
        chkACommentSecurity.Value = False
    End If
    
    '7.9 Enhancement Benefit/Beneficiary
    If chkSecurity(8).Value = False Then
        chkMSecurity(29).Value = False
        chkSecurity(29).Value = False
        chkMSecurity(29).Enabled = False
        chkSecurity(29).Enabled = False
    End If
Else
    '7.9 Enhancement Benefit/Beneficiary
    If chkSecurity(8).Value = True Then
        chkMSecurity(29).Enabled = True
        chkSecurity(29).Enabled = True
    End If
End If

'Ticket #28635 - Add View Own security
If Index = 6 Then
    If chkSecurity(Index).Value = False Then
        chkViewOwnPerform.Value = False
    End If
End If

'Ticket #23923 - View Own security
If Index = 11 Then
    If chkSecurity(Index).Value = False Then
        chkViewOwnFollUp.Value = False
    End If
End If
If Index = 25 Then
    If chkSecurity(Index).Value = False Then
        chkViewOwnComm.Value = False
    End If
End If
If Index = 23 Then
    If chkSecurity(Index).Value = False Then
        chkViewOwnCounsel.Value = False
    End If
End If
If Index = 26 Then
    If chkSecurity(Index).Value = False Then
        chkViewOwnOthInfo.Value = False
    End If
End If


End Sub

Private Sub chkSecurity_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkSecurity(Index).Value = False Then
    chkMSecurity(Index).Value = False
    
    '7.9 Enhancement Benefit/Beneficiary
    If chkSecurity(8).Value = False Then
        chkMSecurity(29).Value = False
        chkSecurity(29).Value = False
        chkMSecurity(29).Enabled = False
        chkSecurity(29).Enabled = False
    End If
Else
    '7.9 Enhancement Benefit/Beneficiary
    If chkSecurity(8).Value = True Then
        chkMSecurity(29).Enabled = True
        chkSecurity(29).Enabled = True
    End If
End If

'Ticket #28635 - Add View Own security
If Index = 6 Then
    If chkSecurity(Index).Value = False Then
        chkViewOwnPerform.Value = False
    End If
End If

'Ticket #23923 - View Own security
If Index = 11 Then
    If chkSecurity(Index).Value = False Then
        chkViewOwnFollUp.Value = False
    End If
End If
If Index = 25 Then
    If chkSecurity(Index).Value = False Then
        chkViewOwnComm.Value = False
    End If
End If
If Index = 23 Then
    If chkSecurity(Index).Value = False Then
        chkViewOwnCounsel.Value = False
    End If
End If
If Index = 26 Then
    If chkSecurity(Index).Value = False Then
        chkViewOwnOthInfo.Value = False
    End If
End If


End Sub

Private Sub chkViewOwnEmpFlags_Click(Value As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnEmpFlags.Value = True Then
        'chkUMSecurity(22).Value = True
        chkUISecurity(22).Value = True
    End If
End Sub

Private Sub chkViewOwnEmpFlags_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkViewOwnEmpFlags_KeyUp(KeyCode As Integer, Shift As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnEmpFlags.Value = True Then
        chkUISecurity(22).Value = True
    End If
End Sub

Private Sub chkViewOwnEmpFlags_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnEmpFlags.Value = True Then
        chkUISecurity(22).Value = True
    End If
End Sub

Private Sub chkViewOwnEmpHis_Click(Value As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnEmpHis.Value = True Then
        'chkUMSecurity(23).Value = True
        chkUISecurity(23).Value = True
    End If
End Sub

Private Sub chkViewOwnEmpHis_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkViewOwnEmpHis_KeyUp(KeyCode As Integer, Shift As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnEmpHis.Value = True Then
        chkUISecurity(23).Value = True
    End If
End Sub

Private Sub chkViewOwnEmpHis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnEmpHis.Value = True Then
        chkUISecurity(23).Value = True
    End If
End Sub

Private Sub chkViewOwnGLDist_Click(Value As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnGLDist.Value = True Then
        'chkUMSecurity(24).Value = True
        chkUISecurity(24).Value = True
    End If
End Sub

Private Sub chkViewOwnGLDist_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkViewOwnGLDist_KeyUp(KeyCode As Integer, Shift As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnGLDist.Value = True Then
        chkUISecurity(24).Value = True
    End If
End Sub

Private Sub chkViewOwnGLDist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnGLDist.Value = True Then
        chkUISecurity(24).Value = True
    End If
End Sub

Private Sub chkViewOwnPerform_Click(Value As Integer)
    'Ticket #28635 - Add View Own security
    If chkViewOwnPerform.Value = True Then
        chkSecurity(6).Value = True
    End If
End Sub

Private Sub chkViewOwnPerform_KeyUp(KeyCode As Integer, Shift As Integer)
    'Ticket #28635 - Add View Own security
    If chkViewOwnPerform.Value = True Then
        chkSecurity(6).Value = True
    End If
End Sub

Private Sub chkViewOwnPerform_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ticket #28635 - Add View Own security
    If chkViewOwnPerform.Value = True Then
        chkSecurity(6).Value = True
    End If
End Sub

Private Sub chkViewOwnSuccPlan_Click(Value As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnSuccPlan.Value = True Then
        'chkUMSecurity(26).Value = True
        chkUISecurity(26).Value = True
    End If
End Sub

Private Sub chkViewOwnSuccPlan_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkViewOwnSuccPlan_KeyUp(KeyCode As Integer, Shift As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnSuccPlan.Value = True Then
        chkUISecurity(26).Value = True
    End If
End Sub

Private Sub chkViewOwnSuccPlan_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnSuccPlan.Value = True Then
        chkUISecurity(26).Value = True
    End If
End Sub

Private Sub chkUISecurity_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If chkUISecurity(Index).Value = False Then
    chkUMSecurity(Index).Value = False
End If

'Ticket #23923 - View Own security
If Index = 26 Then
    If chkUISecurity(Index).Value = False Then
        chkViewOwnSuccPlan.Value = False
    End If
End If

'Ticket #23923 - View Own security
If Index = 22 Then
    If chkUISecurity(Index).Value = False Then
        chkViewOwnEmpFlags.Value = False
    End If
End If
If Index = 23 Then
    If chkUISecurity(Index).Value = False Then
        chkViewOwnEmpHis.Value = False
    End If
End If
If Index = 24 Then
    If chkUISecurity(Index).Value = False Then
        chkViewOwnGLDist.Value = False
    End If
End If

End Sub

Private Sub chkUISecurity_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkUISecurity(Index).Value = False Then
    chkUMSecurity(Index).Value = False
End If

'Ticket #23923 - View Own security
If Index = 26 Then
    If chkUISecurity(Index).Value = False Then
        chkViewOwnSuccPlan.Value = False
    End If
End If

'Ticket #23923 - View Own security
If Index = 22 Then
    If chkUISecurity(Index).Value = False Then
        chkViewOwnEmpFlags.Value = False
    End If
End If
If Index = 23 Then
    If chkUISecurity(Index).Value = False Then
        chkViewOwnEmpHis.Value = False
    End If
End If
If Index = 24 Then
    If chkUISecurity(Index).Value = False Then
        chkViewOwnGLDist.Value = False
    End If
End If

End Sub

Private Sub chkUMSecurity_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If chkUMSecurity(Index).Value = True Then
    chkUISecurity(Index).Value = True
End If
End Sub

Private Sub chkUMSecurity_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkUMSecurity(Index).Value = True Then
    chkUISecurity(Index).Value = True
End If
End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg$, INo&, X%
Dim flgTemplate As Boolean
Dim xTemplate As String

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

If txtUSERID = "999999999" Then
    Msg$ = "You can not delete the master security record."
    Msg$ = Msg$ & Chr(10) & "You can however, change its password."
    MsgBox Msg$, vbExclamation
    Exit Sub
End If

On Error GoTo Del_Err


Msg$ = "Are You Sure You Want To Delete "
'Ticket #20585 - If deleting template then warn accordingly
flgTemplate = False
xTemplate = ""
If cmbSecTemplate = "TEMPLATE" Then
    flgTemplate = True
    xTemplate = lblUSERID   'Template name
    Msg$ = Msg$ & Chr(10) & "This Security Template Record?  "
Else
    flgTemplate = False
    Msg$ = Msg$ & Chr(10) & "This Record?  "
End If

a% = MsgBox(Msg$, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

'????Ticket #24808 - First Update User's Profile using this Template with Template's Profile before deleting the template
If cmbSecTemplate = "TEMPLATE" Then
    Call Update_Users_withthis_Template(xTemplate, , "Delete")
End If

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute "DELETE FROM HR_SECURE_BASIC WHERE USERID='" & Replace(lblUSERID, "'", "''") & "'"
gdbAdoIhr001.Execute "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(lblUSERID, "'", "''") & "'"
gdbAdoIhr001.Execute "DELETE FROM HRPASDEP WHERE PD_USERID='" & Replace(lblUSERID, "'", "''") & "'"
gdbAdoIhr001.Execute "DELETE FROM HR_EMAIL WHERE EM_USERID='" & Replace(lblUSERID, "'", "''") & "'" 'Ticket #13545
gdbAdoIhr001.Execute "DELETE FROM HR_SECURE_COMMENTS WHERE USERID='" & Replace(lblUSERID, "'", "''") & "'"
gdbAdoIhr001.Execute "DELETE FROM HR_SECRPT WHERE USERID='" & Replace(lblUSERID, "'", "''") & "'"
gdbAdoIhr001.Execute "DELETE FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(lblUSERID, "'", "''") & "'"
gdbAdoIhr001.Execute "DELETE FROM HR_SECURE_ATTENDANCE WHERE USERID='" & Replace(lblUSERID, "'", "''") & "'"
gdbAdoIhr001.Execute "DELETE FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(lblUSERID, "'", "''") & "'" 'Release 8.1
gdbAdoIhr001.Execute "DELETE FROM HRA_SECURE_REQUISITION WHERE USERID='" & Replace(lblUSERID, "'", "''") & "'"     'Ticket #30508 - Applicant Tracking Enhancement
gdbAdoIhr001.CommitTrans

'????Ticket #24808 - Also update each of these users with this Template with the Template's Security Profile. I think it should be done before the Deletes above
'Ticket #20585 - Template deleted - clear all user's ref. to this template
If flgTemplate Then
    gdbAdoIhr001.Execute "UPDATE HR_SECURE_BASIC SET SECURE_TEMPLATE = '' WHERE SECURE_TEMPLATE='" & xTemplate & "'"
End If

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

Data1.Refresh

Call SET_UP_MODE
'Call mod_UpdateMode(False)

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_OCC_HEALTH_SAFETY", "Delete")
Call RollBack   '10June99 js

End Sub

Private Sub chkViewOwnComm_Click(Value As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnComm.Value = True Then
        'chkMSecurity(25).Value = True
        chkSecurity(25).Value = True
    End If
End Sub

Private Sub chkViewOwnComm_KeyUp(KeyCode As Integer, Shift As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnComm.Value = True Then
        chkSecurity(25).Value = True
    End If
End Sub

Private Sub chkViewOwnComm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnComm.Value = True Then
        chkSecurity(25).Value = True
    End If
End Sub

Private Sub chkViewOwnCounsel_Click(Value As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnCounsel.Value = True Then
        'chkMSecurity(23).Value = True
        chkSecurity(23).Value = True
    End If
End Sub

Private Sub chkViewOwnCounsel_KeyUp(KeyCode As Integer, Shift As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnCounsel.Value = True Then
        chkSecurity(23).Value = True
    End If
End Sub

Private Sub chkViewOwnCounsel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnCounsel.Value = True Then
        chkSecurity(23).Value = True
    End If
End Sub

Private Sub chkViewOwnFollUp_Click(Value As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnFollUp.Value = True Then
        'chkMSecurity(11).Value = True
        chkSecurity(11).Value = True
    End If
End Sub

Private Sub chkViewOwnFollUp_KeyUp(KeyCode As Integer, Shift As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnFollUp.Value = True Then
        chkSecurity(11).Value = True
    End If
End Sub

Private Sub chkViewOwnFollUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnFollUp.Value = True Then
        chkSecurity(11).Value = True
    End If
End Sub

Private Sub chkViewOwnOthInfo_Click(Value As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnOthInfo.Value = True Then
        'chkMSecurity(26).Value = True
        chkSecurity(26).Value = True
    End If
End Sub

Private Sub chkViewOwnOthInfo_KeyUp(KeyCode As Integer, Shift As Integer)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnOthInfo.Value = True Then
        chkSecurity(26).Value = True
    End If
End Sub

Private Sub chkViewOwnOthInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ticket #23923 - Release 8.0 - View Own
    If chkViewOwnOthInfo.Value = True Then
        chkSecurity(26).Value = True
    End If
End Sub

Private Sub cmbCountry_LostFocus()
lblCountry = cmbCountry
End Sub

Private Sub cmbSecTemplate_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbSecTemplate_LostFocus()
txtSecTemplate = cmbSecTemplate
End Sub

Private Sub cmbTemplate_LostFocus()
If cmbTemplate.Text <> "" Then
    lblTemplate = Trim(Split(cmbTemplate.Text, "-")(0))
End If
End Sub

Private Sub cmdCopySecuritys_Click()
    frmSecuCopy.IsCopyByPayID = False
    frmSecuCopy.txtFromUserID = txtUSERID
    frmSecuCopy.Show 1
    '------------------refreshing the form
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    Data1.Refresh
    Call SET_UP_MODE
    '----------------------
End Sub

Private Sub cmdFindUser_Click()
    Dim SaveID, SaveName, xTxtEEID, xLblEEName

    SaveID = glbLUserID
    SaveName = glbLUserNAME
    
    frmUFIND.Show 1
    
    If glbEEOK Then
        'txtUSERID = glbLUserID
        'txtEEName = glbLUserNAME
        'glbLUserNAME = SaveName
        'glbLUserID = SaveID

        If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
            glbSecUSERID = glbLUserID
                        
            Data1.Recordset.Requery
            Data1.Recordset.Find "USERID='" & Replace(glbLUserID, "'", "''") & "'"
            
            '????Ticket #24808 - If the User is Templated based then retrieve the Template Name of this User to retrieve
            'Template's Profile instead of User's Profile. If the User's Security is not based on Template or is TEMPLATE then
            'retrieve the respective User's record
            If cmbSecTemplate = "" Or cmbSecTemplate = "TEMPLATE" Then
                'User is normal user or is Template itself
                glbSecUSERID = glbLUserID
            Else
                'User's Profile is based on Template
                glbSecUSERID = cmbSecTemplate
            End If
            
            Call Display_Values1
            
            '????Ticket #24808 - Reset this global variable back to User ID
            glbSecUSERID = glbLUserID

        End If
    End If
End Sub

Private Sub cmdGrandInquAT_Click()
Dim X%

For X% = 0 To 13
    chkAISecurity(X%).Value = True
Next X%

End Sub

Private Sub cmdGrantAll_Click()
Dim X%

For X% = 0 To 22
    If X <> 2 Then chkUMSecurity(X%).Value = True
    If X <> 2 Then chkUISecurity(X%).Value = True
Next X%

'Attendance Code Matrix
For X% = 59 To 59
    chkUMSecurity(X%).Value = True
    chkUISecurity(X%).Value = True
Next X%

For X% = 38 To 38
    chkUMSecurity(X%).Value = True
    chkUISecurity(X%).Value = True
Next X%
For X% = 47 To 47
    chkUMSecurity(X%).Value = True
    chkUISecurity(X%).Value = True
Next X%

For X% = 0 To 5 '16    '15    '13 '8
    chkUSecurity(X%).Value = True
Next X%

For X% = 9 To 13
    chkUSecurity(X%).Value = True
Next X%

For X% = 16 To 16
    chkUSecurity(X%).Value = True
Next X%


For X% = 0 To 15
    chkIESecurity(X%).Value = True
Next X%
If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19935
    For X% = 16 To 19 '20 '19
        'no function for 21
        chkIESecurity(X%).Value = True
    Next X%
End If

'Ticket #22220, Ticket #22541, Ticket #23409, Ticket #24655, Ticket #25015
For X% = 64 To 68
    chkUMSecurity(X%).Value = True
    chkUISecurity(X%).Value = True
Next X%

'Ticket #22682 - Release 8.0 - Follow Up Code Email Matrix,Ticket #25746 - Department / GL Number Matrix, Ticket #25922 - OHRS Reporting for CHC
'Ticket #29122 - New Database Setup and Integration Setup securities
For X% = 69 To 73
    chkUMSecurity(X%).Value = True
    chkUISecurity(X%).Value = True
Next X%

'For x% = 70 To 70
'    chkUMSecurity(x%).Value = True
'    chkUISecurity(x%).Value = True
'Next x%

End Sub

Private Sub cmdGrantAll_GotFocus()
    'Hemu - 05/13/2003 Begin
    Call SetPanHelp(ActiveControl)
    'Hemu - 05/13/2003 End
End Sub

Private Sub chkEESecurity_GotFocus()
    'Hemu - 05/13/2003 Begin
    Call SetPanHelp(ActiveControl)
    'Hemu - 05/13/2003 End
End Sub

Private Sub chkEESIN_GotFocus()
    'Hemu - 05/13/2003 Begin
    Call SetPanHelp(ActiveControl)
    'Hemu - 05/13/2003 End
End Sub

Private Sub chkASecurity_GotFocus()
   'tkt10423 jerry said make it available to everyone
   ' If glbCompSerial = "S/N - 2173W" Then
        Call SetPanHelp(ActiveControl)
    'End If
End Sub

Private Sub cmdGrantAllMC_Click()
Dim X%
For X% = 0 To 19
    chkMCSecurity(X%).Value = True
Next X%

End Sub

Private Sub cmdGrantAllR_GotFocus()
    'Hemu - 05/13/2003 Begin
    Call SetPanHelp(ActiveControl)
    'Hemu - 05/13/2003 End
End Sub

Private Sub cmdGrantAllB_GotFocus()
    'Hemu - 05/13/2003 Begin
    Call SetPanHelp(ActiveControl)
    'Hemu - 05/13/2003 End
End Sub

Private Sub cmdGrantAllA_GotFocus()
    'Hemu - 05/13/2003 Begin
    Call SetPanHelp(ActiveControl)
    'Hemu - 05/13/2003 End
End Sub

Private Sub cmdGrantAllA_Click()
Dim X%

For X% = 0 To 13
    chkAMSecurity(X%).Value = True
    chkAISecurity(X%).Value = True
Next X%

End Sub

Private Sub cmdGrantAllB_Click()
Dim X%
 
For X% = 0 To 29
    chkMSecurity(X%).Value = True
    chkSecurity(X%).Value = True
Next X%
chkASecurity.Value = True 'Ticket #22009 Franks 05/10/2012
chkDSecurity.Value = True
chkEEDOB.Value = True
chkEEMarital.Value = True
chkEESIN.Value = True
chkEEADDRESS.Value = True
'Ticket #22682 - Release 8.0
chkNHireSecurity.Value = True

'Ticket #23923 - Release 8.0 - View Own
chkViewOwnFollUp.Value = True
chkViewOwnComm.Value = True
chkViewOwnCounsel.Value = True
chkViewOwnOthInfo.Value = True
'Ticket #28635 - Add View Own security
chkViewOwnPerform.Value = True

'Release 8.1
chkACommentSecurity.Value = True

End Sub

Private Sub cmdGrantAllR_Click()
Dim X%
 
For X% = 0 To 35 '34  'laura nov 17 changed from 31 to 32
    chkSSecurity(X%).Value = True
Next X%

'Overtime & other
For X% = 46 To 98
    If X% = 55 Then
        'Paid Sick hours is not available for Access users
        If glbSQL Or glbOracle Then
            chkSSecurity(55).Value = True
        End If
    Else
        chkSSecurity(X%).Value = True
    End If
Next X%

'Release 8.1
 chkSSecurity(99).Value = True
 
'Ticket #27795 - Friesens Corporation
 chkSSecurity(100).Value = True

End Sub

Sub cmdModify_Click()
Dim X%

newNew = False

On Error GoTo Mod_Err

OUserID = txtUSERID
oldSecTemplate = cmbSecTemplate

'Call mod_UpdateMode(True)

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack   '10June99 js

End Sub

Sub cmdNew_Click()
Dim SQLQ As String

cmdCopySecuritys.Enabled = False

Dim VR

newNew = True
fglbNew% = True

Call ChkCBoxChange(cmbSecTemplate)

If ChangeCBox = True Then
    VR = MsgBox("Do you want to save changes?", MB_YESNO)
    If VR = IDYES Then
        Me.cmdOK_Click
    ElseIf VR = IDNO Then
        'Call Me.cmdCancel_Click
    End If
End If

panEEDESC.Enabled = False

Call SET_UP_MODE

On Error GoTo AddN_Err

Data1.Recordset.AddNew
''' Sam add July 2002 * Remove Binding Control
'Call Set_Control("B", Me)
'rsDATA.AddNew

chkEESecurity = True
lblCNum.Caption = "001"

'Ticket #21685 Franks 03/06/2012
LastExpireDate.Text = Date  'Format(Date, "SHORT DATE")

Call panVisible

txtUSERID.SetFocus
chkEEMarital.Value = True

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_OCC_HEALTH_SAFETY", "Add")
Call RollBack   '10June99 js

End Sub

Sub cmdOK_Click()
Dim xID
Dim xTemplate As String

txtUSERID.SetFocus

DoEvents

If IsNull(Data1.Recordset("PassWord")) Then
    OldPwd = ""
Else
    OldPwd = Data1.Recordset("PassWord")
End If

If gsMultiLang = "YES" Then 'whscc
    If txtPWord.Text <> DecryptPasswordMultiLang(OldPwd) Then
        glbConfPass = txtPWord
        Load frmConfPass
        frmConfPass.Show vbModal
        If glbConfPass = "" Then
            Exit Sub
        End If
    End If
Else
    If txtPWord.Text <> DecryptPassword(OldPwd) Then
        glbConfPass = txtPWord
        Load frmConfPass
        frmConfPass.Show vbModal
        If glbConfPass = "" Then
            Exit Sub
        End If
    End If
End If

On Error GoTo Add_Err

'Ticket #20585 - If Template then update users with this template as well.
'If User and with no template then update that user's profile.
'if User and with Template then do not update user's profile.
'Get the Template Name of this User ID
xTemplate = cmbSecTemplate  'Get_Template(glbSecUSERID)

If xTemplate = "TEMPLATE" Then
    'Update all users with this template. After the changes are saved
        
ElseIf xTemplate = "" Then
    'User - User with no template - don't do anything let system update user's profile
ElseIf xTemplate <> "TEMPLATE" And Not fglbNew% Then
    'User with template - do not allow to save these changes.
    'A User can became template based too.
    
    'Check if security was changed, if so then only display the message below
    Call ChkCBoxChange(xTemplate)
    
    If ChangeCBox = True Then
        MsgBox "This User's Security Profile cannot be changed except for Country, Timesheet & Security Template, Password & Expiration and Department Security. " & vbCrLf & "This user's security profile is based on the '" & xTemplate & "' template.", vbExclamation, "Template based User Security Profile"
    End If
End If

'if Template or User
'If xTemplate = "TEMPLATE" Or xTemplate = "" Then

    If Not chkSecureOk() Then Exit Sub
    
    'Ticket #30299 - Update User's records with templates with the Employee # Based Security value if changed at the template level.
    If xTemplate = "TEMPLATE" Then
        'Update all user with this Template with Employee # Based Security value if changed
        If EmployeeNoBased_Security_Changed(txtUSERID) Then
            'Update user's Security record with Employee # Based Security
            Call Update_Users_EmployeeNoBasedSecurity(txtUSERID)
        End If
    End If

    Call UpdUStats(Me) ' update user's stats (who did it and when)
    
    'Ticket #21629 - Jerry asked the new user's Country should be based on the Company Master
    If newNew Or fglbNew% Then
        If lblCountry = "" Then
            lblCountry = getCompanyMasterData("PC_COUNTRY")
        End If
    End If
        
    panEEDESC.Enabled = True
    
    xID = txtUSERID
        
    If Len(elpEEID) = 0 Then
        Data1.Recordset("EMPNBR") = Null
    Else
        Data1.Recordset("EMPNBR") = getEmpnbr(elpEEID)
    End If
    Data1.Recordset("USERID") = xID
    If Len(txtExpireDays.Text) = 0 Then 'Ticket #21685 Franks 03/06/2012
        Data1.Recordset("PS_EXPIR_DAYS") = 0 'Null
    End If
        
    Data1.Recordset.UpdateBatch
'End If

If newNew Or fglbNew% Then
    'Ticket #20585 - If adding as Template or without Template then follow the normal security Add function
    'otherwise add security profile based on the Template.
    If cmbSecTemplate = "" Or cmbSecTemplate = "TEMPLATE" Then
        'Adding either the Template itself or profile without Template
        Call AddSecAccess
    Else
        '????Ticket #24808 - Do not save the security settings - just save the User Name, Country, TS & Security Template & Dept Security
        'Adding with Template - create user according to the Template profile
        'Call AddSecAccess_From_Template
        Call Template_Based_Security_Profile_Update(xID, cmbSecTemplate.Text, "Add")
    End If
Else
    'Ticket #20585 - If updating as Template or without Template then follow the normal security Add
    'function otherwise update security profile using the Template.
    If oldSecTemplate <> cmbSecTemplate Then
        If cmbSecTemplate = "" Or cmbSecTemplate = "TEMPLATE" Then
            'Updating either the Template itself or profile without Template
            Call UpdSecAccess
            
            'Template changed - update all users with this template since the Template Profile has
            'changed. But if the template became a user then remove the User's association to
            'this template
            If oldSecTemplate = "TEMPLATE" And cmbSecTemplate = "" Then
                '????Ticket #24808 - First update each user who were assigned to this template with Template with this Template's Profile
                Call Update_Users_withthis_Template(xID, , "Delete")
                
                'A Template became a User, remove the association from the Users with this template
                Call Remove_User_Template_Association(xID)
                
            ElseIf cmbSecTemplate = "TEMPLATE" Then
                'User became Template? There won't be any user with this new template but in anycase
                'just check if there are any users with this new template to update with the changes.
                'Call Update_Users_withthis_Template(xID)
            End If
        Else
            '????Ticket #24808 - User's Profile only needs to be updated with Template Name - no other security rights.
            'User's Template changed - update user's profile with the new Template Profile assigned
            'Call Template_Based_Security_Profile_Update(xID, cmbSecTemplate.Text, "Update")
            'Ticket #29544 - Delete the existing profile setup of the User as the user is now part of the template
            Call Delete_Existing_User_Profile
        End If
    Else
        'Template did not change but if any of the security settings has changed then it should only be
        'updated on the User's Profile if not template based or is Template itself.
        If cmbSecTemplate = "" Or cmbSecTemplate = "TEMPLATE" Then
            'Updating either Template itself or profile without Template
            Call UpdSecAccess
                        
            'Template profile has changed - update all users with this template since the template
            'profile has changed. But if the template became a user then remove the User's association to
            'this template
            If oldSecTemplate = "TEMPLATE" And cmbSecTemplate = "" Then
                '????Ticket #24808 - First update each user who were assigned to this template with Template with this Template's Profile
                Call Update_Users_withthis_Template(xID, , "Delete")
                
                'A Template became a User, remove the association from the Users with this template
                Call Remove_User_Template_Association(xID)
                
            ElseIf cmbSecTemplate = "TEMPLATE" Then
                'User became Template? There won't be any user with this new template but in anycase
                'just check if there are any users with this new template to update with the changes.
                If OUserID <> txtUSERID Then
                    'If the Template Name has changed then users associated with the original name
                    'will still be with Old Name so update those first. Later in the code we are
                    'changing the Template Name for all those user as well.
                    Call Update_Users_withthis_Template(OUserID, txtUSERID)
                Else
                    '????Ticket #24808 - Don't save with new Template Profile settings - just save the Security Template Name to the User's Profile
                    'Call Update_Users_withthis_Template(xID)
                End If
                'Ticket #24320 - Warn User that some Users belonging to this template does not have Employee # assigned
                'when Template's Employee # Based Security is checked
                If xEmployeeNoMissing Then
                    MsgBox "Some Users assigned to this Template does not have 'Employee #' assigned. " & vbCrLf & "'Employee #' is mandatory when 'Employee # Based Security' is checked.", vbExclamation, "Employee # Based Security"
                End If
            End If
        Else
            'If only any security was changed of the user
            If ChangeCBox Then
                'User's Profile changed but this user has Template associated, so reset User's profile to
                'Template Profile settings
                Call Template_Based_Security_Profile_Update(xID, cmbSecTemplate.Text, "Reset")
            End If
        End If
    End If
    
    'User ID changed
    If OUserID <> txtUSERID Then
        Call UpdateRelated
        
        'If Template's Name (User ID) changed then update respective Users with the new Template name
        If cmbSecTemplate = "TEMPLATE" Then
            '????Ticket #24808 - Don't save with new Template Profile settings - just save the Security Template Name to the User's Profile
            Call Update_Users_with_NewTemplate_Name(OUserID, txtUSERID)
            
            'Repopulate the Security Template combo box
            Call Populate_Security_Template
        End If
    End If
End If

If fglbNew% Then
    'Ticket #20585 - Add the default Department Security for only those users who are Template or
    'user's without any Template
    If cmbSecTemplate = "" Or cmbSecTemplate = "TEMPLATE" Then
        Call AddSecDept 'v8.0
    End If
End If

If gsSECURED_PSW Then 'Ticket #12707
    If Not (OldPwd = txtSecPassword.Text) Then 'Password changed
        Call UpdPswExpireDatac(xID, OldPwd)
    End If
End If

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

Data1.Refresh
Data1.Recordset.Find "USERID='" & Replace(xID, "'", "''") & "'"

newNew = False
fglbNew% = False

Dim ctylist

ctylist = CountryList
OUserID = txtUSERID
oldSecTemplate = cmbSecTemplate

Call SET_UP_MODE
'Call mod_UpdateMode(False)

cmdCopySecuritys.Enabled = True

'UPDATE TEMPLATE INFO

If Len(Trim(cmbTemplate.Text)) > 0 Then
    UpdateTSTemplate xID, Trim(Split(cmbTemplate.Text, "-")(0))
End If
Call FindTemplate

Exit Sub

Add_Err:
If Err = 3022 Then
    Data1.Recordset.CancelUpdate    ' no dups
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_SECURE", "Update")
Call RollBack   '10June99 js

End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport, X%

'cmdPrint.Enabled = False

If mnu_Sec(3).Checked Then
    Me.vbxCrystal.WindowTitle = "Appplicant Tracking Security Master Report"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 3
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RGSECAPP.rpt"
Else
    Me.vbxCrystal.WindowTitle = "Security Master Report"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 4
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next
        Me.vbxCrystal.DataFiles(5) = glbIHRDBW
        Me.vbxCrystal.DataFiles(6) = glbIHRDB
        Me.vbxCrystal.DataFiles(7) = glbIHRDB
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RGSECURE.rpt"
    Call SECWRK
End If
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub

Sub cmdView_Click()
Dim RHeading As String, xReport, X%

'cmdPrint.Enabled = False

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

If mnu_Sec(3).Checked Then
    Me.vbxCrystal.WindowTitle = "Appplicant Tracking Security Master Report"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 3
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RGSECAPP.rpt"
Else
    Me.vbxCrystal.WindowTitle = "Security Master Report"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 4
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next
        Me.vbxCrystal.DataFiles(5) = glbIHRDBW
        Me.vbxCrystal.DataFiles(6) = glbIHRDB
        Me.vbxCrystal.DataFiles(7) = glbIHRDB
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RGSECURE.rpt"
    Call SECWRK
End If
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub

Private Sub cmdGrantAllWeb_Click()
For X% = 6 To 8
    chkUSecurity(X%).Value = True
Next X%
For X% = 14 To 15
    chkUSecurity(X%).Value = True
Next X%
For X% = 17 To 47   '35 '26
    chkUSecurity(X%).Value = True
Next X%
End Sub

Private Sub cmdGrantCourseAdmin_Click()
Dim X%
For X% = 27 To 34
    chkUMSecurity(X%).Value = True
    chkUISecurity(X%).Value = True
Next X%
For X% = 36 To 45
    chkSSecurity(X%).Value = True
Next X%
chkUMSecurity(2).Value = True
chkUISecurity(2).Value = True

End Sub

Private Sub cmdGrantInqu_Click()
Dim X%
For X% = 0 To 29
    chkSecurity(X%).Value = True
Next X%

End Sub

Private Sub cmdGrantInqu2_Click()
Dim X%
For X% = 22 To 26
    chkUMSecurity(X%).Value = False
    chkUISecurity(X%).Value = True
Next X%
For X% = 35 To 37
    chkUMSecurity(X%).Value = False
    chkUISecurity(X%).Value = True
Next X%
For X% = 39 To 40
    chkUMSecurity(X%).Value = False
    chkUISecurity(X%).Value = True
Next X%
For X% = 42 To 58   '46
    chkUMSecurity(X%).Value = False
    chkUISecurity(X%).Value = True
Next X%

chkUMSecurity(60).Value = False
chkUISecurity(60).Value = True
If glbWSIBModule Then  'WSIB Form 7 - Billable Module
    chkUMSecurity(61).Value = False
    chkUISecurity(61).Value = True
    chkUMSecurity(63).Value = False
    chkUISecurity(63).Value = True
    
    chkUMSecurity(62).Value = False 'Form 9
    chkUISecurity(62).Value = True  'Form 9
End If
'Samuel Ticket #21000 Franks 09/26/2011 - move to Custom Security
'If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #20052 Franks 07/20/2011
'    chkUISecurity(62).Value = True
'End If

'Ticket #25015 - Macaulay - Additional Payroll ID Data
chkUMSecurity(68).Value = False
chkUISecurity(68).Value = True

End Sub

Private Sub cmdGrantInquCourseAdmin_Click()
Dim X%

For X% = 27 To 34
    chkUISecurity(X%).Value = True
Next X%
For X% = 36 To 45
    chkSSecurity(X%).Value = True
Next X%
chkUISecurity(2).Value = True

End Sub

Private Sub cmdGrantInquire_Click()
Dim X%

For X% = 0 To 22
    If X <> 2 Then chkUISecurity(X%).Value = True
Next X%

'Attendance Code Matrix
For X% = 59 To 59
    chkUISecurity(X%).Value = True
Next X%

For X% = 38 To 38
    chkUISecurity(X%).Value = True
Next X%

For X% = 47 To 47
    chkUISecurity(X%).Value = True
Next X%

'Ticket #22220, Ticket #22541, Ticket #23409, Ticket #24655, Ticket #25015
For X% = 64 To 68
    chkUISecurity(X%).Value = True
Next X%

'Ticket #22682 - Release 8.0 - Follow Up Code Email Matrix, 'Ticket #25746 - Department / GL Number Matrix, Ticket #25922 - OHRS Reporting for CHC
'Ticket #29122 - New Database Setup and Integration Setup securities
For X% = 69 To 73
    chkUISecurity(X%).Value = True
Next X%

'Ticket #25746 - Department / GL Number Matrix
'For x% = 70 To 70
'    chkUISecurity(x%).Value = True
'Next x%

End Sub

Private Sub cmdGrantInquire2_Click()
Dim X%
For X% = 22 To 26
    chkUMSecurity(X%).Value = True
    chkUISecurity(X%).Value = True
Next X%
For X% = 35 To 37
    chkUMSecurity(X%).Value = True
    chkUISecurity(X%).Value = True
Next X%
For X% = 39 To 40
    chkUMSecurity(X%).Value = True
    chkUISecurity(X%).Value = True
Next X%
For X% = 42 To 58   '46
    chkUMSecurity(X%).Value = True
    chkUISecurity(X%).Value = True
Next X%

chkUMSecurity(60).Value = True
chkUISecurity(60).Value = True
If glbWSIBModule Then  'WSIB Form 7 - Billable Module
    chkUMSecurity(61).Value = True
    chkUISecurity(61).Value = True
    chkUMSecurity(63).Value = True
    chkUISecurity(63).Value = True
    
    chkUMSecurity(62).Value = True  'Form 9
    chkUISecurity(62).Value = True  'Form 9
End If
'Samuel Ticket #21000 Franks 09/26/2011 - move to Custom Security
'If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #20052 Franks 07/20/2011
'    chkUMSecurity(62).Value = True
'    chkUISecurity(62).Value = True
'End If

'Ticket #23923 - Release 8.0 - View Own
chkViewOwnSuccPlan.Value = True
chkViewOwnEmpFlags.Value = True
chkViewOwnEmpHis.Value = True
chkViewOwnGLDist.Value = True

'Ticket #25015 - Macaulay - Additional Payroll ID Data
chkUMSecurity(68).Value = True
chkUISecurity(68).Value = True

End Sub

Private Sub cmdPageLeft_Click(Index As Integer)
    'Ticket #22295
    'Me.cmdOK_Click
    If Index = 7 Then
        Call mnu_Sec_Click(Index - 2)
    ElseIf Index = 6 Then
        Call mnu_Sec_Click(Index + 1)
    Else
        Call mnu_Sec_Click(Index - 1)
    End If
End Sub

Private Sub cmdPageRight_Click(Index As Integer)
    'Ticket #22295
    'Me.cmdOK_Click
    If Index = 5 Then
        Call mnu_Sec_Click(Index + 2)
    ElseIf Index = 7 Then
        Call mnu_Sec_Click(Index - 1)
    Else
        Call mnu_Sec_Click(Index + 1)
    End If
End Sub

Private Sub cmdRemove_Click()
Dim X%
For X% = 0 To 22
    If X <> 2 Then chkUMSecurity(X%).Value = False
    If X <> 2 Then chkUISecurity(X%).Value = False
Next X%

'Attendance Code Matrix
For X% = 59 To 59
    chkUMSecurity(X%).Value = False
    chkUISecurity(X%).Value = False
Next X%

For X% = 38 To 38
    chkUMSecurity(X%).Value = False
    chkUISecurity(X%).Value = False
Next X%
For X% = 47 To 47
    chkUMSecurity(X%).Value = False
    chkUISecurity(X%).Value = False
Next X%

For X% = 0 To 5 '16    '15    '13 '8
    chkUSecurity(X%).Value = False
Next X%

For X% = 9 To 13
    chkUSecurity(X%).Value = False
Next X%

For X% = 16 To 16
    chkUSecurity(X%).Value = False
Next X%

For X% = 0 To 15
    chkIESecurity(X%).Value = False
Next X%
If glbCompSerial = "S/N - 2382W" Then   'Samuel  - Ticket #19935
    For X% = 16 To 19 '21 '19
        chkIESecurity(X%).Value = False
    Next X%
End If

'Ticket #22220, Ticket #22541, Ticket #23409, Ticket #24655, Ticket #25015
For X% = 64 To 68
    chkUMSecurity(X%).Value = False
    chkUISecurity(X%).Value = False
Next X%

'Ticket #22682 - Release 8.0 - Follow Up Code Email Matrix, Ticket #25746 - Department / GL Number Matrix, Ticket #25922 - OHRS Reporting for CHC
'Ticket #29122 - New Database Setup and Integration Setup securities
For X% = 69 To 73
    chkUMSecurity(X%).Value = False
    chkUISecurity(X%).Value = False
Next X%

'Ticket #25746 - Department / GL Number Matrix
'For x% = 70 To 70
'    chkUMSecurity(x%).Value = False
'    chkUISecurity(x%).Value = False
'Next x%

End Sub

Private Sub cmdRemoveAll_Click()
Dim X%
 
For X% = 0 To 29
    chkMSecurity(X%).Value = False
    chkSecurity(X%).Value = False
Next X%
chkASecurity.Value = False
chkDSecurity.Value = False 'Ticket #22009 Franks 05/10/2012
chkEEDOB.Value = False
chkEEMarital.Value = False
chkEESIN.Value = False
chkEEADDRESS.Value = False
'Ticket #22682 - Release 8.0
chkNHireSecurity.Value = False

'Ticket #23923 - Release 8.0 - View Own
chkViewOwnFollUp.Value = False
chkViewOwnComm.Value = False
chkViewOwnCounsel.Value = False
chkViewOwnOthInfo.Value = False
'Ticket #28635 - Add View Own security
chkViewOwnPerform.Value = False

'Release 8.1
chkACommentSecurity.Value = False

End Sub

Private Sub cmdRemoveAll2_Click()
Dim X%
For X% = 22 To 26
    chkUMSecurity(X%).Value = False
    chkUISecurity(X%).Value = False
Next X%
For X% = 35 To 37
    chkUMSecurity(X%).Value = False
    chkUISecurity(X%).Value = False
Next X%
For X% = 39 To 40
    chkUMSecurity(X%).Value = False
    chkUISecurity(X%).Value = False
Next X%
For X% = 42 To 58   '46
    chkUMSecurity(X%).Value = False
    chkUISecurity(X%).Value = False
Next X%
chkUMSecurity(60).Value = False
chkUISecurity(60).Value = False
If glbWSIBModule Then  'WSIB Form 7 - Billable Module
    chkUMSecurity(61).Value = False
    chkUISecurity(61).Value = False
    chkUMSecurity(63).Value = False
    chkUISecurity(63).Value = False
    
    chkUMSecurity(62).Value = False 'Form 9
    chkUISecurity(62).Value = False 'Form 9
End If

'Samuel Ticket #21000 Franks 09/26/2011 - move to Custom Security
'If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #20052 Franks 07/20/2011
'    chkUMSecurity(62).Value = 0
'    chkUISecurity(62).Value = 0
'End If

'Ticket #23923 - Release 8.0 - View Own
chkViewOwnSuccPlan.Value = False
chkViewOwnEmpFlags.Value = False
chkViewOwnEmpHis.Value = False
chkViewOwnGLDist.Value = False

'Ticket #25015 - Macaulay - Additional Payroll ID Data
chkUMSecurity(68).Value = False
chkUISecurity(68).Value = False

End Sub

Private Sub cmdRemoveAllAT_Click()
Dim X%

For X% = 0 To 13
    chkAMSecurity(X%).Value = False
    chkAISecurity(X%).Value = False
Next X%

End Sub

Private Sub cmdRemoveAllMC_Click()
Dim X%
For X% = 0 To 19
    chkMCSecurity(X%).Value = False
Next X%
End Sub

Private Sub cmdRemoveAllWeb_Click()
For X% = 6 To 8
    chkUSecurity(X%).Value = False
Next X%
For X% = 14 To 15
    chkUSecurity(X%).Value = False
Next X%
For X% = 17 To 47   '35 '26
    chkUSecurity(X%).Value = False
Next X%
End Sub

Private Sub cmdRemoveCourseAdmin_Click()
Dim X%
For X% = 27 To 34
    chkUMSecurity(X%).Value = False
    chkUISecurity(X%).Value = False
Next X%
For X% = 36 To 45
    chkSSecurity(X%).Value = False
Next X%

chkUMSecurity(2).Value = False
chkUISecurity(2).Value = False

End Sub

Private Sub cmdScreenLeft_Click()
    'Ticket #28540 - Security access with left and right moves on security screens.

    If panDetails(0).Visible Then       'Basic 1
        'Do nothing - no other security screen on the left
        cmdScreenRight.Visible = True
        cmdScreenLeft.Visible = False
    ElseIf panDetails(1).Visible Then   'Basic 2
        panDetails(1).Visible = False
        panDetails(0).Visible = True
        
        cmdScreenRight.Visible = True
        cmdScreenLeft.Visible = False
    ElseIf panDetails(2).Visible Then   'Utilities
        panDetails(2).Visible = False
        panDetails(1).Visible = True
        
        cmdScreenRight.Visible = True
        cmdScreenLeft.Visible = True
    ElseIf panDetails(3).Visible Then   'Reports
        panDetails(3).Visible = False
        panDetails(2).Visible = True
        
        cmdScreenRight.Visible = True
        cmdScreenLeft.Visible = True
    ElseIf panDetails(4).Visible Then   'Mass Update
        panDetails(4).Visible = False
        panDetails(3).Visible = True
        
        cmdScreenRight.Visible = True
        cmdScreenLeft.Visible = True
    ElseIf panDetails(5).Visible Then   'Applicant Tracking
        panDetails(5).Visible = False
        panDetails(4).Visible = True
        
        cmdScreenRight.Visible = True
        cmdScreenLeft.Visible = True
    ElseIf panDetails(7).Visible Then   'ESS & Timesheet
        panDetails(7).Visible = False
        panDetails(5).Visible = True
        
        cmdScreenRight.Visible = True
        cmdScreenLeft.Visible = True
    ElseIf panDetails(6).Visible Then   'Course Admin
        panDetails(6).Visible = False
        panDetails(7).Visible = True
        
        'Ticket #22682 - Jerry wants Course Admin visible for Serial #9999 and KPAS
        If (glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 9999W") Then
            cmdScreenRight.Visible = True
        Else
            'Do nothing - last security screen in the order set for the rest of the clients
            cmdScreenRight.Visible = False
        End If
        
        cmdScreenLeft.Visible = True
    End If
End Sub

Private Sub cmdScreenLeft_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdScreenRight_Click()
    'Ticket #28540 - Security access with left and right moves on security screens.

    If panDetails(0).Visible Then       'Basic 1
        panDetails(0).Visible = False
        panDetails(1).Visible = True
        
        cmdScreenRight.Visible = True
        cmdScreenLeft.Visible = True
    ElseIf panDetails(1).Visible Then   'Basic 2
        panDetails(1).Visible = False
        panDetails(2).Visible = True
        
        cmdScreenRight.Visible = True
        cmdScreenLeft.Visible = True
    ElseIf panDetails(2).Visible Then   'Utilities
        panDetails(2).Visible = False
        panDetails(3).Visible = True
        
        cmdScreenRight.Visible = True
        cmdScreenLeft.Visible = True
    ElseIf panDetails(3).Visible Then   'Reports
        panDetails(3).Visible = False
        panDetails(4).Visible = True
        
        cmdScreenRight.Visible = True
        cmdScreenLeft.Visible = True
    ElseIf panDetails(4).Visible Then   'Mass Update
        panDetails(4).Visible = False
        panDetails(5).Visible = True
        
        cmdScreenRight.Visible = True
        cmdScreenLeft.Visible = True
    ElseIf panDetails(5).Visible Then   'Applicant Tracking
        panDetails(5).Visible = False
        panDetails(7).Visible = True
        
        'Ticket #22682 - Jerry wants Course Admin visible for Serial #9999 and KPAS
        If (glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 9999W") Then
            cmdScreenRight.Visible = True
        Else
            cmdScreenRight.Visible = False
        End If
        cmdScreenLeft.Visible = True
    ElseIf panDetails(7).Visible Then   'ESS & Timesheet
        'Ticket #22682 - Jerry wants Course Admin visible for Serial #9999 and KPAS
        If (glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 9999W") Then
            panDetails(7).Visible = False
            panDetails(6).Visible = True    'Course Admin
        Else
            'Do nothing as this screen is the last security screen in the defined order for rest of the clients
        End If
        cmdScreenRight.Visible = False
        cmdScreenLeft.Visible = True
    ElseIf panDetails(6).Visible Then
        'Do nothing as this screen is the last security screen in the defined order
        
        cmdScreenRight.Visible = False
        cmdScreenLeft.Visible = True
    End If
End Sub

Private Sub cmdScreenRight_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Command2_Click()
Dim X%
For X% = 0 To 35
    chkSSecurity(X%).Value = False
Next X%
'Overtime & others
For X% = 46 To 98
    chkSSecurity(X%).Value = False
Next X%

'Release 8.1
chkSSecurity(99).Value = False

'Ticket #27795 - Friesens Corporation
chkSSecurity(100).Value = False
End Sub

Private Sub elpEmpLookup_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
'Call INI_Controls(Me)
Call SET_UP_MODE
glbOnTop = "FRMSECURE"
glbFNo = 2

End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = Me.name

Dim Answer, DefVal, Msg, Title, X%  '  variables.
Dim RFound As Integer ' records found
Dim SQLQ

On Error GoTo SecureLoad_Err

glbOnTop = "FRMSECURE"

Screen.MousePointer = HOURGLASS

newNew = False
fglbNew% = False
Data1.ConnectionString = glbAdoIHRDB

mnu_FollowUp.Caption = lStr("Follow-ups")   'lStr("Follow-Up")
mnu_Comments.Caption = lStr("Comments")

SQLQ = "SELECT *, "
If glbLinamar Then
    SQLQ = SQLQ & " CASE WHEN EMPNBR IS NOT NULL AND LEN(EMPNBR)>2 "
    SQLQ = SQLQ & " THEN RIGHT(EMPNBR,3)+'-'+"
    SQLQ = SQLQ & " LEFT(EMPNBR,LEN(EMPNBR)-3) "
    SQLQ = SQLQ & " ELSE STR(EMPNBR) END "
    SQLQ = SQLQ & " AS SEMPNBR "
Else
    SQLQ = "SELECT * "
    vbxTrueGrid.Columns(1).DataField = "EMPNBR"
End If
SQLQ = SQLQ & " from HR_SECURE_BASIC ORDER BY USERNAME"
Data1.RecordSource = SQLQ
Data1.Refresh

Dim ctylist

ctylist = CountryList
X = 1
Do While X > 0
    X = InStr(ctylist, "&")
    If X > 0 Then
        cmbCountry.AddItem Left(ctylist, X - 1)
        'cmbCountryOfEmp.AddItem Left(ctylist, x - 1)
        ctylist = Mid(ctylist, X + 1)
    Else
        cmbCountry.AddItem ctylist
        'cmbCountryOfEmp.AddItem ctylist
    End If
Loop

cmbCountry.ListIndex = 0

'Mostafa - Template auto selection
lblTemplate.DataField = "TS_TPID"

Dim rsTemplates As New ADODB.Recordset
SQLQ = "SELECT * FROM HR_TEMPLATE where tp_name is not null or Len(ltrim(rtrim(tp_name))) > 0"

cmbTemplate.AddItem ""
rsTemplates.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsTemplates.EOF Then
    Do While Not rsTemplates.EOF
        cmbTemplate.AddItem (rsTemplates("TP_ID") & " - " & rsTemplates("TP_NAME"))
       
        rsTemplates.MoveNext
    Loop
End If
rsTemplates.Close
Set rsTemplates = Nothing

Call FindTemplate
   
'Ticket #20585 - Populate the Security Template combo box
Call Populate_Security_Template

                                              '
If glbLinamar Then
    chkSSecurity(33).Visible = True
    chkUMSecurity(8).Visible = True
    chkUISecurity(8).Visible = True
    chkUSecurity(4).Visible = True
    chkUMSecurity(20).Visible = True
    chkUISecurity(20).Visible = True
    chkUMSecurity(20).Top = chkUMSecurity(14).Top
    chkUISecurity(20).Top = chkUISecurity(14).Top
    chkSecurity(27).Visible = True
    chkMSecurity(27).Visible = True
    chkSecurity(27).Left = 6510     '5670
    chkMSecurity(27).Left = 5535    '4695
ElseIf glbCompSerial = "S/N - 2380W" Then ' For VitalAire Canada Inc. Ticket #26233 Franks 11/20/2014
    chkUISecurity(8).Caption = "Job Classification Tables"
    chkUMSecurity(8).Visible = True
    chkUISecurity(8).Visible = True
ElseIf glbVadim Then 'glbLambton Then
    chkUISecurity(14).Visible = True
    chkUMSecurity(14).Visible = True
    mnu_Cus.Visible = False
ElseIf glbCompSerial = "S/N - 2351W" Then ' For Burlington Tech.
    chkUMSecurity(15).Visible = True
    chkUISecurity(15).Visible = True
ElseIf glbCompSerial = "S/N - 2192W" Then ' county essex

ElseIf glbCompSerial = "S/N - 2407W" Then 'Ticket #18406 - Farmers' Mutual Insurance
    chkPswdLocked.Visible = True
ElseIf glbCompSerial = "S/N - 2382W" Then 'Ticket #23409 - Samuel, Son & Co., Limited - Discipline Audit Table Report
    chkUMSecurity(66).Visible = True
    chkUISecurity(66).Visible = True
    chkUMSecurity(66).Top = chkUMSecurity(14).Top
    chkUISecurity(66).Top = chkUISecurity(14).Top
ElseIf glbCompSerial = "S/N - 2411W" Then 'Ticket #24655 - Wellington-Dufferin-Guelph Public Health - On Call Hours
    chkUMSecurity(67).Visible = True
    chkUISecurity(67).Visible = True
    chkUMSecurity(67).Top = chkUMSecurity(14).Top
    chkUISecurity(67).Top = chkUISecurity(14).Top
    mnu_Attend_Reason.Visible = True
    
    'Ticket #26576 - WDGPHU - Flex Time report - Report in ESS now
    'chkSSecurity(98).Visible = True
ElseIf glbCompSerial = "S/N - 2384W" Then 'Ticket #25746 - Town of St. Marys
    chkUMSecurity(70).Visible = True
    chkUISecurity(70).Visible = True
    chkUISecurity(70).Caption = lStr("Department") & "/" & lStr("G/L") & " Matrix"
    chkUMSecurity(70).Top = chkUMSecurity(14).Top
    chkUISecurity(70).Top = chkUISecurity(14).Top
End If

If Not glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19935
    'These are Samuel's only functions
    chkIESecurity(16).Enabled = False
    chkIESecurity(17).Enabled = False
    chkIESecurity(18).Enabled = False
    chkIESecurity(19).Enabled = False
End If
If glbCompSerial = "S/N - 2382W" Then  'Samuel
    'Ticket #20052 Franks 07/20/2011 - begin
    'chkUMSecurity(62).Top = chkUMSecurity(42).Top
    'chkUISecurity(62).Top = chkUMSecurity(42).Top
    'chkUMSecurity(62).Visible = True
    'chkUISecurity(62).Visible = True
    'chkIESecurity(20).Visible = True
    'chkIESecurity(21).Visible = True
    'chkSSecurity(89).Left = chkSSecurity(69).Left
    'chkSSecurity(89).Visible = True
    'Ticket #20052 Franks 07/20/2011 - end
End If

If glbCompSerial = "S/N - 2425W" Then  'Four Villages - Ticket #21873
    chkSSecurity(90).Visible = True
Else
    chkSSecurity(90).Visible = False
End If

If glbCompSerial = "S/N - 2262W" Then  'County of Wellington - Ticket #22034
    chkSSecurity(91).Top = chkSSecurity(90).Top
    chkSSecurity(91).Visible = True
Else
    chkSSecurity(91).Visible = False
End If

'Ticket #24663 - Showa only
If glbCompSerial = "S/N - 2454W" Then
    chkSSecurity(96).Visible = True
    chkSSecurity(97).Visible = True
Else
    chkSSecurity(96).Visible = False
    chkSSecurity(97).Visible = False
End If

If glbWSIBModule Then  'WSIB Form 7 - Billable Module
    lblHeading(26).Visible = True
    chkUMSecurity(61).Visible = True
    chkUISecurity(61).Visible = True
    chkUMSecurity(63).Visible = True
    chkUISecurity(63).Visible = True
    
    chkUMSecurity(62).Visible = True    'Form 9
    chkUISecurity(62).Visible = True    'Form 9
Else
    lblHeading(26).Visible = False
    chkUMSecurity(61).Visible = False
    chkUISecurity(61).Visible = False
    chkUMSecurity(63).Visible = False
    chkUISecurity(63).Visible = False

    chkUMSecurity(62).Visible = False   'Form 9
    chkUISecurity(62).Visible = False   'Form 9
End If

If Not (glbSQL Or glbOracle) Then
    chkSSecurity(55).Visible = False
End If
    'Security moved out of Essex as controls are available to everyone
    chkUMSecurity(16).Visible = True
    chkUISecurity(16).Visible = True
    chkUMSecurity(17).Visible = True
    chkUISecurity(17).Visible = True
    '***
'If glbCompSerial = "S/N - 2173W" Then
    chkASecurity.Visible = True
'Else
 '   chkASecurity.Visible = False
'End If

'Release 8.1
chkACommentSecurity.Visible = True

'Mostafa Hasheme - show the attendance group code matrix only for Leads and Greenvile
If glbCompSerial = "S/N - 2233W" Then
    chkUISecurity(41).Visible = True
    chkUMSecurity(41).Visible = True
Else
    chkUISecurity(41).Visible = False
    chkUMSecurity(41).Visible = False
End If

'Ticket #16189 - Friesen's Job Files Attachment
If glbCompSerial = "S/N - 2279W" Then
    chkUISecurity(42).Visible = True
    chkUMSecurity(42).Visible = True
    chkUISecurity(43).Visible = True
    chkUMSecurity(43).Visible = True
    chkUISecurity(44).Visible = True
    chkUMSecurity(44).Visible = True
    chkSSecurity(59).Visible = True
    chkSSecurity(60).Visible = True
    
    chkSSecurity(61).Visible = True
    chkSSecurity(62).Visible = True
    chkSSecurity(63).Visible = True
    chkSSecurity(64).Visible = True
    chkSSecurity(65).Visible = True
    chkSSecurity(66).Visible = True
    chkSSecurity(67).Visible = True
    chkSSecurity(68).Visible = True
    chkSSecurity(69).Visible = True
    chkSSecurity(70).Visible = True
    'Ticket #27795 - Friesens Corporation
    chkSSecurity(100).Visible = True
End If

'Ticket #16794 - City of Chatham-Kent - Continuing Education Enhancement
If glbCompSerial = "S/N - 2188W" Then
    chkUISecurity(44).Visible = True
    chkUMSecurity(44).Visible = True
End If

If Not glbWFC Then
    chkMSecurity(24).Visible = False
    chkSecurity(24).Visible = False
'    cmdGrantAllB.Left = 5160
End If
If glbWFC Then
    'Ticket #18566 - begin
    chkUMSecurity(45).Top = chkUMSecurity(42).Top
    chkUISecurity(45).Top = chkUMSecurity(42).Top
    chkUMSecurity(45).Visible = True
    chkUISecurity(45).Visible = True
    chkUMSecurity(46).Top = chkUMSecurity(43).Top
    chkUISecurity(46).Top = chkUMSecurity(43).Top
    chkUMSecurity(46).Visible = True
    chkUISecurity(46).Visible = True
    'Ticket #18566 - end
    
    'Ticket #22009 Franks 05/10/2012
    chkDSecurity.Visible = True
End If

'Ticket #22893 - Security for Year End based on Anniversary Month
'If glbCompSerial = "S/N - 2448W" Then  'For all clients with Security rights
    chkMCSecurity(17).Visible = True
'Else
'    chkMCSecurity(17).Visible = False
'End If

'Ticket #29122 - New Database Setup and Integration Setup securities
If glbVadim Then
    lblHeading(29).Visible = True   'Inquire
    lblHeading(30).Visible = True   'Maintain
    
    chkUMSecurity(72).Visible = True
    chkUISecurity(72).Visible = True
    chkUMSecurity(73).Visible = True
    chkUISecurity(73).Visible = True
Else
    lblHeading(29).Visible = False   'Inquire
    lblHeading(30).Visible = False   'Maintain

    chkUMSecurity(72).Visible = False
    chkUISecurity(72).Visible = False
    chkUMSecurity(73).Visible = False
    chkUISecurity(73).Visible = False
End If

fglbView% = 0 '2
Call displaypanel   '10June99 js
Call setCaption(mnu_Dept)
Call setCaption(chkUISecurity(3))
Call setCaption(chkUISecurity(6))
Call setCaption(chkUISecurity(13))
Call setCaption(chkUISecurity(22))
Call setCaption(chkUISecurity(39))
Call setCaption(chkSSecurity(56)) 'User Defined Report
Call setCaption(chkSecurity(10)) 'Association
Call setCaption(chkSecurity(11)) 'Follow Ups
Call setCaption(chkSecurity(25)) 'Comments
Call setCaption(chkACommentSecurity) 'Add Comments
Call setCaption(chkSecurity(23)) 'Counseling
Call setCaption(chkSecurity(6)) 'Performance
Call setCaption(chkSSecurity(10))  'Salary/Performance Review

'Ticket #22682 - Release 8.0 - Follow Up Code Email Matrix
Call setCaption(chkUISecurity(69)) 'Follow Ups

panWindow.BevelOuter = 0

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

If glbLinamar Or glbWFC Or glbWHSCC Then
    mnu_Cus.Visible = True
End If
If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #21000 Franks 09/26/2011
    mnu_Cus.Visible = True
End If
mnu_Pension.Visible = glbWFC

'Ticket #22682 - Jerry wants Course Admin visible for Serial #9999 and KPAS
If (glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 9999W") Then
    mnu_Sec(6).Visible = True
End If

Me.cmdModify_Click

glbSecUSERID = txtUSERID

'????Ticket #24808 - If the User is Templated based then retrieve the Template Name of this User to retrieve
'Template's Profile instead of User's Profile. If the User's Security is not based on Template or is TEMPLATE then
'retrieve the respective User's record
If cmbSecTemplate = "" Or cmbSecTemplate = "TEMPLATE" Then
    'User is normal user or is Template itself
    glbSecUSERID = txtUSERID
Else
    'User's Profile is based on Template
    glbSecUSERID = cmbSecTemplate
End If

Call Display_Values1

'????Ticket #24808 - Reset this global variable back to User ID
glbSecUSERID = txtUSERID


'Ticket #28540 - Security access with left and right icons on security screens.
cmdScreenRight.Visible = True
cmdScreenLeft.Visible = False

cmdPageLeft(1).Visible = False
cmdPageLeft(2).Visible = False
cmdPageLeft(3).Visible = False
cmdPageLeft(4).Visible = False
cmdPageLeft(5).Visible = False
cmdPageLeft(6).Visible = False
cmdPageLeft(7).Visible = False

cmdPageRight(0).Visible = False
cmdPageRight(1).Visible = False
cmdPageRight(2).Visible = False
cmdPageRight(3).Visible = False
cmdPageRight(4).Visible = False
cmdPageRight(5).Visible = False
cmdPageRight(7).Visible = False
'-----------------------------------------------------------------------------

Exit Sub

SecureLoad_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Load", "HR_SECURE", "Select")
Call RollBack   '10June99 js

End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim VR

Call ChkCBoxChange(cmbSecTemplate)

If ChangeCBox = True Then
        VR = MsgBox("Do you want to save changes?", MB_YESNO)
        If VR = IDYES Then
            Me.cmdOK_Click 'Then Pause (0.5) Else isUpdated = False
        ElseIf VR = IDNO Then
            'Call Me.cmdCancel_Click
        End If
End If
If Not fglbNew Then
    If gsMultiLang = "YES" Then 'whscc
        If Not IsNull(Data1.Recordset("PassWord")) Then
            If txtPWord.Text <> DecryptPasswordMultiLang(Data1.Recordset("PassWord")) Then
                glbConfPass = txtPWord
                Load frmConfPass 'Call ChkPassUnl
                frmConfPass.Show vbModal
            End If
        End If
    Else
        If Not IsNull(Data1.Recordset("PassWord")) Then
            If txtPWord.Text <> DecryptPassword(Data1.Recordset("PassWord")) Then
                glbConfPass = txtPWord
                Load frmConfPass 'Call ChkPassUnl
                frmConfPass.Show vbModal
            End If
        End If
    End If
End If
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Resize()
On Error GoTo Eh
Dim c As Long

If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    panWindow.Height = Me.ScaleHeight - (panEEDESC.Height + SSPanel1.Height + panControls.Height)
    panWindow.Width = Me.ScaleWidth - (scrControl.Width + 200)
    If panWindow.Height >= 7600 Then '6300 Then   '+ 230 Then
        scrControl.Value = 0
        For c = 0 To 7
            panDetails(c).Top = 0
        Next c
        
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        scrControl.Left = Me.ScaleWidth - scrControl.Width
        scrControl.Height = panWindow.Height
    End If


    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 250
    'scrFrame.Height = Me.ScaleHeight - (scrHScroll.Height - 200)  '
    If Me.Width >= 12200 Then '9700 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        If Me.Width < 7000 Then
            scrHScroll.Max = 100
        Else
            scrHScroll.Max = 30
        End If
        scrHScroll.Top = Me.Height - 800
        scrHScroll.Width = Me.Width - 250
    End If
End If

exH:
    Exit Sub
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HR_OCC_HEALTH_SAFETY", "edit/Add")
    Resume exH
End Sub

Private Sub Form_Unload(Cancel As Integer)

MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."

End Sub

Private Sub lblCountry_Change()
    cmbCountry = lblCountry
End Sub

Private Sub lblEEID_Change()
elpEEID.Text = ShowEmpnbr(lblEEID)
If Data1.Recordset.EOF Then Exit Sub
End Sub

Private Sub lblTemplate_Change()
'come back here
'FindTemplate
'cmbTemplate = lblTemplate
Dim intX As Integer
If cmbTemplate.ListCount > 0 Then
    For intX = 1 To cmbTemplate.ListCount - 1
        
             If Trim(Split(cmbTemplate.List(intX))(0)) = lblTemplate.Caption Then
                cmbTemplate.ListIndex = intX
                'cmbTemplate.Refresh
             
                Exit For
             End If
    
        Next
        End If
End Sub

Private Sub mnu_Attend_Reason_Click()
glbSecUSERID = txtUSERID

If Len(lblEEName.Caption) > 0 Then
    glbSecEEName$ = lblEEName.Caption
Else
    glbSecEEName$ = " "
End If

frmSAttendance.Show 1
End Sub

Private Sub mnu_Codes_Click()

glbSecUSERID = txtUSERID

If Len(lblEEName.Caption) > 0 Then
    glbSecEEName$ = lblEEName.Caption
Else
    glbSecEEName$ = " "
End If

frmSCodes.Caption = lStr("Codes Security - ") & lblEEName

frmSCodes.Show 1

End Sub

Private Sub mnu_Comments_Click()

glbSecUSERID = txtUSERID

If Len(lblEEName.Caption) > 0 Then
    glbSecEEName$ = lblEEName.Caption
Else
    glbSecEEName$ = " "
End If

frmSComments.Show 1

End Sub

Private Sub mnu_Cus_Click()

glbSecUSERID = txtUSERID
If Len(lblEEName.Caption) > 0 Then
    glbSecEEName$ = lblEEName.Caption
Else
    glbSecEEName$ = " "
End If
If glbLinamar Then frmSIHRLin.Show
If glbWFC Then frmSIHRWFC.Show
If glbWHSCC Then frmSIHRWHSCC.Show
If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #21000 Franks 09/26/2011
    frmSIHRSamuel.Show
End If
End Sub

Private Sub mnu_CusRpt_Click()
glbSecUSERID = txtUSERID

If Len(lblEEName.Caption) > 0 Then
    glbSecEEName$ = lblEEName.Caption
Else
    glbSecEEName$ = " "
End If
frmSSecRPTs.Show 1
End Sub

Private Sub mnu_Dept_Click()

glbSecUSERID = txtUSERID

If Len(lblEEName.Caption) > 0 Then
    glbSecEEName$ = lblEEName.Caption
Else
    glbSecEEName$ = " "
End If

frmSDept.Show 1

End Sub

Private Sub mnu_DocumentType_Click()
glbSecUSERID = txtUSERID

If Len(lblEEName.Caption) > 0 Then
    glbSecEEName$ = lblEEName.Caption
Else
    glbSecEEName$ = " "
End If

frmSDocType.Show 1
End Sub

Private Sub mnu_Door_Click()
glbSecUSERID = txtUSERID

If Len(lblEEName.Caption) > 0 Then
    glbSecEEName$ = lblEEName.Caption
Else
    glbSecEEName$ = " "
End If
frmLDoors.Show
End Sub

Private Sub mnu_File_ESecure_Click()
    Unload Me
End Sub

Private Sub mnu_File_Exit_Click()
    End
End Sub

Private Sub mnu_FollowUp_Click()
glbSecUSERID = txtUSERID

If Len(lblEEName.Caption) > 0 Then
    glbSecEEName$ = lblEEName.Caption
Else
    glbSecEEName$ = " "
End If

frmSFollowUp.Show 1
End Sub

Private Sub mnu_Pension_Click()

glbSecUSERID = txtUSERID
If Len(lblEEName.Caption) > 0 Then
    glbSecEEName$ = lblEEName.Caption
Else
    glbSecEEName$ = " "
End If
If glbWFC Then frmSIHRWFCPen.Show

End Sub

Private Sub mnu_Requisition_Click()

glbSecUSERID = txtUSERID

If Len(lblEEName.Caption) > 0 Then
    glbSecEEName$ = lblEEName.Caption
Else
    glbSecEEName$ = " "
End If

frmSRequisition.Show 1

End Sub

Private Sub mnu_Sec_Click(Index As Integer)
Dim X%, WExit%

' the pan details array is from 0 to 2 (3 of them)
mnu_Sec(fglbView%).Checked = True
fglbView% = Index

While X% <= 7
    panDetails(X%).Visible = False
    'panDetails(x%).Align = 1
    mnu_Sec(X%).Checked = False
    X% = X% + 1
Wend
    
panDetails(fglbView%).Visible = True

'Ticket #22682 - Jerry wants Course Admin visible for Serial #9999 and KPAS
If fglbView% = 7 And (glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 9999W") Then
    'cmdPageRight(7).Visible = True
    cmdScreenRight.Visible = True
    cmdScreenLeft.Visible = True
ElseIf fglbView% <> 0 Then
    cmdScreenLeft.Visible = True
    If fglbView% = 6 Then
        cmdScreenRight.Visible = False
    Else
        cmdScreenRight.Visible = True
    End If
ElseIf fglbView% = 0 Then
    cmdScreenRight.Visible = True
    cmdScreenLeft.Visible = False
End If

mnu_Sec(fglbView%).Checked = True
End Sub

Private Sub mod_UpdateMode(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

glbOHSEdit% = TF

fUPMode = TF    ' update mode
'Me.vbxTrueGrid.Enabled = FT
'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF
panDetails(0).Enabled = TF      '10June99 js
panDetails(1).Enabled = TF     '
panDetails(2).Enabled = TF     '
panDetails(3).Enabled = TF     '
panDetails(4).Enabled = TF     '
panDetails(5).Enabled = TF
panDetails(6).Enabled = TF
panDetails(7).Enabled = TF

mnu_File.Enabled = TF 'FT
mnu_security.Enabled = TF 'FT
mnu_Dept.Enabled = TF 'FT   ' unenable menu item for depts.

'Ticket #28540 - No Copy Security access when Inquire on Security Master.
cmdCopySecuritys.Enabled = TF

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
txtUSERID.Enabled = TF
elpEEID.Enabled = TF
txtEEName.Enabled = TF

If Len(elpEEID.Text) > 0 Then txtEEName.Enabled = False

If Data1.Recordset.BOF And Data1.Recordset.EOF Then '
'   cmdModify.Enabled = False
'   cmdDelete.Enabled = False
End If                                              '

End Sub

Private Sub elpEEID_Change()
If Len(elpEEID.Text) > 0 Then txtEEName.Enabled = False Else txtEEName.Enabled = True    'And cmdOK.Enabled
If Not UpdateRight Then txtEEName.Enabled = False
End Sub

Private Sub elpEEID_LostFocus()
Dim rsEmp As New ADODB.Recordset
Dim SQLQ
If Len(elpEEID) > 0 Then
    txtEEName = elpEEID.Caption
    'Frank 03/09/2004 Ticket# 5733 for V7.2
    SQLQ = "SELECT ED_COUNTRY FROM HREMP WHERE ED_EMPNBR = " & elpEEID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_COUNTRY")) Then
            cmbCountry = rsEmp("ED_COUNTRY")
            lblCountry.Caption = rsEmp("ED_COUNTRY") 'Ticket #12794
        End If
    End If
    rsEmp.Close
    'Frank 03/09/2004 Ticket# 5733 for V7.2
End If

End Sub

Private Sub scrControl_Change()
Dim c As Long
    For c = 0 To 7
        panDetails(c).Top = 0 - scrControl.Value
    Next c
End Sub

Private Sub scrHScroll_Change()
Dim c As Long
For c = 0 To 7
panDetails(c).Left = 0 - (scrHScroll.Value / 80) * ScaleWidth
Next c
End Sub

Private Sub txtEEName_GotFocus()
    'Hemu - 05/13/2003 Begin
    Call SetPanHelp(ActiveControl)
    'Hemu - 05/13/2003 End
End Sub

Private Sub txtEmpNumber_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtSecPwd_Change()
'txtPassword.Text = DecryptPassword(txtSecPassword.Text)
End Sub

Private Sub txtExpireDays_GotFocus()
    OldExpireDays = 0
    If IsNumeric(txtExpireDays.Text) Then
        OldExpireDays = txtExpireDays.Text
    End If
End Sub

Private Sub txtExpireDays_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo invalidDate_Err
    If IsNumeric(txtExpireDays.Text) Then
        If Not Val(txtExpireDays.Text) = Val(OldExpireDays) Then
            If IsDate(LastExpireDate.Text) Then
                ExpireDate.Text = DateAdd("D", txtExpireDays.Text, LastExpireDate.Text)
                OldExpireDays = txtExpireDays.Text
            Else
                LastExpireDate.Text = Date
                ExpireDate.Text = DateAdd("D", txtExpireDays.Text, LastExpireDate.Text)
                OldExpireDays = txtExpireDays.Text
            End If
        End If
    End If
Exit Sub
invalidDate_Err:
    MsgBox "Exceeding the maximum date value. Reverting the 'Expiration Days' change.", vbExclamation, "Invalid Expiration Date Computation"
    txtExpireDays.Text = OldExpireDays
    ExpireDate.Text = DateAdd("D", txtExpireDays.Text, LastExpireDate.Text)
    Resume Next
End Sub

Private Sub txtPWord_Change()
If gsMultiLang = "YES" Then 'whscc
    txtSecPassword.Text = EncryptPasswordMultiLang(txtPWord.Text)
Else
    txtSecPassword.Text = EncryptPassword(txtPWord.Text)
End If

End Sub

Private Sub txtPWord_GotFocus()
    'Hemu - 05/13/2003 Begin
    Call SetPanHelp(ActiveControl)
    'Hemu - 05/13/2003 End
End Sub

Private Sub txtSecPassword_Change()
If gsMultiLang = "YES" Then 'whscc
    txtPWord.Text = DecryptPasswordMultiLang(txtSecPassword.Text)
Else
    txtPWord.Text = DecryptPassword(txtSecPassword.Text)
End If
End Sub

Private Sub txtSecTemplate_Change()
    If cmbSecTemplate.ListCount > 0 Then
        For X = 1 To cmbSecTemplate.ListCount - 1
            If cmbSecTemplate.List(X) = txtSecTemplate.Text Then
                cmbSecTemplate.ListIndex = X
                Exit For
            Else
                cmbSecTemplate.ListIndex = 0
            End If
            'cmbSecTemplate.Text = txtSecTemplate.Text
        Next
    End If
End Sub

Private Sub txtUserID_GotFocus()
    'Hemu - 05/13/2003 Begin
    Call SetPanHelp(ActiveControl)
    'Hemu - 05/13/2003 End
End Sub

Private Sub txtUsername_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
''Dim Chan
'Dim VR
'Call ChkCBoxChange
'If ChangeCBox = True Then
'        VR = MsgBox("Do you want to save changes?", MB_YESNO)
'        If VR = IDYES Then
'            Me.cmdOK_Click
'        ElseIf VR = IDNO Then
'            Call Me.cmdCancel_Click
'        End If
'End If
'
'If txtPWord.Text <> DecryptPassword(Data1.Recordset("PassWord")) Then
'    glbConfPass = txtPWord
'    Load frmConfPass
'    frmConfPass.Show vbModal
'End If
'If IsNull(Data1.Recordset("COUNTRY")) Then
'    cmbCountry = ""
'Else
'    cmbCountry = Data1.Recordset("COUNTRY")
'End If

'Dim VR
''Call ChkCBoxChange
''If ChangeCBox = True Then
''        VR = MsgBox("Do you want to save changes?", MB_YESNO)
''        If VR = IDYES Then
''            Me.cmdOK_Click 'Then Pause (0.5) Else isUpdated = False
''        ElseIf VR = IDNO Then
''            'Call Me.cmdCancel_Click
''        End If
''End If
If Not fglbNew Then
    If gsMultiLang = "YES" Then 'whscc
        If Not IsNull(Data1.Recordset("PassWord")) Then
            If txtPWord.Text <> DecryptPasswordMultiLang(Data1.Recordset("PassWord")) Then
                'VR = MsgBox("User's Password has changed. Do you want to save changes?", vbExclamation + MB_YESNO)
                'If VR = IDYES Then
                '    glbConfPass = txtPWord
                '    Load frmConfPass 'Call ChkPassUnl
                '    frmConfPass.Show vbModal
                'Else
                    Call Me.cmdCancel_Click
                'End If
            End If
        End If
    Else
        If Not IsNull(Data1.Recordset("PassWord")) Then
            If txtPWord.Text <> DecryptPassword(Data1.Recordset("PassWord")) Then
                'VR = MsgBox("User's Password has changed. Do you want to save changes?", vbExclamation + MB_YESNO)
                'If VR = IDYES Then
                '    glbConfPass = txtPWord
                '    Load frmConfPass 'Call ChkPassUnl
                '    frmConfPass.Show vbModal
                'Else
                    Call Me.cmdCancel_Click
                'End If
            End If
        End If
    End If
End If

End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
    Dim SQLQ As String
    
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    SQLQ = "SELECT *, "
    If glbLinamar Then
        SQLQ = SQLQ & " CASE WHEN EMPNBR IS NOT NULL AND LEN(EMPNBR)>2 "
        SQLQ = SQLQ & " THEN RIGHT(EMPNBR,3)+'-'+"
        SQLQ = SQLQ & " LEFT(EMPNBR,LEN(EMPNBR)-3) "
        SQLQ = SQLQ & " ELSE STR(EMPNBR) END "
        SQLQ = SQLQ & " AS SEMPNBR "
    Else
        SQLQ = "SELECT * "
        vbxTrueGrid.Columns(1).DataField = "EMPNBR"
    End If
    SQLQ = SQLQ & " from HR_SECURE_BASIC "
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
'    End If
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim X%

On Error GoTo Tab1_Err

If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
    '????Ticket #24808 - If the User is Template based then retrieve the Template Name of this User to retrieve
    'Template's Profile instead of User's Profile. If the User's Security is not based on Template or is TEMPLATE then
    'retrieve the respective User's record
    If cmbSecTemplate = "" Or cmbSecTemplate = "TEMPLATE" Then
        'User is normal user or is Template itself
        glbSecUSERID = txtUSERID
    Else
        'User's Profile is based on Template
        glbSecUSERID = cmbSecTemplate
    End If
    
    Call Display_Values1
    
    '????Ticket #24808 - Reset this global variable back to User ID
    glbSecUSERID = txtUSERID
    
End If


'If Not ChkPass Then
'        glbSecUSERID = IDBack
'        Data1.Recordset.Find ("USERID = '" & IDBack & "'")
'        Call Display_Values1
'End If


Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_OCC_HEALTH_SAFETY", "Add")
Call RollBack   '10June99 js

End Sub

Private Function RollBack()
On Error GoTo rr
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
rr:
End Function

Private Function displaypanel() '10June99 js - displays Panel selected

X% = 0
While X% <= 7
    panDetails(X%).Visible = False
    panDetails(X%).Top = 0
    panDetails(X%).Left = 0
    panDetails(X%).Width = 12100
    panDetails(X%).Height = 7780 ' 7680 '5475
    panDetails(X%).BorderStyle = 0
    mnu_Sec(X%).Checked = False
    X% = X% + 1
Wend
panDetails(2).Height = 6000
panDetails(fglbView%).Visible = True
mnu_Sec(fglbView%).Checked = True
panDetails(3).Height = 8655 '7575 '7300 ' 6975
End Function

Private Sub Display_Values1()
Dim rsSR As New ADODB.Recordset
Dim SQLQ

'Ticket #20585 - Repopulate Security Template dropdown list
Call Populate_Security_Template

'????Ticket #24808 - This could be Template's Profile if the User's based Template based or else User's Profile if not template based or TEMPLATE itself.
SQLQ = "SELECT " & Field_SQL("FUNCTION") & ", ACCESSABLE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND Maintainable=0"
' dkostka - 09/28/2001 - Changed from adOpenStatic to adOpenForwardOnly to improve speed.
rsSR.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

Call ResetAll

Do Until rsSR.EOF
    If UCase(rsSR("FUNCTION")) = UCase("Company_Update") Then chkUMSecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Company_Inquiry") Then chkUISecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Security_Update") Then chkUMSecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Security_Inquiry") Then chkUISecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Master_Table_Update") Then chkUMSecurity(2) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Master_Table_Inquiry") Then chkUISecurity(2) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Departments_Update") Then chkUMSecurity(3) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Department_Inquiry") Then chkUISecurity(3) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Audit_Update") Then chkUMSecurity(4) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Audit_Inquiry") Then chkUISecurity(4) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("EmploymentEQT_Update") Then chkUMSecurity(5) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("EmploymentEQT_Inquiry") Then chkUISecurity(5) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Divisions_Update") Then chkUMSecurity(6) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Divisions_Inquiry") Then chkUISecurity(6) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Ledgers_Update") Then chkUMSecurity(7) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Ledgers_Inquiry") Then chkUISecurity(7) = rsSR("ACCESSABLE")
    
    If glbLinamar Then
        If UCase(rsSR("FUNCTION")) = UCase("DoorAccess_Update") Then chkUMSecurity(8) = rsSR("ACCESSABLE")
        If UCase(rsSR("FUNCTION")) = UCase("DoorAccess_Inquiry") Then chkUISecurity(8) = rsSR("ACCESSABLE")
        If UCase(rsSR("FUNCTION")) = UCase("ProductLine_Operation_Update") Then chkUMSecurity(20) = rsSR("ACCESSABLE")
        If UCase(rsSR("FUNCTION")) = UCase("ProductLine_Operation_Inquiry") Then chkUISecurity(20) = rsSR("ACCESSABLE")
    End If
    If glbCompSerial = "S/N - 2380W" Then ' For VitalAire Canada Inc. Ticket #26233 Franks 11/20/2014
        If UCase(rsSR("FUNCTION")) = UCase("DoorAccess_Update") Then chkUMSecurity(8) = rsSR("ACCESSABLE")
        If UCase(rsSR("FUNCTION")) = UCase("DoorAccess_Inquiry") Then chkUISecurity(8) = rsSR("ACCESSABLE")
    End If
    
    'Ticket #23409 - Samuel, Son & Co., Limited - Discipline Audit Table Report
    If glbCompSerial = "S/N - 2382W" Then
        If UCase(rsSR("FUNCTION")) = UCase("CounselAudit_Update") Then chkUMSecurity(66) = rsSR("ACCESSABLE")
        If UCase(rsSR("FUNCTION")) = UCase("CounselAudit_Inquiry") Then chkUISecurity(66) = rsSR("ACCESSABLE")
    End If
    
    'Ticket #24655 - Wellington-Dufferin-Guelph Public Health - On Call Hours
    If glbCompSerial = "S/N - 2411W" Then
        If UCase(rsSR("FUNCTION")) = UCase("On_Call_Hours_Update") Then chkUMSecurity(67) = rsSR("ACCESSABLE")
        If UCase(rsSR("FUNCTION")) = UCase("On_Call_Hours_Inquiry") Then chkUISecurity(67) = rsSR("ACCESSABLE")
    End If
    
    If UCase(rsSR("FUNCTION")) = UCase("CustomReport_Update") Then chkUMSecurity(9) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CustomReport_Inquiry") Then chkUISecurity(9) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Holiday_Update") Then chkUMSecurity(10) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Holiday_Inquiry") Then chkUISecurity(10) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("New_Hire_Update") Then chkUMSecurity(11) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("New_Hire_Inquiry") Then chkUISecurity(11) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Label_Update") Then chkUMSecurity(12) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Label_Inquiry") Then chkUISecurity(12) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Sal_Distribute_Update") Then chkUMSecurity(13) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Sal_Distribute_Inquiry") Then chkUISecurity(13) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Pay_Period_Update") Then chkUMSecurity(19) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Pay_Period_Inquiry") Then chkUISecurity(19) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Email_Setup_Update") Then chkUMSecurity(18) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Email_Setup_Inquiry") Then chkUISecurity(18) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Payroll_Category_Update") Then chkUMSecurity(14) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Payroll_Category_Inquiry") Then chkUISecurity(14) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Charge_Code_Update") Then chkUMSecurity(15) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Charge_Code_Inquiry") Then chkUISecurity(15) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Project_Code_Update") Then chkUMSecurity(16) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Project_Code_Inquiry") Then chkUISecurity(16) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Machine_Update") Then chkUMSecurity(17) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Machine_Inquiry") Then chkUISecurity(17) = rsSR("ACCESSABLE")
    
    'Ticket #25746 - Town of St. Marys
    If UCase(rsSR("FUNCTION")) = UCase("DeptGL_Matrix_Update") Then chkUMSecurity(70) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("DeptGL_Matrix_Inquiry") Then chkUISecurity(70) = rsSR("ACCESSABLE")
    
    '7.6
    If UCase(rsSR("FUNCTION")) = UCase("EMP_FLAGS_Update") Then chkUMSecurity(22) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("EMP_FLAGS_Inquiry") Then chkUISecurity(22) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("EMP_HISTORY_Update") Then chkUMSecurity(23) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("EMP_HISTORY_Inquiry") Then chkUISecurity(23) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("GL_DIST_Update") Then chkUMSecurity(24) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("GL_DIST_Inquiry") Then chkUISecurity(24) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("EMP_LANG_Update") Then chkUMSecurity(25) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("EMP_LANG_Inquiry") Then chkUISecurity(25) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("EMP_SUCCESSION_Update") Then chkUMSecurity(26) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("EMP_SUCCESSION_Inquiry") Then chkUISecurity(26) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Work_Schedule_Update") Then chkUMSecurity(60) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Work_Schedule_Inquiry") Then chkUISecurity(60) = rsSR("ACCESSABLE")
        
    'Attendance Code Matrix
    If UCase(rsSR("FUNCTION")) = UCase("AttendCode_Matrix_Update") Then chkUMSecurity(59) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("AttendCode_Matrix_Inquiry") Then chkUISecurity(59) = rsSR("ACCESSABLE")
    
    'Ticket #22682 - Release 8.0 - Follow Up Code Email Matrix
    If UCase(rsSR("FUNCTION")) = UCase("FollowUpCodeEmail_Matrix_Update") Then chkUMSecurity(69) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("FollowUpCodeEmail_Matrix_Inquiry") Then chkUISecurity(69) = rsSR("ACCESSABLE")

    'Ticket #25922 - OHRS Reporting for CHC
    If UCase(rsSR("FUNCTION")) = UCase("OHRS_Department_Update") Then chkUMSecurity(71) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("OHRS_Department_Inquiry") Then chkUISecurity(71) = rsSR("ACCESSABLE")
    
    'ADP Data
    If UCase(rsSR("FUNCTION")) = UCase("ADP_Data_Update") Then chkUMSecurity(36) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ADP_Data_Inquiry") Then chkUISecurity(36) = rsSR("ACCESSABLE")
    'Sam Added for ESS.NET 07/27/2006 Ticket # 11403
    'If UCase(rsSR("FUNCTION")) = UCase("Archive_VacTimeoff_Update") Then chkUMSecurity(37) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("Archive_VacTimeoff_Inquiry") Then chkUISecurity(37) = rsSR("ACCESSABLE")
    'ends
    
    If UCase(rsSR("FUNCTION")) = UCase("UserDefineTbl_Update") Then chkUMSecurity(39) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("UserDefineTbl_Inquiry") Then chkUISecurity(39) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("PayrollTrans_Update") Then chkUMSecurity(40) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("PayrollTrans_Inquiry") Then chkUISecurity(40) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Emergency_Contacts_Update") Then chkUMSecurity(35) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Emergency_Contacts_Inquiry") Then chkUISecurity(35) = rsSR("ACCESSABLE")
    
    'Course Code Master
    If UCase(rsSR("FUNCTION")) = UCase("CourseCodeMaster_Update") Then chkUMSecurity(38) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CourseCodeMaster_Inquiry") Then chkUISecurity(38) = rsSR("ACCESSABLE")
        
    'Course Admin - Begin
    If UCase(rsSR("FUNCTION")) = UCase("CA_Organization_Update") Then chkUMSecurity(27) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Organization_Inquiry") Then chkUISecurity(27) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Catalog_Update") Then chkUMSecurity(28) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Catalog_Inquiry") Then chkUISecurity(28) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_TrainingLocation_Update") Then chkUMSecurity(29) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_TrainingLocation_Inquiry") Then chkUISecurity(29) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Scheduling_Update") Then chkUMSecurity(30) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Scheduling_Inquiry") Then chkUISecurity(30) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_EnrollRequest_Update") Then chkUMSecurity(31) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_EnrollRequest_Inquiry") Then chkUISecurity(31) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_EnrollApproval_Update") Then chkUMSecurity(32) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_EnrollApproval_Inquiry") Then chkUISecurity(32) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Enrollment_Update") Then chkUMSecurity(33) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Enrollment_Inquiry") Then chkUISecurity(33) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_WaitingList_Update") Then chkUMSecurity(34) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_WaitingList_Inquiry") Then chkUISecurity(34) = rsSR("ACCESSABLE")
    'Course Admin - End

'Moved to 'Display_Values1_Utilities_WebModules' because the compile of this screen was giving 'Procedure too large'
'-------------------------------------------------------------------------------------------------------------------
'    If UCase(rsSR("FUNCTION")) = UCase("Province") Then chkUSecurity(0) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Entitle") Then chkUSecurity(1) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Compress_Fix") Then chkUSecurity(2) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Matrix") Then chkUSecurity(3) = rsSR("ACCESSABLE")
'    If glbLinamar Then
'        If UCase(rsSR("FUNCTION")) = UCase("DoorName") Then chkUSecurity(4) = rsSR("ACCESSABLE")
'        If UCase(rsSR("FUNCTION")) = UCase("Summarize_Attendance") Then chkUSecurity(5) = rsSR("ACCESSABLE")
'    End If
'    If UCase(rsSR("FUNCTION")) = UCase("TimeSheetPrority") Then chkUSecurity(6) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("TimeSheetUserPrority") Then chkUSecurity(7) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("DeleteTimeSheetFile") Then chkUSecurity(8) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("ArchiveTimeSheetFile") Then chkUSecurity(25) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("ApproveTimeSheetFile") Then chkUSecurity(26) = rsSR("ACCESSABLE")
'
'    'If UCase(rsSR("FUNCTION")) = UCase("EssCompTime") Then chkUSecurity(26) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("TSOverrideEmpSecurity") Then chkUSecurity(14) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("DeleteApprovedTimeSheet") Then chkUSecurity(15) = rsSR("ACCESSABLE")
'
'    '7.9 Enhancement
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_Time_Req") Then chkUSecurity(18) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_Vacation_Req") Then chkUSecurity(19) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_Request_Approval") Then chkUSecurity(20) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_Rpt_Request_Approval") Then chkUSecurity(21) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_Rpt_Print_Archive") Then chkUSecurity(22) = rsSR("ACCESSABLE")
'    'If UCase(rsSR("FUNCTION")) = UCase("ESS_Archive_Req") Then chkUSecurity(23) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Archive_VacTimeoff_Update") Then chkUSecurity(23) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_MassDelete_Req") Then chkUSecurity(24) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_ShowAllRequests") Then chkUSecurity(27) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("TS_PUNCHINOUT") Then chkUSecurity(28) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_REQ_OTHER_SUPER") Then chkUSecurity(29) = rsSR("ACCESSABLE")
'
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_DEL_APPR_TIME_REQ") Then chkUSecurity(30) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_DEL_APPR_VAC_REQ") Then chkUSecurity(31) = rsSR("ACCESSABLE")
'
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_CANCEL_REQ") Then chkUSecurity(32) = rsSR("ACCESSABLE")
'
'
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_TS_LIST_RA_ONLY") Then chkUSecurity(33) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_DELETE_FUTURE_VAC_REQ") Then chkUSecurity(34) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_DELETE_FUTURE_TIME_REQ") Then chkUSecurity(35) = rsSR("ACCESSABLE")
'
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_HTML_CALENDAR") Then chkUSecurity(36) = rsSR("ACCESSABLE")
'    'Ticket #23536 - Dashboard ON/OFF
'    If UCase(rsSR("FUNCTION")) = UCase("ESS_DASHBOARDS") Then chkUSecurity(37) = rsSR("ACCESSABLE")
'
'    If UCase(rsSR("FUNCTION")) = UCase("CompanyPreference") Then chkUSecurity(9) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("EmpFlagsSetup") Then chkUSecurity(10) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("MultiDataSourceSetup") Then chkUSecurity(11) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("HelpDescSetup") Then chkUSecurity(12) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("BenefitGroupSetup") Then chkUSecurity(13) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("ChangeYourPassword") Then chkUSecurity(16) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("ITAdmin") Then chkUSecurity(17) = rsSR("ACCESSABLE")
'
'    If UCase(rsSR("FUNCTION")) = UCase("Import_Attendance") Then chkIESecurity(0) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Export_Attendance") Then chkIESecurity(1) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Import_Benefits") Then chkIESecurity(2) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Export_Benefits") Then chkIESecurity(3) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Import_Employee") Then chkIESecurity(4) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Export_Employee") Then chkIESecurity(5) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Import_Salaries") Then chkIESecurity(6) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Export_Salaries") Then chkIESecurity(7) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Import_Table") Then chkIESecurity(8) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Export_Table") Then chkIESecurity(9) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Import_YTD") Then chkIESecurity(10) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Export_YTD") Then chkIESecurity(11) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Import_PayrollTrans") Then chkIESecurity(12) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Export_PayrollTrans") Then chkIESecurity(13) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Import_ContEdu") Then chkIESecurity(14) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Export_ContEdu") Then chkIESecurity(15) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Import_PerfReview") Then chkIESecurity(16) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Export_PerfReview") Then chkIESecurity(17) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Import_EmploymentEquity") Then chkIESecurity(18) = rsSR("ACCESSABLE")
'    If UCase(rsSR("FUNCTION")) = UCase("Export_EmploymentEquity") Then chkIESecurity(19) = rsSR("ACCESSABLE")
'-------------------------------------------------------------------------------------------------------------------

    'Not use - Samuel Ticket #21000 Franks 09/26/2011 - move to Custom Security
    'If UCase(rsSR("FUNCTION")) = UCase("Import_Profit_Sharing") Then chkIESecurity(20) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("Export_Profit_Sharing") Then chkIESecurity(21) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Basic_Update") Then chkMSecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Basic_Inquiry") Then chkSecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Banking_Update") Then chkMSecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Banking_Inquiry") Then chkSecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Dependents_Update") Then chkMSecurity(2) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Dependents_Inquiry") Then chkSecurity(2) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Skills_Update") Then chkMSecurity(3) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Skills_Inquiry") Then chkSecurity(3) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Formal_Education_Update") Then chkMSecurity(4) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Formal_Education_Inquiry") Then chkSecurity(4) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Salary_Update") Then chkMSecurity(5) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Salary_Inquiry") Then chkSecurity(5) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Performance_Update") Then chkMSecurity(6) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Performance_Inquiry") Then chkSecurity(6) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Position_Update") Then chkMSecurity(7) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Position_Inquiry") Then chkSecurity(7) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Benefits_Update") Then chkMSecurity(8) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Benefits_Inquiry") Then chkSecurity(8) = rsSR("ACCESSABLE")
        
    '7.9 Enhancement
    If UCase(rsSR("FUNCTION")) = UCase("Beneficiary_Update") Then chkMSecurity(29) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Beneficiary_Inquiry") Then chkSecurity(29) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Entitlements_Update") Then chkMSecurity(9) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Entitlements_Inquiry") Then chkSecurity(9) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Associations_Update") Then chkMSecurity(10) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Associations_Inquiry") Then chkSecurity(10) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Follow_Ups_Update") Then chkMSecurity(11) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Follow_Ups_Inquiry") Then chkSecurity(11) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Health_Safety_Update") Then chkMSecurity(12) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Health_Safety_Inquiry") Then chkSecurity(12) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Attendance_Update") Then chkMSecurity(13) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Attendance_Inquiry") Then chkSecurity(13) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Other_Entitlements_Update") Then chkMSecurity(14) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Other_Entitlements_Inquiry") Then chkSecurity(14) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Other_Earnings_Update") Then chkMSecurity(15) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Other_Earnings_Inquiry") Then chkSecurity(15) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Terminations_Update") Then chkMSecurity(16) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Termination_Inquiry") Then chkSecurity(16) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Job_Classes_Update") Then chkMSecurity(17) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Job_Classes_Inquiry") Then chkSecurity(17) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Education_Seminars_Update") Then chkMSecurity(18) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Education_Seminars_Inquiry") Then chkSecurity(18) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Job_Skills_Update") Then chkMSecurity(19) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Job_Skills_Inquiry") Then chkSecurity(19) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Job_Eval_Update") Then chkMSecurity(20) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Job_Eval_Inquiry") Then chkSecurity(20) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Job_Master_Update") Then chkMSecurity(21) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Job_Master_Inquiry") Then chkSecurity(21) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Hrly_Entitlements_Update") Then chkMSecurity(22) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Hrly_Entitlements_Inquiry") Then chkSecurity(22) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Show_SIN_SSN") Then chkEESIN = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Show_DOB") Then chkEEDOB = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Show_ADDRESS") Then chkEEADDRESS = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Show_MARITAL") Then chkEEMarital = rsSR("ACCESSABLE")
    'Moved to HR_SECURE_BASIC
    'If glbCompSerial = "S/N - 2407W" Then 'Ticket #18406 - Farmers' Mutual Insurance
    '    If UCase(rsSR("FUNCTION")) = UCase("Lock_Password") Then chkPswdLocked = rsSR("ACCESSABLE")
    'End If
    
    'tkt#10423 Jerry said make it available for everyone
    'If glbCompSerial = "S/N - 2173W" Then
        If UCase(rsSR("FUNCTION")) = UCase("Add_Attendance") Then chkASecurity = rsSR("ACCESSABLE")
    'End If
        
    'Ticket #22682 - Release 8.0
    If UCase(rsSR("FUNCTION")) = UCase("Add_NewHire") Then chkNHireSecurity = rsSR("ACCESSABLE")
    
    'Release 8.1
    If UCase(rsSR("FUNCTION")) = UCase("Add_Comments") Then chkACommentSecurity = rsSR("ACCESSABLE")
        
    'Ticket #23923 - Release 8.0 - View Own
    If UCase(rsSR("FUNCTION")) = UCase("ScsPlan_ViewOwn") Then chkViewOwnSuccPlan = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Comments_ViewOwn") Then chkViewOwnComm = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Counsel_ViewOwn") Then chkViewOwnCounsel = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("FollUp_ViewOwn") Then chkViewOwnFollUp = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("OthInfo_ViewOwn") Then chkViewOwnOthInfo = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("EmpFlags_ViewOwn") Then chkViewOwnEmpFlags = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("EmpHis_ViewOwn") Then chkViewOwnEmpHis = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("GLDist_ViewOwn") Then chkViewOwnGLDist = rsSR("ACCESSABLE")
        
    'Ticket #28635 - Add View Own security
    If UCase(rsSR("FUNCTION")) = UCase("Perform_ViewOwn") Then chkViewOwnPerform = rsSR("ACCESSABLE")
        
    'Ticket #22009 Franks 05/10/2012
    If UCase(rsSR("FUNCTION")) = UCase("Del_Dependents") Then chkDSecurity = rsSR("ACCESSABLE")
    
    ' dkostka - 09/25/2001 - Added security for Counselling screen.
    If UCase(rsSR("FUNCTION")) = UCase("Counselling_Update") Then chkMSecurity(23) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Counselling_Inquiry") Then chkSecurity(23) = rsSR("ACCESSABLE")
    ' Frank - 11/2/2001 - Added security for Salary Grids.
    If glbWFC Then
        If UCase(rsSR("FUNCTION")) = UCase("SalaryGrids_Update") Then
            chkMSecurity(24) = rsSR("ACCESSABLE")
        End If
        If UCase(rsSR("FUNCTION")) = UCase("SalaryGrids_Inquiry") Then
            chkSecurity(24) = rsSR("ACCESSABLE")
        End If
    End If
    If UCase(rsSR("FUNCTION")) = UCase("Comments_Update") Then chkMSecurity(25) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Comments_Inquiry") Then chkSecurity(25) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("OtherInformation_Update") Then chkMSecurity(26) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("OtherInformation_Inquiry") Then chkSecurity(26) = rsSR("ACCESSABLE")
    If glbLinamar Then
        If UCase(rsSR("FUNCTION")) = UCase("LinamarSkills_Update") Then chkMSecurity(27) = rsSR("ACCESSABLE")
        If UCase(rsSR("FUNCTION")) = UCase("LinamarSkills_Inquiry") Then chkSecurity(27) = rsSR("ACCESSABLE")
    End If
    If UCase(rsSR("FUNCTION")) = UCase("Attendance_History_Update") Then chkMSecurity(28) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Attendance_History_Inquiry") Then chkSecurity(28) = rsSR("ACCESSABLE")
    ' Dijana - May 2002 - Add security for Applicant
    
    If UCase(rsSR("FUNCTION")) = UCase("App_Basic_Update") Then chkAMSecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Basic_Inquiry") Then chkAISecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Skills_Update") Then chkAMSecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Skills_Inquiry") Then chkAISecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Formal_Education_Update") Then chkAMSecurity(2) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Formal_Education_Inquiry") Then chkAISecurity(2) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Education_Seminars_Update") Then chkAMSecurity(3) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Education_Seminars_Inquiry") Then chkAISecurity(3) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Interview_Update") Then chkAMSecurity(4) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Interview_Inquiry") Then chkAISecurity(4) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Associations_Update") Then chkAMSecurity(5) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Associations_Inquiry") Then chkAISecurity(5) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_References_Update") Then chkAMSecurity(6) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_References_Inquiry") Then chkAISecurity(6) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Follow_Ups_Update") Then chkAMSecurity(7) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Follow_Ups_Inquiry") Then chkAISecurity(7) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Requisition_Update") Then chkAMSecurity(8) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Requisition_Inquiry") Then chkAISecurity(8) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Recruitment_Update") Then chkAMSecurity(9) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Recruitment_Inquiry") Then chkAISecurity(9) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Employment_Update") Then chkAMSecurity(10) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_Employment_Inquiry") Then chkAISecurity(10) = rsSR("ACCESSABLE")
    'Ticket #30508 - Applicant Tracking Enhancement
    If UCase(rsSR("FUNCTION")) = UCase("App_LetterPosType_Update") Then chkAMSecurity(11) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_LetterPosType_Inquiry") Then chkAISecurity(11) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_FormWorkflow_Update") Then chkAMSecurity(12) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_FormWorkflow_Inquiry") Then chkAISecurity(12) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_FormDefaults_Update") Then chkAMSecurity(13) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("App_FormDefaults_Inquiry") Then chkAISecurity(13) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Report_Age") Then chkSSecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Compensatory_Time") Then chkSSecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Cost_Of_Employment") Then chkSSecurity(2) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Emergecy_Contacts") Then chkSSecurity(3) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Employee_Labels") Then chkSSecurity(4) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Job_List") Then chkSSecurity(5) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Profiles") Then chkSSecurity(6) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Entitlements") Then chkSSecurity(7) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Follow_Ups") Then chkSSecurity(8) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Home_Address") Then chkSSecurity(9) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Salary_Performance") Then chkSSecurity(10) = rsSR("ACCESSABLE")
    'Ticket #27795 - Friesens Corporation
    If UCase(rsSR("FUNCTION")) = UCase("Report_Staff_Profile") Then chkSSecurity(100) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Seniority") Then chkSSecurity(11) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Telephone_Extensions") Then chkSSecurity(12) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Associations") Then chkSSecurity(13) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Master_Attendance") Then chkSSecurity(14) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Bonus_Attendance") Then chkSSecurity(83) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Calendar_Attendance") Then chkSSecurity(84) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Costed_Attendance") Then chkSSecurity(85) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Master_Benefits") Then chkSSecurity(15) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Master_Division") Then chkSSecurity(16) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Master_Termination") Then chkSSecurity(17) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Master_Formal_Education") Then chkSSecurity(18) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Master_Job") Then chkSSecurity(19) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Master_Passwords") Then chkSSecurity(20) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Master_Salaries") Then chkSSecurity(21) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Master_Edu_Seminars") Then chkSSecurity(22) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Master_Table_Codes") Then chkSSecurity(23) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Heatlh_Safety") Then chkSSecurity(24) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Master_DolEnt") Then chkSSecurity(25) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Hourly_Entitlements") Then chkSSecurity(26) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Master_OtherEarn") Then chkSSecurity(27) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("PayEQT_Inquiry") Then chkSSecurity(28) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Dependents") Then chkSSecurity(29) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Skills") Then chkSSecurity(30) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Languages") Then chkSSecurity(31) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Employee_Turnover") Then chkSSecurity(32) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Counselling") Then chkSSecurity(34) = rsSR("ACCESSABLE")
    'Release 8.1
    If UCase(rsSR("FUNCTION")) = UCase("Report_DocumentType") Then chkSSecurity(99) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Emergency_Leave") Then chkSSecurity(35) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_External_Hire") Then chkSSecurity(48) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Internal_Hire") Then chkSSecurity(49) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Key_Workforce") Then chkSSecurity(50) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Manpower_Plan") Then chkSSecurity(51) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Staff_Management") Then chkSSecurity(52) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_WC_Time") Then chkSSecurity(53) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_WC_Work") Then chkSSecurity(54) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Paid_Sick") Then chkSSecurity(55) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_User_Defined_Table") Then chkSSecurity(56) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Future_Entitlement") Then chkSSecurity(57) = rsSR("ACCESSABLE")
    'Overtime
    If UCase(rsSR("FUNCTION")) = UCase("Report_Overtime_Bank") Then chkSSecurity(46) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Overtime_Lost_Hours") Then chkSSecurity(47) = rsSR("ACCESSABLE")
    'More
    If UCase(rsSR("FUNCTION")) = UCase("Report_Employee_Flags") Then chkSSecurity(58) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Temp_CrossTraining") Then chkSSecurity(59) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Required_Course_Hist") Then chkSSecurity(60) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Report_Email_Address") Then chkSSecurity(71) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_LOA") Then chkSSecurity(73) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_POE") Then chkSSecurity(74) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_SINSSN") Then chkSSecurity(75) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Succession") Then chkSSecurity(76) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Gap_Analysis") Then chkSSecurity(72) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Report_GL_Distribution") Then chkSSecurity(86) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Report_Attendance_Hist") Then chkSSecurity(77) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Comments") Then chkSSecurity(78) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Employee_Hist") Then chkSSecurity(79) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Payroll_Transactions") Then chkSSecurity(80) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_AffirmAction") Then chkSSecurity(81) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Report_WorkSchedule") Then chkSSecurity(82) = rsSR("ACCESSABLE")
        
    If UCase(rsSR("FUNCTION")) = UCase("Report_Applicant_Profile") Then chkSSecurity(87) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Applicant_Education") Then chkSSecurity(88) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Training_Plan") Then chkSSecurity(89) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("Report_Profit_Sharing") Then chkSSecurity(89) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Report_AttWrkSch_Descrepancy") Then chkSSecurity(90) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Environmental_Serv") Then chkSSecurity(91) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_ESSReq_TransAudit") Then chkSSecurity(92) = rsSR("ACCESSABLE")
    
    'Release 8.0 - Ticket #22682
    If UCase(rsSR("FUNCTION")) = UCase("Report_Employee_Dates") Then chkSSecurity(93) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Length_of_Service") Then chkSSecurity(94) = rsSR("ACCESSABLE")
    
    'Ticket #24663
    If UCase(rsSR("FUNCTION")) = UCase("Form_Attendance_SignIn") Then chkSSecurity(95) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Form_ATT_Discipline") Then chkSSecurity(96) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Form_COC_Discipline") Then chkSSecurity(97) = rsSR("ACCESSABLE")
        
    'Ticket #26576 - WDGPHU - Flex Time report
    If UCase(rsSR("FUNCTION")) = UCase("Report_FlexTime") Then chkSSecurity(98) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Report_Friesens_IWantToKnowYou") Then chkSSecurity(61) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Friesens_ITHireForm") Then chkSSecurity(62) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Friesens_ITNoticeOfChange") Then chkSSecurity(63) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Friesens_NoticeOfChange") Then chkSSecurity(64) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Friesens_PerfImproveActionPlan") Then chkSSecurity(65) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Friesens_PerformanceReviewRpt") Then chkSSecurity(66) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Friesens_SeparationRpt") Then chkSSecurity(67) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Friesens_TerminationRpt") Then chkSSecurity(68) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Friesens_UpdateMeetingRpt") Then chkSSecurity(69) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Friesens_WarningRpt") Then chkSSecurity(70) = rsSR("ACCESSABLE")
    
    'Course Admin - Begin
    If UCase(rsSR("FUNCTION")) = UCase("CA_Report_Calendar") Then chkSSecurity(36) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Report_Class_List") Then chkSSecurity(37) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Report_Waiting_List") Then chkSSecurity(38) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Report_Conflict") Then chkSSecurity(39) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Report_Course_Catalog") Then chkSSecurity(40) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Report_Course_Per_Position") Then chkSSecurity(41) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Report_Label") Then chkSSecurity(42) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Report_Prerequ_Exception") Then chkSSecurity(43) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Report_CourseNotCompleted") Then chkSSecurity(44) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("CA_Report_Training_Summary") Then chkSSecurity(45) = rsSR("ACCESSABLE")
    'Course Admin - End
    
    If UCase(rsSR("FUNCTION")) = UCase("Codes") Then chkMCSecurity(3) = rsSR("ACCESSABLE")
    
    'Mostafa - Code Group Matrix
    If UCase(rsSR("FUNCTION")) = UCase("Attendance_Group_Code_Matrix_Update") Then chkUMSecurity(41) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Attendance_Group_Code_Matrix_Inquiry") Then chkUISecurity(41) = rsSR("ACCESSABLE")
    
    'Ticket #16189 - Friesens Job Files Attachment and Temporary/Cross Training Position
    If UCase(rsSR("FUNCTION")) = UCase("Job_Files_Attachment_Update") Then chkUMSecurity(42) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Job_Files_Attachment_Inquiry") Then chkUISecurity(42) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Temp_Cross_Training_Update") Then chkUMSecurity(43) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Temp_Cross_Training_Inquiry") Then chkUISecurity(43) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Training_List_Update") Then chkUMSecurity(44) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Training_List_Inquiry") Then chkUISecurity(44) = rsSR("ACCESSABLE")
    
    If glbWFC Then 'Ticket #18566
        If UCase(rsSR("FUNCTION")) = UCase("RetirementProc_Update") Then chkUMSecurity(45) = rsSR("ACCESSABLE")
        If UCase(rsSR("FUNCTION")) = UCase("RetirementProc_Inquiry") Then chkUISecurity(45) = rsSR("ACCESSABLE")
        If UCase(rsSR("FUNCTION")) = UCase("DeathProc_Update") Then chkUMSecurity(46) = rsSR("ACCESSABLE")
        If UCase(rsSR("FUNCTION")) = UCase("DeathProc_Inquiry") Then chkUISecurity(46) = rsSR("ACCESSABLE")
    End If
    
    If UCase(rsSR("FUNCTION")) = UCase("BudgetedManpower_Update") Then chkUMSecurity(47) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("BudgetedManpower_Inquiry") Then chkUISecurity(47) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("WorkScheduleRule_Update") Then chkUMSecurity(64) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WorkScheduleRule_Inquiry") Then chkUISecurity(64) = rsSR("ACCESSABLE")
    
    'Ticket #22541
    If UCase(rsSR("FUNCTION")) = UCase("DashboardSetup_Update") Then chkUMSecurity(65) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("DashboardSetup_Inquiry") Then chkUISecurity(65) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("RequiredCourses_Update") Then chkUMSecurity(50) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("RequiredCourses_Inquiry") Then chkUISecurity(50) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("BudgetedPosition_Update") Then chkUMSecurity(48) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("BudgetedPosition_Inquiry") Then chkUISecurity(48) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ApplicationProcess_Update") Then chkUMSecurity(49) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ApplicationProcess_Inquiry") Then chkUISecurity(49) = rsSR("ACCESSABLE")
        
    'Ticket #25015 - Macaulay
    If UCase(rsSR("FUNCTION")) = UCase("AddPayrollIDData_Update") Then chkUMSecurity(68) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("AddPayrollIDData_Inquiry") Then chkUISecurity(68) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Rehire_Update") Then chkUMSecurity(51) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Rehire_Inquiry") Then chkUISecurity(51) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("EnterLeave_Update") Then chkUMSecurity(52) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("EnterLeave_Inquiry") Then chkUISecurity(52) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("HS_ClaimMed_Update") Then chkUMSecurity(37) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("HS_ClaimMed_Inquiry") Then chkUISecurity(37) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("HS_Contacts_Update") Then chkUMSecurity(53) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("HS_Contacts_Inquiry") Then chkUISecurity(53) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("HS_Cost_Update") Then chkUMSecurity(54) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("HS_Cost_Inquiry") Then chkUISecurity(54) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("HS_CorrectAction_Update") Then chkUMSecurity(55) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("HS_CorrectAction_Inquiry") Then chkUISecurity(55) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("HS_RootCause_Update") Then chkUMSecurity(56) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("HS_RootCause_Inquiry") Then chkUISecurity(56) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("AffirmAction_Data_Update") Then chkUMSecurity(57) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("AffirmAction_Data_Inquiry") Then chkUISecurity(57) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("AffirmAction_Purge_Update") Then chkUMSecurity(58) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("AffirmAction_Purge_Inquiry") Then chkUISecurity(58) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("HS_W7CompanyMaster_Update") Then chkUMSecurity(61) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("HS_W7CompanyMaster_Inquiry") Then chkUISecurity(61) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("HS_W7Injury_Update") Then chkUMSecurity(63) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("HS_W7Injury_Inquiry") Then chkUISecurity(63) = rsSR("ACCESSABLE")
    
    'Form 9
    If UCase(rsSR("FUNCTION")) = UCase("HS_WF9_Update") Then chkUMSecurity(62) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("HS_WF9_Inquiry") Then chkUISecurity(62) = rsSR("ACCESSABLE")
    
    'If UCase(rsSR("FUNCTION")) = UCase("Profit_Sharing_Update") Then chkUMSecurity(62) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("Profit_Sharing_Inquiry") Then chkUISecurity(62) = rsSR("ACCESSABLE")
    
    If glbLinamar Then
        If UCase(rsSR("FUNCTION")) = UCase("Report_DoorAccess") Then chkSSecurity(33) = rsSR("ACCESSABLE")
    End If
    
    rsSR.MoveNext
Loop
rsSR.Close
Set rsSR = Nothing

'More Retrieves - calling from another procedure because this procedure is giving 'Procedure too large' message when
'compiling this form
Call More_Display_Values1

End Sub

Private Sub More_Display_Values1()
'More Retrieves - To resolve the 'Procedure too large' message when compiling this form

'Retrive Utilities and ESS & TimesheetSecurity settings
Call Display_Values1_Utilities_WebModules

'Ticket #23536 - Moved Mass Update Securities to another procedure, getting 'Procedure too large' message when compiling this form
Call GetMassUpdateSecurity_Settings

'Moved it to a procedure as was getting 'Procedure too large' message when compiling this form.
Call Retrieve_User_Country

'Retrieve User's security
Call Retrieve_User_Security

'Here call the find
Call FindTemplate

'Ticket #20585 - show the value on the Security Template dropdown list
Call Refresh_Security_Template

Me.cmdModify_Click
End Sub

Private Sub Display_Values1_Utilities_WebModules()
Dim rsSR As New ADODB.Recordset
Dim SQLQ

'????Ticket #24808 - Retrieve the Template Profile if the User's Security is based on Template. If the User is without Template or is TEMPLATE then retrieve the respective User's record
SQLQ = "SELECT " & Field_SQL("FUNCTION") & ", ACCESSABLE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND Maintainable=0"
' dkostka - 09/28/2001 - Changed from adOpenStatic to adOpenForwardOnly to improve speed.
rsSR.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

Do Until rsSR.EOF
    If UCase(rsSR("FUNCTION")) = UCase("Province") Then chkUSecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Entitle") Then chkUSecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Compress_Fix") Then chkUSecurity(2) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Matrix") Then chkUSecurity(3) = rsSR("ACCESSABLE")
    
    If glbLinamar Then
        If UCase(rsSR("FUNCTION")) = UCase("DoorName") Then chkUSecurity(4) = rsSR("ACCESSABLE")
        If UCase(rsSR("FUNCTION")) = UCase("Summarize_Attendance") Then chkUSecurity(5) = rsSR("ACCESSABLE")
    End If
    If UCase(rsSR("FUNCTION")) = UCase("TimeSheetPrority") Then chkUSecurity(6) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("TimeSheetUserPrority") Then chkUSecurity(7) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("DeleteTimeSheetFile") Then chkUSecurity(8) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ArchiveTimeSheetFile") Then chkUSecurity(25) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ApproveTimeSheetFile") Then chkUSecurity(26) = rsSR("ACCESSABLE")
    
    'If UCase(rsSR("FUNCTION")) = UCase("EssCompTime") Then chkUSecurity(26) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("TSOverrideEmpSecurity") Then chkUSecurity(14) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("DeleteApprovedTimeSheet") Then chkUSecurity(15) = rsSR("ACCESSABLE")
    
    '7.9 Enhancement
    If UCase(rsSR("FUNCTION")) = UCase("ESS_Time_Req") Then chkUSecurity(18) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_Vacation_Req") Then chkUSecurity(19) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_Request_Approval") Then chkUSecurity(20) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_Rpt_Request_Approval") Then chkUSecurity(21) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_Rpt_Print_Archive") Then chkUSecurity(22) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("ESS_Archive_Req") Then chkUSecurity(23) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Archive_VacTimeoff_Update") Then chkUSecurity(23) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_MassDelete_Req") Then chkUSecurity(24) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_ShowAllRequests") Then chkUSecurity(27) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("TS_PUNCHINOUT") Then chkUSecurity(28) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_REQ_OTHER_SUPER") Then chkUSecurity(29) = rsSR("ACCESSABLE")
     
    If UCase(rsSR("FUNCTION")) = UCase("ESS_DEL_APPR_TIME_REQ") Then chkUSecurity(30) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_DEL_APPR_VAC_REQ") Then chkUSecurity(31) = rsSR("ACCESSABLE")
     
    If UCase(rsSR("FUNCTION")) = UCase("ESS_CANCEL_REQ") Then chkUSecurity(32) = rsSR("ACCESSABLE")
       
    If UCase(rsSR("FUNCTION")) = UCase("ESS_TS_LIST_RA_ONLY") Then chkUSecurity(33) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_DELETE_FUTURE_VAC_REQ") Then chkUSecurity(34) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_DELETE_FUTURE_TIME_REQ") Then chkUSecurity(35) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("ESS_HTML_CALENDAR") Then chkUSecurity(36) = rsSR("ACCESSABLE")
    'Ticket #23536 - Dashboard ON/OFF
    If UCase(rsSR("FUNCTION")) = UCase("ESS_DASHBOARDS") Then chkUSecurity(37) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("ESS_QUICKINFO") Then chkUSecurity(38) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_DEMO_MAINTAIN") Then chkUSecurity(39) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_SHOWALLAPPREJ_REQS") Then chkUSecurity(40) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("TS_SUBMISSION") Then chkUSecurity(41) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_REQAPPDEPTSEC") Then chkUSecurity(42) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_APPVIEWOWN") Then chkUSecurity(43) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_MYCOWORKER") Then chkUSecurity(44) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("TS_ENABLE_SUPER") Then chkUSecurity(45) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("TS_RA_ONLY") Then chkUSecurity(46) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ESS_TS_SETUP_CHECKLIST") Then chkUSecurity(47) = rsSR("ACCESSABLE")
     
    
    If UCase(rsSR("FUNCTION")) = UCase("CompanyPreference") Then chkUSecurity(9) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("EmpFlagsSetup") Then chkUSecurity(10) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("MultiDataSourceSetup") Then chkUSecurity(11) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("HelpDescSetup") Then chkUSecurity(12) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("BenefitGroupSetup") Then chkUSecurity(13) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ChangeYourPassword") Then chkUSecurity(16) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("ITAdmin") Then chkUSecurity(17) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Import_Attendance") Then chkIESecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Export_Attendance") Then chkIESecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Import_Benefits") Then chkIESecurity(2) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Export_Benefits") Then chkIESecurity(3) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Import_Employee") Then chkIESecurity(4) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Export_Employee") Then chkIESecurity(5) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Import_Salaries") Then chkIESecurity(6) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Export_Salaries") Then chkIESecurity(7) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Import_Table") Then chkIESecurity(8) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Export_Table") Then chkIESecurity(9) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Import_YTD") Then chkIESecurity(10) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Export_YTD") Then chkIESecurity(11) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Import_PayrollTrans") Then chkIESecurity(12) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Export_PayrollTrans") Then chkIESecurity(13) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Import_ContEdu") Then chkIESecurity(14) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Export_ContEdu") Then chkIESecurity(15) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Import_PerfReview") Then chkIESecurity(16) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Export_PerfReview") Then chkIESecurity(17) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Import_EmploymentEquity") Then chkIESecurity(18) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Export_EmploymentEquity") Then chkIESecurity(19) = rsSR("ACCESSABLE")

    'Ticket #29122 - New Database Setup and Integration Setup securities
    If UCase(rsSR("FUNCTION")) = UCase("IntegrtDBSetup_Update") Then chkUMSecurity(72) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("IntegrtDBSetup_Inquiry") Then chkUISecurity(72) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("IntegrtSetup_Update") Then chkUMSecurity(73) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("IntegrtSetup_Inquiry") Then chkUISecurity(73) = rsSR("ACCESSABLE")

    rsSR.MoveNext
Loop
rsSR.Close
Set rsSR = Nothing

End Sub

Private Sub GetMassUpdateSecurity_Settings()

'Ticket #23536 - Moved out of Display_Values1 procedure, getting 'Procedure too large' message when compiling this form
    
chkMCSecurity(0) = GetMassUpdateSecurities("Attendance_His_MassUpdate", glbSecUSERID)
chkMCSecurity(1) = GetMassUpdateSecurities("Attendance_MassUpdate", glbSecUSERID)
chkMCSecurity(2) = GetMassUpdateSecurities("Benefits_MassUpdate", glbSecUSERID)

chkMCSecurity(4) = GetMassUpdateSecurities("Education_Seminars_MassUpdate", glbSecUSERID)
chkMCSecurity(5) = GetMassUpdateSecurities("Other_Entitlements_MassUpdate", glbSecUSERID)
chkMCSecurity(6) = GetMassUpdateSecurities("Entitlements_MassUpdate", glbSecUSERID)
chkMCSecurity(7) = GetMassUpdateSecurities("Follow_Ups_MassUpdate", glbSecUSERID)
chkMCSecurity(8) = GetMassUpdateSecurities("Hrly_Entitlements_MassUpdate", glbSecUSERID)
chkMCSecurity(9) = GetMassUpdateSecurities("Other_Earnings_MassUpdate", glbSecUSERID)
chkMCSecurity(10) = GetMassUpdateSecurities("Job_Master_MassUpdate", glbSecUSERID)
chkMCSecurity(11) = GetMassUpdateSecurities("Salary_MassUpdate", glbSecUSERID)
chkMCSecurity(12) = GetMassUpdateSecurities("EmployeeNo_MassUpdate", glbSecUSERID)
chkMCSecurity(13) = GetMassUpdateSecurities("OvertimeMaster_MassUpdate", glbSecUSERID)
chkMCSecurity(14) = GetMassUpdateSecurities("Emergency_Leave_MassUpdate", glbSecUSERID)
chkMCSecurity(15) = GetMassUpdateSecurities("Import_Photo_MassUpdate", glbSecUSERID)
chkMCSecurity(16) = GetMassUpdateSecurities("Work_Schedule_MassUpdate", glbSecUSERID)

'Ticket #22893 - Security for Year End based on Anniversary Month
'If glbCompSerial = "S/N - 2448W" Then  'For all clients with Security rights
    chkMCSecurity(17) = GetMassUpdateSecurities("YearEnd_AnniversaryMonth_MassUpdate", glbSecUSERID)
'End If

'Release 8.0 - Ticket #24361: Add Email Address import under Mass Updates menu
chkMCSecurity(18) = GetMassUpdateSecurities("EmailLoad_MassUpdate", glbSecUSERID)

'Release 8.1 - Ticket #27244: Import document Attachment under Mass Updates menu
chkMCSecurity(19) = GetMassUpdateSecurities("ImpAttachment_MassUpdate", glbSecUSERID)

End Sub

Private Sub Retrieve_User_Country()
Dim rsSR As New ADODB.Recordset
Dim SQLQ As String

SQLQ = "SELECT COUNTRY FROM HR_SECURE_BASIC WHERE  USERID='" & Replace(glbSecUSERID, "'", "''") & "'"
rsSR.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
If rsSR.EOF = False And rsSR.BOF = False Then
    If Len(rsSR("COUNTRY")) > 0 Then
        cmbCountry.Text = rsSR("COUNTRY")
    End If
End If
rsSR.Close
End Sub

Private Sub Retrieve_User_Security()
Dim rsSR As New ADODB.Recordset
Dim SQLQ As String

SQLQ = "SELECT LOCK_PASSWORD FROM HR_SECURE_BASIC WHERE  USERID='" & Replace(txtUSERID, "'", "''") & "'"
rsSR.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
If rsSR.EOF = False And rsSR.BOF = False Then
    chkPswdLocked = IIf(IsNull(rsSR("LOCK_PASSWORD")), False, rsSR("LOCK_PASSWORD"))
End If
rsSR.Close
End Sub

Private Sub UpdSecAccess()
Dim SQLQ
SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND CODENAME is NULL"

'don't delete Custom Features security #3581
If glbWFC Then
    SQLQ = SQLQ & " AND NOT (LEFT([FUNCTION],3)='WFC')" 'WFC_ or WFCPEN_
End If
If glbWHSCC Then
    SQLQ = SQLQ & " AND NOT (LEFT([FUNCTION],4)='WHSC')"
End If
If glbSamuel Then 'Ticket #23228 Franks 02/07/2013
    SQLQ = SQLQ & " AND NOT (LEFT([FUNCTION],4)='SAM_')"
End If
'don't delete Custom Features security #3581
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

Call AddSecAccess

End Sub

Private Sub Delete_Existing_User_Profile()
Dim SQLQ
SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND CODENAME is NULL"

'don't delete Custom Features security #3581
If glbWFC Then
    SQLQ = SQLQ & " AND NOT (LEFT([FUNCTION],3)='WFC')" 'WFC_ or WFCPEN_
End If
If glbWHSCC Then
    SQLQ = SQLQ & " AND NOT (LEFT([FUNCTION],4)='WHSC')"
End If
If glbSamuel Then 'Ticket #23228 Franks 02/07/2013
    SQLQ = SQLQ & " AND NOT (LEFT([FUNCTION],4)='SAM_')"
End If
'don't delete Custom Features security #3581
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

End Sub

Private Sub AddSecDept()
Dim SQLQ, sqlI
    SQLQ = "INSERT INTO HRPASDEP (PD_COMPNO,PD_USERID,PD_DEPT,PD_LDATE,PD_LTIME,PD_LUSER) "
    SQLQ = SQLQ & "VALUES('001','" & Replace(Trim(txtUSERID), "'", "''") & "','ALL'," & Date_SQL(Date) & ",'" & Time$ & "','" & glbUserID & "') "
    gdbAdoIhr001.Execute SQLQ
End Sub

Private Sub AddSecAccess()
Dim SQLQ, sqlI

sqlI = "INSERT INTO HR_SECURE_ACCESS(COMPNO,USERID," & Field_SQL("FUNCTION") & ",ACCESSABLE) "
sqlI = sqlI & " VALUES('001','" & Replace(Trim(txtUSERID), "'", "''") & "',"
gdbAdoIhr001.BeginTrans

SQLQ = sqlI & "'Company_Update'," & IIf(chkUMSecurity(0), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'Company_Inquiry'," & IIf(chkUISecurity(0), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Security_Update'," & IIf(chkUMSecurity(1), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Security_Inquiry'," & IIf(chkUISecurity(1), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Master_Table_Update'," & IIf(chkUMSecurity(2), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Master_Table_Inquiry'," & IIf(chkUISecurity(2), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Departments_Update'," & IIf(chkUMSecurity(3), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Department_Inquiry'," & IIf(chkUISecurity(3), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Audit_Update'," & IIf(chkUMSecurity(4), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Audit_Inquiry'," & IIf(chkUISecurity(4), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EmploymentEQT_Update'," & IIf(chkUMSecurity(5), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EmploymentEQT_Inquiry'," & IIf(chkUISecurity(5), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Divisions_Update'," & IIf(chkUMSecurity(6), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Divisions_Inquiry'," & IIf(chkUISecurity(6), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Ledgers_Update'," & IIf(chkUMSecurity(7), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Ledgers_Inquiry'," & IIf(chkUISecurity(7), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
If glbLinamar Then
    SQLQ = sqlI & "'DoorAccess_Update'," & IIf(chkUMSecurity(8), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
    SQLQ = sqlI & "'DoorAccess_Inquiry'," & IIf(chkUISecurity(8), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
     
     SQLQ = sqlI & "'ProductLine_Operation_Update'," & IIf(chkUMSecurity(20), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
    SQLQ = sqlI & "'ProductLine_Operation_Inquiry'," & IIf(chkUISecurity(20), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
End If
If glbCompSerial = "S/N - 2380W" Then ' For VitalAire Canada Inc. Ticket #26233 Franks 11/20/2014
    SQLQ = sqlI & "'DoorAccess_Update'," & IIf(chkUMSecurity(8), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
    SQLQ = sqlI & "'DoorAccess_Inquiry'," & IIf(chkUISecurity(8), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
End If
'Ticket #23409 - Samuel, Son & Co., Limited - Discipline Audit Table Report
If glbCompSerial = "S/N - 2382W" Then
    SQLQ = sqlI & "'CounselAudit_Update'," & IIf(chkUMSecurity(66), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
    SQLQ = sqlI & "'CounselAudit_Inquiry'," & IIf(chkUISecurity(66), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
End If

'Ticket #24655 - Wellington-Dufferin-Guelph Public Health - On Call Hours
If glbCompSerial = "S/N - 2411W" Then
    SQLQ = sqlI & "'On_Call_Hours_Update'," & IIf(chkUMSecurity(67), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
    SQLQ = sqlI & "'On_Call_Hours_Inquiry'," & IIf(chkUISecurity(67), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
End If

SQLQ = sqlI & "'CustomReport_Update'," & IIf(chkUMSecurity(9), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CustomReport_Inquiry'," & IIf(chkUISecurity(9), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Holiday_Update'," & IIf(chkUMSecurity(10), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Holiday_Inquiry'," & IIf(chkUISecurity(10), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'New_Hire_Update'," & IIf(chkUMSecurity(11), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'New_Hire_Inquiry'," & IIf(chkUISecurity(11), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Label_Update'," & IIf(chkUMSecurity(12), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Label_Inquiry'," & IIf(chkUISecurity(12), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Sal_Distribute_Update'," & IIf(chkUMSecurity(13), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Sal_Distribute_Inquiry'," & IIf(chkUISecurity(13), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Pay_Period_Update'," & IIf(chkUMSecurity(19), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Pay_Period_Inquiry'," & IIf(chkUISecurity(19), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Email_Setup_Update'," & IIf(chkUMSecurity(18), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Email_Setup_Inquiry'," & IIf(chkUISecurity(18), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ



SQLQ = sqlI & "'Payroll_Category_Update'," & IIf(chkUMSecurity(14), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Payroll_Category_Inquiry'," & IIf(chkUISecurity(14), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

'Ticket #25746 - Town of St. Marys
SQLQ = sqlI & "'DeptGL_Matrix_Update'," & IIf(chkUMSecurity(70), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'DeptGL_Matrix_Inquiry'," & IIf(chkUISecurity(70), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
  
SQLQ = sqlI & "'Charge_Code_Update'," & IIf(chkUMSecurity(15), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Charge_Code_Inquiry'," & IIf(chkUISecurity(15), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
SQLQ = sqlI & "'Project_Code_Update'," & IIf(chkUMSecurity(16), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Project_Code_Inquiry'," & IIf(chkUISecurity(16), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
SQLQ = sqlI & "'Machine_Update'," & IIf(chkUMSecurity(17), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Machine_Inquiry'," & IIf(chkUISecurity(17), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
'7.6
SQLQ = sqlI & "'EMP_FLAGS_Update'," & IIf(chkUMSecurity(22), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EMP_FLAGS_Inquiry'," & IIf(chkUISecurity(22), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EMP_HISTORY_Update'," & IIf(chkUMSecurity(23), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EMP_HISTORY_Inquiry'," & IIf(chkUISecurity(23), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'GL_DIST_Update'," & IIf(chkUMSecurity(24), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'GL_DIST_Inquiry'," & IIf(chkUISecurity(24), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EMP_LANG_Update'," & IIf(chkUMSecurity(25), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EMP_LANG_Inquiry'," & IIf(chkUISecurity(25), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EMP_SUCCESSION_Update'," & IIf(chkUMSecurity(26), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EMP_SUCCESSION_Inquiry'," & IIf(chkUISecurity(26), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Work_Schedule_Update'," & IIf(chkUMSecurity(60), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Work_Schedule_Inquiry'," & IIf(chkUISecurity(60), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

'Attendance Code Matrix
SQLQ = sqlI & "'AttendCode_Matrix_Update'," & IIf(chkUMSecurity(59), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'AttendCode_Matrix_Inquiry'," & IIf(chkUISecurity(59), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Ticket #22682 - Release 8.0 - Follow Up Code Email Matrix
SQLQ = sqlI & "'FollowUpCodeEmail_Matrix_Update'," & IIf(chkUMSecurity(69), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'FollowUpCodeEmail_Matrix_Inquiry'," & IIf(chkUISecurity(69), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Ticket #25922 - OHRS Reporting for CHC
SQLQ = sqlI & "'OHRS_Department_Update'," & IIf(chkUMSecurity(71), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'OHRS_Department_Inquiry'," & IIf(chkUISecurity(71), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Emergency_Contacts_Update'," & IIf(chkUMSecurity(35), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Emergency_Contacts_Inquiry'," & IIf(chkUISecurity(35), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'ADP Data
SQLQ = sqlI & "'ADP_Data_Update'," & IIf(chkUMSecurity(36), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'ADP_Data_Inquiry'," & IIf(chkUISecurity(36), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Sam Added for ESS.NET 07/27/2006 Ticket # 11403
'SQLQ = sqlI & "'Archive_VacTimeoff_Update'," & IIf(chkUMSecurity(37), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'Archive_VacTimeoff_Inquiry'," & IIf(chkUISecurity(37), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
    
SQLQ = sqlI & "'UserDefineTbl_Update'," & IIf(chkUMSecurity(39), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'UserDefineTbl_Inquiry'," & IIf(chkUISecurity(39), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
    
SQLQ = sqlI & "'PayrollTrans_Update'," & IIf(chkUMSecurity(40), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'PayrollTrans_Inquiry'," & IIf(chkUISecurity(40), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
 
'Course Code Master
SQLQ = sqlI & "'CourseCodeMaster_Update'," & IIf(chkUMSecurity(38), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CourseCodeMaster_Inquiry'," & IIf(chkUISecurity(38), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'ends


'Course Admin - Begin
SQLQ = sqlI & "'CA_Organization_Update'," & IIf(chkUMSecurity(27), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CA_Organization_Inquiry'," & IIf(chkUISecurity(27), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CA_Catalog_Update'," & IIf(chkUMSecurity(28), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CA_Catalog_Inquiry'," & IIf(chkUISecurity(28), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CA_TrainingLocation_Update'," & IIf(chkUMSecurity(29), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CA_TrainingLocation_Inquiry'," & IIf(chkUISecurity(29), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CA_Scheduling_Update'," & IIf(chkUMSecurity(30), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CA_Scheduling_Inquiry'," & IIf(chkUISecurity(30), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CA_EnrollRequest_Update'," & IIf(chkUMSecurity(31), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CA_EnrollRequest_Inquiry'," & IIf(chkUISecurity(31), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CA_EnrollApproval_Update'," & IIf(chkUMSecurity(32), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CA_EnrollApproval_Inquiry'," & IIf(chkUISecurity(32), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CA_Enrollment_Update'," & IIf(chkUMSecurity(33), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CA_Enrollment_Inquiry'," & IIf(chkUISecurity(33), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CA_WaitingList_Update'," & IIf(chkUMSecurity(34), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'CA_WaitingList_Inquiry'," & IIf(chkUISecurity(34), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
'Course Admin - End
SQLQ = sqlI & "'Province'," & IIf(chkUSecurity(0), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Entitle'," & IIf(chkUSecurity(1), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Compress_Fix'," & IIf(chkUSecurity(2), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Matrix'," & IIf(chkUSecurity(3), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
  
If glbLinamar Then
    SQLQ = sqlI & "'DoorName'," & IIf(chkUSecurity(4), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
    SQLQ = sqlI & "'Summarize_Attendance'," & IIf(chkUSecurity(5), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
End If
SQLQ = sqlI & "'TimeSheetPrority'," & IIf(chkUSecurity(6), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'TimeSheetUserPrority'," & IIf(chkUSecurity(7), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'DeleteTimeSheetFile'," & IIf(chkUSecurity(8), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 'ArchiveTimeSheetFile
 SQLQ = sqlI & "'ArchiveTimeSheetFile'," & IIf(chkUSecurity(25), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 'ApproveTimeSheetFile
 SQLQ = sqlI & "'ApproveTimeSheetFile'," & IIf(chkUSecurity(26), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 'EssCompTime
 ' SQLQ = sqlI & "'EssCompTime'," & IIf(chkUSecurity(26), 1, 0) & ")"
' gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'TSOverrideEmpSecurity'," & IIf(chkUSecurity(14), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'DeleteApprovedTimeSheet'," & IIf(chkUSecurity(15), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

'7.9 Enhancement
SQLQ = sqlI & "'ESS_Time_Req'," & IIf(chkUSecurity(18), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'ESS_Vacation_Req'," & IIf(chkUSecurity(19), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'ESS_Request_Approval'," & IIf(chkUSecurity(20), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'ESS_Rpt_Request_Approval'," & IIf(chkUSecurity(21), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'ESS_Rpt_Print_Archive'," & IIf(chkUSecurity(22), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'ESS_Archive_Req'," & IIf(chkUSecurity(23), 1, 0) & ")"
SQLQ = sqlI & "'Archive_VacTimeoff_Update'," & IIf(chkUSecurity(23), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'ESS_MassDelete_Req'," & IIf(chkUSecurity(24), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'ESS_ShowAllRequests'," & IIf(chkUSecurity(27), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

'TS_PUNCHINOUT
SQLQ = sqlI & "'TS_PUNCHINOUT'," & IIf(chkUSecurity(28), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
SQLQ = sqlI & "'ESS_REQ_OTHER_SUPER'," & IIf(chkUSecurity(29), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'ESS_DEL_APPR_TIME_REQ'," & IIf(chkUSecurity(30), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'ESS_DEL_APPR_VAC_REQ'," & IIf(chkUSecurity(31), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
SQLQ = sqlI & "'ESS_CANCEL_REQ'," & IIf(chkUSecurity(32), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'ESS_TS_LIST_RA_ONLY'," & IIf(chkUSecurity(33), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'ESS_DELETE_FUTURE_VAC_REQ'," & IIf(chkUSecurity(34), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'ESS_DELETE_FUTURE_TIME_REQ'," & IIf(chkUSecurity(35), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
SQLQ = sqlI & "'ESS_HTML_CALENDAR'," & IIf(chkUSecurity(36), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

'Ticket #23536 - Dashboard ON/OFF
SQLQ = sqlI & "'ESS_DASHBOARDS'," & IIf(chkUSecurity(37), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
 SQLQ = sqlI & "'ESS_QUICKINFO'," & IIf(chkUSecurity(38), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
SQLQ = sqlI & "'ESS_DEMO_MAINTAIN'," & IIf(chkUSecurity(39), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
 SQLQ = sqlI & "'ESS_SHOWALLAPPREJ_REQS'," & IIf(chkUSecurity(40), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
  SQLQ = sqlI & "'TS_SUBMISSION'," & IIf(chkUSecurity(41), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

 SQLQ = sqlI & "'ESS_REQAPPDEPTSEC'," & IIf(chkUSecurity(42), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'ESS_APPVIEWOWN'," & IIf(chkUSecurity(43), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'ESS_MYCOWORKER'," & IIf(chkUSecurity(44), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'TS_ENABLE_SUPER'," & IIf(chkUSecurity(45), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'TS_RA_ONLY'," & IIf(chkUSecurity(46), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'ESS_TS_SETUP_CHECKLIST'," & IIf(chkUSecurity(47), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ


SQLQ = sqlI & "'CompanyPreference'," & IIf(chkUSecurity(9), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EmpFlagsSetup'," & IIf(chkUSecurity(10), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'MultiDataSourceSetup'," & IIf(chkUSecurity(11), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'HelpDescSetup'," & IIf(chkUSecurity(12), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'BenefitGroupSetup'," & IIf(chkUSecurity(13), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'ChangeYourPassword'," & IIf(chkUSecurity(16), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
 SQLQ = sqlI & "'ITAdmin'," & IIf(chkUSecurity(17), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Import_Attendance'," & IIf(chkIESecurity(0), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Export_Attendance'," & IIf(chkIESecurity(1), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Import_Benefits'," & IIf(chkIESecurity(2), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Export_Benefits'," & IIf(chkIESecurity(3), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Import_Employee'," & IIf(chkIESecurity(4), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Export_Employee'," & IIf(chkIESecurity(5), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Import_Salaries'," & IIf(chkIESecurity(6), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Export_Salaries'," & IIf(chkIESecurity(7), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Import_Table'," & IIf(chkIESecurity(8), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Export_Table'," & IIf(chkIESecurity(9), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Import_YTD'," & IIf(chkIESecurity(10), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Export_YTD'," & IIf(chkIESecurity(11), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Import_PayrollTrans'," & IIf(chkIESecurity(12), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Export_PayrollTrans'," & IIf(chkIESecurity(13), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Import_ContEdu'," & IIf(chkIESecurity(14), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Export_ContEdu'," & IIf(chkIESecurity(15), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Import_PerfReview'," & IIf(chkIESecurity(16), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Export_PerfReview'," & IIf(chkIESecurity(17), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Import_EmploymentEquity'," & IIf(chkIESecurity(18), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Export_EmploymentEquity'," & IIf(chkIESecurity(19), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'Import_Profit_Sharing'," & IIf(chkIESecurity(20), 1, 0) & ")"
' gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'Export_Profit_Sharing'," & IIf(chkIESecurity(21), 1, 0) & ")"
' gdbAdoIhr001.Execute SQLQ

'Ticket #29122 - New Database Setup and Integration Setup securities
SQLQ = sqlI & "'IntegrtDBSetup_Update'," & IIf(chkUMSecurity(72), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'IntegrtDBSetup_Inquiry'," & IIf(chkUISecurity(72), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'IntegrtSetup_Update'," & IIf(chkUMSecurity(73), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'IntegrtSetup_Inquiry'," & IIf(chkUISecurity(73), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ



SQLQ = sqlI & "'Basic_Update'," & IIf(chkMSecurity(0), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Basic_Inquiry'," & IIf(chkSecurity(0), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Banking_Update'," & IIf(chkMSecurity(1), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Banking_Inquiry'," & IIf(chkSecurity(1), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Dependents_Update'," & IIf(chkMSecurity(2), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Dependents_Inquiry'," & IIf(chkSecurity(2), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Skills_Update'," & IIf(chkMSecurity(3), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Skills_Inquiry'," & IIf(chkSecurity(3), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Formal_Education_Update'," & IIf(chkMSecurity(4), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Formal_Education_Inquiry'," & IIf(chkSecurity(4), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Salary_Update'," & IIf(chkMSecurity(5), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Salary_Inquiry'," & IIf(chkSecurity(5), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Performance_Update'," & IIf(chkMSecurity(6), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Performance_Inquiry'," & IIf(chkSecurity(6), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Position_Update'," & IIf(chkMSecurity(7), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Position_Inquiry'," & IIf(chkSecurity(7), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Benefits_Update'," & IIf(chkMSecurity(8), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Benefits_Inquiry'," & IIf(chkSecurity(8), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

'7.9 Enhancement
SQLQ = sqlI & "'Beneficiary_Update'," & IIf(chkMSecurity(29), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Beneficiary_Inquiry'," & IIf(chkSecurity(29), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Entitlements_Update'," & IIf(chkMSecurity(9), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Entitlements_Inquiry'," & IIf(chkSecurity(9), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Associations_Update'," & IIf(chkMSecurity(10), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Associations_Inquiry'," & IIf(chkSecurity(10), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Follow_Ups_Update'," & IIf(chkMSecurity(11), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Follow_Ups_Inquiry'," & IIf(chkSecurity(11), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Health_Safety_Update'," & IIf(chkMSecurity(12), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Health_Safety_Inquiry'," & IIf(chkSecurity(12), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Attendance_Update'," & IIf(chkMSecurity(13), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Attendance_Inquiry'," & IIf(chkSecurity(13), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Other_Entitlements_Update'," & IIf(chkMSecurity(14), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Other_Entitlements_Inquiry'," & IIf(chkSecurity(14), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Other_Earnings_Update'," & IIf(chkMSecurity(15), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Other_Earnings_Inquiry'," & IIf(chkSecurity(15), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Terminations_Update'," & IIf(chkMSecurity(16), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Termination_Inquiry'," & IIf(chkSecurity(16), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Job_Classes_Update'," & IIf(chkMSecurity(17), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Job_Classes_Inquiry'," & IIf(chkSecurity(17), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Education_Seminars_Update'," & IIf(chkMSecurity(18), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Education_Seminars_Inquiry'," & IIf(chkSecurity(18), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Job_Skills_Update'," & IIf(chkMSecurity(19), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Job_Skills_Inquiry'," & IIf(chkSecurity(19), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Job_Eval_Update'," & IIf(chkMSecurity(20), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Job_Eval_Inquiry'," & IIf(chkSecurity(20), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Job_Master_Update'," & IIf(chkMSecurity(21), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Job_Master_Inquiry'," & IIf(chkSecurity(21), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Hrly_Entitlements_Update'," & IIf(chkMSecurity(22), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Hrly_Entitlements_Inquiry'," & IIf(chkSecurity(22), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Show_SIN_SSN'," & IIf(chkEESIN, 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Show_DOB'," & IIf(chkEEDOB, 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Show_ADDRESS'," & IIf(chkEEADDRESS, 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Show_MARITAL'," & IIf(chkEEMarital, 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
'Moved to HR_SECURE_BASIC
'If glbCompSerial = "S/N - 2407W" Then 'Ticket #18406 - Farmers' Mutual Insurance
'    SQLQ = sqlI & "'Lock_Password'," & IIf(chkPswdLocked, 1, 0) & ")"
'     gdbAdoIhr001.Execute SQLQ
'End If

'tkt#10423 Jerry said make it available for everyone
'If glbCompSerial = "S/N - 2173W" Then
    SQLQ = sqlI & "'Add_Attendance'," & IIf(chkASecurity, 1, 0) & ")"
    gdbAdoIhr001.Execute SQLQ
'End If

'Ticket #22682 - Release 8.0
SQLQ = sqlI & "'Add_NewHire'," & IIf(chkNHireSecurity, 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Release 8.1
SQLQ = sqlI & "'Add_Comments'," & IIf(chkACommentSecurity, 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Ticket #23923 - Release 8.0 - View Own
SQLQ = sqlI & "'ScsPlan_ViewOwn'," & IIf(chkViewOwnSuccPlan, 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Comments_ViewOwn'," & IIf(chkViewOwnComm, 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Counsel_ViewOwn'," & IIf(chkViewOwnCounsel, 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'FollUp_ViewOwn'," & IIf(chkViewOwnFollUp, 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'OthInfo_ViewOwn'," & IIf(chkViewOwnOthInfo, 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EmpFlags_ViewOwn'," & IIf(chkViewOwnEmpFlags, 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EmpHis_ViewOwn'," & IIf(chkViewOwnEmpHis, 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'GLDist_ViewOwn'," & IIf(chkViewOwnGLDist, 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Ticket #28635 - Add View Own security
SQLQ = sqlI & "'Perform_ViewOwn'," & IIf(chkViewOwnPerform, 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Ticket #22009 Franks 05/10/2012
SQLQ = sqlI & "'Del_Dependents'," & IIf(chkDSecurity, 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
    
' dkostka - 09/25/2001 - Added security for Counselling screen.
SQLQ = sqlI & "'Counselling_Update'," & IIf(chkMSecurity(23), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Counselling_Inquiry'," & IIf(chkSecurity(23), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
If glbWFC Then
    SQLQ = sqlI & "'SalaryGrids_Update'," & IIf(chkMSecurity(24), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
    SQLQ = sqlI & "'SalaryGrids_Inquiry'," & IIf(chkSecurity(24), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
End If
SQLQ = sqlI & "'Comments_Update'," & IIf(chkMSecurity(25), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Comments_Inquiry'," & IIf(chkSecurity(25), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'OtherInformation_Update'," & IIf(chkMSecurity(26), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'OtherInformation_Inquiry'," & IIf(chkSecurity(26), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
If glbLinamar Then
    SQLQ = sqlI & "'LinamarSkills_Update'," & IIf(chkMSecurity(27), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
    SQLQ = sqlI & "'LinamarSkills_Inquiry'," & IIf(chkSecurity(27), 1, 0) & ")"
     gdbAdoIhr001.Execute SQLQ
End If
SQLQ = sqlI & "'Attendance_History_Update'," & IIf(chkMSecurity(28), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Attendance_History_Inquiry'," & IIf(chkSecurity(28), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

 
' Dijana - May 2002 - Add security for Applicant
SQLQ = sqlI & "'App_Basic_Update'," & IIf(chkAMSecurity(0), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Basic_Inquiry'," & IIf(chkAISecurity(0), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Skills_Update'," & IIf(chkAMSecurity(1), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Skills_Inquiry'," & IIf(chkAISecurity(1), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Formal_Education_Update'," & IIf(chkAMSecurity(2), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Formal_Education_Inquiry'," & IIf(chkAISecurity(2), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Education_Seminars_Update'," & IIf(chkAMSecurity(3), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Education_Seminars_Inquiry'," & IIf(chkAISecurity(3), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Interview_Update'," & IIf(chkAMSecurity(4), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Interview_Inquiry'," & IIf(chkAISecurity(4), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Associations_Update'," & IIf(chkAMSecurity(5), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Associations_Inquiry'," & IIf(chkAISecurity(5), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_References_Update'," & IIf(chkAMSecurity(6), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_References_Inquiry'," & IIf(chkAISecurity(6), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Follow_Ups_Update'," & IIf(chkAMSecurity(7), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Follow_Ups_Inquiry'," & IIf(chkAISecurity(7), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Requisition_Update'," & IIf(chkAMSecurity(8), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Requisition_Inquiry'," & IIf(chkAISecurity(8), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Recruitment_Update'," & IIf(chkAMSecurity(9), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Recruitment_Inquiry'," & IIf(chkAISecurity(9), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'App_Employment_Update'," & IIf(chkAMSecurity(10), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_Employment_Inquiry'," & IIf(chkAISecurity(10), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
'Ticket #30508 - Applicant Tracking Enhancement
SQLQ = sqlI & "'App_LetterPosType_Update'," & IIf(chkAMSecurity(11), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_LetterPosType_Inquiry'," & IIf(chkAISecurity(11), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_FormWorkflow_Update'," & IIf(chkAMSecurity(12), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_FormWorkflow_Inquiry'," & IIf(chkAISecurity(12), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_FormDefaults_Update'," & IIf(chkAMSecurity(13), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'App_FormDefaults_Inquiry'," & IIf(chkAISecurity(13), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
'reports
SQLQ = sqlI & "'Report_Age'," & IIf(chkSSecurity(0), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Compensatory_Time'," & IIf(chkSSecurity(1), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Cost_Of_Employment'," & IIf(chkSSecurity(2), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Emergecy_Contacts'," & IIf(chkSSecurity(3), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Employee_Labels'," & IIf(chkSSecurity(4), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Job_List'," & IIf(chkSSecurity(5), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Profiles'," & IIf(chkSSecurity(6), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Entitlements'," & IIf(chkSSecurity(7), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Follow_Ups'," & IIf(chkSSecurity(8), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Home_Address'," & IIf(chkSSecurity(9), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Salary_Performance'," & IIf(chkSSecurity(10), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
'Ticket #27795 - Friesens Corporation
SQLQ = sqlI & "'Report_Staff_Profile'," & IIf(chkSSecurity(100), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Seniority'," & IIf(chkSSecurity(11), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Telephone_Extensions'," & IIf(chkSSecurity(12), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Associations'," & IIf(chkSSecurity(13), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Master_Attendance'," & IIf(chkSSecurity(14), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Report_Bonus_Attendance'," & IIf(chkSSecurity(83), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Calendar_Attendance'," & IIf(chkSSecurity(84), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Costed_Attendance'," & IIf(chkSSecurity(85), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Report_Master_Benefits'," & IIf(chkSSecurity(15), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Master_Division'," & IIf(chkSSecurity(16), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Master_Termination'," & IIf(chkSSecurity(17), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Master_Formal_Education'," & IIf(chkSSecurity(18), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Master_Job'," & IIf(chkSSecurity(19), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Master_Passwords'," & IIf(chkSSecurity(20), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Master_Salaries'," & IIf(chkSSecurity(21), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Master_Edu_Seminars'," & IIf(chkSSecurity(22), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Master_Table_Codes'," & IIf(chkSSecurity(23), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Heatlh_Safety'," & IIf(chkSSecurity(24), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Master_DolEnt'," & IIf(chkSSecurity(25), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Hourly_Entitlements'," & IIf(chkSSecurity(26), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Master_OtherEarn'," & IIf(chkSSecurity(27), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'PayEQT_Inquiry'," & IIf(chkSSecurity(28), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Dependents'," & IIf(chkSSecurity(29), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Skills'," & IIf(chkSSecurity(30), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Languages'," & IIf(chkSSecurity(31), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Employee_Turnover'," & IIf(chkSSecurity(32), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Counselling', " & IIf(chkSSecurity(34), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
'Release 8.1
SQLQ = sqlI & "'Report_DocumentType', " & IIf(chkSSecurity(99), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Emergency_Leave', " & IIf(chkSSecurity(35), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'Report_External_Hire', " & IIf(chkSSecurity(48), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Internal_Hire', " & IIf(chkSSecurity(49), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Key_Workforce', " & IIf(chkSSecurity(50), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Manpower_Plan', " & IIf(chkSSecurity(51), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Staff_Management', " & IIf(chkSSecurity(52), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_WC_Time', " & IIf(chkSSecurity(53), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_WC_Work', " & IIf(chkSSecurity(54), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Paid_Sick', " & IIf(chkSSecurity(55), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_User_Defined_Table', " & IIf(chkSSecurity(56), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Future_Entitlement', " & IIf(chkSSecurity(57), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Overtime
SQLQ = sqlI & "'Report_Overtime_Bank', " & IIf(chkSSecurity(46), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Overtime_Lost_Hours', " & IIf(chkSSecurity(47), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'More
SQLQ = sqlI & "'Report_Employee_Flags', " & IIf(chkSSecurity(58), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Temp_CrossTraining', " & IIf(chkSSecurity(59), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Required_Course_Hist', " & IIf(chkSSecurity(60), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Report_Email_Address'," & IIf(chkSSecurity(71), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_LOA'," & IIf(chkSSecurity(73), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_POE'," & IIf(chkSSecurity(74), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_SINSSN'," & IIf(chkSSecurity(75), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Succession'," & IIf(chkSSecurity(76), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Gap_Analysis'," & IIf(chkSSecurity(72), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Report_GL_Distribution'," & IIf(chkSSecurity(86), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Report_Attendance_Hist'," & IIf(chkSSecurity(77), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Comments'," & IIf(chkSSecurity(78), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Employee_Hist'," & IIf(chkSSecurity(79), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Payroll_Transactions'," & IIf(chkSSecurity(80), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_AffirmAction', " & IIf(chkSSecurity(81), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Report_WorkSchedule'," & IIf(chkSSecurity(82), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Applicant_Profile'," & IIf(chkSSecurity(87), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Applicant_Education'," & IIf(chkSSecurity(88), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Training_Plan'," & IIf(chkSSecurity(89), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'not use
'SQLQ = sqlI & "'Report_Profit_Sharing'," & IIf(chkSSecurity(89), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Report_AttWrkSch_Descrepancy'," & IIf(chkSSecurity(90), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Report_Environmental_Serv'," & IIf(chkSSecurity(91), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Report_ESSReq_TransAudit'," & IIf(chkSSecurity(92), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Release 8.0 - Ticket #22682
SQLQ = sqlI & "'Report_Employee_Dates'," & IIf(chkSSecurity(93), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Length_of_Service'," & IIf(chkSSecurity(94), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Ticket #24663
SQLQ = sqlI & "'Form_Attendance_SignIn'," & IIf(chkSSecurity(95), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Form_ATT_Discipline'," & IIf(chkSSecurity(96), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Form_COC_Discipline'," & IIf(chkSSecurity(97), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Ticket #26576 - WDGPHU - Flex Time report
SQLQ = sqlI & "'Report_FlexTime'," & IIf(chkSSecurity(98), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

 SQLQ = sqlI & "'Report_Friesens_IWantToKnowYou', " & IIf(chkSSecurity(61), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'Report_Friesens_ITHireForm', " & IIf(chkSSecurity(62), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'Report_Friesens_ITNoticeOfChange', " & IIf(chkSSecurity(63), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'Report_Friesens_NoticeOfChange', " & IIf(chkSSecurity(64), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'Report_Friesens_PerfImproveActionPlan', " & IIf(chkSSecurity(65), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'Report_Friesens_PerformanceReviewRpt', " & IIf(chkSSecurity(66), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'Report_Friesens_SeparationRpt', " & IIf(chkSSecurity(67), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'Report_Friesens_TerminationRpt', " & IIf(chkSSecurity(68), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'Report_Friesens_UpdateMeetingRpt', " & IIf(chkSSecurity(69), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'Report_Friesens_WarningRpt', " & IIf(chkSSecurity(70), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ


'Course Admin - Begin
 SQLQ = sqlI & "'CA_Report_Calendar', " & IIf(chkSSecurity(36), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'CA_Report_Class_List', " & IIf(chkSSecurity(37), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'CA_Report_Waiting_List', " & IIf(chkSSecurity(38), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'CA_Report_Conflict', " & IIf(chkSSecurity(39), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'CA_Report_Course_Catalog', " & IIf(chkSSecurity(40), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'CA_Report_Course_Per_Position', " & IIf(chkSSecurity(41), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'CA_Report_Label', " & IIf(chkSSecurity(42), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'CA_Report_Prerequ_Exception', " & IIf(chkSSecurity(43), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'CA_Report_CourseNotCompleted', " & IIf(chkSSecurity(44), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'CA_Report_Training_Summary', " & IIf(chkSSecurity(45), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
'Course Admin - End
SQLQ = sqlI & "'Attendance_His_MassUpdate'," & IIf(chkMCSecurity(0), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Attendance_MassUpdate'," & IIf(chkMCSecurity(1), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Benefits_MassUpdate'," & IIf(chkMCSecurity(2), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Codes'," & IIf(chkMCSecurity(3), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Education_Seminars_MassUpdate'," & IIf(chkMCSecurity(4), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Other_Entitlements_MassUpdate'," & IIf(chkMCSecurity(5), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Entitlements_MassUpdate'," & IIf(chkMCSecurity(6), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Follow_Ups_MassUpdate'," & IIf(chkMCSecurity(7), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Hrly_Entitlements_MassUpdate'," & IIf(chkMCSecurity(8), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Other_Earnings_MassUpdate'," & IIf(chkMCSecurity(9), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Job_Master_MassUpdate'," & IIf(chkMCSecurity(10), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Salary_MassUpdate'," & IIf(chkMCSecurity(11), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EmployeeNo_MassUpdate'," & IIf(chkMCSecurity(12), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'OvertimeMaster_MassUpdate'," & IIf(chkMCSecurity(13), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Emergency_Leave_MassUpdate'," & IIf(chkMCSecurity(14), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Import_Photo_MassUpdate'," & IIf(chkMCSecurity(15), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
SQLQ = sqlI & "'Work_Schedule_MassUpdate'," & IIf(chkMCSecurity(16), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Ticket #22893 - Security for Year End based on Anniversary Month
'If glbCompSerial = "S/N - 2448W" Then  'For all clients with Security rights
    SQLQ = sqlI & "'YearEnd_AnniversaryMonth_MassUpdate'," & IIf(chkMCSecurity(17), 1, 0) & ")"
    gdbAdoIhr001.Execute SQLQ
'End If

'Release 8.0 - Ticket #24361: Add Email Address import under Mass Updates menu
SQLQ = sqlI & "'EmailLoad_MassUpdate'," & IIf(chkMCSecurity(18), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Release 8.1 - Ticket #27244: Import document Attachment under Mass Updates menu
SQLQ = sqlI & "'ImpAttachment_MassUpdate'," & IIf(chkMCSecurity(19), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

 'Mostafa Attendance Group Code Matrix
 SQLQ = sqlI & "'Attendance_Group_Code_Matrix_Update'," & IIf(chkUMSecurity(41), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'Attendance_Group_Code_Matrix_Inquiry'," & IIf(chkUISecurity(41), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
'Ticket #16189 - Friesens - Job_Files_Attachment_Update and Temp/Cross Training Position
SQLQ = sqlI & "'Job_Files_Attachment_Update'," & IIf(chkUMSecurity(42), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Job_Files_Attachment_Inquiry'," & IIf(chkUISecurity(42), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'Temp_Cross_Training_Update'," & IIf(chkUMSecurity(43), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Temp_Cross_Training_Inquiry'," & IIf(chkUISecurity(43), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Training_List_Update'," & IIf(chkUMSecurity(44), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Training_List_Inquiry'," & IIf(chkUISecurity(44), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'BudgetedManpower_Update'," & IIf(chkUMSecurity(47), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'BudgetedManpower_Inquiry'," & IIf(chkUISecurity(47), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'WorkScheduleRule_Update'," & IIf(chkUMSecurity(64), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WorkScheduleRule_Inquiry'," & IIf(chkUISecurity(64), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Ticket #22541
SQLQ = sqlI & "'DashboardSetup_Update'," & IIf(chkUMSecurity(65), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'DashboardSetup_Inquiry'," & IIf(chkUISecurity(65), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'RequiredCourses_Update'," & IIf(chkUMSecurity(50), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'RequiredCourses_Inquiry'," & IIf(chkUISecurity(50), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'BudgetedPosition_Update'," & IIf(chkUMSecurity(48), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'BudgetedPosition_Inquiry'," & IIf(chkUISecurity(48), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'ApplicationProcess_Update'," & IIf(chkUMSecurity(49), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'ApplicationProcess_Inquiry'," & IIf(chkUISecurity(49), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Ticket #25015 - Macaulay
SQLQ = sqlI & "'AddPayrollIDData_Update'," & IIf(chkUMSecurity(68), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'AddPayrollIDData_Inquiry'," & IIf(chkUISecurity(68), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ


SQLQ = sqlI & "'Rehire_Update'," & IIf(chkUMSecurity(51), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Rehire_Inquiry'," & IIf(chkUISecurity(51), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EnterLeave_Update'," & IIf(chkUMSecurity(52), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EnterLeave_Inquiry'," & IIf(chkUISecurity(52), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'HS_ClaimMed_Update'," & IIf(chkUMSecurity(37), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'HS_ClaimMed_Inquiry'," & IIf(chkUISecurity(37), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'HS_Contacts_Update'," & IIf(chkUMSecurity(53), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'HS_Contacts_Inquiry'," & IIf(chkUISecurity(53), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'HS_Cost_Update'," & IIf(chkUMSecurity(54), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'HS_Cost_Inquiry'," & IIf(chkUISecurity(54), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'HS_CorrectAction_Update'," & IIf(chkUMSecurity(55), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'HS_CorrectAction_Inquiry'," & IIf(chkUISecurity(55), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'HS_RootCause_Update'," & IIf(chkUMSecurity(56), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'HS_RootCause_Inquiry'," & IIf(chkUISecurity(56), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'AffirmAction_Data_Update'," & IIf(chkUMSecurity(57), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'AffirmAction_Data_Inquiry'," & IIf(chkUISecurity(57), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'AffirmAction_Purge_Update'," & IIf(chkUMSecurity(58), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'AffirmAction_Purge_Inquiry'," & IIf(chkUISecurity(58), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

If glbWSIBModule Then   'WSIB Form 7 - Billable Module
    SQLQ = sqlI & "'HS_W7CompanyMaster_Update'," & IIf(chkUMSecurity(61), 1, 0) & ")"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = sqlI & "'HS_W7CompanyMaster_Inquiry'," & IIf(chkUISecurity(61), 1, 0) & ")"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = sqlI & "'HS_W7Injury_Update'," & IIf(chkUMSecurity(63), 1, 0) & ")"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = sqlI & "'HS_W7Injury_Inquiry'," & IIf(chkUISecurity(63), 1, 0) & ")"
    gdbAdoIhr001.Execute SQLQ

    'Form 9
    SQLQ = sqlI & "'HS_WF9_Update'," & IIf(chkUMSecurity(62), 1, 0) & ")"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = sqlI & "'HS_WF9_Inquiry'," & IIf(chkUISecurity(62), 1, 0) & ")"
    gdbAdoIhr001.Execute SQLQ
End If

'If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #20052 Franks 07/20/2011
'    SQLQ = sqlI & "'Profit_Sharing_Update'," & IIf(chkUMSecurity(62), 1, 0) & ")"
'    gdbAdoIhr001.Execute SQLQ
'    SQLQ = sqlI & "'Profit_Sharing_Inquiry'," & IIf(chkUISecurity(62), 1, 0) & ")"
'    gdbAdoIhr001.Execute SQLQ
'End If

If glbWFC Then
    SQLQ = sqlI & "'RetirementProc_Update'," & IIf(chkUMSecurity(45), 1, 0) & ")"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = sqlI & "'RetirementProc_Inquiry'," & IIf(chkUISecurity(45), 1, 0) & ")"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = sqlI & "'DeathProc_Update'," & IIf(chkUMSecurity(46), 1, 0) & ")"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = sqlI & "'DeathProc_Inquiry'," & IIf(chkUISecurity(46), 1, 0) & ")"
    gdbAdoIhr001.Execute SQLQ
End If
If glbLinamar Then
    SQLQ = sqlI & "'Report_DoorAccess'," & IIf(chkSSecurity(33), 1, 0) & ")"
    gdbAdoIhr001.Execute SQLQ
End If
gdbAdoIhr001.CommitTrans
End Sub

Private Function panVisible()   '10June99 js - selects panel to be shown

panDetails(0).Visible = True
panDetails(1).Visible = False
panDetails(2).Visible = False
panDetails(3).Visible = False
panDetails(4).Visible = False
panDetails(7).Visible = False

elpEEID.SetFocus

End Function

Private Sub ResetAll()
Dim X%, StartTimer As Single

'Utilities
For X% = 0 To 26
    chkUMSecurity(X%).Value = 0
    chkUISecurity(X%).Value = 0
Next X%
chkUMSecurity(38).Value = 0
chkUISecurity(38).Value = 0

chkUMSecurity(60).Value = 0
chkUISecurity(60).Value = 0

If glbWSIBModule Then  'WSIB Form 7 - Billable Module
    chkUMSecurity(61).Value = 0
    chkUISecurity(61).Value = 0
    chkUMSecurity(63).Value = 0
    chkUISecurity(63).Value = 0
    
    chkUMSecurity(62).Value = 0 'Form 9
    chkUISecurity(62).Value = 0 'Form 9
End If

'Ticket #22220, Ticket #22541, Ticket #23409, Ticket #24655,Ticket #22682 - Release 8.0, Ticket #25746, 'Ticket #25922 - OHRS Reporting for CHC
'Ticket #29122 - New Database Setup and Integration Setup securities
For X% = 64 To 73
    chkUMSecurity(X%).Value = 0
    chkUISecurity(X%).Value = 0
Next X%

For X% = 0 To 13
    chkAMSecurity(X%).Value = 0
    chkAISecurity(X%).Value = 0
Next X%

For X% = 0 To 47    '35 '27 '26    '15  '13
    chkUSecurity(X%).Value = 0
Next X%

For X% = 0 To 19
    chkMCSecurity(X%).Value = 0
Next X%


For X% = 0 To 15
    chkIESecurity(X%).Value = 0
Next X%
If glbCompSerial = "S/N - 2382W" Then   'Samuel  - Ticket #19935
    For X% = 16 To 19 '21 '19
        chkIESecurity(X%).Value = 0
    Next X%
End If

For X% = 0 To 29
    chkMSecurity(X%).Value = 0
    chkSecurity(X%).Value = 0
Next X%

For X% = 0 To 35 '33  'laura nov 17 changed from 31 to 32
    chkSSecurity(X%).Value = 0
Next X%

For X% = 27 To 34 'Course Admin
    chkUMSecurity(X%).Value = 0
    chkUISecurity(X%).Value = 0
Next X%

For X% = 36 To 45 'Course Admin
    chkSSecurity(X%).Value = 0
Next X%

For X% = 35 To 59 '37 '35 'ADP Data,Archive_VacTimeoff_Update
    chkUMSecurity(X%).Value = 0
    chkUISecurity(X%).Value = 0
Next X%

For X% = 46 To 98   '47 'Overtime & others
    chkSSecurity(X%).Value = 0
Next X%

chkEESIN.Value = 0
chkEEDOB.Value = 0
chkEEADDRESS.Value = 0

If glbCompSerial = "S/N - 2407W" Then 'Ticket #18406 - Farmers' Mutual Insurance
    chkPswdLocked.Value = 0
End If
'If glbCompSerial = "S/N - 2173W" Then
    chkASecurity.Value = 0
'End If

'Release 8.1
chkACommentSecurity.Value = 0

'Ticket #22682 - Release 8.0
chkNHireSecurity.Value = 0

'Ticket #23923 - Release 8.0 - View Own
chkViewOwnSuccPlan.Value = 0
chkViewOwnComm.Value = 0
chkViewOwnCounsel.Value = 0
chkViewOwnFollUp.Value = 0
chkViewOwnOthInfo.Value = 0
chkViewOwnEmpFlags.Value = 0
chkViewOwnEmpHis.Value = 0
chkViewOwnGLDist.Value = 0
'Ticket #28635 - Add View Own security
chkViewOwnPerform.Value = 0

'Ticket #22009 Franks 05/10/2012
chkDSecurity.Value = 0

cmbTemplate.ListIndex = 0

'Release 8.1
chkSSecurity(99).Value = 0

'Ticket #27795 - Friesens Corporation
chkSSecurity(100).Value = 0
End Sub

Private Sub UpdateRelated()
Dim SQLQ
SQLQ = "UPDATE HRPASDEP SET PD_USERID='" & Replace(txtUSERID, "'", "''") & "' WHERE PD_USERID='" & Replace(OUserID, "'", "''") & "'"
gdbAdoIhr001.Execute SQLQ
SQLQ = "UPDATE HR_SECURE_ACCESS SET USERID='" & Replace(txtUSERID, "'", "''") & "' WHERE USERID='" & Replace(OUserID, "'", "''") & "' AND CODENAME IS NOT NULL"
gdbAdoIhr001.Execute SQLQ
SQLQ = "UPDATE HR_SECRPT SET USERID='" & Replace(txtUSERID, "'", "''") & "' WHERE USERID='" & Replace(OUserID, "'", "''") & "'"
gdbAdoIhr001.Execute SQLQ
SQLQ = "UPDATE HR_EMAIL SET EM_USERID='" & Replace(txtUSERID, "'", "''") & "' WHERE EM_USERID='" & Replace(OUserID, "'", "''") & "'"
gdbAdoIhr001.Execute SQLQ
'Ticket #15928
SQLQ = "UPDATE HR_SECURE_COMMENTS SET USERID='" & Replace(txtUSERID, "'", "''") & "' WHERE USERID='" & Replace(OUserID, "'", "''") & "' AND CODENAME IS NOT NULL"
gdbAdoIhr001.Execute SQLQ

SQLQ = "UPDATE HR_SECURE_FOLLOW_UP SET USERID='" & Replace(txtUSERID, "'", "''") & "' WHERE USERID='" & Replace(OUserID, "'", "''") & "' AND CODENAME IS NOT NULL"
gdbAdoIhr001.Execute SQLQ

SQLQ = "UPDATE HR_SECURE_ATTENDANCE SET USERID='" & Replace(txtUSERID, "'", "''") & "' WHERE USERID='" & Replace(OUserID, "'", "''") & "' AND CODENAME IS NOT NULL"
gdbAdoIhr001.Execute SQLQ

'Release 8.1
SQLQ = "UPDATE HR_SECURE_DOCUMENT_TYPE SET USERID='" & Replace(txtUSERID, "'", "''") & "' WHERE USERID='" & Replace(OUserID, "'", "''") & "' AND CODENAME IS NOT NULL"
gdbAdoIhr001.Execute SQLQ

If glbLinamar Then
    SQLQ = "UPDATE LN_SECURE_ACCESS SET USERID='" & Replace(txtUSERID, "'", "''") & "' WHERE USERID='" & Replace(OUserID, "'", "''") & "'"
    gdbAdoIhr001.Execute SQLQ
End If

'Ticket #30508 - Applicant Tracking Enhancement
SQLQ = "UPDATE HRA_SECURE_REQUISITION SET USERID='" & Replace(txtUSERID, "'", "''") & "' WHERE USERID='" & Replace(OUserID, "'", "''") & "'"
End Sub

Private Sub SECWRK()

Dim SQLQ, xField As String, X
Dim xQue
Dim rsSEC As New ADODB.Recordset
Dim rsFun As New ADODB.Recordset
Dim rsSECWrk As New ADODB.Recordset

SQLQ = "select * from HR_SECURE_BASIC where USERID='" & Replace(txtUSERID, "'", "''") & "'"
rsSEC.Open SQLQ, gdbAdoIhr001, adOpenStatic

gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "delete from HRSECWRK"
gdbAdoIhr001W.CommitTrans

rsSECWrk.Open "HRSECWRK", gdbAdoIhr001W, adOpenStatic, adLockPessimistic

Do Until rsSEC.EOF
    rsSECWrk.AddNew
    rsSECWrk("EMPNBR") = rsSEC("EMPNBR")
    rsSECWrk("USERID") = rsSEC("USERID")
    xQue = False
    For X = 1 To rsSECWrk.Fields.count - 1
        xField = rsSECWrk.Fields(X).name
        xQue = True
        If UCase(xField) = UCase("Basic_Inquiry") Then xQue = True
        If UCase(xField) = "PS_CHGDATE" Then xQue = False
        If xQue Then
            If glbOracle Then
                rsFun.Open "select * from HR_SECURE_ACCESS where USERID='" & Replace(rsSEC("USERID"), "'", "''") & "' and FUNCTION='" & xField & "'", gdbAdoIhr001, adOpenStatic
            Else
                rsFun.Open "select * from HR_SECURE_ACCESS where USERID='" & Replace(rsSEC("USERID"), "'", "''") & "' and [FUNCTION]='" & xField & "'", gdbAdoIhr001, adOpenStatic
            End If
            If Not rsFun.EOF Then rsSECWrk(xField) = rsFun("Accessable")
            rsFun.Close
        End If
    Next
    rsSECWrk("WRKEMP") = glbUserID
    rsSECWrk.Update
    rsSEC.MoveNext
Loop
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fglbNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Security
End Property

Public Property Get Addable() As Boolean
Addable = True
End Property

Public Property Get Updateble() As Boolean
Updateble = True
End Property

Public Property Get Deleteble() As Boolean
Deleteble = True
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum

If fglbNew Then
    UpdateState = NewRecord
    TF = True
    vbxTrueGrid.Enabled = False
ElseIf Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
    vbxTrueGrid.Enabled = True
End If

Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False

Call mod_UpdateMode(TF)

End Sub

Private Sub ChkCBoxChange(xTemplate)

Dim rsCH As New ADODB.Recordset
Dim X%, SQLQ
Dim xAccessable
Dim xTemplateEmpNoSec As Integer

'????Ticket #24808 - Retrieve the Template Profile if the User's Security is based on Template, to see if User's Security has changed, otherwise retrieve Normal User or Template Profile itself
SQLQ = "SELECT " & Field_SQL("FUNCTION") & ", ACCESSABLE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND Maintainable=0"
rsCH.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

ChangeCBox = False

Do Until rsCH.EOF
    If glbOracle Then
        If rsCH("ACCESSABLE") = 0 Then
            xAccessable = False
        Else
            xAccessable = True
        End If
    Else
        xAccessable = rsCH("ACCESSABLE")
    End If
    
    If UCase(rsCH("FUNCTION")) = UCase("Company_Update") And xAccessable <> chkUMSecurity(0) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Company_Inquiry") And xAccessable <> chkUISecurity(0) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Security_Update") And xAccessable <> chkUMSecurity(1) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Security_Inquiry") And xAccessable <> chkUISecurity(1) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Departments_Update") And xAccessable <> chkUMSecurity(3) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Department_Inquiry") And xAccessable <> chkUISecurity(3) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Audit_Update") And xAccessable <> chkUMSecurity(4) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Audit_Inquiry") And xAccessable <> chkUISecurity(4) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("EmploymentEQT_Update") And xAccessable <> chkUMSecurity(5) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("EmploymentEQT_Inquiry") And xAccessable <> chkUISecurity(5) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Divisions_Update") And xAccessable <> chkUMSecurity(6) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Divisions_Inquiry") And xAccessable <> chkUISecurity(6) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Ledgers_Update") And xAccessable <> chkUMSecurity(7) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Ledgers_Inquiry") And xAccessable <> chkUISecurity(7) Then GoTo TheEnd
    If glbLinamar Then
        If UCase(rsCH("FUNCTION")) = UCase("DoorAccess_Update") Then chkUMSecurity(8) = rsCH("ACCESSABLE")
        If UCase(rsCH("FUNCTION")) = UCase("DoorAccess_Inquiry") Then chkUISecurity(8) = rsCH("ACCESSABLE")
        If UCase(rsCH("FUNCTION")) = UCase("ProductLine_Operation_Update") Then chkUMSecurity(20) = rsCH("ACCESSABLE")
        If UCase(rsCH("FUNCTION")) = UCase("ProductLine_Operation_Inquiry") Then chkUISecurity(20) = rsCH("ACCESSABLE")
    End If
    If glbCompSerial = "S/N - 2380W" Then ' For VitalAire Canada Inc. Ticket #26233 Franks 11/20/2014
        If UCase(rsCH("FUNCTION")) = UCase("DoorAccess_Update") Then chkUMSecurity(8) = rsCH("ACCESSABLE")
        If UCase(rsCH("FUNCTION")) = UCase("DoorAccess_Inquiry") Then chkUISecurity(8) = rsCH("ACCESSABLE")
    End If
    'Ticket #12204
    If UCase(rsCH("FUNCTION")) = UCase("CourseCodeMaster_Update") Then chkUMSecurity(20) = rsCH("ACCESSABLE")
    If UCase(rsCH("FUNCTION")) = UCase("CourseCodeMaster_Inquiry") Then chkUISecurity(20) = rsCH("ACCESSABLE")
        
    'Ticket #23409 - Samuel, Son & Co., Limited - Discipline Audit Table Report
    If glbCompSerial = "S/N - 2382W" Then
        If UCase(rsCH("FUNCTION")) = UCase("CounselAudit_Update") And xAccessable <> chkUMSecurity(66) Then GoTo TheEnd
        If UCase(rsCH("FUNCTION")) = UCase("CounselAudit_Inquiry") And xAccessable <> chkUISecurity(66) Then GoTo TheEnd
    End If
    
    'Ticket #24655 - Wellington-Dufferin-Guelph Public Health - On Call Hours
    If glbCompSerial = "S/N - 2411W" Then
        If UCase(rsCH("FUNCTION")) = UCase("On_Call_Hours_Update") And xAccessable <> chkUMSecurity(67) Then GoTo TheEnd
        If UCase(rsCH("FUNCTION")) = UCase("On_Call_Hours_Inquiry") And xAccessable <> chkUISecurity(67) Then GoTo TheEnd
    End If
    
    If UCase(rsCH("FUNCTION")) = UCase("CustomReport_Update") And xAccessable <> chkUMSecurity(9) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("CustomReport_Inquiry") And xAccessable <> chkUISecurity(9) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Holiday_Update") And xAccessable <> chkUMSecurity(10) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Holiday_Inquiry") And xAccessable <> chkUISecurity(10) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("New_Hire_Update") And xAccessable <> chkUMSecurity(11) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("New_Hire_Inquiry") And xAccessable <> chkUISecurity(11) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Label_Update") And xAccessable <> chkUMSecurity(12) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Label_Inquiry") And xAccessable <> chkUISecurity(12) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Sal_Distribute_Update") And xAccessable <> chkUMSecurity(13) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Sal_Distribute_Inquiry") And xAccessable <> chkUISecurity(13) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("Pay_Period_Update") And xAccessable <> chkUMSecurity(19) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Pay_Period_Inquiry") And xAccessable <> chkUISecurity(19) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("Email_Setup_Update") And xAccessable <> chkUMSecurity(18) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Email_Setup_Inquiry") And xAccessable <> chkUISecurity(18) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("Payroll_Category_Update") And xAccessable <> chkUMSecurity(14) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Payroll_Category_Inquiry") And xAccessable <> chkUISecurity(14) Then GoTo TheEnd
    
    'Ticket #25746 - Town of St. Marys
    If UCase(rsCH("FUNCTION")) = UCase("DeptGL_Matrix_Update") And xAccessable <> chkUMSecurity(70) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("DeptGL_Matrix_Inquiry") And xAccessable <> chkUISecurity(70) Then GoTo TheEnd

    If UCase(rsCH("FUNCTION")) = UCase("Charge_Code_Update") And xAccessable <> chkUMSecurity(15) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Charge_Code_Inquiry") And xAccessable <> chkUISecurity(15) Then GoTo TheEnd

    If UCase(rsCH("FUNCTION")) = UCase("Project_Code_Update") And xAccessable <> chkUMSecurity(16) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Project_Code_Inquiry") And xAccessable <> chkUISecurity(16) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Machine_Update") And xAccessable <> chkUMSecurity(17) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Machine_Inquiry") And xAccessable <> chkUISecurity(17) Then GoTo TheEnd
    '7.6
    If UCase(rsCH("FUNCTION")) = UCase("EMP_FLAGS_Update") And xAccessable <> chkUMSecurity(22) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("EMP_FLAGS_Inquiry") And xAccessable <> chkUISecurity(22) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("EMP_HISTORY_Update") And xAccessable <> chkUMSecurity(23) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("EMP_HISTORY_Inquiry") And xAccessable <> chkUISecurity(23) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("GL_DIST_Update") And xAccessable <> chkUMSecurity(24) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("GL_DIST_Inquiry") And xAccessable <> chkUISecurity(24) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("EMP_LANG_Update") And xAccessable <> chkUMSecurity(25) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("EMP_LANG_Inquiry") And xAccessable <> chkUISecurity(25) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("EMP_SUCCESSION_Update") And xAccessable <> chkUMSecurity(26) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("EMP_SUCCESSION_Inquiry") And xAccessable <> chkUISecurity(26) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Emergency_Contacts_Update") And xAccessable <> chkUMSecurity(35) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Emergency_Contacts_Inquiry") And xAccessable <> chkUISecurity(35) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("Work_Schedule_Update") And xAccessable <> chkUMSecurity(60) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Work_Schedule_Inquiry") And xAccessable <> chkUISecurity(60) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("Province") And xAccessable <> chkUSecurity(0) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Entitle") And xAccessable <> chkUSecurity(1) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Compress_Fix") And xAccessable <> chkUSecurity(2) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Matrix") And xAccessable <> chkUSecurity(3) Then GoTo TheEnd
    
    If glbLinamar Then
        If UCase(rsCH("FUNCTION")) = UCase("DoorName") Then chkUSecurity(4) = rsCH("ACCESSABLE")
        If UCase(rsCH("FUNCTION")) = UCase("Summarize_Attendance") Then chkUSecurity(5) = rsCH("ACCESSABLE")
    End If
    If UCase(rsCH("FUNCTION")) = UCase("TimeSheetPrority") And xAccessable <> chkUSecurity(6) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("TimeSheetUserPrority") And xAccessable <> chkUSecurity(7) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("DeleteTimeSheetFile") And xAccessable <> chkUSecurity(8) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ArchiveTimeSheetFile") And xAccessable <> chkUSecurity(25) Then GoTo TheEnd
     If UCase(rsCH("FUNCTION")) = UCase("ApproveTimeSheetFile") And xAccessable <> chkUSecurity(26) Then GoTo TheEnd
    'EssCompTime
    'If UCase(rsCH("FUNCTION")) = UCase("EssCompTime") And xAccessable <> chkUSecurity(26) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("TSOverrideEmpSecurity") And xAccessable <> chkUSecurity(14) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("DeleteApprovedTimeSheet") And xAccessable <> chkUSecurity(15) Then GoTo TheEnd
    
    '7.9 Enhancement
    If UCase(rsCH("FUNCTION")) = UCase("ESS_Time_Req") And xAccessable <> chkUSecurity(18) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_Vacation_Req") And xAccessable <> chkUSecurity(19) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_Request_Approval") And xAccessable <> chkUSecurity(20) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_Rpt_Request_Approval") And xAccessable <> chkUSecurity(21) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_Rpt_Print_Archive") And xAccessable <> chkUSecurity(22) Then GoTo TheEnd
    'If UCase(rsCH("FUNCTION")) = UCase("ESS_Archive_Req") And xAccessable <> chkUSecurity(23) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Archive_VacTimeoff_Update") And xAccessable <> chkUSecurity(23) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_MassDelete_Req") And xAccessable <> chkUSecurity(24) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_ShowAllRequests") And xAccessable <> chkUSecurity(27) Then GoTo TheEnd
    'TS_PUNCHINOUT
    If UCase(rsCH("FUNCTION")) = UCase("TS_PUNCHINOUT") And xAccessable <> chkUSecurity(28) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_REQ_OTHER_SUPER") And xAccessable <> chkUSecurity(29) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("ESS_DEL_APPR_TIME_REQ") And xAccessable <> chkUSecurity(30) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_DEL_APPR_VAC_REQ") And xAccessable <> chkUSecurity(31) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("ESS_CANCEL_REQ") And xAccessable <> chkUSecurity(32) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("ESS_TS_LIST_RA_ONLY") And xAccessable <> chkUSecurity(33) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_DELETE_FUTURE_VAC_REQ") And xAccessable <> chkUSecurity(34) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_DELETE_FUTURE_TIME_REQ") And xAccessable <> chkUSecurity(35) Then GoTo TheEnd
        
    If UCase(rsCH("FUNCTION")) = UCase("ESS_HTML_CALENDAR") And xAccessable <> chkUSecurity(36) Then GoTo TheEnd
    'Ticket #23536 - Dashboard ON/OFF
    If UCase(rsCH("FUNCTION")) = UCase("ESS_DASHBOARDS") And xAccessable <> chkUSecurity(37) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_QUICKINFO") And xAccessable <> chkUSecurity(38) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_DEMO_MAINTAIN") And xAccessable <> chkUSecurity(39) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_SHOWALLAPPREJ_REQS") And xAccessable <> chkUSecurity(40) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("TS_SUBMISSION") And xAccessable <> chkUSecurity(41) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_REQAPPDEPTSEC") And xAccessable <> chkUSecurity(42) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_APPVIEWOWN") And xAccessable <> chkUSecurity(43) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ESS_MYCOWORKER") And xAccessable <> chkUSecurity(44) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("TS_ENABLE_SUPER") And xAccessable <> chkUSecurity(45) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("TS_RA_ONLY") And xAccessable <> chkUSecurity(46) Then GoTo TheEnd
    
                 
    If UCase(rsCH("FUNCTION")) = UCase("CompanyPreference") And xAccessable <> chkUSecurity(9) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("EmpFlagsSetup") And xAccessable <> chkUSecurity(10) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("MultiDataSourceSetup") And xAccessable <> chkUSecurity(11) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HelpDescSetup") And xAccessable <> chkUSecurity(12) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("BenefitGroupSetup") And xAccessable <> chkUSecurity(13) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ChangeYourPassword") And xAccessable <> chkUSecurity(16) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ITAdmin") And xAccessable <> chkUSecurity(17) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("Import_Attendance") And xAccessable <> chkIESecurity(0) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Export_Attendance") And xAccessable <> chkIESecurity(1) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Import_Benefits") And xAccessable <> chkIESecurity(2) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Export_Benefits") And xAccessable <> chkIESecurity(3) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Import_Employee") And xAccessable <> chkIESecurity(4) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Export_Employee") And xAccessable <> chkIESecurity(5) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Import_Salaries") And xAccessable <> chkIESecurity(6) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Export_Salaries") And xAccessable <> chkIESecurity(7) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Import_Table") And xAccessable <> chkIESecurity(8) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Export_Table") And xAccessable <> chkIESecurity(9) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Import_YTD") And xAccessable <> chkIESecurity(10) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Export_YTD") And xAccessable <> chkIESecurity(11) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Import_PayrollTrans") And xAccessable <> chkIESecurity(12) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Export_PayrollTrans") And xAccessable <> chkIESecurity(13) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Import_ContEdu") And xAccessable <> chkIESecurity(14) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Export_ContEdu") And xAccessable <> chkIESecurity(15) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Import_PerfReview") And xAccessable <> chkIESecurity(16) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Export_PerfReview") And xAccessable <> chkIESecurity(17) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Import_EmploymentEquity") And xAccessable <> chkIESecurity(18) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Export_EmploymentEquity") And xAccessable <> chkIESecurity(19) Then GoTo TheEnd
    'Not use - Samuel Ticket #21000 Franks 09/26/2011 - move to Custom Security
    'If UCase(rsCH("FUNCTION")) = UCase("Import_Profit_Sharing") And xAccessable <> chkIESecurity(20) Then GoTo TheEnd
    'If UCase(rsCH("FUNCTION")) = UCase("Export_Profit_Sharing") And xAccessable <> chkIESecurity(21) Then GoTo TheEnd
        
    If UCase(rsCH("FUNCTION")) = UCase("Basic_Update") And xAccessable <> chkMSecurity(0) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Basic_Inquiry") And xAccessable <> chkSecurity(0) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Banking_Update") And xAccessable <> chkMSecurity(1) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Banking_Inquiry") And xAccessable <> chkSecurity(1) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Dependents_Update") And xAccessable <> chkMSecurity(2) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Dependents_Inquiry") And xAccessable <> chkSecurity(2) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Skills_Update") And xAccessable <> chkMSecurity(3) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Skills_Inquiry") And xAccessable <> chkSecurity(3) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Formal_Education_Update") And xAccessable <> chkMSecurity(4) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Formal_Education_Inquiry") And xAccessable <> chkSecurity(4) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Salary_Update") And xAccessable <> chkMSecurity(5) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Salary_Inquiry") And xAccessable <> chkSecurity(5) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Performance_Update") And xAccessable <> chkMSecurity(6) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Performance_Inquiry") And xAccessable <> chkSecurity(6) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Position_Update") And xAccessable <> chkMSecurity(7) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Position_Inquiry") And xAccessable <> chkSecurity(7) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Benefits_Update") And xAccessable <> chkMSecurity(8) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Benefits_Inquiry") And xAccessable <> chkSecurity(8) Then GoTo TheEnd
    
    '7.9 Enhancement
    If UCase(rsCH("FUNCTION")) = UCase("Beneficiary_Update") And xAccessable <> chkMSecurity(29) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Beneficiary_Inquiry") And xAccessable <> chkSecurity(29) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("Entitlements_Update") And xAccessable <> chkMSecurity(9) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Entitlements_Inquiry") And xAccessable <> chkSecurity(9) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Associations_Update") And xAccessable <> chkMSecurity(10) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Associations_Inquiry") And xAccessable <> chkSecurity(10) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Follow_Ups_Update") And xAccessable <> chkMSecurity(11) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Follow_Ups_Inquiry") And xAccessable <> chkSecurity(11) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Health_Safety_Update") And xAccessable <> chkMSecurity(12) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Health_Safety_Inquiry") And xAccessable <> chkSecurity(12) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Attendance_Update") And xAccessable <> chkMSecurity(13) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Attendance_Inquiry") And xAccessable <> chkSecurity(13) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Attendance_History_Update") And xAccessable <> chkMSecurity(28) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Attendance_History_Inquiry") And xAccessable <> chkSecurity(28) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Other_Entitlements_Update") And xAccessable <> chkMSecurity(14) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Other_Entitlements_Inquiry") And xAccessable <> chkSecurity(14) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Other_Earnings_Update") And xAccessable <> chkMSecurity(15) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Other_Earnings_Inquiry") And xAccessable <> chkSecurity(15) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Terminations_Update") And xAccessable <> chkMSecurity(16) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Termination_Inquiry") And xAccessable <> chkSecurity(16) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Job_Classes_Update") And xAccessable <> chkMSecurity(17) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Job_Classes_Inquiry") And xAccessable <> chkSecurity(17) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Education_Seminars_Update") And xAccessable <> chkMSecurity(18) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Education_Seminars_Inquiry") And xAccessable <> chkSecurity(18) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Job_Skills_Update") And xAccessable <> chkMSecurity(19) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Job_Skills_Inquiry") And xAccessable <> chkSecurity(19) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Job_Eval_Update") And xAccessable <> chkMSecurity(20) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Job_Eval_Inquiry") And xAccessable <> chkSecurity(20) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Job_Master_Update") And xAccessable <> chkMSecurity(21) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Job_Master_Inquiry") And xAccessable <> chkSecurity(21) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Hrly_Entitlements_Update") And xAccessable <> chkMSecurity(22) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Hrly_Entitlements_Inquiry") And xAccessable <> chkSecurity(22) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Show_SIN_SSN") And xAccessable <> chkEESIN Then GoTo TheEnd
   'tkt10423 jerry said make it available to everyone
   ' If glbCompSerial = "S/N - 2173W" Then
        If UCase(rsCH("FUNCTION")) = UCase("Add_Attendance") And xAccessable <> chkASecurity Then GoTo TheEnd
    'End If
        
    'Ticket #22682 - Release 8.0
    If UCase(rsCH("FUNCTION")) = UCase("Add_NewHire") And xAccessable <> chkNHireSecurity Then GoTo TheEnd
    
    'Release 8.1
    If UCase(rsCH("FUNCTION")) = UCase("Add_Comments") And xAccessable <> chkACommentSecurity Then GoTo TheEnd
    
    'Ticket #23923 - Release 8.0 - View Own
    If UCase(rsCH("FUNCTION")) = UCase("ScsPlan_ViewOwn") And xAccessable <> chkViewOwnSuccPlan Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Comments_ViewOwn") And xAccessable <> chkViewOwnComm Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Counsel_ViewOwn") And xAccessable <> chkViewOwnCounsel Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("FollUp_ViewOwn") And xAccessable <> chkViewOwnFollUp Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("OthInfo_ViewOwn") And xAccessable <> chkViewOwnOthInfo Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("EmpFlags_ViewOwn") And xAccessable <> chkViewOwnEmpFlags Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("EmpHis_ViewOwn") And xAccessable <> chkViewOwnEmpHis Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("GLDist_ViewOwn") And xAccessable <> chkViewOwnGLDist Then GoTo TheEnd
        
    'Ticket #22009 Franks 05/10/2012
    If UCase(rsCH("FUNCTION")) = UCase("Del_Dependents") And xAccessable <> chkDSecurity Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("Counselling_Update") And xAccessable <> chkMSecurity(23) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Counselling_Inquiry") And xAccessable <> chkSecurity(23) Then GoTo TheEnd
    
    If glbWFC Then
        If UCase(rsCH("FUNCTION")) = UCase("SalaryGrids_Update") Then chkMSecurity(24) = rsCH("ACCESSABLE")
        If UCase(rsCH("FUNCTION")) = UCase("SalaryGrids_Inquiry") Then chkSecurity(24) = rsCH("ACCESSABLE")
    End If
    If UCase(rsCH("FUNCTION")) = UCase("Comments_Update") And xAccessable <> chkMSecurity(25) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Comments_Inquiry") And xAccessable <> chkSecurity(25) Then GoTo TheEnd
        
    If UCase(rsCH("FUNCTION")) = UCase("OtherInformation_Update") And xAccessable <> chkMSecurity(26) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("OtherInformation_Inquiry") And xAccessable <> chkSecurity(26) Then GoTo TheEnd
             
    If UCase(rsCH("FUNCTION")) = UCase("LinamarSkills_Update") And xAccessable <> chkMSecurity(27) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("LinamarSkills_Inquiry") And xAccessable <> chkSecurity(27) Then GoTo TheEnd
   
    'Attendance Code Matrix
    If UCase(rsCH("FUNCTION")) = UCase("AttendCode_Matrix_Update") And xAccessable <> chkUMSecurity(59) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("AttendCode_Matrix_Inquiry") And xAccessable <> chkUISecurity(59) Then GoTo TheEnd
    
    'Ticket #22682 - Release 8.0 - Follow Up Code Email Matrix
    If UCase(rsCH("FUNCTION")) = UCase("FollowUpCodeEmail_Matrix_Update") And xAccessable <> chkUMSecurity(69) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("FollowUpCodeEmail_Matrix_Inquiry") And xAccessable <> chkUISecurity(69) Then GoTo TheEnd
    
    'Ticket #25922 - OHRS Reporting for CHC
    If UCase(rsCH("FUNCTION")) = UCase("OHRS_Department_Update") And xAccessable <> chkUMSecurity(71) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("OHRS_Department_Inquiry") And xAccessable <> chkUISecurity(71) Then GoTo TheEnd

    'ADP Data
    If UCase(rsCH("FUNCTION")) = UCase("ADP_Data_Update") And xAccessable <> chkUMSecurity(36) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ADP_Data_Inquiry") And xAccessable <> chkUISecurity(36) Then GoTo TheEnd
    
    'Sam Added for ESS.NET 07/27/2006 Ticket # 11403
    'If UCase(rsCH("FUNCTION")) = UCase("Archive_VacTimeoff_Update") And xAccessable <> chkUMSecurity(37) Then GoTo TheEnd
    'If UCase(rsCH("FUNCTION")) = UCase("Archive_VacTimeoff_Inquiry") And xAccessable <> chkUISecurity(37) Then GoTo TheEnd
    'ends
    
    If UCase(rsCH("FUNCTION")) = UCase("UserDefineTbl_Update") And xAccessable <> chkUMSecurity(39) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("UserDefineTbl_Inquiry") And xAccessable <> chkUISecurity(39) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("PayrollTrans_Update") And xAccessable <> chkUMSecurity(40) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("PayrollTrans_Inquiry") And xAccessable <> chkUISecurity(40) Then GoTo TheEnd
    
    'Course Code Master
    If UCase(rsCH("FUNCTION")) = UCase("CourseCodeMaster_Update") And xAccessable <> chkUMSecurity(38) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("CourseCodeMaster_Inquiry") And xAccessable <> chkUISecurity(38) Then GoTo TheEnd

    If UCase(rsCH("FUNCTION")) = UCase("BudgetedManpower_Update") And xAccessable <> chkUMSecurity(47) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("BudgetedManpower_Inquiry") And xAccessable <> chkUISecurity(47) Then GoTo TheEnd

    If UCase(rsCH("FUNCTION")) = UCase("WorkScheduleRule_Update") And xAccessable <> chkUMSecurity(64) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("WorkScheduleRule_Inquiry") And xAccessable <> chkUISecurity(64) Then GoTo TheEnd
    
    'Ticket #22541
    If UCase(rsCH("FUNCTION")) = UCase("DashboardSetup_Update") And xAccessable <> chkUMSecurity(65) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("DashboardSetup_Inquiry") And xAccessable <> chkUISecurity(65) Then GoTo TheEnd

    If UCase(rsCH("FUNCTION")) = UCase("RequiredCourses_Update") And xAccessable <> chkUMSecurity(50) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("RequiredCourses_Inquiry") And xAccessable <> chkUISecurity(50) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("BudgetedPosition_Update") And xAccessable <> chkUMSecurity(48) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("BudgetedPosition_Inquiry") And xAccessable <> chkUISecurity(48) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ApplicationProcess_Update") And xAccessable <> chkUMSecurity(49) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("ApplicationProcess_Inquiry") And xAccessable <> chkUISecurity(49) Then GoTo TheEnd
    
    'Ticket #25015 - Macaulay
    If UCase(rsCH("FUNCTION")) = UCase("AddPayrollIDData_Update") And xAccessable <> chkUMSecurity(68) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("AddPayrollIDData_Inquiry") And xAccessable <> chkUISecurity(68) Then GoTo TheEnd

    If UCase(rsCH("FUNCTION")) = UCase("Rehire_Update") And xAccessable <> chkUMSecurity(51) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Rehire_Inquiry") And xAccessable <> chkUISecurity(51) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("EnterLeave_Update") And xAccessable <> chkUMSecurity(52) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("EnterLeave_Inquiry") And xAccessable <> chkUISecurity(52) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HS_ClaimMed_Update") And xAccessable <> chkUMSecurity(37) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HS_ClaimMed_Inquiry") And xAccessable <> chkUISecurity(37) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HS_Contacts_Update") And xAccessable <> chkUMSecurity(53) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HS_Contacts_Inquiry") And xAccessable <> chkUISecurity(53) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HS_Cost_Update") And xAccessable <> chkUMSecurity(54) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HS_Cost_Inquiry") And xAccessable <> chkUISecurity(54) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HS_CorrectAction_Update") And xAccessable <> chkUMSecurity(55) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HS_CorrectAction_Inquiry") And xAccessable <> chkUISecurity(55) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HS_RootCause_Update") And xAccessable <> chkUMSecurity(56) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HS_RootCause_Inquiry") And xAccessable <> chkUISecurity(56) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("AffirmAction_Data_Update") And xAccessable <> chkUMSecurity(57) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("AffirmAction_Data_Inquiry") And xAccessable <> chkUISecurity(57) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("AffirmAction_Purge_Update") And xAccessable <> chkUMSecurity(58) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("AffirmAction_Purge_Inquiry") And xAccessable <> chkUISecurity(58) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HS_W7CompanyMaster_Update") And xAccessable <> chkUMSecurity(61) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HS_W7CompanyMaster_Inquiry") And xAccessable <> chkUISecurity(61) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HS_W7Injury_Update") And xAccessable <> chkUMSecurity(63) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HS_W7Injury_Inquiry") And xAccessable <> chkUISecurity(63) Then GoTo TheEnd
    
    'Form 9
    If UCase(rsCH("FUNCTION")) = UCase("HS_WF9_Update") And xAccessable <> chkUMSecurity(62) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("HS_WF9_Inquiry") And xAccessable <> chkUISecurity(62) Then GoTo TheEnd
    
    'If UCase(rsCH("FUNCTION")) = UCase("Profit_Sharing_Update") And xAccessable <> chkUMSecurity(62) Then GoTo TheEnd
    'If UCase(rsCH("FUNCTION")) = UCase("Profit_Sharing_Inquiry") And xAccessable <> chkUISecurity(62) Then GoTo TheEnd

    If UCase(rsCH("FUNCTION")) = UCase("App_Basic_Update") And xAccessable <> chkAMSecurity(0) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Basic_Inquiry") And xAccessable <> chkAISecurity(0) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Skills_Update") And xAccessable <> chkAMSecurity(1) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Skills_Inquiry") And xAccessable <> chkAISecurity(1) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Formal_Education_Update") And xAccessable <> chkAMSecurity(2) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Formal_Education_Inquiry") And xAccessable <> chkAISecurity(2) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Education_Seminars_Update") And xAccessable <> chkAMSecurity(3) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Education_Seminars_Inquiry") And xAccessable <> chkAISecurity(3) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Interview_Update") And xAccessable <> chkAMSecurity(4) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Interview_Inquiry") And xAccessable <> chkAISecurity(4) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Associations_Update") And xAccessable <> chkAMSecurity(5) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Associations_Inquiry") And xAccessable <> chkAISecurity(5) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_References_Update") And xAccessable <> chkAMSecurity(6) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_References_Inquiry") And xAccessable <> chkAISecurity(6) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Follow_Ups_Update") And xAccessable <> chkAMSecurity(7) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Follow_Ups_Inquiry") And xAccessable <> chkAISecurity(7) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Requisition_Update") And xAccessable <> chkAMSecurity(8) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Requisition_Inquiry") And xAccessable <> chkAISecurity(8) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Recruitment_Update") And xAccessable <> chkAMSecurity(9) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_Recruitment_Inquiry") And xAccessable <> chkAISecurity(9) Then GoTo TheEnd
    'Ticket #30508 - Applicant Tracking Enhancement
    If UCase(rsCH("FUNCTION")) = UCase("App_LetterPosType_Update") And xAccessable <> chkAMSecurity(11) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_LetterPosType_Inquiry") And xAccessable <> chkAISecurity(11) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_FormWorkflow_Update") And xAccessable <> chkAMSecurity(12) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_FormWorkflow_Inquiry") And xAccessable <> chkAISecurity(12) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_FormDefaults_Update") And xAccessable <> chkAMSecurity(13) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("App_FormDefaults_Inquiry") And xAccessable <> chkAISecurity(13) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("Report_Age") And xAccessable <> chkSSecurity(0) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Compensatory_Time") And xAccessable <> chkSSecurity(1) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Cost_Of_Employment") And xAccessable <> chkSSecurity(2) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Emergecy_Contacts") And xAccessable <> chkSSecurity(3) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Employee_Labels") And xAccessable <> chkSSecurity(4) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Job_List") And xAccessable <> chkSSecurity(5) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Profiles") And xAccessable <> chkSSecurity(6) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Entitlements") And xAccessable <> chkSSecurity(7) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Follow_Ups") And xAccessable <> chkSSecurity(8) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Home_Address") And xAccessable <> chkSSecurity(9) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Salary_Performance") And xAccessable <> chkSSecurity(10) Then GoTo TheEnd
    'Ticket #27795 - Friesens Corporation
    If UCase(rsCH("FUNCTION")) = UCase("Report_Staff_Profile") And xAccessable <> chkSSecurity(100) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Seniority") And xAccessable <> chkSSecurity(11) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Telephone_Extensions") And xAccessable <> chkSSecurity(12) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Associations") And xAccessable <> chkSSecurity(13) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Master_Attendance") And xAccessable <> chkSSecurity(14) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Bonus_Attendance") And xAccessable <> chkSSecurity(83) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Calendar_Attendance") And xAccessable <> chkSSecurity(84) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Costed_Attendance") And xAccessable <> chkSSecurity(85) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Master_Benefits") And xAccessable <> chkSSecurity(15) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Master_Division") And xAccessable <> chkSSecurity(16) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Master_Termination") And xAccessable <> chkSSecurity(17) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Master_Formal_Education") And xAccessable <> chkSSecurity(18) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Master_Job") And xAccessable <> chkSSecurity(19) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Master_Passwords") And xAccessable <> chkSSecurity(20) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Master_Salaries") And xAccessable <> chkSSecurity(21) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Master_Edu_Seminars") And xAccessable <> chkSSecurity(22) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Master_Table_Codes") And xAccessable <> chkSSecurity(23) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Heatlh_Safety") And xAccessable <> chkSSecurity(24) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Master_DolEnt") And xAccessable <> chkSSecurity(25) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Hourly_Entitlements") And xAccessable <> chkSSecurity(26) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Master_OtherEarn") And xAccessable <> chkSSecurity(27) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("PayEQT_Inquiry") And xAccessable <> chkSSecurity(28) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Dependents") And xAccessable <> chkSSecurity(29) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Skills") And xAccessable <> chkSSecurity(30) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Languages") And xAccessable <> chkSSecurity(31) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Employee_Turnover") And xAccessable <> chkSSecurity(32) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Counselling") And xAccessable <> chkSSecurity(34) Then GoTo TheEnd
    'Release 8.1
    If UCase(rsCH("FUNCTION")) = UCase("Report_DocumentType") And xAccessable <> chkSSecurity(99) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Emergency_Leave") And xAccessable <> chkSSecurity(35) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_External_Hire") And xAccessable <> chkSSecurity(48) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Internal_Hire") And xAccessable <> chkSSecurity(49) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Key_Workforce") And xAccessable <> chkSSecurity(50) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Manpower_Plan") And xAccessable <> chkSSecurity(51) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Staff_Management") And xAccessable <> chkSSecurity(52) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_WC_Time") And xAccessable <> chkSSecurity(53) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_WC_Work") And xAccessable <> chkSSecurity(54) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Paid_Sick") And xAccessable <> chkSSecurity(55) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_User_Defined_Table") And xAccessable <> chkSSecurity(56) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Future_Entitlement") And xAccessable <> chkSSecurity(57) Then GoTo TheEnd
    
    'Overtime
    If UCase(rsCH("FUNCTION")) = UCase("Report_Overtime_Bank") And xAccessable <> chkSSecurity(46) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Overtime_Lost_Hours") And xAccessable <> chkSSecurity(47) Then GoTo TheEnd
    
    'More reports
    If UCase(rsCH("FUNCTION")) = UCase("Report_Employee_Flags") And xAccessable <> chkSSecurity(58) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Temp_CrossTraining") And xAccessable <> chkSSecurity(59) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Required_Course_Hist") And xAccessable <> chkSSecurity(60) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("Report_Email_Address") And xAccessable <> chkSSecurity(71) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_LOA") And xAccessable <> chkSSecurity(73) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_POE") And xAccessable <> chkSSecurity(74) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_SINSSN") And xAccessable <> chkSSecurity(75) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Succession") And xAccessable <> chkSSecurity(76) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Gap_Analysis") And xAccessable <> chkSSecurity(72) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("Report_GL_Distribution") And xAccessable <> chkSSecurity(86) Then GoTo TheEnd
        
    If UCase(rsCH("FUNCTION")) = UCase("Report_Attendance_Hist") And xAccessable <> chkSSecurity(77) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Comments") And xAccessable <> chkSSecurity(78) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Employee_Hist") And xAccessable <> chkSSecurity(79) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Payroll_Transactions") And xAccessable <> chkSSecurity(80) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_AffirmAction") And xAccessable <> chkSSecurity(81) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("Report_WorkSchedule") And xAccessable <> chkSSecurity(82) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_AttWrkSch_Descrepancy") And xAccessable <> chkSSecurity(90) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Environmental_Serv") And xAccessable <> chkSSecurity(91) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_ESSReq_TransAudit") And xAccessable <> chkSSecurity(92) Then GoTo TheEnd
    
    'Release 8.0 - Ticket #22682
    If UCase(rsCH("FUNCTION")) = UCase("Report_Employee_Dates") And xAccessable <> chkSSecurity(93) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Length_of_Service") And xAccessable <> chkSSecurity(94) Then GoTo TheEnd
    
    'Ticket #24663
    If UCase(rsCH("FUNCTION")) = UCase("Form_Attendance_SignIn") And xAccessable <> chkSSecurity(95) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Form_ATT_Discipline") And xAccessable <> chkSSecurity(96) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Form_COC_Discipline") And xAccessable <> chkSSecurity(97) Then GoTo TheEnd
    
    'Ticket #26576 - WDGPHU - Flex Time report
    If UCase(rsCH("FUNCTION")) = UCase("Report_FlexTime") And xAccessable <> chkSSecurity(98) Then GoTo TheEnd
    
    If UCase(rsCH("FUNCTION")) = UCase("Report_Friesens_IWantToKnowYou") And xAccessable <> chkSSecurity(61) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Friesens_ITHireForm") And xAccessable <> chkSSecurity(62) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Friesens_ITNoticeOfChange") And xAccessable <> chkSSecurity(63) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Friesens_NoticeOfChange") And xAccessable <> chkSSecurity(64) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Friesens_PerfImproveActionPlan") And xAccessable <> chkSSecurity(65) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Friesens_PerformanceReviewRpt") And xAccessable <> chkSSecurity(66) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Friesens_SeparationRpt") And xAccessable <> chkSSecurity(67) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Friesens_TerminationRpt") And xAccessable <> chkSSecurity(68) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Friesens_UpdateMeetingRpt") And xAccessable <> chkSSecurity(69) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Report_Friesens_WarningRpt") And xAccessable <> chkSSecurity(70) Then GoTo TheEnd
    
    'Course Admin - Begin
    If UCase(rsCH("FUNCTION")) = UCase("CA_Report_Calendar") And xAccessable <> chkSSecurity(36) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("CA_Report_Class_List") And xAccessable <> chkSSecurity(37) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("CA_Report_Waiting_List") And xAccessable <> chkSSecurity(38) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("CA_Report_Conflict") And xAccessable <> chkSSecurity(39) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("CA_Report_Course_Catalog") And xAccessable <> chkSSecurity(40) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("CA_Report_Course_Per_Position") And xAccessable <> chkSSecurity(41) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("CA_Report_Label") And xAccessable <> chkSSecurity(42) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("CA_Report_Prerequ_Exception") And xAccessable <> chkSSecurity(43) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("CA_Report_CourseNotCompleted") And xAccessable <> chkSSecurity(44) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("CA_Report_Training_Summary") And xAccessable <> chkSSecurity(45) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Master_Table_Update") And xAccessable <> chkUMSecurity(2) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Master_Table_Inquiry") And xAccessable <> chkUISecurity(2) Then GoTo TheEnd
    'Course Admin - End
    
    If UCase(rsCH("FUNCTION")) = UCase("Attendance_His_MassUpdate") And xAccessable <> chkMCSecurity(0) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Attendance_MassUpdate") And xAccessable <> chkMCSecurity(1) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Benefits_MassUpdate") And xAccessable <> chkMCSecurity(2) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Codes") And xAccessable <> chkMCSecurity(3) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Education_Seminars_MassUpdate") And xAccessable <> chkMCSecurity(4) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Other_Entitlements_MassUpdate") And xAccessable <> chkMCSecurity(5) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Entitlements_MassUpdate") And xAccessable <> chkMCSecurity(6) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Follow_Ups_MassUpdate") And xAccessable <> chkMCSecurity(7) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Hrly_Entitlements_MassUpdate") And xAccessable <> chkMCSecurity(8) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Other_Earnings_MassUpdate") And xAccessable <> chkMCSecurity(9) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Job_Master_MassUpdate") And xAccessable <> chkMCSecurity(10) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Salary_MassUpdate") And xAccessable <> chkMCSecurity(11) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("EmployeeNo_MassUpdate") And xAccessable <> chkMCSecurity(12) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("OvertimeMaster_MassUpdate") And xAccessable <> chkMCSecurity(13) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Emergency_Leave_MassUpdate") And xAccessable <> chkMCSecurity(14) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Import_Photo_MassUpdate") And xAccessable <> chkMCSecurity(15) Then GoTo TheEnd
        
    If UCase(rsCH("FUNCTION")) = UCase("Work_Schedule_MassUpdate") And xAccessable <> chkMCSecurity(16) Then GoTo TheEnd
    
    'Ticket #22893 - Security for Year End based on Anniversary Month
    'If glbCompSerial = "S/N - 2448W" Then   'For all clients with Security rights
        If UCase(rsCH("FUNCTION")) = UCase("YearEnd_AnniversaryMonth_MassUpdate") And xAccessable <> chkMCSecurity(17) Then GoTo TheEnd
    'End If
    
    'Release 8.0 - Ticket #24361: Add Email Address import under Mass Updates menu
    If UCase(rsCH("FUNCTION")) = UCase("EmailLoad_MassUpdate") And xAccessable <> chkMCSecurity(18) Then GoTo TheEnd
    
    'Release 8.1 - Ticket #27244: Import document Attachment under Mass Updates menu
    If UCase(rsCH("FUNCTION")) = UCase("ImpAttachment_MassUpdate") And xAccessable <> chkMCSecurity(19) Then GoTo TheEnd
       
    'Ticket #16189 - Friesens  - Job_Files_Attachment_Update and Temp/Cross Training Position
    If UCase(rsCH("FUNCTION")) = UCase("Job_Files_Attachment_Update") And xAccessable <> chkUMSecurity(42) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Job_Files_Attachment_Inquiry") And xAccessable <> chkUISecurity(42) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Temp_Cross_Training_Update") And xAccessable <> chkUMSecurity(43) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Temp_Cross_Training_Inquiry") And xAccessable <> chkUISecurity(43) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Training_List_Update") And xAccessable <> chkUMSecurity(44) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Training_List_Inquiry") And xAccessable <> chkUISecurity(44) Then GoTo TheEnd
        
    rsCH.MoveNext
Loop
rsCH.Close
Set rsCH = Nothing

'Got error that procedure is too large so broke it down to new procedure
If ChkCBoxChange_1(xTemplate) Then GoTo TheEnd

Exit Sub

TheEnd:
ChangeCBox = True

End Sub

Private Function ChkCBoxChange_1(xTemplate)
Dim xAccessable
Dim xTemplateEmpNoSec As Integer

Dim rsCH As New ADODB.Recordset
Dim X%, SQLQ

ChkCBoxChange_1 = False

'????Ticket #24808 - Retrieve the Template Profile if the User's Security is based on Template, to see if User's Security has changed, otherwise retrieve Normal User or Template Profile itself
SQLQ = "SELECT " & Field_SQL("FUNCTION") & ", ACCESSABLE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND Maintainable=0"
rsCH.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
Do Until rsCH.EOF
    If glbOracle Then
        If rsCH("ACCESSABLE") = 0 Then
            xAccessable = False
        Else
            xAccessable = True
        End If
    Else
        xAccessable = rsCH("ACCESSABLE")
    End If

    'Mostafa Attendance group code matrix
    If UCase(rsCH("FUNCTION")) = UCase("Attendance_Group_Code_Matrix_Update") And xAccessable <> chkUMSecurity(41) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("Attendance_Group_Code_Matrix_Inquiry") And xAccessable <> chkUISecurity(41) Then GoTo TheEnd

    If glbWFC Then 'Ticket #18566
        If UCase(rsCH("FUNCTION")) = UCase("RetirementProc_Update") And xAccessable <> chkUMSecurity(45) Then GoTo TheEnd
        If UCase(rsCH("FUNCTION")) = UCase("RetirementProc_Inquiry") And xAccessable <> chkUISecurity(45) Then GoTo TheEnd
        If UCase(rsCH("FUNCTION")) = UCase("DeathProc_Update") And xAccessable <> chkUMSecurity(46) Then GoTo TheEnd
        If UCase(rsCH("FUNCTION")) = UCase("DeathProc_Inquiry") And xAccessable <> chkUISecurity(46) Then GoTo TheEnd
    End If

    If glbLinamar Then
        If UCase(rsCH("FUNCTION")) = UCase("Report_DoorAccess") Then chkSSecurity(33) = rsCH("ACCESSABLE")
    End If

    'Ticket #29122 - New Database Setup and Integration Setup securities
    If UCase(rsCH("FUNCTION")) = UCase("IntegrtDBSetup_Update") And xAccessable <> chkUMSecurity(72) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("IntegrtDBSetup_Inquiry") And xAccessable <> chkUISecurity(72) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("IntegrtSetup_Update") And xAccessable <> chkUMSecurity(73) Then GoTo TheEnd
    If UCase(rsCH("FUNCTION")) = UCase("IntegrtSetup_Inquiry") And xAccessable <> chkUISecurity(73) Then GoTo TheEnd

    'Ticket #28635 - Add View Own security
    If UCase(rsCH("FUNCTION")) = UCase("Perform_ViewOwn") And xAccessable <> chkViewOwnPerform Then GoTo TheEnd

    'Ticket #28795 - Mostafa added this line in the above procedure which started giving Procedure too large error - moved it to this procedure.
    If UCase(rsCH("FUNCTION")) = UCase("ESS_TS_SETUP_CHECKLIST") And xAccessable <> chkUSecurity(47) Then GoTo TheEnd
    
    rsCH.MoveNext
Loop
rsCH.Close
Set rsCH = Nothing

'Ticket #24320 - Check if Employee # Based Security changed
If Len(xTemplate) > 0 And xTemplate <> "TEMPLATE" Then
    xTemplateEmpNoSec = EmployeeNoBasedSecurity(xTemplate)
    If xTemplateEmpNoSec = 0 Then
        xAccessable = False
    ElseIf xTemplateEmpNoSec = 1 Or xTemplateEmpNoSec = 2 Then
        xAccessable = True
    End If
    If xAccessable <> chkEESecurity Then GoTo TheEnd
End If

Exit Function

TheEnd:
ChkCBoxChange_1 = True

End Function

Private Function CountryList() As String
Dim xCountryList As String, ctyFile
xCountryList = ""
ctyFile = glbIHRREPORTS & "CountryList.MTF"

On Error GoTo ErrorHandler

If File(ctyFile) Then
    Open ctyFile For Input As #1
    Input #1, xCountryList
    Close #1
End If

ResumeHere:
'If InStr(xCountryList, BasicCountry) = 0 Then
'    xCountryList = BasicCountry
'End If
If InStr(xCountryList, "ALL&") = 0 Then
    xCountryList = "ALL&" & xCountryList
End If
If InStr(xCountryList, cmbCountry) = 0 And cmbCountry <> "" Then
    xCountryList = xCountryList & "&" & cmbCountry
    cmbCountry.AddItem cmbCountry
'    comCountryOfEmp.AddItem cmbCountry
End If

Open ctyFile For Output As #1
Print #1, xCountryList
Close #1

CountryList = xCountryList

Exit Function

ErrorHandler:
If Err.Number = 62 Then
    ' Corrupted CountryList.MTF, kill it and regenerate
    Close #1
    MsgBox "Found corrupt CountryList.MTF.  info:HR will re-create this file.", vbInformation + vbOKOnly, "Corrupted Country List"
    Kill ctyFile
    Resume ResumeHere
Else
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number & " in CountryList"
    Resume Next
End If
End Function

Sub cmdCancel_Click()
Dim bk

On Error GoTo Can_Err

Data1.Recordset.CancelUpdate

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

vbxTrueGrid.Enabled = True

Data1.Refresh

fglbNew% = False
newNew = False

Call mod_UpdateMode(True)  ' reset screen's attributes

panEEDESC.Enabled = True

Call displaypanel   '10June99 js - displays panel selected

cmdCopySecuritys.Enabled = True

Call SET_UP_MODE

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRSECURE", "Cancel")
Call RollBack   '10June99 js

End Sub

Public Sub UpdateTSTemplate(UserID, templateid)
    
gdbAdoIhr001.Execute "UPDATE HR_SECURE_BASIC SET TS_TPID=" & templateid & " WHERE USERID='" & Replace(UserID, "'", "''") & "'"

End Sub

Public Sub FindTemplate()
    Dim intX
    Dim found
    found = False
    
    If Len(lblTemplate.Caption) > 0 And cmbTemplate.ListCount > 1 Then
    
        For intX = 1 To cmbTemplate.ListCount - 1
        
             If Trim(Split(cmbTemplate.List(intX))(0)) = lblTemplate.Caption Then
                cmbTemplate.ListIndex = intX
                'cmbTemplate.Refresh
                found = True
                Exit For
             End If
        Next
    End If
    
    If Not found Then
        cmbTemplate.ListIndex = 0
    End If
End Sub

Private Sub Populate_Security_Template()
    Dim SQLQ As String
    Dim rsSecBasic As New ADODB.Recordset
    
    'Ticket #20585 - Template based Security. Populate the combo box with Templates.
    
    'Clear the combo box
    cmbSecTemplate.Clear
    
    'Not all users will have a template or needs to have a template
    cmbSecTemplate.AddItem ""
    
    'Add Default value "TEMPLATE". This value is to be chosen when templates are created
    cmbSecTemplate.AddItem "TEMPLATE"
    
    
    'Populate User Template combo box list
    SQLQ = "SELECT USERID, SECURE_TEMPLATE FROM HR_SECURE_BASIC WHERE SECURE_TEMPLATE = 'TEMPLATE'"
    SQLQ = SQLQ & " ORDER BY SECURE_TEMPLATE"
    rsSecBasic.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsSecBasic.EOF
        If Not IsNull(rsSecBasic("USERID")) Then
            cmbSecTemplate.AddItem rsSecBasic("USERID")
        End If
        rsSecBasic.MoveNext
    Loop
    rsSecBasic.Close
    Set rsSecBasic = Nothing
    
End Sub

Public Sub Template_Based_Security_Profile_Update(xUserID, xTemplate, xUpdType)
Dim X, sUSERID
Dim SQLQ As String
Dim rsSecBasicTemp As New ADODB.Recordset
Dim rsSecAccess As New ADODB.Recordset
Dim rsINSERT As New ADODB.Recordset
Dim rsSecTemplate As New ADODB.Recordset
Dim xExpDate
Dim xTemUpt 'Ticket #23787 Franks 05/21/2013
Dim flgExpDate As Boolean

On Error GoTo Err_ExpiryDate
    
    '????Ticket #24808 - User's Profile will only be updated with Template Profile if the Template itself is getting deleted 'Delete'
    '????Ticket #24808 - User's Profile will only be updated up to Employee # Based Security, Password Expiry and Department Security if new User
    'Ticket #20585 - Security Based on Template Profile
    
    If xTemplate = "" Then Exit Sub
    
    'Retrieve Security record of the Template - Employee # based Security only
    SQLQ = "SELECT EmpNBR_Based, PS_EXPIR_DAYS FROM HR_SECURE_BASIC WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    rsSecBasicTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsSecBasicTemp.EOF Then
        MsgBox "This Template has no Security Profile setup.", vbOKOnly, "Error finding Security Profile"
        rsSecBasicTemp.Close
        Set rsSecBasicTemp = Nothing
        Exit Sub
    Else
        'Update User's Security record based on the Template Security record - Employee # based Security only
        'SQLQ = "UPDATE HR_SECURE_BASIC SET EmpNBR_Based = '" & rsSecBasicTemp("EmpNBR_Based") & "' WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
        'Ticket #23787 Franks 05/21/2013
        If IsNull(rsSecBasicTemp("EmpNBR_Based")) Then
            xTemUpt = 0
        Else
            If rsSecBasicTemp("EmpNBR_Based") Then xTemUpt = 1 Else xTemUpt = 0
        End If
        'Ticket #24320 - The following update statement was commented by Frank above to fix an error but he then
        'forgot add the Update statement with the correct syntax. So adding the SQLQ = Update... below.
        SQLQ = "UPDATE HR_SECURE_BASIC SET EmpNBR_Based = " & xTemUpt & " WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
        gdbAdoIhr001.Execute SQLQ
        
        'Ticket #22077 - Begin: Update Expiry Days based on Template Security record,
        'and compute the Expiry Date
        'Update Expiry Days
        SQLQ = "UPDATE HR_SECURE_BASIC SET PS_EXPIR_DAYS = '" & IIf(IsNull(rsSecBasicTemp("PS_EXPIR_DAYS")), 0, rsSecBasicTemp("PS_EXPIR_DAYS")) & "' WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
        gdbAdoIhr001.Execute SQLQ

        'Compute Expiry Date based on Expiry Days
        If IsNumeric(rsSecBasicTemp("PS_EXPIR_DAYS")) Then
            If rsSecBasicTemp("PS_EXPIR_DAYS") <> 0 Then
                xExpDate = ""
                'Ticket #22893 - This is added to avoid the error 'Adding a value to a 'datetime' column caused an overflow'
                'xExpDate = DateAdd("d", rsSecBasicTemp("PS_EXPIR_DAYS"), Date)
                flgExpDate = True
                xExpDate = IIf(IsDate(DateAdd("d", rsSecBasicTemp("PS_EXPIR_DAYS"), Date)), DateAdd("d", rsSecBasicTemp("PS_EXPIR_DAYS"), Date), CVDate(Format("12/31/9999", "mm/dd/yyyy")))
                flgExpDate = False

                'Update the database with Expiry Date
                If IsDate(xExpDate) Then
                    'Ticket #22893 - This is added to avoid the error 'Adding a value to a 'datetime' column caused an overflow'
                    'Check if valid date will be computed then only proceed with the update otherwise update with
                    'upper limit of the date to avoid the error.
                    If Validate_ExpiryDate(xUserID, rsSecBasicTemp("PS_EXPIR_DAYS")) Then
                        'SQLQ = "UPDATE HR_SECURE_BASIC SET PS_EXPIR_DATE = " & Date_SQL(xExpDate) & " WHERE USERID='" & Replace(xUserID, "'", "'+chr(39)+'") & "'"
                        SQLQ = "UPDATE HR_SECURE_BASIC SET PS_EXPIR_DATE = (CASE WHEN PS_EXPIR_DATE IS NULL THEN " & Date_SQL(xExpDate) & " ELSE DATEADD(DAY," & rsSecBasicTemp("PS_EXPIR_DAYS") & ", PS_EXPIR_DATE) END ) WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
                    Else
                        'Ticket #22893 - Update with upper limit for a valid date
                        SQLQ = "UPDATE HR_SECURE_BASIC SET PS_EXPIR_DATE = " & Date_SQL(CVDate(Format("12/31/9999", "mm/dd/yyyy"))) & " WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
                    End If

                    gdbAdoIhr001.Execute SQLQ
                End If
            End If
        End If
        'Ticket #22077 - End
        
    End If
    rsSecBasicTemp.Close
    Set rsSecBasicTemp = Nothing
    
    
    '????Ticket #24808 - Only if Template is getting Deleted
    'Retrieve Template Security Profile from HR_SECURE_ACCESS
    SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    rsSecTemplate.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsSecTemplate.EOF Then
        MsgBox "This Template has no Security Profile setup.", vbOKOnly, "Error finding Security Profile"
        rsSecTemplate.Close
        Set rsSecTemplate = Nothing
        Exit Sub
    Else
        'Delete User's Profile first and then add back based on Template Profile
        SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
        gdbAdoIhr001.Execute SQLQ
        
        
        '????Ticket #24808 - Template getting deleted so save User's record with template Security Profile
        If xUpdType = "Delete" Then
            'Open User's Security record to add back
            SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            rsSecAccess.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
        End If
    End If
    
    MDIMain.panHelp(0).Caption = "Please wait while system updates security profile..."
    

    '????Ticket #24808 - Template getting deleted so save User's record with template Security Profile
    If xUpdType = "Delete" Then

        'Add the user back using the template profile
        Do While Not rsSecTemplate.EOF
             rsSecAccess.AddNew
             rsSecAccess("USERID") = xUserID
             rsSecAccess("FUNCTION") = rsSecTemplate("FUNCTION")
             rsSecAccess("ACCESSABLE") = rsSecTemplate("ACCESSABLE")
             rsSecAccess("Maintainable") = rsSecTemplate("Maintainable")
             rsSecAccess("CODENAME") = rsSecTemplate("CODENAME")
             rsSecAccess("LDATE") = Date
             rsSecAccess("LTIME") = Time$
             rsSecAccess("LUSER") = glbUserID
    
             rsSecAccess.Update
             rsSecTemplate.MoveNext
        Loop
        rsSecAccess.Close
        Set rsSecAccess = Nothing
        
    End If
    rsSecTemplate.Close
    Set rsSecTemplate = Nothing
    
    
    'Department Security --------------------------------------------------------------------------------------------
    'Ticket #21629 - Jerry said not to change/update the Department Security of the user for existing
    'users but for New Users it can take Template's Department Security
    If xUpdType = "Add" Then
        'Add the Department Security
        Dim rsFrmSecDept As New ADODB.Recordset
        Dim rsToSecDept As New ADODB.Recordset
    
        'Retrieve Template's Department Security
        SQLQ = "SELECT * FROM HRPASDEP WHERE PD_USERID='" & Replace(xTemplate, "'", "''") & "'"
        rsFrmSecDept.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    
        'Delete User's Department Security first and then add back based on Template
        SQLQ = "DELETE FROM HRPASDEP WHERE PD_USERID='" & Replace(xUserID, "'", "''") & "'"
        gdbAdoIhr001.Execute SQLQ
    
        'Open User's Department Security record
        SQLQ = "SELECT * FROM HRPASDEP WHERE PD_USERID='" & Replace(xUserID, "'", "''") & "'"
        rsToSecDept.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    
    
        'Add User's Dept Security based on the Template
        Do While Not rsFrmSecDept.EOF
            rsToSecDept.AddNew
            rsToSecDept("PD_COMPNO") = "001"
            rsToSecDept("PD_USERID") = xUserID
            rsToSecDept("PD_DEPT") = rsFrmSecDept("PD_DEPT")
            rsToSecDept("PD_ORG") = rsFrmSecDept("PD_ORG")
            rsToSecDept("PD_DIV") = rsFrmSecDept("PD_DIV")
            rsToSecDept("PD_SECTION") = rsFrmSecDept("PD_SECTION")
            rsToSecDept("PD_ADMINBY") = rsFrmSecDept("PD_ADMINBY")
            'Ticket #22682 - Release 8.0
            rsToSecDept("PD_LOC") = rsFrmSecDept("PD_LOC")
            rsToSecDept("PD_REGION") = rsFrmSecDept("PD_REGION")
            
            rsToSecDept("PD_INCLEMPNBR") = rsFrmSecDept("PD_INCLEMPNBR")
            rsToSecDept("PD_EXCLEMPNBR") = rsFrmSecDept("PD_EXCLEMPNBR")
            rsToSecDept.Update
    
            rsFrmSecDept.MoveNext
        Loop
        rsToSecDept.Close
        Set rsToSecDept = Nothing
        rsFrmSecDept.Close
        Set rsFrmSecDept = Nothing
    End If
    '-----------------------------------------------------------------------------------------------------------------
    
    'Ticket #30508 - Applicant Tracking Enhancement
    'Requisition Security --------------------------------------------------------------------------------------------
        'Ticket #21629 - Jerry said not to change/update the Requisition Security of the user for existing
        'users but for New Users it can take Template's Department Security
    If xUpdType = "Add" Then
        'Add the Requisition Security
        Dim rsFrmSecRequi As New ADODB.Recordset
        Dim rsToSecRequi As New ADODB.Recordset
    
        'Retrieve Template's Requisition Security
        SQLQ = "SELECT * FROM HRA_SECURE_REQUISITION WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
        rsFrmSecRequi.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    
        'Delete User's Requisition Security first and then add back based on Template
        SQLQ = "DELETE FROM HRA_SECURE_REQUISITION WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
        gdbAdoIhr001.Execute SQLQ
    
        'Open User's Requisition Security record
        SQLQ = "SELECT * FROM HRA_SECURE_REQUISITION WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
        rsToSecRequi.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    
    
        'Add User's Requisition Security based on the Template
        Do While Not rsFrmSecRequi.EOF
            rsToSecRequi.AddNew
            rsToSecRequi("COMPNO") = "001"
            rsToSecRequi("USERID") = xUserID
            rsToSecRequi("RS_POSTYPE") = rsFrmSecRequi("RS_POSTYPE")
            rsToSecRequi("RS_ORG") = rsFrmSecRequi("RS_ORG")
            rsToSecRequi("RS_GRPCD") = rsFrmSecRequi("RS_GRPCD")
            rsToSecRequi("RS_STATUS") = rsFrmSecRequi("RS_STATUS")
                        
            rsToSecRequi("RS_INCLJOB") = rsFrmSecRequi("RS_INCLJOB")
            rsToSecRequi("RS_EXCLJOB") = rsFrmSecRequi("RS_EXCLJOB")
            
            rsToSecRequi("LDATE") = Date
            rsToSecRequi("LTIME") = Time$
            rsToSecRequi("LUSER") = glbUserID
            
            rsToSecRequi.Update
    
            rsFrmSecRequi.MoveNext
        Loop
        rsToSecRequi.Close
        Set rsToSecRequi = Nothing
        rsFrmSecRequi.Close
        Set rsFrmSecRequi = Nothing
    End If
    '-----------------------------------------------------------------------------------------------------------------
    
    'Comments Security -----------------------------------------------------------------------------------------------
    '????Ticket #24808 - Only if Template is getting Deleted

    'Add the Comments Security
    Dim rsFrmSecComments As New ADODB.Recordset
    Dim rsToSecComments As New ADODB.Recordset
    
    'Retrieve Template's Comments Security
    SQLQ = "SELECT * FROM HR_SECURE_COMMENTS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    rsFrmSecComments.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    
    'Delete User's Comments Security first and then add back based on Template
    SQLQ = "DELETE FROM HR_SECURE_COMMENTS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
    gdbAdoIhr001.Execute SQLQ
    
    '????Ticket #24808 - Template getting deleted so save User's record with template Security Profile
    If xUpdType = "Delete" Then
    
        'Open User's Department Security record
        SQLQ = "SELECT * FROM HR_SECURE_COMMENTS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
        rsToSecComments.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    
    
        'Add User's Comments Security based on the Template
        Do While Not rsFrmSecComments.EOF
            rsToSecComments.AddNew
            rsToSecComments("COMPNO") = "001"
            rsToSecComments("USERID") = xUserID
            rsToSecComments("ACCESSABLE") = rsFrmSecComments("ACCESSABLE")
            rsToSecComments("MAINTAINABLE") = rsFrmSecComments("MAINTAINABLE")
            rsToSecComments("CODENAME") = rsFrmSecComments("CODENAME")
            rsToSecComments("DESCRIPTION") = rsFrmSecComments("DESCRIPTION")
            rsToSecComments("LDATE") = Date
            rsToSecComments("LTIME") = Time$
            rsToSecComments("LUSER") = glbUserID
            rsToSecComments.Update
            
            rsFrmSecComments.MoveNext
        Loop
        rsToSecComments.Close
        Set rsToSecComments = Nothing
        
    End If
    rsFrmSecComments.Close
    Set rsFrmSecComments = Nothing
    '-----------------------------------------------------------------------------------------------------------------
    
    'Custom Reports Security -----------------------------------------------------------------------------------------
    'Add the Custom Reports Security
    Dim rsFrmSecCustmRpt As New ADODB.Recordset
    Dim rsToSecCustmRpt As New ADODB.Recordset
    
    'Retrieve Template's Custom Reports Security
    SQLQ = "SELECT * FROM HR_SECRPT WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    rsFrmSecCustmRpt.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    
    'Delete User's Custom Reports Security first and then add back based on Template
    SQLQ = "DELETE FROM HR_SECRPT WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
    gdbAdoIhr001.Execute SQLQ
    
    '????Ticket #24808 - Template getting deleted so save User's record with template Security Profile
    If xUpdType = "Delete" Then
    
        'Open User's Custom Reports Security record
        SQLQ = "SELECT * FROM HR_SECRPT WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
        rsToSecCustmRpt.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
        
        
        'Add User's Custom Reports Security based on the Template
        Do While Not rsFrmSecCustmRpt.EOF
            rsToSecCustmRpt.AddNew
            rsToSecCustmRpt("COMPNO") = "001"
            rsToSecCustmRpt("USERID") = xUserID
            rsToSecCustmRpt("FUNCTION") = rsFrmSecCustmRpt("FUNCTION")
            rsToSecCustmRpt("ACCESSABLE") = rsFrmSecCustmRpt("ACCESSABLE")
            rsToSecCustmRpt("Maintainable") = rsFrmSecCustmRpt("Maintainable")
            rsToSecCustmRpt("CODENAME") = rsFrmSecCustmRpt("CODENAME")
            rsToSecCustmRpt("LDATE") = Date
            rsToSecCustmRpt("LTIME") = Time$
            rsToSecCustmRpt("LUSER") = glbUserID
            rsToSecCustmRpt.Update
            
            rsFrmSecCustmRpt.MoveNext
        Loop
        rsToSecCustmRpt.Close
        Set rsToSecCustmRpt = Nothing
        
    End If
    rsFrmSecCustmRpt.Close
    Set rsFrmSecCustmRpt = Nothing
    '-----------------------------------------------------------------------------------------------------------------
    
    
    'Linamar Custom Features -----------------------------------------------------------------------------------------
    'Add Linamar's Custom Features Security
    If glbLinamar Then
        Dim rsFrmSecCustmFeat As New ADODB.Recordset
        Dim rsToSecCustmFeat As New ADODB.Recordset
        
        'Retrieve Template's Linamar's Custom Features Security
        SQLQ = "SELECT * FROM LN_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
        rsFrmSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        
        'Delete User's Linamar's Custom Features Security first and then add back based on Template
        SQLQ = "DELETE FROM LN_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
        gdbAdoIhr001.Execute SQLQ
        
        
        '????Ticket #24808 - Template getting deleted so save User's record with template Security Profile
        If xUpdType = "Delete" Then
        
            'Open User's Linamar's Custom Features Security record
            SQLQ = "SELECT * FROM LN_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            rsToSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
                    
            'Add User's Linamar's Custom Feature Security based on the Template
            Do While Not rsFrmSecCustmFeat.EOF
                rsToSecCustmFeat.AddNew
                rsToSecCustmFeat("COMPNO") = "001"
                rsToSecCustmFeat("USERID") = xUserID
                rsToSecCustmFeat("FUNCTION") = rsFrmSecCustmFeat("FUNCTION")
                rsToSecCustmFeat("ACCESSABLE") = rsFrmSecCustmFeat("ACCESSABLE")
                rsToSecCustmFeat("Maintainable") = rsFrmSecCustmFeat("Maintainable")
                rsToSecCustmFeat("CODENAME") = rsFrmSecCustmFeat("CODENAME")
                rsToSecCustmFeat("LDATE") = Date
                rsToSecCustmFeat("LTIME") = Time$
                rsToSecCustmFeat("LUSER") = glbUserID
                rsToSecCustmFeat.Update
                
                rsFrmSecCustmFeat.MoveNext
            Loop
            rsToSecCustmFeat.Close
            Set rsToSecCustmFeat = Nothing
            
        End If
        rsFrmSecCustmFeat.Close
        Set rsFrmSecCustmFeat = Nothing
    End If
    '-----------------------------------------------------------------------------------------------------------------
    
    
    'WHSCC's Custom Features -----------------------------------------------------------------------------------------
    'Add WHSCC's Custom Features Security
    If glbWHSCC Then
        'Retrieve Template's WHSCC's Custom Features Security
        SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
        SQLQ = SQLQ & " AND LEFT([FUNCTION],4)='WHSC'"
        rsFrmSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        
        'Delete User's WHSCC's Custom Features Security first and then add back based on Template
        SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
        SQLQ = SQLQ & " AND LEFT([FUNCTION],4)='WHSC'"
        gdbAdoIhr001.Execute SQLQ
        
        '????Ticket #24808 - Template getting deleted so save User's record with template Security Profile
        If xUpdType = "Delete" Then
        
            'Open User's WHSCC's Custom Features Security record
            SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            SQLQ = SQLQ & " AND LEFT([FUNCTION],4)='WHSC'"
            rsToSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
                    
            'Add User's WHSCC's Custom Feature Security based on the Template
            Do While Not rsFrmSecCustmFeat.EOF
                rsToSecCustmFeat.AddNew
                rsToSecCustmFeat("COMPNO") = "001"
                rsToSecCustmFeat("USERID") = xUserID
                rsToSecCustmFeat("FUNCTION") = rsFrmSecCustmFeat("FUNCTION")
                rsToSecCustmFeat("ACCESSABLE") = rsFrmSecCustmFeat("ACCESSABLE")
                rsToSecCustmFeat("Maintainable") = rsFrmSecCustmFeat("Maintainable")
                rsToSecCustmFeat("CODENAME") = rsFrmSecCustmFeat("CODENAME")
                rsToSecCustmFeat("LDATE") = Date
                rsToSecCustmFeat("LTIME") = Time$
                rsToSecCustmFeat("LUSER") = glbUserID
                rsToSecCustmFeat.Update
                
                rsFrmSecCustmFeat.MoveNext
            Loop
            rsToSecCustmFeat.Close
            Set rsToSecCustmFeat = Nothing
        
        End If
        rsFrmSecCustmFeat.Close
        Set rsFrmSecCustmFeat = Nothing
    End If
    '----------------------------------------------------------------------------------------------------------------
    
    
    'WFC Custom Features --------------------------------------------------------------------------------------------
    'Add WFC's Custom Features Security
    If glbWFC Then
        'Retrieve Template's WFC's Custom Features Security
        SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
        SQLQ = SQLQ & " AND (LEFT([FUNCTION],7)='WFCPEN_' OR LEFT([FUNCTION],4)='WFC_')"
        rsFrmSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                
        'Delete User's WFC's Custom Features Security first and then add back based on Template
        SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
        SQLQ = SQLQ & " AND (LEFT([FUNCTION],7)='WFCPEN_'"
        SQLQ = SQLQ & " OR LEFT([FUNCTION],4)='WFC_')"
        gdbAdoIhr001.Execute SQLQ
        
        
        '????Ticket #24808 - Template getting deleted so save User's record with template Security Profile
        If xUpdType = "Delete" Then
        
            'Open User's WFC's Custom Features Security record
            SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            SQLQ = SQLQ & " AND (LEFT([FUNCTION],7)='WFCPEN_'"
            SQLQ = SQLQ & " OR LEFT([FUNCTION],4)='WFC_')"
            rsToSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
                    
            'Add User's WFC's Custom Feature Security based on the Template
            Do While Not rsFrmSecCustmFeat.EOF
                rsToSecCustmFeat.AddNew
                rsToSecCustmFeat("COMPNO") = "001"
                rsToSecCustmFeat("USERID") = xUserID
                rsToSecCustmFeat("FUNCTION") = rsFrmSecCustmFeat("FUNCTION")
                rsToSecCustmFeat("ACCESSABLE") = rsFrmSecCustmFeat("ACCESSABLE")
                rsToSecCustmFeat("Maintainable") = rsFrmSecCustmFeat("Maintainable")
                rsToSecCustmFeat("CODENAME") = rsFrmSecCustmFeat("CODENAME")
                rsToSecCustmFeat("LDATE") = Date
                rsToSecCustmFeat("LTIME") = Time$
                rsToSecCustmFeat("LUSER") = glbUserID
                rsToSecCustmFeat.Update
                
                rsFrmSecCustmFeat.MoveNext
            Loop
            rsToSecCustmFeat.Close
            Set rsToSecCustmFeat = Nothing
            
        End If
        rsFrmSecCustmFeat.Close
        Set rsFrmSecCustmFeat = Nothing
    End If
    '----------------------------------------------------------------------------------------------------------------
    
    
    'Samuel's Custom Features ---------------------------------------------------------------------------------------
    'Add Samuel's Custom Features Security
    If glbCompSerial = "S/N - 2382W" Then
        'Retrieve Template's Samuel's Custom Features Security
        SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
        SQLQ = SQLQ & " AND LEFT([FUNCTION],4)='SAM_'"
        rsFrmSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        
        'Delete User's Samuel's Custom Features Security first and then add back based on Template
        SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
        SQLQ = SQLQ & " AND LEFT([FUNCTION],4)='SAM_'"
        gdbAdoIhr001.Execute SQLQ
        
        '????Ticket #24808 - Template getting deleted so save User's record with template Security Profile
        If xUpdType = "Delete" Then
        
            'Open User's Samuel's Custom Features Security record
            SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            SQLQ = SQLQ & " AND LEFT([FUNCTION],4)='SAM_'"
            rsToSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
                    
            'Add User's Samuel's Custom Feature Security based on the Template
            Do While Not rsFrmSecCustmFeat.EOF
                rsToSecCustmFeat.AddNew
                rsToSecCustmFeat("COMPNO") = "001"
                rsToSecCustmFeat("USERID") = xUserID
                rsToSecCustmFeat("FUNCTION") = rsFrmSecCustmFeat("FUNCTION")
                rsToSecCustmFeat("ACCESSABLE") = rsFrmSecCustmFeat("ACCESSABLE")
                rsToSecCustmFeat("Maintainable") = rsFrmSecCustmFeat("Maintainable")
                rsToSecCustmFeat("CODENAME") = rsFrmSecCustmFeat("CODENAME")
                rsToSecCustmFeat("LDATE") = Date
                rsToSecCustmFeat("LTIME") = Time$
                rsToSecCustmFeat("LUSER") = glbUserID
                rsToSecCustmFeat.Update
                
                rsFrmSecCustmFeat.MoveNext
            Loop
            rsToSecCustmFeat.Close
            Set rsToSecCustmFeat = Nothing
            
        End If
        rsFrmSecCustmFeat.Close
        Set rsFrmSecCustmFeat = Nothing
    End If
    '---------------------------------------------------------------------------------------------------------------
    
    
    'Follow Up Code Security ---------------------------------------------------------------------------------------
    'Add Follow Up Security
    Dim rsFrmSecFollowUp As New ADODB.Recordset
    Dim rsToSecFollowUp As New ADODB.Recordset
    
    'Retrieve Template's Follow Up Security
    SQLQ = "SELECT * FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    rsFrmSecFollowUp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    
    'Delete User's Follow Up Security first and then add back based on Template
    SQLQ = "DELETE FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
    gdbAdoIhr001.Execute SQLQ
    
    '????Ticket #24808 - Template getting deleted so save User's record with template Security Profile
    If xUpdType = "Delete" Then
    
        'Open User's Follow Up Security record
        SQLQ = "SELECT * FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
        rsToSecFollowUp.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
        
        
        'Add User's Follow Up Security based on the Template
        Do While Not rsFrmSecFollowUp.EOF
            rsToSecFollowUp.AddNew
            rsToSecFollowUp("COMPNO") = "001"
            rsToSecFollowUp("USERID") = xUserID
            rsToSecFollowUp("ACCESSABLE") = rsFrmSecFollowUp("ACCESSABLE")
            rsToSecFollowUp("MAINTAINABLE") = rsFrmSecFollowUp("MAINTAINABLE")
            rsToSecFollowUp("CODENAME") = rsFrmSecFollowUp("CODENAME")
            rsToSecFollowUp("DESCRIPTION") = rsFrmSecFollowUp("DESCRIPTION")
            rsToSecFollowUp("LDATE") = Date
            rsToSecFollowUp("LTIME") = Time$
            rsToSecFollowUp("LUSER") = glbUserID
            rsToSecFollowUp.Update
            
            rsFrmSecFollowUp.MoveNext
        Loop
        rsToSecFollowUp.Close
        Set rsToSecFollowUp = Nothing
        
    End If
    rsFrmSecFollowUp.Close
    Set rsFrmSecFollowUp = Nothing
    '----------------------------------------------------------------------------------------------------------------
    
    
    'Attendance Reason Code Security --------------------------------------------------------------------------------
    'Add Attendance Reason Code Security
    Dim rsFrmSecAttend As New ADODB.Recordset
    Dim rsToSecAttend As New ADODB.Recordset
    
    'Retrieve Template's Attendance Reason Code Security
    SQLQ = "SELECT * FROM HR_SECURE_ATTENDANCE WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    rsFrmSecAttend.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    
    'Delete User's Attendance Reason Code Security first and then add back based on Template
    SQLQ = "DELETE FROM HR_SECURE_ATTENDANCE WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
    gdbAdoIhr001.Execute SQLQ
    
    '????Ticket #24808 - Template getting deleted so save User's record with template Security Profile
    If xUpdType = "Delete" Then
    
        'Open User's Attendance Reason Code Security record
        SQLQ = "SELECT * FROM HR_SECURE_ATTENDANCE WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
        rsToSecAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
        
        
        'Add User's Attendance Reason Code Security based on the Template
        Do While Not rsFrmSecAttend.EOF
            rsToSecAttend.AddNew
            rsToSecAttend("COMPNO") = "001"
            rsToSecAttend("USERID") = xUserID
            rsToSecAttend("ACCESSABLE") = rsFrmSecAttend("ACCESSABLE")
            rsToSecAttend("MAINTAINABLE") = rsFrmSecAttend("MAINTAINABLE")
            rsToSecAttend("CODENAME") = rsFrmSecAttend("CODENAME")
            rsToSecAttend("DESCRIPTION") = rsFrmSecAttend("DESCRIPTION")
            rsToSecAttend("LDATE") = Date
            rsToSecAttend("LTIME") = Time$
            rsToSecAttend("LUSER") = glbUserID
            rsToSecAttend.Update
            
            rsFrmSecAttend.MoveNext
        Loop
        rsToSecAttend.Close
        Set rsToSecAttend = Nothing
        
    End If
    rsFrmSecAttend.Close
    Set rsFrmSecAttend = Nothing
    '----------------------------------------------------------------------------------------------------------------
    
    
    'Release 8.1
    'Document Type Code Security ---------------------------------------------------------------------------------------
    'Add Document Type Security
    Dim rsFrmSecDocType As New ADODB.Recordset
    Dim rsToSecDocType As New ADODB.Recordset
    
    'Retrieve Template's Document Type Security
    SQLQ = "SELECT * FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    rsFrmSecDocType.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    
    'Delete User's Document Type Security first and then add back based on Template
    SQLQ = "DELETE FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
    gdbAdoIhr001.Execute SQLQ
    
    '????Ticket #24808 - Template getting deleted so save User's record with template Security Profile
    If xUpdType = "Delete" Then
    
        'Open User's Document Type Security record
        SQLQ = "SELECT * FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
        rsToSecDocType.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
        
        
        'Add User's Document Type Security based on the Template
        Do While Not rsFrmSecDocType.EOF
            rsToSecDocType.AddNew
            rsToSecDocType("COMPNO") = "001"
            rsToSecDocType("USERID") = xUserID
            rsToSecDocType("ACCESSABLE") = rsFrmSecDocType("ACCESSABLE")
            rsToSecDocType("MAINTAINABLE") = rsFrmSecDocType("MAINTAINABLE")
            rsToSecDocType("CODENAME") = rsFrmSecDocType("CODENAME")
            rsToSecDocType("DESCRIPTION") = rsFrmSecDocType("DESCRIPTION")
            rsToSecDocType("LDATE") = Date
            rsToSecDocType("LTIME") = Time$
            rsToSecDocType("LUSER") = glbUserID
            rsToSecDocType.Update
            
            rsFrmSecDocType.MoveNext
        Loop
        rsToSecDocType.Close
        Set rsToSecDocType = Nothing
        
    End If
    rsFrmSecDocType.Close
    Set rsFrmSecDocType = Nothing
    '----------------------------------------------------------------------------------------------------------------
    
    
    Screen.MousePointer = DEFAULT
    
    If xUpdType = "Add" Then
        MDIMain.panHelp(0).Caption = "Security Add Done"
        MsgBox "Security Profile added for '" & xUserID & "' successfully.", vbInformation, "Security Added"
    ElseIf xUpdType = "Update" Then
        MDIMain.panHelp(0).Caption = "Security Update Done"
        MsgBox "Security Profile updated for '" & xUserID & "' successfully.", vbInformation, "Security Updated"
    ElseIf xUpdType = "Reset" Then
        MDIMain.panHelp(0).Caption = "Security Reset/Update Done"
        MsgBox "Security Profile has been reset/updated for '" & xUserID & "' successfully based on the '" & xTemplate & "' Security Template.", vbInformation, "Security Reset/Update"
    End If

Exit Sub

Err_ExpiryDate:
    If Err.Number = 13 And flgExpDate Then
        xExpDate = CVDate(Format("12/31/9999", "mm/dd/yyyy"))
        Resume Next
    End If
End Sub

Private Sub Update_Users_withthis_Template(xTemplate, Optional xNewTemplate, Optional xUpdateType)
    Dim SQLQ As String
    Dim rsSecBasic As New ADODB.Recordset
    Dim rsUserSecBasic As New ADODB.Recordset
    
    'Initialise
    xEmployeeNoMissing = False
    
    'Retrieve all users associated with this changed Template
    SQLQ = "SELECT USERID, SECURE_TEMPLATE, EMPNBR FROM HR_SECURE_BASIC WHERE SECURE_TEMPLATE = '" & xTemplate & "'"
    SQLQ = SQLQ & " ORDER BY USERID"
    rsSecBasic.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsSecBasic.EOF
        If Not IsNull(rsSecBasic("USERID")) Then
            'Update each user with this changed Template
            If IsMissing(xNewTemplate) Then
                '????Ticket #24808 - If Deleting the Template then update User's with this Template with Template's Profile
                If IsMissing(xUpdateType) Then
                    Call Template_Based_Security_Profile_Update(rsSecBasic("USERID"), xTemplate, "UpdateBatch")
                Else
                    Call Template_Based_Security_Profile_Update(rsSecBasic("USERID"), xTemplate, xUpdateType)
                End If
            Else
                'Look for the new template security profile because the Template Name has changed and
                'it has already been saved with new name
                Call Template_Based_Security_Profile_Update(rsSecBasic("USERID"), xNewTemplate, "UpdateBatch")
            End If
        End If
        
        'Ticket #24320 - Check if the User has Employee # assigned if Template's Employee # Based Security is checked
        If IsNull(rsSecBasic("EMPNBR")) Or rsSecBasic("EMPNBR") = "" Then
            'User does not have Employee # assigned
            'Check if Template has Employee # Based Checked
            SQLQ = "SELECT EMPNBR, EmpNBR_Based FROM HR_SECURE_BASIC WHERE USERID = '" & IIf(IsMissing(xNewTemplate), xTemplate, xNewTemplate) & "'"
            rsUserSecBasic.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsUserSecBasic.EOF Then
                If rsUserSecBasic("EmpNBR_Based") Then
                     xEmployeeNoMissing = True
                Else
                    xEmployeeNoMissing = False
                End If
            End If
        End If
        
        rsSecBasic.MoveNext
    Loop
    rsSecBasic.Close
    Set rsSecBasic = Nothing
    
End Sub

Private Sub Remove_User_Template_Association(xTemplate)
    'Remove this Template association with users
    gdbAdoIhr001.Execute "UPDATE HR_SECURE_BASIC SET SECURE_TEMPLATE = '' WHERE SECURE_TEMPLATE='" & xTemplate & "'"
End Sub

Private Sub Update_Users_with_NewTemplate_Name(xOldTemplate, xNewTemplate)
    'Update users with new template name
    gdbAdoIhr001.Execute "UPDATE HR_SECURE_BASIC SET SECURE_TEMPLATE = '" & xNewTemplate & "' WHERE SECURE_TEMPLATE='" & xOldTemplate & "'"
End Sub

Private Sub Refresh_Security_Template()
    If cmbSecTemplate.ListCount > 0 Then
        For X = 1 To cmbSecTemplate.ListCount - 1
            If cmbSecTemplate.List(X) = txtSecTemplate.Text Then
                cmbSecTemplate.ListIndex = X
                Exit For
            Else
                cmbSecTemplate.ListIndex = 0
            End If
            'cmbSecTemplate.Text = txtSecTemplate.Text
        Next
    End If
End Sub

Private Function EmployeeNoBasedSecurity(xTemplate) As Integer
    Dim rsSecure As New ADODB.Recordset
    Dim SQLQ As String
    
    EmployeeNoBasedSecurity = 0
    
    'Retrieve Security record of the Template - Employee # based Security only
    SQLQ = "SELECT EmpNBR_Based FROM HR_SECURE_BASIC WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    rsSecure.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsSecure.EOF Then
        MsgBox "This Template has no Security Profile setup.", vbOKOnly, "Error finding Security Profile"
        rsSecure.Close
        Set rsSecure = Nothing
        
        EmployeeNoBasedSecurity = 2
        
        Exit Function
    Else
        'Update User's Security record based on the Template Security record - Employee # based Security only
        If IsNull(rsSecure("EmpNBR_Based")) Then
            EmployeeNoBasedSecurity = 0
        Else
            If rsSecure("EmpNBR_Based") Then EmployeeNoBasedSecurity = 1 Else EmployeeNoBasedSecurity = 0
        End If
    End If
    rsSecure.Close
    Set rsSecure = Nothing
    
End Function

Private Function Validate_ExpiryDate(xUserID, xExpiryDays)
    Dim rsSecure As New ADODB.Recordset
    Dim SQLQ As String
    
    On Error GoTo Err_Validate_ExpiryDate
    'The SQL Server has the limitation in the DateTime field with date range. It can only accomodate the date range
    'from January 1, 1753, through December 31, 9999. Since there is no way I can trap the error while it is trying to
    'update the Expiry Date field computed using the DateAdd function, this function is first computing the date and
    'seeing if the date is invalid. If not then it will update with the upper limit of date range.
    
    Validate_ExpiryDate = False
        
    SQLQ = "SELECT PS_EXPIR_DATE, PS_EXPIR_DAYS FROM HR_SECURE_BASIC WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
    rsSecure.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsSecure.EOF Then
        If Not IsNull(rsSecure("PS_EXPIR_DATE")) Then
            'Check to see if the date is valid
            If Not IsDate(DateAdd("d", " & xExpiryDays & ", rsSecure("PS_EXPIR_DATE"))) Then
                Validate_ExpiryDate = False
            End If
        End If
        Validate_ExpiryDate = True
    Else
        Validate_ExpiryDate = False
    End If
    rsSecure.Close
    Set rsSecure = Nothing

    Exit Function
    
Err_Validate_ExpiryDate:
    If Err.Number = 13 Then Validate_ExpiryDate = False
End Function

Private Sub HideAllShowOnePanDetail(showIndex As Integer)
    Dim X As Integer
    
    For X = 0 To 7
        If X <> showIndex Then
            panDetails(X).Visible = False
            mnu_Sec(X).Checked = False
        End If
    Next
    panDetails(showIndex).Visible = True
    mnu_Sec(showIndex).Checked = True
End Sub

Private Function EmployeeNoBased_Security_Changed(xTemplate)
Dim xAccessable
Dim xTemplateEmpNoSec As Integer

EmployeeNoBased_Security_Changed = False

If Len(xTemplate) > 0 Then
    xTemplateEmpNoSec = EmployeeNoBasedSecurity(xTemplate)
    If xTemplateEmpNoSec = 0 Then
        xAccessable = False
    ElseIf xTemplateEmpNoSec = 1 Or xTemplateEmpNoSec = 2 Then
        xAccessable = True
    End If
    If xAccessable <> chkEESecurity Then
        EmployeeNoBased_Security_Changed = True
    Else
        EmployeeNoBased_Security_Changed = False
    End If
End If

End Function

Private Sub Update_Users_EmployeeNoBasedSecurity(xTemplate)
    Dim rsSecBasic As New ADODB.Recordset
    Dim SQLQ As String
    
    'Update each User's profile who belong to this Template with the Employee # Based Security value.
    SQLQ = "SELECT * FROM HR_SECURE_BASIC WHERE SECURE_TEMPLATE = '" & Replace(xTemplate, "'", "''") & "'"
    SQLQ = SQLQ & " ORDER BY USERID"
    rsSecBasic.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsSecBasic.EOF
        rsSecBasic("EmpNBR_Based") = IIf(chkEESecurity, 1, 0)
        rsSecBasic("LDATE") = Date
        rsSecBasic("LTIME") = Time$
        rsSecBasic("LUSER") = glbUserID
        rsSecBasic.Update
        
        rsSecBasic.MoveNext
    Loop
    rsSecBasic.Close
    Set rsSecBasic = Nothing
    
End Sub
