VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSalPerctg 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Salary % Increase Master"
   ClientHeight    =   10950
   ClientLeft      =   2565
   ClientTop       =   525
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10950
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin VB.Frame VacFram03 
      BorderStyle     =   0  'None
      Height          =   4395
      Left            =   120
      TabIndex        =   111
      Top             =   30
      Width           =   11415
      Begin Threed.SSCheck chkManual 
         Height          =   255
         Left            =   8400
         TabIndex        =   134
         Top             =   3270
         Visible         =   0   'False
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Exclude from Update All"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin INFOHR_Controls.DateLookup dlpAsOf 
         Height          =   285
         Left            =   9900
         TabIndex        =   10
         Tag             =   "40-As of Date"
         Top             =   3420
         Visible         =   0   'False
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   6420
         TabIndex        =   6
         Tag             =   "00-Position Group - Code"
         Top             =   2610
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "JBGC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   945
         TabIndex        =   2
         Tag             =   "00-Enter Union Code"
         Top             =   2640
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   945
         TabIndex        =   1
         Tag             =   "00-Specific Department Desired"
         Top             =   2340
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   7
         LookupType      =   2
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         Height          =   285
         Left            =   945
         TabIndex        =   0
         Tag             =   "00-Specific Division Desired"
         Top             =   2040
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   6420
         TabIndex        =   4
         Tag             =   "00-Specific Employment Status Desired"
         Top             =   2010
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDEM"
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         Height          =   285
         Left            =   6420
         TabIndex        =   5
         Tag             =   "EDPT-Category"
         Top             =   2310
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDPT"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   6420
         TabIndex        =   7
         Tag             =   "00-Section - Code"
         Top             =   2910
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   945
         TabIndex        =   3
         Tag             =   "00-Enter Location Code"
         Top             =   2940
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   0
         Left            =   2100
         TabIndex        =   8
         Tag             =   "40-From Date"
         Top             =   3390
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1210
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   1
         Left            =   3870
         TabIndex        =   9
         Tag             =   "40-To Date"
         Top             =   3390
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1210
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
         Bindings        =   "fsSalPerctg.frx":0000
         Height          =   1695
         Left            =   0
         OleObjectBlob   =   "fsSalPerctg.frx":0014
         TabIndex        =   136
         Top             =   0
         Width           =   9135
      End
      Begin VB.Label lblPeriod 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Attendance Period"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   135
         Top             =   3435
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label lblDiv 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
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
         Left            =   30
         TabIndex        =   124
         Top             =   2040
         Width           =   555
      End
      Begin VB.Label lblDept 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         Left            =   30
         TabIndex        =   123
         Top             =   2340
         Width           =   825
      End
      Begin VB.Label lblUnion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Union"
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
         Left            =   30
         TabIndex        =   122
         Top             =   2670
         Width           =   420
      End
      Begin VB.Label lblCriteria 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employment Status"
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
         Index           =   3
         Left            =   4920
         TabIndex        =   121
         Top             =   2040
         Width           =   1350
      End
      Begin VB.Label lblAsOf 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8550
         TabIndex        =   120
         Top             =   3465
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label textMulti 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "The Union Code and FT/PT/SE/TR/OT will be validated from the Employee Basic Data"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   119
         Top             =   3870
         Visible         =   0   'False
         Width           =   7455
      End
      Begin VB.Label lblCriteria 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Group"
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
         Index           =   5
         Left            =   4920
         TabIndex        =   118
         Top             =   2640
         Width           =   1260
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Ranges (in Hours)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   117
         Top             =   4170
         Width           =   2250
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3960
         TabIndex        =   116
         Top             =   4170
         Width           =   540
      End
      Begin VB.Label lblSelCri 
         Caption         =   "Selection Criteria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   115
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblPT 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         Left            =   4920
         TabIndex        =   114
         Top             =   2340
         Width           =   630
      End
      Begin VB.Label lblSection 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
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
         Left            =   4920
         TabIndex        =   113
         Top             =   2940
         Width           =   1260
      End
      Begin VB.Label lblLocation 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Left            =   30
         TabIndex        =   112
         Top             =   2970
         Width           =   615
      End
   End
   Begin VB.VScrollBar scrControl 
      Height          =   4125
      LargeChange     =   315
      Left            =   10800
      Max             =   100
      SmallChange     =   315
      TabIndex        =   109
      Top             =   4710
      Width           =   300
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   10305
      Width           =   11760
      _Version        =   65536
      _ExtentX        =   20743
      _ExtentY        =   1138
      _StockProps     =   15
      ForeColor       =   0
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
      Font3D          =   1
      Alignment       =   1
      Begin VB.CommandButton cmdUpdateAll 
         Caption         =   "Update All"
         Height          =   375
         Left            =   5400
         TabIndex        =   133
         Top             =   120
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update Entitlement"
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Tag             =   "Change all matching records to the above"
         Top             =   120
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.CommandButton CmdRecalc 
         Appearance      =   0  'Flat
         Caption         =   "R&ecalculate"
         Height          =   375
         Left            =   3600
         TabIndex        =   110
         Tag             =   "Recalculation"
         Top             =   120
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton cmdPrintAll 
         Appearance      =   0  'Flat
         Caption         =   "Print &All"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Tag             =   "Print all Vacation Entitlement Report"
         Top             =   120
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   405
         Left            =   7800
         Top             =   0
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   714
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   1
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         BoundReportHeading=   "RGELIST"
         BoundReportFooter=   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.Frame VacFram 
      BorderStyle     =   0  'None
      Height          =   8500
      Left            =   130
      TabIndex        =   14
      Top             =   4440
      Width           =   11000
      Begin VB.Frame frmAG 
         Height          =   400
         Left            =   5040
         TabIndex        =   137
         Top             =   5
         Width           =   2175
         Begin Threed.SSOption optG 
            Height          =   195
            Left            =   1080
            TabIndex        =   19
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   150
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Grid Step"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optA 
            Height          =   225
            Left            =   120
            TabIndex        =   18
            Tag             =   "Entitlement measured in days"
            Top             =   140
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Actual"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Tag             =   "11-Service is greater than this number"
         Top             =   90
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   16
         Tag             =   "10-Service is less than this number"
         Top             =   105
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   20
         Tag             =   "11-Service is greater than this number"
         Top             =   420
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   21
         Tag             =   "10-Service is less than this number"
         Top             =   427
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   2
         Left            =   0
         TabIndex        =   23
         Tag             =   "11-Service is greater than this number"
         Top             =   735
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   24
         Tag             =   "10-Service is less than this number"
         Top             =   749
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   1
         Left            =   3750
         TabIndex        =   22
         Tag             =   "11-Salary Step/Amount"
         Top             =   442
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   2
         Left            =   3750
         TabIndex        =   25
         Tag             =   "11-Salary Step/Amount"
         Top             =   764
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   3
         Left            =   0
         TabIndex        =   26
         Tag             =   "11-Service is greater than this number"
         Top             =   1050
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   27
         Tag             =   "10-Service is less than this number"
         Top             =   1071
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   4
         Left            =   0
         TabIndex        =   29
         Tag             =   "11-Service is greater than this number"
         Top             =   1380
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   5
         Left            =   0
         TabIndex        =   32
         Tag             =   "11-Service is greater than this number"
         Top             =   1710
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   33
         Tag             =   "10-Service is less than this number"
         Top             =   1715
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   5
         Left            =   3750
         TabIndex        =   34
         Tag             =   "11-Salary Step/Amount"
         Top             =   1730
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   3
         Left            =   3750
         TabIndex        =   28
         Tag             =   "11-Salary Step/Amount"
         Top             =   1086
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   6
         Left            =   0
         TabIndex        =   35
         Tag             =   "11-Service is greater than this number"
         Top             =   2040
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   6
         Left            =   2160
         TabIndex        =   36
         Tag             =   "10-Service is less than this number"
         Top             =   2037
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   7
         Left            =   0
         TabIndex        =   38
         Tag             =   "11-Service is greater than this number"
         Top             =   2355
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   7
         Left            =   2160
         TabIndex        =   39
         Tag             =   "10-Service is less than this number"
         Top             =   2359
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   8
         Left            =   0
         TabIndex        =   41
         Tag             =   "11-Service is greater than this number"
         Top             =   2670
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   8
         Left            =   2160
         TabIndex        =   42
         Tag             =   "10-Service is less than this number"
         Top             =   2681
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   7
         Left            =   3750
         TabIndex        =   40
         Tag             =   "11-Salary Step/Amount"
         Top             =   2374
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   8
         Left            =   3750
         TabIndex        =   43
         Tag             =   "11-Salary Step/Amount"
         Top             =   2696
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   6
         Left            =   3750
         TabIndex        =   37
         Tag             =   "11-Salary Step/Amount"
         Top             =   2052
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   9
         Left            =   0
         TabIndex        =   44
         Tag             =   "11-Service is greater than this number"
         Top             =   2980
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   9
         Left            =   2160
         TabIndex        =   45
         Tag             =   "10-Service is less than this number"
         Top             =   3003
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   9
         Left            =   3750
         TabIndex        =   46
         Tag             =   "11-Salary Step/Amount"
         Top             =   3018
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   10
         Left            =   0
         TabIndex        =   47
         Tag             =   "11-Service is greater than this number"
         Top             =   3300
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   10
         Left            =   2160
         TabIndex        =   48
         Tag             =   "10-Service is less than this number"
         Top             =   3325
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   11
         Left            =   0
         TabIndex        =   50
         Tag             =   "11-Service is greater than this number"
         Top             =   3630
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   11
         Left            =   2160
         TabIndex        =   51
         Tag             =   "10-Service is less than this number"
         Top             =   3647
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   10
         Left            =   3750
         TabIndex        =   49
         Tag             =   "11-Salary Step/Amount"
         Top             =   3340
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   11
         Left            =   3750
         TabIndex        =   52
         Tag             =   "11-Salary Step/Amount"
         Top             =   3662
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   12
         Left            =   0
         TabIndex        =   53
         Tag             =   "11-Service is greater than this number"
         Top             =   3960
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   12
         Left            =   2160
         TabIndex        =   54
         Tag             =   "10-Service is less than this number"
         Top             =   3969
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   13
         Left            =   0
         TabIndex        =   56
         Tag             =   "11-Service is greater than this number"
         Top             =   4275
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   13
         Left            =   2160
         TabIndex        =   57
         Tag             =   "10-Service is less than this number"
         Top             =   4291
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   14
         Left            =   0
         TabIndex        =   59
         Tag             =   "11-Service is greater than this number"
         Top             =   4590
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   14
         Left            =   2160
         TabIndex        =   60
         Tag             =   "10-Service is less than this number"
         Top             =   4613
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   12
         Left            =   3750
         TabIndex        =   55
         Tag             =   "11-Salary Step/Amount"
         Top             =   3984
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   13
         Left            =   3750
         TabIndex        =   58
         Tag             =   "11-Salary Step/Amount"
         Top             =   4306
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   14
         Left            =   3750
         TabIndex        =   61
         Tag             =   "11-Salary Step/Amount"
         Top             =   4628
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   15
         Left            =   0
         TabIndex        =   62
         Tag             =   "11-Service is greater than this number"
         Top             =   4940
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   15
         Left            =   2160
         TabIndex        =   63
         Tag             =   "10-Service is less than this number"
         Top             =   4935
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   16
         Left            =   0
         TabIndex        =   65
         Tag             =   "11-Service is greater than this number"
         Top             =   5260
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   16
         Left            =   2160
         TabIndex        =   66
         Tag             =   "10-Service is less than this number"
         Top             =   5257
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   15
         Left            =   3750
         TabIndex        =   64
         Tag             =   "11-Salary Step/Amount"
         Top             =   4950
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   16
         Left            =   3750
         TabIndex        =   67
         Tag             =   "11-Salary Step/Amount"
         Top             =   5272
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   17
         Left            =   0
         TabIndex        =   68
         Tag             =   "11-Service is greater than this number"
         Top             =   5595
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   17
         Left            =   2160
         TabIndex        =   69
         Tag             =   "10-Service is less than this number"
         Top             =   5595
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   17
         Left            =   3750
         TabIndex        =   70
         Tag             =   "11-Salary Step/Amount"
         Top             =   5595
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   18
         Left            =   0
         TabIndex        =   71
         Tag             =   "11-Service is greater than this number"
         Top             =   5910
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   18
         Left            =   2160
         TabIndex        =   72
         Tag             =   "10-Service is less than this number"
         Top             =   5910
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   19
         Left            =   0
         TabIndex        =   74
         Tag             =   "11-Service is greater than this number"
         Top             =   6240
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   19
         Left            =   2160
         TabIndex        =   75
         Tag             =   "10-Service is less than this number"
         Top             =   6240
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   18
         Left            =   3750
         TabIndex        =   73
         Tag             =   "11-Salary Step/Amount"
         Top             =   5940
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   19
         Left            =   3750
         TabIndex        =   76
         Tag             =   "11-Salary Step/Amount"
         Top             =   6255
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   20
         Left            =   0
         TabIndex        =   77
         Tag             =   "11-Service is greater than this number"
         Top             =   6570
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   20
         Left            =   2160
         TabIndex        =   78
         Tag             =   "10-Service is less than this number"
         Top             =   6570
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   21
         Left            =   0
         TabIndex        =   80
         Tag             =   "11-Service is greater than this number"
         Top             =   6885
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   21
         Left            =   2160
         TabIndex        =   81
         Tag             =   "10-Service is less than this number"
         Top             =   6885
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   22
         Left            =   0
         TabIndex        =   83
         Tag             =   "11-Service is greater than this number"
         Top             =   7200
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   22
         Left            =   2160
         TabIndex        =   84
         Tag             =   "10-Service is less than this number"
         Top             =   7200
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   20
         Left            =   3750
         TabIndex        =   79
         Tag             =   "11-Salary Step/Amount"
         Top             =   6570
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   21
         Left            =   3750
         TabIndex        =   82
         Tag             =   "11-Salary Step/Amount"
         Top             =   6885
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   22
         Left            =   3750
         TabIndex        =   85
         Tag             =   "11-Salary Step/Amount"
         Top             =   7200
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   23
         Left            =   0
         TabIndex        =   86
         Tag             =   "11-Service is greater than this number"
         Top             =   7545
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   23
         Left            =   2160
         TabIndex        =   87
         Tag             =   "10-Service is less than this number"
         Top             =   7545
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   24
         Left            =   0
         TabIndex        =   89
         Tag             =   "11-Service is greater than this number"
         Top             =   7875
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   24
         Left            =   2160
         TabIndex        =   90
         Tag             =   "10-Service is less than this number"
         Top             =   7875
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   23
         Left            =   3750
         TabIndex        =   88
         Tag             =   "11-Salary Step/Amount"
         Top             =   7545
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   24
         Left            =   3750
         TabIndex        =   91
         Tag             =   "11-Salary Step/Amount"
         Top             =   7875
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   0
         Left            =   3750
         TabIndex        =   17
         Tag             =   "11-Salary Step/Amount"
         Top             =   120
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   30
         Tag             =   "10-Service is less than this number"
         Top             =   1393
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   4
         Left            =   3750
         TabIndex        =   31
         Tag             =   "11-Salary Step/Amount"
         Top             =   1408
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   24
         Left            =   975
         TabIndex        =   132
         Top             =   7530
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   23
         Left            =   975
         TabIndex        =   131
         Top             =   5610
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   22
         Left            =   975
         TabIndex        =   130
         Top             =   6270
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   21
         Left            =   975
         TabIndex        =   129
         Top             =   5955
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   20
         Left            =   975
         TabIndex        =   128
         Top             =   7215
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   19
         Left            =   975
         TabIndex        =   127
         Top             =   6915
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   18
         Left            =   975
         TabIndex        =   126
         Top             =   6600
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ">    Service  "
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
         Index           =   17
         Left            =   975
         TabIndex        =   125
         Top             =   7890
         Width           =   915
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   9
         Left            =   980
         TabIndex        =   108
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   8
         Left            =   980
         TabIndex        =   107
         Top             =   2070
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   7
         Left            =   980
         TabIndex        =   106
         Top             =   2385
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   6
         Left            =   980
         TabIndex        =   105
         Top             =   2685
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   5
         Left            =   980
         TabIndex        =   104
         Top             =   1110
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   4
         Left            =   980
         TabIndex        =   103
         Top             =   1425
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   3
         Left            =   980
         TabIndex        =   102
         Top             =   1740
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   2
         Left            =   980
         TabIndex        =   101
         Top             =   780
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   1
         Left            =   980
         TabIndex        =   100
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   0
         Left            =   980
         TabIndex        =   99
         Top             =   150
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   10
         Left            =   980
         TabIndex        =   98
         Top             =   3990
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   11
         Left            =   980
         TabIndex        =   97
         Top             =   4305
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   12
         Left            =   980
         TabIndex        =   96
         Top             =   4605
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   13
         Left            =   980
         TabIndex        =   95
         Top             =   3345
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   14
         Left            =   980
         TabIndex        =   94
         Top             =   3660
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   15
         Left            =   980
         TabIndex        =   93
         Top             =   3000
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   16
         Left            =   980
         TabIndex        =   92
         Top             =   4920
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmSalPerctg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fTablHREMP As New ADODB.Recordset         ' table view of HREMP
Dim snapEntitle As New ADODB.Recordset     'user vier
Dim fglbWDate$, fglbWDateS$
Dim fglbAsOf As Date
Dim Actn

Dim fglbSDate As Variant
Dim fglbMaxRange%
Dim fglbCompMonthly%

Dim fglbMaxRanges%
Dim glbFrmCaption$, glbErrNum&

Dim ControlsShown As Boolean
Dim ODIV, ODept, oOrg, oAsOf, oEMP, oEmpMode, oGRPCE
Dim OSection, OLoc
Dim OFromDate, OToDate
Dim FlagRefresh As Boolean

Dim fglbESQLQ, fglbVSQLQ
Dim fglbNew As Boolean
Dim fglbRunTimes
Dim Memplist1, Memplist2

Private Function chkMUEntitle()
Dim x%, Y%

chkMUEntitle = False

On Error GoTo chkMUEntitle_Err
For x% = 0 To 4
If Len(clpCode(x%).Text) > 0 And clpCode(x%).Caption = "Unassigned" Then
    MsgBox "If Code entered it must be known"
    clpCode(x%).SetFocus
    Exit Function
End If
Next x%

If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "Invalid Department"
     clpDept.SetFocus
    Exit Function
End If
If Len(clpDiv.Text) < 1 Then
    If glbDIVCount = 1 And glbLinamar Then
        MsgBox lStr("Division is required field")
         clpDiv.SetFocus
        Exit Function
    End If
Else
    If clpDiv.Caption = "Unassigned" Then
        MsgBox lStr("Invalid Division")
         clpDiv.SetFocus
        Exit Function
    End If
End If

If clpPT.Caption = "Unassigned" Then
    MsgBox "Invalid " & lblPT.Caption
    clpPT.SetFocus
    Exit Function
End If

'Ticket #15276 - Commented out
'If Len(dlpDateRange(0).Text) > 0 Then
'  If Not IsDate(dlpDateRange(0).Text) Then
'    MsgBox "Invalid Attendance Period From Date"
'    dlpDateRange(0).SetFocus
'    Exit Function
'  End If
'Else
'    MsgBox "Attendance Period From Date is mandatory field"
'    dlpDateRange(0).SetFocus
'    Exit Function
'End If
'
'If Len(dlpDateRange(1).Text) > 0 Then
'  If Not IsDate(dlpDateRange(1).Text) Then
'    MsgBox "Invalid Attendance Period To Date"
'    dlpDateRange(1).SetFocus
'    Exit Function
'  End If
'Else
'    MsgBox "Attendance Period To Date is mandatory field"
'    dlpDateRange(1).SetFocus
'    Exit Function
'End If

'If Len(dlpAsOf.Text) > 0 Then
'  If Not IsDate(dlpAsOf.Text) Then
'    MsgBox "Invalid Effective Date"
'    dlpAsOf.SetFocus
'    Exit Function
'  End If
'Else
'    'If UCase(glbCompEntSick$) = "A" Then
'    '    If glbLinamar Then
'            MsgBox "Effective Date is required field"
'            dlpAsOf.SetFocus
'            Exit Function
'    '    End If
'    'End If
'End If

If Len(medLTServ(0)) < 1 Then
    MsgBox "You must have at least one Service Range Entry."
    If medLTServ(0).Enabled Then medLTServ(0).SetFocus
    Exit Function
End If

'Frank 05/13/2004 Ticket#
If glbWFC Then
    If Len(clpCode(3).Text) = 0 Then
        MsgBox lStr("Section is required field")
        clpCode(3).SetFocus
        Exit Function
    End If
End If

fglbMaxRanges% = 0  ' 0 is first range

Dim intRangesSet%
intRangesSet% = 0    ' 1 to 4 with 0 implying none
If Len(medLTServ(3)) = 0 Then
    medGTServ(3) = ""
Else
    If medLTServ(3) = 0 Then
        medLTServ(3) = ""
        medGTServ(3) = ""
    End If
End If


For x% = 0 To 24
    If Len(medLTServ(x%)) > 0 Then
        If Not IsNumeric(medLTServ(x%)) Then
            MsgBox "Data Entered Must Be Numeric"
            medLTServ(x%).SetFocus
            Exit Function
        End If
    End If
    If Len(medGTServ(x%)) > 0 Then
        If Not IsNumeric(medGTServ(x%)) Then
            MsgBox "Data Entered Must Be Numeric"
            medGTServ(x%).SetFocus
            Exit Function
        End If
    End If
    If Len(medEntitle(x%)) > 0 Then
        If Not IsNumeric(medEntitle(x%)) Then
            MsgBox "Data Entered Must Be Numeric"
            medEntitle(x%).SetFocus
            Exit Function
        End If
    End If

    If Len(medLTServ(x%)) < 1 And Len(medGTServ(x%)) > 1 Then  ' missed one
        MsgBox "Ranges must be sequential"
        medLTServ(x%).SetFocus
        Exit Function
    End If
    If Len(medGTServ(x%)) > 0 Then
        If Val(medLTServ(x%)) > Val(medGTServ(x%)) Then
            MsgBox "Ranges must be sequential"
            medLTServ(x%).SetFocus
            Exit Function
        End If
    End If
    If x% > 0 And Len(medLTServ(x%)) > 0 Then
        If Val(medLTServ(x%)) < Val(medGTServ(x% - 1)) Then
            MsgBox "Ranges must be sequential"
            medLTServ(x%).SetFocus
            Exit Function
        End If
    End If
    If x% > 0 And Len(medGTServ(x%)) > 0 Then
        If Val(medGTServ(x%)) < Val(medGTServ(x% - 1)) And Val(medGTServ(x%)) <> 0 Then
            MsgBox "Ranges must be sequential"
            medLTServ(x%).SetFocus
            Exit Function
        End If
    End If

    If Len(medLTServ(x%)) < 1 Then Exit For  ' missed one
    intRangesSet% = intRangesSet% + 1
Next x%

If intRangesSet% = 0 Then
    MsgBox "At least one Service level must be set"
    medLTServ(0).SetFocus
    Exit Function
End If

'For X% = 0 To 24
'    If Len(medMax(X%)) < 1 Then
'        medMax(X%) = 0
'    End If
'Next X%
chkMUEntitle = True

Exit Function

chkMUEntitle_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEntitle", "HR_SALARY_INCR", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub clpCode_LostFocus(Index As Integer)
        
        'This function only for Vacation, comment by Frank on Mar 2,03
        'If glbWHSCC And Actn = "A" And Index = 0 Then
        '   If (clpCode(0) = "1866" Or clpCode(0) = "946") And clpPT = "FT" Then
        '       medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
        '       medLTServ(1) = 60: medGTServ(1) = 215.99: medEntitle(1) = 1.67
        '       medLTServ(2) = 216: medGTServ(2) = 999: medEntitle(2) = 2.09
        '   End If
        '   If clpCode(0) = "NON" And clpPT = "FT" Then
        '       optD(0).SetFocus
        '       medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
        '       medLTServ(1) = 60: medGTServ(1) = 108.99: medEntitle(1) = 1.67
        '       medLTServ(2) = 109: medGTServ(2) = 119.99: medEntitle(2) = 21
        '       medLTServ(3) = 120: medGTServ(3) = 131.99: medEntitle(3) = 22
        '       medLTServ(4) = 132: medGTServ(4) = 143.99: medEntitle(4) = 23
        '       medLTServ(5) = 144: medGTServ(5) = 155.99: medEntitle(5) = 24
        '       medLTServ(6) = 156: medGTServ(6) = 167.99: medEntitle(6) = 25
        '       medLTServ(7) = 168: medGTServ(7) = 179.99: medEntitle(7) = 26
        '       medLTServ(8) = 180: medGTServ(8) = 191.99: medEntitle(8) = 27
        '       medLTServ(9) = 192: medGTServ(9) = 203.99: medEntitle(9) = 28
        '       medLTServ(10) = 204: medGTServ(10) = 215.99: medEntitle(10) = 29
        '       medLTServ(11) = 216: medGTServ(11) = 999999.99: medEntitle(11) = 30
        '   End If
        '   If clpCode(0) = "PHYS" And clpPT = "FT" Then
        '       medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
        '       medLTServ(1) = 60: medGTServ(1) = 119: medEntitle(1) = 1.67
        '   End If
        '   If clpCode(0) = "NON" And clpPT = "PT" Then
        '       optF(0).SetFocus
        '       medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
        '       medLTServ(1) = 60: medGTServ(1) = 108.99: medEntitle(1) = 1.67
        '       medLTServ(2) = 109: medGTServ(2) = 119.99: medEntitle(2) = 21
        '       medLTServ(3) = 120: medGTServ(3) = 131.99: medEntitle(3) = 22
        '       medLTServ(4) = 132: medGTServ(4) = 143.99: medEntitle(4) = 23
        '       medLTServ(5) = 144: medGTServ(5) = 155.99: medEntitle(5) = 24
        '       medLTServ(6) = 156: medGTServ(6) = 167.99: medEntitle(6) = 25
        '       medLTServ(7) = 168: medGTServ(7) = 179.99: medEntitle(7) = 26
        '       medLTServ(8) = 180: medGTServ(8) = 191.99: medEntitle(8) = 27
        '       medLTServ(9) = 192: medGTServ(9) = 203.99: medEntitle(9) = 28
        '       medLTServ(10) = 204: medGTServ(10) = 215.99: medEntitle(10) = 29
        '       medLTServ(11) = 216: medGTServ(11) = 999999.99: medEntitle(11) = 30
        '   End If
        'End If
        'End Sub
        '
        'Private Sub clpPT_LostFocus()
        'If glbWHSCC And Actn = "A" Then  'And Index = 0 Then
        '   If (clpCode(0) = "1866" Or clpCode(0) = "946") And clpPT = "FT" Then
        '       medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
        '       medLTServ(1) = 60: medGTServ(1) = 215.99: medEntitle(1) = 1.67
        '       medLTServ(2) = 216: medGTServ(2) = 999: medEntitle(2) = 2.09
        '   End If
        '   If clpCode(0) = "NON" And clpPT = "FT" Then
        '       optD(0).SetFocus
        '       medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
        '       medLTServ(1) = 60: medGTServ(1) = 119.99: medEntitle(1) = 1.67
        '       medLTServ(2) = 120: medGTServ(2) = 131.99: medEntitle(2) = 21
        '       medLTServ(3) = 132: medGTServ(3) = 143.99: medEntitle(3) = 22
        '       medLTServ(4) = 144: medGTServ(4) = 155.99: medEntitle(4) = 23
        '       medLTServ(5) = 156: medGTServ(5) = 167.99: medEntitle(5) = 24
        '       medLTServ(6) = 168: medGTServ(6) = 179.99: medEntitle(6) = 25
        '       medLTServ(7) = 180: medGTServ(7) = 191.99: medEntitle(7) = 26
        '       medLTServ(8) = 192: medGTServ(8) = 203.99: medEntitle(8) = 27
        '       medLTServ(9) = 204: medGTServ(9) = 215.99: medEntitle(9) = 28
        '       medLTServ(10) = 216: medGTServ(10) = 227.99: medEntitle(10) = 29
        '       medLTServ(11) = 228: medGTServ(11) = 999999.99: medEntitle(11) = 30
        '   End If
        '   If clpCode(0) = "PHYS" And clpPT = "FT" Then
        '       medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
        '       medLTServ(1) = 60: medGTServ(1) = 119: medEntitle(1) = 1.67
        '   End If
        '   If clpCode(0) = "NON" And clpPT = "PT" Then
        '       optF(0).SetFocus
        '       medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
        '       medLTServ(1) = 60: medGTServ(1) = 119.99: medEntitle(1) = 1.67
        '       medLTServ(2) = 120: medGTServ(2) = 131.99: medEntitle(2) = 21
        '       medLTServ(3) = 132: medGTServ(3) = 143.99: medEntitle(3) = 22
        '       medLTServ(4) = 144: medGTServ(4) = 155.99: medEntitle(4) = 23
        '       medLTServ(5) = 156: medGTServ(5) = 167.99: medEntitle(5) = 24
        '       medLTServ(6) = 168: medGTServ(6) = 179.99: medEntitle(6) = 25
        '       medLTServ(7) = 180: medGTServ(7) = 191.99: medEntitle(7) = 26
        '       medLTServ(8) = 192: medGTServ(8) = 203.99: medEntitle(8) = 27
        '       medLTServ(9) = 204: medGTServ(9) = 215.99: medEntitle(9) = 28
        '       medLTServ(10) = 216: medGTServ(10) = 227.99: medEntitle(10) = 29
        '       medLTServ(11) = 228: medGTServ(11) = 999999.99: medEntitle(11) = 30
        '   End If
        'End If
End Sub

Sub cmdCancel_Click()
    fglbNew = False
    
    Data1.Refresh
    
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    
    Call Display_Value
    
    vbxTrueGrid.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Sub cmdDelete_Click()
    Dim SQLQ, Msg, A%
    
    If Data1.Recordset.BOF And Data1.Recordset.EOF Then
        MsgBox "Nothing to Delete"
        Exit Sub
    End If
    Msg = "Are You Sure You Want To Delete "
    Msg = Msg & Chr(10) & "The Salary Increase Rules?  "
    
    A% = MsgBox(Msg, 36, "Confirm Delete")
    If A% <> 6 Then Exit Sub
    
    Call getWSQLQ("C")
    SQLQ = "DELETE FROM HR_SALARY_INCR WHERE " & fglbVSQLQ
    
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    
    Data1.Refresh
    Call Display_Value
End Sub

Sub cmdModify_Click()
    ODIV = clpDiv.Text
    ODept = clpDept.Text
    oOrg = clpCode(0).Text
    
    'Franks 04/08/03 Ticket# 3943
    'Fix the problem: enter or change Effective Date first, click Edit and then Save,
    'it create another record
    oAsOf = ""
'    If Not Data1.Recordset.EOF Then
'        If Not IsNull(Data1.Recordset("VE_EDATE")) Then
'            oAsOf = Data1.Recordset("VE_EDATE")
'        End If
'    End If
    'Sam 02/02/2006
    'Ticket #15276 - Commented out
'    OFromDate = dlpDateRange(0).Text
'    OToDate = dlpDateRange(1).Text
    'Sam 02/02/2006
    
    OLoc = clpCode(4).Text
    OSection = clpCode(3).Text
    oEMP = clpCode(1).Text
    oEmpMode = clpPT.Text
    oGRPCE = clpCode(2).Text
    Actn = "M"
End Sub

Sub cmdNew_Click()
    Dim x
    For x = 0 To 24
        medLTServ(x) = ""
        medGTServ(x) = ""
        medEntitle(x) = ""
        optA = True
        optG = False
    '    optF(X) = False
    '    medMax(X) = ""
    Next
    
    'Sam 02/2/2006
    'Ticket #15276 - Commented out
'    dlpDateRange(0).Text = ""
'    dlpDateRange(1).Text = ""
    'Sam 02/2/2006
    
    clpDiv.Text = ""
    clpDept.Text = ""
    clpCode(0).Text = ""
    dlpAsOf.Text = ""
    clpCode(1).Text = ""
    clpCode(2).Text = ""
    clpCode(3).Text = ""
    clpCode(4).Text = ""
    clpPT.Text = ""
    Actn = "A"
    fglbNew = True
    
    Call SET_UP_MODE
    clpDiv.SetFocus

End Sub

Sub cmdOK_Click()
    Dim x%, Y%, xUnion, xPT, SQLQ, SQLQW
    Dim xStr
    Dim rsVE As New ADODB.Recordset
    Dim rsVT As New ADODB.Recordset
    Dim glbiOneWhere As Boolean
    Dim bmk As Variant
    
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        bmk = 0 'Ticket #11885 Frank Oct 11th, 2006
    Else
        bmk = Data1.Recordset.Bookmark
    End If
    
    If Not chkMUEntitle() Then Exit Sub
    For x% = 0 To 24
        If Not IsNumeric(medLTServ(x%)) Then Exit For
        If Not IsNumeric(medGTServ(x%)) Then
          medGTServ(x%) = 0
        Else
          If Val(medGTServ(x%)) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
        End If
        If medLTServ(x%) > 0 And medGTServ(x%) = 0 Then medGTServ(x%) = 9999999
    Next
    
    If Actn = "M" Then
        Call getWSQLQ("O")
        SQLQ = "DELETE FROM HR_SALARY_INCR WHERE " & fglbVSQLQ
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
    Else
        Call getWSQLQ("C")
        SQLQ = "SELECT * FROM HR_SALARY_INCR WHERE " & fglbVSQLQ
        rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsVT.EOF Then
            MsgBox "You can not add duplicate record"
             clpDiv.SetFocus
            Exit Sub
        End If
    End If
    gdbAdoIhr001.BeginTrans
    SQLQ = "SELECT * FROM HR_SALARY_INCR"
    rsVE.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    For x% = 0 To 24
        If Len(medLTServ(x%)) > 0 Then
            rsVE.AddNew
            rsVE("SL_ORDER") = x + 1
            rsVE("SL_ORG_TABL") = "EDOR"
            rsVE("SL_ORG") = clpCode(0).Text
            rsVE("SL_PT") = clpPT.Text
            rsVE("SL_DIV") = clpDiv.Text
            rsVE("SL_DEPT") = clpDept.Text
            rsVE("SL_EMP_TABL") = "EDEM"
            rsVE("SL_EMP") = clpCode(1).Text
            rsVE("SL_SECTION") = clpCode(3).Text
            rsVE("SL_LOC") = clpCode(4).Text
            'rsVE("VE_EDATE") = dlpAsOf.Text
            
            'Ticket #15276 - Commented out
'            If Len(dlpDateRange(0).Text) > 0 Then
'                rsVE("SL_FRDATE") = dlpDateRange(0).Text
'            End If
'            If Len(dlpDateRange(1).Text) > 0 Then
'                rsVE("SL_TODATE") = dlpDateRange(1).Text
'            End If
            
            rsVE("SL_GRPCD_TABL") = "JBGC"
            rsVE("SL_GRPCD") = clpCode(2).Text
            rsVE("SL_FROM_HRS") = medLTServ(x%)
            rsVE("SL_TO_HRS") = medGTServ(x%)
            If medEntitle(x%) = "" Then
                rsVE("SL_SALARY") = Null
            Else
                rsVE("SL_SALARY") = medEntitle(x%)
            End If
            If optA Then rsVE("SL_TYPE") = "A"
            If optG Then rsVE("SL_TYPE") = "G"
    '        If optF(X%) Then rsVE("VE_TYPE") = "F"
    '        rsVE("VE_MAX") = medMax(X%)
    '        rsVE("VE_MANUAL") = chkManual.Value
            rsVE.Update
        End If
    Next
    rsVE.Close
    gdbAdoIhr001.CommitTrans
    
    'If Not glbSQL and not glboracle Then Call Pause(0.5)
    
    Data1.Refresh
    
    If Not bmk = 0 Then
        Data1.Recordset.Bookmark = bmk
    End If
    
    fglbNew = False
    
    Call Display_Value

End Sub

Sub cmdPrint_Click()
    Dim RHeading As String, xReport, x%
    Dim SQLQ
    Dim dtYYY%, dtMM%, dtDD%
    'cmdPrint.Enabled = False
    
    Me.vbxCrystal.Reset
    Me.vbxCrystal.WindowTitle = "Salary % Increase Report"
    Call setRptLabel(Me, 0) '1)
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 5
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgsalincrmst.rpt"
    
    SQLQ = "(1=1) "
    If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_DIV} = '" & clpDiv.Text & "'"
    If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_DEPT} = '" & clpDept.Text & "'"
    If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_ORG} = '" & clpCode(0).Text & "'"
'    If Len(dlpAsOf.Text) > 0 Then
'        dtYYY% = Year(dlpAsOf.Text)
'        dtMM% = Month(dlpAsOf.Text)
'        dtDD% = Day(dlpAsOf.Text)
'        SQLQ = SQLQ & " AND {HRSICKENT.VE_EDATE} = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'    End If
    If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_EMP} = '" & clpCode(1).Text & "'"
    If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_PT} = '" & clpPT.Text & "' "
    If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_GRPCD} = '" & clpCode(2).Text & "'"
    If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_SECTION} = '" & clpCode(3).Text & "'"
    If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_LOC} = '" & clpCode(4).Text & "'"
    
    'Ticket #15276 - Commented out
'    If Len(dlpDateRange(0).Text) > 0 Then
'        dtYYY% = Year(dlpDateRange(0).Text)
'        dtMM% = Month(dlpDateRange(0).Text)
'        dtDD% = Day(dlpDateRange(0).Text)
'        SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_FRDATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'    End If
'    If Len(dlpDateRange(1).Text) > 0 Then
'        dtYYY% = Year(dlpDateRange(1).Text)
'        dtMM% = Month(dlpDateRange(1).Text)
'        dtDD% = Day(dlpDateRange(1).Text)
'        SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_TODATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'    End If
    
    Me.vbxCrystal.SelectionFormula = SQLQ
    Me.vbxCrystal.Destination = 1
    Me.vbxCrystal.Action = 1
    
    'cmdPrint.Enabled = True

End Sub

Sub cmdView_Click()
    Dim RHeading As String, xReport, x%
    Dim SQLQ
    Dim dtYYY%, dtMM%, dtDD%
    'cmdPrint.Enabled = False
    
    Me.vbxCrystal.Reset
    Me.vbxCrystal.WindowTitle = "Salary % Increase Report"
    Call setRptLabel(Me, 0) '1)
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 5
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgsalincrmst.rpt"
    
    SQLQ = "(1=1) "
    If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_DIV} = '" & clpDiv.Text & "'"
    If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_DEPT} = '" & clpDept.Text & "'"
    If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_ORG} = '" & clpCode(0).Text & "'"
    'If Len(dlpAsOf.Text) > 0 Then
    '    dtYYY% = Year(dlpAsOf.Text)
    '    dtMM% = Month(dlpAsOf.Text)
    '    dtDD% = Day(dlpAsOf.Text)
    '    SQLQ = SQLQ & " AND {HRSICKENT.VE_EDATE} = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    'End If
    If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_EMP} = '" & clpCode(1).Text & "'"
    If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_PT} = '" & clpPT.Text & "' "
    If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_GRPCD} = '" & clpCode(2).Text & "'"
    If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_SECTION} = '" & clpCode(3).Text & "'"
    If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_LOC} = '" & clpCode(4).Text & "'"
    
    'Ticket #15276 - Commented out
'    If Len(dlpDateRange(0).Text) > 0 Then
'        dtYYY% = Year(dlpDateRange(0).Text)
'        dtMM% = Month(dlpDateRange(0).Text)
'        dtDD% = Day(dlpDateRange(0).Text)
'        SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_FRDATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'    End If
'    If Len(dlpDateRange(1).Text) > 0 Then
'        dtYYY% = Year(dlpDateRange(1).Text)
'        dtMM% = Month(dlpDateRange(1).Text)
'        dtDD% = Day(dlpDateRange(1).Text)
'        SQLQ = SQLQ & " AND {HR_SALARY_INCR.SL_TODATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'    End If
    
    
    Me.vbxCrystal.SelectionFormula = SQLQ
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.Action = 1
    'cmdPrint.Enabled = True
End Sub

Private Sub cmdPrintAll_Click()
    Dim RHeading As String, xReport, x%
    Dim SQLQ
    Dim dtYYY%, dtMM%, dtDD%
    cmdPrintAll.Enabled = False
    
    Me.vbxCrystal.Reset
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    Me.vbxCrystal.WindowTitle = "Salary % Increase Report"
    Call setRptLabel(Me, 0) '1)
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 5
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgsalincrmst.rpt"
    Me.vbxCrystal.Action = 1
    
    cmdPrintAll.Enabled = True
End Sub

'Private Sub cmdRecalc_Click()
'Dim lastday
'Dim flglastdate As Boolean
'Dim lngRecs As Long, pct As Long, prec As Long
'Dim doDate As Date
'Dim bmk As Variant
'
'On Error GoTo EH
'
'bmk = Data1.Recordset.Bookmark
'Screen.MousePointer = vbHourglass
'
'Call getWSQLQ("C")
'Call EntReCalcPeriod(fglbESQLQ, "SICK", , , dlpDateRange(0), dlpDateRange(1))
'
'
'
'If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
'        Data1.Recordset.MoveFirst
'        Do
'            Call Display_Value
'            If Len(dlpAsOf.Text) = 0 Then
'                MsgBox "Effective Date is required field"
'                dlpAsOf.SetFocus
'                GoTo ExH
'            End If
'            If (fglbCompMonthly Or UCase(glbCompEntVac$) = "N") And Not (glbCompSerial = "S/N - 2355W" And chkManual.Value = -1) Then
'                prec = 0
'                Call getWSQLQ("C")
'                gdbAdoIhr001.Execute "UPDATE HREMP SET ED_ANNSICK=0 WHERE " & fglbESQLQ
'                If Not CR_SnapEntitle() Then Exit Sub  ' create snapEntitle (form level recordset)
'                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
'                    While Not snapEntitle.EOF
'                        lngRecs = snapEntitle.RecordCount
'                        prec = prec + 1
'                        pct = Int(100 * (prec / lngRecs))
'                        MDIMain.panHelp(0).FloodPercent = pct
'
'                        doDate = dlpAsOf
'                        fglbAsOf = snapEntitle("ED_EFDATES")
'                        For fglbRunTimes = 1 To 12
'                            If Not modAnnSelection() Then Exit Sub
'                            fglbAsOf = DateAdd("m", 1, CVDate(fglbAsOf))
'                            DoEvents
'                        Next
'                        snapEntitle.MoveNext
'                    Wend
'                    MDIMain.panHelp(0).FloodType = 0
'                End If
'
'            Else
'                prec = 0
'                Call getWSQLQ("C")
'                gdbAdoIhr001.Execute "UPDATE HREMP SET ED_ANNSICK=0 WHERE " & fglbESQLQ
'                If Not CR_SnapEntitle() Then Exit Sub  ' create snapEntitle (form level recordset)
'                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
'                    While Not snapEntitle.EOF
'                        lngRecs = snapEntitle.RecordCount
'                        prec = prec + 1
'                        pct = Int(100 * (prec / lngRecs))
'                        MDIMain.panHelp(0).FloodPercent = pct
'
'                        doDate = dlpAsOf
'                        fglbAsOf = snapEntitle("ED_EFDATES")
'
'                        If Not modAnnSelection() Then Exit Sub
'                        DoEvents
'
'                        snapEntitle.MoveNext
'                    Wend
'                    MDIMain.panHelp(0).FloodType = 0
'                End If
'
'            End If
'            Data1.Recordset.MoveNext
'        Loop Until Data1.Recordset.EOF
'    End If
'    Screen.MousePointer = vbDefault
'    Data1.Recordset.Bookmark = bmk
'    Call Display_Value
'
'ExH:
'    Screen.MousePointer = vbDefault
'    Exit Sub
'EH:
'
'    Resume ExH
'End Sub

'Private Sub cmdUpdate_Click()
'On Error GoTo Mod_Err
'Dim sFlag As Boolean
'
'If Not gSec_Upd_Entitlements Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If
'
'If Not chkMUEntitle() Then Exit Sub
'    'Added by Bryan 25/Oct/05 Ticket#9560
'    'made the code a separate sub because it's being used in two places
''    sFlag = DoWork
'
'Data1.Refresh
'Call Display_Value
'
'MsgBox "Update Completed Successfully", vbInformation + vbOKOnly, "Sick Entitlements"
'
'Screen.MousePointer = DEFAULT
'
'Exit Sub
'
'Mod_Err:
'
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdateAll", "Single", "Modify")
'Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
'     RollBack
'    Resume Next
'Else
'    Unload Me
'End If
'End Sub

'Private Function GetFTEtot(EmpNo, dblFTE)
'Dim rsFTE As New ADODB.Recordset
'Dim SQLQ, xFte
'    xFte = dblFTE
'    If glbMulti Then
'        If Len(Memplist1) > 0 Then
'            If InStr(1, Memplist1, "'" & EmpNo & "'") > 0 Then 'this EmpNo is in Memplist1
'                SQLQ = "SELECT JH_EMPNBR, SUM(JH_FTENUM) AS TOTFTE FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & EmpNo & " "
'                SQLQ = SQLQ & "GROUP BY JH_EMPNBR "
'                rsFTE.Open SQLQ, gdbAdoIhr001, adOpenStatic
'                If Not rsFTE.EOF Then
'                    If Not IsNull(rsFTE("TOTFTE")) Then
'                        xFte = rsFTE("TOTFTE")
'                    End If
'                End If
'                rsFTE.Close
'            End If
'        End If
'    End If
'    GetFTEtot = xFte
'End Function

'Private Function AccuValForMulti(EmpNo, dblEnt) ' Ticket #3304
''For multi positions and annual update, accumulate all entitlement of positions together
''and then replace the entitlement.
'Dim xVal
'    xVal = 0
'    If glbMulti Then
'        If Len(Memplist1) > 0 Then
'            If InStr(1, Memplist1, "'" & EmpNo & "'") > 0 Then 'this EmpNo is in Memplist1
'                If InStr(1, Memplist2, "'" & EmpNo & "'") > 0 Then 'this EmpNo is in Memplist2
'                    'xVal = 0 ' First time replace the Emtitlement with the New one
'                    Memplist2 = Replace(Memplist2, "'" & EmpNo & "',", ",")
'                Else
'                    xVal = dblEnt 'from Second time, accumulate the entitlement
'                End If
'            End If
'        End If
'    End If
'    AccuValForMulti = xVal
'End Function

'Private Function CalcASLRepaid(xEmpNo, xAsofDate, dblEntUpd, dblNewEnt, dblEnt#) '
'Dim rsASL As New ADODB.Recordset
'Dim rsENT As New ADODB.Recordset
'Dim SQLQ, xTaken, xRepaid, xOutStand
'Dim xSickEnt
'
''Hemu
''    Dim tmpTestData As String
''    Dim tmpAdoIHRTest As String
''    Dim sSetting As String
''    Dim sPath1 As String
''    Dim giGar1 As Integer
''
''    sPath1 = REG_NAME & "INFOHR Files"
''
''    sSetting = "IHRREPORTS"  'Compressed database location
''    tmpTestData = glbWorkDir
''    giGar1 = bGetRegistrySetting(lCurrentKey, sPath1, sSetting, tmpTestData)
''    tmpTestData = tmpTestData & IIf(Right$(tmpTestData, 1) <> "\", "\", "") & "TestData.mdb"
''    tmpAdoIHRTest = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=petman;Data Source=" & tmpTestData
''Hemu
'
'
'    xSickEnt = dblEntUpd
'    SQLQ = "SELECT ED_EMPNBR,ED_EFDATES,ED_ETDATES,ED_SICK,ED_SICKT FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
'    rsENT.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    If Not rsENT.EOF Then
'        If IsDate(rsENT("ED_EFDATES")) And IsDate(rsENT("ED_ETDATES")) Then
'            'If rsENT("ED_EFDATES") <= CVDate(xAsofDate) And rsENT("ED_ETDATES") >= CVDate(xAsofDate) Then
'                SQLQ = "SELECT AS_EMPNBR, Sum(AS_HRSTAK) AS TAKEN, Sum(AS_HRSREP) AS REPAID FROM WHSCC_ASL WHERE AS_EMPNBR = " & xEmpNo & " "
'                'Don't check Date Range for ASL T#3304
'                'SQLQ = SQLQ & "AND AS_EFDATES = ('" & Format(rsENT("ED_EFDATES"), "mmm dd,yyyy") & "') "
'                'SQLQ = SQLQ & "AND AS_ETDATES = ('" & Format(rsENT("ED_ETDATES"), "mmm dd,yyyy") & "') "
'                SQLQ = SQLQ & "GROUP BY AS_EMPNBR "
'                rsASL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                xTaken = 0: xRepaid = 0: xOutStand = 0
'                If Not rsASL.EOF Then
'                    If IsNull(rsASL("TAKEN")) Then
'                        xTaken = 0
'                    Else
'                        xTaken = rsASL("TAKEN")
'                    End If
'                    If IsNull(rsASL("REPAID")) Then
'                        xRepaid = 0
'                    Else
'                        xRepaid = rsASL("REPAID")
'                    End If
'                    xOutStand = xTaken - xRepaid
'                End If
'                rsASL.Close
'
'                'Logic changed:
'                'Repaid = Sick Entitlement, before Repaid = ASL Outstanding
'
'                'xOutStand = dblEntUpd
'                If xOutStand > 0 Then
'                    If xOutStand >= dblNewEnt Then
'                        xSickEnt = dblEnt#
'                    Else
'                        xSickEnt = dblEnt# + dblNewEnt - xOutStand
'                        dblNewEnt = xOutStand
'                    End If
''Hemu
''If glbWHSCC Then
'''include the dummy test table here
''    SQLQ = "SELECT * FROM WHSCC_ASL WHERE AS_EMPNBR = " & xEmpNo & " "
''    rsASL.Open SQLQ, tmpAdoIHRTest, adOpenKeyset, adLockOptimistic
''    rsASL.AddNew
''    rsASL("AS_HRSTAK") = 0
''    rsASL("AS_COMPNO") = "001"
''    rsASL("AS_EMPNBR") = xEmpNo
''    rsASL("AS_DOA") = xAsofDate
''    rsASL("AS_CODE") = "REPA"
''    rsASL("AS_HRSREP") = dblNewEnt 'dblEntUpd
''    rsASL("AS_HRSOS") = xOutStand - dblNewEnt 'dblEntUpd
''    rsASL("AS_LDATE") = Format(Now, "SHORT DATE")
''    rsASL("AS_LTIME") = Time$
''    rsASL("AS_LUSER") = glbUserID
''    rsASL.Update
''    rsASL.Close
''
''    GoTo End_Test_Data
''End If
''Hemu
'                    SQLQ = "SELECT * FROM WHSCC_ASL WHERE AS_EMPNBR = " & xEmpNo & " "
'                    'Don't check Date Range for ASL T#3304
'                    'SQLQ = SQLQ & "AND AS_EFDATES = ('" & Format(rsENT("ED_EFDATES"), "mmm dd,yyyy") & "') "
'                    'SQLQ = SQLQ & "AND AS_ETDATES = ('" & Format(rsENT("ED_ETDATES"), "mmm dd,yyyy") & "') "
'                    SQLQ = SQLQ & "AND AS_DOA = ('" & Format(xAsofDate, "mmm dd,yyyy") & "') "
'                    SQLQ = SQLQ & "AND AS_CODE = 'REPA' "
'                    rsASL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                    'If rsASL.EOF Then
'                        rsASL.AddNew
'                        rsASL("AS_HRSTAK") = 0
'                    'Else
'                        'dblEntUpd = dblEntUpd + rsASL("AS_HRSREP")
'                    'End If
'                    rsASL("AS_COMPNO") = "001"
'                    rsASL("AS_EMPNBR") = xEmpNo
'                    rsASL("AS_DOA") = xAsofDate
'                    rsASL("AS_CODE") = "REPA"
'                    rsASL("AS_HRSREP") = dblNewEnt 'dblEntUpd
'                    rsASL("AS_HRSOS") = xOutStand - dblNewEnt 'dblEntUpd
'                    'rsASL("AS_EFDATES") = rsENT("ED_EFDATES")
'                    'rsASL("AS_ETDATES") = rsENT("ED_ETDATES")
'                    rsASL("AS_LDATE") = Format(Now, "SHORT DATE")
'                    rsASL("AS_LTIME") = Time$
'                    rsASL("AS_LUSER") = glbUserID
'                    rsASL.Update
'                    rsASL.Close
'                    Call ReCalcASL(xEmpNo, "")
'                    'SQLQ = "UPDATE WHSCC_ASL SET AS_HRAOS = 0 "
'                    'SQLQ = SQLQ & "WHERE AS_EMPNBR = " & xEmpNo & " "
'                    'SQLQ = SQLQ & "AND AS_EFDATES = ('" & Format(rsENT("ED_EFDATES"), "mmm dd,yyyy") & "') "
'                    'SQLQ = SQLQ & "AND AS_ETDATES = ('" & Format(rsENT("ED_ETDATES"), "mmm dd,yyyy") & "') "
'                    'gdbAdoIhr001.Execute SQLQ
'                End If
'            'End If
'        End If
'    End If
'
''Hemu
''exit after update to dummy table
''End_Test_Data:
''    'Update Test_data table
''    Dim rsTestData As New ADODB.Recordset
''    SQLQ = "SELECT * FROM test_data"
''    rsTestData.Open SQLQ, tmpAdoIHRTest, adOpenKeyset, adLockOptimistic
''    rsTestData.AddNew
''    rsTestData("ED_EMPNBR") = xEmpNo
''    rsTestData("JH_DHRS") = tmpDHrs
''    rsTestData("JH_FTENUM") = tmpFTETotHrs
''    rsTestData("ED_EMP") = txtCode(3).Text
''    rsTestData("ED_PT") = txtPT.Text
''    rsTestData("Max_Entit") = medMax(0).Text
''    rsTestData("Max_Entit_Calc") = tmpNewMax
''    rsTestData("New_Entitlement") = tmpNewEntit
''    rsTestData("Old_Entitlement") = tmpOldEntit
''    rsTestData("Entit_Update") = tmpEntitUpd
''    rsTestData.Update
''    rsTestData.Close
''Hemu
'
'    rsENT.Close
'    CalcASLRepaid = xSickEnt
'End Function
'Private Function modUpdateSelectionWHSCC()
'Dim EmpNo As Long
'Dim dblEntitle#, dblPrevEntitle#, dblTKEEntitle#, strDivision$
'Dim strJob$, dblServiceYears#
'Dim spt As Variant, varStartDate As Variant, lngRecs&
'Dim dblDHours#, intWhereFit&, X%, Y%, z%, dblNewEntitle#
'Dim dblFTEHours#, dblFTEHoursTot#
'Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
'Dim Msg$, Title$, DgDef As Variant
'Dim Response%, pct%
'Dim prec%
'Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
'Dim if_Entitle As Boolean, if_Vacation As Boolean
'Dim ifAnnual As Boolean, dblNewEntAnn#, VacpcNAnn, ifUnionDate As Boolean, ifFirstDate As Boolean, xAsOf 'Frank for WHSCC
'Dim dblServiceYearsYTD, if_NON As Boolean
'Dim NoUptSickList As String
'Dim xComments
'' Entitlements are always valued in HOURS - if you enter days then it
''   works out how many hours (based on average Hrswrked/day found in salary master record)
'On Error GoTo modUpdateSelectionWHSCC_Err
'modUpdateSelectionWHSCC = False
'
'If Len(dlpAsOf.Text) = 0 Then
'    MsgBox "Effective Date is required field"
'    dlpAsOf.SetFocus
'    Exit Function
'End If
'
'If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
''
''If fTablHREMP.State <> 0 Then fTablHREMP.Close
''fTablHREMP.Open "HREMP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
'Screen.MousePointer = DEFAULT
'
'
'If snapEntitle.BOF And snapEntitle.EOF Then
'    MsgBox "Employees for this selection do not exist!"
'    Exit Function
'Else
'    lngRecs& = snapEntitle.RecordCount
'    Msg$ = lngRecs& & " Records to process" & Chr(10) & "Would You Like To Proceed?"
'    Title$ = "Update Entitlements"
'    DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
'    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
'    If Response% = IDNO Then    ' Evaluate response
'        Exit Function
'    End If
'    Screen.MousePointer = HOURGLASS
'End If
'MDIMain.panHelp(0).FloodType = 1
'MDIMain.panHelp(0).FloodPercent = 5
'
''Ticket# 3856
''If the employee's Employment Status is one of those on the list,
''do not update the employee's sick entitlement for that month. Linda Rowland
'NoUptSickList = ",BD,CAS,CLIN,CONT,EIS,LTD,MAT,PAR,STUD,"
'
'For X% = 0 To 24
'    If Not IsNumeric(medLTServ(X%)) Then Exit For ' medLTServ(X%) = 0
'    If Not IsNumeric(medGTServ(X%)) Then
'      medGTServ(X%) = 0
'    Else
'      If Val(medGTServ(X%)) = Int(medGTServ(X%)) Then medGTServ(X%) = medGTServ(X%) + 0.99
'    End If
'    If medLTServ(X%) > 0 And medGTServ(X%) = 0 Then medGTServ(X%) = 9999999
'Next
'
''Hemu
''If Not glbWHSCC Then
''Hemu
'    gdbAdoIhr001.BeginTrans
''End If
'
'While Not snapEntitle.EOF
'    prec% = prec% + 1
'    pct% = Int(100 * (prec% / lngRecs&))
'    MDIMain.panHelp(0).FloodPercent = pct%
'    if_Entitle = False
'    if_Vacation = False
'
'    EmpNo& = snapEntitle("ED_EMPNBR")
'
'    If Not IsNull(snapEntitle("ED_EMP")) Then
'        If InStr(1, NoUptSickList, "," & Trim(snapEntitle("ED_EMP")) & ",") > 0 Then
'            GoTo lblNextRec
'        End If
'    End If
'
'
'    If IsNull(snapEntitle("ED_SICK")) Then
'        dblEntitle# = 0
'    Else
'        dblEntitle# = snapEntitle("ED_SICK")
'    End If
'
'
'    If IsNull(snapEntitle("ED_PSICK")) Then
'        dblPrevEntitle# = 0
'    Else
'        dblPrevEntitle# = snapEntitle("ED_PSICK")
'    End If
'
'    If IsNull(snapEntitle("ED_SICKT")) Then
'        dblTKEEntitle# = 0
'    Else
'        dblTKEEntitle# = snapEntitle("ED_SICKT")
'    End If
'
'    spt = snapEntitle("ED_PT")
'    strDivision$ = snapEntitle("ED_DIV")
'
'    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec
'
'    varStartDate = snapEntitle(fglbWDate$)
'
'    If Not IsNumeric(snapEntitle("JH_DHRS")) Then
'        dblDHours# = 0
'    Else
'        dblDHours# = snapEntitle("JH_DHRS")
'    End If
'
''Hemu
''    tmpDHrs = dblDHours#
''Hemu
'
'    If Not IsNumeric(snapEntitle("JH_FTENUM")) Then
'        dblFTEHours# = 0
'    Else
'        dblFTEHours# = snapEntitle("JH_FTENUM")
'    End If
'    dblFTEHoursTot# = GetFTEtot(EmpNo&, dblFTEHours#) 'For Multi Position, get the Total of FTE for one employee
'
''Hemu
''    tmpFTETotHrs = dblFTEHoursTot#
''Hemu
'
'    'Franks Jul 31, 02 for WHSCC
'    ifAnnual = False
'    ifUnionDate = False
'    ifFirstDate = False
'
'
'    ' dkostka - 08/13/2001 - Changed formula from using number of days / 365 * 12 to using DateDiff
'    '   directly to get number of months.  We don't get decimals here but the value is always correct.
'    '   Using the old formula would cause problems sometimes because it assumes all months have an
'    '   equal number of days, and all years are 365 days.
'    'dblServiceYears# = (DateDiff("d", varStartDate, CVDate(dlpAsOf)) / 365) * 12
'    If Not ifAnnual Then
'        'dblServiceYears# = DateDiff("m", varStartDate, CVDate(dlpAsOf))
'        If Not if_NON Then
'            'dblServiceYears# = DateDiff("m", varStartDate, CVDate(dlpAsOf))
'            dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(dlpAsOf))
'        Else
'            dblServiceYears# = dblServiceYearsYTD
'        End If
'        intWhereFit& = -1   ' first record can be just less than
'
'        For X% = 0 To 24
'            If medGTServ(X%) > 0 Then
'                If dblServiceYears# >= CDbl(medLTServ(X%)) And dblServiceYears# <= CDbl(medGTServ(X%)) Then
'                    intWhereFit& = X%
'                    If Len(medEntitle(X%)) > 0 Then if_Entitle = True
'                    Exit For
'                End If
'            End If
'        Next X%
'
'        If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
'    Else 'Franks Jul 31, 02 for WHSCC
'        xAsOf = CVDate("Jan 1," & Year(dlpAsOf))
'        dblNewEntAnn# = 0
'        VacpcNAnn = 0
'        intWhereFit& = 0
'        For z% = 1 To 12
'            'dblServiceYears# = (DateDiff("d", varStartDate, CVDate(xAsOf)) / 365) * 12
'            dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
'            'If there is date of Union Date or First Day on Status/Dates screen,
'            'use the special vacation rules, otherwise use the rules on the Vacation Master screen
'            If Not (ifUnionDate Or ifFirstDate) Then
'                For X% = 0 To 24
'                    If medGTServ(X%) > 0 Then
'                        If dblServiceYears# >= CDbl(medLTServ(X%)) And dblServiceYears# <= CDbl(medGTServ(X%)) Then
'                            intWhereFit& = X%
'                            If Len(medEntitle(X%)) > 0 Then
'                                if_Entitle = True
'                                dblNewEntAnn# = dblNewEntAnn# + medEntitle(X%)
'                            End If
'                            Exit For
'                        End If
'                    End If
'                Next X%
'            Else
'                If ifUnionDate Then
'                    If dblServiceYears# >= 0 And dblServiceYears# < 48.99 Then
'                            if_Entitle = True
'                            dblNewEntAnn# = dblNewEntAnn# + 1.25
'                    End If
'                    If dblServiceYears# >= 49 And dblServiceYears# < 239.99 Then
'                            if_Entitle = True
'                            dblNewEntAnn# = dblNewEntAnn# + 1.67
'                    End If
'                    If dblServiceYears# >= 240 And dblServiceYears# < 999.99 Then
'                            if_Entitle = True
'                            dblNewEntAnn# = dblNewEntAnn# + 2.09
'                    End If
'                End If
'                If ifFirstDate Then
'                    If dblServiceYears# >= 0 And dblServiceYears# < 11.99 Then
'                            if_Entitle = True
'                            dblNewEntAnn# = dblNewEntAnn# + 1.25
'                    End If
'                    If dblServiceYears# >= 12 And dblServiceYears# < 95.99 Then
'                            if_Entitle = True
'                            dblNewEntAnn# = dblNewEntAnn# + 1.67
'                    End If
'                    If dblServiceYears# >= 96 And dblServiceYears# < 239.99 Then
'                            if_Entitle = True
'                            dblNewEntAnn# = dblNewEntAnn# + 2.09
'                    End If
'                    If dblServiceYears# >= 240 And dblServiceYears# < 999.99 Then
'                            if_Entitle = True
'                            dblNewEntAnn# = dblNewEntAnn# + 2.5
'                    End If
'                End If
'            End If
'            xAsOf = DateAdd("m", 1, xAsOf)
'        Next z%
'    End If 'Franks Jul 31, 02 for WHSCC
'    ' Two variables glbCompEntVac$ = "M" And glbCompEntSick$ = "M"    are 'company' level
'    ' which represents if Sick and Vacation entitlements
'    ' are determined on monthly basis (vs yearly) - these are stored in table hrpasco
'    ' and read on system startup.
'
'    ' In this routine we work independantly of SICK/VACATIon entitlement.
'    '  fglbCompMonthly% - is the independant representation
'        'of glbCompEntVac$ = "M" And glbCompEntSick$ = "M"
'        'Procedure modUpdateSelectionWHSCC is used to set
'        'fglbCompMonthly based on values it finds for global variables
'        ' and what the user wants to manipulate (sick/Vac)
'
'    'optD indicates if Entitlement entered is Daily or yearly based
'    ' if daily then max entitlement is based on entitlement * hours they work.
'
'    ' we have   Entitle = existing entitmenet (stored presently
'    '           NewEntitle = amount entered onto screen = medentitle(index)
'    '           EntitleUpd  = value to update record with
'
'    If if_Entitle Then
'        If ifAnnual Then
'            dblNewEntitle# = dblNewEntAnn#
'            If optD(intWhereFit&) = True Then           ' Entitlements entered in days
'                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblDHours#
'                dblNewEntitle# = dblNewEntitle# * dblDHours#
'                dblEntitleUpd = dblNewEntitle
'            End If
'            If optF(intWhereFit&) = True Then
'                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHoursTot# * dblDHours#
'                dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
'                dblEntitleUpd = dblNewEntitle
'            End If
'            If fglbCompMonthly% Then
'                dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
'            Else
'                'dblEntitleUpd = dblNewEntitle ' rollover is in other utility (to accumulate)
'                dblEntitleUpd# = dblNewEntitle + AccuValForMulti(EmpNo&, dblEntitle#) 'MultiPos Update
'            End If
'            If dblNewMax <> 0 Then          'only do if not zero
'                    If (dblPrevEntitle# + dblEntitle# - dblTKEEntitle# + dblNewEntitle) > dblNewMax Then
'                        dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
'                    End If
'            End If
'        Else
'            dblNewEntitle# = medEntitle(intWhereFit&)
'            dblNewMax# = 0
'            If optD(intWhereFit&) = True Then           ' Entitlements entered in days
'                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblDHours#
'                dblNewEntitle# = dblNewEntitle# * dblDHours#
'                dblEntitleUpd = dblNewEntitle
'            End If
'            If optF(intWhereFit&) = True Then
'                'If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHours# * dblDHours#
'                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHoursTot# * dblDHours#
'                dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
'                dblEntitleUpd = dblNewEntitle
'            End If
'            If optH(intWhereFit&) = True Then
'                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&)
'            End If
'            If fglbCompMonthly% Then
'                dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
'            Else
'                'dblEntitleUpd = dblNewEntitle ' rollover is in other utility (to accumulate)
'                dblEntitleUpd# = dblNewEntitle + AccuValForMulti(EmpNo&, dblEntitle#) 'MultiPos Update
'            End If
'
'            If dblNewMax <> 0 Then          'only do if not zero
'                    If (dblPrevEntitle# + dblEntitle# - dblTKEEntitle# + dblNewEntitle) > dblNewMax Then
'                        dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
'                    End If
'            End If
'        End If
'        DtTm = Now
'    End If
'
'
'    If if_Entitle Then
'
'        'If optSickE.Value Then
'            'For Sick Entitlement update, check the ASL Bank first.
'            'If ASL Bank is greater than 0, take Repaid ASL from it
'            'Otherwise, assign the amount to the Sick Entitlement(ED_SICK)
'        dblEntitleUpd = CalcASLRepaid(EmpNo, CVDate(dlpAsOf), dblEntitleUpd, dblNewEntitle, dblEntitle#) 'dblEntitleUpd)
'
'        xComments = "Current Sick. Ent. Chg from " & snapEntitle("ED_SICK") & " to " & dblEntitleUpd
'        'Hemu - Ticket #11925 - Changed the Accrual Date from Effective Date to Entitlement Start Date
'        'because otherwise it will not update Vadim until the date arrives in case it's not same as the
'        'Entitlement Start Date.
'        'Call Append_Accrual(EmpNo&, "SICK", dlpAsOf, dblEntitleUpd - Val(snapEntitle("ED_SICK") & ""), "U", xComments)
'        Call Append_Accrual(EmpNo&, "SICK", dlpDateRange(0), dblEntitleUpd - Val(snapEntitle("ED_SICK") & ""), "U", xComments)
'
'        snapEntitle("ED_SICK") = dblEntitleUpd
'
'    End If
'    snapEntitle("ED_ANNSICK") = snapEntitle("ED_SICK")
'    snapEntitle.Update
'
'
'
'lblNextRec:
'    snapEntitle.MoveNext
'
'Wend
'modUpdateSelectionWHSCC = True
'MDIMain.panHelp(0).FloodType = 0
'
''Hemu
''If Not glbWHSCC Then
''Hemu
'gdbAdoIhr001.CommitTrans
''End If
'
''fTablHREMP.Close
'
'snapEntitle.Close
'
'Screen.MousePointer = DEFAULT
'
'Exit Function
'
'modUpdateSelectionWHSCC_Err:
''These errors are:
''13=type mismatch
''94=invalid use of null
''3018=couln't find field 'item'
'If Err = 13 Or Err = 94 Or Err = 3018 Then
'   ' MsgBox "Err:" & Str(Err) & Chr(10) & Error$ & Chr(10) & " modUpdateSelectionWHSCC" & Chr(10) & "FORM:FUENTITL.FRM"
'    'commented out by RAUBREY 5/20/97
'    Err = 0
'    Resume Next
'End If
'
'Screen.MousePointer = DEFAULT
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateEntitle", "HR_EMP", "edit/Add")
'Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
'    'Rollback
'    Resume Next
'Else
'    Unload Me
'End If
'End Function

Private Sub cmdUpdate_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

'Private Function CR_SnapEntitle()
'Dim SQLQ As String
'Dim SQLQ2 As String
'Dim snapMultiEmp As New ADODB.Recordset
'
'CR_SnapEntitle = False
'On Error GoTo CR_SnapEntitle_Err
'
'
'Call getWSQLQ("")
'If glbWHSCC Then
'    SQLQ = "SELECT HREMP.ED_EMPNBR, qry_JobCurrent.JB_GRPCD, HREMP.ED_VACPC, HREMP.ED_PVAC, HREMP.ED_VAC, HREMP.ED_VACT, HREMP.ED_ANNVAC, HREMP.ED_ANNSICK, "
'    SQLQ = SQLQ & " HREMP.ED_PSICK, HREMP.ED_SICK, HREMP.ED_SICKT,qry_JobCurrent.JH_DHRS, HREMP.ED_DIV, HREMP.ED_EMP, "
'    SQLQ = SQLQ & " HREMP.ED_DEPTNO, HREMP.ED_PT, HREMP.ED_DOH, HREMP.ED_SENDTE, HREMP.ED_UNION, HREMP.ED_LTHIRE, HREMP.ED_USRDAT1, HREMP.ED_ORG, HREMP.ED_FDAY, qry_JobCurrent.JH_FTENUM, qry_JobCurrent.JH_DHRS, HREMP.ED_SECTION "
'    SQLQ = SQLQ & " FROM HREMP LEFT JOIN qry_JobCurrent ON HREMP.ED_EMPNBR = qry_JobCurrent.JH_EMPNBR "
'    SQLQ = SQLQ & " WHERE " & fglbESQLQ
'Else
'    SQLQ = "SELECT ED_EMPNBR,ED_VACPC,ED_PVAC,ED_VAC,ED_PSICK,ED_SICK,ED_SICKT,ED_EFDATES,ED_ETDATES, HREMP.ED_ANNVAC, HREMP.ED_ANNSICK, "
'    SQLQ = SQLQ & " ED_DIV,ED_PT, ED_SECTION, ED_LOC, ED_EMP,"
'    SQLQ = SQLQ & " ED_HIRECODE," 'County of Brant Ticket #12525
'    SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME "
'    SQLQ = SQLQ & " FROM HREMP WHERE " & fglbESQLQ
'End If
'If Len(clpCode(2).Text) > 0 Then
'    SQLQ = SQLQ & " AND ED_EMPNBR IN "
'    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
'    SQLQ = SQLQ & " WHERE JB_GRPCD = '" & clpCode(2).Text & "') "
'End If
'
''Multi Positions Update #3304
'If glbMulti Then
'    SQLQ2 = "SELECT HREMP.ED_EMPNBR, COUNT(ED_EMPNBR) AS SUMEMP "
'    SQLQ2 = SQLQ2 & " FROM HREMP LEFT JOIN qry_JobCurrent ON HREMP.ED_EMPNBR = qry_JobCurrent.JH_EMPNBR "
'    SQLQ2 = SQLQ2 & " WHERE " & fglbESQLQ
'    If Len(clpCode(2).Text) > 0 Then
'        SQLQ2 = SQLQ2 & " AND ED_EMPNBR IN "
'        SQLQ2 = SQLQ2 & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
'        SQLQ2 = SQLQ2 & " WHERE JB_GRPCD = '" & clpCode(2).Text & "') "
'    End If
'
'    Memplist1 = "": Memplist2 = ""
'    If UCase(glbCompEntVac$) = "A" Or UCase(glbCompEntSick$) = "A" Then
'        'SQLQ2 = SQLQ2 & SQLQ
'        If snapMultiEmp.State <> 0 Then snapMultiEmp.Close
'        SQLQ2 = SQLQ2 & " GROUP BY ED_EMPNBR HAVING COUNT(ED_EMPNBR) > 1 "
'        snapMultiEmp.Open SQLQ2, gdbAdoIhr001, adOpenStatic
'        Do While Not snapMultiEmp.EOF
'            Memplist1 = Memplist1 & "'" & snapMultiEmp("ED_EMPNBR") & "',"
'            Memplist2 = Memplist2 & "'" & snapMultiEmp("ED_EMPNBR") & "',"
'            snapMultiEmp.MoveNext
'        Loop
'        snapMultiEmp.Close
'    End If
'End If
''Multi Positions Update #3304
'
'If snapEntitle.State <> 0 Then snapEntitle.Close
'If glbOracle Then
'    snapEntitle.CursorLocation = adUseServer
'End If
'snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
'
'CR_SnapEntitle = True
'
'Exit Function
'
'CR_SnapEntitle_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapEntitle", "Entitlements/EMP", "Select")
'
'If gintRollBack% = False Then
'    Resume Next
'Else
'    Unload Me
'End If
'
'End Function


'Private Sub cmdUpdateAll_Click()
'On Error GoTo Mod_Err
'
'Dim c As Long
'Dim failed As String
'
'If Not gSec_Upd_Entitlements Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If
'
'failed = ""
'c = 1
'If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
'    Data1.Recordset.MoveFirst
'    Do
'        Call Display_Value
'
'        'made the DoWork a separate sub because it's being used in two places
'        If chkManual.Value = False Then
'            If chkMUEntitle() Then
'                If DoWork = False Then
'                    failed = failed & "Rule " & CStr(c) & ": "
'                    If Not IsNull(Data1.Recordset("VE_DIV")) Then failed = failed & Data1.Recordset("VE_DIV") & ", "
'                    If Not IsNull(Data1.Recordset("VE_DEPT")) Then failed = failed & Data1.Recordset("VE_DEPT") & ", "
'                    If Not IsNull(Data1.Recordset("VE_ORG")) Then failed = failed & Data1.Recordset("VE_ORG") & ", "
'                    If Not IsNull(Data1.Recordset("VE_EDATE")) Then failed = failed & Data1.Recordset("VE_EDATE") & ", "
'                    If Not IsNull(Data1.Recordset("VE_EMP")) Then failed = failed & Data1.Recordset("VE_EMP") & ", "
'                    If Not IsNull(Data1.Recordset("VE_PT")) Then failed = failed & Data1.Recordset("VE_PT") & ", "
'                    If Not IsNull(Data1.Recordset("VE_GRPCD")) Then failed = failed & Data1.Recordset("VE_GRPCD") & ", "
'                    If Not IsNull(Data1.Recordset("VE_LOC")) Then failed = failed & Data1.Recordset("VE_LOC") & ", "
'                    If Not IsNull(Data1.Recordset("VE_SECTION")) Then failed = failed & Data1.Recordset("VE_SECTION") & ", "
'                    If Not IsNull(Data1.Recordset("VE_FRDATE")) Then failed = failed & Data1.Recordset("VE_FRDATE") & ", "
'                    If Not IsNull(Data1.Recordset("VE_TODATE")) Then failed = failed & Data1.Recordset("VE_TODATE") & ", "
'                    failed = Left(failed, Len(failed) - 2) & vbCrLf
'                End If
'            Else
'                failed = failed & "Rule " & CStr(c) & ": "
'                If Not IsNull(Data1.Recordset("VE_DIV")) Then failed = failed & Data1.Recordset("VE_DIV") & ", "
'                If Not IsNull(Data1.Recordset("VE_DEPT")) Then failed = failed & Data1.Recordset("VE_DEPT") & ", "
'                If Not IsNull(Data1.Recordset("VE_ORG")) Then failed = failed & Data1.Recordset("VE_ORG") & ", "
'                If Not IsNull(Data1.Recordset("VE_EDATE")) Then failed = failed & Data1.Recordset("VE_EDATE") & ", "
'                If Not IsNull(Data1.Recordset("VE_EMP")) Then failed = failed & Data1.Recordset("VE_EMP") & ", "
'                If Not IsNull(Data1.Recordset("VE_PT")) Then failed = failed & Data1.Recordset("VE_PT") & ", "
'                If Not IsNull(Data1.Recordset("VE_GRPCD")) Then failed = failed & Data1.Recordset("VE_GRPCD") & ", "
'                If Not IsNull(Data1.Recordset("VE_LOC")) Then failed = failed & Data1.Recordset("VE_LOC") & ", "
'                If Not IsNull(Data1.Recordset("VE_SECTION")) Then failed = failed & Data1.Recordset("VE_SECTION") & ", "
'                If Not IsNull(Data1.Recordset("VE_FRDATE")) Then failed = failed & Data1.Recordset("VE_FRDATE") & ", "
'                If Not IsNull(Data1.Recordset("VE_TODATE")) Then failed = failed & Data1.Recordset("VE_TODATE") & ", "
'                failed = Left(failed, Len(failed) - 2) & vbCrLf
'            End If
'        End If
'        c = c + 1
'        Data1.Recordset.MoveNext
'    Loop Until Data1.Recordset.EOF
'End If
'
'Data1.Refresh
'Call Display_Value
'Screen.MousePointer = DEFAULT
'If Len(failed) = 0 Then
'    MsgBox "All Rules applied", vbInformation + vbOKOnly, "Sick Entitlements"
'Else
'    MsgBox "The Following Rules failed:" & vbCrLf & failed, vbInformation + vbOKOnly, "Sick Entitlements"
'End If
'Exit Sub
'
'Mod_Err:
'
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdateAll", "Single", "Modify")
'Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
'     RollBack
'    Resume Next
'Else
'    Unload Me
'End If
'End Sub

Private Sub Form_Activate()
    Call SET_UP_MODE
    Call INI_Controls(Me)
    
    glbOnTop = "FRMSALPERCTG"

End Sub

Private Sub Form_Load()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    Dim Answer, DefVal, Msg, Title  ' Declare variables.
    Dim RFound As Integer ' records found
    Dim x%
    Dim SQLQ
    
    glbOnTop = "FRMSALPERCTG"
    
    FlagRefresh = False
    
    Data1.ConnectionString = glbAdoIHRDB
    SQLQ = "SELECT DISTINCT SL_DIV,SL_DEPT,SL_ORG,SL_LOC,SL_SECTION,SL_EMP,SL_PT,SL_GRPCD,SL_FRDATE,SL_TODATE FROM HR_SALARY_INCR"
    
    If glbDIVCount = 1 And glbLinamar Then
        SQLQ = SQLQ & " WHERE SL_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
    End If
    Data1.RecordSource = SQLQ
    Data1.Refresh
    
'    If glbCBrant Then
'        'County of Brant using Sick Time Entitlement Outstanding Based Upon to calculate the service months
'        'Ticket #Ticket #12544
'        Select Case glbEntOutStandingS$
'            Case "2": fglbWDate$ = "ED_DOH"
'            Case "3": fglbWDate$ = "ED_SENDTE"
'            Case "4": fglbWDate$ = "ED_LTHIRE"
'            Case "5": fglbWDate$ = "ED_USRDAT1"
'            Case "6": fglbWDate$ = "ED_UNION"
'        End Select
'    Else
'        Select Case glbCompWDate$ ' sets field reference for basic 'which date'
'            Case "O": fglbWDate$ = "ED_DOH"
'            Case "S": fglbWDate$ = "ED_SENDTE"
'            Case "U": fglbWDate$ = "ED_UNION"
'            Case "L": fglbWDate$ = "ED_LTHIRE"
'            Case "D": fglbWDate$ = "ED_USRDAT1"
'        End Select
'    End If
'
'    If UCase(glbCompEntSick$) = "M" Or UCase(glbCompEntSick$) = "N" Then
'        vbxTrueGrid.Columns(5).Visible = False
'    End If
    
    Screen.MousePointer = HOURGLASS
    vbxTrueGrid.Columns(0).Caption = lStr(vbxTrueGrid.Columns(0).Caption)
    vbxTrueGrid.Columns(1).Caption = lStr(vbxTrueGrid.Columns(1).Caption)
    vbxTrueGrid.Columns(2).Caption = lStr(vbxTrueGrid.Columns(2).Caption)
    
    Call setRptCaption(Me)
    
    If glbSyndesis Then
        lblCriteria(5).Caption = "Position Grade"
        vbxTrueGrid.Columns(8).Caption = "Position Grade"
        clpCode(2).Tag = "00-Enter Position Grade"
    End If
    If glbWFC Then
        lblSection.FontBold = True
    End If
    
    Screen.MousePointer = DEFAULT
    
    Call INI_Controls(Me)
    
    If glbMulti Then textMulti.Visible = True
    
    ST_UPD_MODE (False)
    
    Screen.MousePointer = DEFAULT
End Sub

Private Sub Form_LostFocus()
    MDIMain.panHelp(0).Caption = " "
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(2).Caption = " "
    MDIMain.panHelp(3).Caption = " "
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Dim Keepfocus As Boolean
    'If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
    'Keepfocus = Not isUpdated(Me)
    'Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Resize()
    If Me.Height >= 3750 + VacFram.Height + panControls.Height + 230 Then
        scrControl.Value = 0
        VacFram.Top = 4440
        scrControl.Visible = False
        Exit Sub
    End If
    scrControl.Visible = True
    scrControl.Max = VacFram.Height + panControls.Height + 3750 + 550 - Me.Height '250 - Me.Height
    scrControl.Left = Me.Width - scrControl.Width - 120
    If Me.Height - scrControl.Top - panControls.Height - 300 > 0 Then
        scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 300
    Else
        scrControl.Height = 0
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select FROM the menu the appropriate function."
    
    Set frmUEntitle = Nothing  'carmen apr 2000
End Sub

Private Sub medEntitle_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
'If IsNumeric(medEntitle(Index)) Then
'    If Len(medEntitle(Index)) > 0 Then
'        medEntitle(Index) = medEntitle(Index) * 100
'    End If
'End If
End Sub

Private Sub medEntitle_LostFocus(Index As Integer)
'If IsNumeric(medEntitle(Index)) Then
'    If Len(medEntitle(Index)) > 0 Then
'        medEntitle(Index) = medEntitle(Index) / 100
'    End If
'End If
End Sub

Private Sub medGTServ_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medLTServ_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMax_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optA_Click(Value As Integer)
    Dim x
    If optG Then
        For x = 0 To 24
            medEntitle(x).Format = "00"
        Next x
    Else
        For x = 0 To 24
            medEntitle(x).Format = "#,##0.00"
        Next x
    End If
End Sub

Private Sub optG_Click(Value As Integer)
    Dim x
    If optG Then
        For x = 0 To 24
            medEntitle(x).Format = "00"
        Next x
    Else
        For x = 0 To 24
            medEntitle(x).Format = "#,##0.00"
        Next x
    End If
End Sub

'Private Sub modMaximums(TF%)
'Dim X%
'
'For X% = 0 To 24
'    If Not TF Then
'        If IsNumeric(medMax(X%)) Then medMax(X%) = 0
'    End If
'    medMax(X%).Enabled = TF And medMax(X%).Enabled
'Next X%
'
'End Sub


'Private Function modUpdateSelection()
'Dim EmpNo As Long
'Dim dblEntitle#, dblPrevEntitle#, dblTKEEntitle#, strDivision$
'Dim strJob$, dblServiceYears#
'Dim spt As Variant, varStartDate As Variant, lngRecs&
'Dim dblDHours#, intWhereFit&, X%, Y%, z%, dblNewEntitle#
'Dim dblFTEHours#
'Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
'Dim Msg$, Title$, DgDef As Variant
'Dim Response%, pct%
'Dim prec%, xAsOf
''Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
''Dim if_Entitle As Boolean, if_Vacation As Boolean
'Dim xComments
'On Error GoTo modUpdateSelection_Err
'modUpdateSelection = False
'
'If Len(dlpAsOf.Text) = 0 Then
'    MsgBox "Effective Date is required field"
'    dlpAsOf.SetFocus
'    Exit Function
'End If
'If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
'Screen.MousePointer = DEFAULT
'If snapEntitle.BOF And snapEntitle.EOF Then
'    If fglbRunTimes = 1 Then
'        MsgBox "Employees for this selection do not exist!"
'        Exit Function
'    End If
'Else
'    lngRecs& = snapEntitle.RecordCount
'    If fglbRunTimes = 1 Then
'        Msg$ = lngRecs& & " Records to process" & Chr(10) & "Would You Like To Proceed?"
'        Title$ = "Update Entitlements"
'        DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
'        Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
'        If Response% = IDNO Then    ' Evaluate response
'            Exit Function
'        End If
'        Screen.MousePointer = HOURGLASS
'    End If
'End If
'MDIMain.panHelp(0).FloodType = 1
'MDIMain.panHelp(0).FloodPercent = 5
'
'For X% = 0 To 24
'    If Not IsNumeric(medLTServ(X%)) Then
'        medLTServ(X%) = 0
'    End If
'    If Not IsNumeric(medGTServ(X%)) Then
'      medGTServ(X%) = 0
'    Else
'      If Val(medGTServ(X%)) = Int(medGTServ(X%)) Then medGTServ(X%) = medGTServ(X%) + 0.99
'    End If
'    If medLTServ(X%) > 0 And medGTServ(X%) = 0 Then medGTServ(X%) = 9999999
'Next
'
'While Not snapEntitle.EOF
'    prec% = prec% + 1
'    pct% = Int(100 * (prec% / lngRecs&))
'    MDIMain.panHelp(0).FloodPercent = pct%
'
'    'If snapEntitle("ED_EMPNBR") = 3190 Then
'    '    EmpNo& = snapEntitle("ED_EMPNBR")
'    'End If
'
'    EmpNo& = snapEntitle("ED_EMPNBR")
'
'    If IsNull(snapEntitle("ED_SICK")) Then
'        dblEntitle# = 0
'    Else
'        dblEntitle# = snapEntitle("ED_SICK")
'    End If
'
'    If IsNull(snapEntitle("ED_PSICK")) Then
'        dblPrevEntitle# = 0
'    Else
'        dblPrevEntitle# = snapEntitle("ED_PSICK")
'    End If
'
'    If IsNull(snapEntitle("ED_SICKT")) Then
'        dblTKEEntitle# = 0
'    Else
'        dblTKEEntitle# = snapEntitle("ED_SICKT")
'    End If
'
'    spt = snapEntitle("ED_PT")
'
'    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec
'
'    'Ticket #14260 DNSSAB
'    'Check last month attendance records, if there is any record with Incentive checked,
'    'and then skip this employee, also update the Accrual table
'    If glbCompSerial = "S/N - 2388W" Then
'        If IncentiveChecked(EmpNo&, dlpAsOf.Text) Then
'            Call Append_Accrual(EmpNo&, "SICK", dlpAsOf.Text, 0, "N", "No Sick Ent Attendance Found.")
'            GoTo lblNextRec
'        End If
'    End If
'
'    varStartDate = snapEntitle(fglbWDate$)
'
'    Dim rsJOB As New ADODB.Recordset
'    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
'    dblDHours# = 0
'    dblFTEHours# = 0
'    If Not rsJOB.EOF Then
'        If IsNumeric(rsJOB("JH_DHRS")) Then dblDHours# = rsJOB("JH_DHRS")
'        If IsNumeric(rsJOB("JH_FTENUM")) Then dblFTEHours# = rsJOB("JH_FTENUM")
'    End If
'    rsJOB.Close
'    If glbLinamar Then dblDHours# = 8
'
'    xAsOf = dlpAsOf.Text
''    dblServiceYears# = (DateDiff("d", varStartDate, CVDate(xAsOf)) / 365) * 12
'    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
'    intWhereFit& = -1
'
'    For X% = 0 To 24
'        If medGTServ(X%) > 0 Then
'            If dblServiceYears# >= CDbl(medLTServ(X%)) And dblServiceYears# <= CDbl(medGTServ(X%)) Then
'                intWhereFit& = X%
'                Exit For
'            End If
'        End If
'    Next X%
'
'    'Hemu - Added dblServiceYears# < 0 because it gives out entitlement way high which is wrong
'    If intWhereFit& = -1 Or dblServiceYears# < 0 Then GoTo lblNextRec ' skip record if not in any of the ranges
'
'
'    dblNewEntitle# = medEntitle(intWhereFit&)
'    dblNewMax# = 0
'    If optD(intWhereFit&) = True Then           ' Entitlements entered in days
'        If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblDHours#
'        dblNewEntitle# = dblNewEntitle# * dblDHours#
'        dblEntitleUpd = dblNewEntitle
'    End If
'    If optF(intWhereFit&) = True Then
'        If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHours# * dblDHours#
'        dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
'        dblEntitleUpd = dblNewEntitle
'    End If
'    If optH(intWhereFit&) = True Then
'        If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&)
'    End If
'    If fglbCompMonthly Then
'        dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
'    Else
'        dblEntitleUpd = dblNewEntitle ' rollover is in other utility (to accumulate)
'    End If
'
'    If dblNewMax <> 0 Then          'only do if not zero
'        If glbCompSerial = "S/N - 2173W" Or glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2389W" Then
'            'for town of Ajax or City of Timmins or St. Leonard's Community Services(Ticket #15071)
'            If (dblEntitle# + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
'                dblEntitleUpd = dblEntitle#
'            ElseIf (dblEntitleUpd + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
'                dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
'            End If
'        Else
'            If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
'                dblEntitleUpd = dblNewMax - dblPrevEntitle#
'
'                If glbCompSerial = "S/N - 2228W" Then   'Ticket #13359 - Simcoe Muskoka District Health Unit
'                    If dblEntitleUpd < 0 Then
'                        dblEntitleUpd = 0
'                    End If
'                End If
'            End If
'        End If
'    End If
'
'    If glbCBrant Then
'        If snapEntitle("ED_HIRECODE") = "Y" And dblTKEEntitle# > 0 Then
'            dblEntitleUpd = dblEntitleUpd - dblTKEEntitle#
'        End If
'    End If
'    DtTm = Now
'
'    xComments = "Current Sick. Ent. Chg from " & snapEntitle("ED_SICK") & " to " & dblEntitleUpd
'    'Hemu - Ticket #11925 - Changed the Accrual Date from Effective Date to Entitlement Start Date
'    'because otherwise it will not update Vadim until the date arrives in case it's not same as the
'    'Entitlement Start Date.
'    'Call Append_Accrual(EmpNo&, "SICK", dlpAsOf, dblEntitleUpd - Val(snapEntitle("ED_SICK") & ""), "U", xComments)
'    If fglbCompMonthly Then
'        Call Append_Accrual(EmpNo&, "SICK", dlpAsOf, dblEntitleUpd - Val(snapEntitle("ED_SICK") & ""), "U", xComments)
'    Else
'        Call Append_Accrual(EmpNo&, "SICK", dlpDateRange(0), dblEntitleUpd - Val(snapEntitle("ED_SICK") & ""), "U", xComments)
'    End If
'
'    snapEntitle("ED_SICK") = dblEntitleUpd
'    snapEntitle("ED_ANNSICK") = dblEntitleUpd
'    snapEntitle.Update
'lblNextRec:
'    DoEvents
'    Dim xKey
'    xKey = snapEntitle("ED_EMPNBR")
'    'xKey = xKey & "|" & Format(snapEntitle("ED_EFDATES"), "dd-mmm-yyyy")
'    'xKey = xKey & "|" & Format(snapEntitle("ED_ETDATES"), "dd-mmm-yyyy")
'    xKey = xKey & "|" & Format(dlpDateRange(0), "dd-mmm-yyyy")
'    xKey = xKey & "|" & Format(dlpDateRange(1), "dd-mmm-yyyy")
'    xKey = xKey & "|SICK"
'    If dblServiceYears# < 0 Then
'        dblEntitleUpd = 0
'    End If
'    xKey = xKey & "|" & dblEntitleUpd
'    xKey = xKey & "|" & Format(dlpAsOf.Text, "dd-mmm-yyyy") 'Transaction Date
'    Call Entitlements_Master_Integration(xKey, EmpNo&) 'George added for Advance Tracker
'    DoEvents
'    snapEntitle.MoveNext
'
'Wend
'modUpdateSelection = True
'MDIMain.panHelp(0).FloodType = 0
'
'snapEntitle.Close
'Set snapEntitle = Nothing
'Screen.MousePointer = DEFAULT
'
'Exit Function
'
'modUpdateSelection_Err:
''These errors are:
''13=type mismatch
''94=invalid use of null
''3018=couln't find field 'item'
'If Err = 13 Or Err = 94 Or Err = 3018 Then
'   ' MsgBox "Err:" & Str(Err) & Chr(10) & Error$ & Chr(10) & " modUpdateSelection" & Chr(10) & "FORM:FUENTITL.FRM"
'    'commented out by RAUBREY 5/20/97
'    Err = 0
'    Resume Next
'End If
'
'Screen.MousePointer = DEFAULT
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateEntitle", "HR_EMP", "edit/Add")
'Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
'    'Rollback
'    Resume Next
'Else
'    Unload Me
'End If
'End Function

'Private Function modAnnSelection()
'Dim EmpNo As Long
'Dim dblEntitle#, dblPrevEntitle#, dblTKEEntitle#, strDivision$
'Dim strJob$, dblServiceYears#
'Dim spt As Variant, varStartDate As Variant, lngRecs&
'Dim dblDHours#, intWhereFit&, X%, Y%, z%, dblNewEntitle#
'Dim dblFTEHours#
'Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
'Dim Msg$, Title$, DgDef As Variant
'Dim Response%, pct%
'Dim prec%, xAsOf
''Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
''Dim if_Entitle As Boolean, if_Vacation As Boolean
'Dim xComments
'On Error GoTo modUpdateSelection_Err
'modAnnSelection = False
'
'Screen.MousePointer = HOURGLASS
'
'MDIMain.panHelp(0).FloodType = 1
'MDIMain.panHelp(0).FloodPercent = 5
'
'For X% = 0 To 24
'    If Not IsNumeric(medLTServ(X%)) Then
'        medLTServ(X%) = 0
'    End If
'    If Not IsNumeric(medGTServ(X%)) Then
'      medGTServ(X%) = 0
'    Else
'      If Val(medGTServ(X%)) = Int(medGTServ(X%)) And Val(medGTServ(X%)) > 0 Then medGTServ(X%) = medGTServ(X%) + 0.99
'    End If
'    If medLTServ(X%) > 0 And medGTServ(X%) = 0 Then medGTServ(X%) = 9999999
'Next
'
'
'    EmpNo& = snapEntitle("ED_EMPNBR")
'
'    If IsNull(snapEntitle("ED_ANNSICK")) Then
'        dblEntitle# = 0
'    Else
'        dblEntitle# = snapEntitle("ED_ANNSICK")
'    End If
'
'
'    If IsNull(snapEntitle("ED_PSICK")) Then
'        dblPrevEntitle# = 0
'    Else
'        dblPrevEntitle# = snapEntitle("ED_PSICK")
'    End If
'
'    If IsNull(snapEntitle("ED_SICKT")) Then
'        dblTKEEntitle# = 0
'    Else
'        dblTKEEntitle# = snapEntitle("ED_SICKT")
'    End If
'
'    spt = snapEntitle("ED_PT")
'
'    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec
'
'    varStartDate = snapEntitle(fglbWDate$)
'
'    Dim rsJOB As New ADODB.Recordset
'    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
'    dblDHours# = 0
'    dblFTEHours# = 0
'    If Not rsJOB.EOF Then
'        If IsNumeric(rsJOB("JH_DHRS")) Then dblDHours# = rsJOB("JH_DHRS")
'        If IsNumeric(rsJOB("JH_FTENUM")) Then dblFTEHours# = rsJOB("JH_FTENUM")
'    End If
'    rsJOB.Close
'    If glbLinamar Then dblDHours# = 8
'
'    xAsOf = fglbAsOf
''    dblServiceYears# = (DateDiff("d", varStartDate, CVDate(xAsOf)) / 365) * 12
'    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
'    intWhereFit& = -1
'
'    For X% = 0 To 24
'        If medGTServ(X%) > 0 Then
'            If dblServiceYears# >= CDbl(medLTServ(X%)) And dblServiceYears# <= CDbl(medGTServ(X%)) Then
'                intWhereFit& = X%
'                Exit For
'            End If
'        End If
'    Next X%
'
'    If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
'
'
'    dblNewEntitle# = medEntitle(intWhereFit&)
'    dblNewMax# = 0
'    If optD(intWhereFit&) = True Then           ' Entitlements entered in days
'        If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblDHours#
'        dblNewEntitle# = dblNewEntitle# * dblDHours#
'        dblEntitleUpd = dblNewEntitle
'    End If
'    If optF(intWhereFit&) = True Then
'        If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHours# * dblDHours#
'        dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
'        dblEntitleUpd = dblNewEntitle
'    End If
'    If optH(intWhereFit&) = True Then
'        If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&)
'    End If
'
'    dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
'
'
'    If dblNewMax <> 0 Then          'only do if not zero
'        If glbCompSerial = "S/N - 2173W" Or glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2389W" Then
'            'for town of Ajax or City of Timmins or St. Leonard's Community Services
'            If (dblEntitle# + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
'                dblEntitleUpd = dblEntitle#
'            ElseIf (dblEntitleUpd + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
'                dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
'            End If
'        Else
'            If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
'                dblEntitleUpd = dblNewMax - dblPrevEntitle#
'            End If
'        End If
'    End If
'
'    If glbCBrant Then
'        If snapEntitle("ED_HIRECODE") = "Y" And dblTKEEntitle# > 0 Then
'            dblEntitleUpd = dblEntitleUpd - dblTKEEntitle#
'        End If
'    End If
'    DtTm = Now
'
'    xComments = "Current Sick. Ent. Chg from " & snapEntitle("ED_SICK") & " to " & dblEntitleUpd
'
'   snapEntitle("ED_ANNSICK") = dblEntitleUpd
'   snapEntitle.Update
'lblNextRec:
'    DoEvents
'
'
'
'modAnnSelection = True
'Screen.MousePointer = DEFAULT
'
'Exit Function
'
'modUpdateSelection_Err:
''These errors are:
''13=type mismatch
''94=invalid use of null
''3018=couln't find field 'item'
'If Err = 13 Or Err = 94 Or Err = 3018 Then
'    Err = 0
'    Resume Next
'End If
'Screen.MousePointer = DEFAULT
'
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateEntitle", "HR_EMP", "edit/Add")
'
'If gintRollBack% = False Then
'    'Rollback
'    Resume Next
'Else
'    Unload Me
'End If
'End Function


'Private Sub optD_Click(Index As Integer, Value As Integer)
'    Call ST_OPT_VALUE
'End Sub

'Private Sub optD_GotFocus(Index As Integer)
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub optF_Click(Index As Integer, Value As Integer)
'    Call ST_OPT_VALUE
'End Sub

'Private Sub optF_GotFocus(Index As Integer)
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub optH_Click(Index As Integer, Value As Integer)
'    Call ST_OPT_VALUE
'End Sub

'Private Sub optH_GotFocus(Index As Integer)
'Call SetPanHelp(ActiveControl)
'End Sub

Private Sub scrControl_Change()
    VacFram.Top = 4440 - scrControl.Value
End Sub

Sub ST_UPD_MODE(TF As Boolean)
    Dim x, FT
    FT = Not TF
    For x = 0 To 24
        medLTServ(x).Enabled = TF
        medGTServ(x).Enabled = TF
        medEntitle(x).Enabled = TF
    '    If X = 0 Then
        optA.Enabled = TF
        optG.Enabled = TF
    '    optF(X).Enabled = TF
    '    Else
    '    optD(X).Enabled = False
    '    optH(X).Enabled = False
    '    optF(X).Enabled = False
    '    End If
    '    medMax(X).Enabled = TF
    Next
    
    clpDiv.Enabled = TF
    clpDept.Enabled = TF
    clpCode(0).Enabled = TF
'    If Not TF Or glbLinamar Then
'        lblAsOf.FontBold = True
'    Else
'        lblAsOf.FontBold = False
'    End If
'    If glbCompEntSick$ = "M" Or glbCompEntSick$ = "N" Or glbCompEntSick$ = "A" Then
'        dlpAsOf.Enabled = True 'FT
'    Else
'        dlpAsOf.Enabled = True 'Ticket #3419
'    End If
    'If sick Entitlement Outstanding based on "1" then ok, otherwise disenable
'    If glbEntOutStandingS$ = "1" Then
'        CmdRecalc.Enabled = True
'    Else
'        CmdRecalc.Enabled = False
'    End If
    If Not glbWHSCC Then
        clpCode(1).Enabled = TF
    Else
        clpCode(1).Enabled = False
    End If
    clpCode(2).Enabled = TF
    clpCode(3).Enabled = TF
    clpCode(4).Enabled = TF
    clpPT.Enabled = TF
    'Ticket #15276 - Commented out
'    dlpDateRange(0).Enabled = TF
'    dlpDateRange(1).Enabled = TF
    
    'cmdClose.Enabled = FT
    'cmdModify.Enabled = FT
    'cmdDelete.Enabled = FT
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    '    cmdModify.Enabled = False
    '    cmdDelete.Enabled = False
    End If
    'cmdOK.Enabled = TF
    'cmdCancel.Enabled = TF
    'cmdNew.Enabled = FT
    'cmdPrint.Enabled = FT
    ''cmdPrintAll.Enabled = FT
    'cmdUpdate.Enabled = FT
    'vbxTrueGrid.Enabled = FT
    'Call modSetFGlobals("SICK")
End Sub

Sub Display_Value()
    Dim SQLQ, xOrder, nOrder, aa, SQLQW, glbiOneWhere
    Dim rsVE As New ADODB.Recordset
    Dim x, s
    For x = 0 To 24
        medLTServ(x) = ""
        medGTServ(x) = ""
        medEntitle(x) = ""
        optA = True
        optG = False
        If optG Then
            For s = 0 To 24
                medEntitle(s).Format = "00"
            Next s
        Else
            For s = 0 To 24
                medEntitle(s).Format = "#,##0.00"
            Next s
        End If
        
    '    optF(X) = False
    '    medMax(X) = ""
    Next
    clpDiv.Text = ""
    clpDept.Text = ""
    clpCode(0).Text = ""
'    If Not (glbCompEntSick$ = "M" Or glbCompEntSick$ = "N") Then
'        dlpAsOf.Text = ""
'    End If
    clpCode(1).Text = ""
    clpCode(2).Text = ""
    clpCode(3).Text = ""
    clpCode(4).Text = ""
    clpPT.Text = ""
    
    If Not Data1.Recordset.EOF Then
        SQLQ = "SELECT * FROM HR_SALARY_INCR "
        If IsNull(Data1.Recordset("SL_DIV")) Then
            SQLQ = SQLQ & " WHERE SL_DIV IS NULL"
        Else
            SQLQ = SQLQ & " WHERE SL_DIV = '" & Data1.Recordset("SL_DIV") & "'"
        End If
        If IsNull(Data1.Recordset("SL_DEPT")) Then
            SQLQ = SQLQ & " AND SL_DEPT IS NULL"
        Else
            SQLQ = SQLQ & " AND SL_DEPT = '" & Data1.Recordset("SL_DEPT") & "'"
        End If
        If IsNull(Data1.Recordset("SL_ORG")) Then
            SQLQ = SQLQ & " AND SL_ORG IS NULL"
        Else
            SQLQ = SQLQ & " AND SL_ORG = '" & Data1.Recordset("SL_ORG") & "'"
        End If
        If IsNull(Data1.Recordset("SL_LOC")) Then
            SQLQ = SQLQ & " AND SL_LOC IS NULL"
        Else
            SQLQ = SQLQ & " AND SL_LOC = '" & Data1.Recordset("SL_LOC") & "'"
        End If
        If IsNull(Data1.Recordset("SL_SECTION")) Then
            SQLQ = SQLQ & " AND SL_SECTION IS NULL"
        Else
            SQLQ = SQLQ & " AND SL_SECTION = '" & Data1.Recordset("SL_SECTION") & "'"
        End If
'        If Not IsNull(Data1.Recordset("VE_EDATE")) Then
'            SQLQ = SQLQ & " AND SL_EDATE = " & Date_SQL(Data1.Recordset("SL_EDATE"))
'        End If
        If IsNull(Data1.Recordset("SL_EMP")) Then
            SQLQ = SQLQ & " AND SL_EMP IS NULL"
        Else
            SQLQ = SQLQ & " AND SL_EMP = '" & Data1.Recordset("SL_EMP") & "'"
        End If
        If IsNull(Data1.Recordset("SL_PT")) Then
            SQLQ = SQLQ & " AND SL_PT IS NULL"
        Else
            SQLQ = SQLQ & " AND SL_PT = '" & Data1.Recordset("SL_PT") & "' "
        End If
        If IsNull(Data1.Recordset("SL_GRPCD")) Then
            SQLQ = SQLQ & " AND SL_GRPCD IS NULL"
        Else
            SQLQ = SQLQ & " AND SL_GRPCD = '" & Data1.Recordset("SL_GRPCD") & "'"
        End If
        If Not IsNull(Data1.Recordset("SL_FRDATE")) Then
            SQLQ = SQLQ & " AND SL_FRDATE = " & Date_SQL(Data1.Recordset("SL_FRDATE"))
        End If
        If Not IsNull(Data1.Recordset("SL_TODATE")) Then
            SQLQ = SQLQ & " AND SL_TODATE = " & Date_SQL(Data1.Recordset("SL_TODATE"))
        End If
        
        SQLQ = SQLQ & " Order By SL_DIV,SL_DEPT,SL_ORG,SL_EMP,SL_PT,SL_LOC,SL_SECTION,SL_ORDER "
        rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset
        If Not IsNull(Data1.Recordset("SL_DIV")) Then clpDiv.Text = Data1.Recordset("SL_DIV")
        If Not IsNull(Data1.Recordset("SL_DEPT")) Then clpDept.Text = Data1.Recordset("SL_DEPT")
        If Not IsNull(Data1.Recordset("SL_ORG")) Then clpCode(0).Text = Data1.Recordset("SL_ORG")
        'If Not IsNull(Data1.Recordset("SL_EDATE")) Then dlpAsOf.Text = Data1.Recordset("SL_EDATE")
        If Not IsNull(Data1.Recordset("SL_EMP")) Then clpCode(1).Text = Data1.Recordset("SL_EMP")
        If Not IsNull(Data1.Recordset("SL_PT")) Then clpPT.Text = Data1.Recordset("SL_PT")
        If Not IsNull(Data1.Recordset("SL_GRPCD")) Then clpCode(2).Text = Data1.Recordset("SL_GRPCD")
        If Not IsNull(Data1.Recordset("SL_LOC")) Then clpCode(4).Text = Data1.Recordset("SL_LOC")
        If Not IsNull(Data1.Recordset("SL_SECTION")) Then clpCode(3).Text = Data1.Recordset("SL_SECTION")
        'Ticket #15276 - Commented out
'        If Not IsNull(Data1.Recordset("SL_FRDATE")) Then dlpDateRange(0).Text = Data1.Recordset("SL_FRDATE")
'        If Not IsNull(Data1.Recordset("SL_TODATE")) Then dlpDateRange(1).Text = Data1.Recordset("SL_TODATE")
        'If Not IsNull(Data1.Recordset("SL_MANUAL")) Then chkManual.Value = Data1.Recordset("SL_MANUAL")
        If rsVE("SL_TYPE") = "A" Then optA = True
        If rsVE("SL_TYPE") = "G" Then optG = True
                
        Do While Not rsVE.EOF
            xOrder = rsVE("SL_ORDER")
            nOrder = Format(Val(xOrder), "##0") - 1
            If Not (nOrder < 0 Or nOrder > 24) Then
                If Not IsNull(rsVE("SL_FROM_HRS")) Then medLTServ(nOrder) = rsVE("SL_FROM_HRS")
                If Not IsNull(rsVE("SL_TO_HRS")) Then medGTServ(nOrder) = rsVE("SL_TO_HRS")
                If Not IsNull(rsVE("SL_SALARY")) Then medEntitle(nOrder) = rsVE("SL_SALARY")
'                If rsVE("SL_TYPE") = "A" Then optA = True
'                If rsVE("SL_TYPE") = "G" Then optG = True
    '            If rsVE("VE_TYPE") = "F" Then optF(nOrder) = True
    '            If Not IsNull(rsVE("VE_MAX")) Then medMax(nOrder) = rsVE("VE_MAX")
            End If
            rsVE.MoveNext
        Loop
        rsVE.Close
    End If
    Call SET_UP_MODE
    Call cmdModify_Click
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
    Dim SQLQ As String
       
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    SQLQ = "SELECT DISTINCT SL_DIV,SL_DEPT,SL_ORG,SL_LOC,SL_SECTION,SL_EMP,SL_PT,SL_GRPCD,SL_FRDATE,SL_TODATE FROM HR_SALARY_INCR"
    If glbDIVCount = 1 And glbLinamar Then
        SQLQ = SQLQ & " WHERE SL_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Display_Value
End Sub

'Private Sub modSetFGlobals(strTyp$)
'If glbCompEntSick$ = "M" Or UCase(glbCompEntSick$) = "N" Then
'    fglbCompMonthly% = True
'    Call modMaximums(True)
'Else
'    fglbCompMonthly% = False
'    Call modMaximums(False)
'End If
'
'
'End Sub

'Sub ST_OPT_VALUE()
'Dim X, XoptD, XoptH, XoptF
'    XoptD = optD(0).Value
'    XoptH = optH(0).Value
'    XoptF = optF(0).Value
'    For X = 1 To 24
'        optD(X).Value = XoptD
'        optH(X).Value = XoptH
'        optF(X).Value = XoptF
'    Next
'End Sub


Private Sub getWSQLQ(xType)
Dim xDiv, xDept, xORG, xAsOf, xEMP, xEmpMode, xGRPCE
Dim xLoc, xSection
Dim xFromDate
Dim xToDate

fglbESQLQ = glbSeleDeptUn
If Len(clpDept.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND  ED_DEPTNO = '" & clpDept.Text & "' "
If Len(clpDiv.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & clpDiv.Text & "' "
If Len(clpCode(0).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & clpCode(0).Text & "' "
If Len(clpCode(1).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP = '" & clpCode(1).Text & "' "
If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & clpCode(3).Text & "' "
If Len(clpCode(4).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & clpCode(4).Text & "' "


If clpPT.Text <> "" Then fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & clpPT.Text & "' "

If xType = "" Then Exit Sub

If xType = "O" Then
    xDiv = ODIV
    xDept = ODept
    xORG = oOrg
    xAsOf = oAsOf
    xEMP = oEMP
    xEmpMode = oEmpMode
    xGRPCE = oGRPCE
    xLoc = OLoc
    xSection = OSection
    xFromDate = OFromDate
    xToDate = OToDate
Else
    xDiv = clpDiv.Text
    xDept = clpDept.Text
    xORG = clpCode(0).Text
    xAsOf = dlpAsOf.Text
    xEMP = clpCode(1).Text
    xEmpMode = clpPT.Text
    xGRPCE = clpCode(2).Text
    xLoc = clpCode(4).Text
    xSection = clpCode(3).Text
    'Ticket #15276 - Commented out
'    xFromDate = dlpDateRange(0)
'    xToDate = dlpDateRange(1)
End If

If Len(xDiv) = 0 Then
    fglbVSQLQ = " (SL_DIV IS NULL OR SL_DIV='')"
Else
    fglbVSQLQ = "SL_DIV = '" & xDiv & "'"
End If
If Len(xDept) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (SL_DEPT IS NULL OR SL_DEPT='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND SL_DEPT = '" & xDept & "'"
End If
If Len(xORG) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (SL_ORG IS NULL OR SL_ORG='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND SL_ORG = '" & xORG & "'"
End If
'If UCase(glbCompEntSick$) = "A" Then
'    If Len(xAsOf) > 0 Then fglbVSQLQ = fglbVSQLQ & " AND  VE_EDATE = " & Date_SQL(xAsOf)
'End If
If Len(xEMP) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (SL_EMP IS NULL OR SL_EMP='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND SL_EMP = '" & xEMP & "'"
End If
If Len(xEmpMode) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (SL_PT IS NULL OR SL_PT='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND SL_PT = '" & xEmpMode & "' "
End If
If Len(xGRPCE) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (SL_GRPCD IS NULL OR SL_GRPCD='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND SL_GRPCD = '" & xGRPCE & "'"
End If

If Len(xLoc) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (SL_LOC IS NULL OR SL_LOC='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND SL_LOC = '" & xLoc & "'"
End If
If Len(xSection) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (SL_SECTION IS NULL OR SL_SECTION='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND SL_SECTION = '" & xSection & "'"
End If

If Not IsDate(xFromDate) Then
    fglbVSQLQ = fglbVSQLQ & " AND SL_FRDATE IS NULL  "
Else
    fglbVSQLQ = fglbVSQLQ & " AND SL_FRDATE = " & Date_SQL(xFromDate)
End If
If Not IsDate(xToDate) Then
    fglbVSQLQ = fglbVSQLQ & " AND SL_TODATE IS NULL  "
Else
    fglbVSQLQ = fglbVSQLQ & " AND SL_TODATE = " & Date_SQL(xToDate)
End If

End Sub

Public Sub SET_UP_MODE()
    Dim TF As Boolean
    Dim UpdateState As UpdateStateEnum
    If fglbNew Then
        UpdateState = NewRecord
        TF = True
        cmdPrintAll.Enabled = False
        cmdUpdate.Enabled = False
        CmdRecalc.Enabled = False
    ElseIf Me.Data1.Recordset.EOF Then
        UpdateState = NoRecord
        TF = False
        cmdPrintAll.Enabled = True
        cmdUpdate.Enabled = False
        CmdRecalc.Enabled = False
    Else
        UpdateState = OPENING
        TF = True
        cmdPrintAll.Enabled = True
        cmdUpdate.Enabled = True
        CmdRecalc.Enabled = True
    End If
    
    Call ST_UPD_MODE(TF)
    
    Call set_Buttons(UpdateState)
    
    If Not UpdateRight Then TF = False
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
RelateMode = nothingrelate
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Entitlements
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


'Private Function DoWork() As Boolean
'    'Added by Bryan 25/Oct/05 Ticket#9560
'Dim lastday
'Dim flglastdate As Boolean
'Dim lngRecs As Long, pct As Long, prec As Long
'        Screen.MousePointer = DEFAULT
'        DoWork = False
'        If UCase(glbCompEntSick$) = "N" Then
'            For fglbRunTimes = 1 To 12
'                If Not modUpdateSelection() Then Exit Function
'                dlpAsOf = DateAdd("m", 1, CVDate(dlpAsOf.Text))
'                DoEvents
'                If glbCompSerial = "S/N - 2386W" Then 'The Walter Fedy Partnership Ticket #14255
'                    Call Pause(3)
'                    MsgBox "Click OK Button to next month: " & dlpAsOf
'                End If
'            Next
'            dlpAsOf = DateAdd("m", -12, CVDate(dlpAsOf.Text))
'        Else
'            If Not glbWHSCC Then
'                If Not modUpdateSelection() Then Exit Function
'                If fglbCompMonthly Then
'                    gdbAdoIhr001.Execute "UPDATE HREMP SET ED_ANNSICK=0"
'
'                    If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
'                    If snapEntitle.EOF = False And snapEntitle.BOF = False Then
'                        While Not snapEntitle.EOF
'                            lngRecs = snapEntitle.RecordCount
'                            prec = prec + 1
'                            pct = Int(100 * (prec / lngRecs))
'                            If pct > 100 Then pct = 100
'                            MDIMain.panHelp(0).FloodPercent = pct
'                            Dim doDate As Date
'                            doDate = dlpAsOf
'                            If Not IsNull(snapEntitle("ED_EFDATES")) Then 'Ticket #12923
'                                fglbAsOf = snapEntitle("ED_EFDATES")
'                                For fglbRunTimes = 1 To 12
'                                    If Not modAnnSelection() Then Exit Function
'                                    fglbAsOf = DateAdd("m", 1, CVDate(fglbAsOf))
'
'                                    DoEvents
'                                Next
'                            End If
'                            snapEntitle.MoveNext
'                        Wend
'                        MDIMain.panHelp(0).FloodType = 0
'                    End If
'                End If
'            Else
'                If Not modUpdateSelectionWHSCC() Then Exit Function
'            End If
'        End If
'        If Not glbSQL And Not glbOracle Then Call Pause(0.5)
'        Screen.MousePointer = HOURGLASS
'        Call EntReCalc(fglbESQLQ)
'
'        If Not glbSQL And Not glbOracle Then Call Pause(0.5)
'        DoWork = True
'End Function

'Private Function IncentiveChecked(xEmpNo, xEffDate)
'Dim rsAttInc As New ADODB.Recordset
'Dim SQLQ As String
'Dim xDateFrom, xDateEnd
'Dim xMonth As String
'    xMonth = MonthName(Month(xEffDate))
'    xDateFrom = CVDate(xMonth & " 1," & str(Year(xEffDate)))
'    xDateEnd = DateAdd("D", -1, xDateFrom)
'    xDateFrom = DateAdd("M", -1, xDateFrom)
'
'    SQLQ = "SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & xEmpNo & " "
'    SQLQ = SQLQ & "AND AD_DOA >= " & Date_SQL(xDateFrom) & " "
'    SQLQ = SQLQ & "AND AD_DOA <= " & Date_SQL(xDateEnd) & " "
'    SQLQ = SQLQ & "AND NOT (AD_INDICATOR = 0) "
'    rsAttInc.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    If Not rsAttInc.EOF Then
'        IncentiveChecked = True
'    Else
'        IncentiveChecked = False
'    End If
'    rsAttInc.Close
'End Function
