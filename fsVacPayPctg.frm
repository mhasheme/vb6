VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSVacPayPrct 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Vacation Pay Percentage Master"
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
      Height          =   4650
      Left            =   120
      TabIndex        =   112
      Top             =   120
      Width           =   11415
      Begin Threed.SSCheck chkManual 
         Height          =   255
         Left            =   5540
         TabIndex        =   11
         Top             =   3360
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
         Left            =   2100
         TabIndex        =   12
         Tag             =   "40-As of Date"
         Top             =   3705
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
         TabIndex        =   7
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
         TabIndex        =   3
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
         TabIndex        =   2
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   6420
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   8
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
         TabIndex        =   4
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
         TabIndex        =   9
         Tag             =   "40-From Date"
         Top             =   3345
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
         TabIndex        =   10
         Tag             =   "40-To Date"
         Top             =   3345
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1210
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
         Bindings        =   "fsVacPayPctg.frx":0000
         Height          =   1695
         Left            =   0
         OleObjectBlob   =   "fsVacPayPctg.frx":0014
         TabIndex        =   0
         Top             =   0
         Width           =   9135
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         Height          =   285
         Left            =   945
         TabIndex        =   1
         Tag             =   "00-Specific Division Desired"
         Top             =   2040
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
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
         TabIndex        =   118
         Top             =   4410
         Width           =   2250
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vacation Pay %"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3600
         TabIndex        =   117
         Top             =   4410
         Width           =   1335
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
         TabIndex        =   134
         Top             =   4110
         Visible         =   0   'False
         Width           =   7455
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
         TabIndex        =   133
         Top             =   3390
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
         Left            =   30
         TabIndex        =   120
         Top             =   3750
         Visible         =   0   'False
         Width           =   1245
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
         TabIndex        =   119
         Top             =   2640
         Width           =   1260
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
         TabIndex        =   116
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
         TabIndex        =   115
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
         TabIndex        =   114
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
         TabIndex        =   113
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
      TabIndex        =   110
      Top             =   4950
      Width           =   300
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   645
      Left            =   0
      TabIndex        =   13
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
         Left            =   3840
         TabIndex        =   15
         Top             =   120
         Width           =   1665
      End
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update Vac. Pay %"
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Tag             =   "Change all matching records to the above"
         Top             =   120
         Width           =   1905
      End
      Begin VB.CommandButton CmdRecalc 
         Appearance      =   0  'Flat
         Caption         =   "R&ecalculate"
         Height          =   375
         Left            =   6360
         TabIndex        =   111
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
         TabIndex        =   16
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
      TabIndex        =   17
      Top             =   4800
      Width           =   11000
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   18
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
         TabIndex        =   19
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
         TabIndex        =   21
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
         TabIndex        =   22
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   2
         Left            =   0
         TabIndex        =   24
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
         TabIndex        =   25
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   1
         Left            =   3750
         TabIndex        =   23
         Tag             =   "11-Entitlement Amount"
         Top             =   435
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   2
         Left            =   3750
         TabIndex        =   26
         Tag             =   "11-Entitlement Amount"
         Top             =   750
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   0
         Left            =   3750
         TabIndex        =   20
         Tag             =   "11-Vacation Percentage"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   3
         Left            =   0
         TabIndex        =   27
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
         TabIndex        =   28
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   4
         Left            =   0
         TabIndex        =   30
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   31
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   5
         Left            =   0
         TabIndex        =   33
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
         TabIndex        =   34
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   4
         Left            =   3750
         TabIndex        =   32
         Tag             =   "11-Entitlement Amount"
         Top             =   1410
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   5
         Left            =   3750
         TabIndex        =   35
         Tag             =   "11-Entitlement Amount"
         Top             =   1725
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   3
         Left            =   3750
         TabIndex        =   29
         Tag             =   "11-Entitlement Amount"
         Top             =   1080
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   6
         Left            =   0
         TabIndex        =   36
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
         TabIndex        =   37
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   7
         Left            =   0
         TabIndex        =   39
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
         TabIndex        =   40
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   8
         Left            =   0
         TabIndex        =   42
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
         TabIndex        =   43
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   7
         Left            =   3750
         TabIndex        =   41
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   8
         Left            =   3750
         TabIndex        =   44
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   6
         Left            =   3750
         TabIndex        =   38
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   9
         Left            =   0
         TabIndex        =   45
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
         TabIndex        =   46
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   9
         Left            =   3750
         TabIndex        =   47
         Tag             =   "11-Entitlement Amount"
         Top             =   2985
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   10
         Left            =   0
         TabIndex        =   48
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
         TabIndex        =   49
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   11
         Left            =   0
         TabIndex        =   51
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
         TabIndex        =   52
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   10
         Left            =   3750
         TabIndex        =   50
         Tag             =   "11-Entitlement Amount"
         Top             =   3330
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   11
         Left            =   3750
         TabIndex        =   53
         Tag             =   "11-Entitlement Amount"
         Top             =   3645
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   12
         Left            =   0
         TabIndex        =   54
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
         TabIndex        =   55
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   13
         Left            =   0
         TabIndex        =   57
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
         TabIndex        =   58
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   14
         Left            =   0
         TabIndex        =   60
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
         TabIndex        =   61
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   12
         Left            =   3750
         TabIndex        =   56
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   13
         Left            =   3750
         TabIndex        =   59
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   14
         Left            =   3750
         TabIndex        =   62
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   15
         Left            =   0
         TabIndex        =   63
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
         TabIndex        =   64
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   16
         Left            =   0
         TabIndex        =   66
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
         TabIndex        =   67
         Tag             =   "10-Service is less than this number"
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   15
         Left            =   3750
         TabIndex        =   65
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   16
         Left            =   3750
         TabIndex        =   68
         Tag             =   "11-Entitlement Amount"
         Top             =   5265
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   17
         Left            =   0
         TabIndex        =   69
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
         TabIndex        =   70
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
         TabIndex        =   71
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   18
         Left            =   0
         TabIndex        =   72
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
         TabIndex        =   73
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
         TabIndex        =   75
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
         TabIndex        =   76
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
         TabIndex        =   74
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   19
         Left            =   3750
         TabIndex        =   77
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   20
         Left            =   0
         TabIndex        =   78
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
         TabIndex        =   79
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
         TabIndex        =   81
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
         TabIndex        =   82
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
         TabIndex        =   84
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
         TabIndex        =   85
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
         TabIndex        =   80
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   21
         Left            =   3750
         TabIndex        =   83
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   22
         Left            =   3750
         TabIndex        =   86
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   23
         Left            =   0
         TabIndex        =   87
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
         TabIndex        =   88
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
         TabIndex        =   90
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
         TabIndex        =   91
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
         TabIndex        =   89
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   24
         Left            =   3750
         TabIndex        =   92
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "0.00%"
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
         TabIndex        =   109
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
         TabIndex        =   108
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
         TabIndex        =   107
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
         TabIndex        =   106
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
         TabIndex        =   105
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
         TabIndex        =   104
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
         TabIndex        =   103
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
         TabIndex        =   102
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
         TabIndex        =   101
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
         TabIndex        =   100
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
         TabIndex        =   99
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
         TabIndex        =   98
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
         TabIndex        =   97
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
         TabIndex        =   96
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
         TabIndex        =   95
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
         TabIndex        =   94
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
         TabIndex        =   93
         Top             =   4920
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmSVacPayPrct"
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
    MsgBox lStr("Invalid Department")
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
'Hemu - 05/13/2003 Begin
If clpPT.Caption = "Unassigned" Then
    MsgBox "Invalid " & lblPT.Caption
    clpPT.SetFocus
    Exit Function
End If


'Ticket #29617 -  Mississaugas of Scugog Island First Nation
If glbCompSerial = "S/N - 2485W" Then
    If Len(dlpDateRange(0).Text) > 0 Then
        If Not IsDate(dlpDateRange(0).Text) Then
            MsgBox "Invalid Entitlement Period From Date"
            dlpDateRange(0).SetFocus
            Exit Function
        End If
    Else
        MsgBox "Entitlement Period From Date is mandatory field"
        dlpDateRange(0).SetFocus
        Exit Function
    End If
    
    If Len(dlpDateRange(1).Text) > 0 Then
        If Not IsDate(dlpDateRange(1).Text) Then
            MsgBox "Invalid Entitlement Period To Date"
            dlpDateRange(1).SetFocus
            Exit Function
        End If
    Else
        MsgBox "Entitlement Period To Date is mandatory field"
        dlpDateRange(1).SetFocus
        Exit Function
    End If
    
    If IsDate(dlpDateRange(0).Text) And IsDate(dlpDateRange(1).Text) Then
    If CVDate(dlpDateRange(0).Text) > CVDate(dlpDateRange(1).Text) Then
        MsgBox "Entitlement Period From Date cannot be greater than Entitlement Period To Date"
        dlpDateRange(0).SetFocus
        Exit Function
    End If
    End If
Else
    'Sam 02/02/2006
    'Ticket #15276 - Commented
    If Len(dlpDateRange(0).Text) > 0 Then
        If Not IsDate(dlpDateRange(0).Text) Then
            MsgBox "Invalid Attendance Period From Date"
            dlpDateRange(0).SetFocus
            Exit Function
        End If
    Else
        MsgBox "Attendance Period From Date is mandatory field"
        dlpDateRange(0).SetFocus
        Exit Function
    End If
    
    If Len(dlpDateRange(1).Text) > 0 Then
        If Not IsDate(dlpDateRange(1).Text) Then
            MsgBox "Invalid Attendance Period To Date"
            dlpDateRange(1).SetFocus
            Exit Function
        End If
    Else
        MsgBox "Attendance Period To Date is mandatory field"
        dlpDateRange(1).SetFocus
        Exit Function
    End If
    
    If IsDate(dlpDateRange(0).Text) And IsDate(dlpDateRange(1).Text) Then
        If CVDate(dlpDateRange(0).Text) > CVDate(dlpDateRange(1).Text) Then
            MsgBox "Attendance Period From Date cannot be greater than Attedance Period To Date"
            dlpDateRange(0).SetFocus
            Exit Function
        End If
    End If
End If

'Ticket #29617 - Mississaugas of Scugog Island First Nation
If glbCompSerial = "S/N - 2485W" Then
    If Len(dlpAsOf.Text) > 0 Then
      If Not IsDate(dlpAsOf.Text) Then
        MsgBox "Invalid Effective Date"
        dlpAsOf.SetFocus
        Exit Function
      End If
    Else
        'If UCase(glbCompEntSick$) = "A" Then
        '    If glbLinamar Then
                MsgBox "Effective Date is required field"
                dlpAsOf.SetFocus
                Exit Function
        '    End If
        'End If
    End If
End If

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
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkMUEntitle", "HRVACPCTENT", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

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
    Msg = Msg & Chr(10) & "The Vacation Pay Percentage Rules?  "
    
    A% = MsgBox(Msg, 36, "Confirm Delete")
    If A% <> 6 Then Exit Sub
    
    Call getWSQLQ("C")
    SQLQ = "DELETE FROM HRVACPCTENT WHERE " & fglbVSQLQ
    
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
    If Not Data1.Recordset.EOF Then
        If Not IsNull(Data1.Recordset("VP_EDATE")) Then
            oAsOf = Data1.Recordset("VP_EDATE")
        End If
    End If

    'Sam 02/02/2006
    OFromDate = dlpDateRange(0).Text
    OToDate = dlpDateRange(1).Text
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
    '    optD(X) = True
    '    optH(X) = False
    '    optF(X) = False
    '    medMax(X) = ""
    Next
    
    'Sam 02/2/2006
    dlpDateRange(0).Text = ""
    dlpDateRange(1).Text = ""
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
        SQLQ = "DELETE FROM HRVACPCTENT WHERE " & fglbVSQLQ
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
    Else
        Call getWSQLQ("C")
        SQLQ = "SELECT * FROM HRVACPCTENT WHERE " & fglbVSQLQ
        rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsVT.EOF Then
            MsgBox "You can not add duplicate record"
            clpDiv.SetFocus
            Exit Sub
        End If
    End If
    gdbAdoIhr001.BeginTrans
    SQLQ = "SELECT * FROM HRVACPCTENT"
    rsVE.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    For x% = 0 To 24
        If Len(medLTServ(x%)) > 0 Then
            rsVE.AddNew
            rsVE("VP_ORDER") = x + 1
            rsVE("VP_ORG_TABL") = "EDOR"
            rsVE("VP_ORG") = clpCode(0).Text
            rsVE("VP_PT") = clpPT.Text
            rsVE("VP_DIV") = clpDiv.Text
            rsVE("VP_DEPT") = clpDept.Text
            rsVE("VP_EMP_TABL") = "EDEM"
            rsVE("VP_EMP") = clpCode(1).Text
            rsVE("VP_SECTION") = clpCode(3).Text
            rsVE("VP_LOC") = clpCode(4).Text
            
            'Ticket #29617 - Mississaugas of Scugog Island First Nation
            If glbCompSerial = "S/N - 2485W" Then
                rsVE("VP_EDATE") = dlpAsOf.Text
            End If
            
            If Len(dlpDateRange(0).Text) > 0 Then
                rsVE("VP_FRDATE") = dlpDateRange(0).Text
            End If
            If Len(dlpDateRange(1).Text) > 0 Then
                rsVE("VP_TODATE") = dlpDateRange(1).Text
            End If
            
            rsVE("VP_GRPCD_TABL") = "JBGC"
            rsVE("VP_GRPCD") = clpCode(2).Text
            rsVE("VP_BHOUR") = medLTServ(x%)
            rsVE("VP_EHOUR") = medGTServ(x%)
            If medEntitle(x%) = "" Then
                rsVE("VP_PCT") = Null
            Else
                rsVE("VP_PCT") = medEntitle(x%)
            End If
    '        If optD(X%) Then rsVE("VE_TYPE") = "D"
    '        If optH(X%) Then rsVE("VE_TYPE") = "H"
    '        If optF(X%) Then rsVE("VE_TYPE") = "F"
    '        rsVE("VE_MAX") = medMax(X%)
            rsVE("VP_MANUAL") = chkManual.Value
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
    Me.vbxCrystal.WindowTitle = "Vacation Pay Percentage Master Report"
    
    Call setRptLabel(Me, 0) '1)
    
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 5
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    'Ticket #29617 - Mississaugas of Scugog Island First Nation
    If glbCompSerial = "S/N - 2485W" Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgvacpaypctmst1.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgvacpaypctmst.rpt"
    End If
    
    SQLQ = "(1=1) "
    If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_DIV} = '" & clpDiv.Text & "'"
    If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_DEPT} = '" & clpDept.Text & "'"
    If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_ORG} = '" & clpCode(0).Text & "'"
    
    'Ticket #29617 - Mississaugas of Scugog Island First Nation
    If Len(dlpAsOf.Text) > 0 Then
        dtYYY% = Year(dlpAsOf.Text)
        dtMM% = month(dlpAsOf.Text)
        dtDD% = Day(dlpAsOf.Text)
        SQLQ = SQLQ & " AND {HRVACPCTENT.VP_EDATE} = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    End If
    If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_EMP} = '" & clpCode(1).Text & "'"
    If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_PT} = '" & clpPT.Text & "' "
    If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_GRPCD} = '" & clpCode(2).Text & "'"
    If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_SECTION} = '" & clpCode(3).Text & "'"
    If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_LOC} = '" & clpCode(4).Text & "'"
    
    If Len(dlpDateRange(0).Text) > 0 Then
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND {HRVACPCTENT.VP_FRDATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    End If
    If Len(dlpDateRange(1).Text) > 0 Then
        dtYYY% = Year(dlpDateRange(1).Text)
        dtMM% = month(dlpDateRange(1).Text)
        dtDD% = Day(dlpDateRange(1).Text)
        SQLQ = SQLQ & " AND {HRVACPCTENT.VP_TODATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    End If
    
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
    Me.vbxCrystal.WindowTitle = "Vacation Pay Percentage Master Report"
    
    Call setRptLabel(Me, 0) '1)
    
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 5
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    'Ticket #29617 - Mississaugas of Scugog Island First Nation
    If glbCompSerial = "S/N - 2485W" Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgvacpaypctmst1.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgvacpaypctmst.rpt"
    End If
    
    SQLQ = "(1=1) "
    If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_DIV} = '" & clpDiv.Text & "'"
    If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_DEPT} = '" & clpDept.Text & "'"
    If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_ORG} = '" & clpCode(0).Text & "'"
    
    'Ticket #29617 - Mississaugas of Scugog Island First Nation
    If Len(dlpAsOf.Text) > 0 Then
        dtYYY% = Year(dlpAsOf.Text)
        dtMM% = month(dlpAsOf.Text)
        dtDD% = Day(dlpAsOf.Text)
        SQLQ = SQLQ & " AND {HRVACPCTENT.VP_EDATE} = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    End If
    
    If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_EMP} = '" & clpCode(1).Text & "'"
    If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_PT} = '" & clpPT.Text & "' "
    If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_GRPCD} = '" & clpCode(2).Text & "'"
    If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_SECTION} = '" & clpCode(3).Text & "'"
    If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACPCTENT.VP_LOC} = '" & clpCode(4).Text & "'"
    
    If Len(dlpDateRange(0).Text) > 0 Then
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND {HRVACPCTENT.VP_FRDATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    End If
    If Len(dlpDateRange(1).Text) > 0 Then
        dtYYY% = Year(dlpDateRange(1).Text)
        dtMM% = month(dlpDateRange(1).Text)
        dtDD% = Day(dlpDateRange(1).Text)
        SQLQ = SQLQ & " AND {HRVACPCTENT.VP_TODATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    End If
            
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

Me.vbxCrystal.WindowTitle = "Vacation Pay Percentage Master Report"
Call setRptLabel(Me, 0) '1)
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To 5
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next
End If
'Ticket #29617 - Mississaugas of Scugog Island First Nation
If glbCompSerial = "S/N - 2485W" Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgvacpaypctmst1.rpt"
Else
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgvacpaypctmst.rpt"
End If
Me.vbxCrystal.Action = 1

cmdPrintAll.Enabled = True
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo Mod_Err
Dim sFlag As Boolean

If Not gSec_Upd_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

If Not chkMUEntitle() Then Exit Sub

'Added by Bryan 25/Oct/05 Ticket#9560
'made the code a separate sub because it's being used in two places
sFlag = DoWork

Data1.Refresh

Call Display_Value

If sFlag Then
    MsgBox "Update Completed Successfully", vbInformation + vbOKOnly, "Vacation Pay Percentage"
End If

Screen.MousePointer = DEFAULT

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdateAll", "Single", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
     RollBack
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub cmdUpdate_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Function CR_SnapEntitle()
Dim SQLQ As String
Dim SQLQ2 As String
Dim snapMultiEmp As New ADODB.Recordset

CR_SnapEntitle = False
On Error GoTo CR_SnapEntitle_Err


Call getWSQLQ("")

SQLQ = "SELECT ED_EMPNBR,ED_VACPC,ED_PVAC,ED_VAC,ED_PSICK,ED_SICK,ED_SICKT,ED_EFDATES,ED_ETDATES, HREMP.ED_ANNVAC, HREMP.ED_ANNSICK, "
SQLQ = SQLQ & " ED_DIV,ED_PT, ED_SECTION, ED_LOC, ED_EMP,"
SQLQ = SQLQ & " ED_HIRECODE," 'County of Brant Ticket #12525
SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME "
SQLQ = SQLQ & " FROM HREMP WHERE " & fglbESQLQ
If Len(clpCode(2).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_EMPNBR IN "
    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
    SQLQ = SQLQ & " WHERE JB_GRPCD = '" & clpCode(2).Text & "') "
End If

If snapEntitle.State <> 0 Then snapEntitle.Close
If glbOracle Then
    snapEntitle.CursorLocation = adUseServer
End If
snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic

CR_SnapEntitle = True

Exit Function

CR_SnapEntitle_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapEntitle", "VacationPay%/EMP", "Select")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub cmdUpdateAll_Click()
On Error GoTo Mod_Err

Dim c As Long
Dim failed As String

If Not gSec_Upd_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

failed = ""
c = 1
If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
    Data1.Recordset.MoveFirst
    Do
        Call Display_Value

        'made the DoWork a separate sub because it's being used in two places
        If chkManual.Value = False Then
            If chkMUEntitle() Then
                If DoWork = False Then
                    failed = failed & "Rule " & CStr(c) & ": "
                    If Not IsNull(Data1.Recordset("VP_DIV")) Then failed = failed & Data1.Recordset("VP_DIV") & ", "
                    If Not IsNull(Data1.Recordset("VP_DEPT")) Then failed = failed & Data1.Recordset("VP_DEPT") & ", "
                    If Not IsNull(Data1.Recordset("VP_ORG")) Then failed = failed & Data1.Recordset("VP_ORG") & ", "
                    'Ticket #29617 - Mississaugas of Scugog Island First Nation
                    If glbCompSerial = "S/N - 2485W" Then
                        If Not IsNull(Data1.Recordset("VP_EDATE")) Then failed = failed & Data1.Recordset("VP_EDATE") & ", "
                    End If
                    If Not IsNull(Data1.Recordset("VP_EMP")) Then failed = failed & Data1.Recordset("VP_EMP") & ", "
                    If Not IsNull(Data1.Recordset("VP_PT")) Then failed = failed & Data1.Recordset("VP_PT") & ", "
                    If Not IsNull(Data1.Recordset("VP_GRPCD")) Then failed = failed & Data1.Recordset("VP_GRPCD") & ", "
                    If Not IsNull(Data1.Recordset("VP_LOC")) Then failed = failed & Data1.Recordset("VP_LOC") & ", "
                    If Not IsNull(Data1.Recordset("VP_SECTION")) Then failed = failed & Data1.Recordset("VP_SECTION") & ", "
                    If Not IsNull(Data1.Recordset("VP_FRDATE")) Then failed = failed & Data1.Recordset("VP_FRDATE") & ", "
                    If Not IsNull(Data1.Recordset("VP_TODATE")) Then failed = failed & Data1.Recordset("VP_TODATE") & ", "
                    failed = Left(failed, Len(failed) - 2) & vbCrLf
                End If
            Else
                failed = failed & "Rule " & CStr(c) & ": "
                If Not IsNull(Data1.Recordset("VP_DIV")) Then failed = failed & Data1.Recordset("VP_DIV") & ", "
                If Not IsNull(Data1.Recordset("VP_DEPT")) Then failed = failed & Data1.Recordset("VP_DEPT") & ", "
                If Not IsNull(Data1.Recordset("VP_ORG")) Then failed = failed & Data1.Recordset("VP_ORG") & ", "
                'Ticket #29617 - Mississaugas of Scugog Island First Nation
                If glbCompSerial = "S/N - 2485W" Then
                    If Not IsNull(Data1.Recordset("VP_EDATE")) Then failed = failed & Data1.Recordset("VP_EDATE") & ", "
                End If
                If Not IsNull(Data1.Recordset("VP_EMP")) Then failed = failed & Data1.Recordset("VP_EMP") & ", "
                If Not IsNull(Data1.Recordset("VP_PT")) Then failed = failed & Data1.Recordset("VP_PT") & ", "
                If Not IsNull(Data1.Recordset("VP_GRPCD")) Then failed = failed & Data1.Recordset("VP_GRPCD") & ", "
                If Not IsNull(Data1.Recordset("VP_LOC")) Then failed = failed & Data1.Recordset("VP_LOC") & ", "
                If Not IsNull(Data1.Recordset("VP_SECTION")) Then failed = failed & Data1.Recordset("VP_SECTION") & ", "
                If Not IsNull(Data1.Recordset("VP_FRDATE")) Then failed = failed & Data1.Recordset("VP_FRDATE") & ", "
                If Not IsNull(Data1.Recordset("VP_TODATE")) Then failed = failed & Data1.Recordset("VP_TODATE") & ", "
                failed = Left(failed, Len(failed) - 2) & vbCrLf
            End If
        End If
        c = c + 1
        Data1.Recordset.MoveNext
    Loop Until Data1.Recordset.EOF
End If

Data1.Refresh

Call Display_Value

Screen.MousePointer = DEFAULT

If Len(failed) = 0 Then
    MsgBox "All Rules applied", vbInformation + vbOKOnly, "Vacation Pay Percentage"
Else
    MsgBox "The Following Rules failed:" & vbCrLf & failed, vbInformation + vbOKOnly, "Vacation Pay Percentage"
End If

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdateAll", "Single", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
     RollBack
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub Form_Activate()

Call SET_UP_MODE
Call INI_Controls(Me)

glbOnTop = "frmSVacPayPrct"

End Sub

Private Sub Form_Load()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    Dim Answer, DefVal, Msg, Title  ' Declare variables.
    Dim RFound As Integer ' records found
    Dim x%
    Dim SQLQ
    
    glbOnTop = "FRMSVACPAYPRCT"
    
    FlagRefresh = False
    
    Data1.ConnectionString = glbAdoIHRDB
    SQLQ = "SELECT DISTINCT VP_DIV,VP_DEPT,VP_ORG,VP_LOC,VP_SECTION,VP_EMP,VP_PT,VP_GRPCD,VP_FRDATE,VP_TODATE,VP_MANUAL,VP_EDATE FROM HRVACPCTENT "
    
    If glbDIVCount = 1 And glbLinamar Then
        SQLQ = SQLQ & " WHERE VP_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
    End If
    Data1.RecordSource = SQLQ
    Data1.Refresh
    
    'If UCase(glbCompEntSick$) = "M" Or UCase(glbCompEntSick$) = "N" Then
    '    vbxTrueGrid.Columns(5).Visible = False
    'End If
    
    Screen.MousePointer = HOURGLASS
    vbxTrueGrid.Columns(0).Caption = lStr(vbxTrueGrid.Columns(0).Caption)
    vbxTrueGrid.Columns(1).Caption = lStr(vbxTrueGrid.Columns(1).Caption)
    vbxTrueGrid.Columns(2).Caption = lStr(vbxTrueGrid.Columns(2).Caption)
    vbxTrueGrid.Columns(3).Visible = False
    
    Call setRptCaption(Me)
    
    Select Case glbCompWDate$ ' sets field reference for basic 'which date'
        Case "O": fglbWDate$ = "ED_DOH"
        Case "S": fglbWDate$ = "ED_SENDTE"
        Case "U": fglbWDate$ = "ED_UNION"
        Case "L": fglbWDate$ = "ED_LTHIRE"
        Case "D": fglbWDate$ = "ED_USRDAT1"
    End Select
    
    If glbSyndesis Then
        lblCriteria(5).Caption = "Position Grade"
        vbxTrueGrid.Columns(8).Caption = "Position Grade"
        clpCode(2).Tag = "00-Enter Position Grade"
    End If
    If glbWFC Then
        lblSection.FontBold = True
    End If
    
    'Ticket #29617 - Mississaugas of Scugog Island First Nation
    If glbCompSerial = "S/N - 2485W" Then
        lblHeading(0).Caption = "Service Ranges (in Months)"
        'lblPeriod.Visible = False
        'dlpDateRange(0).Visible = False
        'dlpDateRange(1).Visible = False
        lblPeriod.Caption = "Entitlement Period"
        lblAsOf.Visible = True
        dlpAsOf.Visible = True
        
        'lblAsOf.Left = lblPeriod.Left
        'lblAsOf.Top = lblPeriod.Top
        'dlpAsOf.Left = dlpDateRange(0).Left - 100
        'dlpAsOf.Top = dlpDateRange(0).Top
        
        VacFram.Top = 4800
        vbxTrueGrid.Columns(3).Visible = True
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
    If Me.Height >= 4650 + VacFram.Height + panControls.Height + 230 Then
        scrControl.Value = 0
        VacFram.Top = 4800  '4440
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
If IsNumeric(medEntitle(Index)) Then
    If Len(medEntitle(Index)) > 0 Then
        medEntitle(Index) = medEntitle(Index) * 100
    End If
End If
End Sub

Private Sub medEntitle_LostFocus(Index As Integer)
If IsNumeric(medEntitle(Index)) Then
    If Len(medEntitle(Index)) > 0 Then
        medEntitle(Index) = medEntitle(Index) / 100
    End If
End If
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

Private Function modUpdateSelection()
Dim EmpNo As Long
Dim dblServiceHours#
Dim varStartDate As Variant
Dim lngRecs&
Dim dblVacPayPct#, intWhereFit&, x%, Y%, z%
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%, xAsOf
Dim rsAudit As New ADODB.Recordset
Dim xPT As String
Dim xDiv As String
Dim OVACPC

On Error GoTo modUpdateSelection_Err

modUpdateSelection = False

If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)

Screen.MousePointer = DEFAULT

If snapEntitle.BOF And snapEntitle.EOF Then
    MsgBox "Employees for this selection do not exist!"
    Exit Function
Else
    lngRecs& = snapEntitle.RecordCount
    
    Msg$ = lngRecs& & " Records to process" & Chr(10) & "Would You Like To Proceed?"
    Title$ = "Update Vacation Pay Percentage"
    DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        Exit Function
    End If
    
    Screen.MousePointer = HOURGLASS
End If

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 5

For x% = 0 To 24
    If Not IsNumeric(medLTServ(x%)) Then
        medLTServ(x%) = 0
    End If
    If Not IsNumeric(medGTServ(x%)) Then
      medGTServ(x%) = 0
    Else
      If Val(medGTServ(x%)) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
    End If
    If medLTServ(x%) > 0 And medGTServ(x%) = 0 Then medGTServ(x%) = 9999999
Next

While Not snapEntitle.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / lngRecs&))
    MDIMain.panHelp(0).FloodPercent = pct%

    'If snapEntitle("ED_EMPNBR") = 3190 Then
    '    EmpNo& = snapEntitle("ED_EMPNBR")
    'End If

    EmpNo& = snapEntitle("ED_EMPNBR")

    'Ticket #29617 - Mississaugas of Scugog Island First Nation
    'Get the length of service by months
    If glbCompSerial = "S/N - 2485W" Then
        'Vacation / Sick Mass Updated Based Upon
        If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec

        varStartDate = snapEntitle(fglbWDate$)
        
        xAsOf = dlpAsOf.Text
        
        'Length of Service in Months based on Vacation / Sick Mass Update Based Upon
        dblServiceHours# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    Else
        'Get Total Non Absent Hours from Attendance and Attendance History
        dblServiceHours# = Total_NonAbsent_Hours(snapEntitle("ED_EMPNBR"), dlpDateRange(0).Text, dlpDateRange(1).Text)
    End If
    
    intWhereFit& = -1

    For x% = 0 To 24
        If medGTServ(x%) > 0 Then
            If dblServiceHours# >= CDbl(medLTServ(x%)) And dblServiceHours# <= CDbl(medGTServ(x%)) Then
                intWhereFit& = x%
                Exit For
            End If
        End If
    Next x%
    
    If intWhereFit& = -1 Or dblServiceHours# < 0 Then GoTo lblNextRec ' skip record if not in any of the ranges

    dblVacPayPct# = medEntitle(intWhereFit&)
    
    OVACPC = ""
    OVACPC = snapEntitle("ED_VACPC")
    
    snapEntitle("ED_VACPC") = dblVacPayPct# '* 100
    snapEntitle("ED_LDATE") = Now
    snapEntitle("ED_LTIME") = Time$
    snapEntitle("ED_LUSER") = glbUserID
    snapEntitle.Update
    
    
    'Update Audit
    If OVACPC <> dblVacPayPct# Then
        'Retrieve PT and Div from HREMP
        If IsNull(snapEntitle("ED_PT")) Then xPT = "" Else xPT = snapEntitle("ED_PT")
        If IsNull(snapEntitle("ED_DIV")) Then xDiv = "" Else xDiv = snapEntitle("ED_DIV")
                            
        'Add Audit Log
        rsAudit.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        
        rsAudit.AddNew
        rsAudit("AU_LOC_TABL") = "EDLC": rsAudit("AU_SECTION_TABL") = "EDSE": rsAudit("AU_EMP_TABL") = "EDEM": rsAudit("AU_SUPCODE_TABL") = "EDSP": rsAudit("AU_ORG_TABL") = "EDOR": rsAudit("AU_PAYP_TABL") = "SDPP": rsAudit("AU_BCODE_TABL") = "BNCD": rsAudit("AU_TREAS_TABL") = "TERM": rsAudit("AU_DOLENT_TABL") = "EDOL": rsAudit("AU_EARN_TABL") = "EARN"
        rsAudit("AU_ADMINBY_TABL") = "EDAB": rsAudit("AU_LANG1_TABL") = "EDL1": rsAudit("AU_LANG2_TABL") = "EDL1"
        
        rsAudit("AU_NEWEMP") = "N"
        rsAudit("AU_PTUPL") = xPT
        rsAudit("AU_DIVUPL") = xDiv
        rsAudit("AU_COMPNO") = "001"
        rsAudit("AU_EMPNBR") = snapEntitle("ED_EMPNBR")
                
        If OVACPC <> dblVacPayPct# Then
            If IsNumeric(dblVacPayPct#) Then rsAudit("AU_VACPC") = dblVacPayPct# * 100
            If IsNumeric(OVACPC) Then rsAudit("AU_OLDVAC") = OVACPC * 100
        End If
        
        rsAudit("AU_LDATE") = Date
        rsAudit("AU_LUSER") = glbUserID
        rsAudit("AU_LTIME") = Time$
        rsAudit("AU_UPLOAD") = "N"
        rsAudit("AU_TYPE") = "M"
        
        rsAudit.Update
        rsAudit.Close
        Set rsAudit = Nothing
    End If
    
lblNextRec:
    DoEvents
    
    snapEntitle.MoveNext
Wend

modUpdateSelection = True
MDIMain.panHelp(0).FloodType = 0

snapEntitle.Close
Set snapEntitle = Nothing

Screen.MousePointer = DEFAULT

Exit Function

modUpdateSelection_Err:
'These errors are:
'13=type mismatch
'94=invalid use of null
'3018=couln't find field 'item'
If Err = 13 Or Err = 94 Or Err = 3018 Then
   ' MsgBox "Err:" & Str(Err) & Chr(10) & Error$ & Chr(10) & " modUpdateSelection" & Chr(10) & "FORM:FUENTITL.FRM"
    'commented out by RAUBREY 5/20/97
    Err = 0
    Resume Next
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateVacationPay%", "HREMP", "edit/Add")

Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If
End Function

Private Sub scrControl_Change()
    VacFram.Top = 4800 - scrControl.Value
End Sub

Sub ST_UPD_MODE(TF As Boolean)
    Dim x, FT
    FT = Not TF
    For x = 0 To 24
        medLTServ(x).Enabled = TF
        medGTServ(x).Enabled = TF
        medEntitle(x).Enabled = TF
    '    If X = 0 Then
    '    optD(X).Enabled = TF
    '    optH(X).Enabled = TF
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
    'If Not TF Or glbLinamar Then
    '    lblAsOf.FontBold = True
    'Else
    '    lblAsOf.FontBold = False
    'End If
    'If glbCompEntSick$ = "M" Or glbCompEntSick$ = "N" Or glbCompEntSick$ = "A" Then
    '    dlpAsOf.Enabled = True 'FT
    'Else
    '    dlpAsOf.Enabled = True 'Ticket #3419
    'End If
    'If sick Entitlement Outstanding based on "1" then ok, otherwise disenable
    'If glbEntOutStandingS$ = "1" Then
    '    CmdRecalc.Enabled = True
    'Else
    '    CmdRecalc.Enabled = False
    'End If
    If Not glbWHSCC Then
        clpCode(1).Enabled = TF
    Else
        clpCode(1).Enabled = False
    End If
    clpCode(2).Enabled = TF
    clpCode(3).Enabled = TF
    clpCode(4).Enabled = TF
    clpPT.Enabled = TF
    dlpDateRange(0).Enabled = TF
    dlpDateRange(1).Enabled = TF
    
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
    Dim x
    For x = 0 To 24
        medLTServ(x) = ""
        medGTServ(x) = ""
        medEntitle(x) = ""
    '    optD(X) = True
    '    optH(X) = False
    '    optF(X) = False
    '    medMax(X) = ""
    Next
    clpDiv.Text = ""
    clpDept.Text = ""
    clpCode(0).Text = ""
    'If Not (glbCompEntSick$ = "M" Or glbCompEntSick$ = "N") Then
    '    dlpAsOf.Text = ""
    'End If
    clpCode(1).Text = ""
    clpCode(2).Text = ""
    clpCode(3).Text = ""
    clpCode(4).Text = ""
    clpPT.Text = ""
    dlpDateRange(0).Text = ""
    dlpDateRange(1).Text = ""
    
    If Not Data1.Recordset.EOF Then
        SQLQ = "SELECT * FROM HRVACPCTENT "
        If IsNull(Data1.Recordset("VP_DIV")) Then
            SQLQ = SQLQ & " WHERE VP_DIV IS NULL"
        Else
            SQLQ = SQLQ & " WHERE VP_DIV = '" & Data1.Recordset("VP_DIV") & "'"
        End If
        If IsNull(Data1.Recordset("VP_DEPT")) Then
            SQLQ = SQLQ & " AND VP_DEPT IS NULL"
        Else
            SQLQ = SQLQ & " AND VP_DEPT = '" & Data1.Recordset("VP_DEPT") & "'"
        End If
        If IsNull(Data1.Recordset("VP_ORG")) Then
            SQLQ = SQLQ & " AND VP_ORG IS NULL"
        Else
            SQLQ = SQLQ & " AND VP_ORG = '" & Data1.Recordset("VP_ORG") & "'"
        End If
        If IsNull(Data1.Recordset("VP_LOC")) Then
            SQLQ = SQLQ & " AND VP_LOC IS NULL"
        Else
            SQLQ = SQLQ & " AND VP_LOC = '" & Data1.Recordset("VP_LOC") & "'"
        End If
        If IsNull(Data1.Recordset("VP_SECTION")) Then
            SQLQ = SQLQ & " AND VP_SECTION IS NULL"
        Else
            SQLQ = SQLQ & " AND VP_SECTION = '" & Data1.Recordset("VP_SECTION") & "'"
        End If
        'Ticket #29617 - Mississaugas of Scugog Island First Nation
        If Not IsNull(Data1.Recordset("VP_EDATE")) Then
            SQLQ = SQLQ & " AND VP_EDATE = " & Date_SQL(Data1.Recordset("VP_EDATE"))
        End If
        
        If IsNull(Data1.Recordset("VP_EMP")) Then
            SQLQ = SQLQ & " AND VP_EMP IS NULL"
        Else
            SQLQ = SQLQ & " AND VP_EMP = '" & Data1.Recordset("VP_EMP") & "'"
        End If
        If IsNull(Data1.Recordset("VP_PT")) Then
            SQLQ = SQLQ & " AND VP_PT IS NULL"
        Else
            SQLQ = SQLQ & " AND VP_PT = '" & Data1.Recordset("VP_PT") & "' "
        End If
        If IsNull(Data1.Recordset("VP_GRPCD")) Then
            SQLQ = SQLQ & " AND VP_GRPCD IS NULL"
        Else
            SQLQ = SQLQ & " AND VP_GRPCD = '" & Data1.Recordset("VP_GRPCD") & "'"
        End If
        
        If Not IsNull(Data1.Recordset("VP_FRDATE")) Then
            SQLQ = SQLQ & " AND VP_FRDATE = " & Date_SQL(Data1.Recordset("VP_FRDATE"))
        End If
        If Not IsNull(Data1.Recordset("VP_TODATE")) Then
            SQLQ = SQLQ & " AND VP_TODATE = " & Date_SQL(Data1.Recordset("VP_TODATE"))
        End If
        
        SQLQ = SQLQ & " ORDER BY VP_DIV,VP_DEPT,VP_ORG,VP_EMP,VP_PT,VP_LOC,VP_SECTION,VP_ORDER "
        rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset
        If Not IsNull(Data1.Recordset("VP_DIV")) Then clpDiv.Text = Data1.Recordset("VP_DIV")
        If Not IsNull(Data1.Recordset("VP_DEPT")) Then clpDept.Text = Data1.Recordset("VP_DEPT")
        If Not IsNull(Data1.Recordset("VP_ORG")) Then clpCode(0).Text = Data1.Recordset("VP_ORG")
        'Ticket #29617 - Mississaugas of Scugog Island First Nation
        If Not IsNull(Data1.Recordset("VP_EDATE")) Then dlpAsOf.Text = Data1.Recordset("VP_EDATE")
        If Not IsNull(Data1.Recordset("VP_EMP")) Then clpCode(1).Text = Data1.Recordset("VP_EMP")
        If Not IsNull(Data1.Recordset("VP_PT")) Then clpPT.Text = Data1.Recordset("VP_PT")
        If Not IsNull(Data1.Recordset("VP_GRPCD")) Then clpCode(2).Text = Data1.Recordset("VP_GRPCD")
        If Not IsNull(Data1.Recordset("VP_LOC")) Then clpCode(4).Text = Data1.Recordset("VP_LOC")
        If Not IsNull(Data1.Recordset("VP_SECTION")) Then clpCode(3).Text = Data1.Recordset("VP_SECTION")
        
        If Not IsNull(Data1.Recordset("VP_FRDATE")) Then dlpDateRange(0).Text = Data1.Recordset("VP_FRDATE")
        If Not IsNull(Data1.Recordset("VP_TODATE")) Then dlpDateRange(1).Text = Data1.Recordset("VP_TODATE")

        If Not IsNull(Data1.Recordset("VP_MANUAL")) Then chkManual.Value = Data1.Recordset("VP_MANUAL")
        
        Do While Not rsVE.EOF
            xOrder = rsVE("VP_ORDER")
            nOrder = Format(Val(xOrder), "##0") - 1
            If Not (nOrder < 0 Or nOrder > 24) Then
                If Not IsNull(rsVE("VP_BHOUR")) Then medLTServ(nOrder) = rsVE("VP_BHOUR")
                If Not IsNull(rsVE("VP_EHOUR")) Then medGTServ(nOrder) = rsVE("VP_EHOUR")
                If Not IsNull(rsVE("VP_PCT")) Then medEntitle(nOrder) = rsVE("VP_PCT")
    '            If rsVE("VE_TYPE") = "D" Then optD(nOrder) = True
    '            If rsVE("VE_TYPE") = "H" Then optH(nOrder) = True
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
    
    SQLQ = "SELECT DISTINCT VP_DIV,VP_DEPT,VP_ORG,VP_LOC,VP_SECTION,VP_EMP,VP_PT,VP_GRPCD,VP_FRDATE,VP_TODATE,VP_MANUAL,VP_EDATE FROM HRVACPCTENT"
    If glbDIVCount = 1 And glbLinamar Then
        SQLQ = SQLQ & " WHERE VP_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Display_Value
End Sub

Private Sub getWSQLQ(xType)
Dim xDiv, xDept, xORG, xAsOf, xEMP, xEmpMode, xGRPCE
Dim xLoc, xSECTION
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
    xSECTION = OSection
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
    xSECTION = clpCode(3).Text
    
    xFromDate = dlpDateRange(0)
    xToDate = dlpDateRange(1)
End If

If Len(xDiv) = 0 Then
    fglbVSQLQ = " (VP_DIV IS NULL OR VP_DIV='')"
Else
    fglbVSQLQ = "VP_DIV = '" & xDiv & "'"
End If
If Len(xDept) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VP_DEPT IS NULL OR VP_DEPT='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND VP_DEPT = '" & xDept & "'"
End If
If Len(xORG) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VP_ORG IS NULL OR VP_ORG='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VP_ORG = '" & xORG & "'"
End If
'If UCase(glbCompEntSick$) = "A" Then
'    If Len(xAsOf) > 0 Then fglbVSQLQ = fglbVSQLQ & " AND  VE_EDATE = " & Date_SQL(xAsOf)
'End If
If Len(xEMP) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VP_EMP IS NULL OR VP_EMP='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND VP_EMP = '" & xEMP & "'"
End If
If Len(xEmpMode) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VP_PT IS NULL OR VP_PT='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND VP_PT = '" & xEmpMode & "' "
End If
If Len(xGRPCE) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VP_GRPCD IS NULL OR VP_GRPCD='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VP_GRPCD = '" & xGRPCE & "'"
End If

If Len(xLoc) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VP_LOC IS NULL OR VP_LOC='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VP_LOC = '" & xLoc & "'"
End If
If Len(xSECTION) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VP_SECTION IS NULL OR VP_SECTION='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VP_SECTION = '" & xSECTION & "'"
End If

'Sam 02/03/2006
If Not IsDate(xFromDate) Then
    fglbVSQLQ = fglbVSQLQ & " AND VP_FRDATE IS NULL  "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VP_FRDATE = " & Date_SQL(xFromDate)
End If
If Not IsDate(xToDate) Then
    fglbVSQLQ = fglbVSQLQ & " AND VP_TODATE IS NULL  "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VP_TODATE = " & Date_SQL(xToDate)
End If
'Sam 02/03/2006

'Ticket #29617 - Mississaugas of Scugog Island First Nation
If Not IsDate(xAsOf) Then
    fglbVSQLQ = fglbVSQLQ & " AND VP_EDATE IS NULL  "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VP_EDATE = " & Date_SQL(xAsOf)
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
    cmdUpdateAll.Enabled = False
    CmdRecalc.Enabled = False
ElseIf Me.Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
    cmdPrintAll.Enabled = True
    cmdUpdate.Enabled = False
    cmdUpdateAll.Enabled = False
    CmdRecalc.Enabled = False
Else
    UpdateState = OPENING
    TF = True
    cmdPrintAll.Enabled = True
    cmdUpdate.Enabled = True
    cmdUpdateAll.Enabled = True
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

Private Function DoWork() As Boolean
    Dim lastday
    Dim flglastdate As Boolean
    Dim lngRecs As Long, pct As Long, prec As Long

    Screen.MousePointer = DEFAULT
    
    DoWork = False
    
    If Not modUpdateSelection() Then Exit Function
        
    Screen.MousePointer = HOURGLASS

    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    
    DoWork = True
    
End Function

