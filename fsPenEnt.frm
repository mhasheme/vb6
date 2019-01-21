VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSPenEnt 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Pension Entitlement Master"
   ClientHeight    =   10950
   ClientLeft      =   165
   ClientTop       =   -1560
   ClientWidth     =   11400
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
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   3825
      LargeChange     =   315
      Left            =   10920
      Max             =   100
      SmallChange     =   315
      TabIndex        =   111
      Top             =   4140
      Width           =   300
   End
   Begin VB.Frame VacFram03 
      BorderStyle     =   0  'None
      Height          =   4185
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11415
      Begin Threed.SSCheck chkManual 
         Height          =   255
         Left            =   5400
         TabIndex        =   12
         Top             =   3000
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Exclude from Update All     "
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   7020
         TabIndex        =   8
         Tag             =   "00-Position Group - Code"
         Top             =   2400
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "JBGC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   1305
         TabIndex        =   3
         Tag             =   "00-Enter Union Code"
         Top             =   2400
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   1305
         TabIndex        =   2
         Tag             =   "00-Specific Department Desired"
         Top             =   2100
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
         Left            =   1305
         TabIndex        =   1
         Tag             =   "00-Specific Division Desired"
         Top             =   1800
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
         Left            =   7020
         TabIndex        =   6
         Tag             =   "00-Specific Employment Status Desired"
         Top             =   1800
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDEM"
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         Height          =   285
         Left            =   7020
         TabIndex        =   7
         Tag             =   "EDPT-Category"
         Top             =   2100
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDPT"
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
         Bindings        =   "fsPenEnt.frx":0000
         Height          =   1335
         Left            =   0
         OleObjectBlob   =   "fsPenEnt.frx":0014
         TabIndex        =   0
         Top             =   120
         Width           =   9135
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   7020
         TabIndex        =   9
         Tag             =   "00-Section - Code"
         Top             =   2700
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   1305
         TabIndex        =   4
         Tag             =   "00-Enter Location Code"
         Top             =   2700
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.DateLookup dlpAsOf 
         Height          =   285
         Left            =   1305
         TabIndex        =   10
         Tag             =   "40-As of Date"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1390
      End
      Begin Threed.SSCheck chkUseService 
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   3000
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Use Pension Date"
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
      Begin INFOHR_Controls.CodeLookup clpSalDist 
         Height          =   285
         Left            =   7020
         TabIndex        =   5
         Top             =   1480
         Visible         =   0   'False
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   6
         LookupType      =   8
      End
      Begin VB.Label lblSalDist 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Distribution"
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
         Left            =   5400
         TabIndex        =   133
         Top             =   1515
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label lblAsOf 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
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
         TabIndex        =   132
         Top             =   3000
         Visible         =   0   'False
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
         TabIndex        =   115
         Top             =   2730
         Width           =   1215
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
         Left            =   5400
         TabIndex        =   114
         Top             =   2700
         Width           =   1260
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pension Entitlement Percentage"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   3600
         TabIndex        =   113
         Top             =   3960
         Width           =   2730
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
         Left            =   5400
         TabIndex        =   112
         Top             =   2100
         Width           =   630
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
         TabIndex        =   23
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Ranges (in Years)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   22
         Top             =   3960
         Width           =   2235
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
         Left            =   5400
         TabIndex        =   21
         Top             =   2400
         Width           =   1260
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
         TabIndex        =   20
         Top             =   3720
         Visible         =   0   'False
         Width           =   7455
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
         Left            =   5400
         TabIndex        =   19
         Top             =   1800
         Width           =   1350
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
         TabIndex        =   18
         Top             =   2430
         Width           =   420
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
         TabIndex        =   17
         Top             =   2100
         Width           =   825
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
         TabIndex        =   16
         Top             =   1800
         Width           =   555
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   14
      Top             =   10320
      Width           =   11400
      _Version        =   65536
      _ExtentX        =   20108
      _ExtentY        =   1111
      _StockProps     =   15
      ForeColor       =   0
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
      Font3D          =   1
      Alignment       =   1
      Begin VB.CommandButton cmdUpdateAll 
         Caption         =   "Update All"
         Height          =   375
         Left            =   3600
         TabIndex        =   85
         Top             =   120
         Width           =   1665
      End
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update Entitlement"
         Height          =   375
         Left            =   1560
         TabIndex        =   84
         Tag             =   "Change all matching records to the above"
         Top             =   120
         Width           =   1905
      End
      Begin VB.CommandButton cmdPrintAll 
         Appearance      =   0  'Flat
         Caption         =   "Print &All"
         Height          =   375
         Left            =   240
         TabIndex        =   82
         Tag             =   "Print all Vacation Entitlement Report"
         Top             =   120
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   405
         Left            =   7080
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
      Left            =   120
      TabIndex        =   13
      Top             =   4200
      Width           =   11000
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   24
         Tag             =   "11-Service is greater than this number"
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   25
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
         Index           =   2
         Left            =   0
         TabIndex        =   30
         Tag             =   "11-Service is greater than this number"
         Top             =   742
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   0
         Left            =   4275
         TabIndex        =   26
         Tag             =   "10-Vacation Pay Percentage"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   3
         Left            =   4275
         TabIndex        =   35
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   1064
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   4
         Left            =   4275
         TabIndex        =   38
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   1386
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   5
         Left            =   4275
         TabIndex        =   41
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   1708
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   6
         Left            =   4275
         TabIndex        =   44
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   2030
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   7
         Left            =   4275
         TabIndex        =   47
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   2352
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   8
         Left            =   4275
         TabIndex        =   50
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   2674
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   9
         Left            =   4275
         TabIndex        =   53
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   2996
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
         TabIndex        =   33
         Tag             =   "11-Service is greater than this number"
         Top             =   1064
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   34
         Tag             =   "10-Service is less than this number"
         Top             =   1064
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
         TabIndex        =   36
         Tag             =   "11-Service is greater than this number"
         Top             =   1386
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   37
         Tag             =   "10-Service is less than this number"
         Top             =   1386
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
         TabIndex        =   39
         Tag             =   "11-Service is greater than this number"
         Top             =   1708
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   40
         Tag             =   "10-Service is less than this number"
         Top             =   1708
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
         Index           =   6
         Left            =   0
         TabIndex        =   42
         Tag             =   "11-Service is greater than this number"
         Top             =   2030
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   6
         Left            =   2160
         TabIndex        =   43
         Tag             =   "10-Service is less than this number"
         Top             =   2030
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
         TabIndex        =   45
         Tag             =   "11-Service is greater than this number"
         Top             =   2352
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   7
         Left            =   2160
         TabIndex        =   46
         Tag             =   "10-Service is less than this number"
         Top             =   2352
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
         TabIndex        =   48
         Tag             =   "11-Service is greater than this number"
         Top             =   2674
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   8
         Left            =   2160
         TabIndex        =   49
         Tag             =   "10-Service is less than this number"
         Top             =   2674
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
         Index           =   9
         Left            =   0
         TabIndex        =   51
         Tag             =   "11-Service is greater than this number"
         Top             =   2996
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   9
         Left            =   2160
         TabIndex        =   52
         Tag             =   "10-Service is less than this number"
         Top             =   2996
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   10
         Left            =   4275
         TabIndex        =   56
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   3318
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   11
         Left            =   4275
         TabIndex        =   59
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   3640
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   12
         Left            =   4275
         TabIndex        =   62
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   3962
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   13
         Left            =   4275
         TabIndex        =   65
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   4284
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   14
         Left            =   4275
         TabIndex        =   68
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   4606
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
         TabIndex        =   54
         Tag             =   "11-Service is greater than this number"
         Top             =   3318
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   10
         Left            =   2160
         TabIndex        =   55
         Tag             =   "10-Service is less than this number"
         Top             =   3318
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
         TabIndex        =   57
         Tag             =   "11-Service is greater than this number"
         Top             =   3640
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   11
         Left            =   2160
         TabIndex        =   58
         Tag             =   "10-Service is less than this number"
         Top             =   3640
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
         Index           =   12
         Left            =   0
         TabIndex        =   60
         Tag             =   "11-Service is greater than this number"
         Top             =   3962
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   12
         Left            =   2160
         TabIndex        =   61
         Tag             =   "10-Service is less than this number"
         Top             =   3962
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
         TabIndex        =   63
         Tag             =   "11-Service is greater than this number"
         Top             =   4284
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   13
         Left            =   2160
         TabIndex        =   64
         Tag             =   "10-Service is less than this number"
         Top             =   4284
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
         TabIndex        =   66
         Tag             =   "11-Service is greater than this number"
         Top             =   4606
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   14
         Left            =   2160
         TabIndex        =   67
         Tag             =   "10-Service is less than this number"
         Top             =   4606
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   15
         Left            =   4275
         TabIndex        =   71
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   4928
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   16
         Left            =   4275
         TabIndex        =   74
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   5250
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
         TabIndex        =   69
         Tag             =   "11-Service is greater than this number"
         Top             =   4928
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   15
         Left            =   2160
         TabIndex        =   70
         Tag             =   "10-Service is less than this number"
         Top             =   4928
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
         TabIndex        =   72
         Tag             =   "11-Service is greater than this number"
         Top             =   5250
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   16
         Left            =   2160
         TabIndex        =   73
         Tag             =   "10-Service is less than this number"
         Top             =   5250
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   17
         Left            =   4275
         TabIndex        =   77
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   5572
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
         TabIndex        =   75
         Tag             =   "11-Service is greater than this number"
         Top             =   5572
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   17
         Left            =   2160
         TabIndex        =   76
         Tag             =   "10-Service is less than this number"
         Top             =   5572
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   18
         Left            =   4275
         TabIndex        =   80
         Tag             =   "10-Vacation Pay Percentage"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   19
         Left            =   4275
         TabIndex        =   86
         Tag             =   "10-Vacation Pay Percentage"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   20
         Left            =   4275
         TabIndex        =   89
         Tag             =   "10-Vacation Pay Percentage"
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   21
         Left            =   4275
         TabIndex        =   92
         Tag             =   "10-Vacation Pay Percentage"
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   22
         Left            =   4275
         TabIndex        =   124
         Tag             =   "10-Vacation Pay Percentage"
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
         Index           =   18
         Left            =   0
         TabIndex        =   78
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   18
         Left            =   2160
         TabIndex        =   79
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
         TabIndex        =   81
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
         TabIndex        =   83
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   20
         Left            =   0
         TabIndex        =   87
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
         TabIndex        =   88
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
         TabIndex        =   90
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
         TabIndex        =   91
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
         TabIndex        =   93
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
         TabIndex        =   94
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   23
         Left            =   4275
         TabIndex        =   127
         Tag             =   "10-Vacation Pay Percentage"
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   23
         Left            =   0
         TabIndex        =   125
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
         TabIndex        =   126
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
         TabIndex        =   128
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   24
         Left            =   4275
         TabIndex        =   129
         Tag             =   "10-Vacation Pay Percentage"
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   24
         Left            =   2160
         TabIndex        =   130
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   27
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   28
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   31
         Tag             =   "10-Service is less than this number"
         Top             =   742
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
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   1
         Left            =   4275
         TabIndex        =   29
         Tag             =   "10-Vacation Pay Percentage"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPension 
         Height          =   285
         Index           =   2
         Left            =   4275
         TabIndex        =   32
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   742
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
         Index           =   1
         Left            =   980
         TabIndex        =   131
         Top             =   465
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
         Index           =   24
         Left            =   975
         TabIndex        =   123
         Top             =   7590
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
         TabIndex        =   122
         Top             =   5617
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
         TabIndex        =   121
         Top             =   6285
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
         TabIndex        =   120
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
         TabIndex        =   119
         Top             =   7245
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
         TabIndex        =   118
         Top             =   6930
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
         TabIndex        =   117
         Top             =   6615
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
         TabIndex        =   116
         Top             =   7920
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
         TabIndex        =   110
         Top             =   5295
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
         TabIndex        =   109
         Top             =   2075
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
         TabIndex        =   108
         Top             =   2397
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
         TabIndex        =   107
         Top             =   2719
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
         TabIndex        =   106
         Top             =   1109
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
         TabIndex        =   105
         Top             =   1431
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
         TabIndex        =   104
         Top             =   1753
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
         TabIndex        =   103
         Top             =   787
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
         TabIndex        =   102
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
         TabIndex        =   101
         Top             =   4007
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
         TabIndex        =   100
         Top             =   4329
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
         TabIndex        =   99
         Top             =   4651
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
         TabIndex        =   98
         Top             =   3363
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
         TabIndex        =   97
         Top             =   3685
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
         TabIndex        =   96
         Top             =   3041
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
         TabIndex        =   95
         Top             =   4973
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmSPenEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fTablHREMP As New ADODB.Recordset         ' table view of HREMP
Dim snapEntitle As New ADODB.Recordset     'user vier
Dim fglbWDate$, fglbWDateS$
Dim xService(24, 4)
Dim xTypeD(24)
Dim xTypeH(24)
Dim xTypeF(24)

Dim fglbNoDept&
Dim fglb_FindDept
Dim fglbSick%
Dim fglbVac%

Dim Actn

Dim fglbSDate As Variant
Dim fglbMaxRange%
Dim fglbCompMonthly%

Dim fglbMaxRanges%
Dim glbFrmCaption$, glbErrNum&

Dim ControlsShown As Boolean
Dim ODIV, ODept, oOrg, oAsOf, oEMP, oEmpMode, oGRPCE
Dim OSection, OLoc, OSALDIST
Dim OFromDate, OToDate
Dim FlagRefresh As Boolean

Dim fglbESQLQ, fglbVSQLQ
Dim fglbNew As Boolean
Dim fglbRunTimes

Private Function chkMUEntitle()
Dim X%, Y%

chkMUEntitle = False

On Error GoTo chkMUEntitle_Err
For X% = 0 To 4
If Len(clpCode(X%).Text) > 0 And clpCode(X%).Caption = "Unassigned" Then
    MsgBox "If Code entered it must be known"
    clpCode(X%).SetFocus
    Exit Function
End If
Next X%
If glbLinamar Then
    If Len(clpCode(3).Text) = 0 Then
        MsgBox "Vaction Group is required field"
        clpCode(3).SetFocus
        Exit Function
    End If
End If

If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
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
        MsgBox lStr("If Division Entered - it must be known")
         clpDiv.SetFocus
        Exit Function
    End If
End If
'Hemu - 05/13/2003 Begin
If clpPT.Caption = "Unassigned" Then
    MsgBox "If " & lblPT.Caption & " Entered - it must be known"
    clpPT.SetFocus
    Exit Function
End If
'Hemu - 05/13/2003 End


If Len(medLTServ(0)) < 1 Then
    MsgBox "You must have at least one Service Range Entry."
    If medLTServ(0).Enabled Then medLTServ(0).SetFocus
    Exit Function
End If

If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19938 Franks 05/26/2011
    If clpSalDist.Caption = "Unassigned" Then
        MsgBox "If " & lblSalDist.Caption & " Entered - it must be known"
        clpSalDist.SetFocus
        Exit Function
    End If
    If chkUseService.Value = False Then
        If Len(dlpAsOf.Text) = 0 Then
            MsgBox "Effective Date is required field if Use Pension Date is not checked"
            dlpAsOf.SetFocus
            Exit Function
        Else
            If Not IsDate(dlpAsOf.Text) Then
                MsgBox "Invalid Effective Date."
                dlpAsOf.SetFocus
                Exit Function
            End If
        End If
    End If
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


For X% = 0 To 24
    If Len(medLTServ(X%)) > 0 Then
        If Not IsNumeric(medLTServ(X%)) Then
            MsgBox "Data Entered Must Be Numeric"
            medLTServ(X%).SetFocus
            Exit Function
        End If
    End If
    If Len(medGTServ(X%)) > 0 Then
        If Not IsNumeric(medGTServ(X%)) Then
            MsgBox "Data Entered Must Be Numeric"
            medGTServ(X%).SetFocus
            Exit Function
        End If
    End If

    If Len(medPension(X%)) > 0 Then
        If Not IsNumeric(medPension(X%)) Then
            MsgBox "Data Entered Must Be Numeric"
            medPension(X%).SetFocus
            Exit Function
        End If
    End If

    If Len(medLTServ(X%)) < 1 And Len(medGTServ(X%)) > 1 Then  ' missed one
        MsgBox "Ranges must be sequential"
        medLTServ(X%).SetFocus
        Exit Function
    End If
    If Len(medGTServ(X%)) > 0 Then
        If glbFrench Then
            If CDbl(medLTServ(X%)) > CDbl(medGTServ(X%)) Then
                MsgBox "Ranges must be sequential"
                medLTServ(X%).SetFocus
                Exit Function
            End If
        Else
            If Val(medLTServ(X%)) > Val(medGTServ(X%)) Then
                MsgBox "Ranges must be sequential"
                medLTServ(X%).SetFocus
                Exit Function
            End If
        End If
    End If
    If X% > 0 And Len(medLTServ(X%)) > 0 Then
        If glbFrench Then
            If CDbl(medLTServ(X%)) < CDbl(medGTServ(X% - 1)) Then
                MsgBox "Ranges must be sequential"
                medLTServ(X%).SetFocus
                Exit Function
            End If
        Else
            If Val(medLTServ(X%)) < Val(medGTServ(X% - 1)) Then
                MsgBox "Ranges must be sequential"
                medLTServ(X%).SetFocus
                Exit Function
            End If
        End If
    End If
    If X% > 0 And Len(medGTServ(X%)) > 0 Then
        If glbFrench Then
            If CDbl(medGTServ(X%)) < CDbl(medGTServ(X% - 1)) And CDbl(medGTServ(X%)) <> 0 Then
                MsgBox "Ranges must be sequential"
                medLTServ(X%).SetFocus
                Exit Function
            End If
        Else
            If Val(medGTServ(X%)) < Val(medGTServ(X% - 1)) And Val(medGTServ(X%)) <> 0 Then
                MsgBox "Ranges must be sequential"
                medLTServ(X%).SetFocus
                Exit Function
            End If
        End If
    End If
    If Len(medLTServ(X%)) > 0 Or Len(medGTServ(X%)) > 0 Then
        If Len(medPension(X%)) < 1 Then
            MsgBox "Numeric Value Pension Entitlement Percentage Must Be Entered"
            medPension(X%).SetFocus
            Exit Function
        End If
    End If
    If Len(medLTServ(X%)) < 1 Then Exit For  ' missed one
    intRangesSet% = intRangesSet% + 1
Next X%
If intRangesSet% = 0 Then
    MsgBox "At least one Service level must be set"
    medLTServ(0).SetFocus
    Exit Function
End If

chkMUEntitle = True

Exit Function

chkMUEntitle_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEntitle", "HRBENFT", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub chkUseService_Click(Value As Integer)
If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19938 Franks 05/20/2011
    If chkUseService.Value Then
        chkManual.Value = True
        dlpAsOf.Text = ""
    End If
End If
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
     If glbWHSCC And Actn = "A" And Index = 0 Then
        If (clpCode(0) = "1866" Or clpCode(0) = "946") And clpPT = "FT" Then
            medLTServ(0) = 0: medGTServ(0) = 59.99
            medLTServ(1) = 60: medGTServ(1) = 215.99
            medLTServ(2) = 216: medGTServ(2) = 999
        End If
        If clpCode(0) = "NON" And clpPT = "FT" Then
           ' optD(0).SetFocus
            medLTServ(0) = 0: medGTServ(0) = 59.99
            medLTServ(1) = 60: medGTServ(1) = 108.99
            medLTServ(2) = 109: medGTServ(2) = 119.99
            medLTServ(3) = 120: medGTServ(3) = 131.99
            medLTServ(4) = 132: medGTServ(4) = 143.99
            medLTServ(5) = 144: medGTServ(5) = 155.99
            medLTServ(6) = 156: medGTServ(6) = 167.99
            medLTServ(7) = 168: medGTServ(7) = 179.99
            medLTServ(8) = 180: medGTServ(8) = 191.99
            medLTServ(9) = 192: medGTServ(9) = 203.99
            medLTServ(10) = 204: medGTServ(10) = 215.99
            medLTServ(11) = 216: medGTServ(11) = 999999.99
        End If
        If clpCode(0) = "PHYS" And clpPT = "FT" Then
            medLTServ(0) = 0: medGTServ(0) = 59.99
            medLTServ(1) = 60: medGTServ(1) = 119
        End If
        If clpCode(0) = "NON" And clpPT = "PT" Then
            'optF(0).SetFocus
            medLTServ(0) = 0: medGTServ(0) = 59.99
            medLTServ(1) = 60: medGTServ(1) = 108.99
            medLTServ(2) = 109: medGTServ(2) = 119.99
            medLTServ(3) = 120: medGTServ(3) = 131.99
            medLTServ(4) = 132: medGTServ(4) = 143.99
            medLTServ(5) = 144: medGTServ(5) = 155.99
            medLTServ(6) = 156: medGTServ(6) = 167.99
            medLTServ(7) = 168: medGTServ(7) = 179.99
            medLTServ(8) = 180: medGTServ(8) = 191.99
            medLTServ(9) = 192: medGTServ(9) = 203.99
            medLTServ(10) = 204: medGTServ(10) = 215.99
            medLTServ(11) = 216: medGTServ(11) = 999999.99
        End If
     End If
End Sub

Private Sub clpPT_LostFocus()
     If glbWHSCC And Actn = "A" Then  'And Index = 0 Then
        If (clpCode(0) = "1866" Or clpCode(0) = "946") And clpPT = "FT" Then
            medLTServ(0) = 0: medGTServ(0) = 59.99
            medLTServ(1) = 60: medGTServ(1) = 215.99
            medLTServ(2) = 216: medGTServ(2) = 999
        End If
        If clpCode(0) = "NON" And clpPT = "FT" Then
            'optD(0).SetFocus
            medLTServ(0) = 0: medGTServ(0) = 59.99
            medLTServ(1) = 60: medGTServ(1) = 119.99
            medLTServ(2) = 120: medGTServ(2) = 131.99
            medLTServ(3) = 132: medGTServ(3) = 143.99
            medLTServ(4) = 144: medGTServ(4) = 155.99
            medLTServ(5) = 156: medGTServ(5) = 167.99
            medLTServ(6) = 168: medGTServ(6) = 179.99
            medLTServ(7) = 180: medGTServ(7) = 191.99
            medLTServ(8) = 192: medGTServ(8) = 203.99
            medLTServ(9) = 204: medGTServ(9) = 215.99
            medLTServ(10) = 216: medGTServ(10) = 227.99
            medLTServ(11) = 228: medGTServ(11) = 999999.99
        End If
        If clpCode(0) = "PHYS" And clpPT = "FT" Then
            medLTServ(0) = 0: medGTServ(0) = 59.99
            medLTServ(1) = 60: medGTServ(1) = 119
        End If
        If clpCode(0) = "NON" And clpPT = "PT" Then
            'optF(0).SetFocus
            medLTServ(0) = 0: medGTServ(0) = 59.99
            medLTServ(1) = 60: medGTServ(1) = 119.99
            medLTServ(2) = 120: medGTServ(2) = 131.99
            medLTServ(3) = 132: medGTServ(3) = 143.99
            medLTServ(4) = 144: medGTServ(4) = 155.99
            medLTServ(5) = 156: medGTServ(5) = 167.99
            medLTServ(6) = 168: medGTServ(6) = 179.99
            medLTServ(7) = 180: medGTServ(7) = 191.99
            medLTServ(8) = 192: medGTServ(8) = 203.99
            medLTServ(9) = 204: medGTServ(9) = 215.99
            medLTServ(10) = 216: medGTServ(10) = 227.99
            medLTServ(11) = 228: medGTServ(11) = 999999.99
        End If
     End If
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
Dim SQLQ, Msg, a%
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If
Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "The Pension Entitlement Rules?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

Call getWSQLQ("C")
SQLQ = "DELETE FROM HRPENENT WHERE " & fglbVSQLQ

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

Data1.Refresh
Display_Value
End Sub

Sub cmdModify_Click()
ODIV = clpDiv.Text
ODept = clpDept.Text
oOrg = clpCode(0).Text

'Franks 04/08/03 Ticket# 3943
'Fix the problem: enter or change Effective Date first, click Edit and then Save,
'it create another record

OLoc = clpCode(4).Text
OSection = clpCode(3).Text
oEMP = clpCode(1).Text
oEmpMode = clpPT.Text
oGRPCE = clpCode(2).Text
OSALDIST = clpSalDist.Text
Actn = "M"
End Sub

Sub cmdNew_Click()
Dim X
For X = 0 To 24
    medLTServ(X) = ""
    medGTServ(X) = ""
    medPension(X) = ""
Next
clpDiv.Text = ""
clpDept.Text = ""
clpCode(0).Text = ""
clpCode(1).Text = ""
clpCode(2).Text = ""
clpCode(3).Text = ""
clpCode(4).Text = ""
clpPT.Text = ""
clpSalDist.Text = ""
Actn = "A"
fglbNew = True
Call SET_UP_MODE
clpDiv.SetFocus
End Sub
Sub cmdOK_Click()
Dim X%, Y%, xUnion, xPT, SQLQ, SQLQW
Dim xStr
Dim rsVE As New ADODB.Recordset
Dim rsVT As New ADODB.Recordset
Dim glbiOneWhere As Boolean

If Not chkMUEntitle() Then Exit Sub
For X% = 0 To 24
    If Not IsNumeric(medLTServ(X%)) Then Exit For
    If Not IsNumeric(medGTServ(X%)) Then
        medGTServ(X%) = 0
    Else
        If glbFrench Then
            If medGTServ(X%) = Int(medGTServ(X%)) Then medGTServ(X%) = medGTServ(X%) + 0.99
        Else
            If Val(medGTServ(X%)) = Int(medGTServ(X%)) Then medGTServ(X%) = medGTServ(X%) + 0.99
        End If
    End If
    If medLTServ(X%) > 0 And medGTServ(X%) = 0 Then medGTServ(X%) = 9999999
Next

If Actn = "M" Then
    Call getWSQLQ("O")
    SQLQ = "DELETE FROM HRPENENT WHERE " & fglbVSQLQ
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
Else
    Call getWSQLQ("C")
    SQLQ = "SELECT * FROM HRPENENT WHERE " & fglbVSQLQ
    rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsVT.EOF Then
        MsgBox "You can not add duplicate record"
         clpDiv.SetFocus
        Exit Sub
    End If
End If
gdbAdoIhr001.BeginTrans
SQLQ = "SELECT * FROM HRPENENT"
rsVE.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
For X% = 0 To 24
    If Len(medLTServ(X%)) > 0 Then
        rsVE.AddNew
        rsVE("PE_ORDER") = X + 1
        rsVE("PE_ORG_TABL") = "EDOR"
        rsVE("PE_ORG") = clpCode(0).Text
        rsVE("PE_PT") = clpPT.Text
        rsVE("PE_DIV") = clpDiv.Text
        rsVE("PE_DEPT") = clpDept.Text
        rsVE("PE_EMP_TABL") = "EDEM"
        rsVE("PE_EMP") = clpCode(1).Text
        rsVE("PE_SECTION") = clpCode(3).Text
        rsVE("PE_LOC") = clpCode(4).Text
        rsVE("PE_GRPCD_TABL") = "JBGC"
        rsVE("PE_GRPCD") = clpCode(2).Text
        If glbFrench Then
            rsVE("PE_BMONTH") = Replace(medLTServ(X%), ",", ".")
        Else
            rsVE("PE_BMONTH") = medLTServ(X%)
        End If
        If glbFrench Then
            rsVE("PE_EMONTH") = Replace(medGTServ(X%), ",", ".")
        Else
            rsVE("PE_EMONTH") = medGTServ(X%)
        End If
        If medPension(X%) = "" Then
            rsVE("PE_PCT") = Null
        Else
            If glbFrench Then
                rsVE("PE_PCT") = Replace(medPension(X%), ",", ".")
            Else
                rsVE("PE_PCT") = medPension(X%)
            End If
        End If
        rsVE("PE_MANUAL") = chkManual.Value
        If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19938 Franks 05/20/2011
            If IsDate(dlpAsOf.Text) Then
                rsVE("PE_EDATE") = CVDate(dlpAsOf.Text)
            Else
                rsVE("PE_EDATE") = Null
            End If
            rsVE("PE_USESERVICE") = chkUseService.Value
            'Ticket #22084 - Franks 05/25/2012
            If Len(clpSalDist.Text) > 0 Then rsVE("PE_SALDIST") = clpSalDist.Text
        End If
        rsVE("PE_LDATE") = Date
        rsVE("PE_LTIME") = Format(Time, "HH:MM:SS")
        rsVE("PE_LUSER") = glbUserID
        rsVE.Update
    End If
Next
rsVE.Close
gdbAdoIhr001.CommitTrans
'If Not glbSQL and not glboracle Then Call Pause(0.5)
Data1.Refresh
Display_Value
fglbNew = False
End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport, X%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%
'cmdPrint.Enabled = False

Me.vbxCrystal.Reset
Me.vbxCrystal.WindowTitle = "Pension Entitlement Master Report"
Call setRptLabel(Me, 0) '1)
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For X% = 0 To 6
        Me.vbxCrystal.DataFiles(X%) = glbIHRDB
    Next
End If
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgpenent.rpt"

SQLQ = "(1=1) "
If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_DIV} = '" & clpDiv.Text & "'"
If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_DEPT} = '" & clpDept.Text & "'"
If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_ORG} = '" & clpCode(0).Text & "'"
If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_EMP} = '" & clpCode(1).Text & "'"
If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_PT} = '" & clpPT.Text & "' "
If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_GRPCD} = '" & clpCode(2).Text & "'"
If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_SECTION} = '" & clpCode(3).Text & "'"
If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_LOC} = '" & clpCode(4).Text & "'"
Me.vbxCrystal.SelectionFormula = SQLQ
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub
Sub cmdView_Click()
Dim RHeading As String, xReport, X%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%
'cmdPrint.Enabled = False

Me.vbxCrystal.Reset
Me.vbxCrystal.WindowTitle = "Pension Entitlement Master Report"
Call setRptLabel(Me, 0) '1)
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For X% = 0 To 6
        Me.vbxCrystal.DataFiles(X%) = glbIHRDB
    Next
End If
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgpenent.rpt"

SQLQ = "(1=1) "
If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_DIV} = '" & clpDiv.Text & "'"
If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_DEPT} = '" & clpDept.Text & "'"
If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_ORG} = '" & clpCode(0).Text & "'"
If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_EMP} = '" & clpCode(1).Text & "'"
If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_PT} = '" & clpPT.Text & "' "
If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_GRPCD} = '" & clpCode(2).Text & "'"
If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_SECTION} = '" & clpCode(3).Text & "'"
If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND {HRPENENT.PE_LOC} = '" & clpCode(4).Text & "'"


Me.vbxCrystal.SelectionFormula = SQLQ
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
'cmdPrint.Enabled = True
End Sub

Private Sub cmdPrintAll_Click()
Dim RHeading As String, xReport, X%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%
cmdPrintAll.Enabled = False

Me.vbxCrystal.Reset

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.WindowTitle = "Pension Entitlement Master Report"
Call setRptLabel(Me, 0) '1)
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For X% = 0 To 6
        Me.vbxCrystal.DataFiles(X%) = glbIHRDB
    Next
End If
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgpenent.rpt"
Me.vbxCrystal.Action = 1

cmdPrintAll.Enabled = True
End Sub



Private Sub cmdUpdate_Click()
Dim Title$, Msg$, DgDef As Variant, Response%
On Error GoTo Mod_Err

If Not gSec_Upd_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If
If Not chkMUEntitle() Then Exit Sub

If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19938 Franks 05/26/2011
    If chkUseService.Value Then
        MsgBox "You can not do Update Entitlement if Use Pension Date is checked."
        Exit Sub
    End If
End If

    Call modUpdateSelection

Data1.Refresh
Call Display_Value
Screen.MousePointer = DEFAULT

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "Single", "Modify")
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

CR_SnapEntitle = False
On Error GoTo CR_SnapEntitle_Err

Screen.MousePointer = HOURGLASS

Call getWSQLQ("")

SQLQ = "SELECT ED_EMPNBR,ED_PENPCT,"
SQLQ = SQLQ & " ED_DIV,ED_PT, ED_SECTION,ED_SALDIST, ED_LOC, ED_ORG, ED_EMP,"
SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME "
If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19938 Franks 05/20/2011
    SQLQ = SQLQ & " ,ED_PENPCTFIXED,ED_OMERS "
End If
SQLQ = SQLQ & " FROM HREMP WHERE " & fglbESQLQ
If Len(clpCode(2).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_EMPNBR IN "
    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
    SQLQ = SQLQ & " WHERE JB_GRPCD = '" & clpCode(2).Text & "') "
End If
If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19938 Franks 05/20/2011
    SQLQ = SQLQ & "AND (ED_PENPCTFIXED IS NULL OR ED_PENPCTFIXED = 0) "
End If
'SQLQ = SQLQ & " AND ED_EMPNBR=8" 'FOR TESTING
If snapEntitle.State <> 0 Then snapEntitle.Close
snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic

CR_SnapEntitle = True

Exit Function

CR_SnapEntitle_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapEntitle", "Entitlements/EMP", "Select")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function





Private Sub cmdUpdateAll_Click()
'added by Bryan 25/Oct/05 Ticket#9560
Dim c As Long
Dim failed As String

On Error GoTo Mod_Err
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
        If chkMUEntitle() And chkManual.Value = False Then
            'Samuel  - Ticket #19938 Franks 05/20/2011
            'Note: if "Use Pension Date" was checked then "Exclude from Update All" must be checked
            If modUpdateSelection = False Then
                failed = "Rule " & CStr(c) & ": "
                If Not IsNull(Data1.Recordset("PE_DIV")) Then failed = failed & Data1.Recordset("PE_DIV") & ", "
                If Not IsNull(Data1.Recordset("PE_DEPT")) Then failed = failed & Data1.Recordset("PE_DEPT") & ", "
                If Not IsNull(Data1.Recordset("PE_ORG")) Then failed = failed & Data1.Recordset("PE_ORG") & ", "
                'If Not IsNull(Data1.Recordset("PE_EDATE")) Then failed = failed & Data1.Recordset("PE_EDATE") & ", "
                If Not IsNull(Data1.Recordset("PE_EMP")) Then failed = failed & Data1.Recordset("PE_EMP") & ", "
                If Not IsNull(Data1.Recordset("PE_PT")) Then failed = failed & Data1.Recordset("PE_PT") & ", "
                If Not IsNull(Data1.Recordset("PE_GRPCD")) Then failed = failed & Data1.Recordset("PE_GRPCD") & ", "
                If Not IsNull(Data1.Recordset("PE_LOC")) Then failed = failed & Data1.Recordset("PE_LOC") & ", "
                If Not IsNull(Data1.Recordset("PE_SECTION")) Then failed = failed & Data1.Recordset("PE_SECTION") & ", "
                'If Not IsNull(Data1.Recordset("PE_FRDATE")) Then failed = failed & Data1.Recordset("PE_FRDATE") & ", "
                'If Not IsNull(Data1.Recordset("PE_TODATE")) Then failed = failed & Data1.Recordset("PE_TODATE") & ", "
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
    MsgBox "All Rules applied", vbInformation + vbOKOnly, "Pension Entitlements"
Else
    MsgBox "The Following Rules failed:" & vbCrLf & failed, vbInformation + vbOKOnly, "Pension Entitlements"
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
glbOnTop = "FRMSPENENT" 'Zahoor Butt 01/13/2006
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim X%
Dim SQLQ
FlagRefresh = False
glbOnTop = "FRMSPENENT" 'Zahoor Butt 01/13/2006
Data1.ConnectionString = glbAdoIHRDB
SQLQ = "SELECT DISTINCT PE_DIV,PE_DEPT,PE_ORG,PE_LOC,PE_SECTION,PE_EMP,PE_PT,PE_GRPCD, PE_MANUAL,PE_EDATE "
If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19938 Franks 05/20/2011
    SQLQ = SQLQ & ",PE_USESERVICE,PE_SALDIST "
End If
SQLQ = SQLQ & "FROM HRPENENT "
If glbDIVCount = 1 And glbLinamar Then
    SQLQ = SQLQ & " WHERE PE_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
End If

Data1.RecordSource = SQLQ
Data1.Refresh
Select Case glbCompWDate$ ' sets field reference for basic 'which date'
    Case "O": fglbWDate$ = "ED_DOH"
    Case "S": fglbWDate$ = "ED_SENDTE"
    Case "U": fglbWDate$ = "ED_UNION"
    Case "L": fglbWDate$ = "ED_LTHIRE"
    Case "D": fglbWDate$ = "ED_USRDAT1"
End Select

If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #20575 Franks 07/05/2011
    fglbWDate$ = "ED_OMERS"
    lblHeading(0).Caption = "Service Ranges (in Months)" 'Ticket #21160 Franks 11/17/2011
End If

Screen.MousePointer = HOURGLASS
vbxTrueGrid.Columns(0).Caption = lStr(vbxTrueGrid.Columns(0).Caption)
vbxTrueGrid.Columns(1).Caption = lStr(vbxTrueGrid.Columns(1).Caption)
vbxTrueGrid.Columns(2).Caption = lStr(vbxTrueGrid.Columns(2).Caption)

If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19938 Franks 05/20/2011
    lblAsOf.Visible = True
    dlpAsOf.Visible = True
    chkUseService.Visible = True
    'Ticket #22084 - Franks 05/25/2012
    lblSalDist.Caption = lStr("Salary Distribution")
    lblSalDist.Visible = True
    clpSalDist.Visible = True
Else
    vbxTrueGrid.Columns(9).Visible = False 'no effective date for other customers
End If

Call setRptCaption(Me)
If glbLinamar Then
    lblSection = "Vacation Group"
    clpCode(3).LookupType = SalaryDistribution
    lblSection.FontBold = True
End If
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
textMulti.Caption = "The " & lStr("Union") & " and " & lStr("Category") & " will be validated from the Employee Basic Data"

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
If Me.Height >= 4140 + VacFram.Height + panControls.Height + 230 Then
    scrControl.Value = 0
    VacFram.Top = 4140
    scrControl.Visible = False
    Exit Sub
End If
scrControl.Visible = True
scrControl.Max = VacFram.Height + panControls.Height + 3750 + 550 - Me.Height
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

Private Sub medGTServ_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medLTServ_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPension_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
If IsNumeric(medPension(Index)) Then
    If Len(medPension(Index)) > 0 Then
        medPension(Index) = medPension(Index) * 100
    End If
End If

End Sub

Private Sub medPension_LostFocus(Index As Integer)
If IsNumeric(medPension(Index)) Then
    If Len(medPension(Index)) > 0 Then
        medPension(Index) = medPension(Index) / 100
    End If
End If
End Sub


'-----Daily Vacation Calculation-----------------------------------------------------------
Public Function modDailyUpdateSelection(vacFrom, vacTo, currDate, xAutomatic, Optional seleSQL)
Dim empNo As Long
Dim dblEntitle#, dblPrevEntitle#, strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, X%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%, xAsOf
Dim PenpcN, PenpcO, VED_DIV, VED_PT, SQLQW1
Dim if_Entitle As Boolean, if_Pension As Boolean
Dim xComments
Dim flgOnAnniversary, flgOnJan1, flgStubPeriod, flgWithin10
On Error GoTo modDailyUpdateSelection_Err
modDailyUpdateSelection = False

If xAutomatic = "NO" Then
    If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
    
    Screen.MousePointer = DEFAULT
    If snapEntitle.BOF And snapEntitle.EOF Then
        If fglbRunTimes = 1 Then
            MsgBox "Employees for this selection do not exist!"
            Exit Function
        End If
    Else
        lngRecs& = snapEntitle.RecordCount
        If fglbRunTimes = 1 Then
            Msg$ = lngRecs& & " Records to process" & Chr(10) & "Would You Like To Proceed?"
            Title$ = "Update Entitlements"
            DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
            Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
            If Response% = IDNO Then    ' Evaluate response
                Exit Function
            End If
            Screen.MousePointer = HOURGLASS
        End If
    End If
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 5
    
    For X% = 0 To 24
        If Not IsNumeric(medLTServ(X%)) Then
            medLTServ(X%) = 0
        End If
        If Not IsNumeric(medGTServ(X%)) Then
            medGTServ(X%) = 0
        Else
            If glbFrench Then
                If medGTServ(X%) = Int(medGTServ(X%)) Then medGTServ(X%) = medGTServ(X%) + 0.99
            Else
                If Val(medGTServ(X%)) = Int(medGTServ(X%)) Then medGTServ(X%) = medGTServ(X%) + 0.99
            End If
        End If
        If medLTServ(X%) > 0 And medGTServ(X%) = 0 Then medGTServ(X%) = 9999999
    Next
Else
    'Automatic Entitlement Calculation
    Exit Function
End If

gdbAdoIhr001.BeginTrans

While Not snapEntitle.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / lngRecs&))
    MDIMain.panHelp(0).FloodPercent = pct%
    if_Entitle = False
    if_Pension = False

    empNo& = snapEntitle("ED_EMPNBR")
    
   
    spt = snapEntitle("ED_PT")
    
    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec

    varStartDate = snapEntitle(fglbWDate$)  'Date of Hire - ED_DOH
    
    Dim rsJOB As New ADODB.Recordset
    
    'Mitchell Plastics
    xAsOf = currDate    'Current Date
    
    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    
       
    intWhereFit& = -1

    For X% = 0 To 24
        If medGTServ(X%) > 0 Then
            If dblServiceYears# >= CDbl(medLTServ(X%)) And dblServiceYears# <= CDbl(medGTServ(X%)) Then
                intWhereFit& = X%
                If Len(medPension(X%)) > 0 Then if_Pension = True
                Exit For
            End If
        End If
    Next X%
    
    If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
    
    If if_Pension Then
        PenpcN = medPension(intWhereFit&)
        PenpcO = snapEntitle("ED_VACPC")
        VED_DIV = snapEntitle("ED_DIV")
        VED_PT = snapEntitle("ED_PT")
        If IsNumeric(medPension(intWhereFit&)) Then snapEntitle("ED_CHDSUP") = medPension(intWhereFit&)
    End If
Stub_Cont:
    snapEntitle.Update
    
    If if_Pension Then
        If Val(Format(PenpcN)) <> Val(Format(PenpcO)) Then
'*********************************************************************
            SQLQW1 = "INSERT INTO HRAUDIT (AU_TYPE,AU_NEWEMP,AU_EMPNBR,AU_VACPC,AU_OLDVAC, "
            SQLQW1 = SQLQW1 & "AU_DIVUPL,AU_PTUPL,AU_LDATE,AU_LTIME,AU_UPLOAD,AU_LUSER) "
            
            SQLQW1 = SQLQW1 & " VALUES('M','N'," & empNo& & "," & Val(Format(PenpcN)) & "," & Val(Format(PenpcO))
            SQLQW1 = SQLQW1 & ",'" & VED_DIV & "','" & VED_PT & "', "
            SQLQW1 = SQLQW1 & Date_SQL(Date) & ", '"
            
            SQLQW1 = SQLQW1 & Time$ & "', "
            SQLQW1 = SQLQW1 & "'N', "
            SQLQW1 = SQLQW1 & "'" & glbUserID & "'"
            SQLQW1 = SQLQW1 & ")"
            gdbAdoIhr001X.Execute SQLQW1
        End If
    End If
    Dim xKey
    xKey = snapEntitle("ED_EMPNBR")
    xKey = xKey & "|" & Format(snapEntitle("ED_EFDATE"), "dd-mmm-yyyy")
    xKey = xKey & "|" & Format(snapEntitle("ED_ETDATE"), "dd-mmm-yyyy")
    xKey = xKey & "|VAC"
    xKey = xKey & "|" & dblEntitleUpd
    
    Call Entitlements_Master_Integration(xKey, empNo&) 'George added for Advance Tracker

lblNextRec:
    snapEntitle.MoveNext
    DoEvents
Wend
modDailyUpdateSelection = True
MDIMain.panHelp(0).FloodType = 0
gdbAdoIhr001.CommitTrans

snapEntitle.Close

Screen.MousePointer = DEFAULT

Exit Function

modDailyUpdateSelection_Err:
'These errors are:
'13=type mismatch
'94=invalid use of null
'3018=couln't find field 'item'
If Err = 13 Or Err = 94 Or Err = 3018 Then
    Err = 0
    Resume Next
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateEntitle", "HR_EMP", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If
End Function
'===========================================================================================


Private Function modUpdateSelection(Optional isLast As Boolean)
Dim empNo As Long
Dim dblEntitle#, dblPrevEntitle#, strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, X%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%, xAsOf
Dim PenpcN, PenpcO, VED_DIV, VED_PT, SQLQW1
Dim if_Pension As Boolean
Dim xComments
Dim dblEntitleDays

On Error GoTo modUpdateSelection_Err

modUpdateSelection = False

If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
Screen.MousePointer = DEFAULT
If snapEntitle.BOF And snapEntitle.EOF Then
    'If fglbRunTimes = 1 Then
        MsgBox "Employees for this selection do not exist!"
        modUpdateSelection = True
        Exit Function
    'End If
Else
    lngRecs& = snapEntitle.RecordCount
    'If fglbRunTimes = 1 Then
        Msg$ = lngRecs& & " Records to process" & Chr(10) & "Would You Like To Proceed?"
        Title$ = "Update Entitlements"
        DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
        Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
        If Response% = IDNO Then    ' Evaluate response
            modUpdateSelection = True
            Exit Function
        End If
        Screen.MousePointer = HOURGLASS
    'End If
End If
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 5

For X% = 0 To 24
    If Not IsNumeric(medLTServ(X%)) Then
        medLTServ(X%) = 0
    End If
    If Not IsNumeric(medGTServ(X%)) Then
        medGTServ(X%) = 0
    Else
        If glbFrench Then
            If medGTServ(X%) = Int(medGTServ(X%)) Then medGTServ(X%) = medGTServ(X%) + 0.99
        Else
            If Val(medGTServ(X%)) = Int(medGTServ(X%)) Then medGTServ(X%) = medGTServ(X%) + 0.99
        End If
    End If
    If medLTServ(X%) > 0 And medGTServ(X%) = 0 Then medGTServ(X%) = 9999999
Next

gdbAdoIhr001.BeginTrans

While Not snapEntitle.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / lngRecs&))
    MDIMain.panHelp(0).FloodPercent = pct%
    if_Pension = False

    empNo& = snapEntitle("ED_EMPNBR")
    
  
    spt = snapEntitle("ED_PT")
    
    If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19938 Franks 05/26/2011
        If Not IsNull(snapEntitle("ED_PENPCTFIXED")) Then
            If snapEntitle("ED_PENPCTFIXED") Then
                'If the Pension % on the Banking Information is "Fixed". Don't update it.
                GoTo lblNextRec
            End If
        End If
    End If
    
    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec

    varStartDate = snapEntitle(fglbWDate$)
    
   
    xAsOf = Date
    If dlpAsOf.Visible Then 'Samuel  - Ticket #19938 Franks 05/26/2011
        If IsDate(dlpAsOf.Text) Then
            xAsOf = dlpAsOf.Text
        End If
    End If
'    dblServiceYears# = (DateDiff("d", varStartDate, CVDate(xAsOf)) / 365) * 12
    'dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    If glbSamuel Then 'Ticket #21160 Franks 11/17/2011 - use Months instead of Years
        dblServiceYears# = (DateDiff("d", varStartDate, CVDate(xAsOf)) / 365) * 12
    Else
        dblServiceYears# = DateDiff("d", CVDate(varStartDate), CVDate(xAsOf)) / 365
    End If
    
    intWhereFit& = -1

    For X% = 0 To 24
        If medGTServ(X%) > 0 Then
            If dblServiceYears# >= CDbl(medLTServ(X%)) And dblServiceYears# <= CDbl(medGTServ(X%)) Then
                intWhereFit& = X%
                If Len(medPension(X%)) > 0 Then if_Pension = True
                Exit For
            End If
        End If
    Next X%
    
    If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
    
    ' Two variables glbCompEntVac$ = "M" And glbCompEntSick$ = "M"    are 'company' level
    ' which represents if Sick and Vacation entitlements
    ' are determined on monthly basis (vs yearly) - these are stored in table hrpasco
    ' and read on system startup.
        
    ' In this routine we work independantly of SICK/VACATIon entitlement.
    '  fglbCompMonthly% - is the independant representation
        'of glbCompEntVac$ = "M" And glbCompEntSick$ = "M"
        'Procedure modUpdateSelection is used to set
        'fglbCompMonthly based on values it finds for global variables
        ' and what the user wants to manipulate (sick/Vac)
    
    'optD indicates if Entitlement entered is Daily or yearly based
    ' if daily then max entitlement is based on entitlement * hours they work.
    
    ' we have   Entitle = existing entitmenet (stored presently
    '           NewEntitle = amount entered onto screen = medentitle(index)
    '           EntitleUpd  = value to update record with


    If if_Pension Then
        PenpcN = medPension(intWhereFit&)
        PenpcO = snapEntitle("ED_PENPCT")
        VED_DIV = snapEntitle("ED_DIV")
        VED_PT = snapEntitle("ED_PT")
        If glbFrench Then
            If IsNumeric(medPension(intWhereFit&)) Then snapEntitle("ED_PENPCT") = Replace(medPension(intWhereFit&), ",", ".")
        Else
            If IsNumeric(medPension(intWhereFit&)) Then snapEntitle("ED_PENPCT") = medPension(intWhereFit&)
        End If
        
    End If

    snapEntitle.Update
    
    If if_Pension Then
'****************************************************************************
        SQLQW1 = "INSERT INTO HRAUDIT (AU_TYPE,AU_NEWEMP,AU_EMPNBR,AU_PENPCT,AU_OLDPEN, "
        SQLQW1 = SQLQW1 & "AU_DIVUPL,AU_PTUPL,AU_LDATE,AU_LTIME,AU_UPLOAD,AU_LUSER) "
        
        SQLQW1 = SQLQW1 & " VALUES('M','N'," & empNo& & "," & Val(Format(PenpcN)) & "," & Val(Format(PenpcO))
        SQLQW1 = SQLQW1 & ",'" & VED_DIV & "','" & VED_PT & "', "
        SQLQW1 = SQLQW1 & Date_SQL(Date) & ", '"
        
        SQLQW1 = SQLQW1 & Time$ & "', "
        SQLQW1 = SQLQW1 & "'N', "
        SQLQW1 = SQLQW1 & "'" & glbUserID & "'"
        SQLQW1 = SQLQW1 & ")"
        gdbAdoIhr001X.Execute SQLQW1
    End If
    'commented by Bryan
    'Pension Percent only to be integrated for CGL
'    Dim xKey
'    xKey = snapEntitle("ED_EMPNBR")
'    xKey = xKey & "|" & Format(snapEntitle("ED_EFDATE"), "dd-mmm-yyyy")
'    xKey = xKey & "|" & Format(snapEntitle("ED_ETDATE"), "dd-mmm-yyyy")
'    xKey = xKey & "|PEN"
'    xKey = xKey & "|" & dblEntitleUpd
'
'    Call Entitlements_Master_Integration(xKey, EmpNo&) 'George added for Advance Tracker

lblNextRec:
    snapEntitle.MoveNext
    DoEvents
Wend
modUpdateSelection = True
MDIMain.panHelp(0).FloodType = 0
gdbAdoIhr001.CommitTrans

'fTablHREMP.Close

snapEntitle.Close

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
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateEntitle", "HR_EMP", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If
End Function

Private Sub scrControl_Change()
VacFram.Top = 4140 - scrControl.Value
End Sub
Sub ST_UPD_MODE(TF As Boolean)
Dim X, FT
FT = Not TF
For X = 0 To 24
    medLTServ(X).Enabled = TF
    medGTServ(X).Enabled = TF
    medPension(X).Enabled = TF
Next

clpDiv.Enabled = TF
clpDept.Enabled = TF
clpCode(0).Enabled = TF

If Not glbWHSCC Then
    clpCode(1).Enabled = TF
Else
    clpCode(1).Enabled = False
End If
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF
clpPT.Enabled = TF
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

'vbxTrueGrid.Enabled = FT
'Call modSetFGlobals("Vac")
End Sub


Sub Display_Value()
Dim SQLQ, xOrder, nOrder, aa, SQLQW, glbiOneWhere
Dim rsVE As New ADODB.Recordset
Dim X
For X = 0 To 24
    medLTServ(X) = ""
    medGTServ(X) = ""
    medPension(X) = ""
Next
clpDiv.Text = ""
clpDept.Text = ""
clpCode(0).Text = ""
clpCode(1).Text = ""
clpCode(2).Text = ""
clpCode(3).Text = ""
clpCode(4).Text = ""
clpPT.Text = ""
clpSalDist.Text = ""
dlpAsOf.Text = ""

If Not Data1.Recordset.EOF Then
    SQLQ = "SELECT * FROM HRPENENT "
    If IsNull(Data1.Recordset("PE_DIV")) Then
        SQLQ = SQLQ & " WHERE PE_DIV IS NULL"
    Else
        SQLQ = SQLQ & " WHERE PE_DIV = '" & Data1.Recordset("PE_DIV") & "'"
    End If
    If IsNull(Data1.Recordset("PE_DEPT")) Then
        SQLQ = SQLQ & " AND PE_DEPT IS NULL"
    Else
        SQLQ = SQLQ & " AND PE_DEPT = '" & Data1.Recordset("PE_DEPT") & "'"
    End If
    If IsNull(Data1.Recordset("PE_ORG")) Then
        SQLQ = SQLQ & " AND PE_ORG IS NULL"
    Else
        SQLQ = SQLQ & " AND PE_ORG = '" & Data1.Recordset("PE_ORG") & "'"
    End If
    If IsNull(Data1.Recordset("PE_LOC")) Then
        SQLQ = SQLQ & " AND PE_LOC IS NULL"
    Else
        SQLQ = SQLQ & " AND PE_LOC = '" & Data1.Recordset("PE_LOC") & "'"
    End If
    If IsNull(Data1.Recordset("PE_SECTION")) Then
        SQLQ = SQLQ & " AND PE_SECTION IS NULL"
    Else
        SQLQ = SQLQ & " AND PE_SECTION = '" & Data1.Recordset("PE_SECTION") & "'"
    End If
    If glbCompSerial = "S/N - 2382W" Then    'Ticket #22084 - Franks 05/25/2012
        If IsNull(Data1.Recordset("PE_SALDIST")) Then
            SQLQ = SQLQ & " AND PE_SALDIST IS NULL"
        Else
            SQLQ = SQLQ & " AND PE_SALDIST = '" & Data1.Recordset("PE_SALDIST") & "'"
        End If
    End If
    If IsNull(Data1.Recordset("PE_EMP")) Then
        SQLQ = SQLQ & " AND PE_EMP IS NULL"
    Else
        SQLQ = SQLQ & " AND PE_EMP = '" & Data1.Recordset("PE_EMP") & "'"
    End If
    If IsNull(Data1.Recordset("PE_PT")) Then
        SQLQ = SQLQ & " AND PE_PT IS NULL"
    Else
        SQLQ = SQLQ & " AND PE_PT = '" & Data1.Recordset("PE_PT") & "' "
    End If
    If IsNull(Data1.Recordset("PE_GRPCD")) Then
        SQLQ = SQLQ & " AND PE_GRPCD IS NULL"
    Else
        SQLQ = SQLQ & " AND PE_GRPCD = '" & Data1.Recordset("PE_GRPCD") & "'"
    End If

    
    SQLQ = SQLQ & " Order By PE_DIV,PE_DEPT,PE_ORG, PE_EDATE,PE_EMP,PE_PT,PE_LOC,PE_SECTION,PE_ORDER "
    rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    If Not IsNull(Data1.Recordset("PE_DIV")) Then clpDiv.Text = Data1.Recordset("PE_DIV")
    If Not IsNull(Data1.Recordset("PE_DEPT")) Then clpDept.Text = Data1.Recordset("PE_DEPT")
    If Not IsNull(Data1.Recordset("PE_ORG")) Then clpCode(0).Text = Data1.Recordset("PE_ORG")
    If Not IsNull(Data1.Recordset("PE_EMP")) Then clpCode(1).Text = Data1.Recordset("PE_EMP")
    If Not IsNull(Data1.Recordset("PE_PT")) Then clpPT.Text = Data1.Recordset("PE_PT")
    If Not IsNull(Data1.Recordset("PE_GRPCD")) Then clpCode(2).Text = Data1.Recordset("PE_GRPCD")
    If Not IsNull(Data1.Recordset("PE_LOC")) Then clpCode(4).Text = Data1.Recordset("PE_LOC")
    If Not IsNull(Data1.Recordset("PE_SECTION")) Then clpCode(3).Text = Data1.Recordset("PE_SECTION")
    If Not IsNull(Data1.Recordset("PE_MANUAL")) Then
        chkManual.Value = Data1.Recordset("PE_MANUAL")
    End If
    
    If Not IsNull(Data1.Recordset("PE_EDATE")) Then dlpAsOf.Text = Data1.Recordset("PE_EDATE") 'Ticket #19938
    If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19938 Franks 05/20/2011
        If Not IsNull(Data1.Recordset("PE_USESERVICE")) Then
            chkUseService.Value = Data1.Recordset("PE_USESERVICE")
        Else
            chkUseService.Value = 0
        End If
        'Ticket #22084 - Franks 05/25/2012
        If Not IsNull(Data1.Recordset("PE_SALDIST")) Then clpSalDist.Text = Data1.Recordset("PE_SALDIST")
    End If
    
    Do While Not rsVE.EOF
        xOrder = rsVE("PE_ORDER")
        nOrder = Format(Val(xOrder), "##0") - 1
        If Not (nOrder < 0 Or nOrder > 24) Then
            If Not IsNull(rsVE("PE_BMONTH")) Then medLTServ(nOrder) = rsVE("PE_BMONTH")
            If Not IsNull(rsVE("PE_EMONTH")) Then medGTServ(nOrder) = rsVE("PE_EMONTH")
            If Not IsNull(rsVE("PE_PCT")) Then medPension(nOrder) = rsVE("PE_PCT")
        End If
        rsVE.MoveNext
    Loop
    rsVE.Close
End If
SET_UP_MODE
Call cmdModify_Click
End Sub





Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        'SQLQ = "SELECT DISTINCT PE_DIV,PE_DEPT,PE_ORG,PE_LOC,PE_SECTION,PE_EMP,PE_PT,PE_GRPCD, PE_MANUAL FROM HRPENENT "
        SQLQ = "SELECT DISTINCT PE_DIV,PE_DEPT,PE_ORG,PE_LOC,PE_SECTION,PE_EMP,PE_PT,PE_GRPCD, PE_MANUAL,PE_EDATE "
        If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19938 Franks 05/20/2011
            SQLQ = SQLQ & ",PE_USESERVICE,PE_SALDIST "
        End If
        SQLQ = SQLQ & "FROM HRPENENT "
        If glbDIVCount = 1 And glbLinamar Then
            SQLQ = SQLQ & " WHERE PE_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
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
Dim xLoc, xSection, xSALDIST
Dim xFromDate
Dim xToDate
fglbESQLQ = glbSeleDeptUn
If Len(clpDept.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND  ED_DEPTNO = '" & clpDept.Text & "' "
If Len(clpDiv.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & clpDiv.Text & "' "
If Len(clpCode(0).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & clpCode(0).Text & "' "
If Len(clpCode(1).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP = '" & clpCode(1).Text & "' "
If glbLinamar Then
    If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SALDIST = '" & clpCode(3).Text & "' "
Else
If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & clpCode(3).Text & "' "
End If
If Len(clpCode(4).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & clpCode(4).Text & "' "

If clpPT.Text <> "" Then fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & clpPT.Text & "' "

'Ticket #22084 - Franks 05/25/2012
If clpSalDist.Text <> "" Then fglbESQLQ = fglbESQLQ & " AND ED_SALDIST = '" & clpSalDist.Text & "' "

If xType = "" Then Exit Sub

If xType = "O" Then
    xDiv = ODIV
    xDept = ODept
    xORG = oOrg
    xEMP = oEMP
    xEmpMode = oEmpMode
    xGRPCE = oGRPCE
    xLoc = OLoc
    xSection = OSection
    xSALDIST = OSALDIST
Else
    xDiv = clpDiv.Text
    xDept = clpDept.Text
    xORG = clpCode(0).Text
    xEMP = clpCode(1).Text
    xEmpMode = clpPT.Text
    xGRPCE = clpCode(2).Text
    xLoc = clpCode(4).Text
    xSection = clpCode(3).Text
    xSALDIST = clpSalDist.Text
End If

If Len(xDiv) = 0 Then
    fglbVSQLQ = " (PE_DIV IS NULL OR PE_DIV='')"
Else
    fglbVSQLQ = "PE_DIV = '" & xDiv & "'"
End If
If Len(xDept) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (PE_DEPT IS NULL OR PE_DEPT='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND PE_DEPT = '" & xDept & "'"
End If
If Len(xORG) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (PE_ORG IS NULL OR PE_ORG='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND PE_ORG = '" & xORG & "'"
End If
If Len(xEMP) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (PE_EMP IS NULL OR PE_EMP='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND PE_EMP = '" & xEMP & "'"
End If
If Len(xEmpMode) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (PE_PT IS NULL OR PE_PT='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND PE_PT = '" & xEmpMode & "' "
End If
If Len(xGRPCE) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (PE_GRPCD IS NULL OR PE_GRPCD='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND PE_GRPCD = '" & xGRPCE & "'"
End If

If Len(xLoc) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (PE_LOC IS NULL OR PE_LOC='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND PE_LOC = '" & xLoc & "'"
End If
If Len(xSection) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (PE_SECTION IS NULL OR PE_SECTION='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND PE_SECTION = '" & xSection & "'"
End If

'Ticket #22084 - Franks 05/25/2012
If glbCompSerial = "S/N - 2382W" Then
    If Len(xSALDIST) = 0 Then
        fglbVSQLQ = fglbVSQLQ & " AND (PE_SALDIST IS NULL OR PE_SALDIST='') "
    Else
        fglbVSQLQ = fglbVSQLQ & " AND PE_SALDIST = '" & xSALDIST & "'"
    End If
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
ElseIf Me.Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
    cmdPrintAll.Enabled = False
    cmdUpdate.Enabled = False
Else
    UpdateState = OPENING
    TF = True
    cmdPrintAll.Enabled = True
    cmdUpdate.Enabled = True
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
RelateMode = NothingRelate
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

