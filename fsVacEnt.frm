VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSVacEnt 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Vacation Entitlement Master"
   ClientHeight    =   8490
   ClientLeft      =   2565
   ClientTop       =   525
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
   ScaleHeight     =   8490
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   3825
      LargeChange     =   315
      Left            =   10920
      Max             =   100
      SmallChange     =   315
      TabIndex        =   221
      Top             =   4140
      Width           =   300
   End
   Begin VB.Frame VacFram03 
      BorderStyle     =   0  'None
      Height          =   4185
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11415
      Begin VB.Frame fraSamuelType 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   7680
         TabIndex        =   292
         Top             =   3375
         Visible         =   0   'False
         Width           =   3615
         Begin VB.OptionButton optSamuelType 
            Caption         =   "Service Center "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   294
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton optSamuelType 
            Caption         =   "Non Service Center"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   293
            Top             =   0
            Width           =   1815
         End
      End
      Begin MSMask.MaskEdBox medHours 
         Height          =   285
         Left            =   8400
         TabIndex        =   291
         Tag             =   "10-Usual working hours per day"
         Top             =   2370
         Visible         =   0   'False
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
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
      Begin Threed.SSCheck chkManual 
         Height          =   255
         Left            =   5640
         TabIndex        =   12
         Top             =   3075
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
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
         Left            =   2145
         TabIndex        =   11
         Tag             =   "40-As of Date"
         Top             =   3360
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   6780
         TabIndex        =   7
         Tag             =   "00-Position Group - Code"
         Top             =   2370
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
         Left            =   6780
         TabIndex        =   5
         Tag             =   "00-Specific Employment Status Desired"
         Top             =   1770
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDEM"
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         Height          =   285
         Left            =   6780
         TabIndex        =   6
         Tag             =   "EDPT-Category"
         Top             =   2070
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDPT"
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
         Bindings        =   "fsVacEnt.frx":0000
         Height          =   1335
         Left            =   0
         OleObjectBlob   =   "fsVacEnt.frx":0014
         TabIndex        =   0
         Top             =   120
         Width           =   9135
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   0
         Left            =   2145
         TabIndex        =   9
         Tag             =   "40-From Date"
         Top             =   3060
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1210
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   10
         Tag             =   "40-To Date"
         Top             =   3060
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1210
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   6780
         TabIndex        =   8
         Tag             =   "00-Section - Code"
         Top             =   2670
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
      Begin Threed.SSCheck chkRound 
         Height          =   255
         Left            =   5640
         TabIndex        =   13
         Top             =   3375
         Visible         =   0   'False
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Round entitlement"
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
         TabIndex        =   226
         Top             =   2730
         Width           =   615
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
         Left            =   5280
         TabIndex        =   225
         Top             =   2700
         Width           =   1260
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vacation Pay"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   7920
         TabIndex        =   224
         Top             =   3795
         Width           =   1140
      End
      Begin VB.Label lblPeriod 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vacation Entitlement Period"
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
         TabIndex        =   223
         Top             =   3090
         Width           =   1950
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
         Left            =   5280
         TabIndex        =   222
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
         TabIndex        =   28
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   7995
         TabIndex        =   27
         Top             =   3960
         Width           =   990
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   7080
         TabIndex        =   26
         Top             =   3960
         Width           =   660
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Entitlement"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   25
         Top             =   3960
         Width           =   960
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Ranges (in Months)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   24
         Top             =   3960
         Width           =   2370
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
         Left            =   5280
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   3720
         Visible         =   0   'False
         Width           =   7455
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
         TabIndex        =   21
         Top             =   3360
         Width           =   1245
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
         Left            =   5280
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   1800
         Width           =   555
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   15
      Top             =   7860
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
      Begin VB.CommandButton cmdRecaAccVac 
         Appearance      =   0  'Flat
         Caption         =   "Accrued to Date Vacation Update"
         Height          =   495
         Left            =   10080
         TabIndex        =   251
         Tag             =   "Change all matching records to the above"
         Top             =   120
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.CommandButton cmdYearEnd 
         Appearance      =   0  'Flat
         Caption         =   "&Year End"
         Height          =   375
         Left            =   8040
         TabIndex        =   252
         Tag             =   "Year End by Anniversary Month"
         Top             =   120
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton cmdUpdateAll 
         Caption         =   "Update All"
         Height          =   375
         Left            =   5400
         TabIndex        =   250
         Top             =   120
         Width           =   1665
      End
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update Entitlement"
         Height          =   375
         Left            =   1560
         TabIndex        =   248
         Tag             =   "Change all matching records to the above"
         Top             =   120
         Width           =   1905
      End
      Begin VB.CommandButton CmdRecalc 
         Appearance      =   0  'Flat
         Caption         =   "R&ecalculate"
         Height          =   375
         Left            =   3600
         TabIndex        =   249
         Tag             =   "Recalculation"
         Top             =   120
         Width           =   1665
      End
      Begin VB.CommandButton cmdPrintAll 
         Appearance      =   0  'Flat
         Caption         =   "Print &All"
         Height          =   375
         Left            =   240
         TabIndex        =   247
         Tag             =   "Print all Vacation Entitlement Report"
         Top             =   120
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   405
         Left            =   9720
         Top             =   120
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
      Left            =   60
      TabIndex        =   14
      Top             =   4140
      Width           =   11000
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   29
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
         TabIndex        =   30
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
         TabIndex        =   37
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
         TabIndex        =   38
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
         TabIndex        =   45
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   46
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   0
         Left            =   8000
         TabIndex        =   36
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   1
         Left            =   3270
         TabIndex        =   39
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   2
         Left            =   3270
         TabIndex        =   47
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   0
         Left            =   4300
         TabIndex        =   199
         Top             =   20
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   0
            Left            =   1770
            TabIndex        =   34
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   0
            Left            =   930
            TabIndex        =   33
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   32
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Days"
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
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   1
         Left            =   4300
         TabIndex        =   200
         Top             =   330
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   1
            Left            =   1770
            TabIndex        =   42
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   40
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   1
            Left            =   930
            TabIndex        =   41
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   2
         Left            =   4300
         TabIndex        =   201
         Top             =   660
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   2
            Left            =   1770
            TabIndex        =   50
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   48
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   2
            Left            =   930
            TabIndex        =   49
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   0
         Left            =   7050
         TabIndex        =   35
         Tag             =   "10-Maximum Entitlement can be"
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   1
         Left            =   7050
         TabIndex        =   43
         Tag             =   "10-Maximum Entitlement can be"
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   2
         Left            =   7050
         TabIndex        =   51
         Tag             =   "10-Maximum Entitlement can be"
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   0
         Left            =   3270
         TabIndex        =   31
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   1
         Left            =   8000
         TabIndex        =   44
         Tag             =   "10-Vacation Pay Percentage"
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   2
         Left            =   8000
         TabIndex        =   52
         Tag             =   "10-Vacation Pay Percentage"
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   3
         Left            =   8000
         TabIndex        =   60
         Tag             =   "10-Vacation Pay Percentage"
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   4
         Left            =   8000
         TabIndex        =   68
         Tag             =   "10-Vacation Pay Percentage"
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   5
         Left            =   8000
         TabIndex        =   76
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   1740
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   6
         Left            =   8000
         TabIndex        =   84
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   2055
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   7
         Left            =   8000
         TabIndex        =   92
         Tag             =   "10-Vacation Pay Percentage"
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   8
         Left            =   8000
         TabIndex        =   100
         Tag             =   "10-Vacation Pay Percentage"
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   9
         Left            =   7995
         TabIndex        =   108
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   2995
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
         TabIndex        =   53
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   54
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
         TabIndex        =   61
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   62
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
         TabIndex        =   69
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   70
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
         Left            =   3270
         TabIndex        =   63
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   5
         Left            =   3270
         TabIndex        =   71
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   3
         Left            =   4300
         TabIndex        =   202
         Top             =   990
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   3
            Left            =   1770
            TabIndex        =   58
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   3
            Left            =   930
            TabIndex        =   57
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   225
            Index           =   3
            Left            =   90
            TabIndex        =   56
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Days"
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
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   4
         Left            =   4300
         TabIndex        =   203
         Top             =   1320
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   4
            Left            =   1770
            TabIndex        =   66
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   64
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   4
            Left            =   930
            TabIndex        =   65
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   3
         Left            =   7050
         TabIndex        =   59
         Tag             =   "10-Maximum Entitlement can be"
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   4
         Left            =   7050
         TabIndex        =   67
         Tag             =   "10-Maximum Entitlement can be"
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   5
         Left            =   7050
         TabIndex        =   75
         Tag             =   "10-Maximum Entitlement can be"
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   3
         Left            =   3270
         TabIndex        =   55
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   6
         Left            =   0
         TabIndex        =   77
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   6
         Left            =   2160
         TabIndex        =   78
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
         TabIndex        =   85
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   7
         Left            =   2160
         TabIndex        =   86
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
         TabIndex        =   93
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   8
         Left            =   2160
         TabIndex        =   94
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
         Left            =   3270
         TabIndex        =   87
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   8
         Left            =   3270
         TabIndex        =   95
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   6
         Left            =   7050
         TabIndex        =   83
         Tag             =   "10-Maximum Entitlement can be"
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
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   7
         Left            =   7050
         TabIndex        =   91
         Tag             =   "10-Maximum Entitlement can be"
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
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   8
         Left            =   7050
         TabIndex        =   99
         Tag             =   "10-Maximum Entitlement can be"
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   6
         Left            =   3270
         TabIndex        =   79
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   9
         Left            =   0
         TabIndex        =   101
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   9
         Left            =   2160
         TabIndex        =   102
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
         Left            =   3270
         TabIndex        =   103
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   9
         Left            =   7050
         TabIndex        =   107
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   2995
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   10
         Left            =   8000
         TabIndex        =   116
         Tag             =   "10-Vacation Pay Percentage"
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   11
         Left            =   8000
         TabIndex        =   124
         Tag             =   "10-Vacation Pay Percentage"
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   12
         Left            =   8000
         TabIndex        =   132
         Tag             =   "10-Vacation Pay Percentage"
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   13
         Left            =   8000
         TabIndex        =   140
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   4290
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   14
         Left            =   8000
         TabIndex        =   148
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   4605
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
         TabIndex        =   109
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   10
         Left            =   2160
         TabIndex        =   110
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
         TabIndex        =   117
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   11
         Left            =   2160
         TabIndex        =   118
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
         Left            =   3270
         TabIndex        =   111
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   11
         Left            =   3270
         TabIndex        =   119
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   10
         Left            =   7050
         TabIndex        =   115
         Tag             =   "10-Maximum Entitlement can be"
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   11
         Left            =   7050
         TabIndex        =   123
         Tag             =   "10-Maximum Entitlement can be"
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   12
         Left            =   0
         TabIndex        =   125
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   12
         Left            =   2160
         TabIndex        =   126
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
         TabIndex        =   133
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   13
         Left            =   2160
         TabIndex        =   134
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
         TabIndex        =   141
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   14
         Left            =   2160
         TabIndex        =   142
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
         Left            =   3270
         TabIndex        =   127
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   13
         Left            =   3270
         TabIndex        =   135
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   12
         Left            =   7050
         TabIndex        =   131
         Tag             =   "10-Maximum Entitlement can be"
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
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   13
         Left            =   7050
         TabIndex        =   139
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   4290
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
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   14
         Left            =   7050
         TabIndex        =   147
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   4605
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   14
         Left            =   3270
         TabIndex        =   143
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   15
         Left            =   8000
         TabIndex        =   156
         Tag             =   "10-Vacation Pay Percentage"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   16
         Left            =   7995
         TabIndex        =   164
         Tag             =   "10-Vacation Pay Percentage"
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   15
         Left            =   0
         TabIndex        =   149
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   15
         Left            =   2160
         TabIndex        =   150
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
         TabIndex        =   157
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   16
         Left            =   2160
         TabIndex        =   158
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
         Left            =   3270
         TabIndex        =   151
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   15
         Left            =   7050
         TabIndex        =   155
         Tag             =   "10-Maximum Entitlement can be"
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
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   16
         Left            =   7050
         TabIndex        =   163
         Tag             =   "10-Maximum Entitlement can be"
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   16
         Left            =   3270
         TabIndex        =   159
         Tag             =   "11-Entitlement Amount"
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   17
         Left            =   7995
         TabIndex        =   172
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   5610
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
         TabIndex        =   165
         Tag             =   "11-Service is greater than this number"
         Top             =   5580
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
         TabIndex        =   166
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
         Left            =   3270
         TabIndex        =   167
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   17
         Left            =   7050
         TabIndex        =   171
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   5610
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   18
         Left            =   7995
         TabIndex        =   180
         Tag             =   "10-Vacation Pay Percentage"
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   19
         Left            =   7995
         TabIndex        =   185
         Tag             =   "10-Vacation Pay Percentage"
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   20
         Left            =   7995
         TabIndex        =   190
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   21
         Left            =   7995
         TabIndex        =   195
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   6900
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   22
         Left            =   7995
         TabIndex        =   236
         Tag             =   "10-Vacation Pay Percentage"
         Top             =   7215
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
         TabIndex        =   173
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
         TabIndex        =   174
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
         TabIndex        =   181
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
         TabIndex        =   182
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
         Left            =   3270
         TabIndex        =   175
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   19
         Left            =   3270
         TabIndex        =   183
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   18
         Left            =   7050
         TabIndex        =   179
         Tag             =   "10-Maximum Entitlement can be"
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   19
         Left            =   7050
         TabIndex        =   184
         Tag             =   "10-Maximum Entitlement can be"
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
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   20
         Left            =   0
         TabIndex        =   186
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
         TabIndex        =   187
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
         TabIndex        =   191
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
         TabIndex        =   192
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
         TabIndex        =   196
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
         TabIndex        =   197
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
         Left            =   3270
         TabIndex        =   188
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   21
         Left            =   3270
         TabIndex        =   193
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   20
         Left            =   7050
         TabIndex        =   189
         Tag             =   "10-Maximum Entitlement can be"
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
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   21
         Left            =   7050
         TabIndex        =   194
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   6900
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
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   22
         Left            =   7050
         TabIndex        =   235
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   7215
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   22
         Left            =   3270
         TabIndex        =   198
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   23
         Left            =   7995
         TabIndex        =   241
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
      Begin MSMask.MaskEdBox medVacation 
         Height          =   285
         Index           =   24
         Left            =   7995
         TabIndex        =   246
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   23
         Left            =   0
         TabIndex        =   237
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
         TabIndex        =   238
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
         TabIndex        =   242
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
         TabIndex        =   243
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
         Left            =   3270
         TabIndex        =   239
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   23
         Left            =   7050
         TabIndex        =   240
         Tag             =   "10-Maximum Entitlement can be"
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
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   24
         Left            =   7050
         TabIndex        =   245
         Tag             =   "10-Maximum Entitlement can be"
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   24
         Left            =   3270
         TabIndex        =   244
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   5
         Left            =   4320
         TabIndex        =   253
         Top             =   1650
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   5
            Left            =   1770
            TabIndex        =   74
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   5
            Left            =   930
            TabIndex        =   73
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   225
            Index           =   5
            Left            =   90
            TabIndex        =   72
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Days"
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
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   6
         Left            =   4320
         TabIndex        =   254
         Top             =   1965
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   6
            Left            =   1770
            TabIndex        =   82
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   6
            Left            =   90
            TabIndex        =   80
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   6
            Left            =   930
            TabIndex        =   81
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   7
         Left            =   4320
         TabIndex        =   255
         Top             =   2295
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   7
            Left            =   1770
            TabIndex        =   90
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   7
            Left            =   90
            TabIndex        =   88
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   7
            Left            =   930
            TabIndex        =   89
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   8
         Left            =   4320
         TabIndex        =   256
         Top             =   2625
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   8
            Left            =   1770
            TabIndex        =   98
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   8
            Left            =   930
            TabIndex        =   97
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   225
            Index           =   8
            Left            =   90
            TabIndex        =   96
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Days"
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
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   9
         Left            =   4320
         TabIndex        =   257
         Top             =   2955
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   9
            Left            =   1770
            TabIndex        =   106
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   104
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   9
            Left            =   930
            TabIndex        =   105
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   10
         Left            =   4320
         TabIndex        =   258
         Top             =   3270
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   10
            Left            =   1770
            TabIndex        =   114
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   10
            Left            =   930
            TabIndex        =   113
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   225
            Index           =   10
            Left            =   90
            TabIndex        =   112
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Days"
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
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   11
         Left            =   4320
         TabIndex        =   259
         Top             =   3585
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   11
            Left            =   1770
            TabIndex        =   122
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   11
            Left            =   90
            TabIndex        =   120
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   11
            Left            =   930
            TabIndex        =   121
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   12
         Left            =   4320
         TabIndex        =   260
         Top             =   3915
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   12
            Left            =   1770
            TabIndex        =   130
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   12
            Left            =   90
            TabIndex        =   128
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   12
            Left            =   930
            TabIndex        =   129
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   13
         Left            =   4320
         TabIndex        =   261
         Top             =   4245
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   13
            Left            =   1770
            TabIndex        =   138
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   13
            Left            =   930
            TabIndex        =   137
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   225
            Index           =   13
            Left            =   90
            TabIndex        =   136
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Days"
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
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   14
         Left            =   4320
         TabIndex        =   262
         Top             =   4575
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   14
            Left            =   1770
            TabIndex        =   145
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   14
            Left            =   90
            TabIndex        =   146
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   14
            Left            =   930
            TabIndex        =   144
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   15
         Left            =   4320
         TabIndex        =   263
         Top             =   4890
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   15
            Left            =   1770
            TabIndex        =   154
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   15
            Left            =   930
            TabIndex        =   153
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   225
            Index           =   15
            Left            =   90
            TabIndex        =   152
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Days"
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
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   16
         Left            =   4320
         TabIndex        =   264
         Top             =   5205
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   16
            Left            =   1770
            TabIndex        =   162
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   16
            Left            =   90
            TabIndex        =   160
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   16
            Left            =   930
            TabIndex        =   161
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   17
         Left            =   4320
         TabIndex        =   265
         Top             =   5535
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   17
            Left            =   1770
            TabIndex        =   170
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   17
            Left            =   90
            TabIndex        =   168
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   17
            Left            =   930
            TabIndex        =   169
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   18
         Left            =   4320
         TabIndex        =   266
         Top             =   5865
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   18
            Left            =   1770
            TabIndex        =   178
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   18
            Left            =   930
            TabIndex        =   177
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   225
            Index           =   18
            Left            =   90
            TabIndex        =   176
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Days"
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
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   19
         Left            =   4320
         TabIndex        =   267
         Top             =   6195
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   19
            Left            =   1770
            TabIndex        =   268
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   19
            Left            =   90
            TabIndex        =   269
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   19
            Left            =   930
            TabIndex        =   270
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   20
         Left            =   4320
         TabIndex        =   271
         Top             =   6510
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   20
            Left            =   1770
            TabIndex        =   272
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   20
            Left            =   930
            TabIndex        =   273
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   225
            Index           =   20
            Left            =   90
            TabIndex        =   274
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Days"
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
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   21
         Left            =   4320
         TabIndex        =   275
         Top             =   6825
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   21
            Left            =   1770
            TabIndex        =   276
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   21
            Left            =   90
            TabIndex        =   277
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   21
            Left            =   930
            TabIndex        =   278
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   22
         Left            =   4320
         TabIndex        =   279
         Top             =   7155
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   22
            Left            =   1770
            TabIndex        =   280
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   22
            Left            =   90
            TabIndex        =   281
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   22
            Left            =   930
            TabIndex        =   282
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   23
         Left            =   4320
         TabIndex        =   283
         Top             =   7485
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   23
            Left            =   1770
            TabIndex        =   284
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   23
            Left            =   930
            TabIndex        =   285
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   225
            Index           =   23
            Left            =   90
            TabIndex        =   286
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   397
            _StockProps     =   78
            Caption         =   "Days"
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
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   24
         Left            =   4320
         TabIndex        =   287
         Top             =   7815
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   24
            Left            =   1770
            TabIndex        =   288
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   24
            Left            =   90
            TabIndex        =   289
            Tag             =   "Entitlement measured in days"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   24
            Left            =   930
            TabIndex        =   290
            Tag             =   "Entitlement measured in hours"
            Top             =   120
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
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
         TabIndex        =   234
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
         TabIndex        =   233
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
         TabIndex        =   232
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
         TabIndex        =   231
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
         TabIndex        =   230
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
         TabIndex        =   229
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
         TabIndex        =   228
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
         TabIndex        =   227
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
         TabIndex        =   220
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
         TabIndex        =   219
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
         TabIndex        =   218
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
         TabIndex        =   217
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
         TabIndex        =   216
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
         TabIndex        =   215
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
         TabIndex        =   214
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
         TabIndex        =   213
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
         TabIndex        =   212
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
         TabIndex        =   211
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
         TabIndex        =   210
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
         TabIndex        =   209
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
         TabIndex        =   208
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
         TabIndex        =   207
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
         TabIndex        =   206
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
         TabIndex        =   205
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
         TabIndex        =   204
         Top             =   4920
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmSVacEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fTablHREMP As New ADODB.Recordset         ' table view of HREMP
Dim snapEntitle As New ADODB.Recordset     'user vier
Dim fglbWDate$, fglbWDateS$
Dim fglbEntOSDate$
Dim xService(24, 4)
Dim xTypeD(24)
Dim xTypeH(24)
Dim xTypeF(24)

Dim fglbNoDept&
Dim fglb_FindDept
Dim fglbSick%
Dim fglbVac%
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
Dim OManual
Dim FlagRefresh As Boolean

Dim fglbESQLQ, fglbVSQLQ
Dim fglbNew As Boolean
Dim fglbRunTimes
Dim orgEffDate
Dim xFirstMonEnt 'Ticket #23385 Franks 03/25/2013
Dim isConYear As Boolean 'Ticket #23385 Franks 03/25/2013 - Is it the conversion year

Private Function chkMUEntitle(Optional xOKClick)
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
If Len(dlpDateRange(0).Text) > 0 Then
    If Not IsDate(dlpDateRange(0).Text) Then
        MsgBox "Invalid Vacation Entitlement Period From Date"
        dlpDateRange(0).SetFocus
        Exit Function
    End If
Else
    'If blank then default it as date from company master
    If glbEntOutStanding$ = "1" Then
        dlpDateRange(0).Text = glbCompEdFrom
    End If
End If
If Len(dlpDateRange(1).Text) > 0 Then
    If Not IsDate(dlpDateRange(1).Text) Then
        MsgBox "Invalid Vacation Entitlement Period To Date"
        dlpDateRange(1).SetFocus
        Exit Function
    End If
Else
    'If blank then default it as date from company master
    If glbEntOutStanding$ = "1" Then
        dlpDateRange(1).Text = glbCompEdTo
    End If
End If

If IsDate(dlpDateRange(0).Text) And IsDate(dlpDateRange(1).Text) Then
If CVDate(dlpDateRange(0).Text) > CVDate(dlpDateRange(1).Text) Then
    MsgBox "Vacation Entitlement Period From Date cannot be greater than Vacation Entitlement Period To Date"
    dlpDateRange(0).SetFocus
    Exit Function
End If
End If

If Len(dlpAsOf.Text) > 0 Then
    If Not IsDate(dlpAsOf.Text) Then
        MsgBox "Invalid Effective Date"
        dlpAsOf.SetFocus
        Exit Function
  End If
Else
    If UCase(glbCompEntVac$) = "A" Then
        If glbLinamar Then
            MsgBox "Effective Date is required field"
            dlpAsOf.SetFocus
            Exit Function
        End If
    End If
    If Not glbLinamar Then
        MsgBox "Effective Date is required field"
        dlpAsOf.SetFocus
        Exit Function
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
    If Len(medVacation(x%)) > 0 Then
        If Not IsNumeric(medVacation(x%)) Then
            MsgBox "Data Entered Must Be Numeric"
            medVacation(x%).SetFocus
            Exit Function
        End If
    End If

    If Len(medLTServ(x%)) < 1 And Len(medGTServ(x%)) > 1 Then  ' missed one
        MsgBox "Ranges must be sequential"
        medLTServ(x%).SetFocus
        Exit Function
    End If
    If Len(medGTServ(x%)) > 0 Then
        If glbFrench Then
            If CDbl(medLTServ(x%)) > CDbl(medGTServ(x%)) Then
                MsgBox "Ranges must be sequential"
                medLTServ(x%).SetFocus
                Exit Function
            End If
        Else
            If Val(medLTServ(x%)) > Val(medGTServ(x%)) Then
                MsgBox "Ranges must be sequential"
                medLTServ(x%).SetFocus
                Exit Function
            End If
        End If
    End If
    If x% > 0 And Len(medLTServ(x%)) > 0 Then
        If glbFrench Then
            If CDbl(medLTServ(x%)) < CDbl(medGTServ(x% - 1)) Then
                MsgBox "Ranges must be sequential"
                medLTServ(x%).SetFocus
                Exit Function
            End If
        Else
            If Val(medLTServ(x%)) < Val(medGTServ(x% - 1)) Then
                MsgBox "Ranges must be sequential"
                medLTServ(x%).SetFocus
                Exit Function
            End If
        End If
    End If
    If x% > 0 And Len(medGTServ(x%)) > 0 Then
        If glbFrench Then
            If CDbl(medGTServ(x%)) < CDbl(medGTServ(x% - 1)) And CDbl(medGTServ(x%)) <> 0 Then
                MsgBox "Ranges must be sequential"
                medLTServ(x%).SetFocus
                Exit Function
            End If
        Else
            If Val(medGTServ(x%)) < Val(medGTServ(x% - 1)) And Val(medGTServ(x%)) <> 0 Then
                MsgBox "Ranges must be sequential"
                medLTServ(x%).SetFocus
                Exit Function
            End If
        End If
    End If
    If Len(medLTServ(x%)) > 0 Or Len(medGTServ(x%)) > 0 Then
        If Len(medVacation(x%)) < 1 Then
            If Len(medEntitle(x%)) < 1 Then
                MsgBox "Numeric Value For Entitlement or Vacation Pay Percentage Must Be Entered"
                medEntitle(x%).SetFocus
                Exit Function
            End If
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

For x% = 0 To 24
    If Len(medMax(x%)) < 1 Then
        medMax(x%) = 0
    End If
Next x%

If IsMissing(xOKClick) Then
    If orgEffDate <> dlpAsOf.Text Then
        MsgBox "Effective Date has been changed. Please Save the changes before doing the Update."
        Exit Function
    End If
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

Private Function modAnnSelection(isLast As Boolean)
Dim empNo As Long
Dim dblEntitle#, dblPrevEntitle#, strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%, xAsOf
Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
Dim if_Entitle As Boolean, if_Vacation As Boolean
Dim xComments
Dim dblEntitleDays
Dim xTotEmpHours 'Ticket #21843 Franks 04/12/2012

On Error GoTo modUpdateSelection_Err

modAnnSelection = False


Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 5

For x% = 0 To 24
    If Not IsNumeric(medLTServ(x%)) Then
        medLTServ(x%) = 0
    End If
    If Not IsNumeric(medGTServ(x%)) Then
        medGTServ(x%) = 0
    Else
        If glbFrench Then
            If medGTServ(x%) = Int(medGTServ(x%)) And Val(medGTServ(x%)) > 0 Then medGTServ(x%) = medGTServ(x%) + 0.99
        Else
            If Val(medGTServ(x%)) = Int(medGTServ(x%)) And Val(medGTServ(x%)) > 0 Then medGTServ(x%) = medGTServ(x%) + 0.99
        End If
    End If
    If medLTServ(x%) > 0 And medGTServ(x%) = 0 Then medGTServ(x%) = 9999999
Next

gdbAdoIhr001.BeginTrans

    
    if_Entitle = False
    if_Vacation = False

    empNo& = snapEntitle("ED_EMPNBR")
    
    If IsNull(snapEntitle("ED_ANNVAC")) Then
        dblEntitle# = 0
    Else
        dblEntitle# = snapEntitle("ED_ANNVAC")
    End If
    
    If IsNull(snapEntitle("ED_PVAC")) Then
        dblPrevEntitle# = 0
    Else
        dblPrevEntitle# = snapEntitle("ED_PVAC")
    End If
    
    spt = snapEntitle("ED_PT")
    
    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec

    varStartDate = snapEntitle(fglbWDate$)
    
    Dim rsJOB As New ADODB.Recordset
    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
    dblDHours# = 0
    dblFTEHours# = 0
    If Not rsJOB.EOF Then
        If IsNumeric(rsJOB("JH_DHRS")) Then dblDHours# = rsJOB("JH_DHRS")
        If IsNumeric(rsJOB("JH_FTENUM")) Then dblFTEHours# = rsJOB("JH_FTENUM")
    End If
    'rsJOB.Close    'Ticket #22842 -moved below because of calculating the sum of FTEs for multi positions - Frank forgot to add this logic here
    
    If glbLinamar Then dblDHours# = 8
    
    xAsOf = fglbAsOf
'    dblServiceYears# = (DateDiff("d", varStartDate, CVDate(xAsOf)) / 365) * 12
    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    intWhereFit& = -1

    For x% = 0 To 24
        If medGTServ(x%) > 0 Then
            If dblServiceYears# >= CDbl(medLTServ(x%)) And dblServiceYears# <= CDbl(medGTServ(x%)) Then
                intWhereFit& = x%
                If Len(medEntitle(x%)) > 0 Then if_Entitle = True
                If Len(medVacation(x%)) > 0 Then if_Vacation = True
                Exit For
            End If
        End If
    Next x%
    
    If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
    
    'Ticket #22766 - KidsLink - sum up the FTE for multi positions
    'Ticket #22842 - calculating the sum of FTEs for multi positions - Frank forgot to add this logic here
    If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then  'Kerrys Place Ticket #21843 Franks 04/12/2012, they need the total of hours for multiple current positions
        xTotEmpHours = 0
        Do While Not rsJOB.EOF
            If optD(intWhereFit&) = True Then  ' Entitlements entered in days
                If IsNumeric(rsJOB("JH_DHRS")) Then xTotEmpHours = xTotEmpHours + rsJOB("JH_DHRS")
            End If
            If optF(intWhereFit&) = True Then  ' FTE
                If IsNumeric(rsJOB("JH_DHRS")) And IsNumeric(rsJOB("JH_FTENUM")) Then
                    xTotEmpHours = xTotEmpHours + rsJOB("JH_DHRS") * rsJOB("JH_FTENUM")
                End If
            End If
            rsJOB.MoveNext
        Loop
    End If
    rsJOB.Close
    
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

    If if_Entitle Then
        dblNewEntitle# = medEntitle(intWhereFit&)
        dblNewMax# = 0
        If optD(intWhereFit&) = True Then           ' Entitlements entered in days
            'Ticket #22766 - KidsLink - sum up the FTE for multi positions
            If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then  'Kerrys Place Ticket #21843 Franks 04/12/2012
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * xTotEmpHours
                dblNewEntitle# = dblNewEntitle# * xTotEmpHours
            Else
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblDHours#
                dblNewEntitle# = dblNewEntitle# * dblDHours#
            End If
            dblEntitleUpd = dblNewEntitle
        End If
        If optF(intWhereFit&) = True Then
            'Ticket #22766 - KidsLink - sum up the FTE for multi positions
            If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then  'Kerrys Place Ticket #21843 Franks 04/12/2012
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * xTotEmpHours
                dblNewEntitle# = dblNewEntitle# * xTotEmpHours
            Else
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHours# * dblDHours#
                dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
            End If
            dblEntitleUpd = dblNewEntitle
        End If
        If optH(intWhereFit&) = True Then
            If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&)
        End If
        dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values

         If dblNewMax <> 0 Then          'only do if not zero
            'Ticket #23878 - KidsLink/Carizon - their Calculated will be Annualized Vacation not using Prev.
            If glbCompSerial = "S/N - 2430W" Then
                If dblEntitleUpd > dblNewMax Then
                    dblEntitleUpd = dblNewMax
                End If
            Else
                If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
                    dblEntitleUpd = dblNewMax - dblPrevEntitle#
                End If
            End If
        End If
        
        DtTm = Now
    End If

    If if_Vacation Then
        If glbCBrant And Len(clpCode(3).Text) > 0 And snapEntitle("ED_SECTION") >= clpCode(3).Text Then
            VacpcN = medVacation(intWhereFit&) + dblEntitle#
        Else
            VacpcN = medVacation(intWhereFit&)
        End If
        VacpcO = snapEntitle("ED_VACPC")
        VED_DIV = snapEntitle("ED_DIV")
        VED_PT = snapEntitle("ED_PT")
        If IsNumeric(medVacation(intWhereFit&)) Then snapEntitle("ED_VACPC") = medVacation(intWhereFit&)
        
    End If
    If if_Entitle Then
        
        'If glbCompSerial = "S/N - 2188W" Then  'Ticket #8887
        '    dblEntitleUpd = Round(dblEntitleUpd, 0)
        If glbCompSerial = "S/N - 2297W" Then
            If dblEntitleUpd >= 14.9 And dblEntitleUpd <= 15.1 Then
                dblEntitleUpd = 15
            ElseIf dblEntitleUpd >= 19.9 And dblEntitleUpd <= 20.1 Then
                dblEntitleUpd = 20
            ElseIf dblEntitleUpd >= 25.1 And dblEntitleUpd <= 25.1 Then
                dblEntitleUpd = 25
            End If
        End If
        If glbCBrant And Len(clpCode(3).Text) > 0 Then
            dblEntitleUpd = medVacation(intWhereFit&) + dblEntitleUpd 'dblEntitle#
        End If
                                
       
        If isLast And glbCompSerial = "S/N - 2376W" Then '#9536 on Oct 21,2005 George
            If dblDHours# <> 0 Then
                dblEntitleDays = dblEntitleUpd / dblDHours#
                dblEntitleDays = Round((dblEntitleDays / 0.25 + 0.1), 0) * 0.25 ' round to 1/4 days
                dblEntitleUpd = dblEntitleDays * dblDHours#
            End If
        ElseIf isLast And chkRound.Visible = True And chkRound And chkRound Then
            'Round the final entitlement
            If dblDHours# <> 0 And optH(intWhereFit&) = False Then
                dblEntitleDays = dblEntitleUpd / dblDHours#
                
                If glbCompSerial = "S/N - 2344W" Then   'Ticket #27761 - Cascade Canada Ltd - Round to nearest day
                    'dblEntitleDays = Round((dblEntitleDays + 0.5), 0)
                    dblEntitleDays = Round(dblEntitleDays, 1)
                    dblEntitleDays = Round(dblEntitleDays, 0)
                Else
                    dblEntitleDays = Round(dblEntitleDays, 0)
                End If
                
                dblEntitleUpd = dblEntitleDays * dblDHours#
            Else
                dblEntitleUpd = Round(dblEntitleUpd, 0)
            End If
        End If
        
        'Hemu - 12/31/2003 End
        'Added by bryan 13/Jun/06 Ticket#10916
        snapEntitle("ED_ANNVAC") = dblEntitleUpd
    End If
    snapEntitle.Update
    

lblNextRec:
   
modAnnSelection = True

gdbAdoIhr001.CommitTrans

'fTablHREMP.Close



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


'Private Sub chkYearEnd_Click(Value As Integer)
''Ticket #22893 - Year End for Vacation Entitlement Outstanding Based Upon <> Entitlement Date (1)
'If  glbEntOutStanding$ <> "1" And chkYearEnd Then
'    cmbAnnMonth.Visible = True
'    lblAnnMonth.Visible = True
'
'    Call comAnnMonthAdding
'Else
'    cmbAnnMonth.Visible = False
'    lblAnnMonth.Visible = False
'End If

'End Sub

Private Sub clpCode_LostFocus(Index As Integer)
     If glbWHSCC And Actn = "A" And Index = 0 Then
        If (clpCode(0) = "1866" Or clpCode(0) = "946") And clpPT = "FT" Then
            medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
            medLTServ(1) = 60: medGTServ(1) = 215.99: medEntitle(1) = 1.67
            medLTServ(2) = 216: medGTServ(2) = 999: medEntitle(2) = 2.09
        End If
        If clpCode(0) = "NON" And clpPT = "FT" Then
            optD(0).SetFocus
            medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
            medLTServ(1) = 60: medGTServ(1) = 108.99: medEntitle(1) = 1.67
            medLTServ(2) = 109: medGTServ(2) = 119.99: medEntitle(2) = 21
            medLTServ(3) = 120: medGTServ(3) = 131.99: medEntitle(3) = 22
            medLTServ(4) = 132: medGTServ(4) = 143.99: medEntitle(4) = 23
            medLTServ(5) = 144: medGTServ(5) = 155.99: medEntitle(5) = 24
            medLTServ(6) = 156: medGTServ(6) = 167.99: medEntitle(6) = 25
            medLTServ(7) = 168: medGTServ(7) = 179.99: medEntitle(7) = 26
            medLTServ(8) = 180: medGTServ(8) = 191.99: medEntitle(8) = 27
            medLTServ(9) = 192: medGTServ(9) = 203.99: medEntitle(9) = 28
            medLTServ(10) = 204: medGTServ(10) = 215.99: medEntitle(10) = 29
            medLTServ(11) = 216: medGTServ(11) = 999999.99: medEntitle(11) = 30
        End If
        If clpCode(0) = "PHYS" And clpPT = "FT" Then
            medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
            medLTServ(1) = 60: medGTServ(1) = 119: medEntitle(1) = 1.67
        End If
        If clpCode(0) = "NON" And clpPT = "PT" Then
            optF(0).SetFocus
            medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
            medLTServ(1) = 60: medGTServ(1) = 108.99: medEntitle(1) = 1.67
            medLTServ(2) = 109: medGTServ(2) = 119.99: medEntitle(2) = 21
            medLTServ(3) = 120: medGTServ(3) = 131.99: medEntitle(3) = 22
            medLTServ(4) = 132: medGTServ(4) = 143.99: medEntitle(4) = 23
            medLTServ(5) = 144: medGTServ(5) = 155.99: medEntitle(5) = 24
            medLTServ(6) = 156: medGTServ(6) = 167.99: medEntitle(6) = 25
            medLTServ(7) = 168: medGTServ(7) = 179.99: medEntitle(7) = 26
            medLTServ(8) = 180: medGTServ(8) = 191.99: medEntitle(8) = 27
            medLTServ(9) = 192: medGTServ(9) = 203.99: medEntitle(9) = 28
            medLTServ(10) = 204: medGTServ(10) = 215.99: medEntitle(10) = 29
            medLTServ(11) = 216: medGTServ(11) = 999999.99: medEntitle(11) = 30
        End If
     End If
End Sub

Private Sub clpPT_LostFocus()
     If glbWHSCC And Actn = "A" Then  'And Index = 0 Then
        If (clpCode(0) = "1866" Or clpCode(0) = "946") And clpPT = "FT" Then
            medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
            medLTServ(1) = 60: medGTServ(1) = 215.99: medEntitle(1) = 1.67
            medLTServ(2) = 216: medGTServ(2) = 999: medEntitle(2) = 2.09
        End If
        If clpCode(0) = "NON" And clpPT = "FT" Then
            optD(0).SetFocus
            medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
            medLTServ(1) = 60: medGTServ(1) = 119.99: medEntitle(1) = 1.67
            medLTServ(2) = 120: medGTServ(2) = 131.99: medEntitle(2) = 21
            medLTServ(3) = 132: medGTServ(3) = 143.99: medEntitle(3) = 22
            medLTServ(4) = 144: medGTServ(4) = 155.99: medEntitle(4) = 23
            medLTServ(5) = 156: medGTServ(5) = 167.99: medEntitle(5) = 24
            medLTServ(6) = 168: medGTServ(6) = 179.99: medEntitle(6) = 25
            medLTServ(7) = 180: medGTServ(7) = 191.99: medEntitle(7) = 26
            medLTServ(8) = 192: medGTServ(8) = 203.99: medEntitle(8) = 27
            medLTServ(9) = 204: medGTServ(9) = 215.99: medEntitle(9) = 28
            medLTServ(10) = 216: medGTServ(10) = 227.99: medEntitle(10) = 29
            medLTServ(11) = 228: medGTServ(11) = 999999.99: medEntitle(11) = 30
        End If
        If clpCode(0) = "PHYS" And clpPT = "FT" Then
            medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
            medLTServ(1) = 60: medGTServ(1) = 119: medEntitle(1) = 1.67
        End If
        If clpCode(0) = "NON" And clpPT = "PT" Then
            optF(0).SetFocus
            medLTServ(0) = 0: medGTServ(0) = 59.99: medEntitle(0) = 1.25
            medLTServ(1) = 60: medGTServ(1) = 119.99: medEntitle(1) = 1.67
            medLTServ(2) = 120: medGTServ(2) = 131.99: medEntitle(2) = 21
            medLTServ(3) = 132: medGTServ(3) = 143.99: medEntitle(3) = 22
            medLTServ(4) = 144: medGTServ(4) = 155.99: medEntitle(4) = 23
            medLTServ(5) = 156: medGTServ(5) = 167.99: medEntitle(5) = 24
            medLTServ(6) = 168: medGTServ(6) = 179.99: medEntitle(6) = 25
            medLTServ(7) = 180: medGTServ(7) = 191.99: medEntitle(7) = 26
            medLTServ(8) = 192: medGTServ(8) = 203.99: medEntitle(8) = 27
            medLTServ(9) = 204: medGTServ(9) = 215.99: medEntitle(9) = 28
            medLTServ(10) = 216: medGTServ(10) = 227.99: medEntitle(10) = 29
            medLTServ(11) = 228: medGTServ(11) = 999999.99: medEntitle(11) = 30
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
Msg = Msg & Chr(10) & "The Vacation Entitlement Rules?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

Call getWSQLQ("C")
SQLQ = "DELETE FROM HRVACENT WHERE " & fglbVSQLQ

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

Data1.Refresh
Display_Value

orgEffDate = dlpAsOf.Text

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
    If Not IsNull(Data1.Recordset("VE_EDATE")) Then
        oAsOf = Data1.Recordset("VE_EDATE")
    End If
End If
OLoc = clpCode(4).Text
OSection = clpCode(3).Text
OFromDate = dlpDateRange(0).Text
OToDate = dlpDateRange(1).Text
oEMP = clpCode(1).Text
oEmpMode = clpPT.Text
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    oGRPCE = medHours.Text
Else
    oGRPCE = clpCode(2).Text
End If
OManual = chkManual.Value
orgEffDate = dlpAsOf.Text

Actn = "M"

End Sub

Sub cmdNew_Click()
Dim x

For x = 0 To 24
    medLTServ(x) = ""
    medGTServ(x) = ""
    medEntitle(x) = ""
    optD(x) = True
    optH(x) = False
    optF(x) = False
    medMax(x) = ""
    medVacation(x) = ""
Next
clpDiv.Text = ""
clpDept.Text = ""
clpCode(0).Text = ""
dlpAsOf.Text = ""
clpCode(1).Text = ""
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    medHours.Text = ""
Else
    clpCode(2).Text = ""
End If
clpCode(3).Text = ""
clpCode(4).Text = ""
clpPT.Text = ""
dlpDateRange(0).Text = ""
dlpDateRange(1).Text = ""

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
Dim bmk As Variant
Dim glbiOneWhere As Boolean

On Error GoTo AddN_Err

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    bmk = 0 'Ticket #11885 Frank Oct 11th, 2006
Else
    bmk = Data1.Recordset.Bookmark
End If

If Not chkMUEntitle("OKClick") Then Exit Sub

For x% = 0 To 24
    If Not IsNumeric(medLTServ(x%)) Then Exit For
    If Not IsNumeric(medGTServ(x%)) Then
        medGTServ(x%) = 0
    Else
        If glbFrench Then
            If medGTServ(x%) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
        Else
            If Val(medGTServ(x%)) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
        End If
    End If
    If medLTServ(x%) > 0 And medGTServ(x%) = 0 Then medGTServ(x%) = 9999999
Next

If Actn = "M" Then
    Call getWSQLQ("O")
    SQLQ = "DELETE FROM HRVACENT WHERE " & fglbVSQLQ
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
Else
    Call getWSQLQ("C")
    SQLQ = "SELECT * FROM HRVACENT WHERE " & fglbVSQLQ
    rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsVT.EOF Then
        MsgBox "You can not add duplicate record"
         clpDiv.SetFocus
        Exit Sub
    End If
End If

gdbAdoIhr001.BeginTrans
SQLQ = "SELECT * FROM HRVACENT"
rsVE.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
For x% = 0 To 24
    If Len(medLTServ(x%)) > 0 Then
        rsVE.AddNew
        rsVE("VE_ORDER") = x + 1
        rsVE("VE_ORG_TABL") = "EDOR"
        rsVE("VE_ORG") = clpCode(0).Text
        rsVE("VE_PT") = clpPT.Text
        rsVE("VE_DIV") = clpDiv.Text
        rsVE("VE_DEPT") = clpDept.Text
        rsVE("VE_EMP_TABL") = "EDEM"
        rsVE("VE_EMP") = clpCode(1).Text
        rsVE("VE_SECTION") = clpCode(3).Text
        rsVE("VE_LOC") = clpCode(4).Text
        'commented by Bryan Jan/31/2007 ticket#12467
        'On update all even if monthly every record needs an effective date.
'        If UCase(glbCompEntVac$) = "A" Then
'            If Len(dlpAsOf.Text) > 0 Then
                rsVE("VE_EDATE") = dlpAsOf.Text
'            End If
'        Else
'            rsVE("VE_EDATE") = Null
'        End If
        If Len(dlpDateRange(0).Text) > 0 Then
            rsVE("VE_FRDATE") = dlpDateRange(0).Text
        End If
        If Len(dlpDateRange(1).Text) > 0 Then
            rsVE("VE_TODATE") = dlpDateRange(1).Text
        End If
        rsVE("VE_GRPCD_TABL") = "JBGC"
        If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
            rsVE("VE_GRPCD") = medHours.Text
        Else
            rsVE("VE_GRPCD") = clpCode(2).Text
        End If
        If glbFrench Then
            rsVE("VE_BMONTH") = Replace(medLTServ(x%), ",", ".")
            rsVE("VE_EMONTH") = Replace(medGTServ(x%), ",", ".")
        Else
            rsVE("VE_BMONTH") = medLTServ(x%)
            rsVE("VE_EMONTH") = medGTServ(x%)
        End If
        If medEntitle(x%) = "" Then
            rsVE("VE_ENTITLE") = Null
        Else
            If glbFrench Then
                rsVE("VE_ENTITLE") = Replace(medEntitle(x%), ",", ".")
            Else
                rsVE("VE_ENTITLE") = medEntitle(x%)
            End If
        End If
        If optD(x%) Then rsVE("VE_TYPE") = "D"
        If optH(x%) Then rsVE("VE_TYPE") = "H"
        If optF(x%) Then rsVE("VE_TYPE") = "F"
        If glbFrench Then
            rsVE("VE_MAX") = Replace(medMax(x%), ",", ".")
        Else
            rsVE("VE_MAX") = medMax(x%)
        End If
        If medVacation(x%) = "" Then
            rsVE("VE_PCT") = Null
        Else
            If glbFrench Then
                rsVE("VE_PCT") = Replace(medVacation(x%), ",", ".")
            Else
                rsVE("VE_PCT") = medVacation(x%)
            End If
        End If
        rsVE("VE_MANUAL") = chkManual.Value
        If glbSamuel Then 'Ticket #23385 Franks 03/21/2013
            If optSamuelType(0).Value Or optSamuelType(1).Value Then
                If optSamuelType(0).Value Then rsVE("VE_ROUNDENT") = 1
                If optSamuelType(1).Value Then rsVE("VE_ROUNDENT") = 0
            End If
        End If
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

Display_Value

orgEffDate = dlpAsOf.Text

fglbNew = False

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

If Err.Number = -2147217887 Then '01/01/1200 can cause this error Ticket #18227
    MsgBox "    Invalid Date!    "
    gdbAdoIhr001.RollbackTrans
    Exit Sub
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdOK", "VACATION ENTITLEMENTS", "UPDATE")
    Unload Me
End If

End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport, x%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%
'cmdPrint.Enabled = False

Me.vbxCrystal.Reset
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #13979
    Me.vbxCrystal.WindowTitle = "Current Accrued Pay Period Report"
Else
    Me.vbxCrystal.WindowTitle = "Vacation Entitlement Master Report"
End If
Call setRptLabel(Me, 0) '1)
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To 5
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next
End If
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgvacent.rpt"

SQLQ = "(1=1) "
If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_DIV} = '" & clpDiv.Text & "'"
If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_DEPT} = '" & clpDept.Text & "'"
If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_ORG} = '" & clpCode(0).Text & "'"
If Len(dlpAsOf.Text) > 0 Then
    dtYYY% = Year(dlpAsOf.Text)
    dtMM% = month(dlpAsOf.Text)
    dtDD% = Day(dlpAsOf.Text)
    SQLQ = SQLQ & " AND {HRVACENT.VE_EDATE} = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
End If
If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_EMP} = '" & clpCode(1).Text & "'"
If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_PT} = '" & clpPT.Text & "' "
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    If Len(medHours.Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_GRPCD} = '" & medHours.Text & "'"
Else
    If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_GRPCD} = '" & clpCode(2).Text & "'"
End If
If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_SECTION} = '" & clpCode(3).Text & "'"
If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_LOC} = '" & clpCode(4).Text & "'"
If Len(dlpDateRange(0).Text) > 0 Then
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND {HRVACENT.VE_FRDATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
End If
If Len(dlpDateRange(1).Text) > 0 Then
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    SQLQ = SQLQ & " AND {HRVACENT.VE_TODATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
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
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #13979
    Me.vbxCrystal.WindowTitle = "Current Accrued Pay Period Report"
Else
    Me.vbxCrystal.WindowTitle = "Vacation Entitlement Master Report"
End If
Call setRptLabel(Me, 0) '1)
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To 5
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next
End If
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgvacent.rpt"

SQLQ = "(1=1) "
If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_DIV} = '" & clpDiv.Text & "'"
If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_DEPT} = '" & clpDept.Text & "'"
If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_ORG} = '" & clpCode(0).Text & "'"
If Len(dlpAsOf.Text) > 0 Then
    dtYYY% = Year(dlpAsOf.Text)
    dtMM% = month(dlpAsOf.Text)
    dtDD% = Day(dlpAsOf.Text)
    SQLQ = SQLQ & " AND {HRVACENT.VE_EDATE} = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
End If
If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_EMP} = '" & clpCode(1).Text & "'"
If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_PT} = '" & clpPT.Text & "' "
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    If Len(medHours.Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_GRPCD} = '" & medHours.Text & "'"
Else
    If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_GRPCD} = '" & clpCode(2).Text & "'"
End If
If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_SECTION} = '" & clpCode(3).Text & "'"
If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND {HRVACENT.VE_LOC} = '" & clpCode(4).Text & "'"
If Len(dlpDateRange(0).Text) > 0 Then
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND {HRVACENT.VE_FRDATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
End If
If Len(dlpDateRange(1).Text) > 0 Then
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    SQLQ = SQLQ & " AND {HRVACENT.VE_TODATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
End If


Me.vbxCrystal.SelectionFormula = SQLQ
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
'cmdPrint.Enabled = True
End Sub

Private Sub cmbAnnMonth_GotFocus()
Call SetPanHelp(ActiveControl)
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

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #13979
    Me.vbxCrystal.WindowTitle = "Current Accrued Pay Period Report"
Else
    Me.vbxCrystal.WindowTitle = "Vacation Entitlement Master Report"
End If
Call setRptLabel(Me, 0) '1)
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To 5
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next
End If
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgvacent.rpt"
Me.vbxCrystal.Action = 1

cmdPrintAll.Enabled = True
End Sub

Private Sub cmdRecaAccVac_Click()
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%

    If Not gSec_Upd_Entitlements Then
        MsgBox "You Do Not Have Authority For This Transaction"
        Exit Sub
    End If

    Response% = MsgBox("This function will do Accrued to Date Vacation Update for all employees" & Chr(10) & Chr(10) & "Are you sure you want to proceed with this ?", vbExclamation + vbYesNo, "Update")
    If Response% = IDNO Then
        Exit Sub
    End If
            
    Call Auto_AccruedVacEnt_Upd_DurhamCHC_Run
    
    MsgBox "   Finished.   "
End Sub

Private Sub cmdRecalc_Click()
Dim lastday
Dim flglastdate As Boolean
Dim blIsLast As Boolean
Dim lngRecs As Long, pct As Long, prec As Long
Dim doDate As Date
Dim bmk As Variant

bmk = Data1.Recordset.Bookmark

Screen.MousePointer = HOURGLASS

Call getWSQLQ("C")
Call EntReCalcPeriod(fglbESQLQ, "VAC", dlpDateRange(0), dlpDateRange(1))

    If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
        Data1.Recordset.MoveFirst
        Do
            Call Display_Value
            If Len(dlpAsOf.Text) = 0 Then
                MsgBox "Effective Date is required field"
                dlpAsOf.SetFocus
                Exit Sub
            End If
            If (fglbCompMonthly Or UCase(glbCompEntVac$) = "N") And Not (glbCompSerial = "S\N - 2355W" And chkManual.Value = 0) Then
                prec = 0
                Call getWSQLQ("C")
                
                gdbAdoIhr001.Execute "UPDATE HREMP SET ED_ANNVAC=0 WHERE " & fglbESQLQ
                                
                If Not CR_SnapEntitle() Then Exit Sub  ' create snapEntitle (form level recordset)
                
                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
                    MDIMain.panHelp(0).FloodType = 1
                    
                    While Not snapEntitle.EOF
                        lngRecs = snapEntitle.RecordCount
                        prec = prec + 1
                        pct = Int(100 * (prec / lngRecs))
                        MDIMain.panHelp(0).FloodPercent = pct
                        
                        doDate = dlpAsOf
                        'fglbAsOf = snapEntitle("ED_EFDATE")
                        
                        If IsNull(snapEntitle("ED_EFDATE")) Then GoTo nextEmp
                        
                        fglbAsOf = IsValidDate(Format(month(snapEntitle("ED_EFDATE")) & "/" & Day(dlpAsOf) & "/" & Year(snapEntitle("ED_EFDATE")), "mm/dd/yyyy"), Day(dlpAsOf), month(snapEntitle("ED_EFDATE")), Year(snapEntitle("ED_EFDATE")))
                        For fglbRunTimes = 1 To 12
                            blIsLast = False
                            If fglbRunTimes = 12 Then blIsLast = True
                            
                            If Not modAnnSelection(blIsLast) Then Exit Sub
                            
                            fglbAsOf = DateAdd("m", 1, CVDate(fglbAsOf))
                            
                            DoEvents
                            
                        Next
nextEmp:
                        snapEntitle.MoveNext
                    Wend
                    MDIMain.panHelp(0).FloodType = 0
                End If
            
            Else
                prec = 0
                Call getWSQLQ("C")
                
                gdbAdoIhr001.Execute "UPDATE HREMP SET ED_ANNVAC=0 WHERE " & fglbESQLQ
                
                If Not CR_SnapEntitle() Then Exit Sub  ' create snapEntitle (form level recordset)
                
                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
                    MDIMain.panHelp(0).FloodType = 1
                    
                    While Not snapEntitle.EOF
                        lngRecs = snapEntitle.RecordCount
                        prec = prec + 1
                        pct = Int(100 * (prec / lngRecs))
                        MDIMain.panHelp(0).FloodPercent = pct

                        doDate = dlpAsOf
                        If Not IsNull(snapEntitle("ED_EFDATE")) Then
                            fglbAsOf = snapEntitle("ED_EFDATE")
                            If Not modAnnSelection(True) Then Exit Sub
                        End If
            
                        DoEvents
                            
                        snapEntitle.MoveNext
                    Wend
                    MDIMain.panHelp(0).FloodType = 0
                End If
            
            End If
            Data1.Recordset.MoveNext
        Loop Until Data1.Recordset.EOF
    End If
    
    Data1.Recordset.Bookmark = bmk
    Call Display_Value
    
    Screen.MousePointer = DEFAULT
End Sub

Private Sub cmdUpdate_Click()
Dim sFlag As Boolean
Dim bmk As Variant

bmk = Data1.Recordset.Bookmark
On Error GoTo Mod_Err

If Not gSec_Upd_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

If Not chkMUEntitle() Then Exit Sub

'Ticket #22203 - This is because they are using TAKEN as part of Max checking. So when the date range is
'changed the TAKEN should be recalculated so on Update Entitle, the correct TAKEN is used in the formula.
'During Year End, when the date range is changed, saved and Update Entitlement is clicked, the TAKEN of last
'year is still there in ED_VACT and that was being used in the Max comparison formula. This recalculate
'will fix the issue by recalculating the TAKEN.
If glbCompSerial = "S/N - 2430W" Then
    Call getWSQLQ("C")
    Call EntReCalcPeriod(fglbESQLQ, "VAC", , , dlpDateRange(0), dlpDateRange(1))
    Call EntReCalc(fglbESQLQ)
End If

''Ticket #22893 - Do Year End if selected for employee falling in the Anniversary Month
'If chkYearEnd And cmbAnnMonth.ListIndex <> 0 Then
'    If Not AnniversaryMonth_YearEnd Then GoTo ExitSub
'End If

'added by Bryan 25/Oct/05 Ticket#9560
sFlag = DoWork

Data1.Refresh
Data1.Recordset.Bookmark = bmk

Call Display_Value

orgEffDate = dlpAsOf.Text

If sFlag Then
    MsgBox "Update Completed Successfully.", vbInformation + vbOKOnly, "Vacation Entitlements"
End If

ExitSub:

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
Dim xEmplToIncl As String

CR_SnapEntitle = False

On Error GoTo CR_SnapEntitle_Err

Screen.MousePointer = HOURGLASS

'Ticket #24555 - Kerry's Place
'Custom logic to get list of employees to update with the monthly entitlements
If glbCompSerial = "S/N - 2433W" Then
    xEmplToIncl = KerrysPlace_EmployeesToUpdate
    SQLQ = "SELECT ED_EMPNBR,ED_VACPC,ED_PVAC,ED_VAC,ED_VACT,ED_PSICK,ED_SICK,ED_ANNVAC, ED_ANNSICK, ED_EFDATE,ED_ETDATE,"
    SQLQ = SQLQ & " ED_DIV,ED_PT, ED_SECTION,ED_SALDIST, ED_LOC, ED_ORG, ED_EMP,"
    SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME,ED_VADIM1 "
    SQLQ = SQLQ & " ,ED_SALDIST " 'Ticket #18644
    SQLQ = SQLQ & " FROM HREMP WHERE "
    If Len(xEmplToIncl) > 0 Then
        SQLQ = SQLQ & " ED_EMPNBR IN (" & xEmplToIncl & ")"
    Else
        SQLQ = SQLQ & " 1 = 2"
    End If
Else
    Call getWSQLQ("")
    
    'Realized when update is done, everyone in the selection criteria should get the entitlements. The employees
    'who were part of the Anniversary Month update should be updated too as their routine only rolled over, zero
    'out and changed their entitlement period to new year. So the following has been comment out for that reason.
    'Only employees with Anniversary Month matching user input
    'If cmdYearEnd.Visible = True Then
    '    If Len(glbAnnMonth) > 0 Then fglbESQLQ = fglbESQLQ & " AND MONTH(" & fglbEntOSDate$ & ") = " & glbAnnMonth   'cmbAnnMonth.ListIndex
    '    'Because the Entitlement Period has changed from Rollover and Zero Out to new year
    '    If Len(glbAnnMonth) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EFDATE <= " & Date_SQL(MonthLastDate(dlpAsOf.Text))
    '    'If Len(glbAnnMonth) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EFDATE < " & Date_SQL(dlpAsOf.Text)
    '    'If Len(glbAnnMonth) > 0 Then fglbESQLQ = fglbESQLQ & " AND YEAR(ED_EFDATE) < YEAR(" & Date_SQL(dlpAsOf.Text) & ")"
    'End If
    
    SQLQ = "SELECT ED_EMPNBR,ED_VACPC,ED_PVAC,ED_VAC,ED_VACT,ED_PSICK,ED_SICK,ED_ANNVAC, ED_ANNSICK, ED_EFDATE,ED_ETDATE,"
    SQLQ = SQLQ & " ED_DIV,ED_PT, ED_SECTION,ED_SALDIST, ED_LOC, ED_ORG, ED_EMP,"
    SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME,ED_VADIM1 "
    SQLQ = SQLQ & " ,ED_SALDIST " 'Ticket #18644
    SQLQ = SQLQ & " ,ED_EXTRANN " 'Ticket #27765 Franks 03/01/2016 Durham CHC
    SQLQ = SQLQ & " FROM HREMP WHERE " & fglbESQLQ
End If

If Len(clpCode(2).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_EMPNBR IN "
    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
    SQLQ = SQLQ & " WHERE JB_GRPCD = '" & clpCode(2).Text & "') "
    
    'Ticket #13126 Commented by Frank Jun 5th, 07
    'ElseIf glbCompSerial = "S/N - 2376W" Then 'Assembly of First Nations Bryanm 27/Apr/2006 Ticket#10735
    '    SQLQ = SQLQ & " AND ED_EMPNBR IN "
    '    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
    '    SQLQ = SQLQ & " WHERE JB_GRPCD <> 'MGT')"
End If

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    If Len(medHours.Text) > 0 Then
        SQLQ = SQLQ & " AND ED_EMPNBR IN "
        SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
        SQLQ = SQLQ & " WHERE JH_DHRS = " & medHours.Text & ") "
    End If
End If

'SQLQ = SQLQ & " AND ED_EMPNBR=2005048 " 'FOR TESTING
If snapEntitle.State <> 0 Then snapEntitle.Close
snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic

CR_SnapEntitle = True

Screen.MousePointer = DEFAULT

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

Private Function CR_SnapEntitle_Auto(xSeleSQL)
Dim SQLQ As String
Dim rsVacEnt As New ADODB.Recordset

CR_SnapEntitle_Auto = False
On Error GoTo CR_SnapEntitle_Auto_Err

Screen.MousePointer = HOURGLASS

'Call getWSQLQ("")

SQLQ = "SELECT ED_EMPNBR,ED_VACPC,ED_PVAC,ED_VAC,ED_PSICK,ED_SICK,ED_EFDATE,ED_ETDATE,"
SQLQ = SQLQ & " ED_DIV,ED_PT, ED_SECTION,ED_SALDIST, ED_LOC, ED_ORG, ED_EMP,"
SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME,ED_VACT "
SQLQ = SQLQ & " ,ED_EXTRANN " 'Ticket #27765 Franks 03/01/2016 Durham CHC
SQLQ = SQLQ & " FROM HREMP WHERE " & xSeleSQL

If snapEntitle.State <> 0 Then snapEntitle.Close
snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic

CR_SnapEntitle_Auto = True

Screen.MousePointer = DEFAULT

Exit Function

CR_SnapEntitle_Auto_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapEntitle_Auto", "Entitlements/EMP", "Select")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Public Function Auto_AccruedVacEnt_Upd_DurhamCHC_Run() 'Ticket #27765 Franks 02/29/2016
Dim SQLQ As String
Dim rsVacEnt As New ADODB.Recordset
Dim rsVE As New ADODB.Recordset
Dim selSQLQ As String
Dim xOrder, nOrder As Integer
Dim xFrDate, xEfDate, xLaDate
Dim xMonthNo As Integer
Dim isLast As Boolean
Dim TotMonths As Integer
Dim I As Integer

On Error GoTo Auto_AccruedVacEnt_Upd_DurhamCHC_Run_Err

Screen.MousePointer = HOURGLASS

Auto_AccruedVacEnt_Upd_DurhamCHC_Run = True

    SQLQ = "SELECT DISTINCT VE_DIV,VE_DEPT,VE_ORG,VE_LOC,VE_SECTION,VE_EDATE,VE_EMP,VE_PT,VE_GRPCD,VE_FRDATE,VE_TODATE, VE_MANUAL FROM HRVACENT "
    rsVacEnt.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    
    If Not rsVacEnt.EOF Then
        rsVacEnt.MoveFirst
        
        'For each distinct Vacation Entitlement record
        Do While Not rsVacEnt.EOF
        
            '---------- Selection Criteria ---------------------------------------------------------------------
            selSQLQ = glbSeleDeptUn
            If Len(rsVacEnt("VE_DEPT")) > 0 Then selSQLQ = selSQLQ & " AND  ED_DEPTNO = '" & rsVacEnt("VE_DEPT") & "' "
            If Len(rsVacEnt("VE_DIV")) > 0 Then selSQLQ = selSQLQ & " AND ED_DIV = '" & rsVacEnt("VE_DIV") & "' "
            If Len(rsVacEnt("VE_ORG")) > 0 Then selSQLQ = selSQLQ & " AND ED_ORG = '" & rsVacEnt("VE_ORG") & "' "
            If Len(rsVacEnt("VE_EMP")) > 0 Then selSQLQ = selSQLQ & " AND ED_EMP = '" & rsVacEnt("VE_EMP") & "' "
            If Len(rsVacEnt("VE_SECTION")) > 0 Then selSQLQ = selSQLQ & " AND ED_SECTION = '" & rsVacEnt("VE_SECTION") & "' "
            If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #18235
                If Len(rsVacEnt("VE_LOC")) > 0 Then selSQLQ = selSQLQ & " AND ED_VADIM1 = '" & rsVacEnt("VE_LOC") & "' "
            Else
                If Len(rsVacEnt("VE_LOC")) > 0 Then selSQLQ = selSQLQ & " AND ED_LOC = '" & rsVacEnt("VE_LOC") & "' "
            End If
            If Len(rsVacEnt("VE_PT")) > 0 Then selSQLQ = selSQLQ & " AND ED_PT = '" & rsVacEnt("VE_PT") & "' "
            If Len(rsVacEnt("VE_GRPCD")) > 0 Then
                selSQLQ = selSQLQ & " AND ED_EMPNBR IN "
                selSQLQ = selSQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
                selSQLQ = selSQLQ & " WHERE JB_GRPCD = '" & rsVacEnt("VE_GRPCD") & "') "
            End If
            '---------------------------------------------------------------------------------------------------
        
            'For each distinct Vacation Entitlement record read the details - Service ranges
            SQLQ = "SELECT * FROM HRVACENT "
            If IsNull(rsVacEnt("VE_DIV")) Then
                SQLQ = SQLQ & " WHERE VE_DIV IS NULL"
            Else
                SQLQ = SQLQ & " WHERE VE_DIV = '" & rsVacEnt("VE_DIV") & "'"
            End If
            If IsNull(rsVacEnt("VE_DEPT")) Then
                SQLQ = SQLQ & " AND VE_DEPT IS NULL"
            Else
                SQLQ = SQLQ & " AND VE_DEPT = '" & rsVacEnt("VE_DEPT") & "'"
            End If
            If IsNull(rsVacEnt("VE_ORG")) Then
                SQLQ = SQLQ & " AND VE_ORG IS NULL"
            Else
                SQLQ = SQLQ & " AND VE_ORG = '" & rsVacEnt("VE_ORG") & "'"
            End If
            If IsNull(rsVacEnt("VE_LOC")) Then
                SQLQ = SQLQ & " AND VE_LOC IS NULL"
            Else
                SQLQ = SQLQ & " AND VE_LOC = '" & rsVacEnt("VE_LOC") & "'"
            End If
            If IsNull(rsVacEnt("VE_SECTION")) Then
                SQLQ = SQLQ & " AND VE_SECTION IS NULL"
            Else
                SQLQ = SQLQ & " AND VE_SECTION = '" & rsVacEnt("VE_SECTION") & "'"
            End If
            If Not IsNull(rsVacEnt("VE_EDATE")) Then
                SQLQ = SQLQ & " AND VE_EDATE = " & Date_SQL(rsVacEnt("VE_EDATE"))
            End If
            If IsNull(rsVacEnt("VE_EMP")) Then
                SQLQ = SQLQ & " AND VE_EMP IS NULL"
            Else
                SQLQ = SQLQ & " AND VE_EMP = '" & rsVacEnt("VE_EMP") & "'"
            End If
            If IsNull(rsVacEnt("VE_PT")) Then
                SQLQ = SQLQ & " AND VE_PT IS NULL"
            Else
                SQLQ = SQLQ & " AND VE_PT = '" & rsVacEnt("VE_PT") & "' "
            End If
            If IsNull(rsVacEnt("VE_GRPCD")) Then
                SQLQ = SQLQ & " AND VE_GRPCD IS NULL"
            Else
                SQLQ = SQLQ & " AND VE_GRPCD = '" & rsVacEnt("VE_GRPCD") & "'"
            End If
            If Not IsNull(rsVacEnt("VE_FRDATE")) Then
                SQLQ = SQLQ & " AND VE_FRDATE = " & Date_SQL(rsVacEnt("VE_FRDATE"))
            End If
            If Not IsNull(rsVacEnt("VE_TODATE")) Then
                SQLQ = SQLQ & " AND VE_TODATE = " & Date_SQL(rsVacEnt("VE_TODATE"))
            End If
            
            SQLQ = SQLQ & " ORDER BY VE_DIV,VE_DEPT,VE_ORG, VE_EDATE,VE_EMP,VE_PT,VE_LOC,VE_SECTION,VE_ORDER "
            rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                
            'Get service ranges
            Do While Not rsVE.EOF
                xOrder = rsVE("VE_ORDER")
                nOrder = xOrder - 1
                If Not (nOrder < 0 Or nOrder > 24) Then
                    If Not IsNull(rsVE("VE_BMONTH")) Then xService(nOrder, 0) = rsVE("VE_BMONTH") 'medLTServ(nOrder)
                    If Not IsNull(rsVE("VE_EMONTH")) Then xService(nOrder, 1) = rsVE("VE_EMONTH") 'medGTServ(nOrder)
                    If Not IsNull(rsVE("VE_ENTITLE")) Then xService(nOrder, 2) = rsVE("VE_ENTITLE") 'medEntitle(nOrder)
                    If rsVE("VE_TYPE") = "D" Then xTypeD(nOrder) = True    'optD(nOrder)
                    If rsVE("VE_TYPE") = "H" Then xTypeH(nOrder) = True    'optH(nOrder)
                    If rsVE("VE_TYPE") = "F" Then xTypeF(nOrder) = True    'optF(nOrder)
                    If Not IsNull(rsVE("VE_MAX")) Then xService(nOrder, 3) = rsVE("VE_MAX")
                    If Not IsNull(rsVE("VE_PCT")) Then xService(nOrder, 4) = rsVE("VE_PCT")
                End If
                rsVE.MoveNext   'Next detail (Service range) record
            Loop
            rsVE.Close
            
            ''Call the procedure to calculate the Vacation Entitlement for each distinct Vac Ent. Record
            'Call modDailyUpdateSelection_Auto(rsVacEnt("VE_FRDATE"), rsVacEnt("VE_TODATE"), CVDate(Format(Now, "mm/dd/yyyy")), "YES", selSQLQ)
            
            'loop months
            
            xEfDate = rsVacEnt("VE_EDATE")
            xFrDate = rsVacEnt("VE_FRDATE")
            TotMonths = DateDiff("M", rsVacEnt("VE_FRDATE"), Date) + 1
            xLaDate = DateAdd("M", TotMonths, xEfDate)
            isLast = False
            If TotMonths <= 12 Then
                For I = 1 To TotMonths
                    xMonthNo = I
                    If I = TotMonths Then
                        isLast = True
                    End If
                    If I > 1 Then
                        xEfDate = DateAdd("M", 1, xEfDate)
                    End If
                    Call modDailyUpdAccruedVacDurhamCHC_Auto(rsVacEnt("VE_FRDATE"), rsVacEnt("VE_TODATE"), xEfDate, "YES", selSQLQ, xMonthNo, isLast)
                    'Debug.Print xEfDate
                Next
            End If
            
            rsVacEnt.MoveNext   'Next distinct Vacation Entitlement record
        Loop
    End If
    rsVacEnt.Close
    
Exit Function

Auto_AccruedVacEnt_Upd_DurhamCHC_Run_Err:

    MDIMain.panHelp(0).Caption = "An error occurred in Auto_AccruedVacEnt_Upd_DurhamCHC_Run"
    Auto_AccruedVacEnt_Upd_DurhamCHC_Run = False

End Function

Public Function Automatic_VacEntitlement_Update_Run()
Dim SQLQ As String
Dim rsVacEnt As New ADODB.Recordset
Dim rsVE As New ADODB.Recordset
Dim selSQLQ As String
Dim xOrder, nOrder As Integer

On Error GoTo Automatic_VacEntitlement_Update_Run_Err

Screen.MousePointer = HOURGLASS

Automatic_VacEntitlement_Update_Run = True

    SQLQ = "SELECT DISTINCT VE_DIV,VE_DEPT,VE_ORG,VE_LOC,VE_SECTION,VE_EDATE,VE_EMP,VE_PT,VE_GRPCD,VE_FRDATE,VE_TODATE, VE_MANUAL FROM HRVACENT "
    SQLQ = SQLQ & " WHERE VE_DIV = 'ULT'"
    rsVacEnt.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    
    If Not rsVacEnt.EOF Then
        rsVacEnt.MoveFirst
        
        'For each distinct Vacation Entitlement record
        Do While Not rsVacEnt.EOF
        
            '---------- Selection Criteria ---------------------------------------------------------------------
            selSQLQ = glbSeleDeptUn
            If Len(rsVacEnt("VE_DEPT")) > 0 Then selSQLQ = selSQLQ & " AND  ED_DEPTNO = '" & rsVacEnt("VE_DEPT") & "' "
            If Len(rsVacEnt("VE_DIV")) > 0 Then selSQLQ = selSQLQ & " AND ED_DIV = '" & rsVacEnt("VE_DIV") & "' "
            If Len(rsVacEnt("VE_ORG")) > 0 Then selSQLQ = selSQLQ & " AND ED_ORG = '" & rsVacEnt("VE_ORG") & "' "
            If Len(rsVacEnt("VE_EMP")) > 0 Then selSQLQ = selSQLQ & " AND ED_EMP = '" & rsVacEnt("VE_EMP") & "' "
            If Len(rsVacEnt("VE_SECTION")) > 0 Then selSQLQ = selSQLQ & " AND ED_SECTION = '" & rsVacEnt("VE_SECTION") & "' "
            If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #18235
                If Len(rsVacEnt("VE_LOC")) > 0 Then selSQLQ = selSQLQ & " AND ED_VADIM1 = '" & rsVacEnt("VE_LOC") & "' "
            Else
                If Len(rsVacEnt("VE_LOC")) > 0 Then selSQLQ = selSQLQ & " AND ED_LOC = '" & rsVacEnt("VE_LOC") & "' "
            End If
            If Len(rsVacEnt("VE_PT")) > 0 Then selSQLQ = selSQLQ & " AND ED_PT = '" & rsVacEnt("VE_PT") & "' "
            If Len(rsVacEnt("VE_GRPCD")) > 0 Then
                selSQLQ = selSQLQ & " AND ED_EMPNBR IN "
                selSQLQ = selSQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
                selSQLQ = selSQLQ & " WHERE JB_GRPCD = '" & rsVacEnt("VE_GRPCD") & "') "
            End If
            '---------------------------------------------------------------------------------------------------
        
            'For each distinct Vacation Entitlement record read the details - Service ranges
            SQLQ = "SELECT * FROM HRVACENT "
            If IsNull(rsVacEnt("VE_DIV")) Then
                SQLQ = SQLQ & " WHERE VE_DIV IS NULL"
            Else
                SQLQ = SQLQ & " WHERE VE_DIV = '" & rsVacEnt("VE_DIV") & "'"
            End If
            If IsNull(rsVacEnt("VE_DEPT")) Then
                SQLQ = SQLQ & " AND VE_DEPT IS NULL"
            Else
                SQLQ = SQLQ & " AND VE_DEPT = '" & rsVacEnt("VE_DEPT") & "'"
            End If
            If IsNull(rsVacEnt("VE_ORG")) Then
                SQLQ = SQLQ & " AND VE_ORG IS NULL"
            Else
                SQLQ = SQLQ & " AND VE_ORG = '" & rsVacEnt("VE_ORG") & "'"
            End If
            If IsNull(rsVacEnt("VE_LOC")) Then
                SQLQ = SQLQ & " AND VE_LOC IS NULL"
            Else
                SQLQ = SQLQ & " AND VE_LOC = '" & rsVacEnt("VE_LOC") & "'"
            End If
            If IsNull(rsVacEnt("VE_SECTION")) Then
                SQLQ = SQLQ & " AND VE_SECTION IS NULL"
            Else
                SQLQ = SQLQ & " AND VE_SECTION = '" & rsVacEnt("VE_SECTION") & "'"
            End If
            If Not IsNull(rsVacEnt("VE_EDATE")) Then
                SQLQ = SQLQ & " AND VE_EDATE = " & Date_SQL(rsVacEnt("VE_EDATE"))
            End If
            If IsNull(rsVacEnt("VE_EMP")) Then
                SQLQ = SQLQ & " AND VE_EMP IS NULL"
            Else
                SQLQ = SQLQ & " AND VE_EMP = '" & rsVacEnt("VE_EMP") & "'"
            End If
            If IsNull(rsVacEnt("VE_PT")) Then
                SQLQ = SQLQ & " AND VE_PT IS NULL"
            Else
                SQLQ = SQLQ & " AND VE_PT = '" & rsVacEnt("VE_PT") & "' "
            End If
            If IsNull(rsVacEnt("VE_GRPCD")) Then
                SQLQ = SQLQ & " AND VE_GRPCD IS NULL"
            Else
                SQLQ = SQLQ & " AND VE_GRPCD = '" & rsVacEnt("VE_GRPCD") & "'"
            End If
            If Not IsNull(rsVacEnt("VE_FRDATE")) Then
                SQLQ = SQLQ & " AND VE_FRDATE = " & Date_SQL(rsVacEnt("VE_FRDATE"))
            End If
            If Not IsNull(rsVacEnt("VE_TODATE")) Then
                SQLQ = SQLQ & " AND VE_TODATE = " & Date_SQL(rsVacEnt("VE_TODATE"))
            End If
            
            SQLQ = SQLQ & " ORDER BY VE_DIV,VE_DEPT,VE_ORG, VE_EDATE,VE_EMP,VE_PT,VE_LOC,VE_SECTION,VE_ORDER "
            rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                
            'Get service ranges
            Do While Not rsVE.EOF
                xOrder = rsVE("VE_ORDER")
                nOrder = xOrder - 1
                If Not (nOrder < 0 Or nOrder > 24) Then
                    If Not IsNull(rsVE("VE_BMONTH")) Then xService(nOrder, 0) = rsVE("VE_BMONTH") 'medLTServ(nOrder)
                    If Not IsNull(rsVE("VE_EMONTH")) Then xService(nOrder, 1) = rsVE("VE_EMONTH") 'medGTServ(nOrder)
                    If Not IsNull(rsVE("VE_ENTITLE")) Then xService(nOrder, 2) = rsVE("VE_ENTITLE") 'medEntitle(nOrder)
                    If rsVE("VE_TYPE") = "D" Then xTypeD(nOrder) = True    'optD(nOrder)
                    If rsVE("VE_TYPE") = "H" Then xTypeH(nOrder) = True    'optH(nOrder)
                    If rsVE("VE_TYPE") = "F" Then xTypeF(nOrder) = True    'optF(nOrder)
                    If Not IsNull(rsVE("VE_MAX")) Then xService(nOrder, 3) = rsVE("VE_MAX")
                    If Not IsNull(rsVE("VE_PCT")) Then xService(nOrder, 4) = rsVE("VE_PCT")
                End If
                rsVE.MoveNext   'Next detail (Service range) record
            Loop
            rsVE.Close
            
            'Call the procedure to calculate the Vacation Entitlement for each distinct Vac Ent. Record
            Call modDailyUpdateSelection_Auto(rsVacEnt("VE_FRDATE"), rsVacEnt("VE_TODATE"), CVDate(Format(Now, "mm/dd/yyyy")), "YES", selSQLQ)
            
            rsVacEnt.MoveNext   'Next distinct Vacation Entitlement record
        Loop
    End If
    rsVacEnt.Close
    
Exit Function

Automatic_VacEntitlement_Update_Run_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err

'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Automatic_VacEntitlement_Update_Run", "Entitlements/EMP", "Select")

'If gintRollBack% = False Then
'    Resume Next
'Else
'    Unload Me
'End If
    MDIMain.panHelp(0).Caption = "An error occurred in Automatic_VacEntitlement_Update_Run"
    Automatic_VacEntitlement_Update_Run = False
End Function

Private Sub cmdUpdateAll_Click()
'added by Bryan 25/Oct/05 Ticket#9560
Dim failed As String
Dim c As Long

On Error GoTo Mod_Err
If Not gSec_Upd_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

If orgEffDate <> dlpAsOf.Text Then
    MsgBox "Effective Date has been changed. Please Save the changes before doing the Update."
    Exit Sub
End If

failed = ""
c = 1
If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
    Data1.Recordset.MoveFirst
    Do
        Call Display_Value
        If chkManual.Value = False Then
            If chkMUEntitle() Then
                'Ticket #22203 - This is becsuse they are using TAKEN as part of Max checking. So when the date range is
                'changed the TAKEN should be recalculated so on Update Entitle, the correct TAKEN is used in the formula.
                'During Year End, when the date range is changed, saved and Update Entitlement is clicked, the TAKEN of last
                'year is still there in ED_VACT and that was being used in the Max comparison formula. This recalculate
                'will fix the issue by recalculating the TAKEN.
                If glbCompSerial = "S/N - 2430W" Then
                    Call getWSQLQ("C")
                    Call EntReCalcPeriod(fglbESQLQ, "VAC", , , dlpDateRange(0), dlpDateRange(1))
                    Call EntReCalc(fglbESQLQ)
                End If
            
                If DoWork = False Then
                    failed = failed & "Rule " & CStr(c) & ": "
                    If Not IsNull(Data1.Recordset("VE_DIV")) Then failed = failed & Data1.Recordset("VE_DIV") & ", "
                    If Not IsNull(Data1.Recordset("VE_DEPT")) Then failed = failed & Data1.Recordset("VE_DEPT") & ", "
                    If Not IsNull(Data1.Recordset("VE_ORG")) Then failed = failed & Data1.Recordset("VE_ORG") & ", "
                    If Not IsNull(Data1.Recordset("VE_EDATE")) Then failed = failed & Data1.Recordset("VE_EDATE") & ", "
                    If Not IsNull(Data1.Recordset("VE_EMP")) Then failed = failed & Data1.Recordset("VE_EMP") & ", "
                    If Not IsNull(Data1.Recordset("VE_PT")) Then failed = failed & Data1.Recordset("VE_PT") & ", "
                    If Not IsNull(Data1.Recordset("VE_GRPCD")) Then failed = failed & Data1.Recordset("VE_GRPCD") & ", "
                    If Not IsNull(Data1.Recordset("VE_LOC")) Then failed = failed & Data1.Recordset("VE_LOC") & ", "
                    If Not IsNull(Data1.Recordset("VE_SECTION")) Then failed = failed & Data1.Recordset("VE_SECTION") & ", "
                    If Not IsNull(Data1.Recordset("VE_FRDATE")) Then failed = failed & Data1.Recordset("VE_FRDATE") & ", "
                    If Not IsNull(Data1.Recordset("VE_TODATE")) Then failed = failed & Data1.Recordset("VE_TODATE") & ", "
                    failed = Left(failed, Len(failed) - 2) & vbCrLf
                End If
            Else
                    failed = failed & "Rule " & CStr(c) & ": "
                    If Not IsNull(Data1.Recordset("VE_DIV")) Then failed = failed & Data1.Recordset("VE_DIV") & ", "
                    If Not IsNull(Data1.Recordset("VE_DEPT")) Then failed = failed & Data1.Recordset("VE_DEPT") & ", "
                    If Not IsNull(Data1.Recordset("VE_ORG")) Then failed = failed & Data1.Recordset("VE_ORG") & ", "
                    If Not IsNull(Data1.Recordset("VE_EDATE")) Then failed = failed & Data1.Recordset("VE_EDATE") & ", "
                    If Not IsNull(Data1.Recordset("VE_EMP")) Then failed = failed & Data1.Recordset("VE_EMP") & ", "
                    If Not IsNull(Data1.Recordset("VE_PT")) Then failed = failed & Data1.Recordset("VE_PT") & ", "
                    If Not IsNull(Data1.Recordset("VE_GRPCD")) Then failed = failed & Data1.Recordset("VE_GRPCD") & ", "
                    If Not IsNull(Data1.Recordset("VE_LOC")) Then failed = failed & Data1.Recordset("VE_LOC") & ", "
                    If Not IsNull(Data1.Recordset("VE_SECTION")) Then failed = failed & Data1.Recordset("VE_SECTION") & ", "
                    If Not IsNull(Data1.Recordset("VE_FRDATE")) Then failed = failed & Data1.Recordset("VE_FRDATE") & ", "
                    If Not IsNull(Data1.Recordset("VE_TODATE")) Then failed = failed & Data1.Recordset("VE_TODATE") & ", "
                    failed = Left(failed, Len(failed) - 2) & vbCrLf
            End If
        End If
        c = c + 1
        Data1.Recordset.MoveNext
    Loop Until Data1.Recordset.EOF
End If

Data1.Refresh

Call Display_Value

orgEffDate = dlpAsOf.Text

Screen.MousePointer = DEFAULT

If Len(failed) = 0 Then
    MsgBox "All Rules applied", vbInformation + vbOKOnly, "Vacation Entitlements"
Else
    MsgBox "The Following Rules failed:" & vbCrLf & failed, vbInformation + vbOKOnly, "Vacation Entitlements"
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

Private Sub cmdYearEnd_Click()
'Ticket #22893 - Year End for Vacation Entitlement Outstanding Based Upon <> Entitlement Date (1)
If glbEntOutStanding$ <> "1" Then 'And chkYearEnd Then
    'cmbAnnMonth.Visible = True
    'lblAnnMonth.Visible = True
    
    'Call comAnnMonthAdding
    frmAnnMonth.Show 1
    
    'Ticket #22893 - Do Year End if selected for employee falling in the Anniversary Month
    If glbAnnMonth = 999 Then
        MsgBox "Anniversary Month Year End aborted.", vbInformation, "info:HR - Anniversary Month Year End"
    ElseIf glbAnnMonth <> 0 And glbAnnMonth <> -1 Then
        If Not AnniversaryMonth_YearEnd Then Exit Sub
        Call cmdUpdate_Click
    Else
        MsgBox "Anniversary Month Year End aborted. Year End cannot be performed without selecting Anniversary Month.", vbInformation, "info:HR - Anniversary Month Year End"
    End If
Else
   ' cmbAnnMonth.Visible = False
   'lblAnnMonth.Visible = False
End If
End Sub

Private Sub cmdYearEnd_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
Call INI_Controls(Me)
glbOnTop = "FRMSVACENT"
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim x%
Dim SQLQ

glbOnTop = "FRMSVACENT"

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    clpCode(2).Visible = False
    medHours.Left = 7100
    medHours.Top = clpCode(2).Top
    medHours.Visible = True
    lblCriteria(5).Caption = "Hours/Day"
    vbxTrueGrid.Columns(8).Caption = "Hours/Day"
    Me.Caption = "Current Accrued Pay Period Update"    'Ticket #13979
End If

If glbSamuel Then 'Ticket #23385 Franks 03/21/2013
    SamuelScreenSetup
Else
    If glbCompEntVac$ = "M" Or UCase(glbCompEntVac$) = "N" Then
        chkRound.Visible = True
        chkRound.Value = False
    Else
        chkRound.Value = False
        chkRound.Visible = False
    End If
End If

'Ticket #22893 - Year End for Vacation Entitlement Outstanding Based Upon <> Entitlement Date (1)
'If glbCompSerial = "S/N - 2448W" Then  'For all with Security Right
    If glbEntOutStanding$ <> "1" Then
        'chkYearEnd.Visible = True
        cmdYearEnd.Visible = GetMassUpdateSecurities("YearEnd_AnniversaryMonth_MassUpdate", glbUserID) 'True
        cmdUpdate.Enabled = Not cmdYearEnd.Visible 'True
        cmdUpdateAll.Enabled = Not cmdYearEnd.Visible 'True
    End If
'End If

FlagRefresh = False

Data1.ConnectionString = glbAdoIHRDB

'SQLQ = "SELECT DISTINCT VE_DIV,VE_DEPT,VE_ORG,VE_LOC,VE_SECTION,VE_EDATE,VE_EMP,VE_PT,VE_GRPCD,VE_FRDATE,VE_TODATE, VE_MANUAL FROM HRVACENT "
'Ticket #23385 Franks 03/21/2013
SQLQ = "SELECT DISTINCT VE_DIV,VE_DEPT,VE_ORG,VE_LOC,VE_SECTION,VE_EDATE,VE_EMP,VE_PT,VE_GRPCD,VE_FRDATE,VE_TODATE, VE_MANUAL "
If glbSamuel Then 'Ticket #23385 Franks 03/21/2013
    SQLQ = SQLQ & ",VE_ROUNDENT "
End If
SQLQ = SQLQ & "FROM HRVACENT "
'Ticket #23385 Franks 03/21/2013 - end
If glbDIVCount = 1 And glbLinamar Then
    SQLQ = SQLQ & " WHERE VE_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
End If
If glbWFC Then 'Ticket #28553 Franks 05/03/2016
    SQLQ = SQLQ & " WHERE " & getWFCPlantSecurity("VE_SECTION")
End If
Data1.RecordSource = SQLQ
Data1.Refresh

ODIV = ""
ODept = ""
oOrg = ""
OFromDate = ""
OToDate = ""
oAsOf = ""
oEMP = ""
oEmpMode = ""
oGRPCE = ""
OLoc = ""
OSection = ""
orgEffDate = ""
OManual = False

Select Case glbCompWDate$ ' sets field reference for basic 'which date'
    Case "O": fglbWDate$ = "ED_DOH"
    Case "S": fglbWDate$ = "ED_SENDTE"
    Case "U": fglbWDate$ = "ED_UNION"
    Case "L": fglbWDate$ = "ED_LTHIRE"
    Case "D": fglbWDate$ = "ED_USRDAT1"
End Select
'Ticket #27471 - It should be the 'Vacation Entitlement OS Based Upon' date that the Anniversary Month should be compared with
Select Case glbEntOutStanding$ ' sets field reference for basic 'which date'
    Case "2": fglbEntOSDate$ = "ED_DOH"
    Case "3": fglbEntOSDate$ = "ED_SENDTE"
    Case "4": fglbEntOSDate$ = "ED_LTHIRE"
    Case "5": fglbEntOSDate$ = "ED_USRDAT1"
    Case "6": fglbEntOSDate$ = "ED_UNION"
End Select

If UCase(glbCompEntVac$) = "M" Or UCase(glbCompEntVac$) = "N" Then
    vbxTrueGrid.Columns(5).Visible = False
End If

Screen.MousePointer = HOURGLASS
vbxTrueGrid.Columns(0).Caption = lStr(vbxTrueGrid.Columns(0).Caption)
vbxTrueGrid.Columns(1).Caption = lStr(vbxTrueGrid.Columns(1).Caption)
vbxTrueGrid.Columns(2).Caption = lStr(vbxTrueGrid.Columns(2).Caption)

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

'Ticket #18235 - Location to Vadim 1 - Samuel, Son & Co., Limited
If glbCompSerial = "S/N - 2382W" Then
    lblLocation.Caption = lStr("Vadim Field 1")
    vbxTrueGrid.Columns(9).Caption = lStr("Vadim Field 1")
    clpCode(4).TablName = "EDV1"
    clpCode(4).Tag = "00-Enter Vadim 1 Code"
End If
If glbCompSerial = "S/N - 2396W" Then 'Ticket #27765 Franks 03/01/2016 Durham CHC
    cmdRecaAccVac.Left = 7200
    cmdRecaAccVac.Visible = True
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
scrControl.Left = Me.Width - scrControl.Width - 260
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

Private Sub medVacation_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
If IsNumeric(medVacation(Index)) Then
    If Len(medVacation(Index)) > 0 Then
        medVacation(Index) = medVacation(Index) * 100
    End If
End If

End Sub

Private Sub medVacation_LostFocus(Index As Integer)
If IsNumeric(medVacation(Index)) Then
    If Len(medVacation(Index)) > 0 Then
        medVacation(Index) = medVacation(Index) / 100
    End If
End If
End Sub

Private Sub modMaximums(TF%)
Dim x%

For x% = 0 To 24
    If Not TF Then
        If IsNumeric(medMax(x%)) Then medMax(x%) = 0
    End If
    medMax(x%).Enabled = TF And medMax(x%).Enabled
Next x%

End Sub

'-----Daily Vacation Calculation-----------------------------------------------------------
Public Function modDailyUpdateSelection(vacFrom, vacTo, currDate, xAutomatic, Optional seleSQL)
Dim empNo As Long
Dim dblEntitle#, dblPrevEntitle#, strDivision$
Dim strJob$, dblServiceYears#, dblMonthsDOH
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%, xAsOf
Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
Dim if_Entitle As Boolean, if_Vacation, flgStub As Boolean
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
    
    For x% = 0 To 24
        If Not IsNumeric(medLTServ(x%)) Then
            medLTServ(x%) = 0
        End If
        If Not IsNumeric(medGTServ(x%)) Then
            medGTServ(x%) = 0
        Else
            If glbFrench Then
                If medGTServ(x%) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
            Else
                If Val(medGTServ(x%)) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
            End If
        End If
        If medLTServ(x%) > 0 And medGTServ(x%) = 0 Then medGTServ(x%) = 9999999
    Next
Else
    'Automatic Entitlement Calculation
    Exit Function
End If

'Ticket #11992, Don't use BeginTrans because the Integration is called in the loop
'gdbAdoIhr001.BeginTrans

While Not snapEntitle.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / lngRecs&))
    MDIMain.panHelp(0).FloodPercent = pct%
    if_Entitle = False
    if_Vacation = False

    'Hemu - Ticket #14993
    flgOnJan1 = False
    flgOnAnniversary = False
    flgWithin10 = False
    flgStubPeriod = False

    empNo& = snapEntitle("ED_EMPNBR")
    
    If IsNull(snapEntitle("ED_VAC")) Then
        dblEntitle# = 0
    Else
        dblEntitle# = snapEntitle("ED_VAC")
    End If
    
    If IsNull(snapEntitle("ED_PVAC")) Then
        dblPrevEntitle# = 0
    Else
        dblPrevEntitle# = snapEntitle("ED_PVAC")
    End If
    
    spt = snapEntitle("ED_PT")
    
    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec

    varStartDate = snapEntitle(fglbWDate$)  'Date of Hire - ED_DOH
    
    Dim rsJOB As New ADODB.Recordset
    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
    dblDHours# = 0
    dblFTEHours# = 0
    If Not rsJOB.EOF Then
        If IsNumeric(rsJOB("JH_DHRS")) Then dblDHours# = rsJOB("JH_DHRS")
        If IsNumeric(rsJOB("JH_FTENUM")) Then dblFTEHours# = rsJOB("JH_FTENUM")
    End If
    rsJOB.Close
    If glbLinamar Then dblDHours# = 8
    
    'Mitchell Plastics
    xAsOf = currDate    'Current Date
    
    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    
    If glbCompSerial = "S/N - 2335W" And InStr(1, glbSeleDiv, "HSV") = 0 Then   'Mitchell Plastics
        flgStub = True
        If dblServiceYears# >= 12 Then    '# of Months from the date of hire
            'Find out if the current Vacation Year is the Normal Entitl. start period
            dblMonthsDOH = MonthDiff(CVDate(varStartDate), CVDate(vacFrom))   '# of Months from DOH to Vacation From Date
            If dblMonthsDOH >= 12 Then
                flgStub = False
            Else
                flgStub = True
            End If
        End If
        If dblServiceYears# < 23 And flgStub Then    '# of Months from the date of hire
            'Only on beginning of Vac year
            If CVDate(currDate) = CVDate(vacFrom) Then
                If CVDate(varStartDate) >= CVDate(DateAdd("yyyy", -1, vacFrom)) Then    'Last Vacation Period
                    'Get Entitlement for Stub Period - to be taken within 10 months (Jul 1 - May 1)
                    dblEntitleUpd = Calculate_Stub_Period_Entitlement(snapEntitle("ED_EMPNBR"), CVDate(varStartDate), vacFrom) * dblDHours#
                    
                    flgWithin10 = True
                    
                    if_Entitle = True
                    flgStubPeriod = True
                    flgOnJan1 = False
                    flgOnAnniversary = False
                    GoTo Stub_Cont
                Else
                    flgWithin10 = True
                    if_Entitle = True
                    flgStubPeriod = True
                    flgOnJan1 = False
                    flgOnAnniversary = False
                End If
            Else
                'If Current Date < May 1
                'If CVDate(Format(currDate, "mm/dd/yyyy")) < CVDate(Format("05/01/" & Year(vacTo), "mm/dd/yyyy")) Then
                If DateAdd("m", "10", vacFrom) >= CVDate(Format(currDate, "mm/dd/yyyy")) Then
                    flgWithin10 = True
                Else
                    flgWithin10 = False
                End If
            
                'If an employee had not worked until the Vacation From date but then worked after
                'that then their vacation will 0 then but after working they should have
                'something calculated.
                If snapEntitle("ED_VAC") = 0 Or IsNull(snapEntitle("ED_VAC")) Then
                    If CVDate(varStartDate) >= CVDate(DateAdd("yyyy", -1, vacFrom)) Then    'Last Vacation Period
                        'Get Entitlement for Stub Period - to be taken within 10 months (Jul 1 - May 1)
                        dblEntitleUpd = Calculate_Stub_Period_Entitlement(snapEntitle("ED_EMPNBR"), CVDate(varStartDate), vacFrom) * dblDHours#
                        snapEntitle("ED_VAC") = dblEntitleUpd
                    End If
                Else    'Ticket #15130
                    dblEntitleUpd = Calculate_Stub_Period_Entitlement(snapEntitle("ED_EMPNBR"), CVDate(varStartDate), vacFrom) * dblDHours#
                    snapEntitle("ED_VAC") = dblEntitleUpd
                End If
            
                'Continue with the same entitlement
                dblEntitleUpd = snapEntitle("ED_VAC")
            
                if_Entitle = True
                flgStubPeriod = True
                flgOnJan1 = False
                flgOnAnniversary = False
                GoTo Stub_Cont
            End If
        Else
            flgStubPeriod = False 'Ticket #13051 Frank on May 9, 2007. Reset the flag to false
            'Find out if employee should get extra entitlement on Jan 1st or on Anniversary
            If month(varStartDate) >= 7 And month(varStartDate) <= 12 Then
                flgOnAnniversary = True   'Run SQL procedure to update
                flgOnJan1 = False
                If dblServiceYears# < 24 Then
                    flgOnJan1 = True    'Run SQL procedure to update
                    flgOnAnniversary = False
                End If
            Else
                'Recalculate the service months as of end of the period because they are suppose to get then on Jan/01
                dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(vacTo))
                
                flgOnJan1 = True    'Run SQL procedure to update
                flgOnAnniversary = False
            End If
        End If
    End If
        
    intWhereFit& = -1

    For x% = 0 To 24
        If medGTServ(x%) > 0 Then
            If dblServiceYears# >= CDbl(medLTServ(x%)) And dblServiceYears# <= CDbl(medGTServ(x%)) Then
                intWhereFit& = x%
                If Len(medEntitle(x%)) > 0 Then if_Entitle = True
                If Len(medVacation(x%)) > 0 Then if_Vacation = True
                Exit For
            End If
        End If
    Next x%
    
    If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
    
    If if_Entitle Then
        dblNewEntitle# = medEntitle(intWhereFit&)
        dblNewMax# = 0
        If optD(intWhereFit&) = True Then           ' Entitlements entered in days
            If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblDHours#
            dblNewEntitle# = dblNewEntitle# * dblDHours#
            dblEntitleUpd = dblNewEntitle
        End If
        If optF(intWhereFit&) = True Then
            If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHours# * dblDHours#
            dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
            dblEntitleUpd = dblNewEntitle
        End If
        If optH(intWhereFit&) = True Then
            If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&)
        End If
        
        'Type of Vacation Calculation (Monthly/Annualized Monthly or Annually)
        If fglbCompMonthly Then     'Monthly or Annualized Monthly
            dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
        Else
            dblEntitleUpd = dblNewEntitle ' rollover is on other utility (to accumulate)
        End If
        
         If dblNewMax <> 0 Then          'only do if not zero
            If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
                dblEntitleUpd = dblNewMax - dblPrevEntitle#
            End If
        End If
        
        DtTm = Now
    End If

    If if_Vacation Then
        VacpcN = medVacation(intWhereFit&)
        VacpcO = snapEntitle("ED_VACPC")
        VED_DIV = snapEntitle("ED_DIV")
        VED_PT = snapEntitle("ED_PT")
        If IsNumeric(medVacation(intWhereFit&)) Then snapEntitle("ED_VACPC") = medVacation(intWhereFit&)
        
    End If
Stub_Cont:
    If if_Entitle Then
        
        If glbCompSerial = "S/N - 2188W" Then
            dblEntitleUpd = Round(dblEntitleUpd, 0)
        ElseIf glbCompSerial = "S/N - 2297W" Then
            If dblEntitleUpd >= 14.9 And dblEntitleUpd <= 15.1 Then
                dblEntitleUpd = 15
            ElseIf dblEntitleUpd >= 19.9 And dblEntitleUpd <= 20.1 Then
                dblEntitleUpd = 20
            ElseIf dblEntitleUpd >= 25.1 And dblEntitleUpd <= 25.1 Then
                dblEntitleUpd = 25
            End If
        End If
                        
        'Hemu - 12/31/2003 Begin - Ticket #5348 - City of Chatham-Kent
        If (glbCompSerial = "S/N - 2188W" Or glbCompSerial = "S/N - 2228W") And month(CVDate(xAsOf)) = 12 Then
            snapEntitle("ED_VAC") = Round(dblEntitleUpd, 0)      ' base entitlements sic/vacation
        End If
        'Hemu - 12/31/2003 End
        
        'Mitchell Plastics
        If glbCompSerial = "S/N - 2335W" And InStr(1, glbSeleDiv, "HSV") = 0 Then
            'If Anniversary Today or past then give the entitlement
            'Ticket #15130 - Begin - Asked to change the logic for when to update for Anniversary and Jan 1 update.
            'If flgOnAnniversary And (Month(varStartDate) <= Month(currDate)) And (Day(varStartDate) <= Day(currDate)) Then
            If Not flgStubPeriod Then
                'Ticket #22730
                'xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd
                xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd & ". OS: " & (IIf(IsNull(snapEntitle("ED_PVAC")), 0, snapEntitle("ED_PVAC")) + IIf(IsNull(snapEntitle("ED_VAC")), 0, snapEntitle("ED_VAC"))) - IIf(IsNull(snapEntitle("ED_VACT")), 0, snapEntitle("ED_VACT"))
                
                Call Append_Accrual(empNo&, "VAC", currDate, dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
            
                snapEntitle("ED_VAC") = dblEntitleUpd
                
            'If Anniversary after Dec 31st then give the entitlement
            'ElseIf flgOnJan1 Then
            '    xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd
            '    Call Append_Accrual(EmpNo&, "VAC", currDate, dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
            '
            '    snapEntitle("ED_VAC") = dblEntitleUpd
            'Ticket #15130 - End
            ElseIf flgStubPeriod Then
                'Ticket #22730
                'xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd
                xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd & ". OS: " & (IIf(IsNull(snapEntitle("ED_PVAC")), 0, snapEntitle("ED_PVAC")) + IIf(IsNull(snapEntitle("ED_VAC")), 0, snapEntitle("ED_VAC"))) - IIf(IsNull(snapEntitle("ED_VACT")), 0, snapEntitle("ED_VACT"))

                Call Append_Accrual(empNo&, "VAC", currDate, dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
                
                If flgWithin10 Then
                    snapEntitle("ED_VAC") = dblEntitleUpd
                Else
                    'Hemu - Ticket #14993
                    'snapEntitle("ED_VAC") = 0
                End If
            End If
        End If
                    
        If (glbCompSerial <> "S/N - 2335W") Or (glbCompSerial = "S/N - 2335W" And InStr(1, glbSeleDiv, "HSV") > 0) Then      'Not Mitchell Plastics
            'Ticket #22730
            'xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd
            xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd & ". OS: " & (IIf(IsNull(snapEntitle("ED_PVAC")), 0, snapEntitle("ED_PVAC")) + IIf(IsNull(snapEntitle("ED_VAC")), 0, snapEntitle("ED_VAC"))) - IIf(IsNull(snapEntitle("ED_VACT")), 0, snapEntitle("ED_VACT"))

            Call Append_Accrual(empNo&, "VAC", currDate, dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
            snapEntitle("ED_VAC") = dblEntitleUpd       ' base entitlements sic/vacation
        End If
    End If
    snapEntitle.Update
    
    If if_Vacation Then
        If Val(Format(VacpcN)) <> Val(Format(VacpcO)) Then
        
            SQLQW1 = "INSERT INTO HRAUDIT (AU_TYPE,AU_NEWEMP,AU_EMPNBR,AU_VACPC,AU_OLDVAC, "
            SQLQW1 = SQLQW1 & "AU_DIVUPL,AU_PTUPL,AU_LDATE,AU_LTIME,AU_UPLOAD,AU_LUSER) "
            
            SQLQW1 = SQLQW1 & " VALUES('M','N'," & empNo& & "," & Val(Format(VacpcN)) & "," & Val(Format(VacpcO))
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
    xKey = xKey & "|" & Format(dlpAsOf.Text, "dd-mmm-yyyy") 'Transaction Date
    Call Entitlements_Master_Integration(xKey, empNo&) 'George added for Advance Tracker

lblNextRec:
    snapEntitle.MoveNext
    DoEvents
Wend
modDailyUpdateSelection = True
MDIMain.panHelp(0).FloodType = 0
'gdbAdoIhr001.CommitTrans

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

'-----Daily Accrued to Date Vacation Calculation Auto -------------------------------------------------------
Public Function modDailyUpdAccruedVacDurhamCHC_Auto(vacFrom, vacTo, currDate, xAutomatic, seleSQL, xMonthNo As Integer, isLast As Boolean)
Dim empNo As Long
Dim dblEntitle#, dblPrevEntitle#, strDivision$
Dim strJob$, dblServiceYears#, dblMonthsDOH
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%, xAsOf
Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
Dim if_Entitle As Boolean, if_Vacation, flgStub As Boolean
Dim xComments
Dim flgOnAnniversary, flgOnJan1, flgStubPeriod, flgWithin10

On Error GoTo modDailyUpdAccruedVacDurhamCHC_Auto_Err

modDailyUpdAccruedVacDurhamCHC_Auto = False

'Automatic Entitlement Calculation
If Not CR_SnapEntitle_Auto(seleSQL) Then Exit Function

If snapEntitle.EOF Then Exit Function   'No employee to update

lngRecs& = snapEntitle.RecordCount

Select Case glbCompWDate$ ' sets field reference for basic 'which date'
    Case "O": fglbWDate$ = "ED_DOH"
    Case "S": fglbWDate$ = "ED_SENDTE"
    Case "U": fglbWDate$ = "ED_UNION"
    Case "L": fglbWDate$ = "ED_LTHIRE"
    Case "D": fglbWDate$ = "ED_USRDAT1"
End Select

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 5

For x% = 0 To 24
    If Not IsNumeric(xService(x%, 0)) Then
        xService(x%, 0) = 0
    End If
    If Not IsNumeric(xService(x%, 1)) Then
        xService(x%, 1) = 0
    Else
        If glbFrench Then
            If xService(x%, 1) = Int(xService(x%, 1)) Then xService(x%, 1) = xService(x%, 1) + 0.99
        Else
            If Val(xService(x%, 1)) = Int(xService(x%, 1)) Then xService(x%, 1) = xService(x%, 1) + 0.99
        End If
    End If
    If xService(x%, 0) > 0 And xService(x%, 1) = 0 Then xService(x%, 1) = 9999999
Next


gdbAdoIhr001.BeginTrans

While Not snapEntitle.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / lngRecs&))
    MDIMain.panHelp(0).FloodPercent = pct%
    if_Entitle = False
    if_Vacation = False
    
    'Hemu - Ticket #14993
    flgOnJan1 = False
    flgOnAnniversary = False
    flgWithin10 = False
    flgStubPeriod = False
    
    empNo& = snapEntitle("ED_EMPNBR")
    
    If snapEntitle("ED_EMPNBR") = 5 Then
        If isLast Then
        Debug.Print ""
        End If
    End If
    
    If IsNull(snapEntitle("ED_VAC")) Then
        dblEntitle# = 0
    Else
        dblEntitle# = snapEntitle("ED_VAC")
    End If
    
    If IsNull(snapEntitle("ED_PVAC")) Then
        dblPrevEntitle# = 0
    Else
        dblPrevEntitle# = snapEntitle("ED_PVAC")
    End If
    
    spt = snapEntitle("ED_PT")
    
    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec1

    varStartDate = snapEntitle(fglbWDate$)  'Date of Hire - ED_DOH
    
    Dim rsJOB As New ADODB.Recordset
    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
    dblDHours# = 0
    dblFTEHours# = 0
    If Not rsJOB.EOF Then
        If IsNumeric(rsJOB("JH_DHRS")) Then dblDHours# = rsJOB("JH_DHRS")
        If IsNumeric(rsJOB("JH_FTENUM")) Then dblFTEHours# = rsJOB("JH_FTENUM")
    End If
    rsJOB.Close
    
    'Mitchell Plastics
    xAsOf = currDate    'Current Date
    
    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
        
    intWhereFit& = -1

    For x% = 0 To 24
        If xService(x%, 1) > 0 Then
            If dblServiceYears# >= CDbl(xService(x%, 0)) And dblServiceYears# <= CDbl(xService(x%, 1)) Then
                intWhereFit& = x%
                If Len(xService(x%, 2)) > 0 Then if_Entitle = True
                If Len(xService(x%, 4)) > 0 Then if_Vacation = True
                Exit For
            End If
        End If
    Next x%
    
    If intWhereFit& = -1 Then GoTo lblNextRec1  ' skip record if not in any of the ranges
    
    If if_Entitle Then
        dblNewEntitle# = xService(intWhereFit&, 2)
        dblNewMax# = 0
        If xTypeD(intWhereFit&) = True Then            ' Entitlements entered in days
            If xService(intWhereFit&, 3) <> 0 Then dblNewMax# = xService(intWhereFit&, 3) * dblDHours#
            dblNewEntitle# = dblNewEntitle# * dblDHours#
            dblEntitleUpd = dblNewEntitle
        End If
        If xTypeF(intWhereFit&) = True Then
            If xService(intWhereFit&, 3) <> 0 Then dblNewMax# = xService(intWhereFit&, 3) * dblFTEHours# * dblDHours#
            dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
            dblEntitleUpd = dblNewEntitle
        End If
        If xTypeH(intWhereFit&) = True Then
            If xService(intWhereFit&, 3) <> 0 Then dblNewMax# = xService(intWhereFit&, 3)
        End If
        
        ''Type of Vacation Calculation (Monthly/Annualized Monthly or Annually)
        'If fglbCompMonthly Then     'Monthly or Annualized Monthly
        '    dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
        'Else
        '    dblEntitleUpd = dblNewEntitle ' rollover is on other utility (to accumulate)
        'End If
        dblEntitleUpd = dblNewEntitle
        
        'If dblNewMax <> 0 Then          'only do if not zero
        '    If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
        '        dblEntitleUpd = dblNewMax - dblPrevEntitle#
        '    End If
        'End If
        
        DtTm = Now
    End If

    'If if_Vacation Then
    '    VacpcN = xService(intWhereFit&, 4)
    '    VacpcO = snapEntitle("ED_VACPC")
    '    VED_DIV = snapEntitle("ED_DIV")
    '    VED_PT = snapEntitle("ED_PT")
    '    If IsNumeric(xService(intWhereFit&, 4)) Then snapEntitle("ED_VACPC") = xService(intWhereFit&, 4)
    '
    'End If
Stub_Cont1:
    If if_Entitle Then
        
        dblEntitleUpd = Round((dblEntitleUpd / 12), 4) 'each month vacatiton
        
        If xMonthNo = 1 Then
            snapEntitle("ED_EXTRANN") = dblEntitleUpd
        Else
            snapEntitle("ED_EXTRANN") = snapEntitle("ED_EXTRANN") + dblEntitleUpd
        End If
        
        If isLast Then
            If dblNewMax <> 0 Then          'only do if not zero
                If snapEntitle("ED_EXTRANN") > dblNewMax - dblPrevEntitle# Then
                    'dblEntitleUpd = dblNewMax - dblPrevEntitle#
                    snapEntitle("ED_EXTRANN") = dblNewMax - dblPrevEntitle#
                End If
            End If
        End If
        If xMonthNo = 12 Then
            snapEntitle("ED_EXTRANN") = snapEntitle("ED_VAC")
        End If
    End If
    snapEntitle.Update

lblNextRec1:
    snapEntitle.MoveNext
    DoEvents
Wend
modDailyUpdAccruedVacDurhamCHC_Auto = True
gdbAdoIhr001.CommitTrans

snapEntitle.Close

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodPercent = 0
MDIMain.panHelp(0).FloodType = 0
Screen.MousePointer = DEFAULT


Exit Function

modDailyUpdAccruedVacDurhamCHC_Auto_Err:
'These errors are:
'13=type mismatch
'94=invalid use of null
'3018=couln't find field 'item'
If Err = 13 Or Err = 94 Or Err = 3018 Then
    Err = 0
    Resume Next
End If

Screen.MousePointer = DEFAULT

MDIMain.panHelp(0).Caption = "An error occurred in Vacation Entitlement Calculation procedure"

End Function

'-----Daily Vacation Calculation Auto -------------------------------------------------------
Public Function modDailyUpdateSelection_Auto(vacFrom, vacTo, currDate, xAutomatic, Optional seleSQL)
Dim empNo As Long
Dim dblEntitle#, dblPrevEntitle#, strDivision$
Dim strJob$, dblServiceYears#, dblMonthsDOH
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%, xAsOf
Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
Dim if_Entitle As Boolean, if_Vacation, flgStub As Boolean
Dim xComments
Dim flgOnAnniversary, flgOnJan1, flgStubPeriod, flgWithin10

On Error GoTo modDailyUpdateSelection_Auto_Err

modDailyUpdateSelection_Auto = False

If xAutomatic = "NO" Then
    Screen.MousePointer = DEFAULT
    
    Exit Function  ' create snapEntitle (form level recordset)
Else
    'Automatic Entitlement Calculation
    If Not CR_SnapEntitle_Auto(seleSQL) Then Exit Function

    If snapEntitle.EOF Then Exit Function   'No employee to update
    
    lngRecs& = snapEntitle.RecordCount
    
    Select Case glbCompWDate$ ' sets field reference for basic 'which date'
        Case "O": fglbWDate$ = "ED_DOH"
        Case "S": fglbWDate$ = "ED_SENDTE"
        Case "U": fglbWDate$ = "ED_UNION"
        Case "L": fglbWDate$ = "ED_LTHIRE"
        Case "D": fglbWDate$ = "ED_USRDAT1"
    End Select

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 5

    For x% = 0 To 24
        If Not IsNumeric(xService(x%, 0)) Then
            xService(x%, 0) = 0
        End If
        If Not IsNumeric(xService(x%, 1)) Then
            xService(x%, 1) = 0
        Else
            If glbFrench Then
                If xService(x%, 1) = Int(xService(x%, 1)) Then xService(x%, 1) = xService(x%, 1) + 0.99
            Else
                If Val(xService(x%, 1)) = Int(xService(x%, 1)) Then xService(x%, 1) = xService(x%, 1) + 0.99
            End If
        End If
        If xService(x%, 0) > 0 And xService(x%, 1) = 0 Then xService(x%, 1) = 9999999
    Next

End If

gdbAdoIhr001.BeginTrans

While Not snapEntitle.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / lngRecs&))
    MDIMain.panHelp(0).FloodPercent = pct%
    if_Entitle = False
    if_Vacation = False
    
    'Hemu - Ticket #14993
    flgOnJan1 = False
    flgOnAnniversary = False
    flgWithin10 = False
    flgStubPeriod = False
    
    empNo& = snapEntitle("ED_EMPNBR")
    
    'If snapEntitle("ED_EMPNBR") = 3570350 Then
    '    MsgBox "3570350"
    'End If
    
    If IsNull(snapEntitle("ED_VAC")) Then
        dblEntitle# = 0
    Else
        dblEntitle# = snapEntitle("ED_VAC")
    End If
    
    If IsNull(snapEntitle("ED_PVAC")) Then
        dblPrevEntitle# = 0
    Else
        dblPrevEntitle# = snapEntitle("ED_PVAC")
    End If
    
    spt = snapEntitle("ED_PT")
    
    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec1

    varStartDate = snapEntitle(fglbWDate$)  'Date of Hire - ED_DOH
    
    Dim rsJOB As New ADODB.Recordset
    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
    dblDHours# = 0
    dblFTEHours# = 0
    If Not rsJOB.EOF Then
        If IsNumeric(rsJOB("JH_DHRS")) Then dblDHours# = rsJOB("JH_DHRS")
        If IsNumeric(rsJOB("JH_FTENUM")) Then dblFTEHours# = rsJOB("JH_FTENUM")
    End If
    rsJOB.Close
    If glbLinamar Then dblDHours# = 8
    
    'Mitchell Plastics
    xAsOf = currDate    'Current Date
    
    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    
    If glbCompSerial = "S/N - 2335W" And InStr(1, glbSeleDiv, "HSV") = 0 Then   'Mitchell Plastics
        flgStub = True
        If dblServiceYears# >= 12 Then    '# of Months from the date of hire
            'Find out if the current Vacation Year is the Normal Entitl. start period
            dblMonthsDOH = MonthDiff(CVDate(varStartDate), CVDate(vacFrom))   '# of Months from DOH to Vacation From Date
            If dblMonthsDOH >= 12 Then
                flgStub = False
            Else
                flgStub = True
            End If
        End If
        If dblServiceYears# < 23 And flgStub Then   '# of Months from the date of hire
            'Only on beginning of Vac year
            If CVDate(currDate) = CVDate(vacFrom) Then
                If CVDate(varStartDate) >= CVDate(DateAdd("yyyy", -1, vacFrom)) Then    'Last Vacation Period
                    'Get Entitlement for Stub Period - to be taken within 10 months (Jul 1 - May 1)
                    dblEntitleUpd = Calculate_Stub_Period_Entitlement(snapEntitle("ED_EMPNBR"), CVDate(varStartDate), vacFrom) * dblDHours#
                    
                    flgWithin10 = True
                    
                    if_Entitle = True
                    flgStubPeriod = True
                    flgOnJan1 = False
                    flgOnAnniversary = False
                    GoTo Stub_Cont1
                Else
                    flgWithin10 = True
                    if_Entitle = True
                    flgStubPeriod = True
                    flgOnJan1 = False
                    flgOnAnniversary = False
                End If
            Else
                'If Current Date < May 1
                'If CVDate(Format(currDate, "mm/dd/yyyy")) < CVDate(Format("05/01/" & Year(vacTo), "mm/dd/yyyy")) Then
                If DateAdd("m", "10", vacFrom) >= CVDate(Format(currDate, "mm/dd/yyyy")) Then
                    flgWithin10 = True
                Else
                    flgWithin10 = False
                End If
            
                'If an employee had not worked until the Vacation From date but then worked after
                'that then their vacation will 0 then but after working they should have
                'something calculated.
                If snapEntitle("ED_VAC") = 0 Or IsNull(snapEntitle("ED_VAC")) Then
                    If CVDate(varStartDate) >= CVDate(DateAdd("yyyy", -1, vacFrom)) Then    'Last Vacation Period
                        'Get Entitlement for Stub Period - to be taken within 10 months (Jul 1 - May 1)
                        dblEntitleUpd = Calculate_Stub_Period_Entitlement(snapEntitle("ED_EMPNBR"), CVDate(varStartDate), vacFrom) * dblDHours#
                        snapEntitle("ED_VAC") = dblEntitleUpd
                    End If
                Else    'Ticket #15130
                    dblEntitleUpd = Calculate_Stub_Period_Entitlement(snapEntitle("ED_EMPNBR"), CVDate(varStartDate), vacFrom) * dblDHours#
                    snapEntitle("ED_VAC") = dblEntitleUpd
                End If
            
                'Continue with the same entitlement
                dblEntitleUpd = snapEntitle("ED_VAC")
            
                if_Entitle = True
                flgStubPeriod = True
                flgOnJan1 = False
                flgOnAnniversary = False
                GoTo Stub_Cont1
            End If
        Else
            'Find out if employee should get extra entitlement on Jan 1st or on Anniversary
            If month(varStartDate) >= 7 And month(varStartDate) <= 12 Then
                flgOnAnniversary = True   'Run SQL procedure to update
                flgOnJan1 = False
                If dblServiceYears# < 24 Then
                    flgOnJan1 = True    'Run SQL procedure to update
                    flgOnAnniversary = False
                End If
            Else
                'Recalculate the service months as of end of the period because they are suppose to get then on Jan/01
                dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(vacTo))
                
                flgOnJan1 = True    'Run SQL procedure to update
                flgOnAnniversary = False
            End If
        End If
    End If
        
    intWhereFit& = -1

    For x% = 0 To 24
        If xService(x%, 1) > 0 Then
            If dblServiceYears# >= CDbl(xService(x%, 0)) And dblServiceYears# <= CDbl(xService(x%, 1)) Then
                intWhereFit& = x%
                If Len(xService(x%, 2)) > 0 Then if_Entitle = True
                If Len(xService(x%, 4)) > 0 Then if_Vacation = True
                Exit For
            End If
        End If
    Next x%
    
    If intWhereFit& = -1 Then GoTo lblNextRec1  ' skip record if not in any of the ranges
    
    If if_Entitle Then
        dblNewEntitle# = xService(intWhereFit&, 2)
        dblNewMax# = 0
        If xTypeD(intWhereFit&) = True Then            ' Entitlements entered in days
            If xService(intWhereFit&, 3) <> 0 Then dblNewMax# = xService(intWhereFit&, 3) * dblDHours#
            dblNewEntitle# = dblNewEntitle# * dblDHours#
            dblEntitleUpd = dblNewEntitle
        End If
        If xTypeF(intWhereFit&) = True Then
            If xService(intWhereFit&, 3) <> 0 Then dblNewMax# = xService(intWhereFit&, 3) * dblFTEHours# * dblDHours#
            dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
            dblEntitleUpd = dblNewEntitle
        End If
        If xTypeH(intWhereFit&) = True Then
            If xService(intWhereFit&, 3) <> 0 Then dblNewMax# = xService(intWhereFit&, 3)
        End If
        
        'Type of Vacation Calculation (Monthly/Annualized Monthly or Annually)
        If fglbCompMonthly Then     'Monthly or Annualized Monthly
            dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
        Else
            dblEntitleUpd = dblNewEntitle ' rollover is on other utility (to accumulate)
        End If
        
         If dblNewMax <> 0 Then          'only do if not zero
            If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
                dblEntitleUpd = dblNewMax - dblPrevEntitle#
            End If
        End If
        
        DtTm = Now
    End If

    If if_Vacation Then
        VacpcN = xService(intWhereFit&, 4)
        VacpcO = snapEntitle("ED_VACPC")
        VED_DIV = snapEntitle("ED_DIV")
        VED_PT = snapEntitle("ED_PT")
        If IsNumeric(xService(intWhereFit&, 4)) Then snapEntitle("ED_VACPC") = xService(intWhereFit&, 4)
        
    End If
Stub_Cont1:
    If if_Entitle Then
        
        If glbCompSerial = "S/N - 2188W" Then
            dblEntitleUpd = Round(dblEntitleUpd, 0)
        ElseIf glbCompSerial = "S/N - 2297W" Then
            If dblEntitleUpd >= 14.9 And dblEntitleUpd <= 15.1 Then
                dblEntitleUpd = 15
            ElseIf dblEntitleUpd >= 19.9 And dblEntitleUpd <= 20.1 Then
                dblEntitleUpd = 20
            ElseIf dblEntitleUpd >= 25.1 And dblEntitleUpd <= 25.1 Then
                dblEntitleUpd = 25
            End If
        End If
                        
        'Hemu - 12/31/2003 Begin - Ticket #5348 - City of Chatham-Kent
        If (glbCompSerial = "S/N - 2188W" Or glbCompSerial = "S/N - 2228W") And month(CVDate(xAsOf)) = 12 Then
            snapEntitle("ED_VAC") = Round(dblEntitleUpd, 0)      ' base entitlements sic/vacation
        End If
        'Hemu - 12/31/2003 End
        
        'Mitchell Plastics
        If glbCompSerial = "S/N - 2335W" And InStr(1, glbSeleDiv, "HSV") = 0 Then
            'If Anniversary Today or past then give the entitlement
            'Ticket #15130 - Begin - Asked to change the logic for when to update for Anniversary and Jan 1 update.
            'If flgOnAnniversary And (Month(varStartDate) <= Month(currDate)) And (Day(varStartDate) <= Day(currDate)) Then
            If Not flgStubPeriod Then
                'Ticket #22730
                'xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd
                xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd & ". OS: " & (IIf(IsNull(snapEntitle("ED_PVAC")), 0, snapEntitle("ED_PVAC")) + IIf(IsNull(snapEntitle("ED_VAC")), 0, snapEntitle("ED_VAC"))) - IIf(IsNull(snapEntitle("ED_VACT")), 0, snapEntitle("ED_VACT"))

                Call Append_Accrual(empNo&, "VAC", currDate, dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
            
                snapEntitle("ED_VAC") = dblEntitleUpd
                
            'If Anniversary after Dec 31st then give the entitlement
            'ElseIf flgOnJan1 Then
            '    xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd
            '    Call Append_Accrual(EmpNo&, "VAC", currDate, dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
                
            '    snapEntitle("ED_VAC") = dblEntitleUpd
            'Ticket #15130 - End
            ElseIf flgStubPeriod Then
                'Ticket #22730
                'xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd
                xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd & ". OS: " & (IIf(IsNull(snapEntitle("ED_PVAC")), 0, snapEntitle("ED_PVAC")) + IIf(IsNull(snapEntitle("ED_VAC")), 0, snapEntitle("ED_VAC"))) - IIf(IsNull(snapEntitle("ED_VACT")), 0, snapEntitle("ED_VACT"))

                Call Append_Accrual(empNo&, "VAC", currDate, dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
                
                If flgWithin10 Then
                    snapEntitle("ED_VAC") = dblEntitleUpd
                Else
                    'Hemu - Ticket #14993
                    'snapEntitle("ED_VAC") = 0
                End If
            End If
        End If
                    
        If (glbCompSerial <> "S/N - 2335W") Or (glbCompSerial = "S/N - 2335W" And InStr(1, glbSeleDiv, "HSV") > 0) Then     'Not Mitchell Plastics
            'Ticket #22730
            'xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd
            xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd & ". OS: " & (IIf(IsNull(snapEntitle("ED_PVAC")), 0, snapEntitle("ED_PVAC")) + IIf(IsNull(snapEntitle("ED_VAC")), 0, snapEntitle("ED_VAC"))) - IIf(IsNull(snapEntitle("ED_VACT")), 0, snapEntitle("ED_VACT"))

            Call Append_Accrual(empNo&, "VAC", currDate, dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
            snapEntitle("ED_VAC") = dblEntitleUpd       ' base entitlements sic/vacation
        End If
    End If
    snapEntitle.Update
    
    If if_Vacation Then
        If Val(Format(VacpcN)) <> Val(Format(VacpcO)) Then
        
            SQLQW1 = "INSERT INTO HRAUDIT (AU_TYPE,AU_NEWEMP,AU_EMPNBR,AU_VACPC,AU_OLDVAC, "
            SQLQW1 = SQLQW1 & "AU_DIVUPL,AU_PTUPL,AU_LDATE,AU_LTIME,AU_UPLOAD,AU_LUSER) "
            
            SQLQW1 = SQLQW1 & " VALUES('M','N'," & empNo& & "," & Val(Format(VacpcN)) & "," & Val(Format(VacpcO))
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
    xKey = xKey & "|" & Format(dlpAsOf.Text, "dd-mmm-yyyy") 'Transaction Date
    Call Entitlements_Master_Integration(xKey, empNo&) 'George added for Advance Tracker

lblNextRec1:
    snapEntitle.MoveNext
    DoEvents
Wend
modDailyUpdateSelection_Auto = True
gdbAdoIhr001.CommitTrans

snapEntitle.Close

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodPercent = 0
MDIMain.panHelp(0).FloodType = 0
Screen.MousePointer = DEFAULT


Exit Function

modDailyUpdateSelection_Auto_Err:
'These errors are:
'13=type mismatch
'94=invalid use of null
'3018=couln't find field 'item'
If Err = 13 Or Err = 94 Or Err = 3018 Then
    Err = 0
    Resume Next
End If

Screen.MousePointer = DEFAULT
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateEntitle", "HR_EMP", "edit/Add")
'Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
    'Rollback
'    Resume Next
'Else
'    Unload Me
'End If

MDIMain.panHelp(0).Caption = "An error occurred in Vacation Entitlement Calculation procedure"

End Function
'===========================================================================================


Public Function Calculate_Stub_Period_Entitlement(xEmpnbr, xHireDate, vacFrom)
Dim rsAttend As New ADODB.Recordset
Dim rsHRTabl As New ADODB.Recordset
Dim SQLQ As String
Dim xReasonLst As String
Dim xDaysWorked, xNoStubWeeks, xNoOfWeeks
Dim xAvgWorked, xRatioStub, xVacEnt

    '1. --- Average Number of Days worked per week ---
    'Calculate # of Days worked in Stub Period
        '- where Attendance Code Absent Flag = "NO"
    SQLQ = "SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' "
    SQLQ = SQLQ & " AND TB_ABSENCE = 0"
    rsHRTabl.Open SQLQ, gdbAdoIhr001, adOpenStatic

    xReasonLst = ""
    If Not rsHRTabl.EOF Then
        rsHRTabl.MoveFirst
        Do While Not rsHRTabl.EOF
            xReasonLst = xReasonLst & "'" & rsHRTabl("TB_KEY") & "',"
            rsHRTabl.MoveNext
        Loop
        xReasonLst = Mid(xReasonLst, 1, Len(xReasonLst) - 1)    'trim off (,) at the end
    End If
    SQLQ = "SELECT COUNT(*) AS TOTDAYS FROM (SELECT AD_DOA FROM HR_ATTENDANCE "
    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & xEmpnbr & " AND AD_REASON IN (" & xReasonLst & ")"
    SQLQ = SQLQ & " AND (AD_DOA >= " & Date_SQL(xHireDate) & " AND AD_DOA <= " & Date_SQL(DateAdd("d", "-1", vacFrom)) & ")"
    SQLQ = SQLQ & " GROUP BY AD_DOA) AS TEMP"
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsAttend.EOF Then
        xDaysWorked = rsAttend("TOTDAYS")
    Else
        xDaysWorked = 0
    End If

    'Calculate # of Weeks in Stub Period
    xNoStubWeeks = DateDiff("w", xHireDate, DateAdd("d", "-1", vacFrom))

    'Find out the Average Number of Days worked per week
        'Averg = # of Days Worked in Stub Period / # of Weeks in Stub Period
    xAvgWorked = xDaysWorked / xNoStubWeeks

    '2. --- Ratio of the Stub Period to 12 months or 52 weeks ---
    'Calculate # of Months in Stub Period; They suggested us to use Weeks in Stub Period
    'Find out the Ratio of the Stub Period to 52 weeks
        'Ratio = # of Weeks in Stub Period / 52
    xRatioStub = xNoStubWeeks / 52

    '3. --- Vacation Entitlement Amount = Must be taken withing 10 months (by Jul 1st - May 1st) ---
        'VacEnt = Round Up(2 * Averg * Ratio)
    xVacEnt = 2 * xAvgWorked * xRatioStub

    
    If (xVacEnt - Int(xVacEnt)) <= 0.5 And (xVacEnt - Int(xVacEnt)) > 0 Then 'Round up to next .5
        Calculate_Stub_Period_Entitlement = Int(xVacEnt) + 0.5
    Else
        Calculate_Stub_Period_Entitlement = Round(xVacEnt)
    End If
    
    'New Logic - Ticket #15130 - Paid for logic - # of Days based on the month of hire
    Select Case month(xHireDate)
        Case 7: Calculate_Stub_Period_Entitlement = 10
        Case 1: Calculate_Stub_Period_Entitlement = 5
        Case 8: Calculate_Stub_Period_Entitlement = 9
        Case 2: Calculate_Stub_Period_Entitlement = 4
        Case 9: Calculate_Stub_Period_Entitlement = 8
        Case 3: Calculate_Stub_Period_Entitlement = 3
        Case 10: Calculate_Stub_Period_Entitlement = 7
        Case 4: Calculate_Stub_Period_Entitlement = 3
        Case 11: Calculate_Stub_Period_Entitlement = 7
        Case 5: Calculate_Stub_Period_Entitlement = 2
        Case 12: Calculate_Stub_Period_Entitlement = 6
        Case 6: Calculate_Stub_Period_Entitlement = 1
    End Select
    
End Function

Private Function modUpdateSelection(Optional isLast As Boolean)
Dim empNo As Long
Dim dblEntitle#, dblPrevEntitle#, dblTKEEntitle#, strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#, dblWHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%, xAsOf
Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
Dim if_Entitle As Boolean, if_Vacation As Boolean
Dim xComments
Dim dblEntitleDays
Dim xSALDIST As String
Dim xTotEmpHours 'Ticket #21843 Franks 04/12/2012
Dim xNoDaysPerWk    'Ticket #25476 - Family Day Care Services

On Error GoTo modUpdateSelection_Err

modUpdateSelection = False

If Len(dlpAsOf.Text) = 0 Then
    MsgBox "Effective Date is required field"
    dlpAsOf.SetFocus
    Exit Function
End If

If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)

Screen.MousePointer = DEFAULT

If cmdYearEnd.Visible = False Then
    If snapEntitle.BOF And snapEntitle.EOF Then
        'If fglbRunTimes = 1 Then
            MsgBox "Employees for this selection do not exist!"
            Exit Function
        'End If
    Else
        lngRecs& = snapEntitle.RecordCount
        If fglbRunTimes = 1 Or UCase(glbCompEntVac$) <> "N" Then   'Ticket #26777 - Prompt for Annual and Monthly as well
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
End If
lngRecs& = snapEntitle.RecordCount

'Ticket #22682 - Release 8.0: Check Accrual File to see if the update already done for Monthly Updates only. This is
'to avoid multiple updates for the same month.
'Only for Monthly updates
If glbCompEntVac$ = "M" Then
    Do While Not snapEntitle.EOF
        'Ticket #28024 - To fix the error caused by calling this function without '' apostrophes
        'If Accrual_Rec_Exists(snapEntitle("ED_EMPNBR"), "VAC", dlpAsOf.Text, "U") Then
        If Accrual_Rec_Exists(snapEntitle("ED_EMPNBR"), "VAC", dlpAsOf.Text, "'U'") Then
            Response% = MsgBox("'Update Entitlement' already done for at least 1 employee in this selection for the Effective Date: " & dlpAsOf.Text & "." & Chr(10) & Chr(10) & "Are you sure you want to proceed with this Update?", vbExclamation + vbYesNo, "Update Entitlements")
            If Response% = IDNO Then
                Exit Function
            End If
            
            Exit Do
        End If
        
        snapEntitle.MoveNext
        DoEvents
    Loop
End If

snapEntitle.MoveFirst

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 5

For x% = 0 To 24
    If Not IsNumeric(medLTServ(x%)) Then
        medLTServ(x%) = 0
    End If
    If Not IsNumeric(medGTServ(x%)) Then
        medGTServ(x%) = 0
    Else
        If glbFrench Then
            If medGTServ(x%) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
        Else
            If Val(medGTServ(x%)) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
        End If
    End If
    If medLTServ(x%) > 0 And medGTServ(x%) = 0 Then medGTServ(x%) = 9999999
Next

'gdbAdoIhr001.BeginTrans

While Not snapEntitle.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / lngRecs&))
    MDIMain.panHelp(0).FloodPercent = pct%
    if_Entitle = False
    if_Vacation = False

    empNo& = snapEntitle("ED_EMPNBR")
    
    If IsNull(snapEntitle("ED_VAC")) Then
        dblEntitle# = 0
    Else
        dblEntitle# = snapEntitle("ED_VAC")
    End If
    
    If IsNull(snapEntitle("ED_PVAC")) Then
        dblPrevEntitle# = 0
    Else
        dblPrevEntitle# = snapEntitle("ED_PVAC")
    End If
    
    If IsNull(snapEntitle("ED_VACT")) Then
        dblTKEEntitle# = 0
    Else
        dblTKEEntitle# = snapEntitle("ED_VACT")
    End If
    
    spt = snapEntitle("ED_PT")
    
    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec

    varStartDate = snapEntitle(fglbWDate$)
    
    Dim rsJOB As New ADODB.Recordset
    If rsJOB.State <> 0 Then rsJOB.Close
    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
    dblDHours# = 0
    dblFTEHours# = 0
    dblWHours# = 0      'Ticket #25476 - Family Day Care Services
    If Not rsJOB.EOF Then
        If IsNumeric(rsJOB("JH_DHRS")) Then dblDHours# = rsJOB("JH_DHRS")
        If IsNumeric(rsJOB("JH_FTENUM")) Then dblFTEHours# = rsJOB("JH_FTENUM")
        dblWHours# = GetJHData(snapEntitle("ED_EMPNBR"), "JH_WHRS", 0)      'Ticket #25476 - Family Day Care Services
    End If
    'rsJOB.Close - move it to the botton of 2433W section
    If glbLinamar Then dblDHours# = 8
    
    xAsOf = dlpAsOf.Text
'    dblServiceYears# = (DateDiff("d", varStartDate, CVDate(xAsOf)) / 365) * 12
    If glbSamuel Then 'Ticket #23385 Franks 03/22/2013
        dblServiceYears# = getSamMonthDiff(CVDate(varStartDate), CVDate(xAsOf)) ' MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    Else
        dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    End If
    intWhereFit& = -1

    For x% = 0 To 24
        If medGTServ(x%) > 0 Then
            If dblServiceYears# >= CDbl(medLTServ(x%)) And dblServiceYears# <= CDbl(medGTServ(x%)) Then
                intWhereFit& = x%
                If Len(medEntitle(x%)) > 0 Then if_Entitle = True
                If Len(medVacation(x%)) > 0 Then if_Vacation = True
                Exit For
            End If
        End If
    Next x%
    
    'Ticket #16145 - Check for Mitchell Plastics their new hire and less than 12months Seniority logic
    'if true then call procedure to compute the entitlement for < 12 months logic
    'if new hire with Seniority between entitlement date then 0 entitlement
    'Then Goto Contd_Mitchell
    If glbCompSerial = "S/N - 2335W" And InStr(1, glbSeleDiv, "HSV") = 0 Then    'Mitchell Plastics
        If CVDate(varStartDate) >= CVDate(dlpDateRange(0)) And CVDate(varStartDate) <= CVDate(dlpDateRange(1)) Then
            if_Entitle = True
            dblEntitleUpd = 0
            GoTo Contd_Mitchell
        ElseIf dblServiceYears# < 12 And clpDiv.Text = "ULT" Then
            if_Entitle = True
            dblEntitleUpd = Assign_Entitlements_Mitchell(month(CVDate(varStartDate))) * dblDHours#
            GoTo Contd_Mitchell
        ElseIf dblServiceYears# < 12 And clpDiv.Text = "MIT" Then ' 24 -> 12 'Ticket #23034 Franks 01/18/2012
            if_Entitle = True
            dblEntitleUpd = Assign_Entitlements_Mitchell_MIT(month(CVDate(varStartDate))) * dblDHours#
            GoTo Contd_Mitchell
        End If
    End If
    
    If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
    
    'Ticket #22766 - KidsLink - sum up the FTE for multi positions
    If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then  'Kerrys Place Ticket #21843 Franks 04/12/2012, they need the total of hours for multiple current positions
        xTotEmpHours = 0
        Do While Not rsJOB.EOF
            If optD(intWhereFit&) = True Then  ' Entitlements entered in days
                If IsNumeric(rsJOB("JH_DHRS")) Then xTotEmpHours = xTotEmpHours + rsJOB("JH_DHRS")
            End If
            If optF(intWhereFit&) = True Then  ' FTE
                If IsNumeric(rsJOB("JH_DHRS")) And IsNumeric(rsJOB("JH_FTENUM")) Then
                    xTotEmpHours = xTotEmpHours + rsJOB("JH_DHRS") * rsJOB("JH_FTENUM")
                End If
            End If
            rsJOB.MoveNext
        Loop
    End If
    rsJOB.Close
    
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

    If if_Entitle Then
        dblNewEntitle# = medEntitle(intWhereFit&)
        If glbSamuel Then 'Ticket #23385 Franks 03/22/2013
            isConYear = False 'Ticket #23812 Franks 05/23/2013
            If optSamuelType(0).Value Then 'Service Center
                dblNewEntitle# = getSamNewhireMonthVacEnt(dblNewEntitle#, CVDate(varStartDate), CVDate(xAsOf))
                ''Ticket #23385 Franks 03/25/2013 - begin
                'If month(xAsOf) = 1 Then 'first month
                '    xFirstMonEnt = dblNewEntitle#
                '    'isConYear = False
                'Else
                '    If Not xFirstMonEnt = dblNewEntitle# Then
                '    'if one month jumps to next rule during this year, then keep all the original vac ent for all 12 months
                '    'then plus additional vac days
                '        dblNewEntitle# = xFirstMonEnt
                '        isConYear = True
                '    End If
                'End If
                ''Ticket #23385 Franks 03/25/2013 - end
            End If
        End If
        dblNewMax# = 0
        If optD(intWhereFit&) = True Then           ' Entitlements entered in days
            'Ticket #22766 - KidsLink - sum up the FTE for multi positions
            If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then  'Kerrys Place Ticket #21843 Franks 04/12/2012
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * xTotEmpHours
                dblNewEntitle# = dblNewEntitle# * xTotEmpHours
            Else
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblDHours#
                
                'Ticket #25476 - Family Day Care Services. Special formula to compute # of days per week an Employee works and
                'use that to compute Entitlement
                If glbCompSerial = "S/N - 2436W" Then
                    'Compute # of Day per Week an employee works
                    If dblDHours# <> 0 Then
                        xNoDaysPerWk = dblWHours# / dblDHours#
                        
                        'Entitlemnent based on # of Days per Week an employee works
                        If xNoDaysPerWk < 5 Then
                            dblNewEntitle# = (dblNewEntitle# / dblDHours#) * xNoDaysPerWk * dblDHours#
                        Else
                            dblNewEntitle# = dblNewEntitle# * dblDHours#
                        End If
                    Else
                        dblNewEntitle# = 0
                    End If
                Else
                    dblNewEntitle# = dblNewEntitle# * dblDHours#
                End If
            End If
            dblEntitleUpd = dblNewEntitle
        End If
        If optF(intWhereFit&) = True Then
            'Ticket #22766 - KidsLink - sum up the FTE for multi positions
            If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then  'Kerrys Place Ticket #21843 Franks 04/12/2012
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * xTotEmpHours
                dblNewEntitle# = dblNewEntitle# * xTotEmpHours
            Else
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHours# * dblDHours#
                dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
            End If
            dblEntitleUpd = dblNewEntitle
        End If
        If optH(intWhereFit&) = True Then
            If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&)
        End If
        If fglbCompMonthly Then
            If glbCompSerial = "S/N - 2322W" Then   'Family & Children's Services of Guelph and Wellington County
'                If dblDHours# <> 0 Then
'                    dblNewEntitle = Round25(dblNewEntitle / dblDHours#) * dblDHours#
'                End If
                
                If fglbRunTimes = 1 Then
                    dblEntitleUpd = dblNewEntitle
                Else
                    dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
                    If dblDHours# <> 0 And fglbRunTimes = 12 Then
                        dblEntitleUpd# = Round25(dblEntitleUpd# / dblDHours#) * dblDHours#
                    End If
                    
                End If
            Else
                dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
            End If
        Else
            dblEntitleUpd = dblNewEntitle ' rollover is in other utility (to accumulate)
        End If
        
        If dblNewMax <> 0 Then          'only do if not zero
            If glbCompSerial = "S/N - 2423W" Then
                'Accellos Canada Inc.
                xSALDIST = "N"
                If Not IsNull(snapEntitle("ED_SALDIST")) Then
                    If UCase(snapEntitle("ED_SALDIST")) = "Y" Then
                        xSALDIST = "Y"
                    End If
                End If
                If xSALDIST = "Y" Then
                    'Ticket #18644
                    'do not use the Entitlement Maximum if Salary Distribution is "Y"
                Else
                    If (dblEntitle# + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
                        dblEntitleUpd = dblEntitle#
                    ElseIf (dblEntitleUpd + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
                        dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
                    End If
                End If
            Else
                'Ticket #21905 - kidsLINK
                If glbCompSerial = "S/N - 2430W" Then
                    If (dblEntitle# + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
                        dblEntitleUpd = dblEntitle#
                    ElseIf (dblEntitleUpd + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
                        dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
                    End If
                Else
                    If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
                        dblEntitleUpd = dblNewMax - dblPrevEntitle#
                    End If
                End If
            End If
        End If
        
        DtTm = Now
    End If

    If if_Vacation Then
        If glbCBrant And Len(clpCode(3).Text) > 0 And snapEntitle("ED_SECTION") >= clpCode(3).Text Then
            VacpcN = medVacation(intWhereFit&) + dblEntitle#
        Else
            VacpcN = medVacation(intWhereFit&)
        End If
        VacpcO = snapEntitle("ED_VACPC")
        VED_DIV = snapEntitle("ED_DIV")
        VED_PT = snapEntitle("ED_PT")
        If IsNumeric(medVacation(intWhereFit&)) Then snapEntitle("ED_VACPC") = medVacation(intWhereFit&)
        
    End If
    
Contd_Mitchell:
    If if_Entitle Then

        'If glbCompSerial = "S/N - 2188W" Then  'Ticket #8887
        '    dblEntitleUpd = Round(dblEntitleUpd, 0)
        If glbCompSerial = "S/N - 2297W" Then
            If dblEntitleUpd >= 14.9 And dblEntitleUpd <= 15.1 Then
                dblEntitleUpd = 15
            ElseIf dblEntitleUpd >= 19.9 And dblEntitleUpd <= 20.1 Then
                dblEntitleUpd = 20
            ElseIf dblEntitleUpd >= 25.1 And dblEntitleUpd <= 25.1 Then
                dblEntitleUpd = 25
            End If
        End If
        If glbCBrant And Len(clpCode(3).Text) > 0 Then
            'dblEntitleUpd = medVacation(intWhereFit&) + dblEntitle#
            dblEntitleUpd = medVacation(intWhereFit&) + dblEntitleUpd 'Ticket #12480, dblEntitle# was 0
        End If
                                        
        If isLast And glbCompSerial = "S/N - 2376W" Then '#9536 on Oct 21,2005 George
            If dblDHours# <> 0 Then
                dblEntitleDays = dblEntitleUpd / dblDHours#
                dblEntitleDays = Round((dblEntitleDays / 0.25 + 0.1), 0) * 0.25 ' round to 1/4 days
                dblEntitleUpd = dblEntitleDays * dblDHours#
            Else
                dblEntitleUpd = dblEntitleUpd
            End If
        ElseIf isLast And glbSamuel Then 'Ticket #23385 Franks 03/21/2013
            If optSamuelType(0).Value Then 'Service Center
                If dblDHours# <> 0 Then
                    dblEntitleDays = dblEntitleUpd / dblDHours#
                Else
                    dblEntitleUpd = dblEntitleUpd
                End If
                'If isConYear Then 'Ticket #23812 Franks 05/23/2013
                '    dblEntitleDays = dblEntitleDays + getSamAdditionalVac(CVDate(varStartDate), CVDate(xAsOf)) ', medLTServ(1).Text)
                'End If
                dblEntitleDays = Round((dblEntitleDays / 0.5 + 0.01), 0) * 0.5 ' round to 1/2 days
                dblEntitleUpd = dblEntitleDays * dblDHours#
            End If
            If optSamuelType(1).Value Then 'Non Service Center
                If dblDHours# <> 0 Then
                    dblEntitleDays = dblEntitleUpd / dblDHours#
                    dblEntitleDays = Round((dblEntitleDays / 0.5 + 0.01), 0) * 0.5 ' round to 1/2 days
                    dblEntitleUpd = dblEntitleDays * dblDHours#
                Else
                    dblEntitleUpd = dblEntitleUpd
                End If
            End If
        ElseIf isLast And chkRound.Visible = True And chkRound Then
            'Round the final entitlement
            If dblDHours# <> 0 And optH(intWhereFit&) = False Then
                dblEntitleDays = dblEntitleUpd / dblDHours#
                
                If glbCompSerial = "S/N - 2344W" Then   'Ticket #27761 - Cascade Canada Ltd - Round to nearest day
                    'dblEntitleDays = Round((dblEntitleDays + 0.5), 0)
                    dblEntitleDays = Round(dblEntitleDays, 1)
                    dblEntitleDays = Round(dblEntitleDays, 0)
                Else
                    dblEntitleDays = Round(dblEntitleDays, 0)
                End If
                
                dblEntitleUpd = dblEntitleDays * dblDHours#
            Else
                dblEntitleUpd = Round(dblEntitleUpd, 0)
            End If
        Else
            If glbCompEntVac$ = "M" And chkRound.Visible = True And chkRound Then
                'If month(dlpDateRange(1).Text) = month(dlpAsOf.Text) And Year(dlpDateRange(1).Text) = Year(dlpAsOf.Text) Then
                    'Round the final entitlement
                    If dblDHours# <> 0 And optH(intWhereFit&) = False Then
                        dblEntitleDays = dblEntitleUpd / dblDHours#
                        
                        If glbCompSerial = "S/N - 2344W" Then   'Ticket #27761 - Cascade Canada Ltd - Round to nearest day
                            'dblEntitleDays = Round((dblEntitleDays + 0.5), 0)
                            dblEntitleDays = Round(dblEntitleDays, 1)
                            dblEntitleDays = Round(dblEntitleDays, 0)
                        Else
                            dblEntitleDays = Round(dblEntitleDays, 0)
                        End If
                        
                        dblEntitleUpd = dblEntitleDays * dblDHours#
                    Else
                        dblEntitleUpd = Round(dblEntitleUpd, 0)
                    End If
                'Else
                '    dblEntitleUpd = dblEntitleUpd       ' base entitlements sic/vacation
                'End If
            'Else
            '    snapEntitle("ED_VAC") = dblEntitleUpd       ' base entitlements sic/vacation
            End If
        End If
        
        'Ticket #22730
        'xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd
        xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd & ". OS: " & (IIf(IsNull(snapEntitle("ED_PVAC")), 0, snapEntitle("ED_PVAC")) + IIf(IsNull(snapEntitle("ED_VAC")), 0, snapEntitle("ED_VAC"))) - IIf(IsNull(snapEntitle("ED_VACT")), 0, snapEntitle("ED_VACT"))
        
        'Hemu - Ticket #11925 - Changed the Accrual Date from Effective Date to Entitlement Start Date
        'because otherwise it will not update Vadim until the date arrives in case it's not same as the
        'Entitlement Start Date.
        'Call Append_Accrual(EmpNo&, "VAC", dlpAsOf, dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
        If glbSamuel And optSamuelType(0).Value Then 'Service Center
            'will use Samuel_Vac_Ent_Cal 'Ticket #23812 Franks 05/23/2013
        Else
            If fglbCompMonthly Then
                Call Append_Accrual(empNo&, "VAC", dlpAsOf, dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
            Else
                'Annual
                'Ticket #23141
                If glbVadim Then
                    'For Vadim user's we need to send the full value that the employee Annual Accrued, since we are
                    'not doing zero out for Current in the Year End. This is revised steps for Vadim users only for
                    'the Year End.
                    Call Append_Accrual(empNo&, "VAC", dlpDateRange(0), dblEntitleUpd, "U", xComments)
                Else
                    Call Append_Accrual(empNo&, "VAC", dlpDateRange(0), dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
                End If
            End If
        End If
        
        'Hemu - 12/31/2003 Begin - Ticket #5348 - City of Chatham-Kent
        If (glbCompSerial = "S/N - 2188W" Or glbCompSerial = "S/N - 2228W") And month(CVDate(xAsOf)) = 12 Then
            snapEntitle("ED_VAC") = Round(dblEntitleUpd, 0)      ' base entitlements sic/vacation
        Else
            snapEntitle("ED_VAC") = dblEntitleUpd       ' base entitlements sic/vacation
        End If
        'Hemu - 12/31/2003 End
        
        'Added by bryan 13/Jun/06 Ticket#10916
        If glbCompSerial <> "S/N - 2380W" Then  'Ticket #13979 - Don't update for VitalAire - using Annual Vacation Entitlement screen to store the value to ED_ANNVAC
            snapEntitle("ED_ANNVAC") = snapEntitle("ED_VAC")
        End If
    End If
    snapEntitle.Update
    
    If isLast And glbSamuel Then 'Ticket #23812 Franks 05/23/2013
        If optSamuelType(0).Value Then 'Service Center
            Call Samuel_Vac_Ent_Cal(empNo&)
        End If
    End If
    
    If if_Vacation Then
        Dim auVpcn As Double
        If glbCompSerial = "S/N - 2350W" Then 'Listowel Ticket#12299
            auVpcn = Val(Format(VacpcN)) * 100
        Else
            auVpcn = Val(Format(VacpcN))
        End If
        SQLQW1 = "INSERT INTO HRAUDIT (AU_TYPE,AU_NEWEMP,AU_EMPNBR,AU_VACPC,AU_OLDVAC, "
        SQLQW1 = SQLQW1 & "AU_DIVUPL,AU_PTUPL,AU_LDATE,AU_LTIME,AU_UPLOAD,AU_LUSER) "
        
        SQLQW1 = SQLQW1 & " VALUES('M','N'," & empNo& & "," & auVpcn & "," & Val(Format(VacpcO))
        SQLQW1 = SQLQW1 & ",'" & VED_DIV & "','" & VED_PT & "', "
        SQLQW1 = SQLQW1 & Date_SQL(Date) & ", '"
        
        SQLQW1 = SQLQW1 & Time$ & "', "
        SQLQW1 = SQLQW1 & "'N', "
        SQLQW1 = SQLQW1 & "'" & glbUserID & "'"
        SQLQW1 = SQLQW1 & ")"
        gdbAdoIhr001X.Execute SQLQW1
    End If
    Dim xKey
    xKey = snapEntitle("ED_EMPNBR")
    'xKey = xKey & "|" & Format(snapEntitle("ED_EFDATE"), "dd-mmm-yyyy")
    'xKey = xKey & "|" & Format(snapEntitle("ED_ETDATE"), "dd-mmm-yyyy")
    xKey = xKey & "|" & Format(dlpDateRange(0), "dd-mmm-yyyy")
    xKey = xKey & "|" & Format(dlpDateRange(1), "dd-mmm-yyyy")
    xKey = xKey & "|VAC"
    xKey = xKey & "|" & dblEntitleUpd
    xKey = xKey & "|" & Format(dlpAsOf.Text, "dd-mmm-yyyy")  'Transaction Date
    Call Entitlements_Master_Integration(xKey, empNo&) 'George added for Advance Tracker

lblNextRec:
    snapEntitle.MoveNext
    DoEvents
Wend
modUpdateSelection = True
MDIMain.panHelp(0).FloodType = 0
'gdbAdoIhr001.CommitTrans

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

Private Sub optD_Click(Index As Integer, Value As Integer)
    Call ST_OPT_VALUE
End Sub

Private Sub optD_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optF_Click(Index As Integer, Value As Integer)
    Call ST_OPT_VALUE
End Sub

Private Sub optF_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optH_Click(Index As Integer, Value As Integer)
    Call ST_OPT_VALUE
End Sub

Private Sub optH_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub scrControl_Change()
VacFram.Top = 4140 - scrControl.Value
End Sub

Sub ST_UPD_MODE(TF As Boolean)
Dim x, FT
FT = Not TF
For x = 0 To 24
    medLTServ(x).Enabled = TF
    medGTServ(x).Enabled = TF
    medEntitle(x).Enabled = TF
    If x = 0 Then
    optD(x).Enabled = TF
    optH(x).Enabled = TF
    optF(x).Enabled = TF
    Else
    optD(x).Enabled = False
    optH(x).Enabled = False
    optF(x).Enabled = False
    End If
    medMax(x).Enabled = TF
    medVacation(x).Enabled = TF
Next

clpDiv.Enabled = TF
clpDept.Enabled = TF
clpCode(0).Enabled = TF
If Not TF Or glbLinamar Then
    lblAsOf.FontBold = True
Else
    lblAsOf.FontBold = False
End If
If glbCompEntVac$ = "M" Or glbCompEntVac$ = "N" Then
    dlpAsOf.Enabled = TF 'FT
Else
    dlpAsOf.Enabled = True 'Ticket #3419
End If
'If Vacation Entitlement Outstanding based on "1" then ok, otherwise disenable
If glbEntOutStanding$ = "1" Then
    dlpDateRange(0).Enabled = TF
    dlpDateRange(1).Enabled = TF
    CmdRecalc.Enabled = True
Else
    dlpDateRange(0).Enabled = False
    dlpDateRange(1).Enabled = False
    CmdRecalc.Enabled = False
End If
If Not glbWHSCC Then
    clpCode(1).Enabled = TF
Else
    clpCode(1).Enabled = False
End If
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    medHours.Enabled = TF
Else
    clpCode(2).Enabled = TF
End If
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
Call modSetFGlobals("Vac")
End Sub

Sub Display_Value()
Dim SQLQ, xOrder, nOrder, aa, SQLQW, glbiOneWhere
Dim rsVE As New ADODB.Recordset
Dim x

For x = 0 To 24
    medLTServ(x) = ""
    medGTServ(x) = ""
    medEntitle(x) = ""
    optD(x) = True
    optH(x) = False
    optF(x) = False
    medMax(x) = ""
    medVacation(x) = ""
Next
clpDiv.Text = ""
clpDept.Text = ""
clpCode(0).Text = ""
If Not (glbCompEntVac$ = "M" Or glbCompEntVac$ = "N") Then
    dlpAsOf.Text = ""
End If
clpCode(1).Text = ""
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    medHours.Text = ""
Else
    clpCode(2).Text = ""
End If
clpCode(3).Text = ""
clpCode(4).Text = ""
clpPT.Text = ""
dlpDateRange(0).Text = ""
dlpDateRange(1).Text = ""

If Not Data1.Recordset.EOF Then
    SQLQ = "SELECT * FROM HRVACENT "
    If IsNull(Data1.Recordset("VE_DIV")) Then
        SQLQ = SQLQ & " WHERE VE_DIV IS NULL"
    Else
        SQLQ = SQLQ & " WHERE VE_DIV = '" & Data1.Recordset("VE_DIV") & "'"
    End If
    If IsNull(Data1.Recordset("VE_DEPT")) Then
        SQLQ = SQLQ & " AND VE_DEPT IS NULL"
    Else
        SQLQ = SQLQ & " AND VE_DEPT = '" & Data1.Recordset("VE_DEPT") & "'"
    End If
    If IsNull(Data1.Recordset("VE_ORG")) Then
        SQLQ = SQLQ & " AND VE_ORG IS NULL"
    Else
        SQLQ = SQLQ & " AND VE_ORG = '" & Data1.Recordset("VE_ORG") & "'"
    End If
    If IsNull(Data1.Recordset("VE_LOC")) Then
        SQLQ = SQLQ & " AND VE_LOC IS NULL"
    Else
        SQLQ = SQLQ & " AND VE_LOC = '" & Data1.Recordset("VE_LOC") & "'"
    End If
    If IsNull(Data1.Recordset("VE_SECTION")) Then
        SQLQ = SQLQ & " AND VE_SECTION IS NULL"
    Else
        SQLQ = SQLQ & " AND VE_SECTION = '" & Data1.Recordset("VE_SECTION") & "'"
    End If
    If glbWFC Then 'Ticket #28553 Franks 05/03/2016
        SQLQ = SQLQ & " AND " & getWFCPlantSecurity("VE_SECTION")
    End If
    If Not IsNull(Data1.Recordset("VE_EDATE")) Then
        SQLQ = SQLQ & " AND VE_EDATE = " & Date_SQL(Data1.Recordset("VE_EDATE"))
    End If
    If IsNull(Data1.Recordset("VE_EMP")) Then
        SQLQ = SQLQ & " AND VE_EMP IS NULL"
    Else
        SQLQ = SQLQ & " AND VE_EMP = '" & Data1.Recordset("VE_EMP") & "'"
    End If
    If IsNull(Data1.Recordset("VE_PT")) Then
        SQLQ = SQLQ & " AND VE_PT IS NULL"
    Else
        SQLQ = SQLQ & " AND VE_PT = '" & Data1.Recordset("VE_PT") & "' "
    End If
    If IsNull(Data1.Recordset("VE_GRPCD")) Then
        SQLQ = SQLQ & " AND VE_GRPCD IS NULL"
    Else
        SQLQ = SQLQ & " AND VE_GRPCD = '" & Data1.Recordset("VE_GRPCD") & "'"
    End If
    If Not IsNull(Data1.Recordset("VE_FRDATE")) Then
        SQLQ = SQLQ & " AND VE_FRDATE = " & Date_SQL(Data1.Recordset("VE_FRDATE"))
    End If
    If Not IsNull(Data1.Recordset("VE_TODATE")) Then
        SQLQ = SQLQ & " AND VE_TODATE = " & Date_SQL(Data1.Recordset("VE_TODATE"))
    End If
    
    SQLQ = SQLQ & " Order By VE_DIV,VE_DEPT,VE_ORG, VE_EDATE,VE_EMP,VE_PT,VE_LOC,VE_SECTION,VE_ORDER "
    rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    If Not IsNull(Data1.Recordset("VE_DIV")) Then clpDiv.Text = Data1.Recordset("VE_DIV")
    If Not IsNull(Data1.Recordset("VE_DEPT")) Then clpDept.Text = Data1.Recordset("VE_DEPT")
    If Not IsNull(Data1.Recordset("VE_ORG")) Then clpCode(0).Text = Data1.Recordset("VE_ORG")
    If Not IsNull(Data1.Recordset("VE_EDATE")) Then dlpAsOf.Text = Data1.Recordset("VE_EDATE")
    If Not IsNull(Data1.Recordset("VE_EMP")) Then clpCode(1).Text = Data1.Recordset("VE_EMP")
    If Not IsNull(Data1.Recordset("VE_PT")) Then clpPT.Text = Data1.Recordset("VE_PT")
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
        If Not IsNull(Data1.Recordset("VE_GRPCD")) Then medHours.Text = Data1.Recordset("VE_GRPCD")
    Else
        If Not IsNull(Data1.Recordset("VE_GRPCD")) Then clpCode(2).Text = Data1.Recordset("VE_GRPCD")
    End If
    If Not IsNull(Data1.Recordset("VE_LOC")) Then clpCode(4).Text = Data1.Recordset("VE_LOC")
    If Not IsNull(Data1.Recordset("VE_SECTION")) Then clpCode(3).Text = Data1.Recordset("VE_SECTION")
    If Not IsNull(Data1.Recordset("VE_FRDATE")) Then dlpDateRange(0).Text = Data1.Recordset("VE_FRDATE")
    If Not IsNull(Data1.Recordset("VE_TODATE")) Then dlpDateRange(1).Text = Data1.Recordset("VE_TODATE")
    If Not IsNull(Data1.Recordset("VE_MANUAL")) Then
        chkManual.Value = Data1.Recordset("VE_MANUAL")
    End If
    If glbSamuel Then 'Ticket #23385 Franks 03/21/2013
        If Not IsNull(Data1.Recordset("VE_ROUNDENT")) Then
            If Data1.Recordset("VE_ROUNDENT") Then optSamuelType(0).Value = True Else optSamuelType(1).Value = True
        Else
            optSamuelType(0).Value = False
            optSamuelType(1).Value = False
        End If
    End If
    Do While Not rsVE.EOF
        xOrder = rsVE("VE_ORDER")
        nOrder = Format(Val(xOrder), "##0") - 1
        If Not (nOrder < 0 Or nOrder > 24) Then
            If Not IsNull(rsVE("VE_BMONTH")) Then medLTServ(nOrder) = rsVE("VE_BMONTH")
            If Not IsNull(rsVE("VE_EMONTH")) Then medGTServ(nOrder) = rsVE("VE_EMONTH")
            If Not IsNull(rsVE("VE_ENTITLE")) Then medEntitle(nOrder) = rsVE("VE_ENTITLE")
            If rsVE("VE_TYPE") = "D" Then optD(nOrder) = True
            If rsVE("VE_TYPE") = "H" Then optH(nOrder) = True
            If rsVE("VE_TYPE") = "F" Then optF(nOrder) = True
            If Not IsNull(rsVE("VE_MAX")) Then medMax(nOrder) = rsVE("VE_MAX")
            If Not IsNull(rsVE("VE_PCT")) Then medVacation(nOrder) = rsVE("VE_PCT")
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
    
    SQLQ = "SELECT DISTINCT VE_DIV,VE_DEPT,VE_ORG,VE_LOC,VE_SECTION,VE_EDATE,VE_EMP,VE_PT,VE_GRPCD,VE_FRDATE,VE_TODATE, VE_MANUAL FROM HRVACENT "
    If glbDIVCount = 1 And glbLinamar Then
        SQLQ = SQLQ & " WHERE VE_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
    End If
    If glbWFC Then 'Ticket #28553 Franks 05/03/2016
        SQLQ = SQLQ & " WHERE " & getWFCPlantSecurity("VE_SECTION")
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

    Data1.RecordSource = SQLQ
    Data1.Refresh
    
    orgEffDate = dlpAsOf.Text
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim UpdateFlg As Boolean
    Dim Response%
    
    UpdateFlg = False

    If clpDiv.Text <> ODIV Then UpdateFlg = True
    If clpDept.Text <> ODept Then UpdateFlg = True
    If clpCode(0).Text <> oOrg Then UpdateFlg = True
    If dlpAsOf.Text <> oAsOf Then UpdateFlg = True
    If clpCode(1).Text <> oEMP Then UpdateFlg = True
    If clpPT.Text <> oEmpMode Then UpdateFlg = True
    If clpCode(2).Text <> oGRPCE Then UpdateFlg = True
    If clpCode(4).Text <> OLoc Then UpdateFlg = True
    If clpCode(3).Text <> OSection Then UpdateFlg = True
    If dlpDateRange(0).Text <> OFromDate Then UpdateFlg = True
    If dlpDateRange(1).Text <> OToDate Then UpdateFlg = True
    If chkManual.Value <> OManual Then UpdateFlg = True

    If UpdateFlg = True Then
        Response% = MsgBox("Do you want to Save changes?", MB_YESNO, "Save Changes?")    ' Get user response.
        If Response% = IDYES Then     ' Evaluate response
            'Save the changes
            Call cmdOK_Click
            Pause (0.5)
        End If
    End If

Call Display_Value
End Sub

Private Sub modSetFGlobals(strTyp$)
fglbSick% = False
fglbVac% = True
If glbCompEntVac$ = "M" Or UCase(glbCompEntVac$) = "N" Then
    fglbCompMonthly% = True
    Call modMaximums(True)
Else
    fglbCompMonthly% = False
    If glbWHSCC Then
        Call modMaximums(True)
    Else
        Call modMaximums(False)
    End If
End If

End Sub

Sub ST_OPT_VALUE()
Dim x, XoptD, XoptH, XoptF
    XoptD = optD(0).Value
    XoptH = optH(0).Value
    XoptF = optF(0).Value
    For x = 1 To 24
        optD(x).Value = XoptD
        optH(x).Value = XoptH
        optF(x).Value = XoptF
    Next
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
If glbLinamar Then
    If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SALDIST = '" & clpCode(3).Text & "' "
Else
'If Not glbCBrant Then 'added by Bryan 18/Apr/2006 Ticket#10495
    If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & clpCode(3).Text & "' "
'End If
End If
If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #18235
    If Len(clpCode(4).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_VADIM1 = '" & clpCode(4).Text & "' "
Else
    If Len(clpCode(4).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & clpCode(4).Text & "' "
End If


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
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
        xGRPCE = medHours.Text
    Else
        xGRPCE = clpCode(2).Text
    End If
    xLoc = clpCode(4).Text
    xSECTION = clpCode(3).Text
    xFromDate = dlpDateRange(0)
    xToDate = dlpDateRange(1)
End If

If Len(xDiv) = 0 Then
    fglbVSQLQ = " (VE_DIV IS NULL OR VE_DIV='')"
Else
    fglbVSQLQ = "VE_DIV = '" & xDiv & "'"
End If
If Len(xDept) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VE_DEPT IS NULL OR VE_DEPT='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_DEPT = '" & xDept & "'"
End If
If Len(xORG) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VE_ORG IS NULL OR VE_ORG='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_ORG = '" & xORG & "'"
End If
If UCase(glbCompEntVac$) = "A" Then
    If Len(xAsOf) > 0 Then fglbVSQLQ = fglbVSQLQ & " AND  VE_EDATE = " & Date_SQL(xAsOf)
End If
If Len(xEMP) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VE_EMP IS NULL OR VE_EMP='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_EMP = '" & xEMP & "'"
End If
If Len(xEmpMode) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VE_PT IS NULL OR VE_PT='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_PT = '" & xEmpMode & "' "
End If
If Len(xGRPCE) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VE_GRPCD IS NULL OR VE_GRPCD='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_GRPCD = '" & xGRPCE & "'"
End If

If Len(xLoc) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VE_LOC IS NULL OR VE_LOC='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_LOC = '" & xLoc & "'"
End If
If Len(xSECTION) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VE_SECTION IS NULL OR VE_SECTION='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_SECTION = '" & xSECTION & "'"
End If

If Not IsDate(xFromDate) Then
    fglbVSQLQ = fglbVSQLQ & " AND VE_FRDATE IS NULL  "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_FRDATE = " & Date_SQL(xFromDate)
End If
If Not IsDate(xToDate) Then
    fglbVSQLQ = fglbVSQLQ & " AND VE_TODATE IS NULL  "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_TODATE = " & Date_SQL(xToDate)
End If
End Sub

Private Function modUpdateSelectionWHSCC()
Dim empNo As Long
Dim dblEntitle#, dblPrevEntitle#, strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%, xAsOf
Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
Dim if_Entitle As Boolean, if_Vacation As Boolean
Dim ifAnnual As Boolean, dblNewEntAnn#, VacpcNAnn, ifUnionDate As Boolean, ifFirstDate As Boolean 'Frank for WHSCC
Dim dblServiceYearsYTD, if_NON As Boolean
Dim xComments
' Entitlements are always valued in HOURS - if you enter days then it
'   works out how many hours (based on average Hrswrked/day found in salary master record)
On Error GoTo modUpdateSelectionWHSCC_Err
modUpdateSelectionWHSCC = False


If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
'
'If fTablHREMP.State <> 0 Then fTablHREMP.Close
'fTablHREMP.Open "HREMP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
Screen.MousePointer = DEFAULT
If snapEntitle.BOF And snapEntitle.EOF Then
    MsgBox "Employees for this selection do not exist!"
    Exit Function
Else
    lngRecs& = snapEntitle.RecordCount
    Msg$ = lngRecs& & " Records to process" & Chr(10) & "Would You Like To Proceed?"
    Title$ = "Update Entitlements"
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
        If glbFrench Then
            If medGTServ(x%) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
        Else
            If Val(medGTServ(x%)) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
        End If
    End If
    If medLTServ(x%) > 0 And medGTServ(x%) = 0 Then medGTServ(x%) = 9999999
Next

gdbAdoIhr001.BeginTrans

While Not snapEntitle.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / lngRecs&))
    MDIMain.panHelp(0).FloodPercent = pct%
    if_Entitle = False
    if_Vacation = False
    if_NON = False
    
    empNo& = snapEntitle("ED_EMPNBR")
    'If EmpNo& = 16 Then
    'Debug.Print "test"
    'End If
    If IsNull(snapEntitle("ED_VAC")) Then
        dblEntitle# = 0
    Else
        dblEntitle# = snapEntitle("ED_VAC")
    End If
    
    If IsNull(snapEntitle("ED_PVAC")) Then
        dblPrevEntitle# = 0
    Else
        dblPrevEntitle# = snapEntitle("ED_PVAC")
    End If
    
    spt = snapEntitle("ED_PT")
    strDivision$ = snapEntitle("ED_DIV")
    
    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec

    varStartDate = snapEntitle(fglbWDate$)
    
'    If glbLinamar Then
'        dblDHours# = 8
'    Else
'        If Not IsNumeric(snapEntitle("JH_DHRS")) Then
'            dblDHours# = 0
'        Else
'            dblDHours# = snapEntitle("JH_DHRS")
'        End If
'    End If
'    If Not IsNumeric(snapEntitle("JH_FTENUM")) Then
'        dblFTEHours# = 0
'    Else
'        dblFTEHours# = snapEntitle("JH_FTENUM")
'    End If
'
'
    Dim rsJOB As New ADODB.Recordset
    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
    dblDHours# = 0
    dblFTEHours# = 0
    If Not rsJOB.EOF Then
        If IsNumeric(rsJOB("JH_DHRS")) Then dblDHours# = rsJOB("JH_DHRS")
        If IsNumeric(rsJOB("JH_FTENUM")) Then dblFTEHours# = rsJOB("JH_FTENUM")
    End If
    rsJOB.Close
    
    xAsOf = dlpAsOf.Text

'
'    If Len(dlpAsOf) > 0 Then
'        xAsOf = dlpAsOf
'    Else
'        xAsOf = Format(Now, "Short Date")
'    End If
    
    'Franks Jul 31, 02 for WHSCC
    ifAnnual = False
    ifUnionDate = False
    ifFirstDate = False
    'dblServiceYears# = (DateDiff("d", varStartDate, CVDate(xAsOf)) / 365) * 12
    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    
    dblServiceYearsYTD = (DateDiff("d", varStartDate, CVDate(GetMonth("Dec") & " 1," & Year(xAsOf))) / 365) * 12 '(DateDiff("d", varStartDate, CVDate("DEC 31," & Year(xAsOf))) / 365) * 12
    If snapEntitle("ED_ORG") = "1866" And snapEntitle("ED_PT") = "FT" Then
        If snapEntitle("ED_EMP") = "PERM" Or snapEntitle("ED_EMP") = "WCB" Then
            ifAnnual = True
        End If
    End If
    If snapEntitle("ED_ORG") = "946" And snapEntitle("ED_PT") = "FT" Then
        If snapEntitle("ED_EMP") = "PERM" Or snapEntitle("ED_EMP") = "WCB" Then
            ifAnnual = True
        End If
    End If
    If snapEntitle("ED_ORG") = "NON" Then 'And snapEntitle("ED_PT") = "FT" Then
        If snapEntitle("ED_EMP") = "PERM" Or snapEntitle("ED_EMP") = "WCB" Then
            if_NON = True
            If dblServiceYearsYTD < 120 Then 'Less then 10 years, monthly, otherwise yearly
                ifAnnual = True
            End If
            ''As Heather required:
            'You can forget the Long Serv Pay folks
            'as there are only a dozen of them and As Linda and I spoke they will soon disappear form
            'the organization and all folks will use the normal vacation schedule -
            'so if NON or PHYS, FT, either PERMANENT or WORKERS COMP then use the normal vacation and
            'forget the lonf serv pay folks
            'If IsDate(snapEntitle("ED_UNION")) Then
            '    ifAnnual = True
            '    ifUnionDate = True
            'End If
            'If IsDate(snapEntitle("ED_FDAY")) Then
            '    ifAnnual = True
            '    ifFirstDate = True
            'End If
        End If
    End If
    If snapEntitle("ED_ORG") = "PHYS" Then 'And snapEntitle("ED_PT") = "FT" Then
        If snapEntitle("ED_EMP") = "PERM" Or snapEntitle("ED_EMP") = "WCB" Then
            if_NON = True
            If dblServiceYearsYTD < 120 Then 'Less then 10 years, monthly, otherwise yearly
                ifAnnual = True
            End If
            'If IsDate(snapEntitle("ED_UNION")) Then
            '    ifAnnual = True
            '    ifUnionDate = True
            'End If
            'If IsDate(snapEntitle("ED_FDAY")) Then
            '    ifAnnual = True
            '    ifFirstDate = True
            'End If
        End If
    End If
    'Franks Jul 31, 02 for WHSCC
    
    If Not ifAnnual Then
        If Not if_NON Then
'            dblServiceYears# = (DateDiff("d", varStartDate, CVDate(xAsOf)) / 365) * 12
            dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
        Else
            dblServiceYears# = dblServiceYearsYTD
        End If
        intWhereFit& = -1   ' first record can be just less than
    
        For x% = 0 To 24
            If medGTServ(x%) > 0 Then
                If dblServiceYears# >= CDbl(medLTServ(x%)) And dblServiceYears# <= CDbl(medGTServ(x%)) Then
                    intWhereFit& = x%
                    If Len(medEntitle(x%)) > 0 Then if_Entitle = True
                    If Len(medVacation(x%)) > 0 Then if_Vacation = True
                    Exit For
                End If
            End If
        Next x%
        
        If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
    
    Else 'Franks Jul 31, 02 for WHSCC
        xAsOf = CVDate(GetMonth("Jan") & " 1," & Year(xAsOf))
        dblNewEntAnn# = 0
        VacpcNAnn = 0
        intWhereFit& = 0
        For z% = 1 To 12
'            dblServiceYears# = (DateDiff("d", varStartDate, CVDate(xAsOf)) / 365) * 12
            dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))

            'If there is date of Union Date or First Day on Status/Dates screen,
            'use the special vacation rules, otherwise use the rules on the Vacation Master screen
            If Not (ifUnionDate Or ifFirstDate) Then
                For x% = 0 To 24
                    If medGTServ(x%) > 0 Then
                        If dblServiceYears# >= CDbl(medLTServ(x%)) And dblServiceYears# <= CDbl(medGTServ(x%)) Then
                            intWhereFit& = x%
                            If Len(medEntitle(x%)) > 0 Then
                                if_Entitle = True
                                dblNewEntAnn# = dblNewEntAnn# + medEntitle(x%)
                            End If
                            If Len(medVacation(x%)) > 0 Then
                                if_Vacation = True
                                VacpcNAnn = VacpcNAnn + medVacation(intWhereFit&)
                            End If
                            Exit For
                        End If
                    End If
                Next x%
            Else
                If ifUnionDate Then
                    If dblServiceYears# >= 0 And dblServiceYears# < 59.99 Then
                            if_Entitle = True
                            dblNewEntAnn# = dblNewEntAnn# + 1.25
                    End If
                    If dblServiceYears# >= 60 And dblServiceYears# < 239.99 Then
                            if_Entitle = True
                            dblNewEntAnn# = dblNewEntAnn# + 1.67
                    End If
                    If dblServiceYears# >= 240 And dblServiceYears# < 999.99 Then
                            if_Entitle = True
                            dblNewEntAnn# = dblNewEntAnn# + 2.09
                    End If
                End If
                If ifFirstDate Then
                    If dblServiceYears# >= 0 And dblServiceYears# < 11.99 Then
                            if_Entitle = True
                            dblNewEntAnn# = dblNewEntAnn# + 1.25
                    End If
                    If dblServiceYears# >= 12 And dblServiceYears# < 95.99 Then
                            if_Entitle = True
                            dblNewEntAnn# = dblNewEntAnn# + 1.67
                    End If
                    If dblServiceYears# >= 96 And dblServiceYears# < 239.99 Then
                            if_Entitle = True
                            dblNewEntAnn# = dblNewEntAnn# + 2.09
                    End If
                    If dblServiceYears# >= 240 And dblServiceYears# < 999.99 Then
                            if_Entitle = True
                            dblNewEntAnn# = dblNewEntAnn# + 2.5
                    End If
                End If
            End If
            xAsOf = DateAdd("m", 1, xAsOf)
        Next z%
    End If 'Franks Jul 31, 02 for WHSCC
    
    ' Two variables glbCompEntVac$ = "M" And glbCompEntSick$ = "M"    are 'company' level
    ' which represents if Sick and Vacation entitlements
    ' are determined on monthly basis (vs yearly) - these are stored in table hrpasco
    ' and read on system startup.
        
    ' In this routine we work independantly of SICK/VACATIon entitlement.
    '  fglbCompMonthly% - is the independant representation
        'of glbCompEntVac$ = "M" And glbCompEntSick$ = "M"
        'Procedure modUpdateSelectionWHSCC is used to set
        'fglbCompMonthly based on values it finds for global variables
        ' and what the user wants to manipulate (sick/Vac)
    
    'optD indicates if Entitlement entered is Daily or yearly based
    ' if daily then max entitlement is based on entitlement * hours they work.
    
    ' we have   Entitle = existing entitmenet (stored presently
    '           NewEntitle = amount entered onto screen = medentitle(index)
    '           EntitleUpd  = value to update record with


    If if_Entitle Then
        If ifAnnual Then
            dblNewEntitle# = dblNewEntAnn#
            If optD(intWhereFit&) = True Then           ' Entitlements entered in days
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblDHours#
                dblNewEntitle# = dblNewEntitle# * dblDHours#
                dblEntitleUpd = dblNewEntitle
            End If
            If optF(intWhereFit&) = True Then
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHours# * dblDHours#
                dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
                dblEntitleUpd = dblNewEntitle
            End If
            If fglbCompMonthly% Then
                dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
            Else
                dblEntitleUpd = dblNewEntitle ' rollover is in other utility (to accumulate)
            End If
            If dblNewMax <> 0 Then          'only do if not zero
                'If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
                '    dblEntitleUpd = dblNewMax - dblPrevEntitle#
                'End If
                'use Current instead of perious year + current to Maximum
                'ticket #3616
                If dblEntitleUpd > dblNewMax Then
                    dblEntitleUpd = dblNewMax
                End If
                'ticket #3616
            End If
        Else
            dblNewEntitle# = medEntitle(intWhereFit&)
            dblNewMax# = 0
            If optD(intWhereFit&) = True Then           ' Entitlements entered in days
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblDHours#
                dblNewEntitle# = dblNewEntitle# * dblDHours#
                dblEntitleUpd = dblNewEntitle
            End If
            If optF(intWhereFit&) = True Then
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHours# * dblDHours#
                dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
                dblEntitleUpd = dblNewEntitle
            End If
            If optH(intWhereFit&) = True Then
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&)
            End If
            If fglbCompMonthly% Then
                dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
            Else
                dblEntitleUpd = dblNewEntitle ' rollover is in other utility (to accumulate)
            End If
            
            If dblNewMax <> 0 Then          'only do if not zero
                'If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
                '    dblEntitleUpd = dblNewMax - dblPrevEntitle#
                'End If
                'use Current instead of perious year + current to Maximum
                'ticket #3616
                If dblEntitleUpd > dblNewMax Then
                    dblEntitleUpd = dblNewMax
                End If
                'ticket #3616
            End If
            
        End If
        DtTm = Now
    End If

    If if_Vacation Then
        If Not ifAnnual Then
            VacpcN = medVacation(intWhereFit&)
        Else   'Franks Jul 31, 02 for WHSCC
            VacpcN = VacpcNAnn
        End If 'Franks Jul 31, 02 for WHSCC
        VacpcO = snapEntitle("ED_VACPC")
        VED_DIV = snapEntitle("ED_DIV")
        VED_PT = snapEntitle("ED_PT")
        If IsNumeric(medVacation(intWhereFit&)) Then snapEntitle("ED_VACPC") = medVacation(intWhereFit&)
        
    End If
    If if_Entitle Then
        'Ticket #22730
        'xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd
        xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd & ". OS: " & (IIf(IsNull(snapEntitle("ED_PVAC")), 0, snapEntitle("ED_PVAC")) + IIf(IsNull(snapEntitle("ED_VAC")), 0, snapEntitle("ED_VAC"))) - IIf(IsNull(snapEntitle("ED_VACT")), 0, snapEntitle("ED_VACT"))

        'Hemu - Ticket #11925 - Changed the Accrual Date from Effective Date to Entitlement Start Date
        'because otherwise it will not update Vadim until the date arrives in case it's not same as the
        'Entitlement Start Date.
        'Call Append_Accrual(EmpNo&, "VAC", dlpAsOf, dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
        If fglbCompMonthly Then     'Ticket #22730 - Update with Effective Date if Monthly
            Call Append_Accrual(empNo&, "VAC", dlpAsOf, dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
        Else
            Call Append_Accrual(empNo&, "VAC", dlpDateRange(0), dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
        End If
        
        snapEntitle("ED_VAC") = dblEntitleUpd       ' base entitlements sic/vacation
    End If
    snapEntitle("ED_ANNVAC") = snapEntitle("ED_VAC")
    snapEntitle.Update
    
    If if_Vacation Then
        ' INSERT INTO HRAUDIT
        SQLQW1 = "INSERT INTO HRAUDIT (AU_TYPE,AU_NEWEMP,AU_EMPNBR,AU_VACPC,AU_OLDVAC, "
        SQLQW1 = SQLQW1 & "AU_DIVUPL,AU_PTUPL,AU_LDATE,AU_LTIME,AU_UPLOAD,AU_LUSER) "
        
        ' dkostka - 01/09/01 - Added Val(Format()) around vac pay %, removed quotes.  This prevents the 'data type mismatch' error.
        SQLQW1 = SQLQW1 & " VALUES('M','N'," & empNo& & "," & Val(Format(VacpcN)) & "," & Val(Format(VacpcO))
        SQLQW1 = SQLQW1 & ",'" & VED_DIV & "','" & VED_PT & "', "
        SQLQW1 = SQLQW1 & Date_SQL(Now) & " , '"
        SQLQW1 = SQLQW1 & Time$ & "', "
        SQLQW1 = SQLQW1 & "'N', "
        SQLQW1 = SQLQW1 & "'" & glbUserID & "'"
        SQLQW1 = SQLQW1 & ")"
        gdbAdoIhr001X.Execute SQLQW1
    End If

lblNextRec:
    snapEntitle.MoveNext

Wend
modUpdateSelectionWHSCC = True
MDIMain.panHelp(0).FloodType = 0
gdbAdoIhr001.CommitTrans

'fTablHREMP.Close

snapEntitle.Close

Screen.MousePointer = DEFAULT

Exit Function

modUpdateSelectionWHSCC_Err:
'These errors are:
'13=type mismatch
'94=invalid use of null
'3018=couln't find field 'item'
If Err = 13 Or Err = 94 Or Err = 3018 Then
   ' MsgBox "Err:" & Str(Err) & Chr(10) & Error$ & Chr(10) & " modUpdateSelectionWHSCC" & Chr(10) & "FORM:FUENTITL.FRM"
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

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum

If fglbNew Then
    UpdateState = NewRecord
    TF = True
    cmdPrintAll.Enabled = False
    cmdUpdate.Enabled = False
    CmdRecalc.Enabled = False
    cmdUpdateAll.Enabled = False
ElseIf Me.Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
    cmdPrintAll.Enabled = False
    cmdUpdate.Enabled = False
    CmdRecalc.Enabled = False
    cmdUpdateAll.Enabled = False
Else
    UpdateState = OPENING
    TF = True
    cmdPrintAll.Enabled = True
    cmdUpdate.Enabled = True
    CmdRecalc.Enabled = True
    cmdUpdateAll.Enabled = True
End If

Call ST_UPD_MODE(TF)

'Lanark Ticket #17711
'They keep Entitlements in GP, we import the Ent and taken,
'info:HR can not do Ent update, just use Rule to get date range
'Ticket #19782 Franks 02/03/2011 for Frontenac
If glbCompSerial = "S/N - 2172W" Or glbCompSerial = "S/N - 2410W" Then
    cmdUpdate.Enabled = False
    CmdRecalc.Enabled = False
    cmdUpdateAll.Enabled = False
End If

'Ticket #22893 - Year End for Vacation Entitlement Outstanding Based Upon <> Entitlement Date (1)
'If glbCompSerial = "S/N - 2448W" Then  'For all with Security Right
    If glbEntOutStanding$ <> "1" Then
        'chkYearEnd.Visible = True
        cmdYearEnd.Visible = GetMassUpdateSecurities("YearEnd_AnniversaryMonth_MassUpdate", glbUserID) 'True
        cmdUpdate.Enabled = Not cmdYearEnd.Visible 'True
        cmdUpdateAll.Enabled = Not cmdYearEnd.Visible 'True
    End If
'End If

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

Private Function DoWork() As Boolean
'added by Bryan 25/Oct/05 Ticket#9560
'made into a separate sub because it's called twice
Dim lastday
Dim flglastdate As Boolean
Dim blIsLast As Boolean
Dim lngRecs As Long, pct As Long, prec As Long

Screen.MousePointer = DEFAULT

DoWork = False

If glbWHSCC Then
    If Not modUpdateSelectionWHSCC() Then Exit Function
Else
    If UCase(glbCompEntVac$) = "N" Then 'Annualized Monthly
        
        'Hemu - Jerry said if the user enters As of Date as last date of the month in the
        '       Annualized Monthly calculation, then the entitl. should be calculated
        '       as of end of each month of the year. Ticket # 5880
        flglastdate = False
        lastday = MonthLastDate(CVDate(dlpAsOf.Text))
        If CVDate(dlpAsOf.Text) = CVDate(lastday) Then
            flglastdate = True
        End If
        'Hemu
        
        For fglbRunTimes = 1 To 12
            blIsLast = False
            If fglbRunTimes = 12 Then blIsLast = True
            If Not modUpdateSelection(blIsLast) Then Exit Function
            dlpAsOf = DateAdd("m", 1, CVDate(dlpAsOf.Text))
            
            'Hemu - Ticket # 5880 cont'd of above - The As Of Date created above will not
            '       be exactly last day of the month for each month when 1 month is added.
            '       e.g. 02/29/2004 will be 03/29/2004 when 1 month added.
            If flglastdate Then
                dlpAsOf.Text = CVDate(MonthLastDate(CVDate(dlpAsOf.Text)))
            End If
            'Hemu
            DoEvents
            
            'MsgBox ("Month " & fglbRunTimes & " completed")
        Next
        dlpAsOf = DateAdd("m", -12, CVDate(dlpAsOf.Text))
        DoEvents
        
    Else    'Monthly or Annual
    
        'If glbCompSerial = "S/N - 2335W" Then   'Mitchell Plastics
        '    If Not modDailyUpdateSelection(dlpDateRange(0), dlpDateRange(1), dlpAsOf, "NO") Then Exit Function
        'Else
        
            'Vacation Entitlement computation and update
            If Not modUpdateSelection() Then Exit Function
            
            'Annual Vacation computation and update for Monthly Upates only
            If fglbCompMonthly And Not (glbCompSerial = "S\N - 2355W" And chkManual.Value = 0) And (glbCompSerial <> "S/N - 2380W") Then   'Not VitalAire - Ticket #13979
                
                Call getWSQLQ("C")
                
                'Ticket #30154 - This is done here so that the Entitlement period is updated incase it has not done yet. The Annual Vacation computation uses the From Date
                'which was only getting updated at the end of this function. So by calling this function it computes the date range at this level.
                Call EntReCalcPeriod(fglbESQLQ, "VAC")
                
                If Len(clpCode(2).Text) > 0 Then
                    gdbAdoIhr001.Execute "UPDATE HREMP SET ED_ANNVAC=0 WHERE " & fglbESQLQ & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM qry_JobCurrent WHERE JB_GRPCD = '" & clpCode(2).Text & "') "
                Else
                    gdbAdoIhr001.Execute "UPDATE HREMP SET ED_ANNVAC=0 WHERE " & fglbESQLQ
                End If
                
                
                If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
                
                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
                    MDIMain.panHelp(0).FloodType = 1
                    
                    While Not snapEntitle.EOF
                        lngRecs = snapEntitle.RecordCount
                        prec = prec + 1
                        pct = Int(100 * (prec / lngRecs))
                        MDIMain.panHelp(0).FloodPercent = pct
                        Dim doDate As Date
                        doDate = dlpAsOf
                        If Not IsNull(snapEntitle("ED_EFDATE")) Then
                            'fglbAsOf = snapEntitle("ED_EFDATE")
                            fglbAsOf = IsValidDate(Format(month(snapEntitle("ED_EFDATE")) & "/" & Day(dlpAsOf) & "/" & Year(snapEntitle("ED_EFDATE")), "mm/dd/yyyy"), Day(dlpAsOf), month(snapEntitle("ED_EFDATE")), Year(snapEntitle("ED_EFDATE")))
                            'fglbAsOf = Format(month(snapEntitle("ED_EFDATE")) & "/" & Day(dlpAsOf) & "/" & Year(snapEntitle("ED_EFDATE")), "mm/dd/yyyy")
                            For fglbRunTimes = 1 To 12
                                blIsLast = False
                                If fglbRunTimes = 12 Then blIsLast = True
                                If Not modAnnSelection(blIsLast) Then Exit Function
                                fglbAsOf = DateAdd("m", 1, CVDate(fglbAsOf))
                                
                                'Hemu - Ticket # 5880 cont'd of above - The As Of Date created above will not
                                '       be exactly last day of the month for each month when 1 month is added.
                                '       e.g. 02/29/2004 will be 03/29/2004 when 1 month added.
            '                    If flglastdate Then
            '                        fglbAsOf.Text = CVDate(MonthLastDate(CVDate(dlpAsOf.Text)))
            '                    End If
                                'Hemu
                                DoEvents
                                
                                'MsgBox ("Month " & fglbRunTimes & " completed")
                            Next
                        Else
                            'DoEvents
                        End If
        '                dlpAsOf = doDate
                        snapEntitle.MoveNext
                    Wend
                    MDIMain.panHelp(0).FloodType = 0
                End If
            End If
        'End If
    End If
End If

Screen.MousePointer = HOURGLASS
Call EntReCalc(fglbESQLQ)

If Not glbSQL And Not glbOracle Then Call Pause(0.5)
DoWork = True
End Function

Private Function Round25(xNumb)
Dim xInteger, xDecimal, xDecTmp
    xInteger = Int(xNumb)
    xDecimal = xNumb - xInteger
    xDecTmp = 0
    If xDecimal >= 0 And xDecimal < 0.25 Then
        xDecTmp = 0
    End If
    If xDecimal >= 0.25 And xDecimal < 0.75 Then
        xDecTmp = 0.5
    End If
    If xDecimal >= 0.75 Then
        xDecTmp = 1
    End If
'    If xDecimal > 0 And xDecimal <= 0.5 Then
'        xDecTmp = 0.5
'    End If
'    If xDecimal > 0.5 Then
'        xDecTmp = 1
'    End If
    
    Round25 = xInteger + xDecTmp
End Function

Private Function Assign_Entitlements_Mitchell(xMonth)
    
    'New Logic - Ticket #15130 - Paid for logic - # of Days based on the month of hire
    Select Case xMonth
        Case 7: Assign_Entitlements_Mitchell = 10
        Case 1: Assign_Entitlements_Mitchell = 5
        Case 8: Assign_Entitlements_Mitchell = 9
        Case 2: Assign_Entitlements_Mitchell = 4
        Case 9: Assign_Entitlements_Mitchell = 8
        Case 3: Assign_Entitlements_Mitchell = 3
        Case 10: Assign_Entitlements_Mitchell = 7
        Case 4: Assign_Entitlements_Mitchell = 3
        Case 11: Assign_Entitlements_Mitchell = 7
        Case 5: Assign_Entitlements_Mitchell = 2
        Case 12: Assign_Entitlements_Mitchell = 6
        Case 6: Assign_Entitlements_Mitchell = 1
    End Select

End Function

Private Function Assign_Entitlements_Mitchell_MIT(xMonth)
    'New Logic for Mitchell Division - Ticket #18124 - # of Days based on the month of hire
    Select Case xMonth
        Case 7: Assign_Entitlements_Mitchell_MIT = 5
        Case 1: Assign_Entitlements_Mitchell_MIT = 10
        Case 8: Assign_Entitlements_Mitchell_MIT = 4
        Case 2: Assign_Entitlements_Mitchell_MIT = 9
        Case 9: Assign_Entitlements_Mitchell_MIT = 3
        Case 3: Assign_Entitlements_Mitchell_MIT = 8
        Case 10: Assign_Entitlements_Mitchell_MIT = 3
        Case 4: Assign_Entitlements_Mitchell_MIT = 7
        Case 11: Assign_Entitlements_Mitchell_MIT = 2
        Case 5: Assign_Entitlements_Mitchell_MIT = 7
        Case 12: Assign_Entitlements_Mitchell_MIT = 1
        Case 6: Assign_Entitlements_Mitchell_MIT = 6
    End Select

End Function

Private Sub SamuelScreenSetup() 'Ticket #23385 Franks 03/21/2013
    fraSamuelType.Left = lblSection.Left
    fraSamuelType.Visible = True
End Sub

Private Function AnniversaryMonth_YearEnd()
    'Ticket #22893 - Do Year End for Vacation Entitlement Outstanding Based Upon <> 1 and
    'with the Anniversary Month selected
    'Rollover
    'Zero Out
    'Employee's Entitlement Period change to new year
    'Update with new entitlement - not in this function but after this function
    Dim lngRecs&
    Dim Msg$, Title$, DgDef As Variant
    Dim Response%
    Dim xEmpList As String

    AnniversaryMonth_YearEnd = False
    
    'If No Anniversary Month selected then do not proceed
    If glbAnnMonth = 0 Then Exit Function
    
    On Error GoTo AnniversaryMonth_YearEnd_Err
        
    Dim rsHREmp As New ADODB.Recordset
    Dim dblOUTV#
    Dim xComments As String
    
    If Not CR_SnapEntitle_AnniversaryMonth() Then Exit Function
    
    If snapEntitle.BOF And snapEntitle.EOF Then
        'If fglbRunTimes = 1 Then
            'MsgBox "Employees for this selection do not exist!"
            MsgBox "No Employees exists for Anniversary Month Year End for this selection!"
            AnniversaryMonth_YearEnd = True
            Exit Function
        'End If
    Else
        lngRecs& = snapEntitle.RecordCount
        Msg$ = lngRecs& & " Records to process." & Chr(10) & "Would You Like To Proceed?"
        Title$ = "Year End"
        DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
        Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
        If Response% = IDNO Then    ' Evaluate response
            Exit Function
        End If
        Screen.MousePointer = HOURGLASS
    End If
    
    xEmpList = ""
    
    Do While Not snapEntitle.EOF
        'Verify the Anniversary Month
        If month(snapEntitle(fglbEntOSDate$)) = glbAnnMonth Then     'cmbAnnMonth.ListIndex Then
            dblOUTV# = 0
            
            'Rollover -------------------------------------------------------------------------------------------------
            If IsNumeric(snapEntitle("ED_PVAC")) Then
                dblOUTV# = dblOUTV# + snapEntitle("ED_PVAC")
            End If
            If IsNumeric(snapEntitle("ED_VAC")) Then
                dblOUTV# = dblOUTV# + snapEntitle("ED_VAC")
            End If
            If IsNumeric(snapEntitle("ED_VACT")) Then
                dblOUTV# = dblOUTV# - snapEntitle("ED_VACT")
            End If
            
            
            'Ticket #23141 - For Vadim clients Rolling over differently.
            'I will have to clear the balance in Vadim first, i.e. pass -ve OS Bal, so it becomes 0 balance in Vadim
            'and then pass OS to add back the OS. This will show the clear in and out in Accrual file and in Vadim.
            If glbVadim Then
                'Clear the Previous from Vadim first
                'xComments = "Vadim only: Prev. Vac. Ent. Chg " & " to 0" '& dblOUTV#
                xComments = "Vadim OS. Prev. Vac. Ent. Chg from " & dblOUTV# & " to 0" '& dblOUTV#
                Call Append_Accrual(snapEntitle("ED_EMPNBR"), "VAC", snapEntitle("ED_ETDATE"), 0 - dblOUTV#, "R", xComments)
            End If
                            
            If glbVadim Then
                'Ticket #23141 - For Vadim it is actually changing from 0 to OS amount. Add full OS back
                'after clearing above
                xComments = "Prev. Vac. Ent. Chg from 0" & " to " & dblOUTV#
                Call Append_Accrual(snapEntitle("ED_EMPNBR"), "VAC", snapEntitle("ED_ETDATE"), dblOUTV#, "R", xComments)
            Else
                'Update Accrual table
                xComments = "Prev. Vac. Ent. Chg from " & snapEntitle("ED_PVAC") & " to " & dblOUTV#
                Call Append_Accrual(snapEntitle("ED_EMPNBR"), "VAC", snapEntitle("ED_ETDATE"), dblOUTV# - Val(snapEntitle("ED_PVAC") & ""), "R", xComments)
            End If
            
            'Outstanding to Previous
            snapEntitle("ED_PVAC") = dblOUTV#
                        
            'We will have to clear the Current because there is no Zero Out for Vadim clients when doing
            'Year End as they go with the OS. Also if it's Monthly accumulation of entitlements in info:HR,
            'the new year should start with 0 current otherwise it will add to the Current. This clear out will not
            'be passed to Vadim.
            If Not glbVadim Then
                'Zero Out -------------------------------------------------------------------------------------------------
                'Update Accrual table
                xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to 0"
                Call Append_Accrual(snapEntitle("ED_EMPNBR"), "VAC", snapEntitle("ED_ETDATE"), -Val(snapEntitle("ED_VAC") & ""), "Z", xComments)
            End If
            
            'Vacation Current to 0
            snapEntitle("ED_VAC") = 0
            snapEntitle("ED_ANNVAC") = 0

            'Set New Entitlement Period -------------------------------------------------------------------------------
            snapEntitle("ED_EFDATE") = IIf(Not IsNull(snapEntitle("ED_ETDATE")), DateAdd("d", "1", CVDate(snapEntitle("ED_ETDATE"))), Null)
            snapEntitle("ED_ETDATE") = IIf(Not IsNull(snapEntitle("ED_ETDATE")), DateAdd("yyyy", "1", CVDate(snapEntitle("ED_ETDATE"))), Null)
            'snapEntitle("ED_EFDATE") = DateAdd("d", "1", CVDate(snapEntitle("ED_ETDATE")))
            'snapEntitle("ED_ETDATE") = DateAdd("yyyy", "1", CVDate(snapEntitle("ED_ETDATE")))
        
            snapEntitle("ED_LDATE") = Now
            snapEntitle("ED_LTIME") = Time$
            snapEntitle("ED_LUSER") = glbLEE_ID

            'List of employees updated to be used for Recalculate
            If Len(xEmpList) > 0 Then
                xEmpList = xEmpList & "," & snapEntitle("ED_EMPNBR")
            Else
                xEmpList = xEmpList & snapEntitle("ED_EMPNBR")
            End If
            
            snapEntitle.Update
        End If
        
        snapEntitle.MoveNext
    Loop
    snapEntitle.Close
    Set snapEntitle = Nothing

    'Recalculate the Taken
    If Len(xEmpList) > 0 Then
        Call EntReCalc(" ED_EMPNBR IN (" & xEmpList & ")", , "TAKEN ONLY")
    End If

    AnniversaryMonth_YearEnd = True
    
Exit Function

AnniversaryMonth_YearEnd_Err:

AnniversaryMonth_YearEnd = False

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "AnniversaryMonth_YearEnd", "Anniversary Year End", "Select")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
    
End Function

Private Function CR_SnapEntitle_AnniversaryMonth()
Dim SQLQ As String
Dim xEmplToIncl As String

CR_SnapEntitle_AnniversaryMonth = False

On Error GoTo CR_SnapEntitle_AnniversaryMonth_Err

Screen.MousePointer = HOURGLASS

'Ticket #24555 - Kerry's Place
'Custom logic to get list of employees to update with the monthly entitlements
If glbCompSerial = "S/N - 2433W" Then
    xEmplToIncl = KerrysPlace_EmployeesToUpdate
    SQLQ = "SELECT ED_EMPNBR,ED_VACPC,ED_PVAC,ED_VAC,ED_VACT,ED_PSICK,ED_SICK,ED_ANNVAC, ED_ANNSICK, ED_EFDATE,ED_ETDATE,"
    SQLQ = SQLQ & " ED_DIV,ED_PT, ED_SECTION,ED_SALDIST, ED_LOC, ED_ORG, ED_EMP,"
    SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME,ED_VADIM1 "
    SQLQ = SQLQ & " ,ED_SALDIST " 'Ticket #18644
    SQLQ = SQLQ & " FROM HREMP WHERE "
    If Len(xEmplToIncl) > 0 Then
        SQLQ = SQLQ & " ED_EMPNBR IN (" & xEmplToIncl & ")"
    Else
        SQLQ = SQLQ & " 1 = 2"
    End If
Else
    Call getWSQLQ("")

    'Only employees with Anniversary Month matching user input
    If cmdYearEnd.Visible = True Then
        If Len(glbAnnMonth) > 0 Then fglbESQLQ = fglbESQLQ & " AND MONTH(" & fglbEntOSDate$ & ") = " & glbAnnMonth   'cmbAnnMonth.ListIndex
        If Len(glbAnnMonth) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EFDATE < " & Date_SQL(dlpAsOf.Text)
        If Len(glbAnnMonth) > 0 Then fglbESQLQ = fglbESQLQ & " AND YEAR(ED_EFDATE) < YEAR(" & Date_SQL(dlpAsOf.Text) & ")"
    End If

    SQLQ = "SELECT ED_EMPNBR,ED_VACPC,ED_PVAC,ED_VAC,ED_VACT,ED_PSICK,ED_SICK,ED_ANNVAC, ED_ANNSICK, ED_EFDATE,ED_ETDATE,"
    SQLQ = SQLQ & " ED_DIV,ED_PT, ED_SECTION,ED_SALDIST, ED_LOC, ED_ORG, ED_EMP,"
    SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME,ED_VADIM1 "
    SQLQ = SQLQ & " ,ED_SALDIST " 'Ticket #18644
    SQLQ = SQLQ & " FROM HREMP WHERE " & fglbESQLQ
End If

If Len(clpCode(2).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_EMPNBR IN "
    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
    SQLQ = SQLQ & " WHERE JB_GRPCD = '" & clpCode(2).Text & "') "

    'Ticket #13126 Commented by Frank Jun 5th, 07
    'ElseIf glbCompSerial = "S/N - 2376W" Then 'Assembly of First Nations Bryanm 27/Apr/2006 Ticket#10735
    '    SQLQ = SQLQ & " AND ED_EMPNBR IN "
    '    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
    '    SQLQ = SQLQ & " WHERE JB_GRPCD <> 'MGT')"
End If

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    If Len(medHours.Text) > 0 Then
        SQLQ = SQLQ & " AND ED_EMPNBR IN "
        SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
        SQLQ = SQLQ & " WHERE JH_DHRS = " & medHours.Text & ") "
    End If
End If

'SQLQ = SQLQ & " AND ED_EMPNBR=2005048 " 'FOR TESTING
If snapEntitle.State <> 0 Then snapEntitle.Close
snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

CR_SnapEntitle_AnniversaryMonth = True

Screen.MousePointer = DEFAULT

Exit Function

CR_SnapEntitle_AnniversaryMonth_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapEntitle_AnniversaryMonth", "Entitlements/EMP", "Select")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

'Private Sub comAnnMonthAdding()
'    cmbAnnMonth.Clear
'    cmbAnnMonth.AddItem ""
'    cmbAnnMonth.AddItem "Jan"
'    cmbAnnMonth.AddItem "Feb"
'    cmbAnnMonth.AddItem "Mar"
'    cmbAnnMonth.AddItem "Apr"
'    cmbAnnMonth.AddItem "May"
'    cmbAnnMonth.AddItem "Jun"
'    cmbAnnMonth.AddItem "Jul"
'    cmbAnnMonth.AddItem "Aug"
'    cmbAnnMonth.AddItem "Sep"
'    cmbAnnMonth.AddItem "Oct"
'    cmbAnnMonth.AddItem "Nov"
'    cmbAnnMonth.AddItem "Dec"
'End Sub

Private Function KerrysPlace_EmployeesToUpdate()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsHREmpHis As New ADODB.Recordset
    Dim SQLQ As String
    Dim SQLQA As String
    Dim xESQLQ As String
    Dim xlstMonthF As Date
    Dim xlstMonthT As Date
    Dim xEmpToInclude As String
    Dim xEmpFoundDiv As Boolean
    Dim xEmpFoundDept As Boolean
    Dim xEmpFoundOrg As Boolean
    Dim xEmpFoundEmp As Boolean
    Dim xEmpFoundSec As Boolean
    Dim xEmpFoundLoc As Boolean
    Dim xEmpFoundPT As Boolean

    KerrysPlace_EmployeesToUpdate = ""
    xEmpToInclude = ""
    
    'Department Security
    xESQLQ = glbSeleDeptUn
    
    'Get last month's date
    'Ticket #25035 - They are already entering the Effective Date (As of Date) as previous month date so we don't
    'have to compute previous month's date/
    'xlstMonthT = MonthLastDate(DateAdd("m", -1, dlpAsOf.Text))
    xlstMonthT = dlpAsOf.Text
    xlstMonthF = CVDate(month(xlstMonthT) & "/" & "01" & "/" & Year(xlstMonthT))
            
    'List of employees from HREMP based on the Department Security
    SQLQ = "SELECT ED_EMPNBR,ED_DIV,ED_DEPTNO,ED_ORG,ED_SECTION,ED_EMP,ED_LOC,ED_PT, "
    SQLQ = SQLQ & " ED_DIVEDATE,ED_DEPTEDATE,ED_SFDATE,ED_PTEDATE"
    SQLQ = SQLQ & " FROM HREMP WHERE " & xESQLQ
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
    Do While Not rsHREmp.EOF
        xEmpFoundDiv = True
        xEmpFoundDept = True
        xEmpFoundOrg = True
        xEmpFoundEmp = True
        xEmpFoundSec = True
        xEmpFoundLoc = True
        xEmpFoundPT = True
        
        'Build query for Employee History table
        SQLQA = "SELECT TOP 1 * FROM HREMPHIS WHERE EE_EMPNBR = " & rsHREmp("ED_EMPNBR")
        SQLQ = ""
        
        'Check if employee matches the entitlement rule based on the Effective Date
        'Retrieve history based on selection criteria field populated and last month's date range on the
        'Entitlement rule
        If Len(clpDiv.Text) > 0 Then
            SQLQ = SQLQA & " AND (EE_NEWDIV IS NOT NULL)"
            SQLQ = SQLQ & " AND EE_CHGDATE <= " & Date_SQL(xlstMonthT)
            SQLQ = SQLQ & " ORDER BY EE_CHGDATE DESC"
            rsHREmpHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
            If Not rsHREmpHis.EOF Then
                If rsHREmpHis("EE_NEWDIV") = clpDiv.Text Then
                    xEmpFoundDiv = True
                Else
                    xEmpFoundDiv = False
                End If
            Else
                'Check if current value in Employee table matches in case the Employee History does not have any
                'change record for whatever reason
                If Not IsNull(rsHREmp("ED_DIV")) And rsHREmp("ED_DIV") <> "" Then
                    If Not IsNull(rsHREmp("ED_DIVEDATE")) And rsHREmp("ED_DIVEDATE") <> "" Then
                        If rsHREmp("ED_DIV") = clpDiv.Text And CVDate(rsHREmp("ED_DIVEDATE")) <= CVDate(xlstMonthT) Then
                            xEmpFoundDiv = True
                        Else
                            xEmpFoundDiv = False
                        End If
                    Else
                        If rsHREmp("ED_DIV") = clpDiv.Text Then
                            xEmpFoundDiv = True
                        Else
                            xEmpFoundDiv = False
                        End If
                    End If
                Else
                    xEmpFoundDiv = False
                End If
            End If
            rsHREmpHis.Close
            Set rsHREmpHis = Nothing
        End If
                
        If Len(clpDept.Text) > 0 Then
            SQLQ = SQLQA & " AND (EE_NEWDEPT IS NOT NULL)"
            SQLQ = SQLQ & " AND EE_CHGDATE <= " & Date_SQL(xlstMonthT)
            SQLQ = SQLQ & " ORDER BY EE_CHGDATE DESC"
            rsHREmpHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
            If Not rsHREmpHis.EOF Then
                If rsHREmpHis("EE_NEWDEPT") = clpDept.Text Then
                    xEmpFoundDept = True
                Else
                    xEmpFoundDept = False
                End If
            Else
                'Check if current value in Employee table matches in case the Employee History does not have any
                'change record for whatever reason
                If Not IsNull(rsHREmp("ED_DEPTNO")) And rsHREmp("ED_DEPTNO") <> "" Then
                    If Not IsNull(rsHREmp("ED_DEPTEDATE")) And rsHREmp("ED_DEPTEDATE") <> "" Then
                        If rsHREmp("ED_DEPTNO") = clpDept.Text And CVDate(rsHREmp("ED_DEPTEDATE")) <= CVDate(xlstMonthT) Then
                            xEmpFoundDept = True
                        Else
                            xEmpFoundDept = False
                        End If
                    Else
                        If rsHREmp("ED_DEPTNO") = clpDept.Text Then
                            xEmpFoundDept = True
                        Else
                            xEmpFoundDept = False
                        End If
                    End If
                Else
                    xEmpFoundDept = False
                End If
            End If
            rsHREmpHis.Close
            Set rsHREmpHis = Nothing
        End If
            
        If Len(clpCode(0).Text) > 0 Then
            SQLQ = SQLQA & " AND (EE_NEWORG IS NOT NULL)"
            SQLQ = SQLQ & " AND EE_CHGDATE <= " & Date_SQL(xlstMonthT)
            SQLQ = SQLQ & " ORDER BY EE_CHGDATE DESC"
            rsHREmpHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
            If Not rsHREmpHis.EOF Then
                If rsHREmpHis("EE_NEWORG") = clpCode(0).Text Then
                    xEmpFoundOrg = True
                Else
                    xEmpFoundOrg = False
                End If
            Else
                'Check if current value in Employee table matches in case the Employee History does not have any
                'change record for whatever reason
                If Not IsNull(rsHREmp("ED_ORG")) And rsHREmp("ED_ORG") <> "" Then
                    If rsHREmp("ED_ORG") = clpCode(0).Text Then
                        xEmpFoundOrg = True
                    Else
                        xEmpFoundOrg = False
                    End If
                Else
                    xEmpFoundOrg = False
                End If
            End If
            rsHREmpHis.Close
            Set rsHREmpHis = Nothing
        End If
        
        If Len(clpCode(1).Text) > 0 Then
            SQLQ = SQLQA & " AND (EE_NEWSTAT IS NOT NULL)"
            SQLQ = SQLQ & " AND EE_CHGDATE <= " & Date_SQL(xlstMonthT)
            SQLQ = SQLQ & " ORDER BY EE_CHGDATE DESC"
            rsHREmpHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
            If Not rsHREmpHis.EOF Then
                If rsHREmpHis("EE_NEWSTAT") = clpCode(1).Text Then
                    xEmpFoundEmp = True
                Else
                    xEmpFoundEmp = False
                End If
            Else
                'Check if current value in Employee table matches in case the Employee History does not have any
                'change record for whatever reason
                If Not IsNull(rsHREmp("ED_EMP")) And rsHREmp("ED_EMP") <> "" Then
                    If Not IsNull(rsHREmp("ED_SFDATE")) And rsHREmp("ED_SFDATE") <> "" Then
                        If rsHREmp("ED_EMP") = clpCode(1).Text And CVDate(rsHREmp("ED_SFDATE")) <= CVDate(xlstMonthT) Then
                            xEmpFoundEmp = True
                        Else
                            xEmpFoundEmp = False
                        End If
                    Else
                        If rsHREmp("ED_EMP") = clpCode(1).Text Then
                            xEmpFoundEmp = True
                        Else
                            xEmpFoundEmp = False
                        End If
                    End If
                Else
                    xEmpFoundEmp = False
                End If
            End If
            rsHREmpHis.Close
            Set rsHREmpHis = Nothing
        End If
        
        If Len(clpCode(3).Text) > 0 Then
            SQLQ = SQLQA & " AND (EE_NEWSECTION IS NOT NULL)"
            SQLQ = SQLQ & " AND EE_CHGDATE <= " & Date_SQL(xlstMonthT)
            SQLQ = SQLQ & " ORDER BY EE_CHGDATE DESC"
            rsHREmpHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
            If Not rsHREmpHis.EOF Then
                If rsHREmpHis("EE_NEWSECTION") = clpCode(3).Text Then
                    xEmpFoundSec = True
                Else
                    xEmpFoundSec = False
                End If
            Else
                'Check if current value in Employee table matches in case the Employee History does not have any
                'change record for whatever reason
                If Not IsNull(rsHREmp("ED_SECTION")) And rsHREmp("ED_SECTION") <> "" Then
                    If rsHREmp("ED_SECTION") = clpCode(3).Text Then
                        xEmpFoundSec = True
                    Else
                        xEmpFoundSec = False
                    End If
                Else
                    xEmpFoundSec = False
                End If
            End If
            rsHREmpHis.Close
            Set rsHREmpHis = Nothing
        End If
        
        If Len(clpCode(4).Text) > 0 Then
            SQLQ = SQLQA & " AND (EE_NEWLOC IS NOT NULL)"
            SQLQ = SQLQ & " AND EE_CHGDATE <= " & Date_SQL(xlstMonthT)
            SQLQ = SQLQ & " ORDER BY EE_CHGDATE DESC"
            rsHREmpHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
            If Not rsHREmpHis.EOF Then
                If rsHREmpHis("EE_NEWLOC") = clpCode(4).Text Then
                    xEmpFoundLoc = True
                Else
                    xEmpFoundLoc = False
                End If
            Else
                'Check if current value in Employee table matches in case the Employee History does not have any
                'change record for whatever reason
                If Not IsNull(rsHREmp("ED_LOC")) And rsHREmp("ED_LOC") <> "" Then
                    If rsHREmp("ED_LOC") = clpCode(4).Text Then
                        xEmpFoundLoc = True
                    Else
                        xEmpFoundLoc = False
                    End If
                Else
                    xEmpFoundLoc = False
                End If
            End If
            rsHREmpHis.Close
            Set rsHREmpHis = Nothing
        End If
        
        If clpPT.Text <> "" Then
            SQLQ = SQLQA & " AND (EE_NEWPT IS NOT NULL)"
            SQLQ = SQLQ & " AND EE_CHGDATE <= " & Date_SQL(xlstMonthT)
            SQLQ = SQLQ & " ORDER BY EE_CHGDATE DESC"
            rsHREmpHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
            If Not rsHREmpHis.EOF Then
                If rsHREmpHis("EE_NEWPT") = clpPT.Text Then
                    xEmpFoundPT = True
                Else
                    xEmpFoundPT = False
                End If
            Else
                'Check if current value in Employee table matches in case the Employee History does not have any
                'change record for whatever reason
                If Not IsNull(rsHREmp("ED_PT")) And rsHREmp("ED_PT") <> "" Then
                    If Not IsNull(rsHREmp("ED_PTEDATE")) And rsHREmp("ED_PTEDATE") <> "" Then
                        If rsHREmp("ED_PT") = clpPT.Text And CVDate(rsHREmp("ED_PTEDATE")) <= CVDate(xlstMonthT) Then
                            xEmpFoundPT = True
                        Else
                            xEmpFoundPT = False
                        End If
                    Else
                        If rsHREmp("ED_PT") = clpPT.Text Then
                            xEmpFoundPT = True
                        Else
                            xEmpFoundPT = False
                        End If
                    End If
                Else
                    xEmpFoundPT = False
                End If
            End If
            rsHREmpHis.Close
            Set rsHREmpHis = Nothing
        End If
        
        'Employee matches the Entitlement rule
        If xEmpFoundDiv = True And xEmpFoundDept = True And xEmpFoundOrg = True And xEmpFoundEmp = True And _
            xEmpFoundSec = True And xEmpFoundLoc = True And xEmpFoundPT = True Then
            'Add to the list of employees to update from Employee History and Employee tbale based on the
            'Entitlement Rule
            If Len(xEmpToInclude) > 0 Then
                xEmpToInclude = xEmpToInclude & ","
            End If
            xEmpToInclude = xEmpToInclude & rsHREmp("ED_EMPNBR")
        End If
        
        rsHREmp.MoveNext
    Loop
    rsHREmp.Close
    Set rsHREmp = Nothing
    
    KerrysPlace_EmployeesToUpdate = xEmpToInclude
        
End Function


Private Function OshawaPL_Vacation_Update()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsAttend As New ADODB.Recordset
    Dim rsAttPP As New ADODB.Recordset
    Dim SQLQ As String
    Dim xVacEarned As Double
    Dim xFTHsWorked As Double
    Dim xVacEarnedPT As Double
    Dim xVacEarnedFT As Double
    
    
    'For Category = PT employees
    'Get the Total Seniority Hours from HR_ATTENDANCE and HR_ATTENDANCE_HISTORY table
    SQLQ = "SELECT EMPNBR, SUM(TOT_SEN_HRS) AS TOT_SEN_HRS FROM "
    SQLQ = SQLQ & " (SELECT AD_EMPNBR AS EMPNBR, SUM(AD_HRS) AS TOT_SEN_HRS FROM HR_ATTENDANCE WHERE"
    SQLQ = SQLQ & " AD_SEN<>0 "
    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PT = 'PT')"
    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
    SQLQ = SQLQ & " UNION "
    SQLQ = SQLQ & " SELECT AH_EMPNBR AS EMPNBR, SUM(AH_HRS) AS TOT_SEN_HRS FROM HR_ATTENDANCE_HISTORY WHERE"
    SQLQ = SQLQ & " AH_SEN<>0 "
    SQLQ = SQLQ & " AND AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PT = 'PT')"
    SQLQ = SQLQ & " GROUP BY AH_EMPNBR) AS HR_ATTENDANCE"
    SQLQ = SQLQ & " GROUP BY EMPNBR"
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If Not rsAttend.EOF Then
        rsAttend.MoveFirst
                
        Do While Not rsAttend.EOF
            'Initialise
            xVacEarned = 0
            
            'Calculate employee's Seniority Hours for the Pay Period
            SQLQ = "SELECT EMPNBR, SUM(PP_SEN_HRS) AS PP_SEN_HRS FROM "
            SQLQ = SQLQ & " (SELECT AD_EMPNBR AS EMPNBR, SUM(AD_HRS) AS PP_SEN_HRS FROM HR_ATTENDANCE WHERE"
            SQLQ = SQLQ & " AD_SEN<>0 AND AD_EMPNBR = " & rsAttend("EMPNBR")
            'SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(dlpDate(0)) & " AND AD_DOA <= " & Date_SQL(dlpDate(1))
            SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT AH_EMPNBR AS EMPNBR, SUM(AH_HRS) AS PP_SEN_HRS FROM HR_ATTENDANCE_HISTORY WHERE"
            SQLQ = SQLQ & " AH_SEN<>0 AND AD_EMPNBR = " & rsAttend("EMPNBR")
            'SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(dlpDate(0)) & " AND AD_DOA <= " & Date_SQL(dlpDate(1))
            SQLQ = SQLQ & " GROUP BY AH_EMPNBR) AS HR_ATTENDANCE"
            SQLQ = SQLQ & " GROUP BY EMPNBR"
            rsAttPP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsAttPP.EOF Then
            
                'Compute Vacation Earned Hours based on employee's Total Seniority Hours and Pay Period Hours
                '<=9100 then Vacatio Earned Hours = 105/1820 * Pay Period Hours
                If rsAttend("TOT_SEN_HRS") <= 9100 Then
                    xVacEarned = (105 / 1820) * rsAttPP("PP_SEN_HRS")
                Else
                    '>9100 then Vacatio Earned Hours = 140/1820 * Pay Period Hours
                    xVacEarned = (140 / 1820) * rsAttPP("PP_SEN_HRS")
                End If
                
                'Update Employee's Vacation by Vacation Earned based on the Pay Period and Seniority Hours
                SQLQ = "SELECT ED_EMPNBR, ED_VAC FROM HREMP WHERE ED_EMPNBR = " & rsAttend("EMPNBR")
                rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsHREmp.EOF Then
                    If IsNumeric(rsHREmp("ED_VAC")) Then
                        rsHREmp("ED_VAC") = rsHREmp("ED_VAC") + xVacEarned
                    Else
                        rsHREmp("ED_VAC") = xVacEarned
                    End If
                    rsHREmp("ED_LDATE") = Now
                    rsHREmp("ED_LTIME") = Time$
                    rsHREmp("ED_LUSER") = glbLEE_ID
                    rsHREmp.Update
                End If
                rsHREmp.Close
                Set rsHREmp = Nothing
            End If
            rsAttPP.Close
            Set rsAttPP = Nothing
            
            rsAttend.MoveNext
        Loop
    End If
    rsAttend.Close
    Set rsAttend = Nothing
    
    'Initialise
    xFTHsWorked = 70
    
    'For Category = TFT employees
    SQLQ = "SELECT ED_EMPNBR, ED_VAC, ED_VACPC FROM HREMP WHERE ED_PT = 'TFT'"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsHREmp.EOF
        'Initialise
        xVacEarnedPT = 0
        xVacEarnedFT = 0
        
        If IsNumeric(rsHREmp("ED_VACPC")) Then
            xVacEarnedPT = 50 * rsHREmp("ED_VACPC")
            xVacEarnedFT = 20 * 0.04
            
            If IsNumeric(rsHREmp("ED_VAC")) Then
                rsHREmp("ED_VAC") = rsHREmp("ED_VAC") + xVacEarnedPT + xVacEarnedFT
            Else
                rsHREmp("ED_VAC") = xVacEarnedPT + xVacEarnedFT
            End If
            rsHREmp("ED_LDATE") = Now
            rsHREmp("ED_LTIME") = Time$
            rsHREmp("ED_LUSER") = glbLEE_ID
            rsHREmp.Update
        Else
            'xVacEarnedFT = 20 * 0.04
            
            'If IsNumeric(rsHREmp("ED_VAC")) Then
            '    rsHREmp("ED_VAC") = rsHREmp("ED_VAC") + xVacEarnedPT
            'Else
            '    rsHREmp("ED_VAC") = xVacEarnedPT
            'End If
            'rsHREmp("ED_LDATE") = Now
            'rsHREmp("ED_LTIME") = Time$
            'rsHREmp("ED_LUSER") = glbLEE_ID
            'rsHREmp.Update
        End If
        rsHREmp.MoveNext
    Loop
    rsHREmp.Close
    Set rsHREmp = Nothing
    
    
End Function
