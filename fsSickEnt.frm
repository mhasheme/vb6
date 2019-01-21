VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSickEnt 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Sick Entitlement Master"
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
      Left            =   0
      TabIndex        =   154
      Top             =   30
      Width           =   11415
      Begin Threed.SSCheck chkManual 
         Height          =   255
         Left            =   5520
         TabIndex        =   10
         Top             =   3165
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
         TabIndex        =   11
         Tag             =   "40-As of Date"
         Top             =   3500
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   6780
         TabIndex        =   6
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
         Left            =   1185
         TabIndex        =   2
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
         Left            =   1185
         TabIndex        =   1
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
         Left            =   1185
         TabIndex        =   0
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
         TabIndex        =   4
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
         TabIndex        =   5
         Tag             =   "EDPT-Category"
         Top             =   2070
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDPT"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   6780
         TabIndex        =   7
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
         Left            =   1185
         TabIndex        =   3
         Tag             =   "00-Enter Location Code"
         Top             =   2700
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
         Top             =   3150
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
         Top             =   3150
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1210
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
         Bindings        =   "fsSickEnt.frx":0000
         Height          =   1335
         Left            =   0
         OleObjectBlob   =   "fsSickEnt.frx":0014
         TabIndex        =   258
         Top             =   0
         Width           =   9135
      End
      Begin Threed.SSCheck chkRound 
         Height          =   255
         Left            =   5640
         TabIndex        =   261
         Top             =   3515
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
      Begin VB.Label lblPeriod 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sick Entitlement Period"
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
         TabIndex        =   257
         Top             =   3150
         Width           =   1635
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
         TabIndex        =   168
         Top             =   1800
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
         TabIndex        =   167
         Top             =   2100
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
         TabIndex        =   166
         Top             =   2430
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
         Left            =   5280
         TabIndex        =   165
         Top             =   1800
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
         TabIndex        =   164
         Top             =   3545
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
         TabIndex        =   163
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
         Left            =   5280
         TabIndex        =   162
         Top             =   2400
         Width           =   1260
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
         TabIndex        =   161
         Top             =   4170
         Width           =   2370
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
         TabIndex        =   160
         Top             =   4170
         Width           =   960
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
         TabIndex        =   159
         Top             =   4170
         Width           =   660
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
         TabIndex        =   158
         Top             =   1560
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
         Left            =   5280
         TabIndex        =   157
         Top             =   2100
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
         Left            =   5280
         TabIndex        =   156
         Top             =   2700
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
         TabIndex        =   155
         Top             =   2730
         Width           =   615
      End
   End
   Begin VB.VScrollBar scrControl 
      Height          =   4125
      LargeChange     =   315
      Left            =   10800
      Max             =   100
      SmallChange     =   315
      TabIndex        =   136
      Top             =   4710
      Width           =   300
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   645
      Left            =   0
      TabIndex        =   259
      Top             =   10305
      Width           =   11760
      _Version        =   65536
      _ExtentX        =   20743
      _ExtentY        =   1138
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
         Left            =   5400
         TabIndex        =   118
         Top             =   120
         Width           =   1665
      End
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update Entitlement"
         Height          =   375
         Left            =   1560
         TabIndex        =   116
         Tag             =   "Change all matching records to the above"
         Top             =   120
         Width           =   1905
      End
      Begin VB.CommandButton CmdRecalc 
         Appearance      =   0  'Flat
         Caption         =   "R&ecalculate"
         Height          =   375
         Left            =   3600
         TabIndex        =   117
         Tag             =   "Recalculation"
         Top             =   120
         Width           =   1665
      End
      Begin VB.CommandButton cmdPrintAll 
         Appearance      =   0  'Flat
         Caption         =   "Print &All"
         Height          =   375
         Left            =   240
         TabIndex        =   115
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
      Left            =   60
      TabIndex        =   260
      Top             =   4440
      Width           =   11000
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   12
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   13
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
         TabIndex        =   19
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
         TabIndex        =   20
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   24
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
         Left            =   3270
         TabIndex        =   21
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
         TabIndex        =   25
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
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   0
         Left            =   7050
         TabIndex        =   18
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
         TabIndex        =   22
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
         TabIndex        =   26
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
         TabIndex        =   14
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
         TabIndex        =   31
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
         TabIndex        =   32
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
         TabIndex        =   35
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
         TabIndex        =   36
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
         TabIndex        =   33
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
         TabIndex        =   37
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
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   3
         Left            =   7050
         TabIndex        =   30
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
         TabIndex        =   34
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
         TabIndex        =   38
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
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   6
         Left            =   0
         TabIndex        =   39
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
         TabIndex        =   40
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
         TabIndex        =   43
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
         TabIndex        =   44
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
         TabIndex        =   47
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
         TabIndex        =   48
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
         TabIndex        =   45
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
         TabIndex        =   49
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
         TabIndex        =   42
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
         TabIndex        =   46
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
         TabIndex        =   50
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
         TabIndex        =   41
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
         TabIndex        =   51
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
         TabIndex        =   52
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
         TabIndex        =   53
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
         TabIndex        =   54
         Tag             =   "10-Maximum Entitlement can be"
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   10
         Left            =   0
         TabIndex        =   55
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
         TabIndex        =   56
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
         TabIndex        =   59
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
         TabIndex        =   60
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
         TabIndex        =   57
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
         TabIndex        =   61
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
         TabIndex        =   58
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
         TabIndex        =   62
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
         TabIndex        =   63
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
         TabIndex        =   64
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
         TabIndex        =   67
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
         TabIndex        =   68
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
         TabIndex        =   71
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
         TabIndex        =   72
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
         TabIndex        =   65
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
         TabIndex        =   69
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
         TabIndex        =   66
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
         TabIndex        =   70
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
         TabIndex        =   74
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
         TabIndex        =   73
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   15
         Left            =   0
         TabIndex        =   75
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
         TabIndex        =   76
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
         TabIndex        =   79
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
         TabIndex        =   80
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
         TabIndex        =   77
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
         TabIndex        =   78
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
         TabIndex        =   82
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
         TabIndex        =   81
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
      Begin Threed.SSFrame frmDH 
         Height          =   375
         Index           =   0
         Left            =   4300
         TabIndex        =   137
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
            TabIndex        =   17
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
               Size            =   30
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
            TabIndex        =   16
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
               Size            =   30
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
            TabIndex        =   15
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
         TabIndex        =   138
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
            TabIndex        =   139
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
            TabIndex        =   140
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
            TabIndex        =   141
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
         TabIndex        =   142
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
            TabIndex        =   143
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
            TabIndex        =   144
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
            TabIndex        =   145
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
         Index           =   3
         Left            =   4300
         TabIndex        =   146
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
            TabIndex        =   147
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
            TabIndex        =   148
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
            TabIndex        =   149
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
         TabIndex        =   150
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
            TabIndex        =   151
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
            TabIndex        =   152
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
            TabIndex        =   153
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   17
         Left            =   0
         TabIndex        =   83
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   17
         Left            =   2160
         TabIndex        =   84
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
         TabIndex        =   85
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
         TabIndex        =   86
         Tag             =   "10-Maximum Entitlement can be"
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   18
         Left            =   0
         TabIndex        =   87
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
         TabIndex        =   88
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
         TabIndex        =   91
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   19
         Left            =   2160
         TabIndex        =   92
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
         TabIndex        =   89
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
         TabIndex        =   93
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
         TabIndex        =   90
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
         TabIndex        =   94
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
         TabIndex        =   95
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   20
         Left            =   2160
         TabIndex        =   96
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
         TabIndex        =   99
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   21
         Left            =   2160
         TabIndex        =   100
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
         TabIndex        =   103
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   22
         Left            =   2160
         TabIndex        =   104
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
         TabIndex        =   97
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
         TabIndex        =   101
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
         TabIndex        =   98
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
         TabIndex        =   102
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
         TabIndex        =   106
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
         TabIndex        =   105
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   23
         Left            =   0
         TabIndex        =   107
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   23
         Left            =   2160
         TabIndex        =   108
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
         TabIndex        =   111
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   24
         Left            =   2160
         TabIndex        =   112
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
         TabIndex        =   109
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
         TabIndex        =   110
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
         TabIndex        =   114
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
         TabIndex        =   113
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
         Left            =   4290
         TabIndex        =   177
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
            Index           =   5
            Left            =   930
            TabIndex        =   179
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
            TabIndex        =   180
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
         Left            =   4290
         TabIndex        =   181
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
            TabIndex        =   182
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
            TabIndex        =   183
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
            TabIndex        =   184
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
         Left            =   4290
         TabIndex        =   185
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
            TabIndex        =   186
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
            TabIndex        =   187
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
            TabIndex        =   188
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
         Left            =   4290
         TabIndex        =   189
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
            TabIndex        =   190
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
            TabIndex        =   191
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
            TabIndex        =   192
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
         Left            =   4290
         TabIndex        =   193
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
            TabIndex        =   194
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
            TabIndex        =   195
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
            TabIndex        =   196
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
         Left            =   4290
         TabIndex        =   197
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
            TabIndex        =   198
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
            TabIndex        =   199
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
            TabIndex        =   200
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
         Left            =   4290
         TabIndex        =   201
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
            TabIndex        =   202
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
            TabIndex        =   203
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
            TabIndex        =   204
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
         Left            =   4290
         TabIndex        =   205
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
            TabIndex        =   206
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
            TabIndex        =   207
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
            TabIndex        =   208
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
         Left            =   4290
         TabIndex        =   209
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
            TabIndex        =   210
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
            TabIndex        =   211
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
            TabIndex        =   212
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
         Left            =   4290
         TabIndex        =   213
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
            TabIndex        =   214
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
            TabIndex        =   215
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
            TabIndex        =   216
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
         Left            =   4290
         TabIndex        =   217
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
            TabIndex        =   218
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
            TabIndex        =   219
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
            TabIndex        =   220
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
         Left            =   4290
         TabIndex        =   221
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
            TabIndex        =   222
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
            TabIndex        =   223
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
            TabIndex        =   224
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
         Left            =   4290
         TabIndex        =   225
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
            TabIndex        =   226
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
            TabIndex        =   227
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
            TabIndex        =   228
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
         Left            =   4290
         TabIndex        =   229
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
            TabIndex        =   230
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
            TabIndex        =   231
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
            TabIndex        =   232
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
         Left            =   4290
         TabIndex        =   233
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
            TabIndex        =   234
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
            TabIndex        =   235
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
            TabIndex        =   236
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
         Left            =   4290
         TabIndex        =   237
         Top             =   6510
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
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
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   20
            Left            =   1770
            TabIndex        =   238
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
               Size            =   27
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
            TabIndex        =   239
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
               Size            =   27
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
            TabIndex        =   240
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
         Left            =   4290
         TabIndex        =   241
         Top             =   6825
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
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
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   21
            Left            =   1770
            TabIndex        =   242
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
               Size            =   27
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
            TabIndex        =   243
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
               Size            =   27
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
            TabIndex        =   244
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
               Size            =   27
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
         Left            =   4290
         TabIndex        =   245
         Top             =   7155
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
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
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   22
            Left            =   1770
            TabIndex        =   246
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
               Size            =   27
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
            TabIndex        =   247
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
               Size            =   27
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
            TabIndex        =   248
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
               Size            =   27
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
         Left            =   4290
         TabIndex        =   249
         Top             =   7485
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
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
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   23
            Left            =   1770
            TabIndex        =   250
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
               Size            =   27
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
            TabIndex        =   251
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
               Size            =   27
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
            TabIndex        =   252
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
         Left            =   4290
         TabIndex        =   253
         Top             =   7815
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         ForeColor       =   0
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
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   24
            Left            =   1770
            TabIndex        =   254
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
               Size            =   27
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
            TabIndex        =   255
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
               Size            =   27
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
            TabIndex        =   256
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
               Size            =   27
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
         TabIndex        =   176
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
         TabIndex        =   175
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
         TabIndex        =   174
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
         TabIndex        =   173
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
         TabIndex        =   172
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
         TabIndex        =   171
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
         TabIndex        =   170
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
         TabIndex        =   169
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
         TabIndex        =   135
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
         TabIndex        =   134
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
         TabIndex        =   133
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
         TabIndex        =   132
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
         TabIndex        =   131
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
         TabIndex        =   130
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
         TabIndex        =   129
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
         TabIndex        =   128
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
         TabIndex        =   127
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
         TabIndex        =   126
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
         TabIndex        =   125
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
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   122
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
         TabIndex        =   121
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
         TabIndex        =   120
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
         TabIndex        =   119
         Top             =   4920
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmSickEnt"
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
Dim OManual

Dim fglbESQLQ, fglbVSQLQ
Dim fglbNew As Boolean
Dim fglbRunTimes
Dim Memplist1, Memplist2
Dim orgEffDate

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

'Sam 02/02/2006
If Len(dlpDateRange(0).Text) > 0 Then
    If Not IsDate(dlpDateRange(0).Text) Then
        MsgBox "Invalid Sick Entitlement Period From Date"
        dlpDateRange(0).SetFocus
        Exit Function
    End If
Else
    'Only Mandatory if based on Entitlement Date
    If glbEntOutStandingS$ = "1" Then
        MsgBox "Sick Entitlement Period From Date is Mandatory field"
        dlpDateRange(0).SetFocus
        Exit Function
    End If
End If

If Len(dlpDateRange(1).Text) > 0 Then
    If Not IsDate(dlpDateRange(1).Text) Then
        MsgBox "Invalid Sick Entitlement Period To Date"
        dlpDateRange(1).SetFocus
        Exit Function
    End If
Else
    'Only Mandatory if based on Entitlement Date
    If glbEntOutStandingS$ = "1" Then
        MsgBox "Sick Entitlement Period To Date is Mandatory field"
        dlpDateRange(1).SetFocus
        Exit Function
    End If
End If
'Sam 02/02/2006

If IsDate(dlpDateRange(0).Text) And IsDate(dlpDateRange(1).Text) Then
If CVDate(dlpDateRange(0).Text) > CVDate(dlpDateRange(1).Text) Then
    MsgBox "Sick Entitlement Period From Date cannot be greater than Sick Entitlement Period To Date"
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
    'If UCase(glbCompEntSick$) = "A" Then
    '    If glbLinamar Then
            MsgBox "Effective Date is required field"
            dlpAsOf.SetFocus
            Exit Function
    '    End If
    'End If
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

orgEffDate = dlpAsOf.Text

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
Msg = Msg & Chr(10) & "The Sick Entitlement Rules?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

Call getWSQLQ("C")
SQLQ = "DELETE FROM HRSICKENT WHERE " & fglbVSQLQ

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

Data1.Refresh
Call Display_Value

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
'Sam 02/02/2006
OFromDate = dlpDateRange(0).Text
OToDate = dlpDateRange(1).Text
'Sam 02/02/2006

OLoc = clpCode(4).Text
OSection = clpCode(3).Text
oEMP = clpCode(1).Text
oEmpMode = clpPT.Text
oGRPCE = clpCode(2).Text
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

If glbWHSCC Then
    Call SetSickRules
End If

Call SET_UP_MODE

clpDiv.SetFocus
End Sub

Private Sub SetSickRules() 'Whscc only
Dim x
    For x = 0 To 24
        medLTServ(x) = ""
        medGTServ(x) = ""
        medEntitle(x) = ""
        optD(x) = False
        optH(x) = False
        optF(x) = True
        medMax(x) = ""
        'medVacation(x) = ""
    Next
    medLTServ(0) = 0
    medGTServ(0) = 999
    medEntitle(0) = 1.5
    medMax(0) = 240
End Sub

Sub cmdOK_Click()
Dim x%, Y%, xUnion, xPT, SQLQ, SQLQW
Dim xStr
Dim rsVE As New ADODB.Recordset
Dim rsVT As New ADODB.Recordset
Dim glbiOneWhere As Boolean
Dim bmk As Variant

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
    SQLQ = "DELETE FROM HRSICKENT WHERE " & fglbVSQLQ
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
Else
    Call getWSQLQ("C")
    SQLQ = "SELECT * FROM HRSICKENT WHERE " & fglbVSQLQ
    rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsVT.EOF Then
        MsgBox "You can not add duplicate record"
         clpDiv.SetFocus
        Exit Sub
    End If
End If

gdbAdoIhr001.BeginTrans
SQLQ = "SELECT * FROM HRSICKENT"
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
'Ticket #12467
'        If UCase(glbCompEntSick$) = "A" Then
'            If Len(dlpAsOf.Text) > 0 Then
                rsVE("VE_EDATE") = dlpAsOf.Text
'            End If
'        Else
'            rsVE("VE_EDATE") = Null
'        End If
        
        'sam 02/02/2006
        If Len(dlpDateRange(0).Text) > 0 Then
            rsVE("VE_FRDATE") = dlpDateRange(0).Text
        End If
        If Len(dlpDateRange(1).Text) > 0 Then
            rsVE("VE_TODATE") = dlpDateRange(1).Text
        End If
        'sam 02/02/2006
        
        rsVE("VE_GRPCD_TABL") = "JBGC"
        rsVE("VE_GRPCD") = clpCode(2).Text
        If glbFrench Then
            rsVE("VE_BMONTH") = Replace(medLTServ(x%), ",", ".")
        Else
            rsVE("VE_BMONTH") = medLTServ(x%)
        End If
        If glbFrench Then
            rsVE("VE_EMONTH") = Replace(medGTServ(x%), ",", ".")
        Else
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
        rsVE("VE_MANUAL") = chkManual.Value
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

orgEffDate = dlpAsOf.Text

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

If Err.Number = -2147217887 Then '01/01/1200 can cause this error Ticket #18227
    MsgBox "    Invalid Date!    "
    gdbAdoIhr001.RollbackTrans
    Exit Sub
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdOK", "SICK ENTITLEMENTS", "UPDATE")
    Unload Me
End If

End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport, x%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%
'cmdPrint.Enabled = False

Me.vbxCrystal.Reset
Me.vbxCrystal.WindowTitle = "Sick Entitlement Master Report"
Call setRptLabel(Me, 0) '1)
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To 5
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next
End If
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgsickent.rpt"

SQLQ = "(1=1) "
If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_DIV} = '" & clpDiv.Text & "'"
If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_DEPT} = '" & clpDept.Text & "'"
If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_ORG} = '" & clpCode(0).Text & "'"
If Len(dlpAsOf.Text) > 0 Then
    dtYYY% = Year(dlpAsOf.Text)
    dtMM% = month(dlpAsOf.Text)
    dtDD% = Day(dlpAsOf.Text)
    SQLQ = SQLQ & " AND {HRSICKENT.VE_EDATE} = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
End If
If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_EMP} = '" & clpCode(1).Text & "'"
If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_PT} = '" & clpPT.Text & "' "
If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_GRPCD} = '" & clpCode(2).Text & "'"
If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_SECTION} = '" & clpCode(3).Text & "'"
If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_LOC} = '" & clpCode(4).Text & "'"

'sam 02/02/2006
If Len(dlpDateRange(0).Text) > 0 Then
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND {HRSICKENT.VE_FRDATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
End If
If Len(dlpDateRange(1).Text) > 0 Then
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    SQLQ = SQLQ & " AND {HRSICKENT.VE_TODATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
End If
'sam 02/02/2006

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
Me.vbxCrystal.WindowTitle = "Sick Entitlement Master Report"
Call setRptLabel(Me, 0) '1)
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To 5
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next
End If
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgsickent.rpt"

SQLQ = "(1=1) "
If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_DIV} = '" & clpDiv.Text & "'"
If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_DEPT} = '" & clpDept.Text & "'"
If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_ORG} = '" & clpCode(0).Text & "'"
If Len(dlpAsOf.Text) > 0 Then
    dtYYY% = Year(dlpAsOf.Text)
    dtMM% = month(dlpAsOf.Text)
    dtDD% = Day(dlpAsOf.Text)
    SQLQ = SQLQ & " AND {HRSICKENT.VE_EDATE} = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
End If
If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_EMP} = '" & clpCode(1).Text & "'"
If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_PT} = '" & clpPT.Text & "' "
If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_GRPCD} = '" & clpCode(2).Text & "'"
If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_SECTION} = '" & clpCode(3).Text & "'"
If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_LOC} = '" & clpCode(4).Text & "'"

'sam 02/02/2006
If Len(dlpDateRange(0).Text) > 0 Then
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND {HRSICKENT.VE_FRDATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
End If
If Len(dlpDateRange(1).Text) > 0 Then
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    SQLQ = SQLQ & " AND {HRSICKENT.VE_TODATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
End If

'sam 02/02/2006


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

Me.vbxCrystal.WindowTitle = "Sick Entitlement Master Report"
Call setRptLabel(Me, 0) '1)
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To 5
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next
End If
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgsickent.rpt"
Me.vbxCrystal.Action = 1

cmdPrintAll.Enabled = True
End Sub

Private Sub cmdRecalc_Click()
Dim lastday
Dim flglastdate As Boolean
Dim lngRecs As Long, pct As Long, prec As Long
Dim doDate As Date
Dim bmk As Variant
Dim blIsLast As Boolean

On Error GoTo Eh

bmk = Data1.Recordset.Bookmark
Screen.MousePointer = vbHourglass

Call getWSQLQ("C")
Call EntReCalcPeriod(fglbESQLQ, "SICK", , , dlpDateRange(0), dlpDateRange(1))



If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
        Data1.Recordset.MoveFirst
        Do
            Call Display_Value
            
            If Len(dlpAsOf.Text) = 0 Then
                MsgBox "Effective Date is required field"
                dlpAsOf.SetFocus
                GoTo exH
            End If
            
            If (fglbCompMonthly Or UCase(glbCompEntVac$) = "N") And Not (glbCompSerial = "S/N - 2355W" And chkManual.Value = -1) Then
                prec = 0
                Call getWSQLQ("C")
                
                gdbAdoIhr001.Execute "UPDATE HREMP SET ED_ANNSICK=0 WHERE " & fglbESQLQ
                
                If Not CR_SnapEntitle() Then Exit Sub  ' create snapEntitle (form level recordset)
                
                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
                    While Not snapEntitle.EOF
                        lngRecs = snapEntitle.RecordCount
                        prec = prec + 1
                        pct = Int(100 * (prec / lngRecs))
                        MDIMain.panHelp(0).FloodPercent = pct
                        
                        doDate = dlpAsOf
                        'fglbAsOf = snapEntitle("ED_EFDATES")
                        
                        If IsNull(snapEntitle("ED_EFDATES")) Then GoTo nextEmp
                        
                        fglbAsOf = IsValidDate(Format(month(snapEntitle("ED_EFDATES")) & "/" & Day(dlpAsOf) & "/" & Year(snapEntitle("ED_EFDATES")), "mm/dd/yyyy"), Day(dlpAsOf), month(snapEntitle("ED_EFDATES")), Year(snapEntitle("ED_EFDATES")))
                        
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
                gdbAdoIhr001.Execute "UPDATE HREMP SET ED_ANNSICK=0 WHERE " & fglbESQLQ
                If Not CR_SnapEntitle() Then Exit Sub  ' create snapEntitle (form level recordset)
                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
                    While Not snapEntitle.EOF
                        lngRecs = snapEntitle.RecordCount
                        prec = prec + 1
                        pct = Int(100 * (prec / lngRecs))
                        MDIMain.panHelp(0).FloodPercent = pct

                        doDate = dlpAsOf
                        fglbAsOf = snapEntitle("ED_EFDATES")
            
                        If Not modAnnSelection(True) Then Exit Sub
                        DoEvents
                            
                        snapEntitle.MoveNext
                    Wend
                    MDIMain.panHelp(0).FloodType = 0
                End If
            
            End If
            Data1.Recordset.MoveNext
        Loop Until Data1.Recordset.EOF
    End If
    Screen.MousePointer = vbDefault
    Data1.Recordset.Bookmark = bmk
    Call Display_Value
    
exH:
    Screen.MousePointer = vbDefault
    Exit Sub
Eh:
    
    Resume exH
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo Mod_Err
Dim sFlag As Boolean

If Not gSec_Upd_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

If Not chkMUEntitle() Then Exit Sub
    'Ticket #19632 - This is becsuse they are using TAKEN as part of Max checking. So when the date range is
    'changed the TAKEN should be recalculated so on Update Entitle, the correct TAKEN is used in the formula.
    'During Year End, on the date range is changed, saved and Update Entitlement is clicked, the TAKEN of last
    'year is still there in ED_SICT and that was being used in the Max comparison formula. This recalculate
    'will fix the issue by recalculating the TAKEN.
    If glbCompSerial = "S/N - 2173W" Or glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2389W" Or _
        glbCompSerial = "S/N - 2408W" Or glbCompSerial = "S/N - 2412W" Or glbCompSerial = "S/N - 2399W" Or _
        glbCompSerial = "S/N - 2395W" Or glbCompSerial = "S/N - 2430W" Or glbCompSerial = "S/N - 2190W" Or _
        glbCompSerial = "S/N - 2450W" Or glbCompSerial = "S/N - 2436W" Or glbCompSerial = "S/N - 2466W" Or _
        glbCompSerial = "S/N - 2234W" Then
        
        Call getWSQLQ("C")
        Call EntReCalcPeriod(fglbESQLQ, "SICK", , , dlpDateRange(0), dlpDateRange(1))
        Call EntReCalc(fglbESQLQ)
    End If

    'Added by Bryan 25/Oct/05 Ticket#9560
    'made the code a separate sub because it's being used in two places
    sFlag = DoWork

Data1.Refresh
Call Display_Value

orgEffDate = dlpAsOf.Text

If sFlag Then
    MsgBox "Update Completed Successfully.", vbInformation + vbOKOnly, "Sick Entitlements"
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

Private Function GetFTEtot(empNo, dblFTE)
Dim rsFTE As New ADODB.Recordset
Dim SQLQ, xFte
    xFte = dblFTE
    If glbMulti Then
        If Len(Memplist1) > 0 Then
            If InStr(1, Memplist1, "'" & empNo & "'") > 0 Then 'this EmpNo is in Memplist1
                SQLQ = "SELECT JH_EMPNBR, SUM(JH_FTENUM) AS TOTFTE FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & empNo & " "
                SQLQ = SQLQ & "GROUP BY JH_EMPNBR "
                rsFTE.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsFTE.EOF Then
                    If Not IsNull(rsFTE("TOTFTE")) Then
                        xFte = rsFTE("TOTFTE")
                    End If
                End If
                rsFTE.Close
            End If
        End If
    End If
    GetFTEtot = xFte
End Function

Private Function AccuValForMulti(empNo, dblEnt) ' Ticket #3304
'For multi positions and annual update, accumulate all entitlement of positions together
'and then replace the entitlement.
Dim xVal
    xVal = 0
    If glbMulti Then
        If Len(Memplist1) > 0 Then
            If InStr(1, Memplist1, "'" & empNo & "'") > 0 Then 'this EmpNo is in Memplist1
                If InStr(1, Memplist2, "'" & empNo & "'") > 0 Then 'this EmpNo is in Memplist2
                    'xVal = 0 ' First time replace the Emtitlement with the New one
                    Memplist2 = Replace(Memplist2, "'" & empNo & "',", ",")
                Else
                    xVal = dblEnt 'from Second time, accumulate the entitlement
                End If
            End If
        End If
    End If
    AccuValForMulti = xVal
End Function

Private Function CalcASLRepaid(xEmpNo, xAsofDate, dblEntUpd, dblNewEnt, dblEnt#) '
Dim rsASL As New ADODB.Recordset
Dim rsENT As New ADODB.Recordset
Dim SQLQ, xTaken, xRepaid, xOutStand
Dim xSickEnt

'Hemu
'    Dim tmpTestData As String
'    Dim tmpAdoIHRTest As String
'    Dim sSetting As String
'    Dim sPath1 As String
'    Dim giGar1 As Integer
'
'    sPath1 = REG_NAME & "INFOHR Files"
'
'    sSetting = "IHRREPORTS"  'Compressed database location
'    tmpTestData = glbWorkDir
'    giGar1 = bGetRegistrySetting(lCurrentKey, sPath1, sSetting, tmpTestData)
'    tmpTestData = tmpTestData & IIf(Right$(tmpTestData, 1) <> "\", "\", "") & "TestData.mdb"
'    tmpAdoIHRTest = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=petman;Data Source=" & tmpTestData
'Hemu


    xSickEnt = dblEntUpd
    SQLQ = "SELECT ED_EMPNBR,ED_EFDATES,ED_ETDATES,ED_SICK,ED_SICKT FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    rsENT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsENT.EOF Then
        If IsDate(rsENT("ED_EFDATES")) And IsDate(rsENT("ED_ETDATES")) Then
            'If rsENT("ED_EFDATES") <= CVDate(xAsofDate) And rsENT("ED_ETDATES") >= CVDate(xAsofDate) Then
                SQLQ = "SELECT AS_EMPNBR, Sum(AS_HRSTAK) AS TAKEN, Sum(AS_HRSREP) AS REPAID FROM WHSCC_ASL WHERE AS_EMPNBR = " & xEmpNo & " "
                'Don't check Date Range for ASL T#3304
                'SQLQ = SQLQ & "AND AS_EFDATES = ('" & Format(rsENT("ED_EFDATES"), "mmm dd,yyyy") & "') "
                'SQLQ = SQLQ & "AND AS_ETDATES = ('" & Format(rsENT("ED_ETDATES"), "mmm dd,yyyy") & "') "
                SQLQ = SQLQ & "GROUP BY AS_EMPNBR "
                rsASL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                xTaken = 0: xRepaid = 0: xOutStand = 0
                If Not rsASL.EOF Then
                    If IsNull(rsASL("TAKEN")) Then
                        xTaken = 0
                    Else
                        xTaken = rsASL("TAKEN")
                    End If
                    If IsNull(rsASL("REPAID")) Then
                        xRepaid = 0
                    Else
                        xRepaid = rsASL("REPAID")
                    End If
                    xOutStand = xTaken - xRepaid
                End If
                rsASL.Close
                
                'Logic changed:
                'Repaid = Sick Entitlement, before Repaid = ASL Outstanding
                
                'xOutStand = dblEntUpd
                If xOutStand > 0 Then
                    If xOutStand >= dblNewEnt Then
                        xSickEnt = dblEnt#
                    Else
                        xSickEnt = dblEnt# + dblNewEnt - xOutStand
                        dblNewEnt = xOutStand
                    End If
'Hemu
'If glbWHSCC Then
''include the dummy test table here
'    SQLQ = "SELECT * FROM WHSCC_ASL WHERE AS_EMPNBR = " & xEmpNo & " "
'    rsASL.Open SQLQ, tmpAdoIHRTest, adOpenKeyset, adLockOptimistic
'    rsASL.AddNew
'    rsASL("AS_HRSTAK") = 0
'    rsASL("AS_COMPNO") = "001"
'    rsASL("AS_EMPNBR") = xEmpNo
'    rsASL("AS_DOA") = xAsofDate
'    rsASL("AS_CODE") = "REPA"
'    rsASL("AS_HRSREP") = dblNewEnt 'dblEntUpd
'    rsASL("AS_HRSOS") = xOutStand - dblNewEnt 'dblEntUpd
'    rsASL("AS_LDATE") = Format(Now, "SHORT DATE")
'    rsASL("AS_LTIME") = Time$
'    rsASL("AS_LUSER") = glbUserID
'    rsASL.Update
'    rsASL.Close
'
'    GoTo End_Test_Data
'End If
'Hemu
                    SQLQ = "SELECT * FROM WHSCC_ASL WHERE AS_EMPNBR = " & xEmpNo & " "
                    'Don't check Date Range for ASL T#3304
                    'SQLQ = SQLQ & "AND AS_EFDATES = ('" & Format(rsENT("ED_EFDATES"), "mmm dd,yyyy") & "') "
                    'SQLQ = SQLQ & "AND AS_ETDATES = ('" & Format(rsENT("ED_ETDATES"), "mmm dd,yyyy") & "') "
                    SQLQ = SQLQ & "AND AS_DOA = ('" & Format(xAsofDate, "mmm dd,yyyy") & "') "
                    SQLQ = SQLQ & "AND AS_CODE = 'REPA' "
                    rsASL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    'If rsASL.EOF Then
                        rsASL.AddNew
                        rsASL("AS_HRSTAK") = 0
                    'Else
                        'dblEntUpd = dblEntUpd + rsASL("AS_HRSREP")
                    'End If
                    rsASL("AS_COMPNO") = "001"
                    rsASL("AS_EMPNBR") = xEmpNo
                    rsASL("AS_DOA") = xAsofDate
                    rsASL("AS_CODE") = "REPA"
                    rsASL("AS_HRSREP") = dblNewEnt 'dblEntUpd
                    rsASL("AS_HRSOS") = xOutStand - dblNewEnt 'dblEntUpd
                    'rsASL("AS_EFDATES") = rsENT("ED_EFDATES")
                    'rsASL("AS_ETDATES") = rsENT("ED_ETDATES")
                    rsASL("AS_LDATE") = Format(Now, "SHORT DATE")
                    rsASL("AS_LTIME") = Time$
                    rsASL("AS_LUSER") = glbUserID
                    rsASL.Update
                    rsASL.Close
                    Call ReCalcASL(xEmpNo, "")
                    'SQLQ = "UPDATE WHSCC_ASL SET AS_HRAOS = 0 "
                    'SQLQ = SQLQ & "WHERE AS_EMPNBR = " & xEmpNo & " "
                    'SQLQ = SQLQ & "AND AS_EFDATES = ('" & Format(rsENT("ED_EFDATES"), "mmm dd,yyyy") & "') "
                    'SQLQ = SQLQ & "AND AS_ETDATES = ('" & Format(rsENT("ED_ETDATES"), "mmm dd,yyyy") & "') "
                    'gdbAdoIhr001.Execute SQLQ
                End If
            'End If
        End If
    End If

'Hemu
'exit after update to dummy table
'End_Test_Data:
'    'Update Test_data table
'    Dim rsTestData As New ADODB.Recordset
'    SQLQ = "SELECT * FROM test_data"
'    rsTestData.Open SQLQ, tmpAdoIHRTest, adOpenKeyset, adLockOptimistic
'    rsTestData.AddNew
'    rsTestData("ED_EMPNBR") = xEmpNo
'    rsTestData("JH_DHRS") = tmpDHrs
'    rsTestData("JH_FTENUM") = tmpFTETotHrs
'    rsTestData("ED_EMP") = txtCode(3).Text
'    rsTestData("ED_PT") = txtPT.Text
'    rsTestData("Max_Entit") = medMax(0).Text
'    rsTestData("Max_Entit_Calc") = tmpNewMax
'    rsTestData("New_Entitlement") = tmpNewEntit
'    rsTestData("Old_Entitlement") = tmpOldEntit
'    rsTestData("Entit_Update") = tmpEntitUpd
'    rsTestData.Update
'    rsTestData.Close
'Hemu

    rsENT.Close
    CalcASLRepaid = xSickEnt
End Function

Private Function modUpdateSelectionWHSCC()
Dim empNo As Long
Dim dblEntitle#, dblPrevEntitle#, dblTKEEntitle#, strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#, dblFTEHoursTot#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%
Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
Dim if_Entitle As Boolean, if_Vacation As Boolean
Dim ifAnnual As Boolean, dblNewEntAnn#, VacpcNAnn, ifUnionDate As Boolean, ifFirstDate As Boolean, xAsOf 'Frank for WHSCC
Dim dblServiceYearsYTD, if_NON As Boolean
Dim NoUptSickList As String
Dim xComments
Dim rsJOB As New ADODB.Recordset

' Entitlements are always valued in HOURS - if you enter days then it
'   works out how many hours (based on average Hrswrked/day found in salary master record)

On Error GoTo modUpdateSelectionWHSCC_Err

modUpdateSelectionWHSCC = False

If Len(dlpAsOf.Text) = 0 Then
    MsgBox "Effective Date is required field"
    dlpAsOf.SetFocus
    Exit Function
End If

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

'Ticket# 3856
'If the employee's Employment Status is one of those on the list,
'do not update the employee's sick entitlement for that month. Linda Rowland
NoUptSickList = ",BD,CAS,CLIN,CONT,EIS,LTD,MAT,PAR,STUD,"

For x% = 0 To 24
    If Not IsNumeric(medLTServ(x%)) Then Exit For ' medLTServ(X%) = 0
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

'Hemu
'If Not glbWHSCC Then
'Hemu
    gdbAdoIhr001.BeginTrans
'End If

While Not snapEntitle.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / lngRecs&))
    MDIMain.panHelp(0).FloodPercent = pct%
    if_Entitle = False
    if_Vacation = False

    empNo& = snapEntitle("ED_EMPNBR")

    If Not IsNull(snapEntitle("ED_EMP")) Then
        If InStr(1, NoUptSickList, "," & Trim(snapEntitle("ED_EMP")) & ",") > 0 Then
            GoTo lblNextRec
        End If
    End If

    
    If IsNull(snapEntitle("ED_SICK")) Then
        dblEntitle# = 0
    Else
        dblEntitle# = snapEntitle("ED_SICK")
    End If
    

    If IsNull(snapEntitle("ED_PSICK")) Then
        dblPrevEntitle# = 0
    Else
        dblPrevEntitle# = snapEntitle("ED_PSICK")
    End If
    
    If IsNull(snapEntitle("ED_SICKT")) Then
        dblTKEEntitle# = 0
    Else
        dblTKEEntitle# = snapEntitle("ED_SICKT")
    End If
    
    spt = snapEntitle("ED_PT")
    strDivision$ = IIf(IsNull(snapEntitle("ED_DIV")), "", snapEntitle("ED_DIV"))
    
    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec

    varStartDate = snapEntitle(fglbWDate$)
    
    'Ticket #22434
    If rsJOB.State <> 0 Then rsJOB.Close
    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR") & " AND JH_CURRENT <> 0 AND JH_POSITION_CONTROL = 'YES'", gdbAdoIhr001, adOpenForwardOnly
    dblDHours# = 0
    dblFTEHours# = 0
    If Not rsJOB.EOF Then
        If IsNumeric(rsJOB("JH_DHRS")) Then dblDHours# = rsJOB("JH_DHRS")
        If IsNumeric(rsJOB("JH_FTENUM")) Then dblFTEHours# = rsJOB("JH_FTENUM")
    End If
    rsJOB.Close
    '-----------------------------------------------------------------------------------------------------
    'If Not IsNumeric(snapEntitle("JH_DHRS")) Then
    '    dblDHours# = 0
    'Else
    '    dblDHours# = snapEntitle("JH_DHRS")
    'End If
    
    'If Not IsNumeric(snapEntitle("JH_FTENUM")) Then
    '    dblFTEHours# = 0
    'Else
    '    dblFTEHours# = snapEntitle("JH_FTENUM")
    'End If
    dblFTEHoursTot# = GetFTEtot(empNo&, dblFTEHours#) 'For Multi Position, get the Total of FTE for one employee
    '------------------------------------------------------------------------------------------------------

    'Franks Jul 31, 02 for WHSCC
    ifAnnual = False
    ifUnionDate = False
    ifFirstDate = False

    
    ' dkostka - 08/13/2001 - Changed formula from using number of days / 365 * 12 to using DateDiff
    '   directly to get number of months.  We don't get decimals here but the value is always correct.
    '   Using the old formula would cause problems sometimes because it assumes all months have an
    '   equal number of days, and all years are 365 days.
    'dblServiceYears# = (DateDiff("d", varStartDate, CVDate(dlpAsOf)) / 365) * 12
    If Not ifAnnual Then
        'dblServiceYears# = DateDiff("m", varStartDate, CVDate(dlpAsOf))
        If Not if_NON Then
            'dblServiceYears# = DateDiff("m", varStartDate, CVDate(dlpAsOf))
            dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(dlpAsOf))
        Else
            dblServiceYears# = dblServiceYearsYTD
        End If
        intWhereFit& = -1   ' first record can be just less than
    
        For x% = 0 To 24
            If medGTServ(x%) > 0 Then
                If dblServiceYears# >= CDbl(medLTServ(x%)) And dblServiceYears# <= CDbl(medGTServ(x%)) Then
                    intWhereFit& = x%
                    If Len(medEntitle(x%)) > 0 Then if_Entitle = True
                    Exit For
                End If
            End If
        Next x%
        
        If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
    Else 'Franks Jul 31, 02 for WHSCC
        xAsOf = CVDate(GetMonth("Jan") & " 1," & Year(dlpAsOf))
        dblNewEntAnn# = 0
        VacpcNAnn = 0
        intWhereFit& = 0
        For z% = 1 To 12
            'dblServiceYears# = (DateDiff("d", varStartDate, CVDate(xAsOf)) / 365) * 12
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
                            Exit For
                        End If
                    End If
                Next x%
            Else
                If ifUnionDate Then
                    If dblServiceYears# >= 0 And dblServiceYears# < 48.99 Then
                            if_Entitle = True
                            dblNewEntAnn# = dblNewEntAnn# + 1.25
                    End If
                    If dblServiceYears# >= 49 And dblServiceYears# < 239.99 Then
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
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHoursTot# * dblDHours#
                dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
                dblEntitleUpd = dblNewEntitle
            End If
            If fglbCompMonthly% Then
                dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
            Else
                'dblEntitleUpd = dblNewEntitle ' rollover is in other utility (to accumulate)
                dblEntitleUpd# = dblNewEntitle + AccuValForMulti(empNo&, dblEntitle#) 'MultiPos Update
            End If
            If dblNewMax <> 0 Then          'only do if not zero
                    If (dblPrevEntitle# + dblEntitle# - dblTKEEntitle# + dblNewEntitle) > dblNewMax Then
                        dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
                    End If
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
                'If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHours# * dblDHours#
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHoursTot# * dblDHours#
                dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
                dblEntitleUpd = dblNewEntitle
            End If
            If optH(intWhereFit&) = True Then
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&)
            End If
            If fglbCompMonthly% Then
                dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
            Else
                'dblEntitleUpd = dblNewEntitle ' rollover is in other utility (to accumulate)
                dblEntitleUpd# = dblNewEntitle + AccuValForMulti(empNo&, dblEntitle#) 'MultiPos Update
            End If
            
            If dblNewMax <> 0 Then          'only do if not zero
                If (dblPrevEntitle# + dblEntitle# - dblTKEEntitle# + dblNewEntitle) > dblNewMax Then
                    dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
                End If
            End If
        End If
        DtTm = Now
    End If


    If if_Entitle Then
        
        'If optSickE.Value Then
            'For Sick Entitlement update, check the ASL Bank first.
            'If ASL Bank is greater than 0, take Repaid ASL from it
            'Otherwise, assign the amount to the Sick Entitlement(ED_SICK)
        dblEntitleUpd = CalcASLRepaid(empNo, CVDate(dlpAsOf), dblEntitleUpd, dblNewEntitle, dblEntitle#) 'dblEntitleUpd)
        
        'Ticket #22730
        'xComments = "Current Sick. Ent. Chg from " & snapEntitle("ED_SICK") & " to " & dblEntitleUpd
        xComments = "Current Sick. Ent. Chg from " & snapEntitle("ED_SICK") & " to " & dblEntitleUpd & ". OS: " & (IIf(IsNull(snapEntitle("ED_PSICK")), 0, snapEntitle("ED_PSICK")) + IIf(IsNull(snapEntitle("ED_SICK")), 0, snapEntitle("ED_SICK"))) - IIf(IsNull(snapEntitle("ED_SICKT")), 0, snapEntitle("ED_SICKT"))

        'Hemu - Ticket #11925 - Changed the Accrual Date from Effective Date to Entitlement Start Date
        'because otherwise it will not update Vadim until the date arrives in case it's not same as the
        'Entitlement Start Date.
        'Call Append_Accrual(EmpNo&, "SICK", dlpAsOf, dblEntitleUpd - Val(snapEntitle("ED_SICK") & ""), "U", xComments)
        If fglbCompMonthly% Then    'Ticket #22730 - Update with Effective Date if Monthly Update
            Call Append_Accrual(empNo&, "SICK", dlpAsOf, dblEntitleUpd - Val(snapEntitle("ED_SICK") & ""), "U", xComments)
        Else
            Call Append_Accrual(empNo&, "SICK", dlpDateRange(0), dblEntitleUpd - Val(snapEntitle("ED_SICK") & ""), "U", xComments)
        End If
        
        snapEntitle("ED_SICK") = dblEntitleUpd
    
    End If
    snapEntitle("ED_ANNSICK") = snapEntitle("ED_SICK")
    snapEntitle.Update
    


lblNextRec:
    snapEntitle.MoveNext

Wend
modUpdateSelectionWHSCC = True
MDIMain.panHelp(0).FloodType = 0

'Hemu
'If Not glbWHSCC Then
'Hemu
gdbAdoIhr001.CommitTrans
'End If

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
If glbWHSCC Then
    'Ticket #22434 - removing ref. to qry_JobCurrent
    'SQLQ = "SELECT HREMP.ED_EMPNBR, qry_JobCurrent.JB_GRPCD, HREMP.ED_VACPC, HREMP.ED_PVAC, HREMP.ED_VAC, HREMP.ED_VACT, HREMP.ED_ANNVAC, HREMP.ED_ANNSICK, "
    SQLQ = "SELECT HREMP.ED_EMPNBR, HREMP.ED_VACPC, HREMP.ED_PVAC, HREMP.ED_VAC, HREMP.ED_VACT, HREMP.ED_ANNVAC, HREMP.ED_ANNSICK, "
    'Ticket #22434 - removing ref. to qry_JobCurrent
    'SQLQ = SQLQ & " HREMP.ED_PSICK, HREMP.ED_SICK, HREMP.ED_SICKT,qry_JobCurrent.JH_DHRS, HREMP.ED_DIV, HREMP.ED_EMP, "
    SQLQ = SQLQ & " HREMP.ED_PSICK, HREMP.ED_SICK, HREMP.ED_SICKT, HREMP.ED_DIV, HREMP.ED_EMP, "
    'Ticket #22434 - removing ref. to qry_JobCurrent
    'SQLQ = SQLQ & " HREMP.ED_DEPTNO, HREMP.ED_PT, HREMP.ED_DOH, HREMP.ED_SENDTE, HREMP.ED_UNION, HREMP.ED_LTHIRE, HREMP.ED_USRDAT1, HREMP.ED_ORG, HREMP.ED_FDAY, qry_JobCurrent.JH_FTENUM, qry_JobCurrent.JH_DHRS, HREMP.ED_SECTION "
    SQLQ = SQLQ & " HREMP.ED_DEPTNO, HREMP.ED_PT, HREMP.ED_DOH, HREMP.ED_SENDTE, HREMP.ED_UNION, HREMP.ED_LTHIRE, HREMP.ED_USRDAT1, HREMP.ED_ORG, HREMP.ED_FDAY, HREMP.ED_SECTION "
    SQLQ = SQLQ & " FROM HREMP" 'Ticket #22434  LEFT JOIN qry_JobCurrent ON HREMP.ED_EMPNBR = qry_JobCurrent.JH_EMPNBR "
    SQLQ = SQLQ & " WHERE " & fglbESQLQ
Else
    SQLQ = "SELECT ED_EMPNBR,ED_VACPC,ED_PVAC,ED_VAC,ED_PSICK,ED_SICK,ED_SICKT,ED_EFDATES,ED_ETDATES, HREMP.ED_ANNVAC, HREMP.ED_ANNSICK, "
    SQLQ = SQLQ & " ED_DIV,ED_PT, ED_SECTION, ED_LOC, ED_EMP,"
    SQLQ = SQLQ & " ED_HIRECODE," 'County of Brant Ticket #12525
    SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME,ED_VADIM2 "
    SQLQ = SQLQ & " FROM HREMP WHERE " & fglbESQLQ
End If
If Len(clpCode(2).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_EMPNBR IN "
    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
    SQLQ = SQLQ & " WHERE JB_GRPCD = '" & clpCode(2).Text & "') "
End If

'Multi Positions Update #3304
If glbMulti Then
    SQLQ2 = "SELECT HREMP.ED_EMPNBR, COUNT(ED_EMPNBR) AS SUMEMP "
    SQLQ2 = SQLQ2 & " FROM HREMP LEFT JOIN qry_JobCurrent ON HREMP.ED_EMPNBR = qry_JobCurrent.JH_EMPNBR "
    SQLQ2 = SQLQ2 & " WHERE " & fglbESQLQ
    If Len(clpCode(2).Text) > 0 Then
        SQLQ2 = SQLQ2 & " AND ED_EMPNBR IN "
        SQLQ2 = SQLQ2 & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
        SQLQ2 = SQLQ2 & " WHERE JB_GRPCD = '" & clpCode(2).Text & "') "
    End If

    Memplist1 = "": Memplist2 = ""
    If UCase(glbCompEntVac$) = "A" Or UCase(glbCompEntSick$) = "A" Then
        'SQLQ2 = SQLQ2 & SQLQ
        If snapMultiEmp.State <> 0 Then snapMultiEmp.Close
        SQLQ2 = SQLQ2 & " GROUP BY ED_EMPNBR HAVING COUNT(ED_EMPNBR) > 1 "
        snapMultiEmp.Open SQLQ2, gdbAdoIhr001, adOpenStatic
        Do While Not snapMultiEmp.EOF
            Memplist1 = Memplist1 & "'" & snapMultiEmp("ED_EMPNBR") & "',"
            Memplist2 = Memplist2 & "'" & snapMultiEmp("ED_EMPNBR") & "',"
            snapMultiEmp.MoveNext
        Loop
        snapMultiEmp.Close
    End If
End If
'Multi Positions Update #3304

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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapEntitle", "Entitlements/EMP", "Select")

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
        
        'made the DoWork a separate sub because it's being used in two places
        If chkManual.Value = False Then
            If chkMUEntitle() Then
            
                'Ticket #19632 - This is becsuse they are using TAKEN as part of Max checking. So when the date range is
                'changed the TAKEN should be recalculated so on Update Entitle, the correct TAKEN is used in the formula.
                'During Year End, on the date range is changed, saved and Update Entitlement is clicked, the TAKEN of last
                'year is still there in ED_SICT and that was being used in the Max comparison formula. This recalculate
                'will fix the issue by recalculating the TAKEN.
                If glbCompSerial = "S/N - 2173W" Or glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2389W" Or _
                    glbCompSerial = "S/N - 2408W" Or glbCompSerial = "S/N - 2412W" Or glbCompSerial = "S/N - 2399W" Or _
                    glbCompSerial = "S/N - 2395W" Or glbCompSerial = "S/N - 2430W" Or glbCompSerial = "S/N - 2190W" Or _
                    glbCompSerial = "S/N - 2450W" Or glbCompSerial = "S/N - 2436W" Or glbCompSerial = "S/N - 2466W" Or _
                    glbCompSerial = "S/N - 2234W" Then
                    
                    Call getWSQLQ("C")
                    Call EntReCalcPeriod(fglbESQLQ, "SICK", , , dlpDateRange(0), dlpDateRange(1))
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
    MsgBox "All Rules applied", vbInformation + vbOKOnly, "Sick Entitlements"
Else
    MsgBox "The Following Rules failed:" & vbCrLf & failed, vbInformation + vbOKOnly, "Sick Entitlements"
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

glbOnTop = "FRMSICKENT"

End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim x%
Dim SQLQ

glbOnTop = "FRMSICKENT"

FlagRefresh = False

Data1.ConnectionString = glbAdoIHRDB

SQLQ = "SELECT DISTINCT VE_DIV,VE_DEPT,VE_ORG,VE_LOC,VE_SECTION,VE_EDATE,VE_EMP,VE_PT,VE_GRPCD,VE_FRDATE,VE_TODATE, VE_MANUAL FROM HRSICKENT "
If glbWFC Then 'Ticket #28553 Franks 05/03/2016
    SQLQ = SQLQ & " WHERE " & getWFCPlantSecurity("VE_SECTION")
End If
If glbDIVCount = 1 And glbLinamar Then
    SQLQ = SQLQ & " WHERE VE_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
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

If glbCBrant Then
    'County of Brant using Sick Time Entitlement Outstanding Based Upon to calculate the service months
    'Ticket #Ticket #12544
    Select Case glbEntOutStandingS$
        Case "2": fglbWDate$ = "ED_DOH"
        Case "3": fglbWDate$ = "ED_SENDTE"
        Case "4": fglbWDate$ = "ED_LTHIRE"
        Case "5": fglbWDate$ = "ED_USRDAT1"
        Case "6": fglbWDate$ = "ED_UNION"
    End Select
Else
    Select Case glbCompWDate$ ' sets field reference for basic 'which date'
        Case "O": fglbWDate$ = "ED_DOH"
        Case "S": fglbWDate$ = "ED_SENDTE"
        Case "U": fglbWDate$ = "ED_UNION"
        Case "L": fglbWDate$ = "ED_LTHIRE"
        Case "D": fglbWDate$ = "ED_USRDAT1"
    End Select
End If

If glbCompEntSick$ = "M" Or UCase(glbCompEntSick$) = "N" Then
    chkRound.Visible = True
    chkRound.Value = False
Else
    chkRound.Value = False
    chkRound.Visible = False
End If

If UCase(glbCompEntSick$) = "M" Or UCase(glbCompEntSick$) = "N" Then
    vbxTrueGrid.Columns(5).Visible = False
End If

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

'Ticket #18235 - - Location to Vadim 2 - Samuel, Son & Co., Limited
If glbCompSerial = "S/N - 2382W" Then
    lblLocation.Caption = lStr("Vadim Field 2")
    vbxTrueGrid.Columns(9).Caption = lStr("Vadim Field 2")
    clpCode(4).TablName = "EDV2"
    clpCode(4).Tag = "00-Enter Vadim 2 Code"
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
If Me.Height >= 3750 + VacFram.Height + panControls.Height + 230 Then
    scrControl.Value = 0
    VacFram.Top = 4440
    scrControl.Visible = False
    Exit Sub
End If
scrControl.Visible = True
scrControl.Max = VacFram.Height + panControls.Height + 3750 + 550 - Me.Height '250 - Me.Height
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

Private Sub modMaximums(TF%)
Dim x%

For x% = 0 To 24
    If Not TF Then
        If IsNumeric(medMax(x%)) Then medMax(x%) = 0
    End If
    medMax(x%).Enabled = TF And medMax(x%).Enabled
Next x%

End Sub

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
'Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
'Dim if_Entitle As Boolean, if_Vacation As Boolean
Dim xComments
Dim dblEntitleDays
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

If snapEntitle.BOF And snapEntitle.EOF Then
    'If fglbRunTimes = 1 Then
        MsgBox "Employees for this selection do not exist!"
        Exit Function
    'End If
Else
    lngRecs& = snapEntitle.RecordCount
    If fglbRunTimes = 1 Or UCase(glbCompEntSick$) <> "N" Then   'Ticket #26777 - Prompt for Annual and Monthly as well
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

'Ticket #22682 - Release 8.0: Check Accrual File to see if the update already done for Monthly Updates only. This is
'to avoid multiple updates for the same month.
'Only for Monthly updates
If glbCompEntSick$ = "M" Then
    Do While Not snapEntitle.EOF
        'Ticket #28024 - To fix the error caused by calling this function without '' apostrophes
        'If Accrual_Rec_Exists(snapEntitle("ED_EMPNBR"), "SICK", dlpAsOf.Text, "U") Then
        If Accrual_Rec_Exists(snapEntitle("ED_EMPNBR"), "SICK", dlpAsOf.Text, "'U'") Then
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

While Not snapEntitle.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / lngRecs&))
    MDIMain.panHelp(0).FloodPercent = pct%

    'If snapEntitle("ED_EMPNBR") = 3190 Then
    '    EmpNo& = snapEntitle("ED_EMPNBR")
    'End If
    
    empNo& = snapEntitle("ED_EMPNBR")
    
    If IsNull(snapEntitle("ED_SICK")) Then
        dblEntitle# = 0
    Else
        dblEntitle# = snapEntitle("ED_SICK")
    End If
    
    If IsNull(snapEntitle("ED_PSICK")) Then
        dblPrevEntitle# = 0
    Else
        dblPrevEntitle# = snapEntitle("ED_PSICK")
    End If
    
    If IsNull(snapEntitle("ED_SICKT")) Then
        dblTKEEntitle# = 0
    Else
        dblTKEEntitle# = snapEntitle("ED_SICKT")
    End If
    
    spt = snapEntitle("ED_PT")
    
    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec

    'Ticket #14260 DNSSAB
    'Check last month attendance records, if there is any record with Incentive checked,
    'and then skip this employee, also update the Accrual table
    If glbCompSerial = "S/N - 2388W" Then
        If IncentiveChecked(empNo&, dlpAsOf.Text) Then
            Call Append_Accrual(empNo&, "SICK", dlpAsOf.Text, 0, "N", "No Sick Ent Attendance Found.")
            GoTo lblNextRec
        End If
    End If
    
    varStartDate = snapEntitle(fglbWDate$)
    
    'If snapEntitle("ED_EMPNBR") = 1627 Then '41332
    '    DoEvents
    'End If
    
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
    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    intWhereFit& = -1

    For x% = 0 To 24
        If medGTServ(x%) > 0 Then
            If dblServiceYears# >= CDbl(medLTServ(x%)) And dblServiceYears# <= CDbl(medGTServ(x%)) Then
                intWhereFit& = x%
                Exit For
            End If
        End If
    Next x%
    
    'Hemu - Added dblServiceYears# < 0 because it gives out entitlement way high which is wrong
    If intWhereFit& = -1 Or dblServiceYears# < 0 Then GoTo lblNextRec ' skip record if not in any of the ranges
    
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
    
    
    'Ticket #19963 - S.U.C.C.E.S.S
    If glbCompSerial = "S/N - 2422W" Then
        'If employee falls in Second Service Range then get the last year same effective month Sick Time taken
        'This will be the earning for this month, this year not exceeding Maximum.
        If intWhereFit& = 1 Then
            'New entitlement
            dblNewEntitle# = Get_LastYearSickTaken_ThisMonth(empNo&)
            
            'Get Maximum for the range
            dblNewMax# = 0
            If optD(intWhereFit&) = True Then           ' Entitlements entered in days
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblDHours#
            End If
            If optF(intWhereFit&) = True Then
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblFTEHours# * dblDHours#
            End If
            If optH(intWhereFit&) = True Then
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&)
            End If
            GoTo nextStep
        End If
    End If
    
    dblNewEntitle# = medEntitle(intWhereFit&)
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
nextStep:
    If fglbCompMonthly Then
        dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
    Else
        dblEntitleUpd = dblNewEntitle ' rollover is in other utility (to accumulate)
    End If
    
    If dblNewMax <> 0 Then          'only do if not zero
        If glbCompSerial = "S/N - 2173W" Or glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2389W" Or _
            glbCompSerial = "S/N - 2408W" Or glbCompSerial = "S/N - 2412W" Or glbCompSerial = "S/N - 2399W" Or _
            glbCompSerial = "S/N - 2395W" Or glbCompSerial = "S/N - 2430W" Or glbCompSerial = "S/N - 2190W" Or _
            glbCompSerial = "S/N - 2450W" Or glbCompSerial = "S/N - 2436W" Or glbCompSerial = "S/N - 2466W" Or _
            glbCompSerial = "S/N - 2234W" Then
            
            'for town of Ajax or City of Timmins or St. Leonard's Community Services(Ticket #15071)
            'Ticket #17090 - Township of Wilmot
            'Ticket #17160 - NorWest Community Health Centres
            'Ticket #17111 - West Elgin Community Health Centre
            'Ticket #17315 - The Youth Centre
            'Ticket #20653 - kidsLINK
            'Ticket #24051 - Township of Severn - 2450W - Franks 07/12/2013
            'Ticket #24050 - Family Day Care Services 2436W - Franks 07/12/2013
            'Ticket #26573 - Chiefs of Ontario
            'Ticket #26608 - Girl Guides of Canada
            If (dblEntitle# + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
                'OS exceeding the Maximum, do not change the Current - do not give extra entitlement of the month.
                'Just set Current back to same value as it was before
                dblEntitleUpd = dblEntitle#
                
                'Ticket #24590 - NorWest CHC: Reduce the OS to maximum by adding an Attendance record of SICK
                If glbCompSerial = "S/N - 2412W" Then
                    xComments = "System generated Attendance record due to the employee exceeding the Sick Entitlement for the Month as of '" & dlpAsOf & "'."
                    Call Add_Adjustment_Attendance(empNo&, dlpAsOf, "SICK", (dblEntitle# + dblPrevEntitle# - dblTKEEntitle#) - dblNewMax, xComments)
                End If
                
            ElseIf (dblEntitleUpd + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
                dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
            End If
        Else
            If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
                dblEntitleUpd = dblNewMax - dblPrevEntitle#
                
                'Ticket #13359 - Simcoe Muskoka District Health Unit
                'Ticket #17090 - Township of Wilmot
                If glbCompSerial = "S/N - 2228W" Or glbCompSerial = "S/N - 2408W" Then
                    If dblEntitleUpd < 0 Then
                        dblEntitleUpd = 0
                    End If
                End If
            End If
        End If
    End If
    
    If glbCBrant Then
        If snapEntitle("ED_HIRECODE") = "Y" And dblTKEEntitle# > 0 Then
            dblEntitleUpd = dblEntitleUpd - dblTKEEntitle#
        End If
    End If
    DtTm = Now
    
    If isLast And chkRound.Visible = True And chkRound Then
            'Round the final entitlement
            If dblDHours# <> 0 And optH(intWhereFit&) = False Then
                dblEntitleDays = dblEntitleUpd / dblDHours#
                dblEntitleDays = Round(dblEntitleDays, 0)
                dblEntitleUpd = dblEntitleDays * dblDHours#
            Else
                dblEntitleUpd = Round(dblEntitleUpd, 0)
            End If
    ElseIf glbCompEntSick$ = "M" And chkRound.Visible = True And chkRound Then
        'If month(dlpDateRange(1).Text) = month(dlpAsOf.Text) And Year(dlpDateRange(1).Text) = Year(dlpAsOf.Text) Then
            'Round the final entitlement
            If dblDHours# <> 0 And optH(intWhereFit&) = False Then
                dblEntitleDays = dblEntitleUpd / dblDHours#
                dblEntitleDays = Round(dblEntitleDays, 0)
                dblEntitleUpd = dblEntitleDays * dblDHours#
            Else
                dblEntitleUpd = Round(dblEntitleUpd, 0)
            End If
        'Else
        '    dblEntitleUpd = dblEntitleUpd       ' base entitlements sic/vacation
        'End If
    End If
    
    
    'Ticket #22730
    'xComments = "Current Sick. Ent. Chg from " & snapEntitle("ED_SICK") & " to " & dblEntitleUpd
    xComments = "Current Sick. Ent. Chg from " & snapEntitle("ED_SICK") & " to " & dblEntitleUpd & ". OS: " & (IIf(IsNull(snapEntitle("ED_PSICK")), 0, snapEntitle("ED_PSICK")) + IIf(IsNull(snapEntitle("ED_SICK")), 0, snapEntitle("ED_SICK"))) - IIf(IsNull(snapEntitle("ED_SICKT")), 0, snapEntitle("ED_SICKT"))

    'Hemu - Ticket #11925 - Changed the Accrual Date from Effective Date to Entitlement Start Date
    'because otherwise it will not update Vadim until the date arrives in case it's not same as the
    'Entitlement Start Date.
    'Call Append_Accrual(EmpNo&, "SICK", dlpAsOf, dblEntitleUpd - Val(snapEntitle("ED_SICK") & ""), "U", xComments)
    If fglbCompMonthly Then
        Call Append_Accrual(empNo&, "SICK", dlpAsOf, dblEntitleUpd - Val(snapEntitle("ED_SICK") & ""), "U", xComments)
    Else
        'Annual
        'Ticket #23141
        If glbVadim Then
            'For Vadim user's we need to send the full value that the employee Annual Accrued, since we are
            'not doing zero out for Current in the Year End. This is revised steps for Vadim users only for
            'the Year End.
            Call Append_Accrual(empNo&, "SICK", dlpDateRange(0), dblEntitleUpd, "U", xComments)
        Else
            Call Append_Accrual(empNo&, "SICK", dlpDateRange(0), dblEntitleUpd - Val(snapEntitle("ED_SICK") & ""), "U", xComments)
        End If
    End If

    snapEntitle("ED_SICK") = dblEntitleUpd
    snapEntitle("ED_ANNSICK") = dblEntitleUpd
    snapEntitle.Update
    
lblNextRec:
    DoEvents
    Dim xKey
    xKey = snapEntitle("ED_EMPNBR")
    'xKey = xKey & "|" & Format(snapEntitle("ED_EFDATES"), "dd-mmm-yyyy")
    'xKey = xKey & "|" & Format(snapEntitle("ED_ETDATES"), "dd-mmm-yyyy")
    xKey = xKey & "|" & Format(dlpDateRange(0), "dd-mmm-yyyy")
    xKey = xKey & "|" & Format(dlpDateRange(1), "dd-mmm-yyyy")
    xKey = xKey & "|SICK"
    If dblServiceYears# < 0 Then
        dblEntitleUpd = 0
    End If
    xKey = xKey & "|" & dblEntitleUpd
    xKey = xKey & "|" & Format(dlpAsOf.Text, "dd-mmm-yyyy") 'Transaction Date
    Call Entitlements_Master_Integration(xKey, empNo&) 'George added for Advance Tracker
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
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateEntitle", "HR_EMP", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If
End Function

Private Function modAnnSelection(isLast As Boolean)
Dim empNo As Long
Dim dblEntitle#, dblPrevEntitle#, dblTKEEntitle#, strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%, xAsOf
Dim dblEntitleDays
'Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
'Dim if_Entitle As Boolean, if_Vacation As Boolean
Dim xComments
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
                If medGTServ(x%) = Int(medGTServ(x%)) And medGTServ(x%) > 0 Then medGTServ(x%) = medGTServ(x%) + 0.99
            Else
                If Val(medGTServ(x%)) = Int(medGTServ(x%)) And Val(medGTServ(x%)) > 0 Then medGTServ(x%) = medGTServ(x%) + 0.99
            End If
        End If
        If medLTServ(x%) > 0 And medGTServ(x%) = 0 Then medGTServ(x%) = 9999999
    Next


    empNo& = snapEntitle("ED_EMPNBR")
    
    If IsNull(snapEntitle("ED_ANNSICK")) Then
        dblEntitle# = 0
    Else
        dblEntitle# = snapEntitle("ED_ANNSICK")
    End If
    
  
    If IsNull(snapEntitle("ED_PSICK")) Then
        dblPrevEntitle# = 0
    Else
        dblPrevEntitle# = snapEntitle("ED_PSICK")
    End If
    
    If IsNull(snapEntitle("ED_SICKT")) Then
        dblTKEEntitle# = 0
    Else
        dblTKEEntitle# = snapEntitle("ED_SICKT")
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
        If glbCompSerial = "S/N - 2173W" Or glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2389W" Or _
            glbCompSerial = "S/N - 2408W" Or glbCompSerial = "S/N - 2412W" Or glbCompSerial = "S/N - 2399W" Then
            
            'for town of Ajax or City of Timmins or St. Leonard's Community Services
            'Ticket #17090 - Township of Wilmot
            'Ticket #17160 - NorWest Community Health Centres
            'Ticket #17111 - West Elgin Community Health Centre
            If (dblEntitle# + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
                dblEntitleUpd = dblEntitle#
            ElseIf (dblEntitleUpd + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
                dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
            End If
        Else
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
    End If
    
    If glbCBrant Then
        If snapEntitle("ED_HIRECODE") = "Y" And dblTKEEntitle# > 0 Then
            dblEntitleUpd = dblEntitleUpd - dblTKEEntitle#
        End If
    End If
    DtTm = Now
    
    If isLast And chkRound.Visible = True And chkRound Then
        'Round the final entitlement
        If dblDHours# <> 0 And optH(intWhereFit&) = False Then
            dblEntitleDays = dblEntitleUpd / dblDHours#
            dblEntitleDays = Round(dblEntitleDays, 0)
            dblEntitleUpd = dblEntitleDays * dblDHours#
        Else
            dblEntitleUpd = Round(dblEntitleUpd, 0)
        End If
    End If
    
    xComments = "Current Sick. Ent. Chg from " & snapEntitle("ED_SICK") & " to " & dblEntitleUpd

   snapEntitle("ED_ANNSICK") = dblEntitleUpd
   snapEntitle.Update
lblNextRec:
    DoEvents



modAnnSelection = True
Screen.MousePointer = DEFAULT

Exit Function

modUpdateSelection_Err:
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
VacFram.Top = 4380 - scrControl.Value
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
Next

clpDiv.Enabled = TF
clpDept.Enabled = TF
clpCode(0).Enabled = TF

If Not TF Or glbLinamar Then
    lblAsOf.FontBold = True
Else
    lblAsOf.FontBold = False
End If

If glbCompEntSick$ = "M" Or glbCompEntSick$ = "N" Or glbCompEntSick$ = "A" Then
    dlpAsOf.Enabled = True 'FT
Else
    dlpAsOf.Enabled = True 'Ticket #3419
End If

'If sick Entitlement Outstanding based on "1" then ok, otherwise disenable
If glbEntOutStandingS$ = "1" Then
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
''cmdPrintAll.Enabled = FT
'cmdUpdate.Enabled = FT
'vbxTrueGrid.Enabled = FT
Call modSetFGlobals("SICK")

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
Next
clpDiv.Text = ""
clpDept.Text = ""
clpCode(0).Text = ""
If Not (glbCompEntSick$ = "M" Or glbCompEntSick$ = "N") Then
    dlpAsOf.Text = ""
End If
clpCode(1).Text = ""
clpCode(2).Text = ""
clpCode(3).Text = ""
clpCode(4).Text = ""
clpPT.Text = ""
dlpDateRange(0).Text = ""
dlpDateRange(1).Text = ""

If Not Data1.Recordset.EOF Then
    SQLQ = "SELECT * FROM HRSICKENT "
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

    'Sam 02/03/2006
    If Not IsNull(Data1.Recordset("VE_FRDATE")) Then
        SQLQ = SQLQ & " AND VE_FRDATE = " & Date_SQL(Data1.Recordset("VE_FRDATE"))
    End If
    If Not IsNull(Data1.Recordset("VE_TODATE")) Then
        SQLQ = SQLQ & " AND VE_TODATE = " & Date_SQL(Data1.Recordset("VE_TODATE"))
    End If
    'Sam 02/03/2006
    
    SQLQ = SQLQ & " Order By VE_DIV,VE_DEPT,VE_ORG,VE_EDATE,VE_EMP,VE_PT,VE_LOC,VE_SECTION,VE_ORDER "
    rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    If Not IsNull(Data1.Recordset("VE_DIV")) Then clpDiv.Text = Data1.Recordset("VE_DIV")
    If Not IsNull(Data1.Recordset("VE_DEPT")) Then clpDept.Text = Data1.Recordset("VE_DEPT")
    If Not IsNull(Data1.Recordset("VE_ORG")) Then clpCode(0).Text = Data1.Recordset("VE_ORG")
    If Not IsNull(Data1.Recordset("VE_EDATE")) Then dlpAsOf.Text = Data1.Recordset("VE_EDATE")
    If Not IsNull(Data1.Recordset("VE_EMP")) Then clpCode(1).Text = Data1.Recordset("VE_EMP")
    If Not IsNull(Data1.Recordset("VE_PT")) Then clpPT.Text = Data1.Recordset("VE_PT")
    If Not IsNull(Data1.Recordset("VE_GRPCD")) Then clpCode(2).Text = Data1.Recordset("VE_GRPCD")
    If Not IsNull(Data1.Recordset("VE_LOC")) Then clpCode(4).Text = Data1.Recordset("VE_LOC")
    If Not IsNull(Data1.Recordset("VE_SECTION")) Then clpCode(3).Text = Data1.Recordset("VE_SECTION")
    'Sam 02/03/2006
    If Not IsNull(Data1.Recordset("VE_FRDATE")) Then dlpDateRange(0).Text = Data1.Recordset("VE_FRDATE")
    If Not IsNull(Data1.Recordset("VE_TODATE")) Then dlpDateRange(1).Text = Data1.Recordset("VE_TODATE")
    'Sam 02/03/2006
    If Not IsNull(Data1.Recordset("VE_MANUAL")) Then chkManual.Value = Data1.Recordset("VE_MANUAL")
    
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
    
    SQLQ = "SELECT DISTINCT VE_DIV,VE_DEPT,VE_ORG,VE_LOC,VE_SECTION,VE_EDATE,VE_EMP,VE_PT,VE_GRPCD,VE_FRDATE,VE_TODATE, VE_MANUAL FROM HRSICKENT "
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
If glbCompEntSick$ = "M" Or UCase(glbCompEntSick$) = "N" Then
    fglbCompMonthly% = True
    Call modMaximums(True)
Else
    fglbCompMonthly% = False
    Call modMaximums(False)
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
If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & clpCode(3).Text & "' "
If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #18235
    If Len(clpCode(4).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_VADIM2 = '" & clpCode(4).Text & "' "
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
    xGRPCE = clpCode(2).Text
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
If UCase(glbCompEntSick$) = "A" Then
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

'Sam 02/03/2006
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
'Sam 02/03/2006
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
    cmdUpdateAll.Enabled = False
ElseIf Me.Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
    cmdPrintAll.Enabled = True
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
    'Added by Bryan 25/Oct/05 Ticket#9560
    Dim lastday
    Dim flglastdate As Boolean
    Dim lngRecs As Long, pct As Long, prec As Long
    Dim blIsLast As Boolean
    
    Screen.MousePointer = DEFAULT
    DoWork = False
    
    'Annualized Monthly
    If UCase(glbCompEntSick$) = "N" Then
        For fglbRunTimes = 1 To 12
            blIsLast = False
            If fglbRunTimes = 12 Then blIsLast = True
        
            If Not modUpdateSelection(blIsLast) Then Exit Function
            dlpAsOf = DateAdd("m", 1, CVDate(dlpAsOf.Text))
            
            DoEvents
            
            If glbCompSerial = "S/N - 2386W" Then 'The Walter Fedy Partnership Ticket #14255
                Call Pause(3)
                MsgBox "Click OK Button to next month: " & dlpAsOf
            End If
        Next
        dlpAsOf = DateAdd("m", -12, CVDate(dlpAsOf.Text))
    
    'Monthly or Annual
    Else
        If Not glbWHSCC Then
        
            If Not modUpdateSelection() Then Exit Function
            
            If fglbCompMonthly Then
            
                'Ticket #30154 - This is done here so that the Entitlement period is updated incase it has not done yet. The Annual Vacation computation uses the From Date
                'which was only getting updated at the end of this function. So by calling this function it computes the date range at this level.
                Call getWSQLQ("")
                Call EntReCalcPeriod(fglbESQLQ, "SICK")
                
                gdbAdoIhr001.Execute "UPDATE HREMP SET ED_ANNSICK=0 WHERE " & fglbESQLQ
            
                'Ticket #24590 - NorWest CHC: Reduce the OS to maximum by adding an Attendance record of SICK
                'Monthly update already added the Attendance where required, so just do a recalculate to compute
                'Taken so the Annual Entitlement is computed correctly.
                If glbCompSerial = "S/N - 2412W" Then
                    Call getWSQLQ("")
                    If Len(clpCode(2).Text) > 0 Then
                        fglbESQLQ = fglbESQLQ & " AND ED_EMPNBR IN "
                        fglbESQLQ = fglbESQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
                        fglbESQLQ = fglbESQLQ & " WHERE JB_GRPCD = '" & clpCode(2).Text & "') "
                    End If
                    Call EntReCalc(fglbESQLQ, Empty, "TAKEN ONLY")
                End If
            
                If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
                
                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
                    While Not snapEntitle.EOF
                        lngRecs = snapEntitle.RecordCount
                        prec = prec + 1
                        pct = Int(100 * (prec / lngRecs))
                        If pct > 100 Then pct = 100
                        MDIMain.panHelp(0).FloodPercent = pct
                        
                        Dim doDate As Date
                        doDate = dlpAsOf
                        
                        If Not IsNull(snapEntitle("ED_EFDATES")) Then 'Ticket #12923
                            'fglbAsOf = snapEntitle("ED_EFDATES")
                            fglbAsOf = IsValidDate(Format(month(snapEntitle("ED_EFDATES")) & "/" & Day(dlpAsOf) & "/" & Year(snapEntitle("ED_EFDATES")), "mm/dd/yyyy"), Day(dlpAsOf), month(snapEntitle("ED_EFDATES")), Year(snapEntitle("ED_EFDATES")))
                            'fglbAsOf = CVDate(month(snapEntitle("ED_EFDATES")) & "/" & Day(dlpAsOf) & "/" & Year(snapEntitle("ED_EFDATES")))
                            For fglbRunTimes = 1 To 12
                                blIsLast = False
                                If fglbRunTimes = 12 Then blIsLast = True
                                
                                If Not modAnnSelection(blIsLast) Then Exit Function
                                
                                fglbAsOf = DateAdd("m", 1, CVDate(fglbAsOf))

                                DoEvents
                            Next
                        End If
                        snapEntitle.MoveNext
                    Wend
                    MDIMain.panHelp(0).FloodType = 0
                End If
            End If
        Else
            If Not modUpdateSelectionWHSCC() Then Exit Function
        End If
    End If
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    Screen.MousePointer = HOURGLASS
    Call EntReCalc(fglbESQLQ)
    
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    DoWork = True
End Function

Private Function IncentiveChecked(xEmpNo, xEffDate)
Dim rsAttInc As New ADODB.Recordset
Dim SQLQ As String
Dim xDateFrom, xDateEnd
Dim xMonth As String
    xMonth = MonthName(month(xEffDate))
    xDateFrom = CVDate(xMonth & " 1," & Str(Year(xEffDate)))
    xDateEnd = DateAdd("D", -1, xDateFrom)
    xDateFrom = DateAdd("M", -1, xDateFrom)
    
    SQLQ = "SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND AD_DOA >= " & Date_SQL(xDateFrom) & " "
    SQLQ = SQLQ & "AND AD_DOA <= " & Date_SQL(xDateEnd) & " "
    SQLQ = SQLQ & "AND NOT (AD_INDICATOR = 0) "
    rsAttInc.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsAttInc.EOF Then
        IncentiveChecked = True
    Else
        IncentiveChecked = False
    End If
    rsAttInc.Close
End Function

Private Function Get_LastYearSickTaken_ThisMonth(xEmpnbr)
    Dim rsAttend As New ADODB.Recordset
    Dim SQLQ As String
    Dim xTaken As Double
    
    'Calculate & return the Sick Taken last year for the same Effective Month
        
    Get_LastYearSickTaken_ThisMonth = 0
    
    SQLQ = "SELECT SUM(AD_HRS) AS SICKTAKEN FROM HR_ATTENDANCE WHERE "
    SQLQ = SQLQ & " AD_EMPNBR = " & xEmpnbr
    SQLQ = SQLQ & " AND LEFT(AD_REASON,3) = 'SIC' "
    SQLQ = SQLQ & " AND MONTH(AD_DOA) =" & month(dlpAsOf)
    SQLQ = SQLQ & " AND YEAR(AD_DOA) =" & Year(dlpAsOf) - 1
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsAttend.EOF Then
        If Not IsNull(rsAttend("SICKTAKEN")) Then
            xTaken = rsAttend("SICKTAKEN")
        Else
            xTaken = 0
        End If
    Else
        xTaken = 0
    End If
    rsAttend.Close
    Set rsAttend = Nothing
    
    'Return Sick Taken for the Month last year
    Get_LastYearSickTaken_ThisMonth = xTaken
    
End Function

Private Sub Add_Adjustment_Attendance(xEmpnbr, xDate, xReason, xAdjHrs, xComments)
    Dim rsAddAttend As New ADODB.Recordset
    Dim rsTABL As New ADODB.Recordset
    Dim rsCurSal As New ADODB.Recordset
    Dim rsCurPos As New ADODB.Recordset
    Dim SQLQ As String
    Dim xPoint
    Dim xEML As Boolean
    Dim xIncentive As Boolean
    Dim xSen As Boolean
    
    'Add Sick Code if not existing in the Table Master
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_KEY = '" & xReason & "'"
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsTABL.EOF Then
        rsTABL.AddNew
        rsTABL("TB_COMPNO") = "001"
        rsTABL("TB_NAME") = "ADRE"
        rsTABL("TB_KEY") = xReason
        rsTABL("TB_DESC") = xReason
        rsTABL("TB_LDATE") = Date
        rsTABL("TB_LTIME") = Time$
        rsTABL("TB_LUSER") = glbUserID
        rsTABL.Update
    Else
        If rsTABL("TB_USR2") > 0 Then 'Points
            xPoint = rsTABL("TB_USR2")
        Else
            xPoint = Null
        End If
        If rsTABL("TB_USR3") <> 0 Then 'EML
            xEML = rsTABL("TB_USR3")
        Else
            xEML = 0
        End If
        xIncentive = rsTABL("TB_INDICATOR")
        xSen = rsTABL("TB_SEN")
    End If
    rsTABL.Close
    Set rsTABL = Nothing

    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE 1 = 2"
    rsAddAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsAddAttend.AddNew
    rsAddAttend("AD_COMPNO") = "001"
    rsAddAttend("AD_EMPNBR") = xEmpnbr
            
    'Update with Salary info.
    SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & xEmpnbr
    rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsCurSal.BOF Then
        rsAddAttend("AD_SALARY") = rsCurSal("SH_SALARY")
        rsAddAttend("AD_SALCD") = rsCurSal("SH_SALCD")
    End If
    rsCurSal.Close
    Set rsCurSal = Nothing
    
    'Update with Position info.
    SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS,JH_REPTAU,JH_PAYROLL_ID,JH_SHIFT,JH_GLNO,JH_ORG FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & xEmpnbr
    rsCurPos.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsCurPos.EOF Then
        rsAddAttend("AD_JOB") = rsCurPos("JH_JOB")
        rsAddAttend("AD_DHRS") = rsCurPos("JH_DHRS")
        rsAddAttend("AD_WHRS") = rsCurPos("JH_WHRS")
        rsAddAttend("AD_SUPER") = rsCurPos("JH_REPTAU")
        rsAddAttend("AD_PAYROLL_ID") = rsCurPos("JH_PAYROLL_ID")
        rsAddAttend("AD_SHIFT") = rsCurPos("JH_SHIFT")
        rsAddAttend("AD_GLNO") = rsCurPos("JH_GLNO")
        rsAddAttend("AD_ORG") = rsCurPos("JH_ORG")
    End If
    rsCurPos.Close
    Set rsCurPos = Nothing
    
    rsAddAttend("AD_DOA") = xDate
    rsAddAttend("AD_REASON") = xReason
    rsAddAttend("AD_HRS") = xAdjHrs
    rsAddAttend("AD_COMM") = xComments
    'rsAddAttend("AD_BANKHRS_EXP") = IIf(IsDate(xExpiryDate), xExpiryDate, Null)
    'rsAddAttend("AD_CONSUMED") = IIf(IsMissing(xConsumed), Null, xConsumed)
    
    rsAddAttend("AD_INDICATOR") = xIncentive
    rsAddAttend("AD_SEN") = xSen
    rsAddAttend("AD_EMELEA") = xEML
    rsAddAttend("AD_POINT") = xPoint
    
    rsAddAttend("AD_LUSER") = glbUserID
    rsAddAttend("AD_LDATE") = Date
    rsAddAttend("AD_LTIME") = Time$
    rsAddAttend("AD_SOURCE") = "SICADJ"
    rsAddAttend.Update
    rsAddAttend.Close
    Set rsAddAttend = Nothing
End Sub

