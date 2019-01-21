VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRGapAnalysis 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Gap Analysis Report"
   ClientHeight    =   7215
   ClientLeft      =   330
   ClientTop       =   1260
   ClientWidth     =   9960
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
   ScaleHeight     =   7215
   ScaleWidth      =   9960
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkNotMatched 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2085
      TabIndex        =   18
      Top             =   5400
      Width           =   495
   End
   Begin VB.CheckBox chkRelocate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2085
      TabIndex        =   14
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      MaxLength       =   25
      TabIndex        =   15
      Top             =   4680
      Width           =   4935
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2100
      MaxLength       =   4
      TabIndex        =   12
      Tag             =   "00-Employee Position Shift"
      Top             =   3990
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Index           =   1
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Tag             =   "Second level of grouping records"
      Top             =   6525
      Width           =   2325
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
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
      Index           =   0
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Tag             =   "First Level of grouping records"
      Top             =   6180
      Width           =   2325
   End
   Begin Threed.SSCheck chkShowDet 
      Height          =   225
      Index           =   3
      Left            =   4515
      TabIndex        =   21
      Tag             =   "Hide all detail records - show summaries only"
      Top             =   6465
      Visible         =   0   'False
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "   Show Detail"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Value           =   -1  'True
      Font3D          =   3
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1680
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDEM"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2010
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDPT"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1350
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDOR"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1020
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   690
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   360
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   1
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   8
      Left            =   1800
      TabIndex        =   10
      Tag             =   "00-Enter Administered By Code"
      Top             =   3330
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   9
      Left            =   1800
      TabIndex        =   11
      Tag             =   "00-Enter Section Code"
      Top             =   3660
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   7
      Left            =   1800
      TabIndex        =   9
      Tag             =   "00-Enter Region Code"
      Top             =   3000
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3360
      TabIndex        =   8
      Tag             =   "40-Date upto and including this date forward"
      Top             =   2655
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   7
      Tag             =   "40-Date from and including this date forward"
      Top             =   2670
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2340
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6480
      Top             =   6720
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
   End
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   16
      Tag             =   "Detailed Report"
      Top             =   5040
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  Show Last Review"
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
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "Summary Report"
      Top             =   5040
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  All reviews"
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
   Begin INFOHR_Controls.CodeLookup clpPosCode 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Tag             =   "01-Position code"
      Top             =   4320
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   6
      LookupType      =   5
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Show only Not Matched"
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
      Left            =   120
      TabIndex        =   41
      Top             =   5430
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Relocate ?"
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
      Left            =   120
      TabIndex        =   40
      Top             =   4725
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Index           =   18
      Left            =   3120
      TabIndex        =   39
      Top             =   4725
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Preferences"
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
      Left            =   120
      TabIndex        =   38
      Top             =   4365
      Width           =   1650
   End
   Begin VB.Label lblShift 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
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
      Left            =   120
      TabIndex        =   37
      Top             =   4035
      Visible         =   0   'False
      Width           =   645
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
      Left            =   120
      TabIndex        =   36
      Top             =   2055
      Width           =   630
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Left            =   120
      TabIndex        =   35
      Top             =   3705
      Width           =   540
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
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
      Left            =   120
      TabIndex        =   34
      Top             =   3375
      Width           =   1125
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
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
      Left            =   120
      TabIndex        =   33
      Top             =   3045
      Width           =   510
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
      Left            =   120
      TabIndex        =   32
      Top             =   1065
      Width           =   615
   End
   Begin VB.Label lblFromTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From / To Date"
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
      Left            =   120
      TabIndex        =   31
      Top             =   2715
      Width           =   1095
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Final Sort"
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
      Left            =   120
      TabIndex        =   30
      Top             =   6585
      Width           =   660
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #1"
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
      Left            =   120
      TabIndex        =   29
      Top             =   6240
      Width           =   885
   End
   Begin VB.Label lblRepGrp 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Grouping"
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
      Top             =   5970
      Width           =   1575
   End
   Begin VB.Label lblSelCri 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   27
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
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
      Left            =   120
      TabIndex        =   26
      Top             =   2385
      Width           =   1290
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   120
      TabIndex        =   25
      Top             =   1725
      Width           =   450
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
      Left            =   120
      TabIndex        =   24
      Top             =   1395
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
      Left            =   120
      TabIndex        =   23
      Top             =   735
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
      Left            =   120
      TabIndex        =   22
      Top             =   405
      Width           =   555
   End
End
Attribute VB_Name = "frmRGapAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Succession Review Report", Me) Then Exit Sub
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Call set_PrintState(False)
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    Screen.MousePointer = HOURGLASS
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
    Call set_PrintState(True)
End If

Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ENTITLEMENTS", "VIEW")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub comEmpType_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem "Employee Name"
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    comGroup(1).AddItem "Employee Name"
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    If intIdx% = 0 Then strCd$ = "HREMP.ED_LOC"
    If intIdx% = 1 Then strCd$ = "HREMP.ED_ORG"
    If intIdx% = 2 Then strCd$ = "HREMP.ED_EMP"
    If intIdx% = 7 Then strCd$ = "HREMP.ED_REGION"
    If intIdx% = 8 Then strCd$ = "HREMP.ED_ADMINBY"
    If intIdx% = 9 Then strCd$ = "HREMP.ED_SECTION"
        CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
    End If
End If

If Len(CodeCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CodeCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
    End If
    glbiOneWhere = True
End If


End Sub

Private Sub Cri_Div()

Dim DivCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level


If Len(clpDiv.Text) > 0 Then
    DivCri = "({HREMP.ED_DIV} in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
End If

If Len(DivCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = DivCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & DivCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID.Text) > 0 Then
    EECri = "{HREMP.ED_EMPNBR} IN [" & getEmpnbr(elpEEID.Text) & "] "
End If


If Len(EECri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HR_Succession_WRK.EU_CSR_DATE} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
End If

For x% = 0 To 1
    If Len(dlpDateRange(0).Text) > 0 Then
        TempCri = "({HR_Succession_WRK.EU_CSR_DATE}  "
        If x% = 0 Then
            TempCri = TempCri & " >= "
        Else
            TempCri = TempCri & " <= "
        End If
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
Next x%


Cri_FTDatst:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_PT()
Dim EECri As String, OneSet%, x%

If Len(clpPT.Text) < 1 Then Exit Sub

EECri = "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If

glbiOneWhere = True

End Sub

Private Function Cri_SetAll()
Dim x%, strRName$

Cri_SetAll = False
On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
For x% = 0 To 2
    Call Cri_Code(x%)
Next x%
Call Cri_Succession
Call Cri_Code(7)
Call Cri_Code(8)
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
Call Cri_Code(9)
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_FTDates
' report name
Call SETWRK

If comGroup(0) <> "(none)" Then
    strRName$ = glbIHRREPORTS & "rzGapAnalysis.rpt"
Else
    strRName$ = glbIHRREPORTS & "rzGapAnalysis1.rpt"
End If
Me.vbxCrystal.ReportFileName = strRName$
' set to sorting/grouping criteria
x% = Cri_Sorts()   ' returns number of sections formated

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    Me.vbxCrystal.DataFiles(0) = glbIHRDB
    Me.vbxCrystal.DataFiles(1) = glbIHRDB
    Me.vbxCrystal.DataFiles(2) = glbIHRDB
    Me.vbxCrystal.DataFiles(3) = glbIHRDB
    Me.vbxCrystal.DataFiles(4) = glbIHRDB
    Me.vbxCrystal.DataFiles(5) = glbIHRDB
    Me.vbxCrystal.DataFiles(6) = glbIHRDB
    Me.vbxCrystal.DataFiles(7) = glbIHRDB
    Me.vbxCrystal.DataFiles(8) = glbIHRDB
    Me.vbxCrystal.DataFiles(9) = glbIHRDB
    Me.vbxCrystal.DataFiles(10) = glbIHRDB
    Me.vbxCrystal.DataFiles(11) = glbIHRDB
    Me.vbxCrystal.DataFiles(12) = glbIHRDB
    Me.vbxCrystal.DataFiles(13) = glbIHRDBW
    Me.vbxCrystal.DataFiles(14) = glbIHRDB
    Me.vbxCrystal.DataFiles(15) = glbIHRDB
    Me.vbxCrystal.DataFiles(16) = glbIHRDB
End If


' window title if appropriate
Me.vbxCrystal.WindowTitle = "Gap Analysis Report"

Cri_SetAll = True
Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Comp Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub Cri_Succession()
'Laura
Dim EECri As String, x%
Dim strC2, strCx As String
Dim strCa$


strCa$ = "HR_Succession_WRK.EU_JOBPREF1"

If Len(clpPosCode.Text) > 0 Then
      EECri = EECri & "({" & strCa$ & "} = '" & clpPosCode.Text & "')"
End If

If chkRelocate.Value <> 0 Then
    If glbOracle Then
        If Len(EECri) > 0 Then
            EECri = "(" & EECri & ") and ({HR_Succession_WRK.EU_RELOC}=1)"
        Else
            EECri = "({HR_Succession_WRK.EU_RELOC}=1)"
        End If
    Else
        If Len(EECri) > 0 Then
            EECri = "(" & EECri & ") and ({HR_Succession_WRK.EU_RELOC})"
        Else
            EECri = "({HR_Succession_WRK.EU_RELOC})"
        End If
    End If
End If
If Len(txtLocation.Text) > 0 Then
    If Len(EECri) > 0 Then
        EECri = "(" & EECri & ") and ({HR_Succession_WRK.EU_Location} = '" & Trim(txtLocation.Text) & ")"
    Else
        EECri = "({HR_Succession_WRK.EU_Location} = '" & Trim(txtLocation.Text) & ")"
    End If
End If
If optGrouping(0).Value Then
    If glbOracle Then
        If Len(EECri) > 0 Then
            EECri = "(" & EECri & ") and ({HR_Succession_WRK.EU_LAST_RVW}=1)"
        Else
            EECri = "({HR_Succession_WRK.EU_LAST_RVW}=1)"
        End If
    Else
        If Len(EECri) > 0 Then
            EECri = "(" & EECri & ") and ({HR_Succession_WRK.EU_LAST_RVW})"
        Else
            EECri = "({HR_Succession_WRK.EU_LAST_RVW})"
        End If
    End If
End If
'
If chkNotMatched.Value <> 0 Then
    If glbOracle Then
        If Len(EECri) > 0 Then
            EECri = "(" & EECri & ") and ({HR_Succession_WRK.EU_Match}=0)"
        Else
            EECri = "({HR_Succession_WRK.EU_Match}=0)"
        End If
    Else
        If Len(EECri) > 0 Then
            EECri = "(" & EECri & ") and (not {HR_Succession_WRK.EU_Match})"
        Else
            EECri = "(not {HR_Succession_WRK.EU_Match})"
        End If
    End If
End If

If Len(EECri) > 0 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & "(" & EECri & ")"
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, x%

If Len(txtShift.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
'imbeded in report

Cri_Sorts = 0
'first set primary grouping
Y% = 0
    grpField$ = getEGroup(comGroup(0).Text)
    If comGroup(0) = "(none)" Then Exit Function
    
    Y% = x% + 1
    dscGroup$ = comGroup(x%).Text
    dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(x%) = dscGroup$
    
    grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(x%) = grpCond$
    
    strSFormat$ = "GH1;T;T;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1
    strSFormat$ = "GF1;T;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1
    'final sort
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
    End Select
    grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(1) = grpCond$
      
Cri_Sorts = z% ' next section number to format

End Function

Private Function CriCheck()
Dim x%

CriCheck = False

If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    'MsgBox lStr("If Division Entered - it must be known")
    'clpDiv.SetFocus
    Exit Function
End If

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known"
    'clpDept.SetFocus
    Exit Function
End If

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

For x% = 0 To 2
    If Not clpCode(x).ListChecker Then Exit Function
Next x%

For x% = 7 To 9
    If Not clpCode(x).ListChecker Then Exit Function
Next x%


For x% = 0 To 1
 If Len(dlpDateRange(x%).Text) > 0 Then
    If Not IsDate(dlpDateRange(x%).Text) Then
        MsgBox "Not a valid date"
        dlpDateRange(x%).Text = ""
        dlpDateRange(x%).SetFocus
        Exit Function
    End If
 End If
Next x%
If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If
 
CriCheck = True
End Function

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = Me.name

Screen.MousePointer = HOURGLASS

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call comGrpLoad

Call setRptCaption(Me)

If glbLinamar Then clpCode(7).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(7).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub

Private Sub txtShift_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = Reports
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = False
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = False
End Property

Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Call set_Buttons
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub SETWRK()
Dim rsHS As New ADODB.Recordset, rsTmp As New ADODB.Recordset
Dim rsCode As New ADODB.Recordset, rsBODY As New ADODB.Recordset
Dim SQLQ, xNum, xRecNum, SQLQ1, SQLQ2, SQLQ3
Dim xFieldList, xJobCode, iPref
xFieldList = Get_Fields(gdbAdoIhr001, "HR_Succession", "EU_ID,EU_JOBPREF1,EU_JOBPREF2,EU_JOBPREF3,EU_JOBPREF4,EU_JOBPREF5")
'xFieldList = ""

'On Error GoTo AttWrkError
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3
gdbAdoIhr001.CommandTimeout = 600
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).FloodPercent = 15

gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "DELETE FROM HR_Succession_WRK WHERE EU_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001W.CommitTrans

Call Pause(1)

MDIMain.panHelp(0).FloodPercent = 30

'for active employees
SQLQ1 = "INSERT INTO HR_Succession_WRK  (" & xFieldList
SQLQ1 = SQLQ1 & ",EU_SKILL_TABL,EU_SKILL,EU_Match,EU_WRKEMP,EU_JOBPREF1) " & in_SQL(glbIHRDBW)

SQLQ1 = SQLQ1 & " Select " & xFieldList & ","
Select Case gsSystemDb
Case "MS SQL SERVER"
    SQLQ1 = SQLQ1 & " HRJOBSKL.JS_SKILL_TABL as EU_SKILL_TABL,HRJOBSKL.JS_SKILL AS EU_SKILL, "
    SQLQ1 = SQLQ1 & " (CASE WHEN (JS_SKILL IN (SELECT se_skill From hrempskl WHERE se_empnbr = eu_empnbr)) THEN 1 ELSE 0 END) AS EU_Match"
Case "ORACLE"
    SQLQ1 = SQLQ1 & " HRJOBSKL.JS_SKILL_TABL as EU_SKILL_TABL,HRJOBSKL.JS_SKILL AS EU_SKILL, "
    SQLQ1 = SQLQ1 & " (CASE WHEN (JS_SKILL IN (SELECT se_skill From hrempskl WHERE se_empnbr = eu_empnbr)) THEN 1 ELSE 0 END) AS EU_Match"
Case Else
    SQLQ1 = SQLQ1 & " HRJOBSKL.JS_SKILL_TABL as EU_SKILL_TABL,HRJOBSKL.JS_SKILL AS EU_SKILL, "
    SQLQ1 = SQLQ1 & " IIF((JS_SKILL IN (SELECT se_skill From hrempskl WHERE se_empnbr = eu_empnbr)),1, 0) AS EU_Match"
End Select
SQLQ1 = SQLQ1 & ",'" & glbUserID & "' AS EE_WRKEMP, "
    
Select Case gsSystemDb
Case "MS SQL SERVER"
    SQLQ2 = " FROM HR_SUCCESSION LEFT OUTER JOIN HRJOBSKL ON HR_SUCCESSION.EU_JOBPREF"
Case "ORACLE"
    SQLQ2 = " FROM HR_SUCCESSION,HRJOBSKL "
Case Else
    SQLQ2 = " FROM HR_SUCCESSION LEFT OUTER JOIN HRJOBSKL ON HR_SUCCESSION.EU_JOBPREF"
End Select

SQLQ3 = ""
If Len(elpEEID.Text) > 0 Then
    SQLQ3 = SQLQ3 & " and HR_SUCCESSION.EU_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
End If
    
'    SQLQ = SQLQ & xJobCode & " AS EU_JOBPREF1 "

SQLQ = "SELECT EU_EMPNBR,EU_JOBPREF1, EU_JOBPREF2,EU_JOBPREF3, EU_JOBPREF4,EU_JOBPREF5,EU_CSR_DATE FROM HR_Succession "
SQLQ = SQLQ & " WHERE (EU_JOBPREF1 IS NOT NULL) OR (EU_JOBPREF2 IS NOT NULL) OR (EU_JOBPREF3 IS NOT NULL) OR (EU_JOBPREF4 IS NOT NULL) OR (EU_JOBPREF5 IS NOT NULL) "
'SQLQ = SQLQ & ""

If Len(elpEEID.Text) > 0 Then
    SQLQ = SQLQ & " AND HR_SUCCESSION.EU_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
End If
rsTmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly

Do While Not rsTmp.EOF
    For iPref = 1 To 5
    If Not IsNull(rsTmp("EU_JOBPREF" & iPref)) Then
        xJobCode = rsTmp("EU_JOBPREF" & iPref)
        If gsSystemDb = "ORACLE" Then
            SQLQ = SQLQ1 & "'" & xJobCode & "' AS EU_JOBPREF " & SQLQ2 & " where EU_EMPNBR=" & rsTmp("EU_EMPNBR") & " and HR_SUCCESSION.EU_JOBPREF" & iPref & " = HRJOBSKL.JS_CODE  and HR_SUCCESSION.EU_JOBPREF" & iPref & " = '" & xJobCode & "'"
            If Not IsNull(rsTmp("EU_CSR_DATE")) Then SQLQ = SQLQ & " and HR_SUCCESSION.EU_CSR_DATE =" & Date_SQL(rsTmp("EU_CSR_DATE"))
            SQLQ = SQLQ & SQLQ3
        Else
            SQLQ = SQLQ1 & "'" & xJobCode & "' AS EU_JOBPREF " & SQLQ2 & iPref & " = HRJOBSKL.JS_CODE where EU_EMPNBR=" & rsTmp("EU_EMPNBR") & " and HR_SUCCESSION.EU_JOBPREF" & iPref & " = '" & xJobCode & "' "
            If Not IsNull(rsTmp("EU_CSR_DATE")) Then SQLQ = SQLQ & " and HR_SUCCESSION.EU_CSR_DATE =" & Date_SQL(rsTmp("EU_CSR_DATE"))
            SQLQ = SQLQ & SQLQ3
        'SQLQ = SQLQ1 & "'" & xJobCode & "' AS EU_JOBPREF1 " & SQLQ2 & iPref & " = HRJOBSKL.JS_CODE where EU_EMPNBR=" & rsTmp("EU_EMPNBR") & " and HR_SUCCESSION.EU_JOBPREF" & iPref & " = '" & xJobCode & "' " & SQLQ3
        End If
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
    End If
    Next
'    If Not IsNull(rsTmp("EU_JOBPREF2")) Then
'        xJobCode = rsTmp("EU_JOBPREF2")
'        SQLQ = SQLQ1 & "'" & xJobCode & "' AS EU_JOBPREF1 " & SQLQ2 & "2" & " = HRJOBSKL.JS_CODE where EU_EMPNBR=" & rsTmp("EU_EMPNBR") & " " & SQLQ3
'        gdbAdoIhr001.BeginTrans
'        gdbAdoIhr001.Execute SQLQ
'        gdbAdoIhr001.CommitTrans
'    End If
'    If Not IsNull(rsTmp("EU_JOBPREF3")) Then
'        xJobCode = rsTmp("EU_JOBPREF3")
'        SQLQ = SQLQ1 & "'" & xJobCode & "' AS EU_JOBPREF1 " & SQLQ2 & "3" & " = HRJOBSKL.JS_CODE where EU_EMPNBR=" & rsTmp("EU_EMPNBR") & " " & SQLQ3
'        gdbAdoIhr001.BeginTrans
'        gdbAdoIhr001.Execute SQLQ
'        gdbAdoIhr001.CommitTrans
'    End If
'    If Not IsNull(rsTmp("EU_JOBPREF4")) Then
'        xJobCode = rsTmp("EU_JOBPREF4")
'        SQLQ = SQLQ1 & "'" & xJobCode & "' AS EU_JOBPREF1 " & SQLQ2 & "4" & " = HRJOBSKL.JS_CODE where EU_EMPNBR=" & rsTmp("EU_EMPNBR") & " " & SQLQ3
'        gdbAdoIhr001.BeginTrans
'        gdbAdoIhr001.Execute SQLQ
'        gdbAdoIhr001.CommitTrans
'    End If
'    If Not IsNull(rsTmp("EU_JOBPREF5")) Then
'        xJobCode = rsTmp("EU_JOBPREF5")
'        SQLQ = SQLQ1 & "'" & xJobCode & "' AS EU_JOBPREF1 " & SQLQ2 & "5" & " = HRJOBSKL.JS_CODE where EU_EMPNBR=" & rsTmp("EU_EMPNBR") & " " & SQLQ3
'        gdbAdoIhr001.BeginTrans
'        gdbAdoIhr001.Execute SQLQ
'        gdbAdoIhr001.CommitTrans
'    End If
    
    rsTmp.MoveNext
Loop
rsTmp.Close
Set rsTmp = Nothing
MDIMain.panHelp(0).FloodPercent = 50


''for terminated employees
    
Select Case gsSystemDb
Case "MS SQL SERVER"
    SQLQ2 = " FROM TERM_HR_SUCCESSION LEFT OUTER JOIN HRJOBSKL ON TERM_HR_SUCCESSION.EU_JOBPREF"
Case "ORACLE"
    SQLQ2 = " FROM TERM_HR_SUCCESSION,HRJOBSKL ON TERM_HR_SUCCESSION.EU_JOBPREF"
Case Else
    SQLQ2 = " FROM TERM_HR_SUCCESSION LEFT OUTER JOIN HRJOBSKL ON TERM_HR_SUCCESSION.EU_JOBPREF"
End Select

SQLQ3 = ""
If Len(elpEEID.Text) > 0 Then
    SQLQ3 = SQLQ3 & "and TERM_HR_SUCCESSION.EU_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
End If
    
'    SQLQ = SQLQ & xJobCode & " AS EU_JOBPREF1 "

SQLQ = "SELECT TERM_SEQ,EU_EMPNBR,EU_JOBPREF1, EU_JOBPREF2,EU_JOBPREF3, EU_JOBPREF4,EU_JOBPREF5 FROM TERM_HR_SUCCESSION "
'SQLQ = SQLQ & ""

If Len(elpEEID.Text) > 0 Then
    SQLQ = SQLQ & "WHERE TERM_HR_SUCCESSION.EU_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
End If
rsTmp.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly, adLockReadOnly

Do While Not rsTmp.EOF
    If Not IsNull(rsTmp("EU_JOBPREF1")) Then
        xJobCode = rsTmp("EU_JOBPREF1")
        SQLQ = SQLQ1 & "'" & xJobCode & "' AS EU_JOBPREF1 " & SQLQ2 & "1" & " = HRJOBSKL.JS_CODE where TERM_SEQ=" & rsTmp("TERM_SEQ") & " " & SQLQ3
        gdbAdoIhr001X.BeginTrans
        gdbAdoIhr001X.Execute SQLQ
        gdbAdoIhr001X.CommitTrans
    End If
    If Not IsNull(rsTmp("EU_JOBPREF2")) Then
        xJobCode = rsTmp("EU_JOBPREF2")
        SQLQ = SQLQ1 & "'" & xJobCode & "' AS EU_JOBPREF1 " & SQLQ2 & "2" & " = HRJOBSKL.JS_CODE where TERM_SEQ=" & rsTmp("TERM_SEQ") & " " & SQLQ3
        gdbAdoIhr001X.BeginTrans
        gdbAdoIhr001X.Execute SQLQ
        gdbAdoIhr001X.CommitTrans
    End If
    If Not IsNull(rsTmp("EU_JOBPREF3")) Then
        xJobCode = rsTmp("EU_JOBPREF3")
        SQLQ = SQLQ1 & "'" & xJobCode & "' AS EU_JOBPREF1 " & SQLQ2 & "3" & " = HRJOBSKL.JS_CODE where TERM_SEQ=" & rsTmp("TERM_SEQ") & " " & SQLQ3
        gdbAdoIhr001X.BeginTrans
        gdbAdoIhr001X.Execute SQLQ
        gdbAdoIhr001X.CommitTrans
    End If
    If Not IsNull(rsTmp("EU_JOBPREF4")) Then
        xJobCode = rsTmp("EU_JOBPREF4")
        SQLQ = SQLQ1 & "'" & xJobCode & "' AS EU_JOBPREF1 " & SQLQ2 & "4" & " = HRJOBSKL.JS_CODE where TERM_SEQ=" & rsTmp("TERM_SEQ") & " " & SQLQ3
        gdbAdoIhr001X.BeginTrans
        gdbAdoIhr001X.Execute SQLQ
        gdbAdoIhr001X.CommitTrans
    End If
    If Not IsNull(rsTmp("EU_JOBPREF5")) Then
        xJobCode = rsTmp("EU_JOBPREF5")
        SQLQ = SQLQ1 & "'" & xJobCode & "' AS EU_JOBPREF1 " & SQLQ2 & "5" & " = HRJOBSKL.JS_CODE where TERM_SEQ=" & rsTmp("TERM_SEQ") & " " & SQLQ3
        gdbAdoIhr001X.BeginTrans
        gdbAdoIhr001X.Execute SQLQ
        gdbAdoIhr001X.CommitTrans
    End If
    
    rsTmp.MoveNext
Loop
rsTmp.Close
Set rsTmp = Nothing

MDIMain.panHelp(0).FloodPercent = 85
'If chkReport(0) Or optSum(2) Then
    MDIMain.panHelp(0).FloodPercent = 90
    Call Pause(2)
    MDIMain.panHelp(0).FloodPercent = 100
    Call Pause(2)

    Call Pause(1)
'End If

MDIMain.panHelp(0).FloodPercent = 100
Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
End Sub

