VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmRLabels 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employee Labels"
   ClientHeight    =   9240
   ClientLeft      =   465
   ClientTop       =   645
   ClientWidth     =   10395
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
   ScaleHeight     =   9240
   ScaleWidth      =   10395
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkExclCONP 
      Caption         =   "Exclude Employment Status of CONP"
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
      Left            =   120
      TabIndex        =   20
      Tag             =   "Check to Exclude Employees with CONP Employment Status"
      Top             =   4920
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.CheckBox chkExclRET 
      Caption         =   "Exclude Employment Status of RET"
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
      Left            =   120
      TabIndex        =   21
      Tag             =   "Check to Exclude Employees with RET Employment Status"
      Top             =   5220
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.CheckBox chkPrintSalute 
      Caption         =   "Print Name with Salutation"
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
      Left            =   6840
      TabIndex        =   16
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Frame fraTerm 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2160
      TabIndex        =   43
      Top             =   4440
      Visible         =   0   'False
      Width           =   7575
      Begin VB.OptionButton optTerm 
         Caption         =   "Only Terminated Employees"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   19
         Top             =   0
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.OptionButton optTerm 
         Caption         =   "Include Terminated Employees"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.CheckBox chkNaLoc 
      Caption         =   "Print Name and Location"
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
      Left            =   4080
      TabIndex        =   15
      Top             =   4080
      Width           =   2535
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
      Left            =   1890
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Tag             =   "First Level of grouping records"
      Top             =   6930
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
      Index           =   1
      Left            =   1890
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Tag             =   "Second level of grouping records"
      Top             =   7245
      Width           =   2325
   End
   Begin VB.CheckBox chkTermEmp 
      Caption         =   "Terminated Employees"
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
      Left            =   120
      TabIndex        =   17
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CheckBox chkOMERS 
      Caption         =   "OMERS Date"
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
      Left            =   1920
      TabIndex        =   14
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1875
      MaxLength       =   2
      TabIndex        =   10
      Tag             =   "00-Employee Position Shift"
      Top             =   3640
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkName 
      Caption         =   "Print Name only"
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
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   1575
   End
   Begin Threed.SSOption optRtype 
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Tag             =   "You can choose one of the options"
      Top             =   5760
      Width           =   3045
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "    Avery Labels # 5161 (2 Across)"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSOption optRtype 
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "You can choose one of the options"
      Top             =   6060
      Width           =   3045
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "    Avery Labels # 5160 (3 Across)"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optRtype 
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "You can choose one of the options"
      Top             =   6360
      Visible         =   0   'False
      Width           =   3045
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "    Avery Labels # 5267 (4 Across)"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1560
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
      Left            =   1560
      TabIndex        =   6
      Tag             =   "EDPT-Category"
      Top             =   2340
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
      Left            =   1560
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
      Left            =   1560
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
      Left            =   1560
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
      Left            =   1560
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
      Index           =   4
      Left            =   1560
      TabIndex        =   8
      Tag             =   "00-Enter Administered By Code"
      Top             =   3000
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   9
      Tag             =   "00-Enter Section Code"
      Top             =   3330
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   7
      Tag             =   "00-Enter Region Code"
      Top             =   2670
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Tag             =   "10-Enter Employee Number"
      Top             =   2010
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6720
      Top             =   6840
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
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   3
      Left            =   6690
      TabIndex        =   12
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3675
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   2
      Left            =   4980
      TabIndex        =   11
      Tag             =   "40-Date from and including this date forward"
      Top             =   3675
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   503
      TextBoxWidth    =   1215
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
      TabIndex        =   42
      Top             =   6960
      Width           =   885
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
      Index           =   3
      Left            =   120
      TabIndex        =   41
      Top             =   7275
      Width           =   660
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
      TabIndex        =   40
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lblEmplStFrpmTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Status From / To Date"
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
      Left            =   3300
      TabIndex        =   39
      Top             =   3720
      Width           =   1635
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
      TabIndex        =   38
      Top             =   3645
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
      TabIndex        =   37
      Top             =   2340
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
      TabIndex        =   36
      Top             =   3360
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
      TabIndex        =   35
      Top             =   3000
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
      TabIndex        =   34
      Top             =   2670
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
      TabIndex        =   33
      Top             =   1020
      Width           =   615
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
      TabIndex        =   32
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
      TabIndex        =   31
      Top             =   2010
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
      TabIndex        =   30
      Top             =   1680
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
      TabIndex        =   29
      Top             =   1350
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
      TabIndex        =   28
      Top             =   690
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
      TabIndex        =   27
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "frmRLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ODIV, ODivD, xGlbDiv, xGlbDivDesc, xKeyPress, xTxtDiv, xLblDivDesc  'js-26Apr99
Dim LastlastID, LastlastNme, LastFirstNme, xTxtEEID, xLblEEName

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If chkName.Value = 0 Then
        If Not gSec_Show_ADDRESS Then
            MsgBox "You Do Not Have Authority For This Transaction"
            Exit Sub
        End If
    End If

    If Not PrtForm("Employee Labels Criteria", Me) Then Exit Sub
    x% = Cri_SetAll()
      Call set_PrintState(False)
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
      Call set_PrintState(True)
End If
Exit Sub

PrntErr:
MsgBox "Error Printing - check your Windows Printer setup"
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
    If chkName.Value = 0 Then
        If Not gSec_Show_ADDRESS Then
            MsgBox "You Do Not Have Authority For This Transaction"
            Exit Sub
        End If
    End If

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    Screen.MousePointer = HOURGLASS
    x% = Cri_SetAll()
    Call set_PrintState(False)
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ATTEND", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim strCd$
If Len(clpCode(intIdx%).Text) > 0 Then
    If intIdx% = 0 Then strCd$ = "HREMP.ED_LOC"
    If intIdx% = 1 Then strCd$ = "HREMP.ED_ORG"
    If intIdx% = 2 Then strCd$ = "HREMP.ED_EMP"
    If intIdx% = 3 Then strCd$ = "HREMP.ED_REGION"
    If intIdx% = 4 Then strCd$ = "HREMP.ED_ADMINBY"
    If intIdx% = 5 Then strCd$ = "HREMP.ED_SECTION"
        
        
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
    Else
        CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    End If
End If


If Len(CodeCri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
    Else
        glbstrSelCri = CodeCri
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
Dim SQLQ As String, strCriteria
Dim rsIn As New ADODB.Recordset, rsOUT As New ADODB.Recordset
Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_EE
Call Cri_PT
Call Cri_Shift
Call Cri_EmpStatFTDates
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For x% = 0 To 5
    Cri_Code (x%)
Next
Call Cri_Shift
Call Cri_OMERSDate 'Ticket# 7954


If Not (glbSQL Or glbOracle) And chkTermEmp And optTerm(0) Then
    SQLQ = "DELETE FROM HR_HRLABEL_WRK WHERE ED_WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001W.Execute SQLQ
    
    strCriteria = Replace(Replace(Replace(Replace(glbstrSelCri, "{", "("), "}", ")"), "[", "("), "]", ")")
    SQLQ = "SELECT ED_COMPNO, ED_EMPNBR, ED_SURNAME, ED_FNAME, ED_EMPTYPE, ED_ADDR1, ED_ADDR2, ED_CITY, ED_PCODE, ED_PROV, ED_BUSNBR, ED_PHONE, "
    SQLQ = SQLQ & "ED_CELLPHONE, ED_COUNTRY,ED_TITLE , ED_DOH, ED_DIV, ED_DEPTNO, ED_LOC, ED_SECTION, ED_EMP,ED_REGION, ED_SEX "
    SQLQ = SQLQ & "FROM HREMP WHERE " & strCriteria
    rsIn.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rsIn.EOF = False And rsIn.BOF = False Then
        rsOUT.Open "HR_HRLABEL_WRK", gdbAdoIhr001W, adOpenDynamic, adLockOptimistic, adCmdTable
        Do
            rsOUT.AddNew
                rsOUT("ED_COMPNO") = rsIn("ED_COMPNO")
                rsOUT("ED_EMPNBR") = rsIn("ED_EMPNBR")
                rsOUT("ED_SURNAME") = rsIn("ED_SURNAME")
                rsOUT("ED_FNAME") = rsIn("ED_FNAME")
                rsOUT("ED_EMPTYPE") = rsIn("ED_EMPTYPE")
                rsOUT("ED_ADDR1") = rsIn("ED_ADDR1")
                rsOUT("ED_ADDR2") = rsIn("ED_ADDR2")
                rsOUT("ED_CITY") = rsIn("ED_CITY")
                rsOUT("ED_PCODE") = rsIn("ED_PCODE")
                rsOUT("ED_PROV") = rsIn("ED_PROV")
                rsOUT("ED_BUSNBR") = rsIn("ED_BUSNBR")
                rsOUT("ED_PHONE") = rsIn("ED_PHONE")
                rsOUT("ED_CELLPHONE") = rsIn("ED_CELLPHONE")
                rsOUT("ED_COUNTRY") = rsIn("ED_COUNTRY")
                rsOUT("ED_TITLE") = rsIn("ED_TITLE")
                rsOUT("ED_DOH") = rsIn("ED_DOH")
                rsOUT("ED_DIV") = rsIn("ED_DIV")
                rsOUT("ED_DEPTNO") = rsIn("ED_DEPTNO")
                rsOUT("ED_LOC") = rsIn("ED_LOC")
                rsOUT("ED_SECTION") = rsIn("ED_SECTION")
                rsOUT("ED_EMP") = rsIn("ED_EMP")
                rsOUT("ED_REGION") = rsIn("ED_REGION")
                rsOUT("ED_WRKEMP") = glbUserID
                rsOUT("ED_SEX") = rsIn("ED_SEX")
            rsOUT.Update
            rsIn.MoveNext
        Loop Until rsIn.EOF
        rsOUT.Close
    End If
    rsIn.Close
    
    strCriteria = Replace(strCriteria, "HREMP.", "Term_HREMP.")
    SQLQ = "SELECT ED_COMPNO, ED_EMPNBR, ED_SURNAME, ED_FNAME, ED_EMPTYPE, ED_ADDR1, ED_ADDR2, ED_CITY, ED_PCODE, ED_PROV, ED_BUSNBR, ED_PHONE, "
    SQLQ = SQLQ & "ED_CELLPHONE, ED_COUNTRY,ED_TITLE , ED_DOH, ED_DIV, ED_DEPTNO, ED_LOC, ED_SECTION, ED_EMP,ED_REGION, ED_SEX "
    SQLQ = SQLQ & "FROM Term_HREMP WHERE " & strCriteria
    rsIn.Open SQLQ, gdbAdoIhr001X, adOpenStatic, adLockOptimistic, adCmdText
    If rsIn.EOF = False And rsIn.BOF = False Then
        rsOUT.Open "HR_HRLABEL_WRK", gdbAdoIhr001W, adOpenDynamic, adLockOptimistic, adCmdTable
        Do
            rsOUT.AddNew
                rsOUT("ED_COMPNO") = rsIn("ED_COMPNO")
                rsOUT("ED_EMPNBR") = rsIn("ED_EMPNBR")
                rsOUT("ED_SURNAME") = rsIn("ED_SURNAME")
                rsOUT("ED_FNAME") = rsIn("ED_FNAME")
                rsOUT("ED_EMPTYPE") = rsIn("ED_EMPTYPE")
                rsOUT("ED_ADDR1") = rsIn("ED_ADDR1")
                rsOUT("ED_ADDR2") = rsIn("ED_ADDR2")
                rsOUT("ED_CITY") = rsIn("ED_CITY")
                rsOUT("ED_PCODE") = rsIn("ED_PCODE")
                rsOUT("ED_PROV") = rsIn("ED_PROV")
                rsOUT("ED_BUSNBR") = rsIn("ED_BUSNBR")
                rsOUT("ED_PHONE") = rsIn("ED_PHONE")
                rsOUT("ED_CELLPHONE") = rsIn("ED_CELLPHONE")
                rsOUT("ED_COUNTRY") = rsIn("ED_COUNTRY")
                rsOUT("ED_TITLE") = rsIn("ED_TITLE")
                rsOUT("ED_DOH") = rsIn("ED_DOH")
                rsOUT("ED_DIV") = rsIn("ED_DIV")
                rsOUT("ED_DEPTNO") = rsIn("ED_DEPTNO")
                rsOUT("ED_LOC") = rsIn("ED_LOC")
                rsOUT("ED_SECTION") = rsIn("ED_SECTION")
                rsOUT("ED_EMP") = rsIn("ED_EMP")
                rsOUT("ED_REGION") = rsIn("ED_REGION")
                rsOUT("ED_SEX") = rsIn("ED_SEX")
                rsOUT("ED_WRKEMP") = glbUserID
            rsOUT.Update
            rsIn.MoveNext
        Loop Until rsIn.EOF
        rsOUT.Close
    End If
    rsIn.Close

End If


' report name
If chkPrintSalute Then  'Print Salutation infront of the Employee Name
    If chkTermEmp Then 'Terminated Employee
        If optTerm(1) Then 'ONLY Terminated
            If chkName.Value = 0 Then
                If chkNaLoc.Value = 0 Then
                    If optRtype(0).Value Then
                        strRName$ = glbIHRREPORTS & "rzl5161t.rpt"
                    ElseIf optRtype(1).Value Then
                        strRName$ = glbIHRREPORTS & "rzl5160t.rpt"
                    'Else
                    '    strRName$ = glbIHRREPORTS & "rzl5267t.rpt"
                    End If
                Else 'Ticket #12493
                    If optRtype(0).Value Then
                        strRName$ = glbIHRREPORTS & "rzl5161nlt.rpt"
                    Else
                        strRName$ = glbIHRREPORTS & "rzl5160nlt.rpt"
                    End If
                End If
            Else
                If optRtype(0).Value Then
                    strRName$ = glbIHRREPORTS & "rzl5161nt.rpt"
                Else
                    strRName$ = glbIHRREPORTS & "rzl5160nt.rpt"
                End If
            End If
        Else 'Include Terminated
            If chkName.Value = 0 Then
                If chkNaLoc.Value = 0 Then
                    If optRtype(0).Value Then
                        strRName$ = glbIHRREPORTS & "rzl5161ta.rpt"
                    ElseIf optRtype(1).Value Then
                        strRName$ = glbIHRREPORTS & "rzl5160ta.rpt"
                    'Else
                    '    strRName$ = glbIHRREPORTS & "rzl5267t.rpt"
                    End If
                Else 'Ticket #12493
                    If optRtype(0).Value Then
                        strRName$ = glbIHRREPORTS & "rzl5161nlta.rpt"
                    Else
                        strRName$ = glbIHRREPORTS & "rzl5160nlta.rpt"
                    End If
                End If
            Else
                If optRtype(0).Value Then
                    strRName$ = glbIHRREPORTS & "rzl5161nta.rpt"
                Else
                    strRName$ = glbIHRREPORTS & "rzl5160nta.rpt"
                End If
            End If
        End If
    Else 'Active Employee
        If chkName.Value = 0 Then
            If chkNaLoc.Value = 0 Then
                If optRtype(0).Value Then
                    strRName$ = glbIHRREPORTS & "rzl5161.rpt"
                ElseIf optRtype(1).Value Then
                    strRName$ = glbIHRREPORTS & "rzl5160.rpt"
                Else
                    strRName$ = glbIHRREPORTS & "rzl5267.rpt"
                End If
            Else 'Ticket #12493
                If optRtype(0).Value Then
                    strRName$ = glbIHRREPORTS & "rzl5161nl.rpt"
                Else
                    strRName$ = glbIHRREPORTS & "rzl5160nl.rpt"
                End If
            End If
        Else
            If optRtype(0).Value Then
                strRName$ = glbIHRREPORTS & "rzl5161n.rpt"
            Else
                strRName$ = glbIHRREPORTS & "rzl5160n.rpt"
            End If
        End If
    
    End If
Else
    'Do not print Salutation infront of the Employee Name
    If chkTermEmp Then 'Terminated Employee
        If optTerm(1) Then 'ONLY Terminated
            If chkName.Value = 0 Then
                If chkNaLoc.Value = 0 Then
                    If optRtype(0).Value Then
                        strRName$ = glbIHRREPORTS & "rzl5161tS.rpt"
                    ElseIf optRtype(1).Value Then
                        strRName$ = glbIHRREPORTS & "rzl5160tS.rpt"
                    'Else
                    '    strRName$ = glbIHRREPORTS & "rzl5267t.rpt"
                    End If
                Else 'Ticket #12493
                    If optRtype(0).Value Then
                        strRName$ = glbIHRREPORTS & "rzl5161nltS.rpt"
                    Else
                        strRName$ = glbIHRREPORTS & "rzl5160nltS.rpt"
                    End If
                End If
            Else
                If optRtype(0).Value Then
                    strRName$ = glbIHRREPORTS & "rzl5161ntS.rpt"
                Else
                    strRName$ = glbIHRREPORTS & "rzl5160ntS.rpt"
                End If
            End If
        Else 'Include Terminated
            If chkName.Value = 0 Then
                If chkNaLoc.Value = 0 Then
                    If optRtype(0).Value Then
                        strRName$ = glbIHRREPORTS & "rzl5161taS.rpt"
                    ElseIf optRtype(1).Value Then
                        strRName$ = glbIHRREPORTS & "rzl5160taS.rpt"
                    'Else
                    '    strRName$ = glbIHRREPORTS & "rzl5267t.rpt"
                    End If
                Else 'Ticket #12493
                    If optRtype(0).Value Then
                        strRName$ = glbIHRREPORTS & "rzl5161nltaS.rpt"
                    Else
                        strRName$ = glbIHRREPORTS & "rzl5160nltaS.rpt"
                    End If
                End If
            Else
                If optRtype(0).Value Then
                    strRName$ = glbIHRREPORTS & "rzl5161ntaS.rpt"
                Else
                    strRName$ = glbIHRREPORTS & "rzl5160ntaS.rpt"
                End If
            End If
        End If
    Else 'Active Employee
        If chkName.Value = 0 Then
            If chkNaLoc.Value = 0 Then
                If optRtype(0).Value Then
                    strRName$ = glbIHRREPORTS & "rzl5161S.rpt"
                ElseIf optRtype(1).Value Then
                    strRName$ = glbIHRREPORTS & "rzl5160S.rpt"
                Else
                    strRName$ = glbIHRREPORTS & "rzl5267.rpt"
                End If
            Else 'Ticket #12493
                If optRtype(0).Value Then
                    strRName$ = glbIHRREPORTS & "rzl5161nlS.rpt"
                Else
                    strRName$ = glbIHRREPORTS & "rzl5160nlS.rpt"
                End If
            End If
        Else
            If optRtype(0).Value Then
                strRName$ = glbIHRREPORTS & "rzl5161nS.rpt"
            Else
                strRName$ = glbIHRREPORTS & "rzl5160nS.rpt"
            End If
        End If
    End If
End If
Me.vbxCrystal.ReportFileName = strRName$

' set to sorting/grouping criteria
x% = Cri_Sorts()   ' returns number of sections formated

'Ticket #29660 - Contract Employees Enhancement
If glbWFC Then
    If chkExclCONP.Visible And chkExclRET.Visible = True Then
        If chkExclCONP Then
            If Len(glbstrSelCri) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_EMP} <> 'CONP'"
            Else
                glbstrSelCri = "{HREMP.ED_EMP} <> 'CONP'"
            End If
        End If
        If chkExclRET Then
            If Len(glbstrSelCri) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_EMP} <> 'RET'"
            Else
                glbstrSelCri = "{HREMP.ED_EMP} <> 'RET'"
            End If
        End If
    End If
End If

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    If chkTermEmp And optTerm(1) Then 'Terminated ONLY Employee
        For x% = 0 To 4
            If x% = 0 Then
                Me.vbxCrystal.DataFiles(x%) = glbIHRAUDIT
            Else
                Me.vbxCrystal.DataFiles(x%) = glbIHRDB
            End If
        Next x%
    ElseIf chkTermEmp And optTerm(0) Then 'Include Terminated
        For x% = 0 To 4
            If x% = 0 Then
                Me.vbxCrystal.DataFiles(x%) = glbIHRDBW
            Else
                Me.vbxCrystal.DataFiles(x%) = glbIHRDB
            End If
        Next x%
    Else
        For x% = 0 To 4
                Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next x%
    End If
'set security for database
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
End If

'window title if appropriate
Me.vbxCrystal.WindowTitle = "Employee Labels"

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

'    If lblEEName.Caption = "Enter Valid Employee #" Then
'        MsgBox "If Employee Entered - they must exist"
'        txtEEID.SetFocus
'        Exit Function
'    End If

For x% = 0 To 5
    If Not clpCode(x).ListChecker Then Exit Function
Next x%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

If IsDate(dlpDateRange(2)) And IsDate(dlpDateRange(3)) Then
    If DaysBetween(dlpDateRange(2), dlpDateRange(3)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(2).SetFocus                                         '
        Exit Function                                                       '
    End If
End If

CriCheck = True

End Function

Private Sub chkExclCONP_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkExclRET_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkTermEmp_Click()
If chkTermEmp Then
    fraTerm.Visible = True
    elpEEID.LookupType = TERM
Else
    fraTerm.Visible = False
    elpEEID.LookupType = 0
End If
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = "FRMRLABELS"

Screen.MousePointer = HOURGLASS

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call comGrpLoad
Call setRptCaption(Me)

chkOMERS.Caption = lStr(chkOMERS.Caption)
chkNaLoc.Caption = lStr(chkNaLoc.Caption)

If glbLinamar Then clpCode(3).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

'Ticket #29660 - Contract Employees Enhancement
If glbWFC Then
    chkExclCONP.Visible = True
    chkExclRET.Visible = True
Else
    chkExclCONP.Visible = False
    chkExclRET.Visible = False
End If

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

Private Sub optRtype_Click(Index As Integer, Value As Integer)

If optRtype(2) Then
MsgBox "Avery Labels #5267 will not" & Chr(10) & "   display 2nd line of address", 64, "Label Information"
End If

End Sub

Private Sub optRtype_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtShift_GotFocus()
Call SetPanHelp(ActiveControl)
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

Private Sub Cri_OMERSDate()
Dim EECri As String, OneSet%, x%

If chkOMERS Then

    EECri = "Not Isnull({HREMP.ED_OMERS}) "

    
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If
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

Private Sub Cri_EmpStatFTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%, x%
Dim FromDate, ToDate, SQLQ
Dim RsHRPARCO As New ADODB.Recordset

If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
    TempCri = "({HREMP.ED_SFDATE} "
    dtYYY% = Year(dlpDateRange(2).Text)
    dtMM% = month(dlpDateRange(2).Text)
    dtDD% = Day(dlpDateRange(2).Text)
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) and "
    
    dtYYY% = Year(dlpDateRange(3).Text)
    dtMM% = month(dlpDateRange(3).Text)
    dtDD% = Day(dlpDateRange(3).Text)
    TempCri = TempCri & " ({HREMP.ED_STDATE} <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    
    GoTo Cri_FTDatst
End If

If Len(dlpDateRange(2).Text) > 0 Or Len(dlpDateRange(3).Text) > 0 Then
    If Len(dlpDateRange(2).Text) > 0 Then
        TempCri = "({HREMP.ED_STDATE} "
        TempCri = TempCri & " >= "
        dtYYY% = Year(dlpDateRange(2).Text)
        dtMM% = month(dlpDateRange(2).Text)
        dtDD% = Day(dlpDateRange(2).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "

        GoTo Cri_FTDatst
    End If
    
    If Len(dlpDateRange(3).Text) > 0 Then
            TempCri = TempCri & "({HREMP.ED_SFDATE}  "
            TempCri = TempCri & " <= "
            dtYYY% = Year(dlpDateRange(3).Text)
            dtMM% = month(dlpDateRange(3).Text)
            dtDD% = Day(dlpDateRange(3).Text)
            TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "

        GoTo Cri_FTDatst
    End If
Else
    GoTo Cri_FTDatst
End If

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

Private Sub comGrpLoad()
    comGroup(0).AddItem "(none)"
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).ListIndex = 0
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Employee Number"
    comGroup(1).ListIndex = 0

End Sub

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

Cri_Sorts = 0
'first set primary grouping
Y% = 0
grpField$ = getEGroup(comGroup(0).Text)

'Call setRptLabel(Me, 0)
If comGroup(0) = "(none)" Then
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@FullName}"
        Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
    End Select
    grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$

    Exit Function
End If
    
Y% = x% + 1
dscGroup$ = comGroup(x%).Text
dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
'Me.vbxCrystal.Formulas(x%) = dscGroup$

grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(x%) = grpCond$

GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@FullName}"
    Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
End Select
grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$

Cri_Sorts = z% ' next section number to format

End Function
