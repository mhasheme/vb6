VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmREmpDates 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employee Dates Report"
   ClientHeight    =   9360
   ClientLeft      =   375
   ClientTop       =   915
   ClientWidth     =   11010
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
   ScaleHeight     =   9360
   ScaleWidth      =   11010
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
      TabIndex        =   27
      Tag             =   "Check to Exclude Employees with CONP Employment Status"
      Top             =   6960
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
      TabIndex        =   28
      Tag             =   "Check to Exclude Employees with RET Employment Status"
      Top             =   7260
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Frame frmGroup 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      TabIndex        =   43
      Top             =   8040
      Width           =   5055
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
         Left            =   1820
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Tag             =   "First Level of grouping records"
         Top             =   225
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
         Left            =   1820
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Tag             =   "Second level of grouping records"
         Top             =   565
         Width           =   2325
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
         TabIndex        =   46
         Top             =   285
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
         TabIndex        =   45
         Top             =   625
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
         TabIndex        =   44
         Top             =   0
         Width           =   1575
      End
   End
   Begin Threed.SSCheck chkOnlyTerm 
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   6480
      Visible         =   0   'False
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Only Terminated Employees"
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
   Begin Threed.SSCheck chkIncludeTerm 
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   6120
      Visible         =   0   'False
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Include Terminated Employees"
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
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2100
      MaxLength       =   4
      TabIndex        =   10
      Tag             =   "00-Employee Position Shift"
      Top             =   3660
      Visible         =   0   'False
      Width           =   435
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
      Index           =   4
      Left            =   1800
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
      Left            =   1800
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
      Left            =   1800
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
      Left            =   8640
      Top             =   7920
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
   Begin INFOHR_Controls.CodeLookup clpProv 
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Tag             =   "31-Province of Residence - Code"
      Top             =   3990
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   4
   End
   Begin INFOHR_Controls.CodeLookup clpProvEmp 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Tag             =   "31-Province of Employment - Code"
      Top             =   4320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   4
   End
   Begin INFOHR_Controls.DateLookup dlpEmpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   15
      Tag             =   "40-Date upto and including this date forward"
      Top             =   4800
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpEmpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   14
      Tag             =   "40-Date from and including this date forward"
      Top             =   4800
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpPenDateRange 
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   19
      Tag             =   "40-Date upto and including this date forward"
      Top             =   5160
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpPenDateRange 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   18
      Tag             =   "40-Date from and including this date forward"
      Top             =   5160
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpOthDateRange 
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   23
      Tag             =   "40-Date upto and including this date forward"
      Top             =   5520
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpOthDateRange 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   22
      Tag             =   "40-Date from and including this date forward"
      Top             =   5520
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin Threed.SSCheck chkPenDates 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5175
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Pension Dates"
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
   Begin Threed.SSCheck chkOthDates 
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5535
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Other Dates"
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
   Begin Threed.SSCheck chkEmpDates 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4815
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Employment Dates"
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
   Begin Threed.SSCheck chkShowAllPenDates 
      Height          =   225
      Left            =   5280
      TabIndex        =   20
      Top             =   5175
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "Show All Pension Dates"
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
   Begin Threed.SSCheck chkShowAllOthDates 
      Height          =   225
      Left            =   5280
      TabIndex        =   24
      Top             =   5550
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "Show All Other Dates"
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
   Begin Threed.SSCheck chkShowAllEmpDates 
      Height          =   225
      Left            =   5280
      TabIndex        =   16
      Top             =   4815
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "Show All Employment Dates"
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
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Prov. of Employment"
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
      TabIndex        =   48
      Top             =   4365
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Prov. of Residence"
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
      TabIndex        =   47
      Top             =   4035
      Width           =   1365
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
      TabIndex        =   42
      Top             =   3705
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
      TabIndex        =   41
      Top             =   2055
      Width           =   630
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      TabIndex        =   40
      Top             =   3375
      Width           =   540
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      TabIndex        =   39
      Top             =   2715
      Width           =   510
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      TabIndex        =   38
      Top             =   3045
      Width           =   1125
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
      TabIndex        =   37
      Top             =   1065
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
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   34
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
      TabIndex        =   33
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
      TabIndex        =   32
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
      TabIndex        =   31
      Top             =   405
      Width           =   555
   End
End
Attribute VB_Name = "frmREmpDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Excel2007 = 12
Private strGroup As String

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

'''On Error GoTo PrntErr
 
If CriCheck() Then
    'If Not PrtForm("Home Address/Telephone Report Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    X% = Cri_SetAll()
    'Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    'Me.vbxCrystal.Action = 1
    'vbxCrystal.Reset
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
Dim X%
Dim strWHand As String
'''On Error GoTo CRW_Err

If CriCheck() Then
    Call set_PrintState(False)
    'cmdPrint.Enabled = False
    'cmdView.Enabled = False
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    Screen.MousePointer = HOURGLASS
    X% = Cri_SetAll()
    'Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    'Screen.MousePointer = DEFAULT
    'Me.vbxCrystal.Action = 1
    'vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    'cmdPrint.Enabled = True
    'cmdView.Enabled = True
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

Private Sub chkEmpDates_Click(Value As Integer)
    If chkEmpDates.Value = False Then
        dlpEmpDateRange(0).Text = ""
        dlpEmpDateRange(1).Text = ""
    End If
End Sub

Private Sub chkExclCONP_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkExclRET_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkIncludeTerm_Click(Value As Integer)
    If chkIncludeTerm Then
        chkOnlyTerm.Value = False
        elpEEID.LookupType = TERM
    Else
        elpEEID.LookupType = 0
    End If
End Sub

Private Sub chkOnlyTerm_Click(Value As Integer)
    If chkOnlyTerm Then
        chkIncludeTerm.Value = False
        elpEEID.LookupType = TERM
    Else
        elpEEID.LookupType = 0
    End If
End Sub

Private Sub chkOthDates_Click(Value As Integer)
    If chkOthDates.Value = False Then
        dlpOthDateRange(0).Text = ""
        dlpOthDateRange(1).Text = ""
    End If
End Sub

Private Sub chkPenDates_Click(Value As Integer)
    If chkPenDates.Value = False Then
        dlpPenDateRange(0).Text = ""
        dlpPenDateRange(1).Text = ""
    End If
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()

    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Employee Number"
    If Not glbSQL Then
        comGroup(1).AddItem "Date of Hire"
    End If
    
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0


End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    Select Case intIdx%
    Case 0: strCd$ = "HREMP.ED_LOC"
    Case 1: strCd$ = "HREMP.ED_ORG"
    Case 2: strCd$ = "HREMP.ED_EMP"
    Case 3: strCd$ = "HREMP.ED_REGION"
    Case 4: strCd$ = "HREMP.ED_ADMINBY"
    Case 5: strCd$ = "HREMP.ED_SECTION"
    Case 6: strCd$ = "HREMP.ED_SALDIST"
    End Select
        CodeCri = "(" & strCd$ & " in  ('" & Replace(clpCode(intIdx%).Text, ",", "','") & "'))"
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "((" & strCd$ & " = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or (" & strCd$ & " = 'ALL" & clpCode(intIdx%).Text & "') )"
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
Dim countr   As Integer

If Len(clpDiv.Text) > 0 Then
    DivCri = "(HREMP.ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "'))"
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

Private Sub Cri_Dept()
Dim DivCri As String
Dim countr   As Integer

If Len(clpDept.Text) > 0 Then
    DivCri = "(HREMP.ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "'))"
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
    EECri = "HREMP.ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
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
Dim EECri As String, OneSet%, X%

If Len(clpPT.Text) < 1 Then Exit Sub

EECri = "HREMP.ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "')"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_ProvResidence()
Dim EECri As String, OneSet%, X%

If Len(clpProv.Text) < 1 Then Exit Sub
EECri = "HREMP.ED_PROV in ('" & Replace(clpProv.Text, ",", "','") & "')"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_ProvEmployment()
Dim EECri As String, OneSet%, X%

If Len(clpProvEmp.Text) < 1 Then Exit Sub
EECri = "HREMP.ED_PROVEMP in ('" & Replace(clpProvEmp.Text, ",", "','") & "')"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_Employment_Dates()
Dim TempCri As String, OneSet%, X%

    If chkEmpDates.Value = False Then Exit Sub
    
    'Date of Hire
    If Len(dlpEmpDateRange(0).Text) > 0 And Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = " (((HREMP.ED_DOH >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") AND "
        TempCri = TempCri & " (HREMP.ED_DOH <= " & Date_SQL(dlpEmpDateRange(1).Text) & ")) "
    ElseIf Len(dlpEmpDateRange(0).Text) > 0 Then
        TempCri = " (HREMP.ED_DOH >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") "
    ElseIf Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = " (HREMP.ED_DOH <= " & Date_SQL(dlpEmpDateRange(1).Text) & ") "
    End If
        
    'Seniority Date
    If Len(dlpEmpDateRange(0).Text) > 0 And Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR ((HREMP.ED_SENDTE >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") AND "
        TempCri = TempCri & " (HREMP.ED_SENDTE <= " & Date_SQL(dlpEmpDateRange(1).Text) & ")) "
    ElseIf Len(dlpEmpDateRange(0).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_SENDTE >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") "
    ElseIf Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_SENDTE <= " & Date_SQL(dlpEmpDateRange(1).Text) & ") "
    End If
    
    'Last Hire Date
    If Len(dlpEmpDateRange(0).Text) > 0 And Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR ((HREMP.ED_LTHIRE >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") AND "
        TempCri = TempCri & " (HREMP.ED_LTHIRE <= " & Date_SQL(dlpEmpDateRange(1).Text) & ")) "
    ElseIf Len(dlpEmpDateRange(0).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_LTHIRE >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") "
    ElseIf Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_LTHIRE <= " & Date_SQL(dlpEmpDateRange(1).Text) & ") "
    End If
    
    'Union Date
    If Len(dlpEmpDateRange(0).Text) > 0 And Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR ((HREMP.ED_UNION >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") AND "
        TempCri = TempCri & " (HREMP.ED_UNION <= " & Date_SQL(dlpEmpDateRange(1).Text) & ")) "
    ElseIf Len(dlpEmpDateRange(0).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_UNION >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") "
    ElseIf Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_UNION <= " & Date_SQL(dlpEmpDateRange(1).Text) & ") "
    End If
    
    'First Day Date
    If Len(dlpEmpDateRange(0).Text) > 0 And Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR ((HREMP.ED_FDAY >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") AND "
        TempCri = TempCri & " (HREMP.ED_FDAY <= " & Date_SQL(dlpEmpDateRange(1).Text) & ")) "
    ElseIf Len(dlpEmpDateRange(0).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_FDAY >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") "
    ElseIf Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_FDAY <= " & Date_SQL(dlpEmpDateRange(1).Text) & ") "
    End If
    
    'Last Day Date
    If Len(dlpEmpDateRange(0).Text) > 0 And Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR ((HREMP.ED_LDAY >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") AND "
        TempCri = TempCri & " (HREMP.ED_LDAY <= " & Date_SQL(dlpEmpDateRange(1).Text) & ")) "
    ElseIf Len(dlpEmpDateRange(0).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_LDAY >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") "
    ElseIf Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_LDAY <= " & Date_SQL(dlpEmpDateRange(1).Text) & ") "
    End If
    
    'OMERS Date
    If Len(dlpEmpDateRange(0).Text) > 0 And Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR ((HREMP.ED_OMERS >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") AND "
        TempCri = TempCri & " (HREMP.ED_OMERS <= " & Date_SQL(dlpEmpDateRange(1).Text) & ")) "
    ElseIf Len(dlpEmpDateRange(0).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_OMERS >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") "
    ElseIf Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_OMERS <= " & Date_SQL(dlpEmpDateRange(1).Text) & ") "
    End If
    
    'User Define Date
    If Len(dlpEmpDateRange(0).Text) > 0 And Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR ((HREMP.ED_USRDAT1 >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") AND "
        TempCri = TempCri & " (HREMP.ED_USRDAT1 <= " & Date_SQL(dlpEmpDateRange(1).Text) & "))) "
    ElseIf Len(dlpEmpDateRange(0).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_USRDAT1 >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") "
    ElseIf Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_USRDAT1 <= " & Date_SQL(dlpEmpDateRange(1).Text) & ") "
    End If
    
If Len(TempCri) = 0 Then 'Ticket #29378 Frank 10/27/2016 - TempCri = ") " sometime, that causes problems
    Exit Sub
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & TempCri
Else
    glbstrSelCri = TempCri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_Pension_Dates()
Dim TempCri As String, OneSet%, X%
Dim I As Integer

    If chkPenDates.Value = False Then Exit Sub
    
    'Eligibility Date
    If Len(dlpPenDateRange(0).Text) > 0 And Len(dlpPenDateRange(1).Text) > 0 Then
        TempCri = " (((HREMP.ED_ELIGIBLE >= " & Date_SQL(dlpPenDateRange(0).Text) & ") AND "
        TempCri = TempCri & " (HREMP.ED_ELIGIBLE <= " & Date_SQL(dlpPenDateRange(1).Text) & ")) "
    ElseIf Len(dlpEmpDateRange(0).Text) > 0 Then
        TempCri = " (HREMP.ED_ELIGIBLE >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") "
    ElseIf Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = " (HREMP.ED_ELIGIBLE <= " & Date_SQL(dlpEmpDateRange(1).Text) & ") "
    End If

    'Earlier Retirement Date
    If Len(dlpPenDateRange(0).Text) > 0 And Len(dlpPenDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR ((HREMP.ED_EARLYR >= " & Date_SQL(dlpPenDateRange(0).Text) & ") AND "
        TempCri = TempCri & " (HREMP.ED_EARLYR <= " & Date_SQL(dlpPenDateRange(1).Text) & ")) "
    ElseIf Len(dlpEmpDateRange(0).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_EARLYR >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") "
    ElseIf Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_EARLYR <= " & Date_SQL(dlpEmpDateRange(1).Text) & ") "
    End If
        
    'Normal Retirement Date
    If Len(dlpPenDateRange(0).Text) > 0 And Len(dlpPenDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR ((HREMP.ED_NORMALR >= " & Date_SQL(dlpPenDateRange(0).Text) & ") AND "
        TempCri = TempCri & " (HREMP.ED_NORMALR <= " & Date_SQL(dlpPenDateRange(1).Text) & ")) "
    ElseIf Len(dlpEmpDateRange(0).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_NORMALR >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") "
    ElseIf Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_NORMALR <= " & Date_SQL(dlpEmpDateRange(1).Text) & ") "
    End If
    
    'Latest Retirement Date
    If Len(dlpPenDateRange(0).Text) > 0 And Len(dlpPenDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR ((HREMP.ED_LATESTR >= " & Date_SQL(dlpPenDateRange(0).Text) & ") AND "
        TempCri = TempCri & " (HREMP.ED_LATESTR <= " & Date_SQL(dlpPenDateRange(1).Text) & ")) "
    ElseIf Len(dlpEmpDateRange(0).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_LATESTR >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") "
    ElseIf Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = TempCri & "OR (HREMP.ED_LATESTR <= " & Date_SQL(dlpEmpDateRange(1).Text) & ") "
    End If
    
    'Pension Date 1 - 6
    If Len(dlpPenDateRange(0).Text) > 0 And Len(dlpPenDateRange(1).Text) > 0 Then
        For I = 1 To 6
            If Len(dlpPenDateRange(0).Text) > 0 And Len(dlpPenDateRange(1).Text) > 0 Then
                TempCri = TempCri & "OR ((HREMP_OTHER.ER_PENSIONDATE" & I & " >= " & Date_SQL(dlpPenDateRange(0).Text) & ") AND "
                TempCri = TempCri & " (HREMP_OTHER.ER_PENSIONDATE" & I & " <= " & Date_SQL(dlpPenDateRange(1).Text) & ")) "
            ElseIf Len(dlpEmpDateRange(0).Text) > 0 Then
                TempCri = TempCri & "OR (HREMP_OTHER.ER_PENSIONDATE" & I & " >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") "
            ElseIf Len(dlpEmpDateRange(1).Text) > 0 Then
                TempCri = TempCri & "OR (HREMP_OTHER.ER_PENSIONDATE" & I & " <= " & Date_SQL(dlpEmpDateRange(1).Text) & ") "
            End If
            If I = 6 Then
                TempCri = TempCri & ") "
            End If
        Next
    End If
    
If Len(TempCri) = 0 Then 'Ticket #29378 Frank 10/27/2016
    Exit Sub
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & TempCri
Else
    glbstrSelCri = TempCri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_Other_Dates()
Dim TempCri As String, OneSet%, X%
Dim I As Integer

    If chkOthDates.Value = False Then Exit Sub
    
    'Other Date
    If Len(dlpOthDateRange(0).Text) > 0 And Len(dlpOthDateRange(1).Text) > 0 Then
        TempCri = " (((HREMP_OTHER.ER_OTHERDATE1 >= " & Date_SQL(dlpOthDateRange(0).Text) & ") AND "
        TempCri = TempCri & " (HREMP_OTHER.ER_OTHERDATE1 <= " & Date_SQL(dlpOthDateRange(1).Text) & ")) "
    ElseIf Len(dlpEmpDateRange(0).Text) > 0 Then
        TempCri = " (HREMP_OTHER.ER_OTHERDATE1 >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") "
    ElseIf Len(dlpEmpDateRange(1).Text) > 0 Then
        TempCri = " (HREMP_OTHER.ER_OTHERDATE1 <= " & Date_SQL(dlpEmpDateRange(1).Text) & ") "
    End If
    
    'Other Date 2 - 10
    If Len(dlpOthDateRange(0).Text) > 0 And Len(dlpOthDateRange(1).Text) > 0 Then
        For I = 2 To 10
            If Len(dlpOthDateRange(0).Text) > 0 And Len(dlpOthDateRange(1).Text) > 0 Then
                TempCri = TempCri & "OR ((HREMP_OTHER.ER_OTHERDATE" & I & " >= " & Date_SQL(dlpOthDateRange(0).Text) & ") AND "
                TempCri = TempCri & " (HREMP_OTHER.ER_OTHERDATE" & I & " <= " & Date_SQL(dlpOthDateRange(1).Text) & "))"
            ElseIf Len(dlpEmpDateRange(0).Text) > 0 Then
                TempCri = TempCri & "OR (HREMP_OTHER.ER_OTHERDATE" & I & " >= " & Date_SQL(dlpEmpDateRange(0).Text) & ") "
            ElseIf Len(dlpEmpDateRange(1).Text) > 0 Then
                TempCri = TempCri & "OR (HREMP_OTHER.ER_OTHERDATE" & I & " <= " & Date_SQL(dlpEmpDateRange(1).Text) & ") "
            End If
            If I = 10 Then
                TempCri = TempCri & ") "
            End If
        Next
    End If
    
If Len(TempCri) = 0 Then 'Ticket #29378 Frank 10/27/2016 - TempCri = ") " sometime, that causes problems
    Exit Sub
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & TempCri
Else
    glbstrSelCri = TempCri
End If
glbiOneWhere = True

End Sub

Private Function Cri_SetAll()
Dim X%, strRName$
Dim SQLQ As String, strCriteria As String
Dim rsIn As New ADODB.Recordset, rsOUT As New ADODB.Recordset

Cri_SetAll = False

'''On Error GoTo modSetCriteria_Err

Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For X% = 0 To 5
    Call Cri_Code(X%)
Next X%
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_ProvResidence
Call Cri_ProvEmployment
Call Cri_Employment_Dates
Call Cri_Pension_Dates
Call Cri_Other_Dates

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

glbstrSelCri = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")

Cri_SetAll = XLWriter_Employee_Dates

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

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, X%

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
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
'imbeded in report

Cri_Sorts = 0
'first set primary grouping
Y% = 0
grpField$ = getEGroup(comGroup(0).Text)

If comGroup(0) = "(none)" Then
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
        Case 2: grpField$ = "{HREMP.ED_DOH}"
    End Select
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$

    Exit Function
End If
    
Y% = X% + 1
dscGroup$ = comGroup(X%).Text
dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(X%) = dscGroup$

grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(X%) = grpCond$

strSFormat$ = "GH1;T;T;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1
strSFormat$ = "GF1;T;X;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1

GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@EFullName}"
    Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
    Case 2: grpField$ = "{HREMP.ED_DOH}"
End Select
grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$
      

Cri_Sorts = z% ' next section number to format

End Function

Private Function CriCheck()
Dim X%

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

For X% = 0 To 5
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

If Not elpEEID.ListChecker Then
    Exit Function
End If

If clpProv.Caption = "Unassigned" Then
    MsgBox "Invalid Prov. of Residence"
    clpProv.SetFocus
    Exit Function
End If

If clpProvEmp.Caption = "Unassigned" Then
    MsgBox "Invalid Prov. of Employment"
    clpProvEmp.SetFocus
    Exit Function
End If

If chkEmpDates Then
    If Len(Trim(dlpEmpDateRange(0).Text)) = 0 And Len(Trim(dlpEmpDateRange(1).Text)) = 0 Then
        MsgBox "'Employement Dates' date range cannot be blank"
        chkEmpDates.SetFocus
        Exit Function
    Else
        If Len(Trim(dlpEmpDateRange(0).Text)) > 0 Then
            If Not IsDate(dlpEmpDateRange(0).Text) Then
                MsgBox "Invalid From 'Employement Dates'"
                dlpEmpDateRange(0).SetFocus
                Exit Function
            End If
        End If
        If Len(Trim(dlpEmpDateRange(1).Text)) > 0 Then
            If Not IsDate(dlpEmpDateRange(1).Text) Then
                MsgBox "Invalid To 'Employement Dates'"
                dlpEmpDateRange(1).SetFocus
                Exit Function
            End If
        End If
        
        If IsDate(dlpEmpDateRange(0).Text) And IsDate(dlpEmpDateRange(1).Text) Then
            If CVDate(dlpEmpDateRange(0).Text) > CVDate(dlpEmpDateRange(1).Text) Then
                MsgBox "Under 'Employment Dates' From Date cannot be greater thant To Date"
                dlpEmpDateRange(0).SetFocus
                Exit Function
            End If
        End If
    End If
Else
    If Len(Trim(dlpEmpDateRange(0).Text)) > 0 Or Len(Trim(dlpEmpDateRange(1).Text)) > 0 Then
        MsgBox "'Employement Dates' not selected 'Employement Dates' date range must be blank"
        If Len(Trim(dlpEmpDateRange(0).Text)) > 0 Then
            dlpEmpDateRange(0).SetFocus
        Else
            dlpEmpDateRange(1).SetFocus
        End If
        Exit Function
    End If
End If

If chkPenDates Then
    If Len(Trim(dlpPenDateRange(0).Text)) = 0 And Len(Trim(dlpPenDateRange(1).Text)) = 0 Then
        MsgBox "'Pension Dates' date range cannot be blank"
        chkPenDates.SetFocus
        Exit Function
    Else
        If Len(Trim(dlpPenDateRange(0).Text)) > 0 Then
            If Not IsDate(dlpPenDateRange(0).Text) Then
                MsgBox "Invalid From 'Pension Dates'"
                dlpPenDateRange(0).SetFocus
                Exit Function
            End If
        End If
        If Len(Trim(dlpPenDateRange(1).Text)) > 0 Then
            If Not IsDate(dlpPenDateRange(1).Text) Then
                MsgBox "Invalid To 'Pension Dates'"
                dlpPenDateRange(1).SetFocus
                Exit Function
            End If
        End If
    
        If IsDate(dlpPenDateRange(0).Text) And IsDate(dlpPenDateRange(1).Text) Then
            If CVDate(dlpPenDateRange(0).Text) > CVDate(dlpPenDateRange(1).Text) Then
                MsgBox "Under 'Pension Dates' From Date cannot be greater thant To Date"
                dlpPenDateRange(0).SetFocus
                Exit Function
            End If
        End If
    End If
Else
    If Len(Trim(dlpPenDateRange(0).Text)) > 0 Or Len(Trim(dlpPenDateRange(1).Text)) > 0 Then
        MsgBox "'Pension Dates' not selected 'Pension Dates' date range must be blank"
        If Len(Trim(dlpPenDateRange(0).Text)) > 0 Then
            dlpPenDateRange(0).SetFocus
        Else
            dlpPenDateRange(1).SetFocus
        End If
        Exit Function
    End If
End If

If chkOthDates Then
    If Len(Trim(dlpOthDateRange(0).Text)) = 0 And Len(Trim(dlpOthDateRange(1).Text)) = 0 Then
        MsgBox "'Other Dates' date range cannot be blank"
        chkOthDates.SetFocus
        Exit Function
    Else
        If Len(Trim(dlpOthDateRange(0).Text)) > 0 Then
            If Not IsDate(dlpOthDateRange(0).Text) Then
                MsgBox "Invalid From 'Other Dates'"
                dlpOthDateRange(0).SetFocus
                Exit Function
            End If
        End If
        If Len(Trim(dlpOthDateRange(1).Text)) > 0 Then
            If Not IsDate(dlpOthDateRange(1).Text) Then
                MsgBox "Invalid To 'Other Dates'"
                dlpOthDateRange(1).SetFocus
                Exit Function
            End If
        End If
        
        If IsDate(dlpOthDateRange(0).Text) And IsDate(dlpOthDateRange(1).Text) Then
            If CVDate(dlpOthDateRange(0).Text) > CVDate(dlpOthDateRange(1).Text) Then
                MsgBox "Under 'Other Dates' From Date cannot be greater thant To Date"
                dlpOthDateRange(0).SetFocus
                Exit Function
            End If
        End If
    End If
Else
    If Len(Trim(dlpOthDateRange(0).Text)) > 0 Or Len(Trim(dlpOthDateRange(1).Text)) > 0 Then
        MsgBox "'Other Dates' not selected 'Other Dates' date range must be blank"
        If Len(Trim(dlpOthDateRange(0).Text)) > 0 Then
            dlpOthDateRange(0).SetFocus
        Else
            dlpOthDateRange(1).SetFocus
        End If
        Exit Function
    End If
End If

If Not chkEmpDates And Not chkPenDates And Not chkOthDates And Not chkShowAllEmpDates And Not chkShowAllPenDates And Not chkShowAllOthDates Then
    MsgBox "At least one date option must be selected"
    chkEmpDates.SetFocus
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

glbOnTop = "FRMREMPDATES"

Screen.MousePointer = HOURGLASS

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call comGrpLoad
Call setRptCaption(Me)

If FormHomeAddress% = True Then
   frmREmpDates.Caption = "Home Address/Telephone Report"
   chkIncludeTerm.Visible = True
   chkOnlyTerm.Visible = True
End If

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

Private Sub txtShift_GotFocus()
Call SetPanHelp(ActiveControl)
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

Private Function XLWriter_Employee_Dates()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsGroup As New ADODB.Recordset
    Dim rsHREmpOther As New ADODB.Recordset
    Dim SQLQ, sSQLQ As String
    Dim exApp As Object
    Dim exBook As Object
    Dim exSheet As Object
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim xCol As Long
    Dim I, totNum
    Dim xExcelRptPath  As String
    Dim strTime As String
    Dim xRptHeadRange As String

    On Error GoTo Err_XLWriter_Employee_Dates
    
    Screen.MousePointer = HOURGLASS
    
    glbstrSelCri = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
            
    'Ticket #22034 - Get Excel reports path
    If gsTRAININGMATRIX Then
        xExcelRptPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xExcelRptPath) = 0 Then
        xExcelRptPath = glbIHRREPORTS
    End If
            
    'File to export to
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "EmployeeDatesTmp.xls"
    xlsFileMat = xExcelRptPath & IIf(Right(xExcelRptPath, 1) = "\", "", "\") & "EmployeeDates" & Trim(glbUserID) & ".xls"

    If Dir(xlsFileTmp) = "" Then
        Screen.MousePointer = DEFAULT
        MsgBox "There is no " & xlsFileTmp
        Exit Function
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0
    
    FileCopy xlsFileTmp, xlsFileMat

    Dim appVerInt As Double
    
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)

    'Ticket #22166
    appVerInt = Split(exApp.Version, ".")(0)
    If appVerInt - Excel2007 >= 0 Then
        'exApp.ActiveWorkbook.SaveAs (sXLS), 56
        exApp.DisplayAlerts = False
        exBook.SaveAs (xlsFileMat), 56
        exApp.DisplayAlerts = True
    Else
        'exApp.ActiveWorkbook.SaveAs (sXLS), 43
        exApp.DisplayAlerts = False
        exBook.SaveAs (xlsFileMat), 43
        exApp.DisplayAlerts = True
    End If

    'Create Grouping recordset
    SQLQ = getGroupField
    rsGroup.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsGroup.EOF Then
        strTime = Time$
        Do While Not rsGroup.EOF
            If comGroup(0).Text = "(none)" Then
                Set exSheet = exBook.Sheets(1)
                exSheet.name = "All"
            Else
                exBook.Worksheets(1).Copy After:=exBook.Worksheets(exBook.Worksheets.count)
                Set exSheet = exBook.Worksheets(exBook.Worksheets.count)
                
                If IsNull(rsGroup("TB_DESC")) Then
                    exSheet.name = "None"
                Else
                    exSheet.name = Left(rsGroup("TB_DESC"), 30)
                End If
            End If
            
            'Get the Employees for the Report
            SQLQ = "DELETE FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
            gdbAdoIhr001W.Execute SQLQ
            If Not glbSQL And Not glbOracle Then Call Pause(3)
            
            SQLQ = "INSERT INTO HR_EMPLIST_WRK (TT_EMPNBR, TT_WRKEMP) " & in_SQL(glbIHRDBW)
            SQLQ = SQLQ & " SELECT HREMP.ED_EMPNBR, '" & glbUserID & "' as TT_WRKEMP FROM HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR INNER JOIN HREMP_OTHER ON HREMP.ED_EMPNBR = HREMP_OTHER.ER_EMPNBR "
            SQLQ = SQLQ & " WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) And " & glbstrSelCri
            If IsNull(rsGroup("grpField")) Then
                If glbSQL Or glbOracle Then
                    SQLQ = SQLQ & " and " & strGroup & " IS NULL "
                Else
                    SQLQ = SQLQ & " and isnull(" & strGroup & ") = true "
                End If
            Else
                If strGroup = "HR_JOB_HISTORY.JH_REPTAU" Then 'Supervisor is a number
                    SQLQ = SQLQ & " and " & strGroup & "=" & rsGroup("grpField") & " "
                Else
                    SQLQ = SQLQ & " and " & strGroup & "='" & Replace(rsGroup("grpField"), "'", "''") & "' "
                End If
            End If
            gdbAdoIhr001.Execute SQLQ

            If Not glbSQL And Not glbOracle Then Call Pause(3)
            
            'Build recordset of employees to print in the report based on selection criteria a grouping
            SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR IN "
            SQLQ = SQLQ & " (SELECT HR_EMPLIST_WRK.TT_EMPNBR FROM HR_EMPLIST_WRK WHERE HR_EMPLIST_WRK.TT_WRKEMP='" & glbUserID & "') "
            SQLQ = SQLQ & " ORDER BY "
            If comGroup(1).ListIndex = 0 Then
                SQLQ = SQLQ & "HREMP.ED_SURNAME,HREMP.ED_FNAME "
            ElseIf comGroup(1).ListIndex = 1 Then
                SQLQ = SQLQ & "HREMP.ED_EMPNBR"
            End If
            rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            
            'Sheet Labels
            exSheet.Cells(1, 1) = glbCompName
            exSheet.Cells(3, 1) = "Date: " & Format(Now, "mmm dd, yyyy")
            exSheet.Cells(4, 1) = "Time: " & strTime
            exSheet.Cells(6, 1) = "Group by " & comGroup(0).Text
            
            xCol = 4
            xRptHeadRange = ""
            'Show All Employment Dates
            If chkEmpDates Or chkShowAllEmpDates Then
                exSheet.Cells(7, xCol + 3) = "Employment Dates"  'xCol+3 7
                exSheet.Range(exSheet.Cells(7, xCol + 3), exSheet.Cells(7, xCol + 4)).Merge
                exSheet.Range(exSheet.Cells(7, xCol + 3), exSheet.Cells(7, xCol + 4)).HorizontalAlignment = xlCenter
                
                exSheet.Cells(8, xCol) = lStr("Original Hire")      'xCol   4
                exSheet.Cells(8, xCol + 1) = lStr("Seniority")      'xCol+1 5
                exSheet.Cells(8, xCol + 2) = lStr("Last Hire")      'xCol+2
                exSheet.Cells(8, xCol + 3) = lStr("Union Date")     'xCol+3
                exSheet.Cells(8, xCol + 4) = lStr("First Day")      'xCol+4
                exSheet.Cells(8, xCol + 5) = lStr("Last Day")       'xCol+5
                exSheet.Cells(8, xCol + 6) = lStr("OMERS Date")     'xCol+6
                exSheet.Cells(8, xCol + 7) = lStr("User Defined")   'xCol+7
                
                xCol = xCol + 8
                xRptHeadRange = "K1"
            End If
            
            'Show All Pension Dates
            If chkPenDates Or chkShowAllPenDates Then
                exSheet.Cells(7, xCol + 4) = "Pension Dates"    'xCol+4 16
                exSheet.Range(exSheet.Cells(7, xCol + 4), exSheet.Cells(7, xCol + 5)).Merge
                exSheet.Range(exSheet.Cells(7, xCol + 4), exSheet.Cells(7, xCol + 5)).HorizontalAlignment = xlCenter
                
                exSheet.Cells(8, xCol) = lStr("Eligibility")                'xCol   12
                exSheet.Cells(8, xCol + 1) = lStr("Earliest Retirement")    'xCol+1 13
                exSheet.Cells(8, xCol + 2) = lStr("Normal Retirement")      'xCol+2
                exSheet.Cells(8, xCol + 3) = lStr("Latest Retirement")      'xCol+3
                exSheet.Cells(8, xCol + 4) = lStr("Pension Date 1")         'xCol+4
                exSheet.Cells(8, xCol + 5) = lStr("Pension Date 2")         'xCol+5
                exSheet.Cells(8, xCol + 6) = lStr("Pension Date 3")         'xCol+6
                exSheet.Cells(8, xCol + 7) = lStr("Pension Date 4")         'xCol+7
                exSheet.Cells(8, xCol + 8) = lStr("Pension Date 5")         'xCol+8
                exSheet.Cells(8, xCol + 9) = lStr("Pension Date 6")         'xCol+9
                
                xCol = xCol + 10
                
                If Len(xRptHeadRange) > 0 Then
                    xRptHeadRange = "U1"
                Else
                    xRptHeadRange = "M1"
                End If
            End If
            
            'Show All Other Dates
            If chkOthDates Or chkShowAllOthDates Then
                exSheet.Cells(7, xCol + 4) = "Other Dates"          'xCol+4 26
                exSheet.Range(exSheet.Cells(7, xCol + 4), exSheet.Cells(7, xCol + 5)).Merge
                exSheet.Range(exSheet.Cells(7, xCol + 4), exSheet.Cells(7, xCol + 5)).HorizontalAlignment = xlCenter
                
                exSheet.Cells(8, xCol) = lStr("Other Date 1")       'xCol   22
                exSheet.Cells(8, xCol + 1) = lStr("Other Date 2")   'xCol+1 23
                exSheet.Cells(8, xCol + 2) = lStr("Other Date 3")   'xCol+2
                exSheet.Cells(8, xCol + 3) = lStr("Other Date 4")   'xCol+3
                exSheet.Cells(8, xCol + 4) = lStr("Other Date 5")   'xCol+4
                exSheet.Cells(8, xCol + 5) = lStr("Other Date 6")   'xCol+5
                exSheet.Cells(8, xCol + 6) = lStr("Other Date 7")   'xCol+6
                exSheet.Cells(8, xCol + 7) = lStr("Other Date 8")   'xCol+7
                exSheet.Cells(8, xCol + 8) = lStr("Other Date 9")   'xCol+8
                exSheet.Cells(8, xCol + 9) = lStr("Other Date 10")  'xCol+9
            
                If Len(xRptHeadRange) > 0 Then
                    If xCol = 22 Then
                        xRptHeadRange = "AE1"
                    ElseIf xCol = 12 Then
                        xRptHeadRange = "U1"
                    Else
                        xRptHeadRange = "W1"
                    End If
                Else
                    xRptHeadRange = "M1"
                End If
            End If
            
            If Not rsHREmp.EOF Then
                xRow = 9
                Do While Not rsHREmp.EOF
                    exSheet.Cells(xRow, 1) = rsHREmp("ED_EMPNBR")
                    exSheet.Cells(xRow, 2) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
                    exSheet.Cells(xRow, 3) = GetJobData(GetJHData(rsHREmp("ED_EMPNBR"), "JH_JOB", ""), "JB_DESCR")
        
                    xCol = 4
                    'Show All Employment Dates
                    If chkEmpDates Or chkShowAllEmpDates Then
                        exSheet.Cells(xRow, xCol) = rsHREmp("ED_DOH")           '4
                        exSheet.Cells(xRow, xCol + 1) = rsHREmp("ED_SENDTE")    '5
                        exSheet.Cells(xRow, xCol + 2) = rsHREmp("ED_LTHIRE")
                        exSheet.Cells(xRow, xCol + 3) = rsHREmp("ED_UNION")
                        exSheet.Cells(xRow, xCol + 4) = rsHREmp("ED_FDAY")
                        exSheet.Cells(xRow, xCol + 5) = rsHREmp("ED_LDAY")
                        exSheet.Cells(xRow, xCol + 6) = rsHREmp("ED_OMERS")
                        exSheet.Cells(xRow, xCol + 7) = rsHREmp("ED_USRDAT1")
                        
                        xCol = xCol + 8
                    End If
                    
                    'Show All Pension Dates
                    If chkPenDates Or chkShowAllPenDates Then
                        exSheet.Cells(xRow, xCol) = rsHREmp("ED_ELIGIBLE")      '12
                        exSheet.Cells(xRow, xCol + 1) = rsHREmp("ED_EARLYR")    '13
                        exSheet.Cells(xRow, xCol + 2) = rsHREmp("ED_NORMALR")
                        exSheet.Cells(xRow, xCol + 3) = rsHREmp("ED_LATESTR")
                    End If
                    
                    'More Pension Dates & Other Dates from HREMP_OTHER table
                    If chkPenDates Or chkShowAllPenDates Or chkOthDates Or chkShowAllOthDates Then
                        'Get date from HREMP_OTHER table for Pension and Other Dates
                        rsHREmpOther.Open "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & rsHREmp("ED_EMPNBR"), gdbAdoIhr001, adOpenStatic, adLockOptimistic
                        If Not rsHREmpOther.EOF Then
                            If chkPenDates Or chkShowAllPenDates Then
                                exSheet.Cells(xRow, xCol + 4) = rsHREmpOther("ER_PENSIONDATE1")   '16
                                exSheet.Cells(xRow, xCol + 5) = rsHREmpOther("ER_PENSIONDATE2")   '17
                                exSheet.Cells(xRow, xCol + 6) = rsHREmpOther("ER_PENSIONDATE3")
                                exSheet.Cells(xRow, xCol + 7) = rsHREmpOther("ER_PENSIONDATE4")
                                exSheet.Cells(xRow, xCol + 8) = rsHREmpOther("ER_PENSIONDATE5")
                                exSheet.Cells(xRow, xCol + 9) = rsHREmpOther("ER_PENSIONDATE6")
                                
                                xCol = xCol + 10
                            End If
                            
                            'Show All Other Dates
                            If chkOthDates Or chkShowAllOthDates Then
                                exSheet.Cells(xRow, xCol) = rsHREmpOther("ER_OTHERDATE1")       '22
                                exSheet.Cells(xRow, xCol + 1) = rsHREmpOther("ER_OTHERDATE2")   '23
                                exSheet.Cells(xRow, xCol + 2) = rsHREmpOther("ER_OTHERDATE3")
                                exSheet.Cells(xRow, xCol + 3) = rsHREmpOther("ER_OTHERDATE4")
                                exSheet.Cells(xRow, xCol + 4) = rsHREmpOther("ER_OTHERDATE5")
                                exSheet.Cells(xRow, xCol + 5) = rsHREmpOther("ER_OTHERDATE6")
                                exSheet.Cells(xRow, xCol + 6) = rsHREmpOther("ER_OTHERDATE7")
                                exSheet.Cells(xRow, xCol + 7) = rsHREmpOther("ER_OTHERDATE8")
                                exSheet.Cells(xRow, xCol + 8) = rsHREmpOther("ER_OTHERDATE9")
                                exSheet.Cells(xRow, xCol + 9) = rsHREmpOther("ER_OTHERDATE10")
                            End If
                        End If
                        rsHREmpOther.Close
                        Set rsHREmpOther = Nothing
                    End If
                    
                    xRow = xRow + 1
                
                    rsHREmp.MoveNext
                Loop
                
                If Len(xRptHeadRange) > 0 Then
                    exSheet.Range("A1:" & xRptHeadRange).Merge
                    exSheet.Range("A1:" & xRptHeadRange).HorizontalAlignment = xlCenter
                    exSheet.Range("A2:" & Replace(xRptHeadRange, "1", "2")).Merge
                    exSheet.Range("A2:" & Replace(xRptHeadRange, "1", "2")).HorizontalAlignment = xlCenter
                Else
                    exSheet.Range("A1:C1").Merge
                    exSheet.Range("A1:C1").HorizontalAlignment = xlCenter
                    exSheet.Range("A2:C2").Merge
                    exSheet.Range("A2:C2").HorizontalAlignment = xlCenter
                End If
                
                exSheet.PageSetup.Orientation = xlLandscape
                exSheet.PageSetup.Zoom = False
                exSheet.Range("A1").Select
                exBook.Save
                Set exSheet = Nothing
                
            End If
            rsHREmp.Close
            Set rsHREmp = Nothing
            
            rsGroup.MoveNext
        Loop
                
        Set exSheet = exBook.Worksheets(1)
        exApp.DisplayAlerts = False
        exSheet.Delete
        exApp.DisplayAlerts = True
        exBook.Worksheets(1).Activate
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        'exApp.Visible = True
        exApp.Quit
        Set exApp = Nothing

        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If

    Else
        MsgBox "There are no Records matching this criteria", vbInformation + vbOKOnly, "No Records"
        exBook.Close False
        exApp.Quit
        Set exApp = Nothing
    End If
    rsGroup.Close
    Set rsGroup = Nothing

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT

    Screen.MousePointer = vbDefault

    Exit Function
    
Err_XLWriter_Employee_Dates:

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT
    
    If Err = 1004 Then
        Resume Next
    End If
    
    If Err = 75 Then
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Function
    End If
    If Not exBook Is Nothing Then
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
    End If
    If Err = 70 Then
        Set exApp = Nothing
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Function
    End If
    If Err = 76 Then
        MsgBox Err.Description & " to save the Employee Dates Report." & vbCrLf & "Please check Company Preference under Setup Menu."
        Exit Function
    End If
    If Not exApp Is Nothing Then
        If exApp.Visible = False Then
            exApp.Quit
        End If
        Set exApp = Nothing
    End If
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "XLWriter_Employee_Dates", "", "Select")
End Function

Private Function GroupingInExcel(xGroup) As String
    Dim xExcelGrp As String
    
    Select Case xGroup
        Case lStr("Division"):      GroupingInExcel = "HR_DIVISION.Division_Name"
        Case lStr("Department"):    GroupingInExcel = "HRDEPT.DF_NAME"
        Case lStr("Location"):      GroupingInExcel = "HRTABL.TB_DESC"
        Case lStr("Section"):       GroupingInExcel = "tblSec.TB_DESC"
        Case "Employment Type":     GroupingInExcel = "ED_EMPTYPE"
        Case ("Home Line"):         GroupingInExcel = "LN_HOMES.TB_DESC"
        Case "Shift":               GroupingInExcel = "ED_SHIFT"
        Case lStr("Region"):        GroupingInExcel = "tblRegion.TB_DESC"
        Case "(none)":              GroupingInExcel = "(none)"
        Case "Employee Name":       GroupingInExcel = "ED_SURNAME, ED_FNAME"
        Case "Employee Number":     GroupingInExcel = "ED_EMPNBR"
        Case "Date of Hire":        GroupingInExcel = "ED_DOH"
        Case Else:                  GroupingInExcel = "ED_EMPNBR"
    End Select

'if HREMP.ED_EMPTYPE="0" then "0 - Not Applicable" else
'if HREMP.ED_EMPTYPE="1" then "1 - Full Time Salary" else
'if HREMP.ED_EMPTYPE="2" then "2 - Part Time Salary" else
'if HREMP.ED_EMPTYPE="3" then "3 - Full Time Hourly" else
'if HREMP.ED_EMPTYPE="4" then "4 - Part Time Hourly" else
'if HREMP.ED_EMPTYPE="5" then "5 - Casual/Other" else
'if HREMP.ED_EMPTYPE="6" then "6 - Contract Salary"  else
'if HREMP.ED_EMPTYPE="7" then "7 - Contract Hourly"  else
'if HREMP.ED_EMPTYPE="8" then "8 - Salary Pensioners" else
'if HREMP.ED_EMPTYPE="9" then "9 - Salary Elected officials"

End Function

Private Function getGroupField() As String

    Dim strTable As String
    Dim SQLQ As String
    
    Select Case comGroup(0).Text
        Case lStr("Division")
            strTable = ""
            strGroup = "HREMP.ED_DIV"
            SQLQ = "SELECT DISTINCT HREMP.ED_DIV as grpField, HR_DIVISION.Division_Name as TB_DESC FROM ((HREMP LEFT OUTER JOIN HR_DIVISION ON HREMP.ED_DIV = HR_DIVISION.DIV) "
            SQLQ = SQLQ & "INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HREMP_OTHER ON HREMP.ED_EMPNBR = HREMP_OTHER.ER_EMPNBR WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
        Case lStr("Department")
            strTable = ""
            strGroup = "HREMP.ED_DEPTNO"
            SQLQ = "SELECT DISTINCT HREMP.ED_DEPTNO as grpField, HRDEPT.DF_NAME as TB_DESC FROM (HREMP INNER JOIN HRDEPT ON HREMP.ED_DEPTNO = HRDEPT.DF_NBR) INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR INNER JOIN HREMP_OTHER ON HREMP.ED_EMPNBR = HREMP_OTHER.ER_EMPNBR "
            SQLQ = SQLQ & "WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
        Case lStr("Location")
            strGroup = "HREMP.ED_LOC"
            strTable = "HREMP.ED_LOC_TABL"
        Case lStr("Union")
            strGroup = "HREMP.ED_ORG"
            strTable = "HREMP.ED_ORG_TABL"
        Case lStr("Administered By")
            strGroup = "HREMP.ED_ADMINBY"
            strTable = "HREMP.ED_ADMINBY_TABL"
        Case "Employee Name"
            strTable = ""
            strGroup = "HREMP.ED_SURNAME"
            If glbOracle Then
                SQLQ = "SELECT DISTINCT HREMP.ED_SURNAME AS grpfield, CONCAT(CONCAT(HREMP.ED_SURNAME, ', '), HREMP.ED_FNAME) AS TB_DESC FROM HREMP "
            Else
                SQLQ = "SELECT DISTINCT HREMP.ED_SURNAME AS grpfield, ED_SURNAME + ', ' + ED_FNAME AS TB_DESC FROM HREMP "
            End If
            SQLQ = SQLQ & "INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR INNER JOIN HREMP_OTHER ON HREMP.ED_EMPNBR = HREMP_OTHER.ER_EMPNBR WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
        Case lStr("Section")
            strGroup = "HREMP.ED_SECTION"
            strTable = "HREMP.ED_SECTION_TABL"
        Case "Employment Type"
            strTable = ""
            strGroup = "HREMP.ED_EMPTYPE"
            SQLQ = "SELECT DISTINCT HREMP.ED_EMPTYPE as grpfield, ED_EMPTYPE as TB_DESC FROM HREMP "
            SQLQ = SQLQ & "INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR INNER JOIN HREMP_OTHER ON HREMP.ED_EMPNBR = HREMP_OTHER.ER_EMPNBR WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
        Case ("Home Line")
            strGroup = "HREMP.ED_HOMELINE"
            strTable = "HREMP.ED_HOMELINE_TABL"
        Case "Shift"
            strTable = ""
            strGroup = "HR_JOB_HISTORY.JH_SHIFT"
            SQLQ = "SELECT DISTINCT HR_JOB_HISTORY.JH_SHIFT as grpField, HR_JOB_HISTORY.JH_SHIFT as TB_DESC FROM HREMP LEFT OUTER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR INNER JOIN HREMP_OTHER ON HREMP.ED_EMPNBR = HREMP_OTHER.ER_EMPNBR WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
        Case lStr("Rept. Authority 1")
            strTable = ""
            strGroup = "HR_JOB_HISTORY.JH_REPTAU"
            If glbOracle Then
                SQLQ = "SELECT DISTINCT HR_JOB_HISTORY.JH_REPTAU AS grpField, CONCAT(CONCAT(tblSuper.ED_SURNAME, ', '), tblSuper.ED_FNAME) AS TB_DESC "
                SQLQ = SQLQ & "FROM (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HREMP tblSuper ON HR_JOB_HISTORY.JH_REPTAU = tblSuper.ED_EMPNBR INNER JOIN HREMP_OTHER ON HREMP.ED_EMPNBR = HREMP_OTHER.ER_EMPNBR "
                SQLQ = SQLQ & "WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
            Else
                SQLQ = "SELECT DISTINCT HR_JOB_HISTORY.JH_REPTAU AS grpField, tblSuper.ED_SURNAME + ', ' + tblSuper.ED_FNAME AS TB_DESC "
                SQLQ = SQLQ & "FROM (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HREMP AS tblSuper ON HR_JOB_HISTORY.JH_REPTAU = tblSuper.ED_EMPNBR INNER JOIN HREMP_OTHER ON HREMP.ED_EMPNBR = HREMP_OTHER.ER_EMPNBR "
                SQLQ = SQLQ & "WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
            End If
        Case lStr("Region")
            strGroup = "HREMP.ED_REGION"
            strTable = "HREMP.ED_REGION_TABL"
        Case "(none)"
            strGroup = "HREMP.ED_COMPNO"
            strTable = ""
            SQLQ = "SELECT DISTINCT " & strGroup & " as grpField, 'None' as TB_DESC FROM HREMP "
            SQLQ = SQLQ & "INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR INNER JOIN HREMP_OTHER ON HREMP.ED_EMPNBR = HREMP_OTHER.ER_EMPNBR WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
    End Select
    
    If strTable <> "" Then
        SQLQ = "SELECT DISTINCT " & strGroup & " as grpField, HRTABL.TB_DESC as TB_DESC FROM ((HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HREMP_OTHER ON HREMP.ED_EMPNBR = HREMP_OTHER.ER_EMPNBR "
        SQLQ = SQLQ & "LEFT OUTER JOIN HRTABL ON " & strTable & " = HRTABL.TB_NAME AND " & strGroup & " = HRTABL.TB_KEY) "
        SQLQ = SQLQ & "WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
    End If
    
    SQLQ = SQLQ & glbstrSelCri
    SQLQ = SQLQ & " ORDER BY " & strGroup
    
    If Not (glbSQL Or glbOracle) Then
        SQLQ = Replace(SQLQ, "INNER", "LEFT")
    End If
    
    getGroupField = SQLQ
End Function

