VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmRDOB 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   8895
   ClientLeft      =   180
   ClientTop       =   825
   ClientWidth     =   10455
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
   ScaleHeight     =   8895
   ScaleWidth      =   10455
   WindowState     =   2  'Maximized
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
      TabIndex        =   27
      Tag             =   "Check to Exclude Employees with RET Employment Status"
      Top             =   6900
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3315
   End
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
      TabIndex        =   26
      Tag             =   "Check to Exclude Employees with CONP Employment Status"
      Top             =   6600
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.TextBox txtAgeTo 
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
      Height          =   285
      Left            =   3195
      MaxLength       =   2
      TabIndex        =   22
      Tag             =   "10-Age Range To"
      Top             =   5880
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox txtAgeFrom 
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
      Height          =   285
      Left            =   1995
      MaxLength       =   2
      TabIndex        =   21
      Tag             =   "10-Age Range From"
      Top             =   5880
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1995
      MaxLength       =   4
      TabIndex        =   11
      Tag             =   "00-Employee Position Shift"
      Top             =   3590
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   10
      Top             =   3270
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   9
      Top             =   2940
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   8
      Top             =   2610
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin VB.TextBox txtYear 
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
      Height          =   285
      Left            =   1995
      MaxLength       =   4
      TabIndex        =   15
      Tag             =   "10-Birthdays due in this Year"
      Top             =   4200
      Width           =   540
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
      Index           =   2
      Left            =   1875
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Tag             =   "Final Sorting of Records"
      Top             =   8595
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
      ItemData        =   "FZDOB.frx":0000
      Left            =   1875
      List            =   "FZDOB.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Tag             =   "Second level of grouping records"
      Top             =   8280
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
      ItemData        =   "FZDOB.frx":0004
      Left            =   1875
      List            =   "FZDOB.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Tag             =   "First Level of grouping records"
      Top             =   7980
      Width           =   2325
   End
   Begin VB.TextBox txtMonth 
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
      Height          =   285
      Left            =   1995
      MaxLength       =   2
      TabIndex        =   17
      Tag             =   "10-Birthdays due in this month"
      Top             =   4520
      Width           =   540
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   630
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
      Left            =   1680
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   300
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
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   960
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Top             =   1290
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
      Index           =   2
      Left            =   1680
      TabIndex        =   5
      Top             =   1620
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
      Left            =   1680
      TabIndex        =   6
      Top             =   1950
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDPT"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Tag             =   "10-Enter Employee Number"
      Top             =   2280
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6840
      Top             =   8520
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
   Begin Threed.SSCheck chkSupBirthYrAge 
      Height          =   255
      Left            =   3420
      TabIndex        =   16
      Tag             =   "Check to Supress Year of Birth and Age"
      Top             =   4215
      Visible         =   0   'False
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   " Suppress Year of Birth and Age"
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
   Begin Threed.SSCheck chkGender 
      Height          =   225
      Index           =   1
      Left            =   3420
      TabIndex        =   13
      Tag             =   "Show Female"
      Top             =   3915
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "   Female"
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
      Font3D          =   3
   End
   Begin Threed.SSCheck chkGender 
      Height          =   225
      Index           =   0
      Left            =   1980
      TabIndex        =   12
      Tag             =   "Show Male"
      Top             =   3915
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "   Male"
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
      Font3D          =   3
   End
   Begin INFOHR_Controls.DateLookup dlpAsOf 
      Height          =   285
      Left            =   1680
      TabIndex        =   20
      Tag             =   "40-As of Date"
      Top             =   5550
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpProv 
      Height          =   285
      Left            =   1680
      TabIndex        =   18
      Tag             =   "31-Province of Residence - Code"
      Top             =   4850
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   4
   End
   Begin INFOHR_Controls.CodeLookup clpProvEmp 
      Height          =   285
      Left            =   1680
      TabIndex        =   19
      Tag             =   "31-Province of Employment - Code"
      Top             =   5180
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   4
   End
   Begin INFOHR_Controls.DateLookup dlpSenDateRange 
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   24
      Tag             =   "40-Date upto and including this date forward"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpSenDateRange 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   23
      Tag             =   "40-Date from and including this date forward"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   7200
      TabIndex        =   3
      Tag             =   "00-Enter Physical Branch Code"
      Top             =   960
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SUDE"
   End
   Begin Threed.SSCheck chkShowCellPhone 
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Tag             =   "Check to show Cellular Telephone"
      Top             =   5925
      Visible         =   0   'False
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   " Show Cellular Telephone"
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
   Begin Threed.SSCheck chkGender 
      Height          =   225
      Index           =   2
      Left            =   4920
      TabIndex        =   14
      Tag             =   "Show 'Not Disclosed'"
      Top             =   3915
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "   Not Disclosed"
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
      Font3D          =   3
   End
   Begin VB.Label lblActBranch 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Physical Branch"
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
      Left            =   5760
      TabIndex        =   58
      Top             =   1020
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblSeniority 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Seniority"
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
      TabIndex        =   57
      Top             =   6285
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lblToDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   3360
      TabIndex        =   56
      Top             =   6285
      Visible         =   0   'False
      Width           =   195
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
      TabIndex        =   55
      Top             =   5225
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
      TabIndex        =   54
      Top             =   4895
      Width           =   1365
   End
   Begin VB.Label lblAgeTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   2880
      TabIndex        =   53
      Top             =   5925
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblAgeFrom 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Left            =   1440
      TabIndex        =   52
      Top             =   5925
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblAge 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
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
      TabIndex        =   51
      Top             =   5925
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label lblAsOf 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "As of Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   50
      Top             =   5595
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblGender 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      TabIndex        =   49
      Top             =   3915
      Width           =   645
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
      TabIndex        =   48
      Top             =   3635
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblYear 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Year of Birth"
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
      Top             =   4245
      Width           =   870
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
      Left            =   120
      TabIndex        =   46
      Top             =   2325
      Width           =   1290
   End
   Begin VB.Label lblSection 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   120
      TabIndex        =   45
      Top             =   3315
      Width           =   540
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
      TabIndex        =   44
      Top             =   2655
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
      TabIndex        =   43
      Top             =   2985
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
      TabIndex        =   42
      Top             =   1005
      Width           =   615
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
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
      Index           =   2
      Left            =   120
      TabIndex        =   41
      Top             =   8625
      Width           =   825
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #2"
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
      TabIndex        =   40
      Top             =   8310
      Width           =   885
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
      TabIndex        =   39
      Top             =   7995
      Width           =   885
   End
   Begin VB.Label lblMonth 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Month of Birth"
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
      Top             =   4565
      Width           =   990
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
      Top             =   1995
      Width           =   630
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
      TabIndex        =   36
      Top             =   1665
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
      TabIndex        =   35
      Top             =   1335
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
      TabIndex        =   34
      Top             =   675
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
      TabIndex        =   33
      Top             =   345
      Width           =   555
   End
   Begin VB.Label lblRepGrp 
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
      TabIndex        =   32
      Top             =   7680
      Width           =   1575
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
      TabIndex        =   31
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmRDOB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm(frmRDOB.Caption & " Criteria", Me) Then Exit Sub
      Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    X% = Cri_SetAll()
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
MsgBox "Error Printing - check your Windows Printer setup"
'MsgBox Error$ & " Error #:" & Str(Err)

Resume Next

End Sub

Public Sub cmdView_Click()
Dim X%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
    Call set_PrintState(False)
      
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
      
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    X% = Cri_SetAll()
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "HREMP", "SELECT")
Resume Next

End Sub

Private Sub chkExclCONP_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkExclRET_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkShowCellPhone_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkSupBirthYrAge_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    comGroup(0).Clear
    comGroup(1).Clear
    comGroup(2).Clear
    
    If frmRDOB.Caption = "Birthday/Age Report" Then
        comGroup(1).AddItem "Employee Name"
        comGroup(1).AddItem "Month of Birth"
        comGroup(1).AddItem "(none)"
        
        comGroup(2).AddItem "Employee Number"
        comGroup(2).AddItem "Month of Birth"
        
        comGroup(1).ListIndex = 0
        comGroup(2).Enabled = True
    Else
        comGroup(2).AddItem "Employee Name"
        comGroup(2).AddItem "Employee Number"
        comGroup(2).Enabled = True
    End If
    
    comGroup(2).ListIndex = 0
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0

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
    Case 5: strCd$ = "HREMP.ED_SECTION" 'Lucy June 29, 2000
    Case 6: strCd$ = "HREMP.ED_SUBDEPT" 'Ticket #24162 Franks 09/19/2013
    End Select
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

Private Sub Cri_Year()
Dim EECri As String, OneSet%, X%

If Len(txtYear.Text) < 1 Then Exit Sub

EECri = "{@BirthYear}= " & Val(txtYear.Text)

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_Month()
Dim EECri As String, OneSet%, X%

If Len(txtMonth.Text) < 1 Then Exit Sub

EECri = "{@BirthMonth}= " & Val(txtMonth.Text)

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_Age()
Dim EECri As String, OneSet%, X%

If Len(txtAgeFrom.Text) < 1 And Len(txtAgeTo.Text) < 1 Then Exit Sub

If Len(txtAgeFrom.Text) > 0 Then
    EECri = "{@Age}>= " & Val(txtAgeFrom.Text)
    
    If Len(txtAgeTo.Text) > 0 Then
        EECri = EECri & " AND {@Age}< " & Val(txtAgeTo.Text) + 1
    End If
Else
    If Len(txtAgeTo.Text) > 0 Then
        EECri = "{@Age}< " & Val(txtAgeTo.Text) + 1
    End If
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_PT()
Dim EECri As String, OneSet%, X%

If Len(clpPT.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

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

Private Sub Cri_Gender()
Dim gender

'Ticket #29313 - Added 'Not Disclosed' gender option
If chkGender(0).Value = True And chkGender(1).Value = True And chkGender(2).Value = True Then   'All checked
    gender = ""
ElseIf chkGender(0).Value = True And chkGender(1).Value = True Then     'Male and Female checked
    gender = "(({HREMP.ED_SEX} = 'M') OR ({HREMP.ED_SEX} = 'F'))"
ElseIf chkGender(0).Value = True And chkGender(2).Value = True Then     'Male and Not Disclosed checked
    gender = "(({HREMP.ED_SEX} = 'M') OR ({HREMP.ED_SEX} = 'N'))"
ElseIf chkGender(1).Value = True And chkGender(2).Value = True Then     'Female and Not Disclosed checked
    gender = "(({HREMP.ED_SEX} = 'F') OR ({HREMP.ED_SEX} = 'N'))"
ElseIf chkGender(0).Value = True Then
    gender = "({HREMP.ED_SEX} = 'M')"
ElseIf chkGender(1).Value = True Then
    gender = "({HREMP.ED_SEX} = 'F')"
ElseIf chkGender(2).Value = True Then
    gender = "({HREMP.ED_SEX} = 'N')"
Else
    'gender = "(ISNULL({HREMP.ED_SEX}))"
    gender = "((ISNULL({HREMP.ED_SEX})) OR (({HREMP.ED_SEX} <> 'F') AND ({HREMP.ED_SEX} <> 'M') AND ({HREMP.ED_SEX} <> 'N')))"
    'gender = "({HREMP.ED_SEX} <> 'F') AND ({HREMP.ED_SEX} <> 'M') AND ({HREMP.ED_SEX} <> 'N')"
End If

If Len(gender) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = gender
    Else
        glbstrSelCri = glbstrSelCri & " AND " & gender
    End If
    glbiOneWhere = True
End If
        
End Sub

Private Sub Cri_ProvResidence()
Dim EECri As String, OneSet%, X%

If Len(clpProv.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_PROV} in ['" & Replace(clpProv.Text, ",", "','") & "']"

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
EECri = "{HREMP.ED_PROVEMP} in ['" & Replace(clpProvEmp.Text, ",", "','") & "']"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Function Cri_SetAll()
Dim X%, strRName$
Dim dtYYY%, dtMM%, dtDD%
Dim xAsOfDate

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

Call glbCri_DeptUN(clpDept.Text)
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_PT
Call Cri_Shift
If frmRDOB.Caption = "Birthday/Age Report" Then
    Call Cri_Year
    Call Cri_Month
    Call Cri_Age
End If
    
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For X% = 0 To 6 '5 'Ticket #24162 Franks 09/20/2013
    Call Cri_Code(X%)
Next X%
Call Cri_Gender
Call Cri_EE
Call Cri_ProvResidence
Call Cri_ProvEmployment

Call Cri_SenDates 'Ticket #24162 Franks 09/20/2013

X% = Cri_Sorts()   ' returns number of sections formated


If frmRDOB.Caption = "Birthday/Age Report" Then
    
    'Ticket #26681 - Was still an issue with date switching between Month and Day so adding this format like
    'Hemu - Ticket #13778 - If the Computer Date format is yy/mm/dd then the Age was not getting
    'computed properly, therefore formatted the date to mm/dd/yyyy
    'Me.vbxCrystal.Formulas(11) = "AsOfDate=Date('" & CVDate(Format(dlpAsOf.Text, "mm/dd/yyyy")) & "')"
    dtYYY% = Year(dlpAsOf.Text)
    dtMM% = month(dlpAsOf.Text)
    dtDD% = Day(dlpAsOf.Text)
    xAsOfDate = "AsOfDate = DATE(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")"
    Me.vbxCrystal.Formulas(11) = xAsOfDate
    
    If comGroup(0) = "(none)" Then
        'Ticket #23772
        If chkSupBirthYrAge.Visible Then
            'Hemu - 07/10/2003 Begin - To Suppress the Year of Birth and Age
            If chkSupBirthYrAge Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzagenm1.rpt"
                'Hemu - 07/10/2003 End
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzagenme.rpt"
            End If
            
            'Hemu - 07/10/2003 Begin - To Suppress the Year of Birth and Age
            If chkSupBirthYrAge Then
                Me.vbxCrystal.Formulas(10) = "SuppAge='Y'"
            Else
                Me.vbxCrystal.Formulas(10) = "SuppAge='N'"
            End If
            'Hemu - 07/10/2003 End
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzagenm1.rpt"
        End If
    Else
        'Ticket #23772
        If chkSupBirthYrAge.Visible Then
            'Hemu - 07/10/2003 Begin - To Suppress the Year of Birth and Age
            If chkSupBirthYrAge Then
                Me.vbxCrystal.Formulas(10) = "SuppAge='Y'"
            Else
                Me.vbxCrystal.Formulas(10) = "SuppAge='N'"
            End If
            'Hemu - 07/10/2003 End
        End If
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzage.rpt"
    End If
    
    'Ticket #23772
    Me.vbxCrystal.Formulas(12) = "showDOB = " & IIf(gSec_Show_DOB = 0, False, True) & " "
        
Else 'THIS IS THE TELEPHONE EXTENTION REPORT
    'Ticket #28233 - For all
    If chkShowCellPhone Then
        Me.vbxCrystal.Formulas(10) = "ShowCell='Y'"
    Else
        Me.vbxCrystal.Formulas(10) = "ShowCell='N'"
    End If

    If comGroup(0) = "(none)" Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzintte1.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzinttel.rpt"
    End If
End If

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

If Len(glbstrSelCri) >= 0 Then
   Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

Me.vbxCrystal.Connect = RptODBC_SQL

' window title if appropriate
Me.vbxCrystal.WindowTitle = frmRDOB.Caption ' "Birthday/Age Report"

Cri_SetAll = True

Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Comp Report", "Select")
Cri_SetAll = False
Resume Next

End Function

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
' imbeded in report

Cri_Sorts = 0
grpField$ = getEGroup(comGroup(0).Text)

If frmRDOB.Caption = "Birthday/Age Report" Then
    If grpField$ <> "(none)" Then
        dscGroup$ = comGroup(0).Text
        dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
        Me.vbxCrystal.Formulas(0) = dscGroup$
        
        grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(0) = grpCond$
        
        GrpIdx% = comGroup(1).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{@EFullName}"
            Case 1: grpField$ = "{@BirthYear}"
            Case 2: grpField$ = "{@BirthMonth}"
        End Select
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(1) = grpCond$
            
            GrpIdx% = comGroup(2).ListIndex
            Select Case GrpIdx%
                Case 0: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
                Case 1: grpField$ = "{@BirthMonth}"
            End Select
            grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(2) = grpCond$
        End If
    Else
        GrpIdx% = comGroup(1).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{@EFullName}"
            Case 1: grpField$ = "{@BirthYear}"
            Case 2: grpField$ = "{@BirthMonth}"
            Case 3: grpField$ = "(none)"
        End Select
        If grpField$ <> "(none)" Then
            grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(0) = grpCond$
            
            grpCond$ = "GROUP" & CStr(2) & ";" & "{@BirthMonth}" & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(1) = grpCond$
        Else
            GrpIdx% = comGroup(2).ListIndex 'the item chosen in the final sort combo box
            Select Case GrpIdx%
                Case 0: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
                Case 1: grpField$ = "{@BirthMonth}"
            End Select
            'grpCond$ = "GROUP" & CStr(1) & ";" & "{@BirthMonth}" & ";ANYCHANGE;A"
            grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
            Me.vbxCrystal.GroupCondition(0) = grpCond$
        End If
        Me.vbxCrystal.Formulas(9) = "lblLocation='" & lStr("Location") & "'"
        Me.vbxCrystal.Formulas(19) = "lblUnion='" & lStr("Union") & "'" 'Ticket #27366 Franks 07/30/2015
    End If
    
    Call setRptLabel(Me, 0)
    
Else 'THIS IS THE TELEPHONE EXTENTION REPORT
    If grpField$ <> "(none)" Then
        dscGroup$ = comGroup(0).Text
        dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
        Me.vbxCrystal.Formulas(0) = dscGroup$
        
        grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(0) = grpCond$
        
        GrpIdx% = comGroup(2).ListIndex 'the item chosen in the final sort combo box
        Select Case GrpIdx%
            Case 0: grpField$ = "{@EFullName}"       'GROUP ON EMPLOYEE NAME
            Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
        End Select
        grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(1) = grpCond$
    Else
        GrpIdx% = comGroup(2).ListIndex 'the item chosen in the final sort combo box
        Select Case GrpIdx%
            Case 0: grpField$ = "{@EFullName}"       'GROUP ON EMPLOYEE NAME
            Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
        End Select
        grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(0) = grpCond$
    End If
    Me.vbxCrystal.Formulas(1) = "lblCellularTel='" & lStr("Cellular Telephone") & "'"
End If
    
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

For X% = 0 To 5
    If Not clpCode(X).ListChecker Then Exit Function
Next X%


If frmRDOB.Caption = "Birthday/Age Report" Then
  If Len(txtYear) > 0 Then
      If IsNumeric(txtYear) Then
          If Not (Val(txtYear) > 1899 And Val(txtYear) < 3000) Then
              MsgBox "Invalid Year of Birth Entered"
              txtYear.SetFocus
              Exit Function
          End If
      Else
          MsgBox "Year must be Number"
          txtYear.SetFocus
          Exit Function
      End If
  End If
  If Len(txtMonth) > 0 Then
      If IsNumeric(txtMonth) Then
          If Val(txtMonth) > 12 Or Val(txtMonth) < 1 Then
              MsgBox "Month must be between 1 and 12"
              txtMonth.SetFocus
              Exit Function
          End If
      Else
          MsgBox "Month must be between 1 and 12"
          txtMonth.SetFocus
          Exit Function
      End If
  End If
Else
'  If lblEEName.Caption = "Enter Valid Employee #" Then
'    MsgBox "Enter Valid Employee #"
'    txtEEID.SetFocus
'    Exit Function
'  End If
End If

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

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


If frmRDOB.Caption = "Birthday/Age Report" Then
    If Len(dlpAsOf.Text) = 0 Then
        MsgBox "As of Date is a required field!"
        dlpAsOf.SetFocus
        Exit Function
    Else
        If Not IsDate(dlpAsOf.Text) Then
            MsgBox "Invalid As of Date!"
            dlpAsOf.SetFocus
            Exit Function
        End If
    End If
        
    'Age Range
    If Len(txtAgeFrom) > 0 Then
        If Not IsNumeric(txtAgeFrom) Then
            MsgBox "Age From must be numeric"
            txtAgeFrom.SetFocus
            Exit Function
        End If
    End If
    If Len(txtAgeTo) > 0 Then
        If Not IsNumeric(txtAgeTo) Then
            MsgBox "Age To must be numeric"
            txtAgeTo.SetFocus
            Exit Function
        End If
    End If
    If IsNumeric(txtAgeFrom) And IsNumeric(txtAgeTo) Then
        If Val(txtAgeFrom) > Val(txtAgeTo) Then
            MsgBox "Invalid Age range. Age From cannot be greater than To age."
            txtAgeFrom.SetFocus
            Exit Function
        End If
    End If
End If

CriCheck = True

End Function

Private Sub Form_Activate()
Call SET_UP_MODE
'Ticket #13145 Frank 06/11/07
If frmRDOB.Caption = "Birthday/Age Report" Then
    dlpAsOf.Text = Date
    lblAsOf.Visible = True
    dlpAsOf.Visible = True
    
    'Ticket #18668 - Hemu
    lblAge.Visible = True
    lblAgeFrom.Visible = True
    lblAgeTo.Visible = True
    txtAgeFrom.Visible = True
    txtAgeTo.Visible = True
    
    'Ticket #28233
    chkShowCellPhone.Visible = False
Else
    'Ticket #28233
    chkShowCellPhone.Visible = True
    chkShowCellPhone.Top = 5925
    chkShowCellPhone.Left = 120
End If

End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = "FRMRDOB"

Screen.MousePointer = HOURGLASS

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call setRptCaption(Me)
lblSeniority.Caption = lStr("Seniority")

If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
If glbLinamar Then clpCode(3).MaxLength = 8
If glbSamuel Then 'Ticket #24162 Franks 09/20/2013
    Call SamuelScreenSetup
End If

'Ticket #29660 - Contract Employees Enhancement
If glbWFC Then
    chkExclCONP.Visible = True
    chkExclRET.Visible = True
Else
    chkExclCONP.Visible = False
    chkExclRET.Visible = False
End If

chkShowCellPhone.Caption = lStr("Show Cellular Telephone")

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT
End Sub

Private Sub SamuelScreenSetup()
    lblActBranch.Visible = True
    clpCode(6).Visible = True
    lblSeniority.Visible = True
    lblToDate.Visible = True
    dlpSenDateRange(0).Visible = True
    dlpSenDateRange(1).Visible = True
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

Private Sub txtAgeFrom_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtAgeFrom_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then KeyAscii = 0
End Sub

Private Sub txtMonth_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtMonth_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then KeyAscii = 0
End Sub

Private Sub txtShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtYear_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then KeyAscii = 0
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

Private Sub Cri_SenDates() 'Ticket #24162 Franks 09/19/2013
Dim TempCri As String
Dim TempCri2 As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%

If Len(dlpSenDateRange(0).Text) = 0 And Len(dlpSenDateRange(1).Text) = 0 Then
    Exit Sub
End If

If Len(dlpSenDateRange(0).Text) > 0 And Len(dlpSenDateRange(1).Text) > 0 Then
    TempCri = "({HREMP.ED_SENDTE}) "
    TempCri2 = "({HREMP.ED_SENDTE}) "

    dtYYY% = Year(dlpSenDateRange(0).Text)
    dtMM% = month(dlpSenDateRange(0).Text)
    dtDD% = Day(dlpSenDateRange(0).Text)
    'TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpSenDateRange(1).Text)
    dtMM% = month(dlpSenDateRange(1).Text)
    dtDD% = Day(dlpSenDateRange(1).Text)
    TempCri2 = TempCri2 & " <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    TempCri = TempCri & " AND " & TempCri2
    GoTo Cri_SenDatst
End If

For X% = 0 To 1
    TempCri = "({HREMP.ED_SENDTE}) "
    TempCri2 = "({HREMP.ED_SENDTE}) "
        
    If Len(dlpSenDateRange(0).Text) > 0 Then
        TempCri = TempCri & " >= "
        dtYYY% = Year(dlpSenDateRange(0).Text)
        dtMM% = month(dlpSenDateRange(0).Text)
        dtDD% = Day(dlpSenDateRange(0).Text)
    End If
    If Len(dlpSenDateRange(1).Text) > 0 Then
        TempCri = TempCri2 & " <= "
        dtYYY% = Year(dlpSenDateRange(1).Text)
        dtMM% = month(dlpSenDateRange(1).Text)
        dtDD% = Day(dlpSenDateRange(1).Text)
    End If

    TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    GoTo Cri_SenDatst

Next X%



Cri_SenDatst:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
    End If
    glbiOneWhere = True
End If

End Sub
