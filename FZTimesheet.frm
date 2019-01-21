VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmRTimesheet 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Timesheet"
   ClientHeight    =   8145
   ClientLeft      =   375
   ClientTop       =   915
   ClientWidth     =   13260
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
   ScaleHeight     =   8145
   ScaleWidth      =   13260
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame frmAT 
      Height          =   435
      Left            =   120
      TabIndex        =   57
      Top             =   360
      Width           =   9315
      Begin VB.OptionButton optAT 
         Caption         =   "Terminated Employee"
         Height          =   255
         Index           =   1
         Left            =   5250
         TabIndex        =   1
         Top             =   150
         Width           =   2655
      End
      Begin VB.OptionButton optAT 
         Caption         =   "Active Employee"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   0
         Top             =   150
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.CheckBox chkAbsence 
      Caption         =   "Absent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8880
      TabIndex        =   53
      Top             =   10080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.OptionButton optAnnual 
      Caption         =   "Annual"
      Height          =   195
      Left            =   4680
      TabIndex        =   47
      Top             =   8880
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.OptionButton optMonth 
      Caption         =   "Month"
      Height          =   195
      Left            =   6600
      TabIndex        =   46
      Top             =   8880
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.ComboBox cmbMonth 
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
      ItemData        =   "FZTimesheet.frx":0000
      Left            =   6600
      List            =   "FZTimesheet.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   9480
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.TextBox txtYearTo 
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
      Left            =   10560
      MaxLength       =   4
      TabIndex        =   44
      Tag             =   "61- Enter To Year"
      Top             =   9150
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbMonthTo 
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
      ItemData        =   "FZTimesheet.frx":008E
      Left            =   10560
      List            =   "FZTimesheet.frx":00B6
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   9480
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.OptionButton optRolling 
      Caption         =   "Rolling"
      Height          =   195
      Left            =   10200
      TabIndex        =   42
      Top             =   8880
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.OptionButton optWeek 
      Caption         =   "Weekly"
      Height          =   195
      Left            =   8520
      TabIndex        =   41
      Top             =   8910
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.OptionButton optBiWeek 
      Caption         =   "Bi-Weekly"
      Height          =   195
      Left            =   11520
      TabIndex        =   40
      Top             =   8910
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CheckBox chkBlank 
      Caption         =   "Blank Report"
      Height          =   255
      Left            =   11850
      TabIndex        =   39
      Top             =   9180
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txtWeek 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2265
      TabIndex        =   10
      Top             =   3630
      Width           =   1335
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2265
      TabIndex        =   9
      Top             =   3285
      Width           =   1335
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2265
      MaxLength       =   4
      TabIndex        =   16
      Tag             =   "00-Shift"
      Top             =   5295
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1950
      TabIndex        =   15
      Tag             =   "EDSE-Section "
      Top             =   4950
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1950
      TabIndex        =   14
      Tag             =   "EDAB-Administered By"
      Top             =   4620
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1950
      TabIndex        =   13
      Tag             =   "EDRG-Region"
      Top             =   4290
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   1950
      TabIndex        =   7
      Tag             =   "EDPT-Category"
      Top             =   2625
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
      Index           =   3
      Left            =   1950
      TabIndex        =   6
      Top             =   2295
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDEM"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1950
      TabIndex        =   4
      Tag             =   "EDLC-Location"
      Top             =   1620
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1950
      TabIndex        =   3
      Tag             =   "00-Specific Department Desired"
      Top             =   1290
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
      Left            =   1950
      TabIndex        =   2
      Tag             =   "00-Specific Division Desired"
      Top             =   960
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Tag             =   "Second level of grouping records"
      Top             =   6450
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
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Tag             =   "First Level of grouping records"
      Top             =   7935
      Visible         =   0   'False
      Width           =   2325
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1950
      TabIndex        =   8
      Tag             =   "10-Enter Employee Number"
      Top             =   2955
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   0
      Top             =   7000
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1950
      TabIndex        =   5
      Tag             =   "00-Enter Union Code"
      Top             =   1965
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDOR"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpPayP 
      Height          =   285
      Left            =   6360
      TabIndex        =   34
      Top             =   7320
      Visible         =   0   'False
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   503
      ShowDescription =   0   'False
      TABLName        =   "SDPP"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3930
      TabIndex        =   12
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3960
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1950
      TabIndex        =   11
      Tag             =   "40-Date from and including this date forward"
      Top             =   3960
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.DateLookup dlpStartdate 
      Height          =   285
      Left            =   6270
      TabIndex        =   48
      Tag             =   "40-Date from and including this date forward"
      Top             =   9150
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin Threed.SSCheck chkShowEmp 
      Height          =   255
      Left            =   6240
      TabIndex        =   54
      Tag             =   "If X-Show All Employees"
      Top             =   10080
      Visible         =   0   'False
      Width           =   2475
      _Version        =   65536
      _ExtentX        =   4366
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   " Show All Employees"
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   10
      Left            =   6240
      TabIndex        =   55
      Tag             =   "ADRE-Attendance Reason"
      Top             =   8400
      Visible         =   0   'False
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "ADRE"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin VB.Label lblAttCodes 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Attendance Codes"
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
      Left            =   4680
      TabIndex        =   56
      Top             =   8400
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4710
      TabIndex        =   52
      Top             =   9180
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblMonth 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
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
      Left            =   4710
      TabIndex        =   51
      Top             =   9540
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblYearTo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "To Year"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      TabIndex        =   50
      Top             =   9180
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblMonthTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "To Month"
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
      Left            =   9240
      TabIndex        =   49
      Top             =   9540
      Visible         =   0   'False
      Width           =   690
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
      Left            =   150
      TabIndex        =   38
      Top             =   4005
      Width           =   1095
   End
   Begin VB.Image imgIcon 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1920
      Picture         =   "FZTimesheet.frx":011C
      Top             =   3630
      Width           =   240
   End
   Begin VB.Label lblWeek 
      Caption         =   "Pay Period #"
      Height          =   195
      Left            =   150
      TabIndex        =   37
      Top             =   3675
      Width           =   1395
   End
   Begin VB.Label lblYear 
      Caption         =   "Year"
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
      Left            =   150
      TabIndex        =   36
      Top             =   3330
      Width           =   1395
   End
   Begin VB.Label lblBCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Period"
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
      Left            =   4800
      TabIndex        =   35
      Top             =   7380
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblShift 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Left            =   150
      TabIndex        =   33
      Top             =   5340
      Visible         =   0   'False
      Width           =   315
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
      Left            =   150
      TabIndex        =   32
      Top             =   2010
      Width           =   420
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
      Left            =   150
      TabIndex        =   31
      Top             =   2340
      Width           =   450
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
      Left            =   150
      TabIndex        =   30
      Top             =   1665
      Width           =   615
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
      Left            =   150
      TabIndex        =   29
      Top             =   4665
      Width           =   1125
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
      Left            =   150
      TabIndex        =   28
      Top             =   2670
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
      Left            =   150
      TabIndex        =   27
      Top             =   4995
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
      Left            =   150
      TabIndex        =   26
      Top             =   4335
      Width           =   510
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
      TabIndex        =   25
      Top             =   6120
      Width           =   1575
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
      Left            =   60
      TabIndex        =   24
      Top             =   6510
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
      Left            =   4740
      TabIndex        =   23
      Top             =   7965
      Visible         =   0   'False
      Width           =   885
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
      TabIndex        =   22
      Top             =   150
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
      Left            =   150
      TabIndex        =   21
      Top             =   3000
      Width           =   1290
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
      Left            =   150
      TabIndex        =   20
      Top             =   1335
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
      Left            =   150
      TabIndex        =   19
      Top             =   1005
      Width           =   555
   End
End
Attribute VB_Name = "frmRTimesheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fsnapEENames As Recordset
Dim DATE1, DATE2

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Sub cmdClose_Click()
Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdPrint_Click()
Dim X%

'''On Error GoTo PrntErr
 
If CriCheck() Then

    If Not PrtForm(Me.Caption, Me) Then Exit Sub
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
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdPrint_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdView_Click()
Dim X%
Dim strWHand As String
'''On Error GoTo CRW_Err

If CriCheck() Then
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    Screen.MousePointer = HOURGLASS
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
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ENTITLEMENTS", "VIEW")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdView_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "Employee Name"
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    comGroup(1).AddItem "Employee Name"
    comGroup(1).ListIndex = 0

End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    Select Case intIdx%
    Case 0: strCd$ = "HREMP.ED_LOC"
    Case 1: strCd$ = "HREMP.ED_REGION"
    Case 2: strCd$ = "HREMP.ED_SECTION"
    Case 3: strCd$ = "HREMP.ED_EMP"
    Case 4: strCd$ = "HREMP.ED_ADMINBY"
    Case 5: strCd$ = "HREMP.ED_ORG"
    Case 6: strCd$ = "HREMP.ED_PT"
    End Select
    'CodeCri = "(" & strCd$ & " = '" & clpCode(intIdx%).Text & "')"
    CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    
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
Dim countr   As Integer  ' EEList_Snap is definded at form level

If Len(clpDiv.Text) > 0 Then
    DivCri = "(HREMP.ED_DIV in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
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
    EECri = "HREMP.ED_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
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

Private Function Cri_SetAll()
Dim X%, strRName$

Cri_SetAll = False
'''On Error GoTo modSetCriteria_Err

Screen.MousePointer = HOURGLASS

glbiOneWhere = True
glbstrSelCri = ""

Call Cri_Dept
Call Cri_Div

For X% = 0 To 6
    Call Cri_Code(X%)
Next X%
Call Cri_EE

'Ticket #28002 - Opening for all clients and adding Termination employee option as well
If optAT(0) <> 0 Then
    Call AttWrk
Else
    Call AttWrk_Terminated
End If

'Ticket #28002 - Opening for all clients and adding Termination employee option as well
If optAT(0) <> 0 Then
    strRName$ = glbIHRREPORTS & "rztimesheet.rpt"
Else
    strRName$ = glbIHRREPORTS & "rztimesheetT.rpt"
End If

If glbCompSerial = "S/N - 2174W" Then
    strRName$ = glbIHRREPORTS & "SN2174_RZTimesheet.rpt"
End If

Me.vbxCrystal.ReportFileName = strRName$

X% = Cri_Sorts()   ' returns number of sections formated

Me.vbxCrystal.SelectionFormula = "{HR_ATT_TIMESHEET.AD_WRKEMP}='" & glbUserID & "'"
Me.vbxCrystal.WindowTitle = Me.Caption

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = RptODBC_SQL
    Me.vbxCrystal.Connect = "PWD=petman;"
    Me.vbxCrystal.DataFiles(0) = glbIHRDBW
    'For x% = 1 To 7
    '    Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    'Next x%
End If

'Ticket #28002 - Opening for all clients and adding Termination employee option as well
' window title if appropriate
If optAT(0) <> 0 Then
    Me.vbxCrystal.WindowTitle = "Timesheet Report for Active Employees"
Else
    Me.vbxCrystal.WindowTitle = "Timesheet Report for Terminated Employees"
End If


Cri_SetAll = True

Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Timesheet", "Timesheet Report", "Select")
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
Dim strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%
Dim strSMonth$
'for labels - sort by name always
'imbeded in report

Cri_Sorts = 0

'grpField$ = getEGroup(comGroup(0).Text)
'If grpField$ = "(none)" Then grpField$ = "{HRPARCO.PC_CO}"
'If grpField$ = lStr("Region") Then grpField$ = "{@productline}"
'
'If glbCompSerial = "S/N - 2327W" And comGroup(0).Text = "Employee Name" Then
'    dscGroup$ = "Associate Name"
'Else
'    dscGroup$ = comGroup(0).Text
'End If
'dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
'Me.vbxCrystal.Formulas(0) = dscGroup$
'
'grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
'Me.vbxCrystal.GroupCondition(0) = grpCond$

'Hemu - 02/17/2004 Begin
'Custom code for Brant County Health Unit
'If glbCompSerial = "S/N - 2226W" And comGroup(0).ListIndex = 0 Then 'InStr(1, grpCond$, "Division_Name") > 0
'    grpCond$ = "GROUP" & CStr(2) & ";" & "{HRDEPT.DF_NAME}" & ";ANYCHANGE;A"
'    Me.vbxCrystal.GroupCondition(1) = grpCond$
'Else
''Hemu - 02/17/2004 End
'    GrpIdx% = comGroup(1).ListIndex
'    Select Case GrpIdx%
'        Case 0: grpField$ = "{@EFullName}"
'    End Select
'    grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
'    Me.vbxCrystal.GroupCondition(1) = grpCond$
'
'End If


Call setRptLabel(Me, 0)
'Cri_Sorts = z% ' next section number to format
If DATE1 <> "" And DATE2 <> "" Then
    strSFormat$ = "As of " & DATE1 & " through " & DATE2
    Me.vbxCrystal.Formulas(1) = "Daterange = '" & strSFormat$ & "'"
Else
    strSFormat$ = "No date entered"
    Me.vbxCrystal.Formulas(1) = "Daterange = '" & strSFormat$ & "'"
End If
If glbWFC Then
    Me.vbxCrystal.Formulas(2) = "HideJob = False"
End If

End Function

Private Function CriCheck()
Dim X%

CriCheck = False

If Len(txtWeek.Text) = 0 Then
    MsgBox "Pay Period # is requried field"
    txtWeek.SetFocus
    Exit Function
End If

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

For X% = 0 To 6
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

clpCode(10).Text = Replace(clpCode(10).Text, " ", "")

If optWeek Or optBiWeek Then
'    If dlpDateRange(0) = "" Then
'        MsgBox "You have to enter the Start Date!"
'        dlpDateRange(0).SetFocus
'        Exit Function
'    Else
'        If Not IsDate(dlpDateRange(0)) Then
'            MsgBox "Start Date is not a valid date"
'            dlpDateRange(0).SetFocus
'            Exit Function
'        End If
'    End If
    If glbCompSerial = "S/N - 2241W" And optBiWeek Then 'granite club
        If WeekdayName(Weekday(dlpDateRange(0))) <> "Sunday" Then
            MsgBox "Start Date must be Sunday"
            dlpDateRange(0).SetFocus
            Exit Function
        End If
    End If
Else
    If txtYear = "" Then
        MsgBox "Year is required"
        txtYear.SetFocus
        Exit Function
    Else
        If Val(txtYear) > Year(Date) + 100 Or Val(txtYear) < Year(Date) - 100 Then
            MsgBox "Invalid Year"
            txtYear.SetFocus
            Exit Function
        End If
    End If

    If optRolling Then
        If txtYearTo = "" Then
            MsgBox "Year is required"
            txtYearTo.SetFocus
            Exit Function
        Else
            If Val(txtYearTo) > Year(Date) + 100 Or Val(txtYearTo) < Year(Date) - 100 Then
                MsgBox "To Year is invalid"
                txtYearTo.SetFocus
                Exit Function
            End If
        End If
        If CVDate(cmbMonthTo & " 01," & txtYearTo) < CVDate(cmbMonth & " 01," & txtYear) Then
            MsgBox "To Date can not be earlier than From Date"
            txtYear.SetFocus
            Exit Function
        End If
        If DateDiff("m", CVDate(cmbMonth & " 01," & txtYear), CVDate(cmbMonthTo & " 01," & txtYearTo)) > 11 Then
            MsgBox "You can not view this report more than 12 months"
            txtYear.SetFocus
            Exit Function
        End If
    End If
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

'If glbCompSerial = "S/N - 2214W" Then
'    If Len(clpCode(7)) > 0 Then
'        If clpCode(7).Caption = "Unassigned" Then
'            MsgBox "Invalid Attendance-Fund code"
'            clpCode(7).SetFocus
'            Exit Function
'        End If
'    End If
'End If

CriCheck = True

End Function

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()

Screen.MousePointer = HOURGLASS

glbOnTop = Me.name

Call comGrpLoad

'Call setCaption(lblDiv)
'Call setCaption(lblRegion)
'Call setCaption(lblSection)
'Call setCaption(lblDept)
'Call setRptCaption(Me)
'lblFromTo.Caption = lStr("From Date") & " / " & lStr("To Date")

If glbCompSerial = "S/N - 2227W" Then clpCode(1).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

Call INI_Controls(Me)

txtYear = Year(Date)

'chkShowEmp.Visible = True
If glbLinamar Then
    clpCode(1).MaxLength = 8
End If
If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If
If glbCompSerial = "S/N - 2241W" Then 'granite club
    optBiWeek.Visible = True
Else
    optBiWeek.Visible = False
End If
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
gdbAdoIhr001.Execute "DELETE FROM HR_ATT_TIMESHEET " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub

Private Sub optAnnual_Click()
lblMonth.Visible = optMonth
cmbMonth.Visible = optMonth
lblMonthTo.Visible = optRolling
cmbMonthTo.Visible = optRolling
lblYearTo.Visible = optRolling
txtYearTo.Visible = optRolling
lblYear.Caption = "Year"
lblMonth.Caption = "Month"
dlpDateRange(0).Visible = optWeek Or optBiWeek
txtYear.Visible = Not (optWeek Or optBiWeek)
chkBlank.Visible = False
End Sub

Private Sub optAT_Click(Index As Integer)
    If Index = 1 Then
        elpEEID.LookupType = TERM
    Else
        elpEEID.LookupType = 0  '0 = ACTIVE. I cannot put as ACTIVE because it's changing to "Active" and that does not switch the lookup to ACTIVE employees
    End If
End Sub

Private Sub optBiWeek_Click()
lblYear = "Start Date"
dlpDateRange(0).Visible = optWeek Or optBiWeek

txtYear.Visible = Not (optWeek Or optBiWeek)
lblMonth.Visible = Not (optWeek Or optBiWeek)
cmbMonth.Visible = Not (optWeek Or optBiWeek)

lblYearTo.Visible = optRolling
txtYearTo.Visible = optRolling
lblMonthTo.Visible = optRolling
cmbMonthTo.Visible = optRolling
chkBlank.Visible = True
End Sub

Private Sub optMonth_Click()
lblMonth.Visible = optMonth
cmbMonth.Visible = optMonth
lblMonthTo.Visible = optRolling
cmbMonthTo.Visible = optRolling
lblYearTo.Visible = optRolling
txtYearTo.Visible = optRolling
lblYear.Caption = "Year"
lblMonth.Caption = "Month"
dlpDateRange(0).Visible = optWeek Or optBiWeek
txtYear.Visible = Not (optWeek Or optBiWeek)
If cmbMonth.ListIndex = -1 Then cmbMonth.ListIndex = 0
chkBlank.Visible = False
End Sub

Private Sub Cri_Dept()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim DeptCri As String

If Len(clpDept.Text) > 0 Then DeptCri = " AND (ED_DEPTNO in ['" & Replace(clpDept.Text, ",", "','") & "']) "

glbstrSelCri = glbSeleDeptUn & DeptCri

End Sub

Private Sub optRolling_Click()

lblMonth.Visible = optRolling
cmbMonth.Visible = optRolling
lblYearTo.Visible = optRolling
txtYearTo.Visible = optRolling
lblMonthTo.Visible = optRolling
cmbMonthTo.Visible = optRolling

lblYear.Caption = "From Year"
lblMonth.Caption = "From Month"

dlpDateRange(0).Visible = optWeek Or optBiWeek

txtYear.Visible = Not (optWeek Or optBiWeek)

If cmbMonth.ListIndex = -1 Then cmbMonth.ListIndex = 0
If cmbMonthTo.ListIndex = -1 Then cmbMonthTo.ListIndex = 0

chkBlank.Visible = False
End Sub

Private Sub optWeek_Click()
lblYear = "Start Date"
dlpDateRange(0).Visible = optWeek Or optBiWeek

txtYear.Visible = Not (optWeek Or optBiWeek)
lblMonth.Visible = Not (optWeek Or optBiWeek)
cmbMonth.Visible = Not (optWeek Or optBiWeek)

lblYearTo.Visible = optRolling
txtYearTo.Visible = optRolling
lblMonthTo.Visible = optRolling
cmbMonthTo.Visible = optRolling
chkBlank.Visible = False
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

Private Sub txtYear_GotFocus()      ' Serbo
Call SetPanHelp(Me.ActiveControl)   '
End Sub                             '

Function StripChar(StringToStrip, CharToStrip)
    Dim I, buf, OneChar
    
    For I = 1 To Len(StringToStrip)
        OneChar = Mid(StringToStrip, I, 1)
        If OneChar <> CharToStrip Then buf = buf & OneChar
    Next I
    StripChar = buf
End Function

Private Sub txtWeek_Change()
Dim DateRange
DateRange = Split(getDateRange("", txtWeek, txtYear), "|")
    dlpDateRange(0) = DateRange(0)
    dlpDateRange(1) = DateRange(1)
End Sub

Private Sub txtWeek_DblClick()
Call imgIcon_Click
End Sub

Private Sub imgIcon_Click()
frmPayPeriodList.SelectedYear = Val(txtYear)
'frmPayPeriodList.PayPeriodCode = clpPayP.Text
frmPayPeriodList.Show 1
txtWeek = glbWeek
dlpDateRange(0) = glbFrom
dlpDateRange(1) = glbTo
End Sub

Private Sub txtWeek_LostFocus()
If txtWeek = "" Then
    dlpDateRange(0) = ""
    dlpDateRange(1) = ""
Else
    'FIND THE DATA RANGE FROM THE DATABASE FOR THAT WEEK #
End If
End Sub

Private Sub txtYear_Change()
Dim DateRange
DateRange = Split(getDateRange("", txtWeek, txtYear), "|")
    dlpDateRange(0) = DateRange(0)
    dlpDateRange(1) = DateRange(1)
End Sub

Function getDateRange(theClientNumber, thePayNbr, theYear)
Dim rsPayPeriod As New ADODB.Recordset
Dim SQLQ, intNum

On Error Resume Next

getDateRange = "|"

If Not IsNumeric(thePayNbr) Then Exit Function
If Not IsNumeric(theYear) Then Exit Function

SQLQ = "SELECT PP_NBR,PP_YEAR,PP_Start,PP_End FROM HR_payperiod WHERE PP_PAYP='" & theClientNumber & "'"
SQLQ = SQLQ & " and PP_NBR = " & thePayNbr
SQLQ = SQLQ & " and PP_YEAR = '" & theYear & "'"
rsPayPeriod.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

If Not rsPayPeriod.EOF Then
    getDateRange = rsPayPeriod("PP_Start") & "|" & rsPayPeriod("PP_End")
   
End If
rsPayPeriod.Close
Exit Function

End Function

Private Function getWSQLQ(WithAtt As Boolean)
Dim QStr

QStr = glbSeleDeptUn
If clpDiv <> "" Then QStr = QStr & " AND ED_DIV in ('" & Replace(clpDiv, ",", "','") & "')"
If clpDept.Text <> "" Then QStr = QStr & " AND ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')" 'Ticket #30391 Franks 08/10/2017
If clpCode(0) <> "" Then QStr = QStr & " AND ED_LOC='" & clpCode(0) & "'"
If clpCode(1) <> "" Then QStr = QStr & " AND ED_ORG in ('" & Replace(clpCode(1), ",", "','") & "')"
If clpCode(2) <> "" Then QStr = QStr & " AND ED_EMP in ('" & Replace(clpCode(2), ",", "','") & "')"
If clpCode(3) <> "" Then QStr = QStr & " AND ED_REGION='" & clpCode(3) & "'"
If clpCode(4) <> "" Then QStr = QStr & " AND ED_ADMINBY='" & clpCode(4) & "'"
If clpCode(5) <> "" Then QStr = QStr & " AND ED_SECTION='" & clpCode(5) & "'"
If clpCode(6) <> "" Then QStr = QStr & " AND ED_PT in ('" & Replace(clpCode(6), ",", "','") & "')"
If elpEEID.Text <> "" Then QStr = QStr & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
If WithAtt Then
    If IsDate(dlpDateRange(0)) Then QStr = QStr & " AND AD_DOA>=" & Date_SQL(DATE1)
    If IsDate(dlpDateRange(1)) Then QStr = QStr & " AND AD_DOA<=" & Date_SQL(DATE1)
'    If clpAtt <> "" Then QStr = QStr & " AND ES_CTYPE IN ('" & Replace(clpAtt, ",", "','") & "') "
'    If clpPayP <> "" Then QStr = QStr & " AND ES_CRSCODE IN ('" & Replace(clpPayP, ",", "','") & "') "
'    If txtShift <> "" Or clpJob <> "" Or clpPosGroup <> "" Then
'        QStr = QStr & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0"
'        If txtShift <> "" Then QStr = QStr & " AND JH_SHIFT='" & txtShift & "'"
'        If clpJob <> "" Then QStr = QStr & " AND JH_JOB='" & clpJob & "'"
'        If clpPosGroup <> "" Then QStr = QStr & " AND JH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD='" & clpPosGroup & "')"
'        QStr = QStr & ")"
'    End If
End If

getWSQLQ = QStr
End Function

Private Sub AttWrk()
'Dim CoJobCodeS As New Collection
'Dim CoJobCode As New Collection
'Dim rsHours As New ADODB.Recordset
Dim rsAT As New ADODB.Recordset
Dim rsATTCal As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim strStatus

Dim iWorksheet As Integer
Dim iDate, DayNum As Integer
Dim strRANGE As String
Dim sFlag As Boolean
Dim xRegularRate
Dim dtDate As Date
Dim strTableName
'Dim blnFound
Dim xDay, xField
Dim xHours
Dim SQLQ
Dim xEmpnbr
Dim gdbESS As New ADODB.Connection

If glbSQL Or glbOracle Then
    Set gdbESS = gdbAdoIhr001
Else
    gdbESS.Open Replace(glbAdoIHRDB, "IHR001", "IHRESS")
End If

On Error GoTo Err_XLS
    DATE1 = dlpDateRange(0)
    DATE2 = dlpDateRange(1)
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 20
    MDIMain.panHelp(1).Caption = " Please Wait"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute "DELETE FROM HR_ATT_TIMESHEET " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001.CommitTrans
    If Not glbSQL And Not glbOracle Then Pause (0.5)
    'Get Attendace code....
    SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME FROM HREMP where "
    SQLQ = SQLQ & getWSQLQ(False)

    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

    If rsTemp.EOF Then
         MsgBox ("There are no employees based on the search criteria.")
    End If
     
    Do While Not rsTemp.EOF  'processing all employees who have timesheet data in HR_Timesheet or HR_Attendance
        xEmpnbr = rsTemp.Fields("ED_EMPNBR")
        strStatus = getStatus(xEmpnbr, DATE1, DATE2)
        Select Case strStatus
        Case "APPROVED"
            strTableName = "HR_ATTENDANCE"
        Case ""
            'strTableName = "HR_ATTENDANCE"
            GoTo Loopend
        Case Else
            strTableName = "HR_TIMESHEET"
        End Select

        SQLQ = "SELECT AD_COMPNO,'" & glbUserID & "' AS AD_WRKEMP,AD_EMPNBR,AD_DOA,AD_HRS,AD_REASON,  AD_SHIFT "
        SQLQ = SQLQ & " FROM " & strTableName
        '& ", HRTABL "
        SQLQ = SQLQ & " WHERE AD_EMPNBR =" & xEmpnbr
        SQLQ = SQLQ & " AND AD_DOA>=" & Date_SQL(DATE1)
        SQLQ = SQLQ & " AND AD_DOA<=" & Date_SQL(DATE2)
        
        'Ticket #30218 - Filter out the place holder/marker (AD_HRS = 9999) used to identify dates which has one or more attendance records.
        'These are not the hours entered by the user
        SQLQ = SQLQ & " AND AD_HRS <> 9999"
        
        'SQLQ = SQLQ & " AND (AD_REASON = TB_KEY) AND (TB_NAME = 'ADRE')"
        
'        If strTableName = "HR_ATTENDANCE" Then
'            SQLQ = SQLQ & " UNION "
'            SQLQ = SQLQ & " SELECT AH_COMPNO,'" & glbUserID & "' AS AH_WRKEMP,AH_EMPNBR,AH_DOA,AH_HRS,AH_REASON,  AH_SHIFT "
'            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY, HRTABL "
'            SQLQ = SQLQ & " WHERE AH_EMPNBR =" & rsTemp.Fields("ED_EMPNBR")
'            SQLQ = SQLQ & " AND AH_DOA>=" & Date_SQL(DATE1)
'            SQLQ = SQLQ & " AND AH_DOA<=" & Date_SQL(DATE2)
'            SQLQ = SQLQ & " AND (AH_REASON = TB_KEY) AND (TB_NAME = 'ADRE')"
'        End If

        SQLQ = SQLQ & " order by AD_DOA"
        
        If glbSQL Or glbOracle Then
            rsAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockReadOnly
        Else
            If strTableName = "HR_TIMESHEET" Then
                rsAT.Open SQLQ, gdbESS, adOpenKeyset, adLockReadOnly
            Else
                rsAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockReadOnly
            End If
        End If
        '        xxx = rsAT.RecordCount
        '        xx1 = 0
        Do Until rsAT.EOF
'            xx1 = xx1 + 1
'            MDIMain.panHelp(0).FloodPercent = (xx1 / xxx) * 60 + 30
            xDay = DateDiff("d", DATE1, rsAT!AD_DOA) + 1
            xField = "AD_DAY" & xDay
            
            SQLQ = "Select * from HR_ATT_TIMESHEET  "
            SQLQ = SQLQ & " where AD_EMPNBR=" & xEmpnbr
            SQLQ = SQLQ & " AND  AD_WRKEMP='" & glbUserID & "'"
            SQLQ = SQLQ & " AND AD_REASON ='" & rsAT!AD_REASON & "'"
            
            If glbSQL Or glbOracle Then
                rsATTCal.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            Else
                rsATTCal.Open SQLQ, gdbAdoIhr001W, adOpenStatic, adLockOptimistic
            End If
            
            xHours = 0
            If rsATTCal.EOF Then
                rsATTCal.AddNew
                rsATTCal!AD_EMPNBR = xEmpnbr
                rsATTCal!AD_COMPNO = "001"
                rsATTCal!AD_WRKEMP = glbUserID
                rsATTCal!AD_DOA = DATE1
                rsATTCal!AD_REASON = rsAT!AD_REASON
            End If
            rsATTCal!AD_STATUS = strStatus
            If IsNull(rsATTCal(xField)) Then
                xHours = 0
            Else
                xHours = rsATTCal(xField)
            End If
            rsATTCal(xField) = xHours + rsAT!AD_HRS
            rsATTCal.Update
            rsATTCal.Close
            rsAT.MoveNext
        Loop
        rsAT.Close
Loopend:
        rsTemp.MoveNext
    Loop
    rsTemp.Close

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Exit Sub
Err_XLS:
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT

    If Err = 1004 Then
        Resume Next
    End If

    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Timesheet", "", "Select")
Resume Next
End Sub

Private Sub AttWrk_Terminated()

'Ticket #28002 - Opening for all clients and adding Termination employee option as well

'Dim CoJobCodeS As New Collection
'Dim CoJobCode As New Collection
'Dim rsHours As New ADODB.Recordset
Dim rsAT As New ADODB.Recordset
Dim rsATTCal As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim strStatus

Dim iWorksheet As Integer
Dim iDate, DayNum As Integer
Dim strRANGE As String
Dim sFlag As Boolean
Dim xRegularRate
Dim dtDate As Date
Dim strTableName
'Dim blnFound
Dim xDay, xField
Dim xHours
Dim SQLQ
Dim xEmpnbr
Dim gdbESS As New ADODB.Connection

If glbSQL Or glbOracle Then
    Set gdbESS = gdbAdoIhr001
Else
    gdbESS.Open Replace(glbAdoIHRDB, "IHR001", "IHRESS")
End If

On Error GoTo Err_XLS_T

    DATE1 = dlpDateRange(0)
    DATE2 = dlpDateRange(1)
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 20
    MDIMain.panHelp(1).Caption = " Please Wait"
    
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute "DELETE FROM HR_ATT_TIMESHEET " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001.CommitTrans
    
    If Not glbSQL And Not glbOracle Then Pause (0.5)
    'Get Attendace code....
    SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME FROM TERM_HREMP WHERE "
    SQLQ = SQLQ & getWSQLQ(False)

    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

    If rsTemp.EOF Then
         MsgBox ("There are no employees based on the search criteria.")
    End If
     
    Do While Not rsTemp.EOF  'processing all employees who have timesheet data in HR_Timesheet or HR_Attendance
        xEmpnbr = rsTemp.Fields("ED_EMPNBR")
        strStatus = getStatus(xEmpnbr, DATE1, DATE2)
        Select Case strStatus
        Case "APPROVED"
            strTableName = "TERM_ATTENDANCE"
        Case ""
            'strTableName = "TERM_ATTENDANCE"
            GoTo Loopend
        Case Else
            strTableName = "TERM_TIMESHEET"
        End Select

        SQLQ = "SELECT AD_COMPNO,'" & glbUserID & "' AS AD_WRKEMP,AD_EMPNBR,AD_DOA,AD_HRS,AD_REASON,  AD_SHIFT "
        SQLQ = SQLQ & " FROM " & strTableName
        '& ", HRTABL "
        SQLQ = SQLQ & " WHERE AD_EMPNBR =" & xEmpnbr
        SQLQ = SQLQ & " AND AD_DOA>=" & Date_SQL(DATE1)
        SQLQ = SQLQ & " AND AD_DOA<=" & Date_SQL(DATE2)
        
        'Ticket #30218 - Filter out the place holder/marker (AD_HRS = 9999) used to identify dates which has one or more attendance records.
        'These are not the hours entered by the user
        SQLQ = SQLQ & " AND AD_HRS <> 9999"
        
        'SQLQ = SQLQ & " AND (AD_REASON = TB_KEY) AND (TB_NAME = 'ADRE')"
        
'        If strTableName = "HR_ATTENDANCE" Then
'            SQLQ = SQLQ & " UNION "
'            SQLQ = SQLQ & " SELECT AH_COMPNO,'" & glbUserID & "' AS AH_WRKEMP,AH_EMPNBR,AH_DOA,AH_HRS,AH_REASON,  AH_SHIFT "
'            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY, HRTABL "
'            SQLQ = SQLQ & " WHERE AH_EMPNBR =" & rsTemp.Fields("ED_EMPNBR")
'            SQLQ = SQLQ & " AND AH_DOA>=" & Date_SQL(DATE1)
'            SQLQ = SQLQ & " AND AH_DOA<=" & Date_SQL(DATE2)
'            SQLQ = SQLQ & " AND (AH_REASON = TB_KEY) AND (TB_NAME = 'ADRE')"
'        End If

        SQLQ = SQLQ & " ORDER BY AD_DOA"
        
        If glbSQL Or glbOracle Then
            rsAT.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockReadOnly
        Else
            If strTableName = "TERM_TIMESHEET" Then
                rsAT.Open SQLQ, gdbESS, adOpenKeyset, adLockReadOnly
            Else
                rsAT.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockReadOnly
            End If
        End If
        '        xxx = rsAT.RecordCount
        '        xx1 = 0
        Do Until rsAT.EOF
'            xx1 = xx1 + 1
'            MDIMain.panHelp(0).FloodPercent = (xx1 / xxx) * 60 + 30
            xDay = DateDiff("d", DATE1, rsAT!AD_DOA) + 1
            xField = "AD_DAY" & xDay
            
            SQLQ = "SELECT * FROM HR_ATT_TIMESHEET "
            SQLQ = SQLQ & " WHERE AD_EMPNBR=" & xEmpnbr
            SQLQ = SQLQ & " AND AD_WRKEMP='" & glbUserID & "'"
            SQLQ = SQLQ & " AND AD_REASON ='" & rsAT!AD_REASON & "'"
            
            If glbSQL Or glbOracle Then
                rsATTCal.Open SQLQ, gdbAdoIhr001X, adOpenStatic, adLockOptimistic
            Else
                rsATTCal.Open SQLQ, gdbAdoIhr001W, adOpenStatic, adLockOptimistic
            End If
            
            xHours = 0
            If rsATTCal.EOF Then
                rsATTCal.AddNew
                rsATTCal!AD_EMPNBR = xEmpnbr
                rsATTCal!AD_COMPNO = "001"
                rsATTCal!AD_WRKEMP = glbUserID
                rsATTCal!AD_DOA = DATE1
                rsATTCal!AD_REASON = rsAT!AD_REASON
            End If
            rsATTCal!AD_STATUS = strStatus
            If IsNull(rsATTCal(xField)) Then
                xHours = 0
            Else
                xHours = rsATTCal(xField)
            End If
            rsATTCal(xField) = xHours + rsAT!AD_HRS
            rsATTCal.Update
            rsATTCal.Close
            rsAT.MoveNext
        Loop
        rsAT.Close
Loopend:
        rsTemp.MoveNext
    Loop
    rsTemp.Close

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Exit Sub
    
Err_XLS_T:
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT

    If Err = 1004 Then
        Resume Next
    End If

    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Terminated Timesheet", "", "Select")
Resume Next
End Sub

Private Function getStatus(xEmpnbr, strStartDate, strEndDate)
Dim SQLQ, statusFlag
Dim rsDS As New ADODB.Recordset
Dim gdbESS As New ADODB.Connection

    If glbSQL Or glbOracle Then
        Set gdbESS = gdbAdoIhr001
    Else
        gdbESS.Open Replace(glbAdoIHRDB, "IHR001", "IHRESS")
    End If
    
    On Error Resume Next
    'Ticket #28002 - Opening for all clients and adding Termination employee option as well
    If optAT(0) <> 0 Then
        SQLQ = "SELECT DISTINCT AD_APPROVED,AD_UPLOAD FROM HR_TIMESHEET "
    Else
        SQLQ = "SELECT DISTINCT AD_APPROVED,AD_UPLOAD FROM TERM_TIMESHEET "
    End If
    SQLQ = SQLQ & " WHERE AD_EMPNBR =" & xEmpnbr
    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(strStartDate)
    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(strEndDate)

    If glbSQL Or glbOracle Then
        'Ticket #28002 - Opening for all clients and adding Termination employee option as well
        If optAT(0) <> 0 Then
            rsDS.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        Else
            rsDS.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly
        End If
    Else
        rsDS.Open SQLQ, gdbESS, adOpenForwardOnly
    End If
    
    getStatus = ""
    statusFlag = True
    Do While Not rsDS.EOF
        If statusFlag Then
            If IsNull(rsDS("AD_APPROVED")) Then
                If rsDS("AD_UPLOAD") & "" = "Y" Then
                    getStatus = "SUBMITTED"
                Else
                    getStatus = "SAVED"
                End If
            Else
                getStatus = rsDS("AD_APPROVED")
                If getStatus = "RESUBMIT" Then getStatus = "RESUBMITTED"
            End If
            statusFlag = False
        Else
            'getStatus="Inconsistent"
            getStatus = "SAVED"
        End If
        rsDS.MoveNext

    Loop
    rsDS.Close
    If getStatus = "" Then
        'Ticket #28002 - Opening for all clients and adding Termination employee option as well
        If optAT(0) <> 0 Then
            SQLQ = "SELECT DISTINCT AD_EMPNBR,AD_UPLOAD FROM HR_ATTENDANCE "
        Else
            SQLQ = "SELECT DISTINCT AD_EMPNBR,AD_UPLOAD FROM TERM_ATTENDANCE "
        End If
        SQLQ = SQLQ & " WHERE AD_EMPNBR =" & xEmpnbr
        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(strStartDate)
        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(strEndDate)
    
        'Ticket #28002 - Opening for all clients and adding Termination employee option as well
        If optAT(0) <> 0 Then
            rsDS.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        Else
            rsDS.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly
        End If
        
        statusFlag = True
        If Not rsDS.EOF Then
            getStatus = "APPROVED"
        End If
        rsDS.Close
    
    End If
    Set rsDS = Nothing
    If Err.Number <> 0 Then
    End If
End Function

Function getCompTimeBank(xEmpnbr)
    Dim SQLQ, xOTEarned, xCTTaken
    Dim rsAttOT As New ADODB.Recordset
    Dim rsAttCT As New ADODB.Recordset
    SQLQ = ""
    On Error Resume Next
    getCompTimeBank = 0
    
    'EARNED - OT
    SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS AD_HRSTOTAL FROM HR_ATTENDANCE "
    SQLQ = SQLQ & " WHERE AD_EMPNBR =" & xEmpnbr
    
    If glbOracle Then
        SQLQ = SQLQ & " AND SUBSTR(AD_REASON,1,2) = 'OT'"
    Else
        SQLQ = SQLQ & " AND LEFT(AD_REASON,2) = 'OT'"
    End If
    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
    rsAttOT.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    
    xOTEarned = 0
    If rsAttOT.EOF Then
        xOTEarned = 0
    Else
        If Not IsNull(rsAttOT("AD_HRSTOTAL")) Then xOTEarned = CDbl(rsAttOT("AD_HRSTOTAL"))
    End If
    rsAttOT.Close
    
    'TAKEN - CT
    SQLQ = ""
    SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS AD_HRSTOTAL FROM HR_ATTENDANCE "
    SQLQ = SQLQ & " WHERE AD_EMPNBR =" & xEmpnbr
    
    If glbOracle Then
        SQLQ = SQLQ & " AND SUBSTR(AD_REASON,1,2) = 'CT'"
    Else
        SQLQ = SQLQ & " AND LEFT(AD_REASON,2) = 'CT'"
    End If
    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
    rsAttCT.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    
    xCTTaken = 0
    If rsAttCT.EOF Then
        xCTTaken = 0
    Else
        If Not IsNull(rsAttCT("AD_HRSTOTAL")) Then xCTTaken = CDbl(rsAttCT("AD_HRSTOTAL"))
    End If
    rsAttCT.Close
    
    getCompTimeBank = FormatNumber((xOTEarned - xCTTaken), 2)
    
    If Err.Number <> 0 Then
    End If
End Function

Function getVacOutstDay(xEmpnbr)
    Dim SQLQ, xOuts
    Dim rsEmp As New ADODB.Recordset
    
    SQLQ = ""
    
    On Error Resume Next
    
    getVacOutstDay = 0
    
    SQLQ = "SELECT ED_EFDATE,ED_ETDATE, ED_VAC,ED_PVAC,ED_VACT,ED_DHRS FROM HREMP "
    SQLQ = SQLQ & " WHERE ED_EMPNBR =" & xEmpnbr
    
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

    If rsEmp.EOF Then Exit Function

    xOuts = 0
    If Not IsNull(rsEmp("ED_VAC")) Then xOuts = xOuts + CDbl(rsEmp("ED_VAC"))
    If Not IsNull(rsEmp("ED_PVAC")) Then xOuts = xOuts + CDbl(rsEmp("ED_PVAC"))
    If Not IsNull(rsEmp("ED_VACT")) Then xOuts = xOuts - CDbl(rsEmp("ED_VACT"))
    If Not IsNull(rsEmp("ED_DHRS")) Then
        xOuts = FormatNumber(xOuts / CDbl(rsEmp("ED_DHRS")), 2)
    Else
        xOuts = 0
    End If
    getVacOutstDay = xOuts
    
    rsEmp.Close
    If Err.Number <> 0 Then
    End If
End Function

Function getSickTaken(xEmpnbr)
    Dim SQLQ, xTaken
    Dim rsEmp As New ADODB.Recordset
    
    SQLQ = ""
    
    On Error Resume Next
    
    getSickTaken = 0
    SQLQ = "SELECT ED_EFDATES,ED_ETDATES,ED_SICK,ED_PSICK,ED_SICKT,ED_DHRS FROM HREMP "
    SQLQ = SQLQ & " WHERE ED_EMPNBR =" & xEmpnbr
    SQLQ = SQLQ & " AND ED_EFDATES <=" & Date_SQL(dlpDateRange(1))
    SQLQ = SQLQ & " AND ED_ETDATES >=" & Date_SQL(dlpDateRange(0))
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

    If rsEmp.EOF Then Exit Function

    xTaken = 0
    If Not IsNull(rsEmp("ED_PSICK")) Then xTaken = xTaken + CDbl(rsEmp("ED_PSICK"))
    If Not IsNull(rsEmp("ED_SICK")) Then xTaken = xTaken + CDbl(rsEmp("ED_SICK"))
    If Not IsNull(rsEmp("ED_SICKT")) Then xTaken = xTaken - CDbl(rsEmp("ED_SICKT"))
    If Not IsNull(rsEmp("ED_DHRS")) Then
        xTaken = FormatNumber(xTaken / CDbl(rsEmp("ED_DHRS")), 2)
    Else
        xTaken = 0
    End If
    getSickTaken = xTaken
        
    rsEmp.Close
    
    If Err.Number <> 0 Then
    End If
End Function

Function getHrsTaken(xEmpnbr)
    Dim SQLQ2, SQLQ1
    Dim rsHour As New ADODB.Recordset
    Dim rsTBL As New ADODB.Recordset
    Dim xNo, xHDesc, xTaken
    
    xNo = 0
    
    On Error Resume Next
    
    SQLQ2 = ""
    SQLQ2 = "SELECT * from HRENTHRS "
    SQLQ2 = SQLQ2 & " WHERE HE_EMPNBR = " & xEmpnbr
    SQLQ2 = SQLQ2 & " AND HE_FDATE <=" & Date_SQL(dlpDateRange(1))
    SQLQ2 = SQLQ2 & " AND HE_TDATE >=" & Date_SQL(dlpDateRange(0))
    SQLQ2 = SQLQ2 & " ORDER BY HE_FDATE DESC"
    
    rsHour.Open SQLQ2, gdbAdoIhr001, , adOpenForwardOnly

    If rsHour.EOF Then Exit Function

    rsHour.MoveFirst
    
    getHrsTaken = ""
    Do While Not rsHour.EOF
        If Not (glbCompSerial = "S/N - 2257W" And (Left(rsHour("HE_TYPE"), 2) = "BT" Or Left(rsHour("HE_TYPE"), 3) = "MA2")) Then 'not Hamilton CCAS And Reason Code not is BT or MA2
            SQLQ1 = ""
            SQLQ1 = "SELECT TB_KEY,TB_DESC FROM HRTABL "
            SQLQ1 = SQLQ1 & " WHERE (TB_NAME='ADRE') "
            SQLQ1 = SQLQ1 & " AND (TB_KEY='" & rsHour("HE_TYPE") & "')"
            rsTBL.Open SQLQ1, gdbAdoIhr001, , adOpenForwardOnly

            If rsTBL.EOF Then Exit Function

            xHDesc = ""
            If Not IsNull(rsTBL("TB_DESC")) Then xHDesc = xHDesc & rsTBL("TB_DESC") & " Remaining"
            getHrsTaken = getHrsTaken & xHDesc & "|"
    
            xTaken = 0
            If Not IsNull(rsHour("HE_ENTITLE")) Then xTaken = xTaken + CDbl(rsHour("HE_ENTITLE"))
            If Not IsNull(rsHour("HE_TAKEN")) Then xTaken = xTaken - CDbl(rsHour("HE_TAKEN"))
            getHrsTaken = getHrsTaken & FormatNumber(xTaken, 2) & "|"

            rsTBL.Close
        
            xNo = xNo + 1
        End If
        rsHour.MoveNext
    Loop
        
    rsHour.Close

    getHrsTaken = xNo & "|" & getHrsTaken & "|"
    
End Function

