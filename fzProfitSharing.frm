VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRProfitSharing 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Profit Sharing Report"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   1395
   ClientWidth     =   10095
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
   ScaleHeight     =   8565
   ScaleWidth      =   10095
   WindowState     =   2  'Maximized
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
      Top             =   7110
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
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Tag             =   "Second level of grouping records"
      Top             =   7425
      Width           =   2325
   End
   Begin VB.Frame fraRptGroup 
      BorderStyle     =   0  'None
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
      Left            =   2100
      TabIndex        =   42
      Top             =   5880
      Width           =   5175
      Begin Threed.SSOption optRptGrp 
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   43
         Tag             =   "Choose Between Detailed or Summary Report"
         Top             =   0
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "   Active Employees"
         ForeColor       =   16711680
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
      Begin Threed.SSOption optRptGrp 
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   44
         Tag             =   "Choose Between Detailed or Summary Report"
         Top             =   0
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "   Terminated Employees"
         ForeColor       =   16711680
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
      Begin Threed.SSOption optRptGrp 
         Height          =   195
         Index           =   2
         Left            =   4080
         TabIndex        =   45
         Tag             =   "Choose Between Detailed or Summary Report"
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "   Both"
         ForeColor       =   16711680
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
   End
   Begin VB.Frame frmAT 
      Caption         =   "Employee Lookup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2130
      TabIndex        =   38
      Top             =   5160
      Width           =   2535
      Begin VB.OptionButton optActTerm 
         Caption         =   "Terminated"
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
         Index           =   1
         Left            =   1080
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optActTerm 
         Caption         =   "Active"
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
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.ComboBox cmdPSType 
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
      ItemData        =   "fzProfitSharing.frx":0000
      Left            =   9600
      List            =   "fzProfitSharing.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   9240
      Top             =   7440
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
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Tag             =   "10-Enter Employee Number"
      Top             =   4830
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7435
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1650
      Width           =   7755
      _ExtentX        =   13679
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
      Top             =   1980
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDPT"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1320
      Width           =   7755
      _ExtentX        =   13679
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
      Top             =   990
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
      Top             =   660
      Width           =   7755
      _ExtentX        =   13679
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
      Top             =   330
      Width           =   7755
      _ExtentX        =   13679
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
      Index           =   1
      Left            =   1800
      TabIndex        =   6
      Tag             =   "00-Enter Administered By Code"
      Top             =   2310
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   7
      Tag             =   "00-Enter Section Code"
      Top             =   2650
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   8
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
      Left            =   3840
      TabIndex        =   14
      Tag             =   "40-Date upto and including this date forward"
      Top             =   4440
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
      TabIndex        =   13
      Tag             =   "40-Date from and including this date forward"
      Top             =   4440
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin Threed.SSOption optGrouping 
      Height          =   195
      Index           =   1
      Left            =   3900
      TabIndex        =   18
      Tag             =   "Choose Between Detailed or Summary Report"
      Top             =   6240
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "   Summary"
      ForeColor       =   16711680
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
   Begin Threed.SSOption optGrouping 
      Height          =   195
      Index           =   0
      Left            =   2100
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "Choose Between Detailed or Summary Report"
      Top             =   6240
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "   Detailed"
      ForeColor       =   16711680
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
      Index           =   6
      Left            =   1800
      TabIndex        =   10
      Tag             =   "00-Enter TYPE"
      Top             =   3720
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "PSTY"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   7
      Left            =   1800
      TabIndex        =   9
      Tag             =   "00-Enter Physical Branch Code"
      Top             =   3360
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SUDE"
   End
   Begin INFOHR_Controls.DateLookup dlpSenDateRange 
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   12
      Tag             =   "40-Date upto and including this date forward"
      Top             =   4080
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpSenDateRange 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   11
      Tag             =   "40-Date from and including this date forward"
      Top             =   4080
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
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
      Left            =   3480
      TabIndex        =   51
      Top             =   4125
      Width           =   195
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
      TabIndex        =   50
      Top             =   4125
      Width           =   1560
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
      Left            =   120
      TabIndex        =   49
      Top             =   3420
      Width           =   1140
   End
   Begin VB.Label lblNote 
      Caption         =   "Note:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   48
      Top             =   7920
      Width           =   615
   End
   Begin VB.Label lblNote 
      Caption         =   "Purple selection criteria items uses the values in the Profit Sharing table and not the Employee Master Table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   615
      Index           =   0
      Left            =   1800
      TabIndex        =   47
      Top             =   7920
      Width           =   5175
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Include Employees"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   46
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   120
      TabIndex        =   41
      Top             =   6240
      Width           =   1065
   End
   Begin VB.Label lblCostOfEmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Profit Sharing Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Top             =   3780
      Width           =   1485
   End
   Begin VB.Label lblSelectCrit 
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
      Width           =   1695
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
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   375
      Width           =   1395
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
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   705
      Width           =   825
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Top             =   1365
      Width           =   420
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   32
      Top             =   1695
      Width           =   450
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
      Top             =   4875
      Width           =   1290
   End
   Begin VB.Label lblRenewal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Top             =   4485
      Width           =   1095
   End
   Begin VB.Label lblReportGrp 
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
      TabIndex        =   29
      Top             =   6780
      Width           =   1575
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
      TabIndex        =   28
      Top             =   7140
      Width           =   885
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
      TabIndex        =   27
      Top             =   7455
      Width           =   885
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Left            =   120
      TabIndex        =   26
      Top             =   1035
      Width           =   615
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
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   2355
      Width           =   1485
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
      TabIndex        =   24
      Top             =   3055
      Width           =   510
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
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   2695
      Width           =   540
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
      TabIndex        =   22
      Top             =   2025
      Width           =   1455
   End
   Begin VB.Label lblFormalDate 
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
      Left            =   3480
      TabIndex        =   21
      Top             =   4485
      Width           =   195
   End
End
Attribute VB_Name = "frmRProfitSharing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EECri As String, OneSet%, X%

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

  If CriCheck() Then

    Call set_PrintState(False)
    
    If glbFormCaption = "Profit Sharing Report" Then
        X% = Cri_SetAll()
    End If
    
    If glbFormCaption = "Red Circled Report" Then
        X% = Cri_SetRedCircled
    End If
    
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    
    MDIMain.Timer1.Enabled = True
    
    Call set_PrintState(True)
  End If
'End If
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
Dim X%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Call set_PrintState(False)
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    'X% = Cri_SetAll()
    If glbFormCaption = "Profit Sharing Report" Then
        X% = Cri_SetAll()
    End If
    
    If glbFormCaption = "Red Circled Report" Then
        X% = Cri_SetRedCircled
    End If
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString  'laura nov 21, 1997

'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ATTEND", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Private Sub comGroup_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    cmdPSType.Clear
    cmdPSType.AddItem "Quarterly"
    cmdPSType.AddItem "Annual"
    cmdPSType.AddItem "Both"
    cmdPSType.ListIndex = 2
    
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem "Employee Name"
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    comGroup(0).AddItem lStr("Administered By")
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    comGroup(1).AddItem "Employee Name"
    comGroup(0).ListIndex = 0 '5
    comGroup(1).ListIndex = 0
    comGroup(1).Enabled = False
End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    Select Case intIdx%
        Case 0: strCd$ = "HREMP.ED_LOC"
        Case 1: strCd$ = "HR_PROFIT_SHARING.PS_ADMINBY"
        Case 2: strCd$ = "HREMP.ED_ORG"
        Case 3: strCd$ = "HR_PROFIT_SHARING.PS_EMP"
        Case 4: strCd$ = "HR_PROFIT_SHARING.PS_SECTION"
        Case 5: strCd$ = "HREMP.ED_REGION"
        Case 6: strCd$ = "HR_PROFIT_SHARING.PS_PCODE"
        Case 7: strCd$ = "HREMP.ED_SUBDEPT" 'Ticket #24162 Franks 09/19/2013
    End Select
End If

If Len(clpCode(intIdx%).Text) > 0 Then
    CodeCri = "({" & strCd$ & "} in ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
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
Private Sub Cri_Dept()
Dim DeptCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level


If Len(clpDept.Text) > 0 Then
    DeptCri = "({HR_PROFIT_SHARING.PS_DEPTNO} in ['" & Replace(clpDept.Text, ",", "','") & "'])"
End If

If Len(DeptCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = DeptCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & DeptCri
    End If
    glbiOneWhere = True
End If

End Sub
Private Sub Cri_PSType()

Dim TypeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level

'If cmdPSType.Text = "Both" Then Exit Sub
'
'If cmdPSType.Text = "Annual" Then
'    TypeCri = "({HR_PROFIT_SHARING.PS_TYPE}) = 'A' "
'End If
'If cmdPSType.Text = "Quarterly" Then
'    TypeCri = "({HR_PROFIT_SHARING.PS_TYPE}) = 'Q' "
'End If
'If Len(TypeCri) >= 1 Then
'    If Not glbiOneWhere Then
'        glbstrSelCri = TypeCri
'    Else
'        glbstrSelCri = glbstrSelCri & " AND " & TypeCri
'    End If
'    glbiOneWhere = True
'End If


End Sub
Private Sub Cri_Div()

Dim DivCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level


If Len(clpDiv.Text) > 0 Then
    DivCri = "({HR_PROFIT_SHARING.PS_DIV} in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
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

Private Sub Cri_NONEandEXEC() 'Frank Ticket# 6795 Missed Security of -NON and -EXE
Dim EECri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level

If Not (glbNoNONE Or glbNoEXEC) Then
    Exit Sub
End If
If glbNoNONE And glbNoEXEC Then
    EECri = "NOT ({HREMP.ED_ORG} = 'NONE' OR {HREMP.ED_ORG} = 'EXEC') "
ElseIf glbNoNONE Then
    EECri = "NOT ({HREMP.ED_ORG} = 'NONE') "
ElseIf glbNoEXEC Then
    EECri = "NOT ({HREMP.ED_ORG} = 'EXEC') "
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub
Private Sub Cri_ActTerm()
Dim EECri As String
Dim xEmpGroup As Integer '0 - term; 1 - active ; 2- both

xEmpGroup = 1 'default
If optActTerm(1).Value And Len(elpEEID.Text) > 0 Then
    xEmpGroup = 0
Else
    'If chkTerm.Value Then
    '    xEmpGroup = 2
    'End If
    If optRptGrp(0).Value Then xEmpGroup = 1
    If optRptGrp(1).Value Then xEmpGroup = 0
    If optRptGrp(2).Value Then xEmpGroup = 2
End If

If xEmpGroup = 1 Then
    EECri = " left({HR_PROFIT_SHARING.KEY_EMPNBR},1) = '1' "
End If
If xEmpGroup = 0 Then
    EECri = " left({HR_PROFIT_SHARING.KEY_EMPNBR},1) = '0' "
End If
If xEmpGroup = 2 Then
    EECri = " (1=1) "
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

Private Sub Cri_FTDates()
Dim TempCri As String
Dim TempCri2 As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%

If Len(dlpDateRange(0).Text) = 0 And Len(dlpDateRange(1).Text) = 0 Then
    Exit Sub
End If

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HR_PROFIT_SHARING.PS_FDATE}) "
    TempCri2 = "({HR_PROFIT_SHARING.PS_TDATE}) "

    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    'TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri2 = TempCri2 & " <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    TempCri = TempCri & " AND " & TempCri2
    GoTo Cri_FTDatst
End If

For X% = 0 To 1
    TempCri = "({HR_PROFIT_SHARING.PS_FDATE}) "
    TempCri2 = "({HR_PROFIT_SHARING.PS_TDATE}) "
        
    If Len(dlpDateRange(0).Text) > 0 Then
        TempCri = TempCri & " >= "
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
    End If
    If Len(dlpDateRange(1).Text) > 0 Then
        TempCri = TempCri2 & " <= "
        dtYYY% = Year(dlpDateRange(1).Text)
        dtMM% = month(dlpDateRange(1).Text)
        dtDD% = Day(dlpDateRange(1).Text)
    End If

    TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    GoTo Cri_FTDatst

Next X%



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

Private Function Cri_SetRedCircled()
Dim X%, strRName$
Cri_SetRedCircled = False
On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = " (1=1) "

Call glbCri_DeptUN(clpDept.Text)
'Call Cri_Dept
Call Cri_AllDiv
Call Cri_AllPT
Call Cri_AllEE

For X% = 0 To 5
    Call Cri_AllCode(X%)
Next X%

Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZSalRedCir.rpt"

' set to sorting/grouping criteria
X% = Cri_AllSorts()   ' returns number of sections formated
'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If
Me.vbxCrystal.Formulas(10) = "lblSection = '" & lStr("Section") & "'"
Me.vbxCrystal.Formulas(11) = "lblAdminBy = '" & lStr("Administered By") & "'"

Me.vbxCrystal.WindowTitle = "Red Circled Report"
Me.vbxCrystal.Connect = RptODBC_SQL
    
Cri_SetRedCircled = True

Screen.MousePointer = DEFAULT
Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Assoc ", "Assoc Report", "Select")
Cri_SetRedCircled = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function
Private Function Cri_SetAll()
Dim X%, strRName$
Cri_SetAll = False
On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = " (1=1) "

Call glbCri_DeptUN(clpDept.Text) 'Ticket #21202 Franks 11/16/2011
Call Cri_Dept
Call Cri_Div
Call Cri_PT
Call Cri_EE

Call Cri_SenDates 'Ticket #24162 Franks 09/19/2013
Call Cri_FTDates

For X% = 0 To 7 ' 6 '5
    Call Cri_Code(X%)
Next X%
'Call Cri_PSType
Call Cri_ActTerm

Call Cri_NONEandEXEC 'Ticket #22453 Franks 08/24/2012 add Security of -NON and -EXE

' report name
'If comGroup(0) <> "(none)" Then
'    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzassoc.rpt"
'Else
If optGrouping(0).Value Then 'details
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZProfitsD.rpt"
Else 'summary
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZProfitsS.rpt"
End If
' set to sorting/grouping criteria
X% = Cri_Sorts()   ' returns number of sections formated
'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

If optGrouping(0).Value Then 'details
    Me.vbxCrystal.WindowTitle = "Profit Sharing Details Report"
Else
    Me.vbxCrystal.WindowTitle = "Profit Sharing Summary Report"
End If
Me.vbxCrystal.Connect = RptODBC_SQL
    
Cri_SetAll = True

Screen.MousePointer = DEFAULT
Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Assoc ", "Assoc Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function Cri_AllSorts()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

Cri_AllSorts = 0

If comGroup(0).Text = "(none)" Then
    grpField$ = "{HREMP.ED_COMPNO}"
Else
    grpField$ = getEGroup(comGroup(0).Text)
End If

dscGroup$ = comGroup(0).Text
dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(0) = dscGroup$
grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(0) = grpCond$
strSFormat$ = "GH1;T;T;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1
strSFormat$ = "GF1;T;X;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1
    

Cri_AllSorts = z% ' next section number to format

End Function
Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

Cri_Sorts = 0

If comGroup(0).Text = "(none)" Then
    grpField$ = "{HR_PROFIT_SHARING.PS_COMPNO}"
Else
    grpField$ = getEGroup(comGroup(0).Text)
End If

dscGroup$ = comGroup(0).Text
dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(0) = dscGroup$
grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(0) = grpCond$
strSFormat$ = "GH1;T;T;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1
strSFormat$ = "GF1;T;X;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1
    

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
    Exit Function
End If
If Not clpPT.ListChecker Then
    Exit Function
End If
For X% = 0 To 5
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

For X% = 0 To 1
    If Len(dlpDateRange(X%).Text) > 0 Then
       If Not IsDate(dlpDateRange(X%).Text) Then
           MsgBox "Not a valid date"
           dlpDateRange(X%).Text = ""
           dlpDateRange(X%).SetFocus
           Exit Function
       End If
    End If
Next X%

If Len(dlpDateRange(0)) > 0 And Len(dlpDateRange(1)) > 0 Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then
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

Me.Caption = glbFormCaption

glbOnTop = "frmRProfitSharing"

Screen.MousePointer = HOURGLASS

If glbFormCaption = "Red Circled Report" Then
    Call SetScreen4RedCircled
End If

Call comGrpLoad

Call INI_Controls(Me)
Call setRptCaption(Me)
lblSeniority.Caption = lStr("Seniority")

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

Private Sub optActTerm_Click(Index As Integer)
    If Index = 0 Then
        elpEEID.LookupType = 0 'ACTIVE
    End If
    If Index = 1 Then
        elpEEID.LookupType = 1 'TERM
    End If
End Sub

Private Sub SetScreen4RedCircled()
    lblDiv.ForeColor = BLACK
    lblDept.ForeColor = BLACK
    lblStatus.ForeColor = BLACK
    lblAdmin.ForeColor = BLACK
    lblSection.ForeColor = BLACK
    frmAT.Visible = False
    lblEENum(1).Visible = False
    optRptGrp(0).Visible = False
    optRptGrp(1).Visible = False
    optRptGrp(2).Visible = False
    lblTitle(8).Visible = False
    optGrouping(0).Visible = False
    optGrouping(1).Visible = False
    lblNote(0).Visible = False
    lblNote(1).Visible = False
    lblCostOfEmp.Visible = False
    cmdPSType.Visible = False
    lblRenewal.Visible = False
    dlpDateRange(0).Visible = False
    dlpDateRange(1).Visible = False
    lblFormalDate.Visible = False

    'location
    lblEENum(0).Top = lblCostOfEmp.Top
    elpEEID.Top = cmdPSType.Top
    lblReportGrp.Top = 4440
    lblGrp(0).Top = lblReportGrp.Top + 360
    lblGrp(1).Top = lblReportGrp.Top + 360 + 315
    comGroup(0).Top = lblReportGrp.Top + 330
    comGroup(1).Top = lblReportGrp.Top + 360 + 285

End Sub

Private Sub Cri_AllDiv()
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
Private Sub Cri_AllPT()
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


Private Sub Cri_AllCode(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    Select Case intIdx%
    Case 0: strCd$ = "HREMP.ED_LOC"
    Case 1: strCd$ = "HREMP.ED_ADMINBY"
    Case 2: strCd$ = "HREMP.ED_ORG"
    Case 3: strCd$ = "HREMP.ED_EMP"
    Case 4: strCd$ = "HREMP.ED_SECTION"
    Case 5: strCd$ = "HREMP.ED_REGION"
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

Private Sub Cri_AllEE()
Dim EECri As String

If Len(getEmpnbr(elpEEID.Text)) > 0 Then
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
