VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRCsRpt 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Custom Report"
   ClientHeight    =   10230
   ClientLeft      =   180
   ClientTop       =   825
   ClientWidth     =   11280
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10230
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkLasg2PrvMonths 
      Caption         =   "Bring over the last 2 previous months"
      Height          =   252
      Left            =   5520
      TabIndex        =   83
      Top             =   7156
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CheckBox chkInact 
      Caption         =   "Inactive"
      Height          =   252
      Left            =   9850
      TabIndex        =   6
      Top             =   1666
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   890
   End
   Begin VB.CheckBox chkAct 
      Caption         =   "Active"
      Height          =   252
      Left            =   8880
      TabIndex        =   5
      Top             =   1666
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   2115
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Tag             =   "Third level of grouping records"
      Top             =   9120
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.TextBox txtUserText1 
      Appearance      =   0  'Flat
      DataSource      =   " "
      Height          =   280
      Left            =   2115
      MaxLength       =   20
      TabIndex        =   79
      Tag             =   "00-User Text 1"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox txtUserText2 
      Appearance      =   0  'Flat
      DataSource      =   " "
      Height          =   285
      Left            =   2115
      MaxLength       =   20
      TabIndex        =   78
      Tag             =   "00-User Text 2"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1620
   End
   Begin Threed.SSFrame frCasualVac 
      Height          =   855
      Left            =   120
      TabIndex        =   73
      Top             =   9720
      Visible         =   0   'False
      Width           =   8655
      _Version        =   65536
      _ExtentX        =   15266
      _ExtentY        =   1508
      _StockProps     =   14
      Caption         =   "CASUAL or VACATION"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.OptionButton optVac 
         Caption         =   "Outstanding Vacation"
         Height          =   255
         Left            =   5400
         TabIndex        =   77
         Tag             =   "You can choose one of the options"
         Top             =   390
         Width           =   2175
      End
      Begin VB.ComboBox cmbVac 
         Height          =   315
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   360
         Width           =   690
      End
      Begin VB.OptionButton optCas 
         Caption         =   "Outstanding Casual"
         Height          =   375
         Left            =   240
         TabIndex        =   75
         Tag             =   "You can choose one of the options"
         Top             =   330
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.ComboBox cmbCas 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   360
         Width           =   690
      End
   End
   Begin VB.CheckBox chkDoNoLaunch 
      Caption         =   "Do not open the Report"
      Height          =   375
      Left            =   9240
      TabIndex        =   69
      Top             =   5040
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.CheckBox chkInclAtt 
      Caption         =   "Include Attendance History"
      Height          =   375
      Left            =   5160
      TabIndex        =   67
      Top             =   7200
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Frame fraMemo 
      Height          =   735
      Left            =   9240
      TabIndex        =   65
      Top             =   8880
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblFraMeno 
         Caption         =   "Only Plant Can Be Selected For This Report"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame frmActTerm 
      Height          =   615
      Left            =   6000
      TabIndex        =   62
      Top             =   3600
      Visible         =   0   'False
      Width           =   2775
      Begin VB.OptionButton OptAct 
         Caption         =   "Active"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptTerm 
         Caption         =   "Terminated"
         Height          =   255
         Left            =   1320
         TabIndex        =   63
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CheckBox chkExpired 
      Caption         =   "Only Expired"
      Height          =   252
      Left            =   5880
      TabIndex        =   59
      Top             =   8880
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.ComboBox comDateMonth 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Tag             =   "01-Month"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame frmBeneBilling 
      Caption         =   "Benefit Billing Period"
      Height          =   735
      Left            =   5160
      TabIndex        =   46
      Top             =   6360
      Visible         =   0   'False
      Width           =   5295
      Begin VB.ComboBox ComMTH 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Tag             =   "01-Month"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtFiscal 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   21
         Tag             =   "01-Fiscal Year"
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   2880
         TabIndex        =   48
         Top             =   300
         Width           =   690
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   1320
         TabIndex        =   47
         Top             =   300
         Width           =   330
      End
   End
   Begin VB.ComboBox comCountryOfEmp 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2115
      TabIndex        =   13
      Tag             =   "00-Country of Employment"
      Top             =   3960
      Width           =   1440
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2120
      MaxLength       =   4
      TabIndex        =   12
      Tag             =   "00-Employee Position Shift"
      Top             =   3630
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Frame frmDate 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   5880
      Visible         =   0   'False
      Width           =   6075
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   1
         Left            =   3510
         TabIndex        =   20
         Tag             =   "40-Date upto and including this date forward"
         Top             =   0
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   19
         Tag             =   "40-Date from and including this date forward"
         Top             =   0
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin VB.Label lblFromTo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "From / To Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   41
         Top             =   45
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmbReports 
      Height          =   315
      Left            =   2115
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   5400
      Width           =   4395
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1650
      Width           =   7005
      _ExtentX        =   12356
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
      TabIndex        =   7
      Tag             =   "EDPT-Category"
      Top             =   1980
      Width           =   7005
      _ExtentX        =   12356
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
      Top             =   1320
      Width           =   7000
      _ExtentX        =   12356
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
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDLC"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   660
      Width           =   7005
      _ExtentX        =   12356
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
      Width           =   7005
      _ExtentX        =   12356
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
      TabIndex        =   10
      Tag             =   "00-Enter Administered By Code"
      Top             =   2970
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
      TabIndex        =   11
      Tag             =   "00-Enter Section Code"
      Top             =   3300
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDSE"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   9
      Tag             =   "00-Enter Region Code"
      Top             =   2640
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Tag             =   "10-Enter Employee Number"
      Top             =   2310
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   6685
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   9000
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
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   1800
      TabIndex        =   14
      Top             =   4320
      Visible         =   0   'False
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   3
      Left            =   3630
      TabIndex        =   26
      Tag             =   "40-Date upto and including this date forward"
      Top             =   7140
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   2
      Left            =   1810
      TabIndex        =   25
      Tag             =   "40-Date from and including this date forward"
      Top             =   7140
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.CodeLookup clpUser 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   27
      Tag             =   "00-Enter Section Code"
      Top             =   7530
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "COD1"
   End
   Begin INFOHR_Controls.CodeLookup clpUser 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   28
      Tag             =   "00-Enter Section Code"
      Top             =   7905
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "COD2"
   End
   Begin INFOHR_Controls.CodeLookup clpUser 
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   54
      Tag             =   "00-Enter Section Code"
      Top             =   8520
      Visible         =   0   'False
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "COD1"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   55
      Tag             =   "40-Date upto and including this date forward"
      Top             =   8880
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   5
      Left            =   4080
      TabIndex        =   56
      Tag             =   "40-Date from and including this date forward"
      Top             =   8880
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.EmployeeLookup elpSupShow 
      Height          =   285
      Left            =   1800
      TabIndex        =   61
      Tag             =   "10-Employee Number of individual's supervisor"
      Top             =   9360
      Visible         =   0   'False
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   6685
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Tag             =   "00-Enter Position Code "
      Top             =   4650
      Visible         =   0   'False
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   5
      MultiSelect     =   -1  'True
   End
   Begin VB.Frame frBenRate 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   70
      Top             =   4800
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtBenRate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2000
         MaxLength       =   7
         TabIndex        =   16
         Tag             =   "10-Enter Fixed Percentage Rate"
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2640
         TabIndex        =   72
         Top             =   45
         Width           =   120
      End
      Begin VB.Label lblBenRate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   71
         Top             =   45
         Width           =   645
      End
   End
   Begin Crystal.CrystalReport vbxCrystal1 
      Left            =   9480
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
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin INFOHR_Controls.DateLookup dlpAsOf 
      Height          =   285
      Left            =   1810
      TabIndex        =   24
      Tag             =   "40-As of Date"
      Top             =   6780
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Final Sort"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   82
      Top             =   9150
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblUserText1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "User Text 1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   81
      Top             =   5205
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label lblUserText2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "User Text 2"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   80
      Top             =   5685
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label lblPosition 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   68
      Top             =   4695
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblAttSuper 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Rept. Authority 1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   60
      Top             =   9360
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label lblProvince 
      Caption         =   "Province / State: "
      Height          =   255
      Left            =   240
      TabIndex        =   58
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblLic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "License Expiry From / To Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   57
      Top             =   8925
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      Caption         =   "Code 1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   53
      Top             =   7545
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      Caption         =   "Code 2"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   52
      Top             =   7920
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label lblDates 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From / To Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   51
      Top             =   7185
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblMonth 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Month of Birth"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   50
      Top             =   6420
      Visible         =   0   'False
      Width           =   990
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
      TabIndex        =   49
      Top             =   6825
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblCountry 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Country of Employment"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   45
      Top             =   4020
      Width           =   1620
   End
   Begin VB.Label lblBCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Group"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   44
      Top             =   4365
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblShift 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   43
      Top             =   3675
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   42
      Top             =   2025
      Width           =   630
   End
   Begin VB.Label lblReports 
      AutoSize        =   -1  'True
      Caption         =   "Reports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   39
      Top             =   5460
      Width           =   915
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   38
      Top             =   2355
      Width           =   1290
   End
   Begin VB.Label lblSection 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      Height          =   195
      Left            =   120
      TabIndex        =   37
      Top             =   3345
      Width           =   540
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   2685
      Width           =   510
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   3015
      Width           =   1125
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   1035
      Width           =   615
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   33
      Top             =   1695
      Width           =   450
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   32
      Top             =   1365
      Width           =   420
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   31
      Top             =   705
      Width           =   825
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   375
      Width           =   555
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
      TabIndex        =   29
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmRCsRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbEmpTable As String
Dim rsRPT As New ADODB.Recordset
Dim fglbFileName
Dim fglbDateTable
Dim fglbDateField
Dim IsFenwick As Boolean
Dim HisSQL  As String
Dim HisSQLUWR  As String
Dim xWFCBilling As Boolean, xMonNum, xLIndex
Dim WithEvents CN001 As ADODB.Connection
Attribute CN001.VB_VarHelpID = -1
Dim WithEvents CN001W As ADODB.Connection
Attribute CN001W.VB_VarHelpID = -1
Dim rsEntRules As New ADODB.Recordset
Private snapEntitle As New ADODB.Recordset
Private fglbAsOf As String, fglbToDate As String
Dim dblEmpEntitle
Dim xUserFlag As Boolean
Dim StatusCodeCri As String
Const Excel2007 = 12

Private Sub cmbReports_Click()
Dim Sch

If glbSQL Then
    Sch = Replace(cmbReports, "'", "'+CHAR(39)+'")
Else
    Sch = Replace(cmbReports, "'", "'+CHR(39)+'")
End If

If rsRPT.State <> 0 Then rsRPT.Close
rsRPT.Open "SELECT * FROM HR_CUSTOMRPT WHERE RT_RPTNAME='" & Sch & "'", gdbAdoIhr001, adOpenForwardOnly

frmDate.Visible = False

If glbWFC Then 'Ticket #20859 Franks 01/30/2012
    xUserFlag = False
End If

If Not rsRPT.EOF Then
    If Not IsNull(rsRPT("RT_DATETABLE")) And Not IsNull(rsRPT("RT_DATEFIELD")) Then
        If Trim(rsRPT("RT_DATETABLE")) <> "" And Trim(rsRPT("RT_DATEFIELD")) <> "" Then
            frmDate.Visible = True
            If glbCompSerial = "S/N - 2369W" And InStr(1, fglbFileName, "sn2369AttBonus.rpt") > 0 Then
                lblFromTo.FontBold = True
            End If
            fglbDateTable = Trim(rsRPT("RT_DATETABLE"))
            fglbDateField = Trim(rsRPT("RT_DATEFIELD"))
        End If
    End If
    
    If InStr(rsRPT("RT_FILENAME"), ":") = 0 Then
        fglbFileName = glbIHRREPORTS & rsRPT("RT_FILENAME")
    Else
        fglbFileName = rsRPT("RT_FILENAME")
    End If
    
    'Ticket #16544
    'Collectcorp Inc. Ticket #16952
    If (glbCompSerial = "S/N - 2382W" And (InStr(1, fglbFileName, "SN2382_IDLComparison.xls") > 0) Or InStr(1, fglbFileName, "SN2382_IDLComparison_short.xls") > 0) _
        Or (glbCompSerial = "S/N - 2390W" And InStr(1, fglbFileName, "Export_EmployeeList.xls") > 0) Then
        frmActTerm.Visible = True
    Else
        frmActTerm.Visible = False
    End If
    
    'Ticket #16616
    'Ticket #19851 Franks 02/14/2011
    If glbWFC Then
        If InStr(1, UCase(fglbFileName), "RZPAYFORCE.RPT") > 0 Or InStr(1, UCase(fglbFileName), "RZMLFEXP.RPT") > 0 Then
            fraMemo.Top = 4680
            fraMemo.Left = 6600
            fraMemo.Visible = True
        Else
            fraMemo.Visible = False
        End If
    End If
    
    If rsRPT("RT_TERMINATION") Then
        fglbEmpTable = "TERM_HREMP"
        elpEEID.LookupType = 1 'TERM
    Else
        fglbEmpTable = "HREMP"
        elpEEID.LookupType = 0 'ACTIVE
    End If
    
    If glbWFC Then 'Ticket #20859 Franks 01/30/2012
        If Not IsNull(rsRPT("RT_USER_FLAG")) Then
            If rsRPT("RT_USER_FLAG") Then
                xUserFlag = True
            End If
        End If
    End If
    
    Call INI_Controls(Me)
'    cmdView.Enabled = True
'    cmdPrint.Enabled = True
Else
'    cmdView.Enabled = False
'    cmdPrint.Enabled = False
End If

If glbWFC Then
    xWFCBilling = False
    If InStr(1, fglbFileName, "Benefit_OptLife_Billing.rpt") > 0 Or InStr(1, fglbFileName, "Benefit_OptLife_Billing_sum.rpt") > 0 Or InStr(1, fglbFileName, "Benefit_cost_details.rpt") > 0 Or InStr(1, fglbFileName, "Benefit_Costs_summary.rpt") > 0 Then
        xWFCBilling = True
    End If
    If InStr(1, fglbFileName, "Annual_Manulife_Benefits.rpt") > 0 Then
        xWFCBilling = True
    End If
    frmBeneBilling.Visible = xWFCBilling
    'Ticket #19137
    If InStr(1, fglbFileName, "WFC WPS Training Tmp.xls") > 0 Then
        frmDate.Visible = True
    End If
End If

'City of Niagara Falls Ticket #27681 Franks 12/14/2015
If glbCompSerial = "S/N - 2276W" Then
    Call NiagaraFallsSickScreen
End If

'DNSSAB
'Ticket #14795
If glbCompSerial = "S/N - 2388W" And (InStr(1, fglbFileName, "SN2388_PooledBenSumm.rpt") > 0 Or InStr(1, fglbFileName, "SN12288_PooledBenDtl.rpt") > 0) Then
    frmBeneBilling.Visible = True
End If

'Collectcorp Inc. - Display appropriate selection criteria based on the Report selection.
'Ticket #14437
If glbCompSerial = "S/N - 2390W" Then
    lblAsOf.Visible = False
    dlpAsOf.Visible = False
    lblDates.Visible = False
    dlpDateRange(2).Visible = False
    dlpDateRange(3).Visible = False
    
    lblCode(0).Visible = False
    lblCode(1).Visible = False
    clpUser(0).Visible = False
    clpUser(1).Visible = False
    
    'Mostafa - XLS report
    clpUser(2).Visible = False
    lblProvince.Visible = False
    dlpDateRange(4).Visible = False
    dlpDateRange(5).Visible = False
    chkExpired.Visible = False
    lblLic.Visible = False
    
    lblMonth.Visible = False
    comDateMonth.Visible = False

    If InStr(1, fglbFileName, "SN2390_Anniversary.rpt") > 0 Then
        'Show As of Date and DOH
        lblAsOf.Visible = True
        dlpAsOf.Visible = True
        lblAsOf.FontBold = True
        dlpAsOf.Text = Date
        lblMonth.Visible = True
        lblMonth.Caption = lStr("Original Hire Date")
        comDateMonth.Visible = True
    ElseIf InStr(1, fglbFileName, "SN2390_Birthday.rpt") > 0 Then
        'Show Date of Birth
        lblMonth.Visible = True
        lblMonth.Caption = "Month of Birth"     'Ticket #23111
        comDateMonth.Visible = True
    ElseIf InStr(1, fglbFileName, "SN2390_License.xls") > 0 Then  'Mostafa - XLS report
        'Show Date of Birth
        clpUser(2).Visible = True
        lblProvince.Visible = True
        dlpDateRange(4).Visible = True
        dlpDateRange(5).Visible = True
        chkExpired.Visible = True
        lblLic.Visible = True
    ElseIf InStr(1, fglbFileName, "SN2390_LicenseAddr.rpt") > 0 Then
        'Show DOH
        lblDates.Visible = True
        lblDates.Caption = lStr("Original Hire Date")
        dlpDateRange(2).Visible = True
        dlpDateRange(3).Visible = True
    ElseIf InStr(1, fglbFileName, "SN2390_LicenseTerm.rpt") > 0 Then
        'Show Termination Date, License Prov/State, License Status
        lblDates.Visible = True
        lblDates.Caption = "Termination Date"
        dlpDateRange(2).Visible = True
        dlpDateRange(3).Visible = True
        lblCode(0).Visible = True
        lblCode(0).Caption = lStr("Code 1")
        lblCode(1).Visible = True
        lblCode(1).Caption = lStr("Code 2")
        clpUser(0).Visible = True
        clpUser(1).Visible = True
    ElseIf InStr(1, fglbFileName, "SN2390_LicenseAdditions.rpt") > 0 Then
        'Show Date Submitted, License Prov/State, License Status
        lblDates.Visible = True
        lblDates.Caption = lStr("Date 1")
        dlpDateRange(2).Visible = True
        dlpDateRange(3).Visible = True
        lblCode(0).Visible = True
        lblCode(0).Caption = lStr("Code 1")
        lblCode(1).Visible = True
        lblCode(1).Caption = lStr("Code 2")
        clpUser(0).Visible = True
        clpUser(1).Visible = True
    End If
'Ticket #17615 - WDDS
ElseIf glbCompSerial = "S/N - 2190W" And InStr(1, fglbFileName, "SeniorityList_Tmp.xls") > 0 Then
    lblBCode.Visible = False
    clpCode(6).Visible = False
    lblCountry.Visible = False
    lblAsOf.Visible = True
    lblAsOf.Top = 4020
    comCountryOfEmp.Visible = False
    dlpAsOf.Visible = True
    dlpAsOf.Top = 3960
    dlpAsOf.Text = Date
    chkInclAtt.Visible = True
    chkInclAtt.Top = 3930
    chkInclAtt.Left = 4680
Else
    lblAsOf.Visible = False
    dlpAsOf.Visible = False
    lblDates.Visible = False
    dlpDateRange(2).Visible = False
    dlpDateRange(3).Visible = False
    lblCode(0).Visible = False
    lblCode(1).Visible = False
    clpUser(0).Visible = False
    clpUser(1).Visible = False
    lblMonth.Visible = False
    comDateMonth.Visible = False
    
    'Mostafa - XLS report
    clpUser(2).Visible = False
    lblProvince.Visible = False
    dlpDateRange(4).Visible = False
    dlpDateRange(5).Visible = False
    chkExpired.Visible = False
    lblLic.Visible = False
End If

'Chapman's Ice Cream Limited - Ticket #19104
If glbCompSerial = "S/N - 2370W" Then
    If InStr(1, fglbFileName, "AverageHourTmp.xls") > 0 Then
        lblAsOf.Visible = True
        lblAsOf.Caption = "End Date"
        dlpAsOf.Visible = True
        dlpAsOf.Text = Format("12/31/" & Year(Now), "mm/dd/yyyy")
        
        'chkDoNoLaunch.Visible = True    'For testing
        
    ElseIf InStr(1, fglbFileName, "AbsenteeismTmp.xls") > 0 Then
        lblDates.Visible = True
        dlpDateRange(2).Visible = True
        dlpDateRange(3).Visible = True
        
        'chkDoNoLaunch.Visible = False   'For testing
    End If
End If

'Ticket #27081 - Macaulay Child Development Centre
'Ticket #20302 - Surrey Place Centre
If (glbCompSerial = "S/N - 2347W" And InStr(1, fglbFileName, "VacAccrualTmp.xls") > 0) Or _
    (glbCompSerial = "S/N - 2422W" And InStr(1, fglbFileName, "SN2422AccrualTmp.xls") > 0) Or _
    (glbCompSerial = "S/N - 2420W" And InStr(1, fglbFileName, "SN2420AttendUsedTmp.xls") > 0) Then
    frmDate.Visible = True
    
    'Ticket #21277 - No need of From Date
    If (glbCompSerial = "S/N - 2422W" And InStr(1, fglbFileName, "SN2422AccrualTmp.xls") > 0) Then
        lblFromTo.Caption = "Report End Date"
        dlpDateRange(1).Left = dlpDateRange(0).Left
        dlpDateRange(0).Visible = False
    End If
End If

'Ticket #22348 - United Way of Regina
If (glbCompSerial = "S/N - 2444W" And InStr(1, fglbFileName, "SN2444_DivSumAcct.rpt") > 0) Then
    frBenRate.Visible = True
Else
    frBenRate.Visible = False
End If

'Ticket #22548 - Broadcasting Corp. - Moved reports from their Custom Program
'Labels for Time Clock Cards
If glbCompSerial = "S/N - 2235W" Then
    If InStr(1, UCase(fglbFileName), "SN2235Y.RPT") > 0 Then
        If Not gSec_Rpt_Master_Attendance Then
            frCasualVac.Visible = False
            
            MsgBox "You Do Not Have Authority For This Report", , "info:HR"
            Exit Sub
        End If
        lblAsOf.Visible = True
        lblAsOf.Caption = "Pay Period Date"
        dlpAsOf.Visible = True
        lblAsOf.FontBold = True
    Else
        lblAsOf.Visible = False
        dlpAsOf.Visible = False
        lblAsOf.FontBold = False
    End If
End If

'Ticket #22548 - Broadcasting Corp. - Moved reports from their Custom Program
'Vacation / Casual Leave Application Form
If (glbCompSerial = "S/N - 2235W") And (InStr(1, UCase(fglbFileName), "SN2235AA.RPT") > 0 Or InStr(1, UCase(fglbFileName), "SN2235AB.RPT") > 0) Then
    If Not gSec_Rpt_Entitlements Then
        lblAsOf.Visible = False
        dlpAsOf.Visible = False
        lblAsOf.FontBold = False
        
        MsgBox "You Do Not Have Authority For This Report", , "info:HR"
        Exit Sub
    End If

    cmbVac.Clear
    cmbVac.AddItem "Yes"
    cmbVac.AddItem "No"
    cmbVac.ListIndex = 0
    
    cmbCas.Clear
    cmbCas.AddItem "Yes"
    cmbCas.AddItem "No"
    cmbCas.ListIndex = 0

    frCasualVac.Visible = True
    frCasualVac.Top = 6360
    frCasualVac.Left = 120
Else
    frCasualVac.Visible = False
End If

'Ticket #22876 - County of Hastings
If glbCompSerial = "S/N - 2263W" And InStr(1, fglbFileName, "SN2263ManulifeBenefitTmp.xls") > 0 Then
    lblUserText1.Visible = True
    txtUserText1.Visible = True     'ED_USER_TEXT1
    lblUserText1.Top = lblBCode.Top
    txtUserText1.Top = clpCode(6).Top
    
    lblUserText2.Visible = True
    txtUserText2.Visible = True     'ED_USER_TEXT2
    lblUserText2.Top = lblPosition.Top
    txtUserText2.Top = clpJob.Top
    
    Call setCaption(lblUserText1)
    Call setCaption(lblUserText2)
    
    lblTitle(12).Visible = True
    comGroup(0).Visible = True
    Call comGrpLoad
Else
    lblUserText1.Visible = False
    txtUserText1.Visible = False
    lblUserText2.Visible = False
    txtUserText2.Visible = False
    lblTitle(12).Visible = False
    comGroup(0).Visible = False
End If

'Ticket #24374 - Surrey Place Centre
If glbCompSerial = "S/N - 2347W" And InStr(1, fglbFileName, "SN2347_EmpDates.rpt") > 0 Then
    'Show As of Date
    lblAsOf.Visible = True
    dlpAsOf.Visible = True
    lblAsOf.FontBold = True
    dlpAsOf.Text = Date
End If

If glbSamuel Then 'Ticket #24163 Franks 12/05/2013
    If InStr(1, fglbFileName, "SN2382_Employee_Salary.xls") > 0 Then
        lblFromTo.Caption = lStr("Seniority")
    Else
        lblFromTo.Caption = "From / To Date"
    End If
End If

'Ticket #29657 - Renamed the existing Employee Fan Out to General Employee and created a new Employee Fan Out report with new format
'Ticket #27813 - WDGPHU
If glbCompSerial = "S/N - 2411W" And (InStr(1, fglbFileName, "SN2411GeneralEmpTmp.xls") > 0 Or InStr(1, fglbFileName, "SN2411EmpFanOutTmp.xls") > 0 Or _
    InStr(1, fglbFileName, "SN2411EmpTelephoneTmp.xls") > 0 Or InStr(1, fglbFileName, "SN2411LicenseTmp.xls") > 0 Or _
    InStr(1, fglbFileName, "SN2411VacationTmp.xls") > 0 Or InStr(1, fglbFileName, "SN2411PerfManagementTmp.xls") > 0) Then
    'InStr(1, fglbFileName, "SN2411TrainingTmp.xls") > 0 Then
    chkAct.Visible = True
    chkInact.Visible = True
    
    If InStr(1, fglbFileName, "SN2411PerfManagementTmp.xls") > 0 Then
        frmDate.Visible = True
        lblFromTo.Caption = "Review Date"
        lblDates.Visible = True
        lblDates.Caption = "Next Review Date"
        dlpDateRange(2).Visible = True
        dlpDateRange(3).Visible = True
        lblDates.Top = 6420
        dlpDateRange(2).Top = 6360
        dlpDateRange(3).Top = 6360
        chkLasg2PrvMonths.Visible = True
        chkLasg2PrvMonths.Top = 6376
    Else
        lblFromTo.Caption = "From / To Date"
        lblDates.Caption = "From / To Date"
        frmDate.Visible = False
        lblDates.Visible = False
        dlpDateRange(2).Visible = False
        dlpDateRange(3).Visible = False
        chkLasg2PrvMonths.Visible = False
    End If
    
    If InStr(1, fglbFileName, "SN2411VacationTmp.xls") > 0 Then
        lblAsOf.Visible = True
        dlpAsOf.Visible = True
        'dlpAsOf.Text = Format("12/31/" & Year(Now), "mm/dd/yyyy")
    Else
        lblAsOf.Visible = False
        dlpAsOf.Visible = False
    End If
Else
    chkAct.Visible = False
    chkInact.Visible = False
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If glbCompSerial = "S/N - 2276W" Then 'Ticket #27681 Franks 12/15/2015
    Call cmdView_Click
    Exit Sub
End If

If CriCheck() Then
    If Not PrtForm(frmRCsRpt.Caption & " Criteria", Me) Then Exit Sub
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
Resume Next

End Sub

Public Sub cmdView_Click()
Dim X%
Dim strWHand As String
Dim strSFormat$

On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    X% = Cri_SetAll()

    If glbCompSerial = "S/N - 2390W" And InStr(1, fglbFileName, "SN2390_License.xls") > 0 Then
        Call License_Report_XLS_CCORP
    ElseIf glbCompSerial = "S/N - 2382W" And InStr(1, fglbFileName, "SN2382_Employee_Salary.xls") > 0 Then
        'Ticket #24163 Franks 12/04/2013
        Call Samuel_Employee_XLS_Rpt
    ElseIf glbCompSerial = "S/N - 2382W" And InStr(1, fglbFileName, "SN2382_IDLComparison_short.xls") > 0 Then
        'Ticket #21656 Franks 02/29/2012
        Call Samuel_Comparison_XLS_Rpt_Short
    ElseIf glbCompSerial = "S/N - 2382W" And InStr(1, fglbFileName, "SN2382_IDLComparison.xls") > 0 Then
        'Ticket #16544 Samuel
        Call Samuel_Comparison_XLS_Report
    ElseIf glbCompSerial = "S/N - 2390W" And InStr(1, fglbFileName, "Export_EmployeeList.xls") > 0 Then       'Collectcorp Inc. Ticket #16952
        Call Export_Data_to_Excel_CollectCorp
    ElseIf glbCompSerial = "S/N - 2228W" And InStr(1, fglbFileName, "SN2228_Employment_Summary_Report.xls") > 0 Then       'Ticket #26368 Franks 02/23/2015
        Call Export_Data_to_Excel_SMDHU
    ElseIf glbCompSerial = "S/N - 2190W" And InStr(1, fglbFileName, "SeniorityList_Tmp.xls") > 0 Then
        'Ticket #17615 - WDDS
        If clpPT.Text <> "" Then
            If clpPT.Text = "FT" Then
                Call Export_FT_Seniority_Excel
            ElseIf clpPT.Text = "PT" Then
                Call Export_PT_Seniority_Excel
            ElseIf clpPT.Text = "FT,PT" Or clpPT.Text = "PT,FT" Then
                Call Export_Seniority_Excel
            End If
        Else
            Call Export_Seniority_Excel
        End If
    ElseIf glbWFC And InStr(1, fglbFileName, "WFC WPS Training Tmp.xls") > 0 Then 'Ticket #19137
        If isEmptyCourseMaster Then
            Call WFC_WPS_Training_Report
        Else
            Call WFC_WPS_Training_Rpt_WPSCODE 'Ticket #21330 Franks 02/14/2012
        End If
    ElseIf glbCompSerial = "S/N - 2370W" And InStr(1, fglbFileName, "AverageHourTmp.xls") > 0 Then
        'Chapman's Ice Cream Limited - Ticket #19104
        Call ChapmansExcelRpt_AverageHourRpt
    ElseIf glbCompSerial = "S/N - 2370W" And InStr(1, fglbFileName, "AbsenteeismTmp.xls") > 0 Then
        'Chapman's Ice Cream Limited - Ticket #19104
        Call ChapmansExcelRpt_AbsenteeismRpt
    ElseIf glbCompSerial = "S/N - 2347W" And InStr(1, fglbFileName, "VacAccrualTmp.xls") > 0 Then
        'Surrey Place - Ticket #20302
        Call SurreyPlace_3MonthVacationAccrual
    ElseIf glbCompSerial = "S/N - 2422W" And InStr(1, fglbFileName, "SN2422AccrualTmp.xls") > 0 Then
        'S.U.C.C.E.S.S.  - Ticket #21277
        Call SUCCESS_Accrual_XLS_Report
    ElseIf glbCompSerial = "S/N - 2241W" And InStr(1, fglbFileName, "SN2241_EmployeeInfoTmp.xls") > 0 Then       'Collectcorp Inc. Ticket #16952
        'Granite Club Ticket #22166
        Call Export_EmpData_to_Excel_GraniteClub
    ElseIf glbCompSerial = "S/N - 2263W" And InStr(1, fglbFileName, "SN2263ManulifeBenefitTmp.xls") > 0 Then
        'Ticket #22876 - County of Hastings
        Call Export_Manulife_Census_Data_Hastings
    ElseIf glbCompSerial = "S/N - 2290W" And InStr(1, fglbFileName, "SN2290BenefitRptTmp.xls") > 0 Then
        'Ticket #23938 - Goodmans
        Call Goodmans_Benefit_ExcelReport
    ElseIf glbCompSerial = "S/N - 2420W" And InStr(1, fglbFileName, "SN2420AttendUsedTmp.xls") > 0 Then
        'Ticket #27081 - Macaulay Child Development Centre
        Call MacaulayChild_HoursUsed
    ElseIf glbCompSerial = "S/N - 2411W" And InStr(1, fglbFileName, "SN2411GeneralEmpTmp.xls") > 0 Then
        'Ticket #29657 - Renamed the existing Fan Out report to General Employee and created a new format Employee Fan Out report
        'Ticket #27813 - WDGPHU
        'Call WDGPHU_EmployeeFanOut_ExcelReport
        Call WDGPHU_GeneralEmployee_ExcelReport
    ElseIf glbCompSerial = "S/N - 2411W" And InStr(1, fglbFileName, "SN2411EmpFanOutTmp.xls") > 0 Then
        'Ticket #29657 - Renamed the existing Fan Out report to General Employee and created this new format Employee Fan Out report
        Call WDGPHU_EmployeeFanOut_2nd_ExcelReport
    ElseIf glbCompSerial = "S/N - 2411W" And InStr(1, fglbFileName, "SN2411EmpTelephoneTmp.xls") > 0 Then
        'Ticket #27813 - WDGPHU
        Call WDGPHU_EmployeeTelephone_ExcelReport
    ElseIf glbCompSerial = "S/N - 2411W" And InStr(1, fglbFileName, "SN2411VacationTmp.xls") > 0 Then
        'Ticket #27813 - WDGPHU
        Call WDGPHU_Vacation_ExcelReport
    ElseIf glbCompSerial = "S/N - 2411W" And InStr(1, fglbFileName, "SN2411PerfManagementTmp.xls") > 0 Then
        'Ticket #27813 - WDGPHU
        Call WDGPHU_PerfManagement_ExcelReport
    'ElseIf glbCompSerial = "S/N - 2411W" And InStr(1, fglbFileName, "SN2411TrainingTmp.xls") > 0 Then
        'Ticket #27813 - WDGPHU
        'Call WDGPHU_Training_ExcelReport
    ElseIf glbCompSerial = "S/N - 2411W" And InStr(1, fglbFileName, "SN2411LicenseTmp.xls") > 0 Then
        'Ticket #27813 - WDGPHU
        Call WDGPHU_License_ExcelReport
    ElseIf glbCompSerial = "S/N - 2344W" And InStr(1, fglbFileName, "SN2344EmpInfoTmp.xls") > 0 Then
        'Ticket #29695 - Cascade Canada
        Call Cascade_EmployeeInfo_ExcelReport
    ElseIf glbCompSerial = "S/N - 2276W" And InStr(1, fglbFileName, "SN2276_CUPESick.rpt") > 0 Then
        Call NiagaraFallsCUPESickDailyRpt
        Call set_PrintState(True)
        Exit Sub
    Else
        Me.vbxCrystal.Destination = 0
        MDIMain.Timer1.Enabled = False
        Screen.MousePointer = DEFAULT
        Me.vbxCrystal.Action = 1
        vbxCrystal.Reset
        MDIMain.Timer1.Enabled = True
        
        'Ticket #22348 - United Way of Regina
        If (glbCompSerial = "S/N - 2444W" And InStr(1, fglbFileName, "SN2444_DivSumAcct.rpt") > 0) Then
            'Print/View another rpt as well
            Screen.MousePointer = HOURGLASS
            Call set_PrintState(False)
            
            'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
            'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
            Me.vbxCrystal1.WindowShowPrintSetupBtn = True
            
            fglbFileName = Replace(fglbFileName, "DivSumAcct", "DivSum")
            Me.vbxCrystal1.ReportFileName = fglbFileName
            
            Me.vbxCrystal1.SelectionFormula = glbstrSelCri
            Me.vbxCrystal1.Connect = RptODBC_SQL
            
            If dlpDateRange(0).Text <> "" And dlpDateRange(1).Text <> "" Then
                strSFormat$ = "Date Range: " & dlpDateRange(0).Text & " - " & dlpDateRange(1).Text
                Me.vbxCrystal1.Formulas(1) = "daterange='" & strSFormat$ & "'"
            Else
                strSFormat$ = "No date entered"
                Me.vbxCrystal1.Formulas(1) = "daterange='" & strSFormat$ & "'"
            End If
            Me.vbxCrystal1.Formulas(2) = "BenefitRate=" & txtBenRate & ""
            
            Me.vbxCrystal1.WindowTitle = "Divisional Summary Report"
            
            MDIMain.Timer1.Enabled = False
            Screen.MousePointer = DEFAULT
            Me.vbxCrystal1.Action = 1
            vbxCrystal1.Reset
            MDIMain.Timer1.Enabled = True
            fglbFileName = Replace(fglbFileName, "DivSum", "DivSumAcct")
        End If
    End If
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", fglbEmpTable, "SELECT")
Resume Next

End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%)) > 0 Then
    If glbWFC And xUserFlag Then 'Ticket #20859 Franks 01/30/2012
        Select Case intIdx%
        Case 0: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_LOC"
        Case 1: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_ORG"
        Case 2: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_EMP"
        Case 3: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_REGION"
        Case 4: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_ADMINBY"
        Case 5: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_SECTION"   'Lucy June 29, 2000
        Case 6: strCd$ = "HRJOB.JB_GRPCD" 'Fenwick only
        End Select
    Else
        Select Case intIdx%
        Case 0: strCd$ = fglbEmpTable & ".ED_LOC"
        Case 1: strCd$ = fglbEmpTable & ".ED_ORG"
        Case 2: strCd$ = fglbEmpTable & ".ED_EMP"
        Case 3: strCd$ = fglbEmpTable & ".ED_REGION"
        Case 4: strCd$ = fglbEmpTable & ".ED_ADMINBY"
        Case 5: strCd$ = fglbEmpTable & ".ED_SECTION"  'Lucy June 29, 2000
        Case 6: strCd$ = "HRJOB.JB_GRPCD" 'Fenwick only
        End Select
    End If
    
    'Ticket #29657 - Renamed the existing Employee Fan Out to General Employee and created a new Employee Fan Out report with new format
    'Ticket #27813 - WDGPHU
    If glbCompSerial = "S/N - 2411W" And strCd$ = fglbEmpTable & ".ED_EMP" And (InStr(1, fglbFileName, "SN2411GeneralEmpTmp.xls") > 0 Or InStr(1, fglbFileName, "SN2411EmpFanOutTmp.xls") > 0 Or _
        InStr(1, fglbFileName, "SN2411EmpTelephoneTmp.xls") > 0 Or InStr(1, fglbFileName, "SN2411LicenseTmp.xls") > 0 Or _
        InStr(1, fglbFileName, "SN2411VacationTmp.xls") > 0 Or InStr(1, fglbFileName, "SN2411PerfManagementTmp.xls") > 0) Then
        StatusCodeCri = ""
        StatusCodeCri = "'" & UCase(Replace(clpCode(intIdx%).Text, ",", "','")) & "'"
    Else
        'CodeCri = "({" & strCd$ & "} = '" & clpCode(intIdx%) & "')"
        CodeCri = "(Uppercase({" & strCd$ & "}) in  ['" & UCase(Replace(clpCode(intIdx%).Text, ",", "','")) & "'])"
        If glbLinamar And (strCd$ = fglbEmpTable & ".ED_REGION" Or strCd$ = fglbEmpTable & ".ED_SECTION") Then
            CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%) & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%) & "') )"
        End If
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

Private Sub Cri_Dept()
Dim DeptCri As String

If Len(clpDept.Text) > 0 Then DeptCri = "({" & fglbEmpTable & ".ED_DEPTNO} in ['" & Replace(clpDept.Text, ",", "','") & "']) "

If Len(DeptCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = DeptCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & DeptCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_Div()
Dim DivCri As String

If Len(clpDiv.Text) > 0 Then
    'DivCri = "({" & fglbEmpTable & ".ED_DIV} = '" & clpDiv.Text & "')"
    If glbWFC And xUserFlag Then 'Ticket #20859 Franks 01/30/2012
        DivCri = "({HR_OCC_HEALTH_SAFETY.EC_DIV} in  ['" & Replace(clpDiv.Text, ",", "','") & "'])"
    Else
        DivCri = "({" & fglbEmpTable & ".ED_DIV} in  ['" & Replace(clpDiv.Text, ",", "','") & "'])"
    End If
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

Private Sub Cri_AttSuper()
Dim EECri As String

If Len(elpSupShow.Text) > 0 Then
    'Ticket #16407
    'EECri = "{HR_ATTENDANCE.AD_EMPNBR} IN [" & getEmpnbr(elpSupShow.Text) & "] "
    EECri = "{HR_JOB_HISTORY.JH_REPTAU} IN [" & getEmpnbr(elpSupShow.Text) & "] "
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
    EECri = "{" & fglbEmpTable & ".ED_EMPNBR} IN [" & getEmpnbr(elpEEID.Text) & "] "
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

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, X%

If Len(txtShift.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"

'Ticket #22348 - United Way of Regina
If glbCompSerial = "S/N - 2444W" Then
    HisSQLUWR = HisSQLUWR & " AND AD_SHIFT ='" & txtShift.Text & "'"
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
'EECri = "{" & fglbEmpTable & ".ED_PT}= '" & clpPT.Text & "'"
If glbWFC And xUserFlag Then 'Ticket #20859 Franks 01/30/2012
    EECri = "{HR_OCC_HEALTH_SAFETY.EC_PT} IN ['" & Replace(clpPT.Text, ",", "','") & "']"
Else
    EECri = "{" & fglbEmpTable & ".ED_PT} IN ['" & Replace(clpPT.Text, ",", "','") & "']"
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_CountryOfEmployment()
Dim CountryCri As String

If Len(comCountryOfEmp.Text) > 0 Then
    If Not UCase(comCountryOfEmp.Text) = "ALL" Then
        If glbWFC And xUserFlag Then 'Ticket #20859 Franks 01/30/2012
            CountryCri = "({HR_OCC_HEALTH_SAFETY.EC_WORKCOUNTRY} = '" & comCountryOfEmp.Text & "')"
        Else
            CountryCri = "({" & fglbEmpTable & ".ED_WORKCOUNTRY} = '" & comCountryOfEmp.Text & "')"
        End If
    End If
End If

If Len(CountryCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CountryCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CountryCri
    End If
    glbiOneWhere = True
End If
End Sub

Private Sub Cri_UserText1()
Dim EECri As String, OneSet%, X%

If Len(txtUserText1.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_USER_TEXT1}= '" & txtUserText1.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Sub Cri_UserText2()
Dim EECri As String, OneSet%, X%

If Len(txtUserText2.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_USER_TEXT2}= '" & txtUserText2.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Function Cri_SetAll()
Dim X%, strRName$
Dim strSFormat$

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

'Ticket #29657 - WDGPHU wants to ignore the Department Security for this report
If glbCompSerial = "S/N - 2411W" And InStr(1, fglbFileName, "SN2411EmpFanOutTmp.xls") > 0 Then
    'No Department Security but ok with Department Selection
    Call Cri_Dept
Else
    Call glbCri_DeptUN(clpDept.Text)
End If

If glbWFC And xUserFlag Then 'Ticket #20859 Franks 01/30/2012
    glbstrSelCri = Replace(glbstrSelCri, "HREMP.ED_DEPTNO", "HR_OCC_HEALTH_SAFETY.EC_DEPTNO")
Else
    glbstrSelCri = Replace(glbstrSelCri, "HREMP.", fglbEmpTable & ".")
End If

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_AttSuper 'Ticket #16372 SPC
Call Cri_CountryOfEmployment
For X% = 0 To 5
    Call Cri_Code(X%)
Next X%

If IsFenwick Then 'Ticket #12505
    Call Cri_Code(6)
End If

If glbCompSerial = "S/N - 2390W" Then       'Collectcorp Inc. Ticket #14437
    If dlpDateRange(2).Visible Or dlpDateRange(3).Visible Then
        Call Cri_Dates
    End If
    If dlpDateRange(4).Visible Or dlpDateRange(5).Visible Then
        Call Cri_LicDates
    End If
    If clpUser(0).Visible Or clpUser(1).Visible Then
        Call Cri_License_Codes(0)
        Call Cri_License_Codes(1)
    End If
    If clpUser(2).Visible Then
        Call Cri_License_Codes(2)
    End If
    If comDateMonth.Visible Then
        If InStr(1, fglbFileName, "SN2390_Anniversary.rpt") > 0 Then
            Call Cri_DOHMonth
        ElseIf InStr(1, fglbFileName, "SN2390_Birthday.rpt") > 0 Then
            Call Cri_BirthMonth
        End If
    End If
End If

'Chapman's Ice Cream Limited - Ticket #19104
If glbCompSerial = "S/N - 2370W" Then
    Cri_Position
End If

If glbWFC Then 'Ticket #13957
    If InStr(1, fglbFileName, "Benefit_OptLife_Billing.rpt") > 0 Then
        Call WFCOptLifeBilling
        'Exit Function
    End If
    If InStr(1, fglbFileName, "Annual_Manulife_Benefits.rpt") > 0 Then
        Call WFC_Annual_Benefits
    End If
    If xWFCBilling Then 'Ticket #14061
        Me.vbxCrystal.Formulas(10) = "BillYear=" & txtFiscal.Text & ""
        Me.vbxCrystal.Formulas(11) = "BillMonth=" & ComMTH.ListIndex + 1 & ""
    End If
End If

'If glbCompSerial = "S/N - 2276W" Then 'Ticket #27681 Franks 12/14/2015
'    If InStr(1, fglbFileName, "SN2276_CUPESick.rpt") > 0 Then
'        Call NiagaraFallsCUPESickDailyRpt
'    End If
'End If

'DNSSAB - Ticket #14795
If glbCompSerial = "S/N - 2388W" And (InStr(1, fglbFileName, "SN2388_PooledBenSumm.rpt") > 0 Or InStr(1, fglbFileName, "SN12288_PooledBenDtl.rpt") > 0) Then
    Me.vbxCrystal.Formulas(10) = "BillYear=" & txtFiscal.Text & ""
    Me.vbxCrystal.Formulas(11) = "BillMonth=" & ComMTH.ListIndex + 1 & ""
End If

If (glbCompSerial <> "S/N - 2257W" And InStr(1, fglbFileName, "Request Report.rpt") = 0) And _
    (glbCompSerial <> "S/N - 2347W" And InStr(1, fglbFileName, "VacAccrualTmp.xls") = 0) And _
    (glbCompSerial <> "S/N - 2422W" And InStr(1, fglbFileName, "SN2422AccrualTmp.xls") = 0) And _
    (glbCompSerial <> "S/N - 2444W" And InStr(1, fglbFileName, "SN2444DivSumAcct.rpt") = 0) And _
    (glbCompSerial <> "S/N - 2420W" And InStr(1, fglbFileName, "SN2420AttendUsedTmp.xls") = 0) And _
    (glbCompSerial <> "S/N - 2411W" And InStr(1, fglbFileName, "SN2411PerfManagementTmp.xls") = 0) Then
    If frmDate.Visible Then Cri_FTDates
End If

'Ticket #22876 - County of Hastings
If glbCompSerial = "S/N - 2263W" And InStr(1, fglbFileName, "SN2263ManulifeBenefitTmp.xls") > 0 Then
    Call Cri_UserText1
    Call Cri_UserText2
End If

'Ticket #22348 - United Way of Regina
'If (glbCompSerial = "S/N - 2444W" And InStr(1, fglbFileName, "SN2444_DivSumAcct.rpt") = 0) Then
'    If frmDate.Visible Then Cri_FTDates
'End If

'Testing Excel reports
'Call Overtime_Report_XLS_HCAS
'Call Vacation_Report_XLS_HCAS
'Call Attendance_Report_XLS_HCAS
'Exit Function

Me.vbxCrystal.ReportFileName = fglbFileName

If Len(glbstrSelCri) >= 0 Then
    If glbWFC Then
        If InStr(1, fglbFileName, "Benefit_OptLife_Billing.rpt") > 0 Then  'Ticket #13957
            glbstrSelCri = "{WFC_MANULIFE_BENE_WRK.WRKEMP}='" & glbUserID & "'"
        End If
        If InStr(1, fglbFileName, "Annual_Manulife_Benefits.rpt") > 0 Then 'Ticket #14705
            glbstrSelCri = "{WFC_REPORT_WRK.WRKEMP}='" & glbUserID & "'"
        End If
        'Ticket #14031
        If InStr(1, fglbFileName, "Benefit_cost_details.rpt") > 0 Or InStr(1, fglbFileName, "Benefit_OptLife_Billing.rpt") > 0 Then
            If glbNoNONE And glbNoEXEC Then
                glbstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG }<> 'NONE' AND {HREMP.ED_ORG }<> 'EXEC'))"    'Hemu -EXE
            ElseIf glbNoNONE Then
                glbstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'NONE')"
            ElseIf glbNoEXEC Then
                glbstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'EXEC')"
            End If
        End If
        'Ticket #16616
        If InStr(1, UCase(fglbFileName), "RZPAYFORCE.RPT") > 0 Then
            If Len(clpCode(5).Text) > 0 Then
                'glbstrSelCri = "{PAYROLL_TRANSFER_RPT.EDSECTION}='" & clpCode(5).Text & "'"
                glbstrSelCri = "(Uppercase({" & "PAYROLL_TRANSFER_RPT.EDSECTION" & "}) in  ['" & UCase(Replace(clpCode(5).Text, ",", "','")) & "'])"
            Else
                glbstrSelCri = "(1=1)"
            End If
            'Ticket #19347
            'glbstrSelCri = glbstrSelCri & " AND " & "{PAYROLL_TRANSFER_RPT.TABLE}='ADP_Payforce' AND {PAYROLL_TRANSFER_RPT.WRKEMP}='" & glbUserID & "' "
            'Ticket #19492 Frank 11/30/10
            glbstrSelCri = glbstrSelCri & " AND " & "{PAYROLL_TRANSFER_RPT.TABLE}='ADP_Payforce' "
        End If
        'Ticket #19851 Franks 02/14/2011
        If InStr(1, UCase(fglbFileName), "RZMLFEXP.RPT") > 0 Then
            If Len(clpCode(5).Text) > 0 Then
                'glbstrSelCri = "{PAYROLL_TRANSFER_RPT.EDSECTION}='" & clpCode(5).Text & "'"
                glbstrSelCri = "(Uppercase({" & "PAYROLL_TRANSFER_RPT.EDSECTION" & "}) in  ['" & UCase(Replace(clpCode(5).Text, ",", "','")) & "'])"
            Else
                glbstrSelCri = "(1=1)"
            End If
            glbstrSelCri = glbstrSelCri & " AND " & "{PAYROLL_TRANSFER_RPT.TABLE}='Manulife_Export' "
        End If
    End If
    
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
    
    If glbCompSerial = "S/N - 2257W" And InStr(1, fglbFileName, "Request Report.rpt") > 0 Then
        Call HCAS_Request_Report
        Me.vbxCrystal.SelectionFormula = "{HR_REQUEST_RPT.REQ_WRKEMP}='" & glbUserID & "'"
    End If
    
    'Ticket #22348 - United Way of Regina
    If glbCompSerial = "S/N - 2444W" Then
        Call URW_DivisionalSummaryReports
        
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If
End If

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
    
    If glbCompSerial = "S/N - 2288W" And InStr(1, fglbFileName, "SN2288_1.rpt") > 0 Then
        Me.vbxCrystal.SubreportToChange = "counselcount"
        Me.vbxCrystal.Connect = RptODBC_SQL
        Me.vbxCrystal.SubreportToChange = ""
    End If
        'If glbWFC And InStr(1, fglbFileName, "mzINCIDENTWSUB.rpt") > 0 Then
        '    Me.vbxCrystal.SubreportToChange = "Root Causes"
        '    Me.vbxCrystal.Connect = RptODBC_SQL
        '    Me.vbxCrystal.SubreportToChange = "Corrective Action"
        '    Me.vbxCrystal.Connect = RptODBC_SQL
        '    Me.vbxCrystal.SubreportToChange = "Term Root Causes"
        '    Me.vbxCrystal.Connect = RptODBC_SQL
        '    Me.vbxCrystal.SubreportToChange = "Term Corrective Action"
        '    Me.vbxCrystal.Connect = RptODBC_SQL
        '    Me.vbxCrystal.SubreportToChange = ""
        'End If
    
    If glbCompSerial = "S/N - 2385W" And InStr(1, fglbFileName, "SN2385_1.rpt") > 0 Then   'Conservation Halton
        Me.vbxCrystal.Formulas(1) = "lblPT='" & lStr("Category") & "'"
        Me.vbxCrystal.Formulas(2) = "lblUnion='" & lStr("Union") & "'"
    End If
    
    If glbCompSerial = "S/N - 2390W" Then       'Collectcorp Inc. Ticket #14437
        If InStr(1, fglbFileName, "SN2390_Anniversary.rpt") > 0 Then
            Me.vbxCrystal.Formulas(1) = "lblDept='" & lStr("Department") & "'"
            Me.vbxCrystal.Formulas(2) = "lblHireDate='" & lStr("Original Hire Date") & "'"
            Me.vbxCrystal.Formulas(3) = "AsOfDate=Date('" & Format(dlpAsOf.Text, "mm/dd/yyyy") & "')"
            Me.vbxCrystal.Formulas(4) = "DOHMonth='" & lStr("Department") & "'"
        ElseIf InStr(1, fglbFileName, "SN2390_Birthday.rpt") > 0 Then
            Me.vbxCrystal.Formulas(1) = "lblDept='" & lStr("Department") & "'"
        ElseIf InStr(1, fglbFileName, "SN2390_CommPayroll.rpt") > 0 Then
            Me.vbxCrystal.Formulas(1) = "lblDept='" & lStr("Department") & "'"
            Me.vbxCrystal.Formulas(2) = "lblHireDate='" & lStr("Original Hire Date") & "'"
        ElseIf InStr(1, fglbFileName, "SN2390_LicenseAddr.rpt") > 0 Then
            Me.vbxCrystal.Formulas(1) = "lblDept='" & lStr("Department") & "'"
            Me.vbxCrystal.Formulas(2) = "lblHireDate='" & lStr("Original Hire Date") & "'"
            Me.vbxCrystal.Formulas(3) = "lblLocation='" & lStr("Location") & "'"
        ElseIf InStr(1, fglbFileName, "SN2390_LicenseTerm.rpt") > 0 Then
            Me.vbxCrystal.Formulas(1) = "lblDept='" & lStr("Department") & "'"
            Me.vbxCrystal.Formulas(2) = "lblLicNumber='" & lStr("UText 1") & "'"
        ElseIf InStr(1, fglbFileName, "SN2390_LicenseAdditions.rpt") > 0 Then
            Me.vbxCrystal.Formulas(1) = "lblDept='" & lStr("Department") & "'"
            Me.vbxCrystal.Formulas(2) = "lblHireDate='" & lStr("Original Hire Date") & "'"
            Me.vbxCrystal.Formulas(3) = "lblLicStatus='" & lStr("Code 2") & "'"
            Me.vbxCrystal.Formulas(4) = "lblLicProvState='" & lStr("Code 1") & "'"
        End If
    End If
    
    'Ticket #22348 - United Way of Regina
    If glbCompSerial = "S/N - 2444W" Then
        If InStr(1, fglbFileName, "SN2444_DivSumAcct.rpt") > 0 Then
            If dlpDateRange(0).Text <> "" And dlpDateRange(1).Text <> "" Then
                strSFormat$ = "Date Range: " & dlpDateRange(0).Text & " - " & dlpDateRange(1).Text
                Me.vbxCrystal.Formulas(1) = "daterange='" & strSFormat$ & "'"
            Else
                strSFormat$ = "No date entered"
                Me.vbxCrystal.Formulas(1) = "daterange='" & strSFormat$ & "'"
            End If
            Me.vbxCrystal.Formulas(2) = "BenefitRate= " & txtBenRate & ""
        End If
    End If
    
    'Ticket #22548 - Broadcasting Corp. - Moved reports from their Custom Program
    If glbCompSerial = "S/N - 2235W" Then
        If InStr(1, UCase(fglbFileName), "SN2235Y.RPT") > 0 Then
            'Labels for Time Clock Cards
            strSFormat$ = "PayPerDate = '" & dlpAsOf.Text & "'"
            Me.vbxCrystal.Formulas(0) = strSFormat$
        
        ElseIf (InStr(1, UCase(fglbFileName), "SN2235AA.RPT") > 0 Or InStr(1, UCase(fglbFileName), "SN2235AB.RPT") > 0) Then
            'Vacation / Casual Leave Application Form
            'Set which report to call based on the selection criteria
            If optCas.Value = True Then
                If cmbCas.ListIndex = 0 Then
                    strSFormat$ = "YesNo = 'yes'"
                ElseIf cmbCas.ListIndex = 1 Then
                     strSFormat$ = "YesNo = 'no'"
                End If
                strRName$ = glbIHRREPORTS & "SN2235AA.rpt"
            ElseIf optVac.Value = True Then
                If cmbVac.ListIndex = 0 Then
                    strSFormat$ = "YesNo = 'yes'"
                ElseIf cmbVac.ListIndex = 1 Then
                    strSFormat$ = "YesNo = 'no'"
                End If
                strRName$ = glbIHRREPORTS & "SN2235AB.rpt"
            End If
            
            Me.vbxCrystal.Formulas(0) = strSFormat$
            Me.vbxCrystal.ReportFileName = strRName$
        End If
    End If

    'Ticket #24374 - Surrey Place Centre
    If glbCompSerial = "S/N - 2347W" And InStr(1, fglbFileName, "SN2347_EmpDates.rpt") > 0 Then
        Me.vbxCrystal.Formulas(1) = "lblOHireDate='" & lStr("Original Hire Date") & "'"
        Me.vbxCrystal.Formulas(2) = "lblOtherDate1='" & lStr("Other Date 1") & "'"
        Me.vbxCrystal.Formulas(3) = "AsOfDate=Date('" & CVDate(Format(dlpAsOf.Text, "mm/dd/yyyy")) & "')"
        Me.vbxCrystal.Formulas(4) = "showDOB = " & IIf(gSec_Show_DOB = 0, False, True) & " "
    End If
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
           
    If glbCompSerial = "S/N - 2211W" And InStr(1, fglbFileName, "Vac and sick ent.rpt") > 0 Then
        Me.vbxCrystal.SubreportToChange = "HRFDTaken"
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.SubreportToChange = ""
    End If
    If glbCompSerial = "S/N - 2211W" And InStr(1, fglbFileName, "Vac and sick ent.rpt") > 0 Then
        Me.vbxCrystal.SubreportToChange = "OTHours"
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.SubreportToChange = ""
    End If
    If glbCompSerial = "S/N - 2211W" And InStr(1, fglbFileName, "Vac and sick ent.rpt") > 0 Then
        Me.vbxCrystal.SubreportToChange = "VacationTimeUsed"
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.SubreportToChange = ""
    End If
    If glbCompSerial = "S/N - 2211W" And InStr(1, fglbFileName, "Vac and sick ent.rpt") > 0 Then
        Me.vbxCrystal.SubreportToChange = "VacBooked"
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.SubreportToChange = ""
    End If
    If glbCompSerial = "S/N - 2211W" And InStr(1, fglbFileName, "Vac and sick ent.rpt") > 0 Then
        Me.vbxCrystal.SubreportToChange = "SickTimeUsed"
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.SubreportToChange = ""
    End If
    
    If glbCompSerial = "S/N - 2330W" Then   'Town of Marathon
        Me.vbxCrystal.Formulas(1) = "lblTypeVehicle='" & lStr("Type of Vehicle") & "'"
    End If
End If

Me.vbxCrystal.WindowTitle = cmbReports

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

Private Function CriCheck()
Dim X%

CriCheck = False

'Ticket #22548 - Broadcasting Corp. - Moved reports from their Custom Program
'Labels for Time Clock Cards
If (glbCompSerial = "S/N - 2235W" And InStr(1, UCase(fglbFileName), "SN2235Y.RPT") > 0) Then
    If Not gSec_Rpt_Master_Attendance Then
        frCasualVac.Visible = False

        MsgBox "You Do Not Have Authority For This Report", , "info:HR"
        Exit Function
    End If
End If

'Ticket #22548 - Broadcasting Corp. - Moved reports from their Custom Program
'Vacation / Casual Leave Application Form
If (glbCompSerial = "S/N - 2235W") And (InStr(1, UCase(fglbFileName), "SN2235AA.RPT") > 0 Or InStr(1, UCase(fglbFileName), "SN2235AB.RPT") > 0) Then
    If Not gSec_Rpt_Entitlements Then
        lblAsOf.Visible = False
        dlpAsOf.Visible = False
        lblAsOf.FontBold = False
        
        MsgBox "You Do Not Have Authority For This Report", , "info:HR"
        Exit Function
    End If
End If

'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
'    MsgBox lStr("If Division Entered - it must be known")
'     clpDiv.SetFocus
'    Exit Function
'End If

'If glbCompSerial = "S/N - 2347W" And InStr(1, fglbFileName, "SN234718.rpt") > 0 Then
'    If Len(Trim(clpDiv.Text)) = 0 Then
'        MsgBox lStr("Division cannot be left blank")
'        clpDiv.SetFocus
'        Exit Function
'    End If
'End If

'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
'    MsgBox "If Department Entered - it must be known"
'     clpDept.SetFocus
'    Exit Function
'End If

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


If Len(clpPT.Text) > 0 Then
    If Len(clpPT) > 0 And Not clpPT.ListChecker Then
        MsgBox lStr("Category code must be valid")
        clpPT.SetFocus
        Exit Function
    End If
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

'Chapman's Ice Cream Limited - Ticket #19104
If glbCompSerial = "S/N - 2370W" Then
    
    'If Not clpJob.ListChecker Then Exit Function

    'End Date/As of Date
    If dlpAsOf.Visible = True Then
        If Len(dlpAsOf.Text) > 0 Then
            If Not IsDate(dlpAsOf.Text) Then
                MsgBox "Invalid End Date"
                dlpAsOf.SetFocus
                Exit Function
            End If
        Else
            MsgBox "End Date is a required field"
            dlpAsOf.SetFocus
            Exit Function
        End If
    End If
    If dlpDateRange(2).Visible = True Then
        If Len(dlpDateRange(2).Text) > 0 Then
            If Not IsDate(dlpDateRange(2).Text) Then
                MsgBox "Invalid From Date"
                dlpDateRange(2).SetFocus
                Exit Function
            End If
        End If
        If Len(dlpDateRange(3).Text) > 0 Then
            If Not IsDate(dlpDateRange(3).Text) Then
                MsgBox "Invalid To Date"
                dlpDateRange(3).SetFocus
                Exit Function
            End If
        End If
    End If
    
End If

If xWFCBilling Or (glbCompSerial = "S/N - 2388W" And (InStr(1, fglbFileName, "SN2388_PooledBenSumm.rpt") > 0 Or InStr(1, fglbFileName, "SN12288_PooledBenDtl.rpt") > 0)) Then    ''DNSSAB - Ticket #14795
    If Len(txtFiscal) < 1 Then
        MsgBox "Year is a required field"
        txtFiscal.SetFocus
        Exit Function
    Else
        If Val(txtFiscal) < 2000 Then
            MsgBox "Year must be greater than 2000"
            txtFiscal.SetFocus
            Exit Function
        End If
    End If
End If

'Collectcorp Inc. - Ticket #14437
If glbCompSerial = "S/N - 2390W" Then
    If InStr(1, fglbFileName, "SN2390_Anniversary.rpt") > 0 Then
        'As of Date
        If Len(dlpAsOf.Text) > 0 Then
            If Not IsDate(dlpAsOf.Text) Then
                MsgBox "Not a valid As Of Date"
                dlpAsOf.SetFocus
                Exit Function
            End If
        Else
            MsgBox "As of Date is a required field"
            dlpAsOf.SetFocus
            Exit Function
        End If
    ElseIf InStr(1, fglbFileName, "SN2390_LicenseAddr.rpt") > 0 Then
        'Date of Hire
        For X% = 2 To 3
            If Len(dlpDateRange(X%).Text) > 0 Then
                If Not IsDate(dlpDateRange(X%).Text) Then
                    MsgBox "Not a valid " & lStr("Original Hire Date")
                    dlpDateRange(X%).SetFocus
                    Exit Function
                End If
            End If
        Next X%
        If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
            If dlpDateRange(3).Text < dlpDateRange(2).Text Then
                MsgBox "Invalid " & lStr("Original Hire Date") & " range. To " & lStr("Original Hire Date") & " cannot be less than From Date"
                dlpDateRange(3).SetFocus
                Exit Function
            End If
        End If
        
    ElseIf InStr(1, fglbFileName, "SN2390_LicenseTerm.rpt") > 0 Then
        'Termination Date, License Prov/State, License Status
        For X% = 2 To 3
            If Len(dlpDateRange(X%).Text) > 0 Then
                If Not IsDate(dlpDateRange(X%).Text) Then
                    MsgBox "Not a valid Termination Date"
                    dlpDateRange(X%).SetFocus
                    Exit Function
                End If
            End If
        Next X%
        
        If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
            If dlpDateRange(3).Text < dlpDateRange(2).Text Then
                MsgBox "Invalid Termination Date range. To Termination Date cannot be less than From Date"
                dlpDateRange(3).SetFocus
                Exit Function
            End If
        End If
                
        For X% = 0 To 1
            If Not clpUser(X).ListChecker Then Exit Function
        Next X%
        
    ElseIf InStr(1, fglbFileName, "SN2390_LicenseAdditions.rpt") > 0 Then
        'Date Submitted, License Prov/State, License Status
        For X% = 2 To 3
            If Len(dlpDateRange(X%).Text) > 0 Then
                If Not IsDate(dlpDateRange(X%).Text) Then
                    MsgBox "Not a valid " & lStr("Date 1")
                    dlpDateRange(X%).SetFocus
                    Exit Function
                End If
            End If
        Next X%
        If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
            If dlpDateRange(3).Text < dlpDateRange(2).Text Then
                MsgBox "Invalid " & lStr("Date 1") & " range. To" & lStr("Date 1") & " cannot be less than From Date"
                dlpDateRange(3).SetFocus
                Exit Function
            End If
        End If
        
        For X% = 0 To 1
            If Not clpUser(X).ListChecker Then Exit Function
        Next X%
    End If

End If

'Ticket #16616
'Ticket #19851 Franks 02/14/2011
If glbWFC Then
    If InStr(1, UCase(fglbFileName), "RZPAYFORCE.RPT") > 0 Or InStr(1, UCase(fglbFileName), "RZMLFEXP.RPT") > 0 Then
    Dim xCount As Long
    Dim xSTmp As String
        If Len(clpCode(5).Text) = 0 Then
            'No   TB_KEY means it is super user
            'With TB_KEY means it is user with Plant security
            If InStr(1, glbSeleSection, "TB_KEY") > 0 Then
                MsgBox lStr("Section") & " is required.        "
                clpCode(5).SetFocus
                Exit Function
            End If
        End If
        If Len(clpCode(5).Text) > 0 Then
            If InStr(1, glbSeleSection, "TB_KEY") > 0 Then
                xCount = CharCount(clpCode(5).Text, ",") + 1
                If xCount > 0 Then
                    For X% = 1 To xCount
                        xSTmp = "'" & CSVGet(clpCode(5).Text, X%) & "'"
                        If InStr(1, glbSeleSection, xSTmp) = 0 Then
                            MsgBox "You do not have access right for " & lStr("Section") & " " & xSTmp & " "
                            clpCode(5).SetFocus
                            Exit Function
                        End If
                    Next
                End If
            End If
        End If
    End If
End If

'Ticket #20302 - Surrey Place Centre
If (glbCompSerial = "S/N - 2347W" And InStr(1, fglbFileName, "VacAccrualTmp.xls") > 0) Or _
    (glbCompSerial = "S/N - 2422W" And InStr(1, fglbFileName, "SN2422AccrualTmp.xls") > 0) Then
    If frmDate.Visible = True Then
        'Ticket #21277 - No need of From Date for SUCCESS
        If Not (glbCompSerial = "S/N - 2422W" And InStr(1, fglbFileName, "SN2422AccrualTmp.xls") > 0) Then
            If Len(dlpDateRange(0).Text) = 0 Then
                MsgBox "From Date is a required field"
                dlpDateRange(0).SetFocus
                Exit Function
            End If
        
            If Not IsDate(dlpDateRange(0).Text) Then
                MsgBox "Invalid From Date"
                dlpDateRange(0).SetFocus
                Exit Function
            End If
        End If
        
        If Len(dlpDateRange(1).Text) = 0 Then
            If (glbCompSerial = "S/N - 2422W" And InStr(1, fglbFileName, "SN2422AccrualTmp.xls") > 0) Then
                MsgBox "Report End Date is a required field"
            Else
                MsgBox "To Date is a required field"
            End If
            dlpDateRange(1).SetFocus
            Exit Function
        End If
        
        If Not IsDate(dlpDateRange(1).Text) Then
            If (glbCompSerial = "S/N - 2422W" And InStr(1, fglbFileName, "SN2422AccrualTmp.xls") > 0) Then
                MsgBox "Invalid Report End Date"
            Else
                MsgBox "Invalid To Date"
            End If
            dlpDateRange(1).SetFocus
            Exit Function
        End If
    End If
End If

'Ticket #22348 - United Way of Regina
If (glbCompSerial = "S/N - 2444W" And InStr(1, fglbFileName, "SN2444_DivSumAcct.rpt") > 0) Then
    If frmDate.Visible = True Then
        If Len(dlpDateRange(0).Text) > 0 Then
            If Not IsDate(dlpDateRange(0).Text) Then
                MsgBox "Invalid From Date"
                dlpDateRange(0).SetFocus
                Exit Function
            End If
        End If
        
        If Len(dlpDateRange(1).Text) > 0 Then
            If Not IsDate(dlpDateRange(1).Text) Then
                MsgBox "Invalid To Date"
                dlpDateRange(1).SetFocus
                Exit Function
            End If
        End If
        
        If IsDate(dlpDateRange(0).Text) And IsDate(dlpDateRange(1).Text) Then
            If CVDate(dlpDateRange(0).Text) > CVDate(dlpDateRange(1).Text) Then
                MsgBox "From Date cannot be greater than To Date"
                dlpDateRange(0).SetFocus
                Exit Function
            End If
        End If
    End If
    
    If Len(txtBenRate.Text) = 0 Then
        MsgBox "Benefit Rate cannot be blank"
        txtBenRate.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtBenRate.Text) Then
        MsgBox "Invalid Benefit Rate"
        txtBenRate.SetFocus
        Exit Function
    End If
End If

'Ticket #22548 - Broadcasting Corp. - Moved reports from their Custom Program
'Labels for Time Clock Cards
If glbCompSerial = "S/N - 2235W" Then
    If InStr(1, UCase(fglbFileName), "SN2235Y.RPT") > 0 Then
        If dlpAsOf.Visible = True Then
            If Len(dlpAsOf.Text) > 0 Then
                If Not IsDate(dlpAsOf.Text) Then
                    MsgBox "Invalid Pay Period Date"
                    dlpAsOf.SetFocus
                    Exit Function
                End If
            Else
                MsgBox "Pay Period Date is a required field"
                dlpAsOf.SetFocus
                Exit Function
            End If
        End If
    End If
End If
    
'Ticket #27081 - Macaulay Child Development Centre
If (glbCompSerial = "S/N - 2420W" And InStr(1, fglbFileName, "SN2420AttendUsedTmp.xls") > 0) Then
    If frmDate.Visible = True Then
        If Len(dlpDateRange(0).Text) > 0 Then
            If Not IsDate(dlpDateRange(0).Text) Then
                MsgBox "Invalid From Date"
                dlpDateRange(0).SetFocus
                Exit Function
            End If
        End If
        
        If Len(dlpDateRange(1).Text) > 0 Then
            If Not IsDate(dlpDateRange(1).Text) Then
                MsgBox "Invalid To Date"
                dlpDateRange(1).SetFocus
                Exit Function
            End If
        End If
        
        If IsDate(dlpDateRange(0).Text) And IsDate(dlpDateRange(1).Text) Then
            If CVDate(dlpDateRange(0).Text) > CVDate(dlpDateRange(1).Text) Then
                MsgBox "From Date cannot be greater than To Date"
                dlpDateRange(0).SetFocus
                Exit Function
            End If
        End If
    End If
End If

'Ticket #27813 - WDGPHU
If glbCompSerial = "S/N - 2411W" Then
    If InStr(1, fglbFileName, "SN2411PerfManagementTmp.xls") > 0 Then
        If frmDate.Visible = True Then
            If Len(dlpDateRange(0).Text) > 0 Then
                If Not IsDate(dlpDateRange(0).Text) Then
                    MsgBox "Invalid From Review Date"
                    dlpDateRange(0).SetFocus
                    Exit Function
                End If
            End If
            
            If Len(dlpDateRange(1).Text) > 0 Then
                If Not IsDate(dlpDateRange(1).Text) Then
                    MsgBox "Invalid To Review Date"
                    dlpDateRange(1).SetFocus
                    Exit Function
                End If
            End If
            
            If (IsDate(dlpDateRange(0).Text) And Not IsDate(dlpDateRange(1).Text)) Or (Not IsDate(dlpDateRange(0).Text) And IsDate(dlpDateRange(1).Text)) Then
                MsgBox "Enter both 'From Review Date' and 'To Review Date' or none"
                dlpDateRange(0).SetFocus
                Exit Function
            End If
            
            If IsDate(dlpDateRange(0).Text) And IsDate(dlpDateRange(1).Text) Then
                If CVDate(dlpDateRange(0).Text) > CVDate(dlpDateRange(1).Text) Then
                    MsgBox "From Review Date cannot be greater than To Review Date"
                    dlpDateRange(0).SetFocus
                    Exit Function
                End If
            End If
        End If
    
        If lblDates.Visible = True Then
            If Len(dlpDateRange(2).Text) > 0 Then
                If Not IsDate(dlpDateRange(2).Text) Then
                    MsgBox "Invalid From Next Review Date"
                    dlpDateRange(2).SetFocus
                    Exit Function
                End If
            End If
            
            If Len(dlpDateRange(3).Text) > 0 Then
                If Not IsDate(dlpDateRange(3).Text) Then
                    MsgBox "Invalid To Next Review Date"
                    dlpDateRange(3).SetFocus
                    Exit Function
                End If
            End If
            
            If (IsDate(dlpDateRange(2).Text) And Not IsDate(dlpDateRange(3).Text)) Or (Not IsDate(dlpDateRange(2).Text) And IsDate(dlpDateRange(3).Text)) Then
                MsgBox "Enter both 'From Next Review Date' and 'To Next Review Date' or none"
                dlpDateRange(2).SetFocus
                Exit Function
            End If
            
            If IsDate(dlpDateRange(2).Text) And IsDate(dlpDateRange(3).Text) Then
                If CVDate(dlpDateRange(2).Text) > CVDate(dlpDateRange(3).Text) Then
                    MsgBox "From Next Review Date cannot be greater than To Next Review Date"
                    dlpDateRange(2).SetFocus
                    Exit Function
                End If
            End If
            
            If chkLasg2PrvMonths.Visible = True Then
                If chkLasg2PrvMonths.Value = 1 Then
                    If Not IsDate(dlpDateRange(2).Text) Or Not IsDate(dlpDateRange(3).Text) Then
                        MsgBox "Both 'From Next Review Date' and 'To Next Review Date' is required if 'Bring over the last 2 previous months' is checked"
                        dlpDateRange(2).SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
    ElseIf InStr(1, fglbFileName, "SN2411VacationTmp.xls") > 0 Then
        If dlpAsOf.Visible = True Then
            If Len(dlpAsOf.Text) > 0 Then
                If Not IsDate(dlpAsOf.Text) Then
                    MsgBox "Invalid As Of Date"
                    dlpAsOf.SetFocus
                    Exit Function
                End If
            Else
                MsgBox "As Of Date is a required field"
                dlpAsOf.SetFocus
                Exit Function
            End If
        End If
    End If
End If

If glbCompSerial = "S/N - 2276W" Then 'Ticket #27681 Franks 12/14/2015
    If InStr(1, fglbFileName, "SN2276_CUPESick.rpt") > 0 Then
        If Len(dlpDateRange(0).Text) = 0 Then
            MsgBox "From Date is a required field"
            dlpDateRange(0).SetFocus
            Exit Function
        End If
        If Not IsDate(dlpDateRange(0).Text) Then
            MsgBox "Invalid From Date"
            dlpDateRange(0).SetFocus
            Exit Function
        End If
        If Len(dlpDateRange(1).Text) = 0 Then
            MsgBox "To Date is a required field"
            dlpDateRange(1).SetFocus
            Exit Function
        End If
        If Not IsDate(dlpDateRange(1).Text) Then
            MsgBox "Invalid To Date"
            dlpDateRange(1).SetFocus
            Exit Function
        End If
    End If
End If

CriCheck = True
End Function

Private Sub comGroup_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
Dim X%

    comGroup(0).Clear
    comGroup(0).AddItem "Employee Name"
    comGroup(0).AddItem "Employee Number"
    comGroup(0).ListIndex = 0

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = "FRMRCSRPT"

Dim rsALLRPT As New ADODB.Recordset
Dim SQLQ
Dim xTemplate As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbUserID)

Screen.MousePointer = HOURGLASS

If glbOracle Then
    SQLQ = "SELECT * FROM HR_CUSTOMRPT, HR_SECRPT"
    SQLQ = SQLQ & " WHERE HR_CUSTOMRPT.RT_RPTNAME(+)= HR_SECRPT.FUNCTION "
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        SQLQ = SQLQ & " AND USERID='" & Replace(glbUserID, "'", "''") & "'"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        SQLQ = SQLQ & " AND USERID='" & Replace(xTemplate, "'", "''") & "'"
    End If
    SQLQ = SQLQ & " AND ACCESSABLE<>0 "
    SQLQ = SQLQ & " ORDER BY UPPER(FUNCTION)"
Else
    SQLQ = "SELECT * FROM HR_CUSTOMRPT LEFT JOIN HR_SECRPT"
    SQLQ = SQLQ & " ON HR_CUSTOMRPT.RT_RPTNAME= HR_SECRPT.[FUNCTION] "
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        SQLQ = SQLQ & " WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
    Else
        SQLQ = SQLQ & " WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    End If
    SQLQ = SQLQ & " AND ACCESSABLE<>0 "
    SQLQ = SQLQ & " ORDER BY [FUNCTION]"
End If
rsALLRPT.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
cmbReports.Clear
Do Until rsALLRPT.EOF
    cmbReports.AddItem rsALLRPT("RT_RPTNAME")
    rsALLRPT.MoveNext
Loop
rsALLRPT.Close
If cmbReports.ListCount <> 0 Then
    cmbReports.ListIndex = 0
Else
'    cmdView.Enabled = False
'    cmdPrint.Enabled = False
End If

If glbCompSerial = "S/N - 2369W" And InStr(1, fglbFileName, "sn2369AttBonus.rpt") > 0 Then
    lblFromTo.FontBold = True
End If
            
'Ticket #16544
If glbCompSerial = "S/N - 2382W" And (InStr(1, fglbFileName, "SN2382_IDLComparison.xls") > 0 Or InStr(1, fglbFileName, "SN2382_IDLComparison_short.xls") > 0) Then
    frmActTerm.Visible = True
Else
    frmActTerm.Visible = False
End If

If glbCompSerial = "S/N - 2201W" Then
    IsFenwick = True
    lblBCode.Visible = True
    clpCode(6).Visible = True
End If

'Chapman's Ice Cream Limited - Ticket #19104
If glbCompSerial = "S/N - 2370W" Then
    clpJob.Visible = True
    clpJob.Top = clpCode(6).Top
    lblPosition.Visible = True
    lblPosition.Top = lblBCode.Top
End If

Call setRptCaption(Me)

If glbLinamar Then clpCode(3).MaxLength = 8

Call addCountryItems 'Frank 09/07/2007 Ticket #13621

'If glbCompSerial = "S/N - 2347W" Then
'    lblDiv.FontBold = True
'End If

If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

If glbCompSerial = "S/N - 2347W" Then 'Ticket #16372 SPC
    lblAttSuper.Top = 4300
    lblAttSuper.Visible = True
    elpSupShow.Top = 4300
    elpSupShow.Visible = True
End If

Call INI_Controls(Me)
If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

If glbWFC Or glbCompSerial = "S/N - 2388W" Then  'Ticket #14061   ''DNSSAB - Ticket #14795
    frmBeneBilling.Top = cmbReports.Top + 480 ' 5400
    frmBeneBilling.Left = cmbReports.Left
    ComMTH.AddItem "Jan"
    ComMTH.AddItem "Feb"
    ComMTH.AddItem "Mar"
    ComMTH.AddItem "Apr"
    ComMTH.AddItem "May"
    ComMTH.AddItem "Jun"
    ComMTH.AddItem "Jul"
    ComMTH.AddItem "Aug"
    ComMTH.AddItem "Sep"
    ComMTH.AddItem "Oct"
    ComMTH.AddItem "Nov"
    ComMTH.AddItem "Dec"
    txtFiscal.Text = Year(Date)
    xMonNum = month(Date)
    ComMTH.ListIndex = xMonNum - 1
End If

If glbCompSerial = "S/N - 2390W" Then       'Collectcorp Inc. Ticket #14437
    comDateMonth.AddItem "January"
    comDateMonth.AddItem "February"
    comDateMonth.AddItem "March"
    comDateMonth.AddItem "April"
    comDateMonth.AddItem "May"
    comDateMonth.AddItem "June"
    comDateMonth.AddItem "July"
    comDateMonth.AddItem "August"
    comDateMonth.AddItem "September"
    comDateMonth.AddItem "October"
    comDateMonth.AddItem "November"
    comDateMonth.AddItem "December"
    comDateMonth.ListIndex = month(Date) - 1
End If
Screen.MousePointer = DEFAULT

'Hemu - Add Serial # Control
'lblAsOf.Visible = True
'dlpAsOf.Visible = True
'lblAsOf.FontBold = True
'dlpAsOf.Text = Date


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

Private Sub Cri_FTDatSQL() 'Ticket #24163 Franks 12/05/2013
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%

    If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
        TempCri = " " & fglbDateTable & "." & fglbDateField & " >= " & Date_SQL(dlpDateRange(0).Text) & " "
        TempCri = TempCri & "AND " & fglbDateTable & "." & fglbDateField & " <= " & Date_SQL(dlpDateRange(1).Text) & " "
        GoTo Cri_FTDatst
    End If
    
    For X% = 0 To 1
        If Len(dlpDateRange(X).Text) > 0 Then
            TempCri = " " & fglbDateTable & "." & fglbDateField & " "
            If X% = 0 Then
                TempCri = TempCri & " >= "
            Else
                TempCri = TempCri & " <= "
            End If

            TempCri = TempCri & Date_SQL(dlpDateRange(X).Text) & " "
            GoTo Cri_FTDatst
        End If
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

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%
If InStr(1, fglbFileName, "WFC WPS Training Tmp.xls") > 0 Then 'Ticket #19137
    Exit Sub 'do not use it here
End If

'Ticket #24163 Franks 12/05/2013
If glbSamuel And InStr(1, fglbFileName, "SN2382_Employee_Salary.xls") > 0 Then
    Call Cri_FTDatSQL
    Exit Sub
End If

If glbOttawaCCAC And InStr(fglbFileName, "Uptodate_Entitlement") <> 0 Then
    For X% = 0 To 1
        If Len(dlpDateRange(X).Text) > 0 Then
            TempCri = "({" & fglbDateTable & "." & fglbDateField & "} "
            If X% = 0 Then
                TempCri = "FromDate="
            Else
                TempCri = "ToDate="
            End If
            dtYYY% = Year(dlpDateRange(X).Text)
            dtMM% = month(dlpDateRange(X).Text)
            dtDD% = Day(dlpDateRange(X).Text)
            Me.vbxCrystal.Formulas(100 + X%) = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        End If
    Next X%
    Exit Sub
Else
    If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
        'HisSQLUWR = HisSQLUWR & " AND (AD_DOA Between "
        'HisSQLUWR = HisSQLUWR & Date_SQL(dlpDateRange(0)) & "And "
        'HisSQLUWR = HisSQLUWR & Date_SQL(dlpDateRange(1)) & ") "

        TempCri = "({" & fglbDateTable & "." & fglbDateField & "} "
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
        
        If glbCompSerial = "S/N - 2369W" And (InStr(1, fglbFileName, "sn2369AttBonus.rpt") > 0 Or InStr(1, fglbFileName, "sn2369CBonus.rpt") > 0) Then
            Me.vbxCrystal.Formulas(0) = "dteStart= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
        End If
        
        dtYYY% = Year(dlpDateRange(1).Text)
        dtMM% = month(dlpDateRange(1).Text)
        dtDD% = Day(dlpDateRange(1).Text)
        TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        
        If glbCompSerial = "S/N - 2369W" And (InStr(1, fglbFileName, "sn2369AttBonus.rpt") > 0 Or InStr(1, fglbFileName, "sn2369CBonus.rpt") > 0) Then
            Me.vbxCrystal.Formulas(1) = "dteEnd= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
        End If
        
        GoTo Cri_FTDatst
    End If
    
    For X% = 0 To 1
        If Len(dlpDateRange(X).Text) > 0 Then
            TempCri = "({" & fglbDateTable & "." & fglbDateField & "} "
            If X% = 0 Then
                TempCri = TempCri & " >= "
                'HisSQLUWR = HisSQLUWR & "AND (AD_DOA >= " & Date_SQL(dlpDateRange(0).Text) & ") "
            Else
                TempCri = TempCri & " <= "
                'HisSQLUWR = HisSQLUWR & "AND (AD_DOA <= " & Date_SQL(dlpDateRange(1).Text) & " ) "
            End If
            dtYYY% = Year(dlpDateRange(X).Text)
            dtMM% = month(dlpDateRange(X).Text)
            dtDD% = Day(dlpDateRange(X).Text)
            TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
            GoTo Cri_FTDatst
        End If
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
End If
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

If cmbReports.Text = "" Then
Printable = False
Else
Printable = True
End If

End Property

Public Sub SET_UP_MODE()
Call set_Buttons
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub addCountryItems()
Dim ctylist, X

ctylist = CountryList
X = 1
Do While X > 0
    X = InStr(ctylist, "&")
    If X > 0 Then
        comCountryOfEmp.AddItem Left(ctylist, X - 1)
        ctylist = Mid(ctylist, X + 1)
    Else
        comCountryOfEmp.AddItem ctylist
    End If
Loop

End Sub
Private Function CountryList() As String
Dim xCountryList As String, ctyFile
xCountryList = ""
ctyFile = glbIHRREPORTS & "CountryList.MTF"

On Error GoTo ErrorHandler

If File(ctyFile) Then
    Open ctyFile For Input As #1
    Input #1, xCountryList
    Close #1
End If

ResumeHere:
'If InStr(xCountryList, BasicCountry) = 0 Then
'    xCountryList = BasicCountry
'End If
If InStr(xCountryList, comCountryOfEmp) = 0 And comCountryOfEmp <> "" Then
    xCountryList = xCountryList & "&" & comCountryOfEmp
    comCountryOfEmp.AddItem comCountryOfEmp
End If
Open ctyFile For Output As #1
Print #1, xCountryList
Close #1
CountryList = xCountryList
Exit Function

ErrorHandler:
If Err.Number = 62 Then
    ' Corrupted CountryList.MTF, kill it and regenerate
    Close #1
    MsgBox "Found corrupt CountryList.MTF.  info:HR will re-create this file.", vbInformation + vbOKOnly, "Corrupted Country List"
    Kill ctyFile
    Resume ResumeHere
Else
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number & " in CountryList"
    Resume Next
End If
End Function

Private Sub NiagaraFallsCUPESickDailyRpt() 'Ticket #27681 Franks 12/14/2015
Dim rsAtt As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim SQLQ As String, xEmpNbr, xCode
Dim I, totNum
Dim xDHrs, xDay, xLevel
Dim strSFormat$

    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 0
    
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute "DELETE FROM HRATTWRK WHERE AD_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.CommitTrans
    
    SQLQ = "SELECT * FROM HRATTWRK WHERE AD_WRKEMP = '" & glbUserID & "' "
    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    SQLQ = "SELECT * FROM HREMP WHERE (1=1) "
    
    If clpDiv.Text <> "" Then SQLQ = SQLQ & " AND ED_DIV IN ('" & Replace(clpDiv.Text, ",", "','") & "')"
    If clpDept.Text <> "" Then SQLQ = SQLQ & " AND ED_DEPTNO IN ('" & Replace(clpDept.Text, ",", "','") & "')"
    If clpCode(0).Text <> "" Then SQLQ = SQLQ & " AND ED_LOC IN ('" & Replace(clpCode(0).Text, ",", "','") & "')"
    'If clpCode(1).Text <> "" Then SQLQ = SQLQ & " AND ED_ORG IN ('" & Replace(clpCode(1).Text, ",", "','") & "')"
    SQLQ = SQLQ & " AND ED_ORG = 'CUPE' " 'for CUPE only
    If clpCode(2).Text <> "" Then SQLQ = SQLQ & " AND ED_EMP IN ('" & Replace(clpCode(2).Text, ",", "','") & "')"
    If clpCode(3).Text <> "" Then SQLQ = SQLQ & " AND ED_REGION IN ('" & Replace(clpCode(3).Text, ",", "','") & "')"
    If clpCode(4).Text <> "" Then SQLQ = SQLQ & " AND ED_ADMINBY IN ('" & Replace(clpCode(4).Text, ",", "','") & "')"
    If clpCode(5).Text <> "" Then SQLQ = SQLQ & " AND ED_SECTION IN ('" & Replace(clpCode(5).Text, ",", "','") & "')"
    If clpPT.Text <> "" Then SQLQ = SQLQ & " AND ED_PT IN ('" & Replace(clpPT.Text, ",", "','") & "')"
    If elpEEID.Text <> "" Then SQLQ = SQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
    
    If rsEmp.State <> 0 Then rsEmp.Close
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        totNum = rsEmp.RecordCount: I = 0
    End If
    Do While Not rsEmp.EOF
        MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
        I = I + 1
        DoEvents
        SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS TOTHRS FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
        SQLQ = SQLQ & "AND AD_REASON like 'SIC%' "
        SQLQ = SQLQ & "AND AD_DOA >= " & Date_SQL(dlpDateRange(0).Text) & " "
        SQLQ = SQLQ & "AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text) & " "
        SQLQ = SQLQ & "GROUP BY AD_EMPNBR "
        If rsAtt.State <> 0 Then rsAtt.Close
        rsAtt.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsAtt.EOF Then
            If rsAtt("TOTHRS") > 0 Then
                'get HRS/Day from Emp Job His
                xDHrs = getLocEmpPosField(rsEmp("ED_EMPNBR"), 0, "JH_DHRS")
                xDay = 0
                If xDHrs > 0 Then
                    xDay = Round((rsAtt("TOTHRS") / xDHrs), 4)
                End If
                xLevel = 0
                If xDay >= 4 Then
                    If xDay / 4 >= 1 Then 'o   Has been off sick for 4+ days in a calendar year
                        xLevel = 1
                    End If
                    If xDay / 7 >= 1 Then 'o   Has been off sick for 7+ days in a calendar year
                        xLevel = 2
                    End If
                    If xDay / 14 >= 1 Then 'o   Has been off sick for 14+ days in a calendar year
                        xLevel = 3
                    End If
                    If xDay / 84 >= 1 Then 'o   Has been off sick for 84+ days in a calendar year
                        xLevel = 4
                    End If
                    'Debug.Print rsEmp("ED_EMPNBR"), xDay
                    'update HRATTWRK table
                    Call Upt_HRATTWRK(rsEmp("ED_EMPNBR"), xDay, xDHrs, xLevel, rsWRK)
                    
                End If
            End If
        End If
        rsEmp.MoveNext
    Loop
    rsEmp.Close
    If rsWRK.State <> 0 Then rsWRK.Close
    
    Me.vbxCrystal1.WindowShowPrintSetupBtn = True

    Me.vbxCrystal1.ReportFileName = fglbFileName
    
    'Me.vbxCrystal1.SelectionFormula = glbstrSelCri
    Me.vbxCrystal1.Connect = RptODBC_SQL
    
    If dlpDateRange(0).Text <> "" And dlpDateRange(1).Text <> "" Then
        strSFormat$ = "Date Range: " & dlpDateRange(0).Text & " - " & dlpDateRange(1).Text
        Me.vbxCrystal1.Formulas(1) = "daterange='" & strSFormat$ & "'"
    Else
        strSFormat$ = "No date entered"
        Me.vbxCrystal1.Formulas(1) = "daterange='" & strSFormat$ & "'"
    End If
    
    'Me.vbxCrystal1.WindowTitle = "Divisional Summary Report"
    
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal1.Action = 1
    vbxCrystal1.Reset
            
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT
    
End Sub

Private Sub WFC_Annual_Benefits()
Dim rsBeneOpt As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim SQLQ As String, xEmpNbr, xCode
Dim I, totNum
Dim xPolicyNo, xOptAccount
Dim rsDep As New ADODB.Recordset
Dim xBillingDate
Dim BenefitCodeList As String

    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 0
    ''Ticket #19979 Franks 03/09/2011 Benefit codes were changed
    ''BenefitCodeList = " ('LIF', 'LIF1', 'LIF2', 'LIF3', 'LFR', 'LFR1', 'LFR2', 'LFR3','OPLF','OPLS','LTD','AD&D','EHC','DENT')"
    'Ticket #26779 Franks 03/09/2015 replace DN with DEN
    'BenefitCodeList = " ('LIF15', 'LIF2', 'LIF25', 'LIFF', 'LFR', 'LFR1', 'LFR2','OPLF','OPLS','SD','LD','IA','EHC','DN')"
    BenefitCodeList = " ('LIF15', 'LIF2', 'LIF25', 'LIFF', 'LFR', 'LFR1', 'LFR2','OPLF','OPLS','SD','LD','IA','EHC','DEN')"
    
    xBillingDate = CVDate(ComMTH.Text & " 1, " & txtFiscal)
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute "DELETE FROM WFC_REPORT_WRK WHERE WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001W.CommitTrans

    HisSQL = " BF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
    HisSQL = Replace(HisSQL, "Uppercase", "Upper")
    SQLQ = "SELECT * FROM HRBENFT WHERE " & HisSQL & " "
    SQLQ = SQLQ & " AND (BF_EMPNBR in (select ED_EMPNBR from HREMP WHERE ED_COUNTRY = 'CANADA' AND NOT (ED_USER_TEXT1 IS NULL OR ED_USER_TEXT1 = '' ) AND NOT (ED_USER_TEXT2 IS NULL OR ED_USER_TEXT2 = '') AND NOT (ED_USER_NUM1 IS NULL) ))"
    'SQLQ = SQLQ & "AND (BF_BCODE = 'OPLF' OR BF_BCODE = 'OPLS' OR BF_BCODE = 'OPLC') "
    SQLQ = SQLQ & "AND (BF_BCODE IN " & BenefitCodeList & " )"
    'SQLQ = SQLQ & "AND NOT (BF_COVER = 'W' )"
    'SQLQ = SQLQ & "AND (BF_AMT > 0 )"
    'SQLQ = SQLQ & "AND NOT (BF_POLICY IS NULL OR BF_POLICY = '' ) "
    SQLQ = SQLQ & "ORDER BY BF_EMPNBR, BF_BCODE, BF_EDATE "
    rsBeneOpt.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsBeneOpt.EOF Then
       totNum = rsBeneOpt.RecordCount: I = 0
    End If
    Do While Not rsBeneOpt.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
        If Not IsNull(rsBeneOpt("BF_COVER")) Then
            If UCase(rsBeneOpt("BF_COVER")) = "W" Then
                GoTo NexeRec
            End If
        End If
        If Not IsNull(rsBeneOpt("BF_CEASEDATE")) Then
            If IsDate(rsBeneOpt("BF_CEASEDATE")) Then
                If CVDate(xBillingDate) >= CVDate(rsBeneOpt("BF_CEASEDATE")) Then
                    GoTo NexeRec
                End If
            End If
        End If
        
        xEmpNbr = rsBeneOpt("BF_EMPNBR")
        xCode = rsBeneOpt("BF_BCODE")
        xPolicyNo = rsBeneOpt("BF_POLICY")
        'xOptAccount = ""
        'If Len(xPolicyNo) <> 9 Then
        '    GoTo NexeRec 'Invalid Policy number format, it's "#####-###"
        'Else
        '    xOptAccount = Mid(xPolicyNo, 7, 3)
        'End If
        SQLQ = "SELECT * FROM WFC_REPORT_WRK WHERE WRKEMP='" & glbUserID & "'"
        SQLQ = SQLQ & "AND R_EMPNBR = " & xEmpNbr & " "
        If rsWRK.State <> 0 Then rsWRK.Close
        rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsWRK.EOF Then
            rsWRK.AddNew
            rsWRK("R_COMPNO") = "001"
            rsWRK("R_EMPNBR") = xEmpNbr
            rsWRK("WRKEMP") = glbUserID
            rsWRK("R_NUM2") = 0
            rsWRK("R_NUM3") = 0
            rsWRK("R_NUM4") = 0
            rsWRK("R_NUM5") = 0
            rsWRK("R_NUM6") = 0
            rsWRK("R_NUM7") = 0
            rsWRK("R_NUM8") = 0
        End If
        'Ticket #19979 Franks 03/09/2011 Benefit codes were changed
        'If xCode = "LIF" Or xCode = "LIF1" Or xCode = "LIF2" Or xCode = "LIF3" Or xCode = "LFR" Or xCode = "LFR" Or xCode = "LFR1" Or xCode = "LFR2" Or xCode = "LIF3" Then
        If xCode = "LIF15" Or xCode = "LIF2" Or xCode = "LIF25" Or xCode = "LIFF" Or xCode = "LFR" Or xCode = "LFR1" Or xCode = "LFR2" Then
            If Not IsNull(rsBeneOpt("BF_AMT")) Then
                rsWRK("R_NUM2") = rsWRK("R_NUM2") + rsBeneOpt("BF_AMT")
            End If
        End If
        If xCode = "OPLF" Then
            If IsNull(rsBeneOpt("BF_AMT")) Then
                rsWRK("R_NUM3") = 0
            Else
                rsWRK("R_NUM3") = rsBeneOpt("BF_AMT")
            End If
        End If
        'Ticket #19979 Franks 03/09/2011 Benefit codes were changed
        'If xCode = "AD&D" Then
        If xCode = "IA" Then
            If IsNull(rsBeneOpt("BF_AMT")) Then
                rsWRK("R_NUM4") = 0
            Else
                rsWRK("R_NUM4") = rsBeneOpt("BF_AMT")
            End If
        End If
        'Ticket #19979 Franks 03/09/2011 Benefit codes were changed
        'If xCode = "STD" Then
        If xCode = "SD" Then
            If IsNull(rsBeneOpt("BF_AMT")) Then
                rsWRK("R_NUM5") = 0
            Else
                rsWRK("R_NUM5") = rsBeneOpt("BF_AMT")
            End If
        End If
        'Ticket #19979 Franks 03/09/2011 Benefit codes were changed
        'If xCode = "LTD" Then
        If xCode = "LD" Then
            If IsNull(rsBeneOpt("BF_AMT")) Then
                rsWRK("R_NUM6") = 0
            Else
                rsWRK("R_NUM6") = rsBeneOpt("BF_AMT")
            End If
        End If
        If xCode = "OPLS" Then
            If IsNull(rsBeneOpt("BF_AMT")) Then
                rsWRK("R_NUM7") = 0
            Else
                rsWRK("R_NUM7") = rsBeneOpt("BF_AMT")
            End If
        End If
        If xCode = "EHC" Then
            If Not IsNull(rsBeneOpt("BF_COVER")) Then
                rsWRK("R_TEXT1") = rsBeneOpt("BF_COVER")
            End If
        End If
        ''Ticket #19979 Franks 03/09/2011 Benefit codes were changed
        ''If xCode = "DENT" Then
        'Ticket #26779 Franks 03/09/2015 replace DN with DEN
        'If xCode = "DN" Then
        If xCode = "DEN" Then
            If Not IsNull(rsBeneOpt("BF_COVER")) Then
                rsWRK("R_TEXT2") = rsBeneOpt("BF_COVER")
            End If
        End If
        If Not IsNull(rsBeneOpt("BF_GROUP")) Then
            rsWRK("R_TEXT3") = rsBeneOpt("BF_GROUP")
        End If
            
        rsWRK.Update
NexeRec:
        rsBeneOpt.MoveNext
    Loop
    rsBeneOpt.Close
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT

End Sub

Private Sub WFCOptLifeBilling()
Dim rsBeneOpt As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim SQLQ As String, xEmpNbr, xCode
Dim I, totNum
Dim xPolicyNo, xOptAccount
Dim rsDep As New ADODB.Recordset
Dim xBillingDate
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 0
    
    xBillingDate = CVDate(ComMTH.Text & " 1, " & txtFiscal)
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute "DELETE FROM WFC_MANULIFE_BENE_WRK WHERE WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001W.CommitTrans

    HisSQL = " BF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
    SQLQ = "SELECT * FROM HRBENFT WHERE " & HisSQL & " "
    SQLQ = SQLQ & " AND (BF_EMPNBR in (select ED_EMPNBR from HREMP WHERE ED_COUNTRY = 'CANADA' AND NOT (ED_USER_TEXT1 IS NULL OR ED_USER_TEXT1 = '' ) AND NOT (ED_USER_TEXT2 IS NULL OR ED_USER_TEXT2 = '') AND NOT (ED_USER_NUM1 IS NULL) ))"
    SQLQ = SQLQ & "AND (BF_BCODE = 'OPLF' OR BF_BCODE = 'OPLS' OR BF_BCODE = 'OPLC') "
    SQLQ = SQLQ & "AND NOT (BF_POLICY IS NULL OR BF_POLICY = '' ) "
    SQLQ = SQLQ & "ORDER BY BF_EMPNBR, BF_BCODE, BF_EDATE "
    rsBeneOpt.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsBeneOpt.EOF Then
       totNum = rsBeneOpt.RecordCount: I = 0
    End If
    Do While Not rsBeneOpt.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
        If Not IsNull(rsBeneOpt("BF_COVER")) Then
            If UCase(rsBeneOpt("BF_COVER")) = "W" Then
                GoTo NexeRec
            End If
        End If
        If Not IsNull(rsBeneOpt("BF_CEASEDATE")) Then
            If IsDate(rsBeneOpt("BF_CEASEDATE")) Then
                If CVDate(xBillingDate) >= CVDate(rsBeneOpt("BF_CEASEDATE")) Then
                    GoTo NexeRec
                End If
            End If
        End If
        
        xEmpNbr = rsBeneOpt("BF_EMPNBR")
        xCode = rsBeneOpt("BF_BCODE")
        xPolicyNo = rsBeneOpt("BF_POLICY")
        xOptAccount = ""
        If Len(xPolicyNo) <> 9 Then
            GoTo NexeRec 'Invalid Policy number format, it's "#####-###"
        Else
            xOptAccount = Mid(xPolicyNo, 7, 3)
        End If
        SQLQ = "SELECT * FROM WFC_MANULIFE_BENE_WRK WHERE WRKEMP='" & glbUserID & "'"
        SQLQ = SQLQ & "AND WB_EMPNBR = " & xEmpNbr & " "
        'SQLQ = SQLQ & "AND WB_TEXT1 = '" & xOptAccount & "' "
        If rsWRK.State <> 0 Then rsWRK.Close
        rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsWRK.EOF Then
            rsWRK.AddNew
            rsWRK("WB_COMPNO") = "001"
            rsWRK("WB_EMPNBR") = xEmpNbr
            rsWRK("WB_LUSER") = glbUserID
            rsWRK("WB_LDATE") = Date
            rsWRK("WB_LTIME") = Time$
            rsWRK("WRKEMP") = glbUserID
            rsWRK("WB_TEXT1") = xOptAccount
        End If
        If xCode = "OPLF" Then
            rsWRK("WB_BCODE1") = xCode
            If IsNull(rsBeneOpt("BF_AMT")) Then
                rsWRK("WB_AMT1") = 0
            Else
                rsWRK("WB_AMT1") = rsBeneOpt("BF_AMT")
            End If
            If IsNull(rsBeneOpt("BF_ECOST")) Then
                rsWRK("WB_PREMIUM1") = 0
            Else
                rsWRK("WB_PREMIUM1") = rsBeneOpt("BF_MTHECOST")     'rsBeneOpt("BF_ECOST")  - Ticket #18936
            End If
            rsWRK("WB_TAX1") = rsWRK("WB_PREMIUM1") * 0.08
        End If
        If xCode = "OPLS" Then
            rsWRK("WB_BCODE2") = xCode
            If IsNull(rsBeneOpt("BF_AMT")) Then
                rsWRK("WB_AMT2") = 0
            Else
                rsWRK("WB_AMT2") = rsBeneOpt("BF_AMT")
            End If
            If IsNull(rsBeneOpt("BF_ECOST")) Then
                rsWRK("WB_PREMIUM2") = 0
            Else
                rsWRK("WB_PREMIUM2") = rsBeneOpt("BF_MTHECOST")     'rsBeneOpt("BF_ECOST") - Ticket #18936
            End If
            rsWRK("WB_TAX2") = rsWRK("WB_PREMIUM2") * 0.08
            'Get Spouse Sex and Smoker - Begin
            SQLQ = "SELECT * FROM HRDEPEND where DP_EMPNBR = " & xEmpNbr & " "
            SQLQ = SQLQ & "AND (DP_RELATE = 'Wife' OR DP_RELATE = 'Husband' OR DP_RELATE = 'Spouse') "
            If rsDep.State <> 0 Then rsDep.Close
            rsDep.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsDep.EOF Then
                If Not IsNull(rsDep("DP_SEX")) Then
                    rsWRK("WB_TEXT2") = rsDep("DP_SEX")
                End If
                If Not IsNull(rsDep("DP_SMOKER")) Then
                    If rsDep("DP_SMOKER") Then
                        rsWRK("WB_TEXT3") = "Y"
                    Else
                        rsWRK("WB_TEXT3") = "N"
                    End If
                End If
                If IsDate(rsDep("DP_DOB")) Then 'Ticket #14364
                    rsWRK("WB_DATE") = rsDep("DP_DOB")
                End If
            End If

            rsDep.Close
            'Get Spouse Sex and Smoker - End
        End If
        If xCode = "OPLC" Then
            rsWRK("WB_BCODE3") = xCode
            If IsNull(rsBeneOpt("BF_AMT")) Then
                rsWRK("WB_AMT3") = 0
            Else
                rsWRK("WB_AMT3") = rsBeneOpt("BF_AMT")
            End If
            If IsNull(rsBeneOpt("BF_ECOST")) Then
                rsWRK("WB_PREMIUM3") = 0
            Else
                rsWRK("WB_PREMIUM3") = rsBeneOpt("BF_MTHECOST")  'rsBeneOpt("BF_ECOST") - Ticket #18936
            End If
            rsWRK("WB_TAX3") = rsWRK("WB_PREMIUM3") * 0.08
        End If
        rsWRK.Update
NexeRec:
        rsBeneOpt.MoveNext
    Loop
    rsBeneOpt.Close
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT
End Sub

Private Sub HCAS_Request_Report()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsAttend As New ADODB.Recordset
    Dim rsRequest As New ADODB.Recordset
    Dim SQLQ As String
    Dim WRKSQL As String
    Dim I, totNum, xVacAcc
    Dim xEmpNo

    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 0
    
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute "DELETE FROM HR_REQUEST_RPT WHERE REQ_WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001W.CommitTrans

    rsRequest.Open "SELECT * FROM HR_REQUEST_RPT WHERE REQ_WRKEMP='" & glbUserID & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    WRKSQL = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
    SQLQ = "SELECT  AD_EMPNBR, MONTH(AD_DOA) AS MNTH,"
    SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='OT' THEN AD_HRS ELSE 0 END) AS OTEARN,"
    SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='CT' THEN AD_HRS ELSE 0 END) AS CTTAKEN,"
    SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,3)='VAC' THEN AD_HRS ELSE 0 END) AS VACTAKEN"
    SQLQ = SQLQ & " FROM HR_ATTENDANCE"
    SQLQ = SQLQ & " WHERE AD_DOA >= " & Date_SQL(dlpDateRange(0).Text) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text)
    SQLQ = SQLQ & " AND " & WRKSQL
    SQLQ = SQLQ & " GROUP BY AD_EMPNBR,MONTH(AD_DOA)"
    SQLQ = SQLQ & " ORDER BY AD_EMPNBR, MONTH(AD_DOA)"
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsAttend.EOF Then
        totNum = rsAttend.RecordCount: I = 0
        rsAttend.MoveFirst
    End If
    
    xEmpNo = 0
    Do While Not rsAttend.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
            
        
        'Adding records to request report table
        If xEmpNo <> rsAttend("AD_EMPNBR") Then
        
            'Retrieve Vacation Accrued for the year
            Set rsHREmp = Nothing
            SQLQ = "SELECT ED_EMPNBR, ED_VAC, ED_EFDATE, ED_ETDATE FROM HREMP WHERE ED_EMPNBR = " & rsAttend("AD_EMPNBR")
            'SQLQ = SQLQ & " AND ED_EFDATE <= " & " AND ED_ETDATE >="
            rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsHREmp.EOF Then
                xVacAcc = rsHREmp("ED_VAC")
            Else
                xVacAcc = 0
            End If
            rsHREmp.Close
        
            rsRequest.AddNew
            xEmpNo = rsAttend("AD_EMPNBR")
        End If
        
        rsRequest("EMPNBR") = rsAttend("AD_EMPNBR")
        If rsAttend("MNTH") = 1 Then
            rsRequest("OT_ACC_JAN") = rsAttend("OTEARN")
            rsRequest("OT_TKN_JAN") = rsAttend("CTTAKEN")
            rsRequest("VAC_TKN_JAN") = rsAttend("VACTAKEN")
        End If
        
        If rsAttend("MNTH") = 2 Then
            rsRequest("OT_ACC_FEB") = rsAttend("OTEARN")
            rsRequest("OT_TKN_FEB") = rsAttend("CTTAKEN")
            rsRequest("VAC_TKN_FEB") = rsAttend("VACTAKEN")
        End If
        
        If rsAttend("MNTH") = 3 Then
            rsRequest("OT_ACC_MAR") = rsAttend("OTEARN")
            rsRequest("OT_TKN_MAR") = rsAttend("CTTAKEN")
            rsRequest("VAC_TKN_MAR") = rsAttend("VACTAKEN")
        End If
        
        If rsAttend("MNTH") = 4 Then
            rsRequest("OT_ACC_APR") = rsAttend("OTEARN")
            rsRequest("OT_TKN_APR") = rsAttend("CTTAKEN")
            rsRequest("VAC_TKN_APR") = rsAttend("VACTAKEN")
        End If
        
        If rsAttend("MNTH") = 5 Then
            rsRequest("OT_ACC_MAY") = rsAttend("OTEARN")
            rsRequest("OT_TKN_MAY") = rsAttend("CTTAKEN")
            rsRequest("VAC_TKN_MAY") = rsAttend("VACTAKEN")
        End If
        
        If rsAttend("MNTH") = 6 Then
            rsRequest("OT_ACC_JUN") = rsAttend("OTEARN")
            rsRequest("OT_TKN_JUN") = rsAttend("CTTAKEN")
            rsRequest("VAC_TKN_JUN") = rsAttend("VACTAKEN")
        End If
        
        If rsAttend("MNTH") = 7 Then
            rsRequest("OT_ACC_JUL") = rsAttend("OTEARN")
            rsRequest("OT_TKN_JUL") = rsAttend("CTTAKEN")
            rsRequest("VAC_TKN_JUL") = rsAttend("VACTAKEN")
        End If
        
        If rsAttend("MNTH") = 8 Then
            rsRequest("OT_ACC_AUG") = rsAttend("OTEARN")
            rsRequest("OT_TKN_AUG") = rsAttend("CTTAKEN")
            rsRequest("VAC_TKN_AUG") = rsAttend("VACTAKEN")
        End If
        
        If rsAttend("MNTH") = 9 Then
            rsRequest("OT_ACC_SEP") = rsAttend("OTEARN")
            rsRequest("OT_TKN_SEP") = rsAttend("CTTAKEN")
            rsRequest("VAC_TKN_SEP") = rsAttend("VACTAKEN")
        End If
        
        If rsAttend("MNTH") = 10 Then
            rsRequest("OT_ACC_OCT") = rsAttend("OTEARN")
            rsRequest("OT_TKN_OCT") = rsAttend("CTTAKEN")
            rsRequest("VAC_TKN_OCT") = rsAttend("VACTAKEN")
        End If
        
        If rsAttend("MNTH") = 11 Then
            rsRequest("OT_ACC_NOV") = rsAttend("OTEARN")
            rsRequest("OT_TKN_NOV") = rsAttend("CTTAKEN")
            rsRequest("VAC_TKN_NOV") = rsAttend("VACTAKEN")
        End If
        
        If rsAttend("MNTH") = 12 Then
            rsRequest("OT_ACC_DEC") = rsAttend("OTEARN")
            rsRequest("OT_TKN_DEC") = rsAttend("CTTAKEN")
            rsRequest("VAC_TKN_DEC") = rsAttend("VACTAKEN")
        End If
        
        
        If xVacAcc <> 0 And xVacAcc <> "" Then
            rsRequest("VAC_ACC_JAN") = xVacAcc / 12
            rsRequest("VAC_ACC_FEB") = xVacAcc / 12
            rsRequest("VAC_ACC_MAR") = xVacAcc / 12
            rsRequest("VAC_ACC_APR") = xVacAcc / 12
            rsRequest("VAC_ACC_MAY") = xVacAcc / 12
            rsRequest("VAC_ACC_JUN") = xVacAcc / 12
            rsRequest("VAC_ACC_JUL") = xVacAcc / 12
            rsRequest("VAC_ACC_AUG") = xVacAcc / 12
            rsRequest("VAC_ACC_SEP") = xVacAcc / 12
            rsRequest("VAC_ACC_OCT") = xVacAcc / 12
            rsRequest("VAC_ACC_NOV") = xVacAcc / 12
            rsRequest("VAC_ACC_DEC") = xVacAcc / 12
        End If
        
        rsRequest("REQ_WRKEMP") = glbUserID
        rsAttend.MoveNext
        
        If rsAttend.EOF Then
            rsRequest.Update
        Else
            If xEmpNo <> rsAttend("AD_EMPNBR") Then
                rsRequest.Update
            End If
        End If
    Loop
    
    rsAttend.Close
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT
    
End Sub

Private Sub Cri_Dates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%
    If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
        'Collectcorp Inc. - Ticket #14437
        If InStr(1, fglbFileName, "SN2390_LicenseAddr.rpt") > 0 Then
            TempCri = "({HREMP.ED_DOH} "
        ElseIf InStr(1, fglbFileName, "SN2390_LicenseTerm.rpt") > 0 Then
            TempCri = "({TERM_HRTRMEMP.TERM_DOT} "
        ElseIf InStr(1, fglbFileName, "SN2390_LicenseAdditions.rpt") > 0 Then
            TempCri = "({HR_USERDEFINE_TABLE.UD_DATE1} "
        End If
        
        dtYYY% = Year(dlpDateRange(2).Text)
        dtMM% = month(dlpDateRange(2).Text)
        dtDD% = Day(dlpDateRange(2).Text)
        TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
        dtYYY% = Year(dlpDateRange(3).Text)
        dtMM% = month(dlpDateRange(3).Text)
        dtDD% = Day(dlpDateRange(3).Text)
        TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
    
    For X% = 2 To 3
        If Len(dlpDateRange(X).Text) > 0 Then
            'Collectcorp Inc. - Ticket #14437
            If InStr(1, fglbFileName, "SN2390_LicenseAddr.rpt") > 0 Then
                TempCri = "({HREMP.ED_DOH} "
            ElseIf InStr(1, fglbFileName, "SN2390_LicenseTerm.rpt") > 0 Then
                TempCri = "({TERM_HRTRMEMP.TERM_DOT} "
            ElseIf InStr(1, fglbFileName, "SN2390_LicenseAdditions.rpt") > 0 Then
                TempCri = "({HR_USERDEFINE_TABLE.UD_DATE1} "
            End If
            
            If X% = 2 Then
                TempCri = TempCri & " >= "
            Else
                TempCri = TempCri & " <= "
            End If
            dtYYY% = Year(dlpDateRange(X).Text)
            dtMM% = month(dlpDateRange(X).Text)
            dtDD% = Day(dlpDateRange(X).Text)
            TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
            GoTo Cri_FTDatst
        End If
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

Private Sub Cri_LicDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%
    If Len(dlpDateRange(4).Text) > 0 And Len(dlpDateRange(5).Text) > 0 Then
        'Collectcorp Inc. - Ticket #14437
       
        TempCri = "({HR_USERDEFINE_TABLE.UD_DATE2} "
        
        
        
        TempCri = TempCri & " >= " & Date_SQL(dlpDateRange(4).Text) & " and {HR_USERDEFINE_TABLE.UD_DATE2} <= " & Date_SQL(dlpDateRange(5).Text) & ")"
       
        GoTo Cri_LicFTDatst
    End If
    
        
Cri_LicFTDatst:
    If Len(TempCri) >= 1 Then
        If Not glbiOneWhere Then
            glbstrSelCri = TempCri
        Else
            glbstrSelCri = glbstrSelCri & " AND " & TempCri
        End If
        glbiOneWhere = True
    End If

End Sub

Private Sub Cri_License_Codes(intIdx%)
Dim CodeCri As String
Dim countr   As Integer
Dim strCd$

If Len(clpUser(intIdx%)) > 0 Then
    If InStr(1, fglbFileName, "SN2390_LicenseTerm.rpt") > 0 Then
        Select Case intIdx%
            Case 0: strCd$ = "TERM_USERDEFINE_TABLE.UD_CODE1"
            Case 1: strCd$ = "TERM_USERDEFINE_TABLE.UD_CODE2"
        End Select
    Else
        Select Case intIdx%
            Case 0: strCd$ = "HR_USERDEFINE_TABLE.UD_CODE1"
            Case 1: strCd$ = "HR_USERDEFINE_TABLE.UD_CODE2"
            Case 2: strCd$ = "HR_USERDEFINE_TABLE.UD_CODE1"
        End Select
    End If
    If intIdx% = 2 Then
        CodeCri = "(Upper({" & strCd$ & "}) in  ['" & Replace(clpUser(intIdx%).Text, ",", "','") & "'])"
    Else
        CodeCri = "(Ucase({" & strCd$ & "}) in  ['" & Replace(clpUser(intIdx%).Text, ",", "','") & "'])"
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

Private Sub Cri_BirthMonth()
Dim EECri As String, OneSet%, X%
Dim strMonth As String

If Len(comDateMonth.Text) < 1 Then Exit Sub

Select Case comDateMonth.Text
    Case "January": strMonth = "1"
    Case "February": strMonth = "2"
    Case "March": strMonth = "3"
    Case "April": strMonth = "4"
    Case "May": strMonth = "5"
    Case "June": strMonth = "6"
    Case "July": strMonth = "7"
    Case "August": strMonth = "8"
    Case "September": strMonth = "9"
    Case "October": strMonth = "10"
    Case "November": strMonth = "11"
    Case "December": strMonth = "12"
End Select
    
'EECri = "{@BirthMonth}= " & Val(intMonth)
EECri = "Month({HREMP.ED_DOB})= " & Val(strMonth)

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_DOHMonth()
Dim EECri As String, OneSet%, X%
Dim strMonth As String

If Len(comDateMonth.Text) < 1 Then Exit Sub

Select Case comDateMonth.Text
    Case "January": strMonth = "1"
    Case "February": strMonth = "2"
    Case "March": strMonth = "3"
    Case "April": strMonth = "4"
    Case "May": strMonth = "5"
    Case "June": strMonth = "6"
    Case "July": strMonth = "7"
    Case "August": strMonth = "8"
    Case "September": strMonth = "9"
    Case "October": strMonth = "10"
    Case "November": strMonth = "11"
    Case "December": strMonth = "12"
End Select
    
'EECri = "{@DOHMonth}= " & Val(strMonth)
EECri = "Month({HREMP.ED_DOH})= " & Val(strMonth)

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_Position()
Dim PosCri As String

If Len(clpJob.Text) <= 0 Then Exit Sub

PosCri = "({HR_JOB_HISTORY.JH_JOB} IN ['" & Replace(clpJob.Text, ",", "','") & "'])"

If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & PosCri
Else
    glbstrSelCri = PosCri
End If

End Sub

Private Sub Vacation_Report_XLS_HCAS()
    On Error GoTo Vacation_Report_XLS_HCAS_Err

    Dim rsHREmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    Dim xHourlyRate
    Dim xTotVacOut

    'Check the Month End Date is within the Current Vacation Year
    If Len(Trim(dlpAsOf.Text)) = 0 Then
        MsgBox "Month Ending Date cannot be blank"
        dlpAsOf.SetFocus
    ElseIf Not IsDate(dlpAsOf.Text) Then
        MsgBox "Invalid Month Ending Date"
        dlpAsOf.SetFocus
    End If
    
    sSQLQ = Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")")
    SQLQ = "SELECT HREMP.ED_EMPNBR, HREMP.ED_FNAME AS EFNAME, HREMP.ED_SURNAME AS ESURNAME, HREMP.ED_EFDATE, HREMP.ED_ETDATE, "
    SQLQ = SQLQ & " (CASE WHEN HREMP.ED_PVAC IS NULL THEN 0 ELSE HREMP.ED_PVAC END) + "
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AC_HRS) FROM HR_ACCRUAL WHERE AC_EMPNBR = HREMP.ED_EMPNBR AND AC_TYPE = 'VAC' AND AC_ACTION = 'U' AND AC_EDATE >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AC_EDATE <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AC_HRS) FROM HR_ACCRUAL WHERE AC_EMPNBR = HREMP.ED_EMPNBR AND AC_TYPE = 'VAC' AND AC_ACTION = 'U' AND AC_EDATE >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AC_EDATE <=" & Date_SQL(dlpAsOf.Text) & ") END) - "
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,3) = 'VAC' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,3) = 'VAC' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") END) AS ED_VACOUTS, "
    SQLQ = SQLQ & " JH_JOB, JB_DESCR, JH_REPTAU, SH_SALARY, HREMP.ED_PVAC,JH_DHRS,JH_WHRS,SH_SALCD,SUPER.ED_SURNAME AS SSURNAME,SUPER.ED_FNAME AS SFNAME,"
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AC_HRS) FROM HR_ACCRUAL WHERE AC_EMPNBR = HREMP.ED_EMPNBR AND AC_TYPE = 'VAC' AND AC_ACTION = 'U' AND AC_EDATE >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AC_EDATE <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AC_HRS) FROM HR_ACCRUAL WHERE AC_EMPNBR = HREMP.ED_EMPNBR AND AC_TYPE = 'VAC' AND AC_ACTION = 'U' AND AC_EDATE >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AC_EDATE <=" & Date_SQL(dlpAsOf.Text) & ") END) AS VAC,"
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,3) = 'VAC' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,3) = 'VAC' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") END) AS VACT"
    SQLQ = SQLQ & " FROM ((((HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR AND JH_CURRENT <>0)"
    SQLQ = SQLQ & " INNER JOIN HR_SALARY_HISTORY ON HR_SALARY_HISTORY.SH_EMPNBR = ED_EMPNBR AND SH_CURRENT <>0)"
    SQLQ = SQLQ & " INNER JOIN HRJOB ON HRJOB.JB_CODE = HR_JOB_HISTORY.JH_JOB)"
    SQLQ = SQLQ & " INNER JOIN HREMP SUPER ON SUPER.ED_EMPNBR = HR_JOB_HISTORY.JH_REPTAU)"
    SQLQ = SQLQ & " GROUP BY HREMP.ED_EMPNBR,HREMP.ED_FNAME,HREMP.ED_SURNAME,HREMP.ED_EFDATE,HREMP.ED_ETDATE,JH_JOB,JB_DESCR,JH_REPTAU,SH_SALARY,HREMP.ED_PVAC,JH_DHRS,JH_WHRS,SH_SALCD,SUPER.ED_FNAME,SUPER.ED_SURNAME"
    SQLQ = SQLQ & " ORDER BY SSURNAME, SFNAME, ED_VACOUTS ASC"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: I = 0
        rsHREmp.MoveFirst

        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "VacationRptTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "VacationRpt" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
    
        exSheet.Cells(1, 2) = Format(Now, "mmm dd, yyyy")
        exSheet.Cells(2, 2) = Time$
        exSheet.Cells(5, 1) = "Report for the Month Ending: " & Format(dlpAsOf.Text, "dd-mmm-yy")
        
        xTotVacOut = 0
        xRow = 11
        'Columns: 1 - Name, 3 - Job Title, 4 - Supervisor, 5 - Previous Vac, 6 - Current Vac, 7 - Taken Vac, 8 - Outstanding Vac, 9 - Salary, 10 - Cost of Oustanding Vacation
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            xHourlyRate = 0
            If rsHREmp("SH_SALCD") = "H" Then
                xHourlyRate = rsHREmp("SH_SALARY")
            ElseIf rsHREmp("SH_SALCD") = "A" And rsHREmp("JH_WHRS") <> 0 Then
                xHourlyRate = Round((rsHREmp("SH_SALARY") / 52) / rsHREmp("JH_WHRS"), 2)
            End If
            
            exSheet.Cells(xRow, 1) = rsHREmp("ESURNAME") & ", " & rsHREmp("EFNAME")
            exSheet.Cells(xRow, 3) = rsHREmp("JB_DESCR")
            'If rsHREmp("JH_REPTAU") <> "" And Not IsNull(rsHREmp("JH_REPTAU")) Then
                'exSheet.Cells(xRow, 4) = GetEmpData(rsHREmp("JH_REPTAU"), "ED_SURNAME") & ", " & GetEmpData(rsHREmp("JH_REPTAU"), "ED_FNAME")
                exSheet.Cells(xRow, 4) = rsHREmp("SSURNAME") & ", " & rsHREmp("SFNAME")
            'End If
            If rsHREmp("JH_DHRS") <> 0 And Not IsNull(rsHREmp("JH_DHRS")) Then
                exSheet.Cells(xRow, 5) = Round(rsHREmp("ED_PVAC") / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 6) = Round(rsHREmp("VAC") / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 7) = Round(rsHREmp("VACT") / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 8) = Round(rsHREmp("ED_VACOUTS") / rsHREmp("JH_DHRS"), 2)
                xTotVacOut = xTotVacOut + Round(rsHREmp("ED_VACOUTS") / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 9) = Format(Round((rsHREmp("ED_VACOUTS") / rsHREmp("JH_DHRS")), 2) * (xHourlyRate * rsHREmp("JH_DHRS")), "#,##0.00")
            End If
            exSheet.Cells(xRow, 10) = Format(xHourlyRate, "#,##0.00")
            
            
            rsHREmp.MoveNext
            xRow = xRow + 1
        Loop
        
        exSheet.Cells(xRow + 2, 1) = "Total Number of Employees Reported: " & totNum
        exSheet.Cells(xRow + 3, 1) = "Total Number of Days of Vacation Outstanding as at Current Month End: " & xTotVacOut
        exSheet.Rows(xRow + 2).Font.Bold = True
        exSheet.Rows(xRow + 3).Font.Bold = True
        
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    rsHREmp.Close
    
Exit Sub

Vacation_Report_XLS_HCAS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
Resume Next

End Sub
Private Sub WFC_WPS_Training_Rpt_WPSCODE() 'Ticket #21330 Franks 02/14/2012
    On Error GoTo WFC_WPS_Training_Report_Err
    Dim rsEmp As New ADODB.Recordset
    Dim rsTermEmp As New ADODB.Recordset
    Dim rsEdu As New ADODB.Recordset
    Dim rsWRK As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rsDiv As New ADODB.Recordset
    Dim rsCourse As New ADODB.Recordset
    Dim rsCrsMaster As New ADODB.Recordset
    Dim rsEMPHISWRK As New ADODB.Recordset
    Dim rsTmpWrk As New ADODB.Recordset
    'Dim exApp As Excel.Application
    'Dim exBook As Excel.Workbook
    'Dim exSheet As Excel.Worksheet
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim SQLQ, sSQLQ As String
    Dim locSQL As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim xCol As Long
    Dim I, J, K, M, totNum
    Dim xHourlyRate
    Dim xTotVacOut
    Dim setcolumnwidth As Boolean
    Dim rsEmpPos As New ADODB.Recordset
    Dim xTERM_SEQ
    Dim xStartLine As Integer
    Dim xStartColu As Integer
    Dim xPlant As String
    Dim xDiv As String
    Dim xSalHly As String
    Dim xStr As String
    Dim xTotStaff
    Dim xTotPlant
    Dim xPercen As Double
    Dim xPerStr As String
    Dim xCunCourse As Double 'Integer
    Dim xCunLoc As Double 'Integer
    Dim xTotLoc As Double 'Integer
    Dim xDivCourseCun As Double 'Integer
    Dim xDivSQL As String
    Dim xRptSele As String
    Dim xCountFlag As Boolean
    Dim xWPSCode As String
    
    'Note:
    'This program will add all data into a temp table based on the Plant/Location
    'create Plant/Location columns into a table
    'create WPS Course Code columns into a table
    'Write data into Excel file
    
    'Delete Temp tables
    SQLQ = "DELETE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "DELETE FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ
    'Ticket #21330 keep the exclude course code with the same WPS codes
    SQLQ = "DELETE FROM HREMPHIS_WRK WHERE EE_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ

    'open Edu recordset
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    xRptSele = sSQLQ
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE (1=1) "
    SQLQ = SQLQ & "AND " & sSQLQ & " "
    'Employee Status is Active on Employee Status Matrix of WFC Custom Feature
    SQLQ = SQLQ & "AND ED_EMP IN (SELECT  EP_CODE from WFC_HRST_EMPSTATUS WHERE NOT (EP_ACTIVE_FLAG = 0)) "
    locSQL = "(" & SQLQ & ")"
    
    'SQLQ = "SELECT * from HREDSEM"
    'No duplicate course, use DISTINCT
    SQLQ = "SELECT DISTINCT ES_EMPNBR, ES_CRSCODE from HREDSEM"
    SQLQ = SQLQ & " WHERE (1=1) "
    'only employees in the selection scritera
    SQLQ = SQLQ & "AND ES_EMPNBR IN " & locSQL & " "
    'only Course Codes for WPS flag checked
    SQLQ = SQLQ & "AND ES_CRSCODE IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ESCD' AND (NOT TB_LEFLAG = 0)) "
    If IsDate(dlpDateRange(0).Text) Then
        SQLQ = SQLQ & "AND ES_DATCOMP >= " & Date_SQL(dlpDateRange(0).Text) & " "
    End If
    If IsDate(dlpDateRange(1).Text) Then
        SQLQ = SQLQ & "AND ES_DATCOMP <= " & Date_SQL(dlpDateRange(1).Text) & " "
    End If
    SQLQ = SQLQ & " ORDER BY ES_EMPNBR, ES_CRSCODE " 'ES_EMPNBR, ES_CTYPE ASC, ES_DATCOMP DESC "

    'Total = 0
    rsEdu.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    'Margaret asked: if no Edu record found than still show these plants
    'If rsEdu.EOF Then
    '    MsgBox "No record found in this Selection Criteria."
    '    Exit Sub
    'End If
    If Not rsEdu.EOF Then
        rsEdu.MoveFirst
        totNum = rsEdu.RecordCount: I = 0
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
        Screen.MousePointer = HOURGLASS
    
        'Ticket #21330 Franks 02/14/2012 - begin
        'If one employee has multiple courses with same WPS Codes, only count one
        'check which course should be exlcuded
        Do While Not rsEdu.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            'find the match WPS Code
            SQLQ = "SELECT ES_CRSCODE, ES_WPSCODE FROM HR_COURSECODE_MASTER WHERE ES_CRSCODE = '" & rsEdu("ES_CRSCODE") & "' "
            SQLQ = SQLQ & "AND NOT (ES_WPSCODE IS NULL) "
            If rsCrsMaster.State <> 0 Then rsCrsMaster.Close
            rsCrsMaster.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsCrsMaster.EOF Then
                'rsCrsMaster ES_WPSCODE
                If Len(rsCrsMaster("ES_WPSCODE")) > 0 Then
                    SQLQ = "SELECT * FROM HREMPHIS_WRK WHERE EE_EMPNBR = " & rsEdu("ES_EMPNBR") & " "
                    SQLQ = SQLQ & "AND EE_WRKEMP = '" & glbUserID & "' "
                    SQLQ = SQLQ & "AND NOT EE_HISTYPE = '" & rsEdu("ES_CRSCODE") & "' "
                    SQLQ = SQLQ & "AND EE_OLDVALUE = '" & rsCrsMaster("ES_WPSCODE") & "' "
                    'SQLQ = SQLQ & "SELECT * FROM HREMPHIS_WRK WRKEMP = '" & glbUserID & "' "
                    If rsEMPHISWRK.State <> 0 Then rsEMPHISWRK.Close
                    rsEMPHISWRK.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
                    If Not rsEMPHISWRK.EOF Then
                        'found same WPS Code but diff course code, should not count
                        xCountFlag = False
                    Else
                        xCountFlag = True
                    End If
                    'alway add new, if found then this code should exclude; not found count it in the count section
                    rsEMPHISWRK.AddNew
                    rsEMPHISWRK("EE_EMPNBR") = rsEdu("ES_EMPNBR")
                    rsEMPHISWRK("EE_HISTYPE") = rsEdu("ES_CRSCODE")
                    rsEMPHISWRK("EE_OLDVALUE") = rsCrsMaster("ES_WPSCODE")
                    If xCountFlag Then
                        rsEMPHISWRK("EE_SALCD") = "Y"
                    Else
                        rsEMPHISWRK("EE_SALCD") = "N"
                    End If
                    rsEMPHISWRK("EE_WRKEMP") = glbUserID
                    rsEMPHISWRK.Update
                End If
            End If
            rsEdu.MoveNext
        Loop
        rsEdu.MoveFirst
        'Ticket #21330 Franks 02/14/2012 - end
    End If
    
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC WPS Training Tmp.xls"
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC WPS Training(" & Trim(glbUserID) & ").xls"

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    FileCopy xlsFileTmp, xlsFileMat
    I = 0
    Do While Not rsEdu.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents

        'open Employee table
        SQLQ = "SELECT ED_EMPNBR, ED_SECTION, ED_DIV, ED_ORG FROM HREMP WHERE ED_EMPNBR = " & rsEdu("ES_EMPNBR")
        If rsEmp.State <> 0 Then rsEmp.Close
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        xPlant = ""
        xDiv = ""
        xSalHly = "Hourly"
        If Not rsEmp.EOF Then
            If Not IsNull(rsEmp("ED_SECTION")) Then
                xPlant = rsEmp("ED_SECTION")
            End If
            If Not IsNull(rsEmp("ED_DIV")) Then
                xDiv = rsEmp("ED_DIV")
            End If
            If Not IsNull(rsEmp("ED_ORG")) Then
                If rsEmp("ED_ORG") = "NONE" Or rsEmp("ED_ORG") = "EXEC" Then
                    xSalHly = "Salaried"
                End If
            End If
        End If
        rsEmp.Close
        If xPlant = "" Then GoTo next_rec
        If xDiv = "" Then GoTo next_rec
        
        'Plant - TT_PT
        'Div - TT_EMP 'Location
        'Course Code- TT_LANG1
        'Staff count - TT_WHRS (Staff:Salaried; Plant:Hourly)
        'Plant count - TT_DHRS (Staff:Salaried; Plant:Hourly)
        'check if it need to skip based on the WPS Code
        xCountFlag = getEmpCrsCount(rsEdu("ES_EMPNBR"), rsEdu("ES_CRSCODE")) 'Ticket #21330

        If xCountFlag Then
            'Ticket #21330
            'Get the WPS Code if it setup
            xWPSCode = getWPSCodeFromCrsCode(rsEdu("ES_CRSCODE"))
            
            SQLQ = "SELECT TT_WRKEMP, TT_PT,TT_EMP,TT_LANG1,TT_WHRS,TT_DHRS, TT_SEDATE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
            'SQLQ = SQLQ & "AND TT_PT = '" & xPlant & "' " 'not use
            SQLQ = SQLQ & "AND TT_EMP = '" & xDiv & "' "
            'SQLQ = SQLQ & "AND TT_LANG1 = '" & rsEdu("ES_CRSCODE") & "' "
            SQLQ = SQLQ & "AND TT_LANG1 = '" & xWPSCode & "' "
            If rsWRK.State <> 0 Then rsWRK.Close
            rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsWRK.EOF Then
                rsWRK.AddNew
                rsWRK("TT_WRKEMP") = glbUserID
                rsWRK("TT_PT") = xPlant
                rsWRK("TT_EMP") = xDiv
                rsWRK("TT_LANG1") = xWPSCode 'rsEdu("ES_CRSCODE")
                rsWRK("TT_WHRS") = 0
                rsWRK("TT_DHRS") = 0
                rsWRK("TT_SEDATE") = Date
            End If
            If xSalHly = "Salaried" Then
                rsWRK("TT_WHRS") = rsWRK("TT_WHRS") + 1
            End If
            If xSalHly = "Hourly" Then
                rsWRK("TT_DHRS") = rsWRK("TT_DHRS") + 1
            End If
            rsWRK.Update
        End If
        
next_rec:
        rsEdu.MoveNext
    Loop
    rsEdu.Close
    If rsWRK.State <> 0 Then rsWRK.Close
    
    'Margaret asked: if no Edu record found than still show these plants
    'Save Div into Temp table:
    '--Div Desc; Staff count; Plant count
    SQLQ = "SELECT * FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' AND JH_JOB = 'DIV'"
    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
    'SQLQ = "SELECT DISTINCT TT_PT,TT_EMP FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' ORDER BY TT_EMP"
    ''SELECT JH_JOB,JH_COMMENT,JH_COMMENT2,JH_REPTAU2,JH_REPTAU3,WRKEMP FROM HRJOBWRK WHERE WRKEMP = '3142'
    'rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    'Do While Not rsTemp.EOF 'JH_LUSER
    '    rsWRK.AddNew
    '    rsWRK("WRKEMP") = glbUserID
    '    rsWRK("JH_CURRENT") = 0
    '    rsWRK("JH_JOB") = "DIV"
    '    rsWRK("JH_COMMENT") = rsTemp("TT_EMP")
    '    xStr = getDivDesc(rsTemp("TT_EMP"))
    '    rsWRK("JH_COMMENT2") = Left(Trim(xStr), 50) 'Div Desc
    '    xStr = GetTABLDesc("EDSE", rsTemp("TT_PT"))
    '    rsWRK("JH_LUSER") = Left(Trim(xStr), 25) 'Plant Desc
    '    rsWRK("JH_REPTAU2") = GetEmpCountBySele(rsTemp("TT_EMP"), "Salaried") 'Staff count
    '    rsWRK("JH_REPTAU3") = GetEmpCountBySele(rsTemp("TT_EMP"), "Hourly") 'Plant count
    '    rsWRK("JH_LDATE") = Date
    '    rsWRK("JH_LTIME") = Time$
    '    rsWRK.Update
    '    rsTemp.MoveNext
    'Loop
    'rsTemp.Close
    'rsWRK.Close
    xDivSQL = "SELECT DISTINCT ED_SECTION, ED_DIV FROM HREMP WHERE (1=1) "
    xDivSQL = xDivSQL & "AND " & sSQLQ & " "
    xDivSQL = xDivSQL & "AND ED_EMP IN (SELECT  EP_CODE from WFC_HRST_EMPSTATUS WHERE NOT (EP_ACTIVE_FLAG = 0)) "
    SQLQ = xDivSQL
    'SELECT JH_JOB,JH_COMMENT,JH_COMMENT2,JH_REPTAU2,JH_REPTAU3,WRKEMP FROM HRJOBWRK WHERE WRKEMP = '3142'
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsTemp.EOF 'JH_LUSER
        rsWRK.AddNew
        rsWRK("WRKEMP") = glbUserID
        rsWRK("JH_CURRENT") = 0
        rsWRK("JH_JOB") = "DIV"
        rsWRK("JH_COMMENT") = rsTemp("ED_DIV")
        xStr = getDivDesc(rsTemp("ED_DIV"))
        rsWRK("JH_COMMENT2") = Left(Trim(xStr), 50) 'Div Desc
        xStr = GetTABLDesc("EDSE", rsTemp("ED_SECTION"))
        rsWRK("JH_LUSER") = Left(Trim(xStr), 25) 'Plant Desc
        rsWRK("JH_REPTAU2") = GetEmpCountBySele(rsTemp("ED_DIV"), xRptSele, "Salaried") 'Staff count
        rsWRK("JH_REPTAU3") = GetEmpCountBySele(rsTemp("ED_DIV"), xRptSele, "Hourly") 'Plant count
        rsWRK("JH_LDATE") = Date
        rsWRK("JH_LTIME") = Time$
        rsWRK.Update
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    rsWRK.Close
    
        
    'Save Course Codes into Temp table:
    SQLQ = "SELECT * FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' AND JH_JOB = 'COURSE'"
    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
    SQLQ = "SELECT DISTINCT TT_LANG1 FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' ORDER BY TT_LANG1"
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsTemp.EOF
        xStr = GetTABLDesc("ESCD", rsTemp("TT_LANG1"))
        rsWRK.AddNew
        rsWRK("WRKEMP") = glbUserID
        rsWRK("JH_CURRENT") = 0
        rsWRK("JH_JOB") = "COURSE"
        rsWRK("JH_COMMENT") = rsTemp("TT_LANG1")
        rsWRK("JH_COMMENT2") = Left(Trim(xStr), 50) 'Course Desc
        rsWRK("JH_LDATE") = Date
        rsWRK("JH_LTIME") = Time$
        rsWRK.Update
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    rsWRK.Close
    
    'Populate Excel file - begin
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)
            
    exSheet.Cells(2, 2) = "Status of WPS Training for " & Year(Date)
    exSheet.Cells(3, 2) = "(" & Format(Date, "MMM dd, YYYY") & ")"
    'Less than 75% have been trained
    'exSheet.Cells(4, 4).Interior.Color = RED
    '75% and more have been trained
    'exSheet.Cells(5, 4).Interior.Color = RGB(34, 139, 34) 'Dark Green
    
    'Ticket #22481 Franks 08/27/2012 - begin
    exSheet.Cells(1, 1) = "Selection Criteria:"
    If Len(clpCode(2).Text) > 0 Then
        exSheet.Cells(3, 1) = "Status Codes"
        exSheet.Cells(4, 1) = clpCode(2).Text
    End If
    If IsDate(dlpDateRange(0).Text) Or IsDate(dlpDateRange(1).Text) Then
        xStr = "From/To Dates:"
        If IsDate(dlpDateRange(0).Text) Then
            xStr = xStr & " " & dlpDateRange(0).Text
        End If
        If IsDate(dlpDateRange(1).Text) Then
            xStr = xStr & " to " & dlpDateRange(1).Text
        End If
        exSheet.Cells(6, 1) = xStr
    End If
    'Ticket #22481 Franks 08/27/2012 - end
    
    'First line of data
    xStartLine = 10
    
    'Division and Employee columns - begin
    'SQLQ = "SELECT * FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' AND JH_JOB = 'DIV' ORDER BY JH_COMMENT2"
    SQLQ = "SELECT * FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' AND JH_JOB = 'DIV' ORDER BY JH_LUSER,JH_COMMENT2"
    
    If rsDiv.State <> 0 Then rsDiv.Close
    rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsDiv.EOF Then
        totNum = rsDiv.RecordCount: I = 0
    End If
    xRow = xStartLine
    xTotStaff = 0
    xTotPlant = 0
    Do While Not rsDiv.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
        'Plant/Division
        exSheet.Cells(xRow, 1) = rsDiv("JH_LUSER") & "/" & rsDiv("JH_COMMENT2") '"Woodbridge Corporate" ' rsDiv("JH_COMMENT2")
        exSheet.Cells(xRow, 2) = rsDiv("JH_REPTAU2") 'Staff
        xTotStaff = xTotStaff + rsDiv("JH_REPTAU2")
        exSheet.Cells(xRow, 3) = rsDiv("JH_REPTAU3") 'Plant
        xTotPlant = xTotPlant + rsDiv("JH_REPTAU3")
        xRow = xRow + 1
        exSheet.Cells(xRow, 1) = "% trained"
        xRow = xRow + 1
        xRow = xRow + 1
        rsDiv.MoveNext
    Loop
    'Total ----
    If xRow > xStartLine Then
        exSheet.Cells(xRow, 1) = "TOTAL RECORDED"
        exSheet.Cells(xRow, 2) = xTotStaff
        exSheet.Cells(xRow, 3) = xTotPlant
        xRow = xRow + 1
        exSheet.Cells(xRow, 1) = "% trained"
    End If
    'If rsDiv.State <> 0 Then rsDiv.Close
    'Division and Employee columns - end
    
    
    'Coursee columns - begin ----------------------------------------
    SQLQ = "SELECT * FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' AND JH_JOB = 'COURSE' ORDER BY JH_COMMENT"

    If rsCourse.State <> 0 Then rsCourse.Close
    rsCourse.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    xDivCourseCun = 0
    If Not rsCourse.EOF Then
        totNum = rsCourse.RecordCount
        xDivCourseCun = totNum
    End If
    
    For K = 1 To 2
        'K=1: Staff       K=2: Plant
        If K = 1 Then
            xStartColu = 4
            exSheet.Cells(8, xStartColu) = "Staff"
        End If
        If K = 2 Then
            If xDivCourseCun = 0 Then
                xDivCourseCun = 1
            End If
            xStartColu = xStartColu + xDivCourseCun
            exSheet.Cells(8, xStartColu) = "Plant"
        End If
        If Not (rsCourse.EOF And rsCourse.BOF) Then
            rsCourse.MoveFirst
        End If
        M = 0
        xCol = xStartColu
        Do While Not rsCourse.EOF
            exSheet.Cells(9, xCol) = rsCourse("JH_COMMENT")
              
            'loop Course Location - begin ================================
            'SQLQ = "SELECT * FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' AND JH_JOB = 'DIV' ORDER BY JH_COMMENT2"
            'SQLQ = "SELECT * FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' AND JH_JOB = 'DIV' ORDER BY JH_LUSER,JH_COMMENT2"
            'If rsDiv.State <> 0 Then rsDiv.Close
            'rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
            rsDiv.MoveFirst
            If Not rsDiv.EOF Then
                totNum = rsDiv.RecordCount: I = 0
            End If
            xRow = xStartLine
            xTotStaff = 0
            xTotPlant = 0
            xTotLoc = 0
            Do While Not rsDiv.EOF
                If (I / totNum) <= 1 Then
                    MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                    I = I + 1
                End If
                'DoEvents
                If K = 1 Then
                    xTotLoc = xTotLoc + rsDiv("JH_REPTAU2")
                End If
                If K = 2 Then
                    xTotLoc = xTotLoc + rsDiv("JH_REPTAU3")
                End If
                'count this course for this location
                SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
                SQLQ = SQLQ & "AND TT_EMP = '" & rsDiv("JH_COMMENT") & "' "
                SQLQ = SQLQ & "AND TT_LANG1 = '" & rsCourse("JH_COMMENT") & "' "
                If rsTemp.State <> 0 Then rsTemp.Close
                rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                xCunCourse = 0
                xCunLoc = 0
                If Not rsTemp.EOF Then
                    If K = 1 Then 'Staff
                        xCunLoc = rsDiv("JH_REPTAU2")
                        xCunCourse = rsTemp("TT_WHRS")
                        xTotStaff = xTotStaff + rsTemp("TT_WHRS")
                    End If
                    If K = 2 Then 'Plant
                        xCunLoc = rsDiv("JH_REPTAU3")
                        xCunCourse = rsTemp("TT_DHRS")
                        xTotPlant = xTotPlant + rsTemp("TT_DHRS")
                    End If
                End If
                rsTemp.Close
                
                'If xCunLoc > 0 Then 'don't display if the total of employee is 0 for this location
                '    exSheet.Cells(xRow, xCol) = xCunCourse
                'End If
                'xRow = xRow + 1
                'If xCunLoc > 0 Then
                '    xPercen = Round(((xCunCourse / xCunLoc) * 100), 1) '& "%"
                '    exSheet.Cells(xRow, xCol + 125) = GetColorByPerc(xPercen)
                '    xPerStr = xPercen & "%"
                '    'exSheet.Cells(xRow, xCol).Interior.Color = GetColorByPerc(xPercen)
                '    exSheet.Cells(xRow, xCol) = xPerStr
                'End If
                exSheet.Cells(xRow, xCol) = xCunCourse
                xRow = xRow + 1
                If xCunLoc > 0 Then
                    xPercen = Round(((xCunCourse / xCunLoc) * 100), 1) '& "%"
                Else
                    xPercen = 0
                End If
                exSheet.Cells(xRow, xCol + 125) = GetColorByPerc(xPercen)
                xPerStr = xPercen & "%"
                exSheet.Cells(xRow, xCol) = xPerStr
                    
                xRow = xRow + 1
                xRow = xRow + 1
                rsDiv.MoveNext
            Loop
            'Total ----
            If xRow > xStartLine Then
                If K = 1 Then 'Staff
                    exSheet.Cells(xRow, xCol) = xTotStaff
                End If
                If K = 2 Then 'Staff
                    exSheet.Cells(xRow, xCol) = xTotPlant
                End If
                xRow = xRow + 1
                'exSheet.Cells(xRow, xCol) = "% trained"
                If xTotLoc > 0 Then
                    xPercen = 0
                    If K = 1 Then 'Staff
                        xPercen = Round(((xTotStaff / xTotLoc) * 100), 1) '& "%"
                    End If
                    If K = 2 Then 'Plant
                        xPercen = Round(((xTotPlant / xTotLoc) * 100), 1) '& "%"
                    End If
                    exSheet.Cells(xRow, xCol + 125) = GetColorByPerc(xPercen)
                    xPerStr = xPercen & "%"
                    'exSheet.Cells(xRow, xCol).Interior.Color = GetColorByPerc(xPercen)
                    exSheet.Cells(xRow, xCol) = xPerStr

                End If
            End If
            'loop Course Location - end   ================================
            
            xCol = xCol + 1
            rsCourse.MoveNext
        Loop
        MDIMain.panHelp(0).FloodPercent = 100
    Next
    If rsCourse.State <> 0 Then rsCourse.Close
    'Coursee columns - end ----------------------------------------
    
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT

    Call Pause(1)
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
    'Populate Excel file - end
        
    Exit Sub
'-------------- End

WFC_WPS_Training_Report_Err:


glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
Resume Next


End Sub
Private Sub WFC_WPS_Training_Report() 'Ticket #19137
    On Error GoTo WFC_WPS_Training_Report_Err
    Dim rsEmp As New ADODB.Recordset
    Dim rsTermEmp As New ADODB.Recordset
    Dim rsEdu As New ADODB.Recordset
    Dim rsWRK As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rsDiv As New ADODB.Recordset
    Dim rsCourse As New ADODB.Recordset
    'Dim exApp As Excel.Application
    'Dim exBook As Excel.Workbook
    'Dim exSheet As Excel.Worksheet
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim SQLQ, sSQLQ As String
    Dim locSQL As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim xCol As Long
    Dim I, J, K, M, totNum
    Dim xHourlyRate
    Dim xTotVacOut
    Dim setcolumnwidth As Boolean
    Dim rsEmpPos As New ADODB.Recordset
    Dim xTERM_SEQ
    Dim xStartLine As Integer
    Dim xStartColu As Integer
    Dim xPlant As String
    Dim xDiv As String
    Dim xSalHly As String
    Dim xStr As String
    Dim xTotStaff
    Dim xTotPlant
    Dim xPercen As Double
    Dim xPerStr As String
    Dim xCunCourse As Double 'Integer
    Dim xCunLoc As Double 'Integer
    Dim xTotLoc As Double 'Integer
    Dim xDivCourseCun As Double 'Integer
    Dim xDivSQL As String
    Dim xRptSele As String
    
    'Note:
    'This program will add all data into a temp table based on the Plant/Location
    'create Plant/Location columns into a table
    'create WPS Course Code columns into a table
    'Write data into Excel file
    
    'Delete Temp tables
    SQLQ = "DELETE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "DELETE FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ

    'open Edu recordset
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    xRptSele = sSQLQ
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE (1=1) "
    SQLQ = SQLQ & "AND " & sSQLQ & " "
    'Employee Status is Active on Employee Status Matrix of WFC Custom Feature
    SQLQ = SQLQ & "AND ED_EMP IN (SELECT  EP_CODE from WFC_HRST_EMPSTATUS WHERE NOT (EP_ACTIVE_FLAG = 0)) "
    locSQL = "(" & SQLQ & ")"
    
    'SQLQ = "SELECT * from HREDSEM"
    'No duplicate course, use DISTINCT
    SQLQ = "SELECT DISTINCT ES_EMPNBR, ES_CRSCODE from HREDSEM"
    SQLQ = SQLQ & " WHERE (1=1) "
    'only employees in the selection scritera
    SQLQ = SQLQ & "AND ES_EMPNBR IN " & locSQL & " "
    'only Course Codes for WPS flag checked
    SQLQ = SQLQ & "AND ES_CRSCODE IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ESCD' AND (NOT TB_LEFLAG = 0)) "
    If IsDate(dlpDateRange(0).Text) Then
        SQLQ = SQLQ & "AND ES_DATCOMP >= " & Date_SQL(dlpDateRange(0).Text) & " "
    End If
    If IsDate(dlpDateRange(1).Text) Then
        SQLQ = SQLQ & "AND ES_DATCOMP <= " & Date_SQL(dlpDateRange(1).Text) & " "
    End If
    SQLQ = SQLQ & " ORDER BY ES_EMPNBR, ES_CRSCODE " 'ES_EMPNBR, ES_CTYPE ASC, ES_DATCOMP DESC "

    'Total = 0
    rsEdu.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    'Margaret asked: if no Edu record found than still show these plants
    'If rsEdu.EOF Then
    '    MsgBox "No record found in this Selection Criteria."
    '    Exit Sub
    'End If
    If Not rsEdu.EOF Then
        rsEdu.MoveFirst
        totNum = rsEdu.RecordCount: I = 0
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
        Screen.MousePointer = HOURGLASS
    End If
    
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC WPS Training Tmp.xls"
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC WPS Training(" & Trim(glbUserID) & ").xls"

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    FileCopy xlsFileTmp, xlsFileMat
    
    Do While Not rsEdu.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
                    
        'open Employee table
        SQLQ = "SELECT ED_EMPNBR, ED_SECTION, ED_DIV, ED_ORG FROM HREMP WHERE ED_EMPNBR = " & rsEdu("ES_EMPNBR")
        If rsEmp.State <> 0 Then rsEmp.Close
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        xPlant = ""
        xDiv = ""
        xSalHly = "Hourly"
        If Not rsEmp.EOF Then
            If Not IsNull(rsEmp("ED_SECTION")) Then
                xPlant = rsEmp("ED_SECTION")
            End If
            If Not IsNull(rsEmp("ED_DIV")) Then
                xDiv = rsEmp("ED_DIV")
            End If
            If Not IsNull(rsEmp("ED_ORG")) Then
                If rsEmp("ED_ORG") = "NONE" Or rsEmp("ED_ORG") = "EXEC" Then
                    xSalHly = "Salaried"
                End If
            End If
        End If
        rsEmp.Close
        If xPlant = "" Then GoTo next_rec
        If xDiv = "" Then GoTo next_rec
        
        'Plant - TT_PT
        'Div - TT_EMP 'Location
        'Course Code- TT_LANG1
        'Staff count - TT_WHRS (Staff:Salaried; Plant:Hourly)
        'Plant count - TT_DHRS (Staff:Salaried; Plant:Hourly)
        SQLQ = "SELECT TT_WRKEMP, TT_PT,TT_EMP,TT_LANG1,TT_WHRS,TT_DHRS, TT_SEDATE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
        'SQLQ = SQLQ & "AND TT_PT = '" & xPlant & "' " 'not use
        SQLQ = SQLQ & "AND TT_EMP = '" & xDiv & "' "
        SQLQ = SQLQ & "AND TT_LANG1 = '" & rsEdu("ES_CRSCODE") & "' "
        If rsWRK.State <> 0 Then rsWRK.Close
        rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsWRK.EOF Then
            rsWRK.AddNew
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK("TT_PT") = xPlant
            rsWRK("TT_EMP") = xDiv
            rsWRK("TT_LANG1") = rsEdu("ES_CRSCODE")
            rsWRK("TT_WHRS") = 0
            rsWRK("TT_DHRS") = 0
            rsWRK("TT_SEDATE") = Date
        End If
        If xSalHly = "Salaried" Then
            rsWRK("TT_WHRS") = rsWRK("TT_WHRS") + 1
        End If
        If xSalHly = "Hourly" Then
            rsWRK("TT_DHRS") = rsWRK("TT_DHRS") + 1
        End If
        rsWRK.Update
        
next_rec:
        rsEdu.MoveNext
    Loop
    rsEdu.Close
    If rsWRK.State <> 0 Then rsWRK.Close
    
    'Margaret asked: if no Edu record found than still show these plants
    'Save Div into Temp table:
    '--Div Desc; Staff count; Plant count
    SQLQ = "SELECT * FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' AND JH_JOB = 'DIV'"
    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
    'SQLQ = "SELECT DISTINCT TT_PT,TT_EMP FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' ORDER BY TT_EMP"
    ''SELECT JH_JOB,JH_COMMENT,JH_COMMENT2,JH_REPTAU2,JH_REPTAU3,WRKEMP FROM HRJOBWRK WHERE WRKEMP = '3142'
    'rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    'Do While Not rsTemp.EOF 'JH_LUSER
    '    rsWRK.AddNew
    '    rsWRK("WRKEMP") = glbUserID
    '    rsWRK("JH_CURRENT") = 0
    '    rsWRK("JH_JOB") = "DIV"
    '    rsWRK("JH_COMMENT") = rsTemp("TT_EMP")
    '    xStr = getDivDesc(rsTemp("TT_EMP"))
    '    rsWRK("JH_COMMENT2") = Left(Trim(xStr), 50) 'Div Desc
    '    xStr = GetTABLDesc("EDSE", rsTemp("TT_PT"))
    '    rsWRK("JH_LUSER") = Left(Trim(xStr), 25) 'Plant Desc
    '    rsWRK("JH_REPTAU2") = GetEmpCountBySele(rsTemp("TT_EMP"), "Salaried") 'Staff count
    '    rsWRK("JH_REPTAU3") = GetEmpCountBySele(rsTemp("TT_EMP"), "Hourly") 'Plant count
    '    rsWRK("JH_LDATE") = Date
    '    rsWRK("JH_LTIME") = Time$
    '    rsWRK.Update
    '    rsTemp.MoveNext
    'Loop
    'rsTemp.Close
    'rsWRK.Close
    xDivSQL = "SELECT DISTINCT ED_SECTION, ED_DIV FROM HREMP WHERE (1=1) "
    xDivSQL = xDivSQL & "AND " & sSQLQ & " "
    xDivSQL = xDivSQL & "AND ED_EMP IN (SELECT  EP_CODE from WFC_HRST_EMPSTATUS WHERE NOT (EP_ACTIVE_FLAG = 0)) "
    SQLQ = xDivSQL
    'SELECT JH_JOB,JH_COMMENT,JH_COMMENT2,JH_REPTAU2,JH_REPTAU3,WRKEMP FROM HRJOBWRK WHERE WRKEMP = '3142'
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsTemp.EOF 'JH_LUSER
        rsWRK.AddNew
        rsWRK("WRKEMP") = glbUserID
        rsWRK("JH_CURRENT") = 0
        rsWRK("JH_JOB") = "DIV"
        rsWRK("JH_COMMENT") = rsTemp("ED_DIV")
        xStr = getDivDesc(rsTemp("ED_DIV"))
        rsWRK("JH_COMMENT2") = Left(Trim(xStr), 50) 'Div Desc
        xStr = GetTABLDesc("EDSE", rsTemp("ED_SECTION"))
        rsWRK("JH_LUSER") = Left(Trim(xStr), 25) 'Plant Desc
        rsWRK("JH_REPTAU2") = GetEmpCountBySele(rsTemp("ED_DIV"), xRptSele, "Salaried") 'Staff count
        rsWRK("JH_REPTAU3") = GetEmpCountBySele(rsTemp("ED_DIV"), xRptSele, "Hourly") 'Plant count
        rsWRK("JH_LDATE") = Date
        rsWRK("JH_LTIME") = Time$
        rsWRK.Update
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    rsWRK.Close
    
        
    'Save Course Codes into Temp table:
    SQLQ = "SELECT * FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' AND JH_JOB = 'COURSE'"
    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
    SQLQ = "SELECT DISTINCT TT_LANG1 FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' ORDER BY TT_LANG1"
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsTemp.EOF
        xStr = GetTABLDesc("ESCD", rsTemp("TT_LANG1"))
        rsWRK.AddNew
        rsWRK("WRKEMP") = glbUserID
        rsWRK("JH_CURRENT") = 0
        rsWRK("JH_JOB") = "COURSE"
        rsWRK("JH_COMMENT") = rsTemp("TT_LANG1")
        rsWRK("JH_COMMENT2") = Left(Trim(xStr), 50) 'Course Desc
        rsWRK("JH_LDATE") = Date
        rsWRK("JH_LTIME") = Time$
        rsWRK.Update
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    rsWRK.Close
    
    'Populate Excel file - begin
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)
            
    exSheet.Cells(2, 2) = "Status of WPS Training for " & Year(Date)
    exSheet.Cells(3, 2) = "(" & Format(Date, "MMM dd, YYYY") & ")"
    'Less than 75% have been trained
    'exSheet.Cells(4, 4).Interior.Color = RED
    '75% and more have been trained
    'exSheet.Cells(5, 4).Interior.Color = RGB(34, 139, 34) 'Dark Green
    
    'Ticket #22481 Franks 08/27/2012 - begin
    exSheet.Cells(1, 1) = "Selection Criteria:"
    If Len(clpCode(2).Text) > 0 Then
        exSheet.Cells(3, 1) = "Status Codes"
        exSheet.Cells(4, 1) = clpCode(2).Text
    End If
    If IsDate(dlpDateRange(0).Text) Or IsDate(dlpDateRange(1).Text) Then
        xStr = "From/To Dates:"
        If IsDate(dlpDateRange(0).Text) Then
            xStr = xStr & " " & dlpDateRange(0).Text
        End If
        If IsDate(dlpDateRange(1).Text) Then
            xStr = xStr & " to " & dlpDateRange(1).Text
        End If
        exSheet.Cells(6, 1) = xStr
    End If
    'Ticket #22481 Franks 08/27/2012 - end
    
    'First line of data
    xStartLine = 10
    
    'Division and Employee columns - begin
    'SQLQ = "SELECT * FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' AND JH_JOB = 'DIV' ORDER BY JH_COMMENT2"
    SQLQ = "SELECT * FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' AND JH_JOB = 'DIV' ORDER BY JH_LUSER,JH_COMMENT2"
    
    If rsDiv.State <> 0 Then rsDiv.Close
    rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsDiv.EOF Then
        totNum = rsDiv.RecordCount: I = 0
    End If
    xRow = xStartLine
    xTotStaff = 0
    xTotPlant = 0
    Do While Not rsDiv.EOF
        If (I / totNum) <= 1 Then
            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
            I = I + 1
        End If
        DoEvents
        'Plant/Division
        exSheet.Cells(xRow, 1) = rsDiv("JH_LUSER") & "/" & rsDiv("JH_COMMENT2") '"Woodbridge Corporate" ' rsDiv("JH_COMMENT2")
        exSheet.Cells(xRow, 2) = rsDiv("JH_REPTAU2") 'Staff
        xTotStaff = xTotStaff + rsDiv("JH_REPTAU2")
        exSheet.Cells(xRow, 3) = rsDiv("JH_REPTAU3") 'Plant
        xTotPlant = xTotPlant + rsDiv("JH_REPTAU3")
        xRow = xRow + 1
        exSheet.Cells(xRow, 1) = "% trained"
        xRow = xRow + 1
        xRow = xRow + 1
        rsDiv.MoveNext
    Loop
    'Total ----
    If xRow > xStartLine Then
        exSheet.Cells(xRow, 1) = "TOTAL RECORDED"
        exSheet.Cells(xRow, 2) = xTotStaff
        exSheet.Cells(xRow, 3) = xTotPlant
        xRow = xRow + 1
        exSheet.Cells(xRow, 1) = "% trained"
    End If
    'If rsDiv.State <> 0 Then rsDiv.Close
    'Division and Employee columns - end
    
    
    'Coursee columns - begin ----------------------------------------
    SQLQ = "SELECT * FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' AND JH_JOB = 'COURSE' ORDER BY JH_COMMENT"

    If rsCourse.State <> 0 Then rsCourse.Close
    rsCourse.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    xDivCourseCun = 0
    If Not rsCourse.EOF Then
        totNum = rsCourse.RecordCount
        xDivCourseCun = totNum
    End If
    
    For K = 1 To 2
        'K=1: Staff       K=2: Plant
        If K = 1 Then
            xStartColu = 4
            exSheet.Cells(8, xStartColu) = "Staff"
        End If
        If K = 2 Then
            If xDivCourseCun = 0 Then
                xDivCourseCun = 1
            End If
            xStartColu = xStartColu + xDivCourseCun
            exSheet.Cells(8, xStartColu) = "Plant"
        End If
        If Not (rsCourse.EOF And rsCourse.BOF) Then
            rsCourse.MoveFirst
        End If
        M = 0
        xCol = xStartColu
        Do While Not rsCourse.EOF
            exSheet.Cells(9, xCol) = rsCourse("JH_COMMENT")
              
            'loop Course Location - begin ================================
            'SQLQ = "SELECT * FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' AND JH_JOB = 'DIV' ORDER BY JH_COMMENT2"
            'SQLQ = "SELECT * FROM HRJOBWRK WHERE WRKEMP = '" & glbUserID & "' AND JH_JOB = 'DIV' ORDER BY JH_LUSER,JH_COMMENT2"
            'If rsDiv.State <> 0 Then rsDiv.Close
            'rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
            rsDiv.MoveFirst
            If Not rsDiv.EOF Then
                totNum = rsDiv.RecordCount: I = 0
            End If
            xRow = xStartLine
            xTotStaff = 0
            xTotPlant = 0
            xTotLoc = 0
            Do While Not rsDiv.EOF
                If (I / totNum) <= 1 Then
                    MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                    I = I + 1
                End If
                'DoEvents
                If K = 1 Then
                    xTotLoc = xTotLoc + rsDiv("JH_REPTAU2")
                End If
                If K = 2 Then
                    xTotLoc = xTotLoc + rsDiv("JH_REPTAU3")
                End If
                'count this course for this location
                SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
                SQLQ = SQLQ & "AND TT_EMP = '" & rsDiv("JH_COMMENT") & "' "
                SQLQ = SQLQ & "AND TT_LANG1 = '" & rsCourse("JH_COMMENT") & "' "
                If rsTemp.State <> 0 Then rsTemp.Close
                rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                xCunCourse = 0
                xCunLoc = 0
                If Not rsTemp.EOF Then
                    If K = 1 Then 'Staff
                        xCunLoc = rsDiv("JH_REPTAU2")
                        xCunCourse = rsTemp("TT_WHRS")
                        xTotStaff = xTotStaff + rsTemp("TT_WHRS")
                    End If
                    If K = 2 Then 'Plant
                        xCunLoc = rsDiv("JH_REPTAU3")
                        xCunCourse = rsTemp("TT_DHRS")
                        xTotPlant = xTotPlant + rsTemp("TT_DHRS")
                    End If
                End If
                rsTemp.Close
                
                'If xCunLoc > 0 Then 'don't display if the total of employee is 0 for this location
                '    exSheet.Cells(xRow, xCol) = xCunCourse
                'End If
                'xRow = xRow + 1
                'If xCunLoc > 0 Then
                '    xPercen = Round(((xCunCourse / xCunLoc) * 100), 1) '& "%"
                '    exSheet.Cells(xRow, xCol + 125) = GetColorByPerc(xPercen)
                '    xPerStr = xPercen & "%"
                '    'exSheet.Cells(xRow, xCol).Interior.Color = GetColorByPerc(xPercen)
                '    exSheet.Cells(xRow, xCol) = xPerStr
                'End If
                exSheet.Cells(xRow, xCol) = xCunCourse
                xRow = xRow + 1
                If xCunLoc > 0 Then
                    xPercen = Round(((xCunCourse / xCunLoc) * 100), 1) '& "%"
                Else
                    xPercen = 0
                End If
                exSheet.Cells(xRow, xCol + 125) = GetColorByPerc(xPercen)
                xPerStr = xPercen & "%"
                exSheet.Cells(xRow, xCol) = xPerStr
                    
                xRow = xRow + 1
                xRow = xRow + 1
                rsDiv.MoveNext
            Loop
            'Total ----
            If xRow > xStartLine Then
                If K = 1 Then 'Staff
                    exSheet.Cells(xRow, xCol) = xTotStaff
                End If
                If K = 2 Then 'Staff
                    exSheet.Cells(xRow, xCol) = xTotPlant
                End If
                xRow = xRow + 1
                'exSheet.Cells(xRow, xCol) = "% trained"
                If xTotLoc > 0 Then
                    xPercen = 0
                    If K = 1 Then 'Staff
                        xPercen = Round(((xTotStaff / xTotLoc) * 100), 1) '& "%"
                    End If
                    If K = 2 Then 'Plant
                        xPercen = Round(((xTotPlant / xTotLoc) * 100), 1) '& "%"
                    End If
                    exSheet.Cells(xRow, xCol + 125) = GetColorByPerc(xPercen)
                    xPerStr = xPercen & "%"
                    'exSheet.Cells(xRow, xCol).Interior.Color = GetColorByPerc(xPercen)
                    exSheet.Cells(xRow, xCol) = xPerStr

                End If
            End If
            'loop Course Location - end   ================================
            
            xCol = xCol + 1
            rsCourse.MoveNext
        Loop
        MDIMain.panHelp(0).FloodPercent = 100
    Next
    If rsCourse.State <> 0 Then rsCourse.Close
    'Coursee columns - end ----------------------------------------
    
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT

    Call Pause(1)
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
    'Populate Excel file - end
        
    Exit Sub
'-------------- End

WFC_WPS_Training_Report_Err:


glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
Resume Next

End Sub
Private Sub Samuel_Employee_XLS_Rpt()
    On Error GoTo Samuel_Employee_XLS_Report_Err
    Dim rsEmp As New ADODB.Recordset
    Dim rsTermEmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xlsFileMa2 As String
    Dim xRow As Long
    Dim I, totNum, K
    Dim xHourlyRate
    Dim xTotVacOut
    Dim rsTmp As New ADODB.Recordset
    Dim setcolumnwidth As Boolean
    Dim rsEmpPos As New ADODB.Recordset
    Dim rsJOB As New ADODB.Recordset
    Dim xTERM_SEQ
    Dim Total As Integer
    Dim xFilePath
    
    setcolumnwidth = False
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    'If OptAct.Value Then
        SQLQ = "SELECT * FROM HREMP WHERE (1=1) "
    'Else 'Term
    '    SQLQ = "SELECT * FROM TERM_HREMP WHERE (1=1) "
    '    sSQLQ = Replace(sSQLQ, "HREMP.", "TERM_HREMP.")
    'End If
    SQLQ = SQLQ & "AND " & sSQLQ & " "
    SQLQ = SQLQ & "ORDER BY ED_SURNAME,ED_FNAME "

    'Call WriteFile("SQL1=" & SQLQ)
    
    Total = 0
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsEmp.EOF Then
        rsEmp.MoveFirst
        totNum = rsEmp.RecordCount: I = 0
        
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2382_Employee_Salary.xls"
        'xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2382_Employee_Salary" & Trim(glbUserID) & ".xls"
        xFilePath = ""
        'Get the Export path
        If gsTRAININGMATRIX Then
            xFilePath = GetComPreferEmail("TRAININGMATRIX")
        End If
        If Len(xFilePath) = 0 Then
            xFilePath = glbIHRREPORTS
        End If
        'xlsFileTmp = xFilePath & IIf(Right(xFilePath, 1) = "\", "", "\") & "SN2382_Employee_Salary.xls"
        xlsFileMat = xFilePath & IIf(Right(xFilePath, 1) = "\", "", "\") & "SN2382_Employee_Salary" & Trim(glbUserID) & ".xls"

        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
        
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
        
        exSheet.Cells(3, 1) = "As Of " & Format(Now, "mmm dd, yyyy")
        exSheet.Cells(4, 12) = lStr("Organization 1")
        exSheet.Cells(4, 13) = lStr("Organization 1")
        
        xRow = 5
        
        Do While Not rsEmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
           
            'exSheet.Cells(xRow, 1) = rsEmp("ED_ADMINBY") 'Payroll #
            exSheet.Cells(xRow, 1) = rsEmp("ED_EMPNBR")
            'exSheet.Cells(xRow, 1).Font.Bold = True
            exSheet.Cells(xRow, 2) = rsEmp("ED_SURNAME") & ", " & rsEmp("ED_FNAME")
            'exSheet.Cells(xRow, 3) = rsEmp("ED_SURNAME") 'Title
            If Not IsNull(rsEmp("ED_DOH")) Then exSheet.Cells(xRow, 4) = CVDate(rsEmp("ED_DOB"))
            If Not IsNull(rsEmp("ED_SENDTE")) Then exSheet.Cells(xRow, 5) = CVDate(rsEmp("ED_SENDTE"))
            If Not IsNull(rsEmp("ED_DIV")) Then exSheet.Cells(xRow, 8) = rsEmp("ED_DIV") 'Company
            If Not IsNull(rsEmp("ED_LOC")) Then exSheet.Cells(xRow, 9) = rsEmp("ED_LOC") 'Business Divison
            If Not IsNull(rsEmp("ED_SECTION")) Then exSheet.Cells(xRow, 10) = GetTABLCode("EDSE", rsEmp("ED_SECTION")) 'Branch
            If Not IsNull(rsEmp("ED_DEPTNO")) Then exSheet.Cells(xRow, 11) = getDeptDesc(rsEmp("ED_DEPTNO")) 'Dept.
            If Not IsNull(rsEmp("ED_ORGT1")) Then exSheet.Cells(xRow, 12) = GetTABLCode("ORGN", rsEmp("ED_ORGT1")) 'Organization 1
            If Not IsNull(rsEmp("ED_ORGT2")) Then exSheet.Cells(xRow, 13) = GetTABLCode("ORGN", rsEmp("ED_ORGT2")) 'Organization 1
            
            'Salary information
            'If xTERM_SEQ = 0 Then
                SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT (SH_CURRENT = 0) AND SH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
            'Else
            '    SQLQ = "SELECT * FROM TERM_SALARY_HISTORY WHERE NOT (SH_CURRENT = 0) AND SH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
            '    SQLQ = SQLQ & "AND TERM_SEQ = " & xTERM_SEQ & " "
            'End If
            If rsEmpPos.State <> 0 Then rsEmpPos.Close
            rsEmpPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsEmpPos.EOF Then
                If Not IsNull(rsEmpPos("SH_JOB")) Then exSheet.Cells(xRow, 3) = GetJobDesc(rsEmpPos("SH_JOB"))
                exSheet.Cells(xRow, 6) = rsEmpPos("SH_SALARY")
                exSheet.Cells(xRow, 7) = rsEmpPos("SH_EDATE")
            End If
            rsEmpPos.Close
            'If xTERM_SEQ > 0 Then 'rsTermEmp
            '    SQLQ = "SELECT * FROM Term_HRTRMEMP WHERE Employee_Number = " & rsEmp("ED_EMPNBR") & " AND TERM_SEQ = " & xTERM_SEQ & " "
            '    If rsTermEmp.State <> 0 Then rsTermEmp.Close
            '    rsTermEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            '    If Not rsTermEmp.EOF Then
            '        If Not IsNull(rsTermEmp("Term_DOT")) Then exSheet.Cells(xRow, 37) = CVDate(rsTermEmp("Term_DOT"))  'Termination Date
            '        If Not IsNull(rsTermEmp("Term_Reason")) Then exSheet.Cells(xRow, 38) = rsTermEmp("Term_Reason") 'Termination Reason
            '    End If
            '    rsTermEmp.Close
            'End If
            
            xRow = xRow + 1
            rsEmp.MoveNext
        Loop
        
        'exSheet.Cells(xRow + 2, 1) = "Total number of records for this department: " & Total
        'xRow = xRow + 2
        exSheet.Cells(xRow + 2, 1) = "Total number of records in this worksheet: " & totNum
        exSheet.Rows(xRow + 2).Font.Bold = True
        
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(2)
        
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
        If totNum = 0 Then
            MsgBox "No records were found!", , "Message!"
        End If
    Else
        MsgBox "No records were found!", , "Message!"
    End If
        
   
    rsEmp.Close

Exit Sub


Samuel_Employee_XLS_Report_Err:


glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
Resume Next

End Sub

Private Sub Samuel_Comparison_XLS_Rpt_Short()
    On Error GoTo Samuel_Comparison_XLS_Report_Err
    Dim rsEmp As New ADODB.Recordset
    Dim rsTermEmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xlsFileMa2 As String
    Dim xRow As Long
    Dim I, totNum, K
    Dim xHourlyRate
    Dim xTotVacOut
    Dim rsTmp As New ADODB.Recordset
    Dim setcolumnwidth As Boolean
    Dim rsEmpPos As New ADODB.Recordset
    Dim rsJOB As New ADODB.Recordset
    Dim xTERM_SEQ
    
    
    setcolumnwidth = False
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    If OptAct.Value Then
        SQLQ = "SELECT * FROM HREMP WHERE (1=1) "
    Else 'Term
        SQLQ = "SELECT * FROM TERM_HREMP WHERE (1=1) "
        sSQLQ = Replace(sSQLQ, "HREMP.", "TERM_HREMP.")
    End If
    SQLQ = SQLQ & "AND " & sSQLQ & " "
    SQLQ = SQLQ & "ORDER BY ED_SURNAME,ED_FNAME "

    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsEmp.EOF Then
        rsEmp.MoveFirst
        totNum = rsEmp.RecordCount: I = 0
        
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2382_IDLComparison_short.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2382_IDLComparison_short" & Trim(glbUserID) & ".xls"
        'xlsFileMa2 = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2382_IDLComparison_short" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
        'If (Dir(xlsFileMa2)) <> "" Then Kill xlsFileMa2
        
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
        
        exSheet.Cells(3, 1) = "As Of " & Format(Now, "mmm dd, yyyy")
        
        xRow = 5
        
        Do While Not rsEmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
           
            exSheet.Cells(xRow, 1) = rsEmp("ED_ADMINBY") 'Payroll #
            exSheet.Cells(xRow, 2) = rsEmp("ED_EMPNBR")
            'exSheet.Cells(xRow, 1).Font.Bold = True
            exSheet.Cells(xRow, 3) = rsEmp("ED_FNAME")
            exSheet.Cells(xRow, 4) = rsEmp("ED_SURNAME")
            If Not IsNull(rsEmp("ED_SIN")) Then exSheet.Cells(xRow, 5) = rsEmp("ED_SIN")
            If Not IsNull(rsEmp("ED_EMP")) Then exSheet.Cells(xRow, 6) = rsEmp("ED_EMP") 'Employment Status '
            If Not IsNull(rsEmp("ED_MSTAT")) Then exSheet.Cells(xRow, 7) = rsEmp("ED_MSTAT")
            If Not IsNull(rsEmp("ED_SEX")) Then exSheet.Cells(xRow, 8) = rsEmp("ED_SEX") 'Gender
            If Not IsNull(rsEmp("ED_ADDR1")) Then exSheet.Cells(xRow, 9) = rsEmp("ED_ADDR1") 'Address
            If Not IsNull(rsEmp("ED_CITY")) Then exSheet.Cells(xRow, 10) = rsEmp("ED_CITY") '
            If Not IsNull(rsEmp("ED_PROV")) Then exSheet.Cells(xRow, 11) = rsEmp("ED_PROV")
            If Not IsNull(rsEmp("ED_PCODE")) Then exSheet.Cells(xRow, 12) = rsEmp("ED_PCODE")
            If Not IsNull(rsEmp("ED_DOH")) Then exSheet.Cells(xRow, 13) = CVDate(rsEmp("ED_DOH"))
            If Not IsNull(rsEmp("ED_GLNO")) Then exSheet.Cells(xRow, 14) = rsEmp("ED_GLNO")
            If Not IsNull(rsEmp("ED_SECTION")) Then exSheet.Cells(xRow, 15) = rsEmp("ED_SECTION") 'Branch 'GetTABLCode("EDSE", rsEmp("ED_SECTION"))
            If Not IsNull(rsEmp("ED_DEPTNO")) Then exSheet.Cells(xRow, 16) = rsEmp("ED_DEPTNO") 'Dept.
            'exSheet.Cells(xRow, 9) = getDeptDesc(rsEmp("ED_DEPTNO")) ' Dept Desc
            
            If OptAct.Value Then
                xTERM_SEQ = 0
            Else
                xTERM_SEQ = rsEmp("TERM_SEQ")
            End If
            
            'Salary information
            If xTERM_SEQ = 0 Then
                SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT (SH_CURRENT = 0) AND SH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
            Else
                SQLQ = "SELECT * FROM TERM_SALARY_HISTORY WHERE NOT (SH_CURRENT = 0) AND SH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
                SQLQ = SQLQ & "AND TERM_SEQ = " & xTERM_SEQ & " "
            End If
            If rsEmpPos.State <> 0 Then rsEmpPos.Close
            rsEmpPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsEmpPos.EOF Then
                'Pay Type is Pay Period
                If Not IsNull(rsEmpPos("SH_PAYP")) Then
                    exSheet.Cells(xRow, 17) = rsEmpPos("SH_PAYP")
                End If
                'Per Annum/Hourly/Monthly
                If Not IsNull(rsEmpPos("SH_SALCD")) Then
                    If rsEmpPos("SH_SALCD") = "A" Then
                        exSheet.Cells(xRow, 18) = "Annum"
                    End If
                    If rsEmpPos("SH_SALCD") = "H" Then
                        exSheet.Cells(xRow, 18) = "Hourly"
                    End If
                    If rsEmpPos("SH_SALCD") = "M" Then
                        exSheet.Cells(xRow, 18) = "Monthly"
                    End If
                End If
                'Salary Change Reason
                If Not IsNull(rsEmpPos("SH_SREAS1")) Then
                    exSheet.Cells(xRow, 19) = GetTABLCode("SDRC", rsEmpPos("SH_SREAS1"))
                End If
                'Effective Date
                If Not IsNull(rsEmpPos("SH_EDATE")) Then exSheet.Cells(xRow, 20) = CVDate(rsEmpPos("SH_EDATE"))
                'Hours per week
                If Not IsNull(rsEmpPos("SH_WHRS")) Then
                    exSheet.Cells(xRow, 21) = rsEmpPos("SH_WHRS")
                End If
            End If
            rsEmpPos.Close
            If Not IsNull(rsEmp("ED_TD1DOL")) Then exSheet.Cells(xRow, 22) = rsEmp("ED_TD1DOL")
            If Not IsNull(rsEmp("ED_PROVAMT")) Then exSheet.Cells(xRow, 23) = rsEmp("ED_PROVAMT")
            If Not IsNull(rsEmp("ED_TD3")) Then exSheet.Cells(xRow, 24) = rsEmp("ED_TD3")
            If Not IsNull(rsEmp("ED_ExtraTax")) Then exSheet.Cells(xRow, 25) = rsEmp("ED_ExtraTax")
            If Not IsNull(rsEmp("ED_UIC")) Then exSheet.Cells(xRow, 26) = rsEmp("ED_UIC") 'EI Code
            If Not IsNull(rsEmp("ED_CPP")) Then exSheet.Cells(xRow, 27) = rsEmp("ED_CPP")
            If Not IsNull(rsEmp("ED_WCB")) Then exSheet.Cells(xRow, 28) = rsEmp("ED_WCB") 'Status Federal Tax
            If Not IsNull(rsEmp("ED_PROVEMP")) Then exSheet.Cells(xRow, 29) = rsEmp("ED_PROVEMP")
            If Not IsNull(rsEmp("ED_BANK")) Then exSheet.Cells(xRow, 30) = rsEmp("ED_BANK")
            If Not IsNull(rsEmp("ED_BRANCH")) Then exSheet.Cells(xRow, 31) = rsEmp("ED_BRANCH")
            If Not IsNull(rsEmp("ED_ACCOUNT")) Then exSheet.Cells(xRow, 32) = rsEmp("ED_ACCOUNT")
            If Not IsNull(rsEmp("ED_BANK2")) Then exSheet.Cells(xRow, 33) = rsEmp("ED_BANK2")
            If Not IsNull(rsEmp("ED_BRANCH2")) Then exSheet.Cells(xRow, 34) = rsEmp("ED_BRANCH2")
            If Not IsNull(rsEmp("ED_ACCOUNT2")) Then exSheet.Cells(xRow, 35) = rsEmp("ED_ACCOUNT2")
            If Not IsNull(rsEmp("ED_AMTDEPOSIT2")) Then exSheet.Cells(xRow, 36) = rsEmp("ED_AMTDEPOSIT2")
            If xTERM_SEQ > 0 Then 'rsTermEmp
                SQLQ = "SELECT * FROM Term_HRTRMEMP WHERE Employee_Number = " & rsEmp("ED_EMPNBR") & " AND TERM_SEQ = " & xTERM_SEQ & " "
                If rsTermEmp.State <> 0 Then rsTermEmp.Close
                rsTermEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsTermEmp.EOF Then
                    If Not IsNull(rsTermEmp("Term_DOT")) Then exSheet.Cells(xRow, 37) = CVDate(rsTermEmp("Term_DOT"))  'Termination Date
                    If Not IsNull(rsTermEmp("Term_Reason")) Then exSheet.Cells(xRow, 38) = rsTermEmp("Term_Reason") 'Termination Reason
                End If
                rsTermEmp.Close
            End If
            
            xRow = xRow + 1
            rsEmp.MoveNext
        Loop
        
        'exSheet.Cells(xRow + 2, 1) = "Total number of records for this department: " & Total
        'xRow = xRow + 2
        exSheet.Cells(xRow + 2, 1) = "Total number of records in this worksheet: " & totNum
        exSheet.Rows(xRow + 2).Font.Bold = True
        
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(2)
        
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
        If totNum = 0 Then
            MsgBox "No records were found!", , "Message!"
        End If
    Else
        MsgBox "No records were found!", , "Message!"
    End If
        
   
     rsEmp.Close
Exit Sub


Samuel_Comparison_XLS_Report_Err:


glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
Resume Next

End Sub

Private Sub Samuel_Comparison_XLS_Report()
    On Error GoTo Samuel_Comparison_XLS_Report_Err
    Dim rsEmp As New ADODB.Recordset
    Dim rsTermEmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xlsFileMa2 As String
    Dim xRow As Long
    Dim I, totNum, K
    Dim xHourlyRate
    Dim xTotVacOut
    Dim rsTmp As New ADODB.Recordset
    Dim setcolumnwidth As Boolean
    Dim rsEmpPos As New ADODB.Recordset
    Dim rsJOB As New ADODB.Recordset
    Dim xTERM_SEQ
    
    
    setcolumnwidth = False
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    If OptAct.Value Then
        SQLQ = "SELECT * FROM HREMP WHERE (1=1) "
    Else 'Term
        SQLQ = "SELECT * FROM TERM_HREMP WHERE (1=1) "
        sSQLQ = Replace(sSQLQ, "HREMP.", "TERM_HREMP.")
    End If
    SQLQ = SQLQ & "AND " & sSQLQ & " "
    SQLQ = SQLQ & "ORDER BY ED_SURNAME,ED_FNAME "

    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsEmp.EOF Then
        rsEmp.MoveFirst
        totNum = rsEmp.RecordCount: I = 0
        
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2382_IDLComparison.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2382_IDLComparison" & Trim(glbUserID) & ".xls"
        'xlsFileMa2 = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2382_IDLComparison_short" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
        'If (Dir(xlsFileMa2)) <> "" Then Kill xlsFileMa2
        
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
        
        exSheet.Cells(3, 1) = "As Of " & Format(Now, "mmm dd, yyyy")
        
        'Rept. Authority 1,2,3,4
        exSheet.Cells(4, 43) = lStr("Rept. Authority 1")
        exSheet.Cells(4, 44) = lStr("Rept. Authority 2")
        exSheet.Cells(4, 45) = lStr("Rept. Authority 3")
        exSheet.Cells(4, 46) = lStr("Rept. Authority 4")
        
        xRow = 5
        
        Do While Not rsEmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
           
            exSheet.Cells(xRow, 1) = rsEmp("ED_ADMINBY") 'Payroll #
            exSheet.Cells(xRow, 2) = rsEmp("ED_EMPNBR")
            'exSheet.Cells(xRow, 1).Font.Bold = True
            exSheet.Cells(xRow, 3) = rsEmp("ED_FNAME")
            exSheet.Cells(xRow, 4) = rsEmp("ED_SURNAME")
            If Not IsNull(rsEmp("ED_SIN")) Then exSheet.Cells(xRow, 5) = rsEmp("ED_SIN")
            If Not IsNull(rsEmp("ED_EMP")) Then exSheet.Cells(xRow, 6) = rsEmp("ED_EMP") 'Employment Status '
            If Not IsNull(rsEmp("ED_MSTAT")) Then exSheet.Cells(xRow, 7) = rsEmp("ED_MSTAT")
            If Not IsNull(rsEmp("ED_SEX")) Then exSheet.Cells(xRow, 8) = rsEmp("ED_SEX") 'Gender
            If Not IsNull(rsEmp("ED_ADDR1")) Then exSheet.Cells(xRow, 9) = rsEmp("ED_ADDR1") 'Address
            If Not IsNull(rsEmp("ED_CITY")) Then exSheet.Cells(xRow, 10) = rsEmp("ED_CITY") '
            If Not IsNull(rsEmp("ED_PROV")) Then exSheet.Cells(xRow, 11) = rsEmp("ED_PROV")
            If Not IsNull(rsEmp("ED_PCODE")) Then exSheet.Cells(xRow, 12) = rsEmp("ED_PCODE")
            If Not IsNull(rsEmp("ED_DOH")) Then exSheet.Cells(xRow, 13) = CVDate(rsEmp("ED_DOH"))
            If Not IsNull(rsEmp("ED_GLNO")) Then exSheet.Cells(xRow, 14) = rsEmp("ED_GLNO")
            If Not IsNull(rsEmp("ED_SECTION")) Then exSheet.Cells(xRow, 15) = rsEmp("ED_SECTION") 'Branch 'GetTABLCode("EDSE", rsEmp("ED_SECTION"))
            If Not IsNull(rsEmp("ED_DEPTNO")) Then exSheet.Cells(xRow, 16) = rsEmp("ED_DEPTNO") 'Dept.
            'exSheet.Cells(xRow, 9) = getDeptDesc(rsEmp("ED_DEPTNO")) ' Dept Desc
            
            If OptAct.Value Then
                xTERM_SEQ = 0
            Else
                xTERM_SEQ = rsEmp("TERM_SEQ")
            End If
            
            'Salary information
            If xTERM_SEQ = 0 Then
                SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT (SH_CURRENT = 0) AND SH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
            Else
                SQLQ = "SELECT * FROM TERM_SALARY_HISTORY WHERE NOT (SH_CURRENT = 0) AND SH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
                SQLQ = SQLQ & "AND TERM_SEQ = " & xTERM_SEQ & " "
            End If
            If rsEmpPos.State <> 0 Then rsEmpPos.Close
            rsEmpPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsEmpPos.EOF Then
                'Pay Type is Pay Period
                If Not IsNull(rsEmpPos("SH_PAYP")) Then
                    exSheet.Cells(xRow, 17) = rsEmpPos("SH_PAYP")
                End If
                'Per Annum/Hourly/Monthly
                If Not IsNull(rsEmpPos("SH_SALCD")) Then
                    If rsEmpPos("SH_SALCD") = "A" Then
                        exSheet.Cells(xRow, 18) = "Annum"
                    End If
                    If rsEmpPos("SH_SALCD") = "H" Then
                        exSheet.Cells(xRow, 18) = "Hourly"
                    End If
                    If rsEmpPos("SH_SALCD") = "M" Then
                        exSheet.Cells(xRow, 18) = "Monthly"
                    End If
                End If
                'Salary Change Reason
                If Not IsNull(rsEmpPos("SH_SREAS1")) Then
                    exSheet.Cells(xRow, 19) = GetTABLCode("SDRC", rsEmpPos("SH_SREAS1"))
                End If
                'Effective Date
                If Not IsNull(rsEmpPos("SH_EDATE")) Then exSheet.Cells(xRow, 20) = CVDate(rsEmpPos("SH_EDATE"))
                'Hours per week
                If Not IsNull(rsEmpPos("SH_WHRS")) Then
                    exSheet.Cells(xRow, 21) = rsEmpPos("SH_WHRS")
                End If
            End If
            rsEmpPos.Close
            If Not IsNull(rsEmp("ED_TD1DOL")) Then exSheet.Cells(xRow, 22) = rsEmp("ED_TD1DOL")
            If Not IsNull(rsEmp("ED_PROVAMT")) Then exSheet.Cells(xRow, 23) = rsEmp("ED_PROVAMT")
            If Not IsNull(rsEmp("ED_TD3")) Then exSheet.Cells(xRow, 24) = rsEmp("ED_TD3")
            If Not IsNull(rsEmp("ED_ExtraTax")) Then exSheet.Cells(xRow, 25) = rsEmp("ED_ExtraTax")
            If Not IsNull(rsEmp("ED_UIC")) Then exSheet.Cells(xRow, 26) = rsEmp("ED_UIC") 'EI Code
            If Not IsNull(rsEmp("ED_CPP")) Then exSheet.Cells(xRow, 27) = rsEmp("ED_CPP")
            If Not IsNull(rsEmp("ED_WCB")) Then exSheet.Cells(xRow, 28) = rsEmp("ED_WCB") 'Status Federal Tax
            If Not IsNull(rsEmp("ED_PROVEMP")) Then exSheet.Cells(xRow, 29) = rsEmp("ED_PROVEMP")
            If Not IsNull(rsEmp("ED_BANK")) Then exSheet.Cells(xRow, 30) = rsEmp("ED_BANK")
            If Not IsNull(rsEmp("ED_BRANCH")) Then exSheet.Cells(xRow, 31) = rsEmp("ED_BRANCH")
            If Not IsNull(rsEmp("ED_ACCOUNT")) Then exSheet.Cells(xRow, 32) = rsEmp("ED_ACCOUNT")
            If Not IsNull(rsEmp("ED_BANK2")) Then exSheet.Cells(xRow, 33) = rsEmp("ED_BANK2")
            If Not IsNull(rsEmp("ED_BRANCH2")) Then exSheet.Cells(xRow, 34) = rsEmp("ED_BRANCH2")
            If Not IsNull(rsEmp("ED_ACCOUNT2")) Then exSheet.Cells(xRow, 35) = rsEmp("ED_ACCOUNT2")
            If Not IsNull(rsEmp("ED_AMTDEPOSIT2")) Then exSheet.Cells(xRow, 36) = rsEmp("ED_AMTDEPOSIT2")
            If xTERM_SEQ > 0 Then 'rsTermEmp
                SQLQ = "SELECT * FROM Term_HRTRMEMP WHERE Employee_Number = " & rsEmp("ED_EMPNBR") & " AND TERM_SEQ = " & xTERM_SEQ & " "
                If rsTermEmp.State <> 0 Then rsTermEmp.Close
                rsTermEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsTermEmp.EOF Then
                    If Not IsNull(rsTermEmp("Term_DOT")) Then exSheet.Cells(xRow, 37) = CVDate(rsTermEmp("Term_DOT"))  'Termination Date
                    If Not IsNull(rsTermEmp("Term_Reason")) Then exSheet.Cells(xRow, 38) = rsTermEmp("Term_Reason") 'Termination Reason
                End If
                rsTermEmp.Close
            End If
            
            'Ticket #21465 Franks 02/17/2012 - begin
            If Not IsNull(rsEmp("ED_DIV")) Then exSheet.Cells(xRow, 39) = rsEmp("ED_DIV")
            If Not IsNull(rsEmp("ED_PT")) Then exSheet.Cells(xRow, 40) = rsEmp("ED_PT")
            If Not IsNull(rsEmp("ED_REGION")) Then exSheet.Cells(xRow, 41) = rsEmp("ED_REGION")
            'position
            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE NOT (JH_CURRENT = 0) AND JH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
            If rsEmpPos.State <> 0 Then rsEmpPos.Close
            rsEmpPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsEmpPos.EOF Then
                'exSheet.Cells(xRow, 11) = rsEmpPos("JH_JOB") 'NEW Position Code
                exSheet.Cells(xRow, 42) = GetJobDesc(rsEmpPos("JH_JOB"))
                'exSheet.Cells(xRow, 13) = CVDate(rsEmpPos("JH_SDATE")) 'Start Date
                If Not IsNull(rsEmpPos("JH_REPTAU")) Then
                    exSheet.Cells(xRow, 43) = getEEName(rsEmpPos("JH_REPTAU"))  'Reports To
                    'exSheet.Cells(xRow, 15) = rsEmpPos("JH_REPTAU") 'Reports To EE #
                End If
                If Not IsNull(rsEmpPos("JH_REPTAU2")) Then
                    exSheet.Cells(xRow, 44) = getEEName(rsEmpPos("JH_REPTAU2"))  'Reports To
                End If
                If Not IsNull(rsEmpPos("JH_REPTAU3")) Then
                    exSheet.Cells(xRow, 45) = getEEName(rsEmpPos("JH_REPTAU3"))  'Reports To
                End If
                If Not IsNull(rsEmpPos("JH_REPTAU4")) Then
                    exSheet.Cells(xRow, 46) = getEEName(rsEmpPos("JH_REPTAU4"))  'Reports To
                End If
                'exSheet.Cells(xRow, 16) = rsEmpPos("JH_DHRS") '
                'exSheet.Cells(xRow, 17) = rsEmpPos("JH_WHRS")
                'exSheet.Cells(xRow, 18) = rsEmpPos("JH_PHRS") '
            End If
            rsEmpPos.Close
            If Not IsNull(rsEmp("ED_VADIM1")) Then exSheet.Cells(xRow, 47) = rsEmp("ED_VADIM1") 'Vacation Code
            If Not IsNull(rsEmp("ED_SUPCODE")) Then exSheet.Cells(xRow, 48) = rsEmp("ED_SUPCODE") 'Pension Code
            If Not IsNull(rsEmp("ED_USRDAT1")) Then exSheet.Cells(xRow, 49) = CVDate(rsEmp("ED_USRDAT1")) 'Vacation Date
            If Not IsNull(rsEmp("ED_OMERS")) Then exSheet.Cells(xRow, 50) = CVDate(rsEmp("ED_OMERS")) 'Pension Date
            If Not IsNull(rsEmp("ED_SENDTE")) Then exSheet.Cells(xRow, 51) = CVDate(rsEmp("ED_SENDTE")) 'Effective Service Date
            'Ticket #21465 Franks 02/17/2012 - end
            
            xRow = xRow + 1
            rsEmp.MoveNext
        Loop
        
        'exSheet.Cells(xRow + 2, 1) = "Total number of records for this department: " & Total
        'xRow = xRow + 2
        exSheet.Cells(xRow + 2, 1) = "Total number of records in this worksheet: " & totNum
        exSheet.Rows(xRow + 2).Font.Bold = True
        
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(2)
        
        '''Ticket #21465 Franks 02/17/2012 - begin
        '''copy the full file into the short file, then remove the unused info
        ''FileCopy xlsFileMat, xlsFileMa2
        '''Create new WorkBook of Excel
        ''Set exApp = CreateObject("Excel.Application")
        ''Set exBook = exApp.Workbooks.Open(xlsFileMa2)
        ''Set exSheet = exBook.Worksheets(1)
        '''totNum
        ''For I = 4 To totNum
        ''    For K = 39 To 51
        ''        exSheet.Cells(I, K) = ""
        ''    Next K
        ''Next I
        ''exBook.Save
        ''Set exSheet = Nothing
        ''Set exBook = Nothing
        ''exApp.Quit
        ''Set exApp = Nothing
        '''Ticket #21465 Franks 02/17/2012 - end
        
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
        If totNum = 0 Then
            MsgBox "No records were found!", , "Message!"
        End If
    Else
        MsgBox "No records were found!", , "Message!"
    End If
        
   
     rsEmp.Close
Exit Sub


Samuel_Comparison_XLS_Report_Err:


glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
Resume Next
End Sub

Function GetTABLCode(xName, xCode) 'Ticket #16544
Dim rsTABL As New ADODB.Recordset
Dim xStr As String
rsTABL.Open "SELECT * FROM HRTABL WHERE TB_NAME='" & xName & "' AND TB_KEY='" & xCode & "'", gdbAdoIhr001, adOpenStatic, adLockPessimistic
xStr = ""
If Not rsTABL.EOF Then
    xStr = rsTABL("TB_DESC")
End If
rsTABL.Close
GetTABLCode = xStr
End Function
Private Function getDivDesc(xCode) 'Ticket #16544
Dim rsDiv As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = ""
    If Not IsNull(xCode) Then
        SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & xCode & "' "
        rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsDiv.EOF Then
            xRetVal = rsDiv("Division_Name")
        End If
        rsDiv.Close
    End If
    getDivDesc = xRetVal
End Function
Private Function getDeptDesc(xCode) 'Ticket #16544
Dim rsDEPT As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = ""
    If Not IsNull(xCode) Then
        SQLQ = "SELECT * FROM HRDEPT WHERE DF_NBR = '" & xCode & "' "
        rsDEPT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsDEPT.EOF Then
            xRetVal = rsDEPT("DF_NAME")
        End If
        rsDEPT.Close
    End If
    getDeptDesc = xRetVal
End Function
Private Function GetJobDesc(xCode) 'Ticket #16544
Dim rsJOB As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = ""
    If Not IsNull(xCode) Then
        SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & xCode & "' "
        rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsJOB.EOF Then
            xRetVal = rsJOB("JB_DESCR")
        End If
        rsJOB.Close
    End If
    GetJobDesc = xRetVal
End Function

Private Function getEEName(xCode) 'Ticket #16544
Dim rsEEName As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = ""
    If Not IsNull(xCode) Then
        SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME FROM HREMP WHERE ED_EMPNBR = " & xCode & " "
        rsEEName.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsEEName.EOF Then
            xRetVal = rsEEName("ED_SURNAME") & ", " & rsEEName("ED_FNAME")
        End If
        rsEEName.Close
    End If
    getEEName = xRetVal
End Function

Private Sub License_Report_XLS_CCORP()
    On Error GoTo License_Report_XLS_CCORP_Err

    Dim rsHRLic As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    Dim xHourlyRate
    Dim xTotVacOut
    Dim rsTmp As New ADODB.Recordset
    Dim setcolumnwidth As Boolean
    
    setcolumnwidth = False

    'Check the Month End Date is within the Current Vacation Year
    If Len(Trim(dlpDateRange(4).Text)) > 0 Or Len(Trim(dlpDateRange(5).Text)) > 0 Then
        If Not IsDate(dlpDateRange(4).Text) Or Not IsDate(dlpDateRange(5).Text) Then
        MsgBox "Invalid  Date range!"
        dlpDateRange(4).SetFocus
        End If
    End If
    Dim xpirycondtion As String
    xpirycondtion = ""
    If chkExpired.Value = True Then
        xpirycondtion = " AND datediff(day,getdate(),HR_USERDEFINE_TABLE.UD_DATE2) < 0"
    End If
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    'MsgBox "Cri = " & sSQLQ, , "Criteria"
    ''''original Mostafa
    ''SQLQ = " SELECT HREMP.ED_EMPNBR, HREMP.ED_SURNAME, HREMP.ED_FNAME, HREMP.ED_SIN, HREMP.ED_DOH, HREMP.ED_SSN, HRDEPT.DF_NAME, HRDEPT.DF_NBR, HRPARCO.PC_NAME, HRPARCO.PC_SERIAL, HR_JOB_HISTORY.JH_CURRENT, tblLicProvState.TB_DESC, HRJOB.JB_DESCR, tblLicStatus.TB_DESC, HR_USERDEFINE_TABLE.UD_DATE2, HR_USERDEFINE_TABLE.UD_CODE1"
    'SQLQ = SQLQ & " FROM HREMP HREMP INNER JOIN HRPARCO HRPARCO ON HREMP.ED_COMPNO = HRPARCO.PC_CO LEFT OUTER JOIN HR_JOB_HISTORY HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR INNER JOIN HR_USERDEFINE_TABLE HR_USERDEFINE_TABLE ON HREMP.ED_EMPNBR = HR_USERDEFINE_TABLE.UD_EMPNBR LEFT OUTER JOIN HRDEPT HRDEPT ON HREMP.ED_DEPTNO = HRDEPT.DF_NBR INNER JOIN HRJOB HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE LEFT OUTER JOIN HRTABL tblLicStatus ON HR_USERDEFINE_TABLE.UD_CODE2 = tblLicStatus.TB_KEY AND HR_USERDEFINE_TABLE.UD_CODE2_TABL = tblLicStatus.TB_NAME LEFT OUTER JOIN HRTABL tblLicProvState ON HR_USERDEFINE_TABLE.UD_CODE1 = tblLicProvState.TB_KEY AND HR_USERDEFINE_TABLE.UD_CODE1_TABL = tblLicProvState.TB_NAME"
    'SQLQ = SQLQ & " WHERE HR_JOB_HISTORY.JH_CURRENT = 1 And " & sSQLQ & xpirycondtion
    'SQLQ = SQLQ & " GROUP by HRDEPT.DF_NAME,HREMP.ED_EMPNBR, HREMP.ED_SURNAME, HREMP.ED_FNAME, HREMP.ED_SIN, HREMP.ED_DOH, HREMP.ED_SSN, HRDEPT.DF_NBR, HRPARCO.PC_NAME, HRPARCO.PC_SERIAL, HR_JOB_HISTORY.JH_CURRENT, tblLicProvState.TB_DESC, HRJOB.JB_DESCR, tblLicStatus.TB_DESC, HR_USERDEFINE_TABLE.UD_DATE2, HR_USERDEFINE_TABLE.UD_CODE1"
    'SQLQ = SQLQ & " ORDER BY HRDEPT.DF_NAME asc, HREMP.ED_SURNAME asc, HREMP.ED_FNAME asc, tblLicProvState.TB_DESC ASC"
    'MsgBox "SQLQ = " & SQLQ, , "SQLQ"
    'WriteFile SQLQ
    
    SQLQ = "SELECT distinct DF_NBR, HRDEPT.DF_NAME, HREMP.ED_EMPNBR, HREMP.ED_SURNAME, HREMP.ED_FNAME FROM HREMP HREMP"
    SQLQ = SQLQ & " INNER JOIN HRPARCO HRPARCO ON HREMP.ED_COMPNO = HRPARCO.PC_CO LEFT OUTER JOIN HR_JOB_HISTORY"
    SQLQ = SQLQ & " HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR INNER JOIN HR_USERDEFINE_TABLE"
    SQLQ = SQLQ & " HR_USERDEFINE_TABLE ON HREMP.ED_EMPNBR = HR_USERDEFINE_TABLE.UD_EMPNBR LEFT OUTER JOIN HRDEPT"
    SQLQ = SQLQ & " HRDEPT ON HREMP.ED_DEPTNO = HRDEPT.DF_NBR INNER JOIN HRJOB HRJOB ON HR_JOB_HISTORY.JH_JOB ="
    SQLQ = SQLQ & " HRJOB.JB_CODE LEFT OUTER JOIN HRTABL tblLicStatus ON HR_USERDEFINE_TABLE.UD_CODE2 ="
    SQLQ = SQLQ & " tblLicStatus.TB_KEY AND HR_USERDEFINE_TABLE.UD_CODE2_TABL = tblLicStatus.TB_NAME LEFT"
    SQLQ = SQLQ & " OUTER JOIN HRTABL tblLicProvState ON HR_USERDEFINE_TABLE.UD_CODE1 = tblLicProvState.TB_KEY AND"
    SQLQ = SQLQ & " HR_USERDEFINE_TABLE.UD_CODE1_TABL = tblLicProvState.TB_NAME"
    SQLQ = SQLQ & " WHERE HR_JOB_HISTORY.JH_CURRENT = 1 AND " & sSQLQ & xpirycondtion
    SQLQ = SQLQ & " GROUP by HRDEPT.DF_NAME,HREMP.ED_EMPNBR, HREMP.ED_SURNAME, HREMP.ED_FNAME, HREMP.ED_SIN,"
    SQLQ = SQLQ & " HREMP.ED_DOH, HREMP.ED_SSN, HRDEPT.DF_NBR, HRPARCO.PC_NAME, HRPARCO.PC_SERIAL,"
    SQLQ = SQLQ & " HR_JOB_HISTORY.JH_CURRENT, tblLicProvState.TB_DESC, HRJOB.JB_DESCR, tblLicStatus.TB_DESC,"
    SQLQ = SQLQ & " HR_USERDEFINE_TABLE.UD_DATE2, HR_USERDEFINE_TABLE.UD_CODE1 ORDER BY HRDEPT.DF_NAME asc,"
    SQLQ = SQLQ & " HREMP.ED_SURNAME asc, HREMP.ED_FNAME asc"
    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsHRLic.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    If Not rsHRLic.EOF Then
        totNum = rsHRLic.RecordCount: I = 0
        rsHRLic.MoveFirst

        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2390_License.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2390_License" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
         Dim pc, idx
        
        If setcolumnwidth = False Then
               pc = GetCollectCorpStatesCount + 2
               idx = 1
               
               Do While idx <= pc
                    
                   exSheet.Columns(idx).ColumnWidth = 20
                    
                   idx = idx + 1
               Loop
               setcolumnwidth = True
           End If
        exSheet.Cells(2, 1) = "As Of " & Format(Now, "mmm dd, yyyy")
        If Len(dlpDateRange(4).Text) > 0 And Len(dlpDateRange(5).Text) Then
            exSheet.Cells(5, 1) = "License Expiry Range:  "
            exSheet.Cells(5, 2) = Format(dlpDateRange(4).Text, "dd-mmm-yy") & " - " & Format(dlpDateRange(5).Text, "dd-mmm-yy")
        End If
        Dim deptname As String
       
        
        deptname = ""
      
       
        'exSheet.Cells(7, 1) = "Department: " & rsHRLic("DF_NBR") & " - " & rsHRLic("DF_NAME")
         'exSheet.Cells(7, 1).Font.Bold = True
         'exSheet.Range(exSheet.Cells(7, 1), exSheet.Cells(7, 1)).Columns.AutoFit
      xRow = 7
      Dim f
      f = True
      
        Dim rc
        'Columns: 1 - Name, 3 - Job Title, 4 - Supervisor, 5 - Previous Vac, 6 - Current Vac, 7 - Taken Vac, 8 - Outstanding Vac, 9 - Salary, 10 - Cost of Oustanding Vacation
        Do While Not rsHRLic.EOF
            rc = 0
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
           
         
           
           
           If deptname <> rsHRLic("DF_NAME") Then
              
              xRow = xRow + 1
            If (f = False) Then
              exSheet.Cells(xRow, 1) = "Total number of records for this department: " & Total
             
              exSheet.Range(exSheet.Cells(xRow, 1), exSheet.Cells(xRow, 1)).Columns.AutoFit
              xRow = xRow + 2
             

              Total = 0
              
             End If
             exSheet.Cells(xRow, 1) = "Department: " & rsHRLic("DF_NBR") & " - " & rsHRLic("DF_NAME")
             exSheet.Cells(xRow, 1).Font.Bold = True
             exSheet.Range(exSheet.Cells(xRow, 1), exSheet.Cells(xRow, 1)).Columns.AutoFit
             ' Call License_Report_XLS_Headings_CCORP(CInt(xRow), exSheet, rsHRLic("DF_NBR") & " - " & rsHRLic("DF_NAME"))
              xRow = xRow + 3
              deptname = rsHRLic("DF_NAME")
              f = False
           End If
           
            exSheet.Cells(xRow, 1) = rsHRLic("ED_EMPNBR")
             exSheet.Cells(xRow, 1).Font.Bold = True
           
            exSheet.Cells(xRow, 2) = rsHRLic("ED_SURNAME") & ", " & rsHRLic("ED_FNAME")
            exSheet.Cells(xRow, 2).Font.Bold = True
             

            Dim col As Integer
            col = 3
            
            Set rsTmp = GetCollectCorpByEmpnbr(rsHRLic("ED_EMPNBR"), sSQLQ & xpirycondtion)
            
            rc = rsTmp.RecordCount
            If Not rsTmp.EOF Then rsTmp.MoveFirst
            
            Do While Not rsTmp.EOF
                exSheet.Cells(xRow, col) = GetCollectCorpStateName(rsTmp("UD_CODE1"))
                exSheet.Cells(xRow, col).Font.Bold = True
                
                 col = col + 1
                rsTmp.MoveNext
              
            Loop
            xRow = xRow + 1
            If rc > 0 Then rsTmp.MoveFirst
            col = 3
            Do While Not rsTmp.EOF
                exSheet.Cells(xRow, col) = rsTmp("TB_DESC")
                col = col + 1
                rsTmp.MoveNext
              
            Loop
             xRow = xRow + 1
            If rc > 0 Then rsTmp.MoveFirst
            col = 3
            Do While Not rsTmp.EOF
                exSheet.Cells(xRow, col) = Format(rsTmp("UD_DATE2"), "dd-mmm-yy")
                exSheet.Cells(xRow, col).HorizontalAlignment = xlLeft
                col = col + 1
                rsTmp.MoveNext
            
            Loop
             xRow = xRow + 1
            If rc > 0 Then rsTmp.MoveFirst
            col = 3
            Do While Not rsTmp.EOF
               If IsDate(rsTmp("UD_DATE2")) Then
                    If DateDiff("d", Now(), CDate(rsTmp("UD_DATE2"))) < 0 Then
                         exSheet.Cells(xRow, col) = "Expired"
                    
                    End If
                End If
                col = col + 1
                rsTmp.MoveNext
            Loop
            
            'exSheet.Range(exSheet.Cells(xRow, 1), exSheet.Cells(xRow, col)).Columns.AutoFit
              
                
                
                
            If rsTmp.State <> 0 Then rsTmp.Close
            rsHRLic.MoveNext
            
                xRow = xRow + 2
            
            Total = Total + 1
        Loop
      
        exSheet.Cells(xRow + 2, 1) = "Total number of records for this department: " & Total
        xRow = xRow + 2
        exSheet.Cells(xRow + 2, 1) = "Total number of records in this worksheet: " & totNum
        'exSheet.Cells(xRow + 3, 1) = "Total Number of Days of Vacation Outstanding as at Current Month End: " & xTotVacOut
        exSheet.Rows(xRow + 2).Font.Bold = True
        'exSheet.Rows(xRow + 3).Font.Bold = True
        
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
        If totNum = 0 Then
            MsgBox "No records were found!", , "Message!"
        End If
    Else
        MsgBox "No records were found!", , "Message!"
    End If
        
   
     rsHRLic.Close
Exit Sub


License_Report_XLS_CCORP_Err:


glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
Resume Next

End Sub
Public Function GetCollectCorpStatesCount() As Integer
Dim rsCode As New ADODB.Recordset

   
        rsCode.Open "select distinct TB_DESC from HRTABL where TB_NAME = 'COD1'", gdbAdoIhr001, adOpenStatic
        If rsCode.EOF Then
            GetCollectCorpStatesCount = 0
        Else
            GetCollectCorpStatesCount = rsCode.RecordCount
        End If
        If rsCode.State <> 0 Then rsCode.Close
   
    
    
End Function


Public Function GetCollectCorpByEmpnbr(EmpNbr, where) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    
    Dim SQLQ
    
    SQLQ = "SELECT  HREMP.ED_SIN, HREMP.ED_DOH, HREMP.ED_SSN, HRDEPT.DF_NAME, HRDEPT.DF_NBR, HRPARCO.PC_NAME,"
    SQLQ = SQLQ & " HRPARCO.PC_SERIAL, HR_JOB_HISTORY.JH_CURRENT, tblLicProvState.TB_DESC, HRJOB.JB_DESCR, "
    SQLQ = SQLQ & " tblLicStatus.TB_DESC, HR_USERDEFINE_TABLE.UD_DATE2, HR_USERDEFINE_TABLE.UD_CODE1 "
    SQLQ = SQLQ & " FROM HREMP HREMP INNER JOIN HRPARCO HRPARCO ON HREMP.ED_COMPNO = HRPARCO.PC_CO LEFT"
    SQLQ = SQLQ & " OUTER JOIN HR_JOB_HISTORY HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR"
    SQLQ = SQLQ & " INNER JOIN HR_USERDEFINE_TABLE HR_USERDEFINE_TABLE ON HREMP.ED_EMPNBR = "
    SQLQ = SQLQ & " HR_USERDEFINE_TABLE.UD_EMPNBR LEFT OUTER JOIN HRDEPT HRDEPT ON HREMP.ED_DEPTNO = "
    SQLQ = SQLQ & " HRDEPT.DF_NBR INNER JOIN HRJOB HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE LEFT "
    SQLQ = SQLQ & " OUTER JOIN HRTABL tblLicStatus ON HR_USERDEFINE_TABLE.UD_CODE2 = tblLicStatus.TB_KEY "
    SQLQ = SQLQ & " AND HR_USERDEFINE_TABLE.UD_CODE2_TABL = tblLicStatus.TB_NAME LEFT OUTER JOIN HRTABL "
    SQLQ = SQLQ & " tblLicProvState ON HR_USERDEFINE_TABLE.UD_CODE1 = tblLicProvState.TB_KEY AND"
    SQLQ = SQLQ & " HR_USERDEFINE_TABLE.UD_CODE1_TABL = tblLicProvState.TB_NAME WHERE HR_JOB_HISTORY.JH_CURRENT = 1"
    SQLQ = SQLQ & " and hremp.ed_empnbr =" & EmpNbr & " AND " & where & " ORDER BY  tblLicProvState.TB_DESC ASC"
    
    'Call WriteFile("SQL2=" & SQLQ)
            rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
      
    Set GetCollectCorpByEmpnbr = rs
End Function

Public Function GetCollectCorpStateName(ProvCode)
Dim rsCode As New ADODB.Recordset
Dim xDesc
    If Len(ProvCode) = 0 Then
        xDesc = ""
    Else
        rsCode.Open "select * from hrtabl where tb_name = 'COD1' and TB_KEY ='" & ProvCode & "' ", gdbAdoIhr001, adOpenStatic
        If rsCode.EOF Then
            xDesc = ProvCode
        Else
            xDesc = rsCode("TB_DESC")
        End If
        rsCode.Close
    End If
    GetCollectCorpStateName = xDesc
End Function

Function LanchXlsW98(xFileName)
On Error GoTo Error_Deal
    LanchXlsW98 = False
    Shell "Start " & GetShortName(xFileName)
    LanchXlsW98 = True
Exit Function
Error_Deal:

End Function

Private Sub Overtime_Report_XLS_HCAS()
    On Error GoTo Overtime_Report_XLS_HCAS_Err

    Dim rsHREmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    Dim xHourlyRate

    'Check the Month End Date is within the Current Vacation Year
    If Len(Trim(dlpAsOf.Text)) = 0 Then
        MsgBox "Month Ending Date cannot be blank"
        dlpAsOf.SetFocus
    ElseIf Not IsDate(dlpAsOf.Text) Then
        MsgBox "Invalid Month Ending Date"
        dlpAsOf.SetFocus
    End If
    
    sSQLQ = Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")")
    SQLQ = "SELECT HREMP.ED_EMPNBR AS EEMPNBR, HREMP.ED_FNAME AS EFNAME, HREMP.ED_SURNAME AS ESURNAME, "
    SQLQ = SQLQ & " ((CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA <" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA <" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & ") END) - "
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA <" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA <" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & ") END)) + "
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") END) - "
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") END) AS OT_OUTST, "
    SQLQ = SQLQ & " JH_JOB, JB_DESCR, JH_REPTAU, SH_SALARY, JH_DHRS,JH_WHRS,SH_SALCD,SUPER.ED_SURNAME AS SSURNAME,SUPER.ED_FNAME AS SFNAME,"
    SQLQ = SQLQ & " ((CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA <" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA <" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & ") END) - "
    SQLQ = SQLQ & "  (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA <" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA <" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & ") END)) AS OT_PREV, "
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") END) AS OT_CURR,"
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") END) AS OT_TAKEN"
    SQLQ = SQLQ & " FROM ((((HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR AND JH_CURRENT <>0)"
    SQLQ = SQLQ & " INNER JOIN HR_SALARY_HISTORY ON HR_SALARY_HISTORY.SH_EMPNBR = ED_EMPNBR AND SH_CURRENT <>0)"
    SQLQ = SQLQ & " INNER JOIN HRJOB ON HRJOB.JB_CODE = HR_JOB_HISTORY.JH_JOB)"
    SQLQ = SQLQ & " INNER JOIN HREMP SUPER ON SUPER.ED_EMPNBR = HR_JOB_HISTORY.JH_REPTAU)"
    SQLQ = SQLQ & " GROUP BY HREMP.ED_EMPNBR,HREMP.ED_FNAME,HREMP.ED_SURNAME,JH_JOB,JB_DESCR,JH_REPTAU,SH_SALARY,JH_DHRS,JH_WHRS,SH_SALCD,SUPER.ED_FNAME,SUPER.ED_SURNAME"
    SQLQ = SQLQ & " ORDER BY SSURNAME, SFNAME, JB_DESCR, OT_OUTST ASC"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: I = 0
        rsHREmp.MoveFirst

        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "OvertimeRptTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "OvertimeRpt" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
    
        exSheet.Cells(1, 2) = Format(Now, "mmm dd, yyyy")
        exSheet.Cells(2, 2) = Time$
        exSheet.Cells(5, 1) = "Report for the Month Ending: " & Format(dlpAsOf.Text, "dd-mmm-yy")
        
        xRow = 11
        'Columns: 1 - Name, 3 - Job Title, 4 - Supervisor, 5 - Previous Vac, 6 - Current Vac, 7 - Taken Vac, 8 - Outstanding Vac, 9 - Salary, 10 - Cost of Oustanding Vacation
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            xHourlyRate = 0
            If rsHREmp("SH_SALCD") = "H" Then
                xHourlyRate = rsHREmp("SH_SALARY")
            ElseIf rsHREmp("SH_SALCD") = "A" And rsHREmp("JH_WHRS") <> 0 Then
                xHourlyRate = Round((rsHREmp("SH_SALARY") / 52) / rsHREmp("JH_WHRS"), 2)
            End If
            
            exSheet.Cells(xRow, 1) = rsHREmp("ESURNAME") & ", " & rsHREmp("EFNAME")
            exSheet.Cells(xRow, 3) = rsHREmp("JB_DESCR")
            'If rsHREmp("JH_REPTAU") <> "" And Not IsNull(rsHREmp("JH_REPTAU")) Then
                exSheet.Cells(xRow, 4) = rsHREmp("SSURNAME") & ", " & rsHREmp("SFNAME")
            'End If
            'If rsHREmp("JH_DHRS") <> 0 And Not IsNull(rsHREmp("JH_DHRS")) Then
                exSheet.Cells(xRow, 5) = Round(rsHREmp("OT_PREV"), 2) ' / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 6) = Round(rsHREmp("OT_CURR"), 2) '/ rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 7) = Round(rsHREmp("OT_TAKEN"), 2) ' / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 8) = Round(rsHREmp("OT_OUTST"), 2) ' / rsHREmp("JH_DHRS"), 2)
                'exSheet.Cells(xRow, 9) = Format(Round((rsHREmp("OT_OUTST") / rsHREmp("JH_DHRS")), 2) * xHourlyRate, "#,##0.00")
                exSheet.Cells(xRow, 9) = Format(Round(rsHREmp("OT_OUTST"), 2) * xHourlyRate, "#,##0.00")
            'End If
            exSheet.Cells(xRow, 10) = Format(xHourlyRate, "#,##0.00")
            
            rsHREmp.MoveNext
            xRow = xRow + 1
        Loop
        
                
        exSheet.Cells(xRow + 2, 1) = "Total Number of Employees Reported: " & totNum
        exSheet.Rows(xRow + 2).Font.Bold = True

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    rsHREmp.Close
    
Exit Sub

Overtime_Report_XLS_HCAS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
Resume Next

End Sub

Private Sub Attendance_Report_XLS_HCAS()
    On Error GoTo Attendance_Report_XLS_HCAS_Err

    Dim rsHREmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    Dim xHourlyRate
    Dim xAbsentDays, xCostAbsences

    'Check the Month End Date is within the Current Vacation Year
    If Len(Trim(dlpAsOf.Text)) = 0 Then
        MsgBox "Month Ending Date cannot be blank"
        dlpAsOf.SetFocus
    ElseIf Not IsDate(dlpAsOf.Text) Then
        MsgBox "Invalid Month Ending Date"
        dlpAsOf.SetFocus
    End If
    
    sSQLQ = Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")")
    SQLQ = "SELECT HREMP.ED_EMPNBR AS EEMPNBR, HREMP.ED_FNAME AS EFNAME, HREMP.ED_SURNAME AS ESURNAME, "
    SQLQ = SQLQ & " JH_JOB, JB_DESCR, JH_REPTAU, SH_SALARY, JH_DHRS,JH_WHRS,SH_SALCD,SUPER.ED_SURNAME AS SSURNAME,SUPER.ED_FNAME AS SFNAME,"
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND AD_DOA >=" & Date_SQL("01/" & month(dlpAsOf.Text) & "/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='ADRE' AND TB_ABSENCE=1)) IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND AD_DOA >=" & Date_SQL("01/" & month(dlpAsOf.Text) & "/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='ADRE' AND TB_ABSENCE=1)) END) AS CURR_ABSENT,"
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text) & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text) & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='ADRE' AND TB_ABSENCE=1)) IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text) & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text) & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='ADRE' AND TB_ABSENCE=1)) END) AS ABSENT_12"
    SQLQ = SQLQ & " FROM ((((HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR AND JH_CURRENT <>0)"
    SQLQ = SQLQ & " INNER JOIN HR_SALARY_HISTORY ON HR_SALARY_HISTORY.SH_EMPNBR = ED_EMPNBR AND SH_CURRENT <>0)"
    SQLQ = SQLQ & " INNER JOIN HRJOB ON HRJOB.JB_CODE = HR_JOB_HISTORY.JH_JOB)"
    SQLQ = SQLQ & " INNER JOIN HREMP SUPER ON SUPER.ED_EMPNBR = HR_JOB_HISTORY.JH_REPTAU)"
    SQLQ = SQLQ & " GROUP BY HREMP.ED_EMPNBR,HREMP.ED_FNAME,HREMP.ED_SURNAME,JH_JOB,JB_DESCR,JH_REPTAU,SH_SALARY,JH_DHRS,JH_WHRS,SH_SALCD,SUPER.ED_FNAME,SUPER.ED_SURNAME"
    SQLQ = SQLQ & " ORDER BY SSURNAME, SFNAME, JB_DESCR"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: I = 0
        rsHREmp.MoveFirst

        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "AttendanceRptTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "AttendanceRpt" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
    
        exSheet.Cells(1, 2) = Format(Now, "mmm dd, yyyy")
        exSheet.Cells(2, 2) = Time$
        exSheet.Cells(5, 1) = "Report for the Month Ending: " & Format(dlpAsOf.Text, "dd-mmm-yy")
        
        xAbsentDays = 0
        xCostAbsences = 0
        xRow = 11
        'Columns: 1 - Name, 3 - Job Title, 4 - Supervisor, 5 - Previous Vac, 6 - Current Vac, 7 - Taken Vac, 8 - Outstanding Vac, 9 - Salary, 10 - Cost of Oustanding Vacation
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            xHourlyRate = 0
            If rsHREmp("SH_SALCD") = "H" Then
                xHourlyRate = rsHREmp("SH_SALARY")
            ElseIf rsHREmp("SH_SALCD") = "A" And rsHREmp("JH_WHRS") <> 0 Then
                xHourlyRate = Round((rsHREmp("SH_SALARY") / 52) / rsHREmp("JH_WHRS"), 2)
            End If
            
            exSheet.Cells(xRow, 1) = rsHREmp("ESURNAME") & ", " & rsHREmp("EFNAME")
            exSheet.Cells(xRow, 3) = rsHREmp("JB_DESCR")
            exSheet.Cells(xRow, 4) = rsHREmp("SSURNAME") & ", " & rsHREmp("SFNAME")
            
            If rsHREmp("JH_DHRS") <> 0 And Not IsNull(rsHREmp("JH_DHRS")) Then
                exSheet.Cells(xRow, 5) = Round(rsHREmp("CURR_ABSENT") / rsHREmp("JH_DHRS"), 2)
                xAbsentDays = xAbsentDays + Round(rsHREmp("CURR_ABSENT") / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 6) = Round(rsHREmp("ABSENT_12") / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 8) = Format(Round((rsHREmp("ABSENT_12") / rsHREmp("JH_DHRS")), 2) * xHourlyRate, "#,##0.00")
                xCostAbsences = xCostAbsences + Format(Round((rsHREmp("ABSENT_12") / rsHREmp("JH_DHRS")), 2) * xHourlyRate, "#,##0.00")
            End If
            exSheet.Cells(xRow, 7) = Format(xHourlyRate, "#,##0.00")
            
            rsHREmp.MoveNext
            xRow = xRow + 1
        Loop
        
                
        exSheet.Cells(xRow + 2, 1) = "Total Number of Employees Reported: " & totNum
        exSheet.Rows(xRow + 2).Font.Bold = True
        exSheet.Cells(xRow + 3, 1) = "Total Number of Days Absent: " & xAbsentDays
        exSheet.Cells(xRow + 4, 1) = "Average Number of Days Absent: " & Round(xAbsentDays / totNum, 2)
        exSheet.Cells(xRow + 5, 1) = "Cost of Total Absences: " & Round(xCostAbsences, 2)

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    rsHREmp.Close
    
Exit Sub

Attendance_Report_XLS_HCAS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
Resume Next

End Sub


Private Sub Export_Data_to_Excel_CollectCorp()
    Dim rsHREmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
        
    On Error GoTo Err_Export_Data_to_Excel_CollectCorp
    
    
    Screen.MousePointer = HOURGLASS
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    If OptAct.Value Then
        SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO,ED_ADMINBY,ED_LOC,ED_DIV,ED_SECTION,ED_DRIVERLIC,"
        SQLQ = SQLQ & "ED_EMP,ED_PT,ED_DOH,JH_JOB,JH_REPTAU,JH_REPTAU2 "
        SQLQ = SQLQ & " FROM (HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=JH_EMPNBR AND JH_CURRENT <>0) "
        SQLQ = SQLQ & " WHERE 1 = 1"
    Else 'Term
        SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO,ED_ADMINBY,ED_LOC,ED_DIV,ED_SECTION,ED_DRIVERLIC,"
        SQLQ = SQLQ & "ED_EMP,ED_PT,ED_DOH,JH_JOB,JH_REPTAU,JH_REPTAU2 "
        SQLQ = SQLQ & " FROM (Term_HREMP INNER JOIN Term_JOB_HISTORY ON ED_EMPNBR=JH_EMPNBR AND JH_CURRENT <>0) "
        SQLQ = SQLQ & " WHERE 1 = 1"
        sSQLQ = Replace(sSQLQ, "HREMP.", "Term_HREMP.")
    End If
    SQLQ = SQLQ & " AND " & sSQLQ & " "
    SQLQ = SQLQ & " ORDER BY ED_EMPNBR,ED_SURNAME,ED_FNAME "

    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
                
        'File to export to
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "Export_EmployeeList.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "EmployeeList_" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
        
        FileCopy xlsFileTmp, xlsFileMat
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
    
        exSheet.Cells(1, 1) = "Employee #"
        exSheet.Cells(1, 2) = "Department Code"
        exSheet.Cells(1, 3) = "Department Name"
        exSheet.Cells(1, 4) = "Surname"
        exSheet.Cells(1, 5) = "First Name"
        exSheet.Cells(1, 6) = "Administered By" 'Name
        exSheet.Cells(1, 7) = "Location"    'Name
        exSheet.Cells(1, 8) = "Division"    'Name
        exSheet.Cells(1, 9) = "Site"        'Section Name
        exSheet.Cells(1, 10) = "User ID"     'Driver License
        exSheet.Cells(1, 11) = "Status"      'drop down list
        exSheet.Cells(1, 12) = "FT/PT/TEMP/UNKN"
        exSheet.Cells(1, 13) = "Original Hire Date"
        exSheet.Cells(1, 14) = "Reporting Authority 1"   'Name
        exSheet.Cells(1, 15) = "Reporting Authority 2"   'Name
        exSheet.Cells(1, 16) = "Position Code"           'Code
        exSheet.Cells(1, 17) = "Position Name"           'Name
        
        xRow = 2
        
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            exSheet.Cells(xRow, 1) = rsHREmp("ED_EMPNBR")
            exSheet.Cells(xRow, 2) = rsHREmp("ED_DEPTNO")
            exSheet.Cells(xRow, 3) = getDeptDesc(rsHREmp("ED_DEPTNO"))
            exSheet.Cells(xRow, 4) = rsHREmp("ED_SURNAME")
            exSheet.Cells(xRow, 5) = rsHREmp("ED_FNAME")
            exSheet.Cells(xRow, 6) = GetTABLDesc("EDAB", rsHREmp("ED_ADMINBY"))
            exSheet.Cells(xRow, 7) = GetTABLDesc("EDLC", rsHREmp("ED_LOC"))
            exSheet.Cells(xRow, 8) = getDivDesc(rsHREmp("ED_DIV"))
            exSheet.Cells(xRow, 9) = GetTABLDesc("EDSE", rsHREmp("ED_SECTION"))
            exSheet.Cells(xRow, 10) = rsHREmp("ED_DRIVERLIC")
            
            exSheet.Cells(xRow, 11) = GetTABLDesc("EDEM", rsHREmp("ED_EMP"))
            exSheet.Cells(xRow, 12) = GetTABLDesc("EDPT", rsHREmp("ED_PT"))
            exSheet.Cells(xRow, 13) = Format(rsHREmp("ED_DOH"), "mm/dd/yyyy")
                        
            exSheet.Cells(xRow, 14) = getEEName(rsHREmp("JH_REPTAU"))
            exSheet.Cells(xRow, 15) = getEEName(rsHREmp("JH_REPTAU2"))
            exSheet.Cells(xRow, 16) = rsHREmp("JH_JOB")
            exSheet.Cells(xRow, 17) = GetJobDesc(rsHREmp("JH_JOB"))
            
            rsHREmp.MoveNext
            xRow = xRow + 1
        Loop
        
        'exSheet.AutoFilterMode = True
        'exSheet.Range("K1:K20").AutoFilter Field:=1
        exSheet.Columns.AutoFit

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Err_Export_Data_to_Excel_CollectCorp:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
Resume Next

End Sub

Private Sub Export_Seniority_Excel()
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    Dim SQLX, WSQLQ
    Dim rsHRSen As New ADODB.Recordset
    Dim rsHRTabl As New ADODB.Recordset
    Dim xFld As String
    Dim xDeptno As String
    Dim xDeptRow As Integer
    Dim xDeptFT As String
    Dim xn

    On Error GoTo Export_Seniority_Excel_Err
    
    Screen.MousePointer = HOURGLASS
    
    'Hemu
    Set CN001 = New ADODB.Connection
    CN001.CommandTimeout = 600
    CN001.Open glbAdoIHRDB
    
    Set CN001W = New ADODB.Connection
    CN001W.CommandTimeout = 600
    CN001W.Open glbAdoIHRDBW
    
    CN001.BeginTrans
    CN001.Execute "DELETE FROM HRSENHRS " & in_SQL(glbIHRDBW) & " WHERE WRKEMP='" & glbUserID & "'"
    CN001.CommitTrans
        
    WSQLQ = glbSeleDeptUn
    If clpDiv.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DIV IN ('" & Replace(clpDiv.Text, ",", "','") & "')"
    If clpDept.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DEPTNO IN ('" & Replace(clpDept.Text, ",", "','") & "')"
    If clpCode(0).Text <> "" Then WSQLQ = WSQLQ & " AND ED_LOC IN ('" & Replace(clpCode(0).Text, ",", "','") & "')"
    If clpCode(1).Text <> "" Then WSQLQ = WSQLQ & " AND ED_ORG IN ('" & Replace(clpCode(1).Text, ",", "','") & "')"
    If clpCode(2).Text <> "" Then WSQLQ = WSQLQ & " AND ED_EMP IN ('" & Replace(clpCode(2).Text, ",", "','") & "')"
    If clpCode(3).Text <> "" Then WSQLQ = WSQLQ & " AND ED_REGION IN ('" & Replace(clpCode(3).Text, ",", "','") & "')"
    If clpCode(4).Text <> "" Then WSQLQ = WSQLQ & " AND ED_ADMINBY IN ('" & Replace(clpCode(4).Text, ",", "','") & "')"
    If clpCode(5).Text <> "" Then WSQLQ = WSQLQ & " AND ED_SECTION IN ('" & Replace(clpCode(5).Text, ",", "','") & "')"
    If clpPT.Text <> "" Then WSQLQ = WSQLQ & " AND ED_PT IN ('" & Replace(clpPT.Text, ",", "','") & "')"
       
    If elpEEID.Text <> "" Then WSQLQ = WSQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
        
        
    SQLQ = "SELECT '001' AS COMPANY,ED_EMPNBR,'EDOR' AS UNION_TBL,'EDEM' AS EMPLSTAT_TBL,"
    If glbOracle Then
        SQLQ = SQLQ & "SUBSTR(ED_SURNAME || ', ' || ED_FNAME,1,39) AS NAME,"
    Else
        SQLQ = SQLQ & "LEFT(ED_SURNAME +', '+ED_FNAME,39) AS NAME,"
    End If
    SQLQ = SQLQ & "ED_DIV,ED_DEPTNO,ED_ORG,ED_EMP,ED_DOH,0 AS NUMBHRS "
    SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
    SQLQ = SQLQ & "FROM HREMP"
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    
    SQLX = "INSERT INTO HRSENHRS "
    SQLX = SQLX & "(COMPANY,EMPLNUMB,UNION_TBL,EMPLSTAT_TBL,NAME,"
    SQLX = SQLX & "DIVISION,DEPTNO,UNIONS,EMPLSTAT,DOH,NUMBHRS,WRKEMP) "
    SQLX = SQLX & SQLQ & " WHERE " & WSQLQ
    
    
    CN001W.BeginTrans
    CN001W.Execute SQLX
    CN001W.CommitTrans
    
    'Hemu
    
    SQLQ = "SELECT AD_EMPNBR AS EMPNBR, SUM(AD_HRS) AS NUM1"
    SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
    SQLQ = SQLQ & " FROM HR_ATTENDANCE "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    
    SQLQ = SQLQ & " WHERE AD_SEN<>0 AND AD_DOA<=" & Date_SQL(dlpAsOf.Text)
    
    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " WHERE " & WSQLQ & ")"
    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
    If glbSQL Then
        SQLX = " Update HRSENHRS "
        SQLX = SQLX & " SET NUMBHRS = TRSNUM.NUM1"
        SQLX = SQLX & " FROM HRSENHRS INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRSENHRS.EMPLNUMB"
        gdbAdoIhr001.Execute SQLX
    Else
        'Hemu
        CN001W.BeginTrans
        CN001W.Execute "DELETE FROM TRSNUM WHERE WRKEMP='" & glbUserID & "'"
        CN001W.CommitTrans
        'Hemu
                
        SQLX = "INSERT INTO TRSNUM (EMPNBR,NUM1,WRKEMP) " & SQLQ
        'Hemu
        CN001W.BeginTrans
        CN001W.Execute SQLX
        CN001W.CommitTrans
        'Hemu
                
        SQLX = "Update HRSENHRS INNER JOIN TRSNUM ON TRSNUM.EMPNBR=HRSENHRS.EMPLNUMB"
        SQLX = SQLX & " SET NUMBHRS = TRSNUM.NUM1 "
        SQLX = SQLX & "WHERE TRSNUM.WRKEMP='" & glbUserID & "' "
        SQLX = SQLX & "AND HRSENHRS.WRKEMP='" & glbUserID & "' "
        'Hemu
        CN001W.BeginTrans
        CN001W.Execute SQLX
        CN001W.CommitTrans
        'Hemu
    End If
    
    If chkInclAtt Then
        SQLQ = "SELECT AH_EMPNBR AS EMPNBR, SUM(AH_HRS) AS NUM1"
        SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
        SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
        SQLQ = SQLQ & in_SQL(glbIHRDB)
        SQLQ = SQLQ & " WHERE AH_SEN<>0 AND AH_DOA<=" & Date_SQL(dlpAsOf.Text)
        
        SQLQ = SQLQ & " AND AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP "
        SQLQ = SQLQ & in_SQL(glbIHRDB)
        SQLQ = SQLQ & " WHERE " & WSQLQ & ")"
        SQLQ = SQLQ & " GROUP BY AH_EMPNBR"

        If glbSQL Then
            SQLX = " Update HRSENHRS "
            SQLX = SQLX & " SET NUMBHRS = ISNULL(NUMBHRS,0)+TRSNUM.NUM1"
            SQLX = SQLX & " FROM HRSENHRS INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRSENHRS.EMPLNUMB"
            gdbAdoIhr001.Execute SQLX
        Else
            'Hemu
            CN001W.BeginTrans
            CN001W.Execute "DELETE FROM TRSNUM WHERE WRKEMP='" & glbUserID & "'"
            CN001W.CommitTrans
            'Hemu
            
            SQLX = "INSERT INTO TRSNUM (EMPNBR,NUM1,WRKEMP) " & SQLQ
            
            'Hemu
            CN001W.BeginTrans
            CN001W.Execute SQLX
            CN001W.CommitTrans
            'Hemu
            
            SQLX = "Update HRSENHRS INNER JOIN TRSNUM ON TRSNUM.EMPNBR=HRSENHRS.EMPLNUMB"
            SQLX = SQLX & " SET NUMBHRS = NUMBHRS + TRSNUM.NUM1 "
            SQLX = SQLX & "WHERE TRSNUM.WRKEMP='" & glbUserID & "' "
            SQLX = SQLX & "AND HRSENHRS.WRKEMP='" & glbUserID & "' "
            
            'Hemu
            'gdbAdoIhr001W.Execute SQLX
            CN001W.BeginTrans
            CN001W.Execute SQLX
            CN001W.CommitTrans
            'Hemu
        End If
    End If
    
    CN001.Close
    CN001W.Close
    
    Set CN001 = Nothing
    Set CN001W = Nothing


    'Write to Excel Spreadsheet
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME,ED_PT,ED_DIV,ED_DEPTNO,ED_UNION,ED_PHONE,ED_BUSNBR FROM HREMP "
    SQLQ = SQLQ & "WHERE ED_PT IN ('FT','PT') "
    SQLQ = SQLQ & "AND " & WSQLQ
    SQLQ = SQLQ & "ORDER BY ED_DEPTNO ASC, ED_PT ASC,ED_UNION DESC"
    rsHRSen.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    If Not rsHRSen.EOF Then
        totNum = rsHRSen.RecordCount: I = 0
        rsHRSen.MoveFirst
                
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SeniorityList_Tmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SeniorityList" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
    
        exSheet.Cells(2, 1) = Format(Now, "mm/dd/yyyy")
        exSheet.Cells(2, 2) = "as of  " & Format(dlpAsOf.Text, "mmmm dd, yyyy")
        exSheet.Cells(1, 2) = "WDDS Seniority Report By Department"
        
        xDeptno = ""
        xRow = 3
        xDeptRow = 0
        xDeptFT = ""
        Do While Not rsHRSen.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
                    
            If xDeptno <> rsHRSen("ED_DEPTNO") Then
                If xDeptRow <> 0 Then
                    If xDeptFT <> "FT" Then
                        xn = exSheet.Range("A" & xDeptRow & ":E" & xRow - 1).Sort(exSheet.Range("B" & xDeptRow & ":B" & xRow - 1), xlDescending) ', , , , , , xlNo, , , xlSortColumns)
                    Else
                        xn = exSheet.Range("A" & xDeptRow & ":E" & xRow - 1).Sort(exSheet.Range("B" & xDeptRow & ":B" & xRow - 1), xlAscending) ', , , , , , xlNo, , , xlSortColumns)
                    End If
                End If
                xRow = xRow + 1
                exSheet.Cells(xRow, 1) = getDeptDesc(rsHRSen("ED_DEPTNO"))
                exSheet.Rows(xRow).Font.Bold = True
                xRow = xRow + 1
                xDeptno = rsHRSen("ED_DEPTNO")
                xDeptRow = xRow
            End If
            
            'Name
            exSheet.Cells(xRow, 1) = rsHRSen("ED_SURNAME") & ", " & rsHRSen("ED_FNAME")
            
            'If rsHRSen("ED_EMPNBR") = 855 Then
            '    MsgBox ""
            'End If
            
            If rsHRSen("ED_PT") = "FT" Then
                exSheet.Cells(xRow, 2) = IIf(Not IsNull(rsHRSen("ED_UNION")), Format(rsHRSen("ED_UNION"), "mm/dd/yyyy"), "")
                xDeptFT = "FT"
            Else
                'PT
                If rsHRSen("ED_PT") = "PT" Then
                    exSheet.Cells(xRow, 2) = get_Seniority_Hours(rsHRSen("ED_EMPNBR"))
                    xDeptFT = "PT"
                End If
            End If
            
            'Phone #s
            exSheet.Cells(xRow, 3) = Format(rsHRSen("ED_PHONE"), "(###) ###-####")
            exSheet.Cells(xRow, 4) = Format(rsHRSen("ED_BUSNBR"), "(###) ###-####")
            
            'Status
            exSheet.Cells(xRow, 5) = get_Employee_Flag11(rsHRSen("ED_EMPNBR"))
            
            xRow = xRow + 1
            
            rsHRSen.MoveNext
        Loop
        
        'The last group needs to be sorted - in the above loop, the sorting only takes place while not eof.
        If xDeptFT <> "FT" Then
            xn = exSheet.Range("A" & xDeptRow & ":E" & xRow - 1).Sort(exSheet.Range("B" & xDeptRow & ":B" & xRow - 1), xlDescending) ', , , , , , xlNo, , , xlSortColumns)
        Else
            xn = exSheet.Range("A" & xDeptRow & ":E" & xRow - 1).Sort(exSheet.Range("B" & xDeptRow & ":B" & xRow - 1), xlAscending) ', , , , , , xlNo, , , xlSortColumns)
        End If
        
        exSheet.Range("A" & xRow & ":Q423").Delete
        exSheet.Range("A" & xRow - 1 & ":Q" & xRow - 1).Borders(xlEdgeBottom).Weight = xlThick
        
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If

    Screen.MousePointer = DEFAULT

Exit Sub

Export_Seniority_Excel_Err:
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
    Exit Sub
End Sub

Private Function get_Employee_Flag11(xEmpNbr)
    Dim rsEmpFlag As New ADODB.Recordset
    Dim SQLQ As String
    
    get_Employee_Flag11 = ""
    
    SQLQ = "SELECT EF_EMPNBR, EF_FLAGVAL11, EF_FLAGDTE11 FROM HREMP_FLAGS WHERE EF_EMPNBR = " & xEmpNbr
    rsEmpFlag.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    If Not rsEmpFlag.EOF Then
        If Not IsNull(rsEmpFlag("EF_FLAGVAL11")) And rsEmpFlag("EF_FLAGVAL11") <> "" Then
            get_Employee_Flag11 = rsEmpFlag("EF_FLAGVAL11") & " (" & rsEmpFlag("EF_FLAGDTE11") & ")"
        Else
            get_Employee_Flag11 = ""
        End If
    Else
        get_Employee_Flag11 = ""
    End If
    rsEmpFlag.Close
    Set rsEmpFlag = Nothing
End Function

Private Function get_Seniority_Hours(xEmpNbr)
    Dim rsHRSenHrs As New ADODB.Recordset
    Dim SQLQ As String
    
    get_Seniority_Hours = ""
    
    SQLQ = "SELECT EMPLNUMB, NUMBHRS FROM HRSENHRS WHERE EMPLNUMB = " & xEmpNbr
    rsHRSenHrs.Open SQLQ, gdbAdoIhr001W, adOpenDynamic, adLockOptimistic
    If Not rsHRSenHrs.EOF Then
        If Not IsNull(rsHRSenHrs("NUMBHRS")) And rsHRSenHrs("NUMBHRS") <> "" Then
            get_Seniority_Hours = Format(rsHRSenHrs("NUMBHRS"), "#,##0.00")
        Else
            get_Seniority_Hours = ""
        End If
    Else
        get_Seniority_Hours = ""
    End If
    rsHRSenHrs.Close
    Set rsHRSenHrs = Nothing
    
End Function

Private Sub Export_FT_Seniority_Excel()
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    Dim SQLX, WSQLQ
    Dim rsHRSen As New ADODB.Recordset
    Dim rsHRTabl As New ADODB.Recordset
    Dim xFld As String
    Dim xDeptno As String
    Dim xDeptRow As Integer
    Dim xn

    On Error GoTo Export_FT_Seniority_Excel_Err
    
    Screen.MousePointer = HOURGLASS
    
    'Hemu
    Set CN001 = New ADODB.Connection
    CN001.CommandTimeout = 600
    CN001.Open glbAdoIHRDB
    
    Set CN001W = New ADODB.Connection
    CN001W.CommandTimeout = 600
    CN001W.Open glbAdoIHRDBW
    
    CN001.BeginTrans
    CN001.Execute "DELETE FROM HRSENHRS " & in_SQL(glbIHRDBW) & " WHERE WRKEMP='" & glbUserID & "'"
    CN001.CommitTrans
        
    WSQLQ = glbSeleDeptUn
    If clpDiv.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DIV IN ('" & Replace(clpDiv.Text, ",", "','") & "')"
    If clpDept.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DEPTNO IN ('" & Replace(clpDept.Text, ",", "','") & "')"
    If clpCode(0).Text <> "" Then WSQLQ = WSQLQ & " AND ED_LOC IN ('" & Replace(clpCode(0).Text, ",", "','") & "')"
    If clpCode(1).Text <> "" Then WSQLQ = WSQLQ & " AND ED_ORG IN ('" & Replace(clpCode(1).Text, ",", "','") & "')"
    If clpCode(2).Text <> "" Then WSQLQ = WSQLQ & " AND ED_EMP IN ('" & Replace(clpCode(2).Text, ",", "','") & "')"
    If clpCode(3).Text <> "" Then WSQLQ = WSQLQ & " AND ED_REGION IN ('" & Replace(clpCode(3).Text, ",", "','") & "')"
    If clpCode(4).Text <> "" Then WSQLQ = WSQLQ & " AND ED_ADMINBY IN ('" & Replace(clpCode(4).Text, ",", "','") & "')"
    If clpCode(5).Text <> "" Then WSQLQ = WSQLQ & " AND ED_SECTION IN ('" & Replace(clpCode(5).Text, ",", "','") & "')"
    If clpPT.Text <> "" Then WSQLQ = WSQLQ & " AND ED_PT IN ('" & Replace(clpPT.Text, ",", "','") & "')"
       
    If elpEEID.Text <> "" Then WSQLQ = WSQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
        

    'Write to Excel Spreadsheet
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME,ED_PT,ED_DIV,ED_DEPTNO,ED_UNION,ED_PHONE,ED_BUSNBR FROM HREMP "
    SQLQ = SQLQ & "WHERE ED_PT IN ('FT') "
    SQLQ = SQLQ & "AND " & WSQLQ
    SQLQ = SQLQ & "ORDER BY ED_UNION ASC"
    rsHRSen.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    If Not rsHRSen.EOF Then
        totNum = rsHRSen.RecordCount: I = 0
        rsHRSen.MoveFirst
                
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SeniorityList_Tmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SeniorityListFT" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
    
        exSheet.Cells(2, 1) = Format(Now, "mm/dd/yyyy")
        exSheet.Cells(2, 2) = "as of  " & Format(dlpAsOf.Text, "mmmm dd, yyyy")
        exSheet.Cells(1, 2) = "WDDS Seniority Report By Full-Time Seniority Date"
        
        xDeptno = ""
        xRow = 4
        xDeptRow = 0
        Do While Not rsHRSen.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
                    
            'If xDeptNo <> rsHRSen("ED_DEPTNO") Then
            '    If xDeptRow <> 0 Then
            '        xn = exSheet.Range("A" & xDeptRow & ":E" & xRow - 1).Sort(exSheet.Range("B" & xDeptRow & ":B" & xRow - 1), xlDescending) ', , , , , , xlNo, , , xlSortColumns)
            '    End If
            '    xRow = xRow + 1
            '    exSheet.Cells(xRow, 1) = getDeptDesc(rsHRSen("ED_DEPTNO"))
            '    exSheet.Rows(xRow).Font.Bold = True
            '    xRow = xRow + 1
            '    xDeptNo = rsHRSen("ED_DEPTNO")
            '    xDeptRow = xRow
            'End If
            
            'Name
            exSheet.Cells(xRow, 1) = rsHRSen("ED_SURNAME") & ", " & rsHRSen("ED_FNAME")
            
            'If rsHRSen("ED_EMPNBR") = 855 Then
            '    MsgBox ""
            'End If
            
            If rsHRSen("ED_PT") = "FT" Then
                exSheet.Cells(xRow, 2) = IIf(Not IsNull(rsHRSen("ED_UNION")), Format(rsHRSen("ED_UNION"), "mm/dd/yyyy"), "")
            'Else
                'PT
            '    If rsHRSen("ED_PT") = "PT" Then
            '        exSheet.Cells(xRow, 2) = get_Seniority_Hours(rsHRSen("ED_EMPNBR"))
            '    End If
            End If
            
            'Phone #s
            exSheet.Cells(xRow, 3) = Format(rsHRSen("ED_PHONE"), "(###) ###-####")
            exSheet.Cells(xRow, 4) = Format(rsHRSen("ED_BUSNBR"), "(###) ###-####")
            
            'Status
            exSheet.Cells(xRow, 5) = get_Employee_Flag11(rsHRSen("ED_EMPNBR"))
            
            xRow = xRow + 1
            
            rsHRSen.MoveNext
        Loop
        
        exSheet.Range("A" & xRow & ":Q423").Delete
        exSheet.Range("A" & xRow - 1 & ":Q" & xRow - 1).Borders(xlEdgeBottom).Weight = xlThick
        'exSheet.Range("A" & xRow - 1 & ":Q" & xRow - 1).Font.Strikethrough
        
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If

    Screen.MousePointer = DEFAULT

Exit Sub

Export_FT_Seniority_Excel_Err:
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
    Exit Sub
End Sub

Private Sub Export_PT_Seniority_Excel()
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    Dim SQLX, WSQLQ
    Dim rsHRSen As New ADODB.Recordset
    Dim rsHRTabl As New ADODB.Recordset
    Dim xFld As String
    Dim xDeptno As String
    Dim xDeptRow As Integer
    Dim xn

    On Error GoTo Export_FT_Seniority_Excel_Err
    
    Screen.MousePointer = HOURGLASS
    
    'Hemu
    Set CN001 = New ADODB.Connection
    CN001.CommandTimeout = 600
    CN001.Open glbAdoIHRDB
    
    Set CN001W = New ADODB.Connection
    CN001W.CommandTimeout = 600
    CN001W.Open glbAdoIHRDBW
    
    CN001.BeginTrans
    CN001.Execute "DELETE FROM HRSENHRS " & in_SQL(glbIHRDBW) & " WHERE WRKEMP='" & glbUserID & "'"
    CN001.CommitTrans
        
    WSQLQ = glbSeleDeptUn
    If clpDiv.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DIV IN ('" & Replace(clpDiv.Text, ",", "','") & "')"
    If clpDept.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DEPTNO IN ('" & Replace(clpDept.Text, ",", "','") & "')"
    If clpCode(0).Text <> "" Then WSQLQ = WSQLQ & " AND ED_LOC IN ('" & Replace(clpCode(0).Text, ",", "','") & "')"
    If clpCode(1).Text <> "" Then WSQLQ = WSQLQ & " AND ED_ORG IN ('" & Replace(clpCode(1).Text, ",", "','") & "')"
    If clpCode(2).Text <> "" Then WSQLQ = WSQLQ & " AND ED_EMP IN ('" & Replace(clpCode(2).Text, ",", "','") & "')"
    If clpCode(3).Text <> "" Then WSQLQ = WSQLQ & " AND ED_REGION IN ('" & Replace(clpCode(3).Text, ",", "','") & "')"
    If clpCode(4).Text <> "" Then WSQLQ = WSQLQ & " AND ED_ADMINBY IN ('" & Replace(clpCode(4).Text, ",", "','") & "')"
    If clpCode(5).Text <> "" Then WSQLQ = WSQLQ & " AND ED_SECTION IN ('" & Replace(clpCode(5).Text, ",", "','") & "')"
    If clpPT.Text <> "" Then WSQLQ = WSQLQ & " AND ED_PT IN ('" & Replace(clpPT.Text, ",", "','") & "')"

       
    If elpEEID.Text <> "" Then WSQLQ = WSQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
        
        
    SQLQ = "SELECT '001' AS COMPANY,ED_EMPNBR,'EDOR' AS UNION_TBL,'EDEM' AS EMPLSTAT_TBL,"
    If glbOracle Then
        SQLQ = SQLQ & "SUBSTR(ED_SURNAME || ', ' || ED_FNAME,1,39) AS NAME,"
    Else
        SQLQ = SQLQ & "LEFT(ED_SURNAME +', '+ED_FNAME,39) AS NAME,"
    End If
    SQLQ = SQLQ & "ED_DIV,ED_DEPTNO,ED_ORG,ED_EMP,ED_DOH,0 AS NUMBHRS "
    SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
    SQLQ = SQLQ & "FROM HREMP"
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    
    SQLX = "INSERT INTO HRSENHRS "
    SQLX = SQLX & "(COMPANY,EMPLNUMB,UNION_TBL,EMPLSTAT_TBL,NAME,"
    SQLX = SQLX & "DIVISION,DEPTNO,UNIONS,EMPLSTAT,DOH,NUMBHRS,WRKEMP) "
    SQLX = SQLX & SQLQ & " WHERE " & WSQLQ
    
    
    CN001W.BeginTrans
    CN001W.Execute SQLX
    CN001W.CommitTrans
    
    'Hemu
    
    SQLQ = "SELECT AD_EMPNBR AS EMPNBR, SUM(AD_HRS) AS NUM1"
    SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
    SQLQ = SQLQ & " FROM HR_ATTENDANCE "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    
    SQLQ = SQLQ & " WHERE AD_SEN<>0 AND AD_DOA<=" & Date_SQL(dlpAsOf.Text)
    
    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " WHERE " & WSQLQ & ")"
    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
    If glbSQL Then
        SQLX = " Update HRSENHRS "
        SQLX = SQLX & " SET NUMBHRS = TRSNUM.NUM1"
        SQLX = SQLX & " FROM HRSENHRS INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRSENHRS.EMPLNUMB"
        gdbAdoIhr001.Execute SQLX
    Else
        'Hemu
        CN001W.BeginTrans
        CN001W.Execute "DELETE FROM TRSNUM WHERE WRKEMP='" & glbUserID & "'"
        CN001W.CommitTrans
        'Hemu
                
        SQLX = "INSERT INTO TRSNUM (EMPNBR,NUM1,WRKEMP) " & SQLQ
        'Hemu
        CN001W.BeginTrans
        CN001W.Execute SQLX
        CN001W.CommitTrans
        'Hemu
                
        SQLX = "Update HRSENHRS INNER JOIN TRSNUM ON TRSNUM.EMPNBR=HRSENHRS.EMPLNUMB"
        SQLX = SQLX & " SET NUMBHRS = TRSNUM.NUM1 "
        SQLX = SQLX & "WHERE TRSNUM.WRKEMP='" & glbUserID & "' "
        SQLX = SQLX & "AND HRSENHRS.WRKEMP='" & glbUserID & "' "
        'Hemu
        CN001W.BeginTrans
        CN001W.Execute SQLX
        CN001W.CommitTrans
        'Hemu
    End If
    
    If chkInclAtt Then
        SQLQ = "SELECT AH_EMPNBR AS EMPNBR, SUM(AH_HRS) AS NUM1"
        SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
        SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
        SQLQ = SQLQ & in_SQL(glbIHRDB)
        SQLQ = SQLQ & " WHERE AH_SEN<>0 AND AH_DOA<=" & Date_SQL(dlpAsOf.Text)
        
        SQLQ = SQLQ & " AND AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP "
        SQLQ = SQLQ & in_SQL(glbIHRDB)
        SQLQ = SQLQ & " WHERE " & WSQLQ & ")"
        SQLQ = SQLQ & " GROUP BY AH_EMPNBR"

        If glbSQL Then
            SQLX = " Update HRSENHRS "
            SQLX = SQLX & " SET NUMBHRS = ISNULL(NUMBHRS,0)+TRSNUM.NUM1"
            SQLX = SQLX & " FROM HRSENHRS INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRSENHRS.EMPLNUMB"
            gdbAdoIhr001.Execute SQLX
        Else
            'Hemu
            CN001W.BeginTrans
            CN001W.Execute "DELETE FROM TRSNUM WHERE WRKEMP='" & glbUserID & "'"
            CN001W.CommitTrans
            'Hemu
            
            SQLX = "INSERT INTO TRSNUM (EMPNBR,NUM1,WRKEMP) " & SQLQ
            
            'Hemu
            CN001W.BeginTrans
            CN001W.Execute SQLX
            CN001W.CommitTrans
            'Hemu
            
            SQLX = "Update HRSENHRS INNER JOIN TRSNUM ON TRSNUM.EMPNBR=HRSENHRS.EMPLNUMB"
            SQLX = SQLX & " SET NUMBHRS = NUMBHRS + TRSNUM.NUM1 "
            SQLX = SQLX & "WHERE TRSNUM.WRKEMP='" & glbUserID & "' "
            SQLX = SQLX & "AND HRSENHRS.WRKEMP='" & glbUserID & "' "
            
            'Hemu
            'gdbAdoIhr001W.Execute SQLX
            CN001W.BeginTrans
            CN001W.Execute SQLX
            CN001W.CommitTrans
            'Hemu
        End If
    End If
    
    CN001.Close
    CN001W.Close
    
    Set CN001 = Nothing
    Set CN001W = Nothing


    'Write to Excel Spreadsheet
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME,ED_PT,ED_DIV,ED_DEPTNO,ED_UNION,ED_PHONE,ED_BUSNBR FROM HREMP "
    SQLQ = SQLQ & "WHERE ED_PT IN ('PT') "
    SQLQ = SQLQ & "AND " & WSQLQ
    SQLQ = SQLQ & "ORDER BY ED_DOH DESC"
    rsHRSen.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    If Not rsHRSen.EOF Then
        totNum = rsHRSen.RecordCount: I = 0
        rsHRSen.MoveFirst
                
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SeniorityList_Tmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SeniorityListPT" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
    
        exSheet.Cells(2, 1) = Format(Now, "mm/dd/yyyy")
        exSheet.Cells(2, 2) = "as of  " & Format(dlpAsOf.Text, "mmmm dd, yyyy")
        exSheet.Cells(1, 2) = "WDDS Seniority Report By Part-Time Seniority Hours"
        
        xDeptno = ""
        xRow = 4
        xDeptRow = 0
        Do While Not rsHRSen.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
                    
            'If xDeptNo <> rsHRSen("ED_DEPTNO") Then
            '    If xDeptRow <> 0 Then
            '        xn = exSheet.Range("A" & xDeptRow & ":E" & xRow - 1).Sort(exSheet.Range("B" & xDeptRow & ":B" & xRow - 1), xlDescending) ', , , , , , xlNo, , , xlSortColumns)
            '    End If
            '    xRow = xRow + 1
            '    exSheet.Cells(xRow, 1) = getDeptDesc(rsHRSen("ED_DEPTNO"))
            '    exSheet.Rows(xRow).Font.Bold = True
            '    xRow = xRow + 1
            '    xDeptNo = rsHRSen("ED_DEPTNO")
            '    xDeptRow = xRow
            'End If
            
            'Name
            exSheet.Cells(xRow, 1) = rsHRSen("ED_SURNAME") & ", " & rsHRSen("ED_FNAME")
            
            'If rsHRSen("ED_EMPNBR") = 855 Then
            '    MsgBox ""
            'End If
            
            'If rsHRSen("ED_PT") = "FT" Then
            '    exSheet.Cells(xRow, 2) = IIf(Not IsNull(rsHRSen("ED_UNION")), Format(rsHRSen("ED_UNION"), "mm/dd/yyyy"), "")
            'Else
            If rsHRSen("ED_PT") = "PT" Then
                'PT
                If rsHRSen("ED_PT") = "PT" Then
                    exSheet.Cells(xRow, 2) = get_Seniority_Hours(rsHRSen("ED_EMPNBR"))
                End If
            End If
            
            'Phone #s
            exSheet.Cells(xRow, 3) = Format(rsHRSen("ED_PHONE"), "(###) ###-####")
            exSheet.Cells(xRow, 4) = Format(rsHRSen("ED_BUSNBR"), "(###) ###-####")
            
            'Status
            exSheet.Cells(xRow, 5) = get_Employee_Flag11(rsHRSen("ED_EMPNBR"))
            
            xRow = xRow + 1
            
            rsHRSen.MoveNext
        Loop
        
        xn = exSheet.Range("A4:E" & xRow - 1).Sort(exSheet.Range("B4:B" & xRow - 1), xlDescending) ', , , , , , xlNo, , , xlSortColumns)
        
        exSheet.Range("A" & xRow & ":Q423").Delete
        exSheet.Range("A" & xRow - 1 & ":Q" & xRow - 1).Borders(xlEdgeBottom).Weight = xlThick
        
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If

    Screen.MousePointer = DEFAULT

Exit Sub

Export_FT_Seniority_Excel_Err:
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
    Exit Sub
End Sub

Private Function GetEmpCountBySele(xDivCode, xRepSel, Optional xSalHly) 'Ticket #19137
Dim rsLocEmp As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = 0
    'SQLQ = "SELECT COUNT(ED_EMPNBR) AS EMP_COUNT FROM HREMP WHERE ED_DIV = '" & xDivCode & "' "
    SQLQ = "SELECT COUNT(ED_EMPNBR) AS EMP_COUNT FROM HREMP WHERE (1=1) "
    SQLQ = SQLQ & "AND ED_DIV = '" & xDivCode & "' "
    SQLQ = SQLQ & "AND " & xRepSel & " " 'Ticket #20571 07/05/2011
    If Not IsMissing(xSalHly) Then
        If xSalHly = "Salaried" Then
            SQLQ = SQLQ & "AND (ED_ORG = 'NONE' OR ED_ORG = 'EXEC') "
        End If
        If xSalHly = "Hourly" Then
            SQLQ = SQLQ & "AND NOT (ED_ORG = 'NONE' OR ED_ORG = 'EXEC') "
        End If
    End If
    rsLocEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsLocEmp.EOF Then
        xRetVal = rsLocEmp("EMP_COUNT")
    End If
    rsLocEmp.Close
    GetEmpCountBySele = xRetVal
End Function

Private Function GetColorByPerc(xVal) 'Ticket #19137
Dim retval
    'If xVal < 0.75 Then
    If xVal < 75 Then
        retval = "Red"
    Else
        retval = "Green"
    End If
    GetColorByPerc = retval
'    'If xVal < 0.75 Then
'    If xVal < 75 Then
'        retVal = RED
'    Else
'        retVal = RGB(34, 139, 34) 'Dark Green
'    End If
'    GetColorByPerc = retVal
End Function

Private Function ChapmansExcelRpt_AverageHourRpt()
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim I, totNum
    Dim xRow As Integer

    Dim X, Y As Integer
    Dim SQLQ, sSQLQ As String
    Dim rsAttMatrix As New ADODB.Recordset
    Dim rsHREmp As New ADODB.Recordset
    Dim rsAttend As New ADODB.Recordset
    Dim rsHRParco As New ADODB.Recordset
    Dim xRegHrsCode As String
    Dim xVacHrsCode As String
    Dim xAbsentHrsCode As String
    Dim xExtraHrsCode As String
    Dim dtFromDate As Date
    
    'Monthly hours
    Dim RegHrsMonthly(13) As Double
    Dim VacWksMonthly(13) As Double
    Dim WrkWksMonthly(13) As Double
    Dim AbsWksMonthly(13) As Double
    Dim ExtHrsMonthly(13) As Double
    Dim TotHrsMonthly(13) As Double
    Dim AvgHrsMonthly(13) As Double
        
    'Month Columns - to indicate what month is in which column in Excel
    Dim MonthCol(14, 9999) As Integer
    Dim col As Integer
    Dim totmntcol As Integer
        
    On Error GoTo ChapmansExcelRpt_AverageHourRpt_Err
    
    'Selection criteria
    'sSQLQ = Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")")
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 1")
    
    'Compute the 1 year using the End Date
    dlpDateRange(0).Text = DateAdd("yyyy", -1, dlpAsOf.Text) + 1
    dlpDateRange(1).Text = dlpAsOf.Text
    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 2")
    
    'Get Attendance Codes for various sum of hours calculations
    'Get Regular Hours codes
    xRegHrsCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_REG_HRS = 1"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xRegHrsCode) = 0 Then
            xRegHrsCode = xRegHrsCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xRegHrsCode = xRegHrsCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 3")
    
    'Get Vacation Hours Code
    xVacHrsCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_VAC_HRS = 1"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xVacHrsCode) = 0 Then
            xVacHrsCode = xVacHrsCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xVacHrsCode = xVacHrsCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 4")

    'Get Absent Hours Code
    xAbsentHrsCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_ABSENT_HRS = 1"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xAbsentHrsCode) = 0 Then
            xAbsentHrsCode = xAbsentHrsCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xAbsentHrsCode = xAbsentHrsCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 5")

    'Get Extra Hours Code
    xExtraHrsCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_EXTRA_HRS = 1"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xExtraHrsCode) = 0 Then
            xExtraHrsCode = xExtraHrsCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xExtraHrsCode = xExtraHrsCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
        
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 6")
        
    'Initialise/Open Excel Report file
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "AverageHourTmp.xls"
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "AverageHour_" & Trim(glbUserID) & ".xls"

Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 7")

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 8")
        Exit Function
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 9")

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0

    FileCopy xlsFileTmp, xlsFileMat

Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 10")

    Screen.MousePointer = HOURGLASS

    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 10a")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 10b")
    Set exSheet = exBook.Worksheets(1)

Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 11")

    SQLQ = "SELECT PC_NAME FROM HRPARCO"
    rsHRParco.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsHRParco.EOF Then
        exSheet.Cells(1, 1) = rsHRParco("PC_NAME")
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 11a")
    Else
        exSheet.Cells(1, 1) = "CHAPMAN'S ICE CREAM LIMITED"
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 11b")
    End If
    rsHRParco.Close
    Set rsHRParco = Nothing

Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 12")

    'Print the report headers
    exSheet.Cells(2, 1) = Format(Now, "mmm dd, yyyy hh:mm")
    exSheet.Cells(4, 1) = "AVERAGE HOUR REPORT"
    exSheet.Cells(5, 1) = "Report for the Period: " & Format(dlpDateRange(0).Text, "mmmm dd, yyyy") & " To " & Format(dlpDateRange(1).Text, "mmmm dd, yyyy")
                      
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 13")

    'For each employee get the breakdown of hours
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME, JH_DHRS, JH_WHRS "
    SQLQ = SQLQ & " FROM (HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR AND JH_CURRENT <>0)"
    If Len(sSQLQ) > 0 Then
        SQLQ = SQLQ & " WHERE " & sSQLQ
    End If
    SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME"
    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 14: " & SQLQ)

    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: I = 0
        rsHREmp.MoveFirst

        xRow = 7
        Do While Not rsHREmp.EOF
        
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 15")

            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 16: Row:" & xRow)

            'Print the column headings
            exSheet.Cells(xRow, 1) = "NAME"
            exSheet.Cells(xRow, 1).Interior.ColorIndex = 15
                        
            'Print Month Names
            totmntcol = 0
            For X = 1 To 12 '13 '12
                If X <> 1 Then
                    exSheet.Cells(xRow, X + 1) = UCase(Format(DateAdd("m", X - 1, dlpDateRange(0).Text), "MMM"))
                    
                    'Set the Column # in the Months variable so that right hours are updated in the right month col. later
                    MonthCol(month(DateAdd("m", X - 1, dlpDateRange(0).Text)), Year(DateAdd("m", X - 1, dlpDateRange(0).Text))) = X + 1
                Else
                    exSheet.Cells(xRow, X + 1) = UCase(Format(dlpDateRange(0).Text, "MMM"))
                    
                    'Set the Column # in the Months variable so that right hours are updated in the right month col. later
                    MonthCol(month(dlpDateRange(0).Text), Year(dlpDateRange(0).Text)) = X + 1
                End If
                
                'Shading alternate cells
                If (X + 1) / 2 = Int((X + 1) / 2) Then
                    exSheet.Cells(xRow, X + 1).Interior.ColorIndex = 15
                End If
                
                totmntcol = totmntcol + 1
                If X <> 1 And CVDate(DateAdd("m", X - 1, dlpDateRange(0).Text)) > CVDate(dlpDateRange(1).Text) Then
                    Exit For
                End If
                
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 16: " & X)

            Next

Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 17")

            exSheet.Cells(xRow, totmntcol + 2) = "TOTAL"    '14
            exSheet.Cells(xRow, totmntcol + 2).Interior.ColorIndex = 15     '14
            
            exSheet.Rows(xRow).Font.Bold = True
            exSheet.Rows(xRow).HorizontalAlignment = xlCenter
            xRow = xRow + 1
            
            'Print Employee Name
            exSheet.Cells(xRow, 1) = UCase(rsHREmp("ED_SURNAME")) & ", " & rsHREmp("ED_FNAME")
            
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 18")

            'Initialise Monthly variables
            For X = 0 To 13
                RegHrsMonthly(X) = 0
                VacWksMonthly(X) = 0
                WrkWksMonthly(X) = 0
                AbsWksMonthly(X) = 0
                ExtHrsMonthly(X) = 0
                TotHrsMonthly(X) = 0
                AvgHrsMonthly(X) = 0
            Next
            
            col = 0
            'Compute monthly Regular Hours
            If Len(xRegHrsCode) > 0 Then
                'Attendance
                SQLQ = "SELECT SUM(AD_HRS) AS REG_HRS, MONTH(AD_DOA) AS MONTHS, YEAR(AD_DOA) AS YEARS FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xRegHrsCode & ")"
                SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1).Text)
                SQLQ = SQLQ & " GROUP BY MONTH(AD_DOA), YEAR(AD_DOA)"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    'RegHrsMonthly(rsAttend("MONTHS")) = IIf(Not IsNull(rsAttend("REG_HRS")), rsAttend("REG_HRS"), 0)
                    col = col + 1
                    
                    'RegHrsMonthly(col) = IIf(Not IsNull(rsAttend("REG_HRS")), rsAttend("REG_HRS"), 0)
                    RegHrsMonthly(rsAttend("MONTHS")) = IIf(Not IsNull(rsAttend("REG_HRS")), rsAttend("REG_HRS"), 0)
                    
                    'Print Regular Hours
                    'exSheet.Cells(xRow, rsAttend("MONTHS") + 1) = RegHrsMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"))) = RegHrsMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(col)) = RegHrsMonthly(col)
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = RegHrsMonthly(col)
                    exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = RegHrsMonthly(rsAttend("MONTHS"))

Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 18a")
                    
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 19")

                col = 0
                'Attendance History
                SQLQ = "SELECT SUM(AH_HRS) AS REG_HRS, MONTH(AH_DOA) AS MONTHS, YEAR(AH_DOA) AS YEARS FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xRegHrsCode & ")"
                SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(1).Text)
                SQLQ = SQLQ & " GROUP BY MONTH(AH_DOA), YEAR(AH_DOA)"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    'RegHrsMonthly(rsAttend("MONTHS")) = RegHrsMonthly(rsAttend("MONTHS")) + IIf(Not IsNull(rsAttend("REG_HRS")), rsAttend("REG_HRS"), 0)
                    col = col + 1
                    
                    'RegHrsMonthly(col) = RegHrsMonthly(col) + IIf(Not IsNull(rsAttend("REG_HRS")), rsAttend("REG_HRS"), 0)
                    RegHrsMonthly(rsAttend("MONTHS")) = RegHrsMonthly(rsAttend("MONTHS")) + IIf(Not IsNull(rsAttend("REG_HRS")), rsAttend("REG_HRS"), 0)
                    
                    'Print Regular Hours
                    'exSheet.Cells(xRow, rsAttend("MONTHS") + 1) = RegHrsMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"))) = RegHrsMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(col)) = RegHrsMonthly(col)
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = RegHrsMonthly(col)
                    exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = RegHrsMonthly(rsAttend("MONTHS"))
                    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 20")

                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 21")
            End If
            
            
            'Print Vacation heading
            col = 0
            xRow = xRow + 1
            exSheet.Cells(xRow, 1) = "Vacation"
            
            'Compute Vacation Weeks
            If Len(xVacHrsCode) > 0 Then
                'Attendance
                'Ticket #20994 - # of Vacation Attendance records instead of sum of Vacation hours
                '# of Vacation Days / 5
                'SQLQ = "SELECT SUM(AD_HRS) AS VAC_HRS, MONTH(AD_DOA) AS MONTHS, YEAR(AD_DOA) AS YEARS FROM HR_ATTENDANCE "
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS VAC_HRS, MONTH(AD_DOA) AS MONTHS, YEAR(AD_DOA) AS YEARS FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xVacHrsCode & ")"
                SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1).Text)
                SQLQ = SQLQ & " GROUP BY MONTH(AD_DOA), YEAR(AD_DOA)"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    'VacWksMonthly(rsAttend("MONTHS")) = (IIf(Not IsNull(rsAttend("VAC_HRS")), rsAttend("VAC_HRS"), 0) / rsHREmp("JH_DHRS")) / 5
                    col = col + 1
                    
                    'Ticket #20994 - Divide by 5 - since we are getting # of days taken for vacation, we need to
                    'convert that to # of Vacation weeks.
                    'VacWksMonthly(col) = (IIf(Not IsNull(rsAttend("VAC_HRS")), rsAttend("VAC_HRS"), 0) / rsHREmp("JH_DHRS")) / 5
                    'VacWksMonthly(col) = IIf(Not IsNull(rsAttend("VAC_HRS")), rsAttend("VAC_HRS"), 0) / 5
                    VacWksMonthly(rsAttend("MONTHS")) = IIf(Not IsNull(rsAttend("VAC_HRS")), rsAttend("VAC_HRS"), 0) / 5
                    
                    'Print Vacation weeks
                    'exSheet.Cells(xRow, rsAttend("MONTHS") + 1) = VacWksMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"))) = VacWksMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(col)) = VacWksMonthly(col)
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = VacWksMonthly(col)
                    exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = VacWksMonthly(rsAttend("MONTHS"))

Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 22")

                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                col = 0
                'Ticket #20994 - # of Vacation Attendance records instead of sum of Vacation hours
                '# of Vacation Days / 5
                'SQLQ = "SELECT SUM(AH_HRS) AS VAC_HRS, MONTH(AH_DOA) AS MONTHS, YEAR(AH_DOA) AS YEARS FROM HR_ATTENDANCE_HISTORY "
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS VAC_HRS, MONTH(AH_DOA) AS MONTHS, YEAR(AH_DOA) AS YEARS FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xVacHrsCode & ")"
                SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(1).Text)
                SQLQ = SQLQ & " GROUP BY MONTH(AH_DOA), YEAR(AH_DOA)"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    'VacWksMonthly(rsAttend("MONTHS")) = VacWksMonthly(rsAttend("MONTHS")) + (IIf(Not IsNull(rsAttend("VAC_HRS")), rsAttend("VAC_HRS"), 0) / rsHREmp("JH_DHRS")) / 5
                    col = col + 1
                    
                    'Ticket #20994 - Divide by 5 - since we are getting # of days taken for vacation, we need to
                    'convert that to # of Vacation weeks.
                    'VacWksMonthly(col) = VacWksMonthly(col) + (IIf(Not IsNull(rsAttend("VAC_HRS")), rsAttend("VAC_HRS"), 0) / rsHREmp("JH_DHRS")) / 5
                    'VacWksMonthly(col) = VacWksMonthly(col) + (IIf(Not IsNull(rsAttend("VAC_HRS")), rsAttend("VAC_HRS"), 0) / 5)
                    VacWksMonthly(rsAttend("MONTHS")) = VacWksMonthly(rsAttend("MONTHS")) + (IIf(Not IsNull(rsAttend("VAC_HRS")), rsAttend("VAC_HRS"), 0) / 5)
                    
                    'Print Vacation weeks
                    'exSheet.Cells(xRow, rsAttend("MONTHS") + 1) = VacWksMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"))) = VacWksMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(col)) = VacWksMonthly(col)
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = VacWksMonthly(col)
                    exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = VacWksMonthly(rsAttend("MONTHS"))
                    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 23")

                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
            End If
            
                    
            'Print Weeks Worked heading
            col = 0
            xRow = xRow + 1
            exSheet.Cells(xRow, 1) = "Weeks Worked"
                    
            'Compute Weeks Worked
            If Len(xAbsentHrsCode) > 0 Then
                'Attendance
                'Ticket #20994 - # of Absent Attendance records instead of sum of Absent hours
                '(Actual Work Days of the month {excl. w/ends} - Absent Days - Vacation Days) / 5
                'SQLQ = "SELECT SUM(AD_HRS) AS ABS_HRS, MONTH(AD_DOA) AS MONTHS, YEAR(AD_DOA) AS YEARS FROM HR_ATTENDANCE "
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS ABS_HRS, MONTH(AD_DOA) AS MONTHS, YEAR(AD_DOA) AS YEARS FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xAbsentHrsCode & ")"
                SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1).Text)
                SQLQ = SQLQ & " GROUP BY MONTH(AD_DOA), YEAR(AD_DOA)"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    'WrkWksMonthly(rsAttend("MONTHS")) = Round(Day(MonthLastDate(Format(rsAttend("MONTHS") & "/01/" & rsAttend("YEARS"), "mm/dd/yyyy"))) / 7, 2) - IIf(Not IsNull(rsAttend("ABS_HRS")), Round(rsAttend("ABS_HRS") / rsHREmp("JH_WHRS"), 2), 0)
                    col = col + 1
                    
                    'Ticket #20994 - Get # of Work Days in a Month minus # of Absent Days minus (# of Vacation Weeks & 5) divide by 5
                    'WrkWksMonthly(col) = Round(Day(MonthLastDate(Format(rsAttend("MONTHS") & "/01/" & rsAttend("YEARS"), "mm/dd/yyyy"))) / 7, 2) - IIf(Not IsNull(rsAttend("ABS_HRS")), Round(rsAttend("ABS_HRS") / rsHREmp("JH_WHRS"), 2), 0)
                    'WrkWksMonthly(col) = (Weekdays(Format(rsAttend("MONTHS") & "/01/" & rsAttend("YEARS"), "mm/dd/yyyy"), Format(MonthLastDate(Format(rsAttend("MONTHS") & "/01/" & rsAttend("YEARS"), "mm/dd/yyyy")), "mm/dd/yyyy")) - IIf(Not IsNull(rsAttend("ABS_HRS")), rsAttend("ABS_HRS"), 0) - (VacWksMonthly(col) * 5)) / 5
                    WrkWksMonthly(rsAttend("MONTHS")) = (Weekdays(Format(rsAttend("MONTHS") & "/01/" & rsAttend("YEARS"), "mm/dd/yyyy"), Format(MonthLastDate(Format(rsAttend("MONTHS") & "/01/" & rsAttend("YEARS"), "mm/dd/yyyy")), "mm/dd/yyyy")) - IIf(Not IsNull(rsAttend("ABS_HRS")), rsAttend("ABS_HRS"), 0) - (VacWksMonthly(rsAttend("MONTHS")) * 5)) / 5
                    
                    'Print Weeks Worked
                    'exSheet.Cells(xRow, rsAttend("MONTHS") + 1) = WrkWksMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"))) = WrkWksMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(col)) = WrkWksMonthly(col)
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = WrkWksMonthly(col)
                    exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = WrkWksMonthly(rsAttend("MONTHS"))
                    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 24")

                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                col = 0
                'Ticket #20994 - # of Absent Attendance records instead of sum of Absent hours
                '(Actual Work Days {excl. w/ends} - Absent Days - Vacation Days) / 5
                'SQLQ = "SELECT SUM(AH_HRS) AS ABS_HRS, MONTH(AH_DOA) AS MONTHS, YEAR(AH_DOA) AS YEARS FROM HR_ATTENDANCE_HISTORY "
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS ABS_HRS, MONTH(AH_DOA) AS MONTHS, YEAR(AH_DOA) AS YEARS FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xAbsentHrsCode & ")"
                SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(1).Text)
                SQLQ = SQLQ & " GROUP BY MONTH(AH_DOA), YEAR(AH_DOA)"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    'WrkWksMonthly(rsAttend("MONTHS")) = WrkWksMonthly(rsAttend("MONTHS")) + Round(Day(MonthLastDate(Format(rsAttend("MONTHS") & "/01/" & rsAttend("YEARS"), "mm/dd/yyyy"))) / 7, 2) - IIf(Not IsNull(rsAttend("ABS_HRS")), rsAttend("ABS_HRS") / Round(rsHREmp("JH_WHRS"), 2), 0)
                    col = col + 1
                    
                    'Ticket #20994 - Get # of Work Days in a Month minus # of Absent Days minus (# of Vacation Weeks & 5) divide by 5
                    'WrkWksMonthly(col) = WrkWksMonthly(col) + Round(Day(MonthLastDate(Format(rsAttend("MONTHS") & "/01/" & rsAttend("YEARS"), "mm/dd/yyyy"))) / 7, 2) - IIf(Not IsNull(rsAttend("ABS_HRS")), rsAttend("ABS_HRS") / Round(rsHREmp("JH_WHRS"), 2), 0)
                    'WrkWksMonthly(col) = WrkWksMonthly(col) + ((Weekdays(Format(rsAttend("MONTHS") & "/01/" & rsAttend("YEARS"), "mm/dd/yyyy"), Format(MonthLastDate(Format(rsAttend("MONTHS") & "/01/" & rsAttend("YEARS"), "mm/dd/yyyy")), "mm/dd/yyyy")) - IIf(Not IsNull(rsAttend("ABS_HRS")), rsAttend("ABS_HRS"), 0) - (VacWksMonthly(col) * 5)) / 5)
                    WrkWksMonthly(rsAttend("MONTHS")) = WrkWksMonthly(rsAttend("MONTHS")) + ((Weekdays(Format(rsAttend("MONTHS") & "/01/" & rsAttend("YEARS"), "mm/dd/yyyy"), Format(MonthLastDate(Format(rsAttend("MONTHS") & "/01/" & rsAttend("YEARS"), "mm/dd/yyyy")), "mm/dd/yyyy")) - IIf(Not IsNull(rsAttend("ABS_HRS")), rsAttend("ABS_HRS"), 0) - (VacWksMonthly(rsAttend("MONTHS")) * 5)) / 5)
                    
                    'Print Weeks Worked
                    'exSheet.Cells(xRow, rsAttend("MONTHS") + 1) = WrkWksMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"))) = WrkWksMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(col)) = WrkWksMonthly(col)
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = WrkWksMonthly(col)
                    exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = WrkWksMonthly(rsAttend("MONTHS"))
                    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 25")

                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
            End If
            
            
            'Print Weeks Absent heading
            col = 0
            xRow = xRow + 1
            exSheet.Cells(xRow, 1) = "Weeks Absent"
            
            'Compute Weeks Absent
            If Len(xAbsentHrsCode) > 0 Then
                'Attendance
                'Ticket #20994 - # of Absent Attendance records instead of sum of Absent hours
                '# of Absent Days / 5
                'SQLQ = "SELECT SUM(AD_HRS) AS ABS_HRS, MONTH(AD_DOA) AS MONTHS, YEAR(AD_DOA) AS YEARS FROM HR_ATTENDANCE "
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS ABS_HRS, MONTH(AD_DOA) AS MONTHS, YEAR(AD_DOA) AS YEARS FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xAbsentHrsCode & ")"
                SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1).Text)
                SQLQ = SQLQ & " GROUP BY MONTH(AD_DOA), YEAR(AD_DOA)"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    'AbsWksMonthly(rsAttend("MONTHS")) = (IIf(Not IsNull(rsAttend("ABS_HRS")), rsAttend("ABS_HRS"), 0) / rsHREmp("JH_DHRS")) / 5
                    col = col + 1
                    
                    'Ticket #20994 - Divide by 5 - since we are getting # of days absent, we need to
                    'convert that to # of Weeks Absent.
                    'AbsWksMonthly(col) = (IIf(Not IsNull(rsAttend("ABS_HRS")), rsAttend("ABS_HRS"), 0) / rsHREmp("JH_DHRS")) / 5
                    'AbsWksMonthly(col) = IIf(Not IsNull(rsAttend("ABS_HRS")), rsAttend("ABS_HRS"), 0) / 5
                    AbsWksMonthly(rsAttend("MONTHS")) = IIf(Not IsNull(rsAttend("ABS_HRS")), rsAttend("ABS_HRS"), 0) / 5
                    
                    'Print Weeks Absent
                    'exSheet.Cells(xRow, rsAttend("MONTHS") + 1) = AbsWksMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"))) = AbsWksMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(col)) = AbsWksMonthly(col)
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = AbsWksMonthly(col)
                    exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = AbsWksMonthly(rsAttend("MONTHS"))
                    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 26")

                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                col = 0
                'Ticket #20994 - # of Absent Attendance records instead of sum of Absent hours
                '# of Absent Days / 5
                'SQLQ = "SELECT SUM(AH_HRS) AS ABS_HRS, MONTH(AH_DOA) AS MONTHS, YEAR(AH_DOA) AS YEARS FROM HR_ATTENDANCE_HISTORY "
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS ABS_HRS, MONTH(AH_DOA) AS MONTHS, YEAR(AH_DOA) AS YEARS FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xAbsentHrsCode & ")"
                SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(1).Text)
                SQLQ = SQLQ & " GROUP BY MONTH(AH_DOA), YEAR(AH_DOA)"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    'AbsWksMonthly(rsAttend("MONTHS")) = AbsWksMonthly(rsAttend("MONTHS")) + (IIf(Not IsNull(rsAttend("ABS_HRS")), rsAttend("ABS_HRS"), 0) / rsHREmp("JH_DHRS")) / 5
                    col = col + 1
                    
                    'Ticket #20994 - Divide by 5 - since we are getting # of days absent, we need to
                    'convert that to # of Weeks Absent.
                    'AbsWksMonthly(col) = AbsWksMonthly(col) + (IIf(Not IsNull(rsAttend("ABS_HRS")), rsAttend("ABS_HRS"), 0) / rsHREmp("JH_DHRS")) / 5
                    'AbsWksMonthly(col) = AbsWksMonthly(col) + (IIf(Not IsNull(rsAttend("ABS_HRS")), rsAttend("ABS_HRS"), 0) / 5)
                    AbsWksMonthly(rsAttend("MONTHS")) = AbsWksMonthly(rsAttend("MONTHS")) + (IIf(Not IsNull(rsAttend("ABS_HRS")), rsAttend("ABS_HRS"), 0) / 5)
                    
                    'Print Weeks Absent
                    'exSheet.Cells(xRow, rsAttend("MONTHS") + 1) = AbsWksMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"))) = AbsWksMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(col)) = AbsWksMonthly(col)
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = AbsWksMonthly(col)
                    exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = AbsWksMonthly(rsAttend("MONTHS"))
                    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 27")

                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
            End If
            
            
            'Print Extra Hours Paid heading
            col = 0
            xRow = xRow + 1
            exSheet.Cells(xRow, 1) = "Extra Hours Paid"
            
            'Compute Extra Hours Paid
            If Len(xExtraHrsCode) > 0 Then
                'Attendance
                SQLQ = "SELECT SUM(AD_HRS) AS EXTRA_HRS, MONTH(AD_DOA) AS MONTHS, YEAR(AD_DOA) AS YEARS FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xExtraHrsCode & ")"
                SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1).Text)
                SQLQ = SQLQ & " GROUP BY MONTH(AD_DOA), YEAR(AD_DOA)"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    'ExtHrsMonthly(rsAttend("MONTHS")) = IIf(Not IsNull(rsAttend("EXTRA_HRS")), rsAttend("EXTRA_HRS"), 0)
                    col = col + 1
                    
                    'ExtHrsMonthly(col) = IIf(Not IsNull(rsAttend("EXTRA_HRS")), rsAttend("EXTRA_HRS"), 0)
                    ExtHrsMonthly(rsAttend("MONTHS")) = IIf(Not IsNull(rsAttend("EXTRA_HRS")), rsAttend("EXTRA_HRS"), 0)
                    
                    'Print Extra Hours Paid
                    'exSheet.Cells(xRow, rsAttend("MONTHS") + 1) = ExtHrsMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"))) = ExtHrsMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(col)) = ExtHrsMonthly(col)
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = ExtHrsMonthly(col)
                    exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = ExtHrsMonthly(rsAttend("MONTHS"))
                    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 28")

                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                col = 0
                SQLQ = "SELECT SUM(AH_HRS) AS EXTRA_HRS, MONTH(AH_DOA) AS MONTHS, YEAR(AH_DOA) AS YEARS FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xExtraHrsCode & ")"
                SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(1).Text)
                SQLQ = SQLQ & " GROUP BY MONTH(AH_DOA), YEAR(AH_DOA)"
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    'ExtHrsMonthly(rsAttend("MONTHS")) = ExtHrsMonthly(rsAttend("MONTHS")) + IIf(Not IsNull(rsAttend("EXTRA_HRS")), rsAttend("EXTRA_HRS"), 0)
                    col = col + 1
                    
                    'ExtHrsMonthly(col) = ExtHrsMonthly(col) + IIf(Not IsNull(rsAttend("EXTRA_HRS")), rsAttend("EXTRA_HRS"), 0)
                    ExtHrsMonthly(rsAttend("MONTHS")) = ExtHrsMonthly(rsAttend("MONTHS")) + IIf(Not IsNull(rsAttend("EXTRA_HRS")), rsAttend("EXTRA_HRS"), 0)
                    
                    'Print Extra Hours Paid
                    'exSheet.Cells(xRow, rsAttend("MONTHS") + 1) = ExtHrsMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"))) = ExtHrsMonthly(rsAttend("MONTHS"))
                    'exSheet.Cells(xRow, MonthCol(col)) = ExtHrsMonthly(col)
                    'exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = ExtHrsMonthly(col)
                    exSheet.Cells(xRow, MonthCol(rsAttend("MONTHS"), rsAttend("YEARS"))) = ExtHrsMonthly(rsAttend("MONTHS"))
                    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 29")

                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
            End If
            
            
            'Print Tot. Hrs - Extra Hours/#Wks heading
            col = 0
            xRow = xRow + 1
            exSheet.Cells(xRow, 1) = "Tot. Hrs-Extra/#Wks Worked"
            
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 30")

            'Compute & Print Total Hours - Extra Hours / Weeks Worked
            For X = 2 To totmntcol + 1 '13
                If exSheet.Cells(xRow - 3, X) <> "" And exSheet.Cells(xRow - 3, X) <> 0 Then
                    'Ticket #20994 - Regular Hours - Extra Hours Paid / Weeks Worked
                    'exSheet.Cells(xRow, x) = exSheet.Cells(xRow - 5, x) / exSheet.Cells(xRow - 3, x)
                    exSheet.Cells(xRow, X) = (exSheet.Cells(xRow - 5, X) - IIf(IsNull(exSheet.Cells(xRow - 1, X)), 0, exSheet.Cells(xRow - 1, X))) / exSheet.Cells(xRow - 3, X)
                    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 30: " & X)

                End If
            Next
            
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 31")

            'Print Average Hours heading
            xRow = xRow + 1
            exSheet.Cells(xRow, 1) = "Average Hours"
            
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 32")

            'Compute & Print Average Hours
            For X = 2 To totmntcol + 1  '13
                If exSheet.Cells(xRow - 4, X) <> "" And exSheet.Cells(xRow - 4, X) <> 0 Then
                    'Ticket #20994 - Regular Hours / Weeks Worked
                    'exSheet.Cells(xRow, x) = exSheet.Cells(xRow - 5, x) + exSheet.Cells(xRow - 2, x) / exSheet.Cells(xRow - 3, x)
                    exSheet.Cells(xRow, X) = exSheet.Cells(xRow - 6, X) / exSheet.Cells(xRow - 4, X)

Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 32: " & X)
                
                End If
            Next
                      
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 33")

            'Shade the alternate cells for this employee
            For Y = 0 To 6
                For X = 1 To totmntcol  '12
                    If (X + 1) / 2 = Int((X + 1) / 2) Then
                        exSheet.Cells(xRow - Y, X + 1).Interior.ColorIndex = 15
                        
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 33: " & Y & " - " & X)

                    End If
                Next
            Next
            
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 34")

            'Border around
            If totmntcol = 12 Then
            
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 35")

                exSheet.Range("A" & xRow - 7 & ":N" & xRow - 7).Borders(xlEdgeTop).Weight = xlThin
                exSheet.Range("A" & xRow - 7 & ":N" & xRow - 7).Borders(xlEdgeBottom).Weight = xlThin
                exSheet.Range("A" & xRow & ":N" & xRow).Borders(xlEdgeBottom).Weight = xlThin
                exSheet.Range("A" & xRow - 7 & ":A" & xRow).Borders(xlEdgeLeft).Weight = xlThin
                exSheet.Range("N" & xRow - 7 & ":N" & xRow).Borders(xlEdgeRight).Weight = xlThin
                exSheet.Range("A" & xRow & ":N" & xRow).Borders(xlEdgeBottom).Weight = xlThin
                
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 36")

                'Border around the columns
                exSheet.Range("A" & xRow - 7 & ":N" & xRow).Borders(xlOutline).Weight = xlThin
            
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 37")

                'Line under each row
                For Y = 1 To 6
                    exSheet.Range("A" & xRow - Y & ":M" & xRow - Y).Borders(xlEdgeBottom).LineStyle = xlDot
                    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 38: " & Y)

                    'Print Totals for each row
                    exSheet.Range("N" & xRow - 6 & ":N" & xRow).Formula = "=SUM(A" & xRow - Y & ":M" & xRow - Y & ")"
                Next
            
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 39")

            Else
            
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 40")

                exSheet.Range("A" & xRow - 7 & ":O" & xRow - 7).Borders(xlEdgeTop).Weight = xlThin
                exSheet.Range("A" & xRow - 7 & ":O" & xRow - 7).Borders(xlEdgeBottom).Weight = xlThin
                exSheet.Range("A" & xRow & ":O" & xRow).Borders(xlEdgeBottom).Weight = xlThin
                exSheet.Range("A" & xRow - 7 & ":A" & xRow).Borders(xlEdgeLeft).Weight = xlThin
                exSheet.Range("O" & xRow - 7 & ":O" & xRow).Borders(xlEdgeRight).Weight = xlThin
                exSheet.Range("A" & xRow & ":O" & xRow).Borders(xlEdgeBottom).Weight = xlThin
                
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 41")

                'Border around the columns
                exSheet.Range("A" & xRow - 7 & ":O" & xRow).Borders(xlOutline).Weight = xlThin
            
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 42")

                'Line under each row
                For Y = 1 To 6
                    exSheet.Range("A" & xRow - Y & ":N" & xRow - Y).Borders(xlEdgeBottom).LineStyle = xlDot
                    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 43: " & Y)

                    'Print Totals for each row
                    exSheet.Range("O" & xRow - 6 & ":O" & xRow).Formula = "=SUM(A" & xRow - Y & ":N" & xRow - Y & ")"
                Next
            
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 44")

            End If

            
            xRow = xRow + 2

            rsHREmp.MoveNext
        Loop
    
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 45")

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        If chkDoNoLaunch.Visible = True Then
            If chkDoNoLaunch.Value = False Then
                Call Pause(1)
                If Not LanchXlsW98(xlsFileMat) Then
                    Shell "cmd /c " & GetShortName(xlsFileMat)
                End If
            Else
                MsgBox "Report generation complete."
            End If
        Else
            Call Pause(1)
            If Not LanchXlsW98(xlsFileMat) Then
                Shell "cmd /c " & GetShortName(xlsFileMat)
                
                MsgBox "Report generation complete."
            End If
        End If
    Else
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    
Call WriteFile("ChapmansExcelRpt_AverageHourRpt - 46")

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        MsgBox "No employees found in this selection criteria."
    End If
    rsHREmp.Close
    Set rsHREmp = Nothing
    
Exit Function
    
ChapmansExcelRpt_AverageHourRpt_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", "Average Hour", "SELECT")
'Resume Next
Set exSheet = Nothing
Set exBook = Nothing
Set exApp = Nothing

End Function

Private Function ChapmansExcelRpt_AbsenteeismRpt()
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim I, totNum
    Dim xRow As Integer

    Dim X, Y As Integer
    Dim SQLQ, sSQLQ As String
    Dim rsAttMatrix As New ADODB.Recordset
    Dim rsHREmp As New ADODB.Recordset
    Dim rsAttend As New ADODB.Recordset
    Dim rsHRParco As New ADODB.Recordset
    
    Dim xEXCFDCode As String
    Dim xEXCPDCode As String
    Dim xSICFDCode As String
    Dim xSICPDCode As String
    Dim xFSCFDCode As String
    Dim xFSCPDCode As String
    Dim xBRVFDCode As String
    Dim xBRVPDCode As String
    Dim xLATPDCode As String
    Dim xLFEPDCode As String
    Dim xWTHFDCode As String
    Dim xWTHPDCode As String
    Dim xOTHFDCode As String
    Dim xOTHPDCode As String
    
    Dim xNotesCodeEXCF As String
    Dim xNotesCodeEXCP As String
    Dim xNotesCodeSICF As String
    Dim xNotesCodeSICP As String
    Dim xNotesCodeFSCF As String
    Dim xNotesCodeFSCP As String
    Dim xNotesCodeBRVF As String
    Dim xNotesCodeBRVP As String
    Dim xNotesCodeLATP As String
    Dim xNotesCodeLFEP As String
    Dim xNotesCodeWTHF As String
    Dim xNotesCodeWTHP As String
    Dim xNotesCodeOTHF As String
    Dim xNotesCodeOTHP As String
    
    Dim flgShade As Boolean
    
    Dim ExcusedFD As Integer
    Dim ExcusedPD As Integer
    Dim SickFD As Integer
    Dim SickPD As Integer
    Dim FSickFD As Integer
    Dim FSickPD As Integer
    Dim BereaveFD As Integer
    Dim BereavePD As Integer
    Dim LateFD As Integer
    Dim LatePD As Integer
    Dim LfEarlyFD As Integer
    Dim LfEarlyPD As Integer
    Dim WeatherFD As Integer
    Dim WeatherPD As Integer
    Dim OtherFD As Integer
    Dim OtherPD As Integer
    Dim NotesCount As Integer
    
    Dim xNotesStr As String
    Dim xNotesAttCode As String

    On Error GoTo ChapmansExcelRpt_AbsenteeismRpt_Err
    
    'Selection criteria
    'sSQLQ = Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")")
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")

    'Get Excused Full Day codes
    xEXCFDCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_FULL_PARTIAL = 'F' AND AM_CODE_TYPE = 'EXC'"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xEXCFDCode) = 0 Then
            xEXCFDCode = xEXCFDCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xEXCFDCode = xEXCFDCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
        
    'Get Notes Code EXC - F
    xNotesCodeEXCF = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_NOTES = 1 AND AM_FULL_PARTIAL = 'F'"
    SQLQ = SQLQ & " AND AM_CODE_TYPE IN ('EXC')"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xNotesCodeEXCF) = 0 Then
            xNotesCodeEXCF = xNotesCodeEXCF & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xNotesCodeEXCF = xNotesCodeEXCF & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    'Get Excused Partial Day codes
    xEXCPDCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_FULL_PARTIAL = 'P' AND AM_CODE_TYPE = 'EXC'"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xEXCPDCode) = 0 Then
            xEXCPDCode = xEXCPDCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xEXCPDCode = xEXCPDCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    'Get Notes Code EXC - P
    xNotesCodeEXCP = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_NOTES = 1 AND AM_FULL_PARTIAL = 'P'"
    SQLQ = SQLQ & " AND AM_CODE_TYPE IN ('EXC')"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xNotesCodeEXCP) = 0 Then
            xNotesCodeEXCP = xNotesCodeEXCP & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xNotesCodeEXCP = xNotesCodeEXCP & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    
    'Get Sick Full Day codes
    xSICFDCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_FULL_PARTIAL = 'F' AND AM_CODE_TYPE = 'SIC'"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xSICFDCode) = 0 Then
            xSICFDCode = xSICFDCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xSICFDCode = xSICFDCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    'Get Notes Code SIC - F
    xNotesCodeSICF = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_NOTES = 1 AND AM_FULL_PARTIAL = 'F'"
    SQLQ = SQLQ & " AND AM_CODE_TYPE IN ('SIC')"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xNotesCodeSICF) = 0 Then
            xNotesCodeSICF = xNotesCodeSICF & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xNotesCodeSICF = xNotesCodeSICF & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    'Get Sick Partial Day codes
    xSICPDCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_FULL_PARTIAL = 'P' AND AM_CODE_TYPE = 'SIC'"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xSICPDCode) = 0 Then
            xSICPDCode = xSICPDCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xSICPDCode = xSICPDCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    'Get Notes Code SIC - P
    xNotesCodeSICP = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_NOTES = 1 AND AM_FULL_PARTIAL = 'P'"
    SQLQ = SQLQ & " AND AM_CODE_TYPE IN ('SIC')"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xNotesCodeSICP) = 0 Then
            xNotesCodeSICP = xNotesCodeSICP & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xNotesCodeSICP = xNotesCodeSICP & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    
    'Get Family Sick Full Day codes
    xFSCFDCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_FULL_PARTIAL = 'F' AND AM_CODE_TYPE = 'FSC'"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xFSCFDCode) = 0 Then
            xFSCFDCode = xFSCFDCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xFSCFDCode = xFSCFDCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    'Get Notes Code FSC - F
    xNotesCodeFSCF = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_NOTES = 1 AND AM_FULL_PARTIAL = 'F'"
    SQLQ = SQLQ & " AND AM_CODE_TYPE IN ('FSC')"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xNotesCodeFSCF) = 0 Then
            xNotesCodeFSCF = xNotesCodeFSCF & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xNotesCodeFSCF = xNotesCodeFSCF & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    'Get Family Sick Partial Day codes
    xFSCPDCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_FULL_PARTIAL = 'P' AND AM_CODE_TYPE = 'FSC'"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xFSCPDCode) = 0 Then
            xFSCPDCode = xFSCPDCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xFSCPDCode = xFSCPDCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    'Get Notes Code FSC - P
    xNotesCodeFSCP = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_NOTES = 1 AND AM_FULL_PARTIAL = 'P'"
    SQLQ = SQLQ & " AND AM_CODE_TYPE IN ('FSC')"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xNotesCodeFSCP) = 0 Then
            xNotesCodeFSCP = xNotesCodeFSCP & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xNotesCodeFSCP = xNotesCodeFSCP & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    
    'Get Bereavement Full Day codes
    xBRVFDCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_FULL_PARTIAL = 'F' AND AM_CODE_TYPE = 'BRV'"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xBRVFDCode) = 0 Then
            xBRVFDCode = xBRVFDCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xBRVFDCode = xBRVFDCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    'Get Notes Code BRV - F
    xNotesCodeBRVF = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_NOTES = 1 AND AM_FULL_PARTIAL = 'F'"
    SQLQ = SQLQ & " AND AM_CODE_TYPE IN ('BRV')"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xNotesCodeBRVF) = 0 Then
            xNotesCodeBRVF = xNotesCodeBRVF & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xNotesCodeBRVF = xNotesCodeBRVF & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    'Get Bereavement Partial Day codes
    xBRVPDCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_FULL_PARTIAL = 'P' AND AM_CODE_TYPE = 'BRV'"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xBRVPDCode) = 0 Then
            xBRVPDCode = xBRVPDCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xBRVPDCode = xBRVPDCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
        
    'Get Notes Code BRV - P
    xNotesCodeBRVP = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_NOTES = 1 AND AM_FULL_PARTIAL = 'P'"
    SQLQ = SQLQ & " AND AM_CODE_TYPE IN ('BRV')"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xNotesCodeBRVP) = 0 Then
            xNotesCodeBRVP = xNotesCodeBRVP & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xNotesCodeBRVP = xNotesCodeBRVP & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    
    'Get Late Partial Day codes
    xLATPDCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_FULL_PARTIAL = 'P' AND AM_CODE_TYPE = 'LAT'"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xLATPDCode) = 0 Then
            xLATPDCode = xLATPDCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xLATPDCode = xLATPDCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    'Get Notes Code LAT - P
    xNotesCodeLATP = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_NOTES = 1 AND AM_FULL_PARTIAL = 'P'"
    SQLQ = SQLQ & " AND AM_CODE_TYPE IN ('LAT')"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xNotesCodeLATP) = 0 Then
            xNotesCodeLATP = xNotesCodeLATP & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xNotesCodeLATP = xNotesCodeLATP & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    
    'Get Left Early Partial Day codes
    xLFEPDCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_FULL_PARTIAL = 'P' AND AM_CODE_TYPE = 'LFE'"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xLFEPDCode) = 0 Then
            xLFEPDCode = xLFEPDCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xLFEPDCode = xLFEPDCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing

    'Get Notes Code LFE - P
    xNotesCodeLFEP = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_NOTES = 1 AND AM_FULL_PARTIAL = 'P'"
    SQLQ = SQLQ & " AND AM_CODE_TYPE IN ('LFE')"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xNotesCodeLFEP) = 0 Then
            xNotesCodeLFEP = xNotesCodeLFEP & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xNotesCodeLFEP = xNotesCodeLFEP & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing


    'Get Weather Full Day codes
    xWTHFDCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_FULL_PARTIAL = 'F' AND AM_CODE_TYPE = 'WTH'"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xWTHFDCode) = 0 Then
            xWTHFDCode = xWTHFDCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xWTHFDCode = xWTHFDCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    'Get Notes Code WTH - F
    xNotesCodeWTHF = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_NOTES = 1 AND AM_FULL_PARTIAL = 'F'"
    SQLQ = SQLQ & " AND AM_CODE_TYPE IN ('WTH')"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xNotesCodeWTHF) = 0 Then
            xNotesCodeWTHF = xNotesCodeWTHF & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xNotesCodeWTHF = xNotesCodeWTHF & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    'Get Weather Partial Day codes
    xWTHPDCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_FULL_PARTIAL = 'P' AND AM_CODE_TYPE = 'WTH'"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xWTHPDCode) = 0 Then
            xWTHPDCode = xWTHPDCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xWTHPDCode = xWTHPDCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing

    'Get Notes Code WTH - P
    xNotesCodeWTHP = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_NOTES = 1 AND AM_FULL_PARTIAL = 'P'"
    SQLQ = SQLQ & " AND AM_CODE_TYPE IN ('WTH')"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xNotesCodeWTHP) = 0 Then
            xNotesCodeWTHP = xNotesCodeWTHP & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xNotesCodeWTHP = xNotesCodeWTHP & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing


    'Get Other Full Day codes
    xOTHFDCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_FULL_PARTIAL = 'F' AND AM_CODE_TYPE = 'OTH'"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xOTHFDCode) = 0 Then
            xOTHFDCode = xOTHFDCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xOTHFDCode = xOTHFDCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    'Get Notes Code OTH - F
    xNotesCodeOTHF = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_NOTES = 1 AND AM_FULL_PARTIAL = 'F'"
    SQLQ = SQLQ & " AND AM_CODE_TYPE IN ('OTH')"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xNotesCodeOTHF) = 0 Then
            xNotesCodeOTHF = xNotesCodeOTHF & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xNotesCodeOTHF = xNotesCodeOTHF & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing
    
    'Get Other Partial Day codes
    xOTHPDCode = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_FULL_PARTIAL = 'P' AND AM_CODE_TYPE = 'OTH'"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xOTHPDCode) = 0 Then
            xOTHPDCode = xOTHPDCode & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xOTHPDCode = xOTHPDCode & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing

    'Get Notes Code
    xNotesCodeOTHP = ""
    SQLQ = "SELECT AM_REASON FROM HRATT_MATRIX WHERE AM_NOTES = 1 AND AM_FULL_PARTIAL = 'P'"
    SQLQ = SQLQ & " AND AM_CODE_TYPE IN ('OTH')"
    rsAttMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttMatrix.EOF
        If Len(xNotesCodeOTHP) = 0 Then
            xNotesCodeOTHP = xNotesCodeOTHP & "'" & rsAttMatrix("AM_REASON") & "'"
        Else
            xNotesCodeOTHP = xNotesCodeOTHP & ",'" & rsAttMatrix("AM_REASON") & "'"
        End If
        rsAttMatrix.MoveNext
    Loop
    rsAttMatrix.Close
    Set rsAttMatrix = Nothing


    'Initialise/Open Excel Report file
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "AbsenteeismTmp.xls"
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "Absenteeism_" & Trim(glbUserID) & ".xls"

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Function
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0

    FileCopy xlsFileTmp, xlsFileMat

    Screen.MousePointer = HOURGLASS

    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)


    'Print the report headers
    exSheet.Cells(1, 1) = Format(Now, "mmm dd, yyyy hh:mm")
    exSheet.Cells(2, 1) = "ABSENTEEISM - PLANT"
    
    SQLQ = "SELECT PC_NAME FROM HRPARCO"
    rsHRParco.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsHRParco.EOF Then
        exSheet.Cells(3, 1) = rsHRParco("PC_NAME")
    Else
        exSheet.Cells(3, 1) = "CHAPMAN'S ICE CREAM LIMITED"
    End If
    rsHRParco.Close
    Set rsHRParco = Nothing
    
    exSheet.Cells(4, 1) = Format(dlpDateRange(2).Text, "mmmm dd, yyyy") & " To " & Format(dlpDateRange(3).Text, "mmmm dd, yyyy")
                      
    
    'For each employee get the breakdown of hours
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME, JH_DHRS "
    SQLQ = SQLQ & " FROM (HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR AND JH_CURRENT <>0)"
    If Len(sSQLQ) > 0 Then
        SQLQ = SQLQ & " WHERE " & sSQLQ
    End If
    SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: I = 0
        rsHREmp.MoveFirst

        xRow = 8
        flgShade = False
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            'Initialise variables
            ExcusedFD = 0
            ExcusedPD = 0
            SickFD = 0
            SickPD = 0
            FSickFD = 0
            FSickPD = 0
            BereaveFD = 0
            BereavePD = 0
            LateFD = 0
            LatePD = 0
            LfEarlyFD = 0
            LfEarlyPD = 0
            WeatherFD = 0
            WeatherPD = 0
            OtherFD = 0
            OtherPD = 0
            NotesCount = 0
            xNotesStr = ""
            
            'Print Employee Name
            exSheet.Cells(xRow, 1) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
        
            'Compute and Print Excused - Full Day
            If Len(xEXCFDCode) > 0 Then
                'Attendance
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xEXCFDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    ExcusedFD = IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 2) = ExcusedFD
                                        
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xEXCFDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    ExcusedFD = ExcusedFD + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 2) = ExcusedFD
                                                            
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Generate the Notes values
                If Len(xNotesCodeEXCF) > 0 Then
                    xNotesAttCode = ""
                    NotesCount = 0
                
                    SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT, AD_REASON FROM HR_ATTENDANCE "
                    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xNotesCodeEXCF & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    SQLQ = SQLQ & " GROUP BY AD_REASON"
                    SQLQ = SQLQ & " UNION"
                    SQLQ = SQLQ & " SELECT COUNT(AH_EMPNBR) AS TOTCOUNT, AH_REASON AS AD_REASON FROM HR_ATTENDANCE_HISTORY "
                    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xNotesCodeEXCF & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    
                    SQLQ = SQLQ & " GROUP BY AH_REASON"
                    SQLQ = SQLQ & " ORDER BY AD_REASON"
                    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Do While Not rsAttend.EOF
                        'Compute the sum of incidents between Attendance and Attendance History
                        If xNotesAttCode = "" Then xNotesAttCode = rsAttend("AD_REASON")
                        
                        'If the Reason is same (because of Attendance and History in two different rows), add them together
                        If xNotesAttCode = rsAttend("AD_REASON") Then
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        Else
                            'Reason is not same
                            If xNotesStr = "" Then
                                'New Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & NotesCount & "(F) " & xNotesAttCode
                            Else
                                'Append to the Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & ", " & NotesCount & "(F) " & xNotesAttCode
                            End If
                            NotesCount = 0
                            xNotesAttCode = ""
                            
                            'New Count for new code
                            xNotesAttCode = rsAttend("AD_REASON")
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                    
                    'Last Notes Code
                    If xNotesStr = "" And NotesCount <> 0 Then
                        xNotesStr = xNotesStr & NotesCount & "(F) " & xNotesAttCode
                    ElseIf NotesCount <> 0 Then
                        xNotesStr = xNotesStr & ", " & NotesCount & "(F) " & xNotesAttCode
                    End If
                    rsAttend.Close
                    Set rsAttend = Nothing
                End If
            End If
            
            'Compute and Print Excused - Partial Day
            If Len(xEXCPDCode) > 0 Then
                'Attendance
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xEXCPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    ExcusedPD = IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 3) = ExcusedPD
                                        
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xEXCPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    ExcusedPD = ExcusedPD + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 3) = ExcusedPD
                                                            
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                
                'Generate the Notes values
                If Len(xNotesCodeEXCP) > 0 Then
                    xNotesAttCode = ""
                    NotesCount = 0
                
                    SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT, AD_REASON FROM HR_ATTENDANCE "
                    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xNotesCodeEXCP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    SQLQ = SQLQ & " GROUP BY AD_REASON"
                    SQLQ = SQLQ & " UNION"
                    SQLQ = SQLQ & " SELECT COUNT(AH_EMPNBR) AS TOTCOUNT, AH_REASON AS AD_REASON FROM HR_ATTENDANCE_HISTORY "
                    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xNotesCodeEXCP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    
                    SQLQ = SQLQ & " GROUP BY AH_REASON"
                    SQLQ = SQLQ & " ORDER BY AD_REASON"
                    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Do While Not rsAttend.EOF
                        'Compute the sum of incidents between Attendance and Attendance History
                        If xNotesAttCode = "" Then xNotesAttCode = rsAttend("AD_REASON")
                        
                        'If the Reason is same (because of Attendance and History in two different rows), add them together
                        If xNotesAttCode = rsAttend("AD_REASON") Then
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        Else
                            'Reason is not same
                            If xNotesStr = "" Then
                                'New Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                            Else
                                'Append to the Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                            End If
                            NotesCount = 0
                            xNotesAttCode = ""
                            
                            'New Count for new code
                            xNotesAttCode = rsAttend("AD_REASON")
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                    
                    'Last Notes Code
                    If xNotesStr = "" And NotesCount <> 0 Then
                        xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                    ElseIf NotesCount <> 0 Then
                        xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                    End If
                    rsAttend.Close
                    Set rsAttend = Nothing
                End If
            End If
          
            
            
            'Compute and Print Sick - Full Day
            If Len(xSICFDCode) > 0 Then
                'Attendance
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xSICFDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    SickFD = IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 4) = SickFD
                                        
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xSICFDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    SickFD = SickFD + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 4) = SickFD
                                                            
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
            
                'Generate the Notes values
                If Len(xNotesCodeSICF) > 0 Then
                    xNotesAttCode = ""
                    NotesCount = 0
                
                    SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT, AD_REASON FROM HR_ATTENDANCE "
                    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xNotesCodeSICF & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    SQLQ = SQLQ & " GROUP BY AD_REASON"
                    SQLQ = SQLQ & " UNION"
                    SQLQ = SQLQ & " SELECT COUNT(AH_EMPNBR) AS TOTCOUNT, AH_REASON AS AD_REASON FROM HR_ATTENDANCE_HISTORY "
                    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xNotesCodeSICF & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    
                    SQLQ = SQLQ & " GROUP BY AH_REASON"
                    SQLQ = SQLQ & " ORDER BY AD_REASON"
                    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Do While Not rsAttend.EOF
                        'Compute the sum of incidents between Attendance and Attendance History
                        If xNotesAttCode = "" Then xNotesAttCode = rsAttend("AD_REASON")
                        
                        'If the Reason is same (because of Attendance and History in two different rows), add them together
                        If xNotesAttCode = rsAttend("AD_REASON") Then
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        Else
                            'Reason is not same
                            If xNotesStr = "" Then
                                'New Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & NotesCount & "(F) " & xNotesAttCode
                            Else
                                'Append to the Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & ", " & NotesCount & "(F) " & xNotesAttCode
                            End If
                            NotesCount = 0
                            xNotesAttCode = ""
                            
                            'New Count for new code
                            xNotesAttCode = rsAttend("AD_REASON")
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                    
                    'Last Notes Code
                    If xNotesStr = "" And NotesCount <> 0 Then
                        xNotesStr = xNotesStr & NotesCount & "(F) " & xNotesAttCode
                    ElseIf NotesCount <> 0 Then
                        xNotesStr = xNotesStr & ", " & NotesCount & "(F) " & xNotesAttCode
                    End If
                    rsAttend.Close
                    Set rsAttend = Nothing
                End If
            End If
            
            'Compute and Print Sick - Partial Day
            If Len(xSICPDCode) > 0 Then
                'Attendance
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xSICPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    SickPD = IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 5) = SickPD
                                        
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xSICPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    SickPD = SickPD + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 5) = SickPD
                                                            
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
            
                'Generate the Notes values
                If Len(xNotesCodeSICP) > 0 Then
                    xNotesAttCode = ""
                    NotesCount = 0
                
                    SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT, AD_REASON FROM HR_ATTENDANCE "
                    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xNotesCodeSICP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    SQLQ = SQLQ & " GROUP BY AD_REASON"
                    SQLQ = SQLQ & " UNION"
                    SQLQ = SQLQ & " SELECT COUNT(AH_EMPNBR) AS TOTCOUNT, AH_REASON AS AD_REASON FROM HR_ATTENDANCE_HISTORY "
                    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xNotesCodeSICP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    
                    SQLQ = SQLQ & " GROUP BY AH_REASON"
                    SQLQ = SQLQ & " ORDER BY AD_REASON"
                    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Do While Not rsAttend.EOF
                        'Compute the sum of incidents between Attendance and Attendance History
                        If xNotesAttCode = "" Then xNotesAttCode = rsAttend("AD_REASON")
                        
                        'If the Reason is same (because of Attendance and History in two different rows), add them together
                        If xNotesAttCode = rsAttend("AD_REASON") Then
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        Else
                            'Reason is not same
                            If xNotesStr = "" Then
                                'New Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                            Else
                                'Append to the Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                            End If
                            NotesCount = 0
                            xNotesAttCode = ""
                            
                            'New Count for new code
                            xNotesAttCode = rsAttend("AD_REASON")
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                    
                    'Last Notes Code
                    If xNotesStr = "" And NotesCount <> 0 Then
                        xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                    ElseIf NotesCount <> 0 Then
                        xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                    End If
                    rsAttend.Close
                    Set rsAttend = Nothing
                End If
            End If
            
            
            
            'Compute and Print Family Sick - Full Day
            If Len(xFSCFDCode) > 0 Then
                'Attendance
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xFSCFDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    FSickFD = IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 6) = FSickFD
                                        
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xFSCFDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    FSickFD = FSickFD + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 6) = FSickFD
                                                            
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
            
            
                'Generate the Notes values
                If Len(xNotesCodeFSCF) > 0 Then
                    xNotesAttCode = ""
                    NotesCount = 0
                
                    SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT, AD_REASON FROM HR_ATTENDANCE "
                    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xNotesCodeFSCF & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    SQLQ = SQLQ & " GROUP BY AD_REASON"
                    SQLQ = SQLQ & " UNION"
                    SQLQ = SQLQ & " SELECT COUNT(AH_EMPNBR) AS TOTCOUNT, AH_REASON AS AD_REASON FROM HR_ATTENDANCE_HISTORY "
                    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xNotesCodeFSCF & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    
                    SQLQ = SQLQ & " GROUP BY AH_REASON"
                    SQLQ = SQLQ & " ORDER BY AD_REASON"
                    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Do While Not rsAttend.EOF
                        'Compute the sum of incidents between Attendance and Attendance History
                        If xNotesAttCode = "" Then xNotesAttCode = rsAttend("AD_REASON")
                        
                        'If the Reason is same (because of Attendance and History in two different rows), add them together
                        If xNotesAttCode = rsAttend("AD_REASON") Then
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        Else
                            'Reason is not same
                            If xNotesStr = "" Then
                                'New Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & NotesCount & "(F) " & xNotesAttCode
                            Else
                                'Append to the Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & ", " & NotesCount & "(F) " & xNotesAttCode
                            End If
                            NotesCount = 0
                            xNotesAttCode = ""
                            
                            'New Count for new code
                            xNotesAttCode = rsAttend("AD_REASON")
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                    
                    'Last Notes Code
                    If xNotesStr = "" And NotesCount <> 0 Then
                        xNotesStr = xNotesStr & NotesCount & "(F) " & xNotesAttCode
                    ElseIf NotesCount <> 0 Then
                        xNotesStr = xNotesStr & ", " & NotesCount & "(F) " & xNotesAttCode
                    End If
                    rsAttend.Close
                    Set rsAttend = Nothing
                End If
            End If
            
            'Compute and Print Family Sick - Partial Day
            If Len(xFSCPDCode) > 0 Then
                'Attendance
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xFSCPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    FSickPD = IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 7) = FSickPD
                                        
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xFSCPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    FSickPD = FSickPD + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 7) = FSickPD
                                                            
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Generate the Notes values
                If Len(xNotesCodeFSCP) > 0 Then
                    xNotesAttCode = ""
                    NotesCount = 0
                
                    SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT, AD_REASON FROM HR_ATTENDANCE "
                    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xNotesCodeFSCP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    SQLQ = SQLQ & " GROUP BY AD_REASON"
                    SQLQ = SQLQ & " UNION"
                    SQLQ = SQLQ & " SELECT COUNT(AH_EMPNBR) AS TOTCOUNT, AH_REASON AS AD_REASON FROM HR_ATTENDANCE_HISTORY "
                    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xNotesCodeFSCP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    
                    SQLQ = SQLQ & " GROUP BY AH_REASON"
                    SQLQ = SQLQ & " ORDER BY AD_REASON"
                    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Do While Not rsAttend.EOF
                        'Compute the sum of incidents between Attendance and Attendance History
                        If xNotesAttCode = "" Then xNotesAttCode = rsAttend("AD_REASON")
                        
                        'If the Reason is same (because of Attendance and History in two different rows), add them together
                        If xNotesAttCode = rsAttend("AD_REASON") Then
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        Else
                            'Reason is not same
                            If xNotesStr = "" Then
                                'New Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                            Else
                                'Append to the Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                            End If
                            NotesCount = 0
                            xNotesAttCode = ""
                            
                            'New Count for new code
                            xNotesAttCode = rsAttend("AD_REASON")
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                    
                    'Last Notes Code
                    If xNotesStr = "" And NotesCount <> 0 Then
                        xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                    ElseIf NotesCount <> 0 Then
                        xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                    End If
                    rsAttend.Close
                    Set rsAttend = Nothing
                End If
            End If
            
            
            
            'Compute and Print Bereavement - Full Day
            If Len(xBRVFDCode) > 0 Then
                'Attendance
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xBRVFDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    BereaveFD = IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 8) = BereaveFD
                                        
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xBRVFDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    BereaveFD = BereaveFD + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 8) = BereaveFD
                                                            
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
            
                'Generate the Notes values
                If Len(xNotesCodeBRVF) > 0 Then
                    xNotesAttCode = ""
                    NotesCount = 0
                
                    SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT, AD_REASON FROM HR_ATTENDANCE "
                    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xNotesCodeBRVF & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    SQLQ = SQLQ & " GROUP BY AD_REASON"
                    SQLQ = SQLQ & " UNION"
                    SQLQ = SQLQ & " SELECT COUNT(AH_EMPNBR) AS TOTCOUNT, AH_REASON AS AD_REASON FROM HR_ATTENDANCE_HISTORY "
                    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xNotesCodeBRVF & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    
                    SQLQ = SQLQ & " GROUP BY AH_REASON"
                    SQLQ = SQLQ & " ORDER BY AD_REASON"
                    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Do While Not rsAttend.EOF
                        'Compute the sum of incidents between Attendance and Attendance History
                        If xNotesAttCode = "" Then xNotesAttCode = rsAttend("AD_REASON")
                        
                        'If the Reason is same (because of Attendance and History in two different rows), add them together
                        If xNotesAttCode = rsAttend("AD_REASON") Then
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        Else
                            'Reason is not same
                            If xNotesStr = "" Then
                                'New Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & NotesCount & "(F) " & xNotesAttCode
                            Else
                                'Append to the Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & ", " & NotesCount & "(F) " & xNotesAttCode
                            End If
                            NotesCount = 0
                            xNotesAttCode = ""
                            
                            'New Count for new code
                            xNotesAttCode = rsAttend("AD_REASON")
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                    
                    'Last Notes Code
                    If xNotesStr = "" And NotesCount <> 0 Then
                        xNotesStr = xNotesStr & NotesCount & "(F) " & xNotesAttCode
                    ElseIf NotesCount <> 0 Then
                        xNotesStr = xNotesStr & ", " & NotesCount & "(F) " & xNotesAttCode
                    End If
                    rsAttend.Close
                    Set rsAttend = Nothing
                End If
            End If
            
            'Compute and Print Bereavement - Partial Day
            If Len(xBRVPDCode) > 0 Then
                'Attendance
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xBRVPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    BereavePD = IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 9) = BereavePD
                                        
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xBRVPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    BereavePD = BereavePD + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 9) = BereavePD
                                                            
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
            
                'Generate the Notes values
                If Len(xNotesCodeBRVP) > 0 Then
                    xNotesAttCode = ""
                    NotesCount = 0
                
                    SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT, AD_REASON FROM HR_ATTENDANCE "
                    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xNotesCodeBRVP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    SQLQ = SQLQ & " GROUP BY AD_REASON"
                    SQLQ = SQLQ & " UNION"
                    SQLQ = SQLQ & " SELECT COUNT(AH_EMPNBR) AS TOTCOUNT, AH_REASON AS AD_REASON FROM HR_ATTENDANCE_HISTORY "
                    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xNotesCodeBRVP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    
                    SQLQ = SQLQ & " GROUP BY AH_REASON"
                    SQLQ = SQLQ & " ORDER BY AD_REASON"
                    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Do While Not rsAttend.EOF
                        'Compute the sum of incidents between Attendance and Attendance History
                        If xNotesAttCode = "" Then xNotesAttCode = rsAttend("AD_REASON")
                        
                        'If the Reason is same (because of Attendance and History in two different rows), add them together
                        If xNotesAttCode = rsAttend("AD_REASON") Then
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        Else
                            'Reason is not same
                            If xNotesStr = "" Then
                                'New Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                            Else
                                'Append to the Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                            End If
                            NotesCount = 0
                            xNotesAttCode = ""
                            
                            'New Count for new code
                            xNotesAttCode = rsAttend("AD_REASON")
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                    
                    'Last Notes Code
                    If xNotesStr = "" And NotesCount <> 0 Then
                        xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                    ElseIf NotesCount <> 0 Then
                        xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                    End If
                    rsAttend.Close
                    Set rsAttend = Nothing
                End If
            End If
            
            
                        
            'Compute and Print Late - Partial Day
            If Len(xLATPDCode) > 0 Then
                'Attendance
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xLATPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    LatePD = IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 10) = LatePD
                                        
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xLATPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    LatePD = LatePD + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 10) = LatePD
                                                            
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
            
                'Generate the Notes values
                If Len(xNotesCodeLATP) > 0 Then
                    xNotesAttCode = ""
                    NotesCount = 0
                
                    SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT, AD_REASON FROM HR_ATTENDANCE "
                    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xNotesCodeLATP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    SQLQ = SQLQ & " GROUP BY AD_REASON"
                    SQLQ = SQLQ & " UNION"
                    SQLQ = SQLQ & " SELECT COUNT(AH_EMPNBR) AS TOTCOUNT, AH_REASON AS AD_REASON FROM HR_ATTENDANCE_HISTORY "
                    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xNotesCodeLATP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    
                    SQLQ = SQLQ & " GROUP BY AH_REASON"
                    SQLQ = SQLQ & " ORDER BY AD_REASON"
                    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Do While Not rsAttend.EOF
                        'Compute the sum of incidents between Attendance and Attendance History
                        If xNotesAttCode = "" Then xNotesAttCode = rsAttend("AD_REASON")
                        
                        'If the Reason is same (because of Attendance and History in two different rows), add them together
                        If xNotesAttCode = rsAttend("AD_REASON") Then
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        Else
                            'Reason is not same
                            If xNotesStr = "" Then
                                'New Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                            Else
                                'Append to the Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                            End If
                            NotesCount = 0
                            xNotesAttCode = ""
                            
                            'New Count for new code
                            xNotesAttCode = rsAttend("AD_REASON")
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                    
                    'Last Notes Code
                    If xNotesStr = "" And NotesCount <> 0 Then
                        xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                    ElseIf NotesCount <> 0 Then
                        xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                    End If
                    rsAttend.Close
                    Set rsAttend = Nothing
                End If
            End If
            
            
            
            'Compute and Print Left Early - Partial Day
            If Len(xLFEPDCode) > 0 Then
                'Attendance
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xLFEPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    LfEarlyPD = IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 11) = LfEarlyPD
                                        
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xLFEPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    LfEarlyPD = LfEarlyPD + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 11) = LfEarlyPD
                                                            
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
            
                'Generate the Notes values
                If Len(xNotesCodeLFEP) > 0 Then
                    xNotesAttCode = ""
                    NotesCount = 0
                
                    SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT, AD_REASON FROM HR_ATTENDANCE "
                    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xNotesCodeLFEP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    SQLQ = SQLQ & " GROUP BY AD_REASON"
                    SQLQ = SQLQ & " UNION"
                    SQLQ = SQLQ & " SELECT COUNT(AH_EMPNBR) AS TOTCOUNT, AH_REASON AS AD_REASON FROM HR_ATTENDANCE_HISTORY "
                    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xNotesCodeLFEP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    
                    SQLQ = SQLQ & " GROUP BY AH_REASON"
                    SQLQ = SQLQ & " ORDER BY AD_REASON"
                    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Do While Not rsAttend.EOF
                        'Compute the sum of incidents between Attendance and Attendance History
                        If xNotesAttCode = "" Then xNotesAttCode = rsAttend("AD_REASON")
                        
                        'If the Reason is same (because of Attendance and History in two different rows), add them together
                        If xNotesAttCode = rsAttend("AD_REASON") Then
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        Else
                            'Reason is not same
                            If xNotesStr = "" Then
                                'New Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                            Else
                                'Append to the Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                            End If
                            NotesCount = 0
                            xNotesAttCode = ""
                            
                            'New Count for new code
                            xNotesAttCode = rsAttend("AD_REASON")
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                    
                    'Last Notes Code
                    If xNotesStr = "" And NotesCount <> 0 Then
                        xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                    ElseIf NotesCount <> 0 Then
                        xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                    End If
                    rsAttend.Close
                    Set rsAttend = Nothing
                End If
            End If
            
            
            
            'Compute and Print Weather - Full Day
            If Len(xWTHFDCode) > 0 Then
                'Attendance
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xWTHFDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    WeatherFD = IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 12) = WeatherFD
                                        
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xWTHFDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    WeatherFD = WeatherFD + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 12) = WeatherFD
                                                            
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
            
                'Generate the Notes values
                If Len(xNotesCodeWTHF) > 0 Then
                    xNotesAttCode = ""
                    NotesCount = 0
                
                    SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT, AD_REASON FROM HR_ATTENDANCE "
                    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xNotesCodeWTHF & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    SQLQ = SQLQ & " GROUP BY AD_REASON"
                    SQLQ = SQLQ & " UNION"
                    SQLQ = SQLQ & " SELECT COUNT(AH_EMPNBR) AS TOTCOUNT, AH_REASON AS AD_REASON FROM HR_ATTENDANCE_HISTORY "
                    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xNotesCodeWTHF & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    
                    SQLQ = SQLQ & " GROUP BY AH_REASON"
                    SQLQ = SQLQ & " ORDER BY AD_REASON"
                    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Do While Not rsAttend.EOF
                        'Compute the sum of incidents between Attendance and Attendance History
                        If xNotesAttCode = "" Then xNotesAttCode = rsAttend("AD_REASON")
                        
                        'If the Reason is same (because of Attendance and History in two different rows), add them together
                        If xNotesAttCode = rsAttend("AD_REASON") Then
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        Else
                            'Reason is not same
                            If xNotesStr = "" Then
                                'New Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & NotesCount & "(F) " & xNotesAttCode
                            Else
                                'Append to the Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & ", " & NotesCount & "(F) " & xNotesAttCode
                            End If
                            NotesCount = 0
                            xNotesAttCode = ""
                            
                            'New Count for new code
                            xNotesAttCode = rsAttend("AD_REASON")
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                    
                    'Last Notes Code
                    If xNotesStr = "" And NotesCount <> 0 Then
                        xNotesStr = xNotesStr & NotesCount & "(F) " & xNotesAttCode
                    ElseIf NotesCount <> 0 Then
                        xNotesStr = xNotesStr & ", " & NotesCount & "(F) " & xNotesAttCode
                    End If
                    rsAttend.Close
                    Set rsAttend = Nothing
                End If
            End If
            
            'Compute and Print Weather - Partial Day
            If Len(xWTHPDCode) > 0 Then
                'Attendance
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xWTHPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    WeatherPD = IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 13) = WeatherPD
                                        
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xWTHPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    WeatherPD = WeatherPD + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 13) = WeatherPD
                                                            
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
            
                'Generate the Notes values
                If Len(xNotesCodeWTHP) > 0 Then
                    xNotesAttCode = ""
                    NotesCount = 0
                
                    SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT, AD_REASON FROM HR_ATTENDANCE "
                    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xNotesCodeWTHP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    SQLQ = SQLQ & " GROUP BY AD_REASON"
                    SQLQ = SQLQ & " UNION"
                    SQLQ = SQLQ & " SELECT COUNT(AH_EMPNBR) AS TOTCOUNT, AH_REASON AS AD_REASON FROM HR_ATTENDANCE_HISTORY "
                    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xNotesCodeWTHP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    
                    SQLQ = SQLQ & " GROUP BY AH_REASON"
                    SQLQ = SQLQ & " ORDER BY AD_REASON"
                    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Do While Not rsAttend.EOF
                        'Compute the sum of incidents between Attendance and Attendance History
                        If xNotesAttCode = "" Then xNotesAttCode = rsAttend("AD_REASON")
                        
                        'If the Reason is same (because of Attendance and History in two different rows), add them together
                        If xNotesAttCode = rsAttend("AD_REASON") Then
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        Else
                            'Reason is not same
                            If xNotesStr = "" Then
                                'New Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                            Else
                                'Append to the Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                            End If
                            NotesCount = 0
                            xNotesAttCode = ""
                            
                            'New Count for new code
                            xNotesAttCode = rsAttend("AD_REASON")
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                    
                    'Last Notes Code
                    If xNotesStr = "" And NotesCount <> 0 Then
                        xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                    ElseIf NotesCount <> 0 Then
                        xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                    End If
                    rsAttend.Close
                    Set rsAttend = Nothing
                End If
            End If
            
            
            
            'Compute and Print Other - Full Day
            If Len(xOTHFDCode) > 0 Then
                'Attendance
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xOTHFDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    OtherFD = IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 14) = OtherFD
                                        
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xOTHFDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    OtherFD = OtherFD + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 14) = OtherFD
                                                            
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
            
                'Generate the Notes values
                If Len(xNotesCodeOTHF) > 0 Then
                    xNotesAttCode = ""
                    NotesCount = 0
                
                    SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT, AD_REASON FROM HR_ATTENDANCE "
                    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xNotesCodeOTHF & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    SQLQ = SQLQ & " GROUP BY AD_REASON"
                    SQLQ = SQLQ & " UNION"
                    SQLQ = SQLQ & " SELECT COUNT(AH_EMPNBR) AS TOTCOUNT, AH_REASON AS AD_REASON FROM HR_ATTENDANCE_HISTORY "
                    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xNotesCodeOTHF & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    
                    SQLQ = SQLQ & " GROUP BY AH_REASON"
                    SQLQ = SQLQ & " ORDER BY AD_REASON"
                    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Do While Not rsAttend.EOF
                        'Compute the sum of incidents between Attendance and Attendance History
                        If xNotesAttCode = "" Then xNotesAttCode = rsAttend("AD_REASON")
                        
                        'If the Reason is same (because of Attendance and History in two different rows), add them together
                        If xNotesAttCode = rsAttend("AD_REASON") Then
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        Else
                            'Reason is not same
                            If xNotesStr = "" Then
                                'New Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & NotesCount & "(F) " & xNotesAttCode
                            Else
                                'Append to the Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & ", " & NotesCount & "(F) " & xNotesAttCode
                            End If
                            NotesCount = 0
                            xNotesAttCode = ""
                            
                            'New Count for new code
                            xNotesAttCode = rsAttend("AD_REASON")
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                    
                    'Last Notes Code
                    If xNotesStr = "" And NotesCount <> 0 Then
                        xNotesStr = xNotesStr & NotesCount & "(F) " & xNotesAttCode
                    ElseIf NotesCount <> 0 Then
                        xNotesStr = xNotesStr & ", " & NotesCount & "(F) " & xNotesAttCode
                    End If
                    rsAttend.Close
                    Set rsAttend = Nothing
                End If
            End If
            
            'Compute and Print Other - Partial Day
            If Len(xOTHPDCode) > 0 Then
                'Attendance
                SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xOTHPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    OtherPD = IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 15) = OtherPD
                                        
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Attendance History
                SQLQ = "SELECT COUNT(AH_EMPNBR) AS TOTCOUNT FROM HR_ATTENDANCE_HISTORY "
                SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xOTHPDCode & ")"
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                End If
                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsAttend.EOF
                    OtherPD = OtherPD + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                    
                    'Print total Excused Full Day Hours
                    exSheet.Cells(xRow, 15) = OtherPD
                                                            
                    rsAttend.MoveNext
                Loop
                rsAttend.Close
                Set rsAttend = Nothing
                
                'Generate the Notes values
                If Len(xNotesCodeOTHP) > 0 Then
                    xNotesAttCode = ""
                    NotesCount = 0
                
                    SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOTCOUNT, AD_REASON FROM HR_ATTENDANCE "
                    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON IN (" & xNotesCodeOTHP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    SQLQ = SQLQ & " GROUP BY AD_REASON"
                    SQLQ = SQLQ & " UNION"
                    SQLQ = SQLQ & " SELECT COUNT(AH_EMPNBR) AS TOTCOUNT, AH_REASON AS AD_REASON FROM HR_ATTENDANCE_HISTORY "
                    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AH_REASON IN (" & xNotesCodeOTHP & ")"
                    If IsDate(dlpDateRange(2).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(dlpDateRange(2).Text)
                    End If
                    If IsDate(dlpDateRange(3).Text) Then
                        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(dlpDateRange(3).Text)
                    End If
                    
                    SQLQ = SQLQ & " GROUP BY AH_REASON"
                    SQLQ = SQLQ & " ORDER BY AD_REASON"
                    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Do While Not rsAttend.EOF
                        'Compute the sum of incidents between Attendance and Attendance History
                        If xNotesAttCode = "" Then xNotesAttCode = rsAttend("AD_REASON")
                        
                        'If the Reason is same (because of Attendance and History in two different rows), add them together
                        If xNotesAttCode = rsAttend("AD_REASON") Then
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        Else
                            'Reason is not same
                            If xNotesStr = "" Then
                                'New Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                            Else
                                'Append to the Notes string - Incident Count and Attendance Code
                                xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                            End If
                            NotesCount = 0
                            xNotesAttCode = ""
                            
                            'New Count for new code
                            xNotesAttCode = rsAttend("AD_REASON")
                            NotesCount = NotesCount + IIf(Not IsNull(rsAttend("TOTCOUNT")), rsAttend("TOTCOUNT"), 0)
                        End If
                        
                        rsAttend.MoveNext
                    Loop
                    
                    'Last Notes Code
                    If xNotesStr = "" And NotesCount <> 0 Then
                        xNotesStr = xNotesStr & NotesCount & "(P) " & xNotesAttCode
                    ElseIf NotesCount <> 0 Then
                        xNotesStr = xNotesStr & ", " & NotesCount & "(P) " & xNotesAttCode
                    End If
                    rsAttend.Close
                    Set rsAttend = Nothing
                End If
            End If
            
            
            
            'Compute and Print Total
            exSheet.Range("P" & xRow & ":P" & xRow).Formula = "=(B" & xRow & "+D" & xRow & "+F" & xRow & "+H" & xRow & "+L" & xRow & "+N" & xRow & ")"
            exSheet.Range("Q" & xRow & ":Q" & xRow).Formula = "=(C" & xRow & "+E" & xRow & "+G" & xRow & "+I" & xRow & "+J" & xRow & "+K" & xRow & "+M" & xRow & "+O" & xRow & ")"
            
            'Print the Notes string
            If Len(xNotesStr) > 0 Then
                exSheet.Cells(xRow, 18) = xNotesStr
            End If
            
            'Shade or Not - alternative rows to be shaded
            If flgShade Then
                'Shade current row
                exSheet.Range("A" & xRow & ":R" & xRow).Interior.ColorIndex = 15
                
                'Next row do not shade
                flgShade = False
            Else
                'Next row shade
                flgShade = True
            End If
            
            'Border around the columns
            exSheet.Range("B" & xRow & ":R" & xRow).Borders(xlOutline).Weight = xlThin
            
            'Dotted Line under the row
            exSheet.Range("A" & xRow & ":R" & xRow).Borders(xlEdgeBottom).LineStyle = xlDot

            xRow = xRow + 1

            rsHREmp.MoveNext
        Loop
        
        'Solid Line under the last row
        exSheet.Range("A" & xRow - 1 & ":R" & xRow - 1).Borders(xlEdgeBottom).LineStyle = xlSolid
        
        
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    Else
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
        
        MsgBox "No employees found in this selection criteria."
    End If
    rsHREmp.Close
    Set rsHREmp = Nothing

Exit Function

ChapmansExcelRpt_AbsenteeismRpt_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", "Absenteeism", "SELECT")
'Resume Next
Set exSheet = Nothing
Set exBook = Nothing
Set exApp = Nothing

End Function

Private Sub SurreyPlace_3MonthVacationAccrual()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsSal As New ADODB.Recordset
    Dim rsVacAcc As New ADODB.Recordset
    Dim rsAttend As New ADODB.Recordset
    Dim rsHRParco As New ADODB.Recordset
    Dim SQLQ As String
    
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim I, totNum
    Dim xRow As Integer
    
    On Error GoTo SurreyPlace_3MonthVacationAccrual_Err

    'Initialise/Open Excel Report file
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "VacAccrualTmp.xls"
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "VacAccrual_" & Trim(glbUserID) & ".xls"

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0

    FileCopy xlsFileTmp, xlsFileMat

    Screen.MousePointer = HOURGLASS

    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)

    'Print the report headers
    exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy hh:mm")
    
    SQLQ = "SELECT PC_NAME FROM HRPARCO"
    rsHRParco.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsHRParco.EOF Then
        exSheet.Cells(2, 2) = rsHRParco("PC_NAME")
    Else
        exSheet.Cells(2, 2) = "SURREY PLACE CENTRE"
    End If
    rsHRParco.Close
    Set rsHRParco = Nothing
    
    exSheet.Cells(3, 2) = "For the Period: " & Format(dlpDateRange(0).Text, "mmmm dd, yyyy") & " - " & Format(dlpDateRange(1).Text, "mmmm dd, yyyy")
    
    'Column headings
    exSheet.Cells(8, 3) = Format(dlpDateRange(1).Text, "mmm dd/yyyy")
    exSheet.Cells(8, 4) = Format(dlpDateRange(1).Text, "mmm dd/yyyy")
    exSheet.Cells(8, 5) = Format(dlpDateRange(1).Text, "mmm dd/yyyy")
    exSheet.Cells(8, 7) = Format(dlpDateRange(1).Text, "mmm dd/yyyy")

    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME, ED_PVAC FROM HREMP"
    If Len(glbstrSelCri) > 0 Then
        SQLQ = SQLQ & " WHERE " & Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", "("), "}", ")"), "Uppercase", "Upper"), "[", "("), "]", ")")
    End If
    SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME, ED_EMPNBR"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: I = 0
        rsHREmp.MoveFirst

        xRow = 9
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
        
            'Employee #, Name, Previous Vac
            exSheet.Cells(xRow, 1) = rsHREmp("ED_EMPNBR")
            exSheet.Cells(xRow, 2) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
            exSheet.Cells(xRow, 3) = Round(rsHREmp("ED_PVAC"), 2)
    
            'Accrued
            exSheet.Cells(xRow, 4) = Round(Get_EntitlementforPeriod(rsHREmp("ED_EMPNBR"), dlpDateRange(0).Text, dlpDateRange(1).Text), 2)
            
            'Taken
            SQLQ = "SELECT SUM(AD_HRS) AS TAKEN FROM HR_ATTENDANCE "
            SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON LIKE 'VAC%'"
            SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(0).Text)
            SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1).Text)
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsAttend.EOF Then
                'Print Taken
                exSheet.Cells(xRow, 5) = Round(IIf(Not IsNull(rsAttend("TAKEN")), rsAttend("TAKEN"), 0), 2)
            Else
                exSheet.Cells(xRow, 5) = Round(0, 2)
            End If
            rsAttend.Close
            Set rsAttend = Nothing
            
            
            'Total Balance (Prv + Accrued - Taken)
            exSheet.Cells(xRow, 6) = Round((exSheet.Cells(xRow, 3) + exSheet.Cells(xRow, 4)) - exSheet.Cells(xRow, 5), 2)
            
            'Rate - as of To Date
            SQLQ = "SELECT SH_EMPNBR, SH_EDATE, SH_SALARY, SH_SALCD, SH_WHRS FROM HR_SALARY_HISTORY"
            SQLQ = SQLQ & " WHERE SH_EDATE <= " & Date_SQL(dlpDateRange(1).Text)
            SQLQ = SQLQ & " AND SH_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " ORDER BY SH_EDATE DESC, SH_SALARY DESC"
            rsSal.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsSal.EOF Then
                rsSal.MoveFirst
                
                'Print Rate
                If rsSal("SH_SALCD") = "H" Then
                    exSheet.Cells(xRow, 7) = Round2DEC(IIf(Not IsNull(rsSal("SH_SALARY")), rsSal("SH_SALARY"), 0))
                ElseIf rsSal("SH_SALCD") = "A" Then
                    If Not IsNull(rsSal("SH_SALARY")) Then
                        exSheet.Cells(xRow, 7) = Round2DEC((rsSal("SH_SALARY") / Val(rsSal("SH_WHRS"))) / 52)
                    End If
                End If
            Else
                exSheet.Cells(xRow, 7) = Round2DEC(0)
            End If
            rsSal.Close
            Set rsSal = Nothing
            
            'Vacation Balance Accrual (Total Bal * Rate)
            exSheet.Cells(xRow, 8) = Round((exSheet.Cells(xRow, 6) * exSheet.Cells(xRow, 7)), 2)
            
            xRow = xRow + 1
            
            rsHREmp.MoveNext
        Loop
        
        'Print Solid line under the last row and Totals of Vacation Balance Accrual
        exSheet.Range("H" & xRow & ":H" & xRow).Borders(xlEdgeBottom).LineStyle = xlSolid
        exSheet.Cells(xRow + 1, 8).Formula = "=Round(SUM(H9:H" & xRow - 1 & "),2)"

    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT

    Call Pause(1)
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
    
Exit Sub

SurreyPlace_3MonthVacationAccrual_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", "3MonthVacationAccrual", "SELECT")
Resume Next

End Sub

Private Function Get_EntitlementforPeriod(xEmpNbr, xFrom As Date, xTo As Date)
    Dim SQLQ As String
    Dim xRunTimes As Long, blIsLast As Boolean, lngRecs As Long

    On Error GoTo Get_EntitlementforPeriod_Err

    Get_EntitlementforPeriod = 0
    dblEmpEntitle = 0

    SQLQ = "SELECT * FROM HRVACENT ORDER BY VE_ID"
    rsEntRules.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsEntRules.EOF Then
        Do While Not rsEntRules.EOF
            
            If Not CR_SnapEntitle(xEmpNbr, "VAC") Then GoTo nextRule
            
            If (UCase(glbCompEntVac$) = "M" Or UCase(glbCompEntVac$) = "N") Then
                While Not snapEntitle.EOF
                    
                    fglbToDate = dlpDateRange(1).Text
                    fglbAsOf = CVDate(Format(month(dlpDateRange(0).Text) & "/" & Day(rsEntRules("VE_EDATE")) & "/" & Year(dlpDateRange(0).Text), "mm/dd/yyyy"))
                    
                    For xRunTimes = 1 To 12
                        blIsLast = False
                        If xRunTimes = 12 Then blIsLast = True
                        If Not modAnnVacation(blIsLast) Then GoTo nextRule

                        fglbAsOf = DateAdd("m", 1, CVDate(fglbAsOf))
                        fglbToDate = dlpDateRange(1).Text
                        
                        If DateDiff("d", fglbAsOf, dlpDateRange(1).Text) < 0 Then Exit For
                    Next
                    snapEntitle.MoveNext
                Wend
            End If
            snapEntitle.Close
nextRule:
            rsEntRules.MoveNext
        Loop
    
    End If
    rsEntRules.Close
    Set rsEntRules = Nothing
    
    Get_EntitlementforPeriod = dblEmpEntitle
    
Exit Function
    
Get_EntitlementforPeriod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", "Get_EntitlementforPeriod", "SELECT")
Resume Next

End Function

Private Function CR_SnapEntitle(xEmpNbr, Optional xType)
    Dim SQLQ As String
    
    CR_SnapEntitle = False
    
    On Error GoTo CR_SnapEntitle_Err
    
    
    SQLQ = "SELECT ED_EMPNBR, ED_PVAC,ED_VAC,ED_PSICK,ED_SICK,ED_EFDATE,ED_ETDATE,ED_EFDATES,ED_ETDATES,"
    SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1"
    
    If glbOracle Then
        SQLQ = SQLQ & " FROM HREMP, HR_JOB_HISTORY WHERE HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
        SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_CURRENT<>0"
    Else
        SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
        SQLQ = SQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT<>0"
    End If
    
    If Not IsMissing(xType) Then
        SQLQ = SQLQ & " AND " & getWSQLQ(xType)
    Else
        SQLQ = SQLQ & " AND " & getWSQLQ("")
    End If
    
    SQLQ = SQLQ & " AND ED_EMPNBR IN (" & xEmpNbr & ")"
    
    If Len(rsEntRules("VE_GRPCD")) > 0 Then
        SQLQ = SQLQ & " AND ED_EMPNBR IN "
        SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
        SQLQ = SQLQ & " WHERE JB_GRPCD = '" & rsEntRules("VE_GRPCD") & "') "
    End If
    SQLQ = SQLQ & " AND " & Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", "("), "}", ")"), "Uppercase", "Upper"), "[", "("), "]", ")")
    
    If snapEntitle.State <> 0 Then snapEntitle.Close
    snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
    
    If Not snapEntitle.EOF Then
        CR_SnapEntitle = True
    Else
        CR_SnapEntitle = False
    End If
    
Exit Function

CR_SnapEntitle_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapEntitle", "EntitlementforPeriod", "Select")

End Function

Private Function getWSQLQ(xType) As String
Dim xDiv, xDept, xORG, xAsOf, xEMP, xEmpMode, xGRPCE
Dim xLoc, xSection
Dim xFromDate
Dim xToDate, fglbESQLQ As String

fglbESQLQ = glbSeleDeptUn
If Len(rsEntRules("VE_DEPT")) > 0 Then fglbESQLQ = fglbESQLQ & " AND  ED_DEPTNO = '" & rsEntRules("VE_DEPT") & "' "
If Len(rsEntRules("VE_DIV")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & rsEntRules("VE_DIV") & "' "
If Len(rsEntRules("VE_ORG")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & rsEntRules("VE_ORG") & "' "
If Len(rsEntRules("VE_EMP")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP = '" & rsEntRules("VE_EMP") & "' "
If glbLinamar Then
    If Len(rsEntRules("VE_SECTION")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SALDIST = '" & rsEntRules("VE_SECTION") & "' "
Else
    If Not glbCBrant Then 'added by Bryan 18/Apr/2006 Ticket#10495
        If Len(rsEntRules("VE_SECTION")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & rsEntRules("VE_SECTION") & "' "
    End If
End If
If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #18235
    If xType = "VAC" Then
        If Len(rsEntRules("VE_LOC")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_VADIM1 = '" & rsEntRules("VE_LOC") & "' "
    ElseIf xType = "SICK" Then
        If Len(rsEntRules("VE_LOC")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_VADIM2 = '" & rsEntRules("VE_LOC") & "' "
    End If
Else
    If Len(rsEntRules("VE_LOC")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & rsEntRules("VE_LOC") & "' "
End If

If Len(rsEntRules("VE_PT")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & rsEntRules("VE_PT") & "' "

getWSQLQ = fglbESQLQ

End Function

Private Function modAnnVacation(isLast As Boolean)
    Dim empNo As Long
    Dim dblPrevEntitle#, strDivision$
    Dim strJob$, dblServiceYears#
    Dim spt As Variant, varStartDate As Variant, lngRecs&
    Dim dblDHours#, intWhereFit&, X%, Y%, z%, dblNewEntitle#
    Dim dblFTEHours#
    Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
    Dim xAsOf, fglbWDate As String
    Dim if_Entitle As Boolean
    Dim dblEntitleDays
    
    On Error GoTo modUpdateSelection_Err

    modAnnVacation = False
    
    if_Entitle = False
    
    Select Case glbCompWDate$ ' sets field reference for basic 'which date'
        Case "O": fglbWDate$ = "ED_DOH"
        Case "S": fglbWDate$ = "ED_SENDTE"
        Case "U": fglbWDate$ = "ED_UNION"
        Case "L": fglbWDate$ = "ED_LTHIRE"
        Case "D": fglbWDate$ = "ED_USRDAT1"
    End Select

    empNo = snapEntitle("ED_EMPNBR")
    If IsNull(snapEntitle("ED_PVAC")) Then
        dblPrevEntitle# = 0
    Else
        dblPrevEntitle# = snapEntitle("ED_PVAC")
    End If

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
    rsJOB.Close
    Set rsJOB = Nothing
    
    If glbLinamar Then dblDHours# = 8

    xAsOf = fglbAsOf
    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    intWhereFit& = -1

    If rsEntRules("VE_EMONTH") > 0 Then
        If dblServiceYears# >= CDbl(rsEntRules("VE_BMONTH")) And dblServiceYears# <= CDbl(rsEntRules("VE_EMONTH")) Then
            intWhereFit& = X%
            If Len(rsEntRules("VE_ENTITLE")) > 0 Then if_Entitle = True
        End If
    End If

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

    If if_Entitle Then
        dblNewEntitle# = rsEntRules("VE_ENTITLE")
        dblNewMax# = 0
        If rsEntRules("VE_TYPE") = "D" Then           ' Entitlements entered in days
            If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX") * dblDHours#
            dblNewEntitle# = dblNewEntitle# * dblDHours#
            dblEntitleUpd = dblNewEntitle
        End If
        If rsEntRules("VE_TYPE") = "F" Then
            If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX") * dblFTEHours# * dblDHours#
            dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
            dblEntitleUpd = dblNewEntitle
        End If
        If rsEntRules("VE_TYPE") = "H" Then
            If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX")
        End If
        dblEntitleUpd# = dblEmpEntitle + dblNewEntitle  ' accumulate monthly values

        If dblNewMax <> 0 Then          'only do if not zero
            If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
                dblEntitleUpd = dblNewMax - dblPrevEntitle#
            End If
        End If
    End If

    
    dblEmpEntitle = dblEntitleUpd

lblNextRec:

modAnnVacation = True

Exit Function

modUpdateSelection_Err:
    'These errors are:
    '13=type mismatch; 94=invalid use of null; 3018=couln't find field 'item'
    If Err = 13 Or Err = 94 Or Err = 3018 Then
        Err = 0
        Resume Next
    End If

    Screen.MousePointer = DEFAULT
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modAnnVac", "EntitlementforPeriod", "SELECT")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        'Rollback
        Resume Next
    Else
        Unload Me
    End If
End Function

Private Function Round2DEC(tmpNUM, Optional HourlyRate As String)    'laura nov 10, 1997
Dim strNUM As String, X%

If glbFrench Then
    tmpNUM = Replace(Replace(tmpNUM, ",", "."), " ", "")
End If

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
If glbCompSerial = "S/N - 2375W" Then   'City of Timmins
    If GetEmpData(glbLEE_ID, "ED_REGION") <> "S" Then
        Round2DEC = Round(tmpNUM, 2)
    Else
        Round2DEC = Round(tmpNUM, glbCompDecHR)
    End If
Else
    Round2DEC = Round(Val(tmpNUM), glbCompDecHR)
End If
End Function

Private Function getEmpCrsCount(xEmpNo, xCode)
Dim SQLQ As String
Dim rsTemp As New ADODB.Recordset
Dim retval As Boolean
    retval = True
    SQLQ = "SELECT * FROM HREMPHIS_WRK WHERE EE_WRKEMP = '" & glbUserID & "' "
    SQLQ = SQLQ & "AND EE_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND EE_HISTYPE = '" & xCode & "' "
    SQLQ = SQLQ & "AND EE_SALCD = 'N' "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        retval = False
    End If
    rsTemp.Close
                
    getEmpCrsCount = retval
End Function

Private Function isEmptyCourseMaster()
Dim SQLQ As String
Dim rsTemp As New ADODB.Recordset
Dim retval As Boolean
    retval = True
    SQLQ = "SELECT ES_ID, ES_CTYPE FROM HR_COURSECODE_MASTER "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        retval = False
    End If
    rsTemp.Close
    isEmptyCourseMaster = retval
End Function

Private Function getWPSCodeFromCrsCode(xCode)
Dim SQLQ As String
Dim rsTemp As New ADODB.Recordset
Dim retval As String
    retval = xCode
    SQLQ = "SELECT ES_ID,ES_CRSCODE,ES_WPSCODE FROM HR_COURSECODE_MASTER "
    SQLQ = SQLQ & "WHERE ES_CRSCODE = '" & xCode & "' "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("ES_WPSCODE")) Then
            If Len(rsTemp("ES_WPSCODE")) > 0 Then
                retval = rsTemp("ES_WPSCODE")
            End If
        End If
    End If
    rsTemp.Close
    getWPSCodeFromCrsCode = retval
End Function

Private Sub SUCCESS_Accrual_XLS_Report()
    Dim rsHRParco As New ADODB.Recordset
    Dim rsHREmp As New ADODB.Recordset
    Dim rsAttend As New ADODB.Recordset
    Dim rsSal As New ADODB.Recordset
    Dim SQLQ As String

    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim I, totNum
    Dim xRow, xCount, xTotalRow, xFoundRow, xCanaIncRow, xSearchRow As Long
    Dim xFoundStart, xCanaIncStart As Long
    Dim xCurVac, xNoDays, xDailyEntitl, xNoDaysDateRange
    Dim xExcelRptPath  As String
    
    On Error GoTo SUCCESS_Accrual_XLS_Report_Err

    
    'Ticket #22034 - Get Excel reports path
    If gsTRAININGMATRIX Then
        xExcelRptPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xExcelRptPath) = 0 Then
        xExcelRptPath = glbIHRREPORTS
    End If

    'Initialise/Open Excel Report file
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2422AccrualTmp.xls"
    
    'Ticket #22034 - To allow report saving in different path
    'xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "Accrual_" & Trim(glbUserID) & ".xls"
    xlsFileMat = xExcelRptPath & IIf(Right(xExcelRptPath, 1) = "\", "", "\") & "Accrual_" & Trim(glbUserID) & ".xls"

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0

    FileCopy xlsFileTmp, xlsFileMat

    Screen.MousePointer = HOURGLASS

    'Calculate the # of Days between Date Range
    xNoDaysDateRange = 0
    'Ticket #21277 - Changed From Date to Vacation Entitlement From Date. Done in the Employee loop below
    'xNoDaysDateRange = DateDiff("d", CVDate(dlpDateRange(0).Text), CVDate(dlpDateRange(1).Text)) + 1
        
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)

    'Print the report headers
    SQLQ = "SELECT PC_NAME FROM HRPARCO"
    rsHRParco.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsHRParco.EOF Then
        exSheet.Cells(1, 1) = rsHRParco("PC_NAME")
    Else
        exSheet.Cells(1, 1) = "S.U.C.C.E.S.S."
    End If
    rsHRParco.Close
    Set rsHRParco = Nothing
    
    'Ticket #21277 - No need for From Date
    'exSheet.Cells(2, 1) = "For the Date Range: " & Format(dlpDateRange(0).Text, "mmmm dd, yyyy") & " - " & Format(dlpDateRange(1).Text, "mmmm dd, yyyy")
    exSheet.Cells(2, 1) = "As of: " & Format(dlpDateRange(1).Text, "mmmm dd, yyyy")
    
    'Retrieve Employee Records
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME, ED_DEPTNO, ED_EFDATE, ED_ETDATE, ED_VAC, ED_PVAC FROM HREMP"
    If Len(glbstrSelCri) > 0 Then
        SQLQ = SQLQ & " WHERE " & Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", "("), "}", ")"), "Uppercase", "Upper"), "[", "("), "]", ")")
    End If
        
    SQLQ = SQLQ & " ORDER BY ED_EMPNBR, ED_SURNAME, ED_FNAME"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: I = 0
        rsHREmp.MoveFirst

        xRow = 8
        xCount = 0
        Do While Not rsHREmp.EOF
        
            If IsNull(rsHREmp("ED_EFDATE")) Or IsNull(rsHREmp("ED_ETDATE")) Or IsNull(rsHREmp("ED_VAC")) Then
                GoTo skip_emp
            End If
            
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            'Ticket #21277 - Changed From Date to Vacation Entitlement From Date
            xNoDaysDateRange = 0
            xNoDaysDateRange = DateDiff("d", CVDate(rsHREmp("ED_EFDATE")), CVDate(dlpDateRange(1).Text)) + 1
            
            'No.
            xCount = xCount + 1
            exSheet.Cells(xRow, 1) = xCount
            
            'Employee #, Surname Name, First Name
            exSheet.Cells(xRow, 2) = rsHREmp("ED_EMPNBR")
            exSheet.Cells(xRow, 3) = rsHREmp("ED_SURNAME")
            exSheet.Cells(xRow, 4) = rsHREmp("ED_FNAME")
            
            'Current Hours/Week, Hourly Rate
            SQLQ = "SELECT SH_EMPNBR, SH_EDATE, SH_SALARY, SH_SALCD, SH_WHRS FROM HR_SALARY_HISTORY"
            'SQLQ = SQLQ & " WHERE SH_EDATE <= " & Date_SQL(dlpDateRange(1).Text)
            SQLQ = SQLQ & " WHERE SH_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND SH_CURRENT <> 0"
            SQLQ = SQLQ & " ORDER BY SH_EDATE DESC, SH_SALARY DESC"
            rsSal.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsSal.EOF Then
                rsSal.MoveFirst
                
                'Hours/Week
                exSheet.Cells(xRow, 5) = IIf(IsNull(rsSal("SH_WHRS")), "", rsSal("SH_WHRS"))
                
                'Hourly Rate
                If rsSal("SH_SALCD") = "H" Then
                    exSheet.Cells(xRow, 15) = Round2DEC(IIf(Not IsNull(rsSal("SH_SALARY")), rsSal("SH_SALARY"), 0))
                ElseIf rsSal("SH_SALCD") = "A" Then
                    If Not IsNull(rsSal("SH_SALARY")) Then
                        If Not IsNull(rsSal("SH_WHRS")) And rsSal("SH_WHRS") <> 0 Then
                            exSheet.Cells(xRow, 15) = Round2DEC((rsSal("SH_SALARY") / Val(rsSal("SH_WHRS"))) / 52)
                        Else
                            exSheet.Cells(xRow, 15) = Round2DEC(0)
                        End If
                    End If
                End If
            Else
                exSheet.Cells(xRow, 15) = Round2DEC(0)
            End If
            rsSal.Close
            Set rsSal = Nothing
            
            'Department
            exSheet.Cells(xRow, 6) = getDeptDesc(rsHREmp("ED_DEPTNO"))
            
            'Entitlement Period
            exSheet.Cells(xRow, 7) = Format(rsHREmp("ED_EFDATE"), "mm/dd/yyyy")
            exSheet.Cells(xRow, 8) = Format(rsHREmp("ED_ETDATE"), "mm/dd/yyyy")
            
            'Current Vacation
            exSheet.Cells(xRow, 9) = Round(IIf(IsNull(rsHREmp("ED_VAC")), 0, rsHREmp("ED_VAC")), 2)
            
            'Calculate Entitlement as of date range
            'So the Daily Entitlement will be Current Vacation Entitlement (X) divide by # of Days
            'between the Vacation Entitlement Period. The Current Vacation Entitlement is assumed to be
            'Annual Entitlement
            xDailyEntitl = 0
            xCurVac = 0
            xNoDays = 0
            xCurVac = Round(rsHREmp("ED_VAC"), 2)
            xNoDays = DateDiff("d", CVDate(rsHREmp("ED_EFDATE")), CVDate(rsHREmp("ED_ETDATE"))) + 1
            If xNoDays <> 0 Then
                'Removing rounding to avoid slight difference in value from ED_VAC value
                'xDailyEntitl = Round((xCurVac / xNoDays), 2)
                xDailyEntitl = (xCurVac / xNoDays)
            Else
                xDailyEntitl = 0
            End If
            
            'Entitlement as of Date Range (Daily Entitlement * # of Days between Selection Criteria Date Range)
            exSheet.Cells(xRow, 11) = Round(xDailyEntitl * xNoDaysDateRange, 2)
            
            
            'Previous Vacation
            exSheet.Cells(xRow, 12) = Round(IIf(IsNull(rsHREmp("ED_PVAC")), 0, rsHREmp("ED_PVAC")), 2)
                
            'Total Hours Vacation between Date Range
            SQLQ = "SELECT SUM(AD_HRS) AS TAKEN FROM HR_ATTENDANCE "
            SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON LIKE 'VAC%'"
            'Ticket #21277 - Changed From Date to Vacation Entitlement From Date
            'SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(0).Text)
            SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(rsHREmp("ED_EFDATE"))
            SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1).Text)
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsAttend.EOF Then
                'Print Taken
                exSheet.Cells(xRow, 13) = Round(IIf(Not IsNull(rsAttend("TAKEN")), rsAttend("TAKEN"), 0), 2)
            Else
                exSheet.Cells(xRow, 13) = Round(0, 2)
            End If
            rsAttend.Close
            Set rsAttend = Nothing
                        
                        
            'Balance in hrs - A+B-C: (Entitlement as of date range + Prv Vac - Taken)
            exSheet.Cells(xRow, 14) = Round((exSheet.Cells(xRow, 11) + exSheet.Cells(xRow, 12)) - exSheet.Cells(xRow, 13), 2)
            
            
            'Total in $ - D x E: (Balance * Hourly Rate)
            exSheet.Cells(xRow, 16) = Round((exSheet.Cells(xRow, 14) * exSheet.Cells(xRow, 15)), 2)
            
            xRow = xRow + 1
skip_emp:
            rsHREmp.MoveNext
        Loop

        'Print Solid line under the last row and Totals of Balance in Hrs and Total in $
        xTotalRow = xRow
        exSheet.Range("A" & xRow - 1 & ":P" & xRow - 1).Borders(xlEdgeBottom).LineStyle = xlSolid
        exSheet.Rows(xRow).Font.Bold = True
        exSheet.Cells(xRow, 13) = "TOTAL - Vacation Accrual as at " & Format(dlpDateRange(1).Text, "mmmm dd, yyyy")
        exSheet.Cells(xRow, 13).HorizontalAlignment = xlRight
        exSheet.Cells(xRow, 14).Formula = "=Round(SUM(N8:N" & xRow - 1 & "),2)"
        exSheet.Cells(xRow, 15) = "hrs ="
        exSheet.Cells(xRow, 15).HorizontalAlignment = xlCenter
        exSheet.Cells(xRow, 16).Formula = "=Round(SUM(P8:P" & xRow - 1 & "),2)"
        
        'Less Foundation
        xRow = xRow + 1
        xFoundRow = xRow
        
        'Foundation Department
        exSheet.Cells(xRow, 12) = "LESS - "
        exSheet.Cells(xRow, 13) = getDeptDesc("P99")    '"Foundation"
        exSheet.Cells(xRow, 14) = Round(0, 2)
        exSheet.Cells(xRow, 15) = "hrs ="
        exSheet.Cells(xRow, 15).HorizontalAlignment = xlCenter
        exSheet.Cells(xRow, 16) = Round(0, 2)
        
        
        'Less CanaInc
        xRow = xRow + 1
        xCanaIncRow = xRow
        exSheet.Cells(xRow, 12) = "LESS - "
        exSheet.Cells(xRow, 13) = getDeptDesc("W99") '"CanaInc"
        exSheet.Cells(xRow, 14) = Round(0, 2)
        exSheet.Cells(xRow, 15) = "hrs ="
        exSheet.Cells(xRow, 15).HorizontalAlignment = xlCenter
        exSheet.Cells(xRow, 16) = Round(0, 2)
        

        'Print Solid line
        exSheet.Range("N" & xRow & ":P" & xRow).Borders(xlEdgeBottom).LineStyle = xlSolid
        
        'Sum the above totals
        xRow = xRow + 1
        exSheet.Cells(xRow, 14).Formula = "=Round(SUM(N" & xTotalRow & ":N" & xCanaIncRow & "),2)"
        exSheet.Cells(xRow, 15) = "hrs ="
        exSheet.Cells(xRow, 15).HorizontalAlignment = xlCenter
        exSheet.Cells(xRow, 16).Formula = "=Round(SUM(P" & xTotalRow & ":P" & xCanaIncRow & "),2)"
    End If

    rsHREmp.Close
    Set rsHREmp = Nothing

    
    'Print Employees from SUCCESS - Foundation Department
    'Retrieve Employee Records
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME, ED_DEPTNO, ED_EFDATE, ED_ETDATE, ED_VAC, ED_PVAC FROM HREMP"
    If Len(glbstrSelCri) > 0 Then
        SQLQ = SQLQ & " WHERE " & Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", "("), "}", ")"), "Uppercase", "Upper"), "[", "("), "]", ")")
    End If
    'SUCCESS - Foundation Department
    SQLQ = SQLQ & " AND ED_DEPTNO = 'P99'" 'SFND'"
    SQLQ = SQLQ & " ORDER BY ED_EMPNBR, ED_SURNAME, ED_FNAME"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: I = 0
        rsHREmp.MoveFirst

        xRow = xRow + 2
        xFoundStart = xRow
        xCount = 0
        Do While Not rsHREmp.EOF
            If IsNull(rsHREmp("ED_EFDATE")) Or IsNull(rsHREmp("ED_ETDATE")) Or IsNull(rsHREmp("ED_VAC")) Then
                GoTo skip_emp2
            End If
            
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
        
            'No.
            xCount = xCount + 1
            exSheet.Cells(xRow, 1) = xCount
            
            'Employee #, Surname Name, First Name
            exSheet.Cells(xRow, 2) = rsHREmp("ED_EMPNBR")
            exSheet.Cells(xRow, 3) = rsHREmp("ED_SURNAME")
            exSheet.Cells(xRow, 4) = rsHREmp("ED_FNAME")
            
            'Current Hours/Week, Hourly Rate
            SQLQ = "SELECT SH_EMPNBR, SH_EDATE, SH_SALARY, SH_SALCD, SH_WHRS FROM HR_SALARY_HISTORY"
            'SQLQ = SQLQ & " WHERE SH_EDATE <= " & Date_SQL(dlpDateRange(1).Text)
            SQLQ = SQLQ & " WHERE SH_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND SH_CURRENT <> 0"
            SQLQ = SQLQ & " ORDER BY SH_EDATE DESC, SH_SALARY DESC"
            rsSal.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsSal.EOF Then
                rsSal.MoveFirst
                
                'Hours/Week
                exSheet.Cells(xRow, 5) = IIf(IsNull(rsSal("SH_WHRS")), "", rsSal("SH_WHRS"))
                
                'Hourly Rate
                If rsSal("SH_SALCD") = "H" Then
                    exSheet.Cells(xRow, 15) = Round2DEC(IIf(Not IsNull(rsSal("SH_SALARY")), rsSal("SH_SALARY"), 0))
                ElseIf rsSal("SH_SALCD") = "A" Then
                    If Not IsNull(rsSal("SH_SALARY")) Then
                        If Not IsNull(rsSal("SH_WHRS")) And rsSal("SH_WHRS") <> 0 Then
                            exSheet.Cells(xRow, 15) = Round2DEC((rsSal("SH_SALARY") / Val(rsSal("SH_WHRS"))) / 52)
                        Else
                            exSheet.Cells(xRow, 15) = Round2DEC(0)
                        End If
                    End If
                End If
            Else
                exSheet.Cells(xRow, 15) = Round2DEC(0)
            End If
            rsSal.Close
            Set rsSal = Nothing
            
            'Department
            exSheet.Cells(xRow, 6) = getDeptDesc(rsHREmp("ED_DEPTNO"))
            
            'Entitlement Period
            exSheet.Cells(xRow, 7) = Format(rsHREmp("ED_EFDATE"), "mm/dd/yyyy")
            exSheet.Cells(xRow, 8) = Format(rsHREmp("ED_ETDATE"), "mm/dd/yyyy")
            
            'Current Vacation
            exSheet.Cells(xRow, 9) = Round(IIf(IsNull(rsHREmp("ED_VAC")), 0, rsHREmp("ED_VAC")), 2)
            
            'Calculate Entitlement as of date range
            'So the Daily Entitlement will be Current Vacation Entitlement (X) divide by # of Days
            'between the Vacation Entitlement Period. The Current Vacation Entitlement is assumed to be
            'Annual Entitlement
            xDailyEntitl = 0
            xCurVac = 0
            xNoDays = 0
            xCurVac = Round(rsHREmp("ED_VAC"), 2)
            xNoDays = DateDiff("d", CVDate(rsHREmp("ED_EFDATE")), CVDate(rsHREmp("ED_ETDATE"))) + 1
            If xNoDays <> 0 Then
                'Removing rounding to avoid slight difference in value from ED_VAC value
                'xDailyEntitl = Round((xCurVac / xNoDays), 2)
                xDailyEntitl = (xCurVac / xNoDays)
            Else
                xDailyEntitl = 0
            End If
            
            
            'Ticket #21277 - Changed From Date to Vacation Entitlement From Date
            xNoDaysDateRange = 0
            xNoDaysDateRange = DateDiff("d", CVDate(rsHREmp("ED_EFDATE")), CVDate(dlpDateRange(1).Text)) + 1
            
            'Ticket #21277 - Changed From Date to Vacation Entitlement From Date
            'Entitlement as of Date Range (Daily Entitlement * # of Days between Selection Criteria Date Range)
            exSheet.Cells(xRow, 11) = Round(xDailyEntitl * xNoDaysDateRange, 2)
                        
                        
            'Previous Vacation
            exSheet.Cells(xRow, 12) = Round(IIf(IsNull(rsHREmp("ED_PVAC")), 0, rsHREmp("ED_PVAC")), 2)
                
            'Total Hours Vacation between Date Range
            SQLQ = "SELECT SUM(AD_HRS) AS TAKEN FROM HR_ATTENDANCE "
            SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON LIKE 'VAC%'"
            'Ticket #21277 - Changed From Date to Vacation Entitlement From Date
            'SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(0).Text)
            SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(rsHREmp("ED_EFDATE"))
            SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1).Text)
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsAttend.EOF Then
                'Print Taken
                exSheet.Cells(xRow, 13) = Round(IIf(Not IsNull(rsAttend("TAKEN")), rsAttend("TAKEN"), 0), 2)
            Else
                exSheet.Cells(xRow, 13) = Round(0, 2)
            End If
            rsAttend.Close
            Set rsAttend = Nothing
                        
                        
            'Balance in hrs - A+B-C: (Entitlement as of date range + Prv Vac - Taken)
            exSheet.Cells(xRow, 14) = Round((exSheet.Cells(xRow, 11) + exSheet.Cells(xRow, 12)) - exSheet.Cells(xRow, 13), 2)
            
            
            'Total in $ - D x E: (Balance * Hourly Rate)
            exSheet.Cells(xRow, 16) = Round((exSheet.Cells(xRow, 14) * exSheet.Cells(xRow, 15)), 2)
            
            xRow = xRow + 1
skip_emp2:
            rsHREmp.MoveNext
        Loop

        'Print Solid line under the last row and Totals of Balance in Hrs and Total in $
        exSheet.Range("A" & xRow - 1 & ":P" & xRow - 1).Borders(xlEdgeBottom).LineStyle = xlSolid
        exSheet.Rows(xRow).Font.Bold = True
        exSheet.Cells(xRow, 13) = "TOTAL - Foundation - Vacation Accrual as at " & Format(dlpDateRange(1).Text, "mmmm dd, yyyy")
        exSheet.Cells(xRow, 13).HorizontalAlignment = xlRight
        exSheet.Cells(xRow, 14).Formula = "=Round(SUM(N" & xFoundStart & ":N" & xRow - 1 & "),2)"
        exSheet.Cells(xRow, 15) = "hrs ="
        exSheet.Cells(xRow, 15).HorizontalAlignment = xlCenter
        exSheet.Cells(xRow, 16).Formula = "=Round(SUM(P" & xFoundStart & ":P" & xRow - 1 & "),2)"
        
        
        'Print the totals to Less Foundation row. Show in Negative value
        exSheet.Cells(xFoundRow, 14) = 0 - Round(exSheet.Cells(xRow, 14), 2)
        exSheet.Cells(xFoundRow, 16) = 0 - Round(exSheet.Cells(xRow, 16), 2)
    End If

    rsHREmp.Close
    Set rsHREmp = Nothing


    'Print Employees from Cana Inc. Department
    'Retrieve Employee Records
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME, ED_DEPTNO, ED_EFDATE, ED_ETDATE, ED_VAC, ED_PVAC FROM HREMP"
    If Len(glbstrSelCri) > 0 Then
        SQLQ = SQLQ & " WHERE " & Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", "("), "}", ")"), "Uppercase", "Upper"), "[", "("), "]", ")")
    End If
    'SUCCESS - Cana Inc. Department
    SQLQ = SQLQ & " AND ED_DEPTNO = 'W99'" 'CANAINC'"
    SQLQ = SQLQ & " ORDER BY ED_EMPNBR, ED_SURNAME, ED_FNAME"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: I = 0
        rsHREmp.MoveFirst

        xRow = xRow + 2
        xCanaIncStart = xRow
        xCount = 0
        Do While Not rsHREmp.EOF
            If IsNull(rsHREmp("ED_EFDATE")) Or IsNull(rsHREmp("ED_ETDATE")) Or IsNull(rsHREmp("ED_VAC")) Then
                GoTo skip_emp3
            End If
        
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
        
            'No.
            xCount = xCount + 1
            exSheet.Cells(xRow, 1) = xCount
            
            'Employee #, Surname Name, First Name
            exSheet.Cells(xRow, 2) = rsHREmp("ED_EMPNBR")
            exSheet.Cells(xRow, 3) = rsHREmp("ED_SURNAME")
            exSheet.Cells(xRow, 4) = rsHREmp("ED_FNAME")
            
            'Current Hours/Week, Hourly Rate
            SQLQ = "SELECT SH_EMPNBR, SH_EDATE, SH_SALARY, SH_SALCD, SH_WHRS FROM HR_SALARY_HISTORY"
            'SQLQ = SQLQ & " WHERE SH_EDATE <= " & Date_SQL(dlpDateRange(1).Text)
            SQLQ = SQLQ & " WHERE SH_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND SH_CURRENT <> 0"
            SQLQ = SQLQ & " ORDER BY SH_EDATE DESC, SH_SALARY DESC"
            rsSal.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsSal.EOF Then
                rsSal.MoveFirst
                
                'Hours/Week
                exSheet.Cells(xRow, 5) = IIf(IsNull(rsSal("SH_WHRS")), "", rsSal("SH_WHRS"))
                
                'Hourly Rate
                If rsSal("SH_SALCD") = "H" Then
                    exSheet.Cells(xRow, 15) = Round2DEC(IIf(Not IsNull(rsSal("SH_SALARY")), rsSal("SH_SALARY"), 0))
                ElseIf rsSal("SH_SALCD") = "A" Then
                    If Not IsNull(rsSal("SH_SALARY")) Then
                        If Not IsNull(rsSal("SH_WHRS")) And rsSal("SH_WHRS") <> 0 Then
                            exSheet.Cells(xRow, 15) = Round2DEC((rsSal("SH_SALARY") / Val(rsSal("SH_WHRS"))) / 52)
                        Else
                            exSheet.Cells(xRow, 15) = Round2DEC(0)
                        End If
                    End If
                End If
            Else
                exSheet.Cells(xRow, 15) = Round2DEC(0)
            End If
            rsSal.Close
            Set rsSal = Nothing
            
            'Department
            exSheet.Cells(xRow, 6) = getDeptDesc(rsHREmp("ED_DEPTNO"))
            
            'Entitlement Period
            exSheet.Cells(xRow, 7) = Format(rsHREmp("ED_EFDATE"), "mm/dd/yyyy")
            exSheet.Cells(xRow, 8) = Format(rsHREmp("ED_ETDATE"), "mm/dd/yyyy")
            
            'Current Vacation
            exSheet.Cells(xRow, 9) = Round(rsHREmp("ED_VAC"), 2)
            
            'Calculate Entitlement as of date range
            'So the Daily Entitlement will be Current Vacation Entitlement (X) divide by # of Days
            'between the Vacation Entitlement Period. The Current Vacation Entitlement is assumed to be
            'Annual Entitlement
            xDailyEntitl = 0
            xCurVac = 0
            xNoDays = 0
            xCurVac = Round(rsHREmp("ED_VAC"), 2)
            xNoDays = DateDiff("d", CVDate(rsHREmp("ED_EFDATE")), CVDate(rsHREmp("ED_ETDATE"))) + 1
            If xNoDays <> 0 Then
                'Removing rounding to avoid slight difference in value from ED_VAC value
                'xDailyEntitl = Round((xCurVac / xNoDays), 2)
                xDailyEntitl = (xCurVac / xNoDays)
            Else
                xDailyEntitl = 0
            End If
            
            'Ticket #21277 - Changed From Date to Vacation Entitlement From Date
            xNoDaysDateRange = 0
            xNoDaysDateRange = DateDiff("d", CVDate(rsHREmp("ED_EFDATE")), CVDate(dlpDateRange(1).Text)) + 1
            
            'Ticket #21277 - Changed From Date to Vacation Entitlement From Date
            'Entitlement as of Date Range (Daily Entitlement * # of Days between Selection Criteria Date Range)
            exSheet.Cells(xRow, 11) = Round(xDailyEntitl * xNoDaysDateRange, 2)
            
            
            'Previous Vacation
            exSheet.Cells(xRow, 12) = Round(IIf(IsNull(rsHREmp("ED_PVAC")), 0, rsHREmp("ED_PVAC")), 2)
                
            'Total Hours Vacation between Date Range
            SQLQ = "SELECT SUM(AD_HRS) AS TAKEN FROM HR_ATTENDANCE "
            SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND AD_REASON LIKE 'VAC%'"
            'Ticket #21277 - Changed From Date to Vacation Entitlement From Date
            'SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(0).Text)
            SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(rsHREmp("ED_EFDATE"))
            SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1).Text)
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsAttend.EOF Then
                'Print Taken
                exSheet.Cells(xRow, 13) = Round(IIf(Not IsNull(rsAttend("TAKEN")), rsAttend("TAKEN"), 0), 2)
            Else
                exSheet.Cells(xRow, 13) = Round(0, 2)
            End If
            rsAttend.Close
            Set rsAttend = Nothing
                        
                        
            'Balance in hrs - A+B-C: (Entitlement as of date range + Prv Vac - Taken)
            exSheet.Cells(xRow, 14) = Round((exSheet.Cells(xRow, 11) + exSheet.Cells(xRow, 12)) - exSheet.Cells(xRow, 13), 2)
            
            
            'Total in $ - D x E: (Balance * Hourly Rate)
            exSheet.Cells(xRow, 16) = Round((exSheet.Cells(xRow, 14) * exSheet.Cells(xRow, 15)), 2)
            
            xRow = xRow + 1
skip_emp3:
            rsHREmp.MoveNext
        Loop

        'Print Solid line under the last row and Totals of Balance in Hrs and Total in $
        'xTotalRow = xRow
        exSheet.Range("A" & xRow - 1 & ":P" & xRow - 1).Borders(xlEdgeBottom).LineStyle = xlSolid
        exSheet.Rows(xRow).Font.Bold = True
        exSheet.Cells(xRow, 13) = "TOTAL - CanaInc - Vacation Accrual as at " & Format(dlpDateRange(1).Text, "mmmm dd, yyyy")
        exSheet.Cells(xRow, 13).HorizontalAlignment = xlRight
        exSheet.Cells(xRow, 14).Formula = "=Round(SUM(N" & xCanaIncStart & ":N" & xRow - 1 & "),2)"
        exSheet.Cells(xRow, 15) = "hrs ="
        exSheet.Cells(xRow, 15).HorizontalAlignment = xlCenter
        exSheet.Cells(xRow, 16).Formula = "=Round(SUM(P" & xCanaIncStart & ":P" & xRow - 1 & "),2)"
        
        
        'Print the totals to Less CanaInc row. Show in Negative value
        exSheet.Cells(xCanaIncRow, 14) = 0 - Round(exSheet.Cells(xRow, 14), 2)
        exSheet.Cells(xCanaIncRow, 16) = 0 - Round(exSheet.Cells(xRow, 16), 2)
                
    End If

    rsHREmp.Close
    Set rsHREmp = Nothing
        
    'Sum the above totals in the main list
    exSheet.Cells(xCanaIncRow + 1, 14).Formula = "=Round(SUM(N" & xTotalRow & ":N" & xCanaIncRow & "),2)"
    exSheet.Cells(xCanaIncRow + 1, 15) = "hrs ="
    exSheet.Cells(xCanaIncRow + 1, 15) = "hrs ="
    exSheet.Cells(xCanaIncRow + 1, 16).Formula = "=Round(SUM(P" & xTotalRow & ":P" & xCanaIncRow & "),2)"

    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT

    Call Pause(1)
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
    
Exit Sub

SUCCESS_Accrual_XLS_Report_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", "SUCCESS_Accrual_XLS_Report", "SELECT")

Set exSheet = Nothing
Set exBook = Nothing
exApp.Quit
Set exApp = Nothing

Exit Sub

End Sub

Private Sub Export_EmpData_to_Excel_GraniteClub()
    Dim rsHREmp As New ADODB.Recordset
    Dim exApp As Object 'Excel.Application
    Dim exBook As Object 'Excel.Workbook
    Dim exSheet As Object 'Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
        
    On Error GoTo Err_Export_EmpData_to_Excel_GraniteClub
    
    
    Screen.MousePointer = HOURGLASS
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    'If OptAct.Value Then
        SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_ADDR1,ED_ADDR2,ED_CITY,ED_PROV,ED_PCODE,ED_PHONE,ED_CELLPHONE,"
        SQLQ = SQLQ & "ED_EMAIL,ED_ECONT,ED_RELATE,ED_ENBR,ED_ECELLPHONE,ED_ECONT2,ED_RELATE2,ED_ENBR2,ED_ECELLPHONE2"
        SQLQ = SQLQ & " FROM HREMP "
        SQLQ = SQLQ & " WHERE 1 = 1"
    'Else 'Term
    '    SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO,ED_ADMINBY,ED_LOC,ED_DIV,ED_SECTION,ED_DRIVERLIC,"
    '    SQLQ = SQLQ & "ED_EMP,ED_PT,ED_DOH,JH_JOB,JH_REPTAU,JH_REPTAU2 "
    '    SQLQ = SQLQ & " FROM (Term_HREMP INNER JOIN Term_JOB_HISTORY ON ED_EMPNBR=JH_EMPNBR AND JH_CURRENT <>0) "
    '    SQLQ = SQLQ & " WHERE 1 = 1"
    '    sSQLQ = Replace(sSQLQ, "HREMP.", "Term_HREMP.")
    'End If
    SQLQ = SQLQ & " AND " & sSQLQ & " "
    SQLQ = SQLQ & " ORDER BY ED_FNAME, ED_SURNAME "

    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
                
        'File to export to
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2241_EmployeeInfoTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "EmployeeInfo_" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
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
    
        exSheet.Cells(1, 1) = "Employee Name"
        exSheet.Cells(1, 2) = "Home Address"
        exSheet.Cells(1, 3) = "City and Province"
        exSheet.Cells(1, 4) = "Postal Code"
        exSheet.Cells(1, 5) = "Home Number"
        exSheet.Cells(1, 6) = "Cell Phone Number"
        exSheet.Cells(1, 7) = "Email Address"
        exSheet.Cells(1, 8) = "Emergency Contact 1"
        exSheet.Cells(1, 9) = "Relationship"
        exSheet.Cells(1, 10) = "Contact Phone Number"
        exSheet.Cells(1, 11) = "Cell Phone Number"
        exSheet.Cells(1, 12) = "Emergency Contact 2"
        exSheet.Cells(1, 13) = "Relationship"
        exSheet.Cells(1, 14) = "Contact Phone Number"
        exSheet.Cells(1, 15) = "Cell Phone Number"
        
        xRow = 2
        
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            exSheet.Cells(xRow, 1) = rsHREmp("ED_FNAME") & " " & rsHREmp("ED_SURNAME")
            If Len(Trim(rsHREmp("ED_ADDR2"))) > 0 Then
                exSheet.Cells(xRow, 2) = rsHREmp("ED_ADDR1") & " " & rsHREmp("ED_ADDR2")
            Else
                exSheet.Cells(xRow, 2) = rsHREmp("ED_ADDR1")
            End If
            exSheet.Cells(xRow, 3) = rsHREmp("ED_CITY") & ", " & rsHREmp("ED_PROV")
            exSheet.Cells(xRow, 4) = rsHREmp("ED_PCODE")
            exSheet.Cells(xRow, 5) = Format(rsHREmp("ED_PHONE"), "(###) ###-####")
            exSheet.Cells(xRow, 6) = Format(rsHREmp("ED_CELLPHONE"), "(###) ###-####")
            exSheet.Cells(xRow, 7) = rsHREmp("ED_EMAIL")
            
            exSheet.Cells(xRow, 8) = rsHREmp("ED_ECONT")
            exSheet.Cells(xRow, 9) = rsHREmp("ED_RELATE")
            If Not IsNull(rsHREmp("ED_ENBR")) Then
                If Len(Trim(Replace(Replace(rsHREmp("ED_ENBR"), "-", ""), " ", ""))) > 10 Then
                    exSheet.Cells(xRow, 10) = Format(Left(rsHREmp("ED_ENBR"), 10), "(###) ###-####") & " ext. " & Mid(rsHREmp("ED_ENBR"), 11)
                Else
                    exSheet.Cells(xRow, 10) = Format(rsHREmp("ED_ENBR"), "(###) ###-####")
                End If
            End If
            exSheet.Cells(xRow, 11) = Format(rsHREmp("ED_ECELLPHONE"), "(###) ###-####")
            
            exSheet.Cells(xRow, 12) = rsHREmp("ED_ECONT2")
            exSheet.Cells(xRow, 13) = rsHREmp("ED_RELATE2")
            If Not IsNull(rsHREmp("ED_ENBR2")) Then
                If Len(Trim(Replace(Replace(rsHREmp("ED_ENBR2"), "-", ""), " ", ""))) > 10 Then
                    exSheet.Cells(xRow, 14) = Format(Left(rsHREmp("ED_ENBR2"), 10), "(###) ###-####") & " ext. " & Mid(rsHREmp("ED_ENBR2"), 11)
                Else
                    exSheet.Cells(xRow, 14) = Format(rsHREmp("ED_ENBR2"), "(###) ###-####")
                End If
            End If
            exSheet.Cells(xRow, 15) = Format(rsHREmp("ED_ECELLPHONE2"), "(###) ###-####")
            
            rsHREmp.MoveNext
            xRow = xRow + 1
        Loop
        
        'exSheet.AutoFilterMode = True
        'exSheet.Range("K1:K20").AutoFilter Field:=1
        exSheet.Columns.AutoFit

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Err_Export_EmpData_to_Excel_GraniteClub:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")

Set exSheet = Nothing
Set exBook = Nothing
exApp.Quit
Set exApp = Nothing


End Sub

Private Sub URW_DivisionalSummaryReports()
    Dim X%
    Dim WSQLQ As String
    
    WSQLQ = glbSeleDeptUn
    If clpDiv.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DIV IN ('" & Replace(clpDiv.Text, ",", "','") & "')"
    If clpDept.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DEPTNO IN ('" & Replace(clpDept.Text, ",", "','") & "')"
    If clpCode(0).Text <> "" Then WSQLQ = WSQLQ & " AND ED_LOC IN ('" & Replace(clpCode(0).Text, ",", "','") & "')"
    If clpCode(1).Text <> "" Then WSQLQ = WSQLQ & " AND ED_ORG IN ('" & Replace(clpCode(1).Text, ",", "','") & "')"
    If clpCode(2).Text <> "" Then WSQLQ = WSQLQ & " AND ED_EMP IN ('" & Replace(clpCode(2).Text, ",", "','") & "')"
    If clpCode(3).Text <> "" Then WSQLQ = WSQLQ & " AND ED_REGION IN ('" & Replace(clpCode(3).Text, ",", "','") & "')"
    If clpCode(4).Text <> "" Then WSQLQ = WSQLQ & " AND ED_ADMINBY IN ('" & Replace(clpCode(4).Text, ",", "','") & "')"
    If clpCode(5).Text <> "" Then WSQLQ = WSQLQ & " AND ED_SECTION IN ('" & Replace(clpCode(5).Text, ",", "','") & "')"
    If clpPT.Text <> "" Then WSQLQ = WSQLQ & " AND ED_PT IN ('" & Replace(clpPT.Text, ",", "','") & "')"
       
    If elpEEID.Text <> "" Then WSQLQ = WSQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
    
    
    'HisSQLUWR = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
    HisSQLUWR = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & WSQLQ & ")"
    
    If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
        HisSQLUWR = HisSQLUWR & " AND (AD_DOA Between "
        HisSQLUWR = HisSQLUWR & Date_SQL(dlpDateRange(0)) & "And "
        HisSQLUWR = HisSQLUWR & Date_SQL(dlpDateRange(1)) & ") "
    Else
        For X% = 0 To 1
            If Len(dlpDateRange(X).Text) > 0 Then
                If X% = 0 Then
                    HisSQLUWR = HisSQLUWR & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(0).Text) & ") "
                Else
                    HisSQLUWR = HisSQLUWR & " AND (AD_DOA <= " & Date_SQL(dlpDateRange(1).Text) & " ) "
                End If
            End If
        Next X%
    End If
    
    Call SELATTWRK1
    
    glbstrSelCri = " {HRATTWRK.AD_WRKEMP}='" & glbUserID & "'"
End Sub

Private Sub SELATTWRK1()
Dim SQLQ As String
Dim USQLQ As String
Dim xFieldList As String

On Error GoTo AttWrkError
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3

Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).FloodPercent = 15

DoEvents

gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "DELETE FROM HRATTWRK WHERE AD_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001W.CommitTrans

DoEvents

MDIMain.panHelp(0).FloodPercent = 30

xFieldList = Get_Fields(gdbAdoIhr001, "HR_ATTENDANCE", "AD_ATT_ID")
SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS AD_WRKEMP "
If Not glbOracle Then
    If glbMulti Then
        SQLQ = SQLQ & " FROM HR_ATTENDANCE " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) " 'AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB "
    Else
        SQLQ = SQLQ & " FROM HR_ATTENDANCE " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
    End If
Else
    If glbMulti Then
        SQLQ = SQLQ & " FROM HR_ATTENDANCE " ', HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
    Else
        SQLQ = SQLQ & " FROM HR_ATTENDANCE" ', HR_JOB_HISTORY WHERE HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
    End If
End If

If Len(HisSQLUWR) > 1 Then
    If glbOracle Then
        SQLQ = SQLQ & " WHERE (" & HisSQLUWR & ")"
    Else
        SQLQ = SQLQ & " WHERE (" & HisSQLUWR & ")"
    End If
End If

MDIMain.panHelp(0).FloodPercent = 45

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
DoEvents

'If ReportSel = "HIS" Or chkInclHIS Then
'    MDIMain.panHelp(0).FloodPercent = 60
'
'    SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
'    SQLQ = SQLQ & in_SQL(glbIHRDBW)
'    SQLQ = SQLQ & " SELECT " & Replace(xFieldList, "AD_", "AH_") & ",'" & glbUserID & "' AS AD_WRKEMP "
'
'    If Not glbOracle Then
'        If glbMulti Then
'            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) " 'AND HR_ATTENDANCE_HISTORY.AH_JOB = HR_JOB_HISTORY.JH_JOB "
'        Else
'            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY " 'LEFT OUTER JOIN HR_JOB_HISTORY ON (HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
'        End If
'    Else
'        If glbMulti Then
'            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY " ', HR_JOB_HISTORY WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
'        Else
'            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY " ', HR_JOB_HISTORY WHERE HR_ATTENDANCE_HISTORY.AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_CURRENT <> 0 "
'        End If
'    End If
'
'    If Len(HisSQLUWR) > 1 Then
'        If glbOracle Then
'            SQLQ = SQLQ & " WHERE (" & Replace(HisSQLUWR, "AD_", "AH_") & ")"
'        Else
'            SQLQ = SQLQ & "WHERE (" & Replace(HisSQLUWR, "AD_", "AH_") & ")"
'        End If
'    End If
'
'    MDIMain.panHelp(0).FloodPercent = 65
'
'    gdbAdoIhr001.BeginTrans
'    gdbAdoIhr001.Execute SQLQ
'    gdbAdoIhr001.CommitTrans
'
'    DoEvents
'End If

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Exit Sub

AttWrkError:
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub
End Sub

Private Sub Export_Manulife_Census_Data_Hastings()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsBenfit As New ADODB.Recordset
    Dim exApp As Object 'Excel.Application
    Dim exBook As Object 'Excel.Workbook
    Dim exSheet As Object 'Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
        
    On Error GoTo Err_Export_Manulife_Census_Data_Hastings
    
    
    Screen.MousePointer = HOURGLASS
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
    SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO,ED_USER_TEXT1,ED_USER_TEXT2,JH_JOB,ED_SEX,ED_DOB,ED_DOH,SH_SALARY,"
    SQLQ = SQLQ & "ED_VACPC,ED_ORG,ED_COMBINATION "
    SQLQ = SQLQ & "FROM (HREMP LEFT JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR) "
    SQLQ = SQLQ & "LEFT JOIN HR_SALARY_HISTORY ON HR_JOB_HISTORY.JH_EMPNBR=HR_SALARY_HISTORY.SH_EMPNBR "
    SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_JOB=HR_SALARY_HISTORY.SH_JOB "
    SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_CURRENT=HR_SALARY_HISTORY.SH_CURRENT "
    
    SQLQ = SQLQ & " WHERE 1 = 1 AND JH_CURRENT <> 0 AND ED_USER_TEXT1 <> '' AND ED_USER_TEXT1 IS NOT NULL"
    SQLQ = SQLQ & " AND " & sSQLQ & " "
    If comGroup(0).ListIndex = 0 Then
        SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME"
    Else
        SQLQ = SQLQ & " ORDER BY ED_EMPNBR"
    End If

    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
                
        'File to export to
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2263ManulifeBenefitTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "ManulifeBenefit_" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            Screen.MousePointer = DEFAULT
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
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
            
        xRow = 4
        
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
                        
            exSheet.Cells(xRow, 1) = rsHREmp("ED_DEPTNO")
            exSheet.Cells(xRow, 2) = rsHREmp("ED_EMPNBR")
            exSheet.Cells(xRow, 3) = rsHREmp("ED_SURNAME")
            exSheet.Cells(xRow, 4) = rsHREmp("ED_FNAME")
            
            'Benefit info. - Policy info
            'Open Benefit record of the employee
            SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND BF_BCODE IN ('D9FF', 'D9FP', 'D9SF', 'D9SP', 'EFF', 'EFP', 'ESF','ESP')"
            SQLQ = SQLQ & " AND BF_COVER IN ('S','F')"
            rsBenfit.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsBenfit.EOF Then
                exSheet.Cells(xRow, 5) = "31917"
            End If
            rsBenfit.Close
            Set rsBenfit = Nothing
            
            'Benefit info. - Policy info
            'Open Benefit record of the employee
            SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND BF_BCODE IN ('LIFF', 'LIFP', 'LTDF', 'LTDP') AND BF_COVER = 'Y'"
            rsBenfit.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsBenfit.EOF Then
                exSheet.Cells(xRow, 6) = "31939"
            End If
            rsBenfit.Close
            Set rsBenfit = Nothing
                        
            exSheet.Cells(xRow, 7) = rsHREmp("ED_USER_TEXT1")
            exSheet.Cells(xRow, 8) = rsHREmp("ED_USER_TEXT2")
            exSheet.Cells(xRow, 9) = "ON"
            exSheet.Cells(xRow, 10) = GetJobDesc(rsHREmp("JH_JOB"))
            exSheet.Cells(xRow, 11) = rsHREmp("ED_SEX")
            exSheet.Cells(xRow, 12) = Format(rsHREmp("ED_DOB"), "mm/dd/yyyy")
            exSheet.Cells(xRow, 13) = Format(rsHREmp("ED_DOH"), "mm/dd/yyyy")
            exSheet.Cells(xRow, 14) = rsHREmp("SH_SALARY")
            exSheet.Cells(xRow, 15) = rsHREmp("ED_COMBINATION")
            
            'Benefit info. - Employee Life
            'Open Benefit record of the employee
            SQLQ = "SELECT BF_AMT FROM HRBENFT WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND BF_BCODE IN ('LIFF', 'LIFP') AND BF_COVER = 'Y'"
            rsBenfit.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsBenfit.EOF Then
                exSheet.Cells(xRow, 16) = rsBenfit("BF_AMT")
            End If
            rsBenfit.Close
            Set rsBenfit = Nothing
            
            'Benefit info. - Employee Optional Life
            'Open Benefit record of the employee
            SQLQ = "SELECT BF_AMT FROM HRBENFT WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND BF_BCODE IN ('OPTL') AND BF_COVER = 'Y'"
            rsBenfit.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsBenfit.EOF Then
                exSheet.Cells(xRow, 17) = rsBenfit("BF_AMT")
            End If
            rsBenfit.Close
            Set rsBenfit = Nothing
            
            'Benefit info. - Employee Basic AD&D (Same as col. P - LIFE)
            'Open Benefit record of the employee
            SQLQ = "SELECT BF_AMT FROM HRBENFT WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND BF_BCODE IN ('LIFF', 'LIFP') AND BF_COVER = 'Y'"
            rsBenfit.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsBenfit.EOF Then
                exSheet.Cells(xRow, 18) = rsBenfit("BF_AMT")
            End If
            rsBenfit.Close
            Set rsBenfit = Nothing
            
            'Benefit info. - Long Term Disability
            'Open Benefit record of the employee
            SQLQ = "SELECT BF_AMT FROM HRBENFT WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND BF_BCODE IN ('LTDF','LTDP') AND BF_COVER = 'Y'"
            rsBenfit.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsBenfit.EOF Then
                exSheet.Cells(xRow, 19) = rsBenfit("BF_AMT")
            End If
            rsBenfit.Close
            Set rsBenfit = Nothing
            
            'Benefit info. - Health Coverage
            'Open Benefit record of the employee
            SQLQ = "SELECT BF_COVER FROM HRBENFT WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND BF_BCODE IN ('D9FF', 'D9FP', 'D9SF', 'D9SP')"
            SQLQ = SQLQ & " AND BF_COVER IN ('S','F')"
            rsBenfit.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsBenfit.EOF Then
                exSheet.Cells(xRow, 20) = rsBenfit("BF_COVER")
            End If
            rsBenfit.Close
            Set rsBenfit = Nothing
            
            'Benefit info. - Dental Coverage
            'Open Benefit record of the employee
            SQLQ = "SELECT BF_COVER FROM HRBENFT WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND BF_BCODE IN ('EFF', 'EFP', 'ESF','ESP')"
            SQLQ = SQLQ & " AND BF_COVER IN ('S','F')"
            rsBenfit.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsBenfit.EOF Then
                exSheet.Cells(xRow, 21) = rsBenfit("BF_COVER")
            End If
            rsBenfit.Close
            Set rsBenfit = Nothing
            
                        
            exSheet.Cells(xRow, 22) = "2000.00" 'Hard coded for everyone
            
            'Benefit info. - % in Lieu
            'Open Benefit record of the employee
            SQLQ = "SELECT BF_PCC FROM HRBENFT WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND BF_BCODE IN ('%IL') AND BF_COVER = 'Y'"
            rsBenfit.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsBenfit.EOF Then
                exSheet.Cells(xRow, 23) = rsBenfit("BF_PCC")
            Else
                exSheet.Cells(xRow, 23) = "n/a"
            End If
            rsBenfit.Close
            Set rsBenfit = Nothing
            
            'Benefit info. - EI Rebate
            'Open Benefit record of the employee
            SQLQ = "SELECT BF_COVER FROM HRBENFT WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND BF_BCODE IN ('EI')"
            rsBenfit.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsBenfit.EOF Then
                If rsBenfit("BF_COVER") = "Y" Then
                    exSheet.Cells(xRow, 24) = "Reduced"
                ElseIf rsBenfit("BF_COVER") = "N" Then
                    exSheet.Cells(xRow, 24) = "Normal"
                End If
            End If
            rsBenfit.Close
            Set rsBenfit = Nothing
                        
            exSheet.Cells(xRow, 25) = rsHREmp("ED_VACPC") * 100
            
            'Benefit info. - OMERS Member
            'Open Benefit record of the employee
            SQLQ = "SELECT BF_COVER FROM HRBENFT WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND BF_BCODE IN ('OMER')"
            rsBenfit.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsBenfit.EOF Then
                If rsBenfit("BF_COVER") = "Y" Then
                    exSheet.Cells(xRow, 26) = "Yes"
                ElseIf rsBenfit("BF_COVER") = "N" Then
                    exSheet.Cells(xRow, 26) = "No"
                End If
            End If
            rsBenfit.Close
            Set rsBenfit = Nothing
            
            'Benefit info. - CSB Participant
            'Open Benefit record of the employee
            SQLQ = "SELECT BF_PPAMT FROM HRBENFT WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND BF_BCODE IN ('CSB') AND BF_COVER = 'Y'"
            rsBenfit.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsBenfit.EOF Then
                exSheet.Cells(xRow, 27) = rsBenfit("BF_PPAMT")
            End If
            rsBenfit.Close
            Set rsBenfit = Nothing
            
            exSheet.Cells(xRow, 28) = rsHREmp("ED_ORG")
            
Next_Emp:
            rsHREmp.MoveNext
            xRow = xRow + 1
        Loop

        'exSheet.Columns.AutoFit

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Err_Export_Manulife_Census_Data_Hastings:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "Hastings")

Set exSheet = Nothing
Set exBook = Nothing
exApp.Quit
Set exApp = Nothing
End Sub

Private Sub Goodmans_Benefit_ExcelReport()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsBenfit As New ADODB.Recordset
    Dim exApp As Object
    Dim exBook As Object
    Dim exSheet As Object
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    Dim xTotCompCost As Double
    Dim xTotEmpCost As Double
        
    On Error GoTo Err_Goodmans_Benefit_ExcelReport
    
    
    Screen.MousePointer = HOURGLASS
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
    SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_USER_TEXT1,ED_USER_NUM1,ED_USER_NUM2 "
    SQLQ = SQLQ & "FROM HREMP WHERE 1 = 1 "
    SQLQ = SQLQ & " AND " & sSQLQ & " "
    
    'If comGroup(0).ListIndex = 0 Then
        SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME, ED_USER_TEXT1"
    'Else
    '    SQLQ = SQLQ & " ORDER BY ED_EMPNBR, ED_USER_TEXT1"
    'End If

    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
                
        'File to export to
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2290BenefitRptTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "BenefitRpt_" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            Screen.MousePointer = DEFAULT
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
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
            
        xRow = 3
        
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            'Benefit info.
            'Open Benefit record of the employee
            SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " ORDER BY BF_BCODE"
            rsBenfit.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsBenfit.EOF Then
                Do While Not rsBenfit.EOF
                     
                    exSheet.Cells(xRow, 1) = rsHREmp("ED_EMPNBR")
                    exSheet.Cells(xRow, 2) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
                    exSheet.Cells(xRow, 3) = rsHREmp("ED_USER_NUM1")
                    exSheet.Cells(xRow, 4) = rsHREmp("ED_USER_NUM2")
                    exSheet.Cells(xRow, 5) = rsHREmp("ED_USER_TEXT1")
                        
                    exSheet.Cells(xRow, 6) = GetTABLDesc("BNCD", rsBenfit("BF_BCODE"))
                    exSheet.Cells(xRow, 7) = rsBenfit("BF_COVER")
                    exSheet.Cells(xRow, 8) = rsBenfit("BF_AMT")
                    exSheet.Cells(xRow, 9) = rsBenfit("BF_EDATE")
                    
                    'Ticket #25855 Franks 08/14/2014
                    If Not IsNull(rsBenfit("BF_CEASEDATE")) Then
                        exSheet.Cells(xRow, 10) = rsBenfit("BF_CEASEDATE")
                    End If
                    
                    exSheet.Cells(xRow, 11) = rsBenfit("BF_MTHECOST")  'Monthly Employee Cost
                    exSheet.Cells(xRow, 12) = rsBenfit("BF_MTHCCOST")  'Monthly Employer Cost
                    
                            
                    ''Cost (Next Tax) - J+K : (Monthly Employee + Monthly Company)
                    ''exSheet.Cells(xRow, 12) = (exSheet.Cells(xRow, 10) + exSheet.Cells(xRow, 11))
                    'exSheet.Cells(xRow, 12) = "=Round(SUM(J" & xRow & ":K" & xRow & "),2)"
                    exSheet.Cells(xRow, 13) = "=Round(SUM(K" & xRow & ":L" & xRow & "),2)"
                    
                    ''Tax 8% - L x 8%: (Cost (Net Tax) * 8%)
                    ''exSheet.Cells(xRow, 13) = (exSheet.Cells(xRow, 12) * 0.08)
                    'exSheet.Cells(xRow, 13) = "=Round(SUM(L" & xRow & " * 0.08),2)"
                    exSheet.Cells(xRow, 14) = "=Round(SUM(M" & xRow & " * 0.08),2)"
                    
                    ''Cost (Incl Tax) - L+M : (Cost (Net Tax) + Tax 8%)
                    ''exSheet.Cells(xRow, 14) = (exSheet.Cells(xRow, 12) + exSheet.Cells(xRow, 13))
                    'exSheet.Cells(xRow, 14) = "=Round(SUM(L" & xRow & ":M" & xRow & "),2)"
                    exSheet.Cells(xRow, 15) = "=Round(SUM(M" & xRow & ":N" & xRow & "),2)"
                    

                    
                    'Borders
                    'exSheet.Range("A" & xRow & ":N" & xRow).Borders(xlEdgeTop).Weight = xlThin
                    'exSheet.Range("A" & xRow & ":N" & xRow).Borders(xlEdgeBottom).Weight = xlThin
                    'exSheet.Range("A" & xRow & ":A" & xRow).Borders(xlEdgeLeft).Weight = xlThin
                    'exSheet.Range("N" & xRow & ":N" & xRow).Borders(xlEdgeRight).Weight = xlThin
                    'exSheet.Range("A" & xRow & ":N" & xRow).Borders(xlEdgeBottom).Weight = xlThin
                    exSheet.Range("A" & xRow & ":O" & xRow).Borders(xlEdgeTop).Weight = xlThin
                    exSheet.Range("A" & xRow & ":O" & xRow).Borders(xlEdgeBottom).Weight = xlThin
                    exSheet.Range("A" & xRow & ":A" & xRow).Borders(xlEdgeLeft).Weight = xlThin
                    exSheet.Range("O" & xRow & ":O" & xRow).Borders(xlEdgeRight).Weight = xlThin
                    exSheet.Range("A" & xRow & ":O" & xRow).Borders(xlEdgeBottom).Weight = xlThin
                    
                    rsBenfit.MoveNext
                    
                    xRow = xRow + 1
                Loop
            End If
            rsBenfit.Close
            Set rsBenfit = Nothing
            
Next_Emp:
            rsHREmp.MoveNext
        Loop

        'exSheet.Columns.AutoFit

        'Border around the columns
        'exSheet.Range("A3:N" & xRow - 1).Borders(xlOutline).Weight = xlThin
        exSheet.Range("A3:O" & xRow - 1).Borders(xlOutline).Weight = xlThin

        'Sum the above totals in the main list
        'exSheet.Range("A" & xRow - 1 & ":N" & xRow - 1).Borders(xlEdgeBottom).LineStyle = xlSolid
        exSheet.Range("A" & xRow - 1 & ":O" & xRow - 1).Borders(xlEdgeBottom).LineStyle = xlSolid
        exSheet.Rows(xRow + 1).Font.Bold = True
        exSheet.Cells(xRow + 1, 9 + 1) = "Grand Totals:"
        'exSheet.Cells(xRow + 1, 10 ).Formula = "=Round(SUM(J3" & ":J" & xRow - 1 & "),2)"
        'exSheet.Cells(xRow + 1, 11 ).Formula = "=Round(SUM(K3" & ":K" & xRow - 1 & "),2)"
        'exSheet.Cells(xRow + 1, 12 ).Formula = "=Round(SUM(L3" & ":L" & xRow - 1 & "),2)"
        'exSheet.Cells(xRow + 1, 13 ).Formula = "=Round(SUM(M3" & ":M" & xRow - 1 & "),2)"
        'exSheet.Cells(xRow + 1, 14 ).Formula = "=Round(SUM(N3" & ":N" & xRow - 1 & "),2)"
        exSheet.Cells(xRow + 1, 10 + 1).Formula = "=Round(SUM(K3" & ":K" & xRow - 1 & "),2)"
        exSheet.Cells(xRow + 1, 11 + 1).Formula = "=Round(SUM(L3" & ":L" & xRow - 1 & "),2)"
        exSheet.Cells(xRow + 1, 12 + 1).Formula = "=Round(SUM(M3" & ":M" & xRow - 1 & "),2)"
        exSheet.Cells(xRow + 1, 13 + 1).Formula = "=Round(SUM(N3" & ":N" & xRow - 1 & "),2)"
        exSheet.Cells(xRow + 1, 14 + 1).Formula = "=Round(SUM(O3" & ":O" & xRow - 1 & "),2)"
        
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Err_Goodmans_Benefit_ExcelReport:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "GoodmansBenefit")

Set exSheet = Nothing
Set exBook = Nothing
exApp.Quit
Set exApp = Nothing

End Sub

Private Sub Showa_Attendance_SignIn_Form()
    Dim rsHREmp As New ADODB.Recordset
    Dim SQLQ As String
    Dim sSQLQ As String
    
    Dim exApp As Object
    Dim exBook As Object
    Dim exSheet As Object
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    
    On Error GoTo Showa_Attendance_SignIn_Form_Err
    
    Screen.MousePointer = HOURGLASS
    
    'Ticket #22034 - Get Excel reports path
    Dim xExcelRptPath  As String
    If gsTRAININGMATRIX Then
        xExcelRptPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xExcelRptPath) = 0 Then
        xExcelRptPath = glbIHRREPORTS
    End If
    
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
    SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO "
    SQLQ = SQLQ & "FROM HREMP WHERE 1 = 1 "
    SQLQ = SQLQ & " AND " & sSQLQ & " "
    
    SQLQ = SQLQ & " ORDER BY ED_FNAME, ED_SURNAME"

    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
                
        'File to export to
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "AttSignInTmp.xls"
        xlsFileMat = xExcelRptPath & IIf(Right(xExcelRptPath, 1) = "\", "", "\") & "AttSignInForm_" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            Screen.MousePointer = DEFAULT
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
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
            
        xRow = 4
        
        exSheet.Cells(1, 2) = "Date: " & Format(Now, "mmm dd, yyyy hh:mm")
        exSheet.Cells(1, 4) = "Time: " & Format(Now, "hh:mm")
        
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            'Print Name and Department
            exSheet.Cells(xRow, 1) = rsHREmp("ED_FNAME") & " " & rsHREmp("ED_SURNAME")
            exSheet.Cells(xRow, 2) = GetDeptName(rsHREmp("ED_DEPTNO"), "DF_NAME")
            
            'Borders
            exSheet.Range("A" & xRow & ":F" & xRow).Borders(xlEdgeTop).Weight = xlThin
            exSheet.Range("A" & xRow & ":F" & xRow).Borders(xlEdgeBottom).Weight = xlThin
            exSheet.Range("A" & xRow & ":F" & xRow).Borders(xlEdgeLeft).Weight = xlThin
            exSheet.Range("N" & xRow & ":F" & xRow).Borders(xlEdgeRight).Weight = xlThin
            exSheet.Range("A" & xRow & ":F" & xRow).Borders(xlEdgeBottom).Weight = xlThin
            
            xRow = xRow + 1
Next_Emp:
            rsHREmp.MoveNext
        Loop

        'exSheet.Columns.AutoFit

        'Border around the columns
        exSheet.Range("A3:F" & xRow - 1).Borders(xlOutline).Weight = xlThick


        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Showa_Attendance_SignIn_Form_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", "HREMP", "Showa_Attendance_SignIn_Form")

Set exSheet = Nothing
Set exBook = Nothing
exApp.Quit
Set exApp = Nothing
    
End Sub

Private Sub Showa_ATT_Discipline_Form()
    Dim rsCounsel As New ADODB.Recordset
    Dim rsCounselHist As New ADODB.Recordset
    Dim SQLQ As String
    Dim sSQLQ As String

    Dim exApp As Object
    Dim exBook As Object
    Dim exSheet As Object
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    Dim lstEmpNo
    Dim appVerInt As Double
    Dim Total As Integer
    Dim xCounselHist As String
    Dim xCounselHistCount As Integer

    On Error GoTo Showa_ATT_Discipline_Form_Err
    
    Screen.MousePointer = HOURGLASS
    
    'Ticket #22034 - Get Excel reports path
    Dim xExcelRptPath  As String
    If gsTRAININGMATRIX Then
        xExcelRptPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xExcelRptPath) = 0 Then
        xExcelRptPath = glbIHRREPORTS
    End If
    
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
    SQLQ = "SELECT CL_EMPNBR FROM HR_COUNSEL"
    SQLQ = SQLQ & " WHERE CL_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE 1 = 1 AND " & sSQLQ & ")"
    SQLQ = SQLQ & " GROUP BY CL_EMPNBR"
    
    'Call WriteFile("SQL1=" & SQLQ)
        
    Total = 0
    lstEmpNo = 0
    
    rsCounsel.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsCounsel.EOF Then
        rsCounsel.MoveFirst
        totNum = rsCounsel.RecordCount: I = 0
                
        xRow = 4
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
                
        Do While Not rsCounsel.EOF
            
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            'File to export to
            xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2454ATTDiscTmp.xls"
            xlsFileMat = xExcelRptPath & IIf(Right(xExcelRptPath, 1) = "\", "", "\") & "ATTDiscForm_" & Trim(glbUserID) & "_" & rsCounsel("CL_EMPNBR") & ".xls"
        
            If Dir(xlsFileTmp) = "" Then
                Screen.MousePointer = DEFAULT
                MsgBox "There is no " & xlsFileTmp
                Exit Sub
            End If
            If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
                    
            FileCopy xlsFileTmp, xlsFileMat
            
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
            
            xCounselHist = ""
            xCounselHistCount = 0
            'Retrieve the Counsel History of the employee
            SQLQ = "SELECT * FROM HR_COUNSEL"
            SQLQ = SQLQ & " WHERE CL_EMPNBR = " & rsCounsel("CL_EMPNBR")
            SQLQ = SQLQ & " AND CL_TYPE = '" & clpCode(4).Text & "'"
            SQLQ = SQLQ & " AND CL_REASON = '" & clpCode(3).Text & "'"
            SQLQ = SQLQ & " ORDER BY CL_EMPNBR, CL_TYPE, CL_REASON, CL_LEVEL DESC"
            rsCounselHist.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsCounselHist.EOF Then
                '# of Counsel History
                xCounselHistCount = rsCounsel.RecordCount
                
                Do While Not rsCounselHist.EOF
                    'Print field values
                    exSheet.Cells(1, 3) = Format(Now, "mm/dd/yyyy")
                    exSheet.Cells(1, 10) = Format(rsCounsel("CL_INCDATE"), "mm/dd/yyyy")
                    exSheet.Cells(2, 3) = GetEmpData(rsCounsel("CL_EMPNBR"), "ED_FNAME") & " " & GetEmpData(rsCounsel("CL_EMPNBR"), "ED_SURNAME")
                    exSheet.Cells(2, 10) = rsCounsel("CL_EMPNBR")
                                                
                    If rsCounselHist("CL_LEVEL") = 1 Then
                        exSheet.Range("A5:B5").Interior.Color = vbGrayed
                    ElseIf rsCounselHist("CL_LEVEL") = 2 Then
                        exSheet.Range("D5:E5").Interior.Color = vbGrayed
                    ElseIf rsCounselHist("CL_LEVEL") = 3 Then
                        exSheet.Range("G5:H5").Interior.Color = vbGrayed
                    ElseIf rsCounselHist("CL_LEVEL") = 4 Then
                        exSheet.Range("J5:K5").Interior.Color = vbGrayed
                    End If
                    
                    'Previously higher level Incident Date and Level # - C3 and J3
                    If xCounselHistCount = 1 Then
                        exSheet.Cells(3, 3) = Format(rsCounselHist("CL_INCDATE"), "mm/dd/yyyy")
                        exSheet.Cells(3, 10) = rsCounselHist("CL_LEVEL")
                    Else
                        'Check the Incident Date of next highest level
                        rsCounselHist.MoveNext
                        exSheet.Cells(3, 3) = Format(rsCounselHist("CL_INCDATE"), "mm/dd/yyyy")
                        exSheet.Cells(3, 10) = rsCounselHist("CL_LEVEL")
                        
                        'Move back to the previous record - highest level
                        rsCounselHist.MovePrevious
                    End If
                    
                    'Retrieve Counsel History
                    Do While Not rsCounselHist.EOF
                        xCounselHist = "Level: " & rsCounselHist("CL_LEVEL") & ", Incident Date: " & Format(rsCounselHist("CL_INCDATE"), "mm/dd/yyyy") & "; " & xCounselHist
                        
                        rsCounselHist.MoveNext
                    Loop
                    exSheet.Cells(16, 1) = xCounselHist
                    
                    
                    exSheet.Cells(20, 3) = GetEmpData(rsCounsel("CL_EMPNBR"), "ED_FNAME") & " " & GetEmpData(rsCounsel("CL_EMPNBR"), "ED_SURNAME")
                    
                    'Reporting Authority
                    exSheet.Cells(22, 3) = GetEmpData(GetJHData(rsCounsel("CL_EMPNBR"), "JH_REPTAU", 0), "ED_FNAME") & " " & GetEmpData(GetJHData(rsCounsel("CL_EMPNBR"), "JH_REPTAU", 0), "ED_SURNAME")
                    exSheet.Cells(23, 3) = GetEmpData(GetJHData(rsCounsel("CL_EMPNBR"), "JH_REPTAU2", 0), "ED_FNAME") & " " & GetEmpData(GetJHData(rsCounsel("CL_EMPNBR"), "JH_REPTAU2", 0), "ED_SURNAME")
                
                    rsCounselHist.MoveLast
                Loop
            End If
            rsCounselHist.Close
            Set rsCounselHist = Nothing
            
            xRow = xRow + 1
Next_Emp:
            rsCounsel.MoveNext

            'exSheet.Columns.AutoFit
    
            exBook.Save
            Set exSheet = Nothing
            Set exBook = Nothing
            exApp.Quit
            Set exApp = Nothing
        Loop

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    
    rsCounsel.Close
    Set rsCounsel = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Showa_ATT_Discipline_Form_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", "HR_COUNSEL", "Showa_ATT_Discipline_Form")

Set exSheet = Nothing
Set exBook = Nothing
exApp.Quit
Set exApp = Nothing

End Sub

Private Sub Showa_COC_Discipline_Form()
    Dim rsCounsel As New ADODB.Recordset
    Dim rsCounselHist As New ADODB.Recordset
    Dim SQLQ As String
    Dim sSQLQ As String

    Dim exApp As Object
    Dim exBook As Object
    Dim exSheet As Object
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    Dim lstEmpNo
    Dim appVerInt As Double
    Dim Total As Integer
    Dim xCounselHist As String
    Dim xCounselHistCount As Integer

    On Error GoTo Showa_COC_Discipline_Form_Err
    
    Screen.MousePointer = HOURGLASS
    
    'Ticket #22034 - Get Excel reports path
    Dim xExcelRptPath  As String
    If gsTRAININGMATRIX Then
        xExcelRptPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xExcelRptPath) = 0 Then
        xExcelRptPath = glbIHRREPORTS
    End If
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
    SQLQ = "SELECT CL_EMPNBR FROM HR_COUNSEL"
    SQLQ = SQLQ & " WHERE CL_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE 1 = 1 AND " & sSQLQ & ")"
    SQLQ = SQLQ & " GROUP BY CL_EMPNBR"
    
    'Call WriteFile("SQL1=" & SQLQ)
        
    Total = 0
    lstEmpNo = 0
    
    rsCounsel.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsCounsel.EOF Then
        rsCounsel.MoveFirst
        totNum = rsCounsel.RecordCount: I = 0
                
        xRow = 4
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
                
        Do While Not rsCounsel.EOF
            
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            'File to export to
            xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2454COCDiscTmp.xls"
            xlsFileMat = xExcelRptPath & IIf(Right(xExcelRptPath, 1) = "\", "", "\") & "COCDiscForm_" & Trim(glbUserID) & "_" & rsCounsel("CL_EMPNBR") & ".xls"
        
            If Dir(xlsFileTmp) = "" Then
                Screen.MousePointer = DEFAULT
                MsgBox "There is no " & xlsFileTmp
                Exit Sub
            End If
            If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
                    
            FileCopy xlsFileTmp, xlsFileMat
            
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
            
            xCounselHist = ""
            xCounselHistCount = 0
            'Retrieve the Counsel History of the employee
            SQLQ = "SELECT * FROM HR_COUNSEL"
            SQLQ = SQLQ & " WHERE CL_EMPNBR = " & rsCounsel("CL_EMPNBR")
            SQLQ = SQLQ & " AND CL_TYPE = '" & clpCode(4).Text & "'"
            SQLQ = SQLQ & " AND CL_REASON = '" & clpCode(3).Text & "'"
            SQLQ = SQLQ & " ORDER BY CL_EMPNBR, CL_TYPE, CL_REASON, CL_LEVEL DESC"
            rsCounselHist.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsCounselHist.EOF Then
                '# of Counsel History
                xCounselHistCount = rsCounsel.RecordCount
                
                Do While Not rsCounselHist.EOF
                    'Print field values
                    exSheet.Cells(1, 3) = Format(Now, "mm/dd/yyyy")
                    exSheet.Cells(1, 5) = rsCounsel("CL_EMPNBR")
                    exSheet.Cells(1, 7) = Format(GetEmpData(rsCounsel("CL_EMPNBR"), "ED_DOH"), "mm/dd/yyyy")
                    exSheet.Cells(2, 3) = GetEmpData(rsCounsel("CL_EMPNBR"), "ED_FNAME") & " " & GetEmpData(rsCounsel("CL_EMPNBR"), "ED_SURNAME")
                    exSheet.Cells(2, 7) = Format(rsCounsel("CL_INCDATE"), "mm/dd/yyyy")
                                                                
                    If rsCounselHist("CL_LEVEL") = 1 Then
                        exSheet.Range("B4:B4").Interior.Color = vbGrayed
                    ElseIf rsCounselHist("CL_LEVEL") = 2 Then
                        exSheet.Range("C4:C4").Interior.Color = vbGrayed
                    ElseIf rsCounselHist("CL_LEVEL") = 3 Then
                        exSheet.Range("D4:D4").Interior.Color = vbGrayed
                    ElseIf rsCounselHist("CL_LEVEL") = 4 Then
                        exSheet.Range("E4:E4").Interior.Color = vbGrayed
                    ElseIf rsCounselHist("CL_LEVEL") = 5 Then
                        exSheet.Range("F4:F4").Interior.Color = vbGrayed
                    ElseIf rsCounselHist("CL_LEVEL") = 6 Then
                        exSheet.Range("G4:G4").Interior.Color = vbGrayed
                    End If
                    
                    'Stream
                    exSheet.Cells(11, 1) = Get_StreamDesc(rsCounselHist("CL_STREAM"))
                    
                    'Previously higher level Incident Date and Level # - C3 and J3
                    If xCounselHistCount = 1 Then
                        exSheet.Cells(3, 3) = Format(rsCounselHist("CL_INCDATE"), "mm/dd/yyyy")
                        exSheet.Cells(3, 6) = rsCounselHist("CL_LEVEL")
                    Else
                        'Check the Incident Date of next highest level
                        rsCounselHist.MoveNext
                        exSheet.Cells(3, 3) = Format(rsCounselHist("CL_INCDATE"), "mm/dd/yyyy")
                        exSheet.Cells(3, 6) = rsCounselHist("CL_LEVEL")
                        
                        'Move back to the previous record - highest level
                        rsCounselHist.MovePrevious
                    End If
                    
                    'Retrieve Counsel History
                    Do While Not rsCounselHist.EOF
                        xCounselHist = "Level: " & rsCounselHist("CL_LEVEL") & ", Incident Date: " & Format(rsCounselHist("CL_INCDATE"), "mm/dd/yyyy") & "; " & xCounselHist
                        
                        rsCounselHist.MoveNext
                    Loop
                    exSheet.Cells(16, 1) = xCounselHist
                    
                    
                    exSheet.Cells(21, 3) = GetEmpData(rsCounsel("CL_EMPNBR"), "ED_FNAME") & " " & GetEmpData(rsCounsel("CL_EMPNBR"), "ED_SURNAME")
                    
                    'Reporting Authority
                    exSheet.Cells(23, 3) = GetEmpData(GetJHData(rsCounsel("CL_EMPNBR"), "JH_REPTAU", 0), "ED_FNAME") & " " & GetEmpData(GetJHData(rsCounsel("CL_EMPNBR"), "JH_REPTAU", 0), "ED_SURNAME")
                    exSheet.Cells(24, 3) = GetEmpData(GetJHData(rsCounsel("CL_EMPNBR"), "JH_REPTAU2", 0), "ED_FNAME") & " " & GetEmpData(GetJHData(rsCounsel("CL_EMPNBR"), "JH_REPTAU2", 0), "ED_SURNAME")
                
                    rsCounselHist.MoveLast
                Loop
            End If
            rsCounselHist.Close
            Set rsCounselHist = Nothing
            
            xRow = xRow + 1
Next_Emp:
            rsCounsel.MoveNext

            'exSheet.Columns.AutoFit
    
            exBook.Save
            Set exSheet = Nothing
            Set exBook = Nothing
            exApp.Quit
            Set exApp = Nothing
        Loop

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    
    rsCounsel.Close
    Set rsCounsel = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Showa_COC_Discipline_Form_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", "HR_COUNSEL", "Showa_COC_Discipline_Form")

Set exSheet = Nothing
Set exBook = Nothing
exApp.Quit
Set exApp = Nothing

End Sub

Private Function Get_StreamDesc(xStream) As String
    Dim xStreamDesc As String
    
    xStreamDesc = ""
    
    Select Case xStream
        Case 1: xStreamDesc = "1 - Fighting with or attempting to injure another Associate."
        Case 2: xStreamDesc = "2 - Possessing any weapon or related paraphernalia on SCI premises."
        Case 3: xStreamDesc = "3 - Creating an intimidating, hostile, or offensive working environment by threatening, coercing, retaliating  against or using abusive/threatening language towards Associates or visitors on SCI property."
        Case 4: xStreamDesc = "4 - Failing to immediately report a workplace illness, injury, incident or hazard to the Department Supervisor,  Manager or Health & Safety Specialist."
        Case 5: xStreamDesc = "5 - Failing to observe established safety rules including but not limited to, not wearing appropriate PPE as per operation standards."
        Case 6: xStreamDesc = "6 - Smoking or creating a flame outside the designated smoking area without authorization."
        Case 7: xStreamDesc = "7 - Engaging in illegal gambling, horseplay or practical jokes, including but not limited to any prank, contest, feat of strength, unnecessary running or rough and boisterous conduct or games of chance."
        Case 8: xStreamDesc = "8 - Possessing, using/misusing, distributing, selling, purchasing, reporting to work under the influence of, any illegal or prescription drugs/substances or alcohol while on SCI property."
        Case 9: xStreamDesc = "9 - Stealing, willfully damaging, abusing or hiding any property belonging to another Associates or SCI."
        Case 10: xStreamDesc = "10 - Giving false information with respect to absence, sickness, injury claims or personnel files."
        Case 11: xStreamDesc = "11 - Misrepresenting facts or falsifying records or reports, of a business nature."
        Case 12: xStreamDesc = "12 - Inappropriately obtaining or sharing confidential information with anyone outside of the appropriate SCI department."
        Case 13: xStreamDesc = "13 - Interfering, failing to cooperate, or divulging confidential information related to an authorized SCI investigation."
        Case 14: xStreamDesc = "14 - Obtaining property, money, or other privileges from SCI through fraud or misrepresentation or engaging in this type of activity while conducting SCI business."
        Case 15: xStreamDesc = "15 - Willfully scanning or requesting the scanning of yours or another Associate’s identification card."
        Case 16: xStreamDesc = "16 - Interfering with the work or performance of another associate or causing a restriction or slow down of production or tampering with or deliberately misusing company equipment."
        Case 17: xStreamDesc = "17 - Using electronic equipment, cameras, video equipment, or recording devices, without proper authorization or inappropriately while on SCI premises."
        Case 18: xStreamDesc = "18 - Using or carrying personal cell phones, personal paging devices or other electronic equipment on the floor unless otherwise authorized."
        Case 19: xStreamDesc = "19 - Reporting to the work area without wearing the proper SCI uniform including weekends, holidays and shutdowns."
        Case 20: xStreamDesc = "20 - Being absent from work for an unauthorized reason (any incident in which the associate fails to declare the absence as Personal Emergency Leave or the absence does not meet the qualifications under Personal Emergency Leave provisions)."
        Case 21: xStreamDesc = "21 - Failing to contact SCI to report an absence within the scheduled shift (No call / No show)."
        Case 22: xStreamDesc = "22 - Providing late notification of an absence (notification of less than 1 hour prior to shift start)."
        Case 23: xStreamDesc = "23 - Failing to swipe in prior to the scheduled shift start time (late)."
        Case 24: xStreamDesc = "24 - Leaving work prior to the end of the scheduled shift (Early Departure) without declaring it an Emergency Day."
        Case 25: xStreamDesc = "25 - Failing to provide appropriate medical documentation following an absence of 3 or more consecutive days."
        Case 26: xStreamDesc = "26 - Failing to complete the No Swipe Form and receive the Supervisor’s authorization (prior to the shift start meeting or at the end of shift). (Associates will be paid for the earliest time confirmed present at work.)"
        Case 27: xStreamDesc = "27 - Reporting to the work area late and or leaving early at shift start up or after breaks and lunch or prior to shift end."
        Case 28: xStreamDesc = "28 - Leaving the work area during assigned working hours without permission."
        Case 29: xStreamDesc = "29 - Stopping work prior to the shift end."
        Case 30: xStreamDesc = "30 - Sleeping or loafing around while on the job."
        Case 31: xStreamDesc = "31 - Working below the operation standard for quality or quantity, willfully neglecting job responsibilities or refusing to comply with instructions of management."
        Case 32: xStreamDesc = "32 - Removing, defacing, or changing notices or bulletins posted throughout the facility or placing signs, notes, papers, or any materials not required for production in or on products or property."
        Case 33: xStreamDesc = "33 - Bringing or removing tools, materials or equipment to or from SCI premises without proper authorization."
        Case 34: xStreamDesc = "34 - Duplicating any SCI keys without the proper authorization."
        Case 35: xStreamDesc = "35 - Storing or posting inappropriate materials other than personal items and uniforms in SCI issued lockers."
        Case 36: xStreamDesc = "36 - Parking inappropriately or outside the designated Associate parking area during working hours."
    End Select
    
    Get_StreamDesc = xStreamDesc
End Function

Private Sub Export_Data_to_Excel_SMDHU() 'Simcoe Muskoka District Health Unit - Ticket #26368 Franks 02/12/2015
    Dim rsHREmp As New ADODB.Recordset
    Dim rsSal As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    Dim xAnnSal
    
    On Error GoTo Err_Export_Data_to_Excel_CollectCorp
    
    
    Screen.MousePointer = HOURGLASS
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    If OptAct.Value Then
        SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO,ED_ADMINBY,ED_LOC,ED_ORG,ED_DIV,ED_SECTION,ED_DRIVERLIC,"
        SQLQ = SQLQ & "ED_EMP,ED_PT,ED_DOH,ED_GLNO,ED_SENDTE,JH_JOB,JH_REPTAU,JH_REPTAU2,JH_FTENUM "
        SQLQ = SQLQ & " FROM (HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=JH_EMPNBR AND JH_CURRENT <>0) "
        SQLQ = SQLQ & " WHERE 1 = 1"
    Else 'Term
        SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO,ED_ADMINBY,ED_LOC,ED_ORG,ED_DIV,ED_SECTION,ED_DRIVERLIC,"
        SQLQ = SQLQ & "ED_EMP,ED_PT,ED_DOH,ED_GLNO,ED_SENDTE,JH_JOB,JH_REPTAU,JH_REPTAU2,JH_FTENUM,Term_HREMP.TERM_SEQ "
        SQLQ = SQLQ & " FROM (Term_HREMP INNER JOIN Term_JOB_HISTORY ON ED_EMPNBR=JH_EMPNBR AND JH_CURRENT <>0) "
        SQLQ = SQLQ & " WHERE 1 = 1"
        sSQLQ = Replace(sSQLQ, "HREMP.", "Term_HREMP.")
    End If
    SQLQ = SQLQ & " AND " & sSQLQ & " "
    SQLQ = SQLQ & " ORDER BY ED_EMPNBR,ED_SURNAME,ED_FNAME "

    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
                
        'File to export to
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2228_Employment_Summary_Report.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2228_Employment_Summary_Report(" & Trim(glbUserID) & ").xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
        
        FileCopy xlsFileTmp, xlsFileMat
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
        exSheet.Cells(2, 1) = "Date: " & Format(Now, "mmm dd, yyyy")
        'exSheet.Cells(3, 1) = "Time: " & Format(Now, "hh:mm")
    
        'exSheet.Cells(4, 1) = "Employee #"
        'exSheet.Cells(4, 2) = "Name"
        'exSheet.Cells(4, 3) = "Position"           'Desc
        'exSheet.Cells(4, 4) = "Service"
        'exSheet.Cells(4, 5) = "Program"
        'exSheet.Cells(4, 6) = "Manager"
        'exSheet.Cells(4, 7) = "Employment Status"
        'exSheet.Cells(4, 8) = "FTE #"
        'exSheet.Cells(4, 9) = "Salary"
        'exSheet.Cells(4, 10) = "Per"     'Driver License
        'exSheet.Cells(4, 11) = "Hire Date"
        'exSheet.Cells(4, 12) = "Perm DOH"
        'exSheet.Cells(4, 13) = "Union"
        'exSheet.Cells(4, 14) = "Location"   '
        
        xRow = 5
        
        xAnnSal = 0
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            exSheet.Cells(xRow, 1) = rsHREmp("ED_EMPNBR")
            exSheet.Cells(xRow, 2) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
            exSheet.Cells(xRow, 3) = GetJobDesc(rsHREmp("JH_JOB"))
            exSheet.Cells(xRow, 4) = getDivDesc(rsHREmp("ED_DIV"))
            exSheet.Cells(xRow, 5) = getGLDesc(rsHREmp("ED_GLNO"))
            exSheet.Cells(xRow, 6) = getDeptDesc(rsHREmp("ED_DEPTNO"))
            exSheet.Cells(xRow, 7) = GetTABLDesc("EDEM", rsHREmp("ED_EMP"))
            If Not IsNull(rsHREmp("JH_FTENUM")) Then
                exSheet.Cells(xRow, 8) = rsHREmp("JH_FTENUM")
            End If
            'If Not IsNull(rsHREmp("JH_FTENUM")) Then
            '    exSheet.Cells(xRow, 9) = rsHREmp("JH_FTENUM")
            'End If
            
            If OptAct.Value Then
                SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT (SH_CURRENT = 0) AND SH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " "
            Else
                SQLQ = "SELECT * FROM Term_SALARY_HISTORY WHERE NOT (SH_CURRENT = 0) AND SH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " "
                SQLQ = SQLQ & "AND TERM_SEQ = " & rsHREmp("TERM_SEQ") & " "
            End If
            If rsSal.State <> 0 Then rsSal.Close
            rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsSal.EOF Then
                exSheet.Cells(xRow, 9) = rsSal("SH_SALARY")
                exSheet.Cells(xRow, 10) = rsSal("SH_SALCD")
                If rsSal("SH_SALCD") = "H" Then
                    If Not IsNull(rsSal("SH_WHRS")) Then
                        xAnnSal = xAnnSal + rsSal("SH_SALARY") * rsSal("SH_WHRS") * 52
                    End If
                Else
                    xAnnSal = xAnnSal + rsSal("SH_SALARY")
                End If
            End If
            rsSal.Close
            exSheet.Cells(xRow, 11) = Format(rsHREmp("ED_DOH"), "mm/dd/yyyy")
            exSheet.Cells(xRow, 12) = Format(rsHREmp("ED_SENDTE"), "mm/dd/yyyy")
            exSheet.Cells(xRow, 13) = GetTABLDesc("EDOR", rsHREmp("ED_ORG"))
            exSheet.Cells(xRow, 14) = GetTABLDesc("EDLC", rsHREmp("ED_LOC"))
            
            rsHREmp.MoveNext
            xRow = xRow + 1
        Loop
        
        
        exSheet.Cells(xRow + 2, 5) = "Total of Annual Salary "
        exSheet.Cells(xRow + 2, 6) = "$" & Format(xAnnSal, "#,###.##")
        'exSheet.AutoFilterMode = True
        'exSheet.Range("K1:K20").AutoFilter Field:=1
        exSheet.Columns.AutoFit

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Err_Export_Data_to_Excel_CollectCorp:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
Resume Next

End Sub

Private Sub MacaulayChild_HoursUsed()
    'Ticket #27081
    Dim rsHREmp As New ADODB.Recordset
    Dim rsAttend As New ADODB.Recordset
    Dim SQLQ As String
    Dim sSQLQ As String
    
    Dim exApp As Object
    Dim exBook As Object
    Dim exSheet As Object
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum, termRowStart
    
    On Error GoTo MacaulayChild_HoursUsed_Err
    
    Screen.MousePointer = HOURGLASS
    
    'Ticket #22034 - Get Excel reports path
    Dim xExcelRptPath  As String
    If gsTRAININGMATRIX Then
        xExcelRptPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xExcelRptPath) = 0 Then
        xExcelRptPath = glbIHRREPORTS
    End If
    
    'File to export to
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2420AttendUsedTmp.xls"
    xlsFileMat = xExcelRptPath & IIf(Right(xExcelRptPath, 1) = "\", "", "\") & "AttendUsed_" & Trim(glbUserID) & ".xls"

    If Dir(xlsFileTmp) = "" Then
        Screen.MousePointer = DEFAULT
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
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
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
    xRow = 6
    
'ACTIVE EMPLOYEES
    SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO,ED_DOH,ED_DIV "
    SQLQ = SQLQ & "FROM HREMP WHERE 1 = 1 "
    SQLQ = SQLQ & " AND " & sSQLQ & " "
    SQLQ = SQLQ & " ORDER BY ED_FNAME, ED_SURNAME"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
                    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = "Please Wait...Active Employees"
        MDIMain.panHelp(0).FloodPercent = 0
                        
        exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy hh:mm")
        exSheet.Cells(2, 1) = "Time: " & Format(Now, "hh:mm")
        
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            'Print Name and Department
            exSheet.Cells(xRow, 1) = rsHREmp("ED_FNAME") & " " & rsHREmp("ED_SURNAME")
            exSheet.Cells(xRow, 2) = GetJobDesc(GetJHData(rsHREmp("ED_EMPNBR"), "JH_JOB", ""))
            exSheet.Cells(xRow, 3) = GetDeptName(rsHREmp("ED_DEPTNO"), "DF_NAME")
            exSheet.Cells(xRow, 4) = getDivDesc(rsHREmp("ED_DIV"))
            exSheet.Cells(xRow, 5) = rsHREmp("ED_DOH")
            exSheet.Cells(xRow, 6) = GetJHData(rsHREmp("ED_EMPNBR"), "JH_WHRS", "")
            exSheet.Cells(xRow, 7) = getGLDesc(GetJHData(rsHREmp("ED_EMPNBR"), "JH_GLNO", ""))
            
            'Used Hours from Attendance
            SQLQ = "SELECT  "
            'SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,3)='VAC' THEN AD_HRS ELSE 0 END) AS VACUSED,"
            'SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,3)='SIC' THEN AD_HRS ELSE 0 END) AS SICUSED,"
            'SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='OT' THEN AD_HRS ELSE 0 END) AS OTEARN,"
            'SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='CT' THEN AD_HRS ELSE 0 END) AS CTUSED,"
            'SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,3)='LOA' THEN AD_HRS ELSE 0 END) AS LOAUSED,"
            'SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='WCB' THEN AD_HRS ELSE 0 END) AS WSIBUSED"
            
            SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='VC' THEN AD_HRS ELSE 0 END) AS VACUSED,"
            SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='SK' THEN AD_HRS ELSE 0 END) AS SICUSED,"
            SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='FLOA' THEN AD_HRS ELSE 0 END) AS FLOAUSED,"
            'SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='OT' THEN AD_HRS ELSE 0 END) AS OTEARN,"
            SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='CT' THEN AD_HRS ELSE 0 END) AS CTUSED,"
            SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='LOA' THEN AD_HRS ELSE 0 END) AS LOAUSED,"
            SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='WCB' THEN AD_HRS ELSE 0 END) AS WSIBUSED,"
            SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='MAT' THEN AD_HRS ELSE 0 END) AS MATUSED"
            
            SQLQ = SQLQ & " FROM HR_ATTENDANCE"
            SQLQ = SQLQ & " WHERE AD_DOA >= " & Date_SQL(dlpDateRange(0).Text) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text)
            SQLQ = SQLQ & " AND AD_EMPNBR = " & rsHREmp("ED_EMPNBR")
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsAttend.EOF Then
                exSheet.Cells(xRow, 8) = rsAttend("VACUSED")
                exSheet.Cells(xRow, 9) = rsAttend("SICUSED")
                exSheet.Cells(xRow, 10) = rsAttend("FLOAUSED")
                exSheet.Cells(xRow, 11) = rsAttend("CTUSED")
                
                exSheet.Cells(xRow, 13) = rsAttend("LOAUSED")
                exSheet.Cells(xRow, 14) = rsAttend("WSIBUSED")
                exSheet.Cells(xRow, 15) = rsAttend("MATUSED")
            Else
                exSheet.Cells(xRow, 8) = 0
                exSheet.Cells(xRow, 9) = 0
                exSheet.Cells(xRow, 10) = 0
                exSheet.Cells(xRow, 11) = 0
                
                exSheet.Cells(xRow, 13) = 0
                exSheet.Cells(xRow, 14) = 0
                exSheet.Cells(xRow, 15) = 0
            End If
            rsAttend.Close
            Set rsAttend = Nothing
            
            'Salary
            exSheet.Cells(xRow, 12) = GetSHData(rsHREmp("ED_EMPNBR"), "SH_SALARY", "")
            
            'Borders
            'Dotted Line under the row
            exSheet.Range("A" & xRow & ":O" & xRow).Borders(xlEdgeBottom).LineStyle = xlDot

            xRow = xRow + 1
Next_Emp:
            rsHREmp.MoveNext
        Loop
    End If
    'exSheet.Columns.AutoFit

    'Border around the columns
    'Solid Line under the last row
    exSheet.Range("A" & xRow - 1 & ":O" & xRow - 1).Borders(xlEdgeBottom).LineStyle = xlSolid
    exSheet.Range("A6:P" & xRow - 1).Borders(xlOutline).Weight = xlThin
    
    rsHREmp.Close
    Set rsHREmp = Nothing


'**** TERMINATED EMPLOYEES ****

    sSQLQ = Replace(sSQLQ, "HREMP", "Term_HREMP")
    
    SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO,ED_DOH,ED_DIV,TERM_SEQ "
    SQLQ = SQLQ & "FROM Term_HREMP WHERE 1 = 1 "
    SQLQ = SQLQ & " AND " & sSQLQ & " "
    SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT Employee_Number FROM Term_HRTRMEMP "
    SQLQ = SQLQ & " WHERE Term_DOT >=" & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND Term_DOT <=" & Date_SQL(dlpDateRange(1).Text) & ")"
    
    SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
    SQLQ = SQLQ & " WHERE Term_DOT >=" & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND Term_DOT <=" & Date_SQL(dlpDateRange(1).Text) & ")"
    
    SQLQ = SQLQ & " ORDER BY ED_FNAME, ED_SURNAME"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = "Please Wait...Terminated Employees"
        MDIMain.panHelp(0).FloodPercent = 0
        
        termRowStart = 0
                    
        xRow = xRow + 2
        exSheet.Cells(xRow, 1) = "*** TERMINATED EMPLOYEES ***"
        exSheet.Rows(xRow).Font.Bold = True
        exSheet.Range("A" & xRow + 1 & ":O" & xRow + 1).Borders(xlEdgeBottom).LineStyle = xlSolid
        
        xRow = xRow + 2
        termRowStart = xRow
        
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            'Print Name and Department
            exSheet.Cells(xRow, 1) = rsHREmp("ED_FNAME") & " " & rsHREmp("ED_SURNAME")
            exSheet.Cells(xRow, 2) = GetJobDesc(GetTermJHData(rsHREmp("ED_EMPNBR"), "JH_JOB", "", rsHREmp("TERM_SEQ")))
            exSheet.Cells(xRow, 3) = GetDeptName(rsHREmp("ED_DEPTNO"), "DF_NAME")
            exSheet.Cells(xRow, 4) = getDivDesc(rsHREmp("ED_DIV"))
            exSheet.Cells(xRow, 5) = rsHREmp("ED_DOH")
            exSheet.Cells(xRow, 6) = GetTermJHData(rsHREmp("ED_EMPNBR"), "JH_WHRS", "", rsHREmp("TERM_SEQ"))
            exSheet.Cells(xRow, 7) = getGLDesc(GetTermJHData(rsHREmp("ED_EMPNBR"), "JH_GLNO", "", rsHREmp("TERM_SEQ")))
            
            'Used Hours from Attendance
            SQLQ = "SELECT  "
            'SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,3)='VAC' THEN AD_HRS ELSE 0 END) AS VACUSED,"
            'SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,3)='SIC' THEN AD_HRS ELSE 0 END) AS SICUSED,"
            'SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='OT' THEN AD_HRS ELSE 0 END) AS OTEARN,"
            'SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='CT' THEN AD_HRS ELSE 0 END) AS CTUSED,"
            'SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,3)='LOA' THEN AD_HRS ELSE 0 END) AS LOAUSED,"
            'SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='WCB' THEN AD_HRS ELSE 0 END) AS WSIBUSED"
            
            SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='VC' THEN AD_HRS ELSE 0 END) AS VACUSED,"
            SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='SK' THEN AD_HRS ELSE 0 END) AS SICUSED,"
            SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='FLOA' THEN AD_HRS ELSE 0 END) AS FLOAUSED,"
            'SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='OT' THEN AD_HRS ELSE 0 END) AS OTEARN,"
            SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='CT' THEN AD_HRS ELSE 0 END) AS CTUSED,"
            SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='LOA' THEN AD_HRS ELSE 0 END) AS LOAUSED,"
            SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='WCB' THEN AD_HRS ELSE 0 END) AS WSIBUSED,"
            SQLQ = SQLQ & " SUM(CASE WHEN AD_REASON='MAT' THEN AD_HRS ELSE 0 END) AS MATUSED"
            
            SQLQ = SQLQ & " FROM TERM_ATTENDANCE"
            SQLQ = SQLQ & " WHERE AD_DOA >= " & Date_SQL(dlpDateRange(0).Text) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text)
            SQLQ = SQLQ & " AND AD_EMPNBR = " & rsHREmp("ED_EMPNBR")
            SQLQ = SQLQ & " AND TERM_SEQ = " & rsHREmp("TERM_SEQ")
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsAttend.EOF Then
                exSheet.Cells(xRow, 8) = rsAttend("VACUSED")
                exSheet.Cells(xRow, 9) = rsAttend("SICUSED")
                exSheet.Cells(xRow, 10) = rsAttend("FLOAUSED")
                exSheet.Cells(xRow, 11) = rsAttend("CTUSED")
                
                exSheet.Cells(xRow, 13) = rsAttend("LOAUSED")
                exSheet.Cells(xRow, 14) = rsAttend("WSIBUSED")
                exSheet.Cells(xRow, 15) = rsAttend("MATUSED")
                
            Else
                exSheet.Cells(xRow, 8) = 0
                exSheet.Cells(xRow, 9) = 0
                exSheet.Cells(xRow, 10) = 0
                exSheet.Cells(xRow, 11) = 0
                
                exSheet.Cells(xRow, 13) = 0
                exSheet.Cells(xRow, 14) = 0
                exSheet.Cells(xRow, 15) = 0
            End If
            rsAttend.Close
            Set rsAttend = Nothing
            
            'Salary
            exSheet.Cells(xRow, 12) = GetTermSHData(rsHREmp("ED_EMPNBR"), "SH_SALARY", "", rsHREmp("TERM_SEQ"))
            
            'Borders
            'Dotted Line under the row
            exSheet.Range("A" & xRow & ":O" & xRow).Borders(xlEdgeBottom).LineStyle = xlDot

            xRow = xRow + 1
Next_Emp1:
            rsHREmp.MoveNext
        Loop

        'exSheet.Columns.AutoFit
    
        'Border around the columns
        'Solid Line under the last row
        exSheet.Range("A" & xRow - 1 & ":O" & xRow - 1).Borders(xlEdgeBottom).LineStyle = xlSolid
        exSheet.Range("A" & termRowStart & ":P" & xRow - 1).Borders(xlOutline).Weight = xlThin
        
    End If
    
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing


    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT
    
    Call Pause(1)
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

MacaulayChild_HoursUsed_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", "HREMP,Term_HREMP,HR_ATTENDANCE", "Used Attendance Hours")

Set exSheet = Nothing
Set exBook = Nothing
Set exApp = Nothing
    
End Sub

Private Sub WDGPHU_GeneralEmployee_ExcelReport()
    Dim rsHREmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
        
    On Error GoTo Err_WDGPHU_GeneralEmployee_ExcelReport
    
    
    Screen.MousePointer = HOURGLASS
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
    SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO,ED_ADMINBY,ED_LOC,ED_DIV,ED_SECTION,ED_DRIVERLIC,"
    SQLQ = SQLQ & "ED_EMP,ED_PT,ED_DOH,ED_PHONE,ED_ADDR1,ED_CITY,ED_PCODE,JH_JOB,JH_REPTAU,JH_REPTAU2,JH_FTENUM "
    SQLQ = SQLQ & " FROM (HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=JH_EMPNBR AND JH_CURRENT <>0) "
    
    If chkAct.Value = 1 And chkInact.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
    ElseIf chkAct.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
        If Len(StatusCodeCri) > 0 Then
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND ((TB_USR3 = 0) OR (TB_KEY IN ( " & StatusCodeCri & ")))" 'Active - Not on LOA
        Else
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND TB_USR3 = 0)" 'Active - Not on LOA
        End If
    ElseIf chkInact.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
        If Len(StatusCodeCri) > 0 Then
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND ((TB_USR3 <> 0) OR (TB_KEY IN ( " & StatusCodeCri & ")))"   'Inactive - On LOA
        Else
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND TB_USR3 <> 0)"   'Inactive - On LOA
        End If
        'sSQLQ = Replace(sSQLQ, "HREMP.", "Term_HREMP.")
    End If
    SQLQ = SQLQ & " AND " & sSQLQ & " "
    SQLQ = SQLQ & " ORDER BY ED_SURNAME,ED_FNAME "

    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
                
        'File to export to
        'Ticket #29657 - Renamed the existing Employee Fan Out to General Employee and created a new Employee Fan Out report with new format
        'xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2411EmpFanOutTmp.xls"
        'xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "EmpFanOutRpt_" & Trim(glbUserID) & ".xls"
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2411GeneralEmpTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "GeneralEmpRpt_" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait..."
        MDIMain.panHelp(0).FloodPercent = 0
        
        FileCopy xlsFileTmp, xlsFileMat
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
        exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy hh:mm")
        exSheet.Cells(2, 1) = "Time: " & Format(Now, "hh:mm")

        exSheet.Cells(4, 1) = "Employee #"
        exSheet.Cells(4, 2) = "Employee Name"
        exSheet.Cells(4, 3) = "Position"
        exSheet.Cells(4, 4) = "Phone #"
        exSheet.Cells(4, 5) = "Address"
        exSheet.Cells(4, 6) = "City"
        exSheet.Cells(4, 7) = "Postal Code" '
        exSheet.Cells(4, 8) = "Department"
        exSheet.Cells(4, 9) = "Manager"
        exSheet.Cells(4, 10) = "Division"
        exSheet.Cells(4, 11) = "    FTE    "
        exSheet.Cells(4, 12) = "Status"
        exSheet.Cells(4, 13) = "Location"
        exSheet.Cells(4, 14) = " Date of Hire "
        
        xRow = 5
        
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            exSheet.Cells(xRow, 1) = rsHREmp("ED_EMPNBR")
            exSheet.Cells(xRow, 2) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
            exSheet.Cells(xRow, 3) = GetJobDesc(rsHREmp("JH_JOB"))
            exSheet.Cells(xRow, 4) = Format(rsHREmp("ED_PHONE"), "(###) ###-####")
            exSheet.Cells(xRow, 5) = rsHREmp("ED_ADDR1")
            exSheet.Cells(xRow, 6) = rsHREmp("ED_CITY")
            exSheet.Cells(xRow, 7) = rsHREmp("ED_PCODE") 'GetTABLDesc("EDAB", rsHREmp("ED_ADMINBY"))
            exSheet.Cells(xRow, 8) = getDeptDesc(rsHREmp("ED_DEPTNO"))
            exSheet.Cells(xRow, 9) = getEEName(rsHREmp("JH_REPTAU"))
            exSheet.Cells(xRow, 10) = getDivDesc(rsHREmp("ED_DIV"))  'GetTABLDesc("EDSE", rsHREmp("ED_SECTION"))
            'Ticket #28037 - Jerry wants code only to save some space.
            'exSheet.Cells(xRow, 11) = GetTABLDesc("EDPT", rsHREmp("ED_PT"))
            'exSheet.Cells(xRow, 12) = GetTABLDesc("EDEM", rsHREmp("ED_EMP"))
            exSheet.Cells(xRow, 11) = rsHREmp("ED_PT")
            exSheet.Cells(xRow, 12) = rsHREmp("ED_EMP")
            exSheet.Cells(xRow, 13) = GetTABLDesc("EDLC", rsHREmp("ED_LOC"))
            exSheet.Cells(xRow, 14) = Format(rsHREmp("ED_DOH"), "mm/dd/yyyy")
                        
            'exSheet.Cells(xRow, 15) = getEEName(rsHREmp("JH_REPTAU"))
            'exSheet.Cells(xRow, 16) = getEEName(rsHREmp("JH_REPTAU2"))
            'exSheet.Cells(xRow, 17) = rsHREmp("JH_JOB")
            'exSheet.Cells(xRow, 18) = GetJobDesc(rsHREmp("JH_JOB"))
            
            rsHREmp.MoveNext
            
            xRow = xRow + 1
        Loop
        
        'exSheet.AutoFilterMode = True
        'exSheet.Range("K1:K20").AutoFilter Field:=1
        exSheet.Columns.AutoFit

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
        
    Else
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        MsgBox "No employees found in this selection criteria."
    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Err_WDGPHU_GeneralEmployee_ExcelReport:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "General Employee Excel", fglbEmpTable, "SELECT")

Set exSheet = Nothing
Set exBook = Nothing
Set exApp = Nothing

End Sub

'Ticket #29657 - New Report as Employee General report but with limited columns
Private Sub WDGPHU_EmployeeFanOut_2nd_ExcelReport()
    Dim rsHREmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
        
    On Error GoTo Err_WDGPHU_EmployeeFanOut_2nd_ExcelReport
    
    
    Screen.MousePointer = HOURGLASS
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
    SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO,ED_ADMINBY,ED_LOC,ED_DIV,ED_SECTION,ED_DRIVERLIC,"
    SQLQ = SQLQ & "ED_EMP,ED_PT,ED_DOH,ED_PHONE,ED_ADDR1,ED_CITY,ED_PCODE,JH_JOB,JH_REPTAU,JH_REPTAU2,JH_FTENUM "
    SQLQ = SQLQ & " FROM (HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=JH_EMPNBR AND JH_CURRENT <>0) "
    
    If chkAct.Value = 1 And chkInact.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
    ElseIf chkAct.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
        If Len(StatusCodeCri) > 0 Then
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND ((TB_USR3 = 0) OR (TB_KEY IN ( " & StatusCodeCri & ")))" 'Active - Not on LOA
        Else
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND TB_USR3 = 0)" 'Active - Not on LOA
        End If
    ElseIf chkInact.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
        If Len(StatusCodeCri) > 0 Then
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND ((TB_USR3 <> 0) OR (TB_KEY IN ( " & StatusCodeCri & ")))"   'Inactive - On LOA
        Else
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND TB_USR3 <> 0)"   'Inactive - On LOA
        End If
        'sSQLQ = Replace(sSQLQ, "HREMP.", "Term_HREMP.")
    End If
    If Len(sSQLQ) > 0 Then
        SQLQ = SQLQ & " AND " & sSQLQ & " "
    End If
    SQLQ = SQLQ & " ORDER BY ED_SURNAME,ED_FNAME "

    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
                
        'File to export to
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2411EmpFanOutTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "EmpFanOutRpt_" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait..."
        MDIMain.panHelp(0).FloodPercent = 0
        
        FileCopy xlsFileTmp, xlsFileMat
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
        exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy hh:mm")
        exSheet.Cells(2, 1) = "Time: " & Format(Now, "hh:mm")

        exSheet.Cells(4, 1) = "Employee #"
        exSheet.Cells(4, 2) = "Employee Name"
        exSheet.Cells(4, 3) = "Position"
        exSheet.Cells(4, 4) = "Phone #"
        'exSheet.Cells(4, 5) = "Address"
        exSheet.Cells(4, 5) = "City"
        'exSheet.Cells(4, 7) = "Postal Code" '
        exSheet.Cells(4, 6) = "Department"
        exSheet.Cells(4, 7) = "Manager"
        exSheet.Cells(4, 8) = "Division"
        exSheet.Cells(4, 9) = "    FTE    "
        exSheet.Cells(4, 10) = "Status"
        exSheet.Cells(4, 11) = "Location"
        'exSheet.Cells(4, 14) = " Date of Hire "
        
        xRow = 5
        
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            exSheet.Cells(xRow, 1) = rsHREmp("ED_EMPNBR")
            exSheet.Cells(xRow, 2) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
            exSheet.Cells(xRow, 3) = GetJobDesc(rsHREmp("JH_JOB"))
            exSheet.Cells(xRow, 4) = Format(rsHREmp("ED_PHONE"), "(###) ###-####")
            'exSheet.Cells(xRow, 5) = rsHREmp("ED_ADDR1")
            exSheet.Cells(xRow, 5) = rsHREmp("ED_CITY")
            'exSheet.Cells(xRow, 7) = rsHREmp("ED_PCODE") 'GetTABLDesc("EDAB", rsHREmp("ED_ADMINBY"))
            exSheet.Cells(xRow, 6) = getDeptDesc(rsHREmp("ED_DEPTNO"))
            exSheet.Cells(xRow, 7) = getEEName(rsHREmp("JH_REPTAU"))
            exSheet.Cells(xRow, 8) = getDivDesc(rsHREmp("ED_DIV"))  'GetTABLDesc("EDSE", rsHREmp("ED_SECTION"))
            'Ticket #28037 - Jerry wants code only to save some space.
            'exSheet.Cells(xRow, 11) = GetTABLDesc("EDPT", rsHREmp("ED_PT"))
            'exSheet.Cells(xRow, 12) = GetTABLDesc("EDEM", rsHREmp("ED_EMP"))
            exSheet.Cells(xRow, 9) = rsHREmp("ED_PT")
            exSheet.Cells(xRow, 10) = rsHREmp("ED_EMP")
            exSheet.Cells(xRow, 11) = GetTABLDesc("EDLC", rsHREmp("ED_LOC"))
            'exSheet.Cells(xRow, 14) = Format(rsHREmp("ED_DOH"), "mm/dd/yyyy")
                        
            'exSheet.Cells(xRow, 15) = getEEName(rsHREmp("JH_REPTAU"))
            'exSheet.Cells(xRow, 16) = getEEName(rsHREmp("JH_REPTAU2"))
            'exSheet.Cells(xRow, 17) = rsHREmp("JH_JOB")
            'exSheet.Cells(xRow, 18) = GetJobDesc(rsHREmp("JH_JOB"))
            
            rsHREmp.MoveNext
            
            xRow = xRow + 1
        Loop
        
        'exSheet.AutoFilterMode = True
        'exSheet.Range("K1:K20").AutoFilter Field:=1
        exSheet.Columns.AutoFit

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
        
    Else
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        MsgBox "No employees found in this selection criteria."
    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Err_WDGPHU_EmployeeFanOut_2nd_ExcelReport:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Gen. Employee Excel", fglbEmpTable, "SELECT")

Set exSheet = Nothing
Set exBook = Nothing
Set exApp = Nothing

End Sub

Private Sub WDGPHU_EmployeeTelephone_ExcelReport()
    Dim rsHREmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
        
    On Error GoTo Err_WDGPHU_EmployeeTelephone_ExcelReport
    
    
    Screen.MousePointer = HOURGLASS
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
    SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO,ED_ADMINBY,ED_LOC,ED_DIV,ED_SECTION,ED_DRIVERLIC,"
    SQLQ = SQLQ & "ED_EMP,ED_PT,ED_DOH,ED_PHONE,ED_ADDR1,ED_CITY,ED_PCODE,ED_REGION,ED_INTEL,JH_JOB,JH_REPTAU,JH_REPTAU2,JH_FTENUM "
    SQLQ = SQLQ & " FROM (HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=JH_EMPNBR AND JH_CURRENT <>0) "
    
    If chkAct.Value = 1 And chkInact.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
    ElseIf chkAct.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
        If Len(StatusCodeCri) > 0 Then
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND ((TB_USR3 = 0) OR (TB_KEY IN ( " & StatusCodeCri & ")))" 'Active - Not on LOA
        Else
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND TB_USR3 = 0)" 'Active - Not on LOA
        End If
    ElseIf chkInact.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
        If Len(StatusCodeCri) > 0 Then
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND ((TB_USR3 <> 0) OR (TB_KEY IN ( " & StatusCodeCri & ")))"   'Inactive - On LOA
        Else
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND TB_USR3 <> 0)"   'Inactive - On LOA
        End If
        'sSQLQ = Replace(sSQLQ, "HREMP.", "Term_HREMP.")
    End If
    SQLQ = SQLQ & " AND " & sSQLQ & " "
    SQLQ = SQLQ & " ORDER BY ED_SURNAME,ED_FNAME "

    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
                
        'File to export to
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2411EmpTelephoneTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "EmpTelephoneList_" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait..."
        MDIMain.panHelp(0).FloodPercent = 0
        
        FileCopy xlsFileTmp, xlsFileMat
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
        exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy hh:mm")
        exSheet.Cells(2, 1) = "Time: " & Format(Now, "hh:mm")

    
        exSheet.Cells(4, 1) = "Employee Name"
        exSheet.Cells(4, 2) = "Extention #"
        exSheet.Cells(4, 3) = "Position"
        exSheet.Cells(4, 4) = "Location"
        exSheet.Cells(4, 5) = "Program"
        
        xRow = 5
        
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            exSheet.Cells(xRow, 1) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
            exSheet.Cells(xRow, 2) = rsHREmp("ED_INTEL")
            exSheet.Cells(xRow, 3) = GetJobDesc(rsHREmp("JH_JOB"))
            exSheet.Cells(xRow, 4) = GetTABLDesc("EDLC", rsHREmp("ED_LOC"))
            exSheet.Cells(xRow, 5) = GetTABLDesc("EDRG", rsHREmp("ED_REGION"))
            
            rsHREmp.MoveNext
            
            xRow = xRow + 1
        Loop
        
        'exSheet.AutoFilterMode = True
        'exSheet.Range("K1:K20").AutoFilter Field:=1
        exSheet.Columns.AutoFit

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
        
    Else
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        MsgBox "No employees found in this selection criteria."
    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Err_WDGPHU_EmployeeTelephone_ExcelReport:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")

Set exSheet = Nothing
Set exBook = Nothing
Set exApp = Nothing

End Sub

Private Sub WDGPHU_License_ExcelReport()
    Dim rsHREmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
        
    On Error GoTo Err_WDGPHU_License_ExcelReport
    
    
    Screen.MousePointer = HOURGLASS
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
    SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO,ED_ADMINBY,ED_LOC,ED_DIV,ED_SECTION,ED_DRIVERLIC,"
    SQLQ = SQLQ & "ED_EMP,ED_PT,ED_DOH,ED_PHONE,ED_ADDR1,ED_CITY,ED_PCODE,ED_REGION,ED_INTEL,JH_JOB,JH_REPTAU,JH_REPTAU2,JH_FTENUM "
    SQLQ = SQLQ & " FROM (HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=JH_EMPNBR AND JH_CURRENT <>0) "
    
    If chkAct.Value = 1 And chkInact.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
    ElseIf chkAct.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
        If Len(StatusCodeCri) > 0 Then
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND ((TB_USR3 = 0) OR (TB_KEY IN ( " & StatusCodeCri & ")))" 'Active - Not on LOA
        Else
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND TB_USR3 = 0)" 'Active - Not on LOA
        End If
    ElseIf chkInact.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
        If Len(StatusCodeCri) > 0 Then
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND ((TB_USR3 <> 0) OR (TB_KEY IN ( " & StatusCodeCri & ")))"   'Inactive - On LOA
        Else
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND TB_USR3 <> 0)"   'Inactive - On LOA
        End If
        'sSQLQ = Replace(sSQLQ, "HREMP.", "Term_HREMP.")
    End If
    SQLQ = SQLQ & " AND " & sSQLQ & " "
    SQLQ = SQLQ & " ORDER BY ED_SURNAME,ED_FNAME "

    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
                
        'File to export to
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2411LicenseTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "LicenseRpt_" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait..."
        MDIMain.panHelp(0).FloodPercent = 0
        
        FileCopy xlsFileTmp, xlsFileMat
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
        exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy hh:mm")
        exSheet.Cells(2, 1) = "Time: " & Format(Now, "hh:mm")

    
        exSheet.Cells(4, 1) = "Employee Name"
        exSheet.Cells(4, 2) = "Position"
        exSheet.Cells(4, 3) = "    FTE    "
        exSheet.Cells(4, 4) = "Status"
        exSheet.Cells(4, 5) = " Date of Hire "
        exSheet.Cells(4, 6) = "Program" '
        exSheet.Cells(4, 7) = "RA #1 Manager"
        exSheet.Cells(4, 8) = "Expiry Date for DL"
        
        xRow = 5
        
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            exSheet.Cells(xRow, 1) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
            exSheet.Cells(xRow, 2) = GetJobDesc(rsHREmp("JH_JOB"))
            exSheet.Cells(xRow, 3) = GetTABLDesc("EDPT", rsHREmp("ED_PT"))
            exSheet.Cells(xRow, 4) = GetTABLDesc("EDEM", rsHREmp("ED_EMP"))
            exSheet.Cells(xRow, 5) = Format(rsHREmp("ED_DOH"), "mm/dd/yyyy")
            exSheet.Cells(xRow, 6) = GetTABLDesc("EDRG", rsHREmp("ED_REGION"))
            exSheet.Cells(xRow, 7) = getEEName(rsHREmp("JH_REPTAU"))
            'exSheet.Cells(xRow, 8) = get_License_ExpiryDate_FollowUp(rsHREmp("ED_EMPNBR"))
            exSheet.Cells(xRow, 8) = get_Employee_Flag_Value(rsHREmp("ED_EMPNBR"), "EF_FLAGDTE3")
            
            rsHREmp.MoveNext
            
            xRow = xRow + 1
        Loop
        
        'exSheet.AutoFilterMode = True
        'exSheet.Range("K1:K20").AutoFilter Field:=1
        exSheet.Columns.AutoFit

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
        
    Else
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        MsgBox "No employees found in this selection criteria."
    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Err_WDGPHU_License_ExcelReport:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")

Set exSheet = Nothing
Set exBook = Nothing
Set exApp = Nothing

End Sub

Private Function get_License_ExpiryDate_FollowUp(xEmpNbr)
    Dim rsFollowUp As New ADODB.Recordset
    Dim SQLQ As String
    
    'Jerry said to use Employee Flag 3 Date instead - after their data clean up they removed the Follow Up 'DRIV'.
    'EF_FLAGDTE3 = get_Employee_Flag_Value
        
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & xEmpNbr
    SQLQ = SQLQ & " AND EF_FREAS = 'DRIV'"
    SQLQ = SQLQ & " ORDER BY EF_FDATE DESC, EF_COMPLETED ASC"
    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsFollowUp.EOF Then
        rsFollowUp.MoveFirst
        get_License_ExpiryDate_FollowUp = rsFollowUp("EF_FDATE")
    Else
        get_License_ExpiryDate_FollowUp = ""
    End If
    rsFollowUp.Close
    Set rsFollowUp = Nothing
    
End Function

Private Function get_Employee_Flag_Value(xEmpNbr, xFlagField)
    Dim rsEmpFlag As New ADODB.Recordset
    Dim SQLQ As String
    
    get_Employee_Flag_Value = ""
    
    SQLQ = "SELECT EF_EMPNBR, " & xFlagField & " FROM HREMP_FLAGS WHERE EF_EMPNBR = " & xEmpNbr
    rsEmpFlag.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    If Not rsEmpFlag.EOF Then
        If Not IsNull(rsEmpFlag(xFlagField)) And rsEmpFlag(xFlagField) <> "" Then
            get_Employee_Flag_Value = rsEmpFlag(xFlagField)
        Else
            get_Employee_Flag_Value = ""
        End If
    Else
        get_Employee_Flag_Value = ""
    End If
    rsEmpFlag.Close
    Set rsEmpFlag = Nothing
End Function

Private Sub WDGPHU_PerfManagement_ExcelReport()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsPerf As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    Dim lstTwoMonth
    Dim lstLastMonth
        
    On Error GoTo Err_WDGPHU_PerfManagement_ExcelReport
    
    
    Screen.MousePointer = HOURGLASS
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
    SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO,ED_ADMINBY,ED_LOC,ED_DIV,ED_SECTION,ED_DRIVERLIC,"
    SQLQ = SQLQ & "ED_EMP,ED_PT,ED_DOH,ED_PHONE,ED_ADDR1,ED_CITY,ED_PCODE,ED_REGION,ED_INTEL,JH_JOB,JH_REPTAU,JH_REPTAU2,JH_FTENUM "
    SQLQ = SQLQ & " FROM (HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=JH_EMPNBR AND JH_CURRENT <>0) "
    
    If chkAct.Value = 1 And chkInact.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
    ElseIf chkAct.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
        If Len(StatusCodeCri) > 0 Then
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND ((TB_USR3 = 0) OR (TB_KEY IN ( " & StatusCodeCri & ")))" 'Active - Not on LOA
        Else
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND TB_USR3 = 0)" 'Active - Not on LOA
        End If
    ElseIf chkInact.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
        If Len(StatusCodeCri) > 0 Then
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND ((TB_USR3 <> 0) OR (TB_KEY IN ( " & StatusCodeCri & ")))"   'Inactive - On LOA
        Else
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND TB_USR3 <> 0)"   'Inactive - On LOA
        End If
        'sSQLQ = Replace(sSQLQ, "HREMP.", "Term_HREMP.")
    End If
    SQLQ = SQLQ & " AND " & sSQLQ & " "
    SQLQ = SQLQ & " ORDER BY ED_SURNAME,ED_FNAME "

    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
                
        'File to export to
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2411PerfManagementTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "PerfMgmtRpt_" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait..."
        MDIMain.panHelp(0).FloodPercent = 0
        
        FileCopy xlsFileTmp, xlsFileMat
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
        exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy hh:mm")
        exSheet.Cells(2, 1) = "Time: " & Format(Now, "hh:mm")

    
        exSheet.Cells(4, 1) = "Employee Name"
        exSheet.Cells(4, 2) = "Position"
        exSheet.Cells(4, 3) = "Status"
        exSheet.Cells(4, 4) = "    FTE    "
        exSheet.Cells(4, 5) = "RA #1 Manager"
        exSheet.Cells(4, 6) = " Date of Hire "
        exSheet.Cells(4, 7) = "Perf. Review Date"
        exSheet.Cells(4, 8) = "Next Pref. Review Date"
        exSheet.Cells(4, 9) = "Comments"
        
        xRow = 5
        
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            'Check if the employee falls into the Date Range selection criteria
            SQLQ = "SELECT * FROM HR_PERFORM_HISTORY WHERE PH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND PH_CURRENT <> 0 "
            rsPerf.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If rsPerf.EOF Then
                'Employee is missing PA record, include in the report and skip the date range selection criteria condition
                rsPerf.Close
                Set rsPerf = Nothing
                GoTo PrintEmployeeRec
            End If
            rsPerf.Close
            Set rsPerf = Nothing
            
            'Check if the employee falls into the Date Range selection criteria
            SQLQ = "SELECT * FROM HR_PERFORM_HISTORY WHERE PH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " AND PH_CURRENT <> 0 "
            
            'Review Date selection criteria
            If IsDate(dlpDateRange(0).Text) Or IsDate(dlpDateRange(1).Text) Then
                If IsDate(dlpDateRange(0).Text) Then
                    SQLQ = SQLQ & "AND ((PH_PREVIEW >= " & Date_SQL(dlpDateRange(0).Text) & " "
                End If
                If IsDate(dlpDateRange(1).Text) Then
                    SQLQ = SQLQ & "AND PH_PREVIEW <= " & Date_SQL(dlpDateRange(1).Text) & ")"
                End If
                
                SQLQ = SQLQ & " OR (PH_PREVIEW IS NULL OR PH_PREVIEW = ''))"
            End If
            
            'Next Review Date selection criteria
            If IsDate(dlpDateRange(2).Text) Or IsDate(dlpDateRange(3).Text) Then
                If IsDate(dlpDateRange(2).Text) Then
                    SQLQ = SQLQ & "AND (((PH_PNEXT >= " & Date_SQL(dlpDateRange(2).Text) & " "
                End If
                If IsDate(dlpDateRange(3).Text) Then
                    SQLQ = SQLQ & "AND PH_PNEXT <= " & Date_SQL(dlpDateRange(3).Text) & ") "
                End If
                SQLQ = SQLQ & " OR (PH_PNEXT IS NULL OR PH_PNEXT = ''))"
                
                If chkLasg2PrvMonths.Value = 1 Then
                    lstTwoMonth = getFirstDayInMonth(DateAdd("m", -2, CVDate(dlpDateRange(2).Text)))
                    lstLastMonth = getLastDayInMonth(DateAdd("m", -1, CVDate(dlpDateRange(2).Text)))
                    'lstTwoMonth = CVDate(Format(month(DateAdd("m", -2, Now)) & "/" & "01/" & Year(DateAdd("m", -2, Now)), "mm/dd/yyyy"))
                    'lstLastMonth = CVDate(DateAdd("d", -1, DateAdd("m", 2, CVDate(Format(month(DateAdd("m", -2, Now)) & "/" & "01/" & Year(DateAdd("m", -2, Now)), "mm/dd/yyyy")))))
                    SQLQ = SQLQ & " OR (PH_PNEXT >= " & Date_SQL(lstTwoMonth)
                    SQLQ = SQLQ & " AND PH_PNEXT <= " & Date_SQL(lstLastMonth) & "))"
                Else
                    SQLQ = SQLQ & ")"
                End If
            End If
            
            SQLQ = SQLQ & " ORDER BY PH_PREVIEW DESC, PH_PNEXT DESC"
            rsPerf.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsPerf.EOF Then
                exSheet.Cells(xRow, 7) = rsPerf("PH_PREVIEW")
                exSheet.Cells(xRow, 8) = rsPerf("PH_PNEXT")
                exSheet.Cells(xRow, 9) = rsPerf("PH_COMMENTS")
            End If
            rsPerf.Close
            Set rsPerf = Nothing
            
PrintEmployeeRec:
            exSheet.Cells(xRow, 1) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
            exSheet.Cells(xRow, 2) = GetJobDesc(rsHREmp("JH_JOB"))
            'Ticket #28037 - Jerry wants code only to save some space.
            'exSheet.Cells(xRow, 3) = GetTABLDesc("EDEM", rsHREmp("ED_EMP"))
            'exSheet.Cells(xRow, 4) = GetTABLDesc("EDPT", rsHREmp("ED_PT"))
            exSheet.Cells(xRow, 3) = rsHREmp("ED_EMP")
            exSheet.Cells(xRow, 4) = rsHREmp("ED_PT")
            exSheet.Cells(xRow, 5) = getEEName(rsHREmp("JH_REPTAU"))
            exSheet.Cells(xRow, 6) = Format(rsHREmp("ED_DOH"), "mm/dd/yyyy")
                                    
            rsHREmp.MoveNext
            
            xRow = xRow + 1
        Loop
        
        'exSheet.AutoFilterMode = True
        'exSheet.Range("K1:K20").AutoFilter Field:=1
        exSheet.Columns.AutoFit

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
        
    Else
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        MsgBox "No employees found in this selection criteria."
    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Err_WDGPHU_PerfManagement_ExcelReport:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")

Set exSheet = Nothing
Set exBook = Nothing
Set exApp = Nothing
End Sub

Private Sub WDGPHU_Vacation_ExcelReport()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsSal As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    Dim xTaken, xPrvYear
        
    On Error GoTo Err_WDGPHU_Vacation_ExcelReport
    
    
    Screen.MousePointer = HOURGLASS
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
    SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO,ED_ADMINBY,ED_LOC,ED_DIV,ED_SECTION,ED_DRIVERLIC,"
    SQLQ = SQLQ & "ED_EMP,ED_PT,ED_DOH,ED_PHONE,ED_ADDR1,ED_CITY,ED_PCODE,ED_REGION,ED_INTEL,ED_EFDATE,ED_VAC,ED_PVAC,JH_JOB,JH_REPTAU,JH_REPTAU2,JH_FTENUM "
    SQLQ = SQLQ & " FROM (HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=JH_EMPNBR AND JH_CURRENT <>0) "
    
    If chkAct.Value = 1 And chkInact.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
    ElseIf chkAct.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
        If Len(StatusCodeCri) > 0 Then
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND ((TB_USR3 = 0) OR (TB_KEY IN ( " & StatusCodeCri & ")))" 'Active - Not on LOA
        Else
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND TB_USR3 = 0)" 'Active - Not on LOA
        End If
    ElseIf chkInact.Value = 1 Then
        SQLQ = SQLQ & " WHERE 1 = 1"
        If Len(StatusCodeCri) > 0 Then
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND ((TB_USR3 <> 0) OR (TB_KEY IN ( " & StatusCodeCri & ")))"   'Inactive - On LOA
        Else
            SQLQ = SQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND TB_USR3 <> 0)"   'Inactive - On LOA
        End If
        'sSQLQ = Replace(sSQLQ, "HREMP.", "Term_HREMP.")
    End If
    SQLQ = SQLQ & " AND " & sSQLQ & " "
    SQLQ = SQLQ & " ORDER BY ED_SURNAME,ED_FNAME "

    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
                
        'File to export to
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2411VacationTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "VacationRpt_" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait..."
        MDIMain.panHelp(0).FloodPercent = 0
        
        FileCopy xlsFileTmp, xlsFileMat
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
        exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy hh:mm")
        exSheet.Cells(2, 1) = "Time: " & Format(Now, "hh:mm")
        exSheet.Cells(2, 2) = "Vacation Report as of " & Format(dlpAsOf, "mmm dd'yy")
    
        exSheet.Cells(4, 1) = "Employee Name"
        exSheet.Cells(4, 2) = "Program"
        exSheet.Cells(4, 3) = "Position"
        exSheet.Cells(4, 4) = "RA #1 Manager"
        exSheet.Cells(4, 5) = "Status"
        exSheet.Cells(4, 6) = "Current Rate"
        'exSheet.Cells(4, 7) = "Previous Rate"
        'Ticket #29157 - Corrected the column title to be end of Previous as per the As of Date entered.
        'exSheet.Cells(4, 7) = "Balance at " & Format(dlpAsOf, "mmm dd'yy")
        xPrvYear = Year(dlpAsOf) - 1
        exSheet.Cells(4, 7) = "Balance at " & Format("12/31/" & xPrvYear, "mmm dd'yy")
        
        exSheet.Cells(4, 8) = "Jan 1st Entitlement"
        exSheet.Cells(4, 9) = "Taken"
        exSheet.Cells(4, 10) = "Hours O/S"
        exSheet.Cells(4, 11) = "$$ O/S"
        
        
        xRow = 5
        
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            exSheet.Cells(xRow, 1) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
            exSheet.Cells(xRow, 2) = GetTABLDesc("EDRG", rsHREmp("ED_REGION"))
            exSheet.Cells(xRow, 3) = GetJobDesc(rsHREmp("JH_JOB"))
            exSheet.Cells(xRow, 4) = getEEName(rsHREmp("JH_REPTAU"))
            exSheet.Cells(xRow, 5) = GetTABLDesc("EDEM", rsHREmp("ED_EMP"))
            
            'Get Hours Taken as of Date
            xTaken = 0
            xTaken = Get_VacationTaken(rsHREmp("ED_EMPNBR"), rsHREmp("ED_EFDATE"), dlpAsOf.Text)
            If IsNull(xTaken) Then xTaken = 0
            
            exSheet.Cells(xRow, 7) = Round(IIf(IsNull(rsHREmp("ED_PVAC")), 0, rsHREmp("ED_PVAC")), 2) 'Balance at Dec 31st/As of Date
            exSheet.Cells(xRow, 8) = Round(IIf(IsNull(rsHREmp("ED_VAC")), 0, rsHREmp("ED_VAC")), 2) 'Entitl. Jan 1st
            exSheet.Cells(xRow, 9) = Round(IIf(IsNull(xTaken), 0, xTaken), 2)   'Taken as of Date
            exSheet.Cells(xRow, 10) = (rsHREmp("ED_PVAC") + rsHREmp("ED_VAC")) - xTaken 'Hours O/S as of Date
            
            SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " "
            SQLQ = SQLQ & "ORDER BY SH_EDATE DESC, SH_ID DESC, SH_CURRENT " & IIf(glbSQL, "DESC", "")
            rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsSal.EOF Then
                rsSal.MoveFirst
                exSheet.Cells(xRow, 6) = rsSal("SH_SALARY") 'Current Rate
                exSheet.Cells(xRow, 11) = rsSal("SH_SALARY") * ((rsHREmp("ED_PVAC") + rsHREmp("ED_VAC")) - xTaken)  'Dollar O/S as of Date
                'rsSal.MoveNext
                'exSheet.Cells(xRow, 7) = rsSal("SH_SALARY") 'Previous Rate
            End If
            rsSal.Close
            Set rsSal = Nothing
            
            
            rsHREmp.MoveNext
            
            xRow = xRow + 1
        Loop
        
        'exSheet.AutoFilterMode = True
        'exSheet.Range("K1:K20").AutoFilter Field:=1
        exSheet.Columns.AutoFit

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
        
    Else
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        MsgBox "No employees found in this selection criteria."
    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Err_WDGPHU_Vacation_ExcelReport:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")

Set exSheet = Nothing
Set exBook = Nothing
Set exApp = Nothing

End Sub

Private Sub NiagaraFallsSickScreen()
Dim xFDate, xTDate
    xFDate = CVDate("Jan 1, " & Year(Date))
    xTDate = CVDate("Dec 31, " & Year(Date))
    If InStr(1, fglbFileName, "SN2276_CUPESick.rpt") > 0 Then
        dlpDateRange(0).Text = xFDate
        dlpDateRange(1).Text = xTDate
        frmDate.Visible = True
    Else
        frmDate.Visible = False
        dlpDateRange(0).Text = ""
        dlpDateRange(1).Text = ""
    End If
End Sub

Private Function getLocEmpPosField(xEmpNo, xEmpTermSeq, xField)
Dim rsEmpJob As New ADODB.Recordset
Dim rsHRJob As New ADODB.Recordset
Dim SQLQ As String
Dim xJobCode As String
Dim retval As String
    retval = 0
    If xEmpTermSeq = 0 Then
        SQLQ = "SELECT JH_EMPNBR, " & xField & " FROM HR_JOB_HISTORY WHERE NOT (JH_CURRENT = 0) "
        SQLQ = SQLQ & "AND JH_EMPNBR = " & xEmpNo & " "
    Else
        SQLQ = "SELECT JH_EMPNBR, " & xField & " FROM Term_JOB_HISTORY WHERE NOT (JH_CURRENT = 0) "
        SQLQ = SQLQ & "AND JH_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND TERM_SEQ = " & xEmpTermSeq & " "
    End If
    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpJob.EOF Then
        retval = rsEmpJob(xField)
    End If
    rsEmpJob.Close

    getLocEmpPosField = retval
End Function

Private Sub Upt_HRATTWRK(xEmpNo, xDay, xDHrs, xLevel, rsWRK As ADODB.Recordset)
Dim rsAtt As New ADODB.Recordset
'Dim rsWRK As New ADODB.Recordset
Dim SQLQ As String
    
    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND AD_REASON like 'SIC%' "
    SQLQ = SQLQ & "AND AD_DOA >= " & Date_SQL(dlpDateRange(0).Text) & " "
    SQLQ = SQLQ & "AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text) & " "
    SQLQ = SQLQ & "ORDER BY AD_DOA "
    If rsAtt.State <> 0 Then rsAtt.Close
    rsAtt.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsAtt.EOF
        rsWRK.AddNew
        rsWRK("AD_EMPNBR") = rsAtt("AD_EMPNBR")
        rsWRK("AD_REASON") = rsAtt("AD_REASON")
        rsWRK("AD_DOA") = rsAtt("AD_DOA")
        rsWRK("AD_HRS") = rsAtt("AD_HRS")
        rsWRK("AD_MACHINE_HRS") = xDay
        rsWRK("AD_DHRS") = rsAtt("AD_DHRS")
        rsWRK("AD_MACHINE_RATE") = xLevel
        rsWRK("AD_LDATE") = Date
        rsWRK("AD_LTIME") = Time$
        rsWRK("AD_WRKEMP") = glbUserID
        rsWRK.Update
        rsAtt.MoveNext
    Loop
    rsAtt.Close
    'SELECT AD_EMPNBR,AD_REASON,AD_DOA,AD_HRS,AD_MACHINE_HRS,AD_DHRS,AD_MACHINE_RATE,AD_LDATE,AD_LTIME,AD_WRKEMP FROM HR_ATTENDANCE WHERE AD_WRKEMP = '999999999'
End Sub

Function getFirstDayInMonth(Optional dtmDate As Date = 0) As Date
    ' Return the first day in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use the current date.
        dtmDate = Date
    End If
    getFirstDayInMonth = DateSerial(Year(dtmDate), month(dtmDate), 1)
End Function

Function getLastDayInMonth(Optional dtmDate As Date = 0) As Date
    ' Return the last day in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use the current date.
        dtmDate = Date
    End If
    getLastDayInMonth = DateSerial(Year(dtmDate), month(dtmDate) + 1, 0)
End Function

'Ticket #29695 - Cascade Canada
Private Sub Cascade_EmployeeInfo_ExcelReport()
    Dim rsHREmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
        
    On Error GoTo Err_Cascade_EmployeeInfo_ExcelReport
    
    
    Screen.MousePointer = HOURGLASS
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
    SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_DEPTNO,ED_ADMINBY,ED_LOC,ED_DIV,ED_SECTION,ED_DRIVERLIC,"
    SQLQ = SQLQ & "ED_EMP,ED_EMPTYPE,ED_PT,ED_ORG,ED_DOH,ED_DOB,ED_PHONE,ED_ADDR1,ED_CITY,ED_PCODE,JH_JOB,JH_REPTAU,JH_REPTAU2,JH_WHRS,JH_FTENUM,SH_WHRS,SH_SALCD,SH_SALARY "
    SQLQ = SQLQ & " FROM ((HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=JH_EMPNBR AND JH_CURRENT <>0) "
    SQLQ = SQLQ & " INNER JOIN HR_SALARY_HISTORY ON HR_SALARY_HISTORY.SH_EMPNBR = ED_EMPNBR AND SH_CURRENT <>0)"
    SQLQ = SQLQ & " WHERE 1 = 1 "
    SQLQ = SQLQ & " AND " & sSQLQ & " "
    SQLQ = SQLQ & " ORDER BY ED_SURNAME,ED_FNAME "

    'Call WriteFile("SQL1=" & SQLQ)
    Dim Total As Integer
    Total = 0
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsHREmp.EOF Then
        rsHREmp.MoveFirst
        totNum = rsHREmp.RecordCount: I = 0
                
        'File to export to
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2344EmpInfoTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "EmpInfoRpt_" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait..."
        MDIMain.panHelp(0).FloodPercent = 0
        
        FileCopy xlsFileTmp, xlsFileMat
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
        exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy hh:mm")
        exSheet.Cells(2, 1) = "Time: " & Format(Now, "hh:mm")

        exSheet.Cells(4, 1) = "Employee #"
        exSheet.Cells(4, 2) = "Employee Name"
        exSheet.Cells(4, 3) = lStr("Original Hire") '"Date of Hire"
        exSheet.Cells(4, 4) = "Age"
        exSheet.Cells(4, 5) = lStr("Location")
        exSheet.Cells(4, 6) = "Section"
        exSheet.Cells(4, 7) = "Department"
        'exSheet.Cells(4, 8) = "Business Unit"
        exSheet.Cells(4, 8) = lStr("Rept. Authority 1")     '"Reporting Authority 1"
        exSheet.Cells(4, 9) = lStr("Rept. Authority 2")    '"Reporting Authority 2"
        exSheet.Cells(4, 10) = "Employment Status"
        exSheet.Cells(4, 11) = "Employment Type"
        exSheet.Cells(4, 12) = lStr("Category")     '"Active/Inactive"
        exSheet.Cells(4, 13) = lStr("Union")        '"Plant / Office"
        exSheet.Cells(4, 14) = "Position"
        exSheet.Cells(4, 15) = lStr("Hours/Week")
        exSheet.Cells(4, 16) = "Hourly Rate"
        exSheet.Cells(4, 17) = "Annual Salary"
        xRow = 5
        
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            exSheet.Cells(xRow, 1) = rsHREmp("ED_EMPNBR")
            exSheet.Cells(xRow, 2) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
            exSheet.Cells(xRow, 3) = Format(rsHREmp("ED_DOH"), "mm/dd/yyyy")
            exSheet.Cells(xRow, 4) = EmployeeAge(rsHREmp("ED_DOB"))
            exSheet.Cells(xRow, 5) = GetTABLDesc("EDLC", rsHREmp("ED_LOC"))
            exSheet.Cells(xRow, 6) = GetTABLDesc("EDSE", rsHREmp("ED_SECTION"))
            exSheet.Cells(xRow, 7) = getDeptDesc(rsHREmp("ED_DEPTNO"))
            '????exSheet.Cells(xRow, 8) = GetTABLDesc("EDLC", rsHREmp("ED_LOC")) - Business Unit
            exSheet.Cells(xRow, 8) = getEEName(rsHREmp("JH_REPTAU"))
            exSheet.Cells(xRow, 9) = getEEName(rsHREmp("JH_REPTAU2"))
            exSheet.Cells(xRow, 10) = GetTABLDesc("EDEM", rsHREmp("ED_EMP"))
            Select Case rsHREmp("ED_EMPTYPE")
                Case "0": exSheet.Cells(xRow, 11) = "0 - Not Applicable"
                Case "1": exSheet.Cells(xRow, 11) = "1 - Full Time Salary"
                Case "2": exSheet.Cells(xRow, 11) = "2 - Part Time Salary"
                Case "3": exSheet.Cells(xRow, 11) = "3 - Full Time Hourly"
                Case "4": exSheet.Cells(xRow, 11) = "4 - Part Time Hourly"
                Case "5": exSheet.Cells(xRow, 11) = "5 - Casual/Other"
                Case "6": exSheet.Cells(xRow, 11) = "6 - Contract Salary"
                Case "7": exSheet.Cells(xRow, 11) = "7 - Contract Hourly"
                Case "8": exSheet.Cells(xRow, 11) = "8 - Salary Pensioners"
                Case "9": exSheet.Cells(xRow, 11) = "9 - Salary Elected officials"
            End Select
            'exSheet.Cells(xRow, 11) = rsHREmp("ED_EMPTYPE")
            exSheet.Cells(xRow, 12) = GetTABLDesc("EDPT", rsHREmp("ED_PT"))
            exSheet.Cells(xRow, 13) = GetTABLDesc("EDOR", rsHREmp("ED_ORG"))
            exSheet.Cells(xRow, 14) = GetJobDesc(rsHREmp("JH_JOB"))
            exSheet.Cells(xRow, 15) = rsHREmp("SH_WHRS")
            
            'Hourly Rate
            If rsHREmp("SH_SALCD") = "H" Then
                exSheet.Cells(xRow, 16) = Round2DEC(IIf(Not IsNull(rsHREmp("SH_SALARY")), rsHREmp("SH_SALARY"), 0))
            Else
                'Compute Hourly Rate
                If Not IsNull(rsHREmp("SH_WHRS")) And rsHREmp("SH_WHRS") <> 0 Then
                    exSheet.Cells(xRow, 16) = Round2DEC((IIf(Not IsNull(rsHREmp("SH_SALARY")), rsHREmp("SH_SALARY"), 0) / Val(rsHREmp("SH_WHRS"))) / 52)
                Else
                    exSheet.Cells(xRow, 16) = Round2DEC(0)
                End If
            End If
            
            'Annual Salary
            If rsHREmp("SH_SALCD") = "A" Then
                exSheet.Cells(xRow, 17) = Round2DEC(IIf(Not IsNull(rsHREmp("SH_SALARY")), rsHREmp("SH_SALARY"), 0))
            Else
                'Compute Annual Salary
                If Not IsNull(rsHREmp("SH_WHRS")) Then
                    exSheet.Cells(xRow, 17) = Round2DEC(IIf(Not IsNull(rsHREmp("SH_SALARY")), rsHREmp("SH_SALARY"), 0) * rsHREmp("SH_WHRS") * 52)
                End If
            End If
            
            'exSheet.Cells(xRow, 10) = getDivDesc(rsHREmp("ED_DIV"))  'GetTABLDesc("EDSE", rsHREmp("ED_SECTION"))
            'Ticket #28037 - Jerry wants code only to save some space.
            'GetTABLDesc("EDAB", rsHREmp("ED_ADMINBY"))
            'exSheet.Cells(xRow, 11) = GetTABLDesc("EDPT", rsHREmp("ED_PT"))
            'exSheet.Cells(xRow, 12) = GetTABLDesc("EDEM", rsHREmp("ED_EMP"))
            
            rsHREmp.MoveNext
            
            xRow = xRow + 1
        Loop
        
        'exSheet.AutoFilterMode = True
        'exSheet.Range("K1:K20").AutoFilter Field:=1
        exSheet.Columns.AutoFit

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
    

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
        
    Else
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        MsgBox "No employees found in this selection criteria."
    End If
    
    rsHREmp.Close
    Set rsHREmp = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Err_Cascade_EmployeeInfo_ExcelReport:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Employee Info. Excel", fglbEmpTable, "SELECT")

Set exSheet = Nothing
Set exBook = Nothing
Set exApp = Nothing

End Sub

Private Function EmployeeAge(DOB)
Dim birthdate
Dim Age As Double

    EmployeeAge = ""
    
    If IsDate(DOB) Then
        birthdate = CVDate(DOB)
        
        Age = DateDiff("m", birthdate, Now)
        
        If month(birthdate) = month(Now) Then
            If Day(Now) < Day(birthdate) Then
                Age = Age - 1
            End If
        End If
        
        Age = CDbl(Age / 12)

        EmployeeAge = Format(Age, "#0.0")
    End If
End Function
