VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmROvtBank 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Overtime Bank Report"
   ClientHeight    =   9210
   ClientLeft      =   360
   ClientTop       =   1005
   ClientWidth     =   10320
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
   ScaleHeight     =   9210
   ScaleWidth      =   10320
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
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Tag             =   "First Level of grouping records"
      Top             =   6870
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
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Tag             =   "Second level of grouping records"
      Top             =   7225
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
      Index           =   2
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Tag             =   "Third level of grouping records"
      Top             =   7580
      Width           =   2325
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
      Index           =   3
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Tag             =   "Final sort of records"
      Top             =   7935
      Width           =   2325
   End
   Begin VB.TextBox txtBurden 
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
      Left            =   7560
      MaxLength       =   7
      TabIndex        =   59
      Tag             =   "00-Specific Burden Rate Desired"
      Top             =   4446
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox cmbHours 
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
      ItemData        =   "FZOvtBank.frx":0000
      Left            =   2355
      List            =   "FZOvtBank.frx":0002
      TabIndex        =   13
      Tag             =   "Choose display in hours or days"
      Text            =   "Combo1"
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CheckBox chkHrsExceeded 
      Caption         =   "Hours Exceeded"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   17
      Top             =   5340
      Width           =   2355
   End
   Begin VB.TextBox txtChrgCode 
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
      Left            =   7680
      MaxLength       =   15
      TabIndex        =   27
      Tag             =   "00-Enter Charge Code"
      Top             =   8760
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2355
      MaxLength       =   2
      TabIndex        =   11
      Tag             =   "00-Employee Position Shift"
      Top             =   4103
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   2
      Left            =   7245
      TabIndex        =   28
      Tag             =   "40-As of Date"
      Top             =   4103
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.EmployeeLookup elpReptau 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Tag             =   "10-Reporting Authority"
      Top             =   2731
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   8280
      TabIndex        =   24
      Tag             =   "40-Date upto and including this date forward"
      Top             =   7530
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   6480
      TabIndex        =   23
      Tag             =   "40-Date from and including this date forward"
      Top             =   7530
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Left            =   7320
      TabIndex        =   26
      Tag             =   "00-Specific Charge Code Desired"
      Top             =   8400
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   10
      Tag             =   "00-Specific Position Group Desired"
      Top             =   3760
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   3417
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   8
      Top             =   3074
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2045
      Width           =   7395
      _ExtentX        =   13044
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
      Left            =   2040
      TabIndex        =   4
      Tag             =   "00-Employee Status"
      Top             =   1702
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDEM"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1359
      Width           =   7395
      _ExtentX        =   13044
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
      Left            =   2040
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1016
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   673
      Width           =   7395
      _ExtentX        =   13044
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
      Left            =   2040
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   330
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   1
      MultiSelect     =   -1  'True
   End
   Begin VB.CheckBox chkInclAtt 
      Caption         =   "Include Attendance History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   54
      Top             =   7920
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.TextBox txtEarned 
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
      Left            =   6915
      MaxLength       =   7
      TabIndex        =   25
      Tag             =   "00-Specific Months Desired"
      Top             =   7200
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txtExxceed 
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
      Left            =   6360
      MaxLength       =   7
      TabIndex        =   29
      Tag             =   "00- O/S Hours Exceeded"
      Top             =   7980
      Visible         =   0   'False
      Width           =   870
   End
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   0
      Left            =   2085
      TabIndex        =   14
      Tag             =   "Detailed Report"
      Top             =   5415
      Width           =   1110
      _Version        =   65536
      _ExtentX        =   1958
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  Detailed"
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
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   15
      Tag             =   "Summary Report"
      Top             =   5400
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  Summary"
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
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2388
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7075
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   0
      Top             =   8325
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
      Index           =   2
      Left            =   4680
      TabIndex        =   16
      Tag             =   "Summary Report"
      Top             =   5400
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  Cost With Burden Rate"
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
   Begin VB.TextBox txtSubHeading 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2235
      MaxLength       =   50
      TabIndex        =   18
      Tag             =   "Sub heading for the report"
      Top             =   5880
      Width           =   5535
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   12
      Top             =   4446
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin VB.Label lblBurden 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Burden Rate"
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
      Left            =   5940
      TabIndex        =   60
      Top             =   4491
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Display"
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
      Left            =   120
      TabIndex        =   58
      Top             =   4860
      Width           =   510
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2235
      X2              =   7775
      Y1              =   6180
      Y2              =   6180
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Heading for Report :"
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
      Top             =   5880
      Width           =   1770
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
      TabIndex        =   56
      Top             =   4148
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
      TabIndex        =   55
      Top             =   2090
      Width           =   630
   End
   Begin VB.Label lblMonths 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Months"
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
      Left            =   8040
      TabIndex        =   53
      Top             =   7200
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblAsdate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "As of Date"
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
      Left            =   5940
      TabIndex        =   52
      Top             =   4148
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblEarned 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Earned/Not Taken Within"
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
      Left            =   4920
      TabIndex        =   51
      Top             =   7230
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Label lblPosCroup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Left            =   120
      TabIndex        =   50
      Top             =   3805
      Width           =   1035
   End
   Begin VB.Label lblOS 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "O/S Hours Exceeded"
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
      Left            =   4740
      TabIndex        =   49
      Top             =   8055
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label lblChrCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Charge Code"
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
      Left            =   5640
      TabIndex        =   48
      Top             =   8445
      Visible         =   0   'False
      Width           =   930
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
      TabIndex        =   47
      Top             =   3462
      Width           =   1125
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
      TabIndex        =   46
      Top             =   4491
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
      TabIndex        =   45
      Top             =   3119
      Width           =   510
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
      TabIndex        =   44
      Top             =   1061
      Width           =   615
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
      TabIndex        =   43
      Top             =   7995
      Width           =   660
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #3"
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
      TabIndex        =   42
      Top             =   7640
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
      TabIndex        =   41
      Top             =   7285
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
      TabIndex        =   40
      Top             =   6930
      Width           =   885
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
      TabIndex        =   39
      Top             =   6540
      Width           =   1575
   End
   Begin VB.Label lblTypeRep 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   120
      TabIndex        =   38
      Top             =   5445
      Width           =   1065
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
      Left            =   4800
      TabIndex        =   37
      Top             =   7560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblRepAuth 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reporting Authority"
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
      Top             =   2776
      Width           =   1350
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
      TabIndex        =   35
      Top             =   2433
      Width           =   1290
   End
   Begin VB.Label lblCode 
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   34
      Top             =   1747
      Width           =   450
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
      Top             =   1404
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
      Top             =   718
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
      TabIndex        =   30
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmROvtBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WRKSQL
Dim WRKSICSQL
Dim flgUseLabelMst As Boolean

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then
    If glbCompSerial = "S/N - 2425W" Then   'Ticket #18223 - Four Villages CHC
        If Not PrtForm("Extra Time Bank Report Criteria", Me) Then Exit Sub
    Else
        If Not PrtForm("Overtime Bank Report Criteria", Me) Then Exit Sub
    End If
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
     Call set_PrintState(False)
    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    'Me.vbxCrystal.Password = gstrAccPWord$
    'Me.vbxCrystal.UserName = gstrAccUID$
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If
Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
Resume Next
Screen.MousePointer = DEFAULT
End Sub

Public Sub cmdView_Click()
Dim X%
Dim strWHand As String
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
Resume Next
Screen.MousePointer = DEFAULT
End Sub

Private Sub cmbHours_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Public Sub comGrpLoad()
Dim X%

    comGroup(0).Clear
    comGroup(1).Clear
    comGroup(2).Clear
    comGroup(3).Clear

    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")      'Jaddy jun 16, 1999
    comGroup(0).AddItem "Employee Name" 'laura nov 24, 1997
    comGroup(0).AddItem lStr("Section")  'Lucy June 29,2000
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem lStr("Region")
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem lStr("Union")
    comGroup(1).AddItem "(none)"
    comGroup(1).ListIndex = 0
    
    comGroup(2).AddItem "Employee Name"
    comGroup(2).AddItem "(none)"
    comGroup(2).ListIndex = 0
    comGroup(3).AddItem "Attendance Date"
    comGroup(3).ListIndex = 0

    cmbHours.AddItem "Hours"
    cmbHours.AddItem "Days"
    cmbHours.ListIndex = 0
End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String, fld$
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    Select Case intIdx%
    Case 0: strCd$ = "HREMP.ED_LOC"
    Case 1: strCd$ = "HREMP.ED_ORG" 'IIf(glbMulti, "HR_ATTENDANCE.AD_ORG", "HREMP.ED_ORG")
    Case 2: strCd$ = "HREMP.ED_EMP"
    Case 3: strCd$ = "HREMP.ED_REGION"
    Case 4: strCd$ = "HREMP.ED_ADMINBY"
    Case 5: strCd$ = "HREMP.ED_SECTION"
    Case 6: strCd$ = "HRJOB.JB_GRPCD"
    End Select
    CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDiv & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
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

Private Sub Cri_Earned()

Dim TempCri As String
Dim TempCri01 As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%
Dim xDate1, xDate2
    'Exit Sub
    xDate1 = CVDate(dlpDateRange(2).Text)
    xDate2 = DateAdd("m", -Val(txtEarned), xDate1)
    
    TempCri = "(({HR_ATTENDANCE.AD_DOA} "
    TempCri = TempCri & " <= "
    dtYYY% = Year(xDate2)
    dtMM% = month(xDate2)
    dtDD% = Day(xDate2)
    TempCri01 = TempCri & Date_SQL(xDate2) & ") and (LEFT(HR_ATTENDANCE.AD_REASON,2) = 'OT'))"
    TempCri01 = TempCri01 & " OR "
    
    TempCri01 = TempCri01 & TempCri & Date_SQL(xDate1) & ") and (LEFT(HR_ATTENDANCE.AD_REASON,2) = 'CT'))"
    
    TempCri01 = " (" & TempCri01 & ") "
    TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) and (LEFT({HR_ATTENDANCE.AD_REASON},2) = 'OT'))"
        

    If Len(TempCri) >= 1 Then
        If Not glbiOneWhere Then
            glbstrSelCri = TempCri
            If glbOracle Then
                WRKSQL = Replace(Replace(TempCri01, "LEFT", "SUBSTR"), ",2", ",1,2")
            Else
                WRKSQL = TempCri01
            End If
        Else
            glbstrSelCri = glbstrSelCri & " AND " & TempCri
            If glbOracle Then
                WRKSQL = WRKSQL & " AND " & Replace(Replace(TempCri01, "LEFT", "SUBSTR"), ",2", ",1,2")
            Else
                WRKSQL = WRKSQL & " AND " & TempCri01
            End If
        End If
        glbiOneWhere = True
    End If

End Sub


Private Sub Cri_FTDates()

Dim TempCri As String
Dim TempCri01 As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HR_ATTENDANCE.AD_DOA} "
    TempCri01 = "{HR_ATTENDANCE.AD_DOA} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    TempCri01 = "(" & TempCri01 & " >= " & Date_SQL(dlpDateRange(0).Text) & " AND " & TempCri01 & " <= " & Date_SQL(dlpDateRange(1).Text) & ") "
    GoTo Cri_FTDatst
End If

For X% = 0 To 1
    If Len(dlpDateRange(X%).Text) > 0 Then
        TempCri = "({HR_ATTENDANCE.AD_DOA}  "
        If X% = 0 Then
            TempCri = TempCri & " >= "
            
        Else
            TempCri = TempCri & " <= "
        End If
        dtYYY% = Year(dlpDateRange(X%).Text)
        dtMM% = month(dlpDateRange(X%).Text)
        dtDD% = Day(dlpDateRange(X%).Text)
        TempCri01 = TempCri & Date_SQL(dlpDateRange(X%).Text) & ")"
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        
        GoTo Cri_FTDatst
    End If
Next X%



Cri_FTDatst:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
        WRKSQL = TempCri01
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
        WRKSQL = WRKSQL & " AND " & TempCri01
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

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, X%

If Len(txtShift.Text) < 1 Then Exit Sub
EECri = "{HR_ATTENDANCE.AD_SHIFT}= '" & txtShift.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Sub Cri_Reptau()
Dim EECri As String

If Len(elpReptau.Text) > 0 Then
    EECri = "{HR_JOB_HISTORY.JH_REPTAU} = " & getEmpnbr(elpReptau.Text) & " "
End If


If Len(EECri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
        WRKSQL = WRKSQL & " AND " & EECri
    Else
        glbstrSelCri = EECri
        WRKSQL = EECri
    End If
    glbiOneWhere = True
End If

End Sub

Private Function Cri_SetAll()
Dim X%, xNoFile
Dim strSFormat$

Cri_SetAll = False

On Error GoTo modSetCriteria_Err

Screen.MousePointer = HOURGLASS

flgUseLabelMst = True
glbiOneWhere = False
glbstrSelCri = ""
WRKSQL = ""

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_EE
Call Cri_PT
Call Cri_Shift
For X% = 0 To 6 ' 5
    Call Cri_Code(X%)
Next X%
WRKSICSQL = Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")")
WRKSQL = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"

Call Cri_ChargeCode
Call Cri_Position 'Charge Code with Lookup
Call Cri_Reptau
'Call Cri_Code(6)

Call Cri_FTDates
If Len(txtEarned) > 0 Then
    Call Cri_Earned
End If

WRKSQL = Replace(WRKSQL, "{", "(")
WRKSQL = Replace(WRKSQL, "}", ")")
WRKSQL = Replace(WRKSQL, "HR_ATTENDANCE.", "")
    
    If Len(Trim(txtEarned)) > 0 Then 'Eraned/Not Taken Within xx Months Entered
        glbstrSelCri = glbstrSelCri & " AND " & "{HRATTWRK.AD_WRKEMP}= '" & glbUserID & "'"
        
        Call SELATTWRK
    Else
        ' danielk - 03/13/2002 - Added check for Burden Rate, if that's selected we use a different workfile
        '   that doesn't have a WRKEMP field.
      'COMMENTED BY SAM AS JERRY DOES NOT WANT TO USE WORKING TABLE 07/06/2006
'        If optGrouping(2).Value = False Then glbstrSelCri = "{HRCOMPTIME.CT_WRKEMP}= '" & glbUserID & "'"
'
'        If optGrouping(2) Then
'            Call CompTimeWRKCost
'        Else
'            Call CompTimeWRK
'        End If
        ' danielk - 12/02/2002 - Ticket #2686 - 7.0 Priority C changes - wasn't working, added pause, fixed now
        If Not glbSQL And Not glbOracle Then Pause (1)
    End If

xNoFile = 5 + IIf(glbMulti, 1, 0)
If optGrouping(0) Then
    'Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "comptd.rpt"
    If chkHrsExceeded Then
        If cmbHours.ListIndex = 0 Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzOvtBnkExc.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzOvtBnkExcD.rpt"
        End If
    Else
        If cmbHours.ListIndex = 0 Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzOvtBnk.rpt"
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzOvtBnkD.rpt"
        End If
    End If
Else
    flgUseLabelMst = False
    If optGrouping(1) Then
        If comGroup(0) <> "(none)" Then
            'Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "compts.rpt"
            If chkHrsExceeded Then
                If cmbHours.ListIndex = 0 Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzOTBnksExc.rpt"
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzOTBnksExcD.rpt"
                End If
            Else
                If cmbHours.ListIndex = 0 Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzOTBnks.rpt"
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzOTBnksD.rpt"
                End If
            End If
        Else
            'Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "comps1.rpt"   '8
            
            If chkHrsExceeded Then
                If cmbHours.ListIndex = 0 Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzOvtBExc.rpt"
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzOvtBExcD.rpt"
                End If
            Else
                If cmbHours.ListIndex = 0 Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzOvtBs1.rpt"   '8
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzOvtBs1D.rpt"   '8
                End If
            End If
            
        End If
    End If
End If

If Len(Trim(txtEarned)) > 0 Then 'Earned/Not Taken Within xx Months Entered
    flgUseLabelMst = False
    
    If cmbHours.ListIndex = 0 Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZOTBkea.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZOTBkeaD.rpt"
    End If
    
    glbstrSelCri = Replace(glbstrSelCri, "HR_ATTENDANCE", "HRATTWRK")
End If

If Not optGrouping(2) Then 'Add by Franks Mar 13,2002
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

If Len(Trim(txtSubHeading)) > 0 Then
    Me.vbxCrystal.Formulas(2) = "SubHead = '" & txtSubHeading & "'"
End If

If optGrouping(2) Then 'For glbCBrant
    If dlpDateRange(0).Text <> "" And dlpDateRange(1).Text <> "" Then
        strSFormat$ = "As of " & dlpDateRange(0).Text & " through " & dlpDateRange(1).Text
        Me.vbxCrystal.Formulas(0) = "Daterange = '" & strSFormat$ & "'"
    Else
        strSFormat$ = "No date entered"
        Me.vbxCrystal.Formulas(0) = "Daterange = '" & strSFormat$ & "'"
    End If
    Me.vbxCrystal.Connect = "PWD=petman;"
    Me.vbxCrystal.DataFiles(0) = glbIHRDBW
    Me.vbxCrystal.DataFiles(1) = glbIHRDB
    Me.vbxCrystal.WindowTitle = "Overtime Bank Report"
    Cri_SetAll = True
    Screen.MousePointer = DEFAULT
    Exit Function
End If

X% = Cri_Sorts()
'If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
'Else
'    Me.vbxCrystal.Connect = "PWD=petman;"
'    If Not (Len(Trim(txtEarned)) > 0) Then 'Eraned/Not Taken Within xx Months Entered
'        For x% = 0 To xNoFile + 1
'            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
'        Next x%
'        Me.vbxCrystal.DataFiles(x%) = glbIHRDBW
'    Else
'        Me.vbxCrystal.DataFiles(0) = glbIHRDBW
'        For x% = 1 To xNoFile + 1
'            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
'        Next x%
'    End If
'End If

'For x = 0 To 1000
'    If Me.vbxCrystal.Formulas(x) <> "" Then Debug.Print Me.vbxCrystal.Formulas(x)
'Next

If glbCompSerial = "S/N - 2425W" Then   'Ticket #18223 - Four Villages CHC
    Me.vbxCrystal.WindowTitle = "Extra Time Bank Report"
Else
    Me.vbxCrystal.WindowTitle = "Overtime Bank Report"
End If
Cri_SetAll = True
Screen.MousePointer = DEFAULT
Exit Function


modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Overtime Bank Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

Cri_Sorts = 0

'Report label setting
If flgUseLabelMst = True Then
    Call setRptLabel(Me, 4)
End If

' first set primary grouping
grpField$ = getEGroup(comGroup(0).Text)
If grpField$ = "(none)" Then grpField$ = "{HREMP.ED_COMPNO}"

grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(0) = grpCond$

If Not (grpField$ = "{HREMP.ED_COMPNO}") Then 'GrpIdx% <> 5 Then     ' IF Not (none) - SHOW GROUP1  'laura nov 24, 1997 changed from 2 to 3
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup1 = '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    dscGroup$ = "descGroup2 = " & grpField$
    Me.vbxCrystal.Formulas(1) = dscGroup$
    Me.vbxCrystal.SectionFormat(1) = "GF3;T;F;F;X;X;X;X"
    Me.vbxCrystal.SectionFormat(0) = "GH1;T;T;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(2) = "GF1;T;X;F;X;X;X;X"
Else
    dscGroup$ = "descGroup1 = ''"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    dscGroup$ = "descGroup2 = ''"
    Me.vbxCrystal.Formulas(1) = dscGroup$
    Me.vbxCrystal.SectionFormat(0) = "GH1;F;T;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(2) = "GF1;F;X;F;X;X;X;X"
End If

GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@EFullName}"
    Case 1: grpField$ = "{HREMP.ED_ORG}"
    Case 2: grpField$ = "{@EFullName}"    '(none)
End Select

grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$

GrpIdx% = comGroup(2).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@EFullName}"
    Case 1: grpField$ = "{@EFullName}"    '(none)
End Select
        
grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(2) = grpCond$

If comGroup(1).ListIndex > 0 Then
    Me.vbxCrystal.SectionFormat(0) = "GF1;F;X;X;X;X;X;X"
End If

Cri_Sorts = z%
End Function

Private Function CriCheck()
Dim I%

CriCheck = False

For I% = 0 To 6 '5
    If Not clpCode(I).ListChecker Then Exit Function
Next I%

If Not clpDiv.ListChecker Then Exit Function

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known"
    'clpDept.SetFocus
    Exit Function
End If
'If Len(txtEEID) > 0 And lblEEName.Caption = "Enter Valid Employee #" Then
'    MsgBox "If Employee Entered - they must exist"
'    txtEEID.SetFocus
'    Exit Function
'End If


For I% = 0 To 2 '1
    If Len(dlpDateRange(I%).Text) > 0 Then
        If Not IsDate(dlpDateRange(I%).Text) Then
            MsgBox "Not a valid date"
            dlpDateRange(I%).Text = ""
            dlpDateRange(I%).SetFocus
            Exit Function
        End If
    End If
Next I%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

'If Len( clpJob) > 0 And lblJobDesc = "Unassigned" Then
'    MsgBox "If Charge Code Entered - it must be known"
'     clpJob.SetFocus
'    Exit Function
'End If

If Len(txtExxceed) > 0 And Not IsNumeric(txtExxceed) Then
    MsgBox "If O/S Hours Exceeded Entered - it must be number"
    txtExxceed.SetFocus
    Exit Function
End If

If Len(txtBurden) > 0 And Not IsNumeric(txtBurden) Then
    MsgBox "If Burden Rate Entered - it must be number"
    txtBurden.SetFocus
    Exit Function
End If

If Len(txtEarned) > 0 And Not IsNumeric(txtEarned) Then
    MsgBox "If Earned/Not Taken Within Entered - it must be number"
    txtEarned.SetFocus
    Exit Function
End If
If Len(txtEarned) > 0 Then
    If Len(dlpDateRange(2).Text) = 0 Then
        MsgBox "If Earned/Not Taken Within Entered - As of Date is required"
        dlpDateRange(2).SetFocus
        Exit Function
    End If
End If
If Not elpEEID.ListChecker Then Exit Function

If InStr(1, txtSubHeading.Text, "'") > 0 Then
    MsgBox "Sub Heading for Report cannot contain apostrophe (')."
    txtSubHeading.SetFocus
    Exit Function
End If

CriCheck = True
End Function

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()

Screen.MousePointer = HOURGLASS

glbOnTop = "FRMROVTBANK"

If Not glbSQL And Not glbOracle Then
    Call CreateCompTime
    Call CreateCompTimeCost
End If

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call comGrpLoad

Call setRptCaption(Me)
Call setCaption(lblShift)
Call setCaption(lblChrCode)
lblFromTo.Caption = lStr("From Date") & " / " & lStr("To Date")

'If glbSQL Or glbOracle Then
'    lblChrCode.Visible = False
'    lblOS.Visible = False
'    txtExxceed.Visible = False
'    lblPosCroup.Visible = False
'    clpJob.Visible = False
'    clpCode(6).Visible = False
'End If
'If glbCompSerial = "S/N - 2217W" Or glbCompSerial = "S/N - 2214W" Or glbCompSerial = "S/N - 2241W" Then
'    clpJob.Visible = True
'    clpJob.LookupType = Department
'    clpJob.MaxLength = 7
'    txtChrgCode.Visible = False
'ElseIf glbCompSerial = "S/N - 2192W" Then  ' county essex
'    txtChrgCode.Visible = False
'    clpJob.Visible = True
'    clpJob.LookupType = GL
'    clpJob.MaxLength = 25
'ElseIf glbCBrant Then
'    clpJob.Visible = True
'    clpJob.LookupType = Job
'    clpJob.MaxLength = 6
'    txtChrgCode.Visible = False
'End If

'If glbCompSerial = "S/N - 2241W" Then
'    lblChrCode.Caption = "Department"
'    clpJob.Tag = "00-Specific Department Desired"
'End If

If glbCBrant Then
    lblBurden.Visible = True
    txtBurden.Visible = True
Else
    optGrouping(2).Visible = False
    chkHrsExceeded.Left = 5400
End If
If glbLinamar Then clpCode(3).MaxLength = 8
If glbSyndesis Then
    lblPosCroup.Caption = "Position Grade"
    clpCode(6).Tag = "00-Specific Position Grade Desired"
End If
If glbMulti Then lblUnion.ForeColor = &HC000C0
If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

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

Private Sub optGrouping_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtBurden_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtChrgCode_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtEarned_Change()
    If Len(txtEarned) > 0 And IsNumeric(txtEarned) Then
        lblAsdate.Visible = True
        dlpDateRange(2).Visible = True
        optGrouping(0).Value = True
        optGrouping(1).Visible = False
        If glbCBrant Then
            optGrouping(2).Visible = False
            chkHrsExceeded.Left = 5400
        End If
    Else
        lblAsdate.Visible = False
        dlpDateRange(2).Text = ""
        dlpDateRange(2).Visible = False
        optGrouping(1).Visible = True
        If glbCBrant Then
            optGrouping(2).Visible = True
        End If
    End If
End Sub

Private Sub txtEarned_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub


Private Sub txtExxceed_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub


Private Sub Cri_ChargeCode()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim PosCri As String
If Len(txtChrgCode.Text) <= 0 Then Exit Sub
PosCri = "({HR_ATTENDANCE.AD_CHRGCODE} = '" & txtChrgCode.Text & "')"
If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & PosCri
    WRKSQL = WRKSQL & " AND " & PosCri
Else
    glbstrSelCri = PosCri
    WRKSQL = PosCri
End If
End Sub

Private Sub Cri_Position()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim PosCri As String
If Len(clpJob.Text) <= 0 Then Exit Sub
'PosCri = "({HR_JOB_HISTORY.JH_JOB} = '" &  clpJob.Text & "')"
PosCri = "({HR_ATTENDANCE.AD_CHRGCODE} = '" & clpJob.Text & "')"
If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & PosCri
    WRKSQL = WRKSQL & " AND " & PosCri
Else
    glbstrSelCri = PosCri
    WRKSQL = PosCri
End If

End Sub

Private Sub Cri_Exceed()
Dim countr   As Integer
Dim ExceedCri As String
If Len(txtExxceed.Text) <= 0 Then Exit Sub
ExceedCri = "({HR_ATTENDANCE.AD_HRS} > " & txtExxceed.Text & ")"
If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & ExceedCri
Else
    glbstrSelCri = ExceedCri
End If
End Sub

Private Sub CompTimeWRKCost()
Dim RsCompTime As New ADODB.Recordset
Dim RsATTC As New ADODB.Recordset
Dim rsPARCO As New ADODB.Recordset
Dim SQLQ, xEmpnbrN, xEmpnbrO, xEarned, xTaken, xBalance
Dim xDECHR
Dim rsOvtBank As New ADODB.Recordset
Dim OEmpNbr

xDECHR = 2
rsPARCO.Open "HRPARCO", gdbAdoIhr001, adOpenStatic
If (Not rsPARCO.EOF) And (Not rsPARCO.BOF) Then
    xDECHR = rsPARCO("PC_DECHR")
    If xDECHR < 2 Then xDECHR = 2
End If


SQLQ = "DELETE FROM HRCOMPTIME " & in_SQL(glbIHRDBW) & " WHERE CT_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

' danielk - 12/02/2002 - 7.0 Priority C changes
'Frank Aug 7,2002 T#2686
On Error Resume Next
SQLQ = "DELETE * FROM HRCOMPCOST"
gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute SQLQ
gdbAdoIhr001W.CommitTrans
On Error GoTo 0
'Frank Aug 7,2002 T#2686

If glbOracle Then
        SQLQ = "SELECT SH_EMPNBR,SH_SALARY as SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
        SQLQ = SQLQ & " AD_CHRGCODE,AD_DOA,AD_HRS,AD_REASON,JB_GRPCD "
        If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
            SQLQ = SQLQ & ", JB_FTEHRS, ED_ORG, ED_PT, ED_EMP, ED_LOC, ED_REGION, ED_ADMINBY, ED_SECTION "
        End If
        SQLQ = SQLQ & " FROM HR_SALARY_HISTORY,HR_JOB_HISTORY,HREMP,HRJOB,HR_ATTENDANCE "
        SQLQ = SQLQ & " where HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR"
        
        If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
            SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_SALARY_HISTORY.SH_PAYROLL_ID = HR_JOB_HISTORY.JH_PAYROLL_ID "
        End If
        
        SQLQ = SQLQ & " and HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " and HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
        SQLQ = SQLQ & " and HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR "
        SQLQ = SQLQ & " and JH_CURRENT<>0 AND SH_CURRENT<>0 "
        SQLQ = SQLQ & " AND (substr(AD_REASON,1,2) = 'OT' ) "
        If Len(WRKSQL) > 0 Then SQLQ = SQLQ & "AND " & WRKSQL
        If chkInclAtt Then
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT SH_EMPNBR,SH_SALARY as SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
            SQLQ = SQLQ & " AH_CHRGCODE AS AD_CHRGCODE,AH_DOA AS AD_DOA,AH_HRS AS AD_HRS,AH_REASON AS AD_REASON,JB_GRPCD "
            If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                SQLQ = SQLQ & ", JB_FTEHRS, ED_ORG, ED_LOC, ED_REGION, ED_ADMINBY, ED_SECTION "
            End If
            SQLQ = SQLQ & " FROM HR_SALARY_HISTORY, HR_JOB_HISTORY,HREMP,HRJOB,HR_ATTENDANCE_HISTORY "
            SQLQ = SQLQ & " where HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
            SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR)"
            
            If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_SALARY_HISTORY.SH_PAYROLL_ID = HR_JOB_HISTORY.JH_PAYROLL_ID "
            End If
            
            SQLQ = SQLQ & " and HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR "
            SQLQ = SQLQ & " and HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
            SQLQ = SQLQ & " and HREMP.ED_EMPNBR = HR_ATTENDANCE_HISTORY.AH_EMPNBR "
            SQLQ = SQLQ & " and JH_CURRENT<>0 AND SH_CURRENT<>0 "
            SQLQ = SQLQ & " AND (substr(AH_REASON,1,2) = 'OT' ) "
            If Len(WRKSQL) > 0 Then SQLQ = SQLQ & "AND " & Replace(WRKSQL, "AD_", "AH_")
        End If
        SQLQ = SQLQ & " ORDER BY SH_EMPNBR,AD_DOA "

Else

        SQLQ = "SELECT SH_EMPNBR,SH_SALARY as SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
        SQLQ = SQLQ & " AD_CHRGCODE,AD_DOA,AD_HRS,AD_REASON,JB_GRPCD "
        If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
            SQLQ = SQLQ & ", JB_FTEHRS, ED_ORG, ED_LOC, ED_REGION, ED_ADMINBY, ED_SECTION "
        End If
        SQLQ = SQLQ & " FROM (((HR_SALARY_HISTORY INNER JOIN HR_JOB_HISTORY "
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
        If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
            SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_SALARY_HISTORY.SH_PAYROLL_ID = HR_JOB_HISTORY.JH_PAYROLL_ID)"
        Else
            SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR)"
        End If
        SQLQ = SQLQ & " INNER JOIN HREMP ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR)"
        SQLQ = SQLQ & " INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE)"
        SQLQ = SQLQ & " INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR "
        SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND SH_CURRENT<>0 "
        SQLQ = SQLQ & " AND (LEFT(AD_REASON,2) = 'OT' ) "
        If Len(WRKSQL) > 0 Then SQLQ = SQLQ & "AND " & WRKSQL
        If chkInclAtt Then
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT SH_EMPNBR,SH_SALARY as SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
            SQLQ = SQLQ & " AH_CHRGCODE AS AD_CHRGCODE,AH_DOA AS AD_DOA,AH_HRS AS AD_HRS,AH_REASON AS AD_REASON,JB_GRPCD "
            If glbCompSerial = "S/N - 2378W" Then   'Hemu  - Town of Aurora
                SQLQ = SQLQ & ", JB_FTEHRS, ED_ORG, ED_PT, ED_EMP, ED_LOC, ED_REGION, ED_ADMINBY, ED_SECTION "
            End If
            SQLQ = SQLQ & " FROM (((HR_SALARY_HISTORY INNER JOIN HR_JOB_HISTORY "
            SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
            If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_SALARY_HISTORY.SH_PAYROLL_ID = HR_JOB_HISTORY.JH_PAYROLL_ID)"
            Else
                SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR)"
            End If
            SQLQ = SQLQ & " INNER JOIN HREMP ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR)"
            SQLQ = SQLQ & " INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE)"
            SQLQ = SQLQ & " INNER JOIN HR_ATTENDANCE_HISTORY ON HREMP.ED_EMPNBR = HR_ATTENDANCE_HISTORY.AH_EMPNBR "
            SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND SH_CURRENT<>0 "
            SQLQ = SQLQ & " AND (LEFT(AH_REASON,2) = 'OT' ) "
            If Len(WRKSQL) > 0 Then SQLQ = SQLQ & "AND " & Replace(WRKSQL, "AD_", "AH_")
        End If
        SQLQ = SQLQ & " ORDER BY SH_EMPNBR,AD_DOA "

End If
RsATTC.Open SQLQ, gdbAdoIhr001, adOpenStatic

RsCompTime.Open "SELECT * FROM HRCOMPTIME WHERE CT_WRKEMP='" & glbUserID & "' ORDER BY CT_EMPNBR,CT_DOA", gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
OEmpNbr = 0

Do While Not RsATTC.EOF
    '-----------------------------------------------------------------------------------------
    'Add Previous year brought forward record
    If OEmpNbr <> RsATTC("SH_EMPNBR") Then
        OEmpNbr = RsATTC("SH_EMPNBR")
        SQLQ = "SELECT * FROM HR_OVERTIME_BANK WHERE OT_EMPNBR = " & RsATTC("SH_EMPNBR")
        rsOvtBank.Open SQLQ, gdbAdoIhr001, adOpenKeyset
        If Not rsOvtBank.EOF Then
            RsCompTime.AddNew
            RsCompTime("CT_EMPNBR") = RsATTC("SH_EMPNBR")
            RsCompTime("CT_REPTAU") = RsATTC("JH_REPTAU")
            RsCompTime("CT_DOA") = rsOvtBank("OT_EFDATE")
            RsCompTime("CT_REASON") = "OTBF"
            RsCompTime("CT_HRS") = rsOvtBank("OT_PBANK")
            RsCompTime("CT_DHRS") = RsATTC("JH_DHRS")
            If RsATTC("SH_SALCD") = "A" Then
                If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                    RsCompTime("CT_HRSSAL") = Round((RsATTC("SH_SAL") / RsATTC("JB_FTEHRS")), xDECHR) '4)
                Else
                    RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
                End If
            ElseIf RsATTC("SH_SALCD") = "H" Then
                RsCompTime("CT_HRSSAL") = RsATTC("SH_SAL")
            ElseIf RsATTC("SH_SALCD") = "M" Then
                If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                    RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") * 12) / RsATTC("JB_FTEHRS")), xDECHR)  '4)
                Else
                    RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 12) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
                End If
            ElseIf RsATTC("SH_SALCD") = "D" Then
                If GetLeapYear(Year(Date)) Then
                    If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                        RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") * 366) / RsATTC("JB_FTEHRS")), xDECHR)  '4)
                    Else
                        RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 366) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
                    End If
                Else
                    If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                        RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") * 365) / RsATTC("JB_FTEHRS")), xDECHR)  '4)
                    Else
                        RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 365) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
                    End If
                End If
            End If
    
            RsCompTime("CT_GRPCD") = RsATTC("JB_GRPCD")
            RsCompTime("CT_WRKEMP") = glbUserID
            RsCompTime.Update
        End If
        rsOvtBank.Close
    End If
    '-----------------------------------------------------------------------------------------


    RsCompTime.AddNew
    RsCompTime("CT_EMPNBR") = RsATTC("SH_EMPNBR")
    RsCompTime("CT_REPTAU") = RsATTC("JH_REPTAU")
    RsCompTime("CT_DOA") = RsATTC("AD_DOA")
    RsCompTime("CT_REASON") = RsATTC("AD_REASON")
    If Mid(RsATTC("AD_REASON"), 1, 2) = "OT" And glbCompSerial = "S/N - 2378W" Then 'Town of Aurora
        RsCompTime("CT_HRS") = RsATTC("AD_HRS") * Overtime_Multiplier(RsATTC("ED_ORG"), RsATTC("ED_PT"), RsATTC("ED_EMP"), RsATTC("ED_LOC"), RsATTC("ED_REGION"), RsATTC("ED_ADMINBY"), RsATTC("ED_SECTION"))
    Else
        RsCompTime("CT_HRS") = RsATTC("AD_HRS")
    End If
    RsCompTime("CT_DHRS") = RsATTC("JH_DHRS")
    RsCompTime("CT_TEXT1") = RsATTC("AD_CHRGCODE")
    If RsATTC("SH_WHRS") > 0 Then
        If RsATTC("SH_SALCD") = "A" Then
            If glbCompSerial = "S/N - 2378W" Then   'Hemu   - Town of Aurora
                RsCompTime("CT_HRSSAL") = Round((RsATTC("SH_SAL") / RsATTC("JB_FTEHRS")), xDECHR)  '4)
            Else
                RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
            End If
        ElseIf RsATTC("SH_SALCD") = "H" Then
            RsCompTime("CT_HRSSAL") = RsATTC("SH_SAL")
        ElseIf RsATTC("SH_SALCD") = "M" Then
            If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") * 12) / RsATTC("JB_FTEHRS")), xDECHR)  '4)
            Else
                RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 12) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
            End If
        ElseIf RsATTC("SH_SALCD") = "D" Then
            If GetLeapYear(Year(Date)) Then
                If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                    RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") * 366) / RsATTC("JB_FTEHRS")), xDECHR)  '4)
                Else
                    RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 366) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
                End If
            Else
                If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                    RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") * 365) / RsATTC("JB_FTEHRS")), xDECHR)  '4)
                Else
                    RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 365) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
                End If
            End If
        End If

    End If
    If RsCompTime("CT_HRSSAL") > 0 And RsCompTime("CT_HRS") > 0 Then
        RsCompTime("CT_NUM1") = Round((RsCompTime("CT_HRS") * RsCompTime("CT_HRSSAL")), xDECHR)
    End If
    RsCompTime("CT_GRPCD") = RsATTC("JB_GRPCD")
    RsCompTime("CT_WRKEMP") = glbUserID
        
    RsCompTime.Update
    RsATTC.MoveNext
Loop
RsATTC.Close
RsCompTime.Close

'Hemu - 01/22/2004 Begin
SQLQ = "SELECT CT_TEXT1 AS CHRGCODE, SUM(CT_NUM1) AS EXTENDC FROM HRCOMPTIME " 'GROUP BY CT_TEXT1 "
'SQLQ = "SELECT CT_TEXT1 AS CHRGCODE, SUM(CT_BALANCE) AS EXTENDC FROM HRCOMPTIME " 'GROUP BY CT_TEXT1 "
'Hemu - 01/22/2004 End

SQLQ = SQLQ & "WHERE CT_WRKEMP = '" & glbUserID & "' GROUP BY CT_TEXT1 "
'Frank Aug 7,2002 T#2686

RsCompTime.Open SQLQ, gdbAdoIhr001W, adOpenStatic
RsATTC.Open "HRCOMPCOST", gdbAdoIhr001W, adOpenKeyset, adLockOptimistic, adCmdTableDirect
Dim xEXTENDED, xBURDEN
Do While Not RsCompTime.EOF
    If RsCompTime("EXTENDC") > 0 Then
        xEXTENDED = RsCompTime("EXTENDC")
    Else
        xEXTENDED = 0
    End If
    If Len(Trim(txtBurden)) > 0 Then
        xBURDEN = Val(txtBurden)
    Else
        xBURDEN = 0
    End If
    RsATTC.AddNew
    RsATTC("CT_COMPNO") = "001"
    RsATTC("CT_CHRGCODE") = RsCompTime("CHRGCODE")
    RsATTC("CT_EXTEND") = xEXTENDED
    RsATTC("CT_BURDEN") = xBURDEN
    ' dkostka - 11/13/2001 - Changed from adding burden to extended, to adding burden%, like it was supposed to be.
    RsATTC("CT_TOTAL") = xEXTENDED * (1 + xBURDEN / 100)
    RsATTC("CT_WRKEMP") = glbUserID
    RsATTC.Update
    RsCompTime.MoveNext
Loop
RsATTC.Close
RsCompTime.Close
End Sub

Private Sub CompTimeWKsick()
Dim RsCompTime As New ADODB.Recordset
Dim RsATTC As New ADODB.Recordset
Dim rsPARCO As New ADODB.Recordset
Dim SQLQ, xEmpnbrN, xEmpnbrO, xEarned, xTaken, xBalance
Dim xDECHR

xDECHR = 2
rsPARCO.Open "HRPARCO", gdbAdoIhr001, adOpenStatic
If (Not rsPARCO.EOF) And (Not rsPARCO.BOF) Then
    xDECHR = rsPARCO("PC_DECHR")
    If xDECHR < 2 Then xDECHR = 2
End If

SQLQ = "DELETE FROM HRCOMPTIME " & in_SQL(glbIHRDBW) '& " WHERE CT_WRKEMP='" & glbUserID & "'"

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

'Get Sick Entitlement from HREMP - Begin
SQLQ = "SELECT HREMP.ED_EMPNBR,HREMP.ED_EFDATES,HREMP.ED_PSICK,HREMP.ED_SICK,HR_JOB_HISTORY.JH_DHRS FROM ((HREMP "
SQLQ = SQLQ & " INNER JOIN HR_JOB_HISTORY ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR) INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR) "
SQLQ = SQLQ & "WHERE " & WRKSICSQL & " "
SQLQ = SQLQ & "AND JH_CURRENT<>0 "

RsATTC.Open SQLQ, gdbAdoIhr001, adOpenStatic

RsCompTime.Open "SELECT * FROM HRCOMPTIME WHERE CT_WRKEMP='" & glbUserID & "' ORDER BY CT_EMPNBR,CT_DOA", gdbAdoIhr001W, adOpenKeyset, adLockOptimistic  ' , adLockBatchOptimistic
Do While Not RsATTC.EOF
    RsCompTime.AddNew
    RsCompTime("CT_EMPNBR") = RsATTC("ED_EMPNBR")
    RsCompTime("CT_DOA") = RsATTC("ED_EFDATES")
    RsCompTime("CT_REASON") = "OT" 'RsATTC("AD_REASON")
    RsCompTime("CT_HRS") = RsATTC("ED_PSICK") + RsATTC("ED_SICK")
    RsCompTime("CT_DHRS") = RsATTC("JH_DHRS")
    RsCompTime("CT_WRKEMP") = glbUserID
    RsCompTime.Update
    RsATTC.MoveNext 'CT_HRSSAL
Loop
RsATTC.Close
RsCompTime.Requery
'Get Sick Entitlement from HREMP - End


SQLQ = "SELECT SH_EMPNBR,SH_SALARY as SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
SQLQ = SQLQ & " AD_DOA,AD_HRS,AD_REASON,JB_GRPCD "
SQLQ = SQLQ & " FROM (((HR_SALARY_HISTORY INNER JOIN HR_JOB_HISTORY "
SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR)"
SQLQ = SQLQ & " INNER JOIN HREMP ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR)"
SQLQ = SQLQ & " INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE)"
SQLQ = SQLQ & " INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR "
SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND SH_CURRENT<>0 "
SQLQ = SQLQ & " AND (LEFT(AD_REASON,3) = 'SIC')"
If Len(WRKSQL) > 0 Then SQLQ = SQLQ & "AND " & WRKSQL
If chkInclAtt Then
    SQLQ = SQLQ & " UNION "
    SQLQ = SQLQ & " SELECT SH_EMPNBR,SH_SALARY as SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
    SQLQ = SQLQ & " AH_DOA AS AD_DOA,AH_HRS AS AD_HRS,AH_REASON AS AD_REASON,JB_GRPCD "
    SQLQ = SQLQ & " FROM (((HR_SALARY_HISTORY INNER JOIN HR_JOB_HISTORY "
    SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
    SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR)"
    SQLQ = SQLQ & " INNER JOIN HREMP ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR)"
    SQLQ = SQLQ & " INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE)"
    SQLQ = SQLQ & " INNER JOIN HR_ATTENDANCE_HISTORY ON HREMP.ED_EMPNBR = HR_ATTENDANCE_HISTORY.AH_EMPNBR "
    SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND SH_CURRENT<>0 "
    SQLQ = SQLQ & " AND (LEFT(AH_REASON,3) = 'SIC')"
    If Len(WRKSQL) > 0 Then SQLQ = SQLQ & " AND " & Replace(WRKSQL, "AD_", "AH_")
End If
SQLQ = SQLQ & " ORDER BY SH_EMPNBR,AD_DOA "

'added by Bryan 22/Sep/05 Ticket#9343
If glbCompSerial = "S/N - 2373W" Then 'Muskoka
    SQLQ = Replace(SQLQ, "SH_SALARY", "SH_TOTAL", 1, , vbTextCompare)
End If
RsATTC.Open SQLQ, gdbAdoIhr001, adOpenStatic

'RsCompTime.Open "SELECT * FROM HRCOMPTIME WHERE CT_WRKEMP='" & glbUserID & "' ORDER BY CT_EMPNBR,CT_DOA", gdbAdoIhr001W, adOpenKeyset, adLockOptimistic  ' , adLockBatchOptimistic
Do While Not RsATTC.EOF
    RsCompTime.AddNew
    RsCompTime("CT_EMPNBR") = RsATTC("SH_EMPNBR")
    RsCompTime("CT_REPTAU") = RsATTC("JH_REPTAU")
    RsCompTime("CT_DOA") = RsATTC("AD_DOA")
    RsCompTime("CT_REASON") = "CT" 'RsATTC("AD_REASON")
    RsCompTime("CT_HRS") = RsATTC("AD_HRS")
    RsCompTime("CT_DHRS") = RsATTC("JH_DHRS")
    If RsATTC("SH_WHRS") > 0 Then
        If RsATTC("SH_SALCD") = "A" Then
            RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
        ElseIf RsATTC("SH_SALCD") = "H" Then
            RsCompTime("CT_HRSSAL") = RsATTC("SH_SAL")
        ElseIf RsATTC("SH_SALCD") = "M" Then
            RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 12) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
        ElseIf RsATTC("SH_SALCD") = "D" Then
            If GetLeapYear(Year(Date)) Then
                RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 366) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
            Else
                RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 365) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
            End If
        End If

    End If
    RsCompTime("CT_GRPCD") = RsATTC("JB_GRPCD")
    RsCompTime("CT_WRKEMP") = glbUserID
    RsCompTime.Update
    RsATTC.MoveNext 'CT_HRSSAL
Loop
RsATTC.Close
RsCompTime.Requery

xEmpnbrO = 0

Do While Not RsCompTime.EOF
    xEmpnbrN = RsCompTime("CT_EMPNBR")
    If xEmpnbrO <> xEmpnbrN Then
        xBalance = 0
        xEmpnbrO = xEmpnbrN
    End If
    If Not IsNull(RsCompTime("CT_REASON")) Then
        xEarned = 0
        xTaken = 0
        If Mid(RsCompTime("CT_REASON"), 1, 2) = "OT" Then
            ' dkostka - 02/27/2002 - Removed code to make negative amounts zero, not sure
            '   why this was in here, but it shouldn't be.  Checked w/ Linda, she gave
            '   the OK to remove it.
            'If RsCompTime("CT_HRS") > 0 Then
            xEarned = RsCompTime("CT_HRS")
        End If
        If Mid(RsCompTime("CT_REASON"), 1, 2) = "CT" Then
            ' dkostka - 02/27/2002 - Removed code to make negative amounts zero, not sure
            '   why this was in here, but it shouldn't be.  Checked w/ Linda, she gave
            '   the OK to remove it.
            'If RsCompTime("CT_HRS") > 0 Then
            xTaken = RsCompTime("CT_HRS")
        End If
        If IsDate(RsCompTime("CT_DOA")) Then
            RsCompTime("CT_WEEK") = WeekdayName(Weekday(RsCompTime("CT_DOA")))
        End If
        RsCompTime("CT_EARN") = xEarned
        RsCompTime("CT_TAKEN") = xTaken
        xBalance = xBalance + (xEarned - xTaken)
        RsCompTime("CT_BALANCE") = xBalance
        If RsCompTime("CT_DHRS") > 0 Then
            RsCompTime("CT_EQUDAYS") = Round(((xEarned - xTaken) / RsCompTime("CT_DHRS")), xDECHR)
        End If
        'If RsCompTime("CT_HRSSAL") > 0 Then
        '    RsCompTime("CT_EXTDCOST") = Round(((xEarned - xTaken) * RsCompTime("CT_HRSSAL")), xDECHR)
        'Else
        '    RsCompTime("CT_EXTDCOST") = 0
        'End If
        RsCompTime.Update
    End If
    RsCompTime.MoveNext
Loop
RsCompTime.Close

SQLQ = "SELECT CT_EMPNBR AS EMPNBR, SUM(CT_EARN) AS EARN,SUM(CT_TAKEN) AS TAKEN FROM HRCOMPTIME WHERE CT_WRKEMP='" & glbUserID & "' GROUP BY CT_EMPNBR"
RsCompTime.Open SQLQ, gdbAdoIhr001W, adOpenStatic
Do While Not RsCompTime.EOF
    xEmpnbrN = RsCompTime("EMPNBR")
    xBalance = RsCompTime("EARN") - RsCompTime("TAKEN")
    SQLQ = "UPDATE HRCOMPTIME " & in_SQL(glbIHRDBW) & " SET CT_BALANCETOT = " & xBalance & " WHERE CT_EMPNBR = " & xEmpnbrN & " AND CT_WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    RsCompTime.MoveNext
Loop

If Val(txtExxceed) > 0 Then
    SQLQ = "DELETE FROM HRCOMPTIME " & in_SQL(glbIHRDBW) & " WHERE CT_BALANCETOT < " & txtExxceed & " AND CT_WRKEMP='" & glbUserID & "'"
    
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
End If

End Sub

Private Sub CompTimeWRK()
Dim RsCompTime As New ADODB.Recordset
Dim RsATTC As New ADODB.Recordset
Dim rsPARCO As New ADODB.Recordset
Dim SQLQ, xEmpnbrN, xEmpnbrO, xEarned, xTaken, xBalance, xPaid
Dim xDECHR
Dim rsOvtBank As New ADODB.Recordset
Dim OEmpNbr

xDECHR = 2
rsPARCO.Open "HRPARCO", gdbAdoIhr001, adOpenStatic
If (Not rsPARCO.EOF) And (Not rsPARCO.BOF) Then
    xDECHR = rsPARCO("PC_DECHR")
    If xDECHR < 2 Then xDECHR = 2
End If

SQLQ = "DELETE FROM HRCOMPTIME " & in_SQL(glbIHRDBW) & " WHERE CT_WRKEMP='" & glbUserID & "'"

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

If glbOracle Then
    SQLQ = "SELECT SH_EMPNBR,SH_SALARY as SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
    SQLQ = SQLQ & " AD_DOA,AD_HRS,AD_REASON,JB_GRPCD "
    If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
        SQLQ = SQLQ & ", JB_FTEHRS, ED_ORG, ED_PT, ED_EMP, ED_LOC, ED_REGION, ED_ADMINBY, ED_SECTION "
    End If
    SQLQ = SQLQ & " FROM HR_SALARY_HISTORY, HR_JOB_HISTORY,HREMP, HRJOB , HR_ATTENDANCE "
    SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
    SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
    
    If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_SALARY_HISTORY.SH_PAYROLL_ID = HR_JOB_HISTORY.JH_PAYROLL_ID "
    End If
    
    SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR "
    SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
    SQLQ = SQLQ & " AND HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR "
    SQLQ = SQLQ & " AND JH_CURRENT<>0 AND SH_CURRENT<>0 "
    SQLQ = SQLQ & " AND (SUBSTR(AD_REASON,1,2) = 'OT' OR SUBSTR(AD_REASON,1,2) = 'CT')"
    If Len(WRKSQL) > 0 Then SQLQ = SQLQ & " AND " & WRKSQL
    If chkInclAtt Then
        SQLQ = SQLQ & " UNION "
        SQLQ = SQLQ & " SELECT SH_EMPNBR,SH_SALARY as SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
        SQLQ = SQLQ & " AH_DOA AS AD_DOA,AH_HRS AS AD_HRS,AH_REASON AS AD_REASON,JB_GRPCD "
        If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
            SQLQ = SQLQ & ", JB_FTEHRS, ED_ORG, ED_PT, ED_EMP, ED_LOC, ED_REGION, ED_ADMINBY, ED_SECTION "
        End If
        SQLQ = SQLQ & " FROM HR_SALARY_HISTORY, HR_JOB_HISTORY, HREMP, HRJOB,HR_ATTENDANCE_HISTORY "
        SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
        
    If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_SALARY_HISTORY.SH_PAYROLL_ID = HR_JOB_HISTORY.JH_PAYROLL_ID "
    End If
        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
        SQLQ = SQLQ & " AND HREMP.ED_EMPNBR = HR_ATTENDANCE_HISTORY.AH_EMPNBR "
        SQLQ = SQLQ & " AND JH_CURRENT<>0 AND SH_CURRENT<>0 "
        SQLQ = SQLQ & " AND (SUBSTR(AH_REASON,1,2) = 'OT' OR SUBSTR(AH_REASON,1,2) = 'CT')"
        If Len(WRKSQL) > 0 Then SQLQ = SQLQ & " AND " & Replace(WRKSQL, "AD_", "AH_")
    End If
    SQLQ = SQLQ & " ORDER BY SH_EMPNBR,AD_DOA "
Else
    SQLQ = "SELECT SH_EMPNBR,SH_SALARY as SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
    SQLQ = SQLQ & " AD_DOA,AD_HRS,AD_REASON,JB_GRPCD "
    If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
        SQLQ = SQLQ & ", JB_FTEHRS, ED_ORG, ED_PT, ED_EMP, ED_LOC, ED_REGION, ED_ADMINBY, ED_SECTION "
    End If
    SQLQ = SQLQ & " FROM (((HR_SALARY_HISTORY INNER JOIN HR_JOB_HISTORY "
    SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
    If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_SALARY_HISTORY.SH_PAYROLL_ID = HR_JOB_HISTORY.JH_PAYROLL_ID)"
    Else
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR)"
    End If
    SQLQ = SQLQ & " INNER JOIN HREMP ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR)"
    SQLQ = SQLQ & " INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE)"
    SQLQ = SQLQ & " INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR "
    SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND SH_CURRENT<>0 "
    SQLQ = SQLQ & " AND (LEFT(AD_REASON,2) = 'OT' OR LEFT(AD_REASON,2) = 'CT')"
    If Len(WRKSQL) > 0 Then SQLQ = SQLQ & " AND " & WRKSQL
    If chkInclAtt Then
        SQLQ = SQLQ & " UNION "
        SQLQ = SQLQ & " SELECT SH_EMPNBR,SH_SALARY as SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
        SQLQ = SQLQ & " AH_DOA AS AD_DOA,AH_HRS AS AD_HRS,AH_REASON AS AD_REASON,JB_GRPCD "
        If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
            SQLQ = SQLQ & ", JB_FTEHRS, ED_ORG, ED_PT, ED_EMP, ED_LOC, ED_REGION, ED_ADMINBY, ED_SECTION "
        End If
        SQLQ = SQLQ & " FROM (((HR_SALARY_HISTORY INNER JOIN HR_JOB_HISTORY "
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
        If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
            SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_SALARY_HISTORY.SH_PAYROLL_ID = HR_JOB_HISTORY.JH_PAYROLL_ID)"
        Else
            SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR)"
        End If
        SQLQ = SQLQ & " INNER JOIN HREMP ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR)"
        SQLQ = SQLQ & " INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE)"
        SQLQ = SQLQ & " INNER JOIN HR_ATTENDANCE_HISTORY ON HREMP.ED_EMPNBR = HR_ATTENDANCE_HISTORY.AH_EMPNBR "
        SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND SH_CURRENT<>0 "
        SQLQ = SQLQ & " AND (LEFT(AH_REASON,2) = 'OT' OR LEFT(AH_REASON,2) = 'CT')"
        If Len(WRKSQL) > 0 Then SQLQ = SQLQ & " AND " & Replace(WRKSQL, "AD_", "AH_")
    End If
    SQLQ = SQLQ & " ORDER BY SH_EMPNBR,AD_DOA "

End If

RsATTC.Open SQLQ, gdbAdoIhr001, adOpenStatic

RsCompTime.Open "SELECT * FROM HRCOMPTIME WHERE CT_WRKEMP='" & glbUserID & "' ORDER BY CT_EMPNBR,CT_DOA", gdbAdoIhr001W, adOpenKeyset, adLockOptimistic  ' , adLockBatchOptimistic
OEmpNbr = 0
Do While Not RsATTC.EOF
    '-----------------------------------------------------------------------------------------
    'Add Previous year brought forward record for each employee
    If OEmpNbr <> RsATTC("SH_EMPNBR") Then
        OEmpNbr = RsATTC("SH_EMPNBR")
        SQLQ = "SELECT * FROM HR_OVERTIME_BANK WHERE OT_EMPNBR = " & RsATTC("SH_EMPNBR")
        rsOvtBank.Open SQLQ, gdbAdoIhr001, adOpenKeyset
        If Not rsOvtBank.EOF Then
            RsCompTime.AddNew
            RsCompTime("CT_EMPNBR") = RsATTC("SH_EMPNBR")
            RsCompTime("CT_REPTAU") = RsATTC("JH_REPTAU")
            RsCompTime("CT_DOA") = rsOvtBank("OT_EFDATE")
            RsCompTime("CT_REASON") = "OTBF"
            RsCompTime("CT_HRS") = rsOvtBank("OT_PBANK")
            RsCompTime("CT_DHRS") = RsATTC("JH_DHRS")
            If RsATTC("SH_SALCD") = "A" Then
                If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                    RsCompTime("CT_HRSSAL") = Round((RsATTC("SH_SAL") / RsATTC("JB_FTEHRS")), xDECHR) '4)
                Else
                    If RsATTC("SH_WHRS") <> 0 Then
                        RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
                    End If
                End If
            ElseIf RsATTC("SH_SALCD") = "H" Then
                RsCompTime("CT_HRSSAL") = RsATTC("SH_SAL")
            ElseIf RsATTC("SH_SALCD") = "M" Then
                If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                    RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") * 12) / RsATTC("JB_FTEHRS")), xDECHR)  '4)
                Else
                    If RsATTC("SH_WHRS") <> 0 Then
                        RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 12) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
                    End If
                End If
            ElseIf RsATTC("SH_SALCD") = "D" Then
                If GetLeapYear(Year(Date)) Then
                    If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                        RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") * 366) / RsATTC("JB_FTEHRS")), xDECHR)  '4)
                    Else
                        If RsATTC("SH_WHRS") <> 0 Then
                            RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 366) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
                        End If
                    End If
                Else
                    If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                        RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") * 365) / RsATTC("JB_FTEHRS")), xDECHR)  '4)
                    Else
                        If RsATTC("SH_WHRS") <> 0 Then
                            RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 365) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
                        End If
                    End If
                End If
            End If
    
            RsCompTime("CT_GRPCD") = RsATTC("JB_GRPCD")
            RsCompTime("CT_WRKEMP") = glbUserID
            RsCompTime.Update
        End If
        rsOvtBank.Close
    End If
    '-----------------------------------------------------------------------------------------
    
    
    RsCompTime.AddNew
    RsCompTime("CT_EMPNBR") = RsATTC("SH_EMPNBR")
    RsCompTime("CT_REPTAU") = RsATTC("JH_REPTAU")
    RsCompTime("CT_DOA") = RsATTC("AD_DOA")
    RsCompTime("CT_REASON") = RsATTC("AD_REASON")
    If Mid(RsATTC("AD_REASON"), 1, 2) = "OT" And glbCompSerial = "S/N - 2378W" Then 'Town of Aurora
        RsCompTime("CT_HRS") = RsATTC("AD_HRS") * Overtime_Multiplier(RsATTC("ED_ORG"), RsATTC("ED_PT"), RsATTC("ED_EMP"), RsATTC("ED_LOC"), RsATTC("ED_REGION"), RsATTC("ED_ADMINBY"), RsATTC("ED_SECTION"))
    Else
        RsCompTime("CT_HRS") = RsATTC("AD_HRS")
    End If
    RsCompTime("CT_DHRS") = RsATTC("JH_DHRS")
    If RsATTC("SH_WHRS") > 0 Then
        'chanegd by Bryan to include days & months ticket#9354
        If RsATTC("SH_SALCD") = "A" Then
            If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                RsCompTime("CT_HRSSAL") = Round((RsATTC("SH_SAL") / RsATTC("JB_FTEHRS")), xDECHR) '4)
            Else
                RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
            End If
        ElseIf RsATTC("SH_SALCD") = "H" Then
            RsCompTime("CT_HRSSAL") = RsATTC("SH_SAL")
        ElseIf RsATTC("SH_SALCD") = "M" Then
            If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") * 12) / RsATTC("JB_FTEHRS")), xDECHR)  '4)
            Else
                RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 12) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
            End If
        ElseIf RsATTC("SH_SALCD") = "D" Then
            If GetLeapYear(Year(Date)) Then
                If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                    RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") * 366) / RsATTC("JB_FTEHRS")), xDECHR)  '4)
                Else
                    RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 366) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
                End If
            Else
                If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
                    RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") * 365) / RsATTC("JB_FTEHRS")), xDECHR)  '4)
                Else
                    RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 365) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
                End If
            End If
        End If
    End If
    RsCompTime("CT_GRPCD") = RsATTC("JB_GRPCD")
    RsCompTime("CT_WRKEMP") = glbUserID
    RsCompTime.Update
    RsATTC.MoveNext 'CT_HRSSAL
Loop
RsATTC.Close
RsCompTime.Requery

xEmpnbrO = 0

Do While Not RsCompTime.EOF
    xEmpnbrN = RsCompTime("CT_EMPNBR")
    If xEmpnbrO <> xEmpnbrN Then
        xBalance = 0
        xEmpnbrO = xEmpnbrN
    End If
    If Not IsNull(RsCompTime("CT_REASON")) Then
        xEarned = 0
        xTaken = 0
        xPaid = 0
        If Mid(RsCompTime("CT_REASON"), 1, 2) = "OT" Then
            ' dkostka - 02/27/2002 - Removed code to make negative amounts zero, not sure
            '   why this was in here, but it shouldn't be.  Checked w/ Linda, she gave
            '   the OK to remove it.
            'If RsCompTime("CT_HRS") > 0 Then
            xEarned = RsCompTime("CT_HRS")
        End If
        
        If Mid(RsCompTime("CT_REASON"), 1, 2) = "CT" Then
            ' dkostka - 02/27/2002 - Removed code to make negative amounts zero, not sure
            '   why this was in here, but it shouldn't be.  Checked w/ Linda, she gave
            '   the OK to remove it.
            'If RsCompTime("CT_HRS") > 0 Then
            xTaken = RsCompTime("CT_HRS")
        End If
               
        If IsDate(RsCompTime("CT_DOA")) Then
            RsCompTime("CT_WEEK") = WeekdayName(Weekday(RsCompTime("CT_DOA")))
        End If
        RsCompTime("CT_EARN") = xEarned
        RsCompTime("CT_TAKEN") = xTaken
        
        xBalance = xBalance + (xEarned - xTaken)
        
        RsCompTime("CT_BALANCE") = xBalance
                
        If RsCompTime("CT_DHRS") > 0 Then
            RsCompTime("CT_EQUDAYS") = Round(((xEarned - xTaken) / RsCompTime("CT_DHRS")), xDECHR)
        End If
        If RsCompTime("CT_HRSSAL") > 0 Then
            RsCompTime("CT_EXTDCOST") = Round(((xEarned - xTaken) * RsCompTime("CT_HRSSAL")), xDECHR)
        Else
            RsCompTime("CT_EXTDCOST") = 0
        End If
        RsCompTime.Update
    End If
    RsCompTime.MoveNext
Loop
RsCompTime.Close

SQLQ = "SELECT CT_EMPNBR AS EMPNBR, SUM(CT_EARN) AS EARN,SUM(CT_TAKEN) AS TAKEN FROM HRCOMPTIME WHERE CT_WRKEMP='" & glbUserID & "' GROUP BY CT_EMPNBR"

RsCompTime.Open SQLQ, gdbAdoIhr001W, adOpenStatic
Do While Not RsCompTime.EOF
    xEmpnbrN = RsCompTime("EMPNBR")
    xBalance = RsCompTime("EARN") - RsCompTime("TAKEN")
    SQLQ = "UPDATE HRCOMPTIME " & in_SQL(glbIHRDBW) & " SET CT_BALANCETOT = " & xBalance & " WHERE CT_EMPNBR = " & xEmpnbrN & " AND CT_WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    RsCompTime.MoveNext
Loop

If Val(txtExxceed) > 0 Then
    SQLQ = "DELETE FROM HRCOMPTIME " & in_SQL(glbIHRDBW) & " WHERE CT_BALANCETOT < " & txtExxceed & " AND CT_WRKEMP='" & glbUserID & "'"
    
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
End If

End Sub

Public Sub CreateCompTimeCost()
Dim SQLQ, Counter

    SQLQ = "CREATE TABLE HRCOMPCOST (CT_COMPNO CHAR(3),CT_CHRGCODE CHAR(15),"
    SQLQ = SQLQ & "CT_EXTEND DOUBLE, CT_BURDEN DOUBLE, CT_TOTAL DOUBLE, "
    SQLQ = SQLQ & "CT_TEXT1 CHAR(15),CT_TEXT2 CHAR(20),CT_NUM1 DOUBLE, CT_NUM2 DOUBLE, CT_WRKEMP CHAR(25))"
    On Error Resume Next
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute SQLQ
    gdbAdoIhr001W.CommitTrans
    On Error GoTo 0
    On Error Resume Next
    SQLQ = "CREATE INDEX MINDEX ON HRCOMPCOST (CT_CHRGCODE);"
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute SQLQ
    gdbAdoIhr001W.CommitTrans
    On Error GoTo 0

ToEnd1:

    gdbAdoIhr001W.Close
    If gdbAdoIhr001W.State = adStateOpen Then gdbAdoIhr001W.Close
    gdbAdoIhr001W.Mode = adModeReadWrite
    gdbAdoIhr001W.Open glbAdoIHRDBW
End Sub
Public Sub CreateCompTime()
Dim SQLQ, Counter
    SQLQ = "CREATE TABLE HRCOMPTIME (CT_COMPNO CHAR(3),CT_EMPNBR LONG,CT_REPTAU LONG,CT_DOA DATE,CT_REASON CHAR(4),CT_WEEK CHAR(10),CT_GRPCD CHAR(6),"
    SQLQ = SQLQ & "CT_HRS DOUBLE, CT_EARN DOUBLE, CT_TAKEN DOUBLE, CT_BALANCE DOUBLE,CT_BALANCETOT DOUBLE,CT_EQUDAYS DOUBLE, CT_EXTDCOST DOUBLE, "
    SQLQ = SQLQ & "CT_DHRS DOUBLE,CT_HRSSAL DOUBLE, CT_TEXT1 CHAR(15),CT_TEXT2 CHAR(20),CT_NUM1 DOUBLE, CT_NUM2 DOUBLE, CT_WRKEMP CHAR(25))"
    On Error Resume Next
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute SQLQ
    gdbAdoIhr001W.CommitTrans
    On Error GoTo 0
    On Error Resume Next
    SQLQ = "CREATE INDEX MINDEX ON HRCOMPTIME (CT_EMPNBR,CT_DOA);"
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute SQLQ
    gdbAdoIhr001W.CommitTrans
    On Error GoTo 0
 
ToEnd1:
    gdbAdoIhr001W.Close
    If gdbAdoIhr001W.State = adStateOpen Then gdbAdoIhr001W.Close
    gdbAdoIhr001W.Mode = adModeReadWrite
    gdbAdoIhr001W.Open glbAdoIHRDBW
    
End Sub

Private Sub SELATTWRK()
Dim xlen, xxx, xx1
'Dim db001 As Database
Dim RsTempDAO As Recordset
Dim RsTempADO As New ADODB.Recordset
Dim RsTemp1 As New ADODB.Recordset
Dim SQLQ, OEmpNbr, xSumOT, xSumCT
Dim xFieldList
On Error GoTo AttWrkError

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3
gdbAdoIhr001.CommandTimeout = 600

Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).FloodPercent = 15

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute "DELETE FROM HRATTWRK " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.CommitTrans

MDIMain.panHelp(0).FloodPercent = 30

xFieldList = Get_Fields(gdbAdoIhr001, "HR_ATTENDANCE", "AD_ATT_ID")
SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ", AD_WRKEMP)"
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & "SELECT " & xFieldList
SQLQ = SQLQ & ",'" & glbUserID & "' AS AD_WRKEMP "
SQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY "
If Len(WRKSQL) > 1 Then
    If glbCompSerial = "S/N - 2378W" Then   'Hemu - Town of Aurora
        SQLQ = SQLQ & " WHERE (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_ATTENDANCE.AD_JOB = HR_JOB_HISTORY.JH_JOB AND HR_ATTENDANCE.AD_PAYROLL_ID = HR_JOB_HISTORY.JH_PAYROLL_ID) AND HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & WRKSQL & ")"
    Else
        SQLQ = SQLQ & " WHERE (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & WRKSQL & ")"
    End If
End If

MDIMain.panHelp(0).FloodPercent = 45

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

MDIMain.panHelp(0).FloodPercent = 60

If chkInclAtt Then
    SQLQ = "INSERT INTO HRATTWRK (AD_COMPNO, AD_EMPNBR, AD_DOA, "
    SQLQ = SQLQ & "AD_REASON_TABL, AD_REASON, AD_HRS, AD_COMM, "
    SQLQ = SQLQ & "AD_CHRGCODE, AD_SHIFT, AD_SUPER, AD_INCID, AD_SEN, "
    SQLQ = SQLQ & "AD_LDATE, AD_LTIME, AD_LUSER , AD_WRKEMP) "
    SQLQ = SQLQ & in_SQL(glbIHRDBW)
    SQLQ = SQLQ & "SELECT AH_COMPNO, "
    SQLQ = SQLQ & "AH_EMPNBR, AH_DOA, "
    SQLQ = SQLQ & "AH_REASON_TABL, AH_REASON, "
    SQLQ = SQLQ & "AH_HRS, AH_COMM, "
    SQLQ = SQLQ & "AH_CHRGCODE, AH_SHIFT, "
    SQLQ = SQLQ & "AH_SUPER, AH_INCID, "
    SQLQ = SQLQ & "AH_SEN, AH_LDATE, "
    SQLQ = SQLQ & "AH_LTIME, AH_LUSER, " & glbUserID & " " 'AH_LUSER "
    SQLQ = SQLQ & "FROM HR_ATTENDANCE_HISTORY "
    If Len(WRKSQL) > 1 Then
        SQLQ = SQLQ & "WHERE (" & Replace(Replace(WRKSQL, "AD_", "AH_"), "HR_ATTENDANCE", "HR_ATTENDANCE_HISTORY") & ")"
    Else
        SQLQ = SQLQ & ""
    End If
    MDIMain.panHelp(0).FloodPercent = 75
    
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    
    If Not glbSQL And Not glbOracle Then Call Pause(1)
End If

SQLQ = "SELECT AD_EMPNBR FROM HRATTWRK "
SQLQ = SQLQ & "WHERE AD_WRKEMP = '" & glbUserID & "' "
SQLQ = SQLQ & "GROUP BY AD_EMPNBR  "
RsTempADO.Open SQLQ, gdbAdoIhr001W, adOpenStatic
OEmpNbr = -1
Do While Not RsTempADO.EOF
    OEmpNbr = RsTempADO("AD_EMPNBR")
    
    'Update with Multiplier
    SQLQ = "UPDATE HRATTWRK SET AD_HRS = AD_HRS * "
    SQLQ = SQLQ & " (SELECT OM_MULTIPLIER FROM HR_OVERTIME_MASTER WHERE OM_ORG ="
    SQLQ = SQLQ & " (SELECT ED_ORG FROM HREMP WHERE ED_EMPNBR = AD_EMPNBR))"
    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & RsTempADO("AD_EMPNBR")
    SQLQ = SQLQ & " AND AD_REASON LIKE 'OT%' "
    gdbAdoIhr001.Execute SQLQ
    
    
    SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS HRSSUM FROM HRATTWRK "
    SQLQ = SQLQ & "WHERE AD_WRKEMP = '" & glbUserID & "' "
    SQLQ = SQLQ & "AND AD_EMPNBR = " & OEmpNbr & " "
    SQLQ = SQLQ & "AND AD_REASON LIKE 'OT%' "
    SQLQ = SQLQ & "GROUP BY AD_EMPNBR  "
    RsTemp1.Open SQLQ, gdbAdoIhr001W, adOpenStatic
    If RsTemp1.EOF Then
        xSumOT = 0
    Else
        xSumOT = RsTemp1("HRSSUM")
    End If
    RsTemp1.Close
    
    SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS HRSSUM FROM HRATTWRK "
    SQLQ = SQLQ & "WHERE AD_WRKEMP = '" & glbUserID & "' "
    SQLQ = SQLQ & "AND AD_EMPNBR = " & OEmpNbr & " "
    SQLQ = SQLQ & "AND AD_REASON LIKE 'CT%' "
    SQLQ = SQLQ & "GROUP BY AD_EMPNBR  "
    RsTemp1.Open SQLQ, gdbAdoIhr001W, adOpenStatic
    If RsTemp1.EOF Then
        xSumCT = 0
    Else
        xSumCT = RsTemp1("HRSSUM")
    End If
    RsTemp1.Close
    
    If xSumOT <= xSumCT Then 'Delete all the records for this employee
        SQLQ = "DELETE FROM HRATTWRK " & in_SQL(glbIHRDBW)
        SQLQ = SQLQ & " WHERE AD_WRKEMP = '" & glbUserID & "' "
        SQLQ = SQLQ & " AND AD_EMPNBR = " & OEmpNbr & " "
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
    End If
    
    RsTempADO.MoveNext
Loop
RsTempADO.Close

gdbAdoIhr001.CommandTimeout = 600
MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
Exit Sub

AttWrkError:
    gdbAdoIhr001.CommandTimeout = 600
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    'If Err.Number = 2147217871 Then MsgBox Err.Description
    ' dkostka - 04/18/01 - Not sure why the previous line was hiding errors,
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    If False Then
        ' danielk - 06/20/2002 - for debugging purposes
        Resume
    End If
    Exit Sub
End Sub

Private Sub FlexBankWRK()
Dim RsCompTime As New ADODB.Recordset
Dim RsATTC As New ADODB.Recordset
Dim rsPARCO As New ADODB.Recordset
Dim SQLQ, xEmpnbrN, xEmpnbrO, xEarned, xTaken, xBalance
Dim xDECHR

xDECHR = 2
rsPARCO.Open "HRPARCO", gdbAdoIhr001, adOpenStatic
If (Not rsPARCO.EOF) And (Not rsPARCO.BOF) Then
    xDECHR = rsPARCO("PC_DECHR")
    If xDECHR < 2 Then xDECHR = 2
End If

SQLQ = "DELETE FROM HRCOMPTIME " & in_SQL(glbIHRDBW) & " WHERE CT_WRKEMP='" & glbUserID & "'"

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

If glbOracle Then
    SQLQ = "SELECT SH_EMPNBR,SH_SALARY AS SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
    SQLQ = SQLQ & " AD_DOA,AD_HRS,AD_REASON,JB_GRPCD "
    SQLQ = SQLQ & " FROM HR_SALARY_HISTORY, HR_JOB_HISTORY,HREMP, HRJOB , HR_ATTENDANCE "
    SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
    SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
    SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR "
    SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
    SQLQ = SQLQ & " AND HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR "
    SQLQ = SQLQ & " AND JH_CURRENT<>0 AND SH_CURRENT<>0 "
    SQLQ = SQLQ & " AND (AD_REASON = 'FLEX' OR AD_REASON = 'TAKE')"
    If Len(WRKSQL) > 0 Then SQLQ = SQLQ & "AND " & WRKSQL
    If chkInclAtt Then
        SQLQ = SQLQ & " UNION "
        SQLQ = SQLQ & " SELECT SH_EMPNBR,SH_SALARY AS SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
        SQLQ = SQLQ & " AH_DOA AS AD_DOA,AH_HRS AS AD_HRS,AH_REASON AS AD_REASON,JB_GRPCD "
        SQLQ = SQLQ & " FROM HR_SALARY_HISTORY, HR_JOB_HISTORY, HREMP, HRJOB,HR_ATTENDANCE_HISTORY "
        SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
        SQLQ = SQLQ & " AND HREMP.ED_EMPNBR = HR_ATTENDANCE_HISTORY.AH_EMPNBR "
        SQLQ = SQLQ & " AND JH_CURRENT<>0 AND SH_CURRENT<>0 "
        SQLQ = SQLQ & " AND (AH_REASON = 'FLEX' OR AH_REASON = 'TAKE')"
        If Len(WRKSQL) > 0 Then SQLQ = SQLQ & " AND " & Replace(WRKSQL, "AD_", "AH_")
    End If
    SQLQ = SQLQ & " ORDER BY SH_EMPNBR,AD_DOA "

Else
    SQLQ = "SELECT SH_EMPNBR,SH_SALARY AS SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
    SQLQ = SQLQ & " AD_DOA,AD_HRS,AD_REASON,JB_GRPCD "
    If glbCBrant Then SQLQ = SQLQ & ",AD_CHRGCODE "    'Franks 04/30/03 Ticket# 4079
    SQLQ = SQLQ & " FROM (((HR_SALARY_HISTORY INNER JOIN HR_JOB_HISTORY "
    SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
    SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR)"
    SQLQ = SQLQ & " INNER JOIN HREMP ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR)"
    SQLQ = SQLQ & " INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE)"
    SQLQ = SQLQ & " INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR "
    SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND SH_CURRENT<>0 "
    SQLQ = SQLQ & " AND (AD_REASON = 'FLEX' OR AD_REASON = 'TAKE')"
    If Len(WRKSQL) > 0 Then SQLQ = SQLQ & "AND " & WRKSQL
    If chkInclAtt Then
        SQLQ = SQLQ & " UNION "
        SQLQ = SQLQ & " SELECT SH_EMPNBR,SH_SALARY AS SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
        SQLQ = SQLQ & " AH_DOA AS AD_DOA,AH_HRS AS AD_HRS,AH_REASON AS AD_REASON,JB_GRPCD "
        If glbCBrant Then SQLQ = SQLQ & ",AH_CHRGCODE AS AD_CHRGCODE " 'Franks 04/30/03 Ticket# 4079
        SQLQ = SQLQ & " FROM (((HR_SALARY_HISTORY INNER JOIN HR_JOB_HISTORY "
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR)"
        SQLQ = SQLQ & " INNER JOIN HREMP ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR)"
        SQLQ = SQLQ & " INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE)"
        SQLQ = SQLQ & " INNER JOIN HR_ATTENDANCE_HISTORY ON HREMP.ED_EMPNBR = HR_ATTENDANCE_HISTORY.AH_EMPNBR "
        SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND SH_CURRENT<>0 "
        SQLQ = SQLQ & " AND (AH_REASON = 'FLEX' OR AH_REASON = 'TAKE')"
        If Len(WRKSQL) > 0 Then SQLQ = SQLQ & " AND " & Replace(WRKSQL, "AD_", "AH_")
    End If
    SQLQ = SQLQ & " ORDER BY SH_EMPNBR,AD_DOA "

End If
'added by Bryan 22/Sep/05 Ticket#9343
If glbCompSerial = "S/N - 2373W" Then 'Muskoka
    SQLQ = Replace(SQLQ, "SH_SALARY", "SH_TOTAL", 1, , vbTextCompare)
End If
RsATTC.Open SQLQ, gdbAdoIhr001, adOpenStatic

RsCompTime.Open "SELECT * FROM HRCOMPTIME WHERE CT_WRKEMP='" & glbUserID & "' ORDER BY CT_EMPNBR,CT_DOA", gdbAdoIhr001W, adOpenKeyset, adLockOptimistic  ' , adLockBatchOptimistic
Do While Not RsATTC.EOF
    RsCompTime.AddNew
    RsCompTime("CT_EMPNBR") = RsATTC("SH_EMPNBR")
    RsCompTime("CT_REPTAU") = RsATTC("JH_REPTAU")
    RsCompTime("CT_DOA") = RsATTC("AD_DOA")
    RsCompTime("CT_REASON") = RsATTC("AD_REASON")
    RsCompTime("CT_HRS") = RsATTC("AD_HRS")
    RsCompTime("CT_DHRS") = RsATTC("JH_DHRS")
    If RsATTC("SH_WHRS") > 0 Then
        If RsATTC("SH_SALCD") = "A" Then
            RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
        ElseIf RsATTC("SH_SALCD") = "H" Then
            RsCompTime("CT_HRSSAL") = RsATTC("SH_SAL")
        ElseIf RsATTC("SH_SALCD") = "M" Then
            RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 12) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
        ElseIf RsATTC("SH_SALCD") = "D" Then
            If GetLeapYear(Year(Date)) Then
                RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 366) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
            Else
                RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 365) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
            End If
        End If
    End If
    RsCompTime("CT_GRPCD") = RsATTC("JB_GRPCD")
    RsCompTime("CT_WRKEMP") = glbUserID
    If glbCBrant Then RsCompTime("CT_TEXT1") = RsATTC("AD_CHRGCODE")  'Franks 04/30/03 Ticket# 4079
    RsCompTime.Update
    RsATTC.MoveNext 'CT_HRSSAL
Loop
RsATTC.Close
RsCompTime.Requery

xEmpnbrO = 0

Do While Not RsCompTime.EOF
    xEmpnbrN = RsCompTime("CT_EMPNBR")
    If xEmpnbrO <> xEmpnbrN Then
        xBalance = 0
        xEmpnbrO = xEmpnbrN
    End If
    If Not IsNull(RsCompTime("CT_REASON")) Then
        xEarned = 0
        xTaken = 0
        If RsCompTime("CT_REASON") = "FLEX" Then
            ' dkostka - 02/27/2002 - Removed code to make negative amounts zero, not sure
            '   why this was in here, but it shouldn't be.  Checked w/ Linda, she gave
            '   the OK to remove it.
            'If RsCompTime("CT_HRS") > 0 Then
            xEarned = RsCompTime("CT_HRS")
        End If
        If RsCompTime("CT_REASON") = "TAKE" Then
            ' dkostka - 02/27/2002 - Removed code to make negative amounts zero, not sure
            '   why this was in here, but it shouldn't be.  Checked w/ Linda, she gave
            '   the OK to remove it.
            'If RsCompTime("CT_HRS") > 0 Then
            xTaken = RsCompTime("CT_HRS")
        End If
        If IsDate(RsCompTime("CT_DOA")) Then
            RsCompTime("CT_WEEK") = WeekdayName(Weekday(RsCompTime("CT_DOA")))
        End If
        RsCompTime("CT_EARN") = xEarned
        RsCompTime("CT_TAKEN") = xTaken
        xBalance = xBalance + (xEarned - xTaken)
        RsCompTime("CT_BALANCE") = xBalance
        If RsCompTime("CT_DHRS") > 0 Then
            RsCompTime("CT_EQUDAYS") = Round(((xEarned - xTaken) / RsCompTime("CT_DHRS")), xDECHR)
        End If
        If RsCompTime("CT_HRSSAL") > 0 Then
            RsCompTime("CT_EXTDCOST") = Round(((xEarned - xTaken) * RsCompTime("CT_HRSSAL")), xDECHR)
        Else
            RsCompTime("CT_EXTDCOST") = 0
        End If
        RsCompTime.Update
    End If
    RsCompTime.MoveNext
Loop
RsCompTime.Close

SQLQ = "SELECT CT_EMPNBR AS EMPNBR, SUM(CT_EARN) AS EARN,SUM(CT_TAKEN) AS TAKEN FROM HRCOMPTIME WHERE CT_WRKEMP='" & glbUserID & "' GROUP BY CT_EMPNBR"
RsCompTime.Open SQLQ, gdbAdoIhr001W, adOpenStatic
Do While Not RsCompTime.EOF
    xEmpnbrN = RsCompTime("EMPNBR")
    xBalance = RsCompTime("EARN") - RsCompTime("TAKEN")
    SQLQ = "UPDATE HRCOMPTIME " & in_SQL(glbIHRDBW) & " SET CT_BALANCETOT = " & xBalance & " WHERE CT_EMPNBR = " & xEmpnbrN & " AND CT_WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    RsCompTime.MoveNext
Loop

If Val(txtExxceed) > 0 Then
    SQLQ = "DELETE FROM HRCOMPTIME " & in_SQL(glbIHRDBW) & " WHERE CT_BALANCETOT < " & txtExxceed & " AND CT_WRKEMP='" & glbUserID & "'"
    
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
End If

End Sub

Private Sub FlexBank_SELATTWRK()
Dim xlen, xxx, xx1
'Dim db001 As Database
Dim RsTempDAO As Recordset
Dim RsTempADO As New ADODB.Recordset
Dim RsTemp1 As New ADODB.Recordset
Dim SQLQ, OEmpNbr, xSumOT, xSumCT
On Error GoTo AttWrkError
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3
gdbAdoIhr001.CommandTimeout = 600
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).FloodPercent = 15
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute "DELETE FROM HRATTWRK " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.CommitTrans
MDIMain.panHelp(0).FloodPercent = 30
SQLQ = "INSERT INTO HRATTWRK "
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & "SELECT " & Get_Fields(gdbAdoIhr001, "HR_ATTENDANCE", "AD_ATT_ID")
SQLQ = SQLQ & ",'" & glbUserID & "' AS AD_WRKEMP "
SQLQ = SQLQ & " FROM HR_ATTENDANCE, HR_JOB_HISTORY "
If Len(WRKSQL) > 1 Then
    SQLQ = SQLQ & " WHERE (HR_ATTENDANCE.AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) AND HR_JOB_HISTORY.JH_CURRENT <> 0 AND (" & WRKSQL & ")"
End If

MDIMain.panHelp(0).FloodPercent = 45
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
MDIMain.panHelp(0).FloodPercent = 60
If chkInclAtt Then
    SQLQ = "INSERT INTO HRATTWRK (AD_COMPNO, AD_EMPNBR, AD_DOA, "
    SQLQ = SQLQ & "AD_REASON_TABL, AD_REASON, AD_HRS, AD_COMM, "
    SQLQ = SQLQ & "AD_CHRGCODE, AD_SHIFT, AD_SUPER, AD_INCID, AD_SEN, "
    SQLQ = SQLQ & "AD_LDATE, AD_LTIME, AD_LUSER , AD_WRKEMP) "
    SQLQ = SQLQ & in_SQL(glbIHRDBW)
    SQLQ = SQLQ & "SELECT AH_COMPNO, "
    SQLQ = SQLQ & "AH_EMPNBR, AH_DOA, "
    SQLQ = SQLQ & "AH_REASON_TABL, AH_REASON, "
    SQLQ = SQLQ & "AH_HRS, AH_COMM, "
    SQLQ = SQLQ & "AH_CHRGCODE, AH_SHIFT, "
    SQLQ = SQLQ & "AH_SUPER, AH_INCID, "
    SQLQ = SQLQ & "AH_SEN, AH_LDATE, "
    SQLQ = SQLQ & "AH_LTIME, AH_LUSER, " & glbUserID & " " 'AH_LUSER "
    SQLQ = SQLQ & "FROM HR_ATTENDANCE_HISTORY "
    If Len(WRKSQL) > 1 Then
        SQLQ = SQLQ & "WHERE (" & Replace(Replace(WRKSQL, "AD_", "AH_"), "HR_ATTENDANCE", "HR_ATTENDANCE_HISTORY") & ")"
    Else
        SQLQ = SQLQ & ""
    End If
    MDIMain.panHelp(0).FloodPercent = 75
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    If Not glbSQL And Not glbOracle Then Call Pause(1)
End If

SQLQ = "SELECT AD_EMPNBR FROM HRATTWRK "
SQLQ = SQLQ & "WHERE AD_WRKEMP = '" & glbUserID & "' "
SQLQ = SQLQ & "GROUP BY AD_EMPNBR  "
RsTempADO.Open SQLQ, gdbAdoIhr001W, adOpenStatic
OEmpNbr = -1
Do While Not RsTempADO.EOF
    OEmpNbr = RsTempADO("AD_EMPNBR")
    SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS HRSSUM FROM HRATTWRK "
    SQLQ = SQLQ & "WHERE AD_WRKEMP = '" & glbUserID & "' "
    SQLQ = SQLQ & "AND AD_EMPNBR = " & OEmpNbr & " "
    SQLQ = SQLQ & "AND AD_REASON = 'FLEX' "
    SQLQ = SQLQ & "GROUP BY AD_EMPNBR  "
    RsTemp1.Open SQLQ, gdbAdoIhr001W, adOpenStatic
    If RsTemp1.EOF Then
        xSumOT = 0
    Else
        xSumOT = RsTemp1("HRSSUM")
    End If
    RsTemp1.Close
    
    SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS HRSSUM FROM HRATTWRK "
    SQLQ = SQLQ & "WHERE AD_WRKEMP = '" & glbUserID & "' "
    SQLQ = SQLQ & "AND AD_EMPNBR = " & OEmpNbr & " "
    SQLQ = SQLQ & "AND AD_REASON = 'TAKE' "
    SQLQ = SQLQ & "GROUP BY AD_EMPNBR  "
    RsTemp1.Open SQLQ, gdbAdoIhr001W, adOpenStatic
    If RsTemp1.EOF Then
        xSumCT = 0
    Else
        xSumCT = RsTemp1("HRSSUM")
    End If
    RsTemp1.Close
    
    If xSumOT <= xSumCT Then 'Delete all the records for this employee
        SQLQ = "DELETE FROM HRATTWRK " & in_SQL(glbIHRDBW)
        SQLQ = SQLQ & " WHERE AD_WRKEMP = '" & glbUserID & "' "
        SQLQ = SQLQ & " AND AD_EMPNBR = " & OEmpNbr & " "
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
    End If
    
    RsTempADO.MoveNext
Loop
RsTempADO.Close

gdbAdoIhr001.CommandTimeout = 600
MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
Exit Sub

AttWrkError:
    gdbAdoIhr001.CommandTimeout = 600
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    'If Err.Number = 2147217871 Then MsgBox Err.Description
    ' dkostka - 04/18/01 - Not sure why the previous line was hiding errors,
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    If False Then
        ' danielk - 06/20/2002 - for debugging purposes
        Resume
    End If
    Exit Sub
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

