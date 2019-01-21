VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmUEmpPos 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Mass Update Employee/Position"
   ClientHeight    =   9990
   ClientLeft      =   15
   ClientTop       =   1230
   ClientWidth     =   13050
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9990
   ScaleWidth      =   13050
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtComment 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   19
      Tag             =   "00-Position Comments"
      Top             =   7460
      Width           =   2895
   End
   Begin VB.TextBox txtComments2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   20
      Tag             =   "00-Position Notes"
      Top             =   7810
      Width           =   2895
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   270
      Left            =   9090
      TabIndex        =   25
      Top             =   7080
      Width           =   855
   End
   Begin VB.CheckBox chkUpdAttendance 
      Caption         =   "Update Employee's Attendance with New Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Tag             =   "40-Update Attendance records with Salary -y/n"
      Top             =   9120
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1930
      TabIndex        =   8
      Top             =   3055
      Width           =   435
   End
   Begin INFOHR_Controls.DateLookup dlpNSDate 
      Height          =   285
      Left            =   2100
      TabIndex        =   15
      Tag             =   "41-Effective Date of Salary Record"
      Top             =   5898
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin VB.ComboBox comPayType 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9960
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Tag             =   "Annually or Hourly"
      Top             =   9165
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtSALCD 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   9480
      MaxLength       =   1
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   9180
      Visible         =   0   'False
      Width           =   375
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1620
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   710
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
      Left            =   1620
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   375
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
      Index           =   2
      Left            =   1620
      TabIndex        =   3
      Top             =   1380
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
      Index           =   3
      Left            =   1620
      TabIndex        =   4
      Top             =   1715
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
      Left            =   1620
      TabIndex        =   5
      Top             =   2050
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDPT"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Left            =   1620
      TabIndex        =   9
      Tag             =   "00-Enter Position Code"
      Top             =   3405
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   5
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1620
      TabIndex        =   10
      Tag             =   "00-Group - Code"
      Top             =   3740
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "JBGC"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1620
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2385
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpNReason 
      Height          =   285
      Left            =   2100
      TabIndex        =   16
      Tag             =   "01-Reason for change in salary - Code "
      Top             =   6276
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDRC"
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   1800
      Top             =   8400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport vbxCrystal1 
      Left            =   1200
      Top             =   8400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1620
      TabIndex        =   2
      Top             =   1045
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDLC"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   7
      Left            =   1620
      TabIndex        =   7
      Top             =   2720
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDSE"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpGrid 
      Height          =   315
      Left            =   1620
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   556
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "JBGD"
      MaxLength       =   0
      MultiSelect     =   -1  'True
      Object.Height          =   315
   End
   Begin VB.ComboBox comStep 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   18
      Tag             =   "Steps of salary change "
      Top             =   7035
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSMask.MaskEdBox medAmountChng 
      Height          =   285
      Left            =   2400
      TabIndex        =   29
      Tag             =   "21-Amount in dollars of salary change "
      Top             =   7080
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.DateLookup dlpAsofDate 
      Height          =   285
      Left            =   9540
      TabIndex        =   26
      Tag             =   "41-As of Date of Salary to apply the Percent update on"
      Top             =   8760
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.CodeLookup clpNJob 
      Height          =   285
      Left            =   2100
      TabIndex        =   14
      Tag             =   "01-Position code"
      Top             =   5520
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   6
      LookupType      =   5
   End
   Begin INFOHR_Controls.CodeLookup clpNDiv 
      Height          =   285
      Left            =   8835
      TabIndex        =   22
      Tag             =   "00-Division"
      Top             =   5898
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpNDept 
      Height          =   285
      Left            =   8835
      TabIndex        =   21
      Tag             =   "00-Department"
      Top             =   5520
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpNLoc 
      Height          =   285
      Left            =   8835
      TabIndex        =   23
      Tag             =   "00-Location - Code"
      Top             =   6276
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin Threed.SSOption optUpdateType 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   12
      Tag             =   "Salary update - dollars"
      Top             =   4680
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   503
      _StockProps     =   78
      Caption         =   "On Lay Off"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSOption optUpdateType 
      Height          =   285
      Index           =   1
      Left            =   4110
      TabIndex        =   13
      Tag             =   "Salary update -Fixed Dollars"
      Top             =   4680
      Width           =   2355
      _Version        =   65536
      _ExtentX        =   4154
      _ExtentY        =   503
      _StockProps     =   78
      Caption         =   "Return from Lay Off"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin INFOHR_Controls.CodeLookup clpNAdminBy 
      Height          =   285
      Left            =   8835
      TabIndex        =   24
      Tag             =   "00-Administered By"
      Top             =   6654
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.DateLookup dlpExpRtnDate 
      Height          =   285
      Left            =   2100
      TabIndex        =   17
      Tag             =   "41-Expected Return Date"
      Top             =   6654
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   25
      Left            =   7440
      TabIndex        =   59
      Top             =   6699
      Width           =   1125
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes 1"
      Height          =   195
      Left            =   240
      TabIndex        =   58
      Top             =   7505
      Width           =   555
   End
   Begin VB.Label lblComment2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes 2"
      Height          =   195
      Left            =   240
      TabIndex        =   57
      Top             =   7855
      Width           =   555
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Update Type"
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
      Index           =   0
      Left            =   120
      TabIndex        =   56
      Top             =   4680
      Width           =   1680
   End
   Begin VB.Label txtLambtonJob 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6960
      TabIndex        =   55
      Top             =   9240
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   11
      Left            =   7440
      TabIndex        =   54
      Top             =   5565
      Width           =   990
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   23
      Left            =   7440
      TabIndex        =   53
      Top             =   6321
      Width           =   615
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   13
      Left            =   7440
      TabIndex        =   52
      Top             =   5943
      Width           =   675
   End
   Begin VB.Label lblImport 
      Caption         =   "Job Offer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   7440
      TabIndex        =   51
      Top             =   7080
      Width           =   855
   End
   Begin VB.Image imgNoSec 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   8550
      Picture         =   "fuEmpPos.frx":0000
      Top             =   7080
      Width           =   240
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "New Position Code"
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
      Index           =   1
      Left            =   240
      TabIndex        =   27
      Top             =   5565
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mass Update Records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   50
      Top             =   5160
      Width           =   1890
   End
   Begin VB.Image imgHelp 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   9165
      Picture         =   "fuEmpPos.frx":014A
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblAsofDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "As of Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8280
      TabIndex        =   49
      Top             =   8805
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblGrid 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Grid Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   48
      Top             =   4140
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblShift 
      Caption         =   "Shift"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   3076
      Width           =   855
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   46
      Top             =   2744
      Width           =   900
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
      TabIndex        =   45
      Top             =   1084
      Width           =   1095
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
      TabIndex        =   44
      Top             =   2080
      Width           =   630
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   43
      Top             =   2412
      Width           =   1290
   End
   Begin VB.Label lblPosGroup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Group"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   42
      Top             =   3800
      Width           =   1035
   End
   Begin VB.Label lblStep 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Step"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   41
      Top             =   7095
      Width           =   330
   End
   Begin VB.Label lblReason 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason for Change"
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
      Left            =   240
      TabIndex        =   40
      Top             =   6321
      Width           =   1650
   End
   Begin VB.Label lblExpRtnDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Expected Return Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   38
      Top             =   6699
      Width           =   1590
   End
   Begin VB.Label lblEffectiveDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
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
      Left            =   240
      TabIndex        =   37
      Top             =   5943
      Width           =   885
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
      TabIndex        =   36
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblPosTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   3468
      Width           =   975
   End
   Begin VB.Label lblEStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   1748
      Width           =   1350
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   33
      Top             =   1416
      Width           =   840
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
      TabIndex        =   32
      Top             =   752
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
      TabIndex        =   31
      Top             =   420
      Width           =   555
   End
   Begin VB.Image imgSec 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   8550
      Picture         =   "fuEmpPos.frx":058C
      Top             =   7080
      Width           =   240
   End
End
Attribute VB_Name = "frmUEmpPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbCOMPA#, fglbGRADE$
Dim snapNoUpSal As New ADODB.Recordset 'Frank 4/10/2000
Dim fglbSDate As Variant
Dim snapPosSal As New ADODB.Recordset
Dim fglbFrmt
Dim OSalary, NSalary, OEDate, NEDate, ONDate, NNDate, empNo&, dblWHours#, OTOTAL, OPremium
Dim oPayP, NPayp, OJOB1, xGrade, OSalCD, oGrade, oSReas1, oSReas2, oSReas3, oSComment, oSComment2, oCompa
Dim oPayrollID, oGrid
Dim oVGroup, oVStep
Dim lngRecs&, SkipRec&
Dim GLfocus
Dim fglAddDel As String
Dim strEMPLIST 'George Mar 14,2006
Dim MailBody, MailBodyP
Dim fglbDhrs
Dim WSQLQ As String
Dim MsgSal, IfDisplay
Dim oStep
Dim dynSH_Job1 As New ADODB.Recordset

Dim oPHRS, oWHRS, ODHRS, oJob As String, OSDATE
Dim OLeadHand, OLabourCD, OReason
Dim oOrg, oDeptNo, oStatus, oGLNo, oComment, oComment2
Dim oPayCategory
Dim OLambtonJob
Dim OFTE, fOldFTE, fNewFTE, fFTEDate
Dim oLABOUREDATE
Dim oENDDATE, oEndReason
Dim OBillingRate
Dim oSHIFT As String, oREPTAU As String
Dim nJobID
Dim oRepAut, oRepAut2, oRepAut3, oRepAut4
Dim oFTENum, oFTEHrs
Dim oPTFT
Dim oDiv, oDept, oEmp, oRegion, oSect, oPosCtrl, oPayCateg, oBillRate, oLoc, oPosStatus, oAdminBy
Dim nDiv, nDeptNo, nLoc, nAdminBy
Dim oSal

Private Function AUDITPOS(xEmpNo)
Dim rsTA As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim rsTB As New ADODB.Recordset
Dim strFields As String
Dim ACTX As String
Dim UpdateAudit As Boolean
Dim UptPositionDate As Date
Dim HRFields As New Collection
Dim xBatchID
Dim xPayment_oldValue, xPayment_newValue
Dim PayIDs
Dim X  As Integer
Dim Y As Integer

On Error GoTo AUDITPOS_ERR
AUDITPOS = False

rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpNo, gdbAdoIhr001, adOpenKeyset
If Not rsTB.EOF Then
    If IsNull(rsTB("ED_PT")) Then xPT = "" Else xPT = rsTB("ED_PT")
    If IsNull(rsTB("ED_DIV")) Then xDiv = "" Else xDiv = rsTB("ED_DIV")
Else
    xPT = ""
    xDiv = ""
End If

rsTB.Close
Set rsTB = Nothing

MODUPD:

ACTX = "A"
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, AU_DOLENT_TABL, "
strFields = strFields & "AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_PHRS, AU_OLDPHRS, AU_WHRS, AU_OLDWHRS, AU_DHRS, AU_OLDDHRS, "
strFields = strFields & "AU_JOB, AU_SJDATE, AU_JREASON, AU_LEADHAND, AU_LABOURCD, AU_LABOUREDATE,AU_JOBCOMMENT, "
strFields = strFields & "AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID, AU_ORG, AU_BILLINGRATE "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv
rsTA("AU_JOB") = clpNJob.Text
rsTA("AU_SJDATE") = dlpNSDate.Text
rsTA("AU_JREASON") = clpNReason.Text
rsTA("AU_JOBCOMMENT") = Left(oComment, 30)
rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = xEmpNo

If CVDate(dlpNSDate.Text) > CVDate(Date) Then
    rsTA("AU_LDATE") = CVDate(dlpNSDate.Text)
Else
    rsTA("AU_LDATE") = Date
End If

rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX

If glbMulti Then
    rsTA("AU_PAYROLL_ID") = oPayrollID
Else
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & xEmpNo
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then
            rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
            oPayrollID = rsEmp("ED_PAYROLL_ID")
        End If
    End If
    rsEmp.Close
    Set rsEmp = Nothing
End If

'rsTA("AU_PHRS") = oPHRS
'rsTA("AU_OLDPHRS") = oPHRS
'rsTA("AU_WHRS") = oWHRS
'rsTA("AU_OLDWHRS") = oWHRS
'rsTA("AU_DHRS") = ODHRS
'rsTA("AU_OLDDHRS") = ODHRS
'rsTA("AU_LEADHAND") = OLeadHand
'rsTA("AU_LABOURCD") = OLabourCD
'If IsDate(oLABOUREDATE) Then
'    rsTA("AU_LABOUREDATE") = oLABOUREDATE
'Else
'    rsTA("AU_LABOUREDATE") = Null
'End If

rsTA.Update

rsTA.Close
Set rsTA = Nothing

'Vadim transfer
If glbVadim Then
    Dim HRChangs As New Collection
    
    clpNJob.DataField = "JH_JOB"
    dlpNSDate.DataField = "JH_SDATE"
    clpNReason.DataField = "JH_JREASON"
    clpNDiv.DataField = "ED_DIV"
    clpNDept.DataField = "ED_DEPTNO"
    clpNLoc.DataField = "ED_LOC"
    clpNAdminBy.DataField = "ED_ADMINBY"    'Ticket #29480 - Additional fields
    
    If CVDate(Format(dlpNSDate.Text, "mm/dd/yyyy")) > CVDate(Format(Date, "mm/dd/yyyy")) Then
        UptPositionDate = dlpNSDate.Text
    Else
        UptPositionDate = Date
    End If

    If glbLambton Then
        If IsNull(oGrid) Or oGrid = "" Then
            OLambtonJob = oJob
            txtLambtonJob = oJob
        Else
            OLambtonJob = Left(oGrid, 1) & oJob & Mid(oGrid, 2)
            txtLambtonJob = Left(oGrid, 1) & clpNJob.Text & Mid(oGrid, 2)
        End If
    
        If isChanged_Field(HRChangs, OLambtonJob, txtLambtonJob) Then UpdateAudit = True
    Else
        If isChanged_Field(HRChangs, oJob, clpNJob) Then UpdateAudit = True
    End If
    If OSDATE <> "" Then
        If isChanged_Field(HRChangs, Str(OSDATE), dlpNSDate) Then UpdateAudit = True
    Else
        If isChanged_Field(HRChangs, OSDATE, dlpNSDate) Then UpdateAudit = True
    End If
    If isChanged_Field(HRChangs, OReason, clpNReason) Then UpdateAudit = True
        
    If isChanged_Field(HRChangs, oDiv, clpNDiv) Then UpdateAudit = True
    If isChanged_Field(HRChangs, oDeptNo, clpNDept) Then UpdateAudit = True
    If isChanged_Field(HRChangs, oLoc, clpNLoc) Then UpdateAudit = True
    If isChanged_Field(HRChangs, oAdminBy, clpNAdminBy) Then UpdateAudit = True
    
    'Call Passing_Changes(HRChangs, Demographices, "M", UptPositionDate, xEmpNO)
    Call Passing_Changes(HRChangs, Position, "M", UptPositionDate, xEmpNo, oPayrollID)
    
    clpNJob.DataField = ""
    dlpNSDate.DataField = ""
    clpNReason.DataField = ""
    clpNDiv.DataField = ""
    clpNDept.DataField = ""
    clpNLoc.DataField = ""
    clpNAdminBy.DataField = ""  'Ticket #29480 - Additional fields

    'Ticket #29480 - Additional fields
    'When going on Layoff
    If optUpdateType(0) Then
        'Update Vadim with Lay Off Payment Type
        glbChgTermReason = "LAYO"
        glbChgTermDate = dlpNSDate.Text
        
        'This routine will update Vadim with Lay Off Date, Payment Type and Active Flag
        Call TermPayrollEmp(UptPositionDate, xEmpNo, oPayrollID, Termination)
    
        'Ticket #29480 - Additional fields - Update Payment Type in info:HR
        Call UpdPaymentTypeVadim(xEmpNo, "LAYO")
    Else
        'If isTransfer(Rehire) Then
        'Update Vadim withe Payment Type that was there before Lay Off.
        xBatchID = AddBatchVadim("M")
        HRFields.Add "TERM:LO-DATE"
        HRFields.Add "TERM:EMP_DATE"
        xPayment_oldValue = "L"
        
        'Get New Payment Type which is actually what was there before the Layoff (L) from the Employee History
        If Vadim_PayType_field <> "" Then
            'xPayment_newValue = getEmpValue(Vadim_PayType_field, xEmpnbr, xPayID)
            If Vadim_PayType_field = "ED_LOC" Then
                xPayment_newValue = getEmpHistoryValue("LOC", xPayment_oldValue, xEmpNo)
            ElseIf Vadim_PayType_field = "ED_SECTION" Then
                xPayment_newValue = getEmpHistoryValue("SECTION", xPayment_oldValue, xEmpNo)
            ElseIf Vadim_PayType_field = "ED_ADMINBY" Then
                xPayment_newValue = getEmpHistoryValue("ADMINBY", xPayment_oldValue, xEmpNo)
            ElseIf Vadim_PayType_field = "ED_REGION" Then
                xPayment_newValue = getEmpHistoryValue("REGION", xPayment_oldValue, xEmpNo)
            End If
        End If
        
        PayIDs = Split(getPayrollIDs(xEmpNo, , True), "|")
        For X = 0 To UBound(PayIDs)
            xBatchID = AddBatchVadim("M", UptPositionDate)
            For Y = 1 To HRFields.count
                'Clear Layoff Dates
                'City of Timmins - Vadim is not accepting "Null" or Null for Term date to be cleared
                If glbCompSerial = "S/N - 2375W" And HRFields(Y) = "TERM:EMP_DATE" Then
                    Call VadimInterface(xBatchID, PayIDs(X), HRFields(Y), OSDATE, "")
                Else
                    Call VadimInterface(xBatchID, PayIDs(X), HRFields(Y), OSDATE, Null)
                End If
            Next
            
            'Payment Type
            If Vadim_PayType_field <> "" Then
                Call VadimInterface(xBatchID, PayIDs(X), Vadim_PayType_field, xPayment_oldValue, xPayment_newValue)
            End If
            
            'Active Flag
            Call VadimInterface(xBatchID, PayIDs(X), "DFLT:EMP_ACTIVE_FLAG:Y", "N", "Y")
            
            Call CloseBatchVadim(xBatchID)
        Next
        'End If
        
        'Ticket #29480 - Additional fields - Update Payment Type in info:HR with value that was there before Lay Off
        Call UpdPaymentTypeVadim(xEmpNo, xPayment_newValue)
        
        'Ticket #29480 - Update Employee History for this Payment Type change
        If Vadim_PayType_field = "ED_LOC" Then
            If Not EmpHisCalc(2, xEmpNo, "", "", "", "", "", "", "", Date, "LOC", xPayment_newValue, , , , , xPayment_oldValue) Then MsgBox "EMPHIS Error "
        ElseIf Vadim_PayType_field = "ED_SECTION" Then
            If Not EmpHisCalc(2, xEmpNo, "", "", "", "", "", "", "", Date, "SECTION", xPayment_newValue, , , , , xPayment_oldValue) Then MsgBox "EMPHIS Error "
        ElseIf Vadim_PayType_field = "ED_ADMINBY" Then
            If Not EmpHisCalc(2, xEmpNo, "", "", "", "", "", "", "", Date, "ADMINBY", xPayment_newValue, , , , , xPayment_oldValue) Then MsgBox "EMPHIS Error "
        ElseIf Vadim_PayType_field = "ED_REGION" Then
            If Not EmpHisCalc(2, xEmpNo, "", "", "", "", "", "", "", Date, "REGION", xPayment_newValue, , , , , xPayment_oldValue) Then MsgBox "EMPHIS Error "
        End If
    End If
End If

MODNOPOSUPD:

AUDITPOS = True

Exit Function

AUDITPOS_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING POSITION AUDIT RECORD", "AUDIT FILE", "UPDATE")

If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Private Function AUDITSALY(xEmpNo)
Dim TA As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim TB As New ADODB.Recordset
Dim strFields As String
Dim HRChanges As New Collection
Dim UptSalaryDate As Date
Dim UpdateAudit As Boolean

On Error GoTo AUDITSALY_ERR

AUDITSALY = False


TB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpNo, gdbAdoIhr001, adOpenForwardOnly
If Not TB.EOF Then
    If IsNull(TB("ED_PT")) Then
        xPT = ""
    Else
        xPT = TB("ED_PT")
    End If
    If IsNull(TB("ED_DIV")) Then
        xDiv = ""
    Else
        xDiv = TB("ED_DIV")
    End If
Else
    xPT = ""
    xDiv = ""
End If
TB.Close
'TA.Open "HRAUDIT", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdTableDirect
'strfields added by Bryan 02/Dec/05 TICKET#9899
strFields = "AU_LOC_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL,AU_DOLENT_TABL, "
strFields = strFields & "AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_SALARY, AU_OLDSAL, AU_PAYP, AU_OLDPAYP, AU_PAYP, "
strFields = strFields & "AU_OLDPAYP, AU_OLDPAYP, AU_JOB, AU_GRID, AU_PAYROLL_ID, AU_SALCD, AU_WHRS, AU_SEDATE, AU_SNDATE, "
strFields = strFields & "AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_SREASON "
TA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
xADD = False

'If OSalary <> NSalary Then GoTo MODUPD
'If OPayp <> NPayp Then GoTo MODUPD      'laura jan 28, 1998
'If OEDate <> NEDate Then GoTo MODUPD
'If ONDate <> NNDate Then GoTo MODUPD
'GoTo MODNOUPD

MODUPD:
TA.AddNew
TA("AU_LOC_TABL") = "EDLC": TA("AU_EMP_TABL") = "EDEM": TA("AU_SUPCODE_TABL") = "EDSP": TA("AU_ORG_TABL") = "EDOR"
TA("AU_PAYP_TABL") = "SDPP": TA("AU_BCODE_TABL") = "BNCD": TA("AU_TREAS_TABL") = "TERM": TA("AU_DOLENT_TABL") = "EDOL"
TA("AU_EARN_TABL") = "EARN"
TA("AU_NEWEMP") = "N"
TA("AU_PTUPL") = xPT
TA("AU_DIVUPL") = xDiv
TA("AU_SALARY") = NSalary
TA("AU_OLDSAL") = oSal
TA("AU_PAYP") = oPayP ' FRANK 4/5/2000    'NPayp  Laura jan 28, 1998
TA("AU_OLDPAYP") = oPayP    '    ""
TA("AU_GRID") = oGrid
TA("AU_SALCD") = OSalCD
If oWHRS <> "" And Not IsNull(oWHRS) Then TA("AU_WHRS") = oWHRS
'If ONDate <> NNDate Then TA("AU_SNDATE") = IIf(IsDate(NNDate), NNDate, Null)  'Jaddy 11/15/99

'Ticket #23666 - Update with Salary Reason for Change as well.
TA("AU_JOB") = clpNJob.Text
TA("AU_SEDATE") = dlpNSDate.Text
TA("AU_SREASON") = clpNReason.Text

TA("AU_COMPNO") = "001"
TA("AU_EMPNBR") = xEmpNo

'Ticket #23943 - Town of Orangeville noticed the LDATE was not getting updated properly - Jerry asked to fix this as per Salary screen.
If glbCompSerial = "S/N - 2227W" And (xPT = "SE" Or xPT = "OT") Then ' CCAC Kingston, see ticket #3296
    TA("AU_LDATE") = Format(DateAdd("d", 14, dlpNSDate.Text), "SHORT DATE")
Else
    'Ticket #23943 - Town of Orangeville
    If glbCompSerial = "S/N - 2383W" Then
        If CVDate(dlpNSDate.Text) > CVDate(Date) Then
            TA("AU_LDATE") = Format(dlpNSDate.Text, "SHORT DATE")
        Else
            TA("AU_LDATE") = Date
        End If
    Else
        If CVDate(dlpNSDate.Text) > CDate(Date) Then
            TA("AU_LDATE") = Format(dlpNSDate.Text, "SHORT DATE")
        Else
            TA("AU_LDATE") = Date
        End If
    End If
End If
UptSalaryDate = TA("AU_LDATE")

TA("AU_LUSER") = glbUserID
TA("AU_LTIME") = Time$
TA("AU_UPLOAD") = "N"
TA("AU_TYPE") = "A"

If glbMulti Then
    TA("AU_PAYROLL_ID") = oPayrollID
Else
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & xEmpNo
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then
            TA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
            oPayrollID = rsEmp("ED_PAYROLL_ID")
        End If
    End If
    rsEmp.Close
    Set rsEmp = Nothing
End If

TA.Update

TA.Close
Set TA = Nothing

'Vadim transfer
If glbVadim And comStep.Visible = False Then
    If oPayrollID = "" Or IsNull(oPayrollID) Then
        'Do not transfer as there is no Payroll ID
    Else
        dlpNSDate.DataField = "SH_EDATE"
        If isChanged_Field(HRChanges, OEDate, dlpNSDate) Then UpdateAudit = True
        Call Passing_Changes(HRChanges, Salary, "M", UptSalaryDate, xEmpNo, oPayrollID)
        dlpNSDate.DataField = ""
        
        'If isChanged_Field(HRChanges, OReason, clpNReason) Then UpdateAudit = True
        'If isChanged_Field(HRChanges, oJob, clpNJob) Then UpdateAudit = True
    End If
End If

MODNOUPD:

AUDITSALY = True

Exit Function

AUDITSALY_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING SALARY AUDIT RECORD", "AUDIT FILE", "UPDATE")

If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Private Function chkMUEmpPosSal()

Dim SQLQ As String, Msg$, dd&, X%
Dim DgDef As Variant, Title$, Response%, DCurSHDate  As Variant

chkMUEmpPosSal = False

On Error GoTo chkSalH_Err

If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
'    MsgBox lStr("If Division Entered - it must be known")
'     clpDiv.SetFocus
    Exit Function
End If

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
'    MsgBox "If Department Entered - it must be known"
'    clpDept.SetFocus
    Exit Function
End If

For X% = 0 To 3
    If Not clpCode(X).ListChecker Then Exit Function
Next X

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
'    MsgBox lStr("Category code must be valid")
'    clpPT.SetFocus
    Exit Function
End If

'If Len(clpCode(0).Text) > 0 And clpCode(0).Caption = "Unassigned" Then
'    MsgBox lStr("Invalid Location Code")
'    clpCode(0).SetFocus
'    Exit Function
'End If

If Not clpCode(7).ListChecker Then Exit Function
'If Len(clpCode(7).Text) > 0 And clpCode(7).Caption = "Unassigned" Then
'    MsgBox "Invalid Section Code"
'    clpCode(7).SetFocus
'    Exit Function
'End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

If Not clpJob.ListChecker Then Exit Function

If glbMultiGrid And clpGrid.Visible = True Then
    If Not clpGrid.ListChecker Then Exit Function
End If

If Not clpNJob.ListChecker Then Exit Function

If Len(clpNJob.Text) < 1 Then
    MsgBox "New Position is required"
    clpNJob.SetFocus
    Exit Function
End If

If Len(dlpNSDate.Text) < 1 Then
    MsgBox "Start Date is required"
    dlpNSDate.SetFocus
    Exit Function
Else
    If Not IsDate(dlpNSDate.Text) Then
        MsgBox "Invalid Start Date"
        dlpNSDate.SetFocus
        Exit Function
    End If
End If

If Len(clpNReason.Text) < 1 Then
    MsgBox "Reason for Change is required"
    clpNReason.SetFocus
    Exit Function
End If

If Len(clpNReason.Text) > 0 Then
    If clpNReason = "Unassigned" Then
        MsgBox "Invalid Reason for Change Code"
        clpNReason.SetFocus
        Exit Function
    End If
    
    'If optStep Then
    '    If Len(comStep) < 1 Then
    '        MsgBox "Step is required if code entered"
    '        comStep.SetFocus
    '        Exit Function
    '    Else
    '        'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    '        'If (Val(comStep(x%)) < 1 Or Val(comStep(x%)) > 11) And comStep(x%) <> "Next Step" Then
    '        'If (Val(comStep(X%)) < 1 Or Val(comStep(X%)) > 15) And comStep(X%) <> "Next Step" Then
    '        If (Val(comStep) < 1 Or Val(comStep) > 20) And comStep <> "Next Step" Then
    '            MsgBox "Step is not valid"
    '            comStep.SetFocus
    '            Exit Function
    '        End If
    '    End If
    'Else
    '    If Len(medAmountChng(x%)) < 1 Then
    '        MsgBox "Salary is required if code entered"
    '        medAmountChng(x%).SetFocus
    '        Exit Function
    '    Else
    '        If Not IsNumeric(medAmountChng(x%)) Then   'laura jan 12, 1998
    '            MsgBox "Salary is not valid"
    '            Exit Function
    '        End If
    '    End If
    'End If
End If

If optUpdateType(0) Then
    If Len(dlpExpRtnDate.Text) > 0 Then
        If Not IsDate(dlpExpRtnDate.Text) Then
            MsgBox "Expected Return Date is invalid"
            dlpExpRtnDate.SetFocus
            Exit Function
        End If
            
        dd& = DateDiff("d", CVDate(dlpNSDate.Text), CVDate(dlpExpRtnDate.Text))
        
        If dd& < 0 Then
            MsgBox "Expected Return Date cannot preceed Position Start Date"
            dlpNSDate.SetFocus
            Exit Function
        End If
    End If
End If

If optUpdateType(1) Then
    If Len(comStep) < 1 Then
        'Ticket #29382 - They want it to be optional
        'MsgBox "Step is required"
        'comStep.SetFocus
        'Exit Function
    Else
        'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
        'If (Val(comStep(x%)) < 1 Or Val(comStep(x%)) > 11) And comStep(x%) <> "Next Step" Then
        'If (Val(comStep(X%)) < 1 Or Val(comStep(X%)) > 15) And comStep(X%) <> "Next Step" Then
        If (Val(comStep) < 1 Or Val(comStep) > 20) And comStep <> "Next Step" Then
            MsgBox "Step is not valid"
            comStep.SetFocus
            Exit Function
        End If
    End If
End If

If Not clpDiv.ListChecker Then Exit Function

If Not clpDept.ListChecker Then Exit Function

If Not clpNLoc.ListChecker Then Exit Function

If Not clpNAdminBy.ListChecker Then Exit Function


chkMUEmpPosSal = True
Exit Function

chkSalH_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkSal", "HR_JOB_HISTORY, HR_SALARY_HISTORY", "edit/Add")
Resume Next

End Function

Private Sub chkUpdAttendance_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbPrecision_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdClose_Click()
Unload Me

End Sub

Public Sub cmdNew_Click()
Dim Title$, Msg$, DgDef As Variant, Response%
Dim I As Integer

clpDiv.SetFocus

On Error GoTo Mod_Err

If Not gSec_Upd_Salary And Not gSec_Upd_Position Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If


'txtSALCD = Left(comPayType, 1)

fglAddDel = "Add"

If Not chkMUEmpPosSal() Then Exit Sub

If Not modGet_PosSal_Records() Then Exit Sub       ' get selection - form level

Title$ = "Update Employee's Position and Salary"

Msg$ = ""
Msg$ = Msg$ & "Must have Current Salary." & Chr(10)
Msg$ = Msg$ & "Must have Current Position with Position Code and Position Start Date that matches the current Salary record." & Chr(10)
'Msg$ = Msg$ & "Must have Current Postion with Position Code and Position Start Date match current Salary's. " & Chr(10)
'Msg$ = Msg$ & "Must have Position Master to count the Salary Step. " & Chr(10)
Msg$ = Msg$ & Chr(10) & Chr(10)

If snapPosSal.BOF And snapPosSal.EOF Then
    Msg$ = Msg$ & "No Employees with this selection criteria exist!  " & Chr(10)
    'Msg$ = Msg$ & "Please ensure the Hourly/Annually selection box is set correctly to match the group you want to update." & Chr(10)
    Msg$ = Msg$ & Chr(10)
    DgDef = MB_OK + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
    MsgBox Msg$, , Title$
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

lngRecs& = snapPosSal.RecordCount
Msg$ = Msg$ & lngRecs& & " number of Employees will be affected." & Chr(10) & Chr(10)
Msg$ = Msg$ & "Are you sure you want to create a new Employee Position/Salary records for this group of employees?"

DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
If Response% = IDNO Then ' Evaluate response
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

Msg$ = "Do you want to print a list of employees updated?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.

strEMPLIST = ""
If optUpdateType(0) Then
    If Not modUpdateSelection_On_LAYOFF() Then Exit Sub
Else
    If Not modUpdateSelection_ReturnFrom_LAYOFF() Then Exit Sub
End If


Screen.MousePointer = DEFAULT
If SkipRec = 0 Then
    MsgBox "Records Updated Successfully"
Else
    Msg$ = lngRecs& - SkipRec & " record(s) was updated and " & SkipRec & " record(s) was skipped." & Chr(10)
    Msg$ = Msg$ & "Please click on Print or View icon on the Toolbar to see the reports for the updated and skipped employees."
    MsgBox Msg$, , Title$
End If

'Print the list of employees updated.
If Response% = IDYES Then    ' Yes response
    If lngRecs& - SkipRec > 0 Then
        'Call set_PrintState(False)
        Screen.MousePointer = HOURGLASS
        
        'Call getWSQLQ("U")
        
        ' report name
      
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList.rpt"
    
        Me.vbxCrystal.Formulas(0) = "rTitle='Mass Update Employee/Position - Employee Details'"
        'set location for database tables
        If Len(glbstrSelCri) >= 0 Then
            Me.vbxCrystal.SelectionFormula = getWSQLQRPT
        End If
        'If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        'Else
        '    Me.vbxCrystal.Connect = "PWD=petman;"
        '    Me.vbxCrystal.DataFiles(0) = glbIHRDB
        'End If
        
        ' window title if appropriate
        Me.vbxCrystal.WindowTitle = "Employees-updated Report"
        
        Me.vbxCrystal.Destination = 0
        Screen.MousePointer = DEFAULT
        Me.vbxCrystal.Action = 1
        vbxCrystal.Reset
    
    End If
End If

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

    If Not PrtForm("Mass Update Employee/Position Report Criteria", Me) Then Exit Sub
    
 '   cmdPrint.Enabled = False
 '   cmdView.Enabled = False
    Screen.MousePointer = HOURGLASS
    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    Me.vbxCrystal1.Destination = 1
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    
    Me.vbxCrystal1.Action = 1
    vbxCrystal1.Reset
    
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True

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
 '   cmdPrint.Enabled = False
 '   cmdView.Enabled = False
    Screen.MousePointer = HOURGLASS
    X% = Cri_SetAll()
    
'    Me.vbxCrystal.Destination = 0
'    Me.vbxCrystal1.Destination = 0
'    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
'
'    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset

'    Me.vbxCrystal1.Action = 1
    vbxCrystal1.Reset
    MDIMain.Timer1.Enabled = True
 '   cmdPrint.Enabled = True
 '   cmdView.Enabled = True
End Sub

'Private Sub cmbRound_Click()
'    If cmbRound.Text = "Yes" Then
'        cmbPrecision.Visible = True
'        lblDecimal.Visible = True
'        cmbPrecision.Text = glbCompDecHR
'    Else
'        cmbPrecision.Visible = False
'        lblDecimal.Visible = False
'    End If
'End Sub

Private Sub cmbRound_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpNJob_LostFocus()
    If optUpdateType(1) Then
        If Len(clpNJob.Text) > 0 Then
            Call Populate_Steps
        End If
    End If
End Sub


Private Sub cmdImport_Click()
    glbDocNewRecord = True
    glbDocName = "OfferMU"
    glbDocKey = 0
    frmInAttachment.Show 1
    DoEvents
    
    glbDocName = "Offer"
    glbLEE_ID = 0
    
    Call DispimgIcon(Me, "frmUEmpPos")
    
    If glbDocImpFile <> "" Then
        imgSec.Visible = True
        imgNoSec.Visible = False
    Else
        imgSec.Visible = False
        imgNoSec.Visible = True
    End If
End Sub

Private Sub comPayType_GotFocus()
Call SetPanHelp(ActiveControl)

End Sub

Private Sub comStep_Click()
medAmountChng = comStep
End Sub

Private Sub comStep_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpExpRtnDate_LostFocus()
    If IsDate(dlpExpRtnDate) Then
        txtComment.Text = "Expected Return Date:" & dlpExpRtnDate.Text
    End If
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMUEMPPOS"
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim X%, RFound As Integer, xx%

Screen.MousePointer = HOURGLASS

glbOnTop = "FRMUEMPPOS"

Call CRLIST_EECAT

Call setRptCaption(Me)

'Ticket #29480 - Additional fields
Call setCaption(lbltitle(11))
Call setCaption(lbltitle(13))
Call setCaption(lbltitle(23))
Call setCaption(lbltitle(25))
Call setCaption(lblComment)
Call setCaption(lblComment2)

If glbMulti Then lblUnion.ForeColor = &HC000C0: lblPT.ForeColor = &HC000C0

If glbMultiGrid Then
    lblGrid.Visible = True
    clpGrid.Visible = True
End If

If glbSyndesis Then
    lblPosGroup.Caption = "Position Grade"
    clpCode(1).Tag = "00-Grade - Code"
End If

If glbCompSerial = "S/N - 2191W" Then fglbFrmt = "0.0" Else fglbFrmt = "00"

If glbWFC And glbCountry = "AUSTRALIA" Then
    medAmountChng.Format = "$#,##0.0000;($#,##0.0000)"
End If

If glbCompSerial = "S/N - 2172W" Then 'Lanark Ticket #17221 by Frank 08/19/2009
    lblPosGroup.Caption = "Salary Level"
    clpCode(1).TABLTitle = "Salary Level Code"
End If

'comPayType.ListIndex = 0

'Hide / Show controls based on Update Type
If optUpdateType(0) Then
    lblExpRtnDate.Visible = True
    dlpExpRtnDate.Visible = True
    lblStep.Visible = False
    comStep.Visible = False

    'Ticket #28457 City of Niagara Falls - Show the default values
    If glbCompSerial = "S/N - 2276W" Then
        If Len(clpNDept.Text) = 0 Then clpNDept.Text = "4900"
        If Len(clpNDiv.Text) = 0 Then clpNDiv.Text = "LAYO"
        If Len(clpNLoc.Text) = 0 Then clpNLoc.Text = "NW"
        If Len(clpNAdminBy.Text) = 0 Then clpNAdminBy.Text = "4900"     'Ticket #29480 - Additional fields
    End If
Else
    lblExpRtnDate.Visible = False
    dlpExpRtnDate.Visible = False
    lblStep.Visible = True
    comStep.Visible = True
    
    'Ticket #28457 City of Niagara Falls - Show the default values
    If glbCompSerial = "S/N - 2276W" Then
        clpNJob.Text = ""
        dlpNSDate.Text = ""
        clpNReason.Text = ""
        dlpExpRtnDate.Text = ""
        clpNDept.Text = ""
        clpNDiv.Text = ""
        clpNLoc.Text = ""
        clpNAdminBy.Text = ""
        txtComment.Text = ""
        txtComments2.Text = ""
    End If
End If

If glbMulti Then lblUnion.ForeColor = &HC000C0: lblPT.ForeColor = &HC000C0

If Not glbMultiGrid Then lblGrid.Visible = False: clpGrid.Visible = False

Call INI_Controls(Me)

If glbWFC Then 'Ticket #25911 Franks 10/21/2014
    clpJob.TransDiv = glbWFCUserSecList
End If

'Initializing the values
glbDocImpFile = ""
glbDocType = ""
glbDocDesc = ""

GLfocus = False
Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
Set frmUEmpPos = Nothing  'carmen apr 2000
End Sub


Private Sub imgHelp_Click()
Dim MsgStr As String
    MsgStr = "System will find the employee's with Salary Effective Date >= 'As of Date' and will compute the Percentage of Change based on that Salary which will then be applied to the existing Current Salary of the employee to compute New Current Salary."
    MsgBox MsgStr, vbInformation

End Sub

'Private Sub medAmountChng_GotFocus(Index As Integer)
'Call SetPanHelp(ActiveControl)
'' dkostka - 07/10/2001 - Added multiplication code below, this way if the user clicks on the
''   Change Amount fields more than once, it won't keep getting smaller.
'If IsNumeric(medAmountChng(Index)) Then
'    If optPct Then medAmountChng(Index) = medAmountChng(Index) * 100
'End If
'End Sub
'Private Sub medAmountChng_LostFocus(Index As Integer)
'If IsNumeric(medAmountChng(Index)) Then
'    If optPct Then medAmountChng(Index) = medAmountChng(Index) / 100
'End If
'End Sub

Private Function modGet_PosSal_Records()

Dim BD As Integer
Dim SQLQ As String, countr As Integer
Dim Dat1 As Variant, Dat2 As Variant
Dim iOneWhere As Integer, strJob$, strTm$, X%

modGet_PosSal_Records = False
On Error GoTo modGet_PosSal_Records_Err
strTm$ = Time$
Dim Dt As Variant
Dt = Date$

Screen.MousePointer = HOURGLASS

If glbOracle Then
    If glbMultiGrid Then
        SQLQ = "SELECT HR_JOB_HISTORY.*, HRJOB_GRADE.* "
        SQLQ = SQLQ & " FROM HR_JOB_HISTORY, HRJOB_GRADE ,HR_SALARY_HISTORY "
        SQLQ = SQLQ & " WHERE HR_JOB_HISTORY.JH_JOB=HRJOB_GRADE.JB_CODE "
        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_GRID=HRJOB_GRADE.JB_GRID "
        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_EMPNBR=HR_SALARY_HISTORY.SH_EMPNBR "
        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_JOB=HR_SALARY_HISTORY.SH_JOB"
        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_GRID=HR_SALARY_HISTORY.SH_GRID"
        SQLQ = SQLQ & " AND JH_CURRENT<>0 AND SH_CURRENT<>0"
    Else
        SQLQ = "SELECT HR_SALARY_HISTORY.*, HRJOB.* "
        SQLQ = SQLQ & " FROM HR_SALARY_HISTORY, HRJOB ,HR_JOB_HISTORY "
        SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_JOB=HRJOB.JB_CODE "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_JOB=HR_JOB_HISTORY.JH_JOB"
        SQLQ = SQLQ & " AND SH_CURRENT<>0 AND JH_CURRENT<>0"
    End If
Else
    If glbMultiGrid Then
        SQLQ = "SELECT HR_JOB_HISTORY.*, HRJOB_GRADE.* "
        SQLQ = SQLQ & " FROM (HR_JOB_HISTORY INNER JOIN HRJOB_GRADE "
        SQLQ = SQLQ & " ON HR_JOB_HISTORY.JH_JOB=HRJOB_GRADE.JB_CODE  "
        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_GRID=HRJOB_GRADE.JB_GRID) "
        SQLQ = SQLQ & " INNER JOIN HR_SALARY_HISTORY "
        SQLQ = SQLQ & " ON HR_JOB_HISTORY.JH_EMPNBR=HR_SALARY_HISTORY.SH_EMPNBR"
        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_JOB=HR_SALARY_HISTORY.SH_JOB"
        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_GRID=HR_SALARY_HISTORY.SH_GRID"
        SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND SH_CURRENT<>0"
    Else
        SQLQ = "SELECT HR_JOB_HISTORY.*, HRJOB.* "
        SQLQ = SQLQ & " FROM ((HR_JOB_HISTORY INNER JOIN HRJOB "
        SQLQ = SQLQ & " ON HR_JOB_HISTORY.JH_JOB=HRJOB.JB_CODE)  "
        SQLQ = SQLQ & " INNER JOIN HR_SALARY_HISTORY "
        SQLQ = SQLQ & " ON HR_JOB_HISTORY.JH_EMPNBR=HR_SALARY_HISTORY.SH_EMPNBR"
        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_JOB=HR_SALARY_HISTORY.SH_JOB"
        If glbMulti Then
            SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_SDATE=HR_SALARY_HISTORY.SH_SDATE"
        End If
        SQLQ = SQLQ & " )"
        SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND SH_CURRENT<>0"
    End If
End If
If Len(clpJob.Text) > 0 Then SQLQ = SQLQ & " AND JH_JOB IN ('" & Replace(clpJob.Text, ",", "','") & "') "
If Len(clpGrid.Text) > 0 Then SQLQ = SQLQ & " AND JH_GRID IN ('" & Replace(clpGrid.Text, ",", "','") & "') "
'SQLQ = SQLQ & " AND SH_SALCD = '" & txtSALCD & "'"
'Ticket #27555 Franks 09/18/2015 - missing ')', added it
If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND JH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD IN ('" & Replace(clpCode(1).Text, ",", "','") & "')) "

SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn
If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND ED_DEPTNO IN ('" & Replace(clpDept.Text, ",", "','") & "') "
If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND ED_DIV IN ('" & Replace(clpDiv.Text, ",", "','") & "') "
If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND ED_EMP IN ('" & Replace(clpCode(3).Text, ",", "','") & "') "
If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND ED_LOC IN ('" & Replace(clpCode(0).Text, ",", "','") & "') "
If Len(clpCode(7).Text) > 0 Then SQLQ = SQLQ & " AND ED_SECTION IN ('" & Replace(clpCode(7).Text, ",", "','") & "') "
If Len(txtShift.Text) > 0 Then SQLQ = SQLQ & " AND ED_SHIFT = '" & txtShift.Text & "' "

If Len(elpEEID.Text) > 0 Then SQLQ = SQLQ & " AND JH_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
If glbMulti Then SQLQ = SQLQ & ") "
If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND " & IIf(glbMulti, "JH_ORG", "ED_ORG") & " IN ('" & Replace(clpCode(2).Text, ",", "','") & "') "
If Len(Trim(clpPT.Text)) > 0 Then SQLQ = SQLQ & " AND " & IIf(glbMulti, "JH_PT", "ED_PT") & " IN ('" & Replace(clpPT.Text, ",", "','") & "') "
If glbNoNONE Then SQLQ = SQLQ & " AND " & IIf(glbMulti, "JH_ORG", "ED_ORG") & " <> 'NONE' "
If glbNoEXEC Then SQLQ = SQLQ & " AND " & IIf(glbMulti, "JH_ORG", "ED_ORG") & " <> 'EXEC' "  'Hemu -EXE
If Not glbMulti Then SQLQ = SQLQ & ") "


If snapPosSal.State <> 0 Then snapPosSal.Close
snapPosSal.Open SQLQ, gdbAdoIhr001, adOpenStatic

modGet_PosSal_Records = True

Exit Function

modGet_PosSal_Records_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modGet_PosSal_Records", "qry_MU_PosSal", "Select")


End Function

Private Sub modSetCOMPA_GRADE(dblNewSalary, rsSal)
Dim X%, cX$
Dim dblSsalary#, dblHoursPerWeek#, ssalary@
Dim Jb_No#
Dim SQLQ As String
Dim snapJob As New ADODB.Recordset

If glbMultiGrid Then
    SQLQ = "SELECT * FROM HRJOB_GRADE WHERE JB_CODE='" & clpNJob.Text & "' AND JB_GRID='" & rsSal("SH_GRID") & "'"
Else
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE='" & clpNJob.Text & "'"
End If
snapJob.Open SQLQ, gdbAdoIhr001, adOpenStatic

'SET COMPA RATIO
'================
ssalary@ = dblNewSalary
'Franks May 17,2002 for error of Invalid Use of Null Ticket 2252
'dblHoursPerWeek# = rsSal("SH_WHRS")
If IsNull(rsSal("SH_WHRS")) Then
    dblHoursPerWeek# = 0
Else
    dblHoursPerWeek# = rsSal("SH_WHRS")
End If
'Franks May 17,2002 for error of Invalid Use of Null Ticket 2252

'MOnthly and DAily added by Bryan 28/Sep/05 Ticket#9354
If (glbCompSerial = "S/N - 2378W") And (rsSal("SH_SALCD") <> snapJob("JB_SALCD")) And rsSal("SH_SALCD") <> "M" Then   'Town of Aurora
    dblSsalary# = dblNewSalary
Else
If rsSal("SH_SALCD") = "H" Then
    If snapJob("JB_SALCD") = "A" Then
        dblSsalary# = (dblNewSalary * dblHoursPerWeek#) * 52
    ElseIf snapJob("JB_SALCD") = "M" Then
        dblSsalary# = dblNewSalary * 12
    ElseIf snapJob("JB_SALCD") = "D" Then
        If GetLeapYear(Year(Date)) Then
            dblSsalary# = dblNewSalary * 366
        Else
            dblSsalary# = dblNewSalary * 365
        End If
        
        'Ticket #17654 - formula correction
        dblSsalary# = dblNewSalary * fglbDhrs
    Else
        dblSsalary# = dblNewSalary
    End If
ElseIf rsSal("SH_SALCD") = "A" Then
    If snapJob("JB_SALCD") = "H" Then
        'Franks May 17,2002 for error of Invalid Use of Null Ticket 2252
        'dblSsalary# = (dblNewSalary / dblHoursPerWeek#) / 52
        If dblHoursPerWeek# = 0 Then
            dblSsalary# = 0
        Else
            dblSsalary# = (dblNewSalary / dblHoursPerWeek#) / 52
        End If
        'Franks May 17,2002 for error of Invalid Use of Null Ticket 2252
    ElseIf snapJob("JB_SALCD") = "A" Then
        dblSsalary# = dblNewSalary
    ElseIf snapJob("JB_SALCD") = "M" Then
        dblSsalary# = dblNewSalary / 12
    ElseIf snapJob("JB_SALCD") = "D" Then
        If GetLeapYear(Year(Date)) Then
            dblSsalary# = dblNewSalary / 366
        Else
            dblSsalary# = dblNewSalary / 365
        End If
        
        'Ticket #17654 - formula correction
        dblSsalary# = ((dblNewSalary / dblHoursPerWeek#) / 52) * fglbDhrs
    End If
ElseIf rsSal("SH_SALCD") = "M" Then
    If snapJob("JB_SALCD") = "A" Then
        dblSsalary# = dblNewSalary * 12
    ElseIf snapJob("JB_SALCD") = "M" Then
        dblSsalary# = dblNewSalary
    ElseIf snapJob("JB_SALCD") = "D" Then
        If GetLeapYear(Year(Date)) Then
            dblSsalary# = (dblNewSalary * 12) / 366
        Else
            dblSsalary# = (dblNewSalary * 12) / 365
        End If
        
        'Ticket #17654 - formula correction
        dblSsalary# = ((((dblNewSalary * 12) / dblHoursPerWeek#) / 52)) * fglbDhrs
    Else
        If dblHoursPerWeek# = 0 Then
            dblSsalary# = 0
        Else
            dblSsalary# = (dblNewSalary * 12) / dblHoursPerWeek# / 52
        End If
    End If
ElseIf rsSal("SH_SALCD") = "D" Then
    If snapJob("JB_SALCD") = "H" Then
        If dblHoursPerWeek# = 0 Then
            dblSsalary# = 0
        Else
            If GetLeapYear(Year(Date)) Then
                dblSsalary# = (dblNewSalary * 366) / dblHoursPerWeek# / 52
            Else
                dblSsalary# = (dblNewSalary * 365) / dblHoursPerWeek# / 52
            End If
        End If
        
        'Ticket #17654 - formula correction
        If fglbDhrs <> 0 Then
            dblSsalary# = dblNewSalary / fglbDhrs
        Else
            dblSsalary# = 0
        End If
        
    ElseIf snapJob("JB_SALCD") = "M" Then
        If GetLeapYear(Year(Date)) Then
            dblSsalary# = (dblNewSalary * 366) / 12
        Else
            dblSsalary# = (dblNewSalary * 365) / 12
        End If
        
        'Ticket #17654 - formula correction
        dblSsalary# = ((dblNewSalary / fglbDhrs) * dblHoursPerWeek# * 52) / 12
        
    ElseIf snapJob("JB_SALCD") = "A" Then
        If GetLeapYear(Year(Date)) Then
            dblSsalary# = dblNewSalary * 366
        Else
            dblSsalary# = dblNewSalary * 365
        End If
    
        'Ticket #17654 - formula correction
        dblSsalary# = (dblNewSalary / fglbDhrs) * dblHoursPerWeek# * 52
    Else
        dblSsalary# = dblNewSalary
    End If
End If
End If
'end bryan

 ' set COMPA RATIO
 'laura 03/23/98

'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
'If rsSal("JB_MIDPOINT") >= 1 And rsSal("JB_MIDPOINT") <= 11 Then
'If rsSal("JB_MIDPOINT") >= 1 And rsSal("JB_MIDPOINT") <= 15 Then
If snapJob("JB_MIDPOINT") >= 1 And snapJob("JB_MIDPOINT") <= 20 Then
    If glbCompSerial = "S/N - 2378W" And rsSal("SH_SALCD") <> snapJob("JB_SALCD") And rsSal("SH_SALCD") <> "M" Then  'Town of Aurora
        If Not IsNull(snapJob("JB_S" & snapJob("JB_MIDPOINT") & "A")) Then
            Jb_No = snapJob("JB_S" & snapJob("JB_MIDPOINT") & "A")
        End If
    Else
        If Not IsNull(snapJob("JB_S" & snapJob("JB_MIDPOINT"))) Then
            Jb_No = snapJob("JB_S" & snapJob("JB_MIDPOINT"))
        End If
    End If
End If

fglbCOMPA# = 0

If Jb_No <> 0 And dblSsalary# <> 0 Then 'laura 03/23/98
  fglbCOMPA# = (dblSsalary# / Jb_No) * 100
End If

 
If fglbCOMPA# > 999.99 Then
    fglbCOMPA# = 999.99
End If


'Determine Pay Scale individual fits into
'==========================================
fglbGRADE$ = "00"
'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
'For x% = 1 To 11
'For X% = 1 To 15
For X% = 1 To 20
    If glbCompSerial = "S/N - 2378W" And rsSal("SH_SALCD") <> snapJob("JB_SALCD") And rsSal("SH_SALCD") <> "M" Then   'Town of Aurora
        If IsNumeric(snapJob("JB_S" & CStr(X%) & "A")) Then
            If dblNewSalary >= snapJob("JB_S" & CStr(X%) & "A") And snapJob("JB_S" & CStr(X%) & "A") > 0 Then
                cX$ = CStr(X)
                If X% <= 9 Then cX$ = "0" & cX$
                fglbGRADE$ = cX$
            End If
        End If
    Else
        If IsNumeric(snapJob("JB_S" & CStr(X%))) Then
            If dblSsalary# >= snapJob("JB_S" & CStr(X%)) And snapJob("JB_S" & CStr(X%)) > 0 Then
                cX$ = CStr(X)
                If X% <= 9 Then cX$ = "0" & cX$
                fglbGRADE$ = cX$
            End If
        End If
    End If
Next X%

If glbCompSerial = "S/N - 2378W" And rsSal("SH_SALCD") <> snapJob("JB_SALCD") And rsSal("SH_SALCD") <> "M" Then   'Town of Aurora
    If IsNumeric(snapJob("JB_S1A")) Then
        If dblSsalary# < snapJob("JB_S1A") Then
            fglbGRADE$ = "00"
        End If
    End If
Else
    If IsNumeric(snapJob("JB_S1")) Then
        If dblSsalary# < snapJob("JB_S1") Then
            fglbGRADE$ = "00"
        End If
    End If
End If

snapJob.Close
Set snapJob = Nothing

End Sub

Private Function modUpdateSelection_1()
Dim lngLastCurrentID&, X%
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim JobInfo As Boolean
Dim prec%, curSalary#
Dim fTablSalHis  As New ADODB.Recordset
Dim dblPct, dblAmnt, dblOSalary, dblOSalary1, dblOSalary2, dblOSalary3, dblOSalary4, DtTm As Variant, dblNewSalary As Double
Dim salarystep, xGradeF
Dim JobSalCD, xWHRS
Dim SalDetailCode1, SalDetailCode2, SalDetailCode3, SalDetailCode4 As String
Dim EDate1, EDate2, EDate3, EDate4
Dim xSHID 'George Mar 7,2006 #9965
Dim SQLQ
Dim xAsofDateSal As Double

On Error GoTo modUpdateSelection_1_Err

modUpdateSelection_1 = False

Call DelNoUpsal

Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 5

SkipRec = 0

Do While Not snapPosSal.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / (lngRecs&)))
    MDIMain.panHelp(0).FloodPercent = pct%
    
    empNo& = snapPosSal("SH_EMPNBR")
    
    If glbCompSerial = "S/N - 2288W" Then 'MUSASHI AUTO TKT#10845
        If Not IsNull(snapPosSal("SH_SALARY1")) Then
            dblOSalary = snapPosSal("SH_SALARY1")
            OSalary = snapPosSal("SH_SALARY1")
        Else
            dblOSalary = snapPosSal("SH_SALARY")
            OSalary = snapPosSal("SH_SALARY")
        End If

        dblOSalary2 = 0
        dblOSalary3 = 0
        dblOSalary4 = 0
        SalDetailCode1 = ""
        SalDetailCode2 = ""
        SalDetailCode3 = ""
        SalDetailCode4 = ""
        EDate1 = ""
        EDate2 = ""
        EDate3 = ""
        EDate4 = ""
        
        If Not IsNull(snapPosSal("SH_SALARY2")) And snapPosSal("SH_SALARY2") <> 0 Then
            dblOSalary2 = Val(snapPosSal("SH_SALARY2"))
        End If
        If Not IsNull(snapPosSal("SH_SALARY3")) And snapPosSal("SH_SALARY3") <> 0 Then
            dblOSalary3 = Val(snapPosSal("SH_SALARY3"))
        End If
        If Not IsNull(snapPosSal("SH_SALARY4")) And snapPosSal("SH_SALARY4") <> 0 Then
            dblOSalary4 = Val(snapPosSal("SH_SALARY4"))
        End If
        
        If Not IsNull(snapPosSal("SH_DETAILCODE1")) Then
            SalDetailCode1 = snapPosSal("SH_DETAILCODE1")
        End If
        If Not IsNull(snapPosSal("SH_DETAILCODE2")) Then
            SalDetailCode2 = snapPosSal("SH_DETAILCODE2")
        End If
        If Not IsNull(snapPosSal("SH_DETAILCODE3")) Then
            SalDetailCode3 = snapPosSal("SH_DETAILCODE3")
        End If
        If Not IsNull(snapPosSal("SH_DETAILCODE4")) Then
            SalDetailCode4 = snapPosSal("SH_DETAILCODE4")
        End If
        
        If IsDate(snapPosSal("SH_EDATE1")) Then
            EDate1 = snapPosSal("SH_EDATE1")
        End If
        If IsDate(snapPosSal("SH_EDATE2")) Then
            EDate2 = snapPosSal("SH_EDATE2")
        End If
        If IsDate(snapPosSal("SH_EDATE3")) Then
            EDate3 = snapPosSal("SH_EDATE3")
        End If
        If IsDate(snapPosSal("SH_EDATE4")) Then
            EDate4 = snapPosSal("SH_EDATE4")
        End If
    Else
        dblOSalary = snapPosSal("SH_SALARY")
        OSalary = snapPosSal("SH_SALARY")
        OTOTAL = snapPosSal("SH_TOTAL")
    End If
    oPayP = snapPosSal("SH_PAYP")      'laura jan 28, 1998
    OEDate = snapPosSal("SH_EDATE")

    OJOB1 = snapPosSal("SH_JOB")
    oGrid = snapPosSal("SH_GRID")
    oPayrollID = snapPosSal("SH_PAYROLL_ID")
    OSalCD = snapPosSal("SH_SALCD")
    oGrade = snapPosSal("SH_GRADE")
    xWHRS = snapPosSal("SH_WHRS")
    
    If IsNull(snapPosSal("SH_GRADE")) Then xGrade = 0 Else xGrade = Val(snapPosSal("SH_GRADE"))
    
    Call GetGrade(dblOSalary, xGrade, OJOB1, OSalCD, xWHRS)
    
    xGradeF = xGrade

    'Ticket #18668 and Ticket #19154 - Allow same salary effective date update since we are allowing manual
    'update on the salary screen. So changed from >= to >.
    If OEDate > CVDate(dlpNSDate.Text) Then
        Call NoUpSal_Addnew("D")
        GoTo lblNextRec
    End If
    
    ONDate = snapPosSal("SH_NEXTDAT")
    
    If IsNull(snapPosSal("SH_WHRS")) Or Len(snapPosSal("SH_WHRS")) < 1 Then
        dblWHours# = 0
    Else
        dblWHours# = snapPosSal("SH_WHRS")
    End If
    
    lngLastCurrentID& = snapPosSal("SH_ID")
    
    DtTm = Now
    JobInfo = False
    
    'If optStep Then
        'For x% = 4 To 6
            If Len(clpCode(X%).Text) > 0 Then
                'If comStep(x%) = "Next Step" Then
                '    xGrade = xGrade + 1
                '    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
                '    'If xGrade > 11 Then
                '    'If xGrade > 15 Then
                '    If xGrade > 20 Then
                '        JobInfo = True
                '    Else
                '        salarystep = snapPosSal("JB_S" & Format(xGrade, "##"))
                '        JobSalCD = snapPosSal("JB_SALCD")
                '        If glbCompSerial = "S/N - 2378W" And OSalCD <> JobSalCD And OSalCD <> "M" Then  'Town of Aurora
                '            salarystep = snapPosSal("JB_S" & Format(xGrade, "##") & "A")
                '        End If
                '        If OSalCD <> JobSalCD And xWHrs = 0 Then JobInfo = True
                '        If IsNull(salarystep) Then JobInfo = True
                '        If salarystep = 0 Then JobInfo = True
                '    End If
                'Else
                    salarystep = snapPosSal("JB_S" & Format(comStep, "##"))
                    If glbCompSerial = "S/N - 2378W" And OSalCD <> JobSalCD And OSalCD <> "M" Then   'Town of Aurora
                        salarystep = snapPosSal("JB_S" & Format(comStep, "##") & "A")
                    End If
                    If IsNull(salarystep) Then JobInfo = True
                    If salarystep = 0 Then JobInfo = True
                'End If
            End If
        'Next
    'End If
    
    If JobInfo Then
        Call NoUpSal_Addnew(xGrade) ' 4/14/2000
        GoTo lblNextRec
    End If
    '----
    fTablSalHis.Open "SELECT * FROM HR_SALARY_HISTORY WHERE SH_ID=" & lngLastCurrentID&, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    fTablSalHis("SH_CURRENT") = False
    fTablSalHis.Update
        
    'Ticket #21400 - Do not update Vadim's HR_EMP_HISTORY table because the Rate level of the employee is
    'remaining same, only the actual salary is changing and this table only stores the Rate Level
    'Comment enhacement - Ticket #16115
    'City of Niagara Falls - Ticket #15542
    'If glbVadim And glbCompSerial = "S/N - 2276W" Then
    '    'Update previous salary record in Vadim's HR_EMP_HIST table with End Date
    '    Call Update_VadimDB_HR_EMP_HISTORY(oPayrollID, OEDate, "", "", "", "M", DateAdd("d", -1, CVDate(dlpEffectiveDate.Text)))
    'End If
    
    'Get employee's Hours/Day
    fglbDhrs = GetJHData(snapPosSal("SH_EMPNBR"), "JH_DHRS", 0)
    
    fTablSalHis.AddNew
    fTablSalHis("SH_COMPNO") = snapPosSal("SH_COMPNO")
    fTablSalHis("SH_EMPNBR") = snapPosSal("SH_EMPNBR")
    fTablSalHis("SH_EDATE") = CVDate(dlpNSDate.Text)
    fTablSalHis("SH_CURRENT") = True
    fTablSalHis("SH_SDATE") = snapPosSal("SH_SDATE")
    fTablSalHis("SH_SALCD") = txtSALCD
    fTablSalHis("SH_PAYROLL_ID") = snapPosSal("SH_PAYROLL_ID")
    fTablSalHis("SH_WHRS") = snapPosSal("SH_WHRS")

    fTablSalHis("SH_TRANSDATE") = Date
    
    'Ticket #20466 Franks 06/22/2011
    fTablSalHis("SH_VGROUP") = snapPosSal("SH_VGROUP")
    fTablSalHis("SH_VSTEP") = snapPosSal("SH_VSTEP")
    
    fTablSalHis("SH_PAYP") = snapPosSal("SH_PAYP")
    fTablSalHis("SH_PAYP_TABLE") = snapPosSal("SH_PAYP_TABLE")
    fTablSalHis("SH_SREAS_TABLE") = snapPosSal("SH_SREAS_TABLE")
    
    dblNewSalary = dblOSalary
    
    ''Ticket #22893 - WHSC - As of Date of the Salary to apply the Percent update on
    'If glbCompSerial = "S/N - 2448W" And optPct And Not glbMulti Then
    '    If IsDate(dlpAsofDate.Text) Then
    '        'Get Salary As of Date.
    '        xAsofDateSal = Get_AsOfDate_Salary(snapPosSal("SH_EMPNBR"), dlpAsofDate.Text, 0)
    '
    '        'If no salary is retrieved then used the last current salary
    '        If xAsofDateSal = 0 Then
    '            xAsofDateSal = dblOSalary
    '        End If
    '    End If
    'End If
    
    'For x% = 4 To 6
        If Len(clpCode(X%).Text) > 0 Then
            'If optStep Then
                JobSalCD = snapPosSal("JB_SALCD")
                'If comStep(x%) = "Next Step" Then
                 '   xGradeF = xGradeF + 1
                 '   If OSalCD = JobSalCD Then
                 '       dblAmnt = snapPosSal("JB_S" & Format(xGradeF, "##")) - dblNewSalary
                 '       dblPct = dblAmnt / dblOSalary
                 '   Else
                 '       If OSalCD = "H" And JobSalCD = "A" Then
                 '           dblAmnt = snapPosSal("JB_S" & Format(xGradeF, "##")) / (xWHrs * 52) - dblNewSalary
                 '           dblPct = dblAmnt / dblOSalary
                 '       End If
                 '       If OSalCD = "A" And JobSalCD = "H" Then
                 '           dblAmnt = snapPosSal("JB_S" & Format(xGradeF, "##")) * (xWHrs * 52) - dblNewSalary
                 '           dblPct = dblAmnt / dblOSalary
                 '       End If
                 '       If OSalCD = "D" And JobSalCD = "H" Then
                 '           dblAmnt = snapPosSal("JB_S" & Format(xGradeF, "##")) * fglbDhrs - dblNewSalary
                 '           dblPct = dblAmnt / dblOSalary
                 '       End If
                 '       If OSalCD = "D" And JobSalCD = "A" Then
                 '           dblAmnt = (snapPosSal("JB_S" & Format(xGradeF, "##")) / (xWHrs * 52) * fglbDhrs) - dblNewSalary
                 '           dblPct = dblAmnt / dblOSalary
                 '       End If
                 '       If glbCompSerial = "S/N - 2378W" And OSalCD <> "M" Then   'Town of Aurora
                 '           dblAmnt = snapPosSal("JB_S" & Format(xGradeF, "##") & "A") - dblNewSalary
                 '           dblPct = dblAmnt / dblOSalary
                 '       End If
                 '   End If
                'Else
                    'Hemu - testing
                    If txtSALCD <> JobSalCD Then
                        If txtSALCD = "A" And JobSalCD = "H" Then
                            dblAmnt = ((snapPosSal("JB_S" & Format(comStep, "##")) * snapPosSal("SH_WHRS")) * 52) - dblNewSalary
                        ElseIf txtSALCD = "H" And JobSalCD = "A" Then
                            dblAmnt = ((snapPosSal("JB_S" & Format(comStep, "##")) / snapPosSal("SH_WHRS")) / 52) - dblNewSalary
                        ElseIf txtSALCD = "D" And JobSalCD = "H" Then
                            dblAmnt = ((snapPosSal("JB_S" & Format(comStep, "##")) * fglbDhrs)) - dblNewSalary
                        ElseIf txtSALCD = "D" And JobSalCD = "A" Then
                            dblAmnt = (((snapPosSal("JB_S" & Format(comStep, "##")) / snapPosSal("SH_WHRS")) / 52) * fglbDhrs) - dblNewSalary
                        End If
                    ElseIf txtSALCD = JobSalCD Then
                        dblAmnt = snapPosSal("JB_S" & Format(comStep, "##")) - dblNewSalary
                    End If
                    dblPct = dblAmnt / dblOSalary
                    If glbCompSerial = "S/N - 2378W" And OSalCD <> JobSalCD And OSalCD <> "M" Then   'Town of Aurora
                        dblAmnt = snapPosSal("JB_S" & Format(comStep, "##") & "A") - dblNewSalary
                        dblPct = dblAmnt / dblOSalary
                    End If
                'End If
                dblNewSalary = dblNewSalary + dblAmnt
            'ElseIf optDollars Then
            '    dblAmnt = medAmountChng(x%)
            '    dblPct = dblAmnt / dblOSalary
            '    dblNewSalary = dblNewSalary + dblAmnt
            'ElseIf optPct Then
            '    'Ticket #22893 - WHSC - As of Date of the Salary to apply the Percent update on
            '    If glbCompSerial = "S/N - 2448W" And Not glbMulti Then
            '        dblPct = medAmountChng(x%)
            '        dblAmnt = dblPct * xAsofDateSal
            '        dblNewSalary = dblNewSalary + dblAmnt
            '        dblPct = (dblAmnt / dblOSalary) '* 100
            '    Else
            '        dblPct = medAmountChng(x%)
            '        dblAmnt = dblPct * dblOSalary
            '        dblNewSalary = dblNewSalary + dblAmnt
            '    End If
            'ElseIf optFixed Then
            '    dblNewSalary = medAmountChng(x%)
            '    dblAmnt = dblNewSalary - dblOSalary
            '    dblPct = dblAmnt / dblOSalary
            'End If
            
            fTablSalHis("SH_SREAS" & CStr(X% - 3)) = clpCode(X%).Text
            fTablSalHis("SH_SALPC" & CStr(X% - 3)) = dblPct
            fTablSalHis("SH_SALCHG" & CStr(X% - 3)) = dblAmnt
        End If
    'Next x%
    dblNewSalary = Round2DEC(dblNewSalary, snapPosSal("SH_EMPNBR")) 'added by raubrey 8/18/97
    
    If glbCompSerial = "S/N - 2288W" Then 'MUSASHI AUTO TKT#10845
        
        dblOSalary1 = dblNewSalary + dblOSalary2 + dblOSalary3 + dblOSalary4
    End If
    
    Dim RoundSal As Double, strRoundSal As String
    Dim strFirst As String
    'If cmbRound.ListIndex = 0 Then
    '    'dblNewSalary = CLng(dblNewSalary)  'Ticket #14699
    '    dblNewSalary = Round(dblNewSalary, cmbPrecision.Text)   'Ticket #14699
    'Else
        dblNewSalary = dblNewSalary
    'End If
    If glbCompSerial = "S/N - 2288W" Then 'MUSASHI AUTO TKT#10845
        fTablSalHis("SH_SALARY1") = dblNewSalary
        fTablSalHis("SH_SALARY") = dblOSalary1
    Else
        fTablSalHis("SH_SALARY") = dblNewSalary
    End If
    
    'If Len(dlpNextReviewDate.Text) > 0 Then
    '    fTablSalHis("SH_NEXTDAT") = CVDate(dlpNextReviewDate.Text)
    '    UpdateFollowup snapPosSal("SH_EMPNBR"), snapPosSal("SH_NEXTDAT"), CVDate(dlpNextReviewDate.Text), "SREV"
    'Else
    '    If IsDate(snapPosSal("SH_NEXTDAT")) Then
    '        If snapPosSal("SH_NEXTDAT") > CVDate(dlpNSDate) Then
    '            fTablSalHis("SH_NEXTDAT") = snapPosSal("SH_NEXTDAT")
    '        End If
    '    End If
    'End If
    fTablSalHis("SH_JOB") = snapPosSal("SH_JOB")
    fTablSalHis("SH_GRID") = snapPosSal("SH_GRID")
    fTablSalHis("SH_PAYROLL_ID") = snapPosSal("SH_PAYROLL_ID")
    fTablSalHis("SH_JOB_ID") = snapPosSal("SH_JOB_ID")
    'Jaddy changed for WFC Kipling Oct 31, 02
    If glbWFC And (snapPosSal("JB_ORG") = "NONE" Or snapPosSal("JB_ORG") = "EXEC") Then
        fTablSalHis("SH_COMPA_USER") = snapPosSal("SH_COMPA_USER")
        fTablSalHis("SH_COMPA_DOLLAR") = snapPosSal("SH_COMPA_DOLLAR")
        fTablSalHis("SH_BAND") = snapPosSal("SH_BAND")
        fTablSalHis("SH_MARKETLINE") = snapPosSal("SH_MARKETLINE")
        Call Set_WFC_COMPA(dblNewSalary)
    Else
        Call modSetCOMPA_GRADE(dblNewSalary, fTablSalHis) ' sets fglbCOMPA#, and fglbGRADE
    End If
    If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
        fTablSalHis("SH_PREMIUM") = snapPosSal("SH_PREMIUM")
        fTablSalHis("SH_TOTAL") = fTablSalHis("SH_SALARY") + snapPosSal("SH_PREMIUM") 'snapPosSal("SH_TOTAL")
        fTablSalHis("SH_VGROUP") = snapPosSal("SH_VGROUP")
        fTablSalHis("SH_VSTEP") = snapPosSal("SH_VSTEP")
    End If
     If glbCompSerial = "S/N - 2288W" Then 'MUSASHI AUTO TKT#10845
        If Not IsNull(dblOSalary2) Then
            fTablSalHis("SH_SALARY2") = dblOSalary2
        End If
        If Not IsNull(dblOSalary3) Then
            fTablSalHis("SH_SALARY3") = dblOSalary3
        End If
        If Not IsNull(dblOSalary4) Then
            fTablSalHis("SH_SALARY4") = dblOSalary4
        End If
        If Not IsNull(SalDetailCode1) Then
            fTablSalHis("SH_DETAILCODE1") = SalDetailCode1
        End If
        If Not IsNull(SalDetailCode2) Then
            fTablSalHis("SH_DETAILCODE2") = SalDetailCode2
        End If
        If Not IsNull(SalDetailCode3) Then
            fTablSalHis("SH_DETAILCODE3") = SalDetailCode3
        End If
        If Not IsNull(SalDetailCode4) Then
            fTablSalHis("SH_DETAILCODE4") = SalDetailCode4
        End If
        
        fTablSalHis("SH_EDATE1") = CVDate(dlpNSDate.Text)
        If IsDate(EDate2) Then
            fTablSalHis("SH_EDATE2") = EDate2
        End If
        If IsDate(EDate3) Then
            fTablSalHis("SH_EDATE3") = EDate3
        End If
        If IsDate(EDate4) Then
            fTablSalHis("SH_EDATE4") = EDate4
        End If
        
    End If
    
    fTablSalHis("SH_COMPA") = fglbCOMPA#
    fTablSalHis("SH_GRADE") = Format(fglbGRADE$, "00")

    fTablSalHis("SH_LDATE") = Now
    fTablSalHis("SH_LTIME") = Time$
    fTablSalHis("SH_LUSER") = glbUserID 'glbLEE_ID

'    If glbCompSerial = "S/N - 2214W" Then
'        Dim xToDate
'        If IsDate(dlpNextReviewDate.Text) Then
'            xToDate = dlpNextReviewDate.Text
'        Else
'            xToDate = DateAdd("D", -1, DateAdd("YYYY", 1, CVDate(dlpNSDate.Text)))
'        End If
'        Call ChangeOtherEarnAmount(fTablSalHis("SH_EMPNBR"), dblNewSalary, "A", dlpNSDate.Text, xToDate)
'    End If
    fTablSalHis.Update
    
    If glbVadim Then Call Transfer_Salary(fTablSalHis)
    
    'Ticket #28595 - Update employee's Attendance records as welll if selected
    If chkUpdAttendance Then Call Update_Attendance_SalaryInfo(fTablSalHis)
    
    'Ticket #21400 - Do not update Vadim's HR_EMP_HISTORY table because the Rate level of the employee is
    'remaining same, only the actual salary is changing and this table only stores the Rate Level
    'City of Niagara Falls - Ticket #15542
    'If glbVadim And glbCompSerial = "S/N - 2276W" Then
    '    'Add the salary record in Vadim's HR_EMP_HIST table storing the history of Rate changes
    '    Call Update_VadimDB_HR_EMP_HISTORY(fTablSalHis("SH_PAYROLL_ID"), CVDate(dlpEffectiveDate.Text), "", Val(fglbGRADE$), fTablSalHis("SH_JOB"), "A")
    'End If
    
    'George Mar 7,2006 #9965
    xSHID = fTablSalHis.Fields("SH_ID").Value
    'George Mar 7,2006 #9965
    
    fTablSalHis.Close
    
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    
    Call UpSal_Addnew
    Call updBenefitForSalDEPN(empNo&)   'jaddy 9/10/99
    
    'Call Employee_Master_Integration(EmpNo&)
    
    If glbAdv Or glbMediPay Then  'Ticket #12339 'Ticket #16574 for St. Johns Medipay
        Call Employee_Master_Integration(empNo&)
    End If
    
    If glbGP Then
        Call Salary_Integration(empNo&, , False, True, xSHID)
    Else
        Call Salary_Integration(empNo&)
    End If
    
    If glbCompSerial = "S/N - 2288W" Then 'MUSASHI AUTO TKT#10845
        NSalary = dblOSalary1
    Else
        NSalary = dblNewSalary
    End If
    
    NEDate = CVDate(dlpNSDate.Text)
    'If Len(dlpNextReviewDate.Text) > 0 Then
    '    NNDate = CVDate(dlpNextReviewDate.Text)
    'Else
        NNDate = ""
    'End If
'>>>>    If Not AUDITSALY() Then MsgBox "ERROR - AUDIT FILE"
lblNextRec:
    snapPosSal.MoveNext
Loop
'cmdView.Enabled = True
'cmdPrint.Enabled = True

modUpdateSelection_1 = True
MDIMain.panHelp(0).FloodType = 0

snapPosSal.Close
Screen.MousePointer = DEFAULT

If gsEMAIL_ONSALARY Then
    MailBody = ""
    SQLQ = "SELECT TT_EMPNBR,TT_COEFLAG,TT_WRKEMP,TT_GRADE,TT_FLAG,TT_TBEMP FROM HREMPWRK WHERE TT_WRKEMP='" & glbUserID & "'"
    SQLQ = SQLQ & "AND TT_FLAG = 'Y'"
    snapPosSal.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
    If Not snapPosSal.EOF Then
        If snapPosSal.RecordCount = 1 Then
            MailBody = "The following employee's salary has "
        Else
            MailBody = "The following employee's salaries have "
        End If
        'If optDollars Then
        '    MailBody = MailBody & " been increased by $" & medAmountChng(4) & "." & vbCrLf '& vbCrLf
        'End If
        'If optFixed Then
        '    MailBody = MailBody & " been changed to $" & medAmountChng(4) & "." & vbCrLf '& vbCrLf
        'End If
        'If optPct Then
        '    MailBody = MailBody & " been increased by " & medAmountChng(4) * 100 & "%." & vbCrLf '& vbCrLf
        'End If
        'If optStep Then
            MailBody = MailBody & " been changed to Step " & comStep & "." & vbCrLf '& vbCrLf
        'End If
        MailBody = MailBody & "Reason: " & GetTABLDesc("SDRC", clpNReason) & vbCrLf
        MailBody = MailBody & "Effective Date: " & dlpNSDate & vbCrLf & vbCrLf
        
    End If
    Do While Not snapPosSal.EOF
        MailBody = MailBody & GetEmpName(snapPosSal("TT_EMPNBR")) & vbCrLf
        snapPosSal.MoveNext
     Loop
     snapPosSal.Close
     If Len(MailBody) > 0 Then
         Call imgEmail_Click
     End If
End If

Exit Function

modUpdateSelection_1_Err:

MDIMain.panHelp(0).FloodType = 0
Screen.MousePointer = DEFAULT

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Updateentitle", "HREMP", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If
End Function

Private Function GetEmpName(xEmpNo)
Dim rsTemp As New ADODB.Recordset
Dim xStr, SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME FROM HREMP WHERE ED_EMPNBR=" & xEmpNo
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        xStr = "Employee #:" & (xEmpNo) & " Name: " & rsTemp("ED_FNAME") & " " & rsTemp("ED_SURNAME")
    End If
    rsTemp.Close
    GetEmpName = xStr
End Function

Public Sub imgEmail_Click()
Dim xEmail
On Error GoTo Email_Err

    If gsEMAIL_ONSALARY Then
        If Not UserEmailExist Then
            Exit Sub
        End If
        xEmail = GetComPreferEmail("EMAIL_ONSALARY")
        
        If Len(xEmail) > 0 Then
            frmSendEmail.txtTo.Text = xEmail 'GetComPreferEmail("EMAIL_ONSALARY")
            'frmSendEmail.txtCC.Text = GetCurEmpEmail 'xEmail
            frmSendEmail.txtSubject.Text = "info:HR Salary Change Notice"
            frmSendEmail.txtBody.Text = MailBody
            frmSendEmail.Show 1
        Else
            MsgBox "There is no email for Email Notification on Salary on Company Preference screen. "
        End If
    End If
    
    Exit Sub

Email_Err:
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail Salary", "SMTP", "SENDEMAIL")
    Resume Next

End Sub

Public Sub imgEmailP_Click()
Dim xEmail
On Error GoTo Email_Err

    If gsEMAIL_ONPOSITION Then
        If Not UserEmailExist Then
            Exit Sub
        End If
        xEmail = GetComPreferEmail("EMAIL_ONPOSITION")
        
        If Len(xEmail) > 0 Then
            frmSendEmail.txtTo.Text = xEmail 'GetComPreferEmail("EMAIL_ONSALARY")
            'frmSendEmail.txtCC.Text = GetCurEmpEmail 'xEmail
            frmSendEmail.txtSubject.Text = "info:HR Position Change Notice"
            frmSendEmail.txtBody.Text = MailBodyP
            frmSendEmail.Show 1
        Else
            MsgBox "There is no email for Email Notification on Position on Company Preference screen. "
        End If
    End If
    
    Exit Sub

Email_Err:
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail Position", "SMTP", "SENDEMAIL")
    Resume Next

End Sub

'Private Sub optDollars_Click(Value As Integer)
'Call setDollars
'
''Ticket #22893 - WHSC - As of Date of the Salary to apply the Percent update on
'If glbCompSerial = "S/N - 2448W" Then
'    'As of Date
'    Call ShowHide_AsofDate
'End If
'
'End Sub

Private Sub optDollars_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

'Private Sub optFixed_Click(Value As Integer)
'Call setDollars
'
''Ticket #22893 - WHSC - As of Date of the Salary to apply the Percent update on
'If glbCompSerial = "S/N - 2448W" Then
'    'As of Date
'    Call ShowHide_AsofDate
'End If
'End Sub

Private Sub optFixed_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

'Private Sub optPct_Click(Value As Integer)
'Call setDollars
'
''Ticket #22893 - WHSC - As of Date of the Salary to apply the Percent update on
'If glbCompSerial = "S/N - 2448W" Then
'    'As of Date
'    Call ShowHide_AsofDate
'End If
'
'End Sub
Private Sub optPct_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Function DelNoUpsal() 'frank 4/10/2000
Dim iTableCounter, TblName, SQLQ

    'Ticket #16717
    'gdbAdoIhr001.Execute "DELETE FROM HREMPWRK " & in_SQL(glbIHRDBW) & " WHERE TT_WRKEMP='" & glbUserID & "' "
    gdbAdoIhr001.Execute "DELETE FROM HREMPWRK " & in_SQL(glbIHRDBW) & " WHERE (TT_FLAG = 'Y' OR TT_FLAG = 'N') AND TT_WRKEMP = '" & glbUserID & "'"
    
    If snapNoUpSal.State <> 0 Then snapNoUpSal.Close
    snapNoUpSal.Open "SELECT TT_EMPNBR,TT_COEFLAG,TT_WRKEMP,TT_GRADE,TT_FLAG,TT_TBEMP FROM HREMPWRK WHERE TT_WRKEMP='" & glbUserID & "'", gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
End Function

Private Function GetGrade(xOLDSALA, xGRADEt, xSHJOB, xSALCDt, xWHRSt) ' Frank 4/12/2000
    Dim x1%, xSAL_TMP
    Dim xGrid2 As Boolean
'    snapJob.Requery
'    snapJob.Find "JB_CODE='" & xSHJOB & "'"
   
   xGrid2 = False
   
    xSAL_TMP = xOLDSALA
    xGRADEt = 0
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    'For x1% = 1 To 11
    'For x1% = 1 To 15
    For x1% = 1 To 20
        If xSALCDt = snapPosSal("JB_SALCD") Then
            If IsNumeric(snapPosSal("JB_S" & CStr(x1%))) And snapPosSal("JB_S" & CStr(x1%)) > 0 Then
                If xOLDSALA >= snapPosSal("JB_S" & CStr(x1%)) Then
                  xGRADEt = x1%
                End If
            End If
        Else
            If xSALCDt = "H" And snapPosSal("JB_SALCD") = "A" Then
                If xWHRSt = 0 Then
                    xGRADEt = 0
                Else
                    xSAL_TMP = snapPosSal("JB_S" & Format(x1%, "##")) / (xWHRSt * 52)
                End If
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    xSAL_TMP = snapPosSal("JB_S" & Format(x1%, "##") & "A")
                    xGrid2 = True
                End If
            End If
            If xSALCDt = "A" And snapPosSal("JB_SALCD") = "H" Then
                If xWHRSt = 0 Then
                    xGRADEt = 0
                Else
                    xSAL_TMP = snapPosSal("JB_S" & Format(x1%, "##")) * (xWHRSt * 52)
                End If
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    xSAL_TMP = snapPosSal("JB_S" & Format(x1%, "##") & "A")
                    xGrid2 = True
                End If
            End If
            
            If xSALCDt = "D" And snapPosSal("JB_SALCD") = "A" Then
                If xWHRSt = 0 Then
                    xGRADEt = 0
                Else
                    xSAL_TMP = snapPosSal("JB_S" & Format(x1%, "##")) / (xWHRSt * 52) * fglbDhrs
                End If
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    xSAL_TMP = snapPosSal("JB_S" & Format(x1%, "##") & "A")
                    xGrid2 = True
                End If
            End If
            If xSALCDt = "D" And snapPosSal("JB_SALCD") = "H" Then
                If xWHRSt = 0 Then
                    xGRADEt = 0
                Else
                    xSAL_TMP = snapPosSal("JB_S" & Format(x1%, "##")) * fglbDhrs
                End If
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    xSAL_TMP = snapPosSal("JB_S" & Format(x1%, "##") & "A")
                    xGrid2 = True
                End If
            End If
            
            
            If glbCompSerial = "S/N - 2378W" And xGrid2 = True Then 'Town of Aurora
                If xOLDSALA >= xSAL_TMP And snapPosSal("JB_S" & CStr(x1%) & "A") > 0 Then
                  xGRADEt = x1%
                End If
            Else
            If xOLDSALA >= xSAL_TMP And snapPosSal("JB_S" & CStr(x1%)) > 0 Then
              xGRADEt = x1%
                End If
            End If
        End If
    Next x1%
End Function

Private Sub NoUpSal_Addnew(zReason) 'Frank 4/13/2000
    Dim iFieldCounter, FldName
    snapNoUpSal.AddNew
    snapNoUpSal("TT_EMPNBR") = snapPosSal("SH_EMPNBR")
    snapNoUpSal("TT_COEFLAG") = snapPosSal("SH_CURRENT")
    snapNoUpSal("TT_WRKEMP") = glbUserID
    snapNoUpSal("TT_GRADE") = zReason
    snapNoUpSal("TT_FLAG") = "N"
    snapNoUpSal.Update
    SkipRec = SkipRec + 1
End Sub

Private Sub UpSal_Addnew() 'Frank 4/12/2000
    Dim iFieldCounter, FldName
    snapNoUpSal.AddNew
    snapNoUpSal("TT_EMPNBR") = snapPosSal("SH_EMPNBR")
    snapNoUpSal("TT_COEFLAG") = snapPosSal("SH_CURRENT")
    snapNoUpSal("TT_WRKEMP") = glbUserID
    snapNoUpSal("TT_FLAG") = "Y"
    snapNoUpSal.Update
    If Len(strEMPLIST) > 0 Then
        strEMPLIST = strEMPLIST & "," & snapPosSal("SH_EMPNBR")
    Else
        strEMPLIST = snapPosSal("SH_EMPNBR")
    End If
End Sub

Private Sub DelSal_Addnew(xID)
    snapNoUpSal.AddNew
    snapNoUpSal("TT_EMPNBR") = snapPosSal("SH_EMPNBR")
    snapNoUpSal("TT_WRKEMP") = glbUserID
    snapNoUpSal("TT_TBEMP") = xID
    snapNoUpSal("TT_FLAG") = "D"
    snapNoUpSal.Update
    If Len(strEMPLIST) > 0 Then
        strEMPLIST = strEMPLIST & "," & snapPosSal("SH_EMPNBR")
    Else
        strEMPLIST = snapPosSal("SH_EMPNBR")
    End If
End Sub

Private Function Cri_SetAll()
Dim X%, strRName$, strRName1$

Cri_SetAll = False

On Error GoTo modSetCriteria_Err

Call glbCri_DeptUN("")

Screen.MousePointer = HOURGLASS

'Print the Update Log report.
Call ExcelRpt_Log


'strRName$ = glbIHRREPORTS & "rzsalnu.rpt"
'Me.vbxCrystal.ReportFileName = strRName$
'If glbSQL Or glbOracle Then
'    Me.vbxCrystal.Connect = RptODBC_SQL
'Else
'    Me.vbxCrystal.Connect = "PWD=petman;"
'    Me.vbxCrystal.DataFiles(0) = glbIHRDB
'    Me.vbxCrystal.DataFiles(1) = glbIHRDB
'    Me.vbxCrystal.DataFiles(2) = glbIHRDB
'    Me.vbxCrystal.DataFiles(3) = glbIHRDB
'    Me.vbxCrystal.DataFiles(4) = glbIHRDBW
'    Me.vbxCrystal.DataFiles(5) = glbIHRDB
'
'    ' set security for database
'    'Me.vbxCrystal.Password = gstrAccPWord$
'    'Me.vbxCrystal.UserName = gstrAccUID$
'End If
'
''Ticket #25909
''Me.vbxCrystal.SelectionFormula = "{HREMPWRK.TT_FLAG} = 'N' AND " & glbstrSelCri
'Me.vbxCrystal.SelectionFormula = "{HREMPWRK.TT_FLAG} = 'N' AND {HREMPWRK.TT_WRKEMP} = '" & glbUserID & "'" 'AND " & glbstrSelCri
'
'' window title if appropriate
'Me.vbxCrystal.WindowTitle = "Employees Skipped Report"
'
''------------
'strRName1$ = glbIHRREPORTS & "rzsalup.rpt"
'Me.vbxCrystal1.ReportFileName = strRName1$
'If glbSQL Or glbOracle Then
'    Me.vbxCrystal1.Connect = RptODBC_SQL    'Ticket #25909
'Else
'    Me.vbxCrystal1.Connect = "PWD=petman;"
'    Me.vbxCrystal1.DataFiles(0) = glbIHRDB
'    Me.vbxCrystal1.DataFiles(1) = glbIHRDB
'    Me.vbxCrystal1.DataFiles(2) = glbIHRDB
'    Me.vbxCrystal1.DataFiles(3) = glbIHRDB
'    Me.vbxCrystal1.DataFiles(4) = glbIHRDBW
'    Me.vbxCrystal1.DataFiles(5) = glbIHRDB
'
'    ' set security for database
''    Me.vbxCrystal1.Password = gstrAccPWord$
''    Me.vbxCrystal1.UserName = gstrAccUID$
'End If
'
''Ticket #25909
''Me.vbxCrystal.SelectionFormula = "{HREMPWRK.TT_FLAG} = 'Y' AND " & glbstrSelCri
'Me.vbxCrystal1.SelectionFormula = "{HREMPWRK.TT_FLAG} = 'Y' AND {HREMPWRK.TT_WRKEMP} = '" & glbUserID & "'" 'AND " & glbstrSelCri
'
'' window title if appropriate
'Me.vbxCrystal1.WindowTitle = "Employees Updated Report"

Cri_SetAll = True

Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "No Update Salary Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function Round2DEC(tmpNUM, Optional xEmpNo) 'laura nov 10, 1997
Dim strNUM As String, X%

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
If glbCompSerial = "S/N - 2375W" And GetEmpData(xEmpNo, "ED_REGION") <> "S" Then 'City of Timmins
    Round2DEC = Round(tmpNUM, 2)
Else
    Round2DEC = Round(tmpNUM, glbCompDecHR)
End If
If glbWFC And glbCountry = "AUSTRALIA" Then
    Round2DEC = Round(tmpNUM, 4)
End If
End Function

'Private Sub optStep_Click(Value As Integer)
'Call setDollars
'
''Ticket #22893 - WHSC - As of Date of the Salary to apply the Percent update on
'If glbCompSerial = "S/N - 2448W" Then
'    'As of Date
'    Call ShowHide_AsofDate
'End If
'End Sub

Private Sub optStep_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Public Sub UpdateFollowup(EmpNbr, OldDate, NewDate, Code)
    Dim rsFollow As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR=" & EmpNbr
    
    If IsDate(OldDate) Then
        SQLQ = SQLQ & " AND EF_FDATE=" & Date_SQL(OldDate)
    End If
    rsFollow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsFollow.EOF Or rsFollow.BOF Or IsNull(OldDate) Then
        rsFollow.AddNew
        rsFollow("EF_EMPNBR") = EmpNbr
        rsFollow("EF_FREAS") = Code
        rsFollow("EF_COMPLETED") = False
    End If
    rsFollow("EF_FDATE") = NewDate
    rsFollow("EF_LDATE") = Date
    rsFollow("EF_LTIME") = Format(Now, "Medium Time")
    rsFollow.Update
End Sub

Private Sub CRLIST_EECAT()

comPayType.AddItem "Annually"
comPayType.AddItem "Hourly"
'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
'Ticket #17654 - open Daily option
'If glbCompSerial = "S/N - 2282W" Then
    comPayType.AddItem "Daily"
'End If
'cmbRound.AddItem "Yes"
'cmbRound.AddItem "No"
'cmbRound.ListIndex = 1

End Sub

'Private Sub setDollars()
'Dim strFmat$, x%
'
'If optDollars Or optFixed Then
'    strFmat$ = "$#,##0.00;($#,##0.00)"
'    If glbCompDecHR = 3 Then strFmat$ = "$#,##0.000;($#,##0.000)"
'    If glbCompDecHR = 4 Then strFmat$ = "$#,##0.0000;($#,##0.0000)"
'Else
'    strFmat$ = "0.######%"    'Jaddy 11/15
'End If
'
'For x% = 4 To 6
'    medAmountChng(x%) = ""
'    medAmountChng(x%).Format = strFmat$
'    If optFixed And x% > 4 Then
'        medAmountChng(x%).Visible = False
'        comStep(x%).Visible = False
'        clpCode(x%) = ""
'        clpCode(x%).Visible = False
'    Else
'        medAmountChng(x%).Visible = Not optStep
'        comStep(x%).Visible = optStep
'        clpCode(x%).Visible = True
'    End If
'Next x%
'End Sub

Sub Set_WFC_COMPA(dblNewSalary)
Dim xDollear
fglbCOMPA# = 0
If snapPosSal("SH_COMPA_USER") = "U" Then xDollear = snapPosSal("SH_COMPA_DOLLAR") Else xDollear = wfcSalState
If Val(xDollear) <> 0 Then
    fglbCOMPA# = Round((dblNewSalary / xDollear) * 100, 2)
End If
If Val(fglbCOMPA#) > 999.99 Then fglbCOMPA# = "999.99"
End Sub

Private Function wfcSalState() As Double
Dim SQLQ
On Error Resume Next
Dim rsWFC As New ADODB.Recordset
wfcSalState = 0

SQLQ = "SELECT MDOLLARS FROM WFC_Salary_Administration "
SQLQ = SQLQ & " WHERE [BAND]='" & snapPosSal("SH_BAND") & "'"
SQLQ = SQLQ & " AND [MARKETLINE]='" & snapPosSal("SH_MARKETLINE") & "'"
rsWFC.Open SQLQ, gdbAdoIhrWFC, adOpenStatic

If Not rsWFC.EOF Then
    If IsNumeric(rsWFC("MDOLLARS")) Then
        wfcSalState = rsWFC("MDOLLARS")
    End If
End If
rsWFC.Close
End Function

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
TF = True
UpdateState = OPENING
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False

'alpAPPNBR.Enabled = TF
End Sub

Public Property Get RelateMode() As RelateModeEnum
RelateMode = MassChanges
End Property

Public Property Get UpdateRight() As Boolean
'UpdateRight = gSec_Upd_Salary
UpdateRight = GetMassUpdateSecurities("Salary_MassUpdate", glbUserID) And GetMassUpdateSecurities("Position_MassUpdate", glbUserID)
End Property

Public Property Get Addable() As Boolean
Addable = True
End Property

Public Property Get Updateble() As Boolean
Updateble = False
End Property

Public Property Get Deleteble() As Boolean
Deleteble = False 'GetMassUpdateSecurities("Salary_MassUpdate", glbUserID) 'False
End Property

Public Property Get Printable() As Boolean
Printable = True 'False
End Property

'Private Sub Transfer_Position(rsNew As ADODB.Recordset)
'    Dim HRChangs As New Collection
'    Dim xEmpnbr
'    Dim xPayrollID
'    Dim xSDate
'
'    xEmpnbr = rsNew("JH_EMPNBR")
'    If rsNew("JH_PAYROLL_ID") = "" Or IsNull(rsNew("JH_PAYROLL_ID")) Then
'        xPayrollID = GetEmpData(rsNew("JH_EMPNBR"), "ED_PAYROLL_ID")
'    Else
'        xPayrollID = rsNew("JH_PAYROLL_ID")
'    End If
'    xSDate = rsNew("JH_SDATE")
'
'
'    If glbLambton Then
'        If isChanged_Field(HRChangs, OLambtonJob, txtLambtonJob) Then UpdateAudit = True
'    Else
'        If isChanged_Field(HRChangs, oJob, clpJob) Then UpdateAudit = True
'    End If
'    If OSDATE <> "" Then
'        If isChanged_Field(HRChangs, Str(OSDATE), dlpStartDate) Then UpdateAudit = True
'    Else
'        If isChanged_Field(HRChangs, OSDATE, dlpStartDate) Then UpdateAudit = True
'    End If
'    If isChanged_Field(HRChangs, OReason, clpCode(1)) Then UpdateAudit = True
'    If glbMulti Then
'        If isChanged_Field(HRChangs, oPayrollID, txtPayrollID) Then UpdateAudit = True
'        If glbVadim Then
'            If isChanged_Field(HRChangs, oENDDATE, dlpENDDATE) Then UpdateAudit = True
'            If isChanged_Field(HRChangs, oDeptNo, clpDept) Then UpdateAudit = True
'        End If
'    End If
'
'    'If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls - Ticket #14285
'        If CVDate(Format(rsNew("JH_SDATE"), "mm/dd/yyyy")) > CVDate(Format(Date, "mm/dd/yyyy")) Then
'            UptPositionDate = rsNew("JH_SDATE")
'        Else
'            UptPositionDate = Date
'        End If
'        Call Passing_Changes(HRChangs, Position, "M", UptPositionDate, glbLEE_ID, txtPayrollID.Text)
'    'Else
'    '    Call Passing_Changes(HRChangs, Position, "M", Date, glbLEE_ID, txtPayrollID.Text)
'    'End If
'
'    'Ticket #24565- DMuskoka - Transfer the Salary Effective Date at this time too if Same Salary New record, as Last Increment and Probation Date
'    If glbCompSerial = "S/N - 2373W" Then
'        dlpNSDate.DataField = "SH_EDATE"
'        If isChanged_Field(HRChangs1, "", dlpNSDate) Then UpdateAudit = True
'        dlpNSDate.DataField = ""
'        Call Passing_Changes(HRChangs1, Position, "M", dlpNSDate.Text, xEmpnbr, xPayrollID)
'    End If
'
'End Sub

Private Sub Transfer_Salary(rsNew As ADODB.Recordset)
    Dim rsEmpJob As New ADODB.Recordset
    Dim rsSal As New ADODB.Recordset
    Dim HRChanges As New Collection
    Dim UptSalaryDate As Date
    Dim HRSalary As New Collection
    Dim xEmpnbr
    Dim xPayrollID
    Dim xPHrs
    Dim xWHRS, xNiagaraWHRS
    Dim xEDate
    Dim xSalCD
    Dim UpdateAudit
    
    'Employee #
    xEmpnbr = rsNew("SH_EMPNBR")
    
    'Payroll ID
    If rsNew("SH_PAYROLL_ID") = "" Or IsNull(rsNew("SH_PAYROLL_ID")) Then
        xPayrollID = GetEmpData(rsNew("SH_EMPNBR"), "ED_PAYROLL_ID")
    Else
        xPayrollID = rsNew("SH_PAYROLL_ID")
    End If
    
    'Salary Effective Date
    xEDate = rsNew("SH_EDATE")
    
    'Hours per Week, Pay Period
    rsEmpJob.Open "SELECT JH_ID,JH_JOB,JH_DHRS,JH_PHRS,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & xEmpnbr & " AND JH_PAYROLL_ID='" & xPayrollID & "'", gdbAdoIhr001, adOpenForwardOnly
    xPHrs = 0
    xWHRS = 0
    If Not rsEmpJob.EOF Then
        xPHrs = Val(rsEmpJob("JH_PHRS") & "")
        xWHRS = Val(rsEmpJob("JH_WHRS") & "") 'Hemu - it was asssigning JH_DHRS - it should pass Weekly Hours
        xNiagaraWHRS = Val(rsEmpJob("JH_WHRS") & "")
        
        'City of Niagara Falls  = Dhrs = Hours Per Days from Position Master, fglbNiagPhrs = Pay Period
        If glbCompSerial = "S/N - 2276W" Then
            rsSal.Open "SELECT SH_EMPNBR, SH_PAYP, SH_WHRS FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & xEmpnbr & " AND SH_PAYROLL_ID = '" & xPayrollID & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsSal.EOF Then
                xPHrs = Val(rsSal("SH_PAYP") & "")
                xNiagaraWHRS = Val(rsSal("SH_WHRS") & "")
            End If
            rsSal.Close
            Set rsSal = Nothing
            xWHRS = GetJobData(rsEmpJob("JH_JOB"), "JB_DHRS", 1)
            xWHRS = Val(xWHRS & "")
        End If
    End If
    rsEmpJob.Close
   
    If glbCompSerial = "S/N - 2373W" Then   'DMuskoka  - Pass Total which includes Premium
        If isChanged_Salary(HRSalary, OTOTAL, rsNew("SH_TOTAL"), True) Then UpdateAudit = True
    Else
        If isChanged_Salary(HRSalary, oSal, rsNew("SH_SALARY"), True) Then UpdateAudit = True
    End If
    If isChanged_Salary(HRSalary, OSalCD, rsNew("SH_SALCD")) Then UpdateAudit = True
    
    If glbVadim And UpdateAudit Then
        'Ticket #21352 - City of Kawartha Lakes
        If glbCompSerial = "S/N - 2363W" Then
            Call Passing_Salary_Vadim(HRSalary, Salary, Date, xPHrs, xWHRS, xEmpnbr, xPayrollID, , xNiagaraWHRS)
        Else
            Call Passing_Salary_Vadim(HRSalary, Salary, xEDate, xPHrs, xWHRS, xEmpnbr, xPayrollID, , xNiagaraWHRS)
        End If
    End If
    
    'Ticket #24565 - District Municipality of Muskoka
    If glbCompSerial = "S/N - 2373W" Then
        'They want to transfer for 181W as well now - Nov 3rd 2014
        'Ticket #24565 - if Union = '181W' then do not transfer Probation Date, Level and After Probation
        'If GetEmpData(xEmpnbr, "ED_ORG") = "181W" Then
        '    'Do not transfer Probation Date, Level and After Probation
        'Else
            If isChanged_Field(HRChanges, oGrade, rsNew("SH_GRADE"), True) Then Debug.Print "" ' do nothing for the audit transfer
        'End If
    Else
        'Ticket #25469 - City of Campbell River - do not transfer Probation levels
        If glbCompSerial <> "S/N - 2458W" Then
            If isChanged_Field(HRChanges, oGrade, rsNew("SH_GRADE"), True) Then Debug.Print "" ' do nothing for the audit transfer
        End If
    End If
    
    If isChanged_Field(HRChanges, OEDate, rsNew("SH_EDATE")) Then UpdateAudit = True
    
    If glbCompSerial <> "S/N - 2373W" Then 'DMuskoka - Ticket #24565 - Do not transfer Next Review Date
        If isChanged_Field(HRChanges, ONDate, rsNew("SH_NEXTDAT")) Then UpdateAudit = True
    End If
    Call Passing_Changes(HRChanges, Salary, "M", Date, xEmpnbr, xPayrollID)

End Sub

Private Function modDelRecs()
Dim lngLastCurrentID&, X%
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim JobInfo As Boolean
Dim prec%, curSalary#
Dim fTablSalHis  As New ADODB.Recordset
Dim dblPct, dblAmnt, dblOSalary, DtTm  As Variant, dblNewSalary As Double
Dim salarystep, xGradeF
Dim JobSalCD, xWHRS
Dim SQLQ, xDelAmt

Screen.MousePointer = HOURGLASS

Call DelNoUpsal

modDelRecs = False

On Error GoTo CrFollow_Err 'laura nov 12, 1997
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 5
SkipRec = 0
Do While Not snapPosSal.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / (lngRecs&)))
    MDIMain.panHelp(0).FloodPercent = pct%
    
    empNo& = snapPosSal("SH_EMPNBR")
    SQLQ = "SELECT SH_EMPNBR,SH_ID FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & empNo& & " "
    SQLQ = SQLQ & "AND SH_EDATE = " & Date_SQL(dlpNSDate.Text) & " "
    fTablSalHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not fTablSalHis.EOF Then
        Call DelSal_Addnew(fTablSalHis("SH_ID"))
    End If
    fTablSalHis.Close

lblNextRec:
    snapPosSal.MoveNext
Loop
snapPosSal.Close

xDelAmt = 0
'Delete Salary record - Begin
SQLQ = "SELECT TT_EMPNBR,TT_WRKEMP,TT_TBEMP FROM HREMPWRK WHERE TT_WRKEMP='" & glbUserID & "'"
fTablSalHis.Open SQLQ, gdbAdoIhr001W, adOpenStatic
If Not fTablSalHis.EOF Then
    lngRecs& = fTablSalHis.RecordCount
    prec% = -1
    Do While Not fTablSalHis.EOF
        prec% = prec% + 1
        pct% = Int(100 * (prec% / (lngRecs&)))
        MDIMain.panHelp(0).FloodPercent = pct%
        xDelAmt = xDelAmt + 1
        empNo& = fTablSalHis("TT_EMPNBR")
        If Not IsNull(fTablSalHis("TT_TBEMP")) Then
            SQLQ = "DELETE FROM HR_SALARY_HISTORY WHERE SH_ID = " & fTablSalHis("TT_TBEMP")
            gdbAdoIhr001.Execute SQLQ
        End If
        Call Employee_Master_Integration(empNo&)
        'George Mar 9,2006 #9965
        If glbGP Then
            Call Salary_Integration(fTablSalHis("TT_EMPNBR"), , True, False, fTablSalHis("TT_TBEMP"))
        End If
        'George Mar 9,2006 #9965
        fTablSalHis.MoveNext
    Loop
    fTablSalHis.MoveFirst
    prec% = -1
    Do While Not fTablSalHis.EOF
        prec% = prec% + 1
        pct% = Int(100 * (prec% / (lngRecs&)))
        MDIMain.panHelp(0).FloodPercent = pct%
        empNo& = fTablSalHis("TT_EMPNBR")
        Call Set_Current_Flag(empNo&)
        fTablSalHis.MoveNext
    Loop
End If
fTablSalHis.Close
'Delete Salary record - End

'cmdView.Enabled = True
'cmdPrint.Enabled = True


modDelRecs = True
Screen.MousePointer = DEFAULT
If xDelAmt = 0 Then
    MsgBox "No Record Deleted!"
Else
    MsgBox xDelAmt & " Records Deleted Successfully!"
End If

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(0).FloodPercent = 0

Exit Function

CrFollow_Err:


glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Delete RECORDS", "HR_SALARY_HISTORY", "UPDATE TABLES")
Resume Next
End Function

Private Sub Set_Current_Flag(xEmpNo)
Dim SQLQ As String, Msg$
Dim dyn_HRSALHIS As New ADODB.Recordset

On Error GoTo SCFError
If glbMulti Then Exit Sub

SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & xEmpNo & " "
SQLQ = SQLQ & "ORDER BY SH_EDATE DESC, SH_ID DESC "
dyn_HRSALHIS.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic

If dyn_HRSALHIS.RecordCount < 1 Then
    Exit Sub
End If

If dyn_HRSALHIS.RecordCount > 0 Then dyn_HRSALHIS.MoveFirst
dyn_HRSALHIS("SH_CURRENT") = True
dyn_HRSALHIS.Update

Do Until dyn_HRSALHIS.EOF
    dyn_HRSALHIS.MoveNext
    If dyn_HRSALHIS.EOF Then Exit Do
    If dyn_HRSALHIS("SH_CURRENT") <> 0 Then
        dyn_HRSALHIS("SH_CURRENT") = False
        dyn_HRSALHIS.Update
    End If
Loop
dyn_HRSALHIS.Close

Exit Sub

SCFError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_SALARY_HISTORY", "Update")
Resume Next

End Sub

Public Sub cmdDelete_Click()
Dim Title$, Msg$, DgDef As Variant, Response%
Dim I As Integer
clpDiv.SetFocus
On Error GoTo AddN_Err
If Not gSec_Upd_Salary Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If


txtSALCD = Left(comPayType, 1)

fglAddDel = "Delete"

If Not chkMUEmpPosSal() Then Exit Sub

If Not modGet_PosSal_Records() Then Exit Sub       ' get selection - form level

Title$ = "Update Salary"

Msg$ = ""

If snapPosSal.BOF And snapPosSal.EOF Then
    Msg$ = Msg$ & "No Employees with this selection criteria exist!  " & Chr(10)
    Msg$ = Msg$ & "Please ensure the Hourly/Annually selection box is set correctly to match the group you want to update." & Chr(10)
    Msg$ = Msg$ & Chr(10)
    DgDef = MB_OK + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
    MsgBox Msg$, , Title$
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

lngRecs& = snapPosSal.RecordCount
Msg$ = Msg$ & lngRecs& & " Records to process" & Chr(10) & "Proceed?"

DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
If Response% = IDNO Then ' Evaluate response
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

Msg$ = "Do you want to print a list of employees updated?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.

strEMPLIST = ""
If Not modDelRecs() Then Exit Sub

If Response% = IDYES Then    ' Yes response
    If lngRecs& - SkipRec > 0 Then
        'Call set_PrintState(False)
        Screen.MousePointer = HOURGLASS
        
        'Call getWSQLQ("U")
        
        ' report name
      
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList.rpt"
    
        Me.vbxCrystal.Formulas(0) = "rTitle='Mass Update Salary - Employee Details'"
        'set location for database tables
        If Len(glbstrSelCri) >= 0 Then
            Me.vbxCrystal.SelectionFormula = getWSQLQRPT
        End If
        'If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        'Else
        '    Me.vbxCrystal.Connect = "PWD=petman;"
        '            Me.vbxCrystal.DataFiles(0) = glbIHRDB
        'End If
        
        ' window title if appropriate
        Me.vbxCrystal.WindowTitle = "Employees-updated Report"
        
        Me.vbxCrystal.Destination = 0
        Screen.MousePointer = DEFAULT
        Me.vbxCrystal.Action = 1
        vbxCrystal.Reset
    
    End If
End If

Screen.MousePointer = DEFAULT
'MsgBox "Records Added Successfully!"

Exit Sub

AddN_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDelete", "HR_SALARY_HISTORY", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Function getWSQLQRPT() As String
'getWSQLQRPT = glbSeleDeptUn    'Department security removed by Bryan, redundant, this is a list of changes, whether they have security is irrelevant at this point
'If Len(clpDept.Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_DEPTNO} = '" & clpDept.Text & "')"
'If Len(clpDiv.Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_DIV} = '" & clpDiv.Text & "') "
'If Len(clpCode(1).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_LOC} = '" & clpCode(1).Text & "') "
'If Len(clpCode(2).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_ORG} = '" & clpCode(2).Text & "') "
'If Len(clpCode(3).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_EMP} = '" & clpCode(3).Text & "') "
'If Len(clpCode(5).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_REGION} = '" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(5).Text & "') "
'If Len(clpCode(6).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_ADMINBY} = '" & clpCode(6).Text & "') "
'If Len(clpCode(7).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_BENEFIT_GROUP} = '" & clpCode(7).Text & "') "
'If Len(clpPT.Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_PT} = '" & clpPT.Text & "') "
If Len(strEMPLIST) > 0 Then getWSQLQRPT = " ({HREMP.ED_EMPNBR} IN [" & strEMPLIST & "]) "

End Function

'Private Sub ShowHide_AsofDate()
'    'Ticket #22893 - WHSC - As of Date of the Salary to apply the Percent update on
'    If optPct And Not glbMulti Then
'        lblAsofDate.Visible = True
'        dlpAsofDate.Visible = True
'        imgHelp.Visible = True
'    ElseIf Not optPct Then
'        lblAsofDate.Visible = False
'        dlpAsofDate.Visible = False
'        imgHelp.Visible = False
'    End If
'End Sub

Private Function Get_AsOfDate_Salary(xEmpNo, xAsofDate, xDefault)
    Dim rsSal As New ADODB.Recordset
    Dim SQLQ As String
    
    Get_AsOfDate_Salary = xDefault
    
    SQLQ = "SELECT TOP 1 SH_EMPNBR,SH_SALARY,SH_SALCD,SH_EDATE,SH_CURRENT FROM HR_SALARY_HISTORY"
    SQLQ = SQLQ & " WHERE SH_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND SH_EDATE >= " & Date_SQL(xAsofDate)
    SQLQ = SQLQ & " ORDER BY SH_EDATE ASC"
    rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsSal.EOF Then
        Get_AsOfDate_Salary = rsSal("SH_SALARY")
    Else
        Get_AsOfDate_Salary = 0
    End If
    rsSal.Close
    Set rsSal = Nothing
    
End Function

Private Function modUpdateSelection_On_LAYOFF()
    Dim SQLQ As String
    Dim rsJOB As New ADODB.Recordset
    Dim rsSal As New ADODB.Recordset
    Dim rsEmp As New ADODB.Recordset
    Dim lngLastCurrentID&
    Dim prec%, pct%
    Dim xStr As String
    
    'Update Position
    ' - Make current Position Un-Current
    ' - Add a new LayOff Position with default values
    ' - Update HREMP - Department = 4900, Divisio = LAYO, Location = NW, *Section (Admin By) = 4900

    'New Position Record
    On Error GoTo modUpdateSelection_On_LAYOFF_Err
    
    modUpdateSelection_On_LAYOFF = False
    
    'Call DelNoUpsal
    
    Screen.MousePointer = HOURGLASS
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 5
    
    SkipRec = 0
    MailBodyP = ""
    
    'Clear the log table
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute "DELETE FROM HREMPHIS_WRK WHERE EE_WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001W.CommitTrans
    Call Pause(1)
    
    Do While Not snapPosSal.EOF
        prec% = prec% + 1
        pct% = Int(100 * (prec% / (lngRecs&)))
        MDIMain.panHelp(0).FloodPercent = pct%
        
        empNo& = snapPosSal("JH_EMPNBR")
            
        'Clear fields
        oJob = "": OReason = "": OSDATE = "": ODHRS = "": oWHRS = "": oPHRS = "": oRepAut = "": oSHIFT = "": oFTENum = "": oFTEHrs = ""
        oOrg = "": oPTFT = "": oComment = "": oComment2 = "": oRepAut2 = "": oRepAut3 = "": oRepAut4 = "": oDiv = "": oDeptNo = "": oEmp = ""
        oGLNo = "": oSect = "": oRegion = "": oPosCtrl = "": oPayrollID = "": oGrid = "": oPayCateg = "": oBillRate = ""
        oPosStatus = ""
        
        'Retrieve existing Current Position Data
        oJob = snapPosSal("JH_JOB")
        OReason = snapPosSal("JH_JREASON")
        OSDATE = snapPosSal("JH_SDATE")
        If Not IsNull(snapPosSal("JH_DHRS")) Then ODHRS = snapPosSal("JH_DHRS")
        If Not IsNull(snapPosSal("JH_WHRS")) Then oWHRS = snapPosSal("JH_WHRS")
        If Not IsNull(snapPosSal("JH_PHRS")) Then oPHRS = snapPosSal("JH_PHRS")
        If Not IsNull(snapPosSal("JH_REPTAU")) Then oRepAut = snapPosSal("JH_REPTAU")
        If Not IsNull(snapPosSal("JH_SHIFT")) Then oSHIFT = snapPosSal("JH_SHIFT")
        If Not IsNull(snapPosSal("JH_FTENUM")) Then oFTENum = snapPosSal("JH_FTENUM")
        If Not IsNull(snapPosSal("JH_FTEHRS")) Then oFTEHrs = snapPosSal("JH_FTEHRS")
        oOrg = snapPosSal("JH_ORG")
        oPTFT = snapPosSal("JH_PT")
        oComment = snapPosSal("JH_COMMENT")
        oComment2 = snapPosSal("JH_COMMENT2")
        If Not IsNull(snapPosSal("JH_REPTAU2")) Then oRepAut2 = snapPosSal("JH_REPTAU2")
        If Not IsNull(snapPosSal("JH_REPTAU3")) Then oRepAut3 = snapPosSal("JH_REPTAU3")
        If Not IsNull(snapPosSal("JH_REPTAU4")) Then oRepAut4 = snapPosSal("JH_REPTAU4")
        'oLeadHand = rsJOB("JH_LEADHAND")
        'oLabourCD = rsJOB("JH_LABOURCD")
        'oLabourDate = rsJOB("JH_LABOUREDATE")
        'oUsrLabel = rsJOB("JH_USRLABEL")
        'oUsrCheck = rsJOB("JH_USRCHECK")
        'oUsrDate = rsJOB("JH_USREDATE")
        'oUsrLabel2 = rsJOB("JH_USRLABEL2")
        'oUsrCheck2 = rsJOB("JH_USRCHECK2")
        'oUsrDate2 = rsJOB("JH_USREDATE2")
        'oUsrLabel3 = rsJOB("JH_USRLABEL3")
        'oUsrCheck3 = rsJOB("JH_USRCHECK3")
        'oUsrDate3 = rsJOB("JH_USREDATE3")
        oDiv = snapPosSal("JH_DIV")
        oDeptNo = snapPosSal("JH_DEPTNO")
        oEmp = snapPosSal("JH_EMP")
        oGLNo = snapPosSal("JH_GLNO")
        oSect = snapPosSal("JH_SECTION")
        oRegion = snapPosSal("JH_REGION")
        oPosCtrl = snapPosSal("JH_POSITION_CONTROL")
        oPayrollID = snapPosSal("JH_PAYROLL_ID")
        oGrid = snapPosSal("JH_GRID")
        oPayCateg = snapPosSal("JH_PAYROLL_CATEGORY")
        oBillRate = snapPosSal("JH_BILLINGRATE")
        'oENDDATE = snapPosSal("JH_ENDDATE")
        'oEndReason = snapPosSal("JH_ENDREAS")
        oPosStatus = snapPosSal("JH_ESTATUS")

        If IsNull(snapPosSal("JH_GRID")) Then
            OLambtonJob = oJob
        Else
            OLambtonJob = Left(snapPosSal("JH_GRID"), 1) & oJob & Mid(snapPosSal("JH_GRID"), 2)
        End If

        lngLastCurrentID& = snapPosSal("JH_ID")

        'Remove the Current check from the existing current position record
        rsJOB.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID=" & lngLastCurrentID&, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rsJOB("JH_CURRENT") = False
        rsJOB("JH_ENDDATE") = DateAdd("d", -1, dlpNSDate.Text)
        rsJOB("JH_ENDREAS") = clpNReason.Text
        rsJOB("JH_LDATE") = Format(Now, "Short Date")
        rsJOB("JH_LTIME") = Time$
        rsJOB("JH_LUSER") = glbUserID
        rsJOB.Update

        'Add a new Current Position records
        rsJOB.AddNew
        rsJOB("JH_COMPNO") = "001"
        rsJOB("JH_EMPNBR") = empNo&
        rsJOB("JH_JOB") = clpNJob.Text
        rsJOB("JH_SDATE") = dlpNSDate.Text
        If ODHRS <> "" And Not IsNull(ODHRS) Then rsJOB("JH_DHRS") = ODHRS
        If oWHRS <> "" And Not IsNull(oWHRS) Then rsJOB("JH_WHRS") = oWHRS
        If oPHRS <> "" And Not IsNull(oPHRS) Then rsJOB("JH_PHRS") = oPHRS
        rsJOB("JH_JREASON") = clpNReason.Text
        rsJOB("JH_CURRENT") = True
        If oRepAut <> "" And Not IsNull(oRepAut) Then rsJOB("JH_REPTAU") = oRepAut
        rsJOB("JH_SHIFT") = oSHIFT
        If oFTENum <> "" And Not IsNull(oFTENum) Then rsJOB("JH_FTENUM") = oFTENum
        If oFTEHrs <> "" And Not IsNull(oFTEHrs) Then rsJOB("JH_FTEHRS") = oFTEHrs
        rsJOB("JH_ORG") = oOrg
        rsJOB("JH_PT") = oPTFT
        If IsDate(dlpExpRtnDate.Text) Then
            'Ticket #29480 - City of Niagara Falls wants to add their own Notes 1
            If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls only
                'rsJOB("JH_COMMENT") = "Expected Return Date:" & dlpExpRtnDate.Text
                rsJOB("JH_COMMENT") = txtComment.Text
            Else
                'rsJOB("JH_COMMENT") = Left(oComment & " ExpRtnDt:" & dlpExpRtnDate.Text, 50)
                rsJOB("JH_COMMENT") = txtComment.Text
            End If
            'oComment = rsJOB("JH_COMMENT")
        End If
        
        'Ticket #29480 - City of Niagara Falls wants to add their own Notes 1
        rsJOB("JH_COMMENT") = txtComment.Text
        
        'Ticket #29382 - They want Notes 1 and Notes 2 to be user input fields
        'Ticket #29382 - They do not want to carry forward the Notes 2
        'rsJOB("JH_COMMENT2") = oComment2
        rsJOB("JH_COMMENT2") = txtComments2.Text
    
        If oRepAut2 <> "" And Not IsNull(oRepAut2) Then rsJOB("JH_REPTAU2") = oRepAut2
        If oRepAut3 <> "" And Not IsNull(oRepAut3) Then rsJOB("JH_REPTAU3") = oRepAut3
        If oRepAut4 <> "" And Not IsNull(oRepAut4) Then rsJOB("JH_REPTAU4") = oRepAut4
        rsJOB("JH_DIV") = clpNDiv.Text  '"LAYO"
        rsJOB("JH_DEPTNO") = clpNDept.Text  '"4900"
        rsJOB("JH_EMP") = oEmp
        rsJOB("JH_GLNO") = oGLNo
        rsJOB("JH_SECTION") = oSect
        rsJOB("JH_REGION") = oRegion
        rsJOB("JH_POSITION_CONTROL") = oPosCtrl
        rsJOB("JH_PAYROLL_ID") = oPayrollID
        rsJOB("JH_GRID") = oGrid
        rsJOB("JH_PAYROLL_CATEGORY") = oPayCateg
        rsJOB("JH_BILLINGRATE") = oBillRate
        'rsJOB("JH_ENDDATE") = rEndDate
        'rsJOB("JH_ENDREAS") = rEndReason
        rsJOB("JH_ESTATUS") = oPosStatus
        
        rsJOB("JH_LDATE") = Format(Now, "Short Date")
        rsJOB("JH_LTIME") = Time$
        rsJOB("JH_LUSER") = glbUserID
        rsJOB.Update

        'Ticket #29480 - City of Niagara Falls wants email to send out on Position Change to Lay Off
        If gsEMAIL_ONPOSITION Then
            MailBodyP = MailBodyP & GetEmpName(rsJOB("JH_EMPNBR")) & vbCrLf
        End If

        nJobID = rsJOB("JH_ID")
                
        'Document Attachment
        '7.9 Enhancement
        If gsAttachment_DB Then
            If glbDocNewRecord Then 'New Record only
                If Len(glbDocImpFile) > 0 Then
                    'glbJob = xID
                    glbJob = clpNJob.Text
                    glbSDate = dlpNSDate.Text
                    Call AttachmentAdd(empNo&, glbDocImpFile, glbDocType, glbDocDesc)
                End If
            End If
            'glbDocImpFile = ""
        End If

        rsJOB.Close
        Set rsJOB = Nothing

        'Update HREMP with hard coded Dept, Div and Location
        If Not glbMulti Then
            If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls only
                rsEmp.Open "SELECT * FROM HREMP WHERE ED_EMPNBR = " & empNo&, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsEmp.EOF Then
                    If oDeptNo = "" Or IsNull(oDeptNo) Then oDeptNo = rsEmp("ED_DEPTNO")
                    If Not EmpHisCalc(1, empNo&, clpNDept.Text, "", "", "", "", "", "", Date, , , , , , , oDeptNo) Then MsgBox "EMPHIS Error "  'Employee History Update
                    rsEmp("ED_DEPTNO") = clpNDept.Text '"4900"
                    
                    If oDiv = "" Or IsNull(oDiv) Then oDiv = rsEmp("ED_DIV")
                    If Not EmpHisCalc(1, empNo&, "", clpNDiv, "", "", "", "", "", Date, , , , , , , oDiv) Then MsgBox "EMPHIS Error "       'Employee History Update
                    rsEmp("ED_DIV") = clpNDiv.Text  '"LAYO"
                    
                    oLoc = rsEmp("ED_LOC")
                    If Not EmpHisCalc(2, empNo&, "", "", "", "", "", "", "", Date, "LOC", clpNLoc.Text, , , , , oLoc) Then MsgBox "EMPHIS Error "        'Employee History Update
                    rsEmp("ED_LOC") = clpNLoc.Text  '"NW"
                    
                    oAdminBy = rsEmp("ED_ADMINBY")      'Ticket #29480 - Additional fields
                    If Not EmpHisCalc(2, empNo&, "", "", "", "", "", "", "", Date, "ADMINBY", clpNAdminBy.Text, , , , , oAdminBy) Then MsgBox "EMPHIS Error "      'Ticket #29480 - Additional fields
                    rsEmp("ED_ADMINBY") = clpNAdminBy.Text  '"4900"
                    
                    rsEmp("ED_LDATE") = Format(Now, "Short Date")
                    rsEmp("ED_LTIME") = Time$
                    rsEmp("ED_LUSER") = glbUserID
                    rsEmp.Update
                End If
                rsEmp.Close
                Set rsEmp = Nothing
            Else
                rsEmp.Open "SELECT * FROM HREMP WHERE ED_EMPNBR = " & empNo&, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsEmp.EOF Then
                    If oDeptNo = "" Or IsNull(oDeptNo) Then oDeptNo = rsEmp("ED_DEPTNO")
                    If Len(Trim(clpNDept.Text)) > 0 Then
                        If Not EmpHisCalc(1, empNo&, clpNDept.Text, "", "", "", "", "", "", Date, , , , , , , oDeptNo) Then MsgBox "EMPHIS Error "
                        rsEmp("ED_DEPTNO") = clpNDept.Text '"4900"
                    End If
                    
                    If oDiv = "" Or IsNull(oDiv) Then oDiv = rsEmp("ED_DIV")
                    If Len(Trim(clpNDiv.Text)) > 0 Then
                        If Not EmpHisCalc(1, empNo&, "", clpNDiv, "", "", "", "", "", Date, , , , , , , oDiv) Then MsgBox "EMPHIS Error "
                        rsEmp("ED_DIV") = clpNDiv.Text  '"LAYO"
                    End If
                    
                    oLoc = rsEmp("ED_LOC")
                    If Len(Trim(clpNLoc.Text)) > 0 Then
                        If Not EmpHisCalc(2, empNo&, "", "", "", "", "", "", "", Date, "LOC", clpNLoc.Text, , , , , oLoc) Then MsgBox "EMPHIS Error "
                        rsEmp("ED_LOC") = clpNLoc.Text  '"NW"
                    End If
                                        
                    oAdminBy = rsEmp("ED_ADMINBY")      'Ticket #29480 - Additional fields
                    If Len(Trim(clpNAdminBy.Text)) > 0 Then
                        If Not EmpHisCalc(2, empNo&, "", "", "", "", "", "", "", Date, "ADMINBY", clpNAdminBy.Text, , , , , oAdminBy) Then MsgBox "EMPHIS Error "     'Ticket #29480 - Additional fields
                        rsEmp("ED_ADMINBY") = clpNAdminBy.Text  '"4900"
                    End If
                    
                    If Len(Trim(clpNDept.Text)) > 0 Or Len(Trim(clpNDiv.Text)) > 0 Or Len(Trim(clpNLoc.Text)) > 0 Or Len(Trim(clpNAdminBy.Text)) > 0 Then
                        rsEmp("ED_LDATE") = Format(Now, "Short Date")
                        rsEmp("ED_LTIME") = Time$
                        rsEmp("ED_LUSER") = glbUserID
                        rsEmp.Update
                    End If
                End If
                rsEmp.Close
                Set rsEmp = Nothing
            End If
        End If
        
        'Update the Audit Log
        If Not AUDITPOS(empNo&) Then MsgBox "ERROR - AUDIT FILE"
              
        
        'Update Salary
        ' - Make current Salary Un-Current
        ' - Add a new LayOff Salary with default values

        'Create a new Salary record as well with same salary
        SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & empNo& & " AND SH_CURRENT<>0 AND SH_JOB ='" & oJob & "'"
        If glbMulti Then
            SQLQ = SQLQ & " AND SH_SDATE = " & Date_SQL(OSDATE)
        End If
        rsSal.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsSal.EOF Then
            'Clear fields
            OEDate = "": oSal = "": OSalCD = "": oGrade = "": oPayP = "": oSReas1 = "": oSReas2 = "": oSReas3 = "": oCompa = "": ONDate = ""
            oSComment = "": oSComment2 = "": OTOTAL = "": oGrid = "": oPayrollID = "": oWHRS = "": OPremium = "": oVGroup = ""
            oVStep = ""
            
            'Retrieve Data
            OEDate = rsSal("SH_EDATE")
            oSal = rsSal("SH_SALARY")
            OSalCD = rsSal("SH_SALCD")
            oGrade = rsSal("SH_GRADE")
            oPayP = rsSal("SH_PAYP")
            If Not IsNull(rsSal("SH_SREAS1")) Then oSReas1 = rsSal("SH_SREAS1")
            If Not IsNull(rsSal("SH_SREAS2")) Then oSReas2 = rsSal("SH_SREAS2")
            If Not IsNull(rsSal("SH_SREAS3")) Then oSReas3 = rsSal("SH_SREAS3")
            If Not IsNull(rsSal("SH_COMPA")) Then oCompa = rsSal("SH_COMPA")
            If Not IsNull(rsSal("SH_NEXTDAT")) Then ONDate = rsSal("SH_NEXTDAT")
            oSComment = rsSal("SH_COMMENT")
            oSComment2 = rsSal("SH_COMMENT2")
            If Not IsNull(rsSal("SH_TOTAL")) Then OTOTAL = rsSal("SH_TOTAL")
            oGrid = rsSal("SH_GRID")
            oPayrollID = rsSal("SH_PAYROLL_ID")
            If Not IsNull(rsSal("SH_WHRS")) Then oWHRS = rsSal("SH_WHRS")
            OPremium = rsSal("SH_PREMIUM")
            oVGroup = rsSal("SH_VGROUP")
            oVStep = rsSal("SH_VSTEP")
            
            'oSalPC1 = rsSal("SH_SALPC1")
            'oSalChg1 = rsSal("SH_SALCHG1")
            'oSalPC2 = rsSal("SH_SALPC2")
            'oSalChg2 = rsSal("SH_SALCHG2")
            'oSalPC3 = rsSal("SH_SALPC3")
            'oSalChg3 = rsSal("SH_SALCHG3")
            
            'Remove the Current check from the existing current salary record
            rsSal("SH_CURRENT") = False
            rsSal("SH_TRANSDATE") = Now
            rsSal("SH_LDATE") = Format(Now, "Short Date")
            rsSal("SH_LTIME") = Time$
            rsSal("SH_LUSER") = glbUserID
            rsSal.Update
        
            'Add a new Current Salary records
            rsSal.AddNew
            rsSal("SH_COMPNO") = "001"
            rsSal("SH_EMPNBR") = empNo&
            rsSal("SH_CURRENT") = True
            rsSal("SH_JOB") = clpNJob.Text
            rsSal("SH_JOB_ID") = nJobID
            rsSal("SH_SDATE") = dlpNSDate.Text
            rsSal("SH_EDATE") = dlpNSDate.Text
            rsSal("SH_TRANSDATE") = Format(Now, "SHORT DATE")
            rsSal("SH_SALARY") = oSal
            rsSal("SH_SALCD") = OSalCD
            If oWHRS <> "" And Not IsNull(oWHRS) Then rsSal("SH_WHRS") = oWHRS
            rsSal("SH_GRADE") = oGrade
            rsSal("SH_PAYP") = oPayP
            rsSal("SH_SREAS1") = clpNReason.Text
            rsSal("SH_SALPC1") = 0
            rsSal("SH_SALCHG1") = 0
            rsSal("SH_SREAS2") = oSReas2
            rsSal("SH_SALPC2") = 0
            rsSal("SH_SALCHG2") = 0
            rsSal("SH_SREAS3") = oSReas3
            rsSal("SH_SALPC3") = 0
            rsSal("SH_SALCHG3") = 0
            If oCompa <> "" And Not IsNull(oCompa) Then rsSal("SH_COMPA") = oCompa
            If IsDate(ONDate) Then rsSal("SH_NEXTDAT") = ONDate
            
            'Ticket #29480 - City of Niagara Falls do not want to carry forward the Comments
            'rsSal("SH_COMMENT") = oSComment
            
            rsSal("SH_COMMENT2") = oSComment2
            rsSal("SH_PAYROLL_ID") = oPayrollID
            If OTOTAL <> "" And Not IsNull(OTOTAL) Then rsSal("SH_TOTAL") = OTOTAL
            rsSal("SH_GRID") = oGrid
            
            rsSal("SH_TRANSDATE") = Now
            rsSal("SH_LDATE") = Date
            rsSal("SH_LTIME") = Time$
            rsSal("SH_LUSER") = glbUserID
            
            'If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
                rsSal("SH_PREMIUM") = OPremium
                rsSal("SH_TOTAL") = OTOTAL
                rsSal("SH_VGROUP") = oVGroup
                rsSal("SH_VSTEP") = oVStep
            'End If
            
            rsSal.Update
                        
            'For Audit Log
            NSalary = oSal
            
            'Update the Audit Log
            If Not AUDITSALY(empNo&) Then MsgBox "ERROR - AUDIT FILE"
        End If
        rsSal.Close
        Set rsSal = Nothing
    
        'Update Log for the Position change
        Call UpdateLogPosSal(empNo&, OSalCD, oSal, "Position", clpNJob.Text, oJob, dlpNSDate.Text, OSDATE)
    
        'For the Employee List report
        If Len(strEMPLIST) > 0 Then
            strEMPLIST = strEMPLIST & "," & empNo&
        Else
            strEMPLIST = empNo&
        End If


lblNextRec:
    snapPosSal.MoveNext
Loop

    If gsEMAIL_ONPOSITION Then
        If Len(MailBodyP) > 0 Then
            'If glbCompSerial = "S/N - 2443W" Then 'Walters Inc 'Ticket #25952 Franks 11/03/2014
            '    MailBodyP = getWaltersIncEmailBody
            'Else
                If prec% = 1 Then
                    xStr = "This will serve to confirm that the following employee's position title has been changed." & vbCrLf & vbCrLf
                Else
                    xStr = "This will serve to confirm that the following employees position title have been changed." & vbCrLf & vbCrLf
                End If
                xStr = xStr & "New position: " & getPosDesc(clpNJob.Text) & vbCrLf
                xStr = xStr & "Reason: " & GetTABLDesc("SDRC", clpNReason) & vbCrLf
                xStr = xStr & "Effective Date: " & dlpNSDate.Text & vbCrLf & vbCrLf
                MailBodyP = xStr & MailBodyP
                Screen.MousePointer = DEFAULT
                Call imgEmailP_Click
            'End If
        End If
    End If

modUpdateSelection_On_LAYOFF = True

MDIMain.panHelp(0).FloodType = 0

snapPosSal.Close
Set snapPosSal = Nothing

Screen.MousePointer = DEFAULT

Exit Function

modUpdateSelection_On_LAYOFF_Err:

MDIMain.panHelp(0).FloodType = 0
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Add LayOFF Postion & Salary", "HR_JOB_HISTORY / HR_SALARY_HISTORY", "Edit/Add")

Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If

End Function

Private Function modUpdateSelection_ReturnFrom_LAYOFF()
    Dim SQLQ As String
    Dim rsJOB As New ADODB.Recordset
    Dim rsSal As New ADODB.Recordset
    Dim rsEmp As New ADODB.Recordset
    Dim rsJobMst As New ADODB.Recordset
'    Dim dynSH_Job1 As New ADODB.Recordset
    Dim lngLastCurrentID&
    Dim prec%, pct%
    Dim dblPct, dblAmnt As Variant
    Dim dblNewSalary As Double
    Dim JobSalCD
    Dim xSHID
    Dim xStr As String
    Dim xStep
    
    'Update Position
    ' - Make current LayOFF Position Un-Current
    ' - Add a new return to Position with default values
    ' - Update HREMP as well - Department, Divisio, Location

    'New Position Record
    On Error GoTo modUpdateSelection_ReturnFrom_LAYOFF_Err
    
    modUpdateSelection_ReturnFrom_LAYOFF = False
    
    'Call DelNoUpsal
    
    Screen.MousePointer = HOURGLASS
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 5
    
    'Clear the log table
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute "DELETE FROM HREMPHIS_WRK WHERE EE_WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001W.CommitTrans
    Call Pause(1)
    
    If rsJobMst.State <> 0 Then rsJobMst.Close
    rsJobMst.Open "SELECT * FROM HRJOB WHERE JB_CODE = '" & clpNJob.Text & "'", gdbAdoIhr001, adOpenStatic
    
    SkipRec = 0
    MailBodyP = ""
    MailBody = ""
    
    Do While Not snapPosSal.EOF
        prec% = prec% + 1
        pct% = Int(100 * (prec% / (lngRecs&)))
        MDIMain.panHelp(0).FloodPercent = pct%
        
        empNo& = snapPosSal("JH_EMPNBR")
                
        'Clear fields
        oJob = "": OReason = "": OSDATE = "": ODHRS = "": oWHRS = "": oPHRS = "": oRepAut = "": oSHIFT = "": oFTENum = "": oFTEHrs = ""
        oOrg = "": oPTFT = "": oComment = "": oComment2 = "": oRepAut2 = "": oRepAut3 = "": oRepAut4 = "": oDiv = "": oDeptNo = "": oEmp = ""
        oGLNo = "": oSect = "": oRegion = "": oPosCtrl = "": oPayrollID = "": oGrid = "": oPayCateg = "": oBillRate = ""
        oPosStatus = ""
                
        'Retrieve existing Current Position Data
        oJob = snapPosSal("JH_JOB")
        OReason = snapPosSal("JH_JREASON")
        OSDATE = snapPosSal("JH_SDATE")
        If Not IsNull(snapPosSal("JH_DHRS")) Then ODHRS = snapPosSal("JH_DHRS")
        If Not IsNull(snapPosSal("JH_WHRS")) Then oWHRS = snapPosSal("JH_WHRS")
        If Not IsNull(snapPosSal("JH_PHRS")) Then oPHRS = snapPosSal("JH_PHRS")
        If Not IsNull(snapPosSal("JH_REPTAU")) Then oRepAut = snapPosSal("JH_REPTAU")
        If Not IsNull(snapPosSal("JH_SHIFT")) Then oSHIFT = snapPosSal("JH_SHIFT")
        If Not IsNull(snapPosSal("JH_FTENUM")) Then oFTENum = snapPosSal("JH_FTENUM")
        If Not IsNull(snapPosSal("JH_FTEHRS")) Then oFTEHrs = snapPosSal("JH_FTEHRS")
        oOrg = snapPosSal("JH_ORG")
        oPTFT = snapPosSal("JH_PT")
        oComment = snapPosSal("JH_COMMENT")
        oComment2 = snapPosSal("JH_COMMENT2")
        If Not IsNull(snapPosSal("JH_REPTAU2")) Then oRepAut2 = snapPosSal("JH_REPTAU2")
        If Not IsNull(snapPosSal("JH_REPTAU3")) Then oRepAut3 = snapPosSal("JH_REPTAU3")
        If Not IsNull(snapPosSal("JH_REPTAU4")) Then oRepAut4 = snapPosSal("JH_REPTAU4")
        'oLeadHand = rsJOB("JH_LEADHAND")
        'oLabourCD = rsJOB("JH_LABOURCD")
        'oLabourDate = rsJOB("JH_LABOUREDATE")
        'oUsrLabel = rsJOB("JH_USRLABEL")
        'oUsrCheck = rsJOB("JH_USRCHECK")
        'oUsrDate = rsJOB("JH_USREDATE")
        'oUsrLabel2 = rsJOB("JH_USRLABEL2")
        'oUsrCheck2 = rsJOB("JH_USRCHECK2")
        'oUsrDate2 = rsJOB("JH_USREDATE2")
        'oUsrLabel3 = rsJOB("JH_USRLABEL3")
        'oUsrCheck3 = rsJOB("JH_USRCHECK3")
        'oUsrDate3 = rsJOB("JH_USREDATE3")
        oDiv = snapPosSal("JH_DIV")
        oDeptNo = snapPosSal("JH_DEPTNO")
        oEmp = snapPosSal("JH_EMP")
        oGLNo = snapPosSal("JH_GLNO")
        oSect = snapPosSal("JH_SECTION")
        oRegion = snapPosSal("JH_REGION")
        oPosCtrl = snapPosSal("JH_POSITION_CONTROL")
        oPayrollID = snapPosSal("JH_PAYROLL_ID")
        oGrid = snapPosSal("JH_GRID")
        oPayCateg = snapPosSal("JH_PAYROLL_CATEGORY")
        oBillRate = snapPosSal("JH_BILLINGRATE")
        'oENDDATE = snapPosSal("JH_ENDDATE")
        'oEndReason = snapPosSal("JH_ENDREAS")
        oPosStatus = snapPosSal("JH_ESTATUS")
        
        If IsNull(snapPosSal("JH_GRID")) Then
            OLambtonJob = oJob
        Else
            OLambtonJob = Left(snapPosSal("JH_GRID"), 1) & oJob & Mid(snapPosSal("JH_GRID"), 2)
        End If

        lngLastCurrentID& = snapPosSal("JH_ID")

        'Remove the Current check from the existing current position record
        rsJOB.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID=" & lngLastCurrentID&, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rsJOB("JH_CURRENT") = False
        rsJOB("JH_ENDDATE") = DateAdd("d", -1, dlpNSDate.Text)
        rsJOB("JH_ENDREAS") = clpNReason.Text
        rsJOB("JH_LDATE") = Format(Now, "Short Date")
        rsJOB("JH_LTIME") = Time$
        rsJOB("JH_LUSER") = glbUserID
        rsJOB.Update

        'Add a new Current Position records
        rsJOB.AddNew
        rsJOB("JH_COMPNO") = "001"
        rsJOB("JH_EMPNBR") = empNo&
        rsJOB("JH_JOB") = clpNJob.Text
        rsJOB("JH_SDATE") = dlpNSDate.Text
        If ODHRS <> "" And Not IsNull(ODHRS) Then rsJOB("JH_DHRS") = ODHRS
        If oWHRS <> "" And Not IsNull(oWHRS) Then rsJOB("JH_WHRS") = oWHRS
        If oPHRS <> "" And Not IsNull(oPHRS) Then rsJOB("JH_PHRS") = oPHRS
        rsJOB("JH_JREASON") = clpNReason.Text
        rsJOB("JH_CURRENT") = True
        If oRepAut <> "" And Not IsNull(oRepAut) Then rsJOB("JH_REPTAU") = oRepAut
        rsJOB("JH_SHIFT") = oSHIFT
        If oFTENum <> "" And Not IsNull(oFTENum) Then rsJOB("JH_FTENUM") = oFTENum
        If oFTEHrs <> "" And Not IsNull(oFTEHrs) Then rsJOB("JH_FTEHRS") = oFTEHrs
        rsJOB("JH_ORG") = oOrg
        rsJOB("JH_PT") = oPTFT
        'If IsDate(dlpExpRtnDate.Text) Then
        '    If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls only
        '        rsJOB("JH_COMMENT") = "Expected Return Date:" & dlpExpRtnDate.Text
        '    Else
        '        rsJOB("JH_COMMENT") = Left(oComment & " ExpRtnDt:" & dlpExpRtnDate.Text, 50)
        '    End If
        'End If
        'Ticket #29382 - They want Notes 1 and Notes 2 to be user input fields
        rsJOB("JH_COMMENT") = txtComment.Text
        
        'Ticket #29382 - They want Notes 1 and Notes 2 to be user input fields
        'Ticket #29382 - They do not want to carry forward the Notes 2
        'rsJOB("JH_COMMENT2") = oComment2
        rsJOB("JH_COMMENT2") = txtComments2.Text
        
        oComment = ""
        If oRepAut2 <> "" And Not IsNull(oRepAut2) Then rsJOB("JH_REPTAU2") = oRepAut2
        If oRepAut3 <> "" And Not IsNull(oRepAut3) Then rsJOB("JH_REPTAU3") = oRepAut3
        If oRepAut4 <> "" And Not IsNull(oRepAut4) Then rsJOB("JH_REPTAU4") = oRepAut4
        rsJOB("JH_DIV") = clpNDiv.Text
        rsJOB("JH_DEPTNO") = clpNDept.Text
        rsJOB("JH_EMP") = oEmp
        rsJOB("JH_GLNO") = oGLNo
        rsJOB("JH_SECTION") = oSect
        rsJOB("JH_REGION") = oRegion
        rsJOB("JH_POSITION_CONTROL") = oPosCtrl
        rsJOB("JH_PAYROLL_ID") = oPayrollID
        rsJOB("JH_GRID") = oGrid
        rsJOB("JH_PAYROLL_CATEGORY") = oPayCateg
        rsJOB("JH_BILLINGRATE") = oBillRate
        'rsJOB("JH_ENDDATE") = rEndDate
        'rsJOB("JH_ENDREAS") = rEndReason
        rsJOB("JH_ESTATUS") = oPosStatus
        
        rsJOB("JH_LDATE") = Format(Now, "Short Date")
        rsJOB("JH_LTIME") = Time$
        rsJOB("JH_LUSER") = glbUserID
        rsJOB.Update

        If gsEMAIL_ONPOSITION Then
            MailBodyP = MailBodyP & GetEmpName(rsJOB("JH_EMPNBR")) & vbCrLf
        End If

        nJobID = rsJOB("JH_ID")

        'Update Log for the Position change
        Call UpdateLogPosSal(empNo&, "", "", "Position", clpNJob.Text, oJob, dlpNSDate.Text, OSDATE)

        'Document Attachment
        '7.9 Enhancement
        If gsAttachment_DB Then
            If glbDocNewRecord Then 'New Record only
                If Len(glbDocImpFile) > 0 Then
                    'glbJob = xID
                    glbJob = clpNJob.Text
                    glbSDate = dlpNSDate.Text
                    Call AttachmentAdd(empNo&, glbDocImpFile, glbDocType, glbDocDesc)
                End If
            End If
            'glbDocImpFile = ""
        End If

        rsJOB.Close
        Set rsJOB = Nothing

        'Update HREMP with hard coded Dept, Div and Location
        If Not glbMulti Then
            If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls only
                'Employee History
                'If Not EmpHisCalc(1, empNo&, clpNDept.Text, "", "", "", "", "", "", Date) Then MsgBox "EMPHIS Error "
                'If Not EmpHisCalc(1, empNo&, "", clpNDiv, "", "", "", "", "", Date) Then MsgBox "EMPHIS Error "
            
                rsEmp.Open "SELECT * FROM HREMP WHERE ED_EMPNBR = " & empNo&, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsEmp.EOF Then
                    If oDeptNo = "" Or IsNull(oDeptNo) Then oDeptNo = rsEmp("ED_DEPTNO")
                    If Len(Trim(clpNDept.Text)) > 0 Then
                        'If Not EmpHisCalc(1, empNo&, clpNDept.Text, "", "", "", "", "", "", Date,  "DEPT", clpNDept.Text) Then MsgBox "EMPHIS Error "
                        If Not EmpHisCalc(1, empNo&, clpNDept.Text, "", "", "", "", "", "", Date, , , , , , , oDeptNo) Then MsgBox "EMPHIS Error "  'Employee History Update
                        rsEmp("ED_DEPTNO") = clpNDept.Text
                    Else
                        'Get the Previous value prior to the Lay Off from Employee History using their standard Layoff/Current value of the employee
                        nDeptNo = getEmpHistoryValue("DEPT", "4900", empNo&)
                        If Not EmpHisCalc(1, empNo&, nDeptNo, "", "", "", "", "", "", Date, , , , , , , oDeptNo) Then MsgBox "EMPHIS Error "
                        rsEmp("ED_DEPTNO") = nDeptNo
                    End If
                    
                    If Len(Trim(clpNDiv.Text)) > 0 Then
                        If oDiv = "" Or IsNull(oDiv) Then oDiv = rsEmp("ED_DIV")
                        'If Not EmpHisCalc(1, empNo&, "", clpNDiv, "", "", "", "", "", Date, "DIV", clpNDiv.Text) Then MsgBox "EMPHIS Error "
                        If Not EmpHisCalc(1, empNo&, "", clpNDiv, "", "", "", "", "", Date, , , , , , , oDiv) Then MsgBox "EMPHIS Error "       'Employee History Update
                        rsEmp("ED_DIV") = clpNDiv.Text
                    Else
                        'Get the Previous value prior to the Lay Off from Employee History using their standard Layoff/Current value of the employee
                        nDiv = getEmpHistoryValue("DIV", "LAYO", empNo&)
                        If Not EmpHisCalc(1, empNo&, "", nDiv, "", "", "", "", "", Date, , , , , , , oDiv) Then MsgBox "EMPHIS Error "
                        rsEmp("ED_DIV") = nDiv
                    End If
                    
                    oLoc = rsEmp("ED_LOC")
                    If Len(Trim(clpNLoc.Text)) > 0 Then
                        If Not EmpHisCalc(2, empNo&, "", "", "", "", "", "", "", Date, "LOC", clpNLoc.Text, , , , , oLoc) Then MsgBox "EMPHIS Error "        'Employee History Update
                        rsEmp("ED_LOC") = clpNLoc.Text
                    Else
                        'Get the Previous value prior to the Lay Off from Employee History using their standard Layoff/Current value of the employee
                        nLoc = getEmpHistoryValue("LOC", "NW", empNo&)
                        If Not EmpHisCalc(2, empNo&, "", "", "", "", "", "", "", Date, "LOC", nLoc, , , , , oLoc) Then MsgBox "EMPHIS Error "
                        rsEmp("ED_LOC") = nLoc
                    End If
                    
                    oAdminBy = rsEmp("ED_ADMINBY")      'Ticket #29480 - Additional fields
                    If Len(Trim(clpNAdminBy.Text)) > 0 Then
                        If Not EmpHisCalc(2, empNo&, "", "", "", "", "", "", "", Date, "ADMINBY", clpNAdminBy.Text, , , , , oAdminBy) Then MsgBox "EMPHIS Error "      'Ticket #29480 - Additional fields
                        rsEmp("ED_ADMINBY") = clpNAdminBy.Text
                    Else
                        'Get the Previous value prior to the Lay Off from Employee History using their standard Layoff/Current value of the employee
                        nAdminBy = getEmpHistoryValue("ADMINBY", "4900", empNo&)
                        If Not EmpHisCalc(2, empNo&, "", "", "", "", "", "", "", Date, "ADMINBY", nAdminBy, , , , , oAdminBy) Then MsgBox "EMPHIS Error "     'Ticket #29480 - Additional fields
                        rsEmp("ED_ADMINBY") = nAdminBy
                    End If
                    
                    rsEmp("ED_LDATE") = Format(Now, "Short Date")
                    rsEmp("ED_LTIME") = Time$
                    rsEmp("ED_LUSER") = glbUserID
                    rsEmp.Update
                End If
                rsEmp.Close
                Set rsEmp = Nothing
            Else
                rsEmp.Open "SELECT * FROM HREMP WHERE ED_EMPNBR = " & empNo&, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsEmp.EOF Then
                    If oDeptNo = "" Or IsNull(oDeptNo) Then oDeptNo = rsEmp("ED_DEPTNO")
                    If Len(Trim(clpNDept.Text)) > 0 Then
                        If Not EmpHisCalc(1, empNo&, clpNDept.Text, "", "", "", "", "", "", Date, , , , , , , oDeptNo) Then MsgBox "EMPHIS Error "
                        rsEmp("ED_DEPTNO") = clpNDept.Text
                    Else
                        'Get the Previous value prior to the Lay Off from Employee History using the Current value of the employee
                        nDeptNo = getEmpHistoryValue("DEPT", oDeptNo, empNo&)
                        If Not EmpHisCalc(1, empNo&, nDeptNo, "", "", "", "", "", "", Date, , , , , , , oDeptNo) Then MsgBox "EMPHIS Error "
                        rsEmp("ED_DEPTNO") = nDeptNo
                    End If
                    
                    If oDiv = "" Or IsNull(oDiv) Then oDiv = rsEmp("ED_DIV")
                    If Len(Trim(clpNDiv.Text)) > 0 Then
                        If Not EmpHisCalc(1, empNo&, "", clpNDiv, "", "", "", "", "", Date, , , , , , , oDiv) Then MsgBox "EMPHIS Error "
                        rsEmp("ED_DIV") = clpNDiv.Text
                    Else
                        'Get the Previous value prior to the Lay Off from Employee History using the Current value of the employee
                        nDiv = getEmpHistoryValue("DIV", oDiv, empNo&)
                        If Not EmpHisCalc(1, empNo&, "", nDiv, "", "", "", "", "", Date, , , , , , , oDiv) Then MsgBox "EMPHIS Error "
                        rsEmp("ED_DIV") = nDiv
                    End If
                
                    oLoc = rsEmp("ED_LOC")
                    If Len(Trim(clpNLoc.Text)) > 0 Then
                        If Not EmpHisCalc(2, empNo&, "", "", "", "", "", "", "", Date, "LOC", clpNLoc.Text, , , , , oLoc) Then MsgBox "EMPHIS Error "
                        rsEmp("ED_LOC") = clpNLoc.Text
                    Else
                        'Get the Previous value prior to the Lay Off from Employee History using the Current value of the employee
                        nLoc = getEmpHistoryValue("LOC", oLoc, empNo&)
                        If Not EmpHisCalc(2, empNo&, "", "", "", "", "", "", "", Date, "LOC", nLoc, , , , , oLoc) Then MsgBox "EMPHIS Error "
                        rsEmp("ED_LOC") = nLoc
                    End If
                    
                    oAdminBy = rsEmp("ED_ADMINBY")      'Ticket #29480 - Additional fields
                    If Len(Trim(clpNAdminBy.Text)) > 0 Then
                        If Not EmpHisCalc(2, empNo&, "", "", "", "", "", "", "", Date, "ADMINBY", clpNAdminBy.Text, , , , , oAdminBy) Then MsgBox "EMPHIS Error "     'Ticket #29480 - Additional fields
                        rsEmp("ED_ADMINBY") = clpNAdminBy.Text
                    Else
                        'Get the Previous value prior to the Lay Off from Employee History using the Current value of the employee
                        nAdminBy = getEmpHistoryValue("ADMINBY", oAdminBy, empNo&)
                        If Not EmpHisCalc(2, empNo&, "", "", "", "", "", "", "", Date, "ADMINBY", nAdminBy, , , , , oAdminBy) Then MsgBox "EMPHIS Error "     'Ticket #29480 - Additional fields
                        rsEmp("ED_ADMINBY") = nAdminBy
                    End If
                    
                    If Len(Trim(clpNDept.Text)) > 0 Or Len(Trim(clpNDiv.Text)) > 0 Or Len(Trim(clpNLoc.Text)) > 0 Or Len(Trim(clpNAdminBy.Text)) > 0 Then
                        rsEmp("ED_LDATE") = Format(Now, "Short Date")
                        rsEmp("ED_LTIME") = Time$
                        rsEmp("ED_LUSER") = glbUserID
                        rsEmp.Update
                    End If
                End If
                rsEmp.Close
                Set rsEmp = Nothing
            End If
        End If
        
        'Update the Audit Log
        If Not AUDITPOS(empNo&) Then MsgBox "ERROR - AUDIT FILE"
        

        'Update Salary
        ' - Make current LayOFF Salary Un-Current
        ' - Add a new return to Position Salary with new Step# and Salary based on the Position Master.

        SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & empNo& & " AND SH_CURRENT<>0 AND SH_JOB ='" & oJob & "'"
        If glbMulti Then
            SQLQ = SQLQ & " AND SH_SDATE = " & Date_SQL(OSDATE)
        End If
        rsSal.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsSal.EOF Then
            'Clear fields
            OEDate = "": oSal = "": OSalCD = "": oGrade = "": oPayP = "": oSReas1 = "": oSReas2 = "": oSReas3 = "": oCompa = "": ONDate = ""
            oSComment = "": oSComment2 = "": OTOTAL = "": oGrid = "": oPayrollID = "": oWHRS = "": OPremium = "": oVGroup = ""
            oVStep = ""
            
            'Retrieve Data
            OEDate = rsSal("SH_EDATE")
            oSal = rsSal("SH_SALARY")
            OSalCD = rsSal("SH_SALCD")
            oGrade = rsSal("SH_GRADE")
            oPayP = rsSal("SH_PAYP")
            If Not IsNull(rsSal("SH_SREAS1")) Then oSReas1 = rsSal("SH_SREAS1")
            If Not IsNull(rsSal("SH_SREAS2")) Then oSReas2 = rsSal("SH_SREAS2")
            If Not IsNull(rsSal("SH_SREAS3")) Then oSReas3 = rsSal("SH_SREAS3")
            If Not IsNull(rsSal("SH_COMPA")) Then oCompa = rsSal("SH_COMPA")
            If Not IsNull(rsSal("SH_NEXTDAT")) Then ONDate = rsSal("SH_NEXTDAT")
            oSComment = rsSal("SH_COMMENT")
            oSComment2 = rsSal("SH_COMMENT2")
            If Not IsNull(rsSal("SH_TOTAL")) Then OTOTAL = rsSal("SH_TOTAL")
            oGrid = rsSal("SH_GRID")
            oPayrollID = rsSal("SH_PAYROLL_ID")
            If Not IsNull(rsSal("SH_WHRS")) Then oWHRS = rsSal("SH_WHRS")
            OPremium = rsSal("SH_PREMIUM")
            oVGroup = rsSal("SH_VGROUP")
            oVStep = rsSal("SH_VSTEP")
            'oSalPC1 = rsSal("SH_SALPC1")
            'oSalChg1 = rsSal("SH_SALCHG1")
            'oSalPC2 = rsSal("SH_SALPC2")
            'oSalChg2 = rsSal("SH_SALCHG2")
            'oSalPC3 = rsSal("SH_SALPC3")
            'oSalChg3 = rsSal("SH_SALCHG3")
            
            'Remove the Current check from the existing current salary record
            rsSal("SH_CURRENT") = False
            rsSal("SH_TRANSDATE") = Now
            rsSal("SH_LDATE") = Format(Now, "Short Date")
            rsSal("SH_LTIME") = Time$
            rsSal("SH_LUSER") = glbUserID
            rsSal.Update
        
            'Get employee's Hours/Day
            fglbDhrs = snapPosSal("JH_DHRS") 'GetJHData(EmpNo&, "JH_DHRS", 0)
            
            'Add a new Current Salary records
            rsSal.AddNew
            rsSal("SH_COMPNO") = "001"
            rsSal("SH_EMPNBR") = empNo&
            rsSal("SH_CURRENT") = True
            rsSal("SH_JOB") = clpNJob.Text
            rsSal("SH_JOB_ID") = nJobID
            rsSal("SH_SDATE") = dlpNSDate.Text
            rsSal("SH_EDATE") = dlpNSDate.Text
            rsSal("SH_TRANSDATE") = Format(Now, "SHORT DATE")
            rsSal("SH_SALCD") = OSalCD
            If oWHRS <> "" And Not IsNull(oWHRS) Then rsSal("SH_WHRS") = oWHRS
                              
            dblNewSalary = oSal
            JobSalCD = rsJobMst("JB_SALCD")
            
            'Ticket #29382 - Since Step # is optional, if selected then use that Step # to update employee's Salary with otherwise use the
            'Step # of the previous Salary record. In both cases, use the Salary from the Position Master for hte associated Position based on
            'the Step #.
            If Len(Trim(comStep)) <> 0 Then
                xStep = comStep
            Else
                xStep = oGrade
            End If
            
            If OSalCD <> JobSalCD Then
                If xStep <> "00" Then
                    If OSalCD = "A" And JobSalCD = "H" Then
                        dblAmnt = ((rsJobMst("JB_S" & Format(xStep, "##")) * rsSal("SH_WHRS")) * 52) - dblNewSalary
                    ElseIf OSalCD = "H" And JobSalCD = "A" Then
                        If rsSal("SH_WHRS") = 0 Then
                            dblAmnt = 0
                        Else
                            dblAmnt = ((rsJobMst("JB_S" & Format(xStep, "##")) / rsSal("SH_WHRS")) / 52) - dblNewSalary
                        End If
                    ElseIf OSalCD = "D" And JobSalCD = "H" Then
                        dblAmnt = ((rsJobMst("JB_S" & Format(xStep, "##")) * fglbDhrs)) - dblNewSalary
                    ElseIf OSalCD = "D" And JobSalCD = "A" Then
                        If rsSal("SH_WHRS") = 0 Then
                            dblAmnt = 0
                        Else
                            dblAmnt = (((rsJobMst("JB_S" & Format(xStep, "##")) / rsSal("SH_WHRS")) / 52) * fglbDhrs) - dblNewSalary
                        End If
                    End If
                End If
            ElseIf OSalCD = JobSalCD Then
                If xStep <> "00" Then
                    dblAmnt = rsJobMst("JB_S" & Format(xStep, "##")) - dblNewSalary
                End If
            End If
            
            If oSal = 0 Then
                dblPct = 0
            Else
                dblPct = dblAmnt / oSal
            End If
            dblNewSalary = dblNewSalary + dblAmnt
            
            If glbCompSerial = "S/N - 2378W" And OSalCD <> JobSalCD And OSalCD <> "M" Then   'Town of Aurora
                dblAmnt = rsJobMst("JB_S" & Format(xStep, "##") & "A") - dblNewSalary
                If oSal = 0 Then
                    dblPct = 0
                Else
                    dblPct = dblAmnt / oSal
                End If
            End If
                    
            rsSal("SH_SREAS1") = clpNReason.Text
            rsSal("SH_SALPC1") = dblPct
            rsSal("SH_SALCHG1") = dblAmnt
            
            dblNewSalary = Round2DEC(dblNewSalary, rsSal("SH_EMPNBR")) 'added by raubrey 8/18/97
            rsSal("SH_SALARY") = dblNewSalary
            
            'Jaddy changed for WFC Kipling Oct 31, 02
            If glbWFC And (rsJobMst("JB_ORG") = "NONE" Or rsJobMst("JB_ORG") = "EXEC") Then
                rsSal("SH_COMPA_USER") = snapPosSal("SH_COMPA_USER")
                rsSal("SH_COMPA_DOLLAR") = snapPosSal("SH_COMPA_DOLLAR")
                rsSal("SH_BAND") = snapPosSal("SH_BAND")
                rsSal("SH_MARKETLINE") = snapPosSal("SH_MARKETLINE")
                Call Set_WFC_COMPA(dblNewSalary)
            Else
                Call modSetCOMPA_GRADE(Round2DEC(dblNewSalary, rsSal("SH_EMPNBR")), rsSal) ' sets fglbCOMPA#, and fglbGRADE
            End If
            rsSal("SH_COMPA") = Round(fglbCOMPA#, 2)
            rsSal("SH_GRADE") = Format(fglbGRADE$, "00")
            
            rsSal("SH_PAYP") = oPayP
            rsSal("SH_SREAS2") = oSReas2
            rsSal("SH_SALPC2") = 0
            rsSal("SH_SALCHG2") = 0
            rsSal("SH_SREAS3") = oSReas3
            rsSal("SH_SALPC3") = 0
            rsSal("SH_SALCHG3") = 0
                        
            If IsDate(ONDate) Then rsSal("SH_NEXTDAT") = ONDate
            
            'Ticket #29480 - City of Niagara Falls do not want to carry forward the Comments
            'rsSal("SH_COMMENT") = oSComment
            
            rsSal("SH_COMMENT2") = oSComment2
            rsSal("SH_PAYROLL_ID") = oPayrollID
            If OTOTAL <> "" And Not IsNull(OTOTAL) Then rsSal("SH_TOTAL") = OTOTAL
            rsSal("SH_GRID") = oGrid
            
            rsSal("SH_TRANSDATE") = Now
            rsSal("SH_LDATE") = Date
            rsSal("SH_LTIME") = Time$
            rsSal("SH_LUSER") = glbUserID
            
            If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
                rsSal("SH_PREMIUM") = OPremium
                rsSal("SH_TOTAL") = rsSal("SH_SALARY") + OPremium
                rsSal("SH_VGROUP") = oVGroup
                rsSal("SH_VSTEP") = oVStep
            End If
            
            rsSal.Update
                
            'Update Log for the Salary change
            Call UpdateLogPosSal(empNo&, OSalCD, "", "Salary", rsSal("SH_SALARY"), CStr(oSal), dlpNSDate.Text, OEDate)
        
            'Add by Frank Jan 10,2002 As Jerry request
            'If IsDate(ONDate) Then
            '    If CVDate(ONDate) > CVDate(dlpNSDate) Then
            '        UpdateFollowup EmpNo&, CVDate(ONDate), CVDate(ONDate), "SREV"
            '    End If
            'End If
            
            If gsEMAIL_ONSALARY Then
                MailBody = MailBody & GetEmpName(rsSal("SH_EMPNBR")) & vbCrLf
            End If

            xSHID = rsSal("SH_ID") 'George added on MAr 10,2006 #9965

            'Vadim Transfer
            If glbVadim Then Call Transfer_Salary(rsSal)

            'Attendance rates update
            ''Ticket #28595 - Update employee's Attendance records as well if selected
            'If chkUpdAttendance Then Call Update_Attendance_SalaryInfo(rsSal)

            'Salary Dependent Benefit Update
            Call updBenefitForSalDEPN(empNo&)   'jaddy 9/10/99

            'This table only stores the Rate Level
            'City of Niagara Falls - Ticket #15542
            If glbVadim And glbCompSerial = "S/N - 2276W" Then
                'Add the salary record in Vadim's HR_EMP_HIST table storing the history of Rate changes
                Call Update_VadimDB_HR_EMP_HISTORY(rsSal("SH_PAYROLL_ID"), CVDate(dlpNSDate), "", Val(fglbGRADE$), rsSal("SH_JOB"), "A")
            End If
        
            'Other Integration updates
            Call Employee_Master_Integration(empNo&)
            
            If glbGP Then
                Call Salary_Integration(empNo&, , False, True, xSHID) 'George added on MAr 10,2006 #9965
            Else
                Call Salary_Integration(empNo&) 'Ticket #15646
            End If

            NSalary = dblNewSalary
            NEDate = CVDate(dlpNSDate.Text)
            NNDate = ""

            'Update the Audit Log
            If Not AUDITSALY(empNo&) Then MsgBox "ERROR - AUDIT FILE"

        End If
        rsSal.Close
        Set rsSal = Nothing

        'For the Employee List report
        If Len(strEMPLIST) > 0 Then
            strEMPLIST = strEMPLIST & "," & empNo&
        Else
            strEMPLIST = empNo&
        End If

lblNextRec:
    snapPosSal.MoveNext
Loop

    
    If gsEMAIL_ONPOSITION Then
        If Len(MailBodyP) > 0 Then
            'If glbCompSerial = "S/N - 2443W" Then 'Walters Inc 'Ticket #25952 Franks 11/03/2014
            '    MailBodyP = getWaltersIncEmailBody
            'Else
                If prec% = 1 Then
                    xStr = "This will serve to confirm that the following employee's position title has been changed." & vbCrLf & vbCrLf
                Else
                    xStr = "This will serve to confirm that the following employees position title have been changed." & vbCrLf & vbCrLf
                End If
                xStr = xStr & "New position: " & getPosDesc(clpNJob.Text) & vbCrLf
                xStr = xStr & "Reason: " & GetTABLDesc("SDRC", clpNReason) & vbCrLf
                xStr = xStr & "Effective Date: " & dlpNSDate.Text & vbCrLf & vbCrLf
                MailBodyP = xStr & MailBodyP
                Screen.MousePointer = DEFAULT
                Call imgEmailP_Click
            'End If
        End If
    End If


    If gsEMAIL_ONSALARY Then
        If Len(MailBody) > 0 Then
            If prec% = 1 Then
                xStr = "The following employee's salary has "
            Else
                xStr = "The following employee's salaries have "
            End If
            xStr = xStr & " been changed." & vbCrLf 'to " & NSalary & "%." & vbCrLf  '& vbCrLf
            xStr = xStr & "Position: " & GetJobData(clpNJob, "JB_DESCR") & vbCrLf
            xStr = xStr & "Reason: " & GetTABLDesc("SDRC", clpNReason) & vbCrLf
            xStr = xStr & "Effective Date: " & dlpNSDate & vbCrLf & vbCrLf
            MailBody = xStr & MailBody
            Screen.MousePointer = DEFAULT
            Call imgEmail_Click
        End If
    End If

    If prec% > 0 Then
        If prec% = 1 Then
            MsgBox prec% & " employee salary record was updated"
        Else
            MsgBox prec% & " employees salary records were updated"
        End If
    End If


rsJobMst.Close
Set rsJobMst = Nothing

modUpdateSelection_ReturnFrom_LAYOFF = True

MDIMain.panHelp(0).FloodType = 0

snapPosSal.Close
Set snapPosSal = Nothing

Screen.MousePointer = DEFAULT

Exit Function

modUpdateSelection_ReturnFrom_LAYOFF_Err:

MDIMain.panHelp(0).FloodType = 0
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Add Return From LayOff Postion & Salary", "HR_JOB_HISTORY / HR_SALARY_HISTORY", "Edit/Add")

Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub modSetCOMPA_GRADE_1(ByRef dblNewSalary)

Dim X%, cX$, xSalGrade, SQLQ
Dim dblSsalary#, dblHoursPerWeek#, ssalary@
Dim Jb_No#
Dim snapJob As New ADODB.Recordset

Dim xStep As Integer
Dim xGrid2 As Boolean

'SET COMPA RATIO
'================
If glbMultiGrid Then
    SQLQ = "SELECT * FROM HRJOB_GRADE WHERE JB_CODE='" & clpNJob.Text & "' AND JB_GRID='" & snapPosSal("SH_GRID") & "'"
Else
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE='" & clpNJob.Text & "'"
End If
snapJob.Open SQLQ, gdbAdoIhr001, adOpenStatic

ssalary@ = dblNewSalary
dblHoursPerWeek# = snapPosSal("JH_WHRS")    'xWHRS
fglbDhrs = snapPosSal("JH_DHRS") 'GetJHData(snapPosSal("SH_EMPNBR"), "JH_DHRS", 0)

If Len(oStep) > 0 Then
    xStep = CInt(oStep)
Else
    xStep = 0
End If

If snapJob("JB_SALCD") = "H" Then
    If snapPosSal("SH_SALCD") = "H" Then
        If xStep <> 0 Then
            dblSsalary# = snapJob("JB_S" & Format(xStep, "##"))
        Else
            dblSsalary# = 0
        End If
    ElseIf snapPosSal("SH_SALCD") = "M" Then
        If dblHoursPerWeek# = 0 Then
            dblSsalary# = 0
        Else
            dblSsalary# = (dblNewSalary * 12) / (dblHoursPerWeek# * 52)
        End If
    ElseIf snapPosSal("SH_SALCD") = "A" Then
        If dblHoursPerWeek# = 0 Then
            dblSsalary# = 0
        Else
            dblSsalary# = dblNewSalary / (dblHoursPerWeek# * 52)
        End If
    ElseIf snapPosSal("SH_SALCD") = "D" Then
        If GetLeapYear(Year(Date)) Then
            dblSsalary# = (dblNewSalary * 366) / (dblHoursPerWeek# * 52)
        Else
            dblSsalary# = (dblNewSalary * 365) / (dblHoursPerWeek# * 52)
        End If
        
        'Ticket #17654 - formula correction
        If fglbDhrs = 0 Then
            dblSsalary# = 0
        Else
            dblSsalary# = dblNewSalary / fglbDhrs
        End If
    End If
ElseIf snapJob("JB_SALCD") = "A" Then
    If snapPosSal("SH_SALCD") = "H" Then
        dblSsalary# = (dblNewSalary * dblHoursPerWeek#) * 52
    ElseIf snapPosSal("SH_SALCD") = "M" Then
        dblSsalary# = dblNewSalary * 12
    ElseIf snapPosSal("SH_SALCD") = "A" Then
        If xStep <> 0 Then
            dblSsalary# = snapJob("JB_S" & Format(xStep, "##"))
        Else
            dblSsalary# = 0
        End If
    ElseIf snapPosSal("SH_SALCD") = "D" Then
        If GetLeapYear(Year(Date)) Then
            dblSsalary# = (dblNewSalary * 366)
        Else
            dblSsalary# = (dblNewSalary * 365)
        End If
        
        'Ticket #17654 - formula correction
        If fglbDhrs = 0 Then
            dblSsalary# = 0
        Else
            dblSsalary# = (dblNewSalary / fglbDhrs) * dblHoursPerWeek# * 52
        End If
    End If
End If

' dkostka - 02/18/2002 - Added Val(Format(x, "@")) around expression to replace null with 0.
'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
'If dynSH_Job1("JB_MIDPOINT") >= 1 And dynSH_Job1("JB_MIDPOINT") <= 11 Then
'If dynSH_Job1("JB_MIDPOINT") >= 1 And dynSH_Job1("JB_MIDPOINT") <= 15 Then
If snapJob("JB_MIDPOINT") >= 1 And snapJob("JB_MIDPOINT") <= 20 Then
    Jb_No = Val(Format(snapJob("JB_S" & snapJob("JB_MIDPOINT")), "@"))
End If

fglbCOMPA# = 0

If Jb_No <> 0 And dblSsalary# <> 0 Then 'laura 03/23/98
  fglbCOMPA# = (dblSsalary# / Jb_No) * 100
End If

 
If fglbCOMPA# > 999.99 Then fglbCOMPA# = 999.99


fglbGRADE$ = "00"
xSalGrade = dblNewSalary

If xStep <> 0 Then
    If IsNumeric(dynSH_Job1("JB_S" & Format(xStep, "##"))) Then
        If snapJob("JB_SALCD") = "H" Then
            If dynSH_Job1("SH_SALCD") = "H" Then
                dblNewSalary = snapJob("JB_S" & Format(xStep, "##"))
            ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                dblNewSalary = (snapJob("JB_S" & Format(xStep, "##")) * (dblHoursPerWeek# * 52)) / 12
            ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * (dblHoursPerWeek# * 52)
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##") & "A")
                    xGrid2 = True
                End If
            ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                If GetLeapYear(Year(Date)) Then
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * (dblHoursPerWeek# * 52) / 366
                Else
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * (dblHoursPerWeek# * 52) / 365
                End If
                
                'Ticket #17654 - formula correction
                dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * fglbDhrs
            End If
        ElseIf snapJob("JB_SALCD") = "A" Then
            If dynSH_Job1("SH_SALCD") = "H" Then
                If dblHoursPerWeek# = 0 Then
                    dblNewSalary = 0
                Else
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) / (dblHoursPerWeek# * 52)
                End If
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##") & "A")
                    xGrid2 = True
                End If
            ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) / 12
            ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                dblNewSalary = snapJob("JB_S" & Format(xStep, "##"))
            ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                If GetLeapYear(Year(Date)) Then
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * 366
                Else
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * 365
                End If
                
                'Ticket #17654 - formula correction
                dblNewSalary = (snapJob("JB_S" & Format(xStep, "##")) / (dblHoursPerWeek# * 52)) * fglbDhrs
            End If
        End If
        If glbCompSerial = "S/N - 2378W" And xGrid2 = True Then 'Town of Aurora
            If dynSH_Job1("JB_S" & Format(xStep, "##") & "A") > 0 Then
                cX$ = CStr(xStep)
                If X% <= 9 Then cX$ = "0" & cX$
                fglbGRADE$ = cX$
            End If
        Else
            If dynSH_Job1("JB_S" & Format(xStep, "##")) > 0 Then
                cX$ = CStr(xStep)
                If X% <= 9 Then cX$ = "0" & cX$
                fglbGRADE$ = cX$
            End If
        End If
    End If
Else
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    'For X% = 1 To 11
    'For X% = 1 To 15
    For X% = 1 To 20
        'Added D by Bryan 28/Sep/05 Ticket#9354
        If IsNumeric(dynSH_Job1("JB_S" & Format(X%, "##"))) Then
            If snapJob("JB_SALCD") = "H" Then
                If dynSH_Job1("SH_SALCD") = "H" Then
                    xSalGrade = snapJob("JB_S" & Format(X%, "##"))
                ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                    xSalGrade = (snapJob("JB_S" & Format(X%, "##")) * (dblHoursPerWeek# * 52)) / 12
                ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                    xSalGrade = snapJob("JB_S" & Format(X%, "##")) * (dblHoursPerWeek# * 52)
                ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                    If GetLeapYear(Year(Date)) Then
                        xSalGrade = snapJob("JB_S" & Format(X%, "##")) * 366 / (dblHoursPerWeek# * 52)
                    Else
                        xSalGrade = snapJob("JB_S" & Format(X%, "##")) * 365 / (dblHoursPerWeek# * 52)
                    End If
                
                    'Ticket #17654 - formula correction
                    xSalGrade = snapJob("JB_S" & Format(X%, "##")) * fglbDhrs
                End If
            ElseIf snapJob("JB_SALCD") = "A" Then
                If dynSH_Job1("SH_SALCD") = "H" Then
                    If dblHoursPerWeek# = 0 Then
                        xSalGrade = 0
                    Else
                        xSalGrade = snapJob("JB_S" & Format(X%, "##")) / (dblHoursPerWeek# * 52)
                    End If
                ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                    xSalGrade = snapJob("JB_S" & Format(X%, "##")) * 12
                ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                    xSalGrade = snapJob("JB_S" & Format(X%, "##"))
                ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                    If GetLeapYear(Year(Date)) Then
                        xSalGrade = snapJob("JB_S" & Format(X%, "##")) * 366
                    Else
                        xSalGrade = snapJob("JB_S" & Format(X%, "##")) * 365
                    End If
                    
                    'Ticket #17654 - formula correction
                    xSalGrade = snapJob("JB_S" & Format(X%, "##")) / (dblHoursPerWeek# * 52) * fglbDhrs
                End If
            End If
            If dblNewSalary >= xSalGrade And dynSH_Job1("JB_S" & Format(X%, "##")) > 0 Then
                cX$ = CStr(X)
                If X% <= 9 Then cX$ = "0" & cX$
                fglbGRADE$ = cX$
            End If
        End If
    Next X%
End If

If IsNumeric(dynSH_Job1("JB_S1")) Then
    If dblSsalary# < dynSH_Job1("JB_S1") Then
        fglbGRADE$ = "00"
    End If
End If
snapJob.Close
End Sub

Private Sub imgSec_Click()
    If Len(glbDocImpFile) > 0 Then
        Shell "cmd /c " & GetShortName(glbDocImpFile)
    End If
End Sub

Private Sub Populate_Steps()
    Dim rsJobMst As New ADODB.Recordset
    Dim SQLQ As String
    Dim xStep As Integer
    
    'If glbMultiGrid And Len(xSalGrid) > 0 Then
    '    SQLQ = "SELECT * FROM HRJOB_GRADE WHERE JB_CODE='" & xJob & "' AND JB_GRID='" & xSalGrid & "'"
    'Else
        SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE='" & clpNJob.Text & "'"
    'End If
    rsJobMst.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsJobMst.EOF Then
        comStep.Clear
        For xStep = 1 To 20
            If rsJobMst("JB_S" & Trim(Str(xStep))) <> 0 Then
                comStep.AddItem Format(xStep, fglbFrmt)
            End If
        Next
    Else
        comStep.Clear
    End If
    rsJobMst.Close
    Set rsJobMst = Nothing
    
End Sub

Private Sub optUpdateType_Click(Index As Integer, Value As Integer)
    'Hide / Show controls based on Update Type
    If optUpdateType(0) Then
        lblStep.Visible = False
        comStep.Visible = False
        
        lblExpRtnDate.Top = 6699
        dlpExpRtnDate.Top = 6654
        lblStep.Top = 7095
        comStep.Top = 7035
        
        lblExpRtnDate.Visible = True
        dlpExpRtnDate.Visible = True
        
        'Ticket #28457 City of Niagara Falls - Show the default values
        If glbCompSerial = "S/N - 2276W" Then
            If Len(clpNDept.Text) = 0 Then clpNDept.Text = "4900"
            If Len(clpNDiv.Text) = 0 Then clpNDiv.Text = "LAYO"
            If Len(clpNLoc.Text) = 0 Then clpNLoc.Text = "NW"
            If Len(clpNAdminBy.Text) = 0 Then clpNAdminBy.Text = "4900"   'Ticket #29480 - Additional fields
        End If
        
        'Ticket #29382 - They want Notes 2 to be user input fields. Notes 1 is Expected Date of Return
        'txtComment.Enabled = False
    Else
        lblExpRtnDate.Visible = False
        dlpExpRtnDate.Visible = False
        
        lblExpRtnDate.Top = 6699
        dlpExpRtnDate.Top = 6654
        lblStep.Top = 6699
        comStep.Top = 6654
        
        lblStep.Visible = True
        comStep.Visible = True
    
        'Ticket #28457 City of Niagara Falls - Clear values
        If glbCompSerial = "S/N - 2276W" Then
            clpNJob.Text = ""
            dlpNSDate.Text = ""
            clpNReason.Text = ""
            dlpExpRtnDate.Text = ""
            clpNDept.Text = ""
            clpNDiv.Text = ""
            clpNLoc.Text = ""
            clpNAdminBy.Text = ""
            txtComment.Text = ""
            txtComments2.Text = ""
        End If
        
        'Ticket #29382 - They want Notes 1 and Notes 2 to be user input fields
        'txtComment.Enabled = True
    End If
End Sub

Private Sub UpdateLogPosSal(xEmpnbr, xSalCD, xSal, xUptDesc As String, xNew As String, xOld As String, xNewSDate, xOldSDate)
    Dim SQLQ As String
    Dim rsWRK As New ADODB.Recordset
    
    SQLQ = "SELECT * FROM HREMPHIS_WRK WHERE EE_WRKEMP='" & glbUserID & "'"
    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    rsWRK.AddNew
    rsWRK("EE_EMPNBR") = xEmpnbr
    rsWRK("EE_SALCD") = xSalCD
    rsWRK("EE_SALARY") = xSal
    rsWRK("EE_HISTYPE") = Left(xUptDesc, 50) 'Data Being Updated
    rsWRK("EE_NEWVALUE") = Left(xNew, 50)
    rsWRK("EE_OLDVALUE") = Left(xOld, 50)
    rsWRK("EE_CHGDATE") = xNewSDate    'New Start Date
    rsWRK("EE_DOT") = xOldSDate        'Old Start Date
    rsWRK("EE_LUSER") = glbUserID      'By Whom
    rsWRK("EE_LTIME") = Time$          'Time of Change
    rsWRK("EE_LDATE") = Date           'Date of Change
    rsWRK("EE_WRKEMP") = glbUserID
    'rsWrk("TERM_SEQ") = xOrder
    rsWRK.Update
    
    rsWRK.Close
    Set rsWRK = Nothing
End Sub

Private Sub ExcelRpt_Log() 'Ticket #27605 Franks 10/16/2015
Dim exApp As Object, exBook As Object, exSheet As Object
Dim rsWRK As New ADODB.Recordset
Dim xlsFileTmp As String, xlsFileMat As String
Dim SQLQ As String
Dim K As Long
Dim xRow, xRows
Dim xNewEmp As Boolean
Dim xCurName, xNextName
Dim xCurGroup, xNextGroup
Dim rsHRPARCO As New ADODB.Recordset

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""
    

    SQLQ = "SELECT HREMPHIS_WRK.*,ED_SURNAME,ED_FNAME,ED_VADIM2 FROM HREMPHIS_WRK "
    'SQLQ = SQLQ & "LEFT JOIN qry_HREMP ON HREMPHIS_WRK.KEY_EMPNBR = qry_HREMP.KEY_EMPNBR "
    SQLQ = SQLQ & "LEFT JOIN HREMP ON HREMPHIS_WRK.EE_EMPNBR = HREMP.ED_EMPNBR "
    SQLQ = SQLQ & "WHERE EE_WRKEMP='" & glbUserID & "' "
    'SQLQ = SQLQ & "ORDER BY ED_SURNAME,ED_FNAME,EE_LDATE DESC,TERM_SEQ "
    SQLQ = SQLQ & "ORDER BY ED_SURNAME,ED_FNAME,EE_LDATE "

    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsWRK.EOF Then
        'MsgBox "There is no any record in this Selection Criteria"
        rsWRK.Close
        Exit Sub
    End If
    
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "EmpPosSalLogTmp.xls"
    
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "EmpPosSalLog_" & glbUserID & ".xls"
    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    FileCopy xlsFileTmp, xlsFileMat
    
    Set exApp = CreateObject("Excel.Application") 'New Excel.Application
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)
    
    SQLQ = "SELECT PC_NAME FROM HRPARCO"
    rsHRPARCO.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsHRPARCO.EOF Then
        exSheet.Cells(1, 2) = rsHRPARCO("PC_NAME")
    End If
    rsHRPARCO.Close
    Set rsHRPARCO = Nothing

    exSheet.Cells(2, 1).Font.Bold = True
    exSheet.Cells(2, 1) = "Date: " & Date
    exSheet.Cells(3, 1).Font.Bold = True
    exSheet.Cells(3, 1) = "Time: " & Time$

    xRows = rsWRK.RecordCount
    xRow = 0

    K = 5
    xNewEmp = True
    xCurGroup = "**"
    xCurName = "***"
    
    'If Not (xCurName = xNextName) Then
        'K = K + 1
        'exSheet.Cells(K, 1).Font.Bold = True
        'exSheet.Cells(K, 1) = "Employee Number and Name: " & rsWRK("EE_EMPNBR") & " " & xNextName
    '    xCurName = xNextName
    '    K = K + 1
        exSheet.Cells(K, 1) = "Employee #": exSheet.Cells(K, 1).Font.Bold = True
        exSheet.Cells(K, 2) = "Name": exSheet.Cells(K, 2).Font.Bold = True
        exSheet.Cells(K, 3) = "Data Type": exSheet.Cells(K, 3).Font.Bold = True
        exSheet.Cells(K, 4) = "New Data": exSheet.Cells(K, 4).Font.Bold = True
        exSheet.Cells(K, 5) = "Prev. Data": exSheet.Cells(K, 5).Font.Bold = True
        exSheet.Cells(K, 6) = "New Eff. Date": exSheet.Cells(K, 6).Font.Bold = True
        exSheet.Cells(K, 7) = "Prev. Eff. Date": exSheet.Cells(K, 7).Font.Bold = True
        exSheet.Cells(K, 8) = "Per": exSheet.Cells(K, 8).Font.Bold = True
        exSheet.Cells(K, 9) = "Salary": exSheet.Cells(K, 9).Font.Bold = True
        
        exSheet.Cells(K, 10) = "Updated By": exSheet.Cells(K, 10).Font.Bold = True
        exSheet.Cells(K, 11) = "Updated Date": exSheet.Cells(K, 11).Font.Bold = True
        exSheet.Cells(K, 12) = "Updated Time": exSheet.Cells(K, 12).Font.Bold = True
        K = K + 1
    'End If
    
    Do While Not rsWRK.EOF
        MDIMain.panHelp(0).FloodPercent = (xRow / xRows) * 100
        DoEvents
        xRow = xRow + 1

        ''If Not comGroup(0).Text = "(none)" Then
        ''    If IsNull(rsWRK("EE_GROUP_DESC")) Then xNextGroup = "" Else xNextGroup = rsWRK("EE_GROUP_DESC")
        ''    If Not (xCurGroup = xNextGroup) Then
        ''        K = K + 1
        ''        xCurGroup = xNextGroup
        ''        exSheet.Cells(K, 1).Font.Bold = True
        ''        exSheet.Cells(K, 1) = xNextGroup
        ''        K = K + 1
        ''    End If
        ''End If

        xNextName = rsWRK("ED_SURNAME") & ", " & rsWRK("ED_FNAME")
        exSheet.Cells(K, 1) = rsWRK("EE_EMPNBR")
        exSheet.Cells(K, 2) = xNextName
        exSheet.Cells(K, 3) = rsWRK("EE_HISTYPE")
        exSheet.Cells(K, 4) = rsWRK("EE_NEWVALUE")  'New Data
        exSheet.Cells(K, 5) = rsWRK("EE_OLDVALUE")  'Previous Data
        exSheet.Cells(K, 6) = rsWRK("EE_CHGDATE")   'New Effective Date
        exSheet.Cells(K, 7) = rsWRK("EE_DOT")       'Previous Effective Date
        exSheet.Cells(K, 8) = rsWRK("EE_SALCD")     'Per
        exSheet.Cells(K, 9) = rsWRK("EE_SALARY")    'Salary
        exSheet.Cells(K, 10) = rsWRK("EE_LUSER")    'By Whom
        exSheet.Cells(K, 11) = rsWRK("EE_LDATE")    'Date of Change
        exSheet.Cells(K, 12) = rsWRK("EE_LTIME")    'Time of Change
        
        K = K + 1

        rsWRK.MoveNext
    Loop
    rsWRK.Close
    
    If Not exBook Is Nothing Then
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
    End If

    Call Pause(1)
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(2).Caption = ""
    
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
    
End Sub

Private Sub UpdPaymentTypeVadim(xEmpNo, xType)
Dim rsEmp As New ADODB.Recordset
Dim SQLQ

If glbVadim Then
    If Vadim_PayType_field = "" Then Exit Sub
    
    SQLQ = "SELECT " & Vadim_PayType_field & " FROM HREMP WHERE ED_EMPNBR =" & xEmpNo
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmp.EOF Then
        If xType = "LO" Or xType = "LAYO" Then
            rsEmp(Vadim_PayType_field) = "L"
        Else
            rsEmp(Vadim_PayType_field) = xType
        End If
        rsEmp.Update
    End If
    rsEmp.Close
    Set rsEmp = Nothing
End If
End Sub

Private Function getEmpHistoryValue(xFieldName, xCurrentVal, xEmpnbr)
    Dim rsEmpHis As New ADODB.Recordset
    Dim SQLQ As String
    
    SQLQ = "SELECT * FROM HREMPHIS WHERE EE_EMPNBR = " & xEmpnbr & " AND EE_NEW" & xFieldName & " = '" & xCurrentVal & "' ORDER BY EE_CHGDATE DESC"
    rsEmpHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmpHis.EOF Then
        getEmpHistoryValue = rsEmpHis("EE_OLD" & xFieldName)
    Else
        getEmpHistoryValue = ""
    End If
    rsEmpHis.Close
    Set rsEmpHis = Nothing
    
End Function

