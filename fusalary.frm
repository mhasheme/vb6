VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmUSalary 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Mass Update Salary"
   ClientHeight    =   8895
   ClientLeft      =   15
   ClientTop       =   1230
   ClientWidth     =   11400
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8895
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame frDirIndSal 
      Caption         =   "Mass Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   55
      Top             =   4920
      Visible         =   0   'False
      Width           =   3375
      Begin Threed.SSOption optDirIndSal 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Tag             =   "Direct Salary Update"
         Top             =   240
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Direct Salary"
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
      Begin Threed.SSOption optDirIndSal 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   14
         Tag             =   "Indirect Salary Update"
         Top             =   240
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Indirect Salary"
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
      End
   End
   Begin VB.CheckBox chkUpdAttendance 
      Caption         =   "Update Employee's Attendance with New Salary from Effective Date forward"
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
      Left            =   360
      TabIndex        =   33
      Tag             =   "40-Update Attendance records with Salary from Effective Date forward -y/n"
      Top             =   8520
      Width           =   7095
   End
   Begin VB.ComboBox cmbPrecision 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9240
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Tag             =   "Enter Number of Decimal Places for Salary"
      Top             =   6225
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1930
      TabIndex        =   8
      Top             =   3055
      Width           =   435
   End
   Begin INFOHR_Controls.DateLookup dlpNextReviewDate 
      Height          =   285
      Left            =   4680
      TabIndex        =   16
      Tag             =   "40-Next Date to Review Salary"
      Top             =   5820
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpEffectiveDate 
      Height          =   285
      Left            =   1740
      TabIndex        =   15
      Tag             =   "41-Effective Date of Salary Record"
      Top             =   5820
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
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
      Index           =   6
      Left            =   6600
      TabIndex        =   32
      Tag             =   "Steps of salary change "
      Top             =   7725
      Visible         =   0   'False
      Width           =   1035
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
      Index           =   5
      Left            =   6600
      TabIndex        =   29
      Tag             =   "Steps of salary change "
      Top             =   7380
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.ComboBox cmbRound 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Tag             =   "Select Yes / No"
      Top             =   6225
      Width           =   990
   End
   Begin VB.ComboBox comPayType 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "fusalary.frx":0000
      Left            =   6840
      List            =   "fusalary.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Tag             =   "Annually or Hourly"
      Top             =   5805
      Width           =   1395
   End
   Begin VB.TextBox txtSALCD 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   6360
      MaxLength       =   1
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5820
      Visible         =   0   'False
      Width           =   375
   End
   Begin Threed.SSOption optDollars 
      Height          =   285
      Left            =   240
      TabIndex        =   18
      Tag             =   "Salary update - dollars"
      Top             =   6240
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   503
      _StockProps     =   78
      Caption         =   "Dollars"
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
   Begin Threed.SSOption optPct 
      Height          =   285
      Left            =   2610
      TabIndex        =   20
      Tag             =   "Salary update - percentage"
      Top             =   6240
      Width           =   915
      _Version        =   65536
      _ExtentX        =   1614
      _ExtentY        =   503
      _StockProps     =   78
      Caption         =   "Percent"
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
   End
   Begin MSMask.MaskEdBox medAmountChng 
      Height          =   285
      Index           =   5
      Left            =   6600
      TabIndex        =   30
      Tag             =   "21-Amount in dollars of salary change"
      Top             =   7440
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
   Begin MSMask.MaskEdBox medAmountChng 
      Height          =   285
      Index           =   6
      Left            =   6600
      TabIndex        =   31
      Tag             =   "21-Amount in dollars of salary change"
      Top             =   7785
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
   Begin Threed.SSOption optStep 
      Height          =   285
      Left            =   3630
      TabIndex        =   21
      Tag             =   "Salary update - Step Increase"
      Top             =   6240
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   503
      _StockProps     =   78
      Caption         =   "Step Increase"
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   330
      TabIndex        =   25
      Tag             =   "01-Reason for change in salary - Code "
      Top             =   7395
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDRC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   330
      TabIndex        =   26
      Tag             =   "01-Reason for change in salary - Code "
      Top             =   7740
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDRC"
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   330
      TabIndex        =   24
      Tag             =   "01-Reason for change in salary - Code "
      Top             =   7050
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDRC"
   End
   Begin Threed.SSOption optFixed 
      Height          =   285
      Left            =   1230
      TabIndex        =   19
      Tag             =   "Salary update -Fixed Dollars"
      Top             =   6240
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   503
      _StockProps     =   78
      Caption         =   "Fixed Dollars"
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
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   1800
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport vbxCrystal1 
      Left            =   1200
      Top             =   7920
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
      Index           =   4
      Left            =   6600
      TabIndex        =   27
      Tag             =   "Steps of salary change "
      Top             =   7035
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSMask.MaskEdBox medAmountChng 
      Height          =   285
      Index           =   4
      Left            =   6600
      TabIndex        =   28
      Tag             =   "21-Amount in dollars of salary change "
      Top             =   7080
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
      Left            =   1620
      TabIndex        =   12
      Tag             =   "41-As of Date of Salary to apply the Percent update on"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin VB.Image imgHelp 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1240
      Picture         =   "fusalary.frx":0004
      Stretch         =   -1  'True
      Top             =   4455
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
      Left            =   120
      TabIndex        =   54
      Top             =   4485
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblDecimal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Decimal Precision"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7920
      TabIndex        =   53
      Top             =   6255
      Visible         =   0   'False
      Width           =   1335
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
      TabIndex        =   52
      Top             =   4140
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblShift 
      Caption         =   "Shift"
      Height          =   255
      Left            =   120
      TabIndex        =   51
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
      TabIndex        =   50
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
      TabIndex        =   49
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
      TabIndex        =   48
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
      TabIndex        =   47
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
      TabIndex        =   46
      Top             =   3800
      Width           =   1035
   End
   Begin VB.Label lblRound 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Round"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6090
      TabIndex        =   45
      Top             =   6255
      Width           =   615
   End
   Begin VB.Label lblDChg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6660
      TabIndex        =   44
      Top             =   6720
      Width           =   660
   End
   Begin VB.Label lblReason 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   43
      Top             =   6690
      Width           =   660
   End
   Begin VB.Label lblNextReview 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Next Review"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3600
      TabIndex        =   41
      Top             =   5865
      Width           =   915
   End
   Begin VB.Label lblEffectiveDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Effective Date"
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
      Left            =   180
      TabIndex        =   40
      Top             =   5865
      Width           =   1245
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
      TabIndex        =   39
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
      TabIndex        =   38
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
      TabIndex        =   37
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
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   34
      Top             =   420
      Width           =   555
   End
End
Attribute VB_Name = "frmUSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbCOMPA#, fglbGRADE$
Dim snapNoUpSal As New ADODB.Recordset 'Frank 4/10/2000
Dim fglbSDate As Variant
Dim snapSalary As New ADODB.Recordset
Dim fglbFrmt
Dim OSalary, NSalary, OEDate, NEDate, ONDate, NNDate, EmpNo&, dblWHours#, OTOTAL, OIndSalary
Dim oPayP, NPayp, OJOB1, xGrade, OSalCD, oGrade
Dim oPayrollID, oGrid
Dim lngRecs&, SkipRec&
Dim GLfocus
Dim fglAddDel As String
Dim strEMPLIST 'George Mar 14,2006
Dim MailBody
Dim fglbDhrs

Private Function AUDITSALY()
Dim TA As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim TB As New ADODB.Recordset
Dim strFields As String
On Error GoTo AUDIT_ERR
AUDITSALY = False


TB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & EmpNo&, gdbAdoIhr001, adOpenForwardOnly
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

If OSalary <> NSalary Then GoTo MODUPD
'If OPayp <> NPayp Then GoTo MODUPD      'laura jan 28, 1998
If OEDate <> NEDate Then GoTo MODUPD
If ONDate <> NNDate Then GoTo MODUPD
GoTo MODNOUPD

MODUPD:
TA.AddNew
TA("AU_LOC_TABL") = "EDLC": TA("AU_EMP_TABL") = "EDEM": TA("AU_SUPCODE_TABL") = "EDSP": TA("AU_ORG_TABL") = "EDOR"
TA("AU_PAYP_TABL") = "SDPP": TA("AU_BCODE_TABL") = "BNCD": TA("AU_TREAS_TABL") = "TERM": TA("AU_DOLENT_TABL") = "EDOL"
TA("AU_EARN_TABL") = "EARN"
TA("AU_NEWEMP") = "N"
TA("AU_PTUPL") = xPT
TA("AU_DIVUPL") = xDiv
TA("AU_SALARY") = NSalary
TA("AU_OLDSAL") = OSalary
TA("AU_PAYP") = oPayP ' FRANK 4/5/2000    'NPayp  Laura jan 28, 1998
TA("AU_OLDPAYP") = oPayP    '    ""
TA("AU_JOB") = OJOB1         ' FRANK 4/5/2000
TA("AU_GRID") = oGrid
If glbMulti Then TA("AU_PAYROLL_ID") = oPayrollID
TA("AU_SALCD") = OSalCD
TA("AU_WHRS") = dblWHours# 'ADDED BY RAUBREY 7/7/97
If OEDate <> NEDate Then TA("AU_SEDATE") = IIf(IsDate(NEDate), NEDate, Null)   'Jaddy 11/15/99
If ONDate <> NNDate Then TA("AU_SNDATE") = IIf(IsDate(NNDate), NNDate, Null)  'Jaddy 11/15/99

'Ticket #23666 - Update with Salary Reason for Change as well.
TA("AU_SREASON") = clpCode(4).Text

TA("AU_COMPNO") = "001"
TA("AU_EMPNBR") = EmpNo&

'Ticket #23943 - Town of Orangeville noticed the LDATE was not getting updated properly - Jerry asked to fix this as per Salary screen.
If glbCompSerial = "S/N - 2227W" And (xPT = "SE" Or xPT = "OT") Then ' CCAC Kingston, see ticket #3296
    TA("AU_LDATE") = Format(DateAdd("d", 14, NEDate), "SHORT DATE")
Else
    'Ticket #23943 - Town of Orangeville
    If glbCompSerial = "S/N - 2383W" Then
        If CVDate(NEDate) > CVDate(Date) Then
            TA("AU_LDATE") = Format(NEDate, "SHORT DATE")
        Else
            TA("AU_LDATE") = Date
        End If
    Else
        TA("AU_LDATE") = Format(NEDate, "SHORT DATE")
    End If
End If
'TA("AU_LDATE") = Format(NEDate, "SHORT DATE")

TA("AU_LUSER") = glbUserID
TA("AU_LTIME") = Time$
TA("AU_UPLOAD") = "N"
TA("AU_TYPE") = "A"
'If glbSoroc Or glbSyndesis Then
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & EmpNo&
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then TA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
    End If
    rsEmp.Close
'End If
TA.Update


MODNOUPD:
AUDITSALY = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me
End Function

Private Function chkMUSalary()

Dim SQLQ As String, Msg$, dd&, X%
Dim DgDef As Variant, Title$, Response%, DCurSHDate  As Variant

chkMUSalary = False

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

If dlpAsofDate.Visible = True Then
    If Len(dlpAsofDate.Text) > 0 Then
        If Not IsDate(dlpAsofDate.Text) Then
            MsgBox "As of Date is invalid"
            dlpAsofDate.SetFocus
            Exit Function
        End If
    End If
End If


For X% = 4 To 6
    If Len(clpCode(X%).Text) > 0 Then
        If clpCode(X%) = "Unassigned" Then
            MsgBox "Invalid Reason Code"
            clpCode(X%).SetFocus
            Exit Function
        End If
        
        If optStep Then
            If Len(comStep(X%)) < 1 Then
                MsgBox "Step is required if code entered"
                comStep(X%).SetFocus
                Exit Function
            Else
                'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
                'If (Val(comStep(x%)) < 1 Or Val(comStep(x%)) > 11) And comStep(x%) <> "Next Step" Then
                'If (Val(comStep(X%)) < 1 Or Val(comStep(X%)) > 15) And comStep(X%) <> "Next Step" Then
                If (Val(comStep(X%)) < 1 Or Val(comStep(X%)) > 20) And comStep(X%) <> "Next Step" Then
                    MsgBox "Step is not valid"
                    comStep(X%).SetFocus
                    Exit Function
                End If
            End If
        Else
            If Len(medAmountChng(X%)) < 1 Then
                MsgBox "Salary is required if code entered"
                medAmountChng(X%).SetFocus
                Exit Function
            Else
                If Not IsNumeric(medAmountChng(X%)) Then   'laura jan 12, 1998
                    MsgBox "Salary is not valid"
                    Exit Function
                End If
            End If
        End If
    End If
Next X%

If fglAddDel = "Add" Then
    If Len(clpCode(4).Text) < 1 Then
        MsgBox "You must have at least one change"
        clpCode(4).SetFocus
        Exit Function
    End If
End If

If Len(dlpEffectiveDate.Text) < 1 Then
    MsgBox "Effective Date is required"
    dlpEffectiveDate.SetFocus
    Exit Function
Else
    If Not IsDate(dlpEffectiveDate.Text) Then
        MsgBox "Not a Valid Effective Date"
        dlpEffectiveDate.SetFocus
        Exit Function
    End If
End If

If Len(dlpNextReviewDate.Text) < 1 Then
'Jaddy 11/15
'    MsgBox "Next Review Date is required"
'    dlpNextReviewDate.SetFocus
'    Exit Function
Else
    If Not IsDate(dlpNextReviewDate.Text) Then
        MsgBox "Next Review Date is invalid"
        dlpNextReviewDate.SetFocus
        Exit Function
    End If
    
    dd& = DateDiff("d", CVDate(dlpEffectiveDate.Text), CVDate(dlpNextReviewDate.Text))
    
    If dd& < 0 Then
        MsgBox "Next Review preceeds Effective date of salary "
        dlpEffectiveDate.SetFocus
        Exit Function
    End If
End If

chkMUSalary = True
Exit Function

chkSalH_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkSal", "HR SALARY", "edit/Add")
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

Public Sub cmdModify_Click()
Dim Title$, Msg$, DgDef As Variant, Response%
Dim I As Integer

clpDiv.SetFocus

On Error GoTo Mod_Err

If Not gSec_Upd_Salary Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If


txtSALCD = Left(comPayType, 1)

fglAddDel = "Add"

If Not chkMUSalary() Then Exit Sub

If Not modGet_Salary_Records() Then Exit Sub       ' get selection - form level

'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
If glbCompSerial = "S/N - 2494W" Then
    If optDirIndSal(1).Value = True Then
        Title$ = "Update Indirect Salary"
    Else
        Title$ = "Update Direct Salary"
    End If
Else
    Title$ = "Update Salary"
End If

Msg$ = ""
Msg$ = Msg$ & "Must have Current Salary." & Chr(10)
Msg$ = Msg$ & "Must have Current Position with Position Code and Position Start Date that matches the current Salary record." & Chr(10)
'Msg$ = Msg$ & "Must have Current Postion with Position Code and Position Start Date match current Salary's. " & Chr(10)

'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
If glbCompSerial = "S/N - 2494W" Then
    If optDirIndSal(1).Value = True Then
        'Do not display the Salary Step message line
    Else
        Msg$ = Msg$ & "Must have Position Master to count the Salary Step. " & Chr(10)
    End If
Else
    Msg$ = Msg$ & "Must have Position Master to count the Salary Step. " & Chr(10)
End If

Msg$ = Msg$ & Chr(10)

If snapSalary.BOF And snapSalary.EOF Then
    Msg$ = Msg$ & "No Employees with this selection criteria exist!  " & Chr(10)
    Msg$ = Msg$ & "Please ensure the Hourly/Annually selection box is set correctly to match the group you want to update." & Chr(10)
    Msg$ = Msg$ & Chr(10)
    DgDef = MB_OK + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
    MsgBox Msg$, , Title$
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

lngRecs& = snapSalary.RecordCount
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
If Not modUpdateSelection() Then Exit Sub


Screen.MousePointer = DEFAULT
If SkipRec = 0 Then
    MsgBox "Records Updated Successfully"
Else
    Msg$ = lngRecs& - SkipRec & " record(s) was updated and " & SkipRec & " record(s) was skipped." & Chr(10)
    Msg$ = Msg$ & "Please click on Print or View icon on the Toolbar to see the reports for the updated and skipped employees."
    MsgBox Msg$, , Title$
End If

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


    If Not PrtForm("Mass Update Salary Report Criteria", Me) Then Exit Sub
    
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
    
    
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal1.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT

    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset

    Me.vbxCrystal1.Action = 1
    vbxCrystal1.Reset
    MDIMain.Timer1.Enabled = True
 '   cmdPrint.Enabled = True
 '   cmdView.Enabled = True
End Sub

Private Sub cmbRound_Click()
    If cmbRound.Text = "Yes" Then
        cmbPrecision.Visible = True
        lblDecimal.Visible = True
        cmbPrecision.Text = glbCompDecHR
    Else
        cmbPrecision.Visible = False
        lblDecimal.Visible = False
    End If
End Sub

Private Sub cmbRound_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub comPayType_GotFocus()
Call SetPanHelp(ActiveControl)

End Sub

Private Sub comStep_Click(Index As Integer)
medAmountChng(Index) = comStep(Index)
End Sub

Private Sub comStep_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMUSALARY"
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim X%, RFound As Integer, xx%

Screen.MousePointer = HOURGLASS

glbOnTop = "FRMUSALARY"

Call CRLIST_EECAT

Call setRptCaption(Me)

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

For X% = 4 To 6
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    'For xx% = 1 To 11                   'Jaddy 8/31/99
    'For xx% = 1 To 15
    For xx% = 1 To 20
        comStep(X%).AddItem Format(xx%, fglbFrmt)
    Next
    comStep(X%).AddItem "Next Step" 'Frank 4/10/2000
Next X%

If glbWFC And glbCountry = "AUSTRALIA" Then
    medAmountChng(4).Format = "$#,##0.0000;($#,##0.0000)"
    medAmountChng(5).Format = "$#,##0.0000;($#,##0.0000)"
    medAmountChng(6).Format = "$#,##0.0000;($#,##0.0000)"
End If

If glbCompSerial = "S/N - 2172W" Then 'Lanark Ticket #17221 by Frank 08/19/2009
    lblPosGroup.Caption = "Salary Level"
    clpCode(1).TABLTitle = "Salary Level Code"
End If

comPayType.ListIndex = 0

cmbPrecision.AddItem "0"
cmbPrecision.AddItem "2"
cmbPrecision.AddItem "3"
cmbPrecision.AddItem "4"

'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
If glbCompSerial = "S/N - 2494W" Then
    frDirIndSal.Visible = True
Else
    frDirIndSal.Visible = False
End If

If glbMulti Then lblUnion.ForeColor = &HC000C0: lblPT.ForeColor = &HC000C0

If Not glbMultiGrid Then lblGrid.Visible = False: clpGrid.Visible = False

Call INI_Controls(Me)

If glbWFC Then 'Ticket #25911 Franks 10/21/2014
    clpJob.TransDiv = glbWFCUserSecList
End If

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
Set frmUSalary = Nothing  'carmen apr 2000
End Sub


Private Sub imgHelp_Click()
Dim MsgStr As String
    MsgStr = "System will find the employee's with Salary Effective Date >= 'As of Date' and will compute the Percentage of Change based on that Salary which will then be applied to the existing Current Salary of the employee to compute New Current Salary."
    MsgBox MsgStr, vbInformation

End Sub

Private Sub medAmountChng_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
' dkostka - 07/10/2001 - Added multiplication code below, this way if the user clicks on the
'   Change Amount fields more than once, it won't keep getting smaller.
If IsNumeric(medAmountChng(Index)) Then
    If optPct Then medAmountChng(Index) = medAmountChng(Index) * 100
End If
End Sub
Private Sub medAmountChng_LostFocus(Index As Integer)
If IsNumeric(medAmountChng(Index)) Then
    If optPct Then medAmountChng(Index) = medAmountChng(Index) / 100
End If
End Sub

Private Function modGet_Salary_Records()

Dim BD As Integer
Dim SQLQ As String, countr As Integer
Dim Dat1 As Variant, Dat2 As Variant
Dim iOneWhere As Integer, strJob$, strTm$, X%

modGet_Salary_Records = False
On Error GoTo modGet_Salary_Records_Err
strTm$ = Time$
Dim Dt As Variant
Dt = Date$

Screen.MousePointer = HOURGLASS

If glbOracle Then
    If glbMultiGrid Then
        SQLQ = "SELECT HR_SALARY_HISTORY.*, HRJOB_GRADE.* "
        SQLQ = SQLQ & " FROM HR_SALARY_HISTORY, HRJOB_GRADE ,HR_JOB_HISTORY "
        SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_JOB=HRJOB_GRADE.JB_CODE "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_GRID=HRJOB_GRADE.JB_GRID "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_JOB=HR_JOB_HISTORY.JH_JOB"
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_GRID=HR_JOB_HISTORY.JH_GRID"
        SQLQ = SQLQ & " AND SH_CURRENT<>0 AND JH_CURRENT<>0"
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
        SQLQ = "SELECT HR_SALARY_HISTORY.*, HRJOB_GRADE.* "
        SQLQ = SQLQ & " FROM (HR_SALARY_HISTORY INNER JOIN HRJOB_GRADE "
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_JOB=HRJOB_GRADE.JB_CODE  "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_GRID=HRJOB_GRADE.JB_GRID) "
        SQLQ = SQLQ & " INNER JOIN HR_JOB_HISTORY "
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR"
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_JOB=HR_JOB_HISTORY.JH_JOB"
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_GRID=HR_JOB_HISTORY.JH_GRID"
        SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND JH_CURRENT<>0"
    Else
        SQLQ = "SELECT HR_SALARY_HISTORY.*, HRJOB.* "
        SQLQ = SQLQ & " FROM ((HR_SALARY_HISTORY INNER JOIN HRJOB "
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_JOB=HRJOB.JB_CODE)  "
        SQLQ = SQLQ & " INNER JOIN HR_JOB_HISTORY "
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR"
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_JOB=HR_JOB_HISTORY.JH_JOB"
        If glbMulti Then
            SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_SDATE=HR_JOB_HISTORY.JH_SDATE"
        End If
        SQLQ = SQLQ & " )"
        SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND JH_CURRENT<>0"
    End If
End If
If Len(clpJob.Text) > 0 Then SQLQ = SQLQ & " AND SH_JOB IN ('" & Replace(clpJob.Text, ",", "','") & "') "
If Len(clpGrid.Text) > 0 Then SQLQ = SQLQ & " AND SH_GRID IN ('" & Replace(clpGrid.Text, ",", "','") & "') "
SQLQ = SQLQ & " AND SH_SALCD = '" & txtSALCD & "'"
'Ticket #27555 Franks 09/18/2015 - missing ')', added it
If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND SH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD IN ('" & Replace(clpCode(1).Text, ",", "','") & "')) "

SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn
If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND ED_DEPTNO IN ('" & Replace(clpDept.Text, ",", "','") & "') "
If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND ED_DIV IN ('" & Replace(clpDiv.Text, ",", "','") & "') "
If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND ED_EMP IN ('" & Replace(clpCode(3).Text, ",", "','") & "') "
If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND ED_LOC IN ('" & Replace(clpCode(0).Text, ",", "','") & "') "
If Len(clpCode(7).Text) > 0 Then SQLQ = SQLQ & " AND ED_SECTION IN ('" & Replace(clpCode(7).Text, ",", "','") & "') "
If Len(txtShift.Text) > 0 Then SQLQ = SQLQ & " AND ED_SHIFT = '" & txtShift.Text & "' "

If Len(elpEEID.Text) > 0 Then SQLQ = SQLQ & " AND SH_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
If glbMulti Then SQLQ = SQLQ & ") "
If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND " & IIf(glbMulti, "JH_ORG", "ED_ORG") & " IN ('" & Replace(clpCode(2).Text, ",", "','") & "') "
If Len(Trim(clpPT.Text)) > 0 Then SQLQ = SQLQ & " AND " & IIf(glbMulti, "JH_PT", "ED_PT") & " IN ('" & Replace(clpPT.Text, ",", "','") & "') "
If glbNoNONE Then SQLQ = SQLQ & " AND " & IIf(glbMulti, "JH_ORG", "ED_ORG") & " <> 'NONE' "
If glbNoEXEC Then SQLQ = SQLQ & " AND " & IIf(glbMulti, "JH_ORG", "ED_ORG") & " <> 'EXEC' "  'Hemu -EXE
If Not glbMulti Then SQLQ = SQLQ & ") "


If snapSalary.State <> 0 Then snapSalary.Close
snapSalary.Open SQLQ, gdbAdoIhr001, adOpenStatic

modGet_Salary_Records = True

Exit Function

modGet_Salary_Records_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modGet_Salary_Records", "qry_MU_Salary", "Select")
Resume Next

End Function

Private Sub modSetCOMPA_GRADE(dblNewSalary)
Dim X%, cX$
Dim dblSsalary#, dblHoursPerWeek#, ssalary@
Dim Jb_No#

'SET COMPA RATIO
'================
ssalary@ = dblNewSalary
'Franks May 17,2002 for error of Invalid Use of Null Ticket 2252
'dblHoursPerWeek# = snapSalary("SH_WHRS")
If IsNull(snapSalary("SH_WHRS")) Then
    dblHoursPerWeek# = 0
Else
    dblHoursPerWeek# = snapSalary("SH_WHRS")
End If
'Franks May 17,2002 for error of Invalid Use of Null Ticket 2252

'MOnthly and DAily added by Bryan 28/Sep/05 Ticket#9354
If (glbCompSerial = "S/N - 2378W") And (snapSalary("SH_SALCD") <> snapSalary("JB_SALCD")) And snapSalary("SH_SALCD") <> "M" Then   'Town of Aurora
    dblSsalary# = dblNewSalary
Else
If snapSalary("SH_SALCD") = "H" Then
    If snapSalary("JB_SALCD") = "A" Then
        dblSsalary# = (dblNewSalary * dblHoursPerWeek#) * 52
    ElseIf snapSalary("JB_SALCD") = "M" Then
        dblSsalary# = dblNewSalary * 12
    ElseIf snapSalary("JB_SALCD") = "D" Then
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
ElseIf snapSalary("SH_SALCD") = "A" Then
    If snapSalary("JB_SALCD") = "H" Then
        'Franks May 17,2002 for error of Invalid Use of Null Ticket 2252
        'dblSsalary# = (dblNewSalary / dblHoursPerWeek#) / 52
        If dblHoursPerWeek# = 0 Then
            dblSsalary# = 0
        Else
            dblSsalary# = (dblNewSalary / dblHoursPerWeek#) / 52
        End If
        'Franks May 17,2002 for error of Invalid Use of Null Ticket 2252
    ElseIf snapSalary("JB_SALCD") = "A" Then
        dblSsalary# = dblNewSalary
    ElseIf snapSalary("JB_SALCD") = "M" Then
        dblSsalary# = dblNewSalary / 12
    ElseIf snapSalary("JB_SALCD") = "D" Then
        If GetLeapYear(Year(Date)) Then
            dblSsalary# = dblNewSalary / 366
        Else
            dblSsalary# = dblNewSalary / 365
        End If
        
        'Ticket #17654 - formula correction
        dblSsalary# = ((dblNewSalary / dblHoursPerWeek#) / 52) * fglbDhrs
    End If
ElseIf snapSalary("SH_SALCD") = "M" Then
    If snapSalary("JB_SALCD") = "A" Then
        dblSsalary# = dblNewSalary * 12
    ElseIf snapSalary("JB_SALCD") = "M" Then
        dblSsalary# = dblNewSalary
    ElseIf snapSalary("JB_SALCD") = "D" Then
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
ElseIf snapSalary("SH_SALCD") = "D" Then
    If snapSalary("JB_SALCD") = "H" Then
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
        
    ElseIf snapSalary("JB_SALCD") = "M" Then
        If GetLeapYear(Year(Date)) Then
            dblSsalary# = (dblNewSalary * 366) / 12
        Else
            dblSsalary# = (dblNewSalary * 365) / 12
        End If
        
        'Ticket #17654 - formula correction
        dblSsalary# = ((dblNewSalary / fglbDhrs) * dblHoursPerWeek# * 52) / 12
        
    ElseIf snapSalary("JB_SALCD") = "A" Then
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
'If snapSalary("JB_MIDPOINT") >= 1 And snapSalary("JB_MIDPOINT") <= 11 Then
'If snapSalary("JB_MIDPOINT") >= 1 And snapSalary("JB_MIDPOINT") <= 15 Then
If snapSalary("JB_MIDPOINT") >= 1 And snapSalary("JB_MIDPOINT") <= 20 Then
    If glbCompSerial = "S/N - 2378W" And snapSalary("SH_SALCD") <> snapSalary("JB_SALCD") And snapSalary("SH_SALCD") <> "M" Then  'Town of Aurora
        If Not IsNull(snapSalary("JB_S" & snapSalary("JB_MIDPOINT") & "A")) Then
            Jb_No = snapSalary("JB_S" & snapSalary("JB_MIDPOINT") & "A")
        End If
    Else
        If Not IsNull(snapSalary("JB_S" & snapSalary("JB_MIDPOINT"))) Then
            Jb_No = snapSalary("JB_S" & snapSalary("JB_MIDPOINT"))
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
    If glbCompSerial = "S/N - 2378W" And snapSalary("SH_SALCD") <> snapSalary("JB_SALCD") And snapSalary("SH_SALCD") <> "M" Then   'Town of Aurora
        If IsNumeric(snapSalary("JB_S" & CStr(X%) & "A")) Then
            If dblNewSalary >= snapSalary("JB_S" & CStr(X%) & "A") And snapSalary("JB_S" & CStr(X%) & "A") > 0 Then
                cX$ = CStr(X)
                If X% <= 9 Then cX$ = "0" & cX$
                fglbGRADE$ = cX$
            End If
        End If
    Else
    If IsNumeric(snapSalary("JB_S" & CStr(X%))) Then
        If dblSsalary# >= snapSalary("JB_S" & CStr(X%)) And snapSalary("JB_S" & CStr(X%)) > 0 Then
            cX$ = CStr(X)
            If X% <= 9 Then cX$ = "0" & cX$
            fglbGRADE$ = cX$
        End If
    End If
    End If
Next X%

If glbCompSerial = "S/N - 2378W" And snapSalary("SH_SALCD") <> snapSalary("JB_SALCD") And snapSalary("SH_SALCD") <> "M" Then   'Town of Aurora
    If IsNumeric(snapSalary("JB_S1A")) Then
        If dblSsalary# < snapSalary("JB_S1A") Then
            fglbGRADE$ = "00"
        End If
    End If
Else
If IsNumeric(snapSalary("JB_S1")) Then
    If dblSsalary# < snapSalary("JB_S1") Then
        fglbGRADE$ = "00"
    End If
End If
End If
End Sub

Private Function modUpdateSelection()
Dim lngLastCurrentID&, X%
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim JobInfo As Boolean
Dim prec%, curSalary#
Dim fTablSalHis  As New ADODB.Recordset
Dim dblPct, dblAmnt, dblOSalary, dblOIndSalary, dblOSalary1, dblOSalary2, dblOSalary3, dblOSalary4, DtTm As Variant, dblNewSalary As Double
Dim OEDate, OEDate1
Dim salarystep, xGradeF
Dim JobSalCD, xWHRS
Dim SalDetailCode1, SalDetailCode2, SalDetailCode3, SalDetailCode4 As String
Dim EDate1, EDate2, EDate3, EDate4
Dim xSHID 'George Mar 7,2006 #9965
Dim SQLQ
Dim xAsofDateSal As Double

On Error GoTo modUpdateSelection_Err

modUpdateSelection = False

Call DelNoUpsal

Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 5

SkipRec = 0

Do While Not snapSalary.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / (lngRecs&)))
    MDIMain.panHelp(0).FloodPercent = pct%
    
    EmpNo& = snapSalary("SH_EMPNBR")
    If glbCompSerial = "S/N - 2288W" Then 'MUSASHI AUTO TKT#10845
        If Not IsNull(snapSalary("SH_SALARY1")) Then
            dblOSalary = snapSalary("SH_SALARY1")
            OSalary = snapSalary("SH_SALARY1")
        Else
            dblOSalary = snapSalary("SH_SALARY")
            OSalary = snapSalary("SH_SALARY")
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
        
        If Not IsNull(snapSalary("SH_SALARY2")) And snapSalary("SH_SALARY2") <> 0 Then
            dblOSalary2 = Val(snapSalary("SH_SALARY2"))
        End If
        If Not IsNull(snapSalary("SH_SALARY3")) And snapSalary("SH_SALARY3") <> 0 Then
            dblOSalary3 = Val(snapSalary("SH_SALARY3"))
        End If
        If Not IsNull(snapSalary("SH_SALARY4")) And snapSalary("SH_SALARY4") <> 0 Then
            dblOSalary4 = Val(snapSalary("SH_SALARY4"))
        End If
        
        If Not IsNull(snapSalary("SH_DETAILCODE1")) Then
            SalDetailCode1 = snapSalary("SH_DETAILCODE1")
        End If
        If Not IsNull(snapSalary("SH_DETAILCODE2")) Then
            SalDetailCode2 = snapSalary("SH_DETAILCODE2")
        End If
        If Not IsNull(snapSalary("SH_DETAILCODE3")) Then
            SalDetailCode3 = snapSalary("SH_DETAILCODE3")
        End If
        If Not IsNull(snapSalary("SH_DETAILCODE4")) Then
            SalDetailCode4 = snapSalary("SH_DETAILCODE4")
        End If
        
        If IsDate(snapSalary("SH_EDATE1")) Then
            EDate1 = snapSalary("SH_EDATE1")
        End If
        If IsDate(snapSalary("SH_EDATE2")) Then
            EDate2 = snapSalary("SH_EDATE2")
        End If
        If IsDate(snapSalary("SH_EDATE3")) Then
            EDate3 = snapSalary("SH_EDATE3")
        End If
        If IsDate(snapSalary("SH_EDATE4")) Then
            EDate4 = snapSalary("SH_EDATE4")
        End If
    Else
        'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
        If glbCompSerial = "S/N - 2494W" Then
            'Direct Salary
            dblOSalary = snapSalary("SH_SALARY")
            OSalary = snapSalary("SH_SALARY")
            
            'Indirect Salary
            If Not IsNull(snapSalary("SH_SALARY1")) Then
                dblOIndSalary = snapSalary("SH_SALARY1")
                OIndSalary = snapSalary("SH_SALARY1")
            Else
                dblOIndSalary = 0
                OIndSalary = 0
            End If
            
            'Indirect Salary Effective Date 1
            If IsDate(snapSalary("SH_EDATE1")) Then
                OEDate1 = snapSalary("SH_EDATE1")
            Else
                OEDate1 = ""
            End If
        Else
            dblOSalary = snapSalary("SH_SALARY")
            OSalary = snapSalary("SH_SALARY")
        End If
        OTOTAL = snapSalary("SH_TOTAL")
    End If
    
    oPayP = snapSalary("SH_PAYP")      'laura jan 28, 1998
    OEDate = snapSalary("SH_EDATE")

    OJOB1 = snapSalary("SH_JOB")
    oGrid = snapSalary("SH_GRID")
    oPayrollID = snapSalary("SH_PAYROLL_ID")
    OSalCD = snapSalary("SH_SALCD")
    oGrade = snapSalary("SH_GRADE")
    xWHRS = snapSalary("SH_WHRS")
    
    If IsNull(snapSalary("SH_GRADE")) Then xGrade = 0 Else xGrade = Val(snapSalary("SH_GRADE"))
    
    'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
    If glbCompSerial = "S/N - 2494W" Then
        If optDirIndSal(1).Value = True Then
            'Do not recompute the xGrade
        Else
            Call GetGrade(dblOSalary, xGrade, OJOB1, OSalCD, xWHRS)
        End If
    Else
        Call GetGrade(dblOSalary, xGrade, OJOB1, OSalCD, xWHRS)
    End If
    
    xGradeF = xGrade

    'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
    If glbCompSerial = "S/N - 2494W" Then
        If optDirIndSal(1).Value = True Then
            'Check against Indirect Salary Effective Date 1
            'Ticket #18668 and Ticket #19154 - Allow same salary effective date update since we are allowing manual
            'update on the salary screen. So changed from >= to >.
            If OEDate1 > CVDate(dlpEffectiveDate.Text) Then
                Call NoUpSal_Addnew("D")
                GoTo lblNextRec
            End If
        Else
            'Check against Direct Salary Effective Date
            'Ticket #18668 and Ticket #19154 - Allow same salary effective date update since we are allowing manual
            'update on the salary screen. So changed from >= to >.
            If OEDate > CVDate(dlpEffectiveDate.Text) Then
                Call NoUpSal_Addnew("D")
                GoTo lblNextRec
            End If
        End If
    Else
        'Ticket #18668 and Ticket #19154 - Allow same salary effective date update since we are allowing manual
        'update on the salary screen. So changed from >= to >.
        If OEDate > CVDate(dlpEffectiveDate.Text) Then
            Call NoUpSal_Addnew("D")
            GoTo lblNextRec
        End If
    End If
    
    ONDate = snapSalary("SH_NEXTDAT")
    
    If IsNull(snapSalary("SH_WHRS")) Or Len(snapSalary("SH_WHRS")) < 1 Then
        dblWHours# = 0
    Else
        dblWHours# = snapSalary("SH_WHRS")
    End If
    
    lngLastCurrentID& = snapSalary("SH_ID")
    
    DtTm = Now
    JobInfo = False
    
    If optStep Then
        For X% = 4 To 6
            If Len(clpCode(X%).Text) > 0 Then
                If comStep(X%) = "Next Step" Then
                    xGrade = xGrade + 1
                    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
                    'If xGrade > 11 Then
                    'If xGrade > 15 Then
                    If xGrade > 20 Then
                        JobInfo = True
                    Else
                        salarystep = snapSalary("JB_S" & Format(xGrade, "##"))
                        JobSalCD = snapSalary("JB_SALCD")
                        If glbCompSerial = "S/N - 2378W" And OSalCD <> JobSalCD And OSalCD <> "M" Then  'Town of Aurora
                            salarystep = snapSalary("JB_S" & Format(xGrade, "##") & "A")
                        End If
                        If OSalCD <> JobSalCD And xWHRS = 0 Then JobInfo = True
                        If IsNull(salarystep) Then JobInfo = True
                        If salarystep = 0 Then JobInfo = True
                    End If
                Else
                    salarystep = snapSalary("JB_S" & Format(comStep(X%), "##"))
                    If glbCompSerial = "S/N - 2378W" And OSalCD <> JobSalCD And OSalCD <> "M" Then   'Town of Aurora
                        salarystep = snapSalary("JB_S" & Format(comStep(X%), "##") & "A")
                    End If
                    If IsNull(salarystep) Then JobInfo = True
                    If salarystep = 0 Then JobInfo = True
                End If
            End If
        Next
    End If
    
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
    fglbDhrs = GetJHData(snapSalary("SH_EMPNBR"), "JH_DHRS", 0)
    
    fTablSalHis.AddNew
    fTablSalHis("SH_COMPNO") = snapSalary("SH_COMPNO")
    fTablSalHis("SH_EMPNBR") = snapSalary("SH_EMPNBR")
    
    'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
    If glbCompSerial = "S/N - 2494W" Then
        If optDirIndSal(1).Value = True Then
            'Update with New Indirect Salary Effective and Old Direct Salary Effective Date
            fTablSalHis("SH_EDATE1") = CVDate(dlpEffectiveDate.Text)
            fTablSalHis("SH_EDATE") = CVDate(OEDate)
        Else
            'Updat with New Direct Salary Effective Date nad Old Indirect Salary Effective Date
            fTablSalHis("SH_EDATE") = CVDate(dlpEffectiveDate.Text)
            If IsDate(OEDate1) Then
                fTablSalHis("SH_EDATE1") = CVDate(OEDate1)
            End If
        End If
    Else
        fTablSalHis("SH_EDATE") = CVDate(dlpEffectiveDate.Text)
    End If
    
    fTablSalHis("SH_CURRENT") = True
    fTablSalHis("SH_SDATE") = snapSalary("SH_SDATE")
    fTablSalHis("SH_SALCD") = txtSALCD
    fTablSalHis("SH_PAYROLL_ID") = snapSalary("SH_PAYROLL_ID")
    fTablSalHis("SH_WHRS") = snapSalary("SH_WHRS")

    fTablSalHis("SH_TRANSDATE") = Date
    
    'Ticket #20466 Franks 06/22/2011
    fTablSalHis("SH_VGROUP") = snapSalary("SH_VGROUP")
    fTablSalHis("SH_VSTEP") = snapSalary("SH_VSTEP")
    
    fTablSalHis("SH_PAYP") = snapSalary("SH_PAYP")
    fTablSalHis("SH_PAYP_TABLE") = snapSalary("SH_PAYP_TABLE")
    fTablSalHis("SH_SREAS_TABLE") = snapSalary("SH_SREAS_TABLE")
    
    'For New Salary computation, start with assignment of New Salary field with Old Salary
    'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
    If glbCompSerial = "S/N - 2494W" Then
        If optDirIndSal(1).Value = True Then
            'Indirect Salary
            dblNewSalary = dblOIndSalary
        Else
            'Direct Salary
            dblNewSalary = dblOSalary
        End If
    Else
        dblNewSalary = dblOSalary
    End If
    
    'Ticket #22893 - WHSC - As of Date of the Salary to apply the Percent update on
    If glbCompSerial = "S/N - 2448W" And optPct And Not glbMulti Then
        If IsDate(dlpAsofDate.Text) Then
            'Get Salary As of Date.
            xAsofDateSal = Get_AsOfDate_Salary(snapSalary("SH_EMPNBR"), dlpAsofDate.Text, 0)
            
            'If no salary is retrieved then used the last current salary
            If xAsofDateSal = 0 Then
                xAsofDateSal = dblOSalary
            End If
        End If
    End If
    
    'Computation of New Salary, New Amount Change and New Percentage change
    For X% = 4 To 6
        If Len(clpCode(X%).Text) > 0 Then
            If optStep Then
                JobSalCD = snapSalary("JB_SALCD")
                If comStep(X%) = "Next Step" Then
                    xGradeF = xGradeF + 1
                    If OSalCD = JobSalCD Then
                        dblAmnt = snapSalary("JB_S" & Format(xGradeF, "##")) - dblNewSalary
                        dblPct = dblAmnt / dblOSalary
                    Else
                        If OSalCD = "H" And JobSalCD = "A" Then
                            dblAmnt = snapSalary("JB_S" & Format(xGradeF, "##")) / (xWHRS * 52) - dblNewSalary
                            dblPct = dblAmnt / dblOSalary
                        End If
                        If OSalCD = "A" And JobSalCD = "H" Then
                            dblAmnt = snapSalary("JB_S" & Format(xGradeF, "##")) * (xWHRS * 52) - dblNewSalary
                            dblPct = dblAmnt / dblOSalary
                        End If
                        If OSalCD = "D" And JobSalCD = "H" Then
                            dblAmnt = snapSalary("JB_S" & Format(xGradeF, "##")) * fglbDhrs - dblNewSalary
                            dblPct = dblAmnt / dblOSalary
                        End If
                        If OSalCD = "D" And JobSalCD = "A" Then
                            dblAmnt = (snapSalary("JB_S" & Format(xGradeF, "##")) / (xWHRS * 52) * fglbDhrs) - dblNewSalary
                            dblPct = dblAmnt / dblOSalary
                        End If
                        If glbCompSerial = "S/N - 2378W" And OSalCD <> "M" Then   'Town of Aurora
                            dblAmnt = snapSalary("JB_S" & Format(xGradeF, "##") & "A") - dblNewSalary
                            dblPct = dblAmnt / dblOSalary
                        End If
                    End If
                Else
                    'Hemu - testing
                    If txtSALCD <> JobSalCD Then
                        If txtSALCD = "A" And JobSalCD = "H" Then
                            dblAmnt = ((snapSalary("JB_S" & Format(comStep(X%), "##")) * snapSalary("SH_WHRS")) * 52) - dblNewSalary
                        ElseIf txtSALCD = "H" And JobSalCD = "A" Then
                            dblAmnt = ((snapSalary("JB_S" & Format(comStep(X%), "##")) / snapSalary("SH_WHRS")) / 52) - dblNewSalary
                        ElseIf txtSALCD = "D" And JobSalCD = "H" Then
                            dblAmnt = ((snapSalary("JB_S" & Format(comStep(X%), "##")) * fglbDhrs)) - dblNewSalary
                        ElseIf txtSALCD = "D" And JobSalCD = "A" Then
                            dblAmnt = (((snapSalary("JB_S" & Format(comStep(X%), "##")) / snapSalary("SH_WHRS")) / 52) * fglbDhrs) - dblNewSalary
                        End If
                    ElseIf txtSALCD = JobSalCD Then
                        dblAmnt = snapSalary("JB_S" & Format(comStep(X%), "##")) - dblNewSalary
                    End If
                    dblPct = dblAmnt / dblOSalary
                    If glbCompSerial = "S/N - 2378W" And OSalCD <> JobSalCD And OSalCD <> "M" Then   'Town of Aurora
                        dblAmnt = snapSalary("JB_S" & Format(comStep(X%), "##") & "A") - dblNewSalary
                        dblPct = dblAmnt / dblOSalary
                    End If
                End If
                dblNewSalary = dblNewSalary + dblAmnt
            ElseIf optDollars Then
                dblAmnt = medAmountChng(X%)
                'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
                If glbCompSerial = "S/N - 2494W" Then
                    If optDirIndSal(1).Value = True Then
                        'Use Indirect Old Salary
                        dblPct = dblAmnt / dblOIndSalary
                    Else
                        'Use Direct Old Salary
                        dblPct = dblAmnt / dblOSalary
                    End If
                Else
                    dblPct = dblAmnt / dblOSalary
                End If
                dblNewSalary = dblNewSalary + dblAmnt
            ElseIf optPct Then
                'Ticket #22893 - WHSC - As of Date of the Salary to apply the Percent update on
                If glbCompSerial = "S/N - 2448W" And Not glbMulti Then
                    dblPct = medAmountChng(X%)
                    dblAmnt = dblPct * xAsofDateSal
                    dblNewSalary = dblNewSalary + dblAmnt
                    dblPct = (dblAmnt / dblOSalary) '* 100
                Else
                    dblPct = medAmountChng(X%)
                    'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
                    If glbCompSerial = "S/N - 2494W" Then
                        If optDirIndSal(1).Value = True Then
                            'Use Indirect Old Salary
                            dblAmnt = dblPct * dblOIndSalary
                        Else
                            'Use Direct Old Salary
                            dblAmnt = dblPct * dblOSalary
                        End If
                    Else
                        dblAmnt = dblPct * dblOSalary
                    End If
                    dblNewSalary = dblNewSalary + dblAmnt
                End If
            ElseIf optFixed Then
                dblNewSalary = medAmountChng(X%)
                'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
                If glbCompSerial = "S/N - 2494W" Then
                    If optDirIndSal(1).Value = True Then
                        'Do not really need to compute the Amount and Percentage Changed by
                        dblAmnt = dblNewSalary - dblOIndSalary
                        dblPct = dblAmnt / dblOIndSalary
                    Else
                        'Direct Salary
                        dblAmnt = dblNewSalary - dblOSalary
                        dblPct = dblAmnt / dblOSalary
                    End If
                Else
                    dblAmnt = dblNewSalary - dblOSalary
                    dblPct = dblAmnt / dblOSalary
                End If
            End If
            fTablSalHis("SH_SREAS" & CStr(X% - 3)) = clpCode(X%).Text
            'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
            If glbCompSerial = "S/N - 2494W" Then
                If optDirIndSal(1).Value = True Then
                    'For Indirect Salary - Do not update with new Amount and Percentage Changed By
                    fTablSalHis("SH_SALPC" & CStr(X% - 3)) = snapSalary("SH_SALPC" & CStr(X% - 3))
                    fTablSalHis("SH_SALCHG" & CStr(X% - 3)) = snapSalary("SH_SALCHG" & CStr(X% - 3))
                Else
                    'For Direct Salary update Amount and Percentage Changed By
                    fTablSalHis("SH_SALPC" & CStr(X% - 3)) = dblPct
                    fTablSalHis("SH_SALCHG" & CStr(X% - 3)) = dblAmnt
                End If
            Else
                fTablSalHis("SH_SALPC" & CStr(X% - 3)) = dblPct
                fTablSalHis("SH_SALCHG" & CStr(X% - 3)) = dblAmnt
            End If
        End If
    Next X%
    dblNewSalary = Round2DEC(dblNewSalary, snapSalary("SH_EMPNBR")) 'added by raubrey 8/18/97
    
    If glbCompSerial = "S/N - 2288W" Then 'MUSASHI AUTO TKT#10845
        dblOSalary1 = dblNewSalary + dblOSalary2 + dblOSalary3 + dblOSalary4
    End If
    
    'New Salary Update
    Dim RoundSal As Double, strRoundSal As String
    Dim strFirst As String
    If cmbRound.ListIndex = 0 Then
        'dblNewSalary = CLng(dblNewSalary)  'Ticket #14699
        dblNewSalary = Round(dblNewSalary, cmbPrecision.Text)   'Ticket #14699
    Else
        dblNewSalary = dblNewSalary
    End If
    If glbCompSerial = "S/N - 2288W" Then 'MUSASHI AUTO TKT#10845
        fTablSalHis("SH_SALARY1") = dblNewSalary
        fTablSalHis("SH_SALARY") = dblOSalary1
    Else
        'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
        If glbCompSerial = "S/N - 2494W" Then
            If optDirIndSal(1).Value = True Then
                'Indirect Salary changed so update the right field and then update Direct Salary with Previous record's Direct Salary
                fTablSalHis("SH_SALARY1") = dblNewSalary
                fTablSalHis("SH_SALARY") = snapSalary("SH_SALARY")
            Else
                'Direct Salary changed so update the right field adn then update the Indirect Salary with Previous record's Indirect Salary
                fTablSalHis("SH_SALARY") = dblNewSalary
                fTablSalHis("SH_SALARY1") = snapSalary("SH_SALARY1")
            End If
        Else
            fTablSalHis("SH_SALARY") = dblNewSalary
        End If
    End If
    
    If Len(dlpNextReviewDate.Text) > 0 Then
        fTablSalHis("SH_NEXTDAT") = CVDate(dlpNextReviewDate.Text)
        UpdateFollowup snapSalary("SH_EMPNBR"), snapSalary("SH_NEXTDAT"), CVDate(dlpNextReviewDate.Text), "SREV"
    Else
        'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
        If glbCompSerial = "S/N - 2494W" Then
            If optDirIndSal(1).Value = True Then
                'For Indirect Salary update with the Previous Salary record's Next Review Date
                fTablSalHis("SH_NEXTDAT") = snapSalary("SH_NEXTDAT")
            Else
                'For Direct Salary update as needed
                If IsDate(snapSalary("SH_NEXTDAT")) Then
                    If snapSalary("SH_NEXTDAT") > CVDate(dlpEffectiveDate) Then
                        fTablSalHis("SH_NEXTDAT") = snapSalary("SH_NEXTDAT")
                    End If
                End If
            End If
        Else
            If IsDate(snapSalary("SH_NEXTDAT")) Then
                If snapSalary("SH_NEXTDAT") > CVDate(dlpEffectiveDate) Then
                    fTablSalHis("SH_NEXTDAT") = snapSalary("SH_NEXTDAT")
                End If
            End If
        End If
    End If
    fTablSalHis("SH_JOB") = snapSalary("SH_JOB")
    fTablSalHis("SH_GRID") = snapSalary("SH_GRID")
    fTablSalHis("SH_PAYROLL_ID") = snapSalary("SH_PAYROLL_ID")
    fTablSalHis("SH_JOB_ID") = snapSalary("SH_JOB_ID")
    
    'Jaddy changed for WFC Kipling Oct 31, 02
    If glbWFC And (snapSalary("JB_ORG") = "NONE" Or snapSalary("JB_ORG") = "EXEC") Then
        fTablSalHis("SH_COMPA_USER") = snapSalary("SH_COMPA_USER")
        fTablSalHis("SH_COMPA_DOLLAR") = snapSalary("SH_COMPA_DOLLAR")
        fTablSalHis("SH_BAND") = snapSalary("SH_BAND")
        fTablSalHis("SH_MARKETLINE") = snapSalary("SH_MARKETLINE")
        Call Set_WFC_COMPA(dblNewSalary)
    Else
        'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
        If glbCompSerial = "S/N - 2494W" Then
            If optDirIndSal(1).Value = True Then
                'Do not recompute Compa Ratio and Salary Grade
            Else
                Call modSetCOMPA_GRADE(dblNewSalary) ' sets fglbCOMPA#, and fglbGRADE
            End If
        Else
            Call modSetCOMPA_GRADE(dblNewSalary) ' sets fglbCOMPA#, and fglbGRADE
        End If
    End If
    
    If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
        fTablSalHis("SH_PREMIUM") = snapSalary("SH_PREMIUM")
        fTablSalHis("SH_TOTAL") = fTablSalHis("SH_SALARY") + snapSalary("SH_PREMIUM") 'snapSalary("SH_TOTAL")
        fTablSalHis("SH_VGROUP") = snapSalary("SH_VGROUP")
        fTablSalHis("SH_VSTEP") = snapSalary("SH_VSTEP")
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
        
        fTablSalHis("SH_EDATE1") = CVDate(dlpEffectiveDate.Text)
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
    
    'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
    If glbCompSerial = "S/N - 2494W" Then
        If optDirIndSal(1).Value = True Then
            'For Indirect Salary, same as Previous Salary's Compa Ratio and Grade
            fTablSalHis("SH_COMPA") = snapSalary("SH_COMPA")
            fTablSalHis("SH_GRADE") = snapSalary("SH_GRADE")
        Else
            'For Direct Salary New Compa Ratio and Salary Grade
            fTablSalHis("SH_COMPA") = fglbCOMPA#
            fTablSalHis("SH_GRADE") = Format(fglbGRADE$, "00")
        End If
    Else
        fTablSalHis("SH_COMPA") = fglbCOMPA#
        fTablSalHis("SH_GRADE") = Format(fglbGRADE$, "00")
    End If
    
    fTablSalHis("SH_LDATE") = Now
    fTablSalHis("SH_LTIME") = Time$
    fTablSalHis("SH_LUSER") = glbUserID 'glbLEE_ID

    If glbCompSerial = "S/N - 2214W" Then
        Dim xToDate
        If IsDate(dlpNextReviewDate.Text) Then
            xToDate = dlpNextReviewDate.Text
        Else
            xToDate = DateAdd("D", -1, DateAdd("YYYY", 1, CVDate(dlpEffectiveDate.Text)))
        End If
        Call ChangeOtherEarnAmount(fTablSalHis("SH_EMPNBR"), dblNewSalary, "A", dlpEffectiveDate.Text, xToDate)
    End If
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
    
    'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
    If glbCompSerial = "S/N - 2494W" Then
        If optDirIndSal(1).Value = True Then
            'Do not update Benefit for Indirect Salary Update
        Else
            'Direct Salary Update
            Call updBenefitForSalDEPN(EmpNo&)   'jaddy 9/10/99
        End If
    Else
        Call updBenefitForSalDEPN(EmpNo&)   'jaddy 9/10/99
    End If
    
    'Call Employee_Master_Integration(EmpNo&)
    If glbAdv Or glbMediPay Then  'Ticket #12339 'Ticket #16574 for St. Johns Medipay
        Call Employee_Master_Integration(EmpNo&)
    End If
    
    If glbGP Then
        Call Salary_Integration(EmpNo&, , False, True, xSHID)
    Else
        Call Salary_Integration(EmpNo&)
    End If
    
    If glbCompSerial = "S/N - 2288W" Then 'MUSASHI AUTO TKT#10845
        NSalary = dblOSalary1
    Else
        NSalary = dblNewSalary
    End If
    NEDate = CVDate(dlpEffectiveDate.Text)
    If Len(dlpNextReviewDate.Text) > 0 Then
        NNDate = CVDate(dlpNextReviewDate.Text)
    Else
        NNDate = ""
    End If
    
    'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
    If glbCompSerial = "S/N - 2494W" Then
        If optDirIndSal(1).Value = True Then
            'No Audit Update for Indirect Salary
        Else
            'Audit Update for Direct Salary only
            If Not AUDITSALY() Then MsgBox "ERROR - AUDIT FILE"
        End If
    Else
        If Not AUDITSALY() Then MsgBox "ERROR - AUDIT FILE"
    End If
    
lblNextRec:
    snapSalary.MoveNext
Loop

'cmdView.Enabled = True
'cmdPrint.Enabled = True

modUpdateSelection = True
MDIMain.panHelp(0).FloodType = 0

snapSalary.Close
Screen.MousePointer = DEFAULT

If gsEMAIL_ONSALARY Then
    MailBody = ""
    SQLQ = "SELECT TT_EMPNBR,TT_COEFLAG,TT_WRKEMP,TT_GRADE,TT_FLAG,TT_TBEMP FROM HREMPWRK WHERE TT_WRKEMP='" & glbUserID & "'"
    SQLQ = SQLQ & "AND TT_FLAG = 'Y'"
    snapSalary.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
    If Not snapSalary.EOF Then
        If snapSalary.RecordCount = 1 Then
            'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
            If glbCompSerial = "S/N - 2494W" Then
                If optDirIndSal(1).Value = True Then
                    'Indirect Salary change
                    MailBody = "The following employee's Indirect Salary has "
                Else
                    'Direct Salary change
                    MailBody = "The following employee's Direct Salary has "
                End If
            Else
                MailBody = "The following employee's salary has "
            End If
        Else
            'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
            If glbCompSerial = "S/N - 2494W" Then
                If optDirIndSal(1).Value = True Then
                    'Indirect Salary change
                    MailBody = "The following employees Indirect Salaries have "
                Else
                    'Direct Salary change
                    MailBody = "The following employees Direct Salaries have "
                End If
            Else
                MailBody = "The following employees salaries have "
            End If
        End If
        If optDollars Then
            MailBody = MailBody & "been increased by $" & medAmountChng(4) & "." & vbCrLf '& vbCrLf
        End If
        If optFixed Then
            MailBody = MailBody & " been changed to $" & medAmountChng(4) & "." & vbCrLf '& vbCrLf
        End If
        If optPct Then
            MailBody = MailBody & " been increased by " & medAmountChng(4) * 100 & "%." & vbCrLf '& vbCrLf
        End If
        If optStep Then
            MailBody = MailBody & " been changed to Step " & comStep(4) & "." & vbCrLf '& vbCrLf
        End If
        MailBody = MailBody & "Reason: " & GetTABLDesc("SDRC", clpCode(4)) & vbCrLf
        MailBody = MailBody & "Effective Date: " & dlpEffectiveDate & vbCrLf & vbCrLf
    End If
    Do While Not snapSalary.EOF
        MailBody = MailBody & GetEmpName(snapSalary("TT_EMPNBR")) & vbCrLf
        snapSalary.MoveNext
     Loop
     snapSalary.Close
     If Len(MailBody) > 0 Then
         Call imgEmail_Click
     End If
End If

Exit Function

modUpdateSelection_Err:
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
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail", "SMTP", "SENDEMAIL")
    Resume Next

End Sub

Private Sub optDirIndSal_Click(Index As Integer, Value As Integer)
    If optDirIndSal(1).Value = True Then
        lblNextReview.Enabled = False
        dlpNextReviewDate.Enabled = False
        'comPayType.Enabled = False
        optStep.Enabled = False
        chkUpdAttendance.Enabled = False
    Else
        lblNextReview.Enabled = True
        dlpNextReviewDate.Enabled = True
        'comPayType.Enabled = True
        optStep.Enabled = True
        chkUpdAttendance.Enabled = True
    End If

End Sub

Private Sub optDirIndSal_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optDollars_Click(Value As Integer)
Call setDollars

'Ticket #22893 - WHSC - As of Date of the Salary to apply the Percent update on
If glbCompSerial = "S/N - 2448W" Then
    'As of Date
    Call ShowHide_AsofDate
End If

End Sub

Private Sub optDollars_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub optFixed_Click(Value As Integer)
Call setDollars

'Ticket #22893 - WHSC - As of Date of the Salary to apply the Percent update on
If glbCompSerial = "S/N - 2448W" Then
    'As of Date
    Call ShowHide_AsofDate
End If
End Sub

Private Sub optFixed_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optPct_Click(Value As Integer)
Call setDollars

'Ticket #22893 - WHSC - As of Date of the Salary to apply the Percent update on
If glbCompSerial = "S/N - 2448W" Then
    'As of Date
    Call ShowHide_AsofDate
End If

End Sub

Private Sub optPct_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Function DelNoUpsal() 'frank 4/10/2000
Dim iTableCounter, TblName, SQLQ

    'Ticket #16717
    'gdbAdoIhr001.Execute "DELETE FROM HREMPWRK " & in_SQL(glbIHRDBW) & " WHERE TT_WRKEMP='" & glbUserID & "' "
    gdbAdoIhr001.Execute "DELETE FROM HREMPWRK " & in_SQL(glbIHRDBW) & " WHERE (TT_FLAG = 'Y' OR TT_FLAG = 'N') "
    
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
        If xSALCDt = snapSalary("JB_SALCD") Then
            If IsNumeric(snapSalary("JB_S" & CStr(x1%))) And snapSalary("JB_S" & CStr(x1%)) > 0 Then
                If xOLDSALA >= snapSalary("JB_S" & CStr(x1%)) Then
                  xGRADEt = x1%
                End If
            End If
        Else
            If xSALCDt = "H" And snapSalary("JB_SALCD") = "A" Then
                If xWHRSt = 0 Then
                    xGRADEt = 0
                Else
                    xSAL_TMP = snapSalary("JB_S" & Format(x1%, "##")) / (xWHRSt * 52)
                End If
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    xSAL_TMP = snapSalary("JB_S" & Format(x1%, "##") & "A")
                    xGrid2 = True
                End If
            End If
            If xSALCDt = "A" And snapSalary("JB_SALCD") = "H" Then
                If xWHRSt = 0 Then
                    xGRADEt = 0
                Else
                    xSAL_TMP = snapSalary("JB_S" & Format(x1%, "##")) * (xWHRSt * 52)
                End If
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    xSAL_TMP = snapSalary("JB_S" & Format(x1%, "##") & "A")
                    xGrid2 = True
                End If
            End If
            
            If xSALCDt = "D" And snapSalary("JB_SALCD") = "A" Then
                If xWHRSt = 0 Then
                    xGRADEt = 0
                Else
                    xSAL_TMP = snapSalary("JB_S" & Format(x1%, "##")) / (xWHRSt * 52) * fglbDhrs
                End If
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    xSAL_TMP = snapSalary("JB_S" & Format(x1%, "##") & "A")
                    xGrid2 = True
                End If
            End If
            If xSALCDt = "D" And snapSalary("JB_SALCD") = "H" Then
                If xWHRSt = 0 Then
                    xGRADEt = 0
                Else
                    xSAL_TMP = snapSalary("JB_S" & Format(x1%, "##")) * fglbDhrs
                End If
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    xSAL_TMP = snapSalary("JB_S" & Format(x1%, "##") & "A")
                    xGrid2 = True
                End If
            End If
            
            
            If glbCompSerial = "S/N - 2378W" And xGrid2 = True Then 'Town of Aurora
                If xOLDSALA >= xSAL_TMP And snapSalary("JB_S" & CStr(x1%) & "A") > 0 Then
                  xGRADEt = x1%
                End If
            Else
            If xOLDSALA >= xSAL_TMP And snapSalary("JB_S" & CStr(x1%)) > 0 Then
              xGRADEt = x1%
                End If
            End If
        End If
    Next x1%
End Function

Private Sub NoUpSal_Addnew(zReason) 'Frank 4/13/2000
    Dim iFieldCounter, FldName
    snapNoUpSal.AddNew
    snapNoUpSal("TT_EMPNBR") = snapSalary("SH_EMPNBR")
    snapNoUpSal("TT_COEFLAG") = snapSalary("SH_CURRENT")
    snapNoUpSal("TT_WRKEMP") = glbUserID
    snapNoUpSal("TT_GRADE") = zReason
    snapNoUpSal("TT_FLAG") = "N"
    snapNoUpSal.Update
    SkipRec = SkipRec + 1
End Sub

Private Sub UpSal_Addnew() 'Frank 4/12/2000
    Dim iFieldCounter, FldName
    snapNoUpSal.AddNew
    snapNoUpSal("TT_EMPNBR") = snapSalary("SH_EMPNBR")
    snapNoUpSal("TT_COEFLAG") = snapSalary("SH_CURRENT")
    snapNoUpSal("TT_WRKEMP") = glbUserID
    snapNoUpSal("TT_FLAG") = "Y"
    snapNoUpSal.Update
    If Len(strEMPLIST) > 0 Then
        strEMPLIST = strEMPLIST & "," & snapSalary("SH_EMPNBR")
    Else
        strEMPLIST = snapSalary("SH_EMPNBR")
    End If
End Sub

Private Sub DelSal_Addnew(xID)
    snapNoUpSal.AddNew
    snapNoUpSal("TT_EMPNBR") = snapSalary("SH_EMPNBR")
    snapNoUpSal("TT_WRKEMP") = glbUserID
    snapNoUpSal("TT_TBEMP") = xID
    snapNoUpSal("TT_FLAG") = "D"
    snapNoUpSal.Update
    If Len(strEMPLIST) > 0 Then
        strEMPLIST = strEMPLIST & "," & snapSalary("SH_EMPNBR")
    Else
        strEMPLIST = snapSalary("SH_EMPNBR")
    End If
End Sub

Private Function Cri_SetAll()
Dim X%, strRName$, strRName1$

Cri_SetAll = False

On Error GoTo modSetCriteria_Err

Call glbCri_DeptUN("")

Screen.MousePointer = HOURGLASS

strRName$ = glbIHRREPORTS & "rzsalnu.rpt"
Me.vbxCrystal.ReportFileName = strRName$
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    Me.vbxCrystal.DataFiles(0) = glbIHRDB
    Me.vbxCrystal.DataFiles(1) = glbIHRDB
    Me.vbxCrystal.DataFiles(2) = glbIHRDB
    Me.vbxCrystal.DataFiles(3) = glbIHRDB
    Me.vbxCrystal.DataFiles(4) = glbIHRDBW
    Me.vbxCrystal.DataFiles(5) = glbIHRDB
    
    ' set security for database
    'Me.vbxCrystal.Password = gstrAccPWord$
    'Me.vbxCrystal.UserName = gstrAccUID$
End If

'Ticket #25909
'Me.vbxCrystal.SelectionFormula = "{HREMPWRK.TT_FLAG} = 'N' AND " & glbstrSelCri
Me.vbxCrystal.SelectionFormula = "{HREMPWRK.TT_FLAG} = 'N' AND {HREMPWRK.TT_WRKEMP} = '" & glbUserID & "'" 'AND " & glbstrSelCri

' window title if appropriate
Me.vbxCrystal.WindowTitle = "Employees Skipped Report"

'------------
strRName1$ = glbIHRREPORTS & "rzsalup.rpt"
Me.vbxCrystal1.ReportFileName = strRName1$
If glbSQL Or glbOracle Then
    Me.vbxCrystal1.Connect = RptODBC_SQL    'Ticket #25909
Else
    Me.vbxCrystal1.Connect = "PWD=petman;"
    Me.vbxCrystal1.DataFiles(0) = glbIHRDB
    Me.vbxCrystal1.DataFiles(1) = glbIHRDB
    Me.vbxCrystal1.DataFiles(2) = glbIHRDB
    Me.vbxCrystal1.DataFiles(3) = glbIHRDB
    Me.vbxCrystal1.DataFiles(4) = glbIHRDBW
    Me.vbxCrystal1.DataFiles(5) = glbIHRDB
    
    ' set security for database
'    Me.vbxCrystal1.Password = gstrAccPWord$
'    Me.vbxCrystal1.UserName = gstrAccUID$
End If

'Ticket #25909
'Me.vbxCrystal.SelectionFormula = "{HREMPWRK.TT_FLAG} = 'Y' AND " & glbstrSelCri
Me.vbxCrystal1.SelectionFormula = "{HREMPWRK.TT_FLAG} = 'Y' AND {HREMPWRK.TT_WRKEMP} = '" & glbUserID & "'" 'AND " & glbstrSelCri

' window title if appropriate
Me.vbxCrystal1.WindowTitle = "Employees Updated Report"

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

Private Sub optStep_Click(Value As Integer)
Call setDollars

'Ticket #22893 - WHSC - As of Date of the Salary to apply the Percent update on
If glbCompSerial = "S/N - 2448W" Then
    'As of Date
    Call ShowHide_AsofDate
End If
End Sub

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
cmbRound.AddItem "Yes"
cmbRound.AddItem "No"
cmbRound.ListIndex = 1

End Sub



Private Sub setDollars()
Dim strFmat$, X%
If optDollars Or optFixed Then
    strFmat$ = "$#,##0.00;($#,##0.00)"
    If glbCompDecHR = 3 Then strFmat$ = "$#,##0.000;($#,##0.000)"
    If glbCompDecHR = 4 Then strFmat$ = "$#,##0.0000;($#,##0.0000)"
Else
    strFmat$ = "0.######%"    'Jaddy 11/15
End If
For X% = 4 To 6
    medAmountChng(X%) = ""
    medAmountChng(X%).Format = strFmat$
    If optFixed And X% > 4 Then
        medAmountChng(X%).Visible = False
        comStep(X%).Visible = False
        clpCode(X%) = ""
        clpCode(X%).Visible = False
    Else
        medAmountChng(X%).Visible = Not optStep
        comStep(X%).Visible = optStep
        clpCode(X%).Visible = True
    End If
Next X%
End Sub

Sub Set_WFC_COMPA(dblNewSalary)
Dim xDollear
fglbCOMPA# = 0
If snapSalary("SH_COMPA_USER") = "U" Then xDollear = snapSalary("SH_COMPA_DOLLAR") Else xDollear = wfcSalState
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
SQLQ = SQLQ & " WHERE [BAND]='" & snapSalary("SH_BAND") & "'"
SQLQ = SQLQ & " AND [MARKETLINE]='" & snapSalary("SH_MARKETLINE") & "'"
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
UpdateRight = GetMassUpdateSecurities("Salary_MassUpdate", glbUserID)
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = True
End Property

Public Property Get Deleteble() As Boolean
Deleteble = GetMassUpdateSecurities("Salary_MassUpdate", glbUserID) 'False
End Property

Public Property Get Printable() As Boolean
Printable = True 'False
End Property

Private Sub Transfer_Salary(rsNew As ADODB.Recordset)
    Dim rsEmpJob As New ADODB.Recordset
    Dim rsSal As New ADODB.Recordset
    Dim HRChanges As New Collection
    Dim UptSalaryDate As Date
    Dim HRSalary As New Collection
    Dim xEmpNbr
    Dim xPayrollID
    Dim xPHrs
    Dim xWHRS, xNiagaraWHRS
    Dim xEDate
    Dim xSalCD
    Dim UpdateAudit
    
    xEmpNbr = rsNew("SH_EMPNBR")
    If rsNew("SH_PAYROLL_ID") = "" Or IsNull(rsNew("SH_PAYROLL_ID")) Then
        xPayrollID = GetEmpData(rsNew("SH_EMPNBR"), "ED_PAYROLL_ID")
    Else
        xPayrollID = rsNew("SH_PAYROLL_ID")
    End If
    xEDate = rsNew("SH_EDATE")
    
    rsEmpJob.Open "SELECT JH_ID,JH_JOB,JH_DHRS,JH_PHRS,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & xEmpNbr & " AND JH_PAYROLL_ID='" & xPayrollID & "'", gdbAdoIhr001, adOpenForwardOnly
    xPHrs = 0
    xWHRS = 0
    If Not rsEmpJob.EOF Then
        xPHrs = Val(rsEmpJob("JH_PHRS") & "")
        xWHRS = Val(rsEmpJob("JH_WHRS") & "") 'Hemu - it was asssigning JH_DHRS - it should pass Weekly Hours
        xNiagaraWHRS = Val(rsEmpJob("JH_WHRS") & "")
        
        'City of Niagara Falls  = Dhrs = Hours Per Days from Position Master, fglbNiagPhrs = Pay Period
        If glbCompSerial = "S/N - 2276W" Then
            rsSal.Open "SELECT SH_EMPNBR, SH_PAYP, SH_WHRS FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & xEmpNbr & " AND SH_PAYROLL_ID = '" & xPayrollID & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
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
        If isChanged_Salary(HRSalary, OSalary, rsNew("SH_SALARY"), True) Then UpdateAudit = True
    End If
    If isChanged_Salary(HRSalary, OSalCD, rsNew("SH_SALCD")) Then UpdateAudit = True
    
    If glbVadim And UpdateAudit Then
        'Ticket #21352 - City of Kawartha Lakes
        If glbCompSerial = "S/N - 2363W" Then
            Call Passing_Salary_Vadim(HRSalary, Salary, Date, xPHrs, xWHRS, xEmpNbr, xPayrollID, , xNiagaraWHRS)
        Else
            Call Passing_Salary_Vadim(HRSalary, Salary, xEDate, xPHrs, xWHRS, xEmpNbr, xPayrollID, , xNiagaraWHRS)
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
    Call Passing_Changes(HRChanges, Salary, "M", Date, xEmpNbr, xPayrollID)

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
Do While Not snapSalary.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / (lngRecs&)))
    MDIMain.panHelp(0).FloodPercent = pct%
    
    EmpNo& = snapSalary("SH_EMPNBR")
    SQLQ = "SELECT SH_EMPNBR,SH_ID FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & EmpNo& & " "
    SQLQ = SQLQ & "AND SH_EDATE = " & Date_SQL(dlpEffectiveDate.Text) & " "
    fTablSalHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not fTablSalHis.EOF Then
        Call DelSal_Addnew(fTablSalHis("SH_ID"))
    End If
    fTablSalHis.Close

lblNextRec:
    snapSalary.MoveNext
Loop
snapSalary.Close

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
        EmpNo& = fTablSalHis("TT_EMPNBR")
        If Not IsNull(fTablSalHis("TT_TBEMP")) Then
            SQLQ = "DELETE FROM HR_SALARY_HISTORY WHERE SH_ID = " & fTablSalHis("TT_TBEMP")
            gdbAdoIhr001.Execute SQLQ
        End If
        Call Employee_Master_Integration(EmpNo&)
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
        EmpNo& = fTablSalHis("TT_EMPNBR")
        Call Set_Current_Flag(EmpNo&)
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

If Not chkMUSalary() Then Exit Sub

If Not modGet_Salary_Records() Then Exit Sub       ' get selection - form level

Title$ = "Update Salary"

Msg$ = ""

If snapSalary.BOF And snapSalary.EOF Then
    Msg$ = Msg$ & "No Employees with this selection criteria exist!  " & Chr(10)
    Msg$ = Msg$ & "Please ensure the Hourly/Annually selection box is set correctly to match the group you want to update." & Chr(10)
    Msg$ = Msg$ & Chr(10)
    DgDef = MB_OK + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
    MsgBox Msg$, , Title$
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

lngRecs& = snapSalary.RecordCount
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

Private Sub ShowHide_AsofDate()
    'Ticket #22893 - WHSC - As of Date of the Salary to apply the Percent update on
    If optPct And Not glbMulti Then
        lblAsofDate.Visible = True
        dlpAsofDate.Visible = True
        imgHelp.Visible = True
    ElseIf Not optPct Then
        lblAsofDate.Visible = False
        dlpAsofDate.Visible = False
        imgHelp.Visible = False
    End If
End Sub

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


