VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmUSEMINARS 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Continuing Education Mass Change"
   ClientHeight    =   8160
   ClientLeft      =   -45
   ClientTop       =   1425
   ClientWidth     =   11340
   DrawMode        =   1  'Blackness
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8160
   ScaleWidth      =   11340
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   50
      SmallChange     =   4
      TabIndex        =   78
      Top             =   7800
      Width           =   11175
   End
   Begin VB.VScrollBar scrControl 
      Height          =   6075
      LargeChange     =   315
      Left            =   10920
      Max             =   100
      SmallChange     =   315
      TabIndex        =   77
      Top             =   1320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame scrFrame 
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   120
      TabIndex        =   42
      Top             =   120
      Width           =   10575
      Begin VB.TextBox txtCourseName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1815
         MaxLength       =   125
         TabIndex        =   4
         Tag             =   "01-Course Name"
         Top             =   810
         Width           =   4215
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   270
         Left            =   8700
         TabIndex        =   81
         Top             =   4560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame frmCourseCode 
         Height          =   375
         Left            =   1380
         TabIndex        =   2
         Top             =   7350
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox txtMain 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   330
            TabIndex        =   1
            Tag             =   "00-Course Code"
            Top             =   0
            Width           =   990
         End
         Begin VB.Image imgIcon 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   0
            Picture         =   "fuesemnr.frx":0000
            Top             =   30
            Width           =   240
         End
         Begin VB.Label lblDesc 
            Caption         =   "*** NOT ATTACHED ***"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1560
            TabIndex        =   52
            Top             =   30
            Width           =   3075
         End
      End
      Begin VB.TextBox txtExtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1815
         MaxLength       =   60
         TabIndex        =   5
         Tag             =   "00-Course Extended Name"
         Top             =   1140
         Width           =   4215
      End
      Begin VB.TextBox txtKeyword 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "00-Keyword"
         Top             =   450
         Width           =   1215
      End
      Begin VB.TextBox txtCourseHRS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         MaxLength       =   5
         TabIndex        =   18
         Tag             =   "11-Number of Scheduled Course Hours "
         Top             =   780
         Width           =   1215
      End
      Begin VB.TextBox txtAccount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6810
         MaxLength       =   25
         TabIndex        =   16
         Tag             =   "40-Account Number"
         Top             =   120
         Width           =   2085
      End
      Begin VB.CheckBox chkPresentor 
         Alignment       =   1  'Right Justify
         Caption         =   "Presenter"
         Height          =   195
         Left            =   6780
         TabIndex        =   29
         Tag             =   "40-Presentor"
         Top             =   2910
         Width           =   1065
      End
      Begin VB.TextBox txtTrainerName 
         Appearance      =   0  'Flat
         DataField       =   "ES_TRAINNER"
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "00-Trainer Name"
         Top             =   2190
         Width           =   4215
      End
      Begin VB.TextBox txtCompanyName 
         Appearance      =   0  'Flat
         DataField       =   "ES_COMPANYNAME"
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "00-Company Name"
         Top             =   1830
         Width           =   4215
      End
      Begin VB.TextBox txtCEUCred 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         MaxLength       =   5
         TabIndex        =   33
         Tag             =   "11-Number of Scheduled Course Hours "
         Top             =   3540
         Visible         =   0   'False
         Width           =   855
      End
      Begin INFOHR_Controls.DateLookup dlpRenewal 
         Height          =   285
         Left            =   1500
         TabIndex        =   13
         Tag             =   "40-Date when course is to be renewed"
         Top             =   3870
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDatComp 
         Height          =   285
         Left            =   1500
         TabIndex        =   12
         Tag             =   "41-Date when course was completed"
         Top             =   3540
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpStartDate 
         Height          =   285
         Left            =   1500
         TabIndex        =   11
         Tag             =   "41-Date when course was Started"
         Top             =   3210
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpSchDate 
         Height          =   285
         Left            =   1500
         TabIndex        =   10
         Tag             =   "40-Date when course was Scheduled"
         Top             =   2880
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   1500
         TabIndex        =   9
         Tag             =   "00-Results of the Course - Code"
         Top             =   2550
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ESRT"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   1500
         TabIndex        =   6
         Tag             =   "00-Organization/Individual Instructing - Code"
         Top             =   1470
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ESCB"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   1500
         TabIndex        =   3
         Tag             =   "00-Course Code"
         Top             =   460
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ESCD"
         MaxLength       =   8
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   1500
         TabIndex        =   0
         Tag             =   "01-Course Type - Code"
         Top             =   120
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ESCT"
         MaxLength       =   8
      End
      Begin Threed.SSFrame Frame3D1 
         Height          =   1245
         Left            =   0
         TabIndex        =   44
         Top             =   5970
         Width           =   9015
         _Version        =   65536
         _ExtentX        =   15901
         _ExtentY        =   2196
         _StockProps     =   14
         Caption         =   "Employees"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin INFOHR_Controls.EmployeeLookup elpEEID 
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   40
            Top             =   450
            Width           =   7245
            _ExtentX        =   12779
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   6925
            RefreshDescriptionWhen=   2
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.EmployeeLookup elpEEID 
            Height          =   285
            Index           =   1
            Left            =   1560
            TabIndex        =   41
            Top             =   870
            Width           =   7245
            _ExtentX        =   12779
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   6925
            RefreshDescriptionWhen=   2
            MultiSelect     =   -1  'True
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
            TabIndex        =   46
            Top             =   480
            Width           =   1290
         End
         Begin VB.Label lblEENum 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payroll ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   1035
         End
      End
      Begin MSMask.MaskEdBox medEECont 
         Height          =   285
         Index           =   0
         Left            =   7680
         TabIndex        =   19
         Tag             =   "20-Amount Employee Paid"
         Top             =   1110
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Currency"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEECont 
         Height          =   285
         Index           =   2
         Left            =   7680
         TabIndex        =   21
         Tag             =   "20-Other Expenses Paid"
         Top             =   1470
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Currency"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEECont 
         Height          =   285
         Index           =   1
         Left            =   7680
         TabIndex        =   23
         Tag             =   "20-Amount Employer Paid"
         Top             =   1830
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Currency"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medContTotal 
         Height          =   285
         Left            =   7680
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2550
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Currency"
         PromptChar      =   "_"
      End
      Begin Threed.SSFrame frmAttendance 
         Height          =   945
         Left            =   0
         TabIndex        =   43
         Top             =   4890
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   1667
         _StockProps     =   14
         Caption         =   "Attendance Data"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtAttHrs 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1150
            MaxLength       =   5
            TabIndex        =   35
            Tag             =   "11-Number of Hours Spent on Course"
            Top             =   600
            Width           =   615
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   4
            Left            =   840
            TabIndex        =   34
            Tag             =   "00-Attendance Reason - Code"
            Top             =   240
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ADRE"
         End
         Begin Threed.SSCheck chkSeniority 
            Height          =   195
            Left            =   3060
            TabIndex        =   37
            Tag             =   "Hours to be added to employee's seniority."
            Top             =   600
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Seniority    "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin Threed.SSCheck chkIncentive 
            Height          =   195
            Left            =   3060
            TabIndex        =   36
            Tag             =   "Incentive -  Attendance Management"
            Top             =   285
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Incentive  "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin VB.Label lbltitle 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Reason"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   48
            Top             =   290
            Width           =   660
         End
         Begin VB.Label lbltitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hours"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   47
            Top             =   630
            Width           =   510
         End
      End
      Begin Threed.SSFrame frmSkills 
         Height          =   945
         Left            =   4560
         TabIndex        =   49
         Top             =   4890
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   1667
         _StockProps     =   14
         Caption         =   "Skills Data"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtSkillsExp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1530
            MaxLength       =   2
            TabIndex        =   39
            Tag             =   "10-Skill Rating (0-99)"
            Top             =   600
            Width           =   615
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   5
            Left            =   1200
            TabIndex        =   38
            Tag             =   "00-Skills obtained resulting from course - Code"
            Top             =   270
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDSK"
         End
         Begin VB.Label lbltitle 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Skill"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   51
            Top             =   290
            Width           =   375
         End
         Begin VB.Label lbltitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Exp. Factor"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   50
            Top             =   630
            Width           =   990
         End
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   1500
         TabIndex        =   14
         Tag             =   "00-Co-Ordinated By"
         Top             =   4200
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ESCC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   7
         Left            =   1500
         TabIndex        =   15
         Tag             =   "00-Method Used"
         Top             =   4530
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ESMU"
      End
      Begin INFOHR_Controls.CodeLookup clpEmpCur 
         Height          =   285
         Left            =   9060
         TabIndex        =   20
         Top             =   1110
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "ECUR"
      End
      Begin INFOHR_Controls.CodeLookup clpOherCur 
         Height          =   285
         Left            =   9060
         TabIndex        =   22
         Top             =   1470
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "ECUR"
      End
      Begin INFOHR_Controls.CodeLookup clpEmployerCur 
         Height          =   285
         Left            =   9060
         TabIndex        =   24
         Top             =   1830
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "ECUR"
      End
      Begin INFOHR_Controls.CodeLookup clpAcomCur 
         Height          =   285
         Left            =   9060
         TabIndex        =   26
         Top             =   2190
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "ECUR"
      End
      Begin INFOHR_Controls.CodeLookup clpTotCur 
         Height          =   285
         Left            =   9060
         TabIndex        =   28
         Top             =   2550
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "ECUR"
      End
      Begin INFOHR_Controls.CodeLookup clpCEUType 
         Height          =   285
         Left            =   7365
         TabIndex        =   32
         Top             =   3210
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ESUT"
         MaxLength       =   8
      End
      Begin MSMask.MaskEdBox medEECont 
         Height          =   285
         Index           =   3
         Left            =   7680
         TabIndex        =   25
         Tag             =   "20-Accommodation"
         Top             =   2190
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Currency"
         PromptChar      =   "_"
      End
      Begin VB.Frame frmHrlSal 
         Height          =   375
         Left            =   8640
         TabIndex        =   80
         Top             =   2880
         Visible         =   0   'False
         Width           =   1815
         Begin Threed.SSCheck chkHrs 
            Height          =   255
            Left            =   0
            TabIndex        =   30
            Tag             =   "If X-Show Attendance Details"
            Top             =   0
            Width           =   795
            _Version        =   65536
            _ExtentX        =   1402
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Hourly"
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
         Begin Threed.SSCheck chkSal 
            Height          =   255
            Left            =   840
            TabIndex        =   31
            Tag             =   "If X-Show Attendance Details"
            Top             =   0
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Salary"
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
      Begin VB.Image imgNoSec 
         Height          =   240
         Left            =   8280
         Picture         =   "fuesemnr.frx":014A
         Top             =   4560
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSec 
         Height          =   240
         Left            =   8280
         Picture         =   "fuesemnr.frx":0294
         Top             =   4560
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblImport 
         Alignment       =   1  'Right Justify
         Caption         =   "Continuing Education"
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
         Left            =   5880
         TabIndex        =   82
         Top             =   4560
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Accommodation $  "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   20
         Left            =   6120
         TabIndex        =   79
         Top             =   2265
         Width           =   1380
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Course Type"
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
         Left            =   60
         TabIndex        =   76
         Top             =   150
         Width           =   1320
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Course Name"
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
         Index           =   2
         Left            =   60
         TabIndex        =   75
         Top             =   840
         Width           =   1500
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Course Description"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   60
         TabIndex        =   74
         Top             =   1170
         Width           =   1695
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Conducted By"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   73
         Top             =   1500
         Width           =   1200
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Results"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   72
         Top             =   2580
         Width           =   645
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Scheduled Date       "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   60
         TabIndex        =   71
         Top             =   2910
         Width           =   1605
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Completed"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   60
         TabIndex        =   70
         Top             =   3570
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Keyword"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6840
         TabIndex        =   69
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Course Hours"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   6510
         TabIndex        =   68
         Top             =   810
         Width           =   960
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee $"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   6360
         TabIndex        =   67
         Top             =   1140
         Width           =   1065
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Expenses $"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   6000
         TabIndex        =   66
         Top             =   1530
         Width           =   1425
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employer $"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   6000
         TabIndex        =   65
         Top             =   1860
         Width           =   1425
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total $"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   6060
         TabIndex        =   64
         Top             =   2610
         Width           =   1365
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   60
         TabIndex        =   63
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Renewal Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   60
         TabIndex        =   62
         Top             =   3900
         Width           =   1365
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Account #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   19
         Left            =   5640
         TabIndex        =   61
         Top             =   150
         Width           =   1035
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Course Code"
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
         Index           =   21
         Left            =   60
         TabIndex        =   60
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Trainer Name"
         Height          =   195
         Index           =   23
         Left            =   60
         TabIndex        =   59
         Top             =   2250
         Width           =   960
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         Height          =   195
         Index           =   22
         Left            =   60
         TabIndex        =   58
         Top             =   1860
         Width           =   1125
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Co-Ordinated By"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   57
         Top             =   4245
         Width           =   1365
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Method Used"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   24
         Left            =   60
         TabIndex        =   56
         Top             =   4590
         Width           =   1365
      End
      Begin VB.Label lbltitle 
         Alignment       =   2  'Center
         Caption         =   "Currency"
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
         Index           =   26
         Left            =   9420
         TabIndex        =   55
         Top             =   750
         Width           =   975
      End
      Begin VB.Label lblCEUType 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CEU Type"
         Height          =   255
         Left            =   5340
         TabIndex        =   54
         Top             =   3210
         Width           =   1815
      End
      Begin VB.Label lblCEUCred 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CEU Credit"
         Height          =   255
         Left            =   5580
         TabIndex        =   53
         Top             =   3540
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   5040
      Top             =   8280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmUSEMINARS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim fglbNew  As Integer
Dim fglHredsem As String
Dim fglCursName As String
Dim fglExtName As String
Dim fglAddDel As String
Dim strEMPLIST 'George Mar 14,2006
Dim RSEMPLIST As New ADODB.Recordset 'George Mar 14,2006
Dim OldtxtMain As String

Private Function chkFUSemnr() As Boolean
Dim oCode As String, OCodeD As String
Dim Msg, a%

chkFUSemnr = False

If Len(clpCode(1).Text) = 0 Then
    MsgBox "Course Type is a required field"
    clpCode(1).SetFocus
    Exit Function
End If
If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Course Type code must be valid"
    clpCode(1).SetFocus
    Exit Function
End If

'Ticket #22682: Release 8.0 - Made it mandatory to be consistent with Cont. Educ screen.
If Len(Trim(clpCode(0).Text)) = 0 Then
    MsgBox lStr("Course Code is a required field")
    If clpCode(0).Visible Then
        clpCode(0).SetFocus
    Else
        txtMain.SetFocus
    End If
    Exit Function
End If

If Len(clpCode(0).Text) > 0 Then
    If clpCode(0).Caption = "Unassigned" Then
        MsgBox "Course code must be valid"
        If clpCode(0).Visible Then
            clpCode(0).SetFocus
        Else
            txtMain.SetFocus
        End If
        Exit Function
    End If
End If

If Len(txtCourseName) < 1 Then
    MsgBox "Course Name is a required field"
    txtCourseName.SetFocus
    Exit Function
End If

If Len(clpCode(2).Text) > 0 Then
    If clpCode(2).Caption = "Unassigned" Then
        MsgBox "Conducted By code must be valid"
        clpCode(2).SetFocus
        Exit Function
    End If
End If
If Len(clpCode(3).Text) > 0 Then
    If clpCode(3).Caption = "Unassigned" Then
        MsgBox "Result code must be valid"
        clpCode(3).SetFocus
        Exit Function
    End If
End If

If Len(dlpDatComp.Text) > 0 Then
    If Not IsDate(dlpDatComp.Text) Then
        MsgBox "Date Completed is invalid"
        dlpDatComp.SetFocus
        Exit Function
    ElseIf Len(Year(dlpDatComp.Text)) = 3 Then
        MsgBox "Date Completed is invalid"
        dlpDatComp.SetFocus
        Exit Function
    End If
Else
    'MsgBox "Date completed is a required field"
    'dlpDatComp.SetFocus
'    dlpDatComp.Text = dlpStartDate.Text     ' added Sept 21, Laura
    'Exit Function
End If

If Len(medEECont(0)) < 1 Then
    medEECont(0) = 0
Else
    If Not IsNumeric(medEECont(0)) Then
        MsgBox "Employee's Contribution must be numeric"
        medEECont(0).SetFocus
        Exit Function
    End If
End If

If Len(medEECont(1)) < 1 Then
    medEECont(1) = 0
Else
    If Not IsNumeric(medEECont(1)) Then
        MsgBox "Employer's Contribution must be numeric"
        medEECont(1).SetFocus
        Exit Function
    End If

End If
If Len(medEECont(2)) < 1 Then
    medEECont(2) = 0
Else
    If Not IsNumeric(medEECont(2)) Then
        MsgBox "Other Expenses must be numeric"
        medEECont(2).SetFocus
        Exit Function
    End If

End If
If Len(medEECont(3)) < 1 Then
    medEECont(3) = 0
Else
    If Not IsNumeric(medEECont(2)) Then
        MsgBox "Accommodation must be numeric"
        medEECont(3).SetFocus
        Exit Function
    End If

End If

If Len(clpCode(4).Text) > 0 Then
    If clpCode(4).Caption = "Unassigned" Then
        MsgBox "Attendance Reason code must be valid"
        clpCode(4).SetFocus
        Exit Function
    End If
    If fglAddDel = "Add" Then
        If Len(txtAttHrs) <= 0 Then
            MsgBox "Hours is Required Field"
            txtAttHrs.SetFocus
            Exit Function
        End If
    End If
    'Ticket #23856
    If fglAddDel = "Add" Then
        If Not IsDate(dlpDatComp) Then
            MsgBox "Date Completed is required field when adding Attendance data"
            dlpDatComp.SetFocus
            Exit Function
        End If
    End If
End If

If Len(clpCode(5).Text) > 0 Then
    If clpCode(5).Caption = "Unassigned" Then
        MsgBox "Skill code must be valid"
        clpCode(5).SetFocus
        Exit Function
    End If
End If

If Len(txtAttHrs) >= 1 Then
    If Not IsNumeric(txtAttHrs) Then
        MsgBox "Attendance Hours must be numeric"
        txtAttHrs.SetFocus
        Exit Function
    End If
End If

If Len(txtSkillsExp) >= 1 Then
    If Not IsNumeric(txtSkillsExp) Then
        MsgBox "Skills Exp. must be numeric"
        txtSkillsExp.SetFocus
        Exit Function
    End If
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~ added by raubrey 7/29/97 for new data elements~~~~~~~~~~~~~~~
If Len(txtCourseHRS) > 0 Then
  If Not IsNumeric(txtCourseHRS) Then
        MsgBox "Course Hours must be numeric"
        txtCourseHRS.SetFocus
        Exit Function
    End If
End If
If glbLinamar Then
    If Len(txtCourseHRS) = 0 Then
        MsgBox "Course Hours is requried field"
        txtCourseHRS.SetFocus
        Exit Function
    End If
End If

If Len(dlpSchDate.Text) > 0 Then
    If Not IsDate(dlpSchDate.Text) Then
        MsgBox "Scheduled date is invalid"
        dlpSchDate.SetFocus
        Exit Function
    ElseIf Len(Year(dlpSchDate.Text)) = 3 Then
        MsgBox "Scheduled date is invalid"
        dlpSchDate.SetFocus
        Exit Function
    End If
End If

If Len(dlpStartDate.Text) > 0 Then
    If Not IsDate(dlpStartDate.Text) Then
        MsgBox "Start date is invalid"
        dlpStartDate.SetFocus
        Exit Function
    ElseIf Len(Year(dlpStartDate.Text)) = 3 Then
        MsgBox "Start date is invalid"
        dlpStartDate.SetFocus
        Exit Function
    End If
End If
'----------added on Sept 21 by laura
'----------removed on Mar 15 by Bryan
'If Len(dlpStartDate.Text) = 0 Then
'    MsgBox "Start Date is a required field"
'    dlpStartDate.SetFocus
'    Exit Function
'End If

'-----------------
If Len(dlpRenewal.Text) > 0 Then
    If Not IsDate(dlpRenewal.Text) Then
        MsgBox "Renewal date is invalid"
        dlpRenewal.SetFocus
        Exit Function
    ElseIf Len(Year(dlpRenewal.Text)) = 3 Then
        MsgBox "Renewal date is invalid"
        dlpRenewal.SetFocus
        Exit Function
    End If
End If
If Not elpEEID(0).ListChecker Then
    Exit Function
End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''chkFUSemnr = True

If Len(clpEmpCur.Text) > 0 Then
    If clpEmpCur.Caption = "Unassigned" Then
        MsgBox "Employee Currency code must be valid"
        clpEmpCur.SetFocus
        Exit Function
    End If
End If
If Len(clpOherCur.Text) > 0 Then
    If clpOherCur.Caption = "Unassigned" Then
        MsgBox "Other Expenses Currency code must be valid"
        clpOherCur.SetFocus
        Exit Function
    End If
End If
If Len(clpEmployerCur.Text) > 0 Then
    If clpEmployerCur.Caption = "Unassigned" Then
        MsgBox "Employer Currency code must be valid"
        clpEmployerCur.SetFocus
        Exit Function
    End If
End If
If Len(clpAcomCur.Text) > 0 Then
    If clpAcomCur.Caption = "Unassigned" Then
        MsgBox "Accomodations Currency code must be valid"
        clpAcomCur.SetFocus
        Exit Function
    End If
End If
If Len(clpTotCur.Text) > 0 Then
    If clpTotCur.Caption = "Unassigned" Then
        MsgBox "Total Currency code must be valid"
        clpTotCur.SetFocus
        Exit Function
    End If
End If

'Ticket #22682: Release 8.0 - for everyone
'Frank Apr 13, 2007 Ticket #12859 - City of Chatham-Kent
'If glbCompSerial = "S/N - 2188W" Then
    'If Len(clpCEUType.Text) > 0 Then
    '    If clpCEUType.Caption = "Unassigned" Then
    '        MsgBox "CEU Type code must be valid"
    '        clpCEUType.SetFocus
    '        Exit Function
    '    End If
    'End If
'Ticket #24708 Franks 11/28/2013 - make this for all customers
If Len(clpCEUType.Text) > 0 Then
    If clpCEUType.Caption = "Unassigned" Then
        MsgBox lStr("CEU Type") & " code must be valid"
        clpCEUType.SetFocus
        Exit Function
    End If
End If
If Len(txtCEUCred.Text) > 0 Then
    If Not IsNumeric(txtCEUCred.Text) Then
        MsgBox lStr("CEU Credit") & " is not number"
        txtCEUCred.SetFocus
        Exit Function
    End If
End If
'End If
    
'Ticket #22701 - County of Lanark
If glbCompSerial = "S/N - 2172W" Then
    If Len(txtCourseHRS) = 0 Or Val(txtCourseHRS) = 0 Then
        MsgBox "Course Hours is requried field"
        txtCourseHRS.SetFocus
        Exit Function
    End If
End If

If glbWFC Then 'Ticket #13520
    If Len(clpCode(0).Text) = 0 Then
            MsgBox lStr("Course Code is a required field")
            clpCode(0).SetFocus
            Exit Function
    End If
    If IsDate(dlpSchDate.Text) Then '#30081 Franks 06/23/2017
    Else
        If Len(clpCode(3).Text) = 0 Then
                MsgBox lStr("Results is a required field")
                clpCode(3).SetFocus
                Exit Function
        End If
        If Not IsDate(dlpDatComp) Then
                MsgBox lStr("Date Completed is a required field")
                dlpDatComp.SetFocus
                Exit Function
        End If
    End If
    If Len(txtCourseHRS.Text) = 0 Then
            MsgBox lStr("Course Hours is a required field")
            txtCourseHRS.SetFocus
            Exit Function
    Else
        If Val(txtCourseHRS.Text) = 0 Then
            MsgBox lStr("Course Hours is a required field")
            txtCourseHRS.SetFocus
            Exit Function
        End If
    End If
    If Val(Replace(medContTotal.Text, "$", "")) = 0 Then
        Msg = "The Total $ is 0.  "
        Msg = Msg & "Are You Sure there is no cost for this course?"
        
        a% = MsgBox(Msg, 36, "Confirm")
        If a% <> 6 Then Exit Function

    End If
End If

chkFUSemnr = True

End Function

Private Sub chkIncentive_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkSeniority_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
If Index = 4 Then Call ATTCode_Desc(Index)
If Index = 0 Then Call CrsName_Desc
If Index = 0 Then Call CourseCode_Type
End Sub
Private Sub CourseCode_Type()
Dim Response%, Msg$, Title$, DgDef As Double
Dim SQLQ As String, RType
Dim RSTABL As New ADODB.Recordset
On Error GoTo Dept_GL_Err

If glbWFC Then 'Ticket #24767 Franks 12/11/2013 WFC use data from Course Code Master setup
    Exit Sub
End If

If Len(clpCode(0).Text) > 0 Then
    RSTABL.Open "SELECT TB_NAME,TB_KEY,TB_USR1 FROM HRTABL WHERE TB_NAME = 'ESCD' AND TB_KEY='" & clpCode(0).Text & "'", gdbAdoIhr001
    If Not RSTABL.EOF Then
        If IsNull(RSTABL("TB_USR1")) Then
            RType = ""
        Else
            RType = RSTABL("TB_USR1")
        End If
        If Len(RType) > 0 Then
            If clpCode(1).Text <> RType Then
                    Msg$ = lStr("Do you want the associated Course Type?")
                    Title$ = "info:HR"
                    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
                    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                    If Response% = IDYES Then clpCode(1).Text = RType
            End If
        End If
    End If
    RSTABL.Close
End If

Exit Sub

Dept_GL_Err:
If Err = 94 Then
    ' clpGLNum.Text = ""
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Course Code Snap", "Course Code", "SELECT")
Call RollBack '21June99 js
End Sub
Public Sub cmdClose_Click()
Unload Me
If glbOnTop = "FRMUSEMINARS" Then glbOnTop = ""
End Sub

Public Sub cmdDelete_Click()
Dim Title$, DgDef As Variant, Response%, Msg$
Dim SQLQ As String
Dim X
Dim z As Integer, duplicate As Integer
Dim strSelCriteria1 As String
Dim strSelCriteria1a As String
Dim strSelCriteria1b As String
Dim strSelCriteria1c As String
Dim strSelCriteria2 As String
Dim strSelCriteria3 As String
Dim strSelCriteria4 As String
Dim strSelCriteria5 As String
Dim strSelCriteria5a As String
Dim strSelCriteria5b As String

If Not gSec_Upd_Education_Seminars Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If
On Error GoTo AddN_Err
'check for all controls

strEMPLIST = ""
glbstrSelCri = ""   'Ticket #17591

fglAddDel = "Delete"
If Not chkFUSemnr() Then Exit Sub

If Len(Trim(elpEEID(0))) = 0 And Len(Trim(elpEEID(1))) = 0 Then
    MsgBox "You have to enter at least one valid Employee Number or Payroll ID!"
    Exit Sub
End If

'Check if the employees to update to the user has security rights on them
If Len(elpEEID(0).Text) > 0 Then
    If Not Check_EmployeeID_Security Then
        MsgBox "Employee Number field contains invalid Employee Number." & vbCrLf & vbCrLf & "Aborting Mass Update."
        Exit Sub
    End If
End If
If Len(elpEEID(1).Text) > 0 Then
    If Not Check_PayrollID_Security Then
        MsgBox "Payroll ID field contains invalid Payroll ID" & vbCrLf & vbCrLf & "Aborting Mass Update."
        Exit Sub
    End If
End If

Title$ = "Mass Continuing Education Delete"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to delete Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

Msg$ = "Do you want to print a list of employees updated?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.

If Not modDelRecs() Then Exit Sub

If Response% = IDYES Then    ' Yes response
    'Call set_PrintState(False)
    Screen.MousePointer = HOURGLASS
    
    'Call getWSQLQ("U")
    
    
        'Ticket #13496 - City of Chatham Kent wants the selection criteria to be displayed on the report
        If glbCompSerial = "S/N - 2188W" Then
            strSelCriteria1 = ""
            strSelCriteria1a = ""
            strSelCriteria1b = ""
            strSelCriteria1c = ""
            strSelCriteria2 = ""
            strSelCriteria3 = ""
            strSelCriteria4 = ""
            strSelCriteria5 = ""
            strSelCriteria5a = ""
            strSelCriteria5b = ""
            
            If clpCode(1).Text <> "" Then
                strSelCriteria1 = strSelCriteria1 & "Course Type: " & clpCode(1).Text & "; "
            End If
            If clpCode(0).Text <> "" Then
                strSelCriteria1 = strSelCriteria1 & "Course Code: " & clpCode(0).Text & "; "
            End If
            If txtCourseName.Text <> "" Then
                strSelCriteria1 = strSelCriteria1 & "Course Name: " & txtCourseName.Text & "; "
            End If
            If txtExtName.Text <> "" Then
                strSelCriteria1a = strSelCriteria1a & "Course Description: " & txtExtName.Text & "; "
            End If
            If clpCode(2).Text <> "" Then
                strSelCriteria1b = strSelCriteria1b & "Conducted By: " & clpCode(2).Text & "; "
            End If
            If txtCompanyName.Text <> "" Then
                strSelCriteria1b = strSelCriteria1b & "Company Name: " & txtCompanyName.Text & "; "
            End If
            If txtTrainerName.Text <> "" Then
                strSelCriteria1c = strSelCriteria1c & "Trainer Name: " & txtTrainerName.Text & "; "
            End If
            If clpCode(3).Text <> "" Then
                strSelCriteria1c = strSelCriteria1c & "Results: " & clpCode(3).Text & "; "
            End If
            
            If dlpSchDate.Text <> "" Then
                strSelCriteria2 = strSelCriteria2 & "Scheduled Date: " & dlpSchDate.Text & "; "
            End If
            If dlpStartDate.Text <> "" Then
                strSelCriteria2 = strSelCriteria2 & "Start Date: " & dlpStartDate.Text & "; "
            End If
            If dlpDatComp.Text <> "" Then
                strSelCriteria2 = strSelCriteria2 & "Date Completed: " & dlpDatComp.Text & "; "
            End If
            If dlpRenewal.Text <> "" Then
                strSelCriteria2 = strSelCriteria2 & "Renewal Date: " & dlpRenewal.Text & "; "
            End If
            If clpCode(6).Text <> "" Then
                strSelCriteria2 = strSelCriteria2 & "Coordinated By: " & clpCode(6).Text & "; "
            End If
            If clpCode(7).Text <> "" Then
                strSelCriteria2 = strSelCriteria2 & "Method Used: " & clpCode(7).Text & "; "
            End If
            
            If txtAccount.Text <> "" Then
                strSelCriteria3 = strSelCriteria3 & "Account #: " & txtAccount.Text & "; "
            End If
            If txtKeyword.Text <> "" Then
                strSelCriteria3 = strSelCriteria3 & "Keyword #: " & txtKeyword.Text & "; "
            End If
            If txtCourseHRS.Text <> "" Then
                strSelCriteria3 = strSelCriteria3 & "Course Hours: " & txtCourseHRS.Text & "; "
            End If
            If medEECont(0).Text <> "" Then
                strSelCriteria3 = strSelCriteria3 & "Employee $: " & medEECont(0).Text & "; "
            End If
            If medEECont(2).Text <> "" Then
                strSelCriteria3 = strSelCriteria3 & "Other Expenses $: " & medEECont(2).Text & "; "
            End If
            If medEECont(1).Text <> "" Then
                strSelCriteria3 = strSelCriteria3 & "Employer $: " & medEECont(1).Text & "; "
            End If
            If medEECont(3).Text <> "" Then
                strSelCriteria3 = strSelCriteria3 & "Accommodation $: " & medEECont(3).Text & "; "
            End If
            If medContTotal.Text <> "" Then
                strSelCriteria3 = strSelCriteria3 & "Total $: " & medContTotal.Text & "; "
            End If
            strSelCriteria3 = strSelCriteria3 & "Presenter: " & IIf(chkPresentor.Value = "1", "Yes", "No") & "; "
            
            If clpCEUType.Text <> "" Then
                strSelCriteria3 = strSelCriteria3 & "CEU Type: " & clpCEUType.Text & "; "
            End If
            If txtCEUCred.Text <> "" Then
                strSelCriteria3 = strSelCriteria3 & "CEU Credit: " & txtCEUCred.Text & "; "
            End If
            
            If clpCode(4).Text <> "" Then
                strSelCriteria4 = "Attendance Reason: " & clpCode(4).Text & "; "
            End If
            If txtAttHrs.Text <> "" Then
                strSelCriteria4 = strSelCriteria4 & "Hours: " & txtAttHrs.Text & "; "
            End If
            If clpCode(4).Text <> "" Then
                strSelCriteria4 = strSelCriteria4 & "Incentive: " & IIf(chkIncentive.Value = "1", "Yes", "No") & "; "
                strSelCriteria4 = strSelCriteria4 & "Seniority: " & IIf(chkSeniority.Value = "1", "Yes", "No") & "; "
            End If
            
            If clpCode(5).Text <> "" Then
                strSelCriteria4 = strSelCriteria4 & Chr(10) & "Skill: " & clpCode(5).Text & "; "
            End If
            If txtSkillsExp.Text <> "" Then
                strSelCriteria4 = strSelCriteria4 & "Exp. Factor: " & txtSkillsExp.Text & "; "
            End If
        
            If elpEEID(0).Text <> "" Then
                strSelCriteria5 = "Employee Number: " & Replace(elpEEID(0).Text, ",", ", ") & "; "
            End If
            If elpEEID(1).Text <> "" Then
                strSelCriteria5 = strSelCriteria5 & Chr(10) & "Payroll ID: " & elpEEID(1).Text & "; "
            End If
            
            If Len(strSelCriteria1) > 0 Then
                'strSelCriteria = "Criteria: " & Chr(10) & strSelCriteria
                Me.vbxCrystal.Formulas(1) = "Criteria1 = '" & strSelCriteria1 & "' "
            End If
            If Len(strSelCriteria1a) > 0 Then
                'strSelCriteria = "Criteria: " & Chr(10) & strSelCriteria
                Me.vbxCrystal.Formulas(6) = "Criteria1a = '" & strSelCriteria1a & "' "
            End If
            If Len(strSelCriteria1b) > 0 Then
                'strSelCriteria = "Criteria: " & Chr(10) & strSelCriteria
                Me.vbxCrystal.Formulas(7) = "Criteria1b = '" & strSelCriteria1b & "' "
            End If
            If Len(strSelCriteria1c) > 0 Then
                'strSelCriteria = "Criteria: " & Chr(10) & strSelCriteria
                Me.vbxCrystal.Formulas(8) = "Criteria1c = '" & strSelCriteria1c & "' "
            End If
            
            If Len(strSelCriteria2) > 0 Then
                Me.vbxCrystal.Formulas(2) = "Criteria2 = '" & strSelCriteria2 & "' "
            End If
            If Len(strSelCriteria3) > 0 Then
                Me.vbxCrystal.Formulas(3) = "Criteria3 = '" & strSelCriteria3 & "' "
            End If
            If Len(strSelCriteria4) > 0 Then
                Me.vbxCrystal.Formulas(4) = "Criteria4 = '" & strSelCriteria4 & "' "
            End If
            If Len(strSelCriteria5) > 0 And Len(strSelCriteria5) > 254 Then
                Me.vbxCrystal.Formulas(5) = "Criteria5 = '" & Mid(strSelCriteria5, 1, 254) & "' "
                strSelCriteria5a = Mid(strSelCriteria5, 255)
                Me.vbxCrystal.Formulas(6) = "Criteria5a = '" & Mid(strSelCriteria5a, 1, 254) & "' "
                If Len(strSelCriteria5a) > 254 Then
                    strSelCriteria5b = Mid(strSelCriteria5a, 255)
                    Me.vbxCrystal.Formulas(7) = "Criteria5b = '" & strSelCriteria5b & "' "
                End If
            Else
                If Len(strSelCriteria5) > 0 Then
                    Me.vbxCrystal.Formulas(5) = "Criteria5 = '" & strSelCriteria5 & "' "
                End If
            End If
        End If
        
    
    'report name
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList.rpt"

    Me.vbxCrystal.Formulas(0) = "rTitle='Mass Delete Continuing Education - Employee Details'"
    
    'Ticket #29485 - Show the Course Code being updated for all
    If Len(clpCode(1).Text) > 0 Then Me.vbxCrystal.Formulas(21) = "lblnote1 = 'Course Type: " & clpCode(1).Text & " - " & getCodeDesc("ESCT", clpCode(1).Text) & "' "
    If Len(clpCode(0).Text) > 0 Then Me.vbxCrystal.Formulas(22) = "lblnote2 = 'Course Code: " & clpCode(0).Text & " - " & getCodeDesc("ESCD", clpCode(0).Text) & "' "
    
    'set location for database tables
    If Len(glbstrSelCri) >= 0 Then
        If Len(strEMPLIST) > 0 Then
            Me.vbxCrystal.SelectionFormula = getWSQLQRPT
        Else
            Me.vbxCrystal.SelectionFormula = "1=2"
        End If
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

Screen.MousePointer = DEFAULT
'MsgBox "Records Added Successfully!"

Exit Sub

AddN_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDelete", "HREDSEM", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Public Sub cmdNew_Click()
Dim Title$, DgDef As Variant, Response%, Msg$
Dim SQLQ As String
Dim X
Dim z As Integer, duplicate As Integer
Dim strSelCriteria1 As String
Dim strSelCriteria1a As String
Dim strSelCriteria1b As String
Dim strSelCriteria1c As String
Dim strSelCriteria2 As String
Dim strSelCriteria3 As String
Dim strSelCriteria4 As String
Dim strSelCriteria5 As String
Dim strSelCriteria5a As String
Dim strSelCriteria5b As String
Dim strlblNote As String

If Not gSec_Upd_Education_Seminars Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If
On Error GoTo AddN_Err
'check for all controls

strEMPLIST = ""
glbstrSelCri = ""   'Ticket #17591
 
fglAddDel = "Add"
If Not chkFUSemnr() Then Exit Sub
'check if you enter an employee
'If elpEEID = "" Then
'  MsgBox "You have to enter at least one valid employee number!"
'  Exit Sub
'End If
If Len(Trim(elpEEID(0))) = 0 And Len(Trim(elpEEID(1))) = 0 Then
    MsgBox "You have to enter at least one valid Employee Number or Payroll ID!"
    Exit Sub
End If

'Check if the employees to update to the user has security rights on them
If Len(elpEEID(0).Text) > 0 Then
    If Not Check_EmployeeID_Security Then
        MsgBox "Employee Number field contains invalid Employee Number." & vbCrLf & vbCrLf & "Aborting Mass Update."
        Exit Sub
    End If
End If
If Len(elpEEID(1).Text) > 0 Then
    If Not Check_PayrollID_Security Then
        MsgBox "Payroll ID field contains invalid Payroll ID." & vbCrLf & vbCrLf & "Aborting Mass Update."
        Exit Sub
    End If
End If

'Ticket #22682
If Len(clpCode(4).Text) > 0 And IsNumeric(txtAttHrs.Text) Then
    If Not gSec_Add_Attendance Then
        MsgBox "You do not have authority to Add Attendance transaction." & vbCrLf & vbCrLf & "Aborting Mass Update.", vbExclamation, "Adding Attendance"
        Exit Sub
    End If
End If

Title$ = "Mass Continuing Education Add"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to add Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

Msg$ = "Do you want to print a list of employees updated?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.

If Not modInsRecs() Then Exit Sub

If Response% = IDYES Then    ' Yes response
    'Call set_PrintState(False)
    Screen.MousePointer = HOURGLASS
    
    'Call getWSQLQ("U")
    
    'Ticket #13496 - City of Chatham Kent wants the selection criteria to be displayed on the report
    If glbCompSerial = "S/N - 2188W" Then
        strSelCriteria1 = ""
        strSelCriteria1a = ""
        strSelCriteria1b = ""
        strSelCriteria1c = ""
        strSelCriteria2 = ""
        strSelCriteria3 = ""
        strSelCriteria4 = ""
        strSelCriteria5 = ""
        strSelCriteria5a = ""
        strSelCriteria5b = ""
        
        If clpCode(1).Text <> "" Then
            strSelCriteria1 = strSelCriteria1 & "Course Type: " & clpCode(1).Text & "; "
        End If
        If clpCode(0).Text <> "" Then
            strSelCriteria1 = strSelCriteria1 & "Course Code: " & clpCode(0).Text & "; "
        End If
        If txtCourseName.Text <> "" Then
            strSelCriteria1 = strSelCriteria1 & "Course Name: " & txtCourseName.Text & "; "
        End If
        
        If txtExtName.Text <> "" Then
            strSelCriteria1a = strSelCriteria1a & "Course Description: " & txtExtName.Text & "; "
        End If
        If clpCode(2).Text <> "" Then
            strSelCriteria1b = strSelCriteria1b & "Conducted By: " & clpCode(2).Text & "; "
        End If
        
        If txtCompanyName.Text <> "" Then
            strSelCriteria1b = strSelCriteria1b & "Company Name: " & txtCompanyName.Text & "; "
        End If
        
        If txtTrainerName.Text <> "" Then
            strSelCriteria1c = strSelCriteria1c & "Trainer Name: " & txtTrainerName.Text & "; "
        End If
        If clpCode(3).Text <> "" Then
            strSelCriteria1c = strSelCriteria1c & "Results: " & clpCode(3).Text & "; "
        End If
        
        If dlpSchDate.Text <> "" Then
            strSelCriteria2 = strSelCriteria2 & "Scheduled Date: " & dlpSchDate.Text & "; "
        End If
        If dlpStartDate.Text <> "" Then
            strSelCriteria2 = strSelCriteria2 & "Start Date: " & dlpStartDate.Text & "; "
        End If
        If dlpDatComp.Text <> "" Then
            strSelCriteria2 = strSelCriteria2 & "Date Completed: " & dlpDatComp.Text & "; "
        End If
        If dlpRenewal.Text <> "" Then
            strSelCriteria2 = strSelCriteria2 & "Renewal Date: " & dlpRenewal.Text & "; "
        End If
        If clpCode(6).Text <> "" Then
            strSelCriteria2 = strSelCriteria2 & "Coordinated By: " & clpCode(6).Text & "; "
        End If
        If clpCode(7).Text <> "" Then
            strSelCriteria2 = strSelCriteria2 & "Method Used: " & clpCode(7).Text & "; "
        End If
        
        If txtAccount.Text <> "" Then
            strSelCriteria3 = strSelCriteria3 & "Account #: " & txtAccount.Text & "; "
        End If
        If txtKeyword.Text <> "" Then
            strSelCriteria3 = strSelCriteria3 & "Keyword #: " & txtKeyword.Text & "; "
        End If
        If txtCourseHRS.Text <> "" Then
            strSelCriteria3 = strSelCriteria3 & "Course Hours: " & txtCourseHRS.Text & "; "
        End If
        If medEECont(0).Text <> "" Then
            strSelCriteria3 = strSelCriteria3 & "Employee $: " & medEECont(0).Text & "; "
        End If
        If medEECont(2).Text <> "" Then
            strSelCriteria3 = strSelCriteria3 & "Other Expenses $: " & medEECont(2).Text & "; "
        End If
        If medEECont(1).Text <> "" Then
            strSelCriteria3 = strSelCriteria3 & "Employer $: " & medEECont(1).Text & "; "
        End If
        If medEECont(3).Text <> "" Then
            strSelCriteria3 = strSelCriteria3 & "Accommodation $: " & medEECont(3).Text & "; "
        End If
        If medContTotal.Text <> "" Then
            strSelCriteria3 = strSelCriteria3 & "Total $: " & medContTotal.Text & "; "
        End If
        strSelCriteria3 = strSelCriteria3 & "Presenter: " & IIf(chkPresentor.Value = "1", "Yes", "No") & "; "
        
        If clpCEUType.Text <> "" Then
            strSelCriteria3 = strSelCriteria3 & "CEU Type: " & clpCEUType.Text & "; "
        End If
        If txtCEUCred.Text <> "" Then
            strSelCriteria3 = strSelCriteria3 & "CEU Credit: " & txtCEUCred.Text & "; "
        End If
        
        If clpCode(4).Text <> "" Then
            strSelCriteria4 = "Attendance Reason: " & clpCode(4).Text & "; "
        End If
        If txtAttHrs.Text <> "" Then
            strSelCriteria4 = strSelCriteria4 & "Hours: " & txtAttHrs.Text & "; "
        End If
        If clpCode(4).Text <> "" Then
            strSelCriteria4 = strSelCriteria4 & "Incentive: " & IIf(chkIncentive.Value = "1", "Yes", "No") & "; "
            strSelCriteria4 = strSelCriteria4 & "Seniority: " & IIf(chkSeniority.Value = "1", "Yes", "No") & "; "
        End If
        
        If clpCode(5).Text <> "" Then
            strSelCriteria4 = strSelCriteria4 & Chr(10) & "Skill: " & clpCode(5).Text & "; "
        End If
        If txtSkillsExp.Text <> "" Then
            strSelCriteria4 = strSelCriteria4 & "Exp. Factor: " & txtSkillsExp.Text & "; "
        End If
    
        If elpEEID(0).Text <> "" Then
            strSelCriteria5 = "Employee Number: " & Replace(elpEEID(0).Text, ",", ", ") & "; "
        End If
        If elpEEID(1).Text <> "" Then
            strSelCriteria5 = strSelCriteria5 & Chr(10) & "Payroll ID: " & elpEEID(1).Text & "; "
        End If
        
        If Len(strSelCriteria1) > 0 Then
            'strSelCriteria = "Criteria: " & Chr(10) & strSelCriteria
            Me.vbxCrystal.Formulas(1) = "Criteria1 = '" & strSelCriteria1 & "' "
        End If
        If Len(strSelCriteria1a) > 0 Then
            'strSelCriteria = "Criteria: " & Chr(10) & strSelCriteria
            Me.vbxCrystal.Formulas(6) = "Criteria1a = '" & strSelCriteria1a & "' "
        End If
        If Len(strSelCriteria1b) > 0 Then
            'strSelCriteria = "Criteria: " & Chr(10) & strSelCriteria
            Me.vbxCrystal.Formulas(7) = "Criteria1b = '" & strSelCriteria1b & "' "
        End If
        If Len(strSelCriteria1c) > 0 Then
            'strSelCriteria = "Criteria: " & Chr(10) & strSelCriteria
            Me.vbxCrystal.Formulas(8) = "Criteria1c = '" & strSelCriteria1c & "' "
        End If
        
        If Len(strSelCriteria2) > 0 Then
            Me.vbxCrystal.Formulas(2) = "Criteria2 = '" & strSelCriteria2 & "' "
        End If
        If Len(strSelCriteria3) > 0 Then
            Me.vbxCrystal.Formulas(3) = "Criteria3 = '" & strSelCriteria3 & "' "
        End If
        If Len(strSelCriteria4) > 0 Then
            Me.vbxCrystal.Formulas(4) = "Criteria4 = '" & strSelCriteria4 & "' "
        End If
        If Len(strSelCriteria5) > 0 And Len(strSelCriteria5) > 254 Then
            Me.vbxCrystal.Formulas(5) = "Criteria5 = '" & Mid(strSelCriteria5, 1, 254) & "' "
            strSelCriteria5a = Mid(strSelCriteria5, 255)
            Me.vbxCrystal.Formulas(6) = "Criteria5a = '" & Mid(strSelCriteria5a, 1, 254) & "' "
            If Len(strSelCriteria5a) > 254 Then
                strSelCriteria5b = Mid(strSelCriteria5a, 255)
                Me.vbxCrystal.Formulas(7) = "Criteria5b = '" & strSelCriteria5b & "' "
            End If
        Else
            If Len(strSelCriteria5) > 0 Then
                Me.vbxCrystal.Formulas(5) = "Criteria5 = '" & strSelCriteria5 & "' "
            End If
        End If
    End If
    
    ' report name
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList.rpt"
    Me.vbxCrystal.Formulas(0) = "rTitle='Mass Update Continuing Education - Employee Details'"
    If glbWFC Then 'Ticket #22500 Franks 09/10/2012
        If Len(clpCode(1).Text) > 0 Then Me.vbxCrystal.Formulas(21) = "lblnote1 = 'Course Type: " & getCodeDesc("ESCT", clpCode(1).Text) & "' "
        If Len(clpCode(0).Text) > 0 Then Me.vbxCrystal.Formulas(22) = "lblnote2 = 'Course Code: " & getCodeDesc("ESCD", clpCode(0).Text) & "' "
        If Len(txtCourseHRS.Text) > 0 Then
            'If Not txtCourseHRS.Text = "0" Then
            Me.vbxCrystal.Formulas(23) = "lblnote3 = 'Course Hours: " & txtCourseHRS.Text & "' "
            'End If
        End If
        If Len(dlpDatComp.Text) > 0 Then Me.vbxCrystal.Formulas(24) = "lblnote4 = 'Date Completed: " & dlpDatComp.Text & "' "
    Else
        'Ticket #29485 - Show the Course Code being updated for all
        If Len(clpCode(1).Text) > 0 Then Me.vbxCrystal.Formulas(21) = "lblnote1 = 'Course Type: " & clpCode(1).Text & " - " & getCodeDesc("ESCT", clpCode(1).Text) & "' "
        If Len(clpCode(0).Text) > 0 Then Me.vbxCrystal.Formulas(22) = "lblnote2 = 'Course Code: " & clpCode(0).Text & " - " & getCodeDesc("ESCD", clpCode(0).Text) & "' "
    End If
    'set location for database tables
    If Len(glbstrSelCri) >= 0 And Len(strEMPLIST) > 0 Then
        Me.vbxCrystal.SelectionFormula = getWSQLQRPT
    Else
        Me.vbxCrystal.SelectionFormula = "1=2"
    End If
'        If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
'        Else
'            Me.vbxCrystal.Connect = "PWD=petman;"
'            For X = 0 To 4
'                Me.vbxCrystal.DataFiles(X) = glbIHRDB
'            Next X
'        End If
    
    ' window title if appropriate
    Me.vbxCrystal.WindowTitle = "Employees-updated Report"
    
    Me.vbxCrystal.Destination = 0
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset

End If

Screen.MousePointer = DEFAULT
'MsgBox "Records Added Successfully!"


Exit Sub

AddN_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HREDSEM", "Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub ATTCode_Desc(Indx As Integer)
Dim SQLQ As String
Dim rsTA As New ADODB.Recordset
On Error GoTo AttCode_desc_Err


If Len(clpCode(Indx).Text) > 0 And Indx = 4 Then
    SQLQ = "SELECT TB_INDICATOR FROM HRTABL WHERE TB_NAME='ADRE' AND TB_KEY = '" & clpCode(Indx).Text & "'"
    rsTA.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsTA.EOF Then
        If rsTA("TB_INDICATOR") = True Then
            chkIncentive.Value = True
        Else
            chkIncentive.Value = False
        End If
    End If
End If

Exit Sub

AttCode_desc_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Code Snap", "TABL", "SELECT")
Call RollBack '23July99 js

End Sub

Private Sub cmdImport_Click()
    Dim tmpEmp As Integer
    
    'Ticket #18721
    glbDocNewRecord = True
    glbDocName = "EdSem"
    glbDocKey = 0
    tmpEmp = 0
    
    If glbLEE_ID <> 0 Then
        tmpEmp = glbLEE_ID
        glbLEE_ID = 0
    End If
    
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmESEMINARS")
    
    If tmpEmp > 0 Then
        glbLEE_ID = tmpEmp
    End If
End Sub

Private Sub dlpStartDate_LostFocus()
    If Len(dlpDatComp) = 0 Then 'Ticket# 7384
        If IsDate(dlpStartDate) Then
            dlpDatComp = dlpStartDate
        End If
    End If
End Sub

Private Sub EmployeeLookup1_Change()

End Sub

Private Sub empPayrollID_DblClick()
'frmEPayrollID.Show 1
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMUSEMINARS"
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Screen.MousePointer = HOURGLASS

glbOnTop = "FRMUSEMINARS"

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    frmUSEMINARS.Caption = "Continuing Education - Mass Change"
End If

lblTitle(0).Caption = lStr(lblTitle(0).Caption)
lblTitle(1).Caption = lStr(lblTitle(1).Caption)
lblTitle(21).Caption = lStr(lblTitle(21).Caption)
lblTitle(2).Caption = lStr(lblTitle(2).Caption)
lblTitle(16).Caption = lStr(lblTitle(16).Caption)
lblTitle(3).Caption = lStr(lblTitle(3).Caption)
lblTitle(22).Caption = lStr(lblTitle(22).Caption)
lblTitle(23).Caption = lStr(lblTitle(23).Caption)
lblTitle(4).Caption = lStr(lblTitle(4).Caption)
lblTitle(13).Caption = lStr(lblTitle(13).Caption)
lblTitle(14).Caption = lStr(lblTitle(14).Caption)
lblTitle(5).Caption = lStr(lblTitle(5).Caption)
lblTitle(18).Caption = lStr(lblTitle(18).Caption)
lblTitle(24).Caption = lStr(lblTitle(24).Caption)
lblTitle(19).Caption = lStr(lblTitle(19).Caption)
Label1.Caption = lStr(Label1.Caption)
lblTitle(17).Caption = lStr(lblTitle(17).Caption)
lblTitle(6).Caption = lStr(lblTitle(6).Caption)
lblTitle(15).Caption = lStr(lblTitle(15).Caption)
lblTitle(7).Caption = lStr(lblTitle(7).Caption)
lblTitle(20).Caption = lStr(lblTitle(20).Caption)
chkPresentor.Caption = lStr(chkPresentor.Caption)
lblCEUType.Caption = lStr("CEU Type")

If glbCompSerial = "S/N - 2214W" Then
   lblTitle(6).Caption = "Tuition $"
   lblTitle(7).Caption = "Travel $"
   lblTitle(15).Caption = "Daily Allowance $"
End If

'Hemu - 05/29/2003 Begin - Ticket # 4204
If glbCompSerial = "S/N - 2161W" Then
    clpCode(1).TextBoxWidth = 1200
    clpCode(0).TextBoxWidth = 1200
    clpCode(1).MaxLength = 8
    clpCode(0).MaxLength = 8
Else
    clpCode(1).TextBoxWidth = 870
    clpCode(0).TextBoxWidth = 870
    clpCode(1).MaxLength = 8 '4 Ticket #20498 Franks 06/23/2011
    clpCode(0).MaxLength = 8
End If
'Hemu - 05/29/2003 End
If glbOttawaCCAC Then
    lblTitle(6) = "Travel $"
    lblTitle(7) = "Registration $"
    medEECont(0).Tag = "20-Amount Travel"
    medEECont(1).Tag = "20-Amount Registration"
    
End If
If glbLinamar Then
    lblTitle(17).FontBold = True
End If

'If Course Code Master setup, then use it instead of Course Code lookup from HRTABL
'Ticket #12204
If getCrsCodeMasterFlag Then
    clpCode(0).Visible = False
    lblDesc.Caption = ""
    frmCourseCode.Visible = True
    frmCourseCode.Top = clpCode(0).Top
    frmCourseCode.Left = clpCode(0).Left
    frmCourseCode.BorderStyle = 0
End If

'Ticket #22682: Release 8.0 - for everyone
'Frank added Apr 13, 2007 Ticket#12859
'If glbCompSerial = "S/N - 2188W" Then
    'lblCEUType.Visible = True
    lblCEUCred.Visible = True
    txtCEUCred.Visible = True
    'clpCEUType.Visible = True
'End If
'Ticket #24708 Franks 11/28/2013 - make CEU Type show up for WFC and all other customers
lblCEUType.Visible = True
clpCEUType.Visible = True
    
'Ticket #18721
If gsAttachment_DB Then
    'glbJob = ""
    'glbSDate = "01/01/1900"
    lblImport.Visible = True 'False
    imgSec.Visible = False
    imgNoSec.Visible = True 'False
    cmdImport.Visible = True 'False
End If

Call INI_Controls(Me)

If glbWFC Then 'Ticket #15818
    Call WFCScreenSetup
End If

'Ticket #22701 - County of Lanark
If glbCompSerial = "S/N - 2172W" Then
    lblTitle(17).FontBold = True
End If


Screen.MousePointer = DEFAULT


End Sub

Private Sub Form_Resize()
scrFrame.Height = 9800
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 9350 Then
        scrControl.Value = 0
        scrFrame.Top = 120
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 7000 Then
            scrControl.Max = 4000
        Else
            scrControl.Max = 1500
        End If
        scrControl.Left = Me.Width - scrControl.Width - 120
        If Me.Height - scrControl.Top - 780 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - 780
        End If
    End If
    
    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 120
    'scrFrame.Height = Me.ScaleHeight - (scrHScroll.Height - 200)  '
    If Me.Width >= 10900 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        If Me.Width < 9000 Then
            scrHScroll.Max = 100
        Else
            scrHScroll.Max = 30
        End If
        scrHScroll.Top = Me.Height - 800
        scrHScroll.Width = Me.Width - 120
    End If
    scrFrame.Refresh
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmUSEMINARS = Nothing  'carmen apr 2000
End Sub

Private Sub imgIcon_Click()
Call txtMain_DblClick
End Sub

Private Sub imgSec_Click()
'    'Ticket #18721
'    Dim SQLQ
'    SQLQ = getSQL("frmESEMINARS")
    
    'Ticket #28314 - To view the document before doing the mass update.
'    If Len(Trim(SQLQ)) = 0 Then
        If Len(glbDocImpFile) > 0 Then
            Shell "cmd /c " & GetShortName(glbDocImpFile)
        End If
'    Else
'        Call FillMemoFile(SQLQ, "EdSem")
'    End If
End Sub

Private Sub medContTotal_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medEECont_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medEECont_LostFocus(Index As Integer)
Call UpConttotal
End Sub
Private Function getEmpnoFromPayID(xPayIDlist)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ, xStr
    xStr = ""
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn & " "
    SQLQ = SQLQ & "AND ED_PAYROLL_ID IN (" & xPayIDlist & ") "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsTemp.EOF
        xStr = xStr & rsTemp("ED_EMPNBR") & ","
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    getEmpnoFromPayID = xStr
End Function
Private Function modDelRecs()
Dim xEMP, xEmpList, xPayIDlist
Dim SQLQ As String, TSQLQ0 As String, TSQLQ As String, nTable As String, xSHIFT, xSuper
Dim newline As String, Message As String
Dim Msg As String, Edat As String
Dim iRec As Integer
Dim TB As New ADODB.Recordset
Dim TF As New ADODB.Recordset
Dim TA As New ADODB.Recordset
Dim TS As New ADODB.Recordset
Dim TH As New ADODB.Recordset
Dim xDup
Dim rsDup As New ADODB.Recordset
'Dim dynHRAT As New ADODB.Recordset
Dim X As Integer, all As Integer, z As Integer
Dim rsEDSEM As New ADODB.Recordset
Dim SQLQD As String
Dim recCount As Integer
Dim Response%

Screen.MousePointer = HOURGLASS

newline = Chr(13) & Chr(10)
modDelRecs = False
all = True

On Error GoTo CrFollow_Err 'laura nov 12, 1997

xEmpList = getEmpnbr(elpEEID(0))

xPayIDlist = getPayrollID(elpEEID(1))
'xPayIDlist = getPayrollID(xPayIDlist)
If Len(Trim(elpEEID(0))) = 0 Then
    xEmpList = getEmpnoFromPayID(xPayIDlist)
    xEmpList = getEmpnbr(xEmpList)
End If

TSQLQ0 = "SELECT ES_EMPNBR FROM HREDSEM "
TSQLQ = TSQLQ & " WHERE ES_CTYPE = '" & clpCode(1).Text & "' "
TSQLQ = TSQLQ & " AND ES_EMPNBR IN (" & xEmpList & ")"

'Course Code
If Len(clpCode(0).Text) > 0 Then
    TSQLQ = TSQLQ & " AND ES_CRSCODE = '" & clpCode(0).Text & "' "
'Else
'    TSQLQ = TSQLQ & " AND (ES_CRSCODE = '' OR ES_CRSCODE IS NULL) "
End If

'Course Name
If Len(txtCourseName.Text) > 0 Then
    TSQLQ = TSQLQ & " AND ES_COURSE = '" & txtCourseName.Text & "' "
'Else
'    TSQLQ = TSQLQ & " AND (ES_COURSE = '' OR ES_COURSE IS NULL) "
End If

'Course Description
If Len(txtExtName.Text) > 0 Then
    TSQLQ = TSQLQ & " AND ES_EXTNAME = '" & txtExtName.Text & "' "
'Else
'    TSQLQ = TSQLQ & " AND (ES_EXTNAME = '' OR ES_EXTNAME IS NULL) "
End If

'Conducted By
If Len(clpCode(2).Text) > 0 Then
    TSQLQ = TSQLQ & " AND ES_CONDUCT = '" & clpCode(2).Text & "' "
'Else
'    TSQLQ = TSQLQ & " AND (ES_CONDUCT = '' OR ES_CONDUCT IS NULL) "
End If

'Results
If Len(clpCode(3).Text) > 0 Then
    TSQLQ = TSQLQ & " AND ES_RESULTS = '" & clpCode(3).Text & "' "
'Else
'    TSQLQ = TSQLQ & " AND (ES_RESULTS = '' OR ES_RESULTS IS NULL) "
End If

'Scheduled Date
If IsDate(dlpSchDate.Text) Then
    TSQLQ = TSQLQ & " AND ES_SCHEDULED = " & Date_SQL(dlpSchDate.Text)
'Else
'    TSQLQ = TSQLQ & " AND ES_SCHEDULED IS NULL "
End If

'Start Date
If IsDate(dlpStartDate.Text) Then
    TSQLQ = TSQLQ & " AND ES_START = " & Date_SQL(dlpStartDate.Text)
'Else
'    TSQLQ = TSQLQ & " AND ES_START IS NULL "
End If

'Date Completed
If IsDate(dlpDatComp.Text) Then
    TSQLQ = TSQLQ & " AND ES_DATCOMP = " & Date_SQL(dlpDatComp.Text)
'Else
'    TSQLQ = TSQLQ & " AND ES_DATCOMP IS NULL "
End If

'Renewal Date
If IsDate(dlpRenewal.Text) Then
    TSQLQ = TSQLQ & " AND ES_RENEW = " & Date_SQL(dlpRenewal.Text)
'Else
'    TSQLQ = TSQLQ & " AND ES_RENEW IS NULL "
End If

'Co-Ordinated By
If Len(clpCode(6).Text) > 0 Then
    TSQLQ = TSQLQ & " AND ES_COORDINATED = '" & clpCode(6).Text & "' "
'Else
'    TSQLQ = TSQLQ & " AND (ES_COORDINATED = '' OR ES_COORDINATED IS NULL) "
End If

'Method Used
If Len(clpCode(7).Text) > 0 Then
    TSQLQ = TSQLQ & " AND ES_METHODUSED = '" & clpCode(7).Text & "' "
'Else
'    TSQLQ = TSQLQ & " AND (ES_METHODUSED = '' OR ES_METHODUSED IS NULL) "
End If

'CEU Type
If Len(clpCEUType.Text) > 0 Then
    TSQLQ = TSQLQ & " AND ES_CEUTYPE = '" & clpCEUType.Text & "' "
'Else
'    TSQLQ = TSQLQ & " AND (ES_CEUTYPE = '' OR ES_CEUTYPE IS NULL) "
End If

'CEU Credit
If Len(txtCEUCred.Text) > 0 Then
    TSQLQ = TSQLQ & " AND ES_CEUCREDIT = " & txtCEUCred.Text & " "
End If

SQLQ = TSQLQ0 & TSQLQ
rsDup.Open SQLQ, gdbAdoIhr001, adOpenKeyset
If rsDup.EOF Then
    MsgBox "There are no any records to be deleted for this criteria."
    Screen.MousePointer = DEFAULT
    Exit Function
End If
rsDup.Close

recCount = getRecordCount_Delete(xEmpList)
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Continuing Education Record " Else Msg$ = Msg$ & " Continuing Education Records "
    Msg$ = Msg$ & "will be Deleted. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, "Mass Continuing Education Delete")     ' Get user response.
    If Response = IDNO Then
        Screen.MousePointer = DEFAULT
        Exit Function
    End If
Else
    Screen.MousePointer = DEFAULT
    MsgBox "No Continuing Education record found for this selection criteria to delete."
    Exit Function
End If

TSQLQ0 = "SELECT ES_EMPNBR FROM HREDSEM "
SQLQ = TSQLQ0 & TSQLQ
RSEMPLIST.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly
Do While Not RSEMPLIST.EOF
    If Len(strEMPLIST) > 0 Then
        strEMPLIST = strEMPLIST & "," & RSEMPLIST("ES_EMPNBR")
    Else
        strEMPLIST = strEMPLIST & RSEMPLIST("ES_EMPNBR")
    End If
    RSEMPLIST.MoveNext
Loop
RSEMPLIST.Close

'Ticket #18721 - Delete the attachment documents first
If gsAttachment_DB Then
    glbDocName = "EdSem"
    SQLQD = "SELECT ES_EMPNBR,ES_DOCKEY FROM HREDSEM "
    SQLQD = SQLQD & TSQLQ
    rsEDSEM.Open SQLQD, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsEDSEM.EOF
        If Not IsNull(rsEDSEM("ES_DOCKEY")) And rsEDSEM("ES_DOCKEY") <> "" Then
            gdbAdoIhr001_DOC.Execute "delete from HRDOC_EDSEM where ES_TYPE='" & UCase(glbDocName) & "' AND ES_EMPNBR = " & rsEDSEM("ES_EMPNBR") & " and ES_DOCKEY=" & rsEDSEM("ES_DOCKEY") & " "
        End If
        rsEDSEM.MoveNext
    Loop
    rsEDSEM.Close
    Set rsEDSEM = Nothing
End If

'Ticket #22953 - City of Chatham-Kent but for all. Undo the Training Plan records first before deleting the Cont. Edu. records
'Course Code is required to check for Training Plan record
If Len(clpCode(0).Text) > 0 Then
    Call ContEdu_Delete_Undo_TrainingPlan(TSQLQ)
End If

TSQLQ0 = "DELETE FROM HREDSEM "
SQLQ = TSQLQ0 & TSQLQ
gdbAdoIhr001.Execute SQLQ

If dlpRenewal.Text <> "" Then
    SQLQ = "DELETE FROM HR_FOLLOW_UP WHERE EF_EMPNBR IN (" & xEmpList & ") "
    SQLQ = SQLQ & "AND EF_FDATE = " & Date_SQL(CVDate(dlpRenewal.Text)) & " "
    SQLQ = SQLQ & "AND EF_FREAS = 'EDUC' "
    gdbAdoIhr001.Execute SQLQ
End If

If IsDate(dlpDatComp.Text) Then
    If clpCode(4).Text <> "" Then
        SQLQ = "DELETE FROM HR_ATTENDANCE WHERE AD_EMPNBR IN (" & xEmpList & ") "
        SQLQ = SQLQ & "AND AD_DOA = " & Date_SQL(dlpDatComp.Text) & " "
        SQLQ = SQLQ & "AND AD_REASON = '" & clpCode(4).Text & "' "
        gdbAdoIhr001.Execute SQLQ
    End If
    
    If clpCode(5).Text <> "" Then
        SQLQ = "DELETE FROM HREMPSKL WHERE SE_EMPNBR IN (" & xEmpList & ") "
        SQLQ = SQLQ & "AND SE_DATE = " & Date_SQL(dlpDatComp.Text) & " "
        SQLQ = SQLQ & "AND SE_SKILL = '" & clpCode(5).Text & "' "
        gdbAdoIhr001.Execute SQLQ
    End If
End If

modDelRecs = True
Screen.MousePointer = DEFAULT

MsgBox "Records Deleted Successfully!"

Exit Function

CrFollow_Err:


glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DELETE RECORDS", nTable, "UPDATE TABLES")
Resume Next
End Function

Private Function modInsRecs()
Dim xEMP, xEmpList, xPayIDlist
Dim SQLQ As String, TSQLQ As String, nTable As String, xSHIFT, xSuper
Dim newline As String, Message As String
Dim Msg As String, Edat As String
Dim iRec As Integer
Dim TB As New ADODB.Recordset
Dim TF As New ADODB.Recordset
Dim TA As New ADODB.Recordset
Dim TS As New ADODB.Recordset
Dim TH As New ADODB.Recordset
Dim xDup
Dim rsDup As New ADODB.Recordset
'Dim dynHRAT As New ADODB.Recordset
Dim X As Integer, all As Integer, z As Integer
Dim recCount As Integer
Dim Response%

'Ticket #22953
Dim rsHRTrain As New ADODB.Recordset
Dim rsFollowUp As New ADODB.Recordset
Dim xTrainList
Dim xJobCode As String
Dim xTrainListID
Dim xFollowUpID
Dim xComments As String
Dim xRenewalDt
Dim xORenewalDt

Screen.MousePointer = HOURGLASS

newline = Chr(13) & Chr(10)
modInsRecs = False
all = True
'On Error GoTo CrFollow_Err 'laura nov 12, 1997

'Ticket #23510 - Opening the Follow up recordset anyways because we are computing Renewal Date using the Course
'Code Master setup if Renewal Period set for Training List - if the Renewal is not entered and for that Follow Up
'record needs to be created.
'If dlpRenewal.Text <> "" Then
    'Why retrieve all follow up records when we are only adding a new record
    'TF.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    TF.Open "SELECT * FROM HR_FOLLOW_UP WHERE 1 = 2", gdbAdoIhr001, adOpenKeyset, adLockOptimistic ', adCmdTableDirect
'End If
If clpCode(4).Text <> "" Then
    'Ticket #23510 - Not sure why we are opening these two record sets when we are not even using it.
    'TA.Open "HR_ATTENDANCE", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    
    'TH.Open "HR_JOB_HISTORY", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
End If
If clpCode(5).Text <> "" Then
    'Why retrieve all Skills records when we are only adding a new record
    'TS.Open "HREMPSKL", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    TS.Open "SELECT * FROM HREMPSKL WHERE 1 = 2", gdbAdoIhr001, adOpenKeyset, adLockOptimistic ', adCmdTableDirect
End If

'Why retrieve all Skills records when we are only adding a new record
'TB.Open "HREDSEM", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
TB.Open "SELECT * FROM HREDSEM WHERE 1 = 2", gdbAdoIhr001, adOpenKeyset, adLockOptimistic   ', adCmdTableDirect

'laura nov 12, 1997
On Error GoTo CrFollow_Err 'laura nov 12, 1997

'laura nov 12, 1997 remarked out because if the tables are not opened can't make transactions
'If TB.Transactions = False Or TA.Transactions = False Or TH.Transactions = False Or TS.Transactions = False Then
'  MsgBox "The Records Could Not Be Added!"
'  Exit Function
'End If
xEmpList = getEmpnbr(elpEEID(0))

xPayIDlist = getPayrollID(elpEEID(1))
'xPayIDlist = getPayrollID(xPayIDlist) Ticket #11696
If Len(Trim(elpEEID(0))) = 0 Then
    xEmpList = getEmpnoFromPayID(xPayIDlist)
    xEmpList = getEmpnbr(xEmpList)
End If

xDup = False
TSQLQ = "SELECT ES_EMPNBR FROM HREDSEM "
TSQLQ = TSQLQ & " WHERE ES_CTYPE = '" & clpCode(1).Text & "' "
TSQLQ = TSQLQ & " AND ES_EMPNBR IN (" & xEmpList & ")"

If Len(clpCode(0).Text) > 0 Then
    TSQLQ = TSQLQ & " AND ES_CRSCODE = '" & clpCode(0).Text & "' "
Else
    TSQLQ = TSQLQ & " AND (ES_CRSCODE = '' OR ES_CRSCODE IS NULL) "
End If

If IsDate(dlpDatComp.Text) Then
    TSQLQ = TSQLQ & " AND ES_DATCOMP = " & Date_SQL(dlpDatComp.Text)
Else
    TSQLQ = TSQLQ & " AND ES_DATCOMP IS NULL "
End If
rsDup.Open TSQLQ, gdbAdoIhr001, adOpenKeyset
If Not rsDup.EOF Then
    Msg$ = "Type: " & Chr(32) & Chr(32) & clpCode(1).Caption & Chr(10)
    Msg$ = Msg$ & "Code: " & Chr(32) & Chr(32) & IIf(clpCode(0).Text = "", "", clpCode(0).Caption) & Chr(10)
    If IsDate(dlpDatComp.Text) Then
        Msg$ = Msg$ & "Date Completed: " & Chr(32) & dlpDatComp.Text & Chr(10) & Chr(10)
    End If
    Msg$ = Msg$ & rsDup.RecordCount & " duplicates found in Continuing Education Master. " & Chr(10) & Chr(10)
    Msg$ = Msg$ & "Click Yes to post all Continuing Education records including duplicates." & Chr(10)
    Msg$ = Msg$ & "Click No to post all non-duplicate Continuing Education records." & Chr(10)
    If MsgBox(Msg$, vbYesNo, "Duplicates Found") = vbYes Then
        xDup = False
    Else
        xDup = True
        all = False
    End If
End If
If xDup Then
    xEmpList = xEmpList & ","
    Do Until rsDup.EOF
        xEmpList = Replace(xEmpList, rsDup!ES_EMPNBR & ",", "")
        rsDup.MoveNext
    Loop
    If Right(xEmpList, 1) = "," Then xEmpList = Left(xEmpList, Len(xEmpList) - 1)
End If
rsDup.Close

'Get the # of records that will be added.
recCount = getRecordCount_Add(xEmpList)
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Employee's Continuing Education Record " Else Msg$ = Msg$ & " Employee's Continuing Education Records "
    Msg$ = Msg$ & "will be Added. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, "Mass Continuing Education Add")    ' Get user response.
    If Response = IDNO Then
        TB.Close
        'Ticket #23510 - As mentioned above, will have to close this recordset
        'If dlpRenewal.Text <> "" Then TF.Close
        TF.Close
        
        If clpCode(4).Text <> "" Then
            'Ticket #23510 - Not sure why we are opening these two record sets when we are not even using it.
            'TA.Close
            'TH.Close
        End If
        If clpCode(5).Text <> "" Then TS.Close
        
        Screen.MousePointer = DEFAULT
        Exit Function
    End If
Else
    TB.Close
    
    'Ticket #23510 - As mentioned above, will have to close this recordset
    'If dlpRenewal.Text <> "" Then TF.Close
    TF.Close
    
    If clpCode(4).Text <> "" Then
        'Ticket #23510 - Not sure why we are opening these two record sets when we are not even using it.
        'TA.Close
        'TH.Close
    End If
    If clpCode(5).Text <> "" Then TS.Close
    
    Screen.MousePointer = DEFAULT
    
    MsgBox "No employee found to add the Continuing Education record."
    Exit Function
End If


strEMPLIST = xEmpList
Do Until xEmpList = ""
    
    If InStr(xEmpList, ",") = 0 Then
        xEMP = xEmpList
        xEmpList = ""
    Else
        xEMP = Left(xEmpList, InStr(xEmpList, ",") - 1)
        xEmpList = Mid(xEmpList, InStr(xEmpList, ",") + 1)
    End If
    If IsNumeric(xEMP) Then
        'Ticket #22953 - Check if Training List record exists for this Course, get Train ID, Job Code and Follow Up record
        xTrainList = ""
        xTrainListID = ""
        xJobCode = ""
        xFollowUpID = ""
        
        xTrainList = Split(Training_List_Exists(xEMP, clpCode(0).Text), "|")
        
        'Get the Job Code, Training List ID and Follow Up ID
        xTrainListID = xTrainList(0)
        xJobCode = xTrainList(1)
        xFollowUpID = xTrainList(2)
        xORenewalDt = xTrainList(3)
        
        BeginTrans      'in case you have an error the tables will not be updated
        nTable = "HREDSEM"
        
        TB.AddNew
        z = True
        TB("ES_COMPNO") = "001"
        TB("ES_EMPNBR") = xEMP
        TB("ES_CTYPE_TABL") = "ESCT"
        TB("ES_CTYPE") = clpCode(1).Text
        TB("ES_CRSCODE_TABL") = "ESCD"
        TB("ES_CRSCODE") = clpCode(0).Text
        TB("ES_COURSE") = txtCourseName
        
        'Ticket #22953 - Update Job Code if Training List exists
        If Len(xJobCode) > 0 Then
            TB("ES_JOB") = xJobCode
        End If
        
        If txtCompanyName.Text <> "" Then
            TB("ES_COMPANYNAME") = txtCompanyName.Text
        End If
        If txtTrainerName.Text <> "" Then
            TB("ES_TRAINNER") = txtTrainerName.Text
        End If
        TB("ES_ACCOUNTNO") = txtAccount
        TB("ES_TBEMP") = CDbl(medEECont(0))
        TB("ES_TBCO") = CDbl(medEECont(1))
        If IsDate(dlpDatComp.Text) Then
            TB("ES_DATCOMP") = CVDate(dlpDatComp.Text)
        End If
        TB("ES_RESULTS_TABL") = "ESRT"
        TB("ES_RESULTS") = clpCode(3).Text
        
        TB("ES_COORDINATED_TABL") = "ESCC"
        TB("ES_COORDINATED") = clpCode(6).Text
        TB("ES_METHODUSED_TABL") = "ESMU"
        TB("ES_METHODUSED") = clpCode(7).Text
        
        TB("ES_CONDUCT_TABL") = "ESCB"
        TB("ES_CONDUCT") = clpCode(2).Text
        If IsDate(dlpStartDate.Text) Then
            TB("ES_START") = CVDate(dlpStartDate.Text)
        End If
        TB("ES_OTHER") = CDbl(medEECont(2))
        TB("ES_ACCOM") = CDbl(medEECont(3))
        TB("ES_EXTNAME") = txtExtName
        TB("ES_PRESENTOR") = chkPresentor
        TB("ES_EMPCUR") = clpEmpCur.Text
        TB("ES_OTCUR") = clpOherCur.Text
        TB("ES_EMPLOYCUR") = clpEmployerCur.Text
        TB("ES_ACOMCUR") = clpAcomCur.Text
        TB("ES_TOTCUR") = clpTotCur.Text
        
        If dlpSchDate.Text <> "" Then TB("ES_SCHEDULED") = CVDate(dlpSchDate.Text)
        If txtCourseHRS = "" Then txtCourseHRS = "0.00"                 'is not a required field
        TB("ES_HOURS") = CDbl(txtCourseHRS)
        If dlpRenewal.Text <> "" Then TB("ES_RENEW") = CVDate(dlpRenewal.Text)         'is not a required field
        TB("ES_KEYWORD") = txtKeyword
        TB("ES_LDATE") = Date
        TB("ES_LTIME") = Time$
        TB("ES_LUSER") = glbUserID
        
        'Ticket #22682: Release 8.0 - for everyone
        'Frank added Apr 13, 2007 Ticket#12859
        'If glbCompSerial = "S/N - 2188W" Then 'City of Chatham-Kent
            'If Len(clpCEUType.Text) > 0 Then
            '    TB("ES_CEUTYPE") = clpCEUType.Text
            'End If
            If Len(txtCEUCred.Text) > 0 Then
                TB("ES_CEUCREDIT") = txtCEUCred.Text
            End If
        'End If
        'Ticket #24708 Franks 11/28/2013 - make this work for all customers
        If Len(clpCEUType.Text) > 0 Then
            TB("ES_CEUTYPE") = clpCEUType.Text
        End If
            
        'Ticket #22953 - Check if Renewal Period needs to be computed
        If dlpRenewal.Text <> "" Then xRenewalDt = dlpRenewal.Text
        If IsDate(dlpDatComp.Text) And Not IsDate(dlpRenewal.Text) And Len(Trim(clpCode(0).Text)) > 0 Then
            If Len(xJobCode) > 0 Then
                xRenewalDt = Compute_Renewal_Date(xEMP, dlpDatComp.Text, clpCode(0).Text, xJobCode)
            Else
                xRenewalDt = Compute_Renewal_Date(xEMP, dlpDatComp.Text, clpCode(0).Text, "")
            End If
            
            'Update the Renewal Date in the Continuing Education record being updated
            If dlpRenewal.Text = "" And IsDate(xRenewalDt) Then TB("ES_RENEW") = CVDate(xRenewalDt)
        End If
        
        TB.Update
        
        If dlpRenewal.Text <> "" Or IsDate(xRenewalDt) Then
            'Ticket #22953 - If corresponding Training List exists then update the Follow Up record else add a new one
            If Len(xTrainListID) > 0 And Len(xFollowUpID) > 0 Then
                'Update existing Follow Up record
                If Not IsNull(xFollowUpID) Then
                    'Update Follow Up record - Effective Date (new Renewal Date)
                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & xFollowUpID
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsFollowUp("EF_FDATE") = CVDate(xRenewalDt) 'CVDate(dlpRenewal.Text)
                        'rsFollowUp("EF_COMMENTS") = "Course: " & rsCourseCode("ES_CRSCODE") & " - " & GetTABLDesc("ESCD", rsCourseCode("ES_CRSCODE")) & " for Position: " & rsEmpJobs("TW_JOB")
                        rsFollowUp("EF_LDATE") = Date
                        rsFollowUp("EF_LUSER") = glbUserID
                        rsFollowUp("EF_LTIME") = Time$
                        rsFollowUp.Update
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                End If
            Else
                If Len(xTrainListID) > 0 And Len(xFollowUpID) = 0 Then
                    xComments = "Course: " & clpCode(0).Text & " "
                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & xEMP
                    If IsDate(xORenewalDt) Then
                        SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' AND EF_FDATE = " & Date_SQL(xORenewalDt)  'Date_SQL(xRenewalDt)  'Date_SQL(dlpRenewal.Text)
                    Else
                        SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' AND EF_FDATE = " & Date_SQL(xRenewalDt)  'Date_SQL(dlpRenewal.Text)
                    End If
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        xFollowUpID = rsFollowUp("EF_FOLLOWUP_ID")
                        
                        'Update the Follow Up record
                        rsFollowUp("EF_FDATE") = CVDate(xRenewalDt)    'CVDate(dlpRenewal.Text)
                        'rsFollowUp("EF_COMMENTS") = "Course: " & rsCourseCode("ES_CRSCODE") & " - " & GetTABLDesc("ESCD", rsCourseCode("ES_CRSCODE")) & " for Position: " & rsEmpJobs("TW_JOB")
                        rsFollowUp("EF_LDATE") = Date
                        rsFollowUp("EF_LUSER") = glbUserID
                        rsFollowUp("EF_LTIME") = Time$
                        rsFollowUp.Update
                    Else
                        rsFollowUp.Close
                        Set rsFollowUp = Nothing
                        SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                        SQLQ = SQLQ & " WHERE EF_EMPNBR = " & xEMP
                        SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                        SQLQ = SQLQ & " AND EF_COMPLETED <>1 ORDER BY EF_LDATE DESC"
                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsFollowUp.EOF Then
                            xFollowUpID = rsFollowUp("EF_FOLLOWUP_ID")
                        
                            'Update the Follow Up record
                            rsFollowUp("EF_FDATE") = CVDate(xRenewalDt)    'CVDate(dlpRenewal.Text)
                            'rsFollowUp("EF_COMMENTS") = "Course: " & rsCourseCode("ES_CRSCODE") & " - " & GetTABLDesc("ESCD", rsCourseCode("ES_CRSCODE")) & " for Position: " & rsEmpJobs("TW_JOB")
                            rsFollowUp("EF_LDATE") = Date
                            rsFollowUp("EF_LUSER") = glbUserID
                            rsFollowUp("EF_LTIME") = Time$
                            rsFollowUp.Update
                        End If
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                
                    If Len(xFollowUpID) = 0 Then
                        'Add new Follow Up record
                        nTable = "HR_FOLLOW_UP"
                        TF.AddNew
                        z = True
                        TF("EF_COMPNO") = "001"
                        TF("EF_EMPNBR") = xEMP
                        TF("EF_FDATE") = CVDate(xRenewalDt)    'CVDate(dlpRenewal.Text)
                        TF("EF_FREAS_TABL") = "FURE"
                        'Ticket #24257 - Do not update Admin By for them only
                        If glbCompSerial <> "S/N - 2262W" Then
                            TF("EF_ADMINBY_TABL") = "EDAB"
                            TF("EF_ADMINBY") = GetEmpData(xEMP, "ED_ADMINBY", Null)
                        End If
                        TF("EF_FREAS") = "EDUC"
                        TF("EF_COMMENTS") = txtCourseName & newline & txtExtName
                        TF("EF_LDATE") = Date
                        TF("EF_LTIME") = Time$
                        TF("EF_LUSER") = glbUserID
                        
                        'Ticket #22953 - Update Training List record with Follow Up ID, Renewal and Course Taken Date if
                        'Training List exists. Also Update the Comments in Follow Up record in the format required when
                        'Training List exists
                        If Len(xJobCode) > 0 Then
                            TF("EF_COMMENTS") = TF("EF_COMMENTS") & newline & "Course: " & clpCode(0).Text & " - " & GetTABLDesc("ESCD", clpCode(0).Text) & " for Position: " & xJobCode
                        ElseIf Len(xTrainListID) > 0 Then
                            TF("EF_COMMENTS") = TF("EF_COMMENTS") & newline & "Course: " & clpCode(0).Text & " - " & GetTABLDesc("ESCD", clpCode(0).Text)
                        End If
                        
                        TF.Update
                        
                        xFollowUpID = TF("EF_FOLLOWUP_ID")
                    End If
                Else
                    If Len(xFollowUpID) = 0 Then
                        'Add new Follow Up record
                        nTable = "HR_FOLLOW_UP"
                        TF.AddNew
                        z = True
                        TF("EF_COMPNO") = "001"
                        TF("EF_EMPNBR") = xEMP
                        TF("EF_FDATE") = CVDate(xRenewalDt)    'CVDate(dlpRenewal.Text)
                        TF("EF_FREAS_TABL") = "FURE"
                        'Ticket #24257 - Do not update Admin By for them only
                        If glbCompSerial <> "S/N - 2262W" Then
                            TF("EF_ADMINBY_TABL") = "EDAB"
                            TF("EF_ADMINBY") = GetEmpData(xEMP, "ED_ADMINBY", Null)
                        End If
                        TF("EF_FREAS") = "EDUC"
                        TF("EF_COMMENTS") = txtCourseName & newline & txtExtName
                        TF("EF_LDATE") = Date
                        TF("EF_LTIME") = Time$
                        TF("EF_LUSER") = glbUserID
                        
                        'Ticket #22953 - Update Training List record with Follow Up ID, Renewal and Course Taken Date if
                        'Training List exists. Also Update the Comments in Follow Up record in the format required when
                        'Training List exists
                        If Len(xJobCode) > 0 Then
                            TF("EF_COMMENTS") = TF("EF_COMMENTS") & newline & "Course: " & clpCode(0).Text & " - " & GetTABLDesc("ESCD", clpCode(0).Text) & " for Position: " & xJobCode
                        ElseIf Len(xTrainListID) > 0 Then
                            TF("EF_COMMENTS") = TF("EF_COMMENTS") & newline & "Course: " & clpCode(0).Text & " - " & GetTABLDesc("ESCD", clpCode(0).Text)
                        End If
                        
                        TF.Update
                        
                        xFollowUpID = TF("EF_FOLLOWUP_ID")
                    End If
                End If
            End If
        End If
        
        'Ticket #22953 - Update Training List record with Follow Up ID, Renewal and Course Taken Date if Renewal Date and
        'Training List exists. If Renewal Date do not exists then remove the respective Training List record.
        If (dlpRenewal.Text <> "" Or IsDate(xRenewalDt)) And Len(xTrainListID) > 0 Then
            nTable = "HR_TRAIN"
            
            'Update the Training List Record
            SQLQ = "SELECT * FROM HR_TRAIN"
            SQLQ = SQLQ & " WHERE TR_EMPNBR = " & xEMP
            SQLQ = SQLQ & " AND TR_ID = " & xTrainListID
            SQLQ = SQLQ & " AND TR_CRSCODE = '" & clpCode(0).Text & "'"
            rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsHRTrain.EOF Then
                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    xComments = "Course: " & clpCode(0).Text & " "
                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & xEMP
                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' AND EF_FDATE = " & Date_SQL(xRenewalDt)   'Date_SQL(dlpRenewal.Text)
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    Else
                        rsFollowUp.Close
                        Set rsFollowUp = Nothing
                        SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                        SQLQ = SQLQ & " WHERE EF_EMPNBR = " & xEMP
                        SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                        SQLQ = SQLQ & " AND EF_COMPLETED <>1 ORDER BY EF_LDATE DESC"
                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsFollowUp.EOF Then
                            rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                        Else
                            'If Not IsNull(TF("EF_FOLLOWUP_ID")) Then
                            If Not IsNull(xFollowUpID) Then
                                'rsHRTrain("TR_FOLLOWUP_ID") = TF("EF_FOLLOWUP_ID")
                                rsHRTrain("TR_FOLLOWUP_ID") = xFollowUpID
                            End If
                        End If
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                End If
                rsHRTrain("TR_RENEW") = IIf(Trim(xRenewalDt) <> "", CVDate(xRenewalDt), Null) 'IIf(Trim(dlpRenewal.Text) <> "", CVDate(dlpRenewal.Text), Null)
                If IsDate(dlpDatComp.Text) Then
                    rsHRTrain("TR_COURSE_TAKEN") = CVDate(dlpDatComp.Text)
                End If
                rsHRTrain("TR_LDATE") = Date
                rsHRTrain("TR_LUSER") = glbUserID
                rsHRTrain("TR_LTIME") = Time$
                rsHRTrain.Update
            End If
            rsHRTrain.Close
            Set rsHRTrain = Nothing
        Else
            'Remove the Training List record - No renewal of the Course
            If Len(xTrainListID) > 0 Then
                SQLQ = "SELECT * FROM HR_TRAIN"
                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & xEMP
                SQLQ = SQLQ & " AND TR_ID = " & xTrainListID
                SQLQ = SQLQ & " AND TR_CRSCODE = '" & clpCode(0).Text & "'"
                rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsHRTrain.EOF Then
                    'Retrieve the Follow Up record ID
                    If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                        xComments = "Course: " & clpCode(0).Text & " "
                        SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                        SQLQ = SQLQ & " WHERE EF_EMPNBR = " & xEMP
                        SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' AND EF_FDATE = " & Date_SQL(rsHRTrain("TR_RENEW"))  'Date_SQL(xRenewalDt)   'Date_SQL(dlpRenewal.Text)
                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsFollowUp.EOF Then
                            xFollowUpID = rsFollowUp("EF_FOLLOWUP_ID")
                        Else
                            rsFollowUp.Close
                            Set rsFollowUp = Nothing
                            SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_EMPNBR = " & xEMP
                            SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                            SQLQ = SQLQ & " AND EF_COMPLETED <>1 ORDER BY EF_LDATE DESC"
                            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsFollowUp.EOF Then
                                xFollowUpID = rsFollowUp("EF_FOLLOWUP_ID")
                            End If
                        End If
                        rsFollowUp.Close
                        Set rsFollowUp = Nothing
                    End If
                                        
                    'Delete the Training List record - no Renewal Date entered.
                    rsHRTrain.Delete
                    
                    'Ticket #30322 - Mark Follow Up record as completed
                    'Ticket #30404 - Added LDATE, LTIME, LUSER
                    If Not IsNull(xFollowUpID) And xFollowUpID <> "" Then
                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(xORenewalDt) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & xFollowUpID
                        gdbAdoIhr001.Execute SQLQ
                    End If
                End If
            End If
            'rsHRTrain.Close
            If rsHRTrain.State <> 0 Then rsHRTrain.Close 'Ticket #23135 Franks 01/24/2013
            Set rsHRTrain = Nothing
            
            'Ticket #23886 Franks 06/06/2013 - comment out the follow codes to keep the same logc as
            'what we did on employee Continuing Education screen, it caused a problem
            ''Mark Follow Up record as completed
            'If Not IsNull(xFollowUpID) And xFollowUpID <> "" Then
            '    SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(dlpDatComp.Text)
            '    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & xFollowUpID
            '    gdbAdoIhr001.Execute SQLQ
            'End If
        End If
        
        'Update Attendance
        If clpCode(4).Text <> "" Then
            nTable = "HR_ATTENDANCE"
            
            SQLQ = "INSERT INTO HR_ATTENDANCE "
            SQLQ = SQLQ & "(AD_COMPNO,AD_EMPNBR,AD_DOA,AD_REASON,AD_HRS,AD_COMM,"
            SQLQ = SQLQ & "AD_SHIFT,AD_SUPER,AD_JOB,AD_SALARY,AD_SALCD,AD_DHRS,AD_WHRS,AD_ORG,AD_INDICATOR,"
            SQLQ = SQLQ & "AD_SEN, AD_LDATE, AD_LTIME, AD_LUSER) "
            
            SQLQ = SQLQ & "SELECT '001' AS AD_COMPNO, ED_EMPNBR AS AD_EMPNBR,"
            SQLQ = SQLQ & Date_SQL(dlpDatComp.Text) & " AS AD_DOA,"
            SQLQ = SQLQ & "'" & clpCode(4).Text & "' AS AD_REASON,"
            SQLQ = SQLQ & CDbl(txtAttHrs) & " AS AD_HRS,"
            SQLQ = SQLQ & "'" & Replace(txtCourseName, "'", "''") & Chr(13) & Chr(10)
            SQLQ = SQLQ & Replace(txtExtName, "'", "''") & "' AS AD_COMM,"
            SQLQ = SQLQ & "JH_SHIFT,JH_REPTAU,JH_JOB,SH_SALARY,SH_SALCD,JH_DHRS,JH_WHRS,JH_ORG,"
            SQLQ = SQLQ & IIf(chkIncentive, 1, 0) & " AS AD_INDICATOR,"
            SQLQ = SQLQ & IIf(chkSeniority, 1, 0) & " AS AD_SEN,"
            SQLQ = SQLQ & Date_SQL(Date) & " AS AD_LDATE,"
            SQLQ = SQLQ & "'" & Time$ & "' AS AD_LTIME,"
            SQLQ = SQLQ & "'" & glbUserID & "' AS AD_LUSER "
            If glbOracle Then
                SQLQ = SQLQ & "FROM HREMP,HR_JOB_HISTORY,HR_SALARY_HISTORY WHERE HREMP.ED_EMPNBR(+)=HR_JOB_HISTORY.JH_EMPNBR "
                SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_EMPNBR(+)=HR_SALARY_HISTORY.SH_EMPNBR "
                SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_JOB(+)=HR_SALARY_HISTORY.SH_JOB "
                SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_CURRENT(+)=HR_SALARY_HISTORY.SH_CURRENT "
                SQLQ = SQLQ & "AND ED_EMPNBR=" & xEMP & " AND (JH_CURRENT<>0 OR JH_CURRENT is NULL)"
            Else
                SQLQ = SQLQ & "FROM (HREMP LEFT JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR) "
                SQLQ = SQLQ & "LEFT JOIN HR_SALARY_HISTORY ON HR_JOB_HISTORY.JH_EMPNBR=HR_SALARY_HISTORY.SH_EMPNBR "
                SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_JOB=HR_SALARY_HISTORY.SH_JOB "
                SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_CURRENT=HR_SALARY_HISTORY.SH_CURRENT "
                SQLQ = SQLQ & "WHERE ED_EMPNBR=" & xEMP & " AND (JH_CURRENT<>0 OR JH_CURRENT is NULL)"
            End If
            gdbAdoIhr001.Execute SQLQ
        End If
        
        'Update Skills
        If clpCode(5).Text <> "" Then
            nTable = "HREMPSKL"
            TS.AddNew
            z = True
            TS("SE_COMPNO") = "001"
            TS("SE_EMPNBR") = xEMP
            TS("SE_SKILL_TABL") = "EDSK"
            TS("SE_SKILL") = clpCode(5).Text
            If txtSkillsExp = "" Then txtSkillsExp = "0"         'is not a required field
            TS("SE_LEVEL") = CDbl(txtSkillsExp)
            If IsDate(dlpDatComp.Text) Then
                TS("SE_DATE") = CVDate(dlpDatComp.Text)
            End If
            TS("SE_COMM1") = txtCourseName & newline & txtExtName
            TS("SE_LDATE") = Date
            TS("SE_LTIME") = Time$
            TS("SE_LUSER") = glbUserID
            TS.Update
        End If
        
        'Ticket #18721
        If gsAttachment_DB Then
            If glbDocNewRecord Then 'New Record only
                If Len(glbDocImpFile) > 0 Then
                    glbDocKey = TB("ES_ID")
                    Call AttachmentAdd(xEMP, glbDocImpFile, glbDocType, glbDocDesc)
                End If
            End If
        End If
        
        CommitTrans
    Else
        all = False
    End If
    nTable = ""
Loop

glbDocImpFile = ""
imgNoSec.Visible = True

TB.Close

'Ticket #23510 - As mentioned above, will have to close this recordset
'If dlpRenewal.Text <> "" Then TF.Close
TF.Close

If clpCode(4).Text <> "" Then
    'Ticket #23510 - Not sure why we are opening these two record sets when we are not even using it.
    'TA.Close
    'TH.Close
End If
If clpCode(5).Text <> "" Then TS.Close

modInsRecs = True

Screen.MousePointer = DEFAULT

If all = True Then   'if entered all records
    MsgBox "Records Added Successfully!"
ElseIf all = False And z = False Then
    MsgBox "The Record(s) were Not Added!"
ElseIf all = False Then
    MsgBox "Remaining records added successfully."
Else
End If

Exit Function

CrFollow_Err:
' changed to look at database
If gdbAdoIhr001.Errors.count > 0 Then
    If gdbAdoIhr001.Errors(0).SQLState = 3022 Then      ' 3022 = duplicate record
        Message = "Cannot add Employee Skills record for employee " & xEMP & " because a duplicate record already exists."
        MsgBox Message
        all = False
        'z = False
        Err = 0   ' i know will be reset any way - but just in case
        TS.CancelUpdate
        Resume Next
        Exit Function
    End If
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADD RECORDS", nTable & "; Emp#: " & xEMP, "UPDATE TABLES")
Resume Next

End Function

Private Function Compute_Renewal_Date(xEmpnbr, xDateComplete, xCourse, xJob)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsCourseMst As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim SQLQ As String
    Dim xCurRenPrd, xPrvRenPrd
    Dim xCurRenTyp, xPrvRenTyp, xDWMY As String
    Dim flgCrsFound As Boolean
    
    'To display the renewal date on the screen
    
    'Initialise
    xCurRenTyp = ""
    xPrvRenTyp = ""
    xCurRenPrd = 0
    xPrvRenPrd = 0
    flgCrsFound = False
    
    'Find out if this Course is Unique for each Position
    'If Unique for Each Position - then retrieve Renewal Period from Required Courses screen
    'If not Unique for Each Position - then retrieve Renewal Period from Course Code Mst screen
    SQLQ = "SELECT ES_CRSCODE,ES_UNIQUE_FOR_POS,ES_RENEW_CRS_CUR,ES_CUR_PRD_DWMY, ES_RENEW_CRS_PRV,ES_PRV_PRD_DWMY FROM HR_COURSECODE_MASTER"
    SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & xCourse & "'"
    rsCourseMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsCourseMst.EOF Then
        flgCrsFound = True
        'If Unique for Each Position - retrieve from Required Courses screen
        If rsCourseMst("ES_UNIQUE_FOR_POS") <> 0 Then
            'Unique for each Position
            SQLQ = "SELECT PC_CRSCODE,PC_RENEW_CRS_CUR,PC_CUR_PRD_DWMY,PC_RENEW_CRS_PRV,PC_PRV_PRD_DWMY FROM HR_JOB_COURSE "
            SQLQ = SQLQ & " WHERE PC_JOB = '" & xJob & "'"
            SQLQ = SQLQ & " AND PC_CRSCODE = '" & xCourse & "'"
            rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsReqCourse.EOF Then
                If Not IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) And Not IsNull(rsReqCourse("PC_CUR_PRD_DWMY")) And rsReqCourse("PC_RENEW_CRS_CUR") <> 0 And rsReqCourse("PC_CUR_PRD_DWMY") <> "" Then
                    xCurRenTyp = rsReqCourse("PC_CUR_PRD_DWMY")
                    xCurRenPrd = rsReqCourse("PC_RENEW_CRS_CUR")
                End If
                If Not IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) And Not IsNull(rsReqCourse("PC_PRV_PRD_DWMY")) And rsReqCourse("PC_RENEW_CRS_PRV") <> 0 And rsReqCourse("PC_PRV_PRD_DWMY") <> "" Then
                    xPrvRenTyp = rsReqCourse("PC_PRV_PRD_DWMY")
                    xPrvRenPrd = rsReqCourse("PC_RENEW_CRS_PRV")
                End If
            Else
                flgCrsFound = False
            End If
            rsReqCourse.Close
            Set rsReqCourse = Nothing
        Else
            'Not Unique for Each Position
            If Not IsNull(rsCourseMst("ES_RENEW_CRS_CUR")) And Not IsNull(rsCourseMst("ES_CUR_PRD_DWMY")) And rsCourseMst("ES_RENEW_CRS_CUR") <> 0 And rsCourseMst("ES_CUR_PRD_DWMY") <> "" Then
                xCurRenTyp = rsCourseMst("ES_CUR_PRD_DWMY")
                xCurRenPrd = rsCourseMst("ES_RENEW_CRS_CUR")
            End If
            If Not IsNull(rsCourseMst("ES_RENEW_CRS_PRV")) And Not IsNull(rsCourseMst("ES_PRV_PRD_DWMY")) And rsCourseMst("ES_RENEW_CRS_PRV") <> 0 And rsCourseMst("ES_PRV_PRD_DWMY") <> "" Then
                xPrvRenTyp = rsCourseMst("ES_PRV_PRD_DWMY")
                xPrvRenPrd = rsCourseMst("ES_RENEW_CRS_PRV")
            End If
        End If
    Else
        flgCrsFound = False
    End If
    rsCourseMst.Close
    Set rsCourseMst = Nothing
    
    If flgCrsFound Then
        'Compute Renewal Date for Training List record and Follow Up record as well
        SQLQ = "SELECT * FROM HR_TRAIN"
        SQLQ = SQLQ & " WHERE TR_EMPNBR = " & xEmpnbr
        If xJob <> "" And Not IsNull(xJob) Then
            SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
        Else
            SQLQ = SQLQ & " AND (TR_JOB IS NULL OR TR_JOB = '')"
        End If
        SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourse & "'"
        rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsHRTrain.EOF Then
            'Find out the Type of Position to get the Renewal Period for Renewal Date computation
            If rsHRTrain("TR_POS_TYPE") = "C" Or rsHRTrain("TR_POS_TYPE") = "T" And Not IsNull(rsHRTrain("TR_POS_TYPE")) Then
                If xCurRenTyp = "" Then
                    'No Current Renewal Period and so no Renewal Date
                    'Training List and Follow Up for this course should be deleted
                    Compute_Renewal_Date = ""
                Else
                    Select Case xCurRenTyp
                        Case "D"
                            xDWMY = "d"
                        Case "W"
                            xDWMY = "ww"
                        Case "M"
                            xDWMY = "m"
                        Case "Y"
                            xDWMY = "yyyy"
                    End Select
                    Compute_Renewal_Date = DateAdd(xDWMY, xCurRenPrd, CVDate(xDateComplete))
                End If
            ElseIf rsHRTrain("TR_POS_TYPE") = "P" And Not IsNull(rsHRTrain("TR_POS_TYPE")) Then
                If xPrvRenTyp = "" Then
                    'No Previous Renewal Period and so no Renewal Date
                    'Training List and Follow Up for this course should be deleted
                    Compute_Renewal_Date = ""
                Else
                    Select Case xPrvRenTyp
                        Case "D"
                            xDWMY = "d"
                        Case "W"
                            xDWMY = "ww"
                        Case "M"
                            xDWMY = "m"
                        Case "Y"
                            xDWMY = "yyyy"
                    End Select
                    Compute_Renewal_Date = DateAdd(xDWMY, xPrvRenPrd, CVDate(xDateComplete))
                End If
            Else
                'it's an independant course, not required by any current, temp or
                'tracked positions of the employee
                'Compute the date based on the Current Renewal Period
                If xCurRenTyp = "" Then
                    'No Current Renewal Period and so no Renewal Date
                    'Training List record of this course should be deleted
                    Compute_Renewal_Date = ""
                Else
                    Select Case xCurRenTyp
                        Case "D"
                            xDWMY = "d"
                        Case "W"
                            xDWMY = "ww"
                        Case "M"
                            xDWMY = "m"
                        Case "Y"
                            xDWMY = "yyyy"
                    End Select
                    Compute_Renewal_Date = DateAdd(xDWMY, xCurRenPrd, CVDate(xDateComplete))
                End If
            End If
        End If
        rsHRTrain.Close
        Set rsHRTrain = Nothing
    Else
        Compute_Renewal_Date = ""
    End If
    
End Function

Private Sub PayrollID1_DblClick()
frmEPayrollID.Show 1
End Sub

Private Sub scrControl_Change()
scrFrame.Top = 120 - scrControl.Value
End Sub

Private Sub scrHScroll_Change()
scrFrame.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
End Sub

Private Sub txtAccount_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtAttHrs_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtCourseHRS_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtCourseHRS_LostFocus()
If glbWFC Then 'Ticket #15522
    If Not IsNumeric(txtCourseHRS) Then txtCourseHRS = 0
    If chkSal.Value Then ' glbUNION = "NONE" Or glbUNION = "EXEC" Then
        medEECont(1).Text = txtCourseHRS * 50
    Else
        medEECont(1).Text = txtCourseHRS * 35
    End If
End If
End Sub

Private Sub txtCourseName_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtExtName_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtKeyword_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub txtMain_Change()
    clpCode(0).Text = txtMain.Text
    lblDesc = Replace(GetCrsCodeDesc(txtMain), "&", "&&")
    If Len(txtMain) > 0 Then
        If lblDesc.Caption = "" Then lblDesc.Caption = "Unassigned"
        
        If clpCode(0).Caption = "Unassigned" Then
            txtCourseName = ""
        Else
            txtCourseName = Replace(clpCode(0).Caption, "&&", "&")
            Call CourseCode_Type
        End If
    End If
End Sub

Private Sub txtMain_DblClick()
    OldtxtMain = txtMain.Text
    glbCrsCode = txtMain.Text
    glbCrsCodeDesc = lblDesc.Caption
    Call Get_CourseCode(False)
    txtMain.Text = glbCrsCode
    lblDesc.Caption = glbCrsCodeDesc
End Sub

Private Sub txtMain_KeyUp(KeyCode As Integer, Shift As Integer)
    'txtMain.Text = UCase(txtMain.Text)
End Sub

Private Sub txtMain_LostFocus()
    txtMain.Text = UCase(txtMain.Text)
    If Not (OldtxtMain = txtMain.Text) Then
        If glbCrsCodeStrArr(17) = "*" Then ' * - means the changes come from Lookup
            Call SaveArrayInFields
        Else
            If Get_CourseCode_Master_Data(clpCode(1).Text, txtMain.Text) Then
                If glbCrsCodeStrArr(17) = "*" Then ' * - means the changes come from Lookup
                    Call SaveArrayInFields
                End If
            End If
        End If
    End If
End Sub

Private Sub txtSkillsExp_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub UpConttotal()
Dim X%
If IsNumeric(medEECont(0)) Or IsNumeric(medEECont(1)) Or IsNumeric(medEECont(2)) Or IsNumeric(medEECont(3)) Then
    medContTotal = Format(Val(medEECont(0)) + Val(medEECont(1)) + Val(medEECont(2)) + Val(medEECont(3)), "Currency")
Else
    medContTotal = ""
End If

End Sub

Sub CrsName_Desc()
    If clpCode(0).Caption = "Unassigned" Then
        txtCourseName = ""
    Else
        txtCourseName = clpCode(0).Caption
    End If
End Sub



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
'UpdateRight = gSec_Upd_Education_Seminars
UpdateRight = GetMassUpdateSecurities("Education_Seminars_MassUpdate", glbUserID)
End Property

Public Property Get Addable() As Boolean
Addable = True
End Property
Public Property Get Updateble() As Boolean
Updateble = False
End Property
Public Property Get Deleteble() As Boolean
Deleteble = gSec_Upd_Education_Seminars 'False
End Property

Public Property Get Printable() As Boolean
Printable = False
End Property


Private Function getWSQLQRPT() As String
'Ticket #11649, Department security removed by Bryan, redundant, this is a list of changes, whether they have security is irrelevant at this point
'getWSQLQRPT = glbSeleDeptUn
'If Len(clpDept.Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_DEPTNO} = '" & clpDept.Text & "')"
'If Len(clpDiv.Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_DIV} = '" & clpDiv.Text & "') "
'If Len(clpCode(1).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_LOC} = '" & clpCode(1).Text & "') "
'If Len(clpCode(2).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_ORG} = '" & clpCode(2).Text & "') "
'If Len(clpCode(3).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_EMP} = '" & clpCode(3).Text & "') "
'If Len(clpCode(5).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_REGION} = '" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(5).Text & "') "
'If Len(clpCode(6).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_ADMINBY} = '" & clpCode(6).Text & "') "
'If Len(clpCode(7).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_BENEFIT_GROUP} = '" & clpCode(7).Text & "') "
'If Len(clpPT.Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_PT} = '" & clpPT.Text & "') "

'getWSQLQRPT = " ({HREMP.ED_EMPNBR} IN [" & strEMPLIST & "]) "

glbiOneWhere = False    'Ticket #18383 - do not remove this otherwise, when the report gives an error.

Call glbCri_DeptUN("") 'Ticket #15514

If Len(strEMPLIST) > 0 Then
    getWSQLQRPT = " ({HREMP.ED_EMPNBR} IN [" & strEMPLIST & "]) AND " & glbstrSelCri
Else
    getWSQLQRPT = glbstrSelCri
End If
End Function

Private Sub SaveArrayInFields()
Dim K As Integer
    clpCode(1).Text = glbCrsCodeStrArr(1) 'Course Type
    clpCode(6).Text = glbCrsCodeStrArr(2) 'Co-Ordinated By
    txtCompanyName.Text = glbCrsCodeStrArr(3) '
    txtTrainerName.Text = glbCrsCodeStrArr(4) '
    txtCourseHRS.Text = glbCrsCodeStrArr(5) 'Course Hours
    medEECont(0).Text = glbCrsCodeStrArr(6) 'Employee $
    medEECont(2).Text = glbCrsCodeStrArr(7)  'Other Expenses $
    medEECont(1).Text = glbCrsCodeStrArr(8) 'Employer $
    medEECont(3).Text = glbCrsCodeStrArr(9) 'Accommodation $
    'medEECont(4).Text = glbCrsCodeStrArr(10) 'Learning Material $
    clpEmpCur.Text = glbCrsCodeStrArr(11) 'Currency
    clpOherCur.Text = glbCrsCodeStrArr(12) 'Currency
    clpEmployerCur.Text = glbCrsCodeStrArr(13) 'Currency
    clpAcomCur.Text = glbCrsCodeStrArr(14) 'Currency
    'clpLearnCur.Text = glbCrsCodeStrArr(15) 'Currency
    clpTotCur.Text = glbCrsCodeStrArr(16) 'Currency
    'Ticket #24767 Franks 12/11/2013
    clpCode(6).Text = glbCrsCodeStrArr(18) 'Coordinated By
    clpCEUType.Text = glbCrsCodeStrArr(19) 'CEU Type
    clpCode(7).Text = glbCrsCodeStrArr(20) 'Method Used
    For K = 1 To 20 '17
        glbCrsCodeStrArr(K) = ""
    Next K
End Sub

Private Function Check_EmployeeID_Security() As Boolean
    Dim xEmpNo
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ As String
    
    Check_EmployeeID_Security = False
    
    xEmpNo = getEmpnbr(elpEEID(0))
    
    If Len(xEmpNo) > 0 Then
        SQLQ = "SELECT ED_EMPNBR, ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR IN (" & xEmpNo & ") AND " & glbSeleDeptUn
    
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
        If rsEmp.EOF Then
            Check_EmployeeID_Security = False
        Else
            Check_EmployeeID_Security = True
        End If
        rsEmp.Close
    End If
    
End Function

Private Function Check_PayrollID_Security() As Boolean
    Dim xPayIDs
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ As String
    
    Check_PayrollID_Security = False
    
    xPayIDs = getPayrollID(elpEEID(1))
    
    SQLQ = "SELECT ED_EMPNBR, ED_PAYROLL_ID FROM HREMP WHERE ED_PAYROLL_ID IN (" & xPayIDs & ") AND " & glbSeleDeptUn
    
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    If rsEmp.EOF Then
        Check_PayrollID_Security = False
    Else
        Check_PayrollID_Security = True
    End If
    rsEmp.Close
    
End Function

Private Sub chkSal_Click(Value As Integer)
If glbWFC Then
    If chkSal.Value Then
        chkHrs.Value = False
    Else
        chkHrs.Value = True
    End If
End If
End Sub

Private Sub chkHrs_Click(Value As Integer)
If glbWFC Then
    If chkHrs.Value Then
        chkSal.Value = False
    Else
        chkSal.Value = True
    End If
End If
End Sub

Private Function getRecordCount_Add(xEmpList)
    Dim SQLQ As String
    Dim rsEmp As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Add = 0
    recCount = 0

    If Len(xEmpList) > 0 Then
        SQLQ = "SELECT COUNT(ED_EMPNBR) AS TOT_REC FROM HREMP WHERE "
        SQLQ = SQLQ & " ED_EMPNBR IN (" & xEmpList & ")"
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmp.EOF Then
            recCount = rsEmp("TOT_REC")
        Else
            recCount = 0
        End If
        rsEmp.Close
        Set rsEmp = Nothing
    End If
    getRecordCount_Add = recCount

End Function

Private Function getRecordCount_Delete(xEmpList)
    Dim SQLQ As String
    Dim rsEDSEM As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Delete = 0
    recCount = 0

    SQLQ = "SELECT COUNT(ES_EMPNBR) AS TOT_REC FROM HREDSEM "
    SQLQ = SQLQ & " WHERE ES_CTYPE = '" & clpCode(1).Text & "' "
    SQLQ = SQLQ & " AND ES_EMPNBR IN (" & xEmpList & ")"
    
    'Course Code
    If Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ES_CRSCODE = '" & clpCode(0).Text & "' "
    End If
    
    'Course Name
    If Len(txtCourseName.Text) > 0 Then
        SQLQ = SQLQ & " AND ES_COURSE = '" & txtCourseName.Text & "' "
    End If
    
    'Course Description
    If Len(txtExtName.Text) > 0 Then
        SQLQ = SQLQ & " AND ES_EXTNAME = '" & txtExtName.Text & "' "
    End If
    
    'Conducted By
    If Len(clpCode(2).Text) > 0 Then
        SQLQ = SQLQ & " AND ES_CONDUCT = '" & clpCode(2).Text & "' "
    End If
    
    'Results
    If Len(clpCode(3).Text) > 0 Then
        SQLQ = SQLQ & " AND ES_RESULTS = '" & clpCode(3).Text & "' "
    End If
    
    'Scheduled Date
    If IsDate(dlpSchDate.Text) Then
        SQLQ = SQLQ & " AND ES_SCHEDULED = " & Date_SQL(dlpSchDate.Text)
    End If
    
    'Start Date
    If IsDate(dlpStartDate.Text) Then
        SQLQ = SQLQ & " AND ES_START = " & Date_SQL(dlpStartDate.Text)
    End If
    
    'Date Completed
    If IsDate(dlpDatComp.Text) Then
        SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(dlpDatComp.Text)
    End If
    
    'Renewal Date
    If IsDate(dlpRenewal.Text) Then
        SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(dlpRenewal.Text)
    End If
    
    'Co-Ordinated By
    If Len(clpCode(6).Text) > 0 Then
        SQLQ = SQLQ & " AND ES_COORDINATED = '" & clpCode(6).Text & "' "
    End If
    
    'Method Used
    If Len(clpCode(7).Text) > 0 Then
        SQLQ = SQLQ & " AND ES_METHODUSED = '" & clpCode(7).Text & "' "
    End If
    
    'CEU Type
    If Len(clpCEUType.Text) > 0 Then
        SQLQ = SQLQ & " AND ES_CEUTYPE = '" & clpCEUType.Text & "' "
    End If
    
    'CEU Credit
    If Len(txtCEUCred.Text) > 0 Then
        SQLQ = SQLQ & " AND ES_CEUCREDIT = " & txtCEUCred.Text & " "
    End If

    rsEDSEM.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic ', adOpenStatic
    If Not rsEDSEM.EOF Then
        recCount = rsEDSEM("TOT_REC")
    Else
        recCount = 0
    End If
    rsEDSEM.Close
    Set rsEDSEM = Nothing
    
    getRecordCount_Delete = recCount

End Function

Private Function Training_List_Exists(xEmpnbr, xCourse)
    Dim rsHRTrain As New ADODB.Recordset
    Dim SQLQ As String
    Dim xTrainLst As String
    
    xTrainLst = "" & "|" & "" & "|" & "" & "|" & ""
    
    SQLQ = "SELECT * FROM HR_TRAIN"
    SQLQ = SQLQ & " WHERE TR_EMPNBR = " & xEmpnbr
    SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourse & "'"
    rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRTrain.EOF Then
        'Return Training List ID, Job Code and Follow Up ID
        xTrainLst = rsHRTrain("TR_ID") & "|" & IIf(IsNull(rsHRTrain("TR_JOB")) Or rsHRTrain("TR_JOB") = "", "", rsHRTrain("TR_JOB")) & "|" & IIf(IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Or rsHRTrain("TR_FOLLOWUP_ID") = "", "", rsHRTrain("TR_FOLLOWUP_ID")) & "|" & IIf(IsNull(rsHRTrain("TR_RENEW")) Or rsHRTrain("TR_RENEW") = "", "", rsHRTrain("TR_RENEW"))
    End If
    rsHRTrain.Close
    Set rsHRTrain = Nothing
    
    Training_List_Exists = xTrainLst
    
End Function

Private Sub ContEdu_Delete_Undo_TrainingPlan(xSQL As String)
    'Retrieve the Continuing Education records and for each employee and call procedure to undo the Training Plan
    
    Dim rsCondEduDel As New ADODB.Recordset
    Dim SQLQDel As String
    
    SQLQDel = "SELECT * FROM HREDSEM "
    SQLQDel = SQLQDel & xSQL
    rsCondEduDel.Open SQLQDel, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsCondEduDel.EOF Then
        Do While Not rsCondEduDel.EOF
            If IsDate(rsCondEduDel("ES_RENEW")) And IsDate(rsCondEduDel("ES_DATCOMP")) Then
                Call Undo_Training_List_Rec_on_ContEdu_Delete(rsCondEduDel("ES_EMPNBR"), rsCondEduDel("ES_ID"), rsCondEduDel("ES_JOB"), rsCondEduDel("ES_CRSCODE"), rsCondEduDel("ES_RENEW"), rsCondEduDel("ES_DATCOMP"))
            ElseIf IsDate(rsCondEduDel("ES_DATCOMP")) Then
                Call Undo_Training_List_Rec_on_ContEdu_Delete(rsCondEduDel("ES_EMPNBR"), rsCondEduDel("ES_ID"), rsCondEduDel("ES_JOB"), rsCondEduDel("ES_CRSCODE"))
            End If
            
            rsCondEduDel.MoveNext
        Loop
    End If
    rsCondEduDel.Close
    Set rsCondEduDel = Nothing

End Sub

Private Sub Undo_Training_List_Rec_on_ContEdu_Delete(xEmpnbr, xESID, xJob, xCourse, Optional xRenewalDt, Optional xCompleteDt)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsCourseMst As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsEmpJobs As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim SQLQ As String
    Dim xEduRec, xCurRenPrd, xPrvRenPrd, xFlwRenPrd As Integer
    Dim xCurRenTyp, xPrvRenTyp, xFlwRenTyp, xDWMY As String
    Dim flgCrsTakenBefore, flgUnqForPos As Boolean
    Dim xComments As String
    Dim xOrgDate As Date
    

    'Initialise
    xEduRec = 0
    xFlwRenPrd = 0
    xFlwRenTyp = ""
    flgCrsTakenBefore = False
    
    'Course Record being deleted is the one WITH Course Renewal Date
    If Not IsMissing(xRenewalDt) Then
    
        'Check if the course is unique for each position
        SQLQ = "SELECT ES_CRSCODE,ES_UNIQUE_FOR_POS FROM HR_COURSECODE_MASTER"
        SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & xCourse & "'"
        rsCourseMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsCourseMst.EOF Then
            'Course found
            flgUnqForPos = IIf(IsNull(rsCourseMst("ES_UNIQUE_FOR_POS")), False, rsCourseMst("ES_UNIQUE_FOR_POS"))
        Else
            flgUnqForPos = False
        End If
        rsCourseMst.Close
        Set rsCourseMst = Nothing
    
        SQLQ = "SELECT * FROM HR_TRAIN"
        SQLQ = SQLQ & " WHERE TR_EMPNBR = " & xEmpnbr
        If xJob <> "" And Not IsNull(xJob) Then
            SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
        Else
            SQLQ = SQLQ & " AND (TR_JOB IS NULL OR TR_JOB = '')"
        End If
        SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourse & "'"
        rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsHRTrain.EOF Then
            'Check if Course was taken previously
            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE, ES_DATCOMP, ES_RENEW FROM HREDSEM"
            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
            SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourse & "'"
            If flgUnqForPos <> 0 Then
                'Unique for each position course then check if the course was taken for the right position
                If xJob <> "" And Not IsNull(xJob) Then
                    SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
                End If
            End If
            SQLQ = SQLQ & " AND ES_ID <> " & xESID
            SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsContEdu.EOF Then
                'Course Taken records found
                
                rsContEdu.MoveFirst
            
                'Update Training List with the last Course Taken Date and Renewal Date.
                If Not IsNull(rsContEdu("ES_RENEW")) Then
                    rsHRTrain("TR_RENEW") = CVDate(rsContEdu("ES_RENEW"))
                Else
                    rsHRTrain("TR_RENEW") = CVDate(xCompleteDt)
                End If
                rsHRTrain("TR_COURSE_TAKEN") = IIf(Not IsNull(rsContEdu("ES_DATCOMP")), CVDate(rsContEdu("ES_DATCOMP")), Null)
                rsHRTrain("TR_LDATE") = Date
                rsHRTrain("TR_LUSER") = glbUserID
                rsHRTrain("TR_LTIME") = Time$
                
                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & xEmpnbr
                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' AND EF_FDATE = " & Date_SQL(rsContEdu("ES_RENEW"))
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                End If
                
                rsHRTrain.Update
                
                If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    'Update Follow Up record - Effective Date
                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                        'rsFollowUp("EF_COMMENTS") = "Course: " & rsCourseCode("ES_CRSCODE") & " - " & GetTABLDesc("ESCD", rsCourseCode("ES_CRSCODE")) & " for Position: " & rsEmpJobs("TW_JOB")
                        rsFollowUp("EF_LDATE") = Date
                        rsFollowUp("EF_LUSER") = glbUserID
                        rsFollowUp("EF_LTIME") = Time$
                        rsFollowUp.Update
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                End If
            Else
                'No Course Taken records found - (the one being deleted was the first record).
                'Get the Renewal Period and recompute the Renewal Date
                
                If xJob = "" Or IsNull(xJob) Then
                    'If Independant Course - Reset the Renewal Date to Follow Up Period + Position Start Date
                    'Retrieve Renewal Periods from Course Code Master because this is an independant course
                    SQLQ = "SELECT ES_CRSCODE,ES_RENEW_FOLLOWUP,ES_FLWUP_PRD_DWMY FROM HR_COURSECODE_MASTER"
                    SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & xCourse & "'"
                    rsCourseMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsCourseMst.EOF Then
                        'Course found
                        xFlwRenPrd = rsCourseMst("ES_RENEW_FOLLOWUP")
                        xFlwRenTyp = rsCourseMst("ES_FLWUP_PRD_DWMY")
                    End If
                    rsCourseMst.Close
                    Set rsCourseMst = Nothing
                Else
                    'Retrieve renewal period from Required Courses table
                    SQLQ = "SELECT PC_CRSCODE,PC_RENEW_FOLLOWUP,PC_FLWUP_PRD_DWMY FROM HR_JOB_COURSE "
                    SQLQ = SQLQ & " WHERE PC_JOB = '" & xJob & "'"
                    SQLQ = SQLQ & " AND PC_CRSCODE = '" & xCourse & "'"
                    rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsReqCourse.EOF Then
                        'Course found
                        xFlwRenPrd = rsReqCourse("PC_RENEW_FOLLOWUP")
                        xFlwRenTyp = rsReqCourse("PC_FLWUP_PRD_DWMY")
                    End If
                    rsReqCourse.Close
                    Set rsReqCourse = Nothing
                End If
                    
                'Compute Renewal Date
                'Course never taken before - Renewal Date = Follow Up Period + Position Start Date
                Select Case xFlwRenTyp
                    Case "D"
                        xDWMY = "d"
                    Case "W"
                        xDWMY = "ww"
                    Case "M"
                        xDWMY = "m"
                    Case "Y"
                        xDWMY = "yyyy"
                End Select
                
                If IsDate(rsHRTrain("TR_RENEW")) Then
                    xOrgDate = rsHRTrain("TR_RENEW")
                End If
                If Not IsNull(rsHRTrain("TR_SDATE")) Then
                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xFlwRenPrd, CVDate(rsHRTrain("TR_SDATE")))
                Else
                    rsHRTrain("TR_RENEW") = CVDate(xCompleteDt)
                End If
                
                rsHRTrain("TR_COURSE_TAKEN") = Null     'Course never taken before
                rsHRTrain("TR_LDATE") = Date
                rsHRTrain("TR_LUSER") = glbUserID
                rsHRTrain("TR_LTIME") = Time$
                
                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & xEmpnbr
                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                    If IsDate(xOrgDate) Then SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(xOrgDate)
                    SQLQ = SQLQ & " ORDER BY EF_FDATE DESC"
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsFollowUp.MoveFirst
                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                End If
                
                rsHRTrain.Update
                
                If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    'Update Follow Up record - Effective Date
                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                        'rsFollowUp("EF_COMMENTS") = "Course: " & rsCourseCode("ES_CRSCODE") & " - " & GetTABLDesc("ESCD", rsCourseCode("ES_CRSCODE")) & " for Position: " & rsEmpJobs("TW_JOB")
                        rsFollowUp("EF_LDATE") = Date
                        rsFollowUp("EF_LUSER") = glbUserID
                        rsFollowUp("EF_LTIME") = Time$
                        rsFollowUp.Update
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                End If
            End If
            rsContEdu.Close
            Set rsContEdu = Nothing
        End If
        rsHRTrain.Close
        Set rsHRTrain = Nothing
    Else
        'No Course Renewal Date means there is no corresponding Training List record and Follow Up record,
        'which means we may have to create one.
        'Check if course was taken before
            '- if not taken, then check which Current or Tracked Position require this course
                'if Position found - get the Follow Up Renewal period for that Position Course
                'if Position not found - do not add Training List record for this course.
            '- If taken, then check which Current or Tracked Position require this course
                'if Position found - based on the type of Position, Current or Tracked, get the
                    'Renewal Period and compute Renewasl Date based on Course Taken date
                'if Position not found - do not add Training List record for this course
                    'clear the Renewal Date from the last Course Taken record.
                
        'Check if the course is unique for each position
        SQLQ = "SELECT ES_CRSCODE,ES_UNIQUE_FOR_POS FROM HR_COURSECODE_MASTER"
        SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & xCourse & "'"
        rsCourseMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsCourseMst.EOF Then
            'Course found
            flgUnqForPos = IIf(IsNull(rsCourseMst("ES_UNIQUE_FOR_POS")), False, rsCourseMst("ES_UNIQUE_FOR_POS"))
        Else
            flgUnqForPos = False
        End If
        rsCourseMst.Close
        Set rsCourseMst = Nothing
        
        'Course Taken before?
        SQLQ = "SELECT ES_EMPNBR,ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
        SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
        SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourse & "'"
        If flgUnqForPos <> 0 Then
            'Unique for each position course then check if the course was taken for the right position
            If xJob <> "" And Not IsNull(xJob) Then
                SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
            End If
        End If
        SQLQ = SQLQ & " AND ES_ID <> " & xESID
        SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
        rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsContEdu.EOF Then
            'Course Taken before
            rsContEdu.MoveFirst
            flgCrsTakenBefore = True
        Else
            flgCrsTakenBefore = False
        End If
        
        
        'Check which Current or Tracked Position required this Course
        'Get list of Current/Temporary and Tracked Positions of this employee
        SQLQ = "SELECT JH_ID AS TW_ID, JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
        SQLQ = SQLQ & " WHERE JH_EMPNBR = " & xEmpnbr & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
        SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & xCourse & "')"
        
        'if Unique for each Position record that means each Current/Tracked Job requiring this
        'course will have it's own Training List record - of course depending on the Renewal Period
        'Retrieve only the job assigned to the deleted Course Taken record.
        If flgUnqForPos <> 0 Then
            If xJob <> "" And Not IsNull(xJob) Then
                SQLQ = SQLQ & " AND JH_JOB = '" & xJob & "'"
            End If
        End If
        
        SQLQ = SQLQ & " UNION "
        SQLQ = SQLQ & " SELECT TW_ID, TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
        SQLQ = SQLQ & " WHERE TW_EMPNBR = " & xEmpnbr & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
        SQLQ = SQLQ & " AND TW_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & xCourse & "')"
        
        'if Unique for each Position record that means each Current/Tracked Job requiring this
        'course will have it's own Training List record - of course depending on the Renewal Period
        'Retrieve only the job assigned to the deleted Course Taken record.
        If flgUnqForPos <> 0 Then
            If xJob <> "" And Not IsNull(xJob) Then
                SQLQ = SQLQ & " AND TW_JOB = '" & xJob & "'"
            End If
        End If
        
        SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
        rsEmpJobs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmpJobs.EOF Then
            rsEmpJobs.MoveFirst
        
            Do While Not rsEmpJobs.EOF
                'Get the renewal periods of the course
                SQLQ = "SELECT PC_CRSCODE,PC_RENEW_CRS_CUR,PC_CUR_PRD_DWMY,PC_RENEW_CRS_PRV,PC_PRV_PRD_DWMY,PC_RENEW_FOLLOWUP,PC_FLWUP_PRD_DWMY FROM HR_JOB_COURSE "
                SQLQ = SQLQ & " WHERE PC_JOB = '" & rsEmpJobs("TW_JOB") & "'"
                SQLQ = SQLQ & " AND PC_CRSCODE = '" & xCourse & "'"
                rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsReqCourse.EOF Then
                    'Course found
                    xCurRenPrd = rsReqCourse("PC_RENEW_CRS_CUR")
                    xCurRenTyp = rsReqCourse("PC_CUR_PRD_DWMY")
                    xPrvRenPrd = rsReqCourse("PC_RENEW_CRS_PRV")
                    xPrvRenTyp = rsReqCourse("PC_PRV_PRD_DWMY")
                    xFlwRenPrd = rsReqCourse("PC_RENEW_FOLLOWUP")
                    xFlwRenTyp = rsReqCourse("PC_FLWUP_PRD_DWMY")
                End If
                rsReqCourse.Close
                Set rsReqCourse = Nothing
                
                'if Unique for each Position Course check if the Training List existing for this Job
                'already exists - then skip to next Employee Position requiring this course
                If flgUnqForPos <> 0 And xJob <> "" And Not IsNull(xJob) Then
                    SQLQ = "SELECT * FROM HR_TRAIN"
                    SQLQ = SQLQ & " WHERE TR_EMPNBR = " & xEmpnbr
                    SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourse & "'"
                    rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsHRTrain.EOF Then
                        'Skip to next Employee Job because for this Job, the training list record already
                        'exist for this unique for each position course.
                        GoTo next_EmpPosition
                    Else
                        'Continue with the rest of the process
                    End If
                    rsHRTrain.Close
                    Set rsHRTrain = Nothing
                End If
                
                SQLQ = "SELECT * FROM HR_TRAIN WHERE 1=2"
                rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                
                'Course Taken before?
                If flgCrsTakenBefore = True Then
                    'Course Taken before
                    'Compute the Renewal Period for this course and add a Training List and Follow Up record
                    If (rsEmpJobs("POS_TYPE") = "CURRENT" Or rsEmpJobs("POS_TYPE") = "TEMPORARY") And rsEmpJobs("TW_CURRENT") Then
                        'Primary Current/Temporary Position
                        'Based on Current Renewal Period if found
                        If IsNull(xCurRenPrd) Or xCurRenPrd = 0 Or xCurRenPrd = "" Then
                            'No Renewal Period found, clear last course taken record's Renewal Date
                            'There won't be Training List record, because there was no Renewal Date on the
                            'deleted Course Taken record.
                            rsContEdu("ES_RENEW") = Null
                            rsContEdu("ES_LDATE") = Date
                            rsContEdu("ES_LUSER") = glbUserID
                            rsContEdu("ES_LTIME") = Time$
                            rsContEdu.Update
                        
                            'If flgUnqForPos Then
                            '    'Go to next position
                            '    GoTo next_EmpPosition
                            'Else
                                'Exit loop - only the first position gets this course
                                Exit Do
                           ' End If
                        Else
                            'Compute renewal date
                            Select Case xCurRenTyp
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            'Add a new Training List record with Renewal Date based on Current Renewal Period and
                            'Course Taken Date
                            rsHRTrain.AddNew
                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xCurRenPrd, CVDate(rsContEdu("ES_DATCOMP")))
                            rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                            
                            'Update Continuing Education record as well
                            rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                            rsContEdu("ES_JOB") = rsEmpJobs("TW_JOB")
                            rsContEdu("ES_LDATE") = Date
                            rsContEdu("ES_LUSER") = glbUserID
                            rsContEdu("ES_LTIME") = Time$
                            rsContEdu.Update
                        End If
                    Else
                        'Previous position
                        'Based on Previous Renewal period if found
                        If IsNull(xPrvRenPrd) Or xPrvRenPrd = 0 Or xPrvRenPrd = "" Then
                            'No Renewal Period found, clear last course taken record's Renewal Date
                            'There won't be Training List record, because there was no Renewal Date on the
                            'deleted Course Taken record.
                            rsContEdu("ES_RENEW") = Null
                            rsContEdu("ES_LDATE") = Date
                            rsContEdu("ES_LUSER") = glbUserID
                            rsContEdu("ES_LTIME") = Time$
                            rsContEdu.Update
                            
                            Exit Do
                        Else
                            'Compute renewal date
                            Select Case xPrvRenTyp
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            'Add a new Training List record with Renewal Date based on Prev Renewal Period
                            'Course Taken Date
                            rsHRTrain.AddNew
                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xPrvRenPrd, CVDate(rsContEdu("ES_DATCOMP")))
                            rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                            
                            'Update Continuing Education record as well
                            rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                            rsContEdu("ES_JOB") = rsEmpJobs("TW_JOB")
                            rsContEdu("ES_LDATE") = Date
                            rsContEdu("ES_LUSER") = glbUserID
                            rsContEdu("ES_LTIME") = Time$
                            rsContEdu.Update
                        End If
                    End If
                Else
                    'Course not taken before
                    'Compute renewal date based on Follow Up Period
                    Select Case xFlwRenTyp
                        Case "D"
                            xDWMY = "d"
                        Case "W"
                            xDWMY = "ww"
                        Case "M"
                            xDWMY = "m"
                        Case "Y"
                            xDWMY = "yyyy"
                    End Select
                    'Add a new Training List record with Renewal Date based on Follow Up Period
                    rsHRTrain.AddNew
                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xFlwRenPrd, CVDate(rsEmpJobs("TW_SDATE")))
                End If
                
                rsHRTrain("TR_COMPNO") = "001"
                rsHRTrain("TR_EMPNBR") = xEmpnbr
                rsHRTrain("TR_CRSCODE") = xCourse
                
                rsHRTrain("TR_JOB") = rsEmpJobs("TW_JOB")
                rsHRTrain("TR_SDATE") = rsEmpJobs("TW_SDATE")
                If (rsEmpJobs("POS_TYPE") = "CURRENT") And rsEmpJobs("TW_CURRENT") Then
                    rsHRTrain("TR_POS_TYPE") = "C"
                ElseIf (rsEmpJobs("POS_TYPE") = "TEMPORARY") And rsEmpJobs("TW_CURRENT") Then
                    rsHRTrain("TR_POS_TYPE") = "T"
                Else
                    rsHRTrain("TR_POS_TYPE") = "P"
                End If
                rsHRTrain("TR_LDATE") = Date
                rsHRTrain("TR_LTIME") = Time$
                rsHRTrain("TR_LUSER") = glbUserID
                
                'Add a Follow Up record for this Training course
                SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE 1 = 2"
                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                rsFollowUp.AddNew
                rsFollowUp("EF_COMPNO") = "001"
                rsFollowUp("EF_EMPNBR") = xEmpnbr
                rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                rsFollowUp("EF_FREAS_TABL") = "FURE"
                'Ticket #24257 - Do not update Admin By for them only
                If glbCompSerial <> "S/N - 2262W" Then
                    rsFollowUp("EF_ADMINBY_TABL") = "EDAB"
                    rsFollowUp("EF_ADMINBY") = GetEmpData(xEmpnbr, "ED_ADMINBY", Null)
                End If
                rsFollowUp("EF_FREAS") = "EDUC"
                rsFollowUp("EF_COMMENTS") = "Course: " & xCourse & " - " & GetTABLDesc("ESCD", xCourse) & " for Position: " & rsEmpJobs("TW_JOB")
                rsFollowUp("EF_LDATE") = Date
                rsFollowUp("EF_LTIME") = Time$
                rsFollowUp("EF_LUSER") = glbUserID
                rsFollowUp.Update
                
                rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                rsHRTrain.Update
                
                rsFollowUp.Close
                Set rsFollowUp = Nothing
            
                'Update Position record with Follow Up ID
                'if the course code is TRAIN
                If xCourse = "TRAIN" Then
                    'Search HR_JOB_HISTORY or HR_TEMP_WORK table for this Position record
                    'and update with Follow Up Id
                    If (rsEmpJobs("POS_TYPE") = "CURRENT") Then
                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJobs("TW_ID")
                    Else
                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & rsEmpJobs("TW_ID")
                    End If
                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsTJob.EOF Then
                        If (rsEmpJobs("POS_TYPE") = "CURRENT") Then
                            rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                        Else
                            rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                        End If
                        rsTJob.Update
                    End If
                    rsTJob.Close
                    Set rsTJob = Nothing
                End If
                
                rsHRTrain.Close
                Set rsHRTrain = Nothing
                'If flgUnqForPos Then
                '    'Go to next position
                '    GoTo next_EmpPosition
                'Else
                    'Exit loop - only the first position gets this course
                    Exit Do
                'End If
                
next_EmpPosition:
                rsEmpJobs.MoveNext
            Loop
        Else
            'No Current/Temporary or Tracked Positions require this course
            
            'Course Taken before?
            If flgCrsTakenBefore = True Then
                'Course Taken before
                'Clear renewal date if found in the last course taken record
                'There won't be Training List record because the Course Taken record been deleted does
                'not have Renewal Date.
                rsContEdu("ES_RENEW") = Null
                rsContEdu("ES_LDATE") = Date
                rsContEdu("ES_LUSER") = glbUserID
                rsContEdu("ES_LTIME") = Time$
                rsContEdu.Update
            Else
                'Do not do anything, just let the Cont Education record delete.
                'No Current or Tracked Position of this employee require this course.
            End If
            
        End If
        rsEmpJobs.Close
        Set rsEmpJobs = Nothing
        
        rsContEdu.Close
        Set rsContEdu = Nothing
    End If

End Sub

Private Sub WFCScreenSetup() 'Ticket #24767 Franks 12/11/2013
'If glbWFC Then 'Ticket #15818
    lblTitle(21).FontBold = True
    lblTitle(4).FontBold = True
    lblTitle(5).FontBold = True
    lblTitle(17).FontBold = True
    frmHrlSal.Visible = True
    frmHrlSal.BorderStyle = 0
    'Ticket #24767 Franks 12/11/2013 - "   The fields are read only
    lblTitle(1).Enabled = False 'Course Type
    clpCode(1).Enabled = False
    lblTitle(2).Enabled = False 'Course Name
    txtCourseName.Enabled = False
    lblTitle(0).Enabled = False 'Coordinated By
    clpCode(6).Enabled = False
    lblTitle(24).Enabled = False 'Method Used
    clpCode(7).Enabled = False
    lblCEUType.Enabled = False 'CEU Type
    clpCEUType.Enabled = False
    'Ticket #24767 Franks 12/11/2013 - end
'End If
End Sub
