VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEHSCause 
   AutoRedraw      =   -1  'True
   Caption         =   "Root Causes Data"
   ClientHeight    =   10770
   ClientLeft      =   255
   ClientTop       =   765
   ClientWidth     =   13755
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   ScaleHeight     =   10770
   ScaleWidth      =   13755
   WindowState     =   2  'Maximized
   Begin VB.Frame frInvestigation 
      Caption         =   "Investigation Team"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   9240
      TabIndex        =   53
      Top             =   4440
      Visible         =   0   'False
      Width           =   3900
      Begin VB.TextBox txtInvMemName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   2760
         TabIndex        =   83
         Tag             =   "00-Employee Name of individual's supervisor"
         Top             =   1830
         Visible         =   0   'False
         Width           =   1700
      End
      Begin VB.TextBox txtInvMemName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   2760
         TabIndex        =   82
         Tag             =   "00-Employee Name of individual's supervisor"
         Top             =   1470
         Visible         =   0   'False
         Width           =   1700
      End
      Begin VB.TextBox txtInvMemName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   2760
         TabIndex        =   81
         Tag             =   "00-Employee Name of individual's supervisor"
         Top             =   1110
         Visible         =   0   'False
         Width           =   1700
      End
      Begin VB.TextBox txtInvMemName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   2760
         TabIndex        =   80
         Tag             =   "00-Employee Name of individual's supervisor"
         Top             =   750
         Visible         =   0   'False
         Width           =   1700
      End
      Begin VB.TextBox txtInvMemName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   79
         Tag             =   "00-Employee Name of individual's supervisor"
         Top             =   390
         Visible         =   0   'False
         Width           =   1700
      End
      Begin INFOHR_Controls.EmployeeLookup elpInvMem 
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   62
         Tag             =   "00-Employee Number"
         Top             =   1800
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin VB.TextBox txtInvMemSName 
         Appearance      =   0  'Flat
         DataField       =   "RC_INVEST5_SURNAME"
         Height          =   285
         Index           =   4
         Left            =   3240
         TabIndex        =   78
         Tag             =   "00-Employee Surname Name of individual's supervisor"
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMemSName 
         Appearance      =   0  'Flat
         DataField       =   "RC_INVEST4_SURNAME"
         Height          =   285
         Index           =   3
         Left            =   3240
         TabIndex        =   77
         Tag             =   "00-Employee Surname Name of individual's supervisor"
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMemSName 
         Appearance      =   0  'Flat
         DataField       =   "RC_INVEST3_SURNAME"
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   76
         Tag             =   "00-Employee Surname Name of individual's supervisor"
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMemSName 
         Appearance      =   0  'Flat
         DataField       =   "RC_INVEST2_SURNAME"
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   75
         Tag             =   "00-Employee Surname Name of individual's supervisor"
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMemSName 
         Appearance      =   0  'Flat
         DataField       =   "RC_INVEST1_SURNAME"
         Height          =   285
         Index           =   0
         Left            =   3240
         TabIndex        =   74
         Tag             =   "00-Employee Surname Name of individual's supervisor"
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMemFName 
         Appearance      =   0  'Flat
         DataField       =   "RC_INVEST5_FNAME"
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   73
         Tag             =   "00-Employee First Name of individual's supervisor"
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMemFName 
         Appearance      =   0  'Flat
         DataField       =   "RC_INVEST4_FNAME"
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   72
         Tag             =   "00-Employee First Name of individual's supervisor"
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMemFName 
         Appearance      =   0  'Flat
         DataField       =   "RC_INVEST3_FNAME"
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   71
         Tag             =   "00-Employee First Name of individual's supervisor"
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMemFName 
         Appearance      =   0  'Flat
         DataField       =   "RC_INVEST2_FNAME"
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   70
         Tag             =   "00-Employee First Name of individual's supervisor"
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMemFName 
         Appearance      =   0  'Flat
         DataField       =   "RC_INVEST1_FNAME"
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   69
         Tag             =   "00-Employee First Name of individual's supervisor"
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin INFOHR_Controls.EmployeeLookup elpInvMem 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   54
         Tag             =   "00-Employee Number"
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.EmployeeLookup elpInvMem 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   56
         Tag             =   "00-Employee Number"
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.EmployeeLookup elpInvMem 
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   58
         Tag             =   "00-Employee Number"
         Top             =   1080
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.EmployeeLookup elpInvMem 
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   60
         Tag             =   "00-Employee Number"
         Top             =   1440
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin VB.TextBox txtInvMem 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   64
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMem 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   65
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMem 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   66
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMem 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   67
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMem 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   68
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Memeber 5"
         Height          =   195
         Index           =   21
         Left            =   120
         TabIndex        =   63
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Memeber 4"
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   61
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Memeber 3"
         Height          =   195
         Index           =   19
         Left            =   120
         TabIndex        =   59
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Memeber 2"
         Height          =   195
         Index           =   18
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   915
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Memeber 1"
         Height          =   195
         Index           =   17
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame frRootCauseDesc 
      Height          =   975
      Left            =   9240
      TabIndex        =   50
      Top             =   7200
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox txtRCDesc 
         Appearance      =   0  'Flat
         Height          =   885
         Left            =   2640
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Tag             =   "00-Root Cause Description"
         Top             =   0
         Width           =   6255
      End
      Begin VB.Label lblTitle 
         Caption         =   "Root Cause Description"
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
         Index           =   16
         Left            =   240
         TabIndex        =   52
         Top             =   0
         Width           =   2355
      End
   End
   Begin VB.Frame frProblemDesc 
      Height          =   375
      Left            =   9240
      TabIndex        =   49
      Top             =   6720
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox txtProDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "00-Problem Description"
         Top             =   0
         Width           =   6090
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Problem Description"
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
         Index           =   15
         Left            =   240
         TabIndex        =   51
         Top             =   10
         Width           =   2355
      End
   End
   Begin VB.Frame frComments 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      TabIndex        =   31
      Top             =   7200
      Width           =   9135
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy "
         Height          =   330
         Left            =   240
         TabIndex        =   84
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtComments 
         Appearance      =   0  'Flat
         DataField       =   "RC_Comments"
         Height          =   1005
         Left            =   2640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Tag             =   "00-Comments"
         Top             =   60
         Width           =   6255
      End
      Begin VB.Label lblTitle 
         Caption         =   "Comments"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   36
         Top             =   60
         Width           =   1305
      End
      Begin VB.Label lblUpdateDate 
         Caption         =   "Updated Date"
         Height          =   255
         Left            =   6000
         TabIndex        =   35
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label lblUpdateBy 
         Caption         =   "Updated By"
         Height          =   255
         Left            =   2640
         TabIndex        =   34
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label lblUpdDateDesc 
         Height          =   255
         Left            =   7200
         TabIndex        =   33
         Top             =   1140
         Width           =   1935
      End
      Begin VB.Label lblUserDesc 
         Height          =   255
         Left            =   3600
         TabIndex        =   32
         Top             =   1140
         Width           =   2295
      End
   End
   Begin VB.Frame frSecondary 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      TabIndex        =   39
      Top             =   5040
      Visible         =   0   'False
      Width           =   8895
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   2580
         TabIndex        =   9
         Tag             =   "00-Type of Event"
         Top             =   240
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECRE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   2580
         TabIndex        =   10
         Tag             =   "00-Immediate / Direct Causes"
         Top             =   600
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECRI"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   7
         Left            =   2580
         TabIndex        =   11
         Tag             =   "00-Basic / Underlying Causes"
         Top             =   1200
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECRA"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   8
         Left            =   2580
         TabIndex        =   12
         Tag             =   "00-Basic / Underlying Causes"
         Top             =   1560
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECRB"
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Basic/Root Cause"
         Height          =   195
         Index           =   12
         Left            =   0
         TabIndex        =   45
         Top             =   960
         Width           =   2355
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Immediate/Direct Cause"
         Height          =   195
         Index           =   11
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   2355
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Basic / Underlying Causes"
         Height          =   195
         Index           =   10
         Left            =   360
         TabIndex        =   43
         Top             =   1605
         Width           =   2355
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Categories"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   42
         Top             =   1245
         Width           =   2355
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Immediate / Direct Causes"
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   41
         Top             =   645
         Width           =   2355
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Event"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   40
         Top             =   285
         Width           =   2355
      End
   End
   Begin VB.Frame frPrimary 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      TabIndex        =   26
      Top             =   3360
      Width           =   8895
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   2340
         TabIndex        =   5
         Tag             =   "00-Type of Event"
         Top             =   0
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECRE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   2340
         TabIndex        =   6
         Tag             =   "00-Immediate / Direct Causes"
         Top             =   360
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECRI"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   2340
         TabIndex        =   7
         Tag             =   "00-Basic / Underlying Causes"
         Top             =   720
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECRA"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   5
         Left            =   2340
         TabIndex        =   8
         Tag             =   "00-Basic / Underlying Causes"
         Top             =   1080
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECRB"
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Basic/Root Cause"
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
         Index           =   14
         Left            =   0
         TabIndex        =   47
         Top             =   960
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Immediate/Direct Cause"
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
         Index           =   13
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Event"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   45
         Width           =   2355
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Immediate / Direct Causes"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   405
         Width           =   2355
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Categories"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   28
         Top             =   765
         Width           =   2355
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Basic / Underlying Causes"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   1125
         Width           =   2355
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "Fehscase.frx":0000
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "Fehscase.frx":0014
      TabIndex        =   0
      Top             =   480
      Width           =   9495
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "RC_Code"
      Height          =   285
      Index           =   1
      Left            =   2460
      TabIndex        =   4
      Tag             =   "01-Root Cause Code"
      Top             =   3000
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECRC"
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      DataField       =   "RC_Case"
      Height          =   285
      Left            =   4320
      MaxLength       =   8
      TabIndex        =   3
      Tag             =   "11-Incident Number"
      Top             =   2595
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox comShift 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Tag             =   "01-Incident Number"
      Top             =   2580
      Width           =   1575
   End
   Begin VB.TextBox Updstats 
      DataField       =   "RC_LDate"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   11550
      MaxLength       =   25
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   10440
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      DataField       =   "RC_LTime"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   13290
      MaxLength       =   25
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   10440
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      DataField       =   "RC_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   15030
      MaxLength       =   25
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   10440
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   13755
      _Version        =   65536
      _ExtentX        =   24262
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BevelInner      =   2
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin VB.Label lblEEProdLine 
         AutoSize        =   -1  'True
         Caption         =   "Product Line"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   6720
         TabIndex        =   48
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblEENumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   160
         Width           =   1005
      End
      Begin VB.Label lblEENum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1200
         TabIndex        =   20
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2880
         TabIndex        =   19
         Top             =   135
         Width           =   720
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8760
      Top             =   9960
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   10920
      Top             =   9960
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   10560
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowWidth     =   480
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   2
      ReportSource    =   1
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label lblSecondary 
      Caption         =   "Secondary :"
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label lblPrimary 
      Caption         =   "Primary :"
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
      TabIndex        =   37
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Incident Number"
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
      Index           =   2
      Left            =   240
      TabIndex        =   25
      Top             =   2640
      Width           =   1545
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Root Cause Code"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   24
      Top             =   3045
      Width           =   2355
   End
   Begin VB.Label lblEEID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "RC_Empnbr"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   9990
      TabIndex        =   22
      Top             =   10440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "RC_CompNo"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8850
      TabIndex        =   23
      Top             =   10440
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmEHSCause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNewCode
Dim fglbNew
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim rsDATA3 As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control


Function chkHSRootCause()

Dim SQLQ As String, Msg As String, dd#

chkHSRootCause = False

On Error GoTo chkHSRootCause_Err

Dim tTime As Variant
Dim Part1$, Part2$

'~~

If Len(txtShift) < 1 Then
    MsgBox "Incident Number is a required field"
    comShift.SetFocus
    Exit Function
End If
If Not IfIncidentNo(Val(txtShift)) Then
    MsgBox "Incident Number Not Valid"
    clpCode(1).SetFocus
    Exit Function
End If

If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Root Causes code must be valid"
    clpCode(1).SetFocus
    Exit Function
End If

'Ticket #12867 - Begin
'Root Causes Code is not a required field
'If Len(clpCode(1).Text) < 1 Then
'    MsgBox "Root Causes Code is a required field"
'    clpCode(1).SetFocus
'    Exit Function
'End If
If glbWFC Then
    If Len(clpCode(2).Text) < 1 Then
        MsgBox "Type of Event is a required field"
        clpCode(2).SetFocus
        Exit Function
    End If
    If Len(clpCode(3).Text) < 1 Then
        MsgBox "Immediate / Direct Causes is a required field"
        clpCode(3).SetFocus
        Exit Function
    End If
    If Len(clpCode(4).Text) < 1 Then
        MsgBox "Basic Categories is a required field"
        clpCode(4).SetFocus
        Exit Function
    End If
    If Len(clpCode(5).Text) < 1 Then
        MsgBox "Basic / Underlying Causes is a required field"
        clpCode(5).SetFocus
        Exit Function
    End If
End If
'Ticket #12867 - End

If glbLinamar Then 'Ticket #15172
    If Len(txtProDesc.Text) = 0 Then
        MsgBox "Problem Description is a required field"
        txtProDesc.SetFocus
        Exit Function
    End If
    If Len(clpCode(2).Text) < 1 Then
        MsgBox "Substandard Act/Condition is a required field"
        clpCode(2).SetFocus
        Exit Function
    End If
    If Len(clpCode(4).Text) < 1 Then
        MsgBox "Personal/Job Factor is a required field"
        clpCode(4).SetFocus
        Exit Function
    End If
    If Len(txtRCDesc.Text) = 0 Then
        MsgBox "Root Cause Description is a required field"
        txtRCDesc.SetFocus
        Exit Function
    End If
End If

chkHSRootCause = True

Exit Function

chkHSRootCause_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HR_OHS_ROOT_CAUSE", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

'Private Sub cmdCAction_Click()
'frmEHSCorrective.Show
'Unload Me
'End Sub

Private Sub cmdCAction_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Sub cmdCancel_Click()
Dim X
On Error GoTo Can_Err
fglbNew = False
Call Display_Value
'Call SET_UP_MODE
'Call ST_UPD_MODE(False)  ' reset screen's attributes

Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_OHS_ROOT_CAUSE", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If



End Sub

'Sub cmdCancel_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEHSCAUSE" Then glbOnTop = ""

End Sub


'Sub cmdClose_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdContact_Click()
'frmEHSContact.Show
'Unload Me
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, INo&, X

If Not gSec_Upd_HSRootCause Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If


On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001X.CommitTrans
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
End If
If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If
fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_OHS_ROOT_CAUSE", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub


Sub cmdNew_Click()
Dim SQLQ As String

If Not gSec_Upd_HSRootCause Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If
fglbNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
On Error GoTo AddN_Err

'Data1.Recordset.AddNew
'Sam add
Call Set_Control("B", Me)
'Call Set_Control2("B", rsDATA3)



If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"


comShift.SetFocus


Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_OHS_ROOT_CAUSE", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Sub CmdNew_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim X
On Error GoTo Add_Err

If Not chkHSRootCause() Then Exit Sub
rsDATA.Requery
If fglbNew Then
    rsDATA.AddNew
    rsDATA("RC_CODE_TABL") = "ECRC"
End If

If IsNumeric(txtInvMem(0).Text) And elpInvMem(0).Caption <> "Unassigned" Then
    txtInvMemFName(0).Text = GetEmpData(txtInvMem(0).Text, "ED_FNAME")
    txtInvMemSName(0).Text = GetEmpData(txtInvMem(0).Text, "ED_SURNAME")
End If
If IsNumeric(txtInvMem(1).Text) And elpInvMem(1).Caption <> "Unassigned" Then
    txtInvMemFName(1).Text = GetEmpData(txtInvMem(1).Text, "ED_FNAME")
    txtInvMemSName(1).Text = GetEmpData(txtInvMem(1).Text, "ED_SURNAME")
End If
If IsNumeric(txtInvMem(2).Text) And elpInvMem(2).Caption <> "Unassigned" Then
    txtInvMemFName(2).Text = GetEmpData(txtInvMem(2).Text, "ED_FNAME")
    txtInvMemSName(2).Text = GetEmpData(txtInvMem(2).Text, "ED_SURNAME")
End If
If IsNumeric(txtInvMem(3).Text) And elpInvMem(3).Caption <> "Unassigned" Then
    txtInvMemFName(3).Text = GetEmpData(txtInvMem(3).Text, "ED_FNAME")
    txtInvMemSName(3).Text = GetEmpData(txtInvMem(3).Text, "ED_SURNAME")
End If
If IsNumeric(txtInvMem(4).Text) And elpInvMem(4).Caption <> "Unassigned" Then
    txtInvMemFName(4).Text = GetEmpData(txtInvMem(4).Text, "ED_FNAME")
    txtInvMemSName(4).Text = GetEmpData(txtInvMem(4).Text, "ED_SURNAME")
End If

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    Call UpdUStats(Me) ' update user's stats (who did it and when)
    Call Set_Control("U", Me, rsDATA)
    
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
Else
    Call UpdUStats(Me) ' update user's stats (who did it and when)
    Call Set_Control("U", Me, rsDATA)
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
End If
Data1.Refresh
fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
Me.vbxTrueGrid.SetFocus
If NextFormIF("Root Cause") Then
    Call cmdNew_Click
End If
Exit Sub

Add_Err:
If Err = 3022 Then
    Data1.Recordset.CancelUpdate
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OHS_ROOT_CAUSE", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

'Sub cmdOK_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Root Causes"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lblEEName & "'s Root Causes"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub
'Sub cmdPrint_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub


Private Sub clpCode_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
If Not glbLinamar Then  'Ticket #14666
    If glbWFC Or glbLinamar Then 'Ticket #12506
        If Index = 5 Or (glbLinamar And Index = 8) Then
            clpCode(Index).TransDiv = GetTransDiv(Index) '"'1','3'"
        End If
    Else
        If Not (glbCompSerial = "S/N - 2387W") Then
            'Bird Packaging Limited Ticket #13597
            If Index = 3 Or Index = 4 Or Index = 5 Then
                clpCode(Index).TransDiv = GetTransDiv(Index)
            End If
        End If
    End If
End If
End Sub

Private Function GetTransDiv(xNo As Integer)
Dim rsTran As New ADODB.Recordset
Dim SQLQ As String
Dim xFirstVal As String
Dim xFinal As String

    xFirstVal = clpCode(xNo - 1).Text
    xFinal = "'*'"
    If Not Len(xFirstVal) = 0 Then
        SQLQ = "SELECT * FROM "
        If xNo = 3 Then
            SQLQ = SQLQ & "HR_OHS_RLINK_EVENT "
        ElseIf xNo = 4 Then
            SQLQ = SQLQ & "HR_OHS_RLINK_IMMEDIATE "
        Else
            SQLQ = SQLQ & "HR_OHS_RLINK_BASIC "
        End If
        SQLQ = SQLQ & "WHERE RL_FIRSTCODE ='" & xFirstVal & "' "
        rsTran.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsTran.EOF
            xFinal = xFinal & ",'" & rsTran("RL_SECONDCODE") & "'"
            rsTran.MoveNext
        Loop
    End If
    GetTransDiv = xFinal
End Function

Private Sub cmdCopy_Click() 'Ticket #24803 Franks 12/17/2013
Dim a As Integer, Msg As String, INo&, X
If Not Data1.Recordset.EOF Then
    Msg = "Are you sure you want to copy this record? "
    
    a% = MsgBox(Msg, 36, "Confirm Copy")
    If a% <> 6 Then Exit Sub
    Call CopyRootCauses
End If
End Sub

'Sub cmdWCBMed_Click()
'frmEHSWCB.Show
'Unload Me
'End Sub


'Private Sub cmdWCBMed_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdWSIB_Click()
'frmEHSWCBC.Show
'Unload Me
'End Sub

Sub comShift_Click()
'txtShift = comShift  'JDY
End Sub

Sub comShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comShift_LostFocus()
txtShift = comShift  'JDY
End Sub

Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_OHS_ROOT_CAUSE", "SELECT")


End Sub

Function EERetrieve()

Dim SQLQ As String

EERetrieve = False

Screen.MousePointer = HOURGLASS
On Error GoTo EERError


If glbtermopen Then         'Lucy July 5, 2000
    SQLQ = "Select * from Term_OHS_ROOT_CAUSES"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY RC_Case DESC"
Else
    SQLQ = "Select * from HR_OHS_ROOT_CAUSES "
    SQLQ = SQLQ & " where RC_Empnbr = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY RC_Case DESC"
End If

Data1.RecordSource = SQLQ
Data1.Refresh

If glbtermopen Then     'Lucy July 5, 2000
    SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from Term_HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY EC_CASE DESC"

Else
    SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
End If

Data3.RecordSource = SQLQ
Data3.Refresh


EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "OCH Retrieve", "HR_OHS_ROOT_CAUSES", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


Exit Function

End Function

Private Sub elpInvMem_Change(Index As Integer)
    txtInvMem(Index).Text = getEmpnbr(elpInvMem(Index).Text)
    If elpInvMem(Index).Caption = "Unassigned" Then
        If Not Data1.Recordset.EOF Then
            Select Case Index
                Case 0: txtInvMemName(Index).Text = Data1.Recordset!RC_INVEST1_SURNAME & ", " & Data1.Recordset!RC_INVEST1_FNAME
                Case 1: txtInvMemName(Index).Text = Data1.Recordset!RC_INVEST2_SURNAME & ", " & Data1.Recordset!RC_INVEST2_FNAME
                Case 2: txtInvMemName(Index).Text = Data1.Recordset!RC_INVEST3_SURNAME & ", " & Data1.Recordset!RC_INVEST3_FNAME
                Case 3: txtInvMemName(Index).Text = Data1.Recordset!RC_INVEST4_SURNAME & ", " & Data1.Recordset!RC_INVEST4_FNAME
                Case 4: txtInvMemName(Index).Text = Data1.Recordset!RC_INVEST5_SURNAME & ", " & Data1.Recordset!RC_INVEST5_FNAME
                
            End Select
            If Trim(txtInvMemName(Index)) = "," Then
                txtInvMemName(Index).Visible = False
            Else
                txtInvMemName(Index).Visible = True
            End If
        Else
            txtInvMemName(Index).Visible = False
        End If
    Else
        txtInvMemName(Index).Visible = False
    End If
End Sub

Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMEHSCAUSE"

End Sub

Sub Form_GotFocus()
glbOnTop = "FRMEHSCAUSE"
End Sub

Sub Form_Load()
Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim X%
Dim SQLQ1
Dim xLine1 As Integer
Dim xLine2 As Integer
Dim xLine3 As Integer

glbOnTop = "FRMEHSCAUSE"

'Ticket #12867
If glbWFC Then
    Call WFCScreenSetup 'Ticket #24803 Franks 12/17/2013
End If

If glbLinamar Then 'Ticket #12758
    lblTitle(4).FontBold = False
    'Move Root Cause Code underneath Immediate / Direct Causes
    'Hemu - Ticket #14573 - the following line of code is causing layout issue
    'As it is the labels have changed in this ticket so it does not matter where the
    'code placement is
    'xLine1 = clpCode(1).Top
    'xLine2 = clpCode(2).Top
    'xLine3 = clpCode(3).Top
    'lblTitle(1).Top = xLine1
    'clpCode(2).Top = xLine1
    'lblTitle(3).Top = xLine2
    'clpCode(3).Top = xLine2
    'lblTitle(4).Top = xLine3
    'clpCode(1).Top = xLine3
    'clpCode(1).TabIndex = 99
    'clpCode(2).TabIndex = 2
    'clpCode(3).TabIndex = 3
    'clpCode(1).TabIndex = 4
    'Hemu - End - Ticket #14573
    
    'Rename
    'Root Cause Code -> Basic/Root Causes
    'Basic Categories ->Sub Basic/Root Causes 'Sub Category for Basic/Root Causes
    lblTitle(4).Caption = "Basic/Root Causes"
    lblTitle(6).Caption = "Sub Basic/Root Causes"
    vbxTrueGrid.Columns(1).Caption = "Basic/Root Causes"
    
    'Ticket #14573
    lblTitle(1).Caption = "Substandard Act/Condition"
    lblTitle(3).Caption = "Substandard Condition"
    lblTitle(6).Caption = "Personal/Job Factor"
    lblTitle(7).Caption = "Job/System Factor"
    
    lblTitle(0).Caption = "Substandard Act/Condition"
    lblTitle(8).Caption = "Substandard Condition"
    lblTitle(9).Caption = "Personal/Job Factor"
    lblTitle(10).Caption = "Job/System Factor"
    
    'Ticket #14666
    lblTitle(13).Visible = True
    lblTitle(14).Visible = True
    lblTitle(1).Top = 285
    lblTitle(3).Top = 645
    lblTitle(6).Top = 1225
    lblTitle(7).Top = 1585
    clpCode(2).Top = 240
    clpCode(3).Top = 600
    'clpCode(4).Top = 1180
    clpCode(5).Top = 1540
    lblTitle(1).Left = 360
    lblTitle(3).Left = 360
    lblTitle(6).Left = 360
    lblTitle(7).Left = 360
    clpCode(2).Left = 2580
    clpCode(3).Left = 2580
    clpCode(4).Left = 2580
    clpCode(5).Left = 2580
    
    lblPrimary.Visible = True
    frPrimary.Top = 3600    '3720
    frPrimary.Left = 720    '360
    frPrimary.Height = 1300 '1935
    lblSecondary.Visible = True
    lblSecondary.Top = 4950 '5550
    frSecondary.Visible = True
    frSecondary.Left = 720
    frSecondary.Top = 5200 '5800
    frSecondary.Height = 1300
    
    frComments.Top = 7550
    
    'Ticket #15172 - Begin
    lblTitle(4).Visible = False
    clpCode(1).Visible = False
    lblTitle(3).Visible = False
    clpCode(3).Visible = False
    lblTitle(7).Visible = False
    clpCode(5).Visible = False
    lblTitle(8).Visible = False
    clpCode(6).Visible = False
    lblTitle(10).Visible = False
    clpCode(8).Visible = False
    
    lblTitle(9).Top = 960 'lblTitle(12).Top
    clpCode(7).Top = 960 - 45 'lblTitle(12).Top - 45
    lblTitle(12).Top = lblTitle(8).Top
    lblTitle(6).Top = 980 'lblTitle(14).Top
    clpCode(4).Top = 980 - 45 'lblTitle(14).Top - 45
    lblTitle(14).Top = lblTitle(3).Top
    
    frProblemDesc.Left = 0
    frProblemDesc.Top = lblTitle(4).Top - 50
    frProblemDesc.Width = 9135
    frProblemDesc.BorderStyle = 0
    frProblemDesc.Visible = True
    txtProDesc.DataField = "RC_PROBLEM_DESC"
    
    frRootCauseDesc.Left = 0
    frRootCauseDesc.Top = 6600
    frRootCauseDesc.Width = 9135
    frRootCauseDesc.BorderStyle = 0
    frRootCauseDesc.Visible = True
    txtRCDesc.DataField = "RC_ROOTCAUSE_DESC"
    
    frInvestigation.Left = 6720
    frInvestigation.Top = 4300
    frInvestigation.Width = 4620
    'frRootCauseDesc.BorderStyle = 0
    frInvestigation.Visible = True
    txtInvMem(0).DataField = "RC_INVESTIGATION1"
    txtInvMem(1).DataField = "RC_INVESTIGATION2"
    txtInvMem(2).DataField = "RC_INVESTIGATION3"
    txtInvMem(3).DataField = "RC_INVESTIGATION4"
    txtInvMem(4).DataField = "RC_INVESTIGATION5"
    
    vbxTrueGrid.Columns(1).Caption = "Problem Description"
    vbxTrueGrid.Columns(1).DataField = "RC_PROBLEM_DESC"
    'Ticket #15172 - End
Else
    frPrimary.Top = 3360
    frPrimary.Height = 1455
    frComments.Top = 4800
End If

If glbCompSerial = "S/N - 2387W" Then  'Bird Packaging Limited 'Ticket #13636
    lblTitle(3).Caption = "Immediate Causes: Sub Act"
    lblTitle(6).Caption = "Immediate Causes: Sub Cond"
End If

If glbtermopen Then         'Lucy July 5, 2000
    Data1.ConnectionString = glbAdoIHRAUDIT
    Data3.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
    Data3.ConnectionString = glbAdoIHRDB
End If

If glbLinHS Then 'Ticket #12401
    glbLinEmpNo = glbLEE_ID
    If Not glbtermopen Then
        If Len(glbDiv) = 0 Then Call Get_Div(False) 'frmDIVISIONS.Show 1
        If Len(glbDiv) = 0 Then Unload Me: Exit Sub
    Else
        If Len(glbDiv) = 0 Then Call Get_Div(False) 'frmDIVISIONS.Show 1
        If Len(glbDiv) = 0 Then Unload Me: Exit Sub
    End If
    glbLinHSDivNo = Val("999999" & glbDiv)
    glbLEE_ID = glbLinHSDivNo
    glbLEE_SName = glbDivDesc
Else
    If glbLinamar Then
        If glbLEE_ID <> 0 Then
            If Left(Trim(Str(glbLEE_ID)), 6) = "999999" Then
                glbLEE_ID = 0
            End If
        End If
    End If
    If Not glbtermopen Then
        If glbLEE_ID = 0 Then frmEEFIND.Show 1
        If glbLEE_ID = 0 Then Unload Me: Exit Sub
    Else
        If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
        If glbTERM_ID = 0 Then Unload Me: Exit Sub
    End If
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If


If Len(glbLEE_SName) < 1 Then Exit Sub
Screen.MousePointer = HOURGLASS

comShift.Clear
Do Until Data3.Recordset.EOF
  comShift.AddItem Data3.Recordset("EC_CASE")
  Data3.Recordset.MoveNext
Loop

Me.vbxTrueGrid.SetFocus
If glbLinHS Then
    If Len(glbDivDesc) > 0 Then   ' dont do on add new until in
        Me.Caption = "Root Causes Data - " & glbDivDesc
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    lblEENumber.Caption = lStr("Division")
Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
        Me.Caption = "Root Causes Data - " & Left$(glbLEE_SName, 8)
        Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    lblEENum.Caption = ShowEmpnbr(lblEEID)
End If
ST_UPD_MODE (False)

Call Display_Value

If Not gSec_Upd_HSRootCause Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
End If

'If Not glbOracle Then
    clpCode(2).DataField = "RC_EVENT"
    clpCode(3).DataField = "RC_IMMEDIA"
    clpCode(4).DataField = "RC_BASICCATA"
    clpCode(5).DataField = "RC_BASIC"
If glbLinamar Then
    clpCode(0).DataField = "RC_EVENT1"
    clpCode(6).DataField = "RC_IMMEDIA1"
    clpCode(7).DataField = "RC_BASICCATA1"
    clpCode(8).DataField = "RC_BASIC1"
    
    'OH&S Root Causes Codes Type Description to change
    Call Change_Code_Table_Description
End If

'End If
Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
End Sub

Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Sub Form_Unload(Cancel As Integer)

MDIMain.panHelp(0).Caption = "Select function from the menu."
Set frmEHSCause = Nothing 'carmen may 00
Call NextForm
End Sub

Function IfIncidentNo(InciNo As Double)
  IfIncidentNo = False
  If Data3.Recordset.BOF And Data3.Recordset.EOF Then
     Exit Function
  End If
  Data3.Recordset.MoveFirst
  Data3.Recordset.Find "EC_Case=" & InciNo
  If Data3.Recordset.EOF Then Exit Function
  IfIncidentNo = True
End Function


Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

glbOHSEdit% = TF


'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF
txtComments.Enabled = TF
comShift.Enabled = TF
If glbWFC Then 'Ticket #12867
    clpCode(1).Enabled = False
Else
     clpCode(1).Enabled = TF
End If
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF
clpCode(5).Enabled = TF

clpCode(0).Enabled = TF
clpCode(6).Enabled = TF
clpCode(7).Enabled = TF
clpCode(8).Enabled = TF
elpInvMem(0).Enabled = TF
elpInvMem(1).Enabled = TF
elpInvMem(2).Enabled = TF
elpInvMem(3).Enabled = TF
elpInvMem(4).Enabled = TF
txtRCDesc.Enabled = TF
txtProDesc.Enabled = TF

'vbxTrueGrid.Enabled = FT
'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
'cmdWCBMed.Enabled = FT
'cmdIncident.Enabled = FT
'cmdCAction.Enabled = FT
'cmdContact.Enabled = FT
'cmdInjLoc.Enabled = FT
'cmdWSIB.Enabled = FT
'vbxTrueGrid.Enabled = FT
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If
End Sub

Private Sub txtComments_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtInvMem_Change(Index As Integer)
elpInvMem(Index) = ShowEmpnbr(txtInvMem(Index).Text)
If IsNumeric(txtInvMem(Index).Text) And elpInvMem(Index).Caption <> "Unassigned" Then
    txtInvMemFName(Index).Text = GetEmpData(txtInvMem(Index).Text, "ED_FNAME")
    txtInvMemSName(Index).Text = GetEmpData(txtInvMem(Index).Text, "ED_SURNAME")
End If
End Sub

Private Sub txtProDesc_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtRCDesc_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Sub txtShift_Change()
  If Not (Val(txtShift) = 0) Then
    comShift = txtShift
  Else
    comShift = ""
  End If
End Sub

Sub txtShift_DblClick()
Dim oCode As String, OCodeD As String

oCode = txtShift

'frmEHSDisplay.Show 1
If Len(glbCode) < 1 Then
    txtShift = oCode
Else
    txtShift = glbCode
End If
End Sub

Sub txtShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Updstats_Change(Index As Integer)
    If Index = 0 Then
        'If IsDate(Updstats(Index).Text) Then
        lblUpdDateDesc.Caption = Updstats(Index).Text
        'End If
    End If
    If Index = 2 Then
        lblUserDesc.Caption = GetUserDesc(Updstats(Index))
    End If
End Sub

Private Sub Updstats_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Sub vbxTrueGrid_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
 Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        If glbtermopen Then         'Lucy July 5, 2000
            SQLQ = "Select * from Term_OHS_ROOT_CAUSES"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select * from HR_OHS_ROOT_CAUSES "
            SQLQ = SQLQ & " where RC_Empnbr = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
'    End If
End If

End Sub


''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        Call SET_UP_MODE
        Exit Sub
    End If
    
    
If glbtermopen Then
    SQLQ = "Select * from Term_OHS_ROOT_CAUSES"
    SQLQ = SQLQ & " WHERE RC_ID = " & Data1.Recordset!RC_ID
    SQLQ = SQLQ & " ORDER BY RC_Case DESC"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "Select * from HR_OHS_ROOT_CAUSES "
    SQLQ = SQLQ & " where RC_ID = " & Data1.Recordset!RC_ID
    SQLQ = SQLQ & " ORDER BY RC_Case DESC"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If
   
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)



Call SET_UP_MODE
End Sub



Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
End Sub


Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fglbNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_HSRootCause
End Property

Public Property Get Addable() As Boolean
Addable = True
End Property

Public Property Get Updateble() As Boolean
Updateble = True
End Property

Public Property Get Deleteble() As Boolean
Deleteble = True
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fglbNew Then
    UpdateState = NewRecord
    TF = True
ElseIf rsDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call ST_UPD_MODE(TF)
End Sub

Private Sub lblEEID_Change()
If glbLinHS Then
    If Len(glbDivDesc) > 0 Then   ' dont do on add new until in
        Me.Caption = "Root Causes Data - " & glbDivDesc
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = ""
    End If
Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        frmEHSCause.Caption = "Root Causes Data - " & Left$(glbLEE_SName, 5)
        frmEHSCause.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    'lblEEID = glbLEE_ID
    lblEENum = ShowEmpnbr(lblEEID)
    
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = glbLEE_ProdLine
    Else
        lblEEProdLine = ""
    End If
End If
End Sub

Private Sub Change_Code_Table_Description()
Dim rsHRTabDesc As New ADODB.Recordset
Dim SQLQ As String

    SQLQ = "SELECT * FROM HRTABDES WHERE (TD_NAME IN ('ECRE','ECRI','ECRA','ECRB'))"
    rsHRTabDesc.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    If Not rsHRTabDesc.EOF Then
        Do While Not rsHRTabDesc.EOF
            If rsHRTabDesc("TD_NAME") = "ECRE" And rsHRTabDesc("TD_DESC") = "OH&S ROOT CAUSE - TYPE OF EVEN" Then
                rsHRTabDesc("TD_DESC") = "OH&S ROOT CAUSE - SUBSTANDARD ACT"
                rsHRTabDesc.Update
            End If
            If rsHRTabDesc("TD_NAME") = "ECRI" And rsHRTabDesc("TD_DESC") = "OH&S ROOT CAUSE - IMMEDIATE/DIRECT" Then
                rsHRTabDesc("TD_DESC") = "OH&S ROOT CAUSE - SUBSTANDARD COND."
                rsHRTabDesc.Update
            End If
            If rsHRTabDesc("TD_NAME") = "ECRA" And rsHRTabDesc("TD_DESC") = "OH&S ROOT CAUSE - BASIC CATEGORIES" Then
                rsHRTabDesc("TD_DESC") = "OH&S ROOT CAUSE - PERSONAL FACTOR"
                rsHRTabDesc.Update
            End If
            If rsHRTabDesc("TD_NAME") = "ECRB" And rsHRTabDesc("TD_DESC") = "OH&S ROOT CAUSE - BASIC/UNDERLYING" Then
                rsHRTabDesc("TD_DESC") = "OH&S ROOT CAUSE - JOB/SYSTEM FACTOR"
                rsHRTabDesc.Update
            End If
            rsHRTabDesc.MoveNext
        Loop
    End If
    rsHRTabDesc.Close
End Sub

Private Sub CopyRootCauses() 'Ticket #24803 Franks 12/17/2013
Dim rsRoot As New ADODB.Recordset
Dim SQLQ As String

If Not Data1.Recordset.EOF Then
    If glbtermopen Then         'Lucy July 5, 2000
        SQLQ = "Select * from Term_OHS_ROOT_CAUSES"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " ORDER BY RC_Case DESC"
    Else
        SQLQ = "Select * from HR_OHS_ROOT_CAUSES "
        SQLQ = SQLQ & " where RC_Empnbr = " & glbLEE_ID
        SQLQ = SQLQ & " ORDER BY RC_Case DESC"
    End If
    If rsRoot.State <> 0 Then rsRoot.Close
    rsRoot.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsRoot.AddNew
    If glbtermopen Then rsRoot("TERM_SEQ") = glbTERM_Seq
    rsRoot("RC_Empnbr") = glbLEE_ID
    rsRoot("RC_CompNo") = "001"
    rsRoot("RC_Case") = Data1.Recordset("RC_Case")
    rsRoot("RC_Code") = Data1.Recordset("RC_Code")
    rsRoot("RC_Comments") = Data1.Recordset("RC_Comments")
    rsRoot("RC_EVENT") = Data1.Recordset("RC_EVENT")
    rsRoot("RC_IMMEDIA") = Data1.Recordset("RC_IMMEDIA")
    rsRoot("RC_BASIC") = Data1.Recordset("RC_BASIC")
    rsRoot("RC_BASICCATA") = Data1.Recordset("RC_BASICCATA")
    rsRoot("RC_EVENT1") = Data1.Recordset("RC_EVENT1")
    rsRoot("RC_IMMEDIA1") = Data1.Recordset("RC_IMMEDIA1")
    rsRoot("RC_BASIC1") = Data1.Recordset("RC_BASIC1")
    rsRoot("RC_BASICCATA1") = Data1.Recordset("RC_BASICCATA1")
    rsRoot("RC_PROBLEM_DESC") = Data1.Recordset("RC_PROBLEM_DESC")
    rsRoot("RC_ROOTCAUSE_DESC") = Data1.Recordset("RC_ROOTCAUSE_DESC")
    rsRoot("RC_INVESTIGATION1") = Data1.Recordset("RC_INVESTIGATION1")
    rsRoot("RC_INVESTIGATION2") = Data1.Recordset("RC_INVESTIGATION2")
    rsRoot("RC_INVESTIGATION3") = Data1.Recordset("RC_INVESTIGATION3")
    rsRoot("RC_INVESTIGATION4") = Data1.Recordset("RC_INVESTIGATION4")
    rsRoot("RC_INVESTIGATION5") = Data1.Recordset("RC_INVESTIGATION5")
    rsRoot("RC_INVEST1_FNAME") = Data1.Recordset("RC_INVEST1_FNAME")
    rsRoot("RC_INVEST2_FNAME") = Data1.Recordset("RC_INVEST2_FNAME")
    rsRoot("RC_INVEST3_FNAME") = Data1.Recordset("RC_INVEST3_FNAME")
    rsRoot("RC_INVEST4_FNAME") = Data1.Recordset("RC_INVEST4_FNAME")
    rsRoot("RC_INVEST5_FNAME") = Data1.Recordset("RC_INVEST5_FNAME")
    rsRoot("RC_INVEST1_SURNAME") = Data1.Recordset("RC_INVEST1_SURNAME")
    rsRoot("RC_INVEST2_SURNAME") = Data1.Recordset("RC_INVEST2_SURNAME")
    rsRoot("RC_INVEST3_SURNAME") = Data1.Recordset("RC_INVEST3_SURNAME")
    rsRoot("RC_INVEST4_SURNAME") = Data1.Recordset("RC_INVEST4_SURNAME")
    rsRoot("RC_INVEST5_SURNAME") = Data1.Recordset("RC_INVEST5_SURNAME")
    rsRoot("RC_LDate") = Date
    rsRoot("RC_LTime") = Time$
    rsRoot("RC_LUSER") = glbUserID
    rsRoot.Update
    rsRoot.Close
    Data1.Refresh
End If
End Sub

Private Sub WFCScreenSetup()
    lblTitle(4).Enabled = False
    clpCode(1).Enabled = False
    lblTitle(1).FontBold = True
    lblTitle(3).FontBold = True
    lblTitle(6).FontBold = True
    lblTitle(7).FontBold = True
    'Ticket #24803 Franks 12/17/2013 - begin
    cmdCopy.Visible = True
    vbxTrueGrid.Columns(1).Visible = False
    vbxTrueGrid.Columns(2).Visible = False
    'Ticket #24803 Franks 12/17/2013 - end
End Sub
