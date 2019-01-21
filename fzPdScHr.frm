VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRPaidSicHr 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Paid Sick Hours Per Eligible Employee Report"
   ClientHeight    =   8610
   ClientLeft      =   570
   ClientTop       =   1095
   ClientWidth     =   10110
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
   ScaleHeight     =   8610
   ScaleWidth      =   10110
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   7875
      LargeChange     =   315
      Left            =   9600
      Max             =   100
      SmallChange     =   315
      TabIndex        =   48
      Top             =   240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   50
      SmallChange     =   4
      TabIndex        =   47
      Top             =   8295
      Width           =   9855
   End
   Begin VB.Frame scrFrame 
      BorderStyle     =   0  'None
      Height          =   8055
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   9375
      Begin VB.OptionButton optEligible 
         Caption         =   "Benefit Group"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   20
         Tag             =   "Benefit Group"
         Top             =   6480
         Value           =   -1  'True
         Width           =   1515
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Tag             =   "Final Sort of Records"
         Top             =   7590
         Visible         =   0   'False
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Tag             =   "First Level of grouping records"
         Top             =   7260
         Width           =   2325
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1995
         MaxLength       =   4
         TabIndex        =   11
         Tag             =   "00-Employee Position Shift"
         Top             =   3870
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.CheckBox chkTerm 
         Caption         =   "Include Terminated Employee"
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Tag             =   "Check to included Terminated Employees"
         Top             =   5520
         Width           =   3135
      End
      Begin VB.Frame frmTerm 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   3420
         TabIndex        =   25
         Top             =   5490
         Visible         =   0   'False
         Width           =   4695
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   1
            Left            =   3060
            TabIndex        =   18
            Tag             =   "40-Date upto and including this date"
            Top             =   60
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   0
            Left            =   1380
            TabIndex        =   17
            Tag             =   "40-Date from and including this date forward"
            Top             =   60
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
            Top             =   90
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkInclAtt 
         Caption         =   "Include Attendance History"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Tag             =   "Check to include Attendance History"
         Top             =   5880
         Width           =   3075
      End
      Begin VB.ComboBox cmbDateBased 
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
         ItemData        =   "fzPdScHr.frx":0000
         Left            =   1980
         List            =   "fzPdScHr.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Tag             =   "Select Date Range Based On"
         Top             =   4590
         Width           =   2325
      End
      Begin VB.OptionButton optEligible 
         Caption         =   "Current Sick Entitlement"
         Height          =   255
         Index           =   1
         Left            =   4740
         TabIndex        =   21
         Tag             =   "Current Sick Entitlement"
         Top             =   6480
         Width           =   2475
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   1680
         TabIndex        =   7
         Tag             =   "00-Enter Position Group Code"
         Top             =   2550
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
         Index           =   2
         Left            =   1680
         TabIndex        =   4
         Tag             =   "00-Enter Status Code"
         Top             =   1560
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
         TabIndex        =   5
         Tag             =   "EDPT-Category"
         Top             =   1890
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
         Left            =   1680
         TabIndex        =   3
         Tag             =   "00-Enter Union Code"
         Top             =   1230
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
         Left            =   1680
         TabIndex        =   2
         Tag             =   "00-Enter Location Code"
         Top             =   900
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDLC"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Tag             =   "00-Specific Department Desired"
         Top             =   570
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
         Top             =   240
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
         Index           =   5
         Left            =   1680
         TabIndex        =   9
         Tag             =   "00-Enter Administered By Code"
         Top             =   3210
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDAB"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   1680
         TabIndex        =   10
         Tag             =   "00-Enter Section Code"
         Top             =   3540
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDSE"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   1680
         TabIndex        =   8
         Tag             =   "00-Enter Region Code"
         Top             =   2880
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDRG"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Tag             =   "10-Enter Employee Number"
         Top             =   2220
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         TextBoxWidth    =   7195
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   12
         Tag             =   "40-Date from and including this date forward"
         Top             =   4200
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   3
         Left            =   3360
         TabIndex        =   13
         Tag             =   "40-Date upto and including this date / As of Date"
         Top             =   4200
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   7
         Left            =   1680
         TabIndex        =   15
         Tag             =   "ADRE-Attendance Reason"
         Top             =   4980
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "ADRE"
         MaxLength       =   0
         MultiSelect     =   -1  'True
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
         TabIndex        =   46
         Top             =   240
         Width           =   555
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
         TabIndex        =   45
         Top             =   570
         Width           =   825
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
         TabIndex        =   44
         Top             =   1230
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
         Left            =   120
         TabIndex        =   43
         Top             =   1560
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
         TabIndex        =   42
         Top             =   2220
         Width           =   1290
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
         TabIndex        =   41
         Top             =   0
         Width           =   1575
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
         TabIndex        =   40
         Top             =   7065
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
         TabIndex        =   39
         Top             =   7305
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
         TabIndex        =   38
         Top             =   7620
         Visible         =   0   'False
         Width           =   660
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
         Top             =   900
         Width           =   615
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
         TabIndex        =   36
         Top             =   2880
         Width           =   510
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
         TabIndex        =   35
         Top             =   3210
         Width           =   1125
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Group Code"
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
         Top             =   2550
         Width           =   1455
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
         TabIndex        =   33
         Top             =   3510
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
         TabIndex        =   32
         Top             =   1890
         Width           =   630
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
         TabIndex        =   31
         Top             =   3855
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label1 
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
         Top             =   4230
         Width           =   870
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Based on "
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
         TabIndex        =   29
         Top             =   4620
         Width           =   1110
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sick Attendance Code"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   5025
         Width           =   1560
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Eligible Employees are based on "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   6480
         Width           =   2820
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   9600
      Top             =   8160
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
End
Attribute VB_Name = "frmRPaidSicHr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents cnRun  As ADODB.Connection
Attribute cnRun.VB_VarHelpID = -1

Private Sub cmdClose_Click()
Unload Me
End Sub



Public Sub cmdPrint_Click()
Dim x%

'''On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Emergency Contact Report Criteria", Me) Then Exit Sub
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Call set_PrintState(False)
    x% = Cri_SetAll()
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
Dim x%, selected&
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
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    'Text1.Text = Me.vbxCrystal.RecordsPrinted
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    'Text1.Text = Me.vbxCrystal.RecordsPrinted
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ATTEND", "SELECT")
Resume Next

End Sub

Private Sub chkInclAtt_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkTerm_Click()
frmTerm.Visible = chkTerm.Value
End Sub

Private Sub chkTerm_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub clpCode_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmbDateBased_Click()
    'chkInclAtt.Visible = cmbDateBased.ListIndex = 4
End Sub

Private Sub cmbDateBased_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
Dim x%
Dim vPosGroup As String
    
    'Hemu 06/02/2004 Begin
    cmbDateBased.AddItem lStr("Original Hire Date")
    cmbDateBased.AddItem lStr("Seniority Date")
    cmbDateBased.AddItem lStr("Last Hire Date")
    cmbDateBased.AddItem lStr("Union Date")
    cmbDateBased.AddItem lStr("Attendance Date")
    cmbDateBased.AddItem lStr("User Defined Date")
    cmbDateBased.ListIndex = 0
    'Hemu 06/02/2004 End
    
    If Not glbSyndesis Then
        vPosGroup = "Position Group"
    Else
        vPosGroup = "Position Grade"
    End If
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Union")
    comGroup(0).AddItem "Employment Status"
    comGroup(0).AddItem lStr("Category")
    comGroup(0).AddItem vPosGroup '"Position Group Code"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem lStr("Administered By")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem "(none)"
    comGroup(2).AddItem "Employee Name"

    comGroup(0).ListIndex = 0
    comGroup(2).ListIndex = 0
    
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
    Case 5: strCd$ = "HRJOB.JB_GRPCD"
    Case 6: strCd$ = "HREMP.ED_SECTION"
    Case 7: strCd$ = "HR_ATTENDANCE.AD_REASON"
    End Select
    
    'Hemu 06/02/2004 Begin
    '    CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    If glbOracle Then
        CodeCri = "({" & strCd$ & "} IN ['" & getCodes(clpCode(intIdx%).Text) & "'])"
    Else
        CodeCri = "({" & strCd$ & "} IN ('" & getCodes(clpCode(intIdx%).Text) & "'))"
    End If
    'Hemu 06/02/2004 End
    
    'Need clarification for below to incorporate multiple codes - Hemu
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
    'Hemu 06/02/2004 Begin
    'DivCri = "({HREMP.ED_DIV} = '" & clpDiv.Text & "')"
    If glbOracle Then
        DivCri = "({HREMP.ED_DIV} IN ['" & getCodes(clpDiv.Text) & "'])"
    Else
        DivCri = "({HREMP.ED_DIV} IN ('" & getCodes(clpDiv.Text) & "'))"
    End If
    'Hemu 06/02/2004 End
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
    If glbOracle Then
        EECri = "{HREMP.ED_EMPNBR} IN [" & getEmpnbr(elpEEID.Text) & "] "
    Else
        EECri = "{HREMP.ED_EMPNBR} IN (" & getEmpnbr(elpEEID.Text) & ") "
    End If
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
Dim EECri As String, OneSet%, x%

    If Len(clpPT.Text) < 1 Then Exit Sub

    'Hemu 06/02/2004 Begin
    'EECri = "{HREMP.ED_PT}= '" & clpPT.Text & "'"
    If glbOracle Then
        EECri = "{HREMP.ED_PT} IN ['" & getCodes(clpPT.Text) & "']"
    Else
        EECri = "{HREMP.ED_PT} IN ('" & getCodes(clpPT.Text) & "')"
    End If
    'Hemu 06/02/2004 End
    
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True

End Sub

Private Function Cri_SetAll()
Dim x%, strRName$, selected&
Dim selectform
Dim CodeCri

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
For x% = 0 To 4
    Call Cri_Code(x%)
Next x%
Call Cri_Code(6)
'Call Cri_Code(7)
Call Cri_PT
Call Cri_Shift
Call Cri_EE

'Hemu 06/03/2004 Begin
'As of Date = Date Range
If Len(dlpDateRange(2).Text) > 0 Or Len(dlpDateRange(3).Text) > 0 Then
    Select Case cmbDateBased
    Case lStr("Original Hire Date")
        If glbiOneWhere Then
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_DOH} >= " & Date_SQL(dlpDateRange(2)) & " AND {HREMP.ED_DOH} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_DOH} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_DOH} >= " & Date_SQL(dlpDateRange(2))
            End If
        Else
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {HREMP.ED_DOH} >= " & Date_SQL(dlpDateRange(2)) & " AND {HREMP.ED_DOH} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {HREMP.ED_DOH} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = " {HREMP.ED_DOH} >= " & Date_SQL(dlpDateRange(2))
            End If
        End If
        glbiOneWhere = True
    Case lStr("Seniority Date")
        If glbiOneWhere Then
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_SENDTE} >= " & Date_SQL(dlpDateRange(2)) & " AND {HREMP.ED_SENDTE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_SENDTE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_SENDTE} >= " & Date_SQL(dlpDateRange(2))
            End If
        Else
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {HREMP.ED_SENDTE} >= " & Date_SQL(dlpDateRange(2)) & " AND {HREMP.ED_SENDTE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {HREMP.ED_SENDTE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = " {HREMP.ED_SENDTE} >= " & Date_SQL(dlpDateRange(2))
            End If
        End If
        glbiOneWhere = True
    Case lStr("Last Hire Date")
        If glbiOneWhere Then
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_LTHIRE} >= " & Date_SQL(dlpDateRange(2)) & " AND {HREMP.ED_LTHIRE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_LTHIRE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_LTHIRE} >= " & Date_SQL(dlpDateRange(2))
            End If
        Else
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {HREMP.ED_LTHIRE} >= " & Date_SQL(dlpDateRange(2)) & " AND {HREMP.ED_LTHIRE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {HREMP.ED_LTHIRE} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = " {HREMP.ED_LTHIRE} >= " & Date_SQL(dlpDateRange(2))
            End If
        End If
        glbiOneWhere = True
    Case lStr("Union Date")
        If glbiOneWhere Then
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_UNION} >= " & Date_SQL(dlpDateRange(2)) & " AND {HREMP.ED_UNION} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_UNION} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_UNION} >= " & Date_SQL(dlpDateRange(2))
            End If
        Else
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {HREMP.ED_UNION} >= " & Date_SQL(dlpDateRange(2)) & " AND {HREMP.ED_UNION} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {HREMP.ED_UNION} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = " {HREMP.ED_UNION} >= " & Date_SQL(dlpDateRange(2))
            End If
        End If
        glbiOneWhere = True
    Case lStr("User Defined Date")
        If glbiOneWhere Then
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_USRDAT1} >= " & Date_SQL(dlpDateRange(2)) & " AND {HREMP.ED_USRDAT1} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_USRDAT1} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = glbstrSelCri & " AND {HREMP.ED_USRDAT1} >= " & Date_SQL(dlpDateRange(2))
            End If
        Else
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {HREMP.ED_USRDAT1} >= " & Date_SQL(dlpDateRange(2)) & " AND {HREMP.ED_USRDAT1} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                glbstrSelCri = " {HREMP.ED_USRDAT1} <= " & Date_SQL(dlpDateRange(3))
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                glbstrSelCri = " {HREMP.ED_USRDAT1} >= " & Date_SQL(dlpDateRange(2))
            End If
        End If
        glbiOneWhere = True
    End Select
End If
'Hemu 06/03/2004 End

Call SETWRK

If optEligible(0) Then
    Call Get_Sick_Hrs
    If chkTerm Then
        Call Get_Sick_Hrs_Terminated
    End If
ElseIf optEligible(1) Then
    Call Get_Sick_Hrs_Entitle
    If chkTerm Then
        Call Get_Sick_Hrs_Entitle_Terminated
    End If
End If

' report name
If comGroup(0) <> "(none)" Then
  strRName$ = glbIHRREPORTS & "rzPdSic1.rpt"
Else
  strRName$ = glbIHRREPORTS & "rzPdSick.rpt"
End If
Me.vbxCrystal.ReportFileName = strRName$

' set to sorting/grouping criteria
x% = Cri_Sorts()   ' returns number of sections formated

If Len(glbstrSelCri) >= 0 Then
    selectform = " {HREMP.ED_WRKEMP}='" & glbUserID & "'"
        
    If glbCompSerial = "S/N - 2347W" Then   'Surrey Place
        Me.vbxCrystal.SelectionFormula = selectform & " AND {HREMP.ED_PT} <> 'TR'"
    Else
        Me.vbxCrystal.SelectionFormula = selectform
    End If
End If

'set location for database tables
Cont_Average:
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDBW
        Me.vbxCrystal.DataFiles(1) = glbIHRDB
        Me.vbxCrystal.DataFiles(2) = glbIHRDB
        Me.vbxCrystal.DataFiles(3) = glbIHRDB
        Me.vbxCrystal.DataFiles(4) = glbIHRDB
        Me.vbxCrystal.DataFiles(5) = glbIHRDB
        Me.vbxCrystal.DataFiles(6) = glbIHRDB
        Me.vbxCrystal.DataFiles(7) = glbIHRDB
        Me.vbxCrystal.DataFiles(8) = glbIHRDB
        Me.vbxCrystal.DataFiles(9) = glbIHRDB
        Me.vbxCrystal.DataFiles(10) = glbIHRDB
        Me.vbxCrystal.DataFiles(11) = glbIHRDBW
    End If
    ' window title if appropriate
    Me.vbxCrystal.WindowTitle = "Paid Sick Hours Per Eligible Employee Report"
    
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

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, x%

If Len(txtShift.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Sub SETWRK()
Dim SQLQ, xNum, xRecNum, SQLQ1
Dim ESQLQ
Dim rsEMP As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim xFieldList
xFieldList = Get_Fields(gdbAdoIhr001W, "HREMP_HS", "KEY_EMPNBR,ED_WRKEMP,JB_GRPCD_TABL,JB_GRPCD,ED_ID,ED_HOMELINE_TABL,JH_JOB,")
xFieldList = Replace(xFieldList, "ED_LANG1_TABL, ED_LANG1, ED_LANG2_TABL, ED_LANG2, ", "")

Set cnRun = New ADODB.Connection
cnRun.CommandTimeout = 600
cnRun.Open glbAdoIHRDBW

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3
gdbAdoIhr001.CommandTimeout = 600
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).FloodPercent = 15

ESQLQ = glbstrSelCri
ESQLQ = Replace(ESQLQ, "{", "")
ESQLQ = Replace(ESQLQ, "}", "")
ESQLQ = Replace(ESQLQ, "HREMP.", "")

If glbSQL Or glbOracle Then
    ESQLQ = Replace(ESQLQ, "[", "(")
    ESQLQ = Replace(ESQLQ, "]", ")")
End If

cnRun.BeginTrans
cnRun.Execute "DELETE FROM HREMP_HS WHERE ED_WRKEMP='" & glbUserID & "'"
cnRun.CommitTrans


MDIMain.panHelp(0).FloodPercent = 30
'for active employees
SQLQ = "INSERT INTO HREMP_HS (" & xFieldList & ",KEY_EMPNBR,ED_WRKEMP)"
SQLQ = SQLQ & " SELECT " & xFieldList
SQLQ = SQLQ & ",'1_'  AS KEY_EMPNBR "
SQLQ = SQLQ & ",'" & glbUserID & "' AS ED_WRKEMP "
SQLQ = SQLQ & " FROM HREMP "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " WHERE " & ESQLQ
If Len(clpCode(3).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " WHERE JH_CURRENT<>0 "
    SQLQ = SQLQ & " AND JH_JOB IN (SELECT JB_CODE FROM HRJOB "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " WHERE JB_GRPCD='" & clpCode(3).Text & "'))"
End If
cnRun.BeginTrans
cnRun.Execute SQLQ
cnRun.CommitTrans
rsEMP.Open "SELECT ED_EMPNBR,JB_GRPCD_TABL,JB_GRPCD,ED_ID FROM HREMP_HS", cnRun, adOpenStatic, adLockPessimistic
Do Until rsEMP.EOF
    SQLQ1 = "SELECT JB_GRPCD_TABL,JB_GRPCD FROM HRJOB WHERE JB_CODE IN (SELECT JH_JOB FROM HR_JOB_HISTORY "
    SQLQ1 = SQLQ1 & " WHERE JH_EMPNBR=" & rsEMP("ED_EMPNBR") & ")"
    rsJOB.Open SQLQ1, gdbAdoIhr001, adOpenForwardOnly
    If Not rsJOB.EOF Then
        rsEMP("JB_GRPCD_TABL") = "JBGC"
        rsEMP("JB_GRPCD") = rsJOB("JB_GRPCD")
        rsEMP.Update
    End If
    rsJOB.Close
    rsEMP.MoveNext
Loop
rsEMP.Close
MDIMain.panHelp(0).FloodPercent = 50

'for terminated employees
If chkTerm Then
SQLQ = "INSERT INTO HREMP_HS "
    
    SQLQ = "INSERT INTO HREMP_HS (" & xFieldList & ",KEY_EMPNBR,ED_WRKEMP)"
    SQLQ = SQLQ & "SELECT " & xFieldList
    SQLQ = SQLQ & ",'0_'  AS KEY_EMPNBR "
    SQLQ = SQLQ & ",'" & glbUserID & "' AS ED_WRKEMP "
    SQLQ = SQLQ & " FROM Term_HREMP "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & " WHERE " & ESQLQ
    If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
        SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
        SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
        SQLQ = SQLQ & " WHERE Term_DOT >=" & Date_SQL(dlpDateRange(0))
        SQLQ = SQLQ & " AND Term_DOT <=" & Date_SQL(dlpDateRange(1)) & ")"
    Else
        If IsDate(dlpDateRange(0)) Then
            SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
            SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
            SQLQ = SQLQ & " WHERE Term_DOT >=" & Date_SQL(dlpDateRange(0)) & ")"
        End If
        If IsDate(dlpDateRange(1)) Then
            SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
            SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
            SQLQ = SQLQ & " WHERE Term_DOT <=" & Date_SQL(dlpDateRange(1)) & ")"
        End If
    End If
    If Len(clpCode(3).Text) > 0 Then
        SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM TERM_JOB_HISTORY "
        SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
        SQLQ = SQLQ & " WHERE JH_CURRENT<>0 "
        SQLQ = SQLQ & " AND JH_JOB IN (SELECT JB_CODE FROM HRJOB "
        SQLQ = SQLQ & in_SQL(glbIHRDB)
        SQLQ = SQLQ & " WHERE JB_GRPCD='" & clpCode(3).Text & "'))"
    End If
    cnRun.BeginTrans
    cnRun.Execute SQLQ
    cnRun.CommitTrans
    If glbOracle Then
        rsEMP.Open "SELECT ED_EMPNBR,JB_GRPCD_TABL,JB_GRPCD,ED_ID FROM HREMP_HS WHERE SUBSTR(KEY_EMPNBR,1,1)='0'", cnRun, adOpenStatic, adLockPessimistic
    Else
        rsEMP.Open "SELECT ED_EMPNBR,JB_GRPCD_TABL,JB_GRPCD,ED_ID FROM HREMP_HS WHERE LEFT(KEY_EMPNBR,1)='0'", cnRun, adOpenStatic, adLockPessimistic
    End If
    Do Until rsEMP.EOF
        SQLQ = "SELECT JB_GRPCD_TABL, JB_GRPCD FROM HRJOB WHERE JB_CODE IN (SELECT JH_JOB FROM Term_JOB_HISTORY "
        SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
        SQLQ = SQLQ & " WHERE JH_EMPNBR=" & rsEMP("ED_EMPNBR") & ")"
        rsJOB.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsJOB.EOF Then
            rsEMP("JB_GRPCD_TABL") = "JBGC"
            rsEMP("JB_GRPCD") = rsJOB("JB_GRPCD")
        rsEMP.Update
        End If
        rsJOB.Close
        rsEMP.MoveNext
    Loop
    rsEMP.Close
End If
MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
End Sub

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%
Dim SubTotal
Dim wcCode
'for labels - sort by name always
' imbeded in report

Cri_Sorts = 0
' first set primary grouping

x% = 0
grpField$ = getEGroup(comGroup(0).Text)
grpField$ = Replace(grpField$, "HRJOB", "HREMP")

'As of Date
If Len(dlpDateRange(2).Text) > 0 Or Len(dlpDateRange(3).Text) > 0 Then
    Select Case cmbDateBased
        Case lStr("Original Hire Date")
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='Date Range: " & dlpDateRange(2).Text & " - " & dlpDateRange(3).Text & " (" & lStr("Original Hire Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As of Date: " & dlpDateRange(3).Text & " (" & lStr("Original Hire Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As From Date: " & dlpDateRange(2).Text & " (" & lStr("Original Hire Date") & ")'"
            End If
        Case lStr("Seniority Date")
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='Date Range: " & dlpDateRange(2).Text & " - " & dlpDateRange(3).Text & " (" & lStr("Seniority Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As of Date: " & dlpDateRange(3).Text & " (" & lStr("Seniority Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As From Date: " & dlpDateRange(2).Text & " (" & lStr("Seniority Date") & ")'"
            End If
        Case lStr("Last Hire Date")
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='Date Range: " & dlpDateRange(2).Text & " - " & dlpDateRange(3).Text & " (" & lStr("Last Hire Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As of Date: " & dlpDateRange(3).Text & " (" & lStr("Last Hire Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As From Date: " & dlpDateRange(2).Text & " (" & lStr("Last Hire Date") & ")'"
            End If
        Case lStr("Union Date")
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='Date Range: " & dlpDateRange(2).Text & " - " & dlpDateRange(3).Text & " (" & lStr("Union Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As of Date: " & dlpDateRange(3).Text & " (" & lStr("Union Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As From Date: " & dlpDateRange(2).Text & " (" & lStr("Union Date") & ")'"
            End If
        Case "Attendance Date"
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='Date Range: " & dlpDateRange(2).Text & " - " & dlpDateRange(3).Text & " (Attendance Date)'"
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As of Date: " & dlpDateRange(3).Text & " (Attendance Date)'"
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As From Date: " & dlpDateRange(2).Text & " (Attendance Date)'"
            End If
        Case lStr("User Defined Date")
            If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='Date Range: " & dlpDateRange(2).Text & " - " & dlpDateRange(3).Text & " (" & lStr("User Defined Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) = 0 And Len(dlpDateRange(3).Text) > 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As of Date: " & dlpDateRange(3).Text & " (" & lStr("User Defined Date") & ")'"
            ElseIf Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) = 0 Then
                Me.vbxCrystal.Formulas(1) = "AsOfDate='As From Date: " & dlpDateRange(2).Text & " (" & lStr("User Defined Date") & ")'"
            End If
    End Select
End If

Me.vbxCrystal.Formulas(2) = "Title='Paid Sick Hours Per Eligible Employee Report'"

If comGroup(0) = "(none)" Then
    Exit Function
End If

Y% = x% + 1
dscGroup$ = comGroup(x%).Text
dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(x%) = dscGroup$

grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(x%) = grpCond$

'Employee - Count
dscGroup$ = "DistinctCount ({HREMP.ED_EMPNBR}, " & grpField$ & ")"
dscGroup$ = "G1TotalEmp=" & dscGroup$
Me.vbxCrystal.Formulas(3) = dscGroup$

'Average Eligible Employee
dscGroup$ = "Count ({TOTSICKHRS.AD_EMPNBR}, " & grpField$ & ")"
dscGroup$ = "G1AvgEligEmp=" & dscGroup$
Me.vbxCrystal.Formulas(4) = dscGroup$

'Paid Sick Hrs per Eligible Employee
dscGroup$ = "if Sum({@PaidSickHrs}," & grpField$ & ") <> 0 then Sum({@PaidSickHrs}," & grpField$ & ")/{@G1AvgEligEmp}"
dscGroup$ = "G1PaidSickHrsPerEmp=" & dscGroup$
Me.vbxCrystal.Formulas(5) = dscGroup$

Cri_Sorts = z% ' next section number to format

End Function

Private Sub Get_Sick_Hrs_Entitle()
    Dim rsEMP As New ADODB.Recordset
    Dim rsAttnd As New ADODB.Recordset
    Dim rsEmpHis As New ADODB.Recordset
    Dim rsTmpAtt As New ADODB.Recordset
    Dim FromDt, ToDt
    Dim ESQLQ, SQLQHis
    
    Set cnRun = New ADODB.Connection
    cnRun.CommandTimeout = 600
    cnRun.Open glbAdoIHRDBW
    
    cnRun.BeginTrans
    cnRun.Execute "DELETE FROM TOTSICKHRS WHERE AD_WRKEMP='" & glbUserID & "'"
    cnRun.CommitTrans
    
    ESQLQ = glbstrSelCri
    ESQLQ = Replace(ESQLQ, "{", "")
    ESQLQ = Replace(ESQLQ, "}", "")
    ESQLQ = Replace(ESQLQ, "HREMP.", "")
    
    If glbSQL Or glbOracle Then
        ESQLQ = Replace(ESQLQ, "[", "(")
        ESQLQ = Replace(ESQLQ, "]", ")")
    End If
    
    If Trim(ESQLQ) <> "" Then
        If glbCompSerial = "S/N - 2347W" Then   'Surrey Place
            rsEMP.Open "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMP <> 'TR' AND ED_SICK > 0 AND " & ESQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            rsEMP.Open "SELECT ED_EMPNBR FROM HREMP WHERE ED_SICK > 0 AND " & ESQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
    Else
        If glbCompSerial = "S/N - 2347W" Then   'Surrey Place
            rsEMP.Open "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMP <> 'TR' AND ED_SICK > 0 ", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            rsEMP.Open "SELECT ED_EMPNBR FROM HREMP WHERE ED_SICK > 0", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
    End If
    
    Do While Not rsEMP.EOF
        If cmbDateBased = "Attendance Date" Then
            If IsDate(dlpDateRange(2)) And IsDate(dlpDateRange(3)) Then
                rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                SQLQHis = "SELECT SUM(AH_HRS) as TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND (AH_DOA >= " & Date_SQL(dlpDateRange(2)) & " AND AH_DOA <= " & Date_SQL(dlpDateRange(3)) & ")"
            ElseIf IsDate(dlpDateRange(2)) And (Not IsDate(dlpDateRange(3))) Then
                rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                SQLQHis = "SELECT SUM(AH_HRS) as TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND (AH_DOA >= " & Date_SQL(dlpDateRange(2)) & ")"
            ElseIf IsDate(dlpDateRange(3)) And (Not IsDate(dlpDateRange(2))) Then
                rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND (AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                SQLQHis = "SELECT SUM(AH_HRS) as TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND (AH_DOA <= " & Date_SQL(dlpDateRange(3)) & ")"
            Else
                rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "')", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                SQLQHis = "SELECT SUM(AH_HRS) as TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "')"
            End If
        Else
            rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "')", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            SQLQHis = "SELECT SUM(AH_HRS) as TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "')"
        End If
        If Not rsAttnd.EOF Then
            rsTmpAtt.Open "SELECT * FROM TOTSICKHRS WHERE AD_EMPNBR =" & rsEMP("ED_EMPNBR"), gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
            If rsTmpAtt.EOF Then rsTmpAtt.AddNew
            rsTmpAtt("AD_EMPNBR") = rsEMP("ED_EMPNBR")
            rsTmpAtt("AD_HRS") = IIf(IsNull(rsTmpAtt("AD_HRS")), 0, rsTmpAtt("AD_HRS")) + IIf(IsNull(rsAttnd("TOT")), 0, rsAttnd("TOT"))
            rsTmpAtt("AD_WRKEMP") = glbUserID
            rsTmpAtt.Update
            rsTmpAtt.Close
        End If
        rsAttnd.Close

        If chkInclAtt Then
            rsAttnd.Open SQLQHis, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsAttnd.EOF Then
                rsTmpAtt.Open "SELECT * FROM TOTSICKHRS WHERE AD_EMPNBR =" & rsEMP("ED_EMPNBR"), gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
                If rsTmpAtt.EOF Then rsTmpAtt.AddNew
                rsTmpAtt("AD_EMPNBR") = rsEMP("ED_EMPNBR")
                rsTmpAtt("AD_HRS") = IIf(IsNull(rsTmpAtt("AD_HRS")), 0, rsTmpAtt("AD_HRS")) + IIf(IsNull(rsAttnd("TOT")), 0, rsAttnd("TOT"))
                rsTmpAtt("AD_WRKEMP") = glbUserID
                rsTmpAtt.Update
                rsTmpAtt.Close
            End If
            rsAttnd.Close
        End If
                        
        rsEMP.MoveNext
    Loop
    rsEMP.Close

End Sub

Private Sub Get_Sick_Hrs_Entitle_Terminated()
    Dim rsEMP As New ADODB.Recordset
    Dim rsAttnd As New ADODB.Recordset
    Dim rsEmpHis As New ADODB.Recordset
    Dim rsTmpAtt As New ADODB.Recordset
    Dim FromDt, ToDt
    Dim ESQLQ, SQLQHis
        
        
    ESQLQ = glbstrSelCri
    ESQLQ = Replace(ESQLQ, "{", "")
    ESQLQ = Replace(ESQLQ, "}", "")
    ESQLQ = Replace(ESQLQ, "HREMP.", "")
    
    If glbSQL Or glbOracle Then
        ESQLQ = Replace(ESQLQ, "[", "(")
        ESQLQ = Replace(ESQLQ, "]", ")")
    End If
    
    If Trim(ESQLQ) <> "" Then
        'rsEMP.Open "SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_EMP <> 'TR' AND " & ESQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rsEMP.Open "SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_SICK > 0 AND " & ESQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Else
        'rsEMP.Open "SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_EMP <> 'TR'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rsEMP.Open "SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_SICK > 0", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    
    Do While Not rsEMP.EOF
        If cmbDateBased = "Attendance Date" Then
            If IsDate(dlpDateRange(2)) And IsDate(dlpDateRange(3)) Then
                rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            ElseIf IsDate(dlpDateRange(2)) And (Not IsDate(dlpDateRange(3))) Then
                rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            ElseIf IsDate(dlpDateRange(3)) And (Not IsDate(dlpDateRange(2))) Then
                rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND (AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            Else
                rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "')", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            End If
        Else
            rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "')", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        
        If Not rsAttnd.EOF Then
            rsTmpAtt.Open "SELECT * FROM TOTSICKHRS WHERE AD_EMPNBR =" & rsEMP("ED_EMPNBR"), gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
            If rsTmpAtt.EOF Then rsTmpAtt.AddNew
            rsTmpAtt("AD_EMPNBR") = rsEMP("ED_EMPNBR")
            rsTmpAtt("AD_HRS") = IIf(IsNull(rsTmpAtt("AD_HRS")), 0, rsTmpAtt("AD_HRS")) + IIf(IsNull(rsAttnd("TOT")), 0, rsAttnd("TOT"))
            rsTmpAtt("AD_WRKEMP") = glbUserID
            rsTmpAtt.Update
            rsTmpAtt.Close
        End If
        rsAttnd.Close
                
        rsEMP.MoveNext
    Loop
    rsEMP.Close
End Sub

Private Sub Get_Sick_Hrs()
    Dim rsEMP As New ADODB.Recordset
    Dim rsAttnd As New ADODB.Recordset
    Dim rsEmpHis As New ADODB.Recordset
    Dim rsTmpAtt As New ADODB.Recordset
    Dim FromDt, ToDt
    Dim ESQLQ, SQLQHis
    
    Set cnRun = New ADODB.Connection
    cnRun.CommandTimeout = 600
    cnRun.Open glbAdoIHRDBW
    
    cnRun.BeginTrans
    cnRun.Execute "DELETE FROM TOTSICKHRS WHERE AD_WRKEMP='" & glbUserID & "'"
    cnRun.CommitTrans
    
    
    ESQLQ = glbstrSelCri
    ESQLQ = Replace(ESQLQ, "{", "")
    ESQLQ = Replace(ESQLQ, "}", "")
    ESQLQ = Replace(ESQLQ, "HREMP.", "")
    
    If glbSQL Or glbOracle Then
        ESQLQ = Replace(ESQLQ, "[", "(")
        ESQLQ = Replace(ESQLQ, "]", ")")
    End If

    If Trim(ESQLQ) <> "" Then
        If glbCompSerial = "S/N - 2347W" Then   'Surrey Place
            rsEMP.Open "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMP <> 'TR' AND " & ESQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            rsEMP.Open "SELECT ED_EMPNBR FROM HREMP WHERE " & ESQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
    Else
        If glbCompSerial = "S/N - 2347W" Then   'Surrey Place
            rsEMP.Open "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMP <> 'TR'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            rsEMP.Open "SELECT ED_EMPNBR FROM HREMP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
    End If
    
    Do While Not rsEMP.EOF
        
        If IsDate(dlpDateRange(2)) And IsDate(dlpDateRange(3)) Then
            rsEmpHis.Open "SELECT EE_EMPNBR, EE_CHGDATE, EE_OLDBENEGROUP, EE_NEWBENEGROUP FROM HREMPHIS WHERE EE_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND (EE_CHGDATE >= " & Date_SQL(dlpDateRange(2).Text) & " AND EE_CHGDATE <= " & Date_SQL(dlpDateRange(3).Text) & ") ORDER BY EE_CHGDATE, EE_LDATE, EE_LTIME", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        ElseIf IsDate(dlpDateRange(2)) And (Not IsDate(dlpDateRange(3))) Then
            rsEmpHis.Open "SELECT EE_EMPNBR, EE_CHGDATE, EE_OLDBENEGROUP, EE_NEWBENEGROUP FROM HREMPHIS WHERE EE_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND (EE_CHGDATE >= " & Date_SQL(dlpDateRange(2).Text) & ") ORDER BY EE_CHGDATE, EE_LDATE, EE_LTIME", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        ElseIf IsDate(dlpDateRange(3)) And (Not IsDate(dlpDateRange(2))) Then
            rsEmpHis.Open "SELECT EE_EMPNBR, EE_CHGDATE, EE_OLDBENEGROUP, EE_NEWBENEGROUP FROM HREMPHIS WHERE EE_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND (EE_CHGDATE <= " & Date_SQL(dlpDateRange(3).Text) & ") ORDER BY EE_CHGDATE, EE_LDATE, EE_LTIME", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            rsEmpHis.Open "SELECT EE_EMPNBR, EE_CHGDATE, EE_OLDBENEGROUP, EE_NEWBENEGROUP FROM HREMPHIS WHERE EE_EMPNBR = " & rsEMP("ED_EMPNBR") & " ORDER BY EE_CHGDATE, EE_LDATE, EE_LTIME", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        
        If Not rsEmpHis.EOF Then
            FromDt = ""
            ToDt = ""
            Do While Not rsEmpHis.EOF
        
                If rsEmpHis("EE_NEWBENEGROUP") <> "" And (Not IsNull(rsEmpHis("EE_NEWBENEGROUP"))) And FromDt = "" Then
                    FromDt = rsEmpHis("EE_CHGDATE")
                    If ToDt <> "" Then
                        ToDt = ""
                    End If
                ElseIf (rsEmpHis("EE_NEWBENEGROUP") = "" Or (IsNull(rsEmpHis("EE_NEWBENEGROUP")))) And (rsEmpHis("EE_OLDBENEGROUP") <> "" Or (Not IsNull(rsEmpHis("EE_OLDBENEGROUP")))) Then
                    ToDt = rsEmpHis("EE_CHGDATE")
                End If
                                
                If FromDt <> "" And ToDt <> "" Then
                    If cmbDateBased = "Attendance Date" Then
                        If IsDate(dlpDateRange(2)) And IsDate(dlpDateRange(3)) Then
                            rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND AD_DOA <" & Date_SQL(ToDt) & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            SQLQHis = "SELECT SUM(AH_HRS) as TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AH_DOA >=" & Date_SQL(FromDt) & " AND AH_DOA <" & Date_SQL(ToDt) & " AND (AH_DOA >= " & Date_SQL(dlpDateRange(2)) & " AND AH_DOA <= " & Date_SQL(dlpDateRange(3)) & ")"
                        ElseIf IsDate(dlpDateRange(2)) And (Not IsDate(dlpDateRange(3))) Then
                            rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND AD_DOA <" & Date_SQL(ToDt) & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            SQLQHis = "SELECT SUM(AH_HRS) as TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AH_DOA >=" & Date_SQL(FromDt) & " AND AH_DOA <" & Date_SQL(ToDt) & " AND (AH_DOA >= " & Date_SQL(dlpDateRange(2)) & ")"
                        ElseIf IsDate(dlpDateRange(3)) And (Not IsDate(dlpDateRange(2))) Then
                            rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND AD_DOA <" & Date_SQL(ToDt) & " AND (AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            SQLQHis = "SELECT SUM(AH_HRS) as TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AH_DOA >=" & Date_SQL(FromDt) & " AND AH_DOA <" & Date_SQL(ToDt) & " AND (AH_DOA <= " & Date_SQL(dlpDateRange(3)) & ")"
                        Else
                            rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND AD_DOA <" & Date_SQL(ToDt), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            SQLQHis = "SELECT SUM(AH_HRS) as TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AH_DOA >=" & Date_SQL(FromDt) & " AND AH_DOA <" & Date_SQL(ToDt)
                        End If
                    Else
                        rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND AD_DOA <" & Date_SQL(ToDt), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        SQLQHis = "SELECT SUM(AH_HRS) as TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AH_DOA >=" & Date_SQL(FromDt) & " AND AH_DOA <" & Date_SQL(ToDt)
                    End If
                    
                    If Not rsAttnd.EOF Then
                        rsTmpAtt.Open "SELECT * FROM TOTSICKHRS WHERE AD_EMPNBR =" & rsEMP("ED_EMPNBR"), gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
                        If rsTmpAtt.EOF Then rsTmpAtt.AddNew
                        rsTmpAtt("AD_EMPNBR") = rsEMP("ED_EMPNBR")
                        rsTmpAtt("AD_HRS") = IIf(IsNull(rsTmpAtt("AD_HRS")), 0, rsTmpAtt("AD_HRS")) + IIf(IsNull(rsAttnd("TOT")), 0, rsAttnd("TOT"))
                        rsTmpAtt("AD_WRKEMP") = glbUserID
                        rsTmpAtt.Update
                        rsTmpAtt.Close
                    End If
                    rsAttnd.Close

                    If chkInclAtt Then
                        rsAttnd.Open SQLQHis, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsAttnd.EOF Then
                            rsTmpAtt.Open "SELECT * FROM TOTSICKHRS WHERE AD_EMPNBR =" & rsEMP("ED_EMPNBR"), gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
                            If rsTmpAtt.EOF Then rsTmpAtt.AddNew
                            rsTmpAtt("AD_EMPNBR") = rsEMP("ED_EMPNBR")
                            rsTmpAtt("AD_HRS") = IIf(IsNull(rsTmpAtt("AD_HRS")), 0, rsTmpAtt("AD_HRS")) + IIf(IsNull(rsAttnd("TOT")), 0, rsAttnd("TOT"))
                            rsTmpAtt("AD_WRKEMP") = glbUserID
                            rsTmpAtt.Update
                            rsTmpAtt.Close
                        End If
                        rsAttnd.Close
                    End If

                    FromDt = ""
                    ToDt = ""
                End If
                
                rsEmpHis.MoveNext
                
                If Not rsEmpHis.EOF Then
                    If rsEmpHis("EE_EMPNBR") <> rsEMP("ED_EMPNBR") And FromDt <> "" Then
                        If cmbDateBased = "Attendance Date" Then
                            If IsDate(dlpDateRange(2)) And IsDate(dlpDateRange(3)) Then
                                rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                SQLQHis = "SELECT SUM(AH_HRS) AS TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AH_DOA >=" & Date_SQL(FromDt) & " AND (AH_DOA >= " & Date_SQL(dlpDateRange(2)) & " AND AH_DOA <= " & Date_SQL(dlpDateRange(3)) & ")"
                            ElseIf IsDate(dlpDateRange(2)) And (Not IsDate(dlpDateRange(3))) Then
                                rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                SQLQHis = "SELECT SUM(AH_HRS) AS TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AH_DOA >=" & Date_SQL(FromDt) & " AND (AH_DOA >= " & Date_SQL(dlpDateRange(2)) & ")"
                            ElseIf IsDate(dlpDateRange(3)) And (Not IsDate(dlpDateRange(2))) Then
                                rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND (AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                SQLQHis = "SELECT SUM(AH_HRS) AS TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AH_DOA >=" & Date_SQL(FromDt) & " AND (AH_DOA <= " & Date_SQL(dlpDateRange(3)) & ")"
                            Else
                                rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                SQLQHis = "SELECT SUM(AH_HRS) AS TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AH_DOA >=" & Date_SQL(FromDt)
                            End If
                        Else
                            rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            SQLQHis = "SELECT SUM(AH_HRS) AS TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AH_DOA >=" & Date_SQL(FromDt)
                        End If
                        
                        If Not rsAttnd.EOF Then
                            rsTmpAtt.Open "SELECT * FROM TOTSICKHRS WHERE AD_EMPNBR =" & rsEMP("ED_EMPNBR"), gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
                            If rsTmpAtt.EOF Then rsTmpAtt.AddNew
                            rsTmpAtt("AD_EMPNBR") = rsEMP("ED_EMPNBR")
                            rsTmpAtt("AD_HRS") = IIf(IsNull(rsTmpAtt("AD_HRS")), 0, rsTmpAtt("AD_HRS")) + IIf(IsNull(rsAttnd("TOT")), 0, rsAttnd("TOT"))
                            rsTmpAtt("AD_WRKEMP") = glbUserID
                            rsTmpAtt.Update
                            rsTmpAtt.Close
                        End If
                        rsAttnd.Close
                        
                        If chkInclAtt Then
                            rsAttnd.Open SQLQHis, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsAttnd.EOF Then
                                rsTmpAtt.Open "SELECT * FROM TOTSICKHRS WHERE AD_EMPNBR =" & rsEMP("ED_EMPNBR"), gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
                                If rsTmpAtt.EOF Then rsTmpAtt.AddNew
                                rsTmpAtt("AD_EMPNBR") = rsEMP("ED_EMPNBR")
                                rsTmpAtt("AD_HRS") = IIf(IsNull(rsTmpAtt("AD_HRS")), 0, rsTmpAtt("AD_HRS")) + IIf(IsNull(rsAttnd("TOT")), 0, rsAttnd("TOT"))
                                rsTmpAtt("AD_WRKEMP") = glbUserID
                                rsTmpAtt.Update
                                rsTmpAtt.Close
                            End If
                            rsAttnd.Close
                        End If
                        
                        FromDt = ""
                        ToDt = ""
                    End If
                End If
            Loop
            
            If FromDt <> "" Then
                If cmbDateBased = "Attendance Date" Then
                    If IsDate(dlpDateRange(2)) And IsDate(dlpDateRange(3)) Then
                        rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        SQLQHis = "SELECT SUM(AH_HRS) AS TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AH_DOA >=" & Date_SQL(FromDt) & " AND (AH_DOA >= " & Date_SQL(dlpDateRange(2)) & " AND AH_DOA <= " & Date_SQL(dlpDateRange(3)) & ")"
                    ElseIf IsDate(dlpDateRange(2)) And (Not IsDate(dlpDateRange(3))) Then
                        rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        SQLQHis = "SELECT SUM(AH_HRS) AS TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AH_DOA >=" & Date_SQL(FromDt) & " AND (AH_DOA >= " & Date_SQL(dlpDateRange(2)) & ")"
                    ElseIf IsDate(dlpDateRange(3)) And (Not IsDate(dlpDateRange(2))) Then
                        rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND (AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        SQLQHis = "SELECT SUM(AH_HRS) AS TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AH_DOA >=" & Date_SQL(FromDt) & " AND (AH_DOA <= " & Date_SQL(dlpDateRange(3)) & ")"
                    Else
                        rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        SQLQHis = "SELECT SUM(AH_HRS) AS TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AH_DOA >=" & Date_SQL(FromDt)
                    End If
                Else
                    rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    SQLQHis = "SELECT SUM(AH_HRS) AS TOT FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AH_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AH_DOA >=" & Date_SQL(FromDt)
                End If
                
                If Not rsAttnd.EOF Then
                    rsTmpAtt.Open "SELECT * FROM TOTSICKHRS WHERE AD_EMPNBR =" & rsEMP("ED_EMPNBR"), gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
                    If rsTmpAtt.EOF Then rsTmpAtt.AddNew
                    rsTmpAtt("AD_EMPNBR") = rsEMP("ED_EMPNBR")
                    rsTmpAtt("AD_HRS") = IIf(IsNull(rsTmpAtt("AD_HRS")), 0, rsTmpAtt("AD_HRS")) + IIf(IsNull(rsAttnd("TOT")), 0, rsAttnd("TOT"))
                    rsTmpAtt("AD_WRKEMP") = glbUserID
                    rsTmpAtt.Update
                    rsTmpAtt.Close
                End If
                rsAttnd.Close
                
                If chkInclAtt Then
                    rsAttnd.Open SQLQHis, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsAttnd.EOF Then
                        rsTmpAtt.Open "SELECT * FROM TOTSICKHRS WHERE AD_EMPNBR =" & rsEMP("ED_EMPNBR"), gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
                        If rsTmpAtt.EOF Then rsTmpAtt.AddNew
                        rsTmpAtt("AD_EMPNBR") = rsEMP("ED_EMPNBR")
                        rsTmpAtt("AD_HRS") = IIf(IsNull(rsTmpAtt("AD_HRS")), 0, rsTmpAtt("AD_HRS")) + IIf(IsNull(rsAttnd("TOT")), 0, rsAttnd("TOT"))
                        rsTmpAtt("AD_WRKEMP") = glbUserID
                        rsTmpAtt.Update
                        rsTmpAtt.Close
                    End If
                    rsAttnd.Close
                End If
                
                FromDt = ""
                ToDt = ""
            End If
            
        End If
        rsEmpHis.Close
        
        rsEMP.MoveNext
    Loop
    rsEMP.Close
End Sub

Private Sub Get_Sick_Hrs_Terminated()
    Dim rsEMP As New ADODB.Recordset
    Dim rsAttnd As New ADODB.Recordset
    Dim rsEmpHis As New ADODB.Recordset
    Dim rsTmpAtt As New ADODB.Recordset
    Dim FromDt, ToDt
    Dim ESQLQ, SQLQHis
        
        
    ESQLQ = glbstrSelCri
    ESQLQ = Replace(ESQLQ, "{", "")
    ESQLQ = Replace(ESQLQ, "}", "")
    ESQLQ = Replace(ESQLQ, "HREMP.", "")
    
    If glbSQL Or glbOracle Then
        ESQLQ = Replace(ESQLQ, "[", "(")
        ESQLQ = Replace(ESQLQ, "]", ")")
    End If
    
    If Trim(ESQLQ) <> "" Then
        'rsEMP.Open "SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_EMP <> 'TR' AND " & ESQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rsEMP.Open "SELECT ED_EMPNBR FROM Term_HREMP WHERE " & ESQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Else
        'rsEMP.Open "SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_EMP <> 'TR'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rsEMP.Open "SELECT ED_EMPNBR FROM Term_HREMP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    
    Do While Not rsEMP.EOF
        
        If IsDate(dlpDateRange(2)) And IsDate(dlpDateRange(3)) Then
            rsEmpHis.Open "SELECT EE_EMPNBR, EE_CHGDATE, EE_OLDBENEGROUP, EE_NEWBENEGROUP FROM Term_HREMPHIS WHERE EE_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND (EE_CHGDATE >= " & Date_SQL(dlpDateRange(2).Text) & " AND EE_CHGDATE <= " & Date_SQL(dlpDateRange(3).Text) & ") ORDER BY EE_CHGDATE, EE_LDATE, EE_LTIME", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        ElseIf IsDate(dlpDateRange(2)) And (Not IsDate(dlpDateRange(3))) Then
            rsEmpHis.Open "SELECT EE_EMPNBR, EE_CHGDATE, EE_OLDBENEGROUP, EE_NEWBENEGROUP FROM Term_HREMPHIS WHERE EE_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND (EE_CHGDATE >= " & Date_SQL(dlpDateRange(2).Text) & ") ORDER BY EE_CHGDATE, EE_LDATE, EE_LTIME", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        ElseIf IsDate(dlpDateRange(3)) And (Not IsDate(dlpDateRange(2))) Then
            rsEmpHis.Open "SELECT EE_EMPNBR, EE_CHGDATE, EE_OLDBENEGROUP, EE_NEWBENEGROUP FROM Term_HREMPHIS WHERE EE_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND (EE_CHGDATE <= " & Date_SQL(dlpDateRange(3).Text) & ") ORDER BY EE_CHGDATE, EE_LDATE, EE_LTIME", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            rsEmpHis.Open "SELECT EE_EMPNBR, EE_CHGDATE, EE_OLDBENEGROUP, EE_NEWBENEGROUP FROM Term_HREMPHIS WHERE EE_EMPNBR = " & rsEMP("ED_EMPNBR") & " ORDER BY EE_CHGDATE, EE_LDATE, EE_LTIME", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        
        If Not rsEmpHis.EOF Then
            FromDt = ""
            ToDt = ""
            Do While Not rsEmpHis.EOF
        
                If rsEmpHis("EE_NEWBENEGROUP") <> "" And (Not IsNull(rsEmpHis("EE_NEWBENEGROUP"))) And FromDt = "" Then
                    FromDt = rsEmpHis("EE_CHGDATE")
                    If ToDt <> "" Then
                        ToDt = ""
                    End If
                ElseIf (rsEmpHis("EE_NEWBENEGROUP") = "" Or (IsNull(rsEmpHis("EE_NEWBENEGROUP")))) And (rsEmpHis("EE_OLDBENEGROUP") <> "" Or (Not IsNull(rsEmpHis("EE_OLDBENEGROUP")))) Then
                    ToDt = rsEmpHis("EE_CHGDATE")
                End If
                                
                If FromDt <> "" And ToDt <> "" Then
                    If cmbDateBased = "Attendance Date" Then
                        If IsDate(dlpDateRange(2)) And IsDate(dlpDateRange(3)) Then
                            rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND AD_DOA <" & Date_SQL(ToDt) & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        ElseIf IsDate(dlpDateRange(2)) And (Not IsDate(dlpDateRange(3))) Then
                            rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND AD_DOA <" & Date_SQL(ToDt) & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        ElseIf IsDate(dlpDateRange(3)) And (Not IsDate(dlpDateRange(2))) Then
                            rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND AD_DOA <" & Date_SQL(ToDt) & " AND (AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        Else
                            rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND AD_DOA <" & Date_SQL(ToDt), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        End If
                    Else
                        rsAttnd.Open "SELECT SUM(AD_HRS) as TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND AD_DOA <" & Date_SQL(ToDt), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    End If
                    
                    If Not rsAttnd.EOF Then
                        rsTmpAtt.Open "SELECT * FROM TOTSICKHRS WHERE AD_EMPNBR =" & rsEMP("ED_EMPNBR"), gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
                        If rsTmpAtt.EOF Then rsTmpAtt.AddNew
                        rsTmpAtt("AD_EMPNBR") = rsEMP("ED_EMPNBR")
                        rsTmpAtt("AD_HRS") = IIf(IsNull(rsTmpAtt("AD_HRS")), 0, rsTmpAtt("AD_HRS")) + IIf(IsNull(rsAttnd("TOT")), 0, rsAttnd("TOT"))
                        rsTmpAtt("AD_WRKEMP") = glbUserID
                        rsTmpAtt.Update
                        rsTmpAtt.Close
                    End If
                    rsAttnd.Close

                    FromDt = ""
                    ToDt = ""
                End If
                
                rsEmpHis.MoveNext
                
                If Not rsEmpHis.EOF Then
                    If rsEmpHis("EE_EMPNBR") <> rsEMP("ED_EMPNBR") And FromDt <> "" Then
                        If cmbDateBased = "Attendance Date" Then
                            If IsDate(dlpDateRange(2)) And IsDate(dlpDateRange(3)) Then
                                rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            ElseIf IsDate(dlpDateRange(2)) And (Not IsDate(dlpDateRange(3))) Then
                                rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            ElseIf IsDate(dlpDateRange(3)) And (Not IsDate(dlpDateRange(2))) Then
                                rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND (AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            Else
                                rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            End If
                        Else
                            rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        End If
                        
                        If Not rsAttnd.EOF Then
                            rsTmpAtt.Open "SELECT * FROM TOTSICKHRS WHERE AD_EMPNBR =" & rsEMP("ED_EMPNBR"), gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
                            If rsTmpAtt.EOF Then rsTmpAtt.AddNew
                            rsTmpAtt("AD_EMPNBR") = rsEMP("ED_EMPNBR")
                            rsTmpAtt("AD_HRS") = IIf(IsNull(rsTmpAtt("AD_HRS")), 0, rsTmpAtt("AD_HRS")) + IIf(IsNull(rsAttnd("TOT")), 0, rsAttnd("TOT"))
                            rsTmpAtt("AD_WRKEMP") = glbUserID
                            rsTmpAtt.Update
                            rsTmpAtt.Close
                        End If
                        rsAttnd.Close
                                                
                        FromDt = ""
                        ToDt = ""
                    End If
                End If
            Loop
            
            If FromDt <> "" Then
                If cmbDateBased = "Attendance Date" Then
                    If IsDate(dlpDateRange(2)) And IsDate(dlpDateRange(3)) Then
                        rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    ElseIf IsDate(dlpDateRange(2)) And (Not IsDate(dlpDateRange(3))) Then
                        rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(2)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    ElseIf IsDate(dlpDateRange(3)) And (Not IsDate(dlpDateRange(2))) Then
                        rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt) & " AND (AD_DOA <= " & Date_SQL(dlpDateRange(3)) & ")", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Else
                        rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    End If
                Else
                    rsAttnd.Open "SELECT SUM(AD_HRS) AS TOT FROM Term_ATTENDANCE WHERE AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " AND AD_REASON IN ('" & getCodes(clpCode(7).Text) & "') AND AD_DOA >=" & Date_SQL(FromDt), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                End If
                
                If Not rsAttnd.EOF Then
                    rsTmpAtt.Open "SELECT * FROM TOTSICKHRS WHERE AD_EMPNBR =" & rsEMP("ED_EMPNBR"), gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
                    If rsTmpAtt.EOF Then rsTmpAtt.AddNew
                    rsTmpAtt("AD_EMPNBR") = rsEMP("ED_EMPNBR")
                    rsTmpAtt("AD_HRS") = IIf(IsNull(rsTmpAtt("AD_HRS")), 0, rsTmpAtt("AD_HRS")) + IIf(IsNull(rsAttnd("TOT")), 0, rsAttnd("TOT"))
                    rsTmpAtt("AD_WRKEMP") = glbUserID
                    rsTmpAtt.Update
                    rsTmpAtt.Close
                End If
                rsAttnd.Close
                                
                FromDt = ""
                ToDt = ""
            End If
            
        End If
        rsEmpHis.Close
        
        rsEMP.MoveNext
    Loop
    rsEMP.Close
End Sub

Private Function CriCheck()
Dim x%

CriCheck = False
If Not clpDiv.ListChecker Then
    Exit Function
End If

If Not clpDept.ListChecker Then
    Exit Function
End If

For x% = 0 To 6
    If Not clpCode(x%).ListChecker Then
        Exit Function
    End If
Next x%

If Not clpPT.ListChecker Then
    Exit Function
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

If Not clpCode(7).ListChecker Then
    Exit Function
End If
If Len(Trim(clpCode(7).Text)) = 0 Then
    MsgBox "Sick Attendance Code cannot be blank"
    clpCode(7).SetFocus
    Exit Function
End If

CriCheck = True
End Function

Private Sub dlpDateRange_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = "FRMRPAIDSICHR"

Screen.MousePointer = HOURGLASS
If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If
Call comGrpLoad
Call setRptCaption(Me)
If glbLinamar Then clpCode(3).MaxLength = 8

If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6

If glbSyndesis Then
    Label2.Caption = "Position Grade"
    clpCode(5).Tag = "00-Enter Position Grade"
End If
Call INI_Controls(Me)
Screen.MousePointer = DEFAULT

'Display today's date for the Male vs Female report
dlpDateRange(3).Text = Format(Now, "Short Date")

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Resize()
scrFrame.Height = 8055
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 8400 Then
        scrControl.Value = 0
        scrFrame.Top = 120
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 5000 Then
            scrControl.Max = 5000
        Else
            scrControl.Max = 3200
        End If
        scrControl.Left = Me.Width - scrControl.Width - 120
        If Me.Height - scrControl.Top - 780 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - 780
        End If
    End If
    
    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 120
    'scrFrame.Height = Me.ScaleHeight - (scrHScroll.Height - 200)  '
    If Me.Width >= 9700 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        If Me.Width < 7000 Then
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
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub

Private Sub optEligible_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub scrControl_Change()
scrFrame.Top = 120 - scrControl.Value
End Sub

Private Sub scrHScroll_Change()
scrFrame.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
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
