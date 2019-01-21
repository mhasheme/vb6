VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRExtHire 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "External Hire Rate Report"
   ClientHeight    =   8745
   ClientLeft      =   570
   ClientTop       =   1095
   ClientWidth     =   10200
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
   ScaleHeight     =   8745
   ScaleWidth      =   10200
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   50
      SmallChange     =   4
      TabIndex        =   52
      Top             =   8415
      Width           =   9855
   End
   Begin VB.VScrollBar scrControl 
      Height          =   8115
      LargeChange     =   315
      Left            =   9600
      Max             =   100
      SmallChange     =   315
      TabIndex        =   51
      Top             =   240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame scrFrame 
      BorderStyle     =   0  'None
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
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
         TabIndex        =   11
         Tag             =   "Final Sort of Records"
         Top             =   7770
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
         TabIndex        =   10
         Tag             =   "First Level of grouping records"
         Top             =   7455
         Width           =   2325
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1995
         MaxLength       =   4
         TabIndex        =   8
         Tag             =   "00-Employee Position Shift"
         Top             =   3570
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.CheckBox chkTerm 
         Caption         =   "Include Terminated Employee"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Tag             =   "Check to include Terminated Employees"
         Top             =   4740
         Width           =   3135
      End
      Begin VB.Frame frmTerm 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   3420
         TabIndex        =   3
         Top             =   4710
         Visible         =   0   'False
         Width           =   4695
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   1
            Left            =   3060
            TabIndex        =   4
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
            TabIndex        =   5
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
            TabIndex        =   6
            Top             =   90
            Width           =   1095
         End
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
         Left            =   4530
         TabIndex        =   2
         Tag             =   "Check to include Attendance History"
         Top             =   4260
         Visible         =   0   'False
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
         ItemData        =   "fzExHire.frx":0000
         Left            =   1980
         List            =   "fzExHire.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Select Date Range Based On"
         Top             =   4290
         Width           =   2325
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   1680
         TabIndex        =   9
         Tag             =   "00-Enter Position Group Code"
         Top             =   2250
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
         TabIndex        =   12
         Tag             =   "00-Enter Status Code"
         Top             =   5820
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
         TabIndex        =   13
         Tag             =   "EDPT-Category"
         Top             =   1590
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
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   16
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
         TabIndex        =   17
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
         TabIndex        =   18
         Tag             =   "00-Enter Administered By Code"
         Top             =   2910
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
         TabIndex        =   19
         Tag             =   "00-Enter Section Code"
         Top             =   3240
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
         TabIndex        =   20
         Tag             =   "00-Enter Region Code"
         Top             =   2580
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
         TabIndex        =   21
         Tag             =   "10-Enter Employee Number"
         Top             =   1920
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
         TabIndex        =   22
         Tag             =   "40-Date from and including this date forward"
         Top             =   3900
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
         TabIndex        =   23
         Tag             =   "40-Date upto and including this date / As of Date"
         Top             =   3900
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
         TabIndex        =   24
         Tag             =   "01-Reason for change in position - Code"
         Top             =   5340
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "SDRC"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   8
         Left            =   1680
         TabIndex        =   25
         Tag             =   "01-Reason for change in position - Code"
         Top             =   6780
         Visible         =   0   'False
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "SDRC"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   9
         Left            =   1680
         TabIndex        =   26
         Tag             =   "00-Enter Status Code"
         Top             =   6300
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDEM"
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
         Top             =   6000
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
         TabIndex        =   46
         Top             =   1920
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
         TabIndex        =   45
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
         TabIndex        =   44
         Top             =   7245
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
         TabIndex        =   43
         Top             =   7485
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
         TabIndex        =   42
         Top             =   7800
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
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   2580
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
         TabIndex        =   39
         Top             =   2910
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
         TabIndex        =   38
         Top             =   2250
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
         TabIndex        =   37
         Top             =   3210
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
         TabIndex        =   36
         Top             =   1590
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
         TabIndex        =   35
         Top             =   3555
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
         TabIndex        =   34
         Top             =   3930
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
         TabIndex        =   33
         Top             =   4320
         Width           =   1110
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Reason"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   5505
         Width           =   1395
      End
      Begin VB.Label lblHireType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "External Hire"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   5280
         Width           =   1110
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Full Time &&  Part Time"
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
         Top             =   5760
         Width           =   1530
      End
      Begin VB.Label lblIntTrans 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Internal Transfers"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   6660
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label lblIntPos 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Reason"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   6885
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Status"
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
         TabIndex        =   27
         Top             =   6345
         Width           =   885
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   9600
      Top             =   8400
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
Attribute VB_Name = "frmRExtHire"
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

On Error GoTo PrntErr

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
On Error GoTo CRW_Err

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
    'cmbDateBased.AddItem lStr("Attendance Date")
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
Dim xType
Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For x% = 0 To 4
    If x% <> 2 Then
        Call Cri_Code(x%)
    End If
Next x%
Call Cri_Code(6)
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

' report name
If comGroup(0) <> "(none)" Then
    If Me.Caption = "External Hire Rate Report" Then
        strRName$ = glbIHRREPORTS & "rzExHir1.rpt"
    Else
        strRName$ = glbIHRREPORTS & "rzInHir1.rpt"
    End If
Else
    If Me.Caption = "External Hire Rate Report" Then
        strRName$ = glbIHRREPORTS & "rzExHire.rpt"
    Else
        strRName$ = glbIHRREPORTS & "rzInHire.rpt"
    End If
End If
Me.vbxCrystal.ReportFileName = strRName$

' set to sorting/grouping criteria
x% = Cri_Sorts()

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    If glbCompSerial = "S/N - 2347W" Then   'Surrey Place
        Me.vbxCrystal.SelectionFormula = "{HREMP.ED_WRKEMP}='" & glbUserID & "' AND {HREMP.ED_EMP} <> 'TERM' AND {HREMP.ED_PT} <> 'TR'"
    Else
        Me.vbxCrystal.SelectionFormula = "{HREMP.ED_WRKEMP}='" & glbUserID & "' AND {HREMP.ED_EMP} <> 'TERM'"
    End If
End If

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
        If comGroup(0) <> "(none)" Then
            Me.vbxCrystal.DataFiles(11) = glbIHRDB
            Me.vbxCrystal.DataFiles(12) = glbIHRDB
        End If
    End If
    ' window title if appropriate
    'Me.vbxCrystal.WindowTitle = "External Hire Rate Report"
    
    If Me.Caption = "External Hire Rate Report" Then
        Me.vbxCrystal.WindowTitle = "External Hire Rate Report"
    Else
        Me.vbxCrystal.WindowTitle = "Internal Transfers to Total Hires Ratio Report"
    End If
    
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
    MsgBox "External Hire Position Reason cannot be blank"
    clpCode(7).SetFocus
    Exit Function
End If

If Not clpCode(2).ListChecker Then
    Exit Function
End If
If Not clpCode(9).ListChecker Then
    Exit Function
End If

If Len(Trim(clpCode(2).Text)) = 0 And Len(Trim(clpCode(9).Text)) = 0 Then
    MsgBox "Both Full Time & Part Time Status and Other Status cannot be blank"
    clpCode(2).SetFocus
    Exit Function
End If
'If Len(Trim(clpCode(9).Text)) = 0 Then
'    MsgBox "Other Status cannot be blank"
'    clpCode(9).SetFocus
'    Exit Function
'End If

If Me.Caption <> "External Hire Rate Report" Then
    If Not clpCode(8).ListChecker Then
        Exit Function
    End If
    If Len(Trim(clpCode(8).Text)) = 0 Then
        MsgBox "Internal Transfers Position Reason cannot be blank"
        clpCode(8).SetFocus
        Exit Function
    End If
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
glbOnTop = "FRMREXTHIRE"

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
scrFrame.Height = 8295
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 8400 Then
        scrControl.Value = 0
        scrFrame.Top = 120
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 7000 Then
            scrControl.Max = 6000
        Else
            scrControl.Max = 2000
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


Private Sub optAvgServLvl_Click()
    'If optAvgServLvl.Value = True Then
    '    comGroup(0).Text = "(none)"
    'End If
End Sub

Private Sub optGender_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optMStatus_GotFocus()
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

Private Sub Calculate_Average_Age()
Dim ESQLQ, SQLQ, SQLQ1, SQLQ2
Dim rsEMP As New ADODB.Recordset
Dim TotEmp
Dim Age, TotAge
Dim AvgAge
        
    ESQLQ = glbstrSelCri
    ESQLQ = Replace(ESQLQ, "{", "")
    ESQLQ = Replace(ESQLQ, "}", "")
    ESQLQ = Replace(ESQLQ, "HREMP.", "")
    
    comGroup(0) = "(none)"
    
    'Active Employees
    SQLQ2 = "SELECT COUNT(ED_EMPNBR) AS EMP_COUNT FROM HREMP WHERE " & ESQLQ
    SQLQ1 = "SELECT ED_DOB FROM HREMP WHERE " & ESQLQ
    
    If Len(clpCode(3).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY "
        SQLQ = SQLQ & in_SQL(glbIHRDB)
        SQLQ = SQLQ & " WHERE JH_CURRENT<>0 "
        SQLQ = SQLQ & " AND JH_JOB IN (SELECT JB_CODE FROM HRJOB "
        SQLQ = SQLQ & in_SQL(glbIHRDB)
        If glbOracle Then
            SQLQ = SQLQ & " WHERE JB_GRPCD IN ['" & getCodes(clpCode(3).Text) & "']))"
        Else
            SQLQ = SQLQ & " WHERE JB_GRPCD IN ('" & getCodes(clpCode(3).Text) & "')))"
        End If
    End If

    'Total Employees
    SQLQ2 = SQLQ2 & SQLQ
    rsEMP.Open SQLQ2, gdbAdoIhr001, adOpenKeyset
    If Not rsEMP.EOF Then
        TotEmp = rsEMP("EMP_COUNT")
    Else
        TotEmp = 0
    End If
    rsEMP.Close
    
    'Calculate Total Age
    Age = 0
    TotAge = 0
    
    SQLQ1 = SQLQ1 & SQLQ
    rsEMP.Open SQLQ1, gdbAdoIhr001, adOpenKeyset
    If Not rsEMP.EOF Then
        rsEMP.MoveFirst
        
        Do While Not rsEMP.EOF
            If Not IsNull(rsEMP("ED_DOB")) Then
                Age = DateDiff("m", rsEMP("ED_DOB"), Now)
                If month(rsEMP("ED_DOB")) = month(Now) Then
                    If Day(Now) < Day(rsEMP("ED_DOB")) Then
                        Age = Age - 1
                    End If
                End If
                
                Age = Age / 12
                TotAge = TotAge + Age
                
            End If
            rsEMP.MoveNext
        Loop
    End If
    rsEMP.Close
    
    
    'Terminated Employees
    If chkTerm Then
        SQLQ2 = "SELECT COUNT(ED_EMPNBR) AS EMP_COUNT FROM Term_HREMP WHERE " & ESQLQ
        SQLQ1 = "SELECT ED_DOB FROM Term_HREMP WHERE " & ESQLQ
        
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
            If glbOracle Then
                SQLQ = SQLQ & " WHERE JB_GRPCD IN ['" & getCodes(clpCode(3).Text) & "']))"
            Else
                SQLQ = SQLQ & " WHERE JB_GRPCD IN ('" & getCodes(clpCode(3).Text) & "')))"
            End If
        End If
        
        'Calcualate Total Employees
        SQLQ2 = SQLQ2 & SQLQ
        rsEMP.Open SQLQ2, gdbAdoIhr001, adOpenKeyset
        If Not rsEMP.EOF Then
            TotEmp = TotEmp + rsEMP("EMP_COUNT")
        End If
        rsEMP.Close
        
        'Calculate Total Age
        Age = 0
        
        SQLQ1 = SQLQ1 & SQLQ
        rsEMP.Open SQLQ1, gdbAdoIhr001, adOpenKeyset
        If Not rsEMP.EOF Then
            rsEMP.MoveFirst
            
            Do While Not rsEMP.EOF
                If Not IsNull(rsEMP("ED_DOB")) Then
                    Age = DateDiff("m", rsEMP("ED_DOB"), Now)
                    If month(rsEMP("ED_DOB")) = month(Now) Then
                        If Day(Now) < Day(rsEMP("ED_DOB")) Then
                            Age = Age - 1
                        End If
                    End If
                    
                    Age = Age / 12
                    TotAge = TotAge + Age
                    
                End If
                rsEMP.MoveNext
            Loop
        End If
        rsEMP.Close
    End If
    
    'Average Age
    AvgAge = TotAge / TotEmp
    
    If chkTerm Then
        Me.vbxCrystal.Formulas(1) = "TotalDescEmployee='Total Number of Active and Terminated Employees: " & TotEmp & "'"
        Me.vbxCrystal.Formulas(2) = "TotalAge='Total Age of the Active and Terminated Employees: " & Format(TotAge, "#0.0") & " years '"
    Else
        Me.vbxCrystal.Formulas(1) = "TotalDescEmployee='Total Number of Employees: " & TotEmp & "'"
        Me.vbxCrystal.Formulas(2) = "TotalAge='Total Age of the Employees: " & Format(TotAge, "#0.0") & " years '"
    End If
    Me.vbxCrystal.Formulas(3) = "AvgAge='Average Employee Age: " & Format(AvgAge, "#0.0") & " years '"
    Me.vbxCrystal.Formulas(4) = "Title='Average Employee Age'"
        
End Sub


Private Sub Calculate_Average_Service_Level()
Dim ESQLQ, SQLQ, SQLQ1, SQLQ2
Dim fldVal
Dim rsEMP As New ADODB.Recordset
Dim TotEmp
Dim Service, TotService
Dim AvgServLvl
        
    ESQLQ = glbstrSelCri
    ESQLQ = Replace(ESQLQ, "{", "")
    ESQLQ = Replace(ESQLQ, "}", "")
    ESQLQ = Replace(ESQLQ, "HREMP.", "")
    
    comGroup(0) = "(none)"
    
    'Active Employees
    SQLQ2 = "SELECT COUNT(ED_EMPNBR) AS EMP_COUNT FROM HREMP WHERE " & ESQLQ
    SQLQ1 = "SELECT ED_DOH, ED_SENDTE, ED_LTHIRE, ED_UNION, ED_USRDAT1 FROM HREMP WHERE " & ESQLQ
    
    If Len(clpCode(3).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY "
        SQLQ = SQLQ & in_SQL(glbIHRDB)
        SQLQ = SQLQ & " WHERE JH_CURRENT<>0 "
        SQLQ = SQLQ & " AND JH_JOB IN (SELECT JB_CODE FROM HRJOB "
        SQLQ = SQLQ & in_SQL(glbIHRDB)
        
        If glbOracle Then
            SQLQ = SQLQ & " WHERE JB_GRPCD IN ['" & getCodes(clpCode(3).Text) & "']))"
        Else
            SQLQ = SQLQ & " WHERE JB_GRPCD IN ('" & getCodes(clpCode(3).Text) & "')))"
        End If
    End If

    'Total Employees
    SQLQ2 = SQLQ2 & SQLQ
    rsEMP.Open SQLQ2, gdbAdoIhr001, adOpenKeyset
    If Not rsEMP.EOF Then
        TotEmp = rsEMP("EMP_COUNT")
    Else
        TotEmp = 0
    End If
    rsEMP.Close
    
    'Calculate Total Service
    Service = 0
    TotService = 0
    
    SQLQ1 = SQLQ1 & SQLQ
    rsEMP.Open SQLQ1, gdbAdoIhr001, adOpenKeyset
    If Not rsEMP.EOF Then
        rsEMP.MoveFirst
        
        Do While Not rsEMP.EOF
            Select Case cmbDateBased
                Case lStr("Original Hire Date")
                    fldVal = rsEMP("ED_DOH")
                Case lStr("Seniority Date")
                    fldVal = rsEMP("ED_SENDTE")
                Case lStr("Last Hire Date")
                    fldVal = rsEMP("ED_LTHIRE")
                Case lStr("Union Date")
                    fldVal = rsEMP("ED_UNION")
                Case lStr("User Defined Date")
                    fldVal = rsEMP("ED_USRDAT1")
            End Select
        
            If (Not IsNull(fldVal)) And fldVal <> "" Then
                Service = DateDiff("m", fldVal, Now)
                If month(fldVal) = month(Now) Then
                    If Day(Now) < Day(fldVal) Then
                        Service = Service - 1
                    End If
                End If
                
                Service = Service / 12
                TotService = TotService + Service
                
            End If
            rsEMP.MoveNext
        Loop
    End If
    rsEMP.Close
    
    
    'Terminated Employees
    If chkTerm Then
        SQLQ2 = "SELECT COUNT(ED_EMPNBR) AS EMP_COUNT FROM Term_HREMP WHERE " & ESQLQ
        SQLQ1 = "SELECT ED_DOH, ED_SENDTE, ED_LTHIRE, ED_UNION, ED_USRDAT1 FROM Term_HREMP WHERE " & ESQLQ
        
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
            
            If glbOracle Then
                SQLQ = SQLQ & " WHERE JB_GRPCD IN ['" & getCodes(clpCode(3).Text) & "']))"
            Else
                SQLQ = SQLQ & " WHERE JB_GRPCD IN ('" & getCodes(clpCode(3).Text) & "')))"
            End If
        End If
        
        'Calcualate Total Employees
        SQLQ2 = SQLQ2 & SQLQ
        rsEMP.Open SQLQ2, gdbAdoIhr001, adOpenKeyset
        If Not rsEMP.EOF Then
            TotEmp = TotEmp + rsEMP("EMP_COUNT")
        End If
        rsEMP.Close
        
        'Calculate Total Service
        Service = 0
        
        SQLQ1 = SQLQ1 & SQLQ
        rsEMP.Open SQLQ1, gdbAdoIhr001, adOpenKeyset
        If Not rsEMP.EOF Then
            rsEMP.MoveFirst
            
            Do While Not rsEMP.EOF
                Select Case cmbDateBased
                    Case lStr("Original Hire Date")
                        fldVal = rsEMP("ED_DOH")
                    Case lStr("Seniority Date")
                        fldVal = rsEMP("ED_SENDTE")
                    Case lStr("Last Hire Date")
                        fldVal = rsEMP("ED_LTHIRE")
                    Case lStr("Union Date")
                        fldVal = rsEMP("ED_UNION")
                    Case lStr("User Defined Date")
                        fldVal = rsEMP("ED_USRDAT1")
                End Select
                
                If (Not IsNull(fldVal)) And fldVal <> "" Then
                    Service = DateDiff("m", fldVal, Now)
                    If month(fldVal) = month(Now) Then
                        If Day(Now) < Day(fldVal) Then
                            Service = Service - 1
                        End If
                    End If
                    
                    Service = Service / 12
                    TotService = TotService + Service
                    
                End If
                rsEMP.MoveNext
            Loop
        End If
        rsEMP.Close
    End If
    
    'Average Service Level
    AvgServLvl = TotService / TotEmp
    
    If chkTerm Then
        Me.vbxCrystal.Formulas(1) = "TotalDescEmployee='Total Number of Active and Terminated Employees: " & TotEmp & "'"
        Me.vbxCrystal.Formulas(2) = "TotalService='Total Service of the Active and Terminated Employees: " & Format(TotService, "#0.0") & " years '"
    Else
        Me.vbxCrystal.Formulas(1) = "TotalDescEmployee='Total Number of Employees: " & TotEmp & "'"
        Me.vbxCrystal.Formulas(2) = "TotalService='Total Service of the Employees: " & Format(TotService, "#0.0") & " years '"
    End If
    Me.vbxCrystal.Formulas(3) = "AvgServiceLvl='Average Employee Service Level: " & Format(AvgServLvl, "#0.0") & " years '"
    Me.vbxCrystal.Formulas(4) = "Title='Average Employee Service Level'"
    
    If Len(dlpDateRange(3).Text) > 0 Then
        Me.vbxCrystal.Formulas(5) = "AsOfDate='As of Date: " & dlpDateRange(3).Text & "'"
    End If
        
End Sub


Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%, dscGroup1
Dim SubTotal
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

If Me.Caption = "External Hire Rate Report" Then
    Me.vbxCrystal.Formulas(2) = "Title='External Hire Rate Report'"
Else
    Me.vbxCrystal.Formulas(2) = "Title='Internal Transfers to Total Hires Ratio Report'"
End If

If comGroup(0) = "(none)" Then
    GoTo Set_Criteria
    Exit Function
End If

Y% = x% + 1
dscGroup$ = comGroup(x%).Text
dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(x%) = dscGroup$

grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(x%) = grpCond$

'Total Employee - Count
dscGroup$ = "Count ({HREMP.ED_EMPNBR}, " & grpField$ & ")"
dscGroup$ = "G1TotalEmp=" & dscGroup$
Me.vbxCrystal.Formulas(3) = dscGroup$

Set_Criteria:
'External - FT/PT Status Codes:
'If Len(Trim(clpCode(2).Text)) > 0 Then
    'If InStr(getCodes(clpCode(2).Text), ",") <> 0 Then
        dscGroup$ = Replace(getCodes(clpCode(2).Text), ",", " OR {HREMP.ED_EMP} = ")
        dscGroup1 = Replace(getCodes(clpCode(7).Text), ",", " OR {HR_JOB_HISTORY.JH_JREASON} = ")
        dscGroup$ = "TotFTPT=if ({HREMP.ED_EMP} = '" & dscGroup$ & "') AND ({HR_JOB_HISTORY.JH_JREASON} = '" & dscGroup1 & "') then 1 else 0"
    'Else
    '    dscGroup$ = "TotFTPT=if ({HREMP.ED_EMP} = '" & getCodes(clpCode(2).Text) & "') AND ({HR_JOB_HISTORY.JH_JREASON} = '" & getCodes(clpCode(7).Text) & "') then 1 else 0"
    'End If
    Me.vbxCrystal.Formulas(5) = dscGroup$
'Else
'    If Len(Trim(clpCode(7).Text)) > 0 Then
'        dscGroup1 = Replace(getCodes(clpCode(7).Text), ",", " OR {HR_JOB_HISTORY.JH_JREASON} = ")
'        dscGroup$ = "TotFTPT=if ({HR_JOB_HISTORY.JH_JREASON} = '" & dscGroup1 & "') then 1 else 0"
'    Else
'        dscGroup$ = "TotFTPT=1"
'    End If
'    Me.vbxCrystal.Formulas(5) = dscGroup$
'End If


'External - NOT FT/PT Status Codes:
'If Len(Trim(clpCode(2).Text)) > 0 Then
'    If InStr(getCodes(clpCode(2).Text), ",") <> 0 Then
        dscGroup$ = Replace(getCodes(clpCode(9).Text), ",", " OR {HREMP.ED_EMP} = ")
        dscGroup1 = Replace(getCodes(clpCode(7).Text), ",", " OR {HR_JOB_HISTORY.JH_JREASON} = ")
        dscGroup$ = "TotOthStatus=if ({HREMP.ED_EMP} = '" & dscGroup$ & "') AND ({HR_JOB_HISTORY.JH_JREASON} = '" & dscGroup1 & "') then 1 else 0"
'    Else
'        dscGroup$ = "TotOthStatus=if ({HREMP.ED_EMP} = '" & getCodes(clpCode(9).Text) & "') AND ({HR_JOB_HISTORY.JH_JREASON} = '" & getCodes(clpCode(7).Text) & "') then 1 else 0"
'    End If
    Me.vbxCrystal.Formulas(6) = dscGroup$
'Else
'    If Len(Trim(clpCode(7).Text)) > 0 Then
'        dscGroup1 = Replace(getCodes(clpCode(7).Text), ",", " OR {HR_JOB_HISTORY.JH_JREASON} = ")
'        dscGroup$ = "TotOthStatus=if ({HR_JOB_HISTORY.JH_JREASON} = '" & dscGroup1 & "') then 1 else 0"
'    Else
'        dscGroup$ = "TotOthStatus=1"
'    End If
'    Me.vbxCrystal.Formulas(6) = dscGroup$
'End If

'Total External
If comGroup(0) = "(none)" Then
    dscGroup$ = "Sum({@TotFTPT}) + Sum({@TotOthStatus})"
    dscGroup$ = "TotExternal=" & dscGroup$
Else
    dscGroup$ = "Sum({@TotFTPT}, " & grpField$ & ") + Sum({@TotOthStatus}, " & grpField$ & ")"
    dscGroup$ = "G1TotExternal=" & dscGroup$
End If
Me.vbxCrystal.Formulas(4) = dscGroup$


If Me.Caption = "Internal Transfers to Total Hires Ratio Report" Then
    'For Internal Transfers
    'FT/PT Status Codes:
    'If Len(Trim(clpCode(2).Text)) > 0 Then
    '    If InStr(getCodes(clpCode(2).Text), ",") <> 0 Then
            dscGroup$ = Replace(getCodes(clpCode(2).Text), ",", " OR {HREMP.ED_EMP} = ")
            dscGroup1 = Replace(getCodes(clpCode(8).Text), ",", " OR {HR_JOB_HISTORY.JH_JREASON} = ")
            dscGroup$ = "TotFTPT1=if ({HREMP.ED_EMP} = '" & dscGroup$ & "') AND ({HR_JOB_HISTORY.JH_JREASON} = '" & dscGroup1 & "') then 1 else 0"
   '     Else
   '         dscGroup$ = "TotFTPT1=if ({HREMP.ED_EMP} = '" & getCodes(clpCode(2).Text) & "') AND ({HR_JOB_HISTORY.JH_JREASON} = '" & getCodes(clpCode(8).Text) & "') then 1 else 0"
   '     End If
        Me.vbxCrystal.Formulas(7) = dscGroup$
'    Else
'        If Len(Trim(clpCode(8).Text)) > 0 Then
'            dscGroup1 = Replace(getCodes(clpCode(8).Text), ",", " OR {HR_JOB_HISTORY.JH_JREASON} = ")
'            dscGroup$ = "TotFTPT1=if ({HR_JOB_HISTORY.JH_JREASON} = '" & dscGroup1 & "') then 1 else 0"
'        Else
'            dscGroup$ = "TotFTPT1=1"
'        End If
'        Me.vbxCrystal.Formulas(7) = dscGroup$
    'End If
    
    'NOT FT/PT Status Codes:
    'If Len(Trim(clpCode(2).Text)) > 0 Then
    '    If InStr(getCodes(clpCode(2).Text), ",") <> 0 Then
            dscGroup$ = Replace(getCodes(clpCode(9).Text), ",", " OR {HREMP.ED_EMP} = ")
            dscGroup1 = Replace(getCodes(clpCode(8).Text), ",", " OR {HR_JOB_HISTORY.JH_JREASON} = ")
            dscGroup$ = "TotOthStatus1=if ({HREMP.ED_EMP} = '" & dscGroup$ & "') AND ({HR_JOB_HISTORY.JH_JREASON} = '" & dscGroup1 & "') then 1 else 0"
    '    Else
    '        dscGroup$ = "TotOthStatus1=if ({HREMP.ED_EMP} = '" & getCodes(clpCode(2).Text) & "') AND ({HR_JOB_HISTORY.JH_JREASON} = '" & getCodes(clpCode(8).Text) & "') then 1 else 0"
    '    End If
        Me.vbxCrystal.Formulas(8) = dscGroup$
'    Else
'        If Len(Trim(clpCode(8).Text)) > 0 Then
'            dscGroup1 = Replace(getCodes(clpCode(8).Text), ",", " OR {HR_JOB_HISTORY.JH_JREASON} = ")
'            dscGroup$ = "TotOthStatus1=if ({HR_JOB_HISTORY.JH_JREASON} = '" & dscGroup1 & "') then 1 else 0"
'        Else
'            dscGroup$ = "TotOthStatus1=1"
'        End If
'        Me.vbxCrystal.Formulas(8) = dscGroup$
    'End If

    'Total Internal
    If comGroup(0) = "(none)" Then
        dscGroup$ = "Sum({@TotFTPT1}) + Sum({@TotOthStatus1})"
        dscGroup$ = "TotInternal=" & dscGroup$
    Else
        dscGroup$ = "Sum({@TotFTPT1}, " & grpField$ & ") + Sum({@TotOthStatus1}, " & grpField$ & ")"
        dscGroup$ = "G1TotInternal=" & dscGroup$
    End If
    Me.vbxCrystal.Formulas(9) = dscGroup$

End If

Cri_Sorts = z% ' next section number to format

End Function

