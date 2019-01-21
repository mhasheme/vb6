VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRDocType 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Document Type Report"
   ClientHeight    =   12210
   ClientLeft      =   375
   ClientTop       =   915
   ClientWidth     =   10695
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
   ScaleHeight     =   12210
   ScaleWidth      =   10695
   WindowState     =   2  'Maximized
   Begin VB.Frame scrFrame 
      BorderStyle     =   0  'None
      Height          =   12015
      Left            =   120
      TabIndex        =   39
      Top             =   0
      Width           =   10095
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
         TabIndex        =   36
         Tag             =   "First Level of grouping records"
         Top             =   11205
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
         TabIndex        =   37
         Tag             =   "Second level of grouping records"
         Top             =   11520
         Width           =   2325
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2355
         MaxLength       =   4
         TabIndex        =   12
         Tag             =   "00-Employee Position Shift"
         Top             =   4155
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.ListBox lstEmpFlagsList 
         Height          =   1425
         ItemData        =   "fzDocType.frx":0000
         Left            =   1935
         List            =   "fzDocType.frx":0002
         MultiSelect     =   1  'Simple
         TabIndex        =   16
         Top             =   6000
         Width           =   2655
      End
      Begin VB.CheckBox chkNEDocType 
         Caption         =   "Not"
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
         Left            =   1395
         TabIndex        =   10
         Tag             =   "00-Document Type Not Equal To"
         Top             =   3825
         Width           =   615
      End
      Begin VB.CheckBox chkResume 
         Caption         =   "Resume (Status/Dates screen)"
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
         Left            =   360
         TabIndex        =   13
         Tag             =   "00-Keyword Not Equal To"
         Top             =   4920
         Width           =   4095
      End
      Begin VB.CheckBox chkOtherInfo 
         Caption         =   "Other Information (Other Information screen)"
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
         Left            =   360
         TabIndex        =   14
         Tag             =   "00-Keyword Not Equal To"
         Top             =   5280
         Width           =   4095
      End
      Begin VB.CheckBox chkEmpFlags 
         Caption         =   "Employee Flags"
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
         Left            =   360
         TabIndex        =   15
         Tag             =   "00-Keyword Not Equal To"
         Top             =   5640
         Width           =   4095
      End
      Begin VB.CheckBox chkJobOffer 
         Caption         =   "Job Offer (Position screen)"
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
         Left            =   360
         TabIndex        =   18
         Tag             =   "00-Keyword Not Equal To"
         Top             =   7560
         Width           =   4095
      End
      Begin VB.CheckBox chkPerfReview 
         Caption         =   "Performance"
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
         Left            =   360
         TabIndex        =   20
         Tag             =   "00-Keyword Not Equal To"
         Top             =   7920
         Width           =   4095
      End
      Begin VB.CheckBox chkDollarEnt 
         Caption         =   "Dollar Entitlements"
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
         Left            =   360
         TabIndex        =   22
         Tag             =   "00-Keyword Not Equal To"
         Top             =   8280
         Width           =   4095
      End
      Begin VB.CheckBox chkAttendance 
         Caption         =   "Attendance"
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
         Left            =   360
         TabIndex        =   24
         Tag             =   "00-Keyword Not Equal To"
         Top             =   8640
         Width           =   4095
      End
      Begin VB.CheckBox chkAssociation 
         Caption         =   "Associations"
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
         Left            =   360
         TabIndex        =   26
         Tag             =   "00-Keyword Not Equal To"
         Top             =   9000
         Width           =   4095
      End
      Begin VB.CheckBox chkContEdu 
         Caption         =   "Continuing Education"
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
         Left            =   360
         TabIndex        =   28
         Tag             =   "00-Keyword Not Equal To"
         Top             =   9360
         Width           =   4095
      End
      Begin VB.CheckBox chkFormalEdu 
         Caption         =   "Formal Education"
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
         Left            =   360
         TabIndex        =   30
         Tag             =   "00-Keyword Not Equal To"
         Top             =   9720
         Width           =   4095
      End
      Begin VB.CheckBox chkCounselling 
         Caption         =   "Counselling"
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
         Left            =   360
         TabIndex        =   32
         Tag             =   "00-Keyword Not Equal To"
         Top             =   10080
         Width           =   4095
      End
      Begin VB.CheckBox chkComments 
         Caption         =   "Comments"
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
         Left            =   360
         TabIndex        =   34
         Tag             =   "00-Keyword Not Equal To"
         Top             =   10440
         Width           =   4095
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Top             =   510
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
         Left            =   2040
         TabIndex        =   9
         Tag             =   "00-Enter Section Code"
         Top             =   3480
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   2040
         TabIndex        =   8
         Tag             =   "00-Enter Administered By Code"
         Top             =   3150
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   7
         Tag             =   "00-Enter Region Code"
         Top             =   2820
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Tag             =   "EDPT-Category"
         Top             =   2160
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
         Index           =   2
         Left            =   2040
         TabIndex        =   4
         Tag             =   "00-Enter Status Code"
         Top             =   1830
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
         Index           =   1
         Left            =   2040
         TabIndex        =   3
         Tag             =   "00-Enter Union Code"
         Top             =   1500
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
         Left            =   2040
         TabIndex        =   2
         Tag             =   "00-Enter Location Code"
         Top             =   1170
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
         Top             =   840
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
      Begin Threed.SSCheck chkShowEmp 
         Height          =   225
         Left            =   6720
         TabIndex        =   40
         Tag             =   "If X-Show All Employees"
         Top             =   3000
         Visible         =   0   'False
         Width           =   2835
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   " Show All Employees"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27.5
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
         Top             =   2490
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
         Index           =   6
         Left            =   2040
         TabIndex        =   11
         Tag             =   "01-Document Type Code "
         Top             =   3810
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "DOCT"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   7
         Left            =   6180
         TabIndex        =   23
         Tag             =   "01-Entitlement - Code"
         Top             =   8265
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOL"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   8
         Left            =   6180
         TabIndex        =   25
         Tag             =   "01-Attendance Reason"
         Top             =   8625
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ADRE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   9
         Left            =   6180
         TabIndex        =   27
         Tag             =   "01-Association/Membership- Code"
         Top             =   8985
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "TDCD"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   10
         Left            =   6180
         TabIndex        =   29
         Tag             =   "00-Course Code"
         Top             =   9345
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ESCD"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   11
         Left            =   6180
         TabIndex        =   31
         Tag             =   "01-School - Code"
         Top             =   9705
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EUSC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   12
         Left            =   6180
         TabIndex        =   33
         Tag             =   "01-Counselling Type- Code"
         Top             =   10065
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "CETY"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   13
         Left            =   6180
         TabIndex        =   35
         Tag             =   "01-Comment Type- Code"
         Top             =   10425
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECOM"
      End
      Begin Threed.SSCheck chkShowAllEmpFlags 
         Height          =   225
         Left            =   4800
         TabIndex        =   17
         Tag             =   "If X-Show All Employee Flags"
         Top             =   7200
         Visible         =   0   'False
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   " Show All Employee Flags"
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   14
         Left            =   6180
         TabIndex        =   19
         Tag             =   "01-Reason for change in position - Code"
         Top             =   7530
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "SDRC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   15
         Left            =   6180
         TabIndex        =   21
         Tag             =   "00-Performance Rating - Code "
         Top             =   7905
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "SDPC"
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Left            =   4710
         TabIndex        =   67
         Top             =   7950
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Left            =   4710
         TabIndex        =   66
         Top             =   7575
         Visible         =   0   'False
         Width           =   375
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
         TabIndex        =   65
         Top             =   555
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
         TabIndex        =   64
         Top             =   885
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
         TabIndex        =   63
         Top             =   1545
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
         TabIndex        =   62
         Top             =   1875
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
         TabIndex        =   61
         Top             =   2535
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
         TabIndex        =   60
         Top             =   120
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
         Left            =   240
         TabIndex        =   59
         Top             =   11265
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
         Left            =   240
         TabIndex        =   58
         Top             =   11580
         Width           =   660
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
         Left            =   120
         TabIndex        =   57
         Top             =   10995
         Width           =   1575
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
         TabIndex        =   56
         Top             =   1215
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
         Left            =   120
         TabIndex        =   55
         Top             =   3195
         Width           =   1125
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
         Left            =   120
         TabIndex        =   54
         Top             =   2865
         Width           =   510
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
         Left            =   120
         TabIndex        =   53
         Top             =   3525
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
         TabIndex        =   52
         Top             =   2205
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
         TabIndex        =   51
         Top             =   4200
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Index           =   7
         Left            =   4710
         TabIndex        =   50
         Top             =   8310
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document Type"
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
         TabIndex        =   49
         Top             =   3855
         Width           =   1140
      End
      Begin VB.Label lblEmpFlagList 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select at least ONE Employee Flag :"
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
         Height          =   585
         Left            =   570
         TabIndex        =   48
         Top             =   6000
         Width           =   1215
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Document:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   4560
         Width           =   1635
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Index           =   8
         Left            =   4710
         TabIndex        =   46
         Top             =   8670
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Index           =   9
         Left            =   4710
         TabIndex        =   45
         Top             =   9030
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Index           =   10
         Left            =   4710
         TabIndex        =   44
         Top             =   9390
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Index           =   11
         Left            =   4710
         TabIndex        =   43
         Top             =   9750
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Index           =   12
         Left            =   4710
         TabIndex        =   42
         Top             =   10110
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Index           =   13
         Left            =   4710
         TabIndex        =   41
         Top             =   10470
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.VScrollBar scrControl 
      Height          =   12015
      LargeChange     =   300
      Left            =   10250
      Max             =   4000
      SmallChange     =   300
      TabIndex        =   38
      Top             =   0
      Width           =   285
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   10080
      Top             =   12240
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
End
Attribute VB_Name = "frmRDocType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fNEflag As Boolean
Private strGroup As String
Const Excel2007 = 12

Private Sub chkAssociation_Click()
    If chkAssociation Then
        lblCode(9).Visible = True
        'lblCode(9).Top = 9030
        lblCode(9).Caption = lStr("Associations")
        'clpCode(7).Visible = False
        'clpCode(8).Visible = False
        clpCode(9).Visible = True
        'clpCode(10).Visible = False
        'clpCode(11).Visible = False
        'clpCode(12).Visible = False
        'clpCode(13).Visible = False
    Else
        lblCode(9).Visible = False
        clpCode(9).Visible = False
    End If
End Sub

Private Sub chkAttendance_Click()
    If chkAttendance Then
        lblCode(8).Visible = True
        'lblCode(8).Top = 8670
        lblCode(8).Caption = "Reason"
        'clpCode(7).Visible = False
        clpCode(8).Visible = True
        'clpCode(9).Visible = False
        'clpCode(10).Visible = False
        'clpCode(11).Visible = False
        'clpCode(12).Visible = False
        'clpCode(13).Visible = False
    Else
        lblCode(8).Visible = False
        clpCode(8).Visible = False
    End If
End Sub

Private Sub chkComments_Click()
    If chkComments Then
        lblCode(13).Visible = True
        'lblCode(13).Top = 10470
        lblCode(13).Caption = "Type"
        'clpCode(7).Visible = False
        'clpCode(8).Visible = False
        'clpCode(9).Visible = False
        'clpCode(10).Visible = False
        'clpCode(11).Visible = False
        'clpCode(12).Visible = False
        clpCode(13).Visible = True
    Else
        lblCode(13).Visible = False
        clpCode(13).Visible = False
    End If
End Sub

Private Sub chkContEdu_Click()
    If chkContEdu Then
        lblCode(10).Visible = True
        'lblCode(10).Top = 9390
        lblCode(10).Caption = "Course Code"
        'clpCode(7).Visible = False
        'clpCode(8).Visible = False
        'clpCode(9).Visible = False
        clpCode(10).Visible = True
        'clpCode(11).Visible = False
        'clpCode(12).Visible = False
        'clpCode(13).Visible = False
    Else
        lblCode(10).Visible = False
        clpCode(10).Visible = False
    End If
End Sub

Private Sub chkCounselling_Click()
    If chkCounselling Then
        lblCode(12).Visible = True
        'lblCode(12).Top = 10110
        lblCode(12).Caption = "Type"
        'clpCode(7).Visible = False
        'clpCode(8).Visible = False
        'clpCode(9).Visible = False
        'clpCode(10).Visible = False
        'clpCode(11).Visible = False
        clpCode(12).Visible = True
        'clpCode(13).Visible = False
    Else
        lblCode(12).Visible = False
        clpCode(12).Visible = False
    End If
End Sub

Private Sub chkDollarEnt_Click()
    If chkDollarEnt Then
        lblCode(7).Visible = True
        'lblCode(7).Top = 8310
        lblCode(7).Caption = "Entitlement"
        clpCode(7).Visible = True
        'clpCode(8).Visible = False
        'clpCode(9).Visible = False
        'clpCode(10).Visible = False
        'clpCode(11).Visible = False
        'clpCode(12).Visible = False
        'clpCode(13).Visible = False
    Else
        lblCode(7).Visible = False
        clpCode(7).Visible = False
    End If
End Sub

Private Sub chkEmpFlags_Click()
    If chkEmpFlags Then
        'clpCode(7).Visible = False
        'clpCode(8).Visible = False
        'clpCode(9).Visible = False
        'clpCode(10).Visible = False
        'clpCode(11).Visible = False
        'clpCode(12).Visible = False
        'clpCode(13).Visible = False
        chkShowAllEmpFlags.Visible = True
    Else
        chkShowAllEmpFlags.Visible = False
    End If
End Sub

Private Sub chkFormalEdu_Click()
    If chkFormalEdu Then
        lblCode(11).Visible = True
        'lblCode(11).Top = 9750
        lblCode(11).Caption = "School"
        'clpCode(7).Visible = False
        'clpCode(8).Visible = False
        'clpCode(9).Visible = False
        'clpCode(10).Visible = False
        clpCode(11).Visible = True
        'clpCode(12).Visible = False
        'clpCode(13).Visible = False
    Else
        lblCode(11).Visible = False
        clpCode(11).Visible = False
    End If
End Sub

Private Sub chkJobOffer_Click()
    If chkJobOffer Then
        lblCode(0).Visible = True
        'lblCode(7).Top = 8310
        lblCode(0).Caption = "Reason for Change"
        clpCode(14).Visible = True
        'clpCode(8).Visible = False
        'clpCode(9).Visible = False
        'clpCode(10).Visible = False
        'clpCode(11).Visible = False
        'clpCode(12).Visible = False
        'clpCode(13).Visible = False
    Else
        lblCode(0).Visible = False
        clpCode(14).Visible = False
    End If
End Sub

Private Sub chkPerfReview_Click()
    If chkPerfReview Then
        lblCode(1).Visible = True
        'lblCode(7).Top = 8310
        lblCode(1).Caption = lStr("Performance Rating")
        clpCode(15).Visible = True
        'clpCode(8).Visible = False
        'clpCode(9).Visible = False
        'clpCode(10).Visible = False
        'clpCode(11).Visible = False
        'clpCode(12).Visible = False
        'clpCode(13).Visible = False
    Else
        lblCode(1).Visible = False
        clpCode(15).Visible = False
    End If
End Sub

Private Sub chkShowAllEmpFlags_Click(Value As Integer)
    Dim Y As Integer
    
    If chkShowAllEmpFlags Then
        For Y = 0 To lstEmpFlagsList.ListCount - 1
            lstEmpFlagsList.selected(Y) = True
        Next Y
        lstEmpFlagsList.ListIndex = 0
    Else
        For Y = 0 To lstEmpFlagsList.ListCount - 1
            lstEmpFlagsList.selected(Y) = False
        Next Y
        lstEmpFlagsList.ListIndex = 1
        lstEmpFlagsList.ListIndex = 0
    End If
End Sub

Private Sub chkShowAllEmpFlags_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkShowEmp_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then

    If Not PrtForm("Document Type Report Criteria", Me) Then Exit Sub
    
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

Public Sub cmdView_Click()
Dim X%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
      Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Screen.MousePointer = HOURGLASS
    
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
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ENTITLEMENTS", "VIEW")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub comEmpType_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Document Type")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Employee Number"
    comGroup(1).ListIndex = 0

End Sub

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
    Case 5: strCd$ = "HREMP.ED_SECTION"
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

Private Sub Cri_DocType()

Dim EECri As String, OneSet%, X%
Dim strC2, strCx As String
Dim strCa$
Dim NECri As String

OneSet% = False

If Len(clpCode(6).Text) = 0 Then Exit Sub
strCa$ = "HR_EMPDOCS_WRK.DC_DOCTYPE"

If chkNEDocType Then
    If Len(clpCode(6).Text) > 0 Then
        fNEflag = True
        NECri = NECri & " or (" & strCa$ & " IN ('" & Replace(clpCode(6).Text, ",", "','") & "'))"
    End If
Else
    If Len(clpCode(6).Text) > 0 Then
        EECri = EECri & "({" & strCa$ & "} IN ['" & Replace(clpCode(6).Text, ",", "','") & "'])"
        OneSet% = OneSet% - 1
        If OneSet% > 0 Then
            EECri = EECri & " OR "
        Else
            EECri = EECri
        End If
    End If
End If

If Len(EECri) > 0 Then
    ' dkostka - 10/19/2001 - Add brackets around this criteria if there's an OR in it.  If
    '   we don't, things get confused and security gets ignored.
    If InStr(EECri, "OR") > 0 Then EECri = "(" & EECri & ")"
    
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & "(" & EECri & ")"
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If

End Sub

Private Function Cri_SetAll()
Dim X%, strRName$

Cri_SetAll = False
On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

fNEflag = False
glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div
For X% = 0 To 5
    Call Cri_Code(X%)
Next X%
Call Cri_PT
Call Cri_Shift
Call Cri_EE

'glbstrSelCri = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
'Cri_SetAll = XLWriter_Employee_DocumentType

'Populate the work table
Call Create_EmployeeDocumentList

'Report Name
If comGroup(0) <> "(none)" Then
    strRName$ = glbIHRREPORTS & "rzEmpDoc.rpt"
Else
    strRName$ = glbIHRREPORTS & "rzEmpDoc1.rpt"
End If
Me.vbxCrystal.ReportFileName = strRName$

'set to sorting/grouping criteria
X% = Cri_Sorts()   ' returns number of sections formated

If Len(clpCode(6).Text) > 0 Then
    Call Cri_DocType
End If

If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If
 
Me.vbxCrystal.Connect = RptODBC_SQL

  
Cri_SetAll = True
Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Document Type", "Document Type Report", "Select")
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

'for labels - sort by name always
'imbeded in report

Cri_Sorts = 0
'first set primary grouping
Y% = 0
grpField$ = getEGroup(comGroup(0).Text)

Call setRptLabel(Me, 0)

If comGroup(0) = "(none)" Then
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
    End Select
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$

    Exit Function
End If
    
Y% = X% + 1
dscGroup$ = comGroup(X%).Text
dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(X%) = dscGroup$

grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(X%) = grpCond$

strSFormat$ = "GH1;T;T;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1
strSFormat$ = "GF1;T;X;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1

GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@EFullName}"
    Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
End Select
grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$

Cri_Sorts = z% ' next section number to format

End Function

Private Function CriCheck()
Dim X%
Dim empFlagSel As Boolean
Dim Y As Integer

CriCheck = False

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

If Not clpPT.ListChecker Then
'If Len(clpPT) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

For X% = 0 To 5
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

If Not elpEEID.ListChecker Then
    Exit Function
End If

'If chkNEDocType Then
'    If Len(clpCode(6).Text) = 0 Then
'        MsgBox lStr("Document Type cannot be blank when 'NOT' equal to is selected.")
'        clpCode(6).SetFocus
'        Exit Function
'    End If
'End If

    'Mandatory fields
    If chkEmpFlags Then
        empFlagSel = False
        'Go through the flags list to look for at least one selected Flag
        For Y = 0 To lstEmpFlagsList.ListCount - 1
            If lstEmpFlagsList.selected(Y) Then
                empFlagSel = True
                Exit For
            Else
                empFlagSel = False
            End If
        Next Y
        
        If empFlagSel = False Then
            MsgBox "At least one Employee Flag must be selected."
            lstEmpFlagsList.SetFocus
            Exit Function
        End If
    End If
    
    If chkJobOffer Then
        'lblCode.Caption = "Entitlement"
        If Len(clpCode(14)) > 0 Then
            If clpCode(14).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode(0).Caption & " code."
                clpCode(14).SetFocus
                Exit Function
            End If
        'Else
        '    MsgBox lblCode(7).Caption & " code is required"
        '    clpCode(7).SetFocus
        '    Exit Function
        End If
    End If
    
    If chkPerfReview Then
        'lblCode.Caption = "Entitlement"
        If Len(clpCode(15)) > 0 Then
            If clpCode(15).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode(1).Caption & " code."
                clpCode(15).SetFocus
                Exit Function
            End If
        'Else
        '    MsgBox lblCode(7).Caption & " code is required"
        '    clpCode(7).SetFocus
        '    Exit Function
        End If
    End If
    
    If chkDollarEnt Then
        'lblCode.Caption = "Entitlement"
        If Len(clpCode(7)) > 0 Then
            If clpCode(7).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode(7).Caption & " code."
                clpCode(7).SetFocus
                Exit Function
            End If
        'Else
        '    MsgBox lblCode(7).Caption & " code is required"
        '    clpCode(7).SetFocus
        '    Exit Function
        End If
    End If
    
    'Attendance
    If chkAttendance Then
        'lblCode.Caption = "Reason"
        If Len(clpCode(8)) > 0 Then
            If clpCode(8).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode(8).Caption & " code."
                clpCode(8).SetFocus
                Exit Function
            End If
        End If
    End If
        
    'Association
    If chkAssociation Then
        'lblCode.Caption = "Associations"
        If Len(clpCode(9)) > 0 Then
            If clpCode(9).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode(9).Caption & " code."
                clpCode(9).SetFocus
                Exit Function
            End If
        End If
    End If
        
    'Continuing Education
    If chkContEdu Then
        'lblCode.Caption = "Course Code"
        If Len(clpCode(10)) > 0 Then
            If clpCode(10).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode(10).Caption & " code."
                clpCode(10).SetFocus
                Exit Function
            End If
        End If
    End If
    
    'Formal Education
    If chkFormalEdu Then
        'lblCode.Caption = "School"
        If Len(clpCode(11)) > 0 Then
            If clpCode(11).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode(11).Caption & " code."
                clpCode(11).SetFocus
                Exit Function
            End If
        End If
    End If
    
    'Counseling
    If chkCounselling Then
        'lblCode.Caption = "Type"
        If Len(clpCode(12)) > 0 Then
            If clpCode(12).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode(12).Caption & " code."
                clpCode(12).SetFocus
                Exit Function
            End If
        End If
    End If
    
    'Comments
    If chkComments Then
        'lblCode.Caption = "Type"
        If Len(clpCode(13)) > 0 Then
            If clpCode(13).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode(13).Caption & " code."
                clpCode(13).SetFocus
                Exit Function
            End If
        End If
    End If

If chkResume Or chkOtherInfo Or chkEmpFlags Or chkJobOffer Or chkPerfReview Or chkDollarEnt Or _
    chkAttendance Or chkAssociation Or chkContEdu Or chkFormalEdu Or chkCounselling Or chkComments Then
    'do not thing - negative statement not working that's why doing the checking this way
Else
    MsgBox "At least one Type of Document must be selected."
    chkResume.SetFocus
    Exit Function
End If

CriCheck = True

End Function

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim Y As Integer

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = "FRMRDOCTYPE"

Screen.MousePointer = HOURGLASS

frmRDocType.Caption = "Document Type Report"

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If
Call comGrpLoad
Call setRptCaption(Me)

chkPerfReview.Caption = lStr(chkPerfReview.Caption)
chkAssociation.Caption = lStr(chkAssociation.Caption)
chkComments.Caption = lStr(chkComments.Caption)
chkCounselling.Caption = lStr("Counseling")

If glbLinamar Then clpCode(3).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

'Load Employee Flags list
lstEmpFlagsList.Clear
For Y = 1 To 20
    If lStr("Employee Flag " & Y) <> "" Then
        lstEmpFlagsList.AddItem lStr("Employee Flag " & Y)
    End If
Next Y

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT


End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Resize()
On Error GoTo Eh
Dim c As Long

If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    If Me.Height >= 13000 Then
        scrControl.Value = 0
        
        scrFrame.Top = 120
        
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        scrControl.Left = Me.ScaleWidth - scrControl.Width - 20
        scrControl.Height = Me.Height - 1000
        
        scrControl.Max = 6500
    End If
End If

exH:
    Exit Sub
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "Document Type Report", "View")
    Resume exH

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub

Private Sub lstEmpFlagsList_Click()
    'chkShowAllEmpFlags.Value = False
End Sub

Private Sub scrControl_Change()
    scrFrame.Top = 120 - scrControl.Value
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

Private Function XLWriter_Employee_DocumentType()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsGroup As New ADODB.Recordset
    Dim rsHRDoc_DocType As New ADODB.Recordset
    Dim rsHREmpOther As New ADODB.Recordset
    Dim SQLQ, sSQLQ As String
    Dim exApp As Object
    Dim exBook As Object
    Dim exSheet As Object
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim xCol As Long
    Dim I, totNum
    Dim xExcelRptPath  As String
    Dim strTime As String
    Dim xRptHeadRange As String

    On Error GoTo Err_XLWriter_Employee_DocumentType
    
    Screen.MousePointer = HOURGLASS
    
    glbstrSelCri = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
            
    'Ticket #22034 - Get Excel reports path
    If gsTRAININGMATRIX Then
        xExcelRptPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xExcelRptPath) = 0 Then
        xExcelRptPath = glbIHRREPORTS
    End If
            
    'File to export to
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "DocumentTypeTmp.xls"
    xlsFileMat = xExcelRptPath & IIf(Right(xExcelRptPath, 1) = "\", "", "\") & "DocumentType" & Trim(glbUserID) & ".xls"

    If Dir(xlsFileTmp) = "" Then
        Screen.MousePointer = DEFAULT
        MsgBox "There is no " & xlsFileTmp
        Exit Function
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

    'Create Grouping recordset
    SQLQ = getGroupField
    rsGroup.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsGroup.EOF Then
        strTime = Time$
        Do While Not rsGroup.EOF
            If comGroup(0).Text = "(none)" Then
                Set exSheet = exBook.Sheets(1)
                exSheet.name = "All"
            Else
                exBook.Worksheets(1).Copy After:=exBook.Worksheets(exBook.Worksheets.count)
                Set exSheet = exBook.Worksheets(exBook.Worksheets.count)
                
                If IsNull(rsGroup("TB_DESC")) Then
                    exSheet.name = "None"
                Else
                    exSheet.name = Left(rsGroup("TB_DESC"), 30)
                End If
            End If
            
            'Get the Employees for the Report
            SQLQ = "DELETE FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
            gdbAdoIhr001W.Execute SQLQ
            If Not glbSQL And Not glbOracle Then Call Pause(3)
            
            SQLQ = "INSERT INTO HR_EMPLIST_WRK (TT_EMPNBR, TT_WRKEMP) " & in_SQL(glbIHRDBW)
            SQLQ = SQLQ & " SELECT HREMP.ED_EMPNBR, '" & glbUserID & "' as TT_WRKEMP FROM HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR INNER JOIN HREMP_OTHER ON HREMP.ED_EMPNBR = HREMP_OTHER.ER_EMPNBR "
            SQLQ = SQLQ & " WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) And " & glbstrSelCri
            If IsNull(rsGroup("grpField")) Then
                If glbSQL Or glbOracle Then
                    SQLQ = SQLQ & " and " & strGroup & " IS NULL "
                Else
                    SQLQ = SQLQ & " and isnull(" & strGroup & ") = true "
                End If
            Else
                If strGroup = "HR_JOB_HISTORY.JH_REPTAU" Then 'Supervisor is a number
                    SQLQ = SQLQ & " and " & strGroup & "=" & rsGroup("grpField") & " "
                Else
                    SQLQ = SQLQ & " and " & strGroup & "='" & Replace(rsGroup("grpField"), "'", "''") & "' "
                End If
            End If
            gdbAdoIhr001.Execute SQLQ

            If Not glbSQL And Not glbOracle Then Call Pause(3)
            
            'Build recordset of employees to print in the report based on selection criteria a grouping
            SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR IN "
            SQLQ = SQLQ & " (SELECT HR_EMPLIST_WRK.TT_EMPNBR FROM HR_EMPLIST_WRK WHERE HR_EMPLIST_WRK.TT_WRKEMP='" & glbUserID & "') "
            SQLQ = SQLQ & " ORDER BY "
            If comGroup(1).ListIndex = 0 Then
                SQLQ = SQLQ & "HREMP.ED_SURNAME,HREMP.ED_FNAME "
            ElseIf comGroup(1).ListIndex = 1 Then
                SQLQ = SQLQ & "HREMP.ED_EMPNBR"
            End If
            rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            
            'Sheet Labels
            exSheet.Cells(1, 1) = glbCompName
            exSheet.Cells(3, 1) = "Date: " & Format(Now, "mmm dd, yyyy")
            exSheet.Cells(4, 1) = "Time: " & strTime
            exSheet.Cells(6, 1) = "Group by " & comGroup(0).Text
            
            xCol = 4
            xRptHeadRange = ""
            
            'Show All Employment Dates
            'If chkEmpDates Or chkShowAllEmpDates Then
                exSheet.Cells(7, xCol + 3) = "Employment Dates"  'xCol+3 7
                exSheet.Range(exSheet.Cells(7, xCol + 3), exSheet.Cells(7, xCol + 4)).Merge
                exSheet.Range(exSheet.Cells(7, xCol + 3), exSheet.Cells(7, xCol + 4)).HorizontalAlignment = xlCenter
                
                exSheet.Cells(8, xCol) = lStr("Original Hire")      'xCol   4
                exSheet.Cells(8, xCol + 1) = lStr("Seniority")      'xCol+1 5
                exSheet.Cells(8, xCol + 2) = lStr("Last Hire")      'xCol+2
                exSheet.Cells(8, xCol + 3) = lStr("Union Date")     'xCol+3
                exSheet.Cells(8, xCol + 4) = lStr("First Day")      'xCol+4
                exSheet.Cells(8, xCol + 5) = lStr("Last Day")       'xCol+5
                exSheet.Cells(8, xCol + 6) = lStr("OMERS Date")     'xCol+6
                exSheet.Cells(8, xCol + 7) = lStr("User Defined")   'xCol+7
                
                xCol = xCol + 8
                xRptHeadRange = "K1"
            'End If
            
            'Show All Pension Dates
            'If chkPenDates Or chkShowAllPenDates Then
                exSheet.Cells(7, xCol + 4) = "Pension Dates"    'xCol+4 16
                exSheet.Range(exSheet.Cells(7, xCol + 4), exSheet.Cells(7, xCol + 5)).Merge
                exSheet.Range(exSheet.Cells(7, xCol + 4), exSheet.Cells(7, xCol + 5)).HorizontalAlignment = xlCenter
                
                exSheet.Cells(8, xCol) = lStr("Eligibility")                'xCol   12
                exSheet.Cells(8, xCol + 1) = lStr("Earliest Retirement")    'xCol+1 13
                exSheet.Cells(8, xCol + 2) = lStr("Normal Retirement")      'xCol+2
                exSheet.Cells(8, xCol + 3) = lStr("Latest Retirement")      'xCol+3
                exSheet.Cells(8, xCol + 4) = lStr("Pension Date 1")         'xCol+4
                exSheet.Cells(8, xCol + 5) = lStr("Pension Date 2")         'xCol+5
                exSheet.Cells(8, xCol + 6) = lStr("Pension Date 3")         'xCol+6
                exSheet.Cells(8, xCol + 7) = lStr("Pension Date 4")         'xCol+7
                exSheet.Cells(8, xCol + 8) = lStr("Pension Date 5")         'xCol+8
                exSheet.Cells(8, xCol + 9) = lStr("Pension Date 6")         'xCol+9
                
                xCol = xCol + 10
                
                If Len(xRptHeadRange) > 0 Then
                    xRptHeadRange = "U1"
                Else
                    xRptHeadRange = "M1"
                End If
            'End If
            
            'Show All Other Dates
            'If chkOthDates Or chkShowAllOthDates Then
                exSheet.Cells(7, xCol + 4) = "Other Dates"          'xCol+4 26
                exSheet.Range(exSheet.Cells(7, xCol + 4), exSheet.Cells(7, xCol + 5)).Merge
                exSheet.Range(exSheet.Cells(7, xCol + 4), exSheet.Cells(7, xCol + 5)).HorizontalAlignment = xlCenter
                
                exSheet.Cells(8, xCol) = lStr("Other Date 1")       'xCol   22
                exSheet.Cells(8, xCol + 1) = lStr("Other Date 2")   'xCol+1 23
                exSheet.Cells(8, xCol + 2) = lStr("Other Date 3")   'xCol+2
                exSheet.Cells(8, xCol + 3) = lStr("Other Date 4")   'xCol+3
                exSheet.Cells(8, xCol + 4) = lStr("Other Date 5")   'xCol+4
                exSheet.Cells(8, xCol + 5) = lStr("Other Date 6")   'xCol+5
                exSheet.Cells(8, xCol + 6) = lStr("Other Date 7")   'xCol+6
                exSheet.Cells(8, xCol + 7) = lStr("Other Date 8")   'xCol+7
                exSheet.Cells(8, xCol + 8) = lStr("Other Date 9")   'xCol+8
                exSheet.Cells(8, xCol + 9) = lStr("Other Date 10")  'xCol+9
            
                If Len(xRptHeadRange) > 0 Then
                    If xCol = 22 Then
                        xRptHeadRange = "AE1"
                    ElseIf xCol = 12 Then
                        xRptHeadRange = "U1"
                    Else
                        xRptHeadRange = "W1"
                    End If
                Else
                    xRptHeadRange = "M1"
                End If
            'End If
            
            If Not rsHREmp.EOF Then
                xRow = 9
                Do While Not rsHREmp.EOF
                    exSheet.Cells(xRow, 1) = rsHREmp("ED_EMPNBR")
                    exSheet.Cells(xRow, 2) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
                    exSheet.Cells(xRow, 3) = GetJobData(GetJHData(rsHREmp("ED_EMPNBR"), "JH_JOB", ""), "JB_DESCR")
        
                    xCol = 4
                    'Show All Employment Dates
                    'If chkEmpDates Or chkShowAllEmpDates Then
                        exSheet.Cells(xRow, xCol) = rsHREmp("ED_DOH")           '4
                        exSheet.Cells(xRow, xCol + 1) = rsHREmp("ED_SENDTE")    '5
                        exSheet.Cells(xRow, xCol + 2) = rsHREmp("ED_LTHIRE")
                        exSheet.Cells(xRow, xCol + 3) = rsHREmp("ED_UNION")
                        exSheet.Cells(xRow, xCol + 4) = rsHREmp("ED_FDAY")
                        exSheet.Cells(xRow, xCol + 5) = rsHREmp("ED_LDAY")
                        exSheet.Cells(xRow, xCol + 6) = rsHREmp("ED_OMERS")
                        exSheet.Cells(xRow, xCol + 7) = rsHREmp("ED_USRDAT1")
                        
                        xCol = xCol + 8
                    'End If
                    
                    'Show All Pension Dates
                    'If chkPenDates Or chkShowAllPenDates Then
                        exSheet.Cells(xRow, xCol) = rsHREmp("ED_ELIGIBLE")      '12
                        exSheet.Cells(xRow, xCol + 1) = rsHREmp("ED_EARLYR")    '13
                        exSheet.Cells(xRow, xCol + 2) = rsHREmp("ED_NORMALR")
                        exSheet.Cells(xRow, xCol + 3) = rsHREmp("ED_LATESTR")
                    'End If
                    
                    'More Pension Dates & Other Dates from HREMP_OTHER table
                    'If chkPenDates Or chkShowAllPenDates Or chkOthDates Or chkShowAllOthDates Then
                        'Get date from HREMP_OTHER table for Pension and Other Dates
                        rsHREmpOther.Open "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & rsHREmp("ED_EMPNBR"), gdbAdoIhr001, adOpenStatic, adLockOptimistic
                        If Not rsHREmpOther.EOF Then
                            'If chkPenDates Or chkShowAllPenDates Then
                                exSheet.Cells(xRow, xCol + 4) = rsHREmpOther("ER_PENSIONDATE1")   '16
                                exSheet.Cells(xRow, xCol + 5) = rsHREmpOther("ER_PENSIONDATE2")   '17
                                exSheet.Cells(xRow, xCol + 6) = rsHREmpOther("ER_PENSIONDATE3")
                                exSheet.Cells(xRow, xCol + 7) = rsHREmpOther("ER_PENSIONDATE4")
                                exSheet.Cells(xRow, xCol + 8) = rsHREmpOther("ER_PENSIONDATE5")
                                exSheet.Cells(xRow, xCol + 9) = rsHREmpOther("ER_PENSIONDATE6")
                                
                                xCol = xCol + 10
                            'End If
                            
                            'Show All Other Dates
                            'If chkOthDates Or chkShowAllOthDates Then
                                exSheet.Cells(xRow, xCol) = rsHREmpOther("ER_OTHERDATE1")       '22
                                exSheet.Cells(xRow, xCol + 1) = rsHREmpOther("ER_OTHERDATE2")   '23
                                exSheet.Cells(xRow, xCol + 2) = rsHREmpOther("ER_OTHERDATE3")
                                exSheet.Cells(xRow, xCol + 3) = rsHREmpOther("ER_OTHERDATE4")
                                exSheet.Cells(xRow, xCol + 4) = rsHREmpOther("ER_OTHERDATE5")
                                exSheet.Cells(xRow, xCol + 5) = rsHREmpOther("ER_OTHERDATE6")
                                exSheet.Cells(xRow, xCol + 6) = rsHREmpOther("ER_OTHERDATE7")
                                exSheet.Cells(xRow, xCol + 7) = rsHREmpOther("ER_OTHERDATE8")
                                exSheet.Cells(xRow, xCol + 8) = rsHREmpOther("ER_OTHERDATE9")
                                exSheet.Cells(xRow, xCol + 9) = rsHREmpOther("ER_OTHERDATE10")
                            'End If
                        End If
                        rsHREmpOther.Close
                        Set rsHREmpOther = Nothing
                    'End If
                    
                    xRow = xRow + 1
                
                    rsHREmp.MoveNext
                Loop
                
                If Len(xRptHeadRange) > 0 Then
                    exSheet.Range("A1:" & xRptHeadRange).Merge
                    exSheet.Range("A1:" & xRptHeadRange).HorizontalAlignment = xlCenter
                    exSheet.Range("A2:" & Replace(xRptHeadRange, "1", "2")).Merge
                    exSheet.Range("A2:" & Replace(xRptHeadRange, "1", "2")).HorizontalAlignment = xlCenter
                Else
                    exSheet.Range("A1:C1").Merge
                    exSheet.Range("A1:C1").HorizontalAlignment = xlCenter
                    exSheet.Range("A2:C2").Merge
                    exSheet.Range("A2:C2").HorizontalAlignment = xlCenter
                End If
                
                exSheet.PageSetup.Orientation = xlLandscape
                exSheet.PageSetup.Zoom = False
                exSheet.Range("A1").Select
                exBook.Save
                Set exSheet = Nothing
                
            End If
            rsHREmp.Close
            Set rsHREmp = Nothing
            
            rsGroup.MoveNext
        Loop
                
        Set exSheet = exBook.Worksheets(1)
        exApp.DisplayAlerts = False
        exSheet.Delete
        
        exApp.DisplayAlerts = True
        exBook.Worksheets(1).Activate
        exBook.Save
        
        Set exSheet = Nothing
        Set exBook = Nothing
        'exApp.Visible = True
        
        exApp.Quit
        
        Set exApp = Nothing

        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If

    Else
        MsgBox "There are no Records matching this criteria", vbInformation + vbOKOnly, "No Records"
        exBook.Close False
        exApp.Quit
        Set exApp = Nothing
    End If
    rsGroup.Close
    Set rsGroup = Nothing

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT

    Screen.MousePointer = vbDefault

    Exit Function
    
Err_XLWriter_Employee_DocumentType:

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT
    
    If Err = 1004 Then
        Resume Next
    End If
    
    If Err = 75 Then
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Function
    End If
    If Not exBook Is Nothing Then
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
    End If
    If Err = 70 Then
        Set exApp = Nothing
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Function
    End If
    If Err = 76 Then
        MsgBox Err.Description & " to save the Document Type Report." & vbCrLf & "Please check Company Preference under Setup Menu."
        Exit Function
    End If
    If Not exApp Is Nothing Then
        If exApp.Visible = False Then
            exApp.Quit
        End If
        Set exApp = Nothing
    End If
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "XLWriter_Employee_DocumentType", "", "Select")
End Function

Private Function GroupingInExcel(xGroup) As String
    Dim xExcelGrp As String
    
    Select Case xGroup
        Case lStr("Division"):      GroupingInExcel = "HR_DIVISION.Division_Name"
        Case lStr("Department"):    GroupingInExcel = "HRDEPT.DF_NAME"
        Case lStr("Location"):      GroupingInExcel = "HRTABL.TB_DESC"
        Case lStr("Section"):       GroupingInExcel = "tblSec.TB_DESC"
        Case "Employment Type":     GroupingInExcel = "ED_EMPTYPE"
        Case ("Home Line"):         GroupingInExcel = "LN_HOMES.TB_DESC"
        Case "Shift":               GroupingInExcel = "ED_SHIFT"
        Case lStr("Region"):        GroupingInExcel = "tblRegion.TB_DESC"
        Case "(none)":              GroupingInExcel = "(none)"
        Case "Employee Name":       GroupingInExcel = "ED_SURNAME, ED_FNAME"
        Case "Employee Number":     GroupingInExcel = "ED_EMPNBR"
        Case "Date of Hire":        GroupingInExcel = "ED_DOH"
        Case Else:                  GroupingInExcel = "ED_EMPNBR"
    End Select

    'if HREMP.ED_EMPTYPE="0" then "0 - Not Applicable" else
    'if HREMP.ED_EMPTYPE="1" then "1 - Full Time Salary" else
    'if HREMP.ED_EMPTYPE="2" then "2 - Part Time Salary" else
    'if HREMP.ED_EMPTYPE="3" then "3 - Full Time Hourly" else
    'if HREMP.ED_EMPTYPE="4" then "4 - Part Time Hourly" else
    'if HREMP.ED_EMPTYPE="5" then "5 - Casual/Other" else
    'if HREMP.ED_EMPTYPE="6" then "6 - Contract Salary"  else
    'if HREMP.ED_EMPTYPE="7" then "7 - Contract Hourly"  else
    'if HREMP.ED_EMPTYPE="8" then "8 - Salary Pensioners" else
    'if HREMP.ED_EMPTYPE="9" then "9 - Salary Elected officials"

End Function

Private Function getGroupField() As String

    Dim strTable As String
    Dim SQLQ As String
    
    Select Case comGroup(0).Text
        Case lStr("Division")
            strTable = ""
            strGroup = "HREMP.ED_DIV"
            SQLQ = "SELECT DISTINCT HREMP.ED_DIV as grpField, HR_DIVISION.Division_Name as TB_DESC FROM ((HREMP LEFT OUTER JOIN HR_DIVISION ON HREMP.ED_DIV = HR_DIVISION.DIV) "
            SQLQ = SQLQ & "INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
        Case lStr("Department")
            strTable = ""
            strGroup = "HREMP.ED_DEPTNO"
            SQLQ = "SELECT DISTINCT HREMP.ED_DEPTNO as grpField, HRDEPT.DF_NAME as TB_DESC FROM (HREMP INNER JOIN HRDEPT ON HREMP.ED_DEPTNO = HRDEPT.DF_NBR) INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
            SQLQ = SQLQ & "WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
        Case lStr("Location")
            strGroup = "HREMP.ED_LOC"
            strTable = "HREMP.ED_LOC_TABL"
        Case lStr("Union")
            strGroup = "HREMP.ED_ORG"
            strTable = "HREMP.ED_ORG_TABL"
        Case lStr("Administered By")
            strGroup = "HREMP.ED_ADMINBY"
            strTable = "HREMP.ED_ADMINBY_TABL"
        Case "Employee Name"
            strTable = ""
            strGroup = "HREMP.ED_SURNAME"
            If glbOracle Then
                SQLQ = "SELECT DISTINCT HREMP.ED_SURNAME AS grpfield, CONCAT(CONCAT(HREMP.ED_SURNAME, ', '), HREMP.ED_FNAME) AS TB_DESC FROM HREMP "
            Else
                SQLQ = "SELECT DISTINCT HREMP.ED_SURNAME AS grpfield, ED_SURNAME + ', ' + ED_FNAME AS TB_DESC FROM HREMP "
            End If
            SQLQ = SQLQ & "INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
        Case lStr("Section")
            strGroup = "HREMP.ED_SECTION"
            strTable = "HREMP.ED_SECTION_TABL"
        Case "Employment Type"
            strTable = ""
            strGroup = "HREMP.ED_EMPTYPE"
            SQLQ = "SELECT DISTINCT HREMP.ED_EMPTYPE as grpfield, ED_EMPTYPE as TB_DESC FROM HREMP "
            SQLQ = SQLQ & "INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
        Case ("Home Line")
            strGroup = "HREMP.ED_HOMELINE"
            strTable = "HREMP.ED_HOMELINE_TABL"
        Case "Shift"
            strTable = ""
            strGroup = "HR_JOB_HISTORY.JH_SHIFT"
            SQLQ = "SELECT DISTINCT HR_JOB_HISTORY.JH_SHIFT as grpField, HR_JOB_HISTORY.JH_SHIFT as TB_DESC FROM HREMP LEFT OUTER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
        Case lStr("Rept. Authority 1")
            strTable = ""
            strGroup = "HR_JOB_HISTORY.JH_REPTAU"
            If glbOracle Then
                SQLQ = "SELECT DISTINCT HR_JOB_HISTORY.JH_REPTAU AS grpField, CONCAT(CONCAT(tblSuper.ED_SURNAME, ', '), tblSuper.ED_FNAME) AS TB_DESC "
                SQLQ = SQLQ & "FROM (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HREMP tblSuper ON HR_JOB_HISTORY.JH_REPTAU = tblSuper.ED_EMPNBR "
                SQLQ = SQLQ & "WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "

            Else
                SQLQ = "SELECT DISTINCT HR_JOB_HISTORY.JH_REPTAU AS grpField, tblSuper.ED_SURNAME + ', ' + tblSuper.ED_FNAME AS TB_DESC "
                SQLQ = SQLQ & "FROM (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HREMP AS tblSuper ON HR_JOB_HISTORY.JH_REPTAU = tblSuper.ED_EMPNBR "
                SQLQ = SQLQ & "WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
            End If
        Case lStr("Region")
            strGroup = "HREMP.ED_REGION"
            strTable = "HREMP.ED_REGION_TABL"
        Case "(none)"
            strGroup = "HREMP.ED_COMPNO"
            strTable = ""
            SQLQ = "SELECT DISTINCT " & strGroup & " as grpField, 'None' as TB_DESC FROM HREMP "
            SQLQ = SQLQ & "INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
    End Select
    
    If strTable <> "" Then
        SQLQ = "SELECT DISTINCT " & strGroup & " as grpField, HRTABL.TB_DESC as TB_DESC FROM ((HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
        SQLQ = SQLQ & "LEFT OUTER JOIN HRTABL ON " & strTable & " = HRTABL.TB_NAME AND " & strGroup & " = HRTABL.TB_KEY) "
        SQLQ = SQLQ & "WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
    End If
    
    SQLQ = SQLQ & glbstrSelCri
    SQLQ = SQLQ & " ORDER BY " & strGroup
    
    If Not (glbSQL Or glbOracle) Then
        SQLQ = Replace(SQLQ, "INNER", "LEFT")
    End If
    
    getGroupField = SQLQ
End Function

Public Function AttachmentExists(xEmpnbr, xDocName, Optional xDocKey, Optional xJob, Optional xSDate) As Boolean
    'Check if the attachment exists for an employee in the respective table
    Dim rsDoc As New ADODB.Recordset
    Dim SQLQ As String
    
    AttachmentExists = False
    
    Select Case xDocName
        Case "Resume"
            SQLQ = "SELECT * FROM HRDOC_EMP WHERE RE_TYPE='" & UCase(xDocName) & "' AND RE_EMPNBR IN (SELECT ED_EMPNRB FROM HREMP WHERE 1=1 AND " & glbstrSelCri & ")"
            
            rsDoc.Open "SELECT * FROM HRDOC_EMP WHERE RE_TYPE='" & UCase(xDocName) & "' AND RE_EMPNBR=" & xEmpnbr, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsDoc.EOF Then AttachmentExists = True
            
        Case "OtherInfo"
            'If glbtermopen Then
            '    rsDoc.Open "SELECT * FROM Term_HRDOC_HREMP_OTHER WHERE ER_TYPE='" & UCase(xDocName) & "' AND ER_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            '    If Not rsDoc.EOF Then AttachmentExists = True
            'Else
                rsDoc.Open "SELECT * FROM HRDOC_HREMP_OTHER WHERE ER_TYPE='" & UCase(xDocName) & "' AND ER_EMPNBR=" & xEmpnbr, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
                If Not rsDoc.EOF Then AttachmentExists = True
            'End If
        
        Case "EmployeeFlag"
            rsDoc.Open "SELECT * FROM HRDOC_EMP_FLAGS WHERE EF_FLAG = " & glbEmpFlagNo & " AND EF_TYPE='" & UCase(xDocName) & "' AND EF_EMPNBR=" & xEmpnbr, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsDoc.EOF Then AttachmentExists = True
        
        Case "Offer"
            rsDoc.Open "SELECT * FROM HRDOC_JOB_HISTORY WHERE DJ_TYPE='" & UCase(xDocName) & "' AND DJ_EMPNBR=" & xEmpnbr & " AND DJ_JOB= '" & xJob & "' AND DJ_SDATE =" & Date_SQL(xSDate), gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsDoc.EOF Then AttachmentExists = True
        
        Case "Performance"
            rsDoc.Open "SELECT * FROM HRDOC_PERFORM_HISTORY WHERE DH_TYPE='" & UCase(xDocName) & "' AND DH_EMPNBR=" & xEmpnbr & " AND DH_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsDoc.EOF Then AttachmentExists = True
        
        Case "DollarEnt"
            rsDoc.Open "SELECT * FROM HRDOC_HRDOLENT WHERE DE_TYPE='" & UCase(xDocName) & "' AND DE_EMPNBR=" & xEmpnbr & " AND DE_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsDoc.EOF Then AttachmentExists = True
            
        Case "Attendance"
            'If glbtermopen Then
            '    rsDoc.Open "SELECT * FROM Term_HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(xDocName) & "' AND AD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND AD_REASON='" & glbAttReason & "' AND AD_DOA=" & Date_SQL(glbAttDOA) & " AND AD_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            '    If Not rsDoc.EOF Then AttachmentExists = True
            'Else
                rsDoc.Open "SELECT * FROM HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(xDocName) & "' AND AD_EMPNBR=" & xEmpnbr & " AND AD_DOCKEY =" & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
                If Not rsDoc.EOF Then AttachmentExists = True
            'End If
        
        Case "Associations"
            'If glbtermopen Then
            '    rsDoc.Open "SELECT * FROM Term_HRDOC_TRADE WHERE TD_TYPE='" & UCase(xDocName) & "' AND TD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND TD_CODE='" & glbAssocCode & "' AND TD_BEGINDT=" & Date_SQL(glbBeginDt)    '" AND TD_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            '    If Not rsDoc.EOF Then AttachmentExists = True
            'Else
                rsDoc.Open "SELECT * FROM HRDOC_TRADE WHERE TD_TYPE ='" & UCase(xDocName) & "' AND TD_EMPNBR=" & xEmpnbr & " AND TD_DOCKEY =" & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
                If Not rsDoc.EOF Then AttachmentExists = True
            'End If
        
        Case "EdSem"
            rsDoc.Open "SELECT * FROM HRDOC_EDSEM WHERE ES_TYPE='" & UCase(xDocName) & "' AND ES_EMPNBR=" & xEmpnbr & " AND ES_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsDoc.EOF Then AttachmentExists = True
        
        Case "FormalEdu"
            'If glbtermopen Then
            '    rsDoc.Open "SELECT * FROM Term_HRDOC_HREDU WHERE EU_TYPE='" & UCase(xDocName) & "' AND EU_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND EU_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            '    If Not rsDoc.EOF Then AttachmentExists = True
            'Else
                rsDoc.Open "SELECT * FROM HRDOC_HREDU WHERE EU_TYPE='" & UCase(xDocName) & "' AND EU_EMPNBR=" & xEmpnbr & " AND EU_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
                If Not rsDoc.EOF Then AttachmentExists = True
            'End If
        
        Case "Counsel"
            'If glbtermopen Then
            '    rsDoc.Open "SELECT * FROM Term_HRDOC_COUNSEL WHERE DC_TYPE='" & UCase(xDocName) & "' AND DC_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND DC_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            '    If Not rsDoc.EOF Then AttachmentExists = True
            'Else
                rsDoc.Open "SELECT * FROM HRDOC_COUNSEL WHERE DC_TYPE='" & UCase(xDocName) & "' AND DC_EMPNBR=" & xEmpnbr & " AND DC_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
                If Not rsDoc.EOF Then AttachmentExists = True
            'End If
        
        Case "Comments"
            'If glbtermopen Then
            '    rsDoc.Open "SELECT * FROM Term_HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(xDocName) & "' AND DO_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND DO_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            '    If Not rsDoc.EOF Then AttachmentExists = True
            'Else
                rsDoc.Open "SELECT * FROM HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(xDocName) & "' AND DO_EMPNBR=" & xEmpnbr & " AND DO_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
                If Not rsDoc.EOF Then AttachmentExists = True
            'End If
        
        Case "Jobdescription"
            rsDoc.Open "SELECT * FROM HRDOC_JOB WHERE DB_TYPE='" & UCase(xDocName) & "' AND DB_JOB= '" & xJob & "'", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsDoc.EOF Then AttachmentExists = True
            
    End Select
End Function

Private Sub Create_EmployeeDocumentList()
    Dim rsDocs As New ADODB.Recordset
    Dim SQLQ As String
    Dim xShowFlag As Boolean
    Dim xFlagNo As Integer
    Dim strSelCri As String
    
    'For each Type of Document selected, prepare a list of employee with that document
    'Include the filter for specific group of employees and specific Document Types.
    'If NOT EQUAL selected, prepare Employee List who do not have specific Document Types
    
'EQUAL TO Doc Type
'SELECT * FROM HRDOC_EMP WHERE RE_TYPE = 'RESUME' AND RE_DOCTYPE IN ('RESU') AND
'RE_EMPNBR IN (SELECT ED_EMPNBR FROM [SUPPORT1\SQLEXPRESS].IHRDEMO80.DBO.HREMP)
    
'NOT EQUAL TO Doc Type
'SELECT * FROM HREMP WHERE ED_EMPNBR NOT IN
'(SELECT RE_EMPNBR FROM [SUPPORT1\SQLEXPRESS].IHRDEMO80_DOC.DBO.HRDOC_EMP
'WHERE RE_TYPE = 'RESUME' AND RE_DOCTYPE IN ('RESU') )
'ORDER BY ED_EMPNBR
    
'Call in_SQL(glbAdoIHRDB_DOC)
'gdbAdoIhr001_DOC
    
    'Clear the existing records in the EmpDocs table for the existing User
    SQLQ = "DELETE FROM HR_EMPDOCS_WRK WHERE DC_WRKEMP='" & Replace(glbUserID, "'", "''") & "'"
    gdbAdoIhr001.Execute SQLQ
    
    'Change syntax to match TSQL from Crystal Reports syntax
    strSelCri = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
    'Ticket #28791
    strSelCri = Replace(Replace(strSelCri, "HREMP.", SQLDatabaseName & ".DBO.HREMP."), "HR_EMPDOCS_WRK.", SQLDatabaseName & ".DBO.HR_EMPDOCS_WRK.")
    
    'For each Type of Document selected, create the SQL Script, add to the temporary Employee Doc Table
    
    'RESUME (RE)
    If chkResume Then
        'HRDOC_EMP
        SQLQ = ""
        
        If chkNEDocType Then
            'NOT EQUAL TO Document Type
            'Ticket #28791
            'SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR NOT IN (SELECT RE_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_EMP"
            SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR NOT IN (SELECT RE_EMPNBR FROM " & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_EMP"
            
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = SQLQ & " WHERE RE_TYPE = 'RESUME' AND RE_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "'))"
            Else
                SQLQ = SQLQ & " WHERE RE_TYPE = 'RESUME') "
            End If
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " AND " & strSelCri
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            'EQUAL TO Document Type
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = "SELECT * FROM HRDOC_EMP WHERE RE_TYPE = 'RESUME' AND RE_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
            Else
                SQLQ = "SELECT * FROM HRDOC_EMP WHERE RE_TYPE = 'RESUME' "
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND RE_EMPNBR IN (SELECT ED_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HREMP "
            SQLQ = SQLQ & " AND RE_EMPNBR IN (SELECT ED_EMPNBR FROM " & SQLDatabaseName & ".DBO.HREMP "
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " WHERE " & strSelCri & ")"
            Else
                SQLQ = SQLQ & ")"
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        End If
        
        'Read each record and add in the temporary table.
        Do While Not rsDocs.EOF
            If chkNEDocType Then
                Call Add_EmpDocs_WorkTable(rsDocs("ED_EMPNBR"), "", "", "", "RESUME", "", "", "", "", "A", True)
            Else
                Call Add_EmpDocs_WorkTable(rsDocs("RE_EMPNBR"), rsDocs("RE_DOCTYPE"), rsDocs("RE_USRDESC"), rsDocs("RE_FILEEXT"), "RESUME", rsDocs("RE_LDATE"), rsDocs("RE_LTIME"), rsDocs("RE_LUSER"), "", "A", False)
            End If
            rsDocs.MoveNext
        Loop
        rsDocs.Close
        Set rsDocs = Nothing
    End If
    
    
    'OTHERINFO (ER)
    If chkOtherInfo Then
        'HRDOC_HREMP_OTHER
        SQLQ = ""
        
        If chkNEDocType Then
            'NOT EQUAL TO Document Type
            'Ticket #28791
            'SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR NOT IN (SELECT ER_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_HREMP_OTHER"
            SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR NOT IN (SELECT ER_EMPNBR FROM " & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_HREMP_OTHER"
            
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = SQLQ & " WHERE ER_TYPE = 'OTHERINFO' AND ER_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "'))"
            Else
                SQLQ = SQLQ & " WHERE ER_TYPE = 'OTHERINFO') "
            End If
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " AND " & strSelCri
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            'EQUAL TO Document Type
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = "SELECT * FROM HRDOC_HREMP_OTHER WHERE ER_TYPE = 'OTHERINFO' AND ER_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
            Else
                SQLQ = "SELECT * FROM HRDOC_HREMP_OTHER WHERE ER_TYPE = 'OTHERINFO' "
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND ER_EMPNBR IN (SELECT ED_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HREMP "
            SQLQ = SQLQ & " AND ER_EMPNBR IN (SELECT ED_EMPNBR FROM " & SQLDatabaseName & ".DBO.HREMP "
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " WHERE " & strSelCri & ")"
            Else
                SQLQ = SQLQ & ")"
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        End If
        
        'Read each record and add in the temporary table.
        Do While Not rsDocs.EOF
            If chkNEDocType Then
                Call Add_EmpDocs_WorkTable(rsDocs("ED_EMPNBR"), "", "", "", "OTHERINFO", "", "", "", "", "A", True)
            Else
                Call Add_EmpDocs_WorkTable(rsDocs("ER_EMPNBR"), rsDocs("ER_DOCTYPE"), rsDocs("ER_USRDESC"), rsDocs("ER_FILEEXT"), "OTHERINFO", rsDocs("ER_LDATE"), rsDocs("ER_LTIME"), rsDocs("ER_LUSER"), "", "A", False)
            End If
            rsDocs.MoveNext
        Loop
        rsDocs.Close
        Set rsDocs = Nothing
    End If
    
    'EMPLOYEEFLAG (EF)
    If chkEmpFlags Then
        'HRDOC_EMP_FLAGS
        SQLQ = ""
        xShowFlag = False
        
        For xFlagNo = 0 To lstEmpFlagsList.ListCount - 1 '20
            'Only show flags which are selected to show
            If lstEmpFlagsList.selected(xFlagNo) Then
                'This flag is selected to show
                xShowFlag = True
                
                'xFlagNo
                
                'Recordset for only flags selected to show
                If chkNEDocType Then
                    'NOT EQUAL TO Document Type
                    'Ticket #28791
                    'SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR NOT IN (SELECT EF_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_EMP_FLAGS"
                    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR NOT IN (SELECT EF_EMPNBR FROM " & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_EMP_FLAGS"
                    
                    If Len(clpCode(6).Text) > 0 Then
                        SQLQ = SQLQ & " WHERE EF_TYPE = 'EMPLOYEEFLAG' AND EF_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "')"
                    Else
                        SQLQ = SQLQ & " WHERE EF_TYPE = 'EMPLOYEEFLAG'"
                    End If
                    
                    SQLQ = SQLQ & " AND EF_FLAG = " & xFlagNo
                    SQLQ = SQLQ & ")"
                    If Len(strSelCri) > 0 Then
                        SQLQ = SQLQ & " AND " & strSelCri
                    End If
                    rsDocs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Else
                    'EQUAL TO Document Type
                    If Len(clpCode(6).Text) > 0 Then
                        SQLQ = "SELECT * FROM HRDOC_EMP_FLAGS WHERE EF_TYPE = 'EMPLOYEEFLAG' AND EF_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
                    Else
                        SQLQ = "SELECT * FROM HRDOC_EMP_FLAGS WHERE EF_TYPE = 'EMPLOYEEFLAG' "
                    End If
                    
                    SQLQ = SQLQ & " AND EF_FLAG = " & xFlagNo
                    
                    'Ticket #28791
                    'SQLQ = SQLQ & " AND EF_EMPNBR IN (SELECT ED_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HREMP "
                    SQLQ = SQLQ & " AND EF_EMPNBR IN (SELECT ED_EMPNBR FROM " & SQLDatabaseName & ".DBO.HREMP "
                    
                    If Len(strSelCri) > 0 Then
                        SQLQ = SQLQ & " WHERE " & strSelCri & ")"
                    Else
                        SQLQ = SQLQ & ")"
                    End If
                    rsDocs.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
                End If
                
                'Read each record and add in the temporary table.
                Do While Not rsDocs.EOF
                    If chkNEDocType Then
                        Call Add_EmpDocs_WorkTable(rsDocs("ED_EMPNBR"), "", "", "", "EMPLOYEEFLAG", "", "", "", "", "A", True)
                    Else
                        Call Add_EmpDocs_WorkTable(rsDocs("EF_EMPNBR"), rsDocs("EF_DOCTYPE"), rsDocs("EF_USRDESC"), rsDocs("EF_FILEEXT"), "EMPLOYEEFLAG", rsDocs("EF_LDATE"), rsDocs("EF_LTIME"), rsDocs("EF_LUSER"), rsDocs("EF_DOCKEY"), "A", False)
                    End If
                    rsDocs.MoveNext
                Loop
                rsDocs.Close
                Set rsDocs = Nothing
            End If
        Next
            
        'No flags selected to show
        If xShowFlag = False Then
        End If

    End If
    
    'OFFER (DJ)
    If chkJobOffer Then
        'HRDOC_JOB_HISTORY
        SQLQ = ""
        
        If chkNEDocType Then
            'NOT EQUAL TO Document Type
            SQLQ = "SELECT * FROM HREMP WHERE 1 = 1 "
            
            'Code selection
            'Ticket #29047
            If Len(clpCode(14).Text) > 0 Then
                SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_JREASON = '" & clpCode(14).Text & "')"
            End If
            
            'Ticket #28791
            'SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR NOT IN (SELECT DJ_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_JOB_HISTORY"
            'Ticket #29047
            'SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR NOT IN (SELECT DJ_EMPNBR FROM " & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_JOB_HISTORY"
            SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT DJ_EMPNBR FROM " & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_JOB_HISTORY"
            
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = SQLQ & " WHERE DJ_TYPE = 'OFFER' AND DJ_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "'))"
            Else
                SQLQ = SQLQ & " WHERE DJ_TYPE = 'OFFER') "
            End If
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " AND " & strSelCri
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            'EQUAL TO Document Type
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = "SELECT * FROM HRDOC_JOB_HISTORY WHERE DJ_TYPE = 'OFFER' AND DJ_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
            Else
                SQLQ = "SELECT * FROM HRDOC_JOB_HISTORY WHERE DJ_TYPE = 'OFFER' "
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND DJ_EMPNBR IN (SELECT ED_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HREMP "
            SQLQ = SQLQ & " AND DJ_EMPNBR IN (SELECT ED_EMPNBR FROM " & SQLDatabaseName & ".DBO.HREMP "
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " WHERE " & strSelCri & ")"
            Else
                SQLQ = SQLQ & ")"
            End If
            
            'Code selection
            If Len(clpCode(14).Text) > 0 Then
                'Ticket #29047
                SQLQ = SQLQ & " AND DJ_JOB IN (SELECT JH_JOB FROM " & SQLDatabaseName & ".DBO.HR_JOB_HISTORY WHERE JH_JREASON = '" & clpCode(14).Text & "')"
                SQLQ = SQLQ & " AND DJ_SDATE IN (SELECT JH_SDATE FROM " & SQLDatabaseName & ".DBO.HR_JOB_HISTORY WHERE JH_JREASON = '" & clpCode(14).Text & "')"
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        End If
        
        'Read each record and add in the temporary table.
        Do While Not rsDocs.EOF
            If chkNEDocType Then
                Call Add_EmpDocs_WorkTable(rsDocs("ED_EMPNBR"), "", "", "", "OFFER", "", "", "", "", "A", True)
            Else
                Call Add_EmpDocs_WorkTable(rsDocs("DJ_EMPNBR"), rsDocs("DJ_DOCTYPE"), rsDocs("DJ_USRDESC"), rsDocs("DJ_FILEEXT"), "OFFER", rsDocs("DJ_LDATE"), rsDocs("DJ_LTIME"), rsDocs("DJ_LUSER"), "", "A", False)
            End If
            rsDocs.MoveNext
        Loop
        rsDocs.Close
        Set rsDocs = Nothing
    End If
    
    'PERFORMANCE (DH)
    If chkPerfReview Then
        'HRDOC_PERFORM_HISTORY
        SQLQ = ""
        
        If chkNEDocType Then
            'NOT EQUAL TO Document Type
            SQLQ = "SELECT * FROM HREMP WHERE 1 = 1 "
            
            'Code selection
            'Ticket #29047
            If Len(clpCode(15).Text) > 0 Then
                SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT PH_EMPNBR FROM HR_PERFORM_HISTORY WHERE PH_PCODE = '" & clpCode(15).Text & "')"
            End If
            
            'Ticket #28791
            'SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR NOT IN (SELECT DH_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_PERFORM_HISTORY"
            'Ticket #29047
            'SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR NOT IN (SELECT DH_EMPNBR FROM " & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_PERFORM_HISTORY"
            SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT DH_EMPNBR FROM " & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_PERFORM_HISTORY"
            
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = SQLQ & " WHERE DH_TYPE = 'PERFORMANCE' AND DH_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "'))"
            Else
                SQLQ = SQLQ & " WHERE DH_TYPE = 'PERFORMANCE') "
            End If
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " AND " & strSelCri
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            'EQUAL TO Document Type
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = "SELECT * FROM HRDOC_PERFORM_HISTORY WHERE DH_TYPE = 'PERFORMANCE' AND DH_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
            Else
                SQLQ = "SELECT * FROM HRDOC_PERFORM_HISTORY WHERE DH_TYPE = 'PERFORMANCE' "
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND DH_EMPNBR IN (SELECT ED_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HREMP "
            SQLQ = SQLQ & " AND DH_EMPNBR IN (SELECT ED_EMPNBR FROM " & SQLDatabaseName & ".DBO.HREMP "
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " WHERE " & strSelCri & ")"
            Else
                SQLQ = SQLQ & ")"
            End If
            
            'Code selection
            If Len(clpCode(15).Text) > 0 Then
                'Ticket #29047
                SQLQ = SQLQ & " AND DH_DOCKEY IN (SELECT PH_DOCKEY FROM " & SQLDatabaseName & ".DBO.HR_PERFORM_HISTORY WHERE PH_PCODE = '" & clpCode(15).Text & "')"
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        End If
        
        'Read each record and add in the temporary table.
        Do While Not rsDocs.EOF
            If chkNEDocType Then
                Call Add_EmpDocs_WorkTable(rsDocs("ED_EMPNBR"), "", "", "", "PERFORMANCE", "", "", "", "", "A", True)
            Else
                Call Add_EmpDocs_WorkTable(rsDocs("DH_EMPNBR"), rsDocs("DH_DOCTYPE"), rsDocs("DH_USRDESC"), rsDocs("DH_FILEEXT"), "PERFORMANCE", rsDocs("DH_LDATE"), rsDocs("DH_LTIME"), rsDocs("DH_LUSER"), rsDocs("DH_DOCKEY"), "A", False)
            End If
            rsDocs.MoveNext
        Loop
        rsDocs.Close
        Set rsDocs = Nothing
    End If
    
    'DOLLARENT (DE)
    If chkDollarEnt Then
        'HRDOC_HRDOLENT
        SQLQ = ""
        
        If chkNEDocType Then
            'NOT EQUAL TO Document Type
            SQLQ = "SELECT * FROM HREMP WHERE 1 = 1 "
            
            'Code selection
            If Len(clpCode(7).Text) > 0 Then
                SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT DE_EMPNBR FROM HRDOLENT WHERE DE_TYPE = '" & clpCode(7).Text & "')"
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT DE_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_HRDOLENT"
            SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT DE_EMPNBR FROM " & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_HRDOLENT"
            
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = SQLQ & " WHERE DE_TYPE = 'DOLLARENT' AND DE_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "'))"
            Else
                SQLQ = SQLQ & " WHERE DE_TYPE = 'DOLLARENT') "
            End If
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " AND " & strSelCri
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            'EQUAL TO Document Type
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = "SELECT * FROM HRDOC_HRDOLENT WHERE DE_TYPE = 'DOLLARENT' AND DE_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
            Else
                SQLQ = "SELECT * FROM HRDOC_HRDOLENT WHERE DE_TYPE = 'DOLLARENT' "
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND DE_EMPNBR IN (SELECT ED_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HREMP "
            SQLQ = SQLQ & " AND DE_EMPNBR IN (SELECT ED_EMPNBR FROM " & SQLDatabaseName & ".DBO.HREMP "
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " WHERE " & strSelCri & ")"
            Else
                SQLQ = SQLQ & ")"
            End If
            
            'Code selection
            If Len(clpCode(7).Text) > 0 Then
                'Ticket #28791
                'SQLQ = SQLQ & " AND DE_EMPNBR IN (SELECT DE_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HRDOLENT WHERE DE_TYPE = '" & clpCode(7).Text & "')"
                SQLQ = SQLQ & " AND DE_EMPNBR IN (SELECT DE_EMPNBR FROM " & SQLDatabaseName & ".DBO.HRDOLENT WHERE DE_TYPE = '" & clpCode(7).Text & "')"
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        End If
        
        'Read each record and add in the temporary table.
        Do While Not rsDocs.EOF
            If chkNEDocType Then
                Call Add_EmpDocs_WorkTable(rsDocs("ED_EMPNBR"), "", "", "", "DOLLARENT", "", "", "", "", "A", True)
            Else
                Call Add_EmpDocs_WorkTable(rsDocs("DE_EMPNBR"), rsDocs("DE_DOCTYPE"), rsDocs("DE_USRDESC"), rsDocs("DE_FILEEXT"), "DOLLARENT", rsDocs("DE_LDATE"), rsDocs("DE_LTIME"), rsDocs("DE_LUSER"), rsDocs("DE_DOCKEY"), "A", False)
            End If
            rsDocs.MoveNext
        Loop
        rsDocs.Close
        Set rsDocs = Nothing
    End If
    
    'ATTENDANCE (AD)
    If chkAttendance Then
        'HRDOC_ATTENDANCE
        SQLQ = ""
    
        If chkNEDocType Then
            'NOT EQUAL TO Document Type
            SQLQ = "SELECT * FROM HREMP WHERE 1 = 1 "
            
            'Code selection
            If Len(clpCode(8).Text) > 0 Then
                SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE AD_REASON = '" & clpCode(8).Text & "')"
            End If
                       
            'Ticket #28791
            'SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT AD_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_ATTENDANCE"
            SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT AD_EMPNBR FROM " & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_ATTENDANCE"
            
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = SQLQ & " WHERE AD_TYPE = 'ATTENDANCE' AND AD_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "')"
            Else
                SQLQ = SQLQ & " WHERE AD_TYPE = 'ATTENDANCE' "
            End If
            
            'Ticket #28851 - Code selection
            If Len(clpCode(8).Text) > 0 Then
                SQLQ = SQLQ & " AND " & "AD_REASON = '" & clpCode(8).Text & "')"
            Else
                SQLQ = SQLQ & ")"
            End If
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " AND " & strSelCri
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            'EQUAL TO Document Type
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = "SELECT * FROM HRDOC_ATTENDANCE WHERE AD_TYPE = 'ATTENDANCE' AND AD_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
            Else
                SQLQ = "SELECT * FROM HRDOC_ATTENDANCE WHERE AD_TYPE = 'ATTENDANCE' "
            End If
            
            'Ticket #28851 - Code selection
            If Len(clpCode(8).Text) > 0 Then
                SQLQ = SQLQ & " AND " & "AD_REASON = '" & clpCode(8).Text & "'"
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HREMP "
            SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM " & SQLDatabaseName & ".DBO.HREMP "
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " WHERE " & strSelCri & ")"
            Else
                SQLQ = SQLQ & ")"
            End If
            
            'Code selection
            If Len(clpCode(8).Text) > 0 Then
                'Ticket #28791
                'SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT AD_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HR_ATTENDANCE WHERE AD_REASON = '" & clpCode(8).Text & "')"
                SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT AD_EMPNBR FROM " & SQLDatabaseName & ".DBO.HR_ATTENDANCE WHERE AD_REASON = '" & clpCode(8).Text & "')"
                
                'Ticket #28851 - All check for Dockey
                SQLQ = SQLQ & " AND AD_DOCKEY IN (SELECT AD_DOCKEY FROM " & SQLDatabaseName & ".DBO.HR_ATTENDANCE WHERE AD_REASON = '" & clpCode(8).Text & "')"
            Else
                'Ticket #28851 - All check for Dockey
                SQLQ = SQLQ & " AND AD_DOCKEY IN (SELECT AD_DOCKEY FROM " & SQLDatabaseName & ".DBO.HR_ATTENDANCE)"
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        End If
    
        'Read each record and add in the temporary table.
        Do While Not rsDocs.EOF
            If chkNEDocType Then
                Call Add_EmpDocs_WorkTable(rsDocs("ED_EMPNBR"), "", "", "", "ATTENDANCE", "", "", "", "", "A", True)
            Else
                Call Add_EmpDocs_WorkTable(rsDocs("AD_EMPNBR"), rsDocs("AD_DOCTYPE"), rsDocs("AD_USRDESC"), rsDocs("AD_FILEEXT"), "ATTENDANCE", rsDocs("AD_LDATE"), rsDocs("AD_LTIME"), rsDocs("AD_LUSER"), rsDocs("AD_DOCKEY"), "A", False)
            End If
            rsDocs.MoveNext
        Loop
        rsDocs.Close
        Set rsDocs = Nothing
    End If
    
    'ASSOCIATIONS (TD)
    If chkAssociation Then
        'HRDOC_TRADE
        SQLQ = ""
        
        If chkNEDocType Then
            'NOT EQUAL TO Document Type
            SQLQ = "SELECT * FROM HREMP WHERE 1 = 1 "
            
            'Code selection
            If Len(clpCode(9).Text) > 0 Then
                SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT TD_EMPNBR FROM HRTRADE WHERE TD_CODE = '" & clpCode(9).Text & "')"
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT TD_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_TRADE"
            SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT TD_EMPNBR FROM " & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_TRADE"
            
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = SQLQ & " WHERE TD_TYPE = 'ASSOCIATIONS' AND TD_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "')"
            Else
                SQLQ = SQLQ & " WHERE TD_TYPE = 'ASSOCIATIONS' "
            End If
            
            'Ticket #28851 - Code selection
            If Len(clpCode(9).Text) > 0 Then
                SQLQ = SQLQ & " AND " & "TD_CODE = '" & clpCode(9).Text & "')"
            Else
                SQLQ = SQLQ & ")"
            End If
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " AND " & strSelCri
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            'EQUAL TO Document Type
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = "SELECT * FROM HRDOC_TRADE WHERE TD_TYPE = 'ASSOCIATIONS' AND TD_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
            Else
                SQLQ = "SELECT * FROM HRDOC_TRADE WHERE TD_TYPE = 'ASSOCIATIONS' "
            End If
            
            'Ticket #28851 - Code selection
            If Len(clpCode(9).Text) > 0 Then
                SQLQ = SQLQ & " AND " & "TD_CODE = '" & clpCode(9).Text & "'"
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND TD_EMPNBR IN (SELECT ED_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HREMP "
            SQLQ = SQLQ & " AND TD_EMPNBR IN (SELECT ED_EMPNBR FROM " & SQLDatabaseName & ".DBO.HREMP "
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " WHERE " & strSelCri & ")"
            Else
                SQLQ = SQLQ & ")"
            End If
            
            'Code selection
            If Len(clpCode(9).Text) > 0 Then
                'Ticket #28791
                'SQLQ = SQLQ & " AND TD_EMPNBR IN (SELECT TD_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HRTRADE WHERE TD_CODE = '" & clpCode(9).Text & "')"
                SQLQ = SQLQ & " AND TD_EMPNBR IN (SELECT TD_EMPNBR FROM " & SQLDatabaseName & ".DBO.HRTRADE WHERE TD_CODE = '" & clpCode(9).Text & "')"
                
                'Ticket #28851 - All check for Dockey
                SQLQ = SQLQ & " AND TD_DOCKEY IN (SELECT TD_DOCKEY FROM " & SQLDatabaseName & ".DBO.HRTRADE WHERE TD_CODE = '" & clpCode(9).Text & "')"
            Else
                'Ticket #28851 - All check for Dockey
                SQLQ = SQLQ & " AND TD_DOCKEY IN (SELECT TD_DOCKEY FROM " & SQLDatabaseName & ".DBO.HRTRADE)"
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        End If
        
        'Read each record and add in the temporary table.
        Do While Not rsDocs.EOF
            If chkNEDocType Then
                Call Add_EmpDocs_WorkTable(rsDocs("ED_EMPNBR"), "", "", "", "ASSOCIATIONS", "", "", "", "", "A", True)
            Else
                Call Add_EmpDocs_WorkTable(rsDocs("TD_EMPNBR"), rsDocs("TD_DOCTYPE"), rsDocs("TD_USRDESC"), rsDocs("TD_FILEEXT"), "ASSOCIATIONS", rsDocs("TD_LDATE"), rsDocs("TD_LTIME"), rsDocs("TD_LUSER"), rsDocs("TD_DOCKEY"), "A", False)
            End If
            rsDocs.MoveNext
        Loop
        rsDocs.Close
        Set rsDocs = Nothing
    End If
    
    'EDSEM (ES)
    If chkContEdu Then
        'HRDOC_EDSEM
        SQLQ = ""
        
        If chkNEDocType Then
            'NOT EQUAL TO Document Type
            SQLQ = "SELECT * FROM HREMP WHERE 1 = 1 "
            
            'Code selection
            If Len(clpCode(10).Text) > 0 Then
                SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT ES_EMPNBR FROM HREDSEM WHERE ES_CRSCODE = '" & clpCode(10).Text & "')"
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT ES_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_EDSEM"
            SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT ES_EMPNBR FROM " & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_EDSEM"
            
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = SQLQ & " WHERE ES_TYPE = 'EDSEM' AND ES_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "'))"
            Else
                SQLQ = SQLQ & " WHERE ES_TYPE = 'EDSEM') "
            End If
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " AND " & strSelCri
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            'EQUAL TO Document Type
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = "SELECT * FROM HRDOC_EDSEM WHERE ES_TYPE = 'EDSEM' AND ES_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
            Else
                SQLQ = "SELECT * FROM HRDOC_EDSEM WHERE ES_TYPE = 'EDSEM' "
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND ES_EMPNBR IN (SELECT ED_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HREMP "
            SQLQ = SQLQ & " AND ES_EMPNBR IN (SELECT ED_EMPNBR FROM " & SQLDatabaseName & ".DBO.HREMP "
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " WHERE " & strSelCri & ")"
            Else
                SQLQ = SQLQ & ")"
            End If
            
            'Code selection
            If Len(clpCode(10).Text) > 0 Then
                'Ticket #28791
                'SQLQ = SQLQ & " AND ES_EMPNBR IN (SELECT ES_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HREDSEM WHERE ES_CRSCODE = '" & clpCode(10).Text & "')"
                SQLQ = SQLQ & " AND ES_EMPNBR IN (SELECT ES_EMPNBR FROM " & SQLDatabaseName & ".DBO.HREDSEM WHERE ES_CRSCODE = '" & clpCode(10).Text & "')"
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        End If
        
        'Read each record and add in the temporary table.
        Do While Not rsDocs.EOF
            If chkNEDocType Then
                Call Add_EmpDocs_WorkTable(rsDocs("ED_EMPNBR"), "", "", "", "EDSEM", "", "", "", "", "A", True)
            Else
                Call Add_EmpDocs_WorkTable(rsDocs("ES_EMPNBR"), rsDocs("ES_DOCTYPE"), rsDocs("ES_USRDESC"), rsDocs("ES_FILEEXT"), "EDSEM", rsDocs("ES_LDATE"), rsDocs("ES_LTIME"), rsDocs("ES_LUSER"), rsDocs("ES_DOCKEY"), "A", False)
            End If
            rsDocs.MoveNext
        Loop
        rsDocs.Close
        Set rsDocs = Nothing
    End If
    
    'FORMALEDU (EU)
    If chkFormalEdu Then
        'HRDOC_HREDU
        SQLQ = ""
        
        If chkNEDocType Then
            'NOT EQUAL TO Document Type
            SQLQ = "SELECT * FROM HREMP WHERE 1= 1 "
            
            'Code selection
            If Len(clpCode(11).Text) > 0 Then
                SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT EU_EMPNBR FROM HREDU WHERE EU_SCHOOL = '" & clpCode(11).Text & "')"
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT EU_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_HREDU"
            SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT EU_EMPNBR FROM " & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_HREDU"
            
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = SQLQ & " WHERE EU_TYPE = 'FORMALEDU' AND EU_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "'))"
            Else
                SQLQ = SQLQ & " WHERE EU_TYPE = 'FORMALEDU') "
            End If
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " AND " & strSelCri
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            'EQUAL TO Document Type
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = "SELECT * FROM HRDOC_HREDU WHERE EU_TYPE = 'FORMALEDU' AND EU_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
            Else
                SQLQ = "SELECT * FROM HRDOC_HREDU WHERE EU_TYPE = 'FORMALEDU' "
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND EU_EMPNBR IN (SELECT ED_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HREMP "
            SQLQ = SQLQ & " AND EU_EMPNBR IN (SELECT ED_EMPNBR FROM " & SQLDatabaseName & ".DBO.HREMP "
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " WHERE " & strSelCri & ")"
            Else
                SQLQ = SQLQ & ")"
            End If
            
            'Code selection
            If Len(clpCode(11).Text) > 0 Then
                'Ticket #28791
                'SQLQ = SQLQ & " AND EU_EMPNBR IN (SELECT EU_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HREDU WHERE EU_SCHOOL = '" & clpCode(11).Text & "')"
                SQLQ = SQLQ & " AND EU_EMPNBR IN (SELECT EU_EMPNBR FROM " & SQLDatabaseName & ".DBO.HREDU WHERE EU_SCHOOL = '" & clpCode(11).Text & "')"
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        End If
        
        'Read each record and add in the temporary table.
        Do While Not rsDocs.EOF
            If chkNEDocType Then
                Call Add_EmpDocs_WorkTable(rsDocs("ED_EMPNBR"), "", "", "", "FORMALEDU", "", "", "", "", "A", True)
            Else
                Call Add_EmpDocs_WorkTable(rsDocs("EU_EMPNBR"), rsDocs("EU_DOCTYPE"), rsDocs("EU_USRDESC"), rsDocs("EU_FILEEXT"), "FORMALEDU", rsDocs("EU_LDATE"), rsDocs("EU_LTIME"), rsDocs("EU_LUSER"), rsDocs("EU_DOCKEY"), "A", False)
            End If
            rsDocs.MoveNext
        Loop
        rsDocs.Close
        Set rsDocs = Nothing
    End If
    
    'COUNSEL (DC)
    If chkCounselling Then
        'HRDOC_COUNSEL
        SQLQ = ""
        
        If chkNEDocType Then
            'NOT EQUAL TO Document Type
            SQLQ = "SELECT * FROM HREMP WHERE 1 = 1 "
            
            'Code selection
            If Len(clpCode(12).Text) > 0 Then
                SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT CL_EMPNBR FROM HR_COUNSEL WHERE CL_TYPE = '" & clpCode(12).Text & "')"
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT DC_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_COUNSEL"
            SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT DC_EMPNBR FROM " & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_COUNSEL"
            
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = SQLQ & " WHERE DC_TYPE = 'COUNSEL' AND DC_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "'))"
            Else
                SQLQ = SQLQ & " WHERE DC_TYPE = 'COUNSEL') "
            End If
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " AND " & strSelCri
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            'EQUAL TO Document Type
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = "SELECT * FROM HRDOC_COUNSEL WHERE DC_TYPE = 'COUNSEL' AND DC_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
            Else
                SQLQ = "SELECT * FROM HRDOC_COUNSEL WHERE DC_TYPE = 'COUNSEL' "
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND DC_EMPNBR IN (SELECT ED_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HREMP "
            SQLQ = SQLQ & " AND DC_EMPNBR IN (SELECT ED_EMPNBR FROM " & SQLDatabaseName & ".DBO.HREMP "
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " WHERE " & strSelCri & ")"
            Else
                SQLQ = SQLQ & ")"
            End If
            
            'Code selection
            If Len(clpCode(12).Text) > 0 Then
                'Ticket #28791
                'SQLQ = SQLQ & " AND DC_EMPNBR IN (SELECT CL_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HR_COUNSEL WHERE CL_TYPE = '" & clpCode(12).Text & "')"
                SQLQ = SQLQ & " AND DC_EMPNBR IN (SELECT CL_EMPNBR FROM " & SQLDatabaseName & ".DBO.HR_COUNSEL WHERE CL_TYPE = '" & clpCode(12).Text & "')"
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        End If
        
        'Read each record and add in the temporary table.
        Do While Not rsDocs.EOF
            If chkNEDocType Then
                Call Add_EmpDocs_WorkTable(rsDocs("ED_EMPNBR"), "", "", "", "COUNSEL", "", "", "", "", "A", True)
            Else
                Call Add_EmpDocs_WorkTable(rsDocs("DC_EMPNBR"), rsDocs("DC_DOCTYPE"), rsDocs("DC_USRDESC"), rsDocs("DC_FILEEXT"), "COUNSEL", rsDocs("DC_LDATE"), rsDocs("DC_LTIME"), rsDocs("DC_LUSER"), rsDocs("DC_DOCKEY"), "A", False)
            End If
            rsDocs.MoveNext
        Loop
        rsDocs.Close
        Set rsDocs = Nothing
    End If
    
    'COMMENTS (D0)
    If chkComments Then
        'HRDOC_COMMENTS
        SQLQ = ""
        
        If chkNEDocType Then
            'NOT EQUAL TO Document Type
            SQLQ = "SELECT * FROM HREMP WHERE 1 = 1 "
            
            'Code selection
            If Len(clpCode(13).Text) > 0 Then
                SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT CO_EMPNBR FROM HR_COMMENTS WHERE CO_TYPE = '" & clpCode(13).Text & "')"
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT DO_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_COMMENTS"
            SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT DO_EMPNBR FROM " & SQLDatabaseName & "_DOC" & ".DBO.HRDOC_COMMENTS"
            
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = SQLQ & " WHERE DO_TYPE = 'COMMENTS' AND DO_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "'))"
            Else
                SQLQ = SQLQ & " WHERE DO_TYPE = 'COMMENTS') "
            End If
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " AND " & strSelCri
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            'EQUAL TO Document Type
            If Len(clpCode(6).Text) > 0 Then
                SQLQ = "SELECT * FROM HRDOC_COMMENTS WHERE DO_TYPE = 'COMMENTS' AND DO_DOCTYPE IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
            Else
                SQLQ = "SELECT * FROM HRDOC_COMMENTS WHERE DO_TYPE = 'COMMENTS' "
            End If
            
            'Ticket #28791
            'SQLQ = SQLQ & " AND DO_EMPNBR IN (SELECT ED_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HREMP "
            SQLQ = SQLQ & " AND DO_EMPNBR IN (SELECT ED_EMPNBR FROM " & SQLDatabaseName & ".DBO.HREMP "
            
            If Len(strSelCri) > 0 Then
                SQLQ = SQLQ & " WHERE " & strSelCri & ")"
            Else
                SQLQ = SQLQ & ")"
            End If
            
            'Code selection
            If Len(clpCode(13).Text) > 0 Then
                'Ticket #28791
                'SQLQ = SQLQ & " AND DO_EMPNBR IN (SELECT CO_EMPNBR FROM [" & SQLServerName & "]." & SQLDatabaseName & ".DBO.HR_COMMENTS WHERE CO_TYPE = '" & clpCode(13).Text & "')"
                SQLQ = SQLQ & " AND DO_EMPNBR IN (SELECT CO_EMPNBR FROM " & SQLDatabaseName & ".DBO.HR_COMMENTS WHERE CO_TYPE = '" & clpCode(13).Text & "')"
            End If
            rsDocs.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        End If
        
        'Read each record and add in the temporary table.
        Do While Not rsDocs.EOF
            If chkNEDocType Then
                Call Add_EmpDocs_WorkTable(rsDocs("ED_EMPNBR"), "", "", "", "COMMENTS", "", "", "", "", "A", True)
            Else
                Call Add_EmpDocs_WorkTable(rsDocs("DO_EMPNBR"), rsDocs("DO_DOCTYPE"), rsDocs("DO_USRDESC"), rsDocs("DO_FILEEXT"), "COMMENTS", rsDocs("DO_LDATE"), rsDocs("DO_LTIME"), rsDocs("DO_LUSER"), rsDocs("DO_DOCKEY"), "A", False)
            End If
            rsDocs.MoveNext
        Loop
        rsDocs.Close
        Set rsDocs = Nothing
    End If
    
End Sub

Private Sub Add_EmpDocs_WorkTable(xEmpnbr, xDocType, xDocDesc, xFileExt, xDocName, xLDate, xLTime, xLUser, xDocKey, xActTerm, xMissing As Boolean)
    Dim rsEmpDocs As New ADODB.Recordset
    Dim SQLQ As String
    
    'Add record to the Work table
    SQLQ = "SELECT * FROM HR_EMPDOCS_WRK WHERE 1 = 2"
    rsEmpDocs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsEmpDocs.EOF Then
        rsEmpDocs.AddNew
        rsEmpDocs("DC_COMPNO") = "001"
        rsEmpDocs("DC_EMPNBR") = xEmpnbr
        rsEmpDocs("DC_FNAME") = GetEmpData(xEmpnbr, "ED_SURNAME", "")
        rsEmpDocs("DC_SURNAME") = GetEmpData(xEmpnbr, "ED_FNAME", "")
        rsEmpDocs("DC_DOCTYPE") = xDocType
        rsEmpDocs("DC_USRDESC") = xDocDesc
        rsEmpDocs("DC_FILEEXT") = xFileExt
        rsEmpDocs("DC_TYPE") = xDocName
        If Len(xLDate) Then
            rsEmpDocs("DC_LDATE") = xLDate
        End If
        rsEmpDocs("DC_LTIME") = xLTime
        rsEmpDocs("DC_LUSER") = xLUser
        If Len(xDocKey) > 0 Then
            rsEmpDocs("DC_DOCKEY") = xDocKey
        End If
        rsEmpDocs("DC_ACTTERM") = xActTerm
        rsEmpDocs("DC_MISSING") = xMissing
        rsEmpDocs("DC_WRKEMP") = glbUserID
        
        rsEmpDocs.Update
    End If
    rsEmpDocs.Close
    Set rsEmpDocs = Nothing
    
End Sub
