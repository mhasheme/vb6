VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmAUDITAttend 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance Audit Master File Update"
   ClientHeight    =   9840
   ClientLeft      =   4380
   ClientTop       =   3915
   ClientWidth     =   11985
   DrawMode        =   1  'Blackness
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9840
   ScaleWidth      =   11985
   Tag             =   "Audit Master File Update"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcAttAudit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   120
      ScaleHeight     =   9255
      ScaleWidth      =   11295
      TabIndex        =   31
      Top             =   240
      Width           =   11295
      Begin INFOHR_Controls.CodeLookup clpDIV 
         Height          =   285
         Left            =   2070
         TabIndex        =   0
         Top             =   930
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         TABLName        =   "n/a"
         LookupType      =   1
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
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Tag             =   "First Level of grouping records"
         Top             =   8310
         Width           =   2325
      End
      Begin VB.Frame frmAT 
         Height          =   470
         Left            =   210
         TabIndex        =   32
         Top             =   390
         Width           =   5475
         Begin VB.OptionButton optAT 
            Caption         =   "Active Employee"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   1
            Top             =   150
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optAT 
            Caption         =   "Terminated Employee"
            Height          =   255
            Index           =   1
            Left            =   2730
            TabIndex        =   2
            Top             =   150
            Width           =   2655
         End
      End
      Begin VB.ComboBox cmbUpload 
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
         Left            =   2385
         TabIndex        =   22
         Tag             =   "Choose Upload flag."
         Text            =   "Combo1"
         Top             =   6705
         Width           =   975
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
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Tag             =   "Final sorting of records - no totals"
         Top             =   8670
         Width           =   2325
      End
      Begin VB.TextBox txtShift 
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
         Left            =   9135
         MaxLength       =   4
         TabIndex        =   13
         Tag             =   "00-Shift"
         Top             =   4455
         Width           =   450
      End
      Begin VB.ComboBox comCountryOfEmp 
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
         Left            =   8145
         TabIndex        =   19
         Tag             =   "00-Country of Employment"
         Top             =   5880
         Width           =   1440
      End
      Begin VB.ComboBox comCountry 
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
         Left            =   2385
         TabIndex        =   18
         Tag             =   "00-Country"
         Top             =   5880
         Width           =   1440
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Left            =   5670
         TabIndex        =   21
         Tag             =   "40-Date upto and including this date forward"
         Top             =   6270
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Left            =   2070
         TabIndex        =   20
         Tag             =   "40-Date from and including this date forward"
         Top             =   6270
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin Threed.SSCheck chkPage 
         Height          =   225
         Left            =   2400
         TabIndex        =   24
         Tag             =   "Page break after Employee changes"
         Top             =   7470
         Width           =   225
         _Version        =   65536
         _ExtentX        =   397
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Page Break"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
         Font3D          =   3
      End
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   285
         Left            =   2070
         TabIndex        =   9
         Tag             =   "10-Enter Employee Number"
         Top             =   3405
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         TextBoxWidth    =   7195
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.EmployeeLookup elpUser 
         Height          =   315
         Left            =   7935
         TabIndex        =   23
         Top             =   6705
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         ShowDescription =   0   'False
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.CodeLookup clpDiv1 
         Height          =   285
         Left            =   2070
         TabIndex        =   3
         Tag             =   "00-Specific Division Desired"
         Top             =   1275
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
         Left            =   2070
         TabIndex        =   14
         Top             =   4815
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
         Left            =   2070
         TabIndex        =   11
         Tag             =   "00-Enter Region Code"
         Top             =   4110
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDRG"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   2070
         TabIndex        =   5
         Tag             =   "00-Enter Location Code"
         Top             =   1980
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
         Index           =   10
         Left            =   2070
         TabIndex        =   15
         Tag             =   "ADRE-Attendance Reason"
         Top             =   5175
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "ADRE"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   2070
         TabIndex        =   12
         Tag             =   "EDAB-Administered By"
         Top             =   4470
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
      End
      Begin INFOHR_Controls.EmployeeLookup elpSUP 
         Height          =   285
         Index           =   1
         Left            =   2070
         TabIndex        =   10
         Tag             =   "00-Employee Number "
         Top             =   3750
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TextBoxWidth    =   7195
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   2070
         TabIndex        =   4
         Tag             =   "00-Specific Department Desired"
         Top             =   1620
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   2070
         TabIndex        =   8
         Tag             =   "EDPT-Category"
         Top             =   3045
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
         Index           =   0
         Left            =   2070
         TabIndex        =   7
         Top             =   2685
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
         Index           =   5
         Left            =   2070
         TabIndex        =   6
         Tag             =   "00-Enter Union Code"
         Top             =   2340
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDOR"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpProv 
         Height          =   285
         Left            =   2070
         TabIndex        =   16
         Tag             =   "31-Province of Residence - Code"
         Top             =   5520
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin INFOHR_Controls.CodeLookup clpProvEmp 
         Height          =   285
         Left            =   8400
         TabIndex        =   17
         Tag             =   "31-Province of Employment - Code"
         Top             =   5520
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
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
         Left            =   180
         TabIndex        =   59
         Top             =   1320
         Width           =   1755
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Facility"
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
         Index           =   0
         Left            =   180
         TabIndex        =   58
         Top             =   945
         Visible         =   0   'False
         Width           =   1035
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
         Left            =   120
         TabIndex        =   57
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblReprtGrping 
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
         TabIndex        =   56
         Top             =   8070
         Width           =   1695
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
         Left            =   180
         TabIndex        =   55
         Top             =   8370
         Width           =   885
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Number  "
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
         Left            =   180
         TabIndex        =   54
         Top             =   3450
         Width           =   1380
      End
      Begin VB.Label lblFromTo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Update Date Range"
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
         Left            =   180
         TabIndex        =   53
         Top             =   6315
         Width           =   1785
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Upload Flag"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   52
         Top             =   6765
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Page Break on Employee"
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
         Left            =   180
         TabIndex        =   51
         Top             =   7470
         Width           =   1800
      End
      Begin VB.Label lblTo 
         Appearance      =   0  'Flat
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
         Left            =   5160
         TabIndex        =   50
         Top             =   6315
         Width           =   240
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
         Left            =   180
         TabIndex        =   49
         Top             =   8730
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User"
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
         Left            =   7380
         TabIndex        =   48
         Top             =   6735
         Width           =   330
      End
      Begin VB.Label lblSection 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   180
         TabIndex        =   47
         Top             =   4860
         Width           =   1620
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
         Left            =   180
         TabIndex        =   46
         Top             =   4155
         Width           =   1710
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
         Left            =   180
         TabIndex        =   45
         Top             =   2025
         Width           =   1695
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "AttSupervisor"
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
         Left            =   180
         TabIndex        =   44
         Top             =   3795
         Width           =   945
      End
      Begin VB.Label lblShift 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Left            =   8460
         TabIndex        =   43
         Top             =   4500
         Width           =   315
      End
      Begin VB.Label lblAttCodes 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Attendance Codes"
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
         Left            =   180
         TabIndex        =   42
         Top             =   5220
         Width           =   1320
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
         Left            =   180
         TabIndex        =   41
         Top             =   4515
         Width           =   1125
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
         Left            =   180
         TabIndex        =   40
         Top             =   1665
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
         Left            =   180
         TabIndex        =   39
         Top             =   2385
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
         Left            =   180
         TabIndex        =   38
         Top             =   2730
         Width           =   450
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
         Left            =   180
         TabIndex        =   37
         Top             =   3090
         Width           =   630
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Prov. of Residence"
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
         Left            =   180
         TabIndex        =   36
         Top             =   5565
         Width           =   1365
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Prov. of Employment"
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
         Left            =   6780
         TabIndex        =   35
         Top             =   5565
         Width           =   1455
      End
      Begin VB.Label lblCountry 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Country of Employment"
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
         Left            =   6300
         TabIndex        =   34
         Top             =   5940
         Width           =   1620
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
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
         Left            =   180
         TabIndex        =   33
         Top             =   5940
         Width           =   540
      End
   End
   Begin VB.VScrollBar scrControl 
      Height          =   9015
      LargeChange     =   315
      Left            =   11640
      Max             =   100
      SmallChange     =   315
      TabIndex        =   30
      Top             =   120
      Width           =   340
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   7680
      Top             =   10560
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7080
      Top             =   10560
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
   Begin INFOHR_Controls.CodeLookup clpPP 
      DataField       =   "SH_PAYP"
      Height          =   285
      Left            =   2205
      TabIndex        =   27
      Tag             =   "00-Enter pay period code"
      Top             =   10785
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDPP"
   End
   Begin VB.Label lblPP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Period"
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
      Left            =   300
      TabIndex        =   29
      Top             =   10830
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Left            =   10560
      TabIndex        =   28
      Top             =   10800
      Visible         =   0   'False
      Width           =   420
   End
End
Attribute VB_Name = "frmAUDITAttend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DeletedRecs As Long

Private Function chkAudit()
Dim dd As Long
Dim X%

chkAudit = False

On Error GoTo chkEOTHERE_Err

If glbLinamar Then
    If Len(clpDIV) > 0 Then
        If clpDIV.Caption = "Unassigned" Then
            MsgBox "If Facility Entered - they must exist"
            clpDIV.SetFocus
            Exit Function
        End If
    End If
Else
    If Not clpDiv1.ListChecker Then
        Exit Function
    End If
End If

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known"
    'clpDept.SetFocus
    Exit Function
End If

For X% = 0 To 6
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

If Not elpEEID.ListChecker Then
    Exit Function
End If

If Not elpSUP(1).ListChecker Then Exit Function
If Not clpCode(10).ListChecker Then Exit Function

If Len(dlpFrom.Text) > 0 Then
    If Not IsDate(dlpFrom.Text) Then
        MsgBox "Invalid From date"
        dlpFrom.SetFocus
        Exit Function
    End If
End If
If Len(dlpTo.Text) > 0 Then
    If Not IsDate(dlpTo.Text) Then
        MsgBox "Invalid To date"
        dlpTo.SetFocus
        Exit Function
    End If
End If
If Len(dlpFrom.Text) > 0 And Len(dlpTo.Text) > 0 Then
    dd = DateDiff("d", CVDate(dlpFrom.Text), CVDate(dlpTo.Text))
    If dd < 0 Then
        MsgBox "From date must be earlier than To Date"
        dlpFrom.SetFocus
        Exit Function
    End If
End If

If clpProv.Caption = "Unassigned" Then
    MsgBox "Invalid Prov. of Residence"
    clpProv.SetFocus
    Exit Function
End If

If clpProvEmp.Caption = "Unassigned" Then
    MsgBox "Invalid Prov. of Employment"
    clpProvEmp.SetFocus
    Exit Function
End If

chkAudit = True
Exit Function

chkEOTHERE_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkAudit", "HRAUDIT_ATTEND", "Update")
Resume Next

End Function

Private Sub chkPage_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbUpload_GotFocus()
    Call SetPanHelp(ActiveControl)
    MDIMain.panHelp(2).Caption = "Req."
End Sub

Public Sub cmdClose_Click()
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub Cri_PP()
    Dim PPCri As String
    
    If Len(clpPP.Text) > 0 Then
      PPCri = "{HR_SALARY_HISTORY.SH_PAYP} in ['" & clpPP.Text & "'] "
      If glbOracle Then
        PPCri = PPCri & "AND {HR_SALARY_HISTORY.SH_CURRENT}<>0 "
      Else
        PPCri = PPCri & "AND {HR_SALARY_HISTORY.SH_CURRENT} "
      End If
      If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
      glbstrSelCri = glbstrSelCri & PPCri
    End If
End Sub

Private Sub Cri_AdminBy()
    Dim AdminByCri As String
    
    If Len(clpCode(1).Text) > 0 Then
      AdminByCri = "{HREMP.ED_ADMINBY} = ['" & clpCode(1).Text & "'] "
      If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
      glbstrSelCri = glbstrSelCri & AdminByCri
    End If
End Sub

'Public Sub cmdDelete_Click()
'Dim X As Integer
'Dim DgDef, Title As String, Msg As String, Response As Integer
'
'If glbLinamar Then
'    If Len(clpDiv) = 0 Then
'        MsgBox "Facility is a required field"
'        clpDiv.SetFocus
'        Exit Sub
'    End If
'End If
'
'Title = "Mass Audit File Delete"
'DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
'Msg = "Are You Sure You Want To Delete ALL records for this criteria?"
'Response = MsgBox(Msg, DgDef, Title)    ' Get user response.
'
'If Response = IDNO Then    ' Evaluate response
'    Exit Sub
'End If
'
'Screen.MousePointer = HOURGLASS
'
'X = modDelRecs()
'
'Screen.MousePointer = DEFAULT
'
'If DeletedRecs = 0 Then
'    MsgBox "No records found for given selection criteria."
'Else
'    MsgBox DeletedRecs & " records deleted successfully"
'End If
'
'Exit Sub
'
'Del_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRAUDIT_ATTEND", "Delete")
'Screen.MousePointer = DEFAULT
'
'If gintRollBack% = False Then
'    Resume Next
'Else
'    Unload Me
'End If
'
'End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdPrint_Click()
On Error GoTo PrntErr
Dim X As Integer

Screen.MousePointer = HOURGLASS
If chkAudit() Then
    If Not PrtForm("Audit Master Update Criteria", Me) Then
        Exit Sub
    End If
    ' cmdView.Enabled = False
    ' cmdPrint.Enabled = False
    ' cmdDelete.Enabled = False
     X = cri_SetAll()
     Me.vbxCrystal.Destination = 1
     MDIMain.Timer1.Enabled = False
     Me.vbxCrystal.Action = 1
     vbxCrystal.Reset
     MDIMain.Timer1.Enabled = True
    '  cmdView.Enabled = True
    '  cmdPrint.Enabled = True
    '  If gSec_Upd_Audit Then cmdDelete.Enabled = True
End If
Screen.MousePointer = DEFAULT

Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdView_Click()
Dim X As Integer

On Error GoTo ViewErr

Screen.MousePointer = HOURGLASS

If chkAudit() Then
    '  cmdView.Enabled = False
    '  cmdPrint.Enabled = False
    '  cmdDelete.Enabled = False
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    X = cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    '  cmdView.Enabled = True
    '  cmdPrint.Enabled = True
    '  If gSec_Upd_Audit Then cmdDelete.Enabled = True
End If

Screen.MousePointer = DEFAULT

Exit Sub

ViewErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Cri_Dept()
    Dim countr   As Integer  ' EEList_Snap is definded at form level
    Dim DeptCri As String
    If Len(clpDept.Text) > 0 Then DeptCri = " AND (ED_DEPTNO IN ['" & Replace(clpDept.Text, ",", "','") & "')] "
    glbstrSelCri = glbSeleDeptUn & DeptCri
End Sub

Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID.Text) > 0 Then
    EECri = "{HR_ATTENDANCE.AD_EMPNBR} in [" & getEmpnbr(elpEEID.Text) & "] "
    
    If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
    
    glbstrSelCri = glbstrSelCri & EECri
End If

End Sub

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY As Integer, dtMM As Integer, dtDD As Integer


If Len(dlpFrom.Text) = 0 And Len(dlpTo.Text) = 0 Then Exit Sub

TempCri = "({HR_ATTENDANCE.AD_LDATE} "
If Len(dlpFrom.Text) > 0 And Len(dlpTo.Text) > 0 Then
    dtYYY = Year(dlpFrom.Text)
    dtMM = month(dlpFrom.Text)
    dtDD = Day(dlpFrom.Text)
    TempCri = TempCri & " in Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ") "
    dtYYY = Year(dlpTo.Text)
    dtMM = month(dlpTo.Text)
    dtDD = Day(dlpTo.Text)
    TempCri = TempCri & " to Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ")) "
Else
    If Len(dlpFrom.Text) > 0 Then
        TempCri = TempCri & " >= "
        dtYYY = Year(dlpFrom.Text)
        dtMM = month(dlpFrom.Text)
        dtDD = Day(dlpFrom.Text)
        TempCri = TempCri & " Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ")) "
    End If
    If Len(dlpTo.Text) > 0 Then
        TempCri = TempCri & " <= "
        dtYYY = Year(dlpTo.Text)
        dtMM = month(dlpTo.Text)
        dtDD = Day(dlpTo.Text)
        TempCri = TempCri & " Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ")) "
    End If
End If
If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
glbstrSelCri = glbstrSelCri & TempCri

End Sub

Private Sub Cri_Sup()
Dim EECri As String

If Len(elpSUP(1).Text) > 0 Then
    'EECri = "{HREMP.ED_EMPNBR} IN (SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE AD_SUPER IN [" & getEmpnbr(elpSUP(1).Text) & "]) "
    EECri = "{HR_ATTENDANCE.AD_SUPER} IN [" & getEmpnbr(elpSUP(1).Text) & "] "
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

Private Sub Cri_Reason()
Dim ReasonCri As String
Dim countr   As Integer

If Len(clpCode(10).Text) > 0 Then
    ReasonCri = " {HR_ATTENDANCE.AD_REASON} IN ['" & getCodes(clpCode(10).Text) & "'] "
End If

If Len(ReasonCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = ReasonCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & ReasonCri
    End If
    glbiOneWhere = True
End If
End Sub

Private Sub Cri_Shift()
Dim EECri As String

If Len(txtShift.Text) > 0 Then
    EECri = "{HR_ATTENDANCE.AD_SHIFT} = '" & txtShift.Text & "' " 'laura dec 16, 1997
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

Private Function cri_SetAll()
On Error GoTo modSetCriteria_Err
Dim X As Integer
Dim xTitle As String

cri_SetAll = False

Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

'Ticket #27177
'Call glbCri_DeptUN("")
Call glbCri_DeptUN(clpDept.Text)

' call cri models set both glbiONeWhere and strSelCri
If glbLinamar Then
    Call Cri_Div
Else
    Call Cri_Div1
End If

'Ticket #27177
'Call Cri_Dept

For X% = 0 To 6
    Call Cri_Code(X%)
Next X%

'Call Cri_Loc
'Call Cri_Section 'Ticket #19437 11/12/2010 Frank
'Call Cri_Region
'Call Cri_AdminBy 'Ticket #18352 04/27/2010 Frank

Call Cri_EE
Call Cri_Sup

'Call Cri_PP
Call Cri_FTDates
Call Cri_Upload

Call Cri_Shift
Call Cri_Reason
Call Cri_ProvResidence
Call Cri_ProvEmployment
Call Cri_Country
Call Cri_CountryOfEmployment

Call Cri_Checks

Call Cri_Sorts

Call Cri_User

'Call setRptLabel(Me, 2)

If optAT(0) <> 0 Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzAttAudit2.rpt"
    xTitle = lStr("Attendance Audit Report")
Else
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzAttAudit3.rpt"
    xTitle = lStr("Terminated Attendance Audit Report")
End If
Me.vbxCrystal.Formulas(4) = "lblTitle='" & xTitle & "'"

'From Label Master
Me.vbxCrystal.Formulas(1) = "lblAttFromDate='" & lStr("From Date") & "'"
Me.vbxCrystal.Formulas(2) = "lblAttReason='" & lStr("Reason") & "'"
Me.vbxCrystal.Formulas(3) = "lblAttHours='" & lStr("Hours") & "'"

Me.vbxCrystal.SelectionFormula = glbstrSelCri

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    If optAT(0) <> 0 Then
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
    Else
        Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
    End If
    Me.vbxCrystal.DataFiles(1) = glbIHRDB
    Me.vbxCrystal.DataFiles(2) = glbIHRAUDIT
    Me.vbxCrystal.DataFiles(3) = glbIHRDB
End If

If chkPage Then
    Me.vbxCrystal.SectionFormat(0) = "GH1;T;F;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = "GF1;X;X;T;X;X;X;X"
Else
    Me.vbxCrystal.SectionFormat(0) = "GH1;T;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = "GF1;X;F;X;X;X;X;X"
End If

If glbSQL Then 'Ticket #18267, make this function for Samuel and all SQL customers
'If glbWFC Then 'Ticket #12867
    Me.vbxCrystal.Formulas(10) = "WFCNoEXECuser = " & glbNoEXEC & " "
    Me.vbxCrystal.Formulas(11) = "WFCNoNONEuser = " & glbNoNONE & " "
End If

' window title if appropriate
If optAT(0) <> 0 Then
    Me.vbxCrystal.WindowTitle = lStr("Attendance Audit Master File Report")
Else
    Me.vbxCrystal.WindowTitle = lStr("Terminated Attendance Audit Master File Report")
End If

cri_SetAll = True

Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Audit Master", "HRAUDIT_ATTEND Report", "Select")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Private Sub Cri_Upload()
Dim EECri As String

If cmbUpload.ListIndex > 0 Then
    If cmbUpload.ListIndex = 1 Then
        EECri = "{HR_ATTENDANCE.AD_UPLOAD} = 'Y' "
    End If
    
    If cmbUpload.ListIndex = 2 Then
        EECri = "{HR_ATTENDANCE.AD_UPLOAD} = 'N' "
    End If
    
    If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
    
    glbstrSelCri = glbstrSelCri & EECri
End If
End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMAUDITATTEND"
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim SQLQ As String

glbOnTop = "FRMAUDITATTEND"
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Me.Caption = lStr("Attendance Audit Master File Update")

If glbLinamar Then
    lblDiv.Visible = False
    clpDiv1.Visible = False
Else
    Call setCaption(lblDiv)
End If


'lblSection.Caption = lStr("Section")
'lblLocation.Caption = lStr("Location")
'lblRegion.Caption = lStr("Region")
Call setRptCaption(Me)
lblEENum(1).Caption = lStr("AttSupervisor")

Screen.MousePointer = HOURGLASS

If glbLinamar Then
    clpPP.Visible = False
    lblPP.Visible = False
End If
Data1.ConnectionString = glbAdoIHRAUDIT

cmbUpload.AddItem "All"
cmbUpload.AddItem "Yes"
cmbUpload.AddItem "No"
cmbUpload.ListIndex = 0
   
comGroup(0).Clear
comGroup(0).AddItem lStr("Division")
comGroup(0).AddItem lStr("Department")
comGroup(0).AddItem lStr("Location")
comGroup(0).AddItem "Employee Name"
comGroup(0).AddItem lStr("G/L")
comGroup(0).AddItem lStr("Section")
comGroup(0).AddItem lStr("Region")
comGroup(0).AddItem lStr("AttSupervisor")
comGroup(0).AddItem lStr("Administered By")
comGroup(0).AddItem "(none)"
comGroup(0).ListIndex = 0
    
comGroup(1).Clear
comGroup(1).AddItem "Attendance Date"
comGroup(1).AddItem "Date Changed"
comGroup(1).AddItem "Employee Number"
comGroup(1).AddItem "Employee Name"
'Ticket #22682 - Release 8.0: Add User to the final sort
comGroup(1).AddItem "User"
comGroup(1).AddItem "(none)"
comGroup(1).ListIndex = 0

If Not gSec_Upd_Audit Then     'May99 js
'    cmdDelete.Enabled = False   '
End If                          '
If glbLinamar Then
    lblTitle(0).Visible = True
    clpDIV.Visible = True
    frmAT.Visible = True
End If
elpUser.LookupType = 2

'If glbCompSerial = "S/N - 2382W" Then 'Ticket #18352 Samuel - add Admin By
'    lblAdmin.Caption = lStr("Administered By")
'    lblAdmin.Top = 650
'    lblAdmin.Visible = True
'    clpCode(1).Top = 650
'    clpCode(1).Visible = True
'    frmAT.Top = 120
'End If

Call INI_Controls(Me)

Call addCountryItems

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
    MDIMain.panHelp(0).Caption = " "
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(2).Caption = " "
    MDIMain.panHelp(3).Caption = " "
End Sub

Private Sub Form_Resize()
Dim c As Long

On Error GoTo Eh

If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    If Me.Height >= 10500 Then
        scrControl.Value = 0
        
        pcAttAudit.Top = 120
        
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        scrControl.Left = Me.ScaleWidth - scrControl.Width
        scrControl.Height = Me.Height - 1000
        
        scrControl.Max = 5000
    End If

'    'Horizontal Scroll
'    scrHScroll.Width = Me.Width - 200
'    If Me.Width >= 11190 Then '9700 Then
'        scrHScroll.Value = 0
'        scrHScroll.Visible = False
'    Else
'        scrHScroll.Visible = True
'        scrHScroll.Top = Me.Height - 700
'        scrHScroll.Width = Me.Width - 120
'    End If
    
End If

exH:
    Exit Sub
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form_Resize", "Attendance Audit Report", "Resize")
    Resume exH
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmAUDITAttend = Nothing  'carmen may 2000
End Sub

'Private Function modDelRecs()
''''On Error GoTo cmdDel_Err
'Dim SQLQ As String, SQLW As String, SQL1 As String, SQLQ1 As String
'Dim TmpDeletedRecs As Long, DeletedRecs1 As Long, TmpDeletedRecs1 As Long, TmpDeletedRecs2 As Long, DeletedEmp0Recs As Long, DeletedEmp0Recs2 As Long
'Dim SQLQ2, SQLQ_0 As String
'
'modDelRecs = False
'
'glbstrSelCri = ""
'Screen.MousePointer = HOURGLASS
'
'SQLQ = "Delete FROM HRAUDIT_COUNSEL WHERE 1=1 "
'
'' do selection for pay period if they entered one
'If Len(clpPP.Text) > 0 Then
'    SQLQ = SQLQ & "AND AU_EMPNBR IN (SELECT SH_EMPNBR FROM HR_SALARY_HISTORY "
'    If Not glbSQL Then
'        SQLQ = SQLQ & in_SQL(glbIHRDB)
'    End If
'    SQLQ = SQLQ & "WHERE SH_CURRENT<>0 AND SH_PAYP='" & clpPP.Text & "') "
'End If
'
'' pay period selection end
'If glbLinamar Then
'    ' do selection for only emps we have security for
'    SQLQ = SQLQ & "AND AU_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP "
'    SQLQ = SQLQ & in_SQL(glbIHRDB)
'    SQLW = "WHERE " & glbSeleDeptUn & ")"
'Else
'    SQLW = ""
'End If
'
'SQLQ1 = SQLQ
'
'If Len(elpEEID.Text) > 0 Then SQLW = SQLW & " AND AU_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
'If Len(dlpFrom.Text) > 0 Then SQLW = SQLW & " AND AU_LDATE >= " & Date_SQL(dlpFrom.Text)
'If Len(dlpTo.Text) > 0 Then SQLW = SQLW & " AND AU_LDATE <= " & Date_SQL(dlpTo.Text)
'If glbLinamar Then
'    If Len(clpDIV) > 0 Then SQLW = SQLW & " AND RIGHT(AU_EMPNBR,3)=" & clpDIV
'Else
'    If Len(clpDiv1.Text) > 0 Then SQLW = SQLW & " AND AU_DIVUPL IN ('" & getCodes(clpDiv1.Text) & "') "
'End If
'If Len(elpUser.Text) > 0 Then SQLW = SQLW & "AND AU_LUSER = '" & elpUser.Text & "' "
'If cmbUpload.ListIndex > 0 Then
'  If cmbUpload.ListIndex = 1 Then SQLW = SQLW + " AND AU_UPLOAD = 'Y' "
'  If cmbUpload.ListIndex = 2 Then SQLW = SQLW + " AND AU_UPLOAD = 'N' "
'End If
'
'glbstrSelCri = ""
'If glbSQL Or glbOracle Then
'    Call glbCri_DeptUN("")
'    glbstrSelCri = Trim(Replace(Replace(glbstrSelCri, "{", ""), "}", ""))
'    If LCase(Left(Trim(glbstrSelCri), 3)) = "and" Then
'        glbstrSelCri = Mid(glbstrSelCri, 4, Len(glbstrSelCri) - 3)
'    End If
'    glbstrSelCri = " AND (AU_EMPNBR in (SELECT ED_EMPNBR FROM HREMP WHERE " & glbstrSelCri & ") OR AU_EMPNBR in (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & Replace(glbstrSelCri, "HREMP.", "Term_HREMP.") & ")  )"
'
'    SQLW = SQLW & glbstrSelCri
'End If
'
'If glbLinamar Then
'    SQLW = SQLW & " AND AU_TYPE<>'R'"
'End If
'
'SQLQ = SQLQ & SQLW
'gdbAdoIhr001X.Execute SQLQ, DeletedRecs
'
''--------------------------------------------------------------------------------------------
''Delete Audit records with AU_DIVUPL = blank or null
'If Not glbLinamar Or Len(clpDiv1.Text) > 0 Then
'    SQL1 = ""
'    If Len(elpEEID.Text) > 0 Then SQL1 = SQL1 & " AND AU_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
'    If Len(dlpFrom.Text) > 0 Then SQL1 = SQL1 & " AND AU_LDATE >= " & Date_SQL(dlpFrom.Text)
'    If Len(dlpTo.Text) > 0 Then SQL1 = SQL1 & " AND AU_LDATE <= " & Date_SQL(dlpTo.Text)
'
'    'If Len(clpDiv1.Text) > 0 Then SQL1 = SQL1 & " AND AU_DIVUPL IN ('" & getCodes(clpDiv1.Text) & "') "
'    If Len(clpDiv1.Text) > 0 Then SQL1 = SQL1 & " AND ((AU_DIVUPL IS NULL OR AU_DIVUPL = '') AND AU_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_DIV IN ('" & getCodes(clpDiv1.Text) & "')))"
'
'    If Len(elpUser.Text) > 0 Then SQL1 = SQL1 & " AND AU_LUSER = '" & elpUser.Text & "' "
'    If cmbUpload.ListIndex > 0 Then
'      If cmbUpload.ListIndex = 1 Then SQL1 = SQL1 + " AND AU_UPLOAD = 'Y' "
'      If cmbUpload.ListIndex = 2 Then SQL1 = SQL1 + " AND AU_UPLOAD = 'N' "
'    End If
'    SQL1 = SQL1 & glbstrSelCri
'    SQLQ1 = SQLQ1 & SQL1
'    gdbAdoIhr001X.Execute SQLQ1, DeletedRecs1
'End If
''--------------------------------------------------------------------------------------------
'
'' dkostka - 08/20/2001 - Added code to remove records for terminated emps too
'SQLQ = "DELETE FROM HRAUDIT_COUNSEL WHERE 1=1 "
'If glbLinamar Then
'    SQLQ = SQLQ & "AND AU_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP "
'End If
'SQLQ = SQLQ & SQLW
'
'' do selection for pay period if they entered one
'If Len(clpPP.Text) > 0 Then
'    SQLQ = SQLQ & "AND AU_EMPNBR IN (SELECT SH_EMPNBR FROM Term_SALARY_HISTORY "
'    SQLQ = SQLQ & "WHERE SH_CURRENT<>0 AND SH_PAYP='" & clpPP.Text & "') "
'End If
'' pay period selection end
'
'SQLQ = SQLQ & glbstrSelCri
'SQLQ2 = SQLQ
'
''gdbAdoIhr001X.Execute SQLQ, TmpDeletedRecs
'gdbAdoIhr001X.Execute SQLQ2, TmpDeletedRecs
''DeletedRecs = DeletedRecs + TmpDeletedRecs
'
''--------------------------------------------------------------------------------------------
''Delete Audit records with AU_DIVUPL = blank or null - Terminated employees
'If Not glbLinamar Or Len(clpDiv1.Text) > 0 Then
'    SQLQ2 = "DELETE FROM HRAUDIT_COUNSEL WHERE 1=1 "
'    SQLQ2 = SQLQ2 & SQL1
'
'    ' do selection for pay period if they entered one
'    If Len(clpPP.Text) > 0 Then
'        SQLQ2 = SQLQ2 & "AND AU_EMPNBR IN (SELECT SH_EMPNBR FROM Term_SALARY_HISTORY "
'        SQLQ2 = SQLQ2 & "WHERE SH_CURRENT<>0 AND SH_PAYP='" & clpPP.Text & "') "
'    End If
'
'    SQLQ2 = SQLQ2 & glbstrSelCri
'
'    ' pay period selection end
'    gdbAdoIhr001X.Execute SQLQ2, TmpDeletedRecs1
'    'DeletedRecs = DeletedRecs + TmpDeletedRecs1 + DeletedRecs1
'End If
''--------------------------------------------------------------------------------------------
'
''Ticket #16768
'SQLQ_0 = "DELETE FROM HRAUDIT_COUNSEL WHERE AU_EMPNBR = 0"
'gdbAdoIhr001X.Execute SQLQ_0, DeletedEmp0Recs
'
'
'DeletedRecs = DeletedRecs + DeletedRecs1 + DeletedEmp0Recs2 + TmpDeletedRecs + TmpDeletedRecs1 + TmpDeletedRecs2
'
'
'modDelRecs = True
'
'Exit Function
'
'cmdDel_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modDelRecs", "HRAUDIT_COUNSEL", "Delete")
'
'Screen.MousePointer = DEFAULT
'
'If gintRollBack% = False Then
'    RollBack
'    Resume Next
'Else
'    Unload Me
'End If
'
'End Function

Private Sub Cri_Div()

Dim DivCri As String

If Len(clpDIV.Text) > 0 Then
    DivCri = "(RIGHT(TOTEXT({HREMP.ED_EMPNBR},0),3) = '" & clpDIV.Text & "')"
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

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    Select Case intIdx%
        Case 0: strCd$ = "HREMP.ED_LOC"
        Case 1: strCd$ = "HREMP.ED_REGION"
        Case 2: strCd$ = "HREMP.ED_SECTION"
        Case 3: strCd$ = "HREMP.ED_EMP"
        Case 4: strCd$ = "HREMP.ED_ADMINBY"
        Case 5: strCd$ = "HREMP.ED_ORG"
        Case 6: strCd$ = "HREMP.ED_PT"
    End Select
    
    'CodeCri = "(" & strCd$ & " = '" & clpCode(intIdx%).Text & "')"
    CodeCri = "({" & strCd$ & "} in ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDIV.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
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
'MDIMain.MainToolBar.ButtonS(10).Visible = True
'MDIMain.MainToolBar.ButtonS(10).Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Cri_Checks()
'Added by Bryan 6/Jul/05 for ticket#8857
    Dim TempCri As String
        
    If Not glbLinamar Then
        If Not clpDiv1.ListChecker Then
            Exit Sub
        End If
    End If

  If Len(glbstrSelCri) > 3 And Len(TempCri) >= 1 Then glbstrSelCri = glbstrSelCri & " AND "
  glbstrSelCri = glbstrSelCri & TempCri
    
End Sub

Private Sub Cri_Sorts()
'Added by Bryan on Sep 7, 2005 Ticket#9279
Dim grpField As String
Dim grpCond As String
Dim z%
    
grpField$ = getEGroup(comGroup(0).Text)
If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #21854 Franks 04/09/2012, they need code instead of desc
    If comGroup(0).Text = lStr("Administered By") Then
        grpField$ = "{HREMP.ED_ADMINBY}"
    End If
End If
If grpField$ = "(none)" Then grpField$ = "{HREMP.ED_COMPNO}"

If comGroup(0).Text = "Shift" Then grpField$ = "{HR_ATTENDANCE.AD_SHIFT}"
If comGroup(0).Text = lStr("AttSupervisor") Then
    grpField$ = "{@fldADSuper}"
End If
'SavGrp1 = grpField$
If Not (grpField$ = "{HREMP.ED_COMPNO}") Then 'If GrpIdx% < 5 Then
    If glbCompSerial = "S/N - 2327W" And comGroup(0).Text = "Employee Name" Then
        Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = 'Associate Name'"
    Else
        Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = '" & comGroup(0).Text & "'"
    End If
    'Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = " & grpField$
Else
    Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = ''"
    'Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = ''"
    Me.vbxCrystal.SectionFormat(z%) = "GF1;F;X;X;X;X;X;X"
    z% = z% + 1
End If
Me.vbxCrystal.GroupCondition(0) = "GROUP1;" & grpField$ & ";ANYCHANGE;A"
    
    
If grpField$ = "(none)" Then
    'If optAT(0) <> 0 Then  'Ticket #18668
        Select Case comGroup(1).ListIndex
            Case 0:
                grpField = "{HR_ATTENDANCE.AD_DOA}"
                grpCond = "GROUP" & CStr(1) & ";{HR_ATTENDANCE.AD_DOA};ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(0) = grpCond
            Case 1:
                grpField = "{HR_ATTENDANCE.AD_LDATE}"
                grpCond = "GROUP" & CStr(1) & ";{HR_ATTENDANCE.AD_LDATE};ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(0) = grpCond
                'grpCond = "GROUP" & CStr(2) & ";{@EFullName};ANYCHANGE;A"
                'Me.vbxCrystal.GroupCondition(1) = grpCond
                'Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = 'Date of Change:'"
                'Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {HR_ATTENDANCE.AD_LDATE}"
                'Me.vbxCrystal.Formulas(2) = "DESCGROUP3 = ''"
                'Me.vbxCrystal.Formulas(3) = "lblEMPNO = ''"
            Case 2:
                grpCond = "GROUP" & CStr(1) & ";{HR_ATTENDANCE.AD_EMPNBR};ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(0) = grpCond
'                grpCond = "GROUP" & CStr(2) & ";{HR_ATTENDANCE.AD_LDATE};ANYCHANGE;A"
'                Me.vbxCrystal.GroupCondition(1) = grpCond
                'Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = 'Employee:'"
                'Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {@EFullName}"
                'Me.vbxCrystal.Formulas(2) = "DESCGROUP3 = 'Number:'"
            Case 3:
                grpCond = "GROUP" & CStr(1) & ";{@EFullName};ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(0) = grpCond
'                grpCond = "GROUP" & CStr(2) & ";{HR_ATTENDANCE.AD_LDATE};ANYCHANGE;A"
'                Me.vbxCrystal.GroupCondition(1) = grpCond
                grpField = "{@EFullName}"
                'Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = 'Employee:'"
                'Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {@EFullName}"
                'Me.vbxCrystal.Formulas(2) = "DESCGROUP3 = 'Number:'"
            Case 4:
                grpCond = "GROUP" & CStr(1) & ";{HR_ATTENDANCE.AD_LUSER};ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(0) = grpCond
'                grpCond = "GROUP" & CStr(2) & ";{HR_ATTENDANCE.AD_LDATE};ANYCHANGE;A"
'                Me.vbxCrystal.GroupCondition(1) = grpCond
                'Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = 'Employee:'"
                'Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {@EFullName}"
                'Me.vbxCrystal.Formulas(2) = "DESCGROUP3 = 'Number:'"
            Case Else: grpField = "(none)"
        End Select
    'End If
Else
    Select Case comGroup(1).ListIndex
        Case 0:
            grpField = "{HR_ATTENDANCE.AD_DOA}"
            Me.vbxCrystal.SortFields(0) = "+" & grpField
            Me.vbxCrystal.SortFields(1) = "+{@EFullName}"
            'grpCond = "GROUP" & CStr(2) & ";{HR_ATTENDANCE.AD_DOA};ANYCHANGE;A"
            'Me.vbxCrystal.GroupCondition(1) = grpCond
        Case 1:
            grpField = "{HR_ATTENDANCE.AD_LDATE}"
            Me.vbxCrystal.SortFields(0) = "+" & grpField
            Me.vbxCrystal.SortFields(1) = "+{@EFullName}"
            'grpCond = "GROUP" & CStr(2) & ";{HR_ATTENDANCE.AD_LDATE};ANYCHANGE;A"
            'Me.vbxCrystal.GroupCondition(1) = grpCond
        Case 2:
            grpField = "{HR_ATTENDANCE.AD_EMPNBR}"
            Me.vbxCrystal.SortFields(0) = "+" & grpField
            Me.vbxCrystal.SortFields(1) = "+{@EFullName}"
            'grpCond = "GROUP" & CStr(2) & ";{HR_ATTENDANCE.AD_EMPNBR};ANYCHANGE;A"
            'Me.vbxCrystal.GroupCondition(1) = grpCond
        Case 3:
            'grpCond = "GROUP" & CStr(2) & ";{@EFullName};ANYCHANGE;A"
            'Me.vbxCrystal.GroupCondition(1) = grpCond
            grpField = "{@EFullName}"
            Me.vbxCrystal.SortFields(0) = "+" & grpField
        Case 4:
            'grpCond = "GROUP" & CStr(2) & ";{HR_ATTENDANCE.AD_LUSER};ANYCHANGE;A"
            'Me.vbxCrystal.GroupCondition(1) = grpCond
        Case Else: grpField = "(none)"
    End Select

End If

End Sub

Private Sub Cri_Country()
Dim CountryCri As String

If Len(comCountry.Text) > 0 Then
    CountryCri = "({HREMP.ED_COUNTRY} = '" & comCountry.Text & "')"
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

Private Sub Cri_CountryOfEmployment()
Dim CountryCri As String

If Len(comCountryOfEmp.Text) > 0 Then
    CountryCri = "({HREMP.ED_WORKCOUNTRY} = '" & comCountryOfEmp.Text & "')"
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

Private Sub Cri_ProvResidence()
Dim EECri As String, OneSet%, X%

If Len(clpProv.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_PROV} in ['" & Replace(clpProv.Text, ",", "','") & "']"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_ProvEmployment()
Dim EECri As String, OneSet%, X%

If Len(clpProvEmp.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_PROVEMP} in ['" & Replace(clpProvEmp.Text, ",", "','") & "']"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_User()
Dim EECri As String

If Len(elpUser.Text) > 0 Then
    EECri = "LowerCase({HR_ATTENDANCE.AD_LUSER}) ='" & LCase(elpUser.Text) & "' "
    If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
    glbstrSelCri = glbstrSelCri & EECri
End If

End Sub

'Private Sub Cri_Region() 'Ticket #22423
'Dim RegionCri As String
'Dim countr   As Integer
'
'If Len(clpCode(4).Text) > 0 Then
'      RegionCri = " {HREMP.ED_REGION} IN ['" & getCodes(clpCode(4).Text) & "'] "
'End If
'
'If Len(RegionCri) >= 1 Then
'    If Not glbiOneWhere Then
'        glbstrSelCri = RegionCri
'    Else
'        glbstrSelCri = glbstrSelCri & " AND " & RegionCri
'    End If
'    glbiOneWhere = True
'End If
'End Sub
'
'Private Sub Cri_Loc() 'Ticket #22423
'Dim LocCri As String
'Dim countr   As Integer
'
'If Len(clpCode(3).Text) > 0 Then
'      LocCri = " {HREMP.ED_LOC} IN ['" & getCodes(clpCode(3).Text) & "'] "
'End If
'
'If Len(LocCri) >= 1 Then
'    If Not glbiOneWhere Then
'        glbstrSelCri = LocCri
'    Else
'        glbstrSelCri = glbstrSelCri & " AND " & LocCri
'    End If
'    glbiOneWhere = True
'End If
'
'End Sub
'
'Private Sub Cri_Section() 'Ticket #19437
'Dim SectionCri As String
'Dim countr   As Integer  ' EEList_Snap is definded at form level
'
'If Len(clpCode(2).Text) > 0 Then
'      SectionCri = " {HREMP.ED_SECTION} IN ['" & getCodes(clpCode(2).Text) & "'] "
'End If
'
'If Len(SectionCri) >= 1 Then
'    If Not glbiOneWhere Then
'        glbstrSelCri = SectionCri
'    Else
'        glbstrSelCri = glbstrSelCri & " AND " & SectionCri
'    End If
'    glbiOneWhere = True
'End If
'
'End Sub

Private Sub Cri_Div1()

Dim DivCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level


If Len(clpDiv1.Text) > 0 Then
    'DivCri = "({HREMP.ED_DIV} = '" & clpDiv.Text & "')"
    If glbOracle Then
        DivCri = "({HREMP.ED_DIV} IN ['" & getCodes(clpDiv1.Text) & "'])"
    Else
        DivCri = "({HREMP.ED_DIV} IN ['" & getCodes(clpDiv1.Text) & "'])"
    End If
        
    'Ticket #12843
    'DivCri = "({HRAUDIT_COUNSEL.AU_DIVUPL} IN ('" & getCodes(clpDiv1.Text) & "'))"
    'Ticket #13540 Frank, come AU_DIVUPL values were null or blank, but still showup on the report
    'DivCri = "(Length({HRAUDIT_COUNSEL.AU_DIVUPL})>0  AND ({HRAUDIT_COUNSEL.AU_DIVUPL} IN ('" & getCodes(clpDiv1.Text) & "')))"
    'DivCri = "(Length({HRAUDIT_COUNSEL.AU_DIVUPL})>0  AND ({HRAUDIT_COUNSEL.AU_DIVUPL} IN ['" & getCodes(clpDiv1.Text) & "']))"
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

Private Sub optAT_Click(Index As Integer)
    'Ticket #15483
    If Index = 1 Then
        elpEEID.LookupType = TERM
    Else
        elpEEID.LookupType = 0  '0 = ACTIVE. I cannot put as ACTIVE because it's changing to "Active" and that does not switch the lookup to ACTIVE employees
    End If
End Sub

Private Sub addCountryItems()
Dim ctylist, X

ctylist = CountryList
X = 1
Do While X > 0
    X = InStr(ctylist, "&")
    If X > 0 Then
        comCountryOfEmp.AddItem Left(ctylist, X - 1)
        comCountry.AddItem Left(ctylist, X - 1)
        ctylist = Mid(ctylist, X + 1)
    Else
        comCountryOfEmp.AddItem ctylist
        comCountry.AddItem ctylist
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
    comCountry.AddItem comCountry
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

Private Sub scrControl_Change()
pcAttAudit.Top = 120 - scrControl.Value
End Sub
