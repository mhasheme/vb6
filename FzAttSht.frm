VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRAttSht 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Attendance Calendar"
   ClientHeight    =   9915
   ClientLeft      =   375
   ClientTop       =   915
   ClientWidth     =   11850
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
   ScaleHeight     =   9915
   ScaleWidth      =   11850
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   9315
      LargeChange     =   315
      Left            =   11520
      Max             =   100
      SmallChange     =   315
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame scrFrame 
      BorderStyle     =   0  'None
      Height          =   9375
      Left            =   0
      TabIndex        =   32
      Top             =   120
      Width           =   11415
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Tag             =   "First Level of grouping records"
         Top             =   8415
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Tag             =   "Second level of grouping records"
         Top             =   8730
         Width           =   2325
      End
      Begin VB.TextBox txtYear 
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
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   24
         Tag             =   "61- Enter Year"
         Top             =   6960
         Width           =   855
      End
      Begin VB.OptionButton optAnnual 
         Caption         =   "Annual"
         Height          =   195
         Left            =   300
         TabIndex        =   17
         Top             =   6450
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "Month"
         Height          =   195
         Left            =   2085
         TabIndex        =   18
         Top             =   6450
         Width           =   1275
      End
      Begin VB.ComboBox cmbMonth 
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
         ItemData        =   "FzAttSht.frx":0000
         Left            =   1680
         List            =   "FzAttSht.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   7290
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2265
         MaxLength       =   4
         TabIndex        =   11
         Tag             =   "00-Shift"
         Top             =   3900
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox txtYearTo 
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
         Left            =   5880
         MaxLength       =   4
         TabIndex        =   26
         Tag             =   "61- Enter To Year"
         Top             =   6960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbMonthTo 
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
         ItemData        =   "FzAttSht.frx":008E
         Left            =   5880
         List            =   "FzAttSht.frx":00B6
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   7290
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.OptionButton optRolling 
         Caption         =   "Rolling"
         Height          =   195
         Left            =   5670
         TabIndex        =   20
         Top             =   6480
         Width           =   1275
      End
      Begin VB.CheckBox chkAbsence 
         Caption         =   "Absent"
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
         Left            =   3000
         TabIndex        =   14
         Top             =   4800
         Width           =   1215
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "Weekly"
         Height          =   195
         Left            =   3870
         TabIndex        =   19
         Top             =   6480
         Width           =   1275
      End
      Begin VB.OptionButton optBiWeek 
         Caption         =   "Bi-Weekly"
         Height          =   195
         Left            =   7455
         TabIndex        =   21
         Top             =   6480
         Width           =   1275
      End
      Begin VB.CheckBox chkBlank 
         Caption         =   "Blank Report"
         Height          =   255
         Left            =   7455
         TabIndex        =   34
         Top             =   6975
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.OptionButton optScheVAC 
         Caption         =   "Scheduled Vacation"
         Height          =   195
         Left            =   9270
         TabIndex        =   22
         Top             =   6480
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Frame frRptType 
         Caption         =   "Type of Report"
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   180
         TabIndex        =   33
         Top             =   5400
         Width           =   3375
         Begin VB.OptionButton optCrystal 
            Caption         =   "Crystal Report"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optExcel 
            Caption         =   "Excel"
            Height          =   195
            Left            =   2280
            TabIndex        =   16
            Top             =   360
            Width           =   855
         End
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   10
         Left            =   1950
         TabIndex        =   12
         Tag             =   "ADRE-Attendance Reason"
         Top             =   4230
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "ADRE"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   1950
         TabIndex        =   10
         Tag             =   "EDSE-Section "
         Top             =   3570
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   1950
         TabIndex        =   9
         Tag             =   "EDAB-Administered By"
         Top             =   3240
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   1950
         TabIndex        =   8
         Tag             =   "EDRG-Region"
         Top             =   2910
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   1950
         TabIndex        =   5
         Tag             =   "EDPT-Category"
         Top             =   1890
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
         Index           =   3
         Left            =   1950
         TabIndex        =   4
         Top             =   1560
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
         Index           =   0
         Left            =   1950
         TabIndex        =   2
         Tag             =   "EDLC-Location"
         Top             =   900
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   1950
         TabIndex        =   1
         Tag             =   "00-Specific Department Desired"
         Top             =   570
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
         Left            =   1950
         TabIndex        =   0
         Tag             =   "00-Specific Division Desired"
         Top             =   240
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
      Begin Threed.SSCheck chkShowEmp 
         Height          =   255
         Left            =   270
         TabIndex        =   13
         Tag             =   "If X-Show All Employees"
         Top             =   4800
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   " Show All Employees"
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
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   285
         Left            =   1950
         TabIndex        =   6
         Tag             =   "10-Enter Employee Number"
         Top             =   2220
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   503
         ShowUnassigned  =   1
         TextBoxWidth    =   7075
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.DateLookup dlpStartdate 
         Height          =   285
         Left            =   1350
         TabIndex        =   23
         Tag             =   "40-Date from and including this date forward"
         Top             =   6960
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   5
         Left            =   1950
         TabIndex        =   3
         Tag             =   "00-Enter Union Code"
         Top             =   1230
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDOR"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin Threed.SSCheck chkShowEmpNo 
         Height          =   255
         Left            =   4920
         TabIndex        =   30
         Tag             =   "If X-Show Employee Numbers only"
         Top             =   8760
         Visible         =   0   'False
         Width           =   2475
         _Version        =   65536
         _ExtentX        =   4366
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   " Show Employee Number only"
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
      Begin INFOHR_Controls.EmployeeLookup elpSUP 
         Height          =   285
         Index           =   1
         Left            =   1950
         TabIndex        =   7
         Tag             =   "00-Employee Number "
         Top             =   2565
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   503
         ShowUnassigned  =   1
         TextBoxWidth    =   7075
         RefreshDescriptionWhen=   2
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
         Left            =   180
         TabIndex        =   55
         Top             =   285
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
         Left            =   180
         TabIndex        =   54
         Top             =   615
         Width           =   825
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
         Left            =   180
         TabIndex        =   53
         Top             =   2280
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
         TabIndex        =   52
         Top             =   0
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
         Left            =   180
         TabIndex        =   51
         Top             =   8445
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
         Left            =   180
         TabIndex        =   50
         Top             =   8760
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
         Left            =   0
         TabIndex        =   49
         Top             =   8085
         Width           =   1575
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
         TabIndex        =   48
         Top             =   2955
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
         Left            =   180
         TabIndex        =   47
         Top             =   3615
         Width           =   540
      End
      Begin VB.Label lblYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   300
         TabIndex        =   46
         Top             =   6990
         Width           =   975
      End
      Begin VB.Label lblMonth 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
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
         TabIndex        =   45
         Top             =   7350
         Visible         =   0   'False
         Width           =   930
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
         TabIndex        =   44
         Top             =   1905
         Width           =   630
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
         TabIndex        =   43
         Top             =   3285
         Width           =   1125
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
         TabIndex        =   42
         Top             =   945
         Width           =   615
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
         TabIndex        =   41
         Top             =   1605
         Width           =   450
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
         TabIndex        =   40
         Top             =   1275
         Width           =   420
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
         TabIndex        =   39
         Top             =   4275
         Width           =   1320
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
         Left            =   180
         TabIndex        =   38
         Top             =   3945
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblYearTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "To Year"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   37
         Top             =   6990
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblMonthTo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "To Month"
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
         Left            =   4560
         TabIndex        =   36
         Top             =   7350
         Visible         =   0   'False
         Width           =   690
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
         TabIndex        =   35
         Top             =   2610
         Width           =   945
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   11280
      Top             =   9405
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
Attribute VB_Name = "frmRAttSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fsnapEENames As Recordset
Dim DATE1, DATE2
Dim xVacEnt, xVacTaken
Dim xAllReason As String
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Sub cmdClose_Click()
Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr
 
If CriCheck() Then

    If Not PrtForm(Me.Caption, Me) Then Exit Sub
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    x% = Cri_SetAll()
    If optScheVAC Or optExcel Then  'Excel file show up
        Call set_PrintState(True)
        Exit Sub
    End If
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

'Private Sub cmdPrint_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Call set_PrintState(False)
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Screen.MousePointer = HOURGLASS
    x% = Cri_SetAll()
    If optScheVAC Or optExcel Then 'Excel file show up
        Call set_PrintState(True)
        Exit Sub
    End If
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

Private Sub clpCode_Change(Index As Integer)
    'Added by Bryan 10/08/05 Ticket#8924
    'commented by Bryan Sep 1, 2005
    'code for making month boxes show on Annual report
'    If Index = 10 And glbCompSerial = "S/N - 2362W" Then
'        If Len(clpCode(10).Text) > 0 Then
'            lblMonth.Visible = optRolling
'            cmbMonth.Visible = optRolling
'            lblMonthTo.Visible = optAnnual Or optRolling
'            cmbMonthTo.Visible = optAnnual Or optRolling
'            lblMonth.Caption = "From Month"
'            txtYear.Visible = Not (optWeek Or optBiWeek)
'            If cmbMonth.ListIndex = -1 Then cmbMonth.ListIndex = 0
'            If cmbMonthTo.ListIndex = -1 Then cmbMonthTo.ListIndex = 0
'        End If
'    End If
End Sub

Private Sub comGroup_Click(Index As Integer)
    If Index = 1 Then
        If comGroup(1).ListIndex = 1 And comGroup(1).Text = "Employee Number" Then
            chkShowEmpNo.Visible = True
            If glbCompSerial = "S/N - 2335W" And InStr(1, glbSeleDiv, "HSV") = 0 Then   'Mitchell Plastics (Ultra Manufacturing)
                chkShowEmpNo.Value = True
            Else
                chkShowEmpNo.Value = False
            End If
        Else
            chkShowEmpNo.Visible = False
        End If
    End If
End Sub

'Private Sub cmdView_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "Employee Name"
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem "Rept. Authority 1"
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Employee Number"
    comGroup(1).ListIndex = 0

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
    CodeCri = "(" & strCd$ & " in ('" & Replace(clpCode(intIdx%).Text, ",", "','") & "'))"
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "((" & strCd$ & " = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or (" & strCd$ & " = 'ALL" & clpCode(intIdx%).Text & "') )"
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
    DivCri = "(HREMP.ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "'))"
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
    EECri = "HREMP.ED_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
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

Private Sub Cri_Sup()
Dim EECri As String

If Len(elpSUP(1).Text) > 0 Then
    EECri = "HREMP.ED_EMPNBR in (SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE AD_SUPER IN (" & getEmpnbr(elpSUP(1).Text) & ")) "
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

Private Function Cri_SetAll()
Dim x%, strRName$

Cri_SetAll = False

On Error GoTo modSetCriteria_Err

Screen.MousePointer = HOURGLASS

glbiOneWhere = True
glbstrSelCri = ""

Call Cri_Dept
Call Cri_Div
For x% = 0 To 6
    Call Cri_Code(x%)
Next x%
Call Cri_EE
Call Cri_Sup

Call AttWrk

'Added and commented out by Bryan Aug 25, Sep 1
'If optAnnual And Len(clpCode(10).Text) > 0 And glbCompSerial = "S/N - 2362W" Then
'    Call AttWrk2
'End If

If optMonth And Not optExcel Then
    If glbLinamar Then
        strRName$ = glbIHRREPORTS & "rlattcam.rpt"
    Else
        strRName$ = glbIHRREPORTS & "rzattcam.rpt"
    End If
ElseIf optWeek Then
    strRName$ = glbIHRREPORTS & "rzattcaw.rpt"
ElseIf optBiWeek Then
    strRName$ = glbIHRREPORTS & "rzattbw.rpt"
ElseIf optScheVAC Then
    Call WriteTo_XLS_VacSchedule
    Exit Function
ElseIf optExcel Then
    Call WriteTo_XLS_AttendanceCalendar
    Exit Function
Else
    If glbLinamar Then
        strRName$ = glbIHRREPORTS & "rlattcal.rpt"
    'Added by Bryan Aug 31 Ticket #9231
    'what they asked for is not what they wanted
    'commented out by Bryan on Sep 1, 2005
'    ElseIf glbCompSerial = "S/N - 2362W" And Len(clpCode(10).Text) > 0 And optRolling Then
'        strRName$ = glbIHRREPORTS & "rzattcal1.rpt"
    Else
        strRName$ = glbIHRREPORTS & "rzattcal.rpt"
        
        If glbCompSerial = "S/N - 2226W" And comGroup(0).ListIndex = 0 And optAnnual Then
            strRName$ = glbIHRREPORTS & "rzattcl1.rpt"
        End If
        
        If glbCompSerial = "S/N - 2226W" And optRolling Then
            'Excel Spreadsheet
            Dim dtFrom As Date, dtTo As Date
            dtFrom = cmbMonth.Text & " 1, " & txtYear.Text
            dtTo = cmbMonthTo.Text & " 1, " & txtYearTo.Text
            If Abs(DateDiff("m", dtTo, dtFrom)) > 6 Then
                MsgBox "Please choose a smaller date range, Excel can only handle 256 columns", vbInformation + vbOKOnly, "Too many days"
            Else
                Call Export_XLSWriter_BrantCounty
            End If
        End If
        
    End If
End If

Me.vbxCrystal.ReportFileName = strRName$

x% = Cri_Sorts()   ' returns number of sections formated

Me.vbxCrystal.SelectionFormula = "{HR_ATTCAL.AD_WRKEMP}='" & glbUserID & "'"

Me.vbxCrystal.WindowTitle = Me.Caption

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
    
    'Franks - 04/29/03 ticket# 4071 Begin - didn't transfer Connect to sub report
    Me.vbxCrystal.SubreportToChange = "Reason_Desc"
    Me.vbxCrystal.Connect = RptODBC_SQL
    
    'Added by Bryan 10/08/05 ticket#8924 - added a second subreport
    If strRName$ = "rzattcal1.rpt" Then
        Me.vbxCrystal.SubreportToChange = ""
        Me.vbxCrystal.SubreportToChange = "rzattcal2.rpt"
        Me.vbxCrystal.SelectionFormula = "{HR_ATTCAL.AD_WRKEMP}='" & glbUserID & "'"
        Me.vbxCrystal.Connect = RptODBC_SQL
    End If
    
    Me.vbxCrystal.SubreportToChange = ""
    'Franks - 04/29/03 ticket# 4071 End
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    Me.vbxCrystal.DataFiles(0) = glbIHRDBW
    For x% = 1 To 6
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next x%
    Me.vbxCrystal.SubreportToChange = "Reason_Desc"
    Me.vbxCrystal.Connect = "PWD=petman;"
    Me.vbxCrystal.DataFiles(0) = glbIHRDB
    
    'Added by Bryan 10/08/05 ticket#8924 - added a second subreport
    If strRName$ = "rzattcal1.rpt" Then
        Me.vbxCrystal.SubreportToChange = ""
        Me.vbxCrystal.SubreportToChange = "rzattcal2.rpt"
        Me.vbxCrystal.SelectionFormula = "{HR_ATTCAL.AD_WRKEMP}='" & glbUserID & "'"
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDBW
        For x% = 1 To 6
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next x%
    End If
    
    Me.vbxCrystal.SubreportToChange = ""
End If

Cri_SetAll = True

Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Cri_SetAll", "Attendance Calendar", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

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

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%
Dim strSMonth$
'for labels - sort by name always
'imbeded in report

Cri_Sorts = 0

'Added by Bryan 30/May/05 Ticket#10902
If comGroup(0).Text = "Rept. Authority 1" Then
    grpField$ = "{@fldADSuper}"
Else
    grpField$ = getEGroup(comGroup(0).Text)
    If grpField$ = "(none)" Then grpField$ = "{HRPARCO.PC_CO}"
    If glbLinamar Then
        If grpField$ = lStr("Region") Then grpField$ = "{@productline}"
    End If
End If
If glbCompSerial = "S/N - 2327W" And comGroup(0).Text = "Employee Name" Then
    dscGroup$ = "Associate Name"
Else
    dscGroup$ = comGroup(0).Text
End If

dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(0) = dscGroup$
      
grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(0) = grpCond$

'Hemu - 02/17/2004 Begin
'Custom code for Brant County Health Unit
If glbCompSerial = "S/N - 2226W" And comGroup(0).ListIndex = 0 Then 'InStr(1, grpCond$, "Division_Name") > 0
    grpCond$ = "GROUP" & CStr(2) & ";" & "{HRDEPT.DF_NAME}" & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(1) = grpCond$
Else
'Hemu - 02/17/2004 End
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{HR_ATTCAL.AD_EMPNBR}"
    End Select
    grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(1) = grpCond$

    If Not glbLinamar Then
        If GrpIdx% = 1 And comGroup(1).Text = "Employee Number" Then
            If chkShowEmpNo Then
                Me.vbxCrystal.Formulas(3) = "ShowEmpNo = True"
            End If
        End If
    End If
End If

Call setRptLabel(Me, 0)

'Cri_Sorts = z% ' next section number to format
If DATE1 <> "" And DATE2 <> "" Then
    strSFormat$ = "As of " & Format(DATE1, "MMM dd, YYYY") & " through " & Format(DATE2, "MMM dd, YYYY")
    Me.vbxCrystal.Formulas(1) = "Daterange = '" & strSFormat$ & "'"
    'added by Bryan Ticket #9231
'    If glbCompSerial = "S/N - 2362W" And Len(clpCode(10).Text) > 0 And optAnnual Then
'        strSFormat$ = "For Year " & "Jan 1, " & txtYear & " through " & "Dec 31, " & txtYear & " All Attendance Codes"
'        Me.vbxCrystal.SubreportToChange = "RZATTCAL2.rpt"
'        Me.vbxCrystal.Formulas(2) = "Daterange2 = '" & strSFormat$ & "'"
'        Me.vbxCrystal.SubreportToChange = ""
'    End If
Else
    strSFormat$ = "No date entered"
    Me.vbxCrystal.Formulas(1) = "Daterange = '" & strSFormat$ & "'"
End If
If glbWFC Then
    Me.vbxCrystal.Formulas(2) = "HideJob = False"
End If

End Function

Private Function CriCheck()
Dim x%

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

For x% = 0 To 6
    If Not clpCode(x).ListChecker Then Exit Function
Next x%

If glbWFC Then   'Ticket #13256
    If Len(Trim(clpCode(2).Text)) = 0 Then
        MsgBox lStr("Section is required")
        clpCode(2).SetFocus
        Exit Function
    End If
End If

clpCode(10).Text = Replace(clpCode(10).Text, " ", "")
If optWeek Or optBiWeek Then
    If dlpStartDate = "" Then
        MsgBox "You have to enter the Start Date!"
        dlpStartDate.SetFocus
        Exit Function
    Else
        If Not IsDate(dlpStartDate) Then
            MsgBox "Start Date is not a valid date"
            dlpStartDate.SetFocus
            Exit Function
        End If
    End If
    If glbCompSerial = "S/N - 2241W" And optBiWeek Then 'granite club
        If WeekdayName(Weekday(dlpStartDate)) <> "Sunday" Then
            MsgBox "Start Date must be Sunday"
            dlpStartDate.SetFocus
            Exit Function
        End If
    End If
Else
    If txtYear = "" Then
        MsgBox "You have to enter a year!"
        txtYear.SetFocus
        Exit Function
    Else
        If Val(txtYear) > Year(Date) + 100 Or Val(txtYear) < Year(Date) - 100 Then
            MsgBox "Year is not a valid year"
            txtYear.SetFocus
            Exit Function
        End If
    End If

    If optRolling Then
        If txtYearTo = "" Then
            MsgBox "You have to enter To Year!"
            txtYearTo.SetFocus
            Exit Function
        Else
            If Val(txtYearTo) > Year(Date) + 100 Or Val(txtYearTo) < Year(Date) - 100 Then
                MsgBox "To Year is not a valid year"
                txtYearTo.SetFocus
                Exit Function
            End If
        End If
        If CVDate(cmbMonthTo & " 01," & txtYearTo) < CVDate(cmbMonth & " 01," & txtYear) Then
            MsgBox "To Date can not be earlier than From Date"
            txtYear.SetFocus
            Exit Function
        End If
        
        '7.9 Enhancement
        If optExcel Then
            If DateDiff("m", CVDate(cmbMonth & " 01," & txtYear), CVDate(cmbMonthTo & " 01," & txtYearTo)) > 7 Then
                MsgBox "You can not view this report for more than 8 months"
                txtYear.SetFocus
                Exit Function
            End If
        Else
            If DateDiff("m", CVDate(cmbMonth & " 01," & txtYear), CVDate(cmbMonthTo & " 01," & txtYearTo)) > 11 Then
                MsgBox "You can not view this report for more than 12 months"
                txtYear.SetFocus
                Exit Function
            End If
        End If
    End If
    
    If optScheVAC Then 'Ticket #15212
        If txtYearTo = "" Then
            MsgBox "You have to enter To Year!"
            txtYearTo.SetFocus
            Exit Function
        Else
            If Val(txtYearTo) > Year(Date) + 100 Or Val(txtYearTo) < Year(Date) - 100 Then
                MsgBox "To Year is not a valid year"
                txtYearTo.SetFocus
                Exit Function
            End If
        End If
        If CVDate(cmbMonthTo & " 01," & txtYearTo) < CVDate(cmbMonth & " 01," & txtYear) Then
            MsgBox "To Date can not be earlier than From Date"
            txtYear.SetFocus
            Exit Function
        End If
        If DateDiff("m", CVDate(cmbMonth & " 01," & txtYear), CVDate(cmbMonthTo & " 01," & txtYearTo)) > 5 Then
            MsgBox "You can not view this report for more than 6 months"
            txtYear.SetFocus
            Exit Function
        End If
    End If
    
End If

If Not elpEEID.ListChecker Then Exit Function
If Not elpSUP(1).ListChecker Then Exit Function
If Not clpCode(10).ListChecker Then Exit Function

'If glbCompSerial = "S/N - 2214W" Then
'    If Len(clpCode(7)) > 0 Then
'        If clpCode(7).Caption = "Unassigned" Then
'            MsgBox "Invalid Attendance-Fund code"
'            clpCode(7).SetFocus
'            Exit Function
'        End If
'    End If
'End If
CriCheck = True
End Function

Private Sub Form_Activate()
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()

Screen.MousePointer = HOURGLASS

glbOnTop = "FRMRATTSHT"

Call comGrpLoad
Call setCaption(lblDiv)
Call setCaption(lblRegion)
Call setCaption(lblSection)
Call setCaption(lblDept)
Call setCaption(lblEENum(1))
Call setCaption(lblShift)
Call setRptCaption(Me)

If lblEENum(1).Caption = "AttSupervisor" Then lblEENum(1).Caption = "Supervisor"
lblAttCodes.Caption = lStr("Reason")

If glbCompSerial = "S/N - 2227W" Then clpCode(1).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

Call INI_Controls(Me)

chkShowEmp.Visible = True
If glbLinamar Then
    clpCode(1).MaxLength = 8
End If
If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If
If glbCompSerial = "S/N - 2241W" Then 'granite club
    optBiWeek.Visible = True
Else
    optBiWeek.Visible = False
End If
If glbCompSerial = "S/N - 2386W" Then 'The Walter Fedy Partnership Ticket #15212
    optScheVAC.Visible = True
    optScheVAC.Left = optBiWeek.Left
End If
If glbWFC Then   'Ticket #13256
    lblSection.FontBold = True
End If

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Screen.MousePointer = DEFAULT


End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Resize()
scrFrame.Height = 9375
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 9750 Then
        scrControl.Value = 0
        scrFrame.Top = 120
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 5600 Then
            scrControl.Max = 6300
        Else
            scrControl.Max = 3200
        End If
        scrControl.Left = Me.Width - scrControl.Width - 220
        If Me.Height - scrControl.Top - 780 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - 780
        End If
    End If
    scrFrame.Refresh
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
gdbAdoIhr001.Execute "DELETE FROM HR_ATTCAL " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub

Private Sub optAnnual_Click()
'If Added by Bryan 10/08/05 Ticket#8924
'comment by Bryan on Sep 1, 2005
'If glbCompSerial = "S/N - 2362W" And Len(clpCode(10).Text) > 0 Then
'    lblMonth.Visible = optAnnual
'    cmbMonth.Visible = optAnnual
'    lblMonthTo.Visible = optAnnual
'    cmbMonthTo.Visible = optAnnual
'    lblMonth.Caption = "From Month"
'    txtYear.Visible = Not (optWeek Or optBiWeek)
'    If cmbMonth.ListIndex = -1 Then cmbMonth.ListIndex = 0
'    If cmbMonthTo.ListIndex = -1 Then cmbMonthTo.ListIndex = 0
'Else
    lblMonth.Visible = optMonth
    cmbMonth.Visible = optMonth
    lblMonthTo.Visible = optRolling
    cmbMonthTo.Visible = optRolling
    lblYearTo.Visible = optRolling
    txtYearTo.Visible = optRolling
    lblYear.Caption = "Year"
    lblMonth.Caption = "Month"
    dlpStartDate.Visible = optWeek Or optBiWeek
    txtYear.Visible = Not (optWeek Or optBiWeek)
    chkBlank.Visible = False
'End If
End Sub

Private Sub optBiWeek_Click()
    lblYear = "Start Date"
    dlpStartDate.Visible = optWeek Or optBiWeek
    
    txtYear.Visible = Not (optWeek Or optBiWeek)
    lblMonth.Visible = Not (optWeek Or optBiWeek)
    cmbMonth.Visible = Not (optWeek Or optBiWeek)
    
    lblYearTo.Visible = optRolling
    txtYearTo.Visible = optRolling
    lblMonthTo.Visible = optRolling
    cmbMonthTo.Visible = optRolling
    chkBlank.Visible = True
End Sub

Private Sub optCrystal_Click()
    If optCrystal Then
        optAnnual.Visible = True
        optWeek.Visible = True
        
        If glbCompSerial = "S/N - 2241W" Then 'granite club
            optBiWeek.Visible = True
        Else
            optBiWeek.Visible = False
        End If
        
        If glbCompSerial = "S/N - 2386W" Then 'The Walter Fedy Partnership Ticket #15212
            optScheVAC.Visible = True
            optScheVAC.Left = optBiWeek.Left
        Else
            optScheVAC.Visible = False
        End If
    End If
End Sub

Private Sub optExcel_Click()
    If optExcel Then
        optAnnual.Visible = False
        optWeek.Visible = False
        optBiWeek.Visible = False
        optMonth.Value = True
        Call optMonth_Click
    End If
End Sub

Private Sub optMonth_Click()
    lblMonth.Visible = optMonth
    cmbMonth.Visible = optMonth
    lblMonthTo.Visible = optRolling
    cmbMonthTo.Visible = optRolling
    lblYearTo.Visible = optRolling
    txtYearTo.Visible = optRolling
    lblYear.Caption = "Year"
    lblMonth.Caption = "Month"
    dlpStartDate.Visible = optWeek Or optBiWeek
    txtYear.Visible = Not (optWeek Or optBiWeek)
    If cmbMonth.ListIndex = -1 Then cmbMonth.ListIndex = 0
    chkBlank.Visible = False
End Sub

Private Sub Cri_Dept()
    Dim countr   As Integer  ' EEList_Snap is definded at form level
    Dim DeptCri As String
    If Len(clpDept.Text) > 0 Then DeptCri = " AND (ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')) "
    glbstrSelCri = glbSeleDeptUn & DeptCri
End Sub

Private Sub AttWrk()
Dim SQLQ, ISQLQ
Dim rsAT As New ADODB.Recordset
Dim rsAW As New ADODB.Recordset
Dim rsHL As New ADODB.Recordset
Dim xEMPNBR, xDOA, xField
Dim xxx, xx1, x
Dim Y
Dim xDate
Dim xWeekDay
Dim xMons, DATEx, z
Dim xDay, xBuf
Dim date4, date5
Dim strCode As String
Dim strFilter As String

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 20
MDIMain.panHelp(1).Caption = " Please Wait"

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute "DELETE FROM HR_ATTCAL " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.CommitTrans

If optMonth Then
    DATE1 = cmbMonth & " 1, " & txtYear
    SQLQ = " SELECT ED_COMPNO AS AD_COMPNO,'" & glbUserID & "' AS AD_WRKEMP,ED_EMPNBR AS AD_EMPNBR ,"
    SQLQ = SQLQ & Date_SQL(DATE1) & " AS AD_DOA"
    SQLQ = SQLQ & " FROM HREMP "
    SQLQ = SQLQ & " WHERE " & glbstrSelCri
    If Not chkShowEmp Then
        SQLQ = SQLQ & " AND ED_EMPNBR IN ("
        If glbOracle Then
            SQLQ = SQLQ & " SELECT DISTINCT AD_EMPNBR FROM HR_ATTENDANCE "
            SQLQ = SQLQ & " WHERE TO_CHAR(AD_DOA,'YYYY')=" & txtYear
            SQLQ = SQLQ & " AND TO_CHAR(AD_DOA,'MM')=" & (cmbMonth.ListIndex + 1)
            If Len(txtShift.Text) > 0 Then
                SQLQ = SQLQ & " AND AD_SHIFT= '" & txtShift.Text & "'"
            End If
            
            If Len(clpCode(10).Text) > 0 Then SQLQ = SQLQ & " AND AD_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
            If chkAbsence.Value Then
                SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
            End If
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT DISTINCT AH_EMPNBR FROM HR_ATTENDANCE_HISTORY "
            SQLQ = SQLQ & " WHERE TO_CHAR(AH_DOA,'YYYY')=" & txtYear
            SQLQ = SQLQ & " AND TO_CHAR(AH_DOA,'MM')=" & (cmbMonth.ListIndex + 1)
            If Len(txtShift.Text) > 0 Then
                SQLQ = SQLQ & " AND AH_SHIFT= '" & txtShift.Text & "'"
            End If
            
            If Len(clpCode(10).Text) > 0 Then SQLQ = SQLQ & " AND AH_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
            If chkAbsence.Value Then
                SQLQ = SQLQ & " AND AH_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
            End If
        ElseIf glbSQL Then
        
            SQLQ = SQLQ & " SELECT DISTINCT AD_EMPNBR FROM HR_ATTENDANCE "
            SQLQ = SQLQ & " WHERE YEAR(AD_DOA)=" & txtYear
            SQLQ = SQLQ & " AND MONTH(AD_DOA)=" & (cmbMonth.ListIndex + 1)
            If Len(txtShift.Text) > 0 Then
                SQLQ = SQLQ & " AND AD_SHIFT= '" & txtShift.Text & "'"
            End If
            
            If Len(clpCode(10).Text) > 0 Then SQLQ = SQLQ & " AND AD_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
            If chkAbsence.Value Then
                SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
            End If
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT DISTINCT AH_EMPNBR FROM HR_ATTENDANCE_HISTORY "
            SQLQ = SQLQ & " WHERE YEAR(AH_DOA)=" & txtYear
            SQLQ = SQLQ & " AND MONTH(AH_DOA)=" & (cmbMonth.ListIndex + 1)
            If Len(txtShift.Text) > 0 Then
                SQLQ = SQLQ & " AND AH_SHIFT= '" & txtShift.Text & "'"
            End If
            
            If Len(clpCode(10).Text) > 0 Then SQLQ = SQLQ & " AND AH_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
            If chkAbsence.Value Then
                SQLQ = SQLQ & " AND AH_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
            End If
        Else
            SQLQ = SQLQ & " SELECT DISTINCT AD_EMPNBR FROM qry_EMPNBR_ATTENDANCE WHERE YEAR(AD_DOA)=" & txtYear
            SQLQ = SQLQ & " AND MONTH(AD_DOA)=" & (cmbMonth.ListIndex + 1)
            If Len(txtShift.Text) > 0 Then
                SQLQ = SQLQ & " AND AD_SHIFT= '" & txtShift.Text & "'"
            End If
            
            If Len(clpCode(10).Text) > 0 Then SQLQ = SQLQ & " AND AD_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
            If chkAbsence.Value Then
                SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
            End If
        End If
        SQLQ = SQLQ & " )"
    End If
    
    ISQLQ = "INSERT INTO HR_ATTCAL (AD_COMPNO,AD_WRKEMP,AD_EMPNBR,AD_DOA) " & in_SQL(glbIHRDBW) & SQLQ
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute ISQLQ
    gdbAdoIhr001.CommitTrans
    
    For x = 1 To 31
        xDate = cmbMonth & " " & x & " ," & txtYear
        If IsDate(xDate) Then
            'xWeekDay = Left(WeekdayName(Weekday(CVDate(xDATE))), 1)
            xWeekDay = Left(WeekdayName(Weekday(CVDate(xDate))), 3)
            xBuf = ""
            'If xWeekDay = "S" Then
            If UCase(xWeekDay) = "SAT" Or UCase(xWeekDay) = "SUN" Then
                'gdbAdoIhr001.Execute "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & " SET AD_DAY" & X & "='S '"
                'gdbAdoIhr001.Execute "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & " SET AD_DAY" & X & "='" & UCase(xWeekDay) & "'"
                 xBuf = UCase(xWeekDay)
            End If
            If glbWFC Then   'Ticket #13256
                rsHL.Open "SELECT * FROM HR_HOLIDAY WHERE HL_DATE=" & Date_SQL(xDate) & " AND HL_SECTION = '" & clpCode(2).Text & "'", gdbAdoIhr001, adOpenStatic
            Else
                rsHL.Open "SELECT * FROM HR_HOLIDAY WHERE HL_DATE=" & Date_SQL(xDate), gdbAdoIhr001, adOpenStatic
            End If
            If Not rsHL.EOF Then
                'gdbAdoIhr001.Execute "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & " SET AD_DAY" & X & "='H '"
                If Len(xBuf) > 0 Then
                    xBuf = xBuf & Chr$(10) & " H "
                Else
                    xBuf = xBuf & " H "
                End If
            End If
            rsHL.Close
            If Len(xBuf) > 0 Then
                gdbAdoIhr001.Execute "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & " SET AD_DAY" & x & "='" & xBuf & "'"
            End If
            DATE2 = xDate
        Else
            gdbAdoIhr001.Execute "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & " SET AD_DAY" & x & "='*'"
        End If
    Next
ElseIf optWeek Or optBiWeek Then
    DATE1 = dlpStartDate
    If optWeek Then
        DATE2 = DateAdd("d", 6, dlpStartDate)
    Else
        DATE2 = DateAdd("d", 13, dlpStartDate)
    End If
    
    SQLQ = " SELECT ED_COMPNO AS AD_COMPNO,'" & glbUserID & "' AS AD_WRKEMP,ED_EMPNBR AS AD_EMPNBR ,"
    'SQLQ = SQLQ & IIf(glbSQL, "", "CVDATE") & "('" & DATE1 & "') AS AD_DOA"
    SQLQ = SQLQ & Date_SQL(DATE1) & " AS AD_DOA"
    SQLQ = SQLQ & " FROM HREMP "
    SQLQ = SQLQ & " WHERE " & glbstrSelCri
    If Not chkShowEmp Then
        SQLQ = SQLQ & " AND ED_EMPNBR IN ("
        If glbOracle Then
            SQLQ = SQLQ & " SELECT DISTINCT AD_EMPNBR FROM HR_ATTENDANCE "
            SQLQ = SQLQ & " WHERE AD_DOA>=" & Date_SQL(DATE1) & " "
            SQLQ = SQLQ & " AND   AD_DOA<=" & Date_SQL(DATE2) & " "
            If Len(txtShift.Text) > 0 Then
                SQLQ = SQLQ & " AND AD_SHIFT= '" & txtShift.Text & "'"
            End If
            
            If Len(clpCode(10).Text) > 0 Then SQLQ = SQLQ & " AND AD_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
            If chkAbsence.Value Then
                SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
            End If
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT DISTINCT AH_EMPNBR FROM HR_ATTENDANCE_HISTORY "
            SQLQ = SQLQ & " WHERE AH_DOA>=" & Date_SQL(DATE1) & " "
            SQLQ = SQLQ & " AND   AH_DOA<=" & Date_SQL(DATE2) & " "
            If Len(txtShift.Text) > 0 Then
                SQLQ = SQLQ & " AND AH_SHIFT= '" & txtShift.Text & "'"
            End If
            
            If Len(clpCode(10).Text) > 0 Then SQLQ = SQLQ & " AND AH_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
            If chkAbsence.Value Then
                SQLQ = SQLQ & " AND AH_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
            End If
        ElseIf glbSQL Then
            SQLQ = SQLQ & " SELECT DISTINCT AD_EMPNBR FROM HR_ATTENDANCE "
            SQLQ = SQLQ & " WHERE AD_DOA>='" & Format(DATE1, "mmm dd, yyyy") & "' "
            SQLQ = SQLQ & " AND   AD_DOA<='" & Format(DATE2, "mmm dd, yyyy") & "' "
            If Len(txtShift.Text) > 0 Then
                SQLQ = SQLQ & " AND AD_SHIFT= '" & txtShift.Text & "'"
            End If
            
            If Len(clpCode(10)) > 0 Then SQLQ = SQLQ & " AND AD_REASON in ('" & Replace(clpCode(10), ",", "','") & "') "
            If chkAbsence.Value Then
                SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
            End If
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT DISTINCT AH_EMPNBR FROM HR_ATTENDANCE_HISTORY "
            SQLQ = SQLQ & " WHERE AH_DOA>='" & Format(DATE1, "mmm dd, yyyy") & "' "
            SQLQ = SQLQ & " AND  AH_DOA<='" & Format(DATE2, "mmm dd, yyyy") & "' "
            If Len(txtShift.Text) > 0 Then
                SQLQ = SQLQ & " AND AH_SHIFT= '" & txtShift.Text & "'"
            End If
            
            If Len(clpCode(10)) > 0 Then SQLQ = SQLQ & " AND AH_REASON in ('" & Replace(clpCode(10), ",", "','") & "') "
            If chkAbsence.Value Then
                SQLQ = SQLQ & " AND AH_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
            End If
        Else
            SQLQ = SQLQ & " SELECT DISTINCT AD_EMPNBR FROM qry_EMPNBR_ATTENDANCE  "
            SQLQ = SQLQ & " WHERE AD_DOA>=CVDATE('" & Format(DATE1, "mmm dd, yyyy") & "')"
            SQLQ = SQLQ & " AND   AD_DOA<=CVDATE('" & Format(DATE2, "mmm dd, yyyy") & "') "
            If Len(txtShift.Text) > 0 Then
                SQLQ = SQLQ & " AND AD_SHIFT= '" & txtShift.Text & "'"
            End If
            
            If Len(clpCode(10)) > 0 Then SQLQ = SQLQ & " AND AD_REASON in ('" & Replace(clpCode(10), ",", "','") & "') "
            If chkAbsence.Value Then
                SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
            End If
        End If
        SQLQ = SQLQ & " ) "
    End If
    
    ISQLQ = "INSERT INTO HR_ATTCAL (AD_COMPNO,AD_WRKEMP,AD_EMPNBR,AD_DOA) " & in_SQL(glbIHRDBW) & SQLQ
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute ISQLQ
    gdbAdoIhr001.CommitTrans
    
    For x = 1 To IIf(optWeek, 7, 14)
        If Not (glbCompSerial = "S/N - 2241W" And optBiWeek) Then
            xDate = DateAdd("d", x - 1, dlpStartDate)
            If IsDate(xDate) Then
                'xWeekDay = Left(WeekdayName(Weekday(CVDate(xDATE))), 1)
                xWeekDay = Left(WeekdayName(Weekday(CVDate(xDate))), 3)
                xBuf = ""
                'If xWeekDay = "S" Then
                If UCase(xWeekDay) = "SAT" Or UCase(xWeekDay) = "SUN" Then
                    'gdbAdoIhr001.Execute "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & " SET AD_DAY" & X & "='S '"
                    'gdbAdoIhr001.Execute "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & " SET AD_DAY" & X & "='" & UCase(xWeekDay) & "'"
                    xBuf = UCase(xWeekDay)
                End If
                'rsHL.Open "SELECT * FROM HR_HOLIDAY WHERE HL_DATE=" & IIf(glbSQL, "", "CVDATE") & "('" & Format(xDATE, "mmm dd, yyyy") & "')", gdbAdoIhr001, adOpenStatic
                If glbWFC Then   'Ticket #13256
                    rsHL.Open "SELECT * FROM HR_HOLIDAY WHERE HL_DATE=" & Date_SQL(xDate) & " AND HL_SECTION = '" & clpCode(2).Text & "'", gdbAdoIhr001, adOpenStatic
                Else
                    rsHL.Open "SELECT * FROM HR_HOLIDAY WHERE HL_DATE=" & Date_SQL(xDate), gdbAdoIhr001, adOpenStatic
                End If
                If Not rsHL.EOF Then
                    'gdbAdoIhr001.Execute "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & " SET AD_DAY" & X & "='H '"
                    If Len(xBuf) > 0 Then
                        xBuf = xBuf & Chr$(10) & " H "
                    Else
                        xBuf = xBuf & " H "
                    End If
                End If
                rsHL.Close
                If Len(xBuf) > 0 Then
                    gdbAdoIhr001.Execute "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & " SET AD_DAY" & x & "='" & xBuf & "'"
                End If
            Else
                gdbAdoIhr001.Execute "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & " SET AD_DAY" & x & "='*'"
            End If
        End If
    Next
Else
    If optAnnual Then
        DATE1 = GetMonth("Jan") & " 1, " & txtYear
        DATE2 = GetMonth("Dec") & " 31, " & txtYear
        DATEx = DATE1
        xMons = 11
    ElseIf optRolling And glbCompSerial <> "S/N - 2362W" Then
        DATE1 = cmbMonth & " 1, " & txtYear
        DATE2 = cmbMonthTo & " 1, " & txtYearTo
        DATEx = DATE1
        xMons = DateDiff("m", DATE1, DATE2)
        DATE2 = DateAdd("m", 1, DATE2)
        DATE2 = DateAdd("d", -1, DATE2)
    ElseIf optScheVAC Then
        DATE1 = cmbMonth & " 1, " & txtYear
        DATE2 = cmbMonthTo & " 1, " & txtYearTo
        DATEx = DATE1
        xMons = DateDiff("m", DATE1, DATE2)
        DATE2 = DateAdd("m", 1, DATE2)
        DATE2 = DateAdd("d", -1, DATE2)
    Else
        DATE1 = GetMonth("Jan") & " 1, " & txtYear
        DATE2 = GetMonth("Dec") & " 31, " & txtYear
        DATEx = DATE1
        xMons = 11
        date4 = cmbMonth & " 1, " & txtYear
        date5 = cmbMonthTo & " 1, " & txtYearTo
        date5 = DateAdd("m", 1, date5)
        date5 = DateAdd("d", -1, date5)
    End If
    For x = 0 To xMons '
        z = month(DATEx) - 1
        SQLQ = " SELECT ED_COMPNO AS AD_COMPNO,'" & glbUserID & "' AS AD_WRKEMP,ED_EMPNBR AS AD_EMPNBR ,"
        SQLQ = SQLQ & Date_SQL(cmbMonth.List(z) & " 1," & Year(DATEx)) & " AS AD_DOA"
        SQLQ = SQLQ & " FROM HREMP "
        SQLQ = SQLQ & " WHERE " & glbstrSelCri
        'Frank uncomment the chkShowEmp function, Ticket #13442, but WFC will not have this function since Ticket #12982
        'Hemu - Begin - Ticket #12982 - Show the months in the report - if part of date range and do not have any attendance records.
        If Not chkShowEmp And Not glbWFC Then
            SQLQ = SQLQ & " AND ED_EMPNBR IN ("
            If glbOracle Then
                SQLQ = SQLQ & " SELECT DISTINCT AD_EMPNBR FROM HR_ATTENDANCE WHERE TO_CHAR(AD_DOA,'YYYY')=" & Year(DATEx)

                If Len(txtShift.Text) > 0 Then
                    SQLQ = SQLQ & " AND AD_SHIFT= '" & txtShift.Text & "'"
                End If
                If Len(clpCode(10).Text) > 0 Then
                        SQLQ = SQLQ & " AND AD_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
                End If
                If chkAbsence.Value Then
                    SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
                End If
                SQLQ = SQLQ & " UNION "
                SQLQ = SQLQ & " SELECT DISTINCT AH_EMPNBR FROM HR_ATTENDANCE_HISTORY WHERE TO_CHAR(AH_DOA,'YYYY')=" & Year(DATEx)

                If Len(txtShift.Text) > 0 Then
                    SQLQ = SQLQ & " AND AH_SHIFT= '" & txtShift.Text & "'"
                End If

                If Len(clpCode(10).Text) > 0 Then SQLQ = SQLQ & " AND AH_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
                If chkAbsence.Value Then
                    SQLQ = SQLQ & " AND AH_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
                End If
            ElseIf glbSQL Then
                SQLQ = SQLQ & " SELECT DISTINCT AD_EMPNBR FROM HR_ATTENDANCE WHERE "

                If optRolling And glbCompSerial = "S/N - 2362W" Then
                    SQLQ = SQLQ & "AD_DOA >= " & Date_SQL(date4) & " AND AD_DOA <= " & Date_SQL(date5)
                ElseIf optScheVAC Then
                    SQLQ = SQLQ & "AD_DOA >= " & Date_SQL(date4) & " AND AD_DOA <= " & Date_SQL(date5)
                Else
                    SQLQ = SQLQ & "Year(AD_DOA) = " & Year(DATEx)
                End If

                If Len(txtShift.Text) > 0 Then
                    SQLQ = SQLQ & " AND AD_SHIFT= '" & txtShift.Text & "'"
                End If

                If Len(clpCode(10).Text) > 0 Then
                        SQLQ = SQLQ & " AND AD_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
                End If
                If chkAbsence.Value Then
                    SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
                End If
                If Len(elpSUP(1).Text) > 0 Then SQLQ = SQLQ & " AND AD_SUPER IN (" & getEmpnbr(elpSUP(1).Text) & ") "
                SQLQ = SQLQ & " UNION "
                SQLQ = SQLQ & " SELECT DISTINCT AH_EMPNBR FROM HR_ATTENDANCE_HISTORY WHERE "

                If optRolling And glbCompSerial = "S/N - 2362W" Then
                    SQLQ = SQLQ & "AH_DOA >= " & Date_SQL(date4) & " AND AH_DOA <= " & Date_SQL(date5)
                ElseIf optScheVAC Then
                    SQLQ = SQLQ & "AH_DOA >= " & Date_SQL(date4) & " AND AH_DOA <= " & Date_SQL(date5)
                Else
                    SQLQ = SQLQ & "Year(AH_DOA) = " & Year(DATEx)
                End If

                If Len(txtShift.Text) > 0 Then
                    SQLQ = SQLQ & " AND AH_SHIFT= '" & txtShift.Text & "'"
                End If

                If Len(clpCode(10).Text) > 0 Then
                    SQLQ = SQLQ & " AND AH_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
                End If
                If chkAbsence.Value Then
                    SQLQ = SQLQ & " AND AH_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
                End If
            Else
                SQLQ = SQLQ & " SELECT DISTINCT AD_EMPNBR FROM QRY_EMPNBR_ATTENDANCE WHERE YEAR(AD_DOA)=" & Year(DATEx)

                If Len(txtShift.Text) > 0 Then
                    SQLQ = SQLQ & " AND AD_SHIFT= '" & txtShift.Text & "'"
                End If

                If Len(clpCode(10).Text) > 0 Then
                        SQLQ = SQLQ & " AND AD_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
                End If

                If chkAbsence.Value Then
                    SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
                End If
            End If
            SQLQ = SQLQ & " )"
        End If
        'Hemu - End - Ticket #12982

        ISQLQ = "INSERT INTO HR_ATTCAL(AD_COMPNO,AD_WRKEMP,AD_EMPNBR,AD_DOA) " & in_SQL(glbIHRDBW) & SQLQ
        'Debug.Print ISQLQ
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute ISQLQ
        gdbAdoIhr001.CommitTrans
        For Y = 1 To 31
            'xDATE = cmbMonth.List(X) & " " & Y & ", " & txtYear
            xDate = cmbMonth.List(z) & " " & Y & ", " & Year(DATEx)
            If IsDate(xDate) Then
                'xWeekDay = Left(WeekdayName(Weekday(CVDate(xDATE))), 1)
                xWeekDay = Left(WeekdayName(Weekday(CVDate(xDate))), 3)
                xBuf = ""
                'If xWeekDay = "S" Then
                If UCase(xWeekDay) = "SAT" Or UCase(xWeekDay) = "SUN" Then
                    'SQLQ = "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & "SET AD_DAY" & Y & "='S ' "
'                    SQLQ = "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & "SET AD_DAY" & Y & "='" & UCase(xWeekDay) & "'"
'                    If glbOracle Then
'                        SQLQ = SQLQ & " WHERE TO_NUMBER(TO_CHAR(AD_DOA,'MM'))= " & (z + 1)
'                    Else
'                        SQLQ = SQLQ & " WHERE MONTH(AD_DOA)= " & (z + 1)
'                    End If
'                    gdbAdoIhr001.Execute SQLQ
                    xBuf = UCase(xWeekDay)
                End If
                If glbWFC Then   'Ticket #13256
                    rsHL.Open "SELECT * FROM HR_HOLIDAY WHERE HL_DATE=" & Date_SQL(xDate) & " AND HL_SECTION = '" & clpCode(2).Text & "'", gdbAdoIhr001, adOpenStatic
                Else
                    rsHL.Open "SELECT * FROM HR_HOLIDAY WHERE HL_DATE=" & Date_SQL(xDate), gdbAdoIhr001, adOpenStatic
                End If
                If Not rsHL.EOF Then
'                        SQLQ = "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & " SET AD_DAY" & Y & "='H '"
'                        If glbOracle Then
'                             SQLQ = SQLQ & " WHERE TO_NUMBER(TO_CHAR(AD_DOA,'MM'))=" & (z + 1)
'                        Else
'                             SQLQ = SQLQ & " WHERE MONTH(AD_DOA)=" & (z + 1)
'                        End If
'                        gdbAdoIhr001.Execute SQLQ
                    If Len(xBuf) > 0 Then
                        xBuf = xBuf & Chr$(10) & " H "
                    Else
                        xBuf = xBuf & " H "
                    End If
                End If
                rsHL.Close
                If Len(xBuf) > 0 Then
                    SQLQ = "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & " SET AD_DAY" & Y & "='" & xBuf & "'"
                    If glbOracle Then
                         SQLQ = SQLQ & " WHERE TO_NUMBER(TO_CHAR(AD_DOA,'MM'))=" & (z + 1)
                    Else
                         SQLQ = SQLQ & " WHERE MONTH(AD_DOA)=" & (z + 1)
                    End If
                    gdbAdoIhr001.Execute SQLQ
                End If
             Else
                SQLQ = "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & " SET AD_DAY" & Y & "='*'"
                If glbOracle Then
                    SQLQ = SQLQ & " WHERE TO_NUMBER(TO_CHAR(AD_DOA,'MM'))=" & (z + 1)
                Else
                    SQLQ = SQLQ & " WHERE MONTH(AD_DOA)=" & (z + 1)
                End If
                gdbAdoIhr001.Execute SQLQ
            End If
        Next
        DATEx = DateAdd("M", 1, DATEx)
    Next
End If
MDIMain.panHelp(0).FloodPercent = 30
SQLQ = "SELECT AD_COMPNO,'" & glbUserID & "' AS AD_WRKEMP,AD_EMPNBR,AD_DOA,AD_HRS,AD_REASON, AD_SHIFT "
SQLQ = SQLQ & " FROM HR_ATTENDANCE "
SQLQ = SQLQ & " WHERE AD_EMPNBR IN (SELECT AD_EMPNBR FROM HR_ATTCAL" & in_SQL(glbIHRDBW) & ") "
SQLQ = SQLQ & " AND AD_DOA>=" & Date_SQL(DATE1)
SQLQ = SQLQ & " AND AD_DOA<=" & Date_SQL(DATE2)

'If Len(clpCode(10).Text) > 0 And Not (glbCompSerial = "S/N - 2362W" And optRolling) Then
If Len(clpCode(10).Text) > 0 Then 'Ticket #13442
    SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
End If

'Franks 06/20/2003 ticket# 4125 Absence
If chkAbsence.Value Then
    SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
End If

'-----------------------------------------------------------------------------------------
'Hemu - Ticket #13616 - filter records as per the selection criteria to speed up the process
If Len(elpSUP(1).Text) > 0 Then     'Supervisor
    SQLQ = SQLQ & " AND AD_SUPER IN (" & getEmpnbr(elpSUP(1).Text) & ") "
End If

'From HREMP
strFilter = ""
If Len(clpDiv.Text) > 0 Then    'Division
    strFilter = strFilter & " (ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "'))"
End If
If Len(clpDept.Text) > 0 Then   'Department
    If Len(strFilter) > 0 Then
        strFilter = strFilter & " AND (ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')) "
    Else
        strFilter = strFilter & " (ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')) "
    End If
End If

For x = 0 To 6
    If Len(clpCode(x).Text) > 0 Then
        Select Case x
            Case 0: strCode = "ED_LOC"
            Case 1: strCode = "ED_REGION"
            Case 2: strCode = "ED_SECTION"
            Case 3: strCode = "ED_EMP"
            Case 4: strCode = "ED_ADMINBY"
            Case 5: strCode = "ED_ORG"
            Case 6: strCode = "ED_PT"
        End Select
        
        If glbLinamar And (strCode = "ED_REGION" Or strCode = "ED_SECTION") Then
            If Len(strFilter) > 0 Then
                strFilter = strFilter & " AND (" & strCode & " = ('" & clpDiv.Text & clpCode(x).Text & "') or (" & strCode & " = 'ALL" & clpCode(x).Text & "'))"
            Else
                strFilter = strFilter & " (" & strCode & " = ('" & clpDiv.Text & clpCode(x).Text & "') or (" & strCode & " = 'ALL" & clpCode(x).Text & "'))"
            End If
        Else
            If Len(strFilter) > 0 Then
                strFilter = strFilter & " AND (" & strCode & " in ('" & Replace(clpCode(x).Text, ",", "','") & "'))"
            Else
                strFilter = strFilter & "(" & strCode & " in ('" & Replace(clpCode(x).Text, ",", "','") & "'))"
            End If
        End If
    End If
Next x

If Len(elpEEID.Text) > 0 Then   'Employee
    If Len(strFilter) > 0 Then
        strFilter = strFilter & " AND ED_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
    Else
        strFilter = strFilter & " ED_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
    End If
End If

If Len(strFilter) > 0 Then
    strFilter = " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & strFilter & ")"
    SQLQ = SQLQ & strFilter
End If
'-----------------------------------------------------------------------------------------

SQLQ = SQLQ & " UNION "
SQLQ = SQLQ & " SELECT AH_COMPNO,'" & glbUserID & "' AS AH_WRKEMP,AH_EMPNBR,AH_DOA,AH_HRS,AH_REASON, AH_SHIFT "
SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
SQLQ = SQLQ & " WHERE AH_EMPNBR IN (SELECT AD_EMPNBR FROM HR_ATTCAL " & in_SQL(glbIHRDBW) & ") "
SQLQ = SQLQ & " AND AH_DOA>=" & Date_SQL(DATE1)
SQLQ = SQLQ & " AND AH_DOA<=" & Date_SQL(DATE2)

'If Len(clpCode(10).Text) > 0 And Not (glbCompSerial = "S/N - 2362W" And optRolling) Then
If Len(clpCode(10).Text) > 0 Then 'Ticket #13442
    SQLQ = SQLQ & " AND HR_ATTENDANCE_HISTORY.AH_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
End If
'Franks 06/20/2003 ticket# 4125 Absence
If chkAbsence.Value Then
    SQLQ = SQLQ & " AND HR_ATTENDANCE_HISTORY.AH_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
End If

'-----------------------------------------------------------------------------------------
'Hemu - Ticket #13616 - filter records as per the selection criteria to speed up the process
If Len(elpSUP(1).Text) > 0 Then     'Supervisor
    SQLQ = SQLQ & " AND AH_SUPER IN (" & getEmpnbr(elpSUP(1).Text) & ") "
End If

'From HREMP
strFilter = ""
If Len(clpDiv.Text) > 0 Then    'Division
    strFilter = strFilter & " (ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "'))"
End If
If Len(clpDept.Text) > 0 Then   'Department
    If Len(strFilter) > 0 Then
        strFilter = strFilter & " AND (ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')) "
    Else
        strFilter = strFilter & " (ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')) "
    End If
End If

For x = 0 To 6
    If Len(clpCode(x).Text) > 0 Then
        Select Case x
            Case 0: strCode = "ED_LOC"
            Case 1: strCode = "ED_REGION"
            Case 2: strCode = "ED_SECTION"
            Case 3: strCode = "ED_EMP"
            Case 4: strCode = "ED_ADMINBY"
            Case 5: strCode = "ED_ORG"
            Case 6: strCode = "ED_PT"
        End Select
        
        If glbLinamar And (strCode = "ED_REGION" Or strCode = "ED_SECTION") Then
            If Len(strFilter) > 0 Then
                strFilter = strFilter & " AND (" & strCode & " = ('" & clpDiv.Text & clpCode(x).Text & "') or (" & strCode & " = 'ALL" & clpCode(x).Text & "'))"
            Else
                strFilter = strFilter & " (" & strCode & " = ('" & clpDiv.Text & clpCode(x).Text & "') or (" & strCode & " = 'ALL" & clpCode(x).Text & "'))"
            End If
        Else
            If Len(strFilter) > 0 Then
                strFilter = strFilter & " AND (" & strCode & " in ('" & Replace(clpCode(x).Text, ",", "','") & "'))"
            Else
                strFilter = strFilter & "(" & strCode & " in ('" & Replace(clpCode(x).Text, ",", "','") & "'))"
            End If
        End If
    End If
Next x

If Len(elpEEID.Text) > 0 Then   'Employee
    If Len(strFilter) > 0 Then
        strFilter = strFilter & " AND ED_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
    Else
        strFilter = strFilter & " ED_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
    End If
End If


If Len(strFilter) > 0 Then
    strFilter = " AND AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & strFilter & ")"
    SQLQ = SQLQ & strFilter
End If
'-----------------------------------------------------------------------------------------

rsAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockReadOnly
xxx = rsAT.RecordCount
xx1 = 0
Do Until rsAT.EOF
    xx1 = xx1 + 1
    MDIMain.panHelp(0).FloodPercent = (xx1 / xxx) * 60 + 30
    'xField = "AD_DAY" & Day(rsAT!AD_DOA)
    If optWeek Or optBiWeek Then
         xDay = DateDiff("d", DATE1, rsAT!AD_DOA) + 1
    Else
        xDay = Day(rsAT!AD_DOA)
    End If
    xField = "AD_DAY" & xDay
        
    SQLQ = "UPDATE HR_ATTCAL " & in_SQL(glbIHRDBW) & " SET " & xField & "="
    If glbSQL Or glbOracle Then
        SQLQ = SQLQ & " (CASE WHEN " & xField & " IS NULL THEN '' ELSE " & xField & " END )"
    Else
        SQLQ = SQLQ & " IIF(" & xField & " IS NULL,''," & xField & ")"
    End If
    
    If glbOracle Then
        SQLQ = SQLQ & " || Chr(13) || Chr(10) || '" & rsAT!AD_HRS & "' || Chr(13)|| Chr(10)||'" & rsAT!AD_REASON & "'"
        SQLQ = SQLQ & " WHERE AD_EMPNBR=" & rsAT!AD_EMPNBR
        SQLQ = SQLQ & " AND TO_CHAR(AD_DOA,'MM')= TO_CHAR(" & Date_SQL(rsAT!AD_DOA) & ",'MM') "
        SQLQ = SQLQ & " AND TO_CHAR(AD_DOA,'YYYY')=TO_CHAR(" & Date_SQL(rsAT!AD_DOA) & ",'YYYY')"
    Else
        'Hemu - 02/16/2004 Begin - Ticket # 5600
        If glbCompSerial = "S/N - 2226W" Then
            SQLQ = SQLQ & " +'" & Chr(13) & Chr(10) & rsAT!AD_HRS & rsAT!AD_SHIFT & Chr(13) & Chr(10) & rsAT!AD_REASON & "'"
        ElseIf glbCompSerial = "S/N - 2241W" And optBiWeek Then
            If chkBlank Then
                SQLQ = SQLQ & " +''"
            Else
                SQLQ = SQLQ & " +(CASE WHEN " & xField & " IS NULL THEN '' ELSE '" & Chr(13) & Chr(10) & "' END )+'" & rsAT!AD_HRS & Chr(13) & Chr(10) & rsAT!AD_REASON & "'"
            End If
        Else
            SQLQ = SQLQ & " +'" & Chr(13) & Chr(10) & rsAT!AD_HRS & Chr(13) & Chr(10) & rsAT!AD_REASON & "'"
        End If
        'Hemu - 02/16/2004 End
        
        SQLQ = SQLQ & " WHERE AD_EMPNBR=" & rsAT!AD_EMPNBR
        If optWeek Or optBiWeek Then
            SQLQ = SQLQ & " AND AD_DOA=" & Date_SQL(DATE1)
        Else
        
            SQLQ = SQLQ & " AND MONTH(AD_DOA)=" & month(rsAT!AD_DOA)
            SQLQ = SQLQ & " AND YEAR(AD_DOA)=" & Year(rsAT!AD_DOA)
        End If
    End If
    

    SQLQ = SQLQ & " AND AD_WRKEMP='" & glbUserID & "'"
    
    gdbAdoIhr001.Execute SQLQ
    
    rsAT.MoveNext
Loop
MDIMain.panHelp(0).FloodPercent = 90

'7.9 Enhancement - Get the list of Reason Codes to display in the Excel Spreadsheet report as legend
xAllReason = ""
If rsAT.RecordCount <> 0 Then
    rsAT.Requery
    rsAT.MoveFirst
End If
Do While Not rsAT.EOF
    If InStr(1, xAllReason, rsAT!AD_REASON) = 0 Then
        If Len(xAllReason) = 0 Then
            xAllReason = xAllReason & "'" & rsAT!AD_REASON & "'"
        Else
            xAllReason = xAllReason & ",'" & rsAT!AD_REASON & "'"
        End If
    End If
    rsAT.MoveNext
Loop

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
End Sub

Private Sub optRolling_Click()
    lblMonth.Visible = optRolling
    cmbMonth.Visible = optRolling
    lblYearTo.Visible = optRolling
    txtYearTo.Visible = optRolling
    lblMonthTo.Visible = optRolling
    cmbMonthTo.Visible = optRolling
    
    lblYear.Caption = "From Year"
    lblMonth.Caption = "From Month"
    
    dlpStartDate.Visible = optWeek Or optBiWeek
    txtYear.Visible = Not (optWeek Or optBiWeek)
    
    If cmbMonth.ListIndex = -1 Then cmbMonth.ListIndex = 0
    If cmbMonthTo.ListIndex = -1 Then cmbMonthTo.ListIndex = 0
    
    chkBlank.Visible = False
End Sub

Private Sub optScheVAC_Click()
    lblMonth.Visible = optScheVAC
    cmbMonth.Visible = optScheVAC
    lblYearTo.Visible = optScheVAC
    txtYearTo.Visible = optScheVAC
    lblMonthTo.Visible = optScheVAC
    cmbMonthTo.Visible = optScheVAC
    
    lblYear.Caption = "From Year"
    lblMonth.Caption = "From Month"
    
    dlpStartDate.Visible = optWeek Or optBiWeek
    txtYear.Visible = Not (optWeek Or optBiWeek)
    
    If cmbMonth.ListIndex = -1 Then cmbMonth.ListIndex = 0
    If cmbMonthTo.ListIndex = -1 Then cmbMonthTo.ListIndex = 0
    
    chkBlank.Visible = False
    
    clpCode(10).Text = "VAC,SCHV"
End Sub

Private Sub optWeek_Click()
    lblYear = "Start Date"
    dlpStartDate.Visible = optWeek Or optBiWeek
    
    txtYear.Visible = Not (optWeek Or optBiWeek)
    lblMonth.Visible = Not (optWeek Or optBiWeek)
    cmbMonth.Visible = Not (optWeek Or optBiWeek)
    
    lblYearTo.Visible = optRolling
    txtYearTo.Visible = optRolling
    lblMonthTo.Visible = optRolling
    cmbMonthTo.Visible = optRolling
    chkBlank.Visible = False
End Sub

Private Sub scrControl_Change()
scrFrame.Top = 120 - scrControl.Value
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

Private Sub txtYear_GotFocus()      ' Serbo
Call SetPanHelp(Me.ActiveControl)   '
End Sub                             '

Private Sub Export_XLSWriter_BrantCounty()
    Dim startdt, enddt
    Dim Month1, Month2
    Dim xlsFileTmp, xlsFileMat
    Dim colno, lineno, x, mons, lstEmpNo
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim strTemp, SQLQ, adDay
    Dim rsHRAtt As New ADODB.Recordset
    Dim rsHREmp As New ADODB.Recordset
    
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "RollCTmp.xls"
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "RollCaln.xls"

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."

    FileCopy xlsFileTmp, xlsFileMat
    
    
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)

    'Print Titles
    exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy")
    exSheet.Cells(2, 1) = "Time: " & Time$
    
    strTemp = ""
    If DATE1 <> "" And DATE2 <> "" Then
        strTemp = "As of " & DATE1 & " through " & Format(DATE2, "mmmm d, yyyy")
    Else
        strTemp = "No date entered"
    End If
    exSheet.Cells(3, 2) = strTemp
    
    'Print Division Name
    If Len(clpDiv.Text) > 0 Then
        exSheet.Cells(4, 2) = clpDiv.Caption
    End If

    'Column Headings
    exSheet.Cells(5, 1) = "Employee Name"
    
    startdt = cmbMonth.Text & " " & txtYear.Text
    enddt = cmbMonthTo.Text & " " & txtYearTo.Text
    
    colno = 1
    mons = 0    'Months to add
    'exSheet.Cells(5, 2) = cmbMonth.Text & " " & txtYear 'First Month
    Month1 = DATE1
    
    
    'Ran only once if the Start period and End period is same - it won't do the following
    'do while loop
    If enddt = startdt Then
        'Print Month and year
        colno = colno + 1
        startdt = Format(DateAdd("m", mons, DATE1), "mmmm yyyy")
        exSheet.Cells(5, colno) = "'" & startdt
        
        colno = colno - 1
        
        'Print Day of the month
        'Get last date of the month
        Month2 = DateAdd("m", mons + 1, DATE1)
        Month2 = DateAdd("d", -1, Month2)
        For x = 1 To Day(Month2)
            If Weekday(Format(month(Month2) & "/" & x & "/" & Year(Month2), "mm/dd/yyyy")) <> vbSaturday And _
                Weekday(Format(month(Month2) & "/" & x & "/" & Year(Month2), "mm/dd/yyyy")) <> vbSunday Then
                colno = colno + 1
                exSheet.Cells(6, colno) = x
            End If
        Next
    
        startdt = cmbMonth.Text & " " & txtYear.Text
        enddt = cmbMonthTo.Text & " " & txtYearTo.Text
    End If
    
    'Rest of the months
    Do While enddt <> startdt
        'Print Month and year
        colno = colno + 1
        startdt = Format(DateAdd("m", mons, DATE1), "mmmm yyyy")
        exSheet.Cells(5, colno) = "'" & startdt
        
        colno = colno - 1
        
        'Print Day of the month
        'Get last date of the month
        Month2 = DateAdd("m", mons + 1, DATE1)
        Month2 = DateAdd("d", -1, Month2)
        For x = 1 To Day(Month2)
            If Weekday(Format(month(Month2) & "/" & x & "/" & Year(Month2), "mm/dd/yyyy")) <> vbSaturday And _
                Weekday(Format(month(Month2) & "/" & x & "/" & Year(Month2), "mm/dd/yyyy")) <> vbSunday Then
                colno = colno + 1
                exSheet.Cells(6, colno) = x
            End If
        Next
        mons = mons + 1
    Loop
        
    'Get recordset to print the data in excel spreadsheet
'    SQLQ = "SELECT HR_ATTCAL.*, HREMP.ED_EMPNBR,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREMP.ED_DIV, "
'    SQLQ = SQLQ & "HREMP.ED_DEPTNO,HREMP.ED_DOH "
'    SQLQ = SQLQ & "FROM HR_ATTCAL " & in_SQL(glbIHRDBW) & ",HREMP "
'    SQLQ = SQLQ & "WHERE "
'    SQLQ = SQLQ & " (HR_ATTCAL.AD_EMPNBR = HREMP.ED_EMPNBR) "
'    SQLQ = SQLQ & "ORDER BY HREMP.ED_DIV,HREMP.ED_DEPTNO,HREMP.ED_SURNAME,HREMP.ED_FNAME,HR_ATTCAL.AD_DOA "

    'Update HR_ATTCAL with Employee Name, Division Name and Department Name
    Dim rsDiv As New ADODB.Recordset
    Dim rsDEPT As New ADODB.Recordset
    Dim SSNo
    Dim tmpDiv, tmpDept
    
    SQLQ = "SELECT * FROM HR_ATTCAL WHERE AD_WRKEMP = '" & glbUserID & "' ORDER BY AD_EMPNBR, AD_DOA"
    rsHRAtt.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
    Do While Not rsHRAtt.EOF
        
        rsHREmp.Open "SELECT ED_FNAME, ED_SURNAME, ED_DIV, ED_DEPTNO FROM HREMP WHERE ED_EMPNBR = " & rsHRAtt("AD_EMPNBR"), gdbAdoIhr001, adOpenStatic
        rsHRAtt("Emp_Name") = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
        'rsHREmp.Close
        
        rsDiv.Open "SELECT Division_Name FROM HR_DIVISION WHERE DIV = '" & rsHREmp("ED_DIV") & "'", gdbAdoIhr001, adOpenStatic
        If rsDiv.EOF = False Then
            rsHRAtt("Div_Name") = rsDiv("Division_Name")
        End If
        rsDiv.Close
        
        rsDEPT.Open "SELECT DF_NAME FROM HRDEPT WHERE DF_NBR = '" & rsHREmp("ED_DEPTNO") & "'", gdbAdoIhr001, adOpenStatic
        If rsDEPT.EOF = False Then
            rsHRAtt("Dept_Name") = rsDEPT("DF_NAME")
        End If
        rsDEPT.Close
        rsHREmp.Close

        rsHRAtt.Update
        
        rsHRAtt.MoveNext
    Loop
    rsHRAtt.Close

    MDIMain.panHelp(0).FloodPercent = 40

    'SQLQ = "SELECT * FROM HR_ATTCAL WHERE AD_WRKEMP = '" & glbUserID & "' ORDER BY AD_EMPNBR, AD_DOA"
    SQLQ = "SELECT * FROM HR_ATTCAL WHERE AD_WRKEMP = '" & glbUserID & "' ORDER BY Div_Name, Dept_Name, Emp_Name, AD_EMPNBR, AD_DOA"
    rsHRAtt.Open SQLQ, gdbAdoIhr001W, adOpenStatic
    lineno = 6
    lstEmpNo = 0
    tmpDiv = ""
    tmpDept = ""
    Do While Not rsHRAtt.EOF
        If lstEmpNo <> rsHRAtt("AD_EMPNBR") Then
            lineno = lineno + 1
            lstEmpNo = rsHRAtt("AD_EMPNBR")
            colno = 1
            
            'rsHREmp.Open "SELECT ED_SURNAME,ED_FNAME FROM HREMP WHERE ED_EMPNBR = " & rsHRAtt("AD_EMPNBR"), gdbAdoIhr001, adOpenStatic
            'If Not rsHREmp.EOF Then
                'exSheet.Cells(lineno, colno) = rsHREmp("ED_SURNAME") & ", " & rsHREmp("ED_FNAME")
                
                If tmpDiv <> rsHRAtt("Div_Name") Then
                    If lineno <> 7 Then
                        'exSheet.Cells(lineno, colno) = "Division: " & rsHRAtt("Div_Name")
                        exSheet.Cells(lineno, colno) = "."
                        lineno = lineno + 1
                    End If
                    tmpDiv = rsHRAtt("Div_Name")
                ElseIf tmpDept <> rsHRAtt("Dept_Name") Then
                    If lineno <> 7 Then
                        'exSheet.Cells(lineno, colno) = "Department: " & rsHRAtt("Dept_Name")
                        exSheet.Cells(lineno, colno) = "."
                        lineno = lineno + 1
                    End If
                    tmpDept = rsHRAtt("Dept_Name")
                End If
                exSheet.Cells(lineno, colno) = rsHRAtt("Emp_Name")
                colno = colno + 1
            'End If
            'rsHREmp.Close
        End If
        
        SSNo = 1
        For x = 1 To 31
            If (rsHRAtt("AD_DAY" & x) <> "*" And Left(Trim(rsHRAtt("AD_DAY" & x)), 1) <> "S" And Trim(rsHRAtt("AD_DAY" & x)) <> "S ") Or IsNull(rsHRAtt("AD_DAY" & x)) Then
                If Not IsNull(rsHRAtt("AD_DAY" & x)) Then
                    adDay = StripChar(rsHRAtt("AD_DAY" & x), Chr(10))
                    adDay = StripChar(rsHRAtt("AD_DAY" & x), Chr(13))
                Else
                    adDay = " "
                End If
                exSheet.Cells(lineno, colno) = adDay
                colno = colno + 1
            ElseIf Left(Trim(rsHRAtt("AD_DAY" & x)), 1) = "S" Or Trim(rsHRAtt("AD_DAY" & x)) = "S " Then
                If SSNo <> 2 Then
                    exSheet.Cells(999, colno) = "S"
                    SSNo = SSNo + 1
                Else
                    SSNo = 1
                End If
            End If
        Next
        rsHRAtt.MoveNext
    Loop
    rsHRAtt.Close
    
    MDIMain.panHelp(0).FloodPercent = 90
    
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Visible = True
    Set exApp = Nothing
        
    Call Pause(1)
    'launch Excel file
    'Shell "Start " & GetShortName(xlsFileMat)
'    If Not LanchXlsW98(xlsFileMat) Then
'        Shell "cmd /c " & GetShortName(xlsFileMat)
'    End If
    
    MDIMain.panHelp(0).FloodPercent = 100

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Exit Sub
    
End Sub

Function LanchXlsW98(xFileName)
On Error GoTo Error_Deal
    LanchXlsW98 = False
    Shell "Start " & GetShortName(xFileName)
    LanchXlsW98 = True
Exit Function
Error_Deal:

End Function

Public Function GetShortName(ByVal sLongFileName As String) As String
    Dim lRetVal As Long, sShortPathName As String, iLen As Integer
    'Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)

    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    'Strip away unwanted characters.
    GetShortName = Left(sShortPathName, lRetVal)
End Function

Function StripChar(StringToStrip, CharToStrip)
    Dim I, buf, OneChar
    
    For I = 1 To Len(StringToStrip)
        OneChar = Mid(StringToStrip, I, 1)
        If OneChar <> CharToStrip Then buf = buf & OneChar
    Next I
    StripChar = buf
End Function


Private Sub AttWrk2()
 Dim SQLQ, ISQLQ
Dim rsAT As New ADODB.Recordset
Dim rsAW As New ADODB.Recordset
Dim rsHL As New ADODB.Recordset
Dim xEMPNBR, xDOA, xField
Dim xxx, xx1, x
Dim Y
Dim xDate
Dim xWeekDay
Dim xMons, DATEx, z
Dim xDay, xBuf
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 20
MDIMain.panHelp(1).Caption = " Please Wait"
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute "DELETE FROM HR_ATTCAL2 " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.CommitTrans


    DATE1 = cmbMonth & " 1, " & txtYear
    DATE2 = cmbMonthTo & " 1, " & txtYear
    DATEx = DATE1
    xMons = DateDiff("m", DATE1, DATE2)
    DATE2 = DateAdd("m", 1, DATE2)
    DATE2 = DateAdd("d", -1, DATE2)

    For x = 0 To xMons '
        z = month(DATEx) - 1
        SQLQ = " SELECT ED_COMPNO AS AD_COMPNO,'" & glbUserID & "' AS AD_WRKEMP,ED_EMPNBR AS AD_EMPNBR ,"
        SQLQ = SQLQ & Date_SQL(cmbMonth.List(z) & " 1," & Year(DATEx)) & " AS AD_DOA"
        SQLQ = SQLQ & " FROM HREMP "
        SQLQ = SQLQ & " WHERE " & glbstrSelCri
        If Not chkShowEmp Then
            SQLQ = SQLQ & " AND ED_EMPNBR IN ("
            If glbOracle Then
                SQLQ = SQLQ & " SELECT DISTINCT AD_EMPNBR FROM HR_ATTENDANCE WHERE TO_CHAR(AD_DOA,'YYYY')=" & Year(DATEx)
                
                If Len(txtShift.Text) > 0 Then
                    SQLQ = SQLQ & " AND AD_SHIFT= '" & txtShift.Text & "'"
                End If
                If Len(clpCode(10).Text) > 0 And glbCompSerial <> "S/N - 2362W" Then
                    SQLQ = SQLQ & " AND AD_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
                End If
                If chkAbsence.Value Then
                    SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
                End If
                SQLQ = SQLQ & " UNION "
                SQLQ = SQLQ & " SELECT DISTINCT AH_EMPNBR FROM HR_ATTENDANCE_HISTORY WHERE TO_CHAR(AH_DOA,'YYYY')=" & Year(DATEx)
            
                If Len(txtShift.Text) > 0 Then
                    SQLQ = SQLQ & " AND AH_SHIFT= '" & txtShift.Text & "'"
                End If
                
                If Len(clpCode(10).Text) > 0 Then SQLQ = SQLQ & " AND AH_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
                If chkAbsence.Value Then
                    SQLQ = SQLQ & " AND AH_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
                End If
            ElseIf glbSQL Then
                SQLQ = SQLQ & " SELECT DISTINCT AD_EMPNBR FROM HR_ATTENDANCE WHERE YEAR(AD_DOA)=" & Year(DATEx)
                
                If Len(txtShift.Text) > 0 Then
                    SQLQ = SQLQ & " AND AD_SHIFT= '" & txtShift.Text & "'"
                End If
                
                If Len(clpCode(10).Text) > 0 And glbCompSerial <> "S/N - 2362W" Then
                    SQLQ = SQLQ & " AND AD_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
                End If
                If chkAbsence.Value Then
                    SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
                End If
                SQLQ = SQLQ & " UNION "
                SQLQ = SQLQ & " SELECT DISTINCT AH_EMPNBR FROM HR_ATTENDANCE_HISTORY WHERE YEAR(AH_DOA)=" & Year(DATEx)
                
                If Len(txtShift.Text) > 0 Then
                    SQLQ = SQLQ & " AND AH_SHIFT= '" & txtShift.Text & "'"
                End If
                
                If Len(clpCode(10).Text) > 0 And glbCompSerial <> "S/N - 2362W" Then
                    SQLQ = SQLQ & " AND AH_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
                End If
                If chkAbsence.Value Then
                    SQLQ = SQLQ & " AND AH_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
                End If
            Else
                SQLQ = SQLQ & " SELECT DISTINCT AD_EMPNBR FROM QRY_EMPNBR_ATTENDANCE WHERE YEAR(AD_DOA)=" & Year(DATEx)
                
                If Len(txtShift.Text) > 0 Then
                    SQLQ = SQLQ & " AND AD_SHIFT= '" & txtShift.Text & "'"
                End If
                    
                If Len(clpCode(10).Text) > 0 And glbCompSerial <> "S/N - 2362W" Then
                    SQLQ = SQLQ & " AND AD_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
                End If
                
                If chkAbsence.Value Then
                    SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
                End If
            End If
            SQLQ = SQLQ & " )"
        End If
        ISQLQ = "INSERT INTO HR_ATTCAL2(AD_COMPNO,AD_WRKEMP,AD_EMPNBR,AD_DOA) " & in_SQL(glbIHRDBW) & SQLQ
        'Debug.Print ISQLQ
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute ISQLQ
        gdbAdoIhr001.CommitTrans
        For Y = 1 To 31
            'xDATE = cmbMonth.List(X) & " " & Y & ", " & txtYear
            xDate = cmbMonth.List(z) & " " & Y & ", " & Year(DATEx)
            If IsDate(xDate) Then
                'xWeekDay = Left(WeekdayName(Weekday(CVDate(xDATE))), 1)
                xWeekDay = Left(WeekdayName(Weekday(CVDate(xDate))), 3)
                xBuf = ""
                'If xWeekDay = "S" Then
                If UCase(xWeekDay) = "SAT" Or UCase(xWeekDay) = "SUN" Then
                    'SQLQ = "UPDATE HR_ATTCAL2 " & in_SQL(glbIHRDBW) & "SET AD_DAY" & Y & "='S ' "
'                    SQLQ = "UPDATE HR_ATTCAL2 " & in_SQL(glbIHRDBW) & "SET AD_DAY" & Y & "='" & UCase(xWeekDay) & "'"
'                    If glbOracle Then
'                        SQLQ = SQLQ & " WHERE TO_NUMBER(TO_CHAR(AD_DOA,'MM'))= " & (z + 1)
'                    Else
'                        SQLQ = SQLQ & " WHERE MONTH(AD_DOA)= " & (z + 1)
'                    End If
'                    gdbAdoIhr001.Execute SQLQ
                    xBuf = UCase(xWeekDay)
                End If
                rsHL.Open "SELECT * FROM HR_HOLIDAY WHERE HL_DATE=" & Date_SQL(xDate), gdbAdoIhr001, adOpenStatic
                If Not rsHL.EOF Then
'                    SQLQ = "UPDATE HR_ATTCAL2 " & in_SQL(glbIHRDBW) & " SET AD_DAY" & Y & "='H '"
'                    If glbOracle Then
'                         SQLQ = SQLQ & " WHERE TO_NUMBER(TO_CHAR(AD_DOA,'MM'))=" & (z + 1)
'                    Else
'                         SQLQ = SQLQ & " WHERE MONTH(AD_DOA)=" & (z + 1)
'                    End If
'                    gdbAdoIhr001.Execute SQLQ
                    If Len(xBuf) > 0 Then
                        xBuf = xBuf & Chr$(10) & " H "
                    Else
                        xBuf = xBuf & " H "
                    End If
                End If
                rsHL.Close
                If Len(xBuf) > 0 Then
                    SQLQ = "UPDATE HR_ATTCAL2 " & in_SQL(glbIHRDBW) & " SET AD_DAY" & Y & "='" & xBuf & "'"
                    If glbOracle Then
                         SQLQ = SQLQ & " WHERE TO_NUMBER(TO_CHAR(AD_DOA,'MM'))=" & (z + 1)
                    Else
                         SQLQ = SQLQ & " WHERE MONTH(AD_DOA)=" & (z + 1)
                    End If
                    gdbAdoIhr001.Execute SQLQ
                End If
             Else
                SQLQ = "UPDATE HR_ATTCAL2 " & in_SQL(glbIHRDBW) & " SET AD_DAY" & Y & "='*'"
                If glbOracle Then
                    SQLQ = SQLQ & " WHERE TO_NUMBER(TO_CHAR(AD_DOA,'MM'))=" & (z + 1)
                Else
                    SQLQ = SQLQ & " WHERE MONTH(AD_DOA)=" & (z + 1)
                End If
                gdbAdoIhr001.Execute SQLQ
            End If
        Next
        DATEx = DateAdd("M", 1, DATEx)
    Next

MDIMain.panHelp(0).FloodPercent = 30
SQLQ = "SELECT AD_COMPNO,'" & glbUserID & "' AS AD_WRKEMP,AD_EMPNBR,AD_DOA,AD_HRS,AD_REASON, AD_SHIFT "
SQLQ = SQLQ & " FROM HR_ATTENDANCE "
SQLQ = SQLQ & " WHERE AD_EMPNBR IN (SELECT AD_EMPNBR FROM HR_ATTCAL2" & in_SQL(glbIHRDBW) & ") "
SQLQ = SQLQ & " AND AD_DOA>=" & Date_SQL(DATE1)
SQLQ = SQLQ & " AND AD_DOA<=" & Date_SQL(DATE2)

If Len(clpCode(10).Text) > 0 Then SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
'Franks 06/20/2003 ticket# 4125 Absence
If chkAbsence.Value Then
    SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
End If
SQLQ = SQLQ & " UNION "
SQLQ = SQLQ & " SELECT AH_COMPNO,'" & glbUserID & "' AS AH_WRKEMP,AH_EMPNBR,AH_DOA,AH_HRS,AH_REASON, AH_SHIFT "
SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
SQLQ = SQLQ & " WHERE AH_EMPNBR IN (SELECT AD_EMPNBR FROM HR_ATTCAL2 " & in_SQL(glbIHRDBW) & ") "
SQLQ = SQLQ & " AND AH_DOA>=" & Date_SQL(DATE1)
SQLQ = SQLQ & " AND AH_DOA<=" & Date_SQL(DATE2)

If Len(clpCode(10).Text) > 0 Then
    SQLQ = SQLQ & " AND HR_ATTENDANCE_HISTORY.AH_REASON in ('" & Replace(clpCode(10).Text, ",", "','") & "') "
End If
'Franks 06/20/2003 ticket# 4125 Absence
If chkAbsence.Value Then
    SQLQ = SQLQ & " AND HR_ATTENDANCE_HISTORY.AH_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0) "
End If

rsAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockReadOnly
xxx = rsAT.RecordCount
xx1 = 0
Do Until rsAT.EOF
    xx1 = xx1 + 1
    MDIMain.panHelp(0).FloodPercent = (xx1 / xxx) * 60 + 30
    'xField = "AD_DAY" & Day(rsAT!AD_DOA)
    If optWeek Or optBiWeek Then
         xDay = DateDiff("d", DATE1, rsAT!AD_DOA) + 1
    Else
        xDay = Day(rsAT!AD_DOA)
    End If
    xField = "AD_DAY" & xDay

    SQLQ = "UPDATE HR_ATTCAL2 " & in_SQL(glbIHRDBW) & " SET " & xField & "="
    If glbSQL Or glbOracle Then
        SQLQ = SQLQ & " (CASE WHEN " & xField & " IS NULL THEN '' ELSE " & xField & " END )"
    Else
        SQLQ = SQLQ & " IIF(" & xField & " IS NULL,''," & xField & ")"
    End If

    If glbOracle Then
        SQLQ = SQLQ & " || Chr(13) || Chr(10) || '" & rsAT!AD_HRS & "' || Chr(13)|| Chr(10)||'" & rsAT!AD_REASON & "'"
        SQLQ = SQLQ & " WHERE AD_EMPNBR=" & rsAT!AD_EMPNBR
        SQLQ = SQLQ & " AND TO_CHAR(AD_DOA,'MM')= TO_CHAR(" & Date_SQL(rsAT!AD_DOA) & ",'MM') "
        SQLQ = SQLQ & " AND TO_CHAR(AD_DOA,'YYYY')=TO_CHAR(" & Date_SQL(rsAT!AD_DOA) & ",'YYYY')"
    Else
        'Hemu - 02/16/2004 Begin - Ticket # 5600
        If glbCompSerial = "S/N - 2226W" Then
            SQLQ = SQLQ & " +'" & Chr(13) & Chr(10) & rsAT!AD_HRS & rsAT!AD_SHIFT & Chr(13) & Chr(10) & rsAT!AD_REASON & "'"
        ElseIf glbCompSerial = "S/N - 2241W" And optBiWeek Then
            If chkBlank Then
                SQLQ = SQLQ & " +''"
            Else
                SQLQ = SQLQ & " +(CASE WHEN " & xField & " IS NULL THEN '' ELSE '" & Chr(13) & Chr(10) & "' END )+'" & rsAT!AD_HRS & Chr(13) & Chr(10) & rsAT!AD_REASON & "'"
            End If
        Else
            SQLQ = SQLQ & " +'" & Chr(13) & Chr(10) & rsAT!AD_HRS & Chr(13) & Chr(10) & rsAT!AD_REASON & "'"
        End If
        'Hemu - 02/16/2004 End

        SQLQ = SQLQ & " WHERE AD_EMPNBR=" & rsAT!AD_EMPNBR
        If optWeek Or optBiWeek Then
            SQLQ = SQLQ & " AND AD_DOA=" & Date_SQL(DATE1)
        Else

            SQLQ = SQLQ & " AND MONTH(AD_DOA)=" & month(rsAT!AD_DOA)
            SQLQ = SQLQ & " AND YEAR(AD_DOA)=" & Year(rsAT!AD_DOA)
        End If
    End If


    SQLQ = SQLQ & " AND AD_WRKEMP='" & glbUserID & "'"

    gdbAdoIhr001.Execute SQLQ
    rsAT.MoveNext
Loop
MDIMain.panHelp(0).FloodPercent = 90

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
End Sub

Private Sub WriteTo_XLS_VacSchedule()
    Dim rsHREmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum, x, z
    Dim xDate
    Dim xHourlyRate
    Dim xReptAuth As String
    Dim xAccFrwDate As Date
    Dim xTotOTBal, xTotVacBal
    Dim xSumOTBal, xSumVacBal
    Dim xMonthNum, xMonthName
    Dim xMonth, xYear
    Dim xEmpNoOld, xEmpNoCur
    Dim xEmpListIndex, xTemp, xEmpName As String
    Dim xGroupField As String, xGroupName As String, xOldGroupName As String
    
    Dim xExcelRptPath  As String
    
    'Ticket #22034 - Get Excel reports path
    If gsTRAININGMATRIX Then
        xExcelRptPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xExcelRptPath) = 0 Then
        xExcelRptPath = glbIHRREPORTS
    End If
    
    xGroupField = getGroupField(comGroup(0).Text)
    
    SQLQ = "SELECT HR_ATTCAL.*, HREMP.* FROM HR_ATTCAL LEFT JOIN HREMP ON HR_ATTCAL.AD_EMPNBR = HREMP.ED_EMPNBR WHERE AD_WRKEMP = '" & glbUserID & "' "
    SQLQ = SQLQ & "ORDER BY  "
    If Len(xGroupField) > 0 Then
        SQLQ = SQLQ & xGroupField & ", "
    End If
    SQLQ = SQLQ & "ED_SURNAME, ED_FNAME "
    SQLQ = SQLQ & ",AD_EMPNBR, AD_DOA "
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: I = 0
        rsHREmp.MoveFirst

        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "VacScheduleTmp.xls"
        
        'Ticket #22034 - May need to save report in different path
        'xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "VacSchedule" & Trim(glbUserID) & ".xls"
        xlsFileMat = xExcelRptPath & IIf(Right(xExcelRptPath, 1) = "\", "", "\") & "VacSchedule" & Trim(glbUserID) & ".xls"
    
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
    
        'Sweep the default monthes
        exSheet.Cells(3, 4 + 32 * 0) = ""  'month 1
        exSheet.Cells(3, 4 + 32 * 1) = ""  'month 2
        exSheet.Cells(3, 4 + 32 * 2) = ""  'month 3
        exSheet.Cells(3, 4 + 32 * 3) = ""  'month 4
        exSheet.Cells(3, 4 + 32 * 4) = ""  'month 5
        exSheet.Cells(3, 4 + 32 * 5) = ""  'month 6
        
        'Color list: Pink - 7; lime green - 10; Green - 4
        xOldGroupName = "*"
        xMonthNum = 0: xMonthName = ""
        xEmpNoOld = 0: xEmpNoCur = 0
        xEmpListIndex = 5
        Do While Not rsHREmp.EOF
            xEmpNoCur = rsHREmp("AD_EMPNBR")

            If xEmpNoCur <> xEmpNoOld Then
                xEmpName = GetEmpName(xEmpNoCur)
                xEmpListIndex = xEmpListIndex + 1
                xMonthNum = 1
                xMonthName = MonthName(month(rsHREmp("AD_DOA"))) '& " " & Year(rsHREmp("AD_DOA"))
                xEmpNoOld = xEmpNoCur
                xMonth = month(rsHREmp("AD_DOA")): xYear = Year(rsHREmp("AD_DOA"))
                exSheet.Cells(3, 4) = xMonthName & " " & xYear
            Else 'Same employee but different month
                xMonthNum = xMonthNum + 1
                xMonthName = MonthName(month(rsHREmp("AD_DOA")))
                xMonth = month(rsHREmp("AD_DOA")): xYear = Year(rsHREmp("AD_DOA"))
            End If
            
            'Group
            If Len(xGroupField) > 0 Then
                xGroupName = GetGroupName(xEmpNoCur, xGroupField)
                If Not (xOldGroupName = xGroupName) Then
                    'If xOldGroupName = "*" Then
                    '    exSheet.Cells(3, 1) = xGroupName
                    'Else
                        exSheet.Cells(xEmpListIndex, 1) = xGroupName
                        exSheet.Cells(xEmpListIndex, 1).Font.Bold = True
                        xEmpListIndex = xEmpListIndex + 1
                        '.Font.Underline = True
                        '.Font.name = "Times New Roman"
                        '.Font.Bold = True
                        '.Font.Size = 15
                    'End If
                    xOldGroupName = xGroupName
                End If
            End If
            
            'Set the Month + Year
            exSheet.Cells(3, 4 + 32 * (xMonthNum - 1)) = xMonthName & " " & xYear
            If xMonthNum = 1 Then
                exSheet.Cells(xEmpListIndex, 1) = xEmpName
                If xVacEnt > 0 Then exSheet.Cells(xEmpListIndex, 2) = xVacEnt
                If xVacTaken > 0 Then exSheet.Cells(xEmpListIndex, 3) = xVacTaken
            Else
                exSheet.Cells(xEmpListIndex, 3 + 32 * (xMonthNum - 1)) = xEmpName
            End If
            For x = 1 To 31
                If Not IsNull(rsHREmp("AD_DAY" & x)) Then
                    xTemp = rsHREmp("AD_DAY" & x)
                    If xTemp = "SAT" Or xTemp = "SUN" Then
                        exSheet.Cells(xEmpListIndex, 3 + x + 32 * (xMonthNum - 1)) = ""
                        exSheet.Cells(xEmpListIndex, 3 + x + 32 * (xMonthNum - 1)).Interior.ColorIndex = 36
                    End If
                    
                    'Delete the SCHV if this day has bee taken: both VAC and SCHV on this day
                    xTemp = Replace(xTemp, Chr$(10), "")
                    xTemp = Replace(xTemp, Chr$(13), "")
                    xTemp = Replace(xTemp, "SAT", "")
                    xTemp = Replace(xTemp, "SUN", "")
                    If InStr(xTemp, "VAC") > 0 And InStr(xTemp, "SCHV") > 0 Then
                        '7.5SCHV7.5VAC or 7.5VAC7.5SCHV
                        If Trim(Right(xTemp, 3)) = "VAC" Then
                            I = InStr(xTemp, "SCHV") + 4
                            xTemp = Mid(xTemp, I, Len(xTemp) - I + 1)
                        End If
                        If Trim(Right(xTemp, 4)) = "SCHV" Then
                            I = InStr(xTemp, "VAC") + 3
                            xTemp = Mid(xTemp, 1, Len(xTemp) - I)
                        End If
                    End If
                    
                    If Trim(Right(xTemp, 3)) = "VAC" Then
                        xTemp = Trim(Left(xTemp, Len(xTemp) - 3))
                        exSheet.Cells(xEmpListIndex, 3 + x + 32 * (xMonthNum - 1)) = xTemp
                        exSheet.Cells(xEmpListIndex, 3 + x + 32 * (xMonthNum - 1)).Interior.ColorIndex = 7
                    End If
                    If Trim(Right(xTemp, 4)) = "SCHV" Then
                        xTemp = Trim(Left(xTemp, Len(xTemp) - 4))
                        exSheet.Cells(xEmpListIndex, 3 + x + 32 * (xMonthNum - 1)) = xTemp
                        exSheet.Cells(xEmpListIndex, 3 + x + 32 * (xMonthNum - 1)).Interior.ColorIndex = 10
                    End If
                End If
            Next
            
            rsHREmp.MoveNext
        Loop
        
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
    
End Sub

Private Function GetEmpName(xEmpNo)
Dim rsTemp As New ADODB.Recordset
Dim xStr, SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME, ED_PVAC,ED_VAC,ED_VACT FROM HREMP WHERE ED_EMPNBR=" & xEmpNo
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        xStr = rsTemp("ED_SURNAME") & ", " & rsTemp("ED_FNAME")
        If Not IsNull(rsTemp("ED_VAC")) And Not IsNull(rsTemp("ED_PVAC")) Then
            xVacEnt = rsTemp("ED_PVAC") + rsTemp("ED_VAC")
        Else
            xVacEnt = 0
        End If
        If Not IsNull(rsTemp("ED_VACT")) Then
            xVacTaken = rsTemp("ED_VACT")
        Else
            xVacTaken = 0
        End If
    End If
    rsTemp.Close
    GetEmpName = xStr
End Function

Private Function getGroupField(xGrptxt)
Dim xTTemp As String
    xTTemp = ""
    Select Case xGrptxt
    Case lStr("Division")
        xTTemp = "ED_DIV"
    Case lStr("Department")
        xTTemp = "ED_DEPTNO"
    Case lStr("Region")
        xTTemp = "ED_REGION"
    End Select
    
    getGroupField = xTTemp
End Function

Private Function getGroupField1(xGrptxt)
Dim xTTemp As String
    xTTemp = ""
    Select Case xGrptxt
    Case lStr("Division")
        xTTemp = "ED_DIV"
    Case lStr("Department")
        xTTemp = "ED_DEPTNO"
    Case lStr("Region")
        xTTemp = "ED_REGION"
    Case "Employee Name"
        xTTemp = "ED_SURNAME, ED_FNAME"
    Case "Shift"
        xTTemp = "ED_SHIFT"
    'Case "Rept. Authority 1"
    '    xTTemp = "AD_SUPER"
    Case "(none)"
        xTTemp = ""
    End Select
    
    getGroupField1 = xTTemp
End Function

Private Function GetGroupName(xEmpNo, xGrpCode)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xCode As String, xCodeDesc As String
    
    xCodeDesc = ""
    SQLQ = "SELECT " & xGrpCode & " FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp(xGrpCode)) Then
            xCode = rsTemp(xGrpCode)
        End If
    End If
    rsTemp.Close
    If Len(xCode) > 0 Then
        If xGrpCode = "ED_DIV" Then
            SQLQ = "SELECT Division_Name FROM HR_DIVISION WHERE DIV = '" & xCode & "' "
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTemp.EOF Then
                xCodeDesc = rsTemp("Division_Name")
            End If
            rsTemp.Close
        End If
        If xGrpCode = "ED_DEPTNO" Then
            SQLQ = "SELECT DF_NAME FROM HRDEPT WHERE DF_NBR = '" & xCode & "' "
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTemp.EOF Then
                xCodeDesc = rsTemp("DF_NAME")
            End If
            rsTemp.Close
        End If
        If xGrpCode = "ED_REGION" Then
            SQLQ = "SELECT TB_DESC FROM HRTABL WHERE TB_NAME = 'EDRG' AND TB_KEY = '" & xCode & "' "
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTemp.EOF Then
                xCodeDesc = rsTemp("TB_DESC")
            End If
            rsTemp.Close
        End If
    End If
    GetGroupName = xCodeDesc
End Function

Private Sub WriteTo_XLS_AttendanceCalendar()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsATTCal As New ADODB.Recordset
    Dim rsHRTable As New ADODB.Recordset
    Dim exApp As Object 'Excel.Application
    Dim exBook As Object 'Excel.Workbook
    Dim exSheet As Object  'Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum, x, z
    Dim xDate
    Dim xHourlyRate
    Dim xReptAuth As String
    Dim xAccFrwDate As Date
    Dim xTotOTBal, xTotVacBal
    Dim xSumOTBal, xSumVacBal
    Dim xMonthNum, xMonthName
    Dim xMonth, xYear
    Dim xEmpNoOld, xEmpNoCur
    Dim xEmpListIndex, xTemp, xEmpName As String
    Dim xGroupField As String, xGroupName As String, xOldGroupName As String
    
    Dim xMonthCnt As Integer
    Dim xDays As Integer
    Dim xMonthYear(11) As Long
    Dim xMonthYearCol(11) As Integer
    Dim xStartCol As Integer
    
    Dim xExcelRptPath  As String
    
    'Ticket #29802 - Added Employee # column - shifted the rest of the columns.
    
    'Ticket #22034 - Get Excel reports path
    If gsTRAININGMATRIX Then
        xExcelRptPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xExcelRptPath) = 0 Then
        xExcelRptPath = glbIHRREPORTS
    End If
    
    xGroupField = getGroupField1(comGroup(0).Text)
    
    SQLQ = "SELECT HR_ATTCAL.*, HREMP.ED_SURNAME, HREMP.ED_FNAME " & IIf(Len(xGroupField) > 0, "," & xGroupField & "", "") & " FROM HR_ATTCAL LEFT JOIN HREMP ON HR_ATTCAL.AD_EMPNBR = HREMP.ED_EMPNBR WHERE AD_WRKEMP = '" & glbUserID & "' "
    SQLQ = SQLQ & "ORDER BY  "
    If Len(xGroupField) > 0 Then
        If xGroupField <> "ED_SURNAME, ED_FNAME" Then
            SQLQ = SQLQ & xGroupField & ", "
            SQLQ = SQLQ & "ED_SURNAME, ED_FNAME "
        Else
            SQLQ = SQLQ & "ED_SURNAME, ED_FNAME "
        End If
        SQLQ = SQLQ & ",AD_EMPNBR"
    Else
        If comGroup(1).Text = "Employee Name" Then
            SQLQ = SQLQ & "ED_SURNAME, ED_FNAME "
            SQLQ = SQLQ & ",AD_EMPNBR"
        ElseIf comGroup(1).Text = "Employee Number" Then
            'SQLQ = SQLQ & "ED_SURNAME, ED_FNAME "
            SQLQ = SQLQ & "AD_EMPNBR"
        End If
    End If
    SQLQ = SQLQ & ",AD_DOA "
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: I = 0
        rsHREmp.MoveFirst

        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "CalendarTmp.xls"
        
        'Ticket #22034 - May need to save the report in different path
        'xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "Calendar" & Trim(glbUserID) & ".xls"
        xlsFileMat = xExcelRptPath & IIf(Right(xExcelRptPath, 1) = "\", "", "\") & "Calendar" & Trim(glbUserID) & ".xls"
    
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
    
        'Display the Report Heading
        If optMonth Then
            exSheet.Cells(1, 1) = "Attendance Calendar for Period: " & cmbMonth.Text & " " & txtYear
        Else
            exSheet.Cells(1, 1) = "Attendance Calendar for Period: " & cmbMonth.Text & " " & txtYear & " - " & cmbMonthTo.Text & " " & txtYearTo
        End If
        
        'Get the list of distinct months and years to setup the columns headings
        SQLQ = "SELECT MONTH(AD_DOA) AS DOA_MONTH,YEAR(AD_DOA) AS DOA_YEAR"
        SQLQ = SQLQ & " FROM HR_ATTCAL"
        SQLQ = SQLQ & " WHERE AD_WRKEMP = '" & glbUserID & "' "
        SQLQ = SQLQ & " GROUP BY MONTH(AD_DOA), YEAR(AD_DOA)"
        SQLQ = SQLQ & " ORDER BY YEAR(AD_DOA), MONTH(AD_DOA) ASC"
        rsATTCal.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsATTCal.EOF Then
            '# of Months
            xMonthCnt = 3   '2 - Ticket #29813 - Start from column after as we are adding Employee #
            
            'Clear the array
            For x = 0 To 11
                xMonthYear(x) = 0
                xMonthYearCol(x) = 0
            Next
            
            x = 0
            
            'Set Month/Year Names
            Do While Not rsATTCal.EOF
                'Display the Month Name & Year
                exSheet.Cells(3, xMonthCnt) = MonthName(rsATTCal("DOA_MONTH")) & " " & rsATTCal("DOA_YEAR")
                                
                'Populate the arrays containing the monthyear and starting column
                xMonthYear(x) = rsATTCal("DOA_MONTH") & rsATTCal("DOA_YEAR")
                xMonthYearCol(x) = xMonthCnt
                
                'Merge the Month Name & Year cell and Center
                'exSheet.Range(exSheet.Cells(3, xMonthCnt), exSheet.Cells(3, xMonthCnt + DaysInMonth(CVDate(Format(rsAttCal("DOA_MONTH") & "/" & "01/" & rsAttCal("DOA_YEAR"), "mm/dd/yyyy"))))).Merge
                'exSheet.Range(exSheet.Cells(3, xMonthCnt), exSheet.Cells(3, xMonthCnt + DaysInMonth(CVDate(Format(rsAttCal("DOA_MONTH") & "/" & "01/" & rsAttCal("DOA_YEAR"), "mm/dd/yyyy"))))).HorizontalAlignment = xlCenter
                
                exSheet.Range(exSheet.Cells(3, xMonthCnt), exSheet.Cells(3, xMonthCnt + 30)).Merge
                exSheet.Range(exSheet.Cells(3, xMonthCnt), exSheet.Cells(3, xMonthCnt + 30)).HorizontalAlignment = xlCenter
                
                'Display the dates in a month in each cell
                'For xDays = 1 To DaysInMonth(CVDate(Format(rsAttCal("DOA_MONTH") & "/" & "01/" & rsAttCal("DOA_YEAR"), "mm/dd/yyyy")))
                For xDays = 1 To 31
                    exSheet.Cells(4, xMonthCnt + xDays - 1) = xDays
                Next
                
                'Add # of Days in a Month to MonthCnt to get the next cell to display Month Name & Year
                'xMonthCnt = xMonthCnt + DaysInMonth(CVDate(Format(rsAttCal("DOA_MONTH") & "/" & "01/" & rsAttCal("DOA_YEAR"), "mm/dd/yyyy")))
                xMonthCnt = xMonthCnt + 31 + 1
                
                x = x + 1
                rsATTCal.MoveNext
            Loop
        End If
        rsATTCal.Close
        Set rsATTCal = Nothing
        
        'Color list: Pink - 7; lime green - 10; Yellow - 4
        xOldGroupName = "*"
        xMonthNum = 0: xMonthName = ""
        xEmpNoOld = 0: xEmpNoCur = 0
        xEmpListIndex = 5
        
        Do While Not rsHREmp.EOF
            xEmpNoCur = rsHREmp("AD_EMPNBR")

            If xEmpNoCur <> xEmpNoOld Then
                xEmpName = GetEmpName(xEmpNoCur)
                xEmpListIndex = xEmpListIndex + 1
                xMonthNum = 1
                xMonthName = MonthName(month(rsHREmp("AD_DOA"))) '& " " & Year(rsHREmp("AD_DOA"))
                xEmpNoOld = xEmpNoCur
                
                xMonth = month(rsHREmp("AD_DOA")): xYear = Year(rsHREmp("AD_DOA"))
                'exSheet.Cells(3, 4) = xMonthName & " " & xYear
            Else 'Same employee but different month
                xMonthNum = xMonthNum + 1
                xMonthName = MonthName(month(rsHREmp("AD_DOA")))
                xMonth = month(rsHREmp("AD_DOA")): xYear = Year(rsHREmp("AD_DOA"))
            End If
            
            'Group
            If Len(xGroupField) > 0 And xGroupField <> "ED_SURNAME, ED_FNAME" Then
                xGroupName = GetGroupName(xEmpNoCur, xGroupField)
                If Not (xOldGroupName = xGroupName) Then
                    'If xOldGroupName = "*" Then
                    '    exSheet.Cells(3, 1) = xGroupName
                    'Else
                        exSheet.Cells(xEmpListIndex, 1) = xGroupName
                        exSheet.Cells(xEmpListIndex, 1).Font.Bold = True
                        xEmpListIndex = xEmpListIndex + 1
                        '.Font.Underline = True
                        '.Font.name = "Times New Roman"
                        '.Font.Bold = True
                        '.Font.Size = 15
                    'End If
                    xOldGroupName = xGroupName
                End If
            End If
            
            'Dispplay employee name
            If comGroup(1).ListIndex = 1 And comGroup(1).Text = "Employee Number" Then
                If chkShowEmpNo Then
                    exSheet.Cells(xEmpListIndex, 1) = rsHREmp("AD_EMPNBR")
                    'Ticket #29802 - Since added the Employee # column to the default report, if user choses to show only employee # then hide the Employee Name column.
                    exSheet.Range(exSheet.Cells(xEmpListIndex, 2), exSheet.Cells(xEmpListIndex, 2)).EntireColumn.Hidden = True
                Else
                    exSheet.Cells(xEmpListIndex, 1) = rsHREmp("AD_EMPNBR")
                    exSheet.Cells(xEmpListIndex, 2) = xEmpName
                End If
            Else
                exSheet.Cells(xEmpListIndex, 1) = rsHREmp("AD_EMPNBR")
                exSheet.Cells(xEmpListIndex, 2) = xEmpName
            End If
            
            'Retrieve the Column to start populating the hours for the day in the month/year
            xStartCol = 0
            For x = 0 To 11
                If xMonthYear(x) = month(rsHREmp("AD_DOA")) & Year(rsHREmp("AD_DOA")) Then
                    xStartCol = xMonthYearCol(x)
                    Exit For
                End If
            Next
            
            'Retrieving values from the fields for each day
            For x = 1 To 31
                If Not IsNull(rsHREmp("AD_DAY" & x)) Then
                    xTemp = rsHREmp("AD_DAY" & x)
                    If InStr(1, xTemp, "SAT") > 0 Or InStr(1, xTemp, "SUN") > 0 Then
                        exSheet.Cells(xEmpListIndex, xStartCol) = ""
                        exSheet.Cells(xEmpListIndex, xStartCol).Interior.ColorIndex = 15
                    ElseIf InStr(1, xTemp, " H ") > 0 Then
                        exSheet.Cells(xEmpListIndex, xStartCol) = ""
                        exSheet.Cells(xEmpListIndex, xStartCol).Interior.ColorIndex = 22
                    End If
                                        
                    'Clear unwanted text and spaces from the xTemp
                    xTemp = Replace(xTemp, Chr$(10), " ")
                    xTemp = Replace(xTemp, Chr$(13), "")
                    xTemp = Replace(xTemp, "SAT", "")
                    xTemp = Replace(xTemp, "SUN", "")
                    xTemp = Replace(xTemp, " H ", "")
                    xTemp = Replace(xTemp, "  ", " ")
                    
                    'Display the hours/reason code
                    exSheet.Cells(xEmpListIndex, xStartCol) = Trim(xTemp)
                End If
                
                xStartCol = xStartCol + 1
            Next
            
            rsHREmp.MoveNext
        Loop
        
        If Len(xAllReason) > 0 Then
            x = 5
            exSheet.Cells(xEmpListIndex + x, 1) = "Legend:"
            exSheet.Cells(xEmpListIndex + x, 1).Font.Bold = True
            SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_KEY IN (" & xAllReason & ")"
            rsHRTable.Open SQLQ, gdbAdoIhr001, adOpenStatic
            Do While Not rsHRTable.EOF
                x = x + 1
                exSheet.Cells(xEmpListIndex + x, 1) = rsHRTable("TB_KEY") & " = " & rsHRTable("TB_DESC")
                
                rsHRTable.MoveNext
            Loop
            rsHRTable.Close
            Set rsHRTable = Nothing
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
    End If
    rsHREmp.Close
    Set rsHREmp = Nothing
    
End Sub

